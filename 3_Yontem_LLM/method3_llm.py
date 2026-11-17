"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: 3. Yöntem - Büyük Dil Modeli (LLM) Tabanlı Duygu Analizi (EVREN API)

Bu modül:
1. EVREN OpenAI-uyumlu API ('https://evren-llmapi.ssyz.org.tr/v1') üzerinden
   'qwen3.8-flash-next' (ve yedek olarak 'gemma-4-31b') modellerini çağırır.
2. Yapılandırılmış istem mühendisliği (Structured Batch Prompting) ile
   9.000 İngilizce ve 9.000 Türkçe restoran/kafe yorumunu sınıflandırır.
3. Her iki dil ve duygu sınıfı (Positive, Neutral, Negative) bazında yapılan
   hatalı tahmin sayılarını ve performans metriklerini hesaplar.
"""

import concurrent.futures
import json
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import requests
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix


# ============================================================================
# SABİTLER VE KONFİGÜRASYON (CONSTANTS & CONFIGURATION)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
DATA_DIR: Path = BASE_DIR.parent / "Veri_Setleri"

ENGLISH_INPUT_CSV: Path = DATA_DIR / "english_restaurant_reviews_9000.csv"
TURKISH_INPUT_CSV: Path = DATA_DIR / "turkish_restaurant_reviews_9000.csv"

ENGLISH_OUTPUT_CSV: Path = BASE_DIR / "llm_english_results.csv"
TURKISH_OUTPUT_CSV: Path = BASE_DIR / "llm_turkish_results.csv"
SUMMARY_METRICS_JSON: Path = BASE_DIR / "llm_summary_metrics.json"

EVREN_API_URL: str = "https://evren-llmapi.ssyz.org.tr/v1/chat/completions"
EVREN_API_KEY: str = os.environ.get("EVREN_API_KEY", "")

PRIMARY_MODEL: str = "qwen3.8-flash-next"
FALLBACK_MODEL: str = "gemma-4-31b"

LABELS: List[str] = ["Positive", "Neutral", "Negative"]

LLM_BATCH_SIZE: int = 50
LLM_MAX_WORKERS: int = 8
REQUEST_TIMEOUT_SEC: int = 45


# ============================================================================
# İSTEM MÜHENDİSLİĞİ VE API İLETİŞİM FONKSİYONLARI
# ============================================================================

def normalize_sentiment_label(raw_label: str) -> str:
    """LLM çıktısındaki duygu etiketini standart 3 sınıftan birine normalize eder."""
    cleaned = str(raw_label).strip().lower()
    if "pos" in cleaned or "olumlu" in cleaned or "pozitif" in cleaned:
        return "Positive"
    if "neg" in cleaned or "olumsuz" in cleaned or "negatif" in cleaned:
        return "Negative"
    return "Neutral"


def build_batch_prompt(texts: List[str]) -> str:
    """Toplu yorum sınıflandırması için yapılandırılmış sıfır-örnek (zero-shot) istem üretir."""
    batch_len = len(texts)
    numbered_reviews = [
        f"{idx + 1}. {str(text).replace(chr(10), ' ')}"
        for idx, text in enumerate(texts)
    ]
    return (
        f"Classify the sentiment of each of the {batch_len} numbered restaurant reviews below "
        "into EXACTLY one of: Positive, Neutral, or Negative.\n"
        "- Positive: satisfied customer, praising food/service/atmosphere (4-5 star tone).\n"
        "- Neutral: mixed feelings, average/mediocre experience, pros and cons balanced (3 star tone).\n"
        "- Negative: dissatisfied customer, complaining about food/service/hygiene/price (1-2 star tone).\n"
        f"Return ONLY a valid JSON array of {batch_len} strings "
        '(each "Positive", "Neutral", or "Negative") in the exact same order. '
        "Do NOT include markdown formatting or any explanation.\n\n"
        + "\n".join(numbered_reviews)
    )


def classify_batch_with_llm(texts: List[str], api_key: str, max_retries: int = 4) -> List[str]:
    """
    Verilen yorum paketini EVREN LLM API'ye göndererek duygu tahminlerini alır.
    Paket boyutu uyuşmazlığı durumunda paketi ikiye bölerek özyinelemeli (recursive) tamamlar.
    """
    if not api_key:
        raise ValueError(
            "EVREN_API_KEY ortam değişkeni tanımlanmamış. "
            "Lütfen 'export EVREN_API_KEY=evren_llm_...' komutunu çalıştırın."
        )

    batch_len = len(texts)
    prompt = build_batch_prompt(texts)
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    for attempt in range(max_retries):
        active_model = PRIMARY_MODEL if attempt % 2 == 0 else FALLBACK_MODEL
        payload = {
            "model": active_model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.0,
            "max_tokens": max(800, batch_len * 12),
            "chat_template_kwargs": {"enable_thinking": False},
        }
        try:
            response = requests.post(
                EVREN_API_URL,
                headers=headers,
                json=payload,
                timeout=REQUEST_TIMEOUT_SEC,
            )
            if response.status_code == 429:
                time.sleep(2.0 * (attempt + 1))
                continue
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"].get("content", "")
                match = re.search(r"\[.*\]", content.strip(), re.DOTALL)
                if match:
                    parsed_array = json.loads(match.group(0))
                    if isinstance(parsed_array, list) and len(parsed_array) == batch_len:
                        return [normalize_sentiment_label(item) for item in parsed_array]
        except Exception:
            time.sleep(1.5 * (attempt + 1))

    if batch_len > 1:
        midpoint = batch_len // 2
        return (
            classify_batch_with_llm(texts[:midpoint], api_key, max_retries=3)
            + classify_batch_with_llm(texts[midpoint:], api_key, max_retries=3)
        )
    return ["Neutral"]


# ============================================================================
# DEĞERLENDİRME VE HATA ANALİZİ FONKSİYONLARI
# ============================================================================

def evaluate_predictions(
    df: pd.DataFrame, pred_col: str, title: str, model_name: str
) -> Dict[str, Any]:
    """
    LLM tahminlerini gerçek etiketlerle karşılaştırarak sınıf bazlı hatalı tahmin
    adetlerini, karmaşıklık matrisini ve F1 skorlarını hesaplar.
    """
    y_true = df["true_label"]
    y_pred = df[pred_col]

    accuracy = accuracy_score(y_true, y_pred)
    cm = confusion_matrix(y_true, y_pred, labels=LABELS)
    report = classification_report(y_true, y_pred, labels=LABELS, output_dict=True, digits=4)

    total_correct = int((y_true == y_pred).sum())
    total_wrong = int((y_true != y_pred).sum())

    print(f"\n{'=' * 60}")
    print(f" {title}")
    print(f" Model: {model_name} (EVREN API)")
    print(f"{'=' * 60}")
    print(f"Genel Doğruluk (Accuracy)   : %{accuracy * 100:.2f} ({total_correct:,} / {len(df):,} doğru)")
    print(f"Toplam Hatalı Tahmin Sayısı : {total_wrong:,} / {len(df):,}\n")

    class_errors: Dict[str, Dict[str, Any]] = {}
    print("Sınıf Bazlı Hata Analizi (Gerçek Sınıfa Göre):")
    for idx, label in enumerate(LABELS):
        total_cls = int(cm[idx].sum())
        correct_cls = int(cm[idx, idx])
        wrong_cls = total_cls - correct_cls
        error_rate_pct = (wrong_cls / total_cls) * 100 if total_cls else 0.0

        misclassified_as = {
            LABELS[col_idx]: int(cm[idx, col_idx])
            for col_idx in range(len(LABELS))
            if col_idx != idx
        }
        class_errors[label] = {
            "total": total_cls,
            "correct": correct_cls,
            "wrong": wrong_cls,
            "error_rate_pct": round(error_rate_pct, 2),
            "misclassified_as": misclassified_as,
        }
        breakdown_str = ", ".join([f"{k}'e kayan: {v}" for k, v in misclassified_as.items()])
        print(
            f"  - [{label:8s}] Toplam: {total_cls} | Doğru: {correct_cls:4d} | "
            f"Hatalı: {wrong_cls:4d} (%{error_rate_pct:5.2f} hata) -> ({breakdown_str})"
        )

    print("\nKarmaşıklık Matrisi (Confusion Matrix):")
    cm_df = pd.DataFrame(
        cm,
        index=[f"Gerçek {lbl}" for lbl in LABELS],
        columns=[f"Tahmin {lbl}" for lbl in LABELS],
    )
    print(cm_df)

    return {
        "title": title,
        "model_name": model_name,
        "accuracy": round(accuracy, 4),
        "macro_f1": round(report["macro avg"]["f1-score"], 4),
        "total_samples": len(df),
        "total_correct": total_correct,
        "total_wrong": total_wrong,
        "class_errors": class_errors,
        "classification_report": report,
        "confusion_matrix": cm.tolist(),
    }


def run_llm_pipeline(input_csv: Path, output_csv: Path, title: str) -> Dict[str, Any]:
    """
    Veri setini paralel paketler halinde LLM ile sınıflandırır (veya mevcut sonuç
    dosyası varsa önbellekten yükler) ve hata analizini gerçekleştirir.
    """
    if output_csv.exists():
        df_cached = pd.read_csv(output_csv)
        if "llm_pred" in df_cached.columns and df_cached["llm_pred"].notnull().all():
            print(f"Önbellekteki tamamlanmış sonuç dosyası kullanılıyor: {output_csv.name}")
            return evaluate_predictions(df_cached, "llm_pred", title, PRIMARY_MODEL)

    df = pd.read_csv(input_csv)
    texts = df["text"].astype(str).tolist()
    batches = [
        (idx, texts[idx : idx + LLM_BATCH_SIZE])
        for idx in range(0, len(texts), LLM_BATCH_SIZE)
    ]
    predictions: List[str] = [""] * len(texts)

    print(f"\n{title} başlatılıyor ({len(texts):,} yorum, {len(batches)} paket)...")
    start_time = time.time()
    completed = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=LLM_MAX_WORKERS) as executor:
        future_to_start_idx = {
            executor.submit(classify_batch_with_llm, batch_texts, EVREN_API_KEY): start_idx
            for start_idx, batch_texts in batches
        }
        for future in concurrent.futures.as_completed(future_to_start_idx):
            start_idx = future_to_start_idx[future]
            batch_result = future.result()
            predictions[start_idx : start_idx + len(batch_result)] = batch_result
            completed += 1
            if completed % 30 == 0 or completed == len(batches):
                elapsed = time.time() - start_time
                print(f"   İlerleme: {completed}/{len(batches)} paket tamamlandı ({elapsed:.1f} sn)")

    df["llm_pred"] = predictions
    df.to_csv(output_csv, index=False, encoding="utf-8")
    print(f"   -> Kaydedildi: {output_csv.name}")

    return evaluate_predictions(df, "llm_pred", title, PRIMARY_MODEL)


def main() -> None:
    """İngilizce ve Türkçe LLM analizlerini çalıştırır ve özet metrikleri kaydeder."""
    english_metrics = run_llm_pipeline(
        ENGLISH_INPUT_CSV,
        ENGLISH_OUTPUT_CSV,
        "İNGİLİZCE VERİ SETİ - LLM TABANLI DUYGU ANALİZİ SONUÇLARI (9.000 Yorum)",
    )
    turkish_metrics = run_llm_pipeline(
        TURKISH_INPUT_CSV,
        TURKISH_OUTPUT_CSV,
        "TÜRKÇE VERİ SETİ - LLM TABANLI DUYGU ANALİZİ SONUÇLARI (9.000 Yorum)",
    )

    summary_payload = {
        "english_llm": english_metrics,
        "turkish_llm": turkish_metrics,
    }
    with open(SUMMARY_METRICS_JSON, "w", encoding="utf-8") as file:
        json.dump(summary_payload, file, ensure_ascii=False, indent=2)
    print(f"\nÖzet metrikler kaydedildi: {SUMMARY_METRICS_JSON.name}")


if __name__ == "__main__":
    main()
