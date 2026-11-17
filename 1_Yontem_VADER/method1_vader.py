"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: 1. Yöntem - VADER Duygu Analizi (Kural ve Sözlük Tabanlı Yaklaşım)

Bu modül:
1. İngilizce veri seti (9.000 yorum) üzerinde doğrudan VADER duygu analizi uygular.
2. Türkçe veri seti (9.000 yorum) üzerinde hem çevirisiz (ham metin) hem de
   İngilizceye çevrilmiş metinler üzerinde VADER duygu analizi uygular.
3. Her iki dil ve duygu sınıfı (Positive, Neutral, Negative) bazında yapılan
   hatalı tahmin sayılarını (yanlış sınıflandırma adetlerini) hesaplar.
"""

import concurrent.futures
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

import pandas as pd
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer


# ============================================================================
# SABİTLER VE KONFİGÜRASYON (CONSTANTS & CONFIGURATION)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
DATA_DIR: Path = BASE_DIR.parent / "Veri_Setleri"

ENGLISH_INPUT_CSV: Path = DATA_DIR / "english_restaurant_reviews_9000.csv"
TURKISH_INPUT_CSV: Path = DATA_DIR / "turkish_restaurant_reviews_9000.csv"

ENGLISH_OUTPUT_CSV: Path = BASE_DIR / "vader_english_results.csv"
TURKISH_OUTPUT_CSV: Path = BASE_DIR / "vader_turkish_results.csv"
SUMMARY_METRICS_JSON: Path = BASE_DIR / "vader_summary_metrics.json"

LABELS: List[str] = ["Positive", "Neutral", "Negative"]

POSITIVE_THRESHOLD: float = 0.05
NEGATIVE_THRESHOLD: float = -0.05

TRANSLATION_BATCH_SIZE: int = 50
TRANSLATION_MAX_WORKERS: int = 4
TRANSLATION_CLIENTS: Tuple[str, ...] = ("dict-chrome-ex", "at")


# ============================================================================
# VADER SINIFLANDIRMA VE ÇEVİRİ FONKSİYONLARI
# ============================================================================

def classify_compound_score(compound_score: float) -> str:
    """
    VADER bileşik (compound) skorunu standart eşik değerlerine göre
    3 sınıflı duygu etiketine dönüştürür.
    """
    if compound_score >= POSITIVE_THRESHOLD:
        return "Positive"
    if compound_score <= NEGATIVE_THRESHOLD:
        return "Negative"
    return "Neutral"


def compute_vader_scores(
    texts: Sequence[str], analyzer: SentimentIntensityAnalyzer
) -> Tuple[List[float], List[str]]:
    """Verilen metin listesi için VADER compound skorlarını ve sınıf tahminlerini üretir."""
    compounds = [analyzer.polarity_scores(str(text))["compound"] for text in texts]
    predictions = [classify_compound_score(score) for score in compounds]
    return compounds, predictions


def translate_batch_tr_to_en(texts: List[str], max_retries: int = 4) -> List[str]:
    """
    Türkçe metin listesini Google Translate API üzerinden toplu (batch) olarak
    İngilizceye çevirir. Hata durumunda paketi ikiye bölerek yeniden dener.
    """
    for attempt in range(max_retries):
        client_name = TRANSLATION_CLIENTS[attempt % len(TRANSLATION_CLIENTS)]
        url = f"https://clients5.google.com/translate_a/t?client={client_name}&sl=tr&tl=en"
        encoded_payload = urllib.parse.urlencode([("q", str(t)) for t in texts]).encode("utf-8")

        request = urllib.request.Request(
            url,
            data=encoded_payload,
            headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
                "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                result = json.loads(response.read().decode("utf-8"))
                if isinstance(result, list) and len(result) == len(texts):
                    return [
                        str(item[0]) if isinstance(item, list) else str(item)
                        for item in result
                    ]
        except Exception:
            time.sleep(1.0 * (attempt + 1))

    if len(texts) > 1:
        midpoint = len(texts) // 2
        return (
            translate_batch_tr_to_en(texts[:midpoint])
            + translate_batch_tr_to_en(texts[midpoint:])
        )
    return [str(texts[0])]


def translate_corpus_parallel(texts: List[str]) -> List[str]:
    """Tüm Türkçe yorum korpusunu paralel iş parçacıklarıyla İngilizceye çevirir."""
    batches = [
        (idx, texts[idx : idx + TRANSLATION_BATCH_SIZE])
        for idx in range(0, len(texts), TRANSLATION_BATCH_SIZE)
    ]
    translated_corpus: List[str] = [""] * len(texts)

    print(f"   {len(texts):,} Türkçe yorum İngilizceye çevriliyor ({len(batches)} paket)...")
    start_time = time.time()
    completed = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=TRANSLATION_MAX_WORKERS) as executor:
        future_to_start_idx = {
            executor.submit(translate_batch_tr_to_en, batch_texts): start_idx
            for start_idx, batch_texts in batches
        }
        for future in concurrent.futures.as_completed(future_to_start_idx):
            start_idx = future_to_start_idx[future]
            batch_result = future.result()
            translated_corpus[start_idx : start_idx + len(batch_result)] = batch_result
            completed += 1
            if completed % 30 == 0 or completed == len(batches):
                elapsed = time.time() - start_time
                print(f"   Çeviri ilerlemesi: {completed}/{len(batches)} paket ({elapsed:.1f} sn)")

    return translated_corpus


# ============================================================================
# DEĞERLENDİRME VE HATA ANALİZİ FONKSİYONLARI
# ============================================================================

def evaluate_predictions(df: pd.DataFrame, pred_col: str, title: str) -> Dict[str, Any]:
    """
    Gerçek etiketler ile model tahminlerini karşılaştırarak sınıf bazlı hatalı
    tahmin sayılarını, karmaşıklık matrisini ve F1 metriklerini raporlar.
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
        "accuracy": round(accuracy, 4),
        "total_samples": len(df),
        "total_correct": total_correct,
        "total_wrong": total_wrong,
        "class_errors": class_errors,
        "classification_report": report,
        "confusion_matrix": cm.tolist(),
    }


# ============================================================================
# İŞ AKIŞI YÜRÜTÜCÜLERİ (WORKFLOW RUNNERS)
# ============================================================================

def run_english_vader() -> Dict[str, Any]:
    """İngilizce veri seti üzerinde VADER analizini çalıştırır ve sonuçları kaydeder."""
    print("1) İngilizce veri seti (9.000 yorum) üzerinde VADER çalıştırılıyor...")
    df_en = pd.read_csv(ENGLISH_INPUT_CSV)
    analyzer = SentimentIntensityAnalyzer()

    start_time = time.time()
    compounds, predictions = compute_vader_scores(df_en["text"].tolist(), analyzer)

    df_en["vader_compound"] = compounds
    df_en["vader_pred"] = predictions
    df_en.to_csv(ENGLISH_OUTPUT_CSV, index=False, encoding="utf-8")

    elapsed = time.time() - start_time
    print(f"   Tamamlandı ({elapsed:.2f} sn) -> {ENGLISH_OUTPUT_CSV.name} kaydedildi.")
    return evaluate_predictions(
        df_en, "vader_pred", "İNGİLİZCE VERİ SETİ - VADER SONUÇLARI (9.000 Yorum)"
    )


def load_or_translate_turkish_dataset() -> pd.DataFrame:
    """
    Önceden çevrilmiş Türkçe sonuç dosyası varsa yükler; yoksa ham Türkçe veri setini
    okuyarak İngilizceye çevirir.
    """
    if TURKISH_OUTPUT_CSV.exists():
        df_cached = pd.read_csv(TURKISH_OUTPUT_CSV)
        if "translated_text" in df_cached.columns and df_cached["translated_text"].notnull().all():
            print(f"   Önbellekteki çevrilmiş dosya kullanılıyor: {TURKISH_OUTPUT_CSV.name}")
            return df_cached

    df_tr = pd.read_csv(TURKISH_INPUT_CSV)
    df_tr["translated_text"] = translate_corpus_parallel(df_tr["text"].astype(str).tolist())
    return df_tr


def run_turkish_vader() -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    Türkçe veri seti üzerinde hem çevirisiz (ham) hem de İngilizceye çevrilmiş
    metinlerle VADER analizini çalıştırır.
    """
    print("\n2) Türkçe veri seti (9.000 yorum) üzerinde VADER çalıştırılıyor...")
    df_tr = load_or_translate_turkish_dataset()
    analyzer = SentimentIntensityAnalyzer()

    raw_compounds, raw_preds = compute_vader_scores(df_tr["text"].tolist(), analyzer)
    df_tr["vader_raw_compound"] = raw_compounds
    df_tr["vader_raw_pred"] = raw_preds

    trans_compounds, trans_preds = compute_vader_scores(df_tr["translated_text"].tolist(), analyzer)
    df_tr["vader_compound"] = trans_compounds
    df_tr["vader_pred"] = trans_preds

    df_tr.to_csv(TURKISH_OUTPUT_CSV, index=False, encoding="utf-8")
    print(f"   -> {TURKISH_OUTPUT_CSV.name} kaydedildi.")

    metrics_raw = evaluate_predictions(
        df_tr, "vader_raw_pred", "TÜRKÇE VERİ SETİ - ÇEVİRİSİZ (HAM) VADER SONUÇLARI (9.000 Yorum)"
    )
    metrics_translated = evaluate_predictions(
        df_tr, "vader_pred", "TÜRKÇE VERİ SETİ - İNGİLİZCEYE ÇEVRİLMİŞ VADER SONUÇLARI (9.000 Yorum)"
    )
    return metrics_raw, metrics_translated


def main() -> None:
    """VADER duygu analizi iş akışını çalıştırır ve özet metrikleri JSON olarak kaydeder."""
    english_metrics = run_english_vader()
    turkish_raw_metrics, turkish_translated_metrics = run_turkish_vader()

    summary_payload = {
        "english_vader": english_metrics,
        "turkish_vader_raw": turkish_raw_metrics,
        "turkish_vader_translated": turkish_translated_metrics,
    }
    with open(SUMMARY_METRICS_JSON, "w", encoding="utf-8") as file:
        json.dump(summary_payload, file, ensure_ascii=False, indent=2)
    print(f"\nÖzet metrikler kaydedildi: {SUMMARY_METRICS_JSON.name}")


if __name__ == "__main__":
    main()
