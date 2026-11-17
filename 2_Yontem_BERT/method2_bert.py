"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: 2. Yöntem - BERT & BERTurk Tabanlı Duygu Analizi (Derin Öğrenme Yaklaşımı)

Bu modül:
1. İngilizce veri seti (9.000 yorum) üzerinde 3 sınıflı fine-tuned BERT modeli
   ('gohbwj/sentiment-fine-tuned-yelp-3L') ile duygu tahmini yapar.
2. Türkçe veri seti (9.000 yorum) üzerinde 3 sınıflı fine-tuned BERTurk modeli
   ('incidelen/bert-base-turkish-sentiment-analysis-cased') ile duygu tahmini yapar.
3. Sınıf bazlı hatalı tahmin sayılarını, karmaşıklık matrisini ve F1 skorlarını hesaplar.
"""

import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Union

import pandas as pd
import torch
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from transformers import logging as hf_logging
from transformers import pipeline

hf_logging.set_verbosity_error()


# ============================================================================
# SABİTLER VE KONFİGÜRASYON (CONSTANTS & CONFIGURATION)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
DATA_DIR: Path = BASE_DIR.parent / "Veri_Setleri"

ENGLISH_INPUT_CSV: Path = DATA_DIR / "english_restaurant_reviews_9000.csv"
TURKISH_INPUT_CSV: Path = DATA_DIR / "turkish_restaurant_reviews_9000.csv"

ENGLISH_OUTPUT_CSV: Path = BASE_DIR / "bert_english_results.csv"
TURKISH_OUTPUT_CSV: Path = BASE_DIR / "bert_turkish_results.csv"
SUMMARY_METRICS_JSON: Path = BASE_DIR / "bert_summary_metrics.json"

LABELS: List[str] = ["Positive", "Neutral", "Negative"]
MAX_TOKEN_LENGTH: int = 512
INFERENCE_BATCH_SIZE: int = 64


@dataclass(frozen=True)
class BertTaskConfig:
    """Dil bazlı BERT model ve dosya yolu konfigürasyonu."""

    title: str
    model_name: str
    input_csv: Path
    output_csv: Path
    label_map: Dict[str, str]


ENGLISH_BERT_CONFIG = BertTaskConfig(
    title="İNGİLİZCE VERİ SETİ - BERT TABANLI DUYGU ANALİZİ SONUÇLARI (9.000 Yorum)",
    model_name="gohbwj/sentiment-fine-tuned-yelp-3L",
    input_csv=ENGLISH_INPUT_CSV,
    output_csv=ENGLISH_OUTPUT_CSV,
    label_map={
        "LABEL_0": "Negative",
        "LABEL_1": "Positive",
        "LABEL_2": "Neutral",
        "negative": "Negative",
        "neutral": "Neutral",
        "positive": "Positive",
    },
)

TURKISH_BERT_CONFIG = BertTaskConfig(
    title="TÜRKÇE VERİ SETİ - BERTurk TABANLI DUYGU ANALİZİ SONUÇLARI (9.000 Yorum)",
    model_name="incidelen/bert-base-turkish-sentiment-analysis-cased",
    input_csv=TURKISH_INPUT_CSV,
    output_csv=TURKISH_OUTPUT_CSV,
    label_map={
        "Negative": "Negative",
        "Neutral": "Neutral",
        "Positive": "Positive",
        "LABEL_0": "Negative",
        "LABEL_1": "Neutral",
        "LABEL_2": "Positive",
    },
)


# ============================================================================
# DONANIM SEÇİMİ VE DEĞERLENDİRME FONKSİYONLARI
# ============================================================================

def detect_compute_device() -> Union[int, str]:
    """Kullanılabilir en hızlı donanım hızlandırıcısını (CUDA GPU, Apple MPS veya CPU) seçer."""
    if torch.cuda.is_available():
        return 0
    if torch.backends.mps.is_available():
        return "mps"
    return -1


def evaluate_predictions(
    df: pd.DataFrame, pred_col: str, title: str, model_name: str
) -> Dict[str, Any]:
    """
    Model tahminlerini gerçek etiketlerle karşılaştırarak sınıf bazlı hatalı tahmin
    adetlerini, karmaşıklık matrisini ve sınıflandırma raporunu üretir.
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
    print(f" Model: {model_name}")
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


def run_bert_pipeline(config: BertTaskConfig, device: Union[int, str]) -> Dict[str, Any]:
    """Belirtilen dil ve BERT konfigürasyonu için çıkarım (inference) ve değerlendirme yapar."""
    print(f"\nBERT analizi çalıştırılıyor: {config.model_name} (Cihaz: {device})...")
    df = pd.read_csv(config.input_csv)

    start_time = time.time()
    classifier = pipeline(
        "sentiment-analysis",
        model=config.model_name,
        device=device,
        truncation=True,
        max_length=MAX_TOKEN_LENGTH,
    )
    texts = df["text"].astype(str).tolist()
    outputs = classifier(texts, batch_size=INFERENCE_BATCH_SIZE)

    df["bert_raw_label"] = [item["label"] for item in outputs]
    df["bert_confidence"] = [round(float(item["score"]), 4) for item in outputs]
    df["bert_pred"] = [config.label_map.get(item["label"], "Neutral") for item in outputs]

    df.to_csv(config.output_csv, index=False, encoding="utf-8")
    elapsed = time.time() - start_time
    print(f"   Tamamlandı ({elapsed:.1f} sn) -> {config.output_csv.name} kaydedildi.")

    return evaluate_predictions(df, "bert_pred", config.title, config.model_name)


def main() -> None:
    """İngilizce ve Türkçe BERT modellerini çalıştırır ve özet metrikleri kaydeder."""
    device = detect_compute_device()
    english_metrics = run_bert_pipeline(ENGLISH_BERT_CONFIG, device)
    turkish_metrics = run_bert_pipeline(TURKISH_BERT_CONFIG, device)

    summary_payload = {
        "english_bert": english_metrics,
        "turkish_bert": turkish_metrics,
    }
    with open(SUMMARY_METRICS_JSON, "w", encoding="utf-8") as file:
        json.dump(summary_payload, file, ensure_ascii=False, indent=2)
    print(f"\nÖzet metrikler kaydedildi: {SUMMARY_METRICS_JSON.name}")


if __name__ == "__main__":
    main()
