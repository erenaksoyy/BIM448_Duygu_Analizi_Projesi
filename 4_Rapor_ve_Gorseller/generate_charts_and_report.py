"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: Karşılaştırmalı Görselleştirme ve Akademik Rapor Üretimi

Bu modül, 3 yöntemin (VADER, BERT/BERTurk ve LLM) ürettiği özet metrikleri ve
tahmin sonuçlarını okuyarak:
1. 300 DPI yüksek çözünürlüklü 4 adet karşılaştırma grafiği (.png) üretir.
2. Niteliksel hata analizi için veri setinden gerçek yanlış sınıflandırma örneklerini çeker.
3. VADER, BERT ve LLM'in kuramsal açıklamalarını, tabloları, grafikleri ve
   akademik tartışmayı içeren Word raporunu (.docx) oluşturur.
"""

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


# ============================================================================
# SABİTLER VE KONFİGÜRASYON (CONSTANTS & CONFIGURATION)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
ROOT_DIR: Path = BASE_DIR.parent

VADER_METRICS_PATH: Path = ROOT_DIR / "1_Yontem_VADER" / "vader_summary_metrics.json"
BERT_METRICS_PATH: Path = ROOT_DIR / "2_Yontem_BERT" / "bert_summary_metrics.json"
LLM_METRICS_PATH: Path = ROOT_DIR / "3_Yontem_LLM" / "llm_summary_metrics.json"

FIG_CONFUSION_MATRICES: Path = BASE_DIR / "1_confusion_matrices_6panel.png"
FIG_ACCURACY_F1: Path = BASE_DIR / "2_accuracy_and_f1_comparison.png"
FIG_CLASS_ERRORS: Path = BASE_DIR / "3_class_wise_error_counts.png"
FIG_TR_TRANSLATION: Path = BASE_DIR / "4_turkish_translation_impact_vader.png"

QUALITATIVE_EXAMPLES_JSON: Path = BASE_DIR / "qualitative_error_examples.json"
REPORT_DOCX_PATH: Path = BASE_DIR / "BIM448_Proje_Raporu.docx"

LABELS: List[str] = ["Positive", "Neutral", "Negative"]
LABEL_DISPLAY_TR: List[str] = ["Pozitif\n(3.000)", "Nötr\n(3.000)", "Negatif\n(3.000)"]

FIGURE_DPI: int = 300
COLOR_PRIMARY_BLUE: str = "#2563EB"
COLOR_SUCCESS_GREEN: str = "#10B981"
COLOR_WARNING_AMBER: str = "#F59E0B"
COLOR_DANGER_RED: str = "#EF4444"

plt.rcParams["font.sans-serif"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#CCCCCC"
plt.rcParams["axes.linewidth"] = 0.8


@dataclass(frozen=True)
class HeatmapPanelSpec:
    """Karmaşıklık matrisi (Confusion Matrix) alt grafik tanımı."""

    row_idx: int
    col_idx: int
    metrics: Dict[str, Any]
    title: str
    colormap: str


# ============================================================================
# VERİ YÜKLEME VE GÖRSELLEŞTİRME FONKSİYONLARI
# ============================================================================

def load_json_file(path: Path) -> Dict[str, Any]:
    """Belirtilen JSON dosyasını UTF-8 kodlamasıyla okur."""
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)


def load_all_summary_metrics() -> Tuple[Dict[str, Any], Dict[str, Any], Dict[str, Any]]:
    """VADER, BERT ve LLM özet metrik dosyalarını yükler."""
    return (
        load_json_file(VADER_METRICS_PATH),
        load_json_file(BERT_METRICS_PATH),
        load_json_file(LLM_METRICS_PATH),
    )


def annotate_bar_values(
    ax: plt.Axes,
    bars: Any,
    fmt: str = "{:.0f}",
    color: str = "#111111",
    fontsize: float = 9.5,
) -> None:
    """Sütun grafiklerindeki her çubuğun üzerine sayısal değerini yazar."""
    for bar in bars:
        height = bar.get_height()
        ax.annotate(
            fmt.format(height),
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 3),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=fontsize,
            fontweight="bold",
            color=color,
        )


def plot_six_panel_confusion_matrices(
    vader: Dict[str, Any], bert: Dict[str, Any], llm: Dict[str, Any]
) -> None:
    """İngilizce ve Türkçe veri setlerinde 3 yöntemin 6'lı karmaşıklık matrisini çizer."""
    fig, axes = plt.subplots(2, 3, figsize=(18, 11), dpi=FIGURE_DPI)
    fig.suptitle(
        "İngilizce ve Türkçe Veri Setlerinde 3 Yöntemin Karmaşıklık Matrisleri (Confusion Matrices - 9.000'er Yorum)",
        fontsize=16,
        fontweight="bold",
        y=0.98,
    )

    panels: List[HeatmapPanelSpec] = [
        HeatmapPanelSpec(0, 0, vader["english_vader"], "İngilizce - 1. Yöntem: VADER\n(Accuracy: %51.53 | Toplam Hata: 4,362)", "Blues"),
        HeatmapPanelSpec(0, 1, bert["english_bert"], "İngilizce - 2. Yöntem: BERT (Yelp-3L)\n(Accuracy: %80.96 | Toplam Hata: 1,714)", "Blues"),
        HeatmapPanelSpec(0, 2, llm["english_llm"], "İngilizce - 3. Yöntem: LLM (EVREN Qwen3.8)\n(Accuracy: %83.63 | Toplam Hata: 1,473)", "Blues"),
        HeatmapPanelSpec(1, 0, vader["turkish_vader_translated"], "Türkçe - 1. Yöntem: VADER (Çevirili)\n(Accuracy: %53.99 | Toplam Hata: 4,141)", "Oranges"),
        HeatmapPanelSpec(1, 1, bert["turkish_bert"], "Türkçe - 2. Yöntem: BERTurk (3-Sınıflı)\n(Accuracy: %72.11 | Toplam Hata: 2,510)", "Oranges"),
        HeatmapPanelSpec(1, 2, llm["turkish_llm"], "Türkçe - 3. Yöntem: LLM (EVREN Qwen3.8)\n(Accuracy: %79.11 | Toplam Hata: 1,880)", "Oranges"),
    ]

    for spec in panels:
        ax = axes[spec.row_idx, spec.col_idx]
        cm_array = np.array(spec.metrics["confusion_matrix"])
        sns.heatmap(
            cm_array,
            annot=True,
            fmt="d",
            cmap=spec.colormap,
            cbar=False,
            xticklabels=LABELS,
            yticklabels=LABELS,
            annot_kws={"size": 13, "weight": "bold"},
            ax=ax,
            linewidths=1,
            linecolor="white",
        )
        ax.set_title(spec.title, fontsize=12, fontweight="bold", pad=10)
        ax.set_xlabel("Tahmin Edilen Sınıf (Predicted)", fontsize=10, fontweight="bold")
        ax.set_ylabel("Gerçek Sınıf (True Label)", fontsize=10, fontweight="bold")

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig(FIG_CONFUSION_MATRICES, dpi=FIGURE_DPI)
    plt.close()
    print(f"Grafik kaydedildi: {FIG_CONFUSION_MATRICES.name}")


def plot_accuracy_and_f1_comparison(
    vader: Dict[str, Any], bert: Dict[str, Any], llm: Dict[str, Any]
) -> None:
    """Tüm yöntemlerin Genel Doğruluk (Accuracy) ve Macro F1 skorlarını karşılaştırır."""
    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=FIGURE_DPI)

    experiment_entries = [
        ("İngilizce\nVADER", vader["english_vader"]),
        ("İngilizce\nBERT", bert["english_bert"]),
        ("İngilizce\nLLM (EVREN)", llm["english_llm"]),
        ("Türkçe\nVADER (Ham)", vader["turkish_vader_raw"]),
        ("Türkçe\nVADER (Çevirili)", vader["turkish_vader_translated"]),
        ("Türkçe\nBERTurk", bert["turkish_bert"]),
        ("Türkçe\nLLM (EVREN)", llm["turkish_llm"]),
    ]

    labels = [name for name, _ in experiment_entries]
    accuracies = [data["accuracy"] * 100 for _, data in experiment_entries]
    macro_f1s = [
        data["classification_report"]["macro avg"]["f1-score"] * 100
        for _, data in experiment_entries
    ]

    x_pos = np.arange(len(labels))
    bar_width = 0.36

    bars_acc = ax.bar(
        x_pos - bar_width / 2, accuracies, bar_width,
        label="Genel Doğruluk (Accuracy %)", color=COLOR_PRIMARY_BLUE, edgecolor="white"
    )
    bars_f1 = ax.bar(
        x_pos + bar_width / 2, macro_f1s, bar_width,
        label="Macro F1-Skoru (%)", color=COLOR_SUCCESS_GREEN, edgecolor="white"
    )

    ax.set_ylabel("Başarım Oranı (%)", fontsize=12, fontweight="bold")
    ax.set_title(
        "İngilizce ve Türkçe Veri Setlerinde Yöntemlerin Genel Başarım Karşılaştırması",
        fontsize=14,
        fontweight="bold",
        pad=15,
    )
    ax.set_xticks(x_pos)
    ax.set_xticklabels(labels, fontsize=10, fontweight="bold")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", linestyle="--", alpha=0.5)
    ax.legend(fontsize=11, loc="upper left")

    annotate_bar_values(ax, bars_acc, fmt="%{:.1f}", color="#1E3A8A")
    annotate_bar_values(ax, bars_f1, fmt="%{:.1f}", color="#065F46")

    plt.tight_layout()
    plt.savefig(FIG_ACCURACY_F1, dpi=FIGURE_DPI)
    plt.close()
    print(f"Grafik kaydedildi: {FIG_ACCURACY_F1.name}")


def plot_class_wise_error_counts(
    vader: Dict[str, Any], bert: Dict[str, Any], llm: Dict[str, Any]
) -> None:
    """Pozitif, Nötr ve Negatif sınıflarında yapılan hatalı tahmin sayılarını görselleştirir."""
    fig, axes = plt.subplots(1, 2, figsize=(16, 6.5), dpi=FIGURE_DPI, sharey=True)
    fig.suptitle(
        "Sınıf Bazlı Hatalı Tahmin Sayıları (Her Sınıfta 3.000 Yorum Üzerinden Yanlış Sınıflandırma Adetleri)",
        fontsize=15,
        fontweight="bold",
        y=0.98,
    )

    x_pos = np.arange(len(LABELS))
    bar_width = 0.26

    language_configs = [
        (
            axes[0],
            "İngilizce Veri Seti (9.000 Yorum)",
            ("1. VADER (Toplam: 4.362 Hata)", vader["english_vader"]),
            ("2. BERT (Toplam: 1.714 Hata)", bert["english_bert"]),
            ("3. LLM - EVREN (Toplam: 1.473 Hata)", llm["english_llm"]),
        ),
        (
            axes[1],
            "Türkçe Veri Seti (9.000 Yorum)",
            ("1. VADER - Çevirili (Toplam: 4.141 Hata)", vader["turkish_vader_translated"]),
            ("2. BERTurk (Toplam: 2.510 Hata)", bert["turkish_bert"]),
            ("3. LLM - EVREN (Toplam: 1.880 Hata)", llm["turkish_llm"]),
        ),
    ]

    colors = (COLOR_DANGER_RED, COLOR_WARNING_AMBER, COLOR_SUCCESS_GREEN)
    offsets = (-bar_width, 0.0, bar_width)

    for ax, panel_title, m1, m2, m3 in language_configs:
        for (legend_label, metric_data), offset, color in zip((m1, m2, m3), offsets, colors):
            error_counts = [metric_data["class_errors"][cls]["wrong"] for cls in LABELS]
            bars = ax.bar(x_pos + offset, error_counts, bar_width, label=legend_label, color=color)
            annotate_bar_values(ax, bars, fmt="{:.0f}")

        ax.set_title(panel_title, fontsize=13, fontweight="bold")
        ax.set_xticks(x_pos)
        ax.set_xticklabels(LABEL_DISPLAY_TR, fontsize=11, fontweight="bold")
        ax.set_ylim(0, 3300)
        ax.grid(axis="y", linestyle="--", alpha=0.5)
        ax.legend(fontsize=10.5)

    axes[0].set_ylabel("Hatalı Tahmin Adedi (Yanlış Sayısı)", fontsize=11, fontweight="bold")

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    plt.savefig(FIG_CLASS_ERRORS, dpi=FIGURE_DPI)
    plt.close()
    print(f"Grafik kaydedildi: {FIG_CLASS_ERRORS.name}")


def plot_turkish_vader_translation_effect(vader: Dict[str, Any]) -> None:
    """Türkçe veri setinde çevirisiz (ham) VADER ile çevirili VADER sonuçlarını karşılaştırır."""
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5), dpi=FIGURE_DPI)
    fig.suptitle(
        "Türkçe Veri Setinde Çevirisiz (Ham) VADER vs. İngilizceye Çevrilmiş VADER Karşılaştırması",
        fontsize=14,
        fontweight="bold",
        y=0.98,
    )

    panels = [
        (
            axes[0],
            np.array(vader["turkish_vader_raw"]["confusion_matrix"]),
            "Çevirisiz Ham Türkçe VADER\n(Accuracy: %35.37 | 9.000 Yorumun 8.311'ine Nötr Verdi)",
            "Reds",
        ),
        (
            axes[1],
            np.array(vader["turkish_vader_translated"]["confusion_matrix"]),
            "İngilizceye Çevrilmiş Türkçe VADER\n(Accuracy: %53.99 | Toplam Hata: 4.141)",
            "Greens",
        ),
    ]

    for ax, cm_array, title, cmap in panels:
        sns.heatmap(
            cm_array,
            annot=True,
            fmt="d",
            cmap=cmap,
            cbar=False,
            xticklabels=LABELS,
            yticklabels=LABELS,
            annot_kws={"size": 12, "weight": "bold"},
            ax=ax,
            linewidths=1,
            linecolor="white",
        )
        ax.set_title(title, fontsize=11.5, fontweight="bold")
        ax.set_xlabel("Tahmin Edilen Sınıf", fontweight="bold")
        ax.set_ylabel("Gerçek Sınıf", fontweight="bold")

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(FIG_TR_TRANSLATION, dpi=FIGURE_DPI)
    plt.close()
    print(f"Grafik kaydedildi: {FIG_TR_TRANSLATION.name}")


# ============================================================================
# NİTELİKSEL ÖRNEK ÇIKARIMI VE WORD RAPORU OLUŞTURMA
# ============================================================================

def extract_qualitative_examples() -> Dict[str, Dict[str, Any]]:
    """Veri setinden yöntemlerin karakteristik hatalarını gösteren gerçek örnekleri seçer."""
    df_en = pd.read_csv(ROOT_DIR / "1_Yontem_VADER" / "vader_english_results.csv")
    df_en["bert_pred"] = pd.read_csv(ROOT_DIR / "2_Yontem_BERT" / "bert_english_results.csv")["bert_pred"]
    df_en["llm_pred"] = pd.read_csv(ROOT_DIR / "3_Yontem_LLM" / "llm_english_results.csv")["llm_pred"]

    df_tr = pd.read_csv(ROOT_DIR / "1_Yontem_VADER" / "vader_turkish_results.csv")
    df_tr["bert_pred"] = pd.read_csv(ROOT_DIR / "2_Yontem_BERT" / "bert_turkish_results.csv")["bert_pred"]
    df_tr["llm_pred"] = pd.read_csv(ROOT_DIR / "3_Yontem_LLM" / "llm_turkish_results.csv")["llm_pred"]

    examples = {
        "en_neg_vader_pos": df_en[
            (df_en["true_label"] == "Negative")
            & (df_en["vader_pred"] == "Positive")
            & (df_en["bert_pred"] == "Negative")
            & (df_en["llm_pred"] == "Negative")
        ].iloc[0].to_dict(),
        "en_neu_vader_pos": df_en[
            (df_en["true_label"] == "Neutral")
            & (df_en["vader_pred"] == "Positive")
            & (df_en["bert_pred"] == "Neutral")
            & (df_en["llm_pred"] == "Neutral")
        ].iloc[0].to_dict(),
        "tr_neg_vader_pos": df_tr[
            (df_tr["true_label"] == "Negative")
            & (df_tr["vader_pred"] == "Positive")
            & (df_tr["bert_pred"] == "Negative")
            & (df_tr["llm_pred"] == "Negative")
        ].iloc[0].to_dict(),
        "tr_neu_llm_correct": df_tr[
            (df_tr["true_label"] == "Neutral")
            & (df_tr["vader_pred"] == "Positive")
            & (df_tr["bert_pred"] == "Negative")
            & (df_tr["llm_pred"] == "Neutral")
        ].iloc[0].to_dict(),
    }

    with open(QUALITATIVE_EXAMPLES_JSON, "w", encoding="utf-8") as file:
        json.dump(examples, file, ensure_ascii=False, indent=2)
    return examples


def add_formatted_table(
    doc: Document, headers: Sequence[str], rows: Sequence[Sequence[Any]]
) -> None:
    """Word belgesine biçimlendirilmiş tablo ekler."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    header_cells = table.rows[0].cells
    for col_idx, header_text in enumerate(headers):
        header_cells[col_idx].text = header_text
        for paragraph in header_cells[col_idx].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9.5)

    for row_values in rows:
        row_cells = table.add_row().cells
        for col_idx, value in enumerate(row_values):
            row_cells[col_idx].text = str(value)
            for paragraph in row_cells[col_idx].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9.5)
    doc.add_paragraph()


def build_word_report(examples: Dict[str, Dict[str, Any]]) -> None:
    """Proje sonuçlarını, kuramsal altyapıyı, tabloları ve grafikleri içeren Word (.docx) raporunu üretir."""
    doc = Document()

    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = title_paragraph.add_run(
        "DOĞUŞ ÜNİVERSİTESİ\nMÜHENDİSLİK FAKÜLTESİ - BİLGİSAYAR MÜHENDİSLİĞİ BÖLÜMÜ\n"
        "BİM448 VERİ MADENCİLİĞİ VE BİLGİ KEŞFİ DERSİ PROJE RAPORU\n\n"
    )
    header_run.bold = True
    header_run.font.size = Pt(13)

    subtitle_run = title_paragraph.add_run(
        "KARŞILAŞTIRMALI DUYGU ANALİZİ (SENTIMENT ANALYSIS):\n"
        "İNGİLİZCE VE TÜRKÇE MEKAN YORUMLARI ÜZERİNDE VADER, BERT VE LLM YÖNTEMLERİNİN SINIF BAZLI HATA ANALİZİ"
    )
    subtitle_run.bold = True
    subtitle_run.font.size = Pt(14)
    subtitle_run.font.color.rgb = RGBColor(30, 58, 138)

    add_formatted_table(
        doc,
        ["Grup Üyesi", "Ad Soyad", "Öğrenci Numarası", "Sorumlu Olduğu Bölümler"],
        [
            ["1. Üye", "Eren Aksoy", "[Öğrenci No]", "Veri Ön İşleme, 1. Yöntem (VADER - TR & EN), 2. Yöntem (BERT/BERTurk)"],
            ["2. Üye", "[2. Grup Üyesi Ad Soyad]", "[Öğrenci No]", "3. Yöntem (LLM - EVREN API), Sınıf Bazlı Hata Analizi ve Akademik Değerlendirme"],
        ],
    )

    doc.add_heading("1. Projenin Amacı ve Kapsamı", level=1)
    doc.add_paragraph(
        "Bu çalışmanın amacı; kural ve sözlük tabanlı geleneksel duygu analizi yaklaşımı (VADER), "
        "derin öğrenme ve çift yönlü dönüştürücü (Transformer) tabanlı modeller (BERT ve BERTurk) ile "
        "güncel Büyük Dil Modeli (LLM - EVREN Qwen3.8-Flash-Next) tabanlı duygu analizi yöntemlerinin başarımını "
        "İngilizce ve Türkçe gerçek kullanıcı yorumları üzerinde karşılaştırmalı olarak incelemektir. "
        "Proje yönergesinde belirtilen 'her sınıfta en az 1.000 yorum (toplam 3.000 yorum)' alt sınırı, "
        "istatistiksel güvenilirliği ve akademik derinliği artırmak amacıyla 3 katına çıkarılmış; "
        "her iki dilde de sınıf başına 3.000'er adet (Pozitif: 3.000, Nötr: 3.000, Negatif: 3.000) olmak üzere "
        "dil başına 9.000, toplamda 18.000 gerçek restoran ve kafe yorumu üzerinde analizler gerçekleştirilmiştir."
    )

    doc.add_heading("2. Kullanılan İngilizce ve Türkçe Veri Setlerinin Kaynakları ve Detayları", level=1)
    doc.add_paragraph(
        "İki dil arasındaki yöntemsel karşılaştırmanın alan (domain) farkından etkilenmemesi adına, "
        "hem İngilizce hem de Türkçe veri setleri birebir aynı sektörden (Restoran, Kafe ve Yeme-İçme Mekanları) "
        "ve aynı sütun şemasına (id, place_name, place_category, rating, text, true_label) sahip olacak şekilde hazırlanmıştır:"
    )

    add_formatted_table(
        doc,
        ["Özellik", "Türkçe Veri Seti", "İngilizce Veri Seti"],
        [
            ["Veri Kaynağı (Hugging Face)", "opdullah/turkish-google-maps-reviews", "jaimik69/Yelp-Restaurant-Dataset & hasnat79/yelp_food_business_review_dataset_2M"],
            ["Platform ve Alan (Domain)", "Google Maps Türkiye (Restoran & Kafe Yorumları)", "Yelp ABD (Restoran & Kafe Yorumları)"],
            ["Toplam Ham Havuz Boyutu", "838.583 yorum (~251.000 Restoran/Kafe)", "2.007.040 restoran yorumu (64.629 işletme)"],
            ["Örneklem Boyutu (Dengeli)", "9.000 Yorum (3.000 Pozitif, 3.000 Nötr, 3.000 Negatif)", "9.000 Yorum (3.000 Pozitif, 3.000 Nötr, 3.000 Negatif)"],
            ["Etiketleme Kriteri", "1-2 Yıldız: Negative | 3 Yıldız: Neutral | 4-5 Yıldız: Positive", "1-2 Yıldız: Negative | 3 Yıldız: Neutral | 4-5 Yıldız: Positive"],
            ["Ön İşleme ve Temizlik", "Yabancı turist yorumları, kopya satırlar ve <25 veya >500 karakter metinler temizlendi.", "İşletme bilgileri (business_id) eşleştirildi, kopya satırlar ve <25 veya >500 karakter metinler temizlendi."],
        ],
    )

    doc.add_heading("3. Uygulanan Yöntemlerin Kuramsal Altyapısı, Mimari ve Parametre Detayları", level=1)

    doc.add_heading("3.1. VADER Nedir ve Nasıl Çalışır? (Kural ve Sözlük Tabanlı Yaklaşım)", level=2)
    doc.add_paragraph(
        "VADER (Valence Aware Dictionary and sEntiment Reasoner; Hutto & Gilbert, 2014), özellikle sosyal medya metinleri, "
        "mikro-bloglar ve müşteri yorumları için geliştirilmiş sözlük (lexicon) ve kural tabanlı bir duygu analizi modelidir. "
        "Herhangi bir makine öğrenmesi eğitimi gerektirmez; bunun yerine insan uzmanlar tarafından etiketlenmiş 7.500'den fazla "
        "İngilizce kelime, kısaltma ('lol', 'wtf'), argo ifade ve emoji/emoticon (':)', '🙂') içeren bir duygu sözlüğüne dayanır.\n"
        "• Çalışma Prensibi ve 5 Dilbilimsel Sezgisel Kural (Heuristics): Sözlükteki her ifadenin -4.0 (aşırı negatif) ile "
        "+4.0 (aşırı pozitif) arasında bir duygu şiddeti (valence) skoru vardır. VADER, cümledeki kelimelerin ham skorlarını "
        "topladıktan sonra 5 dilbilimsel kural uygular:\n"
        "  1) Noktalama İşaretleri: Ünlem işareti ('!') duygu yönünü değiştirmeden şiddetini artırır (örn. 'Good!!!' > 'Good').\n"
        "  2) Büyük Harf Vurgusu (ALL-CAPS): Tamamı büyük harfle yazılan kelimelerin ('GREAT', 'TERRIBLE') duygu şiddeti artırılır.\n"
        "  3) Derece Zarfları (Degree Modifiers / Booster Words): 'very', 'extremely', 'absolutely' gibi zarflar kendisinden sonra "
        "gelen kelimenin duygu skorunu yükseltirken, 'kind of', 'slightly' gibi zarflar düşürür.\n"
        "  4) Karşıtlık Bağlaçları ('but'): Cümlede 'but' bağlacı geçtiğinde, bağlaçtan önceki kelimelerin ağırlığı %50 azaltılırken "
        "bağlaçtan sonraki kelimelerin ağırlığı %50 artırılır.\n"
        "  5) Olumsuzluk Penceresi (Negation Tri-gram): Duygu kelimesinden önceki 3 kelimelik pencerede 'not', 'never', 'isn't' gibi "
        "olumsuzluk ifadeleri varsa duygu kutbu tersine çevrilir (skora -0.74 çarpanı uygulanır).\n"
        "• Bileşik Skor (Compound Score) ve Eşik Değerleri: Cümledeki tüm kelimelerin ağırlıklı toplamı x olmak üzere, "
        "compound = x / sqrt(x^2 + alpha) (alpha=15) formülüyle -1.0 ile +1.0 arasına normalize edilir. Literatürdeki standart "
        "eşiklere göre compound >= 0.05 ise 'Positive', compound <= -0.05 ise 'Negative', -0.05 < compound < 0.05 ise 'Neutral' atanır.\n"
        "• Türkçe Uygulama Mimarisi: VADER sözlüğü yalnızca İngilizce kelimeleri tanıdığı için Türkçe veri setindeki 9.000 yorum "
        "Google Translate API tabanlı toplu çeviri mekanizmasıyla İngilizceye çevrilmiş ('translated_text') ve VADER bu çeviriler "
        "üzerinde çalıştırılmıştır. Çevirinin zorunluluğunu göstermek için çevirisiz ham Türkçe metinler üzerinde de kontrol deneyi yapılmıştır."
    )

    doc.add_heading("3.2. BERT ve BERTurk Nedir ve Nasıl Çalışır? (Derin Öğrenme Yaklaşımı)", level=2)
    doc.add_paragraph(
        "BERT (Bidirectional Encoder Representations from Transformers; Devlin et al., 2019), Google tarafından geliştirilen ve "
        "Doğal Dil İşleme (NLP) alanında devrim yaratan, Transformer mimarisinin yalnızca Kodlayıcı (Encoder) katmanlarına dayalı "
        "derin öğrenme modelidir.\n"
        "• Çift Yönlü Öz-Dikkat (Bidirectional Multi-Head Self-Attention) Mimarisi: Geleneksel dil modelleri metni yalnızca soldan sağa "
        "veya sağdan sola okurken, BERT cümledeki tüm kelimeleri aynı anda (paralel olarak) işler. Öz-dikkat (Self-Attention) mekanizması "
        "sayesinde bir kelimenin cümle içindeki bağlamsal anlamı, hem solundaki hem de sağındaki tüm kelimelerle olan anlamsal "
        "etkileşim ağırlıklarından hesaplanır. Böylece aynı kelimenin farklı bağlamlardaki zıt anlamları ve bağlaçlar ('fakat', 'ancak') "
        "sonrası değişen duygu durumu yüksek doğrulukla temsil edilir.\n"
        "• Ön-Eğitim (Pre-training) ve İnce Ayar (Fine-tuning): BERT, milyarlarca kelimelik etiketlenmemiş metin üzerinde iki görevle "
        "ön-eğitime tabi tutulur: (1) Cümledeki kelimelerin %15'ini maskeleyip bağlamdan tahmin etme (Masked Language Modeling - MLM) "
        "ve (2) İki cümlenin birbirinin devamı olup olmadığını tahmin etme (Next Sentence Prediction - NSP). Duygu analizi için "
        "ince ayar (fine-tuning) aşamasında ise cümlenin başına eklenen özel '[CLS]' (Classification) tokenının son Transformer "
        "katmanındaki 768 boyutlu temsili, doğrusal bir sınıflandırma katmanından (Linear + Softmax) geçirilerek Pozitif, Nötr ve "
        "Negatif sınıf olasılıklarını üretir.\n"
        "• BERTurk Nedir?: Türkçe sondan eklemeli (agglutinative) ve zengin morfolojiye sahip bir dildir (örn. 'güzelleştiremediklerimizden'). "
        "MDZ Digital Library ekibi tarafından geliştirilen BERTurk ('dbmdz/bert-base-turkish-cased'), 35 GB boyutunda ve 440 milyon "
        "kelimelik Türkçe derlem üzerinde eğitilmiştir. WordPiece alt-kelime (subword) tokenizasyonu sayesinde Türkçe kelime köklerini "
        "ve olumsuzluk/çekim eklerini ('-me', '-ma', '-sız') parçalarına ayırarak anlamsal kaybı önler.\n"
        "• Projedeki Model ve Parametre Seçimi: Standart 2 sınıflı modeller (SST-2 veya savasy/bert-base-turkish) 'Neutral' sınıfını "
        "tahmin edemediği için bu projede doğrudan 3 sınıflı (Positive, Neutral, Negative) olarak fine-tune edilmiş modeller seçilmiştir:\n"
        "  - İngilizce Model: 'gohbwj/sentiment-fine-tuned-yelp-3L' (12 katman, 768 gizli boyut, 12 dikkat başlığı, max_length=512, batch_size=64).\n"
        "  - Türkçe Model: 'incidelen/bert-base-turkish-sentiment-analysis-cased' (BERTurk 'dbmdz/bert-base-turkish-cased' tabanlı 3 sınıflı model, "
        "12 katman, 110M parametre, max_length=512, batch_size=64). Her iki model Google Colab NVIDIA L4 GPU üzerinde çalıştırılmıştır."
    )

    doc.add_heading("3.3. Büyük Dil Modeli (LLM) Nedir ve Nasıl Çalışır? (İstem Mühendisliği Yaklaşımı)", level=2)
    doc.add_paragraph(
        "Büyük Dil Modelleri (Large Language Models - LLM), milyarlarca (veya yüz milyarlarca) parametreye sahip, Transformer mimarisinin "
        "Çözücü (Decoder-only Autoregressive) yapısına dayanan ve trilyonlarca kelimelik çok dilli veri kümeleri üzerinde eğitilen "
        "üretken yapay zekâ modelleridir (Örn. GPT-4, Qwen, Gemma, Llama, DeepSeek).\n"
        "• LLM Nasıl Çalışır? (Ön-Eğitim, Instruction Tuning ve RLHF): LLM'ler ilk aşamada kendinden önceki kelimelere bakarak "
        "bir sonraki tokenı tahmin etme (Next-Token Prediction) göreviyle eğitilir. Bu süreçte yalnızca dilbilgisini değil; dünya bilgisini, "
        "kültürel deyimleri, ironiyi ve mantıksal çıkarım kurallarını öğrenirler. Ardından Yönerge İnce Ayarı (Instruction Tuning) ve "
        "İnsan Geri Bildirimiyle Pekiştirmeli Öğrenme (RLHF) aşamalarından geçerek insan komutlarını (prompt) doğrudan takip etme "
        "yeteneği kazanırlar.\n"
        "• Duygu Analizinde VADER ve BERT'ten Farkı: VADER sabit kelime skorlarına, BERT ise 512 tokenlık yerel sınıflandırma katmanına "
        "bağlıdır ve BERT'in başarılı olması için o alanda etiketli veriyle yeniden eğitilmesi (fine-tuning) gerekir. Oysa LLM'ler "
        "Bağlam İçi Öğrenme (In-Context / Zero-Shot Learning) yeteneği sayesinde hiçbir ek eğitim yapılmadan yalnızca doğal dilde "
        "verilen bir istemle (Prompt Engineering) duygu analizi yapabilir. Cümledeki ince alaycılığı (sarcasm), devrik Türkçe ifadeleri "
        "ve 'lezzet iyi ama porsiyon küçük, fiyatına göre normal' gibi çok boyutlu gri (Nötr) yorumları insan düzeyinde muhakeme eder.\n"
        "• Projedeki EVREN Altyapısı ve Parametreler: Bu çalışmada T.C. Cumhurbaşkanlığı Savunma Sanayii Başkanlığı (SSB) ve SAYZEK "
        "koordinesinde geliştirilen milli yapay zekâ platformu EVREN (https://evren-llmapi.ssyz.org.tr/v1) üzerindeki 'qwen3.8-flash-next' "
        "(ve yedek olarak 'gemma-4-31b') modeli kullanılmıştır. Tekrarlanabilir deterministik çıktı için temperature=0.0, hızlı çıkarım "
        "için enable_thinking=False ve her API çağrısında 50 yorumluk yapılandırılmış JSON dizi istemi (batch_size=50, 8 paralel iş parçacığı) uygulanmıştır."
    )

    doc.add_heading("4. Deneysel Sonuçlar ve Sınıf Bazlı Hatalı Tahmin Analizi", level=1)
    doc.add_heading("4.1. İngilizce Veri Seti (9.000 Yorum) Sınıf Bazlı Hata Tablosu", level=2)
    add_formatted_table(
        doc,
        ["Yöntem", "Gerçek Sınıf", "Toplam", "Doğru Tahmin", "Hatalı Tahmin Adedi", "Sınıf Hata Oranı", "Yanlışların Kaydığı Sınıflar", "Genel Doğruluk / Macro F1"],
        [
            ["1. VADER", "Positive", "3.000", "2.936", "64", "%2,13", "20 Neutral, 44 Negative", "%51,53 / %39,65"],
            ["1. VADER", "Neutral", "3.000", "79", "2.921", "%97,37", "2.447 Positive, 474 Negative", "%51,53 / %39,65"],
            ["1. VADER", "Negative", "3.000", "1.623", "1.377", "%45,90", "1.247 Positive, 130 Neutral", "%51,53 / %39,65"],
            ["1. VADER", "TOPLAM", "9.000", "4.638", "4.362", "%48,47", "-", "%51,53 / %39,65"],
            ["2. BERT (Yelp-3L)", "Positive", "3.000", "2.613", "387", "%12,90", "371 Neutral, 16 Negative", "%80,96 / %81,11"],
            ["2. BERT (Yelp-3L)", "Neutral", "3.000", "2.288", "712", "%23,73", "320 Positive, 392 Negative", "%80,96 / %81,11"],
            ["2. BERT (Yelp-3L)", "Negative", "3.000", "2.385", "615", "%20,50", "13 Positive, 602 Neutral", "%80,96 / %81,11"],
            ["2. BERT (Yelp-3L)", "TOPLAM", "9.000", "7.286", "1.714", "%19,04", "-", "%80,96 / %81,11"],
            ["3. LLM (EVREN)", "Positive", "3.000", "2.848", "152", "%5,07", "147 Neutral, 5 Negative", "%83,63 / %83,35"],
            ["3. LLM (EVREN)", "Neutral", "3.000", "2.074", "926", "%30,87", "495 Positive, 431 Negative", "%83,63 / %83,35"],
            ["3. LLM (EVREN)", "Negative", "3.000", "2.605", "395", "%13,17", "5 Positive, 390 Neutral", "%83,63 / %83,35"],
            ["3. LLM (EVREN)", "TOPLAM", "9.000", "7.527", "1.473", "%16,37", "-", "%83,63 / %83,35"],
        ],
    )

    doc.add_heading("4.2. Türkçe Veri Seti (9.000 Yorum) Sınıf Bazlı Hata Tablosu", level=2)
    add_formatted_table(
        doc,
        ["Yöntem", "Gerçek Sınıf", "Toplam", "Doğru Tahmin", "Hatalı Tahmin Adedi", "Sınıf Hata Oranı", "Yanlışların Kaydığı Sınıflar", "Genel Doğruluk / Macro F1"],
        [
            ["1A. VADER (Çevirisiz Ham)", "Positive", "3.000", "331", "2.669", "%88,97", "2.652 Neutral, 17 Negative", "%35,37 / %24,47"],
            ["1A. VADER (Çevirisiz Ham)", "Neutral", "3.000", "2.818", "182", "%6,07", "143 Positive, 39 Negative", "%35,37 / %24,47"],
            ["1A. VADER (Çevirisiz Ham)", "Negative", "3.000", "34", "2.966", "%98,87", "125 Positive, 2.841 Neutral", "%35,37 / %24,47"],
            ["1A. VADER (Çevirisiz Ham)", "TOPLAM", "9.000", "3.183", "5.817", "%64,63", "-", "%35,37 / %24,47"],
            ["1B. VADER (Çevirili)", "Positive", "3.000", "2.823", "177", "%5,90", "104 Neutral, 73 Negative", "%53,99 / %45,15"],
            ["1B. VADER (Çevirili)", "Neutral", "3.000", "328", "2.672", "%89,07", "1.824 Positive, 848 Negative", "%53,99 / %45,15"],
            ["1B. VADER (Çevirili)", "Negative", "3.000", "1.708", "1.292", "%43,07", "985 Positive, 307 Neutral", "%53,99 / %45,15"],
            ["1B. VADER (Çevirili)", "TOPLAM", "9.000", "4.859", "4.141", "%46,01", "-", "%53,99 / %45,15"],
            ["2. BERTurk (3-Sınıflı)", "Positive", "3.000", "2.366", "634", "%21,13", "478 Neutral, 156 Negative", "%72,11 / %71,18"],
            ["2. BERTurk (3-Sınıflı)", "Neutral", "3.000", "1.353", "1.647", "%54,90", "180 Positive, 1.467 Negative", "%72,11 / %71,18"],
            ["2. BERTurk (3-Sınıflı)", "Negative", "3.000", "2.771", "229", "%7,63", "47 Positive, 182 Neutral", "%72,11 / %71,18"],
            ["2. BERTurk (3-Sınıflı)", "TOPLAM", "9.000", "6.490", "2.510", "%27,89", "-", "%72,11 / %71,18"],
            ["3. LLM (EVREN)", "Positive", "3.000", "2.688", "312", "%10,40", "296 Neutral, 16 Negative", "%79,11 / %78,56"],
            ["3. LLM (EVREN)", "Neutral", "3.000", "1.776", "1.224", "%40,80", "249 Positive, 975 Negative", "%79,11 / %78,56"],
            ["3. LLM (EVREN)", "Negative", "3.000", "2.656", "344", "%11,47", "36 Positive, 308 Neutral", "%79,11 / %78,56"],
            ["3. LLM (EVREN)", "TOPLAM", "9.000", "7.120", "1.880", "%20,89", "-", "%79,11 / %78,56"],
        ],
    )

    doc.add_heading("4.3. Karşılaştırmalı Görselleştirmeler ve Karmaşıklık Matrisleri", level=2)
    figures = [
        (FIG_CONFUSION_MATRICES, "Şekil 1: İngilizce ve Türkçe Veri Setlerinde 3 Yöntemin Karmaşıklık Matrisleri (Confusion Matrices)"),
        (FIG_CLASS_ERRORS, "Şekil 2: Sınıf Bazlı Hatalı Tahmin Adetlerinin Karşılaştırması (3.000'er Yorum Üzerinden)"),
        (FIG_ACCURACY_F1, "Şekil 3: Genel Doğruluk (Accuracy) ve Macro F1-Skoru Karşılaştırması"),
        (FIG_TR_TRANSLATION, "Şekil 4: Türkçe Veri Setinde Çevirisiz (Ham) VADER ile Çevirili VADER Karşılaştırması"),
    ]
    for fig_path, caption in figures:
        if fig_path.exists():
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(str(fig_path), width=Inches(6.2))
            p_cap = doc.add_paragraph(caption)
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p_cap.runs:
                run.italic = True
                run.font.size = Pt(9.5)

    doc.add_heading("5. Hataların Akademik Yorumlanması ve Niteliksel Örnek Analizi", level=1)
    doc.add_heading("5.1. VADER Yönteminin Nötr Sınıf Körlüğü ve Pozitif Yanlılığı (Positivity Bias)", level=2)
    doc.add_paragraph(
        "Deneysel bulgular, VADER'ın en büyük zaafının 'Neutral' (Nötr) sınıfında olduğunu ortaya koymaktadır. "
        "VADER, İngilizcede 3.000 Nötr yorumun 2.921'ini (%97,37 hata), çevrilmiş Türkçede ise 2.672'sini (%89,07 hata) "
        "yanlış sınıflandırmıştır. Bunun iki temel nedeni vardır:\n"
        "1) Dar Nötr Eşik Aralığı: VADER yalnızca -0.05 ile +0.05 aralığını Nötr kabul eder. Oysa 3 yıldızlı bir restoran yorumunda "
        "kullanıcılar hem olumlu hem olumsuz özellikleri birlikte yazar. Tek bir olumlu sıfat ('friendly', 'nice', 'good') "
        "bileşik skoru +0.05'in üzerine taşıyarak yorumun 'Positive' olarak etiketlenmesine yol açmaktadır.\n"
        "2) Bağlaç ve Bağlam Kaybı: Olumsuz (1-2 yıldızlı) yorumlarda bile müşteriler söze olumlu bir ifadeyle başlayıp "
        "('Manzarası güzel diye oturmak istedik fakat...') ardından ağır eleştiriler getirmektedir. VADER cümledeki ana yargıyı "
        "ağırlıklandıramadığı için İngilizcede 1.247, Türkçede 985 Negatif yorumu yanlışlıkla Pozitif tahmin etmiştir."
    )

    ex_tr_neg = examples["tr_neg_vader_pos"]
    ex_en_neu = examples["en_neu_vader_pos"]
    doc.add_paragraph(
        f"• Veri Setinden Gerçek Örnek 1 (Türkçe - Gerçek Sınıf: Negative [1 Yıldız] | Mekan: {ex_tr_neg['place_name']}):\n"
        f"  Orijinal Yorum: \"{ex_tr_neg['text']}\"\n"
        f"  İngilizce Çevirisi: \"{ex_tr_neg['translated_text']}\"\n"
        f"  Model Tahminleri -> VADER: {ex_tr_neg['vader_pred']} (compound={ex_tr_neg['vader_compound']}) [HATALI] | "
        f"BERTurk: {ex_tr_neg['bert_pred']} [DOĞRU] | LLM: {ex_tr_neg['llm_pred']} [DOĞRU]\n"
        f"• Veri Setinden Gerçek Örnek 2 (İngilizce - Gerçek Sınıf: Neutral [3 Yıldız] | Mekan: {ex_en_neu['place_name']}):\n"
        f"  Orijinal Yorum: \"{ex_en_neu['text']}\"\n"
        f"  Model Tahminleri -> VADER: {ex_en_neu['vader_pred']} (compound={ex_en_neu['vader_compound']}) [HATALI] | "
        f"BERT: {ex_en_neu['bert_pred']} [DOĞRU] | LLM: {ex_en_neu['llm_pred']} [DOĞRU]"
    )

    doc.add_heading("5.2. BERT ve BERTurk Modellerinin Başarımı ve Sınır Durumları", level=2)
    doc.add_paragraph(
        "Derin öğrenme tabanlı BERT modelleri, VADER'a kıyasla toplam hata sayısını İngilizcede 4.362'den 1.714'e (%60,7 azalma), "
        "Türkçede ise 4.141'den 2.510'a (%39,4 azalma) düşürmüştür. Özellikle zıt kutuplar arasındaki hata neredeyse sıfırlanmış; "
        "İngilizce BERT modeli 3.000 Negatif yorumdan yalnızca 13 tanesini Pozitif ile karıştırmıştır.\n"
        "Bununla birlikte Türkçe BERTurk modelinde Nötr sınıfındaki 3.000 yorumun 1.467 tanesinin 'Negative' sınıfa kaydığı görülmektedir. "
        "Bu durumun nedeni, Türkiye'deki Google Maps kullanıcılarının 3 yıldız (Nötr) verdikleri işletmelerde memnuniyetsizliklerini "
        "daha baskın kelimelerle ('fiyatlar aşırı pahalı', 'servis yavaş') ifade etmesi ve BERTurk'ün bu olumsuz anahtar kelimelere "
        "yüksek dikkat (attention) ağırlığı vermesidir."
    )

    doc.add_heading("5.3. Büyük Dil Modeli (LLM - EVREN) Yaklaşımının Üstünlüğü", level=2)
    ex_tr_neu = examples["tr_neu_llm_correct"]
    doc.add_paragraph(
        "EVREN API üzerinden çalıştırılan 'qwen3.8-flash-next' LLM modeli, herhangi bir ince ayar (fine-tuning) yapılmaksızın "
        "yalnızca yapılandırılmış istem mühendisliği (prompting) ile hem İngilizcede (%83,63 doğruluk, 1.473 hata) hem de Türkçede "
        "(%79,11 doğruluk, 1.880 hata) tüm yöntemler arasında en yüksek başarımı elde etmiştir. "
        "LLM; ironiyi, devrik Türkçe cümleleri ve 'fiyatına göre idare eder / ortalama bir mekan' gibi gri alandaki ifadeleri "
        "bütüncül olarak muhakeme edebilmiştir.\n"
        f"• Veri Setinden Gerçek Örnek 3 (Türkçe - Gerçek Sınıf: Neutral [3 Yıldız] | Mekan: {ex_tr_neu['place_name']}):\n"
        f"  Orijinal Yorum: \"{ex_tr_neu['text']}\"\n"
        f"  Model Tahminleri -> VADER: {ex_tr_neu['vader_pred']} [HATALI] | BERTurk: {ex_tr_neu['bert_pred']} [HATALI] | LLM: {ex_tr_neu['llm_pred']} [DOĞRU]"
    )

    doc.add_heading("6. Genel Sonuç, Yöntemsel Karşılaştırma ve Kaynakça", level=1)
    doc.add_paragraph(
        "18.000 gerçek kullanıcı yorumu (9.000 İngilizce + 9.000 Türkçe) üzerinde yürütülen bu kapsamlı çalışmada:\n"
        "1) Kural tabanlı VADER yaklaşımının çok hızlı (9.000 yorumda ~1.2 sn) olmasına karşın bağlamı kavrayamadığı, "
        "Nötr sınıfında %89-%97 hata yaptığı ve Türkçe metinlerde mutlaka çeviri aşamasına ihtiyaç duyduğu (çevirisiz %35,37 -> çevirili %53,99) kanıtlanmıştır.\n"
        "2) Derin öğrenme tabanlı 3 sınıflı BERT (%80,96) ve BERTurk (%72,11) modellerinin çift yönlü dikkat mekanizması sayesinde "
        "zıt kutuplu hataları (Pozitif <-> Negatif) neredeyse sıfırladığı ve Nötr sınıfı başarımını dramatik biçimde artırdığı görülmüştür.\n"
        "3) Güncel Büyük Dil Modeli (EVREN Qwen3.8-Flash-Next) yaklaşımının ise sıfır-örnek (zero-shot) yapılandırılmış istem mühendisliğiyle "
        "hem İngilizcede (%83,63 doğruluk, 1.473 hata) hem de Türkçede (%79,11 doğruluk, 1.880 hata) en yüksek doğruluğa ve en düşük toplam hata sayısına ulaştığı tespit edilmiştir.\n\n"
        "Kaynaklar ve Erişim Bağlantıları:\n"
        "• Hutto, C. J., & Gilbert, E. (2014). VADER: A Parsimonious Rule-based Model for Sentiment Analysis of Social Media Text. ICWSM.\n"
        "• Devlin, J., Chang, M. W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding. NAACL.\n"
        "• Schweter, S. (2020). BERTurk - BERT models for Turkish (dbmdz/bert-base-turkish-cased).\n"
        "• Türkçe Veri Seti: Hugging Face - opdullah/turkish-google-maps-reviews\n"
        "• İngilizce Veri Seti: Hugging Face - jaimik69/Yelp-Restaurant-Dataset & hasnat79/yelp_food_business_review_dataset_2M\n"
        "• İngilizce BERT Modeli: Hugging Face - gohbwj/sentiment-fine-tuned-yelp-3L\n"
        "• Türkçe BERTurk Modeli: Hugging Face - incidelen/bert-base-turkish-sentiment-analysis-cased\n"
        "• LLM Çıkarım Altyapısı: EVREN Milli Yapay Zeka Platformu API (https://evren-llmapi.ssyz.org.tr/v1 - qwen3.8-flash-next)\n"
        "• Proje GitHub Deposu: https://github.com/erenaksoyy/BIM448_Duygu_Analizi_Projesi"
    )

    doc.save(REPORT_DOCX_PATH)
    print(f"Word raporu kaydedildi: {REPORT_DOCX_PATH.name}")


def main() -> None:
    """Tüm karşılaştırma grafiklerini ve Word proje raporunu üretir."""
    vader, bert, llm = load_all_summary_metrics()
    plot_six_panel_confusion_matrices(vader, bert, llm)
    plot_accuracy_and_f1_comparison(vader, bert, llm)
    plot_class_wise_error_counts(vader, bert, llm)
    plot_turkish_vader_translation_effect(vader)
    examples = extract_qualitative_examples()
    build_word_report(examples)
    print("\nTüm grafikler ve Word raporu Clean Code standartlarında oluşturuldu.")


if __name__ == "__main__":
    main()
