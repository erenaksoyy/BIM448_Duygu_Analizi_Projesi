"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: 10 Dakikalık Akademik Sunum Dosyası (.pptx) Üreticisi

Bu modül, 2 kişilik proje grubunun 10 dakikalık sunum videosunda kullanacağı
12 slaytlık geniş ekran (16:9) PowerPoint sunumunu ('BIM448_Proje_Sunumu.pptx')
yüksek çözünürlüklü grafikler, karşılaştırma tabloları ve her slayta özel
konuşmacı notlarıyla (Speaker Notes) birlikte otomatik olarak üretir.
"""

from pathlib import Path
from typing import List, Sequence

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ============================================================================
# SABİTLER VE TASARIM PALETİ (CONSTANTS & THEME PALETTE)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
OUTPUT_PPTX_PATH: Path = BASE_DIR / "BIM448_Proje_Sunumu.pptx"

FIG_CONFUSION_MATRICES: Path = BASE_DIR / "1_confusion_matrices_6panel.png"
FIG_ACCURACY_F1: Path = BASE_DIR / "2_accuracy_and_f1_comparison.png"
FIG_CLASS_ERRORS: Path = BASE_DIR / "3_class_wise_error_counts.png"
FIG_TR_TRANSLATION: Path = BASE_DIR / "4_turkish_translation_impact_vader.png"

SLIDE_WIDTH: Inches = Inches(13.333)
SLIDE_HEIGHT: Inches = Inches(7.5)

COLOR_NAVY_DARK = RGBColor(15, 23, 42)
COLOR_BLUE_ACCENT = RGBColor(37, 99, 235)
COLOR_LIGHT_BG = RGBColor(248, 250, 252)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_TEXT_DARK = RGBColor(30, 41, 59)
COLOR_TEXT_MUTED = RGBColor(100, 116, 139)
COLOR_TABLE_HEADER = RGBColor(30, 58, 138)
COLOR_TABLE_ALT = RGBColor(241, 245, 249)


# ============================================================================
# YARDIMCI SLAYT BİLEŞEN FONKSİYONLARI
# ============================================================================

def add_speaker_notes(slide: object, note_text: str) -> None:
    """Slaytın konuşmacı notları (Speaker Notes) bölümüne sunum konuşma metnini ekler."""
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = note_text


def apply_slide_header(
    slide: object, title_text: str, subtitle_text: str, presenter_badge: str
) -> None:
    """Standart üst başlık şeridi, alt başlık ve konuşmacı rozetini slayta ekler."""
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = COLOR_LIGHT_BG
    bg_shape.line.fill.background()

    header_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, Inches(1.15)
    )
    header_bar.fill.solid()
    header_bar.fill.fore_color.rgb = COLOR_NAVY_DARK
    header_bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9.5), Inches(0.55))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_WHITE

    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.68), Inches(9.5), Inches(0.38))
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = subtitle_text
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = RGBColor(147, 197, 253)

    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.2), Inches(0.28), Inches(2.65), Inches(0.55)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_BLUE_ACCENT
    badge.line.fill.background()
    tf_badge = badge.text_frame
    p_badge = tf_badge.paragraphs[0]
    p_badge.alignment = PP_ALIGN.CENTER
    p_badge.text = presenter_badge
    p_badge.font.size = Pt(12)
    p_badge.font.bold = True
    p_badge.font.color.rgb = COLOR_WHITE


def add_bullet_card(
    slide: object,
    left: float,
    top: float,
    width: float,
    height: float,
    card_title: str,
    bullets: Sequence[str],
) -> None:
    """Beyaz arka planlı, başlıklı ve madde işaretli içerik kartı ekler."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = RGBColor(203, 213, 225)
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(
        Inches(left + 0.25), Inches(top + 0.18), Inches(width - 0.5), Inches(height - 0.35)
    )
    tf = tb.text_frame
    tf.word_wrap = True

    p_head = tf.paragraphs[0]
    p_head.text = card_title
    p_head.font.size = Pt(18)
    p_head.font.bold = True
    p_head.font.color.rgb = COLOR_TABLE_HEADER
    p_head.space_after = Pt(10)

    for bullet_text in bullets:
        p = tf.add_paragraph()
        p.text = f"•  {bullet_text}"
        p.font.size = Pt(14)
        p.font.color.rgb = COLOR_TEXT_DARK
        p.space_after = Pt(8)


def add_styled_slide_table(
    slide: object,
    left: float,
    top: float,
    width: float,
    height: float,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    font_size: int = 12,
) -> None:
    """Slayta kurumsal renkli ve okunaklı karşılaştırma tablosu ekler."""
    table_shape = slide.shapes.add_table(
        len(rows) + 1, len(headers), Inches(left), Inches(top), Inches(width), Inches(height)
    )
    table = table_shape.table

    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_TABLE_HEADER
        cell.text = header_text
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(font_size)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE
            p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT

    for row_idx, row_data in enumerate(rows):
        is_alt = row_idx % 2 == 1
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_TABLE_ALT if is_alt else COLOR_WHITE
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_size)
                p.font.bold = "TOPLAM" in str(row_data[0]) or col_idx == 0
                p.font.color.rgb = COLOR_TEXT_DARK
                p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT


# ============================================================================
# ANA SUNUM OLUŞTURUCU (12 SLAYT)
# ============================================================================

def build_presentation() -> None:
    """12 slaytlık proje sunumunu ve konuşmacı notlarını oluşturur."""
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]

    # ------------------------------------------------------------------------
    # SLAYT 1: KAPAK SLAYTI
    # ------------------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_NAVY_DARK
    bg1.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.3), Inches(5.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_uni = tf1.paragraphs[0]
    p_uni.text = "DOĞUŞ ÜNİVERSİTESİ • BİLGİSAYAR MÜHENDİSLİĞİ BÖLÜMÜ\nBİM448 VERİ MADENCİLİĞİ VE BİLGİ KEŞFİ DERSİ PROJE SUNUMU"
    p_uni.font.size = Pt(16)
    p_uni.font.bold = True
    p_uni.font.color.rgb = RGBColor(147, 197, 253)
    p_uni.space_after = Pt(24)

    p_main = tf1.add_paragraph()
    p_main.text = "Karşılaştırmalı Duygu Analizi (Sentiment Analysis):\nİngilizce ve Türkçe Mekan Yorumları Üzerinde\nVADER, BERT ve LLM Yöntemlerinin İncelenmesi"
    p_main.font.size = Pt(34)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_WHITE
    p_main.space_after = Pt(28)

    p_meta = tf1.add_paragraph()
    p_meta.text = (
        "• Veri Ölçeği: 9.000 İngilizce + 9.000 Türkçe = 18.000 Gerçek Restoran & Kafe Yorumu\n"
        "• Yöntemler: 1) VADER (Kural Tabanlı)   2) BERT & BERTurk (Derin Öğrenme)   3) EVREN LLM (Qwen3.8-Flash-Next)\n"
        "• Sunum Süresi Planı: 1. Üye (Slayt 1–6 | 0:00–5:00 dk)  •  2. Üye (Slayt 7–12 | 5:00–10:00 dk)"
    )
    p_meta.font.size = Pt(16)
    p_meta.font.color.rgb = RGBColor(203, 213, 225)

    add_speaker_notes(
        s1,
        "[1. ÜYE - 0:00 - 0:45]\n"
        "Merhaba Hocam, BİM448 Veri Madenciliği ve Bilgi Keşfi dersi kapsamında hazırladığımız "
        "Karşılaştırmalı Duygu Analizi projemizin sunumuna hoş geldiniz. Bu projede geleneksel kural tabanlı VADER, "
        "derin öğrenme tabanlı BERT/BERTurk ve Büyük Dil Modeli tabanlı EVREN LLM yöntemlerini hem İngilizce hem de Türkçe "
        "gerçek restoran ve kafe yorumları üzerinde karşılaştırdık. Sunumun ilk yarısında ben veri setlerimizi, VADER ve BERT "
        "sonuçlarımızı aktaracağım; ikinci yarısında ise grup arkadaşım LLM sonuçlarını ve sınıf bazlı hata analizimizi sunacak."
    )

    # ------------------------------------------------------------------------
    # SLAYT 2: PROJENİN AMACI, KAPSAMI VE MİMARİ AKIŞI
    # ------------------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s2,
        "1. Projenin Amacı, Kapsamı ve Yöntemsel Çerçeve",
        "Üç farklı kuşak doğal dil işleme yaklaşımının iki dil ve üç duygu sınıfı üzerinde kıyaslanması",
        "1. Üye (0:45–1:30)",
    )
    add_bullet_card(
        s2, 0.5, 1.45, 6.0, 5.5,
        "🎯 Projenin Temel Hedefleri",
        [
            "Geleneksel Sözlük/Kural Tabanlı (VADER), Derin Öğrenme (BERT/BERTurk) ve Güncel LLM (EVREN API) yaklaşımlarını adil koşullarda karşılaştırmak.",
            "Ödevdeki 'sınıf başına en az 1.000 yorum (3.000 toplam)' alt sınırını 3 katına çıkararak her sınıfta 3.000'er adet (dil başına 9.000, toplamda 18.000 yorum) ile çalışmak.",
            "Her iki dilde ve her duygu sınıfında (Positive, Neutral, Negative) yapılan hatalı tahmin sayılarını tek tek hesaplamak.",
            "Modellerin neden hata yaptığını (nötr sınıf körlüğü, bağlaçlar, çeviri kaybı, ironi) gerçek yorum örnekleriyle açıklamak.",
        ],
    )
    add_bullet_card(
        s2, 6.8, 1.45, 6.0, 5.5,
        "⚙️ Uygulanan 3 Yöntem Mimarisi",
        [
            "1. Yöntem – VADER (Kural Tabanlı): İngilizceye doğrudan; Türkçeye ise Google Translate API ile İngilizceye çevrilerek (+ çevirisiz kontrol deneyiyle) uygulandı.",
            "2. Yöntem – BERT & BERTurk (Derin Öğrenme): Standart 2 sınıflı modeller Nötr tahmin edemediği için doğrudan 3 sınıflı fine-tuned BERT ve BERTurk modelleri Colab L4 GPU'da çalıştırıldı.",
            "3. Yöntem – LLM (EVREN API - qwen3.8-flash-next): Türkiye'nin milli yapay zeka platformu EVREN üzerinden yapılandırılmış toplu istem mühendisliği (batch prompting) uygulandı.",
        ],
    )
    add_speaker_notes(
        s2,
        "[1. ÜYE - 0:45 - 1:30]\n"
        "Projemizin temel amacı kural tabanlı, derin öğrenme tabanlı ve LLM tabanlı duygu analizi yaklaşımlarını İngilizce ve Türkçe "
        "metinler üzerinde karşılaştırmaktır. İstatistiksel güvenilirliği artırmak için minimum 1.000'er yorum sınırıyla yetinmedik; "
        "her sınıfta 3.000'er adet olmak üzere dil başına 9.000, toplamda 18.000 yorum kullandık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 3: VERİ SETLERİNİN KAYNAKLARI VE ÖN İŞLEME
    # ------------------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s3,
        "2. İngilizce ve Türkçe Veri Setleri (18.000 Mekan Yorumu)",
        "Aynı konu alanından (Restoran & Kafe) ve birebir aynı şemada dengeli veri seti tasarımı",
        "1. Üye (1:30–2:30)",
    )
    add_styled_slide_table(
        s3, 0.5, 1.45, 12.33, 3.3,
        ["Özellik", "Türkçe Veri Seti (Google Maps Türkiye)", "İngilizce Veri Seti (Yelp ABD)"],
        [
            ["Hugging Face Kaynağı", "opdullah/turkish-google-maps-reviews", "jaimik69/Yelp-Restaurant-Dataset + hasnat79/yelp_food_2M"],
            ["Toplam Ham Veri Havuzu", "838.583 yorum (~251.000 Restoran & Kafe)", "2.007.040 restoran yorumu (64.629 işletme)"],
            ["Seçilen Dengeli Örneklem", "9.000 Yorum (3.000 Pozitif, 3.000 Nötr, 3.000 Negatif)", "9.000 Yorum (3.000 Pozitif, 3.000 Nötr, 3.000 Negatif)"],
            ["Sınıf Etiketleme Kriteri", "1-2★: Negative | 3★: Neutral | 4-5★: Positive", "1-2★: Negative | 3★: Neutral | 4-5★: Positive"],
            ["Ortak Sütun Şeması", "id, place_name, place_category, rating, text, true_label", "id, place_name, place_category, rating, text, true_label"],
        ],
        font_size=12,
    )
    add_bullet_card(
        s3, 0.5, 4.95, 12.33, 2.1,
        "🧹 Veri Ön İşleme ve Temizleme Adımları (prepare_datasets.py)",
        [
            "Sektör Eşitlemesi: Hastane/okul gibi alakasız kategoriler çıkarılarak yalnızca Restoran, Kafe ve Yeme-İçme mekanları filtrelendi.",
            "Doğal Dil & Gürültü Filtresi: Türkiye'deki mekanlara turistlerin yazdığı Arapça/Rusça/İngilizce yorumlar, tekrar eden kopyalar ve <25 veya >500 karakter metinler elendi.",
            "Tekrarlanabilirlik: Sabit rastgelelik tohumu (RANDOM_SEED = 42) ile her sınıftan tam 3.000'er adet örneklem çekildi (0 eksik veri).",
        ],
    )
    add_speaker_notes(
        s3,
        "[1. ÜYE - 1:30 - 2:30]\n"
        "İki dil arasındaki başarı farkının konu farkından etkilenmemesi için her iki dilde de Restoran ve Kafe yorumlarını seçtik. "
        "Türkçede 838 bin yorumluk Google Maps veri setinden, İngilizcede ise 2 milyon yorumluk Yelp veri setinden işletme isimlerini ve "
        "kategorilerini eşleştirerek 3.000 Pozitif, 3.000 Nötr ve 3.000 Negatif olmak üzere 9.000'er yorumluk tertemiz iki veri seti hazırladık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 4: 1. YÖNTEM - VADER SONUÇLARI (İNGİLİZCE & TÜRKÇE)
    # ------------------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s4,
        "3. Yöntem 1: VADER Duygu Analizi ve Sınıf Bazlı Hatalar",
        "Kural tabanlı yaklaşım (compound >= 0.05: Positive, <= -0.05: Negative, arası: Neutral)",
        "1. Üye (2:30–3:45)",
    )
    add_styled_slide_table(
        s4, 0.5, 1.4, 12.33, 3.6,
        ["Dil ve Uygulama Şekli", "Gerçek Sınıf", "Toplam", "Doğru", "Hatalı Tahmin Adedi", "Hata Oranı", "Yanlışların Kaydığı Sınıflar"],
        [
            ["İngilizce VADER (%51,53)", "Positive", "3.000", "2.936", "64", "%2,13", "20 Neutral, 44 Negative"],
            ["İngilizce VADER (%51,53)", "Neutral", "3.000", "79", "2.921", "%97,37", "2.447 Positive, 474 Negative"],
            ["İngilizce VADER (%51,53)", "Negative", "3.000", "1.623", "1.377", "%45,90", "1.247 Positive, 130 Neutral"],
            ["İngilizce VADER TOPLAM", "Tüm Sınıflar", "9.000", "4.638", "4.362", "%48,47", "Macro F1: %39,65"],
            ["Türkçe VADER - Çevirili (%53,99)", "Positive", "3.000", "2.823", "177", "%5,90", "104 Neutral, 73 Negative"],
            ["Türkçe VADER - Çevirili (%53,99)", "Neutral", "3.000", "328", "2.672", "%89,07", "1.824 Positive, 848 Negative"],
            ["Türkçe VADER - Çevirili (%53,99)", "Negative", "3.000", "1.708", "1.292", "%43,07", "985 Positive, 307 Neutral"],
            ["Türkçe VADER (Çevirili) TOPLAM", "Tüm Sınıflar", "9.000", "4.859", "4.141", "%46,01", "Macro F1: %45,15"],
        ],
        font_size=11,
    )
    add_bullet_card(
        s4, 0.5, 5.2, 12.33, 1.9,
        "🔍 VADER Temel Bulguları: Nötr Sınıf Körlüğü ve Pozitif Yanlılığı",
        [
            "Dar Nötr Aralığı (-0.05 ile +0.05): 3 yıldızlı Nötr yorumlarda geçen tek bir olumlu kelime ('friendly', 'nice') skoru +0.05 üstüne çıkardığı için VADER İngilizcede 2.921, Türkçede 2.672 Nötr yorumu hatalı sınıflandırmıştır.",
            "Bağlaç Körlüğü: 'Manzarası güzel fakat servis berbat' gibi yorumlarda başlangıçtaki olumlu kelime yüzünden Negatif yorumların ~%44'ü Pozitif tahmin edilmiştir.",
        ],
    )
    add_speaker_notes(
        s4,
        "[1. ÜYE - 2:30 - 3:45]\n"
        "İlk yöntemimiz olan VADER'da İngilizce veri setinde %51,53 doğruluk ve 4.362 hatalı tahmin; İngilizceye çevrilmiş Türkçe "
        "veri setinde ise %53,99 doğruluk ve 4.141 hatalı tahmin elde ettik. Tabloda görüldüğü üzere VADER Pozitif sınıfında çok az hata "
        "yaparken, Nötr sınıfında İngilizcede %97,37 (2.921 hata), Türkçede %89,07 (2.672 hata) oranında yanılmıştır."
    )

    # ------------------------------------------------------------------------
    # SLAYT 5: TÜRKÇEDE ÇEVİRİSİZ VS ÇEVİRİLİ VADER ETKİSİ
    # ------------------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s5,
        "4. Türkçe VADER Analizinde Çevirinin Etkisi (Ham vs. Çevirili)",
        "VADER İngilizce sözlük tabanlı olduğu için Türkçe metinlerin çevrilmesi zorunludur",
        "1. Üye (3:45–4:15)",
    )
    if FIG_TR_TRANSLATION.exists():
        s5.shapes.add_picture(str(FIG_TR_TRANSLATION), Inches(0.5), Inches(1.4), width=Inches(7.7))
    add_bullet_card(
        s5, 8.4, 1.4, 4.4, 5.6,
        "📌 Deneysel Kanıt",
        [
            "Çevirisiz Ham Türkçe VADER: 9.000 yorumun 8.311 tanesine 0.0000 (Neutral) skoru vermiştir (Doğruluk: %35,37 | Toplam Hata: 5.817).",
            "Sadece emoji 🙂 ve 'super/ok' gibi kelimelerden 331 Pozitif yakalayabilmiştir.",
            "Google Translate ile Çeviri Sonrası: Doğruluk %35,37'den %53,99'a yükselmiş, Pozitif sınıf doğru tahmin sayısı 331'den 2.823'e, Negatif sınıf doğru tahmini 34'ten 1.708'e çıkmıştır.",
        ],
    )
    add_speaker_notes(
        s5,
        "[1. ÜYE - 3:45 - 4:15]\n"
        "Ders içi uygulamamızda vurgulandığı gibi VADER yalnızca İngilizce sözlükle çalışır. Bunu sayısal olarak kanıtlamak için "
        "9.000 Türkçe yorumu hem çevirmeden hem de İngilizceye çevirerek test ettik. Çevirisiz kullanımda VADER 9.000 yorumun 8.311'ine "
        "sıfır puan verip Nötr derken, çeviri sonrasında doğruluk %35,37'den %53,99'a yükselmiştir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 6: 2. YÖNTEM - BERT & BERTurk SONUÇLARI
    # ------------------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s6,
        "5. Yöntem 2: BERT ve BERTurk Derin Öğrenme Sonuçları",
        "Google Colab NVIDIA L4 GPU üzerinde 3 sınıflı fine-tuned Transformer modelleri",
        "1. Üye (4:15–5:00)",
    )
    add_styled_slide_table(
        s6, 0.5, 1.4, 12.33, 3.6,
        ["Dil ve Model Mimarisi", "Gerçek Sınıf", "Toplam", "Doğru", "Hatalı Tahmin Adedi", "Hata Oranı", "Yanlışların Kaydığı Sınıflar"],
        [
            ["İngilizce BERT (gohbwj/yelp-3L)", "Positive", "3.000", "2.613", "387", "%12,90", "371 Neutral, sadece 16 Negative"],
            ["İngilizce BERT (gohbwj/yelp-3L)", "Neutral", "3.000", "2.288", "712", "%23,73", "320 Positive, 392 Negative"],
            ["İngilizce BERT (gohbwj/yelp-3L)", "Negative", "3.000", "2.385", "615", "%20,50", "sadece 13 Positive, 602 Neutral"],
            ["İngilizce BERT TOPLAM (%80,96)", "Tüm Sınıflar", "9.000", "7.286", "1.714", "%19,04", "Macro F1: %81,11 (VADER'dan 2.648 az hata)"],
            ["Türkçe BERTurk (incidelen/3-class)", "Positive", "3.000", "2.366", "634", "%21,13", "478 Neutral, 156 Negative"],
            ["Türkçe BERTurk (incidelen/3-class)", "Neutral", "3.000", "1.353", "1.647", "%54,90", "180 Positive, 1.467 Negative"],
            ["Türkçe BERTurk (incidelen/3-class)", "Negative", "3.000", "2.771", "229", "%7,63", "47 Positive, 182 Neutral"],
            ["Türkçe BERTurk TOPLAM (%72,11)", "Tüm Sınıflar", "9.000", "6.490", "2.510", "%27,89", "Macro F1: %71,18 (VADER'dan 1.631 az hata)"],
        ],
        font_size=11,
    )
    add_bullet_card(
        s6, 0.5, 5.2, 12.33, 1.9,
        "🚀 Derin Öğrenme (Self-Attention) Kazanımları",
        [
            "3 Sınıflı Model Seçimi: Popüler SST-2 veya savasy/bert-base-turkish modelleri sadece 2 sınıflı olduğu için Nötr tahmin edemez; projemizde 3 sınıflı fine-tuned BERT ve BERTurk modelleri kullanılmıştır.",
            "Zıt Kutupların Ayrışması: İngilizce BERT'te 3.000 Negatif yorumdan yalnızca 13'ü Pozitif'e kaymıştır (VADER'da 1.247 idi!).",
        ],
    )
    add_speaker_notes(
        s6,
        "[1. ÜYE - 4:15 - 5:00]\n"
        "İkinci yöntemimiz olan BERT aşamasında çok önemli bir tasarım kararı aldık: Hazır SST-2 ve savasy modelleri yalnızca Pozitif ve Negatif "
        "ürettiği için Nötr sınıfını tahmin edememektedir. Biz doğrudan 3 sınıflı olarak fine-tune edilmiş İngilizce BERT ve Türkçe BERTurk "
        "modellerini Colab L4 GPU üzerinde çalıştırdık. İngilizcede doğruluk %80,96'ya çıktı ve toplam hata 4.362'den 1.714'e düştü; "
        "Türkçede ise BERTurk %72,11 doğruluk sağladı. Şimdi sözü 3. yöntemimiz olan LLM ve hata analizi için arkadaşıma bırakıyorum."
    )

    # ------------------------------------------------------------------------
    # SLAYT 7: 3. YÖNTEM - EVREN LLM MİMARİSİ VE SONUÇLARI
    # ------------------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s7,
        "6. Yöntem 3: Büyük Dil Modeli (LLM – EVREN API) Sonuçları",
        "Milli Yapay Zeka Platformu EVREN (qwen3.8-flash-next) üzerinde Yapılandırılmış İstem Mühendisliği",
        "2. Üye (5:00–6:15)",
    )
    add_styled_slide_table(
        s7, 0.5, 1.4, 12.33, 3.6,
        ["Dil ve LLM Modeli", "Gerçek Sınıf", "Toplam", "Doğru", "Hatalı Tahmin Adedi", "Hata Oranı", "Yanlışların Kaydığı Sınıflar"],
        [
            ["İngilizce LLM (EVREN Qwen3.8)", "Positive", "3.000", "2.848", "152", "%5,07", "147 Neutral, sadece 5 Negative"],
            ["İngilizce LLM (EVREN Qwen3.8)", "Neutral", "3.000", "2.074", "926", "%30,87", "495 Positive, 431 Negative"],
            ["İngilizce LLM (EVREN Qwen3.8)", "Negative", "3.000", "2.605", "395", "%13,17", "sadece 5 Positive, 390 Neutral"],
            ["İngilizce LLM TOPLAM (%83,63)", "Tüm Sınıflar", "9.000", "7.527", "1.473", "%16,37", "Macro F1: %83,35 (En Düşük Toplam Hata!)"],
            ["Türkçe LLM (EVREN Qwen3.8)", "Positive", "3.000", "2.688", "312", "%10,40", "296 Neutral, sadece 16 Negative"],
            ["Türkçe LLM (EVREN Qwen3.8)", "Neutral", "3.000", "1.776", "1.224", "%40,80", "249 Positive, 975 Negative"],
            ["Türkçe LLM (EVREN Qwen3.8)", "Negative", "3.000", "2.656", "344", "%11,47", "sadece 36 Positive, 308 Neutral"],
            ["Türkçe LLM TOPLAM (%79,11)", "Tüm Sınıflar", "9.000", "7.120", "1.880", "%20,89", "Macro F1: %78,56 (Türkçede En Yüksek Başarı!)"],
        ],
        font_size=11,
    )
    add_bullet_card(
        s7, 0.5, 5.2, 12.33, 1.9,
        "🤖 İstem Mühendisliği (Prompt Engineering) ve Parametre Detayları",
        [
            "Parametreler: Deterministik sonuç için temperature=0.0, hızlı çıkarım için enable_thinking=False ve 50'şerli JSON dizi çıktısı (batch_size=50, 8 paralel iş parçacığı).",
            "Üstünlük: Hiçbir fine-tuning yapılmadan (Zero-Shot) hem İngilizcede (%83,63) hem Türkçede (%79,11) tüm yöntemler arasında 1. sıraya yerleşmiştir.",
        ],
    )
    add_speaker_notes(
        s7,
        "[2. ÜYE - 5:00 - 6:15]\n"
        "Teşekkür ederim. Üçüncü yöntemimiz olan Büyük Dil Modeli aşamasında Savunma Sanayii Başkanlığı ve SAYZEK tarafından geliştirilen "
        "EVREN platformundaki qwen3.8-flash-next modelini OpenAI uyumlu API üzerinden kullandık. Temperature değerini 0.0 olarak sabitleyip "
        "yapılandırılmış JSON istem mühendisliği uyguladık. Sonuç olarak LLM, İngilizcede %83,63 doğruluk ve sadece 1.473 hata ile; "
        "Türkçede ise %79,11 doğruluk ve 1.880 hata ile her iki dilde de en başarılı yöntem oldu."
    )

    # ------------------------------------------------------------------------
    # SLAYT 8: 6'LI CONFUSION MATRIX KARŞILAŞTIRMASI
    # ------------------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s8,
        "7. Altılı Karmaşıklık Matrisi (Confusion Matrices) Karşılaştırması",
        "Üst Satır: İngilizce (VADER, BERT, LLM)  •  Alt Satır: Türkçe (VADER Çevirili, BERTurk, LLM)",
        "2. Üye (6:15–7:15)",
    )
    if FIG_CONFUSION_MATRICES.exists():
        s8.shapes.add_picture(str(FIG_CONFUSION_MATRICES), Inches(0.8), Inches(1.3), width=Inches(11.7))
    add_speaker_notes(
        s8,
        "[2. ÜYE - 6:15 - 7:15]\n"
        "Bu slaytta her iki dil ve üç yöntem için ürettiğimiz 6 adet Confusion Matrix'i bir arada görüyoruz. Soldan sağa (VADER'dan BERT'e "
        "ve LLM'e) geçildikçe matrisin ana köşegenindeki (diagonal) doğru tahmin sayılarının nasıl koyulaşıp arttığını ve köşegen dışındaki "
        "hataların nasıl azaldığını net biçimde gözlemleyebiliyoruz."
    )

    # ------------------------------------------------------------------------
    # SLAYT 9: SINIF BAZLI HATALI TAHMİN SAYILARI GRAFİĞİ
    # ------------------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s9,
        "8. Sınıf Bazlı Hatalı Tahmin Sayılarının Karşılaştırması",
        "Pozitif (3.000), Nötr (3.000) ve Negatif (3.000) sınıflarında yapılan yanlış sınıflandırma adetleri",
        "2. Üye (7:15–8:00)",
    )
    if FIG_CLASS_ERRORS.exists():
        s9.shapes.add_picture(str(FIG_CLASS_ERRORS), Inches(0.5), Inches(1.35), width=Inches(12.3))
    add_Bullet_note = (
        "[2. ÜYE - 7:15 - 8:00]\n"
        "Ödev yönergesinde özellikle istenen sınıf bazlı hatalı tahmin sayılarını bu grafikte özetledik. Kırmızı sütunlar VADER'ı, "
        "turuncu sütunlar BERT'i, yeşil sütunlar ise LLM'i gösteriyor. En dikkat çekici bulgu Nötr sınıfındadır: VADER Nötr sınıfında "
        "2.921 ve 2.672 hata yaparken, BERT ve LLM bu hatayı üçte bir oranına kadar indirmiştir. Türkçe Nötr sınıfında en az hatayı "
        "1.224 hata ile EVREN LLM modelimiz yapmıştır."
    )
    add_speaker_notes(s9, add_Bullet_note)

    # ------------------------------------------------------------------------
    # SLAYT 10: GENEL DOĞRULUK VE MACRO F1 KARŞILAŞTIRMASI
    # ------------------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s10,
        "9. Genel Doğruluk (Accuracy) ve Macro F1-Skoru Karşılaştırması",
        "Tüm deneylerin bütüncül performans grafiği",
        "2. Üye (8:00–8:45)",
    )
    if FIG_ACCURACY_F1.exists():
        s10.shapes.add_picture(str(FIG_ACCURACY_F1), Inches(1.2), Inches(1.35), width=Inches(10.9))
    add_speaker_notes(
        s10,
        "[2. ÜYE - 8:00 - 8:45]\n"
        "Genel doğruluk ve Macro F1 skorlarına baktığımızda: İngilizcede VADER %51,5 doğruluk ve %39,6 F1 skorunda kalırken, BERT %81,0'e, "
        "LLM ise %83,6'ya ulaşmıştır. Türkçede ise çevirisiz VADER %35,4, çevirili VADER %54,0, BERTurk %72,1 ve EVREN LLM %79,1 doğruluk "
        "elde etmiştir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 11: NİTELİKSEL HATA ANALİZİ (GERÇEK YORUM ÖRNEKLERİ)
    # ------------------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s11,
        "10. Akademik Hata Analizi: Modeller Neden Yanıldı?",
        "Veri setimizden seçilen gerçek yanlış sınıflandırma vakaları (Case Studies)",
        "2. Üye (8:45–9:30)",
    )
    add_bullet_card(
        s11, 0.5, 1.4, 12.33, 2.7,
        "📌 Vaka 1: Karşıtlık Bağlacı ('fakat / but') ve VADER'ın Pozitif Yanlılığı",
        [
            "Mekan: Kahve Dünyası - Büyükada (Gerçek Etiket: Negative [1★])",
            "Yorum: \"Manzarası güzel diye oturmak istedik fakat hem personelin tavrı bakımından hem tuvaletlerin ve masaların temizliği açısından berbat...\"",
            "Tahminler: VADER -> Positive (compound: +0.5719) ❌  |  BERTurk -> Negative ✅  |  LLM -> Negative ✅",
            "Analiz: VADER cümle başındaki 'beautiful/güzel' kelimesine aldanırken, BERTurk ve LLM 'fakat' bağlacından sonraki ana yargıyı yakalamıştır.",
        ],
    )
    add_bullet_card(
        s11, 0.5, 4.3, 12.33, 2.8,
        "📌 Vaka 2: Türkçe 3 Yıldızlı (Nötr) Yorumlarda Eleştirel Ton ve LLM'in Muhakeme Gücü",
        [
            "Mekan: Seyr-i İstanbul Cafe (Gerçek Etiket: Neutral [3★])",
            "Yorum: \"Manzarası güzel fakat hizmetin o kadarda güzel olmadığını ve fiyatlarının yemek lezzetlerine istinaden aşırı pahalı olduğunu söyleyebilirim.\"",
            "Tahminler: VADER -> Positive (+0.8225) ❌  |  BERTurk -> Negative ❌  |  LLM (EVREN) -> Neutral ✅",
            "Analiz: BERTurk 'aşırı pahalı' ifadesine yüksek ağırlık verip Negative tahmin ederken, LLM hem manzara övgüsünü hem fiyat eleştirisini dengeleyip bunun tam bir 3 yıldızlı Nötr yorum olduğunu doğru bilmiştir.",
        ],
    )
    add_speaker_notes(
        s11,
        "[2. ÜYE - 8:45 - 9:30]\n"
        "Modellerin neden hata yaptığını veri setimizdeki gerçek örneklerle incelediğimizde çok çarpıcı iki durum görüyoruz. "
        "İlk örnekte müşteri 1 yıldız verdiği yorumuna 'Manzarası güzel diye oturmak istedik fakat...' diyerek başlıyor. VADER 'güzel' "
        "kelimesi yüzünden buna Pozitif derken BERTurk ve LLM Negatif olduğunu anlıyor. İkinci örnekte ise 3 yıldızlı Nötr bir yorumda "
        "hem VADER hem BERTurk yanılırken, sadece LLM olumlu ve olumsuz yönlerin dengesini kurup doğru Nötr tahminini yapıyor."
    )

    # ------------------------------------------------------------------------
    # SLAYT 12: SONUÇ VE GENEL DEĞERLENDİRME
    # ------------------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s12,
        "11. Genel Sonuç, Karşılaştırma Özeti ve Çıkarımlar",
        "Hız, maliyet, bağlamsal derinlik ve dil desteği açısından 3 yöntemin değerlendirmesi",
        "2. Üye (9:30–10:00)",
    )
    add_styled_slide_table(
        s12, 0.5, 1.4, 12.33, 2.9,
        ["Kriter", "1. VADER (Kural Tabanlı)", "2. BERT & BERTurk (Derin Öğrenme)", "3. LLM - EVREN (Qwen3.8-Flash)"],
        [
            ["İngilizce Doğruluk / Hata", "%51,53  (4.362 Hata)", "%80,96  (1.714 Hata)", "%83,63  (1.473 Hata) 🏆"],
            ["Türkçe Doğruluk / Hata", "%53,99  (4.141 Hata - Çevirili)", "%72,11  (2.510 Hata)", "%79,11  (1.880 Hata) 🏆"],
            ["Çalışma Süresi (9.000 Yorum)", "~1.2 sn (Çeviri hariç)", "~33 sn (Colab L4 GPU)", "~120 sn (8 Paralel Batch API)"],
            ["Nötr Sınıf Başarımı", "Çok Düşük (%2,6 - %10,9)", "Yüksek (EN: %76,3 | TR: %45,1)", "En Dengeli (EN: %69,1 | TR: %59,2)"],
            ["Bağlam, İroni ve Çok Dillilik", "Yok (Sadece İngilizce Sözlük)", "Yüksek (Dile Özgü Fine-Tuning Şart)", "Çok Yüksek (Zero-Shot Çok Dilli)"],
        ],
        font_size=12,
    )
    add_bullet_card(
        s12, 0.5, 4.55, 12.33, 2.5,
        "🎓 Sonuç ve Teşekkür",
        [
            "Kural tabanlı VADER yalnızca hız gerektiren ve Nötr sınıfın önemsiz olduğu basit İngilizce görevlerde kullanılabilir.",
            "Alan içi (in-domain) 3 sınıflı fine-tuned BERT modelleri, GPU üzerinde saniyeler içinde %80+ doğruluk sunarak mükemmel bir hız/performans dengesi sağlar.",
            "Güncel Büyük Dil Modelleri (EVREN LLM), hiçbir eğitim gerektirmeden (zero-shot) özellikle Türkçe gibi sondan eklemeli ve bağlamı zengin dillerde en yüksek duygu analizi başarımını (%79,11 TR / %83,63 EN) sunmaktadır. Dinlediğiniz için teşekkür ederiz!",
        ],
    )
    add_speaker_notes(
        s12,
        "[2. ÜYE - 9:30 - 10:00]\n"
        "Sonuç olarak; 18.000 yorum üzerindeki deneylerimiz, kural tabanlı VADER'dan derin öğrenme tabanlı BERT'e ve oradan Büyük Dil "
        "Modellerine geçildikçe toplam hata sayısının İngilizcede 4.362'den 1.473'e, Türkçede ise 5.817'den 1.880'e düştüğünü "
        "kanıtlamıştır. Bizi dinlediğiniz için teşekkür ederiz."
    )

    prs.save(OUTPUT_PPTX_PATH)
    print(f"Sunum dosyası başarıyla oluşturuldu: {OUTPUT_PPTX_PATH.name} (12 Slayt)")


if __name__ == "__main__":
    build_presentation()
