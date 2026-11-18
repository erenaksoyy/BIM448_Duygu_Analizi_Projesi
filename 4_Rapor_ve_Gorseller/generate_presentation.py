"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: Akademik Sunum Dosyası (.pptx) Üreticisi (15 Slayt)

Bu modül, proje grubunun sunumda kullanacağı 15 slaytlık geniş ekran (16:9)
PowerPoint sunumunu ('BIM448_Proje_Sunumu.pptx'); VADER, BERT/BERTurk ve LLM
kavramlarının kuramsal açıklamaları, yüksek çözünürlüklü grafikler,
karşılaştırma tabloları ve konuşmacı notlarıyla (Speaker Notes) birlikte üretir.
"""

from pathlib import Path
from typing import Sequence

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
    slide: object, title_text: str, subtitle_text: str, section_badge: str
) -> None:
    """Standart üst başlık şeridi, alt başlık ve bölüm rozetini slayta ekler."""
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
    p_title.font.size = Pt(23)
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
    p_badge.text = section_badge
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
    font_size: int = 13.5,
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
        Inches(left + 0.25), Inches(top + 0.16), Inches(width - 0.5), Inches(height - 0.3)
    )
    tf = tb.text_frame
    tf.word_wrap = True

    p_head = tf.paragraphs[0]
    p_head.text = card_title
    p_head.font.size = Pt(17)
    p_head.font.bold = True
    p_head.font.color.rgb = COLOR_TABLE_HEADER
    p_head.space_after = Pt(8)

    for bullet_text in bullets:
        p = tf.add_paragraph()
        p.text = f"•  {bullet_text}"
        p.font.size = Pt(font_size)
        p.font.color.rgb = COLOR_TEXT_DARK
        p.space_after = Pt(7)


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
# ANA SUNUM OLUŞTURUCU (15 SLAYT)
# ============================================================================

def build_presentation() -> None:
    """15 slaytlık proje sunumunu ve konuşmacı notlarını oluşturur."""
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

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(0.85), Inches(11.3), Inches(5.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_uni = tf1.paragraphs[0]
    p_uni.text = "DOĞUŞ ÜNİVERSİTESİ • YAZILIM MÜHENDİSLİĞİ BÖLÜMÜ\nVERİ MADENCİLİĞİ VE BİLGİ KEŞFİ DERSİ PROJE SUNUMU"
    p_uni.font.size = Pt(16)
    p_uni.font.bold = True
    p_uni.font.color.rgb = RGBColor(147, 197, 253)
    p_uni.space_after = Pt(20)

    p_main = tf1.add_paragraph()
    p_main.text = "Karşılaştırmalı Duygu Analizi (Sentiment Analysis):\nİngilizce ve Türkçe Mekan Yorumları Üzerinde\nVADER, BERT ve LLM Yöntemlerinin İncelenmesi"
    p_main.font.size = Pt(32)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_WHITE
    p_main.space_after = Pt(22)

    p_authors_title = tf1.add_paragraph()
    p_authors_title.text = "Hazırlayanlar:"
    p_authors_title.font.size = Pt(18)
    p_authors_title.font.bold = True
    p_authors_title.font.color.rgb = RGBColor(96, 165, 250)
    p_authors_title.space_after = Pt(6)

    p_authors = tf1.add_paragraph()
    p_authors.text = (
        "•  Eren Aksoy — 202403011009\n"
        "•  Emre Kayıkcı — 202403011068"
    )
    p_authors.font.size = Pt(17)
    p_authors.font.bold = True
    p_authors.font.color.rgb = COLOR_WHITE
    p_authors.space_after = Pt(20)

    p_meta = tf1.add_paragraph()
    p_meta.text = (
        "• Veri Ölçeği: 9.000 İngilizce + 9.000 Türkçe = 18.000 Gerçek Restoran & Kafe Yorumu\n"
        "• İncelenen Yöntemler: 1) VADER (Kural Tabanlı)   2) BERT & BERTurk (Derin Öğrenme)   3) EVREN LLM (Qwen3.8-Flash-Next)"
    )
    p_meta.font.size = Pt(15)
    p_meta.font.color.rgb = RGBColor(203, 213, 225)

    add_speaker_notes(
        s1,
        "Merhaba Hocam, BİM448 Veri Madenciliği ve Bilgi Keşfi dersi kapsamında hazırladığımız "
        "Karşılaştırmalı Duygu Analizi projemizin sunumuna hoş geldiniz. Bu çalışmada geleneksel kural tabanlı VADER, "
        "derin öğrenme tabanlı BERT/BERTurk ve Büyük Dil Modeli tabanlı EVREN LLM yöntemlerini hem İngilizce hem de Türkçe "
        "gerçek restoran ve kafe yorumları üzerinde karşılaştırdık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 2: PROJENİN AMACI, KAPSAMI VE MİMARİ AKIŞI
    # ------------------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s2,
        "1. Projenin Amacı, Kapsamı ve Yöntemsel Çerçeve",
        "Üç farklı kuşak doğal dil işleme yaklaşımının iki dil ve üç duygu sınıfı üzerinde kıyaslanması",
        "Giriş ve Kapsam",
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
        "⚙️ Uygulanan 3 Kuşak NLP Yaklaşımı",
        [
            "1. Kuşak – Kural ve Sözlük Tabanlı (VADER): Kelime kutupluluk sözlüğü ve dilbilimsel kurallarla bileşik skor üretir.",
            "2. Kuşak – Derin Öğrenme / Transformer Encoder (BERT & BERTurk): Çift yönlü öz-dikkat (bidirectional self-attention) ile bağlamsal kelime temsilleri öğrenir.",
            "3. Kuşak – Büyük Dil Modeli / Üretken Yapay Zeka (EVREN LLM): Milyarlarca parametreli ön-eğitim ve sıfır-örnek (zero-shot) istem mühendisliği ile metni bütüncül muhakeme eder.",
        ],
    )
    add_speaker_notes(
        s2,
        "Projemizin temel amacı; kural tabanlı, derin öğrenme tabanlı ve LLM tabanlı duygu analizi yaklaşımlarını İngilizce ve Türkçe "
        "metinler üzerinde karşılaştırmaktır. İstatistiksel güvenilirliği artırmak için minimum 1.000'er yorum sınırıyla yetinmeyip "
        "her sınıfta 3.000'er adet olmak üzere dil başına 9.000, toplamda 18.000 gerçek kullanıcı yorumu kullandık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 3: VERİ SETLERİNİN KAYNAKLARI VE ÖN İŞLEME
    # ------------------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s3,
        "2. İngilizce ve Türkçe Veri Setleri (18.000 Mekan Yorumu)",
        "Aynı konu alanından (Restoran & Kafe) ve birebir aynı şemada dengeli veri seti tasarımı",
        "Veri Setleri",
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
        "İki dil arasındaki başarı farkının konu farkından etkilenmemesi için her iki dilde de Restoran ve Kafe yorumlarını seçtik. "
        "Türkçede 838 bin yorumluk Google Maps veri setinden, İngilizcede ise 2 milyon yorumluk Yelp veri setinden işletme isimlerini ve "
        "kategorilerini eşleştirerek 3.000 Pozitif, 3.000 Nötr ve 3.000 Negatif olmak üzere 9.000'er yorumluk iki veri seti hazırladık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 4 (YENİ): VADER NEDİR VE NASIL ÇALIŞIR?
    # ------------------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s4,
        "3. Yöntem 1 Kuramsal Altyapı: VADER Nedir ve Nasıl Çalışır?",
        "Valence Aware Dictionary and sEntiment Reasoner (Hutto & Gilbert, 2014)",
        "Yöntem 1: VADER",
    )
    add_bullet_card(
        s4, 0.5, 1.45, 6.0, 5.5,
        "📖 VADER Nedir? (Sözlük ve Kural Yapısı)",
        [
            "Sosyal medya ve kısa kullanıcı yorumları için geliştirilmiş, eğitim gerektirmeyen sözlük (lexicon) ve kural tabanlı bir duygu analizi modelidir.",
            "Altın Standart Duygu Sözlüğü: İnsan uzmanlarca etiketlenmiş 7.500+ İngilizce kelime, argo ('lol', 'meh') ve emoji/emoticon (':)', '🙂') içerir.",
            "Her ifadenin -4.0 (aşırı olumsuz) ile +4.0 (aşırı olumlu) arasında bir duygu şiddeti (valence) puanı vardır (Örn: 'great': +3.1, 'horrible': -2.5).",
            "Bileşik Skor (Compound): Kelime puanlarının toplamı x olmak üzere x / sqrt(x² + 15) formülüyle -1.0 ile +1.0 arasına normalize edilir.",
        ],
    )
    add_bullet_card(
        s4, 6.8, 1.45, 6.0, 5.5,
        "⚙️ VADER'ın 5 Dilbilimsel Sezgisel Kuralı",
        [
            "1. Noktalama İşaretleri ('!'): Ünlem işareti duygu yönünü değiştirmeden şiddetini artırır ('Good!!!' > 'Good').",
            "2. Büyük Harf Vurgusu (ALL-CAPS): Tamamı büyük harfle yazılan kelimelerin ('AMAZING', 'AWFUL') duygu şiddeti yükseltilir.",
            "3. Derece Zarfları (Booster Words): 'very', 'extremely' duygu şiddetini artırırken; 'slightly', 'kind of' düşürür.",
            "4. Karşıtlık Bağlacı ('but'): 'but' kelimesinden önceki ifadelerin ağırlığı %50 azaltılır, sonrasındakilerin %50 artırılır.",
            "5. Olumsuzluk (Negation): Kelimeden önceki 3 kelimelik pencerede 'not', 'never' varsa skor -0.74 ile çarpılarak ters çevrilir.",
        ],
    )
    add_speaker_notes(
        s4,
        "İlk yöntemimiz olan VADER, sosyal medya ve müşteri yorumları için tasarlanmış sözlük ve kural tabanlı bir modeldir. "
        "İçerisinde 7.500'den fazla İngilizce kelime ve emojinin -4 ile +4 arasında duygu puanı yer alır. VADER bu puanları toplarken "
        "ünlem işaretlerini, büyük harf vurgusunu, 'very' gibi derece zarflarını ve 'not' gibi olumsuzluk eklerini hesaba katarak "
        "-1 ile +1 arasında bir compound skor üretir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 5: 1. YÖNTEM - VADER UYGULAMA SONUÇLARI
    # ------------------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s5,
        "4. Yöntem 1 Uygulama Sonuçları: VADER Sınıf Bazlı Hatalar",
        "Eşik Değerleri: compound >= 0.05 (Positive)  |  compound <= -0.05 (Negative)  |  Arası (Neutral)",
        "Yöntem 1: VADER",
    )
    add_styled_slide_table(
        s5, 0.5, 1.4, 12.33, 3.6,
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
        s5, 0.5, 5.2, 12.33, 1.9,
        "🔍 VADER Temel Bulguları: Nötr Sınıf Körlüğü ve Pozitif Yanlılığı",
        [
            "Dar Nötr Aralığı (-0.05 ile +0.05): 3 yıldızlı Nötr yorumlarda geçen tek bir olumlu kelime ('friendly', 'nice') skoru +0.05 üstüne çıkardığı için VADER İngilizcede 2.921, Türkçede 2.672 Nötr yorumu hatalı sınıflandırmıştır.",
            "Bağlam Körlüğü: 'Manzarası güzel fakat servis berbat' gibi yorumlarda başlangıçtaki olumlu kelime yüzünden Negatif yorumların ~%44'ü Pozitif tahmin edilmiştir.",
        ],
    )
    add_speaker_notes(
        s5,
        "VADER'ı 9.000 İngilizce ve İngilizceye çevrilmiş 9.000 Türkçe yorum üzerinde çalıştırdığımızda; İngilizcede %51,53 doğruluk ve "
        "4.362 hatalı tahmin, Türkçede ise %53,99 doğruluk ve 4.141 hatalı tahmin elde ettik. VADER Pozitif sınıfında çok başarılı "
        "görünse de aslında aşırı pozitif yanlısı (positivity bias) davranmaktadır ve Nötr yorumların %89 ila %97'sini kaçırmaktadır."
    )

    # ------------------------------------------------------------------------
    # SLAYT 6: TÜRKÇEDE ÇEVİRİSİZ VS ÇEVİRİLİ VADER ETKİSİ
    # ------------------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s6,
        "5. Türkçe VADER Analizinde Çevirinin Etkisi (Ham vs. Çevirili)",
        "VADER İngilizce sözlük tabanlı olduğu için Türkçe metinlerin çevrilmesi zorunludur",
        "Yöntem 1: VADER",
    )
    if FIG_TR_TRANSLATION.exists():
        s6.shapes.add_picture(str(FIG_TR_TRANSLATION), Inches(0.5), Inches(1.4), width=Inches(7.7))
    add_bullet_card(
        s6, 8.4, 1.4, 4.4, 5.6,
        "📌 Deneysel Kanıt",
        [
            "Çevirisiz Ham Türkçe VADER: 9.000 yorumun 8.311 tanesine 0.0000 (Neutral) skoru vermiştir (Doğruluk: %35,37 | Toplam Hata: 5.817).",
            "Sadece emoji 🙂 ve 'super/ok' gibi kelimelerden 331 Pozitif yakalayabilmiştir.",
            "Google Translate ile Çeviri Sonrası: Doğruluk %35,37'den %53,99'a yükselmiş, Pozitif sınıf doğru tahmin sayısı 331'den 2.823'e, Negatif sınıf doğru tahmini 34'ten 1.708'e çıkmıştır.",
        ],
    )
    add_speaker_notes(
        s6,
        "VADER sözlüğü yalnızca İngilizce kelimeleri içerdiği için Türkçe metinleri doğrudan verdiğimizde 9.000 yorumun 8.311'ine "
        "0.0000 puan vererek Nötr tahmin etmektedir. Google Translate API ile İngilizceye çevirip VADER uyguladığımızda ise "
        "doğruluk %35,37'den %53,99'a yükselmiştir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 7 (YENİ): BERT VE BERTurk NEDİR VE NASIL ÇALIŞIR?
    # ------------------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s7,
        "6. Yöntem 2 Kuramsal Altyapı: BERT ve BERTurk Nedir?",
        "Bidirectional Encoder Representations from Transformers (Devlin et al., 2019)",
        "Yöntem 2: BERT",
    )
    add_bullet_card(
        s7, 0.5, 1.45, 6.0, 5.5,
        "🧠 BERT Nedir ve Nasıl Çalışır?",
        [
            "Transformer mimarisinin Kodlayıcı (Encoder) katmanlarına dayanan derin öğrenme modelidir (12 katman, 768 gizli boyut, 12 dikkat başlığı, 110M parametre).",
            "Çift Yönlü Öz-Dikkat (Bidirectional Self-Attention): Metni tek yönlü değil, hem soldan sağa hem sağdan sola aynı anda işler; her kelimenin anlamını cümledeki diğer tüm kelimelerle etkileşiminden çıkarır.",
            "Ön-Eğitim Görevleri: Maskelenmiş Dil Modeli (MLM - %15 kelime maskeleme) ve Sonraki Cümle Tahmini (NSP) ile derin dil temsili öğrenir.",
            "Sınıflandırma (Fine-Tuning): Cümle başındaki özel [CLS] tokenının vektörü Softmax katmanından geçirilerek 3 sınıfa olasılık atar.",
        ],
    )
    add_bullet_card(
        s7, 6.8, 1.45, 6.0, 5.5,
        "🇹🇷 BERTurk Nedir? Neden 3 Sınıflı Model Seçtik?",
        [
            "BERTurk (dbmdz/bert-base-turkish-cased): 35 GB (440 milyon kelime) Türkçe metin üzerinde eğitilmiş Türkçe Transformer modelidir.",
            "WordPiece Alt-Kelime Tokenizasyonu: Türkçe sondan eklemeli bir dil olduğu için kelime köklerini ve olumsuzluk eklerini ('-me', '-ma', '-sız') parçalarına ayırarak doğru temsil eder.",
            "Kritik Tasarım Kararı: Popüler SST-2 ve savasy/bert-base-turkish modelleri yalnızca 2 sınıflıdır (Pozitif/Negatif) ve Nötr tahmin edemez.",
            "Bu nedenle projemizde doğrudan 3 sınıflı (Positive, Neutral, Negative) olarak fine-tune edilmiş BERT ve BERTurk modelleri kullanılmıştır.",
        ],
    )
    add_speaker_notes(
        s7,
        "İkinci yöntemimiz olan BERT, Google tarafından geliştirilen ve Transformer Encoder mimarisine dayanan bir derin öğrenme modelidir. "
        "VADER'ın aksine kelimelere sabit puan vermez; Çift Yönlü Öz-Dikkat mekanizmasıyla bir kelimenin hem solundaki hem sağındaki "
        "tüm kelimelerle bağlamsal ilişkisini hesaplar. Türkçe için ise 35 GB Türkçe derlem üzerinde eğitilen ve Türkçe ekleri WordPiece "
        "ile çözümleyen BERTurk mimarisini kullandık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 8: 2. YÖNTEM - BERT & BERTurk UYGULAMA SONUÇLARI
    # ------------------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s8,
        "7. Yöntem 2 Uygulama Sonuçları: BERT ve BERTurk Hata Analizi",
        "Google Colab NVIDIA L4 GPU üzerinde 3 sınıflı fine-tuned Transformer modelleri",
        "Yöntem 2: BERT",
    )
    add_styled_slide_table(
        s8, 0.5, 1.4, 12.33, 3.6,
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
        s8, 0.5, 5.2, 12.33, 1.9,
        "🚀 Derin Öğrenme Kazanımları ve Sınıf Davranışı",
        [
            "Nötr Sınıf Devrimi: İngilizce Nötr sınıfında VADER 2.921 hata yaparken BERT sadece 712 hata yapmıştır (2.288 doğru tahmin).",
            "Zıt Kutupların Ayrışması: İngilizce BERT'te 3.000 Negatif yorumdan yalnızca 13'ü Pozitif'e kaymıştır (VADER'da 1.247 idi!). Türkçede ise Negatif sınıf hata oranı %7,63'e (229 hata) inmiştir.",
        ],
    )
    add_speaker_notes(
        s8,
        "BERT ve BERTurk modellerini Colab L4 GPU üzerinde 9.000'er yorumda çalıştırdığımızda; İngilizce doğruluk %51,53'ten %80,96'ya "
        "fırlamış ve toplam hata sayısı 4.362'den 1.714'e düşmüştür. Türkçede ise BERTurk %72,11 doğruluk elde etmiş ve Negatif "
        "sınıfta 3.000 yorumun 2.771'ini doğru bilerek sadece 229 hata yapmıştır."
    )

    # ------------------------------------------------------------------------
    # SLAYT 9 (YENİ): BÜYÜK DİL MODELİ (LLM) NEDİR VE NASIL ÇALIŞIR?
    # ------------------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s9,
        "8. Yöntem 3 Kuramsal Altyapı: Büyük Dil Modeli (LLM) Nedir?",
        "Üretken Transformer Decoder Mimarisi ve İstem Mühendisliği (Prompt Engineering)",
        "Yöntem 3: LLM",
    )
    add_bullet_card(
        s9, 0.5, 1.45, 6.0, 5.5,
        "🤖 Büyük Dil Modeli (LLM) Nedir?",
        [
            "Milyarlarca parametreye sahip, Transformer Çözücü (Decoder-only Autoregressive) mimarisine dayanan üretken yapay zekâ modelleridir (GPT-4, Qwen, Gemma, Llama, DeepSeek).",
            "Trilyonlarca token üzerinde 'sonraki kelime tahmini' (Next-Token Prediction) ile eğitilerek yalnızca dilbilgisini değil; dünya bilgisini, deyimleri, ironiyi ve mantıksal muhakemeyi öğrenir.",
            "Instruction Tuning & RLHF: İnsan yönergelerini takip etmek üzere hizalanmıştır; bu sayede ek eğitim (fine-tuning) gerektirmeden Sıfır-Örnek (Zero-Shot) görev çözebilir.",
        ],
    )
    add_bullet_card(
        s9, 6.8, 1.45, 6.0, 5.5,
        "⚡ Duygu Analizinde LLM & EVREN Platformu",
        [
            "VADER ve BERT'ten Farkı: Kelime puanlarına veya yerel kalıplara takılmaz; cümlenin tamamındaki alaycılığı (sarcasm), devrik Türkçe yapıları ve 'fiyatına göre idare eder' gibi gri (Nötr) yargıları bütüncül muhakeme eder.",
            "EVREN Milli Yapay Zeka Platformu: T.C. Cumhurbaşkanlığı SSB & SAYZEK koordinesinde Türkiye'deki NVIDIA H200 GPU kümeleri üzerinde çalışan OpenAI uyumlu LLM altyapısıdır.",
            "Projemizde EVREN üzerindeki 'qwen3.8-flash-next' modeli, temperature=0.0 ve yapılandırılmış JSON istem mühendisliği ile kullanılmıştır.",
        ],
    )
    add_speaker_notes(
        s9,
        "Üçüncü yöntemimiz olan Büyük Dil Modelleri, milyarlarca parametreli Transformer Decoder mimarisine dayanır. BERT'ten farklı olarak "
        "herhangi bir etiketli veriyle yeniden eğitilmeye (fine-tuning) ihtiyaç duymaz; Zero-Shot Prompt Engineering yani İstem Mühendisliği "
        "ile kendisine verilen görevi doğrudan anlar ve cümledeki ironiyi, bağlamı ve gri alanları muhakeme eder."
    )

    # ------------------------------------------------------------------------
    # SLAYT 10: 3. YÖNTEM - EVREN LLM UYGULAMA SONUÇLARI
    # ------------------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s10,
        "9. Yöntem 3 Uygulama Sonuçları: EVREN LLM Hata Analizi",
        "EVREN API (qwen3.8-flash-next, temperature=0.0, Yapılandırılmış Batch Prompting)",
        "Yöntem 3: LLM",
    )
    add_styled_slide_table(
        s10, 0.5, 1.4, 12.33, 3.6,
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
        s10, 0.5, 5.2, 12.33, 1.9,
        "🏆 LLM Yaklaşımının Öne Çıkan Bulguları",
        [
            "Genel Liderlik: Hiçbir fine-tuning yapılmadan hem İngilizcede (%83,63 doğruluk, 1.473 hata) hem Türkçede (%79,11 doğruluk, 1.880 hata) 1. sıraya yerleşmiştir.",
            "Sıfıra Yakın Kutup Hatası: İngilizcede 3.000 Pozitif yorumdan yalnızca 5'i Negatif'e, 3.000 Negatif yorumdan yalnızca 5'i Pozitif'e kaymıştır.",
        ],
    )
    add_speaker_notes(
        s10,
        "EVREN API üzerindeki qwen3.8-flash-next modelimizle 18.000 yorumu analiz ettiğimizde; İngilizcede %83,63 doğruluk ve sadece "
        "1.473 toplam hata, Türkçede ise %79,11 doğruluk ve 1.880 toplam hata ile her iki dilde de en yüksek başarıya ulaştık."
    )

    # ------------------------------------------------------------------------
    # SLAYT 11: 6'LI CONFUSION MATRIX KARŞILAŞTIRMASI
    # ------------------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s11,
        "10. Altılı Karmaşıklık Matrisi (Confusion Matrices) Karşılaştırması",
        "Üst Satır: İngilizce (VADER, BERT, LLM)  •  Alt Satır: Türkçe (VADER Çevirili, BERTurk, LLM)",
        "Karşılaştırma",
    )
    if FIG_CONFUSION_MATRICES.exists():
        s11.shapes.add_picture(str(FIG_CONFUSION_MATRICES), Inches(0.8), Inches(1.3), width=Inches(11.7))
    add_speaker_notes(
        s11,
        "Bu slaytta her iki dil ve üç yöntem için ürettiğimiz 6 adet Confusion Matrix'i bir arada görüyoruz. Soldan sağa (VADER'dan BERT'e "
        "ve LLM'e) geçildikçe matrisin ana köşegenindeki doğru tahmin sayılarının nasıl arttığını ve köşegen dışındaki hataların "
        "nasıl azaldığını net biçimde gözlemleyebiliyoruz."
    )

    # ------------------------------------------------------------------------
    # SLAYT 12: SINIF BAZLI HATALI TAHMİN SAYILARI GRAFİĞİ
    # ------------------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s12,
        "11. Sınıf Bazlı Hatalı Tahmin Sayılarının Karşılaştırması",
        "Pozitif (3.000), Nötr (3.000) ve Negatif (3.000) sınıflarında yapılan yanlış sınıflandırma adetleri",
        "Hata Analizi",
    )
    if FIG_CLASS_ERRORS.exists():
        s12.shapes.add_picture(str(FIG_CLASS_ERRORS), Inches(0.5), Inches(1.35), width=Inches(12.3))
    add_speaker_notes(
        s12,
        "Ödev yönergesinde özellikle istenen sınıf bazlı hatalı tahmin sayılarını bu grafikte özetledik. Kırmızı sütunlar VADER'ı, "
        "turuncu sütunlar BERT'i, yeşil sütunlar ise LLM'i gösteriyor. VADER Nötr sınıfında 2.921 ve 2.672 hata yaparken, "
        "BERT ve LLM bu hatayı üçte bir oranına kadar indirmiştir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 13: GENEL DOĞRULUK VE MACRO F1 KARŞILAŞTIRMASI
    # ------------------------------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s13,
        "12. Genel Doğruluk (Accuracy) ve Macro F1-Skoru Karşılaştırması",
        "Tüm deneylerin bütüncül performans grafiği",
        "Genel Başarım",
    )
    if FIG_ACCURACY_F1.exists():
        s13.shapes.add_picture(str(FIG_ACCURACY_F1), Inches(1.2), Inches(1.35), width=Inches(10.9))
    add_speaker_notes(
        s13,
        "Genel doğruluk ve Macro F1 skorlarına baktığımızda: İngilizcede VADER %51,5 doğruluk ve %39,6 F1 skorunda kalırken, BERT %81,0'e, "
        "LLM ise %83,6'ya ulaşmıştır. Türkçede ise çevirisiz VADER %35,4, çevirili VADER %54,0, BERTurk %72,1 ve EVREN LLM %79,1 doğruluk "
        "elde etmiştir."
    )

    # ------------------------------------------------------------------------
    # SLAYT 14: NİTELİKSEL HATA ANALİZİ (GERÇEK YORUM ÖRNEKLERİ)
    # ------------------------------------------------------------------------
    s14 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s14,
        "13. Akademik Hata Analizi: Modeller Neden Yanıldı?",
        "Veri setimizden seçilen gerçek yanlış sınıflandırma vakaları (Case Studies)",
        "Niteliksel Analiz",
    )
    add_bullet_card(
        s14, 0.5, 1.4, 12.33, 2.7,
        "📌 Vaka 1: Karşıtlık Bağlacı ('fakat / but') ve VADER'ın Pozitif Yanlılığı",
        [
            "Mekan: Kahve Dünyası - Büyükada (Gerçek Etiket: Negative [1★])",
            "Yorum: \"Manzarası güzel diye oturmak istedik fakat hem personelin tavrı bakımından hem tuvaletlerin ve masaların temizliği açısından berbat...\"",
            "Tahminler: VADER -> Positive (compound: +0.5719) ❌  |  BERTurk -> Negative ✅  |  LLM -> Negative ✅",
            "Analiz: VADER cümle başındaki 'beautiful/güzel' kelimesine aldanırken, BERTurk ve LLM 'fakat' bağlacından sonraki ana yargıyı yakalamıştır.",
        ],
    )
    add_bullet_card(
        s14, 0.5, 4.3, 12.33, 2.8,
        "📌 Vaka 2: Türkçe 3 Yıldızlı (Nötr) Yorumlarda Eleştirel Ton ve LLM'in Muhakeme Gücü",
        [
            "Mekan: Seyr-i İstanbul Cafe (Gerçek Etiket: Neutral [3★])",
            "Yorum: \"Manzarası güzel fakat hizmetin o kadarda güzel olmadığını ve fiyatlarının yemek lezzetlerine istinaden aşırı pahalı olduğunu söyleyebilirim.\"",
            "Tahminler: VADER -> Positive (+0.8225) ❌  |  BERTurk -> Negative ❌  |  LLM (EVREN) -> Neutral ✅",
            "Analiz: BERTurk 'aşırı pahalı' ifadesine yüksek ağırlık verip Negative tahmin ederken, LLM hem manzara övgüsünü hem fiyat eleştirisini dengeleyip bunun tam bir 3 yıldızlı Nötr yorum olduğunu doğru bilmiştir.",
        ],
    )
    add_speaker_notes(
        s14,
        "Modellerin neden hata yaptığını veri setimizdeki gerçek örneklerle incelediğimizde çok çarpıcı iki durum görüyoruz. "
        "İlk örnekte müşteri 1 yıldız verdiği yorumuna 'Manzarası güzel diye oturmak istedik fakat...' diyerek başlıyor. VADER 'güzel' "
        "kelimesi yüzünden buna Pozitif derken BERTurk ve LLM Negatif olduğunu anlıyor. İkinci örnekte ise 3 yıldızlı Nötr bir yorumda "
        "hem VADER hem BERTurk yanılırken, sadece LLM olumlu ve olumsuz yönlerin dengesini kurup doğru Nötr tahminini yapıyor."
    )

    # ------------------------------------------------------------------------
    # SLAYT 15: SONUÇ VE GENEL DEĞERLENDİRME
    # ------------------------------------------------------------------------
    s15 = prs.slides.add_slide(blank_layout)
    apply_slide_header(
        s15,
        "14. Genel Sonuç, Karşılaştırma Özeti ve Çıkarımlar",
        "Hız, maliyet, bağlamsal derinlik ve dil desteği açısından 3 yöntemin değerlendirmesi",
        "Sonuç",
    )
    add_styled_slide_table(
        s15, 0.5, 1.4, 12.33, 2.9,
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
        s15, 0.5, 4.55, 12.33, 2.5,
        "🎓 Sonuç ve Teşekkür",
        [
            "Kural tabanlı VADER yalnızca hız gerektiren ve Nötr sınıfın önemsiz olduğu basit İngilizce görevlerde kullanılabilir.",
            "Alan içi (in-domain) 3 sınıflı fine-tuned BERT modelleri, GPU üzerinde saniyeler içinde %80+ doğruluk sunarak mükemmel bir hız/performans dengesi sağlar.",
            "Güncel Büyük Dil Modelleri (EVREN LLM), hiçbir eğitim gerektirmeden (zero-shot) özellikle Türkçe gibi sondan eklemeli ve bağlamı zengin dillerde en yüksek duygu analizi başarımını (%79,11 TR / %83,63 EN) sunmaktadır. Dinlediğiniz için teşekkür ederiz!",
        ],
    )
    add_speaker_notes(
        s15,
        "Sonuç olarak; 18.000 yorum üzerindeki deneylerimiz, kural tabanlı VADER'dan derin öğrenme tabanlı BERT'e ve oradan Büyük Dil "
        "Modellerine geçildikçe toplam hata sayısının İngilizcede 4.362'den 1.473'e, Türkçede ise 5.817'den 1.880'e düştüğünü "
        "kanıtlamıştır. Bizi dinlediğiniz için teşekkür ederiz."
    )

    prs.save(OUTPUT_PPTX_PATH)
    print(f"Sunum dosyası başarıyla oluşturuldu: {OUTPUT_PPTX_PATH.name} (15 Slayt)")


if __name__ == "__main__":
    build_presentation()
