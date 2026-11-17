"""
BİM448 Veri Madenciliği ve Bilgi Keşfi - Proje Ödevi
Modül: Veri Ön İşleme ve Dengeli Örneklem Oluşturma (Data Preprocessing)

Bu modül, Hugging Face üzerindeki Türkçe (Google Maps) ve İngilizce (Yelp)
restoran/kafe yorum veri setlerini indirir, gürültülü verileri temizler ve
her sınıftan (Positive, Neutral, Negative) eşit sayıda (3.000'er adet)
örneklem seçerek standart şemada CSV dosyaları üretir.
"""

import csv
import io
import random
import re
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set, Tuple

import pandas as pd
import pyarrow.parquet as pq


# ============================================================================
# SABİTLER VE KONFİGÜRASYON (CONSTANTS & CONFIGURATION)
# ============================================================================

BASE_DIR: Path = Path(__file__).resolve().parent
RANDOM_SEED: int = 42
SAMPLES_PER_CLASS: int = 3000
MIN_CHAR_LENGTH: int = 25
MAX_CHAR_LENGTH: int = 500
REQUEST_TIMEOUT_SEC: int = 60

TURKISH_DATASET_URL: str = (
    "https://huggingface.co/datasets/opdullah/turkish-google-maps-reviews/resolve/main/train.csv"
)
ENGLISH_BUSINESS_URL: str = (
    "https://huggingface.co/datasets/jaimik69/Yelp-Restaurant-Dataset/resolve/"
    "refs%2Fconvert%2Fparquet/default/train/0000.parquet"
)
ENGLISH_REVIEWS_URL: str = (
    "https://huggingface.co/datasets/hasnat79/yelp_food_business_review_dataset_2M/resolve/"
    "refs%2Fconvert%2Fparquet/default/test/0000.parquet"
)

TURKISH_OUTPUT_CSV: Path = BASE_DIR / "turkish_restaurant_reviews_9000.csv"
ENGLISH_OUTPUT_CSV: Path = BASE_DIR / "english_restaurant_reviews_9000.csv"

SENTIMENT_CLASSES: Tuple[str, ...] = ("Positive", "Negative", "Neutral")

TURKISH_CHARS: Set[str] = set("çğıöşüÇĞİÖŞÜ")
TURKISH_COMMON_WORDS: Set[str] = {
    "ve", "bir", "bu", "çok", "için", "ama", "de", "da", "ile", "gibi",
    "daha", "en", "güzel", "iyi", "kötü", "lezzetli", "servis", "mekan",
    "fiyat", "yemek", "kahve", "çay", "garson", "kesinlikle", "tavsiye",
    "olarak", "kadar", "sonra", "hiç", "var", "yok", "gayet", "ortalama",
}

INCLUDED_CATEGORY_KEYWORDS: Tuple[str, ...] = (
    "restoran", "kafe", "kahve", "cafe", "restaurant", "coffee",
    "pastane", "bistro", "lokanta", "kebap", "pide", "pizza", "burger", "tatlı",
)
EXCLUDED_CATEGORY_KEYWORDS: Tuple[str, ...] = (
    "hastane", "klinik", "doktor", "diş", "otel", "okul", "üniversite", "banka", "eczane",
)


@dataclass(frozen=True)
class ReviewRecord:
    """Standartlaştırılmış tekil restoran/kafe yorum kaydı."""

    place_name: str
    place_category: str
    rating: int
    text: str
    true_label: str


# ============================================================================
# YARDIMCI FONKSİYONLAR (HELPER FUNCTIONS)
# ============================================================================

def normalize_whitespace(text: str) -> str:
    """Metindeki fazla boşlukları ve satır sonlarını tek boşluğa indirger."""
    return re.sub(r"\s+", " ", (text or "").strip())


def is_valid_length(text: str) -> bool:
    """Metin uzunluğunun belirlenen karakter sınırları içinde olup olmadığını denetler."""
    return MIN_CHAR_LENGTH <= len(text) <= MAX_CHAR_LENGTH


def is_natural_turkish_text(text: str) -> bool:
    """
    Metnin doğal Türkçe karakter ve kelime içerdiğini, Arapça/Kiril gibi
    farklı alfabeler barındırmadığını doğrular.
    """
    if not is_valid_length(text):
        return False
    if re.search(r"[\u0400-\u04FF\u0600-\u06FF]", text):
        return False

    words = set(re.findall(r"[a-zA-ZçğıöşüÇĞİÖŞÜ]+", text.lower()))
    has_turkish_char = any(char in TURKISH_CHARS for char in text)
    has_common_word = bool(words.intersection(TURKISH_COMMON_WORDS))
    return has_turkish_char and has_common_word


def is_food_and_beverage_venue(category: str, category_types: str) -> bool:
    """İşletmenin restoran, kafe veya yeme-içme mekanı olup olmadığını filtreler."""
    cat_lower = (category or "").lower()
    types_lower = (category_types or "").lower()

    if any(excluded in cat_lower for excluded in EXCLUDED_CATEGORY_KEYWORDS):
        return False
    return any(keyword in cat_lower or keyword in types_lower for keyword in INCLUDED_CATEGORY_KEYWORDS)


def map_star_rating_to_sentiment(rating: int) -> str:
    """
    1-5 arası yıldız puanını 3 sınıflı duygu etiketine dönüştürür:
    1-2 Yıldız -> Negative, 3 Yıldız -> Neutral, 4-5 Yıldız -> Positive.
    """
    if rating in (1, 2):
        return "Negative"
    if rating == 3:
        return "Neutral"
    if rating in (4, 5):
        return "Positive"
    raise ValueError(f"Geçersiz yıldız puanı: {rating}")


def sample_balanced_dataframe(buckets: Dict[str, List[ReviewRecord]], seed: int) -> pd.DataFrame:
    """
    Her duygu sınıfından eşit sayıda (SAMPLES_PER_CLASS) rastgele örneklem seçer,
    karıştırır ve standart sütun yapısında DataFrame döndürür.
    """
    rng = random.Random(seed)
    sampled_records: List[Dict[str, object]] = []

    for label in SENTIMENT_CLASSES:
        pool = buckets[label]
        if len(pool) < SAMPLES_PER_CLASS:
            raise ValueError(
                f"[{label}] sınıfında yeterli veri yok: {len(pool)} < {SAMPLES_PER_CLASS}"
            )
        for record in rng.sample(pool, SAMPLES_PER_CLASS):
            sampled_records.append({
                "place_name": record.place_name,
                "place_category": record.place_category,
                "rating": record.rating,
                "text": record.text,
                "true_label": record.true_label,
            })

    rng.shuffle(sampled_records)
    df = pd.DataFrame(sampled_records)
    df.insert(0, "id", range(1, len(df) + 1))
    return df


# ============================================================================
# ANA VERİ HAZIRLAMA FONKSİYONLARI (CORE PIPELINES)
# ============================================================================

def prepare_turkish_dataset() -> pd.DataFrame:
    """Türkçe Google Maps Restoran & Kafe veri setini indirir, filtreler ve kaydeder."""
    print("1) Türkçe Google Maps Restoran & Kafe veri seti hazırlanıyor...")
    start_time = time.time()

    buckets: Dict[str, List[ReviewRecord]] = {label: [] for label in SENTIMENT_CLASSES}
    seen_texts: Set[str] = set()

    with urllib.request.urlopen(TURKISH_DATASET_URL, timeout=REQUEST_TIMEOUT_SEC) as response:
        text_stream = io.TextIOWrapper(response, encoding="utf-8", errors="replace")
        reader = csv.DictReader(text_stream)

        for row in reader:
            category = row.get("place_category", "")
            category_types = row.get("place_category_types", "")
            if not is_food_and_beverage_venue(category, category_types):
                continue

            text = normalize_whitespace(row.get("review_text", ""))
            if not is_natural_turkish_text(text):
                continue

            try:
                rating = int(float(row.get("review_rating", 0)))
                label = map_star_rating_to_sentiment(rating)
            except ValueError:
                continue

            text_key = text.lower()
            if text_key in seen_texts:
                continue
            seen_texts.add(text_key)

            buckets[label].append(
                ReviewRecord(
                    place_name=row.get("place_name", "").strip(),
                    place_category=category.strip(),
                    rating=rating,
                    text=text,
                    true_label=label,
                )
            )

    df_turkish = sample_balanced_dataframe(buckets, RANDOM_SEED)
    df_turkish.to_csv(TURKISH_OUTPUT_CSV, index=False, encoding="utf-8")

    elapsed = time.time() - start_time
    print(f"   -> Tamamlandı ({elapsed:.1f} sn): {TURKISH_OUTPUT_CSV.name} ({len(df_turkish):,} satır)")
    return df_turkish


def load_english_business_metadata() -> Dict[str, Tuple[str, str]]:
    """Yelp işletme tablosunu okuyarak business_id -> (mekan_adı, kategori) eşlemesini döndürür."""
    with urllib.request.urlopen(ENGLISH_BUSINESS_URL, timeout=REQUEST_TIMEOUT_SEC) as response:
        parquet_bytes = io.BytesIO(response.read())
        df_biz = pq.read_table(
            parquet_bytes, columns=["business_id", "name", "categories"]
        ).to_pandas()

    business_map: Dict[str, Tuple[str, str]] = {}
    for biz_id, name, categories in zip(df_biz["business_id"], df_biz["name"], df_biz["categories"]):
        if biz_id and name:
            business_map[str(biz_id)] = (str(name).strip(), str(categories or "Restaurants").strip())
    return business_map


def prepare_english_dataset() -> pd.DataFrame:
    """İngilizce Yelp Restoran & Kafe veri setini işletme bilgileriyle eşleştirir ve kaydeder."""
    print("\n2) İngilizce Yelp Restoran & Kafe veri seti hazırlanıyor...")
    start_time = time.time()

    business_map = load_english_business_metadata()

    with urllib.request.urlopen(ENGLISH_REVIEWS_URL, timeout=REQUEST_TIMEOUT_SEC) as response:
        parquet_bytes = io.BytesIO(response.read())
        df_reviews = pq.read_table(
            parquet_bytes, columns=["business_id", "stars", "text"]
        ).to_pandas()

    buckets: Dict[str, List[ReviewRecord]] = {label: [] for label in SENTIMENT_CLASSES}
    seen_texts: Set[str] = set()

    for biz_id, stars_raw, text_raw in zip(
        df_reviews["business_id"], df_reviews["stars"], df_reviews["text"]
    ):
        biz_id_str = str(biz_id)
        if biz_id_str not in business_map:
            continue

        try:
            rating = int(round(float(stars_raw)))
            label = map_star_rating_to_sentiment(rating)
        except (ValueError, TypeError):
            continue

        text = normalize_whitespace(str(text_raw or ""))
        if not is_valid_length(text):
            continue

        text_key = text.lower()
        if text_key in seen_texts:
            continue
        seen_texts.add(text_key)

        place_name, place_category = business_map[biz_id_str]
        buckets[label].append(
            ReviewRecord(
                place_name=place_name,
                place_category=place_category,
                rating=rating,
                text=text,
                true_label=label,
            )
        )

    df_english = sample_balanced_dataframe(buckets, RANDOM_SEED)
    df_english.to_csv(ENGLISH_OUTPUT_CSV, index=False, encoding="utf-8")

    elapsed = time.time() - start_time
    print(f"   -> Tamamlandı ({elapsed:.1f} sn): {ENGLISH_OUTPUT_CSV.name} ({len(df_english):,} satır)")
    return df_english


def main() -> None:
    """Veri hazırlama iş akışını sırasıyla çalıştırır."""
    prepare_turkish_dataset()
    prepare_english_dataset()
    print("\nTüm veri setleri Clean Code standartlarında başarıyla oluşturuldu.")


if __name__ == "__main__":
    main()
