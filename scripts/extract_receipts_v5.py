"""
extract_receipts_v5.py
----------------------
Extracts madrasa data from the 6 new images (July 21, 2026) and appends
to extracted_data.csv. These images contain:
  - 5 individual madrasa receipts/bills with institutional details
  - 1 Urdu Raabta examination table with 17 madrasa rows + 1 summary row

The data is manually extracted from careful reading of each image.
After appending, it moves processed images to images/done/.

Usage:
    python scripts/extract_receipts_v5.py
"""

import pandas as pd
import shutil
import datetime
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
CSV_PATH = ROOT / "extracted_data.csv"
ERROR_CSV_PATH = ROOT / "error_tracker.csv"
IMAGES_DIR = ROOT / "images"
DONE_DIR = IMAGES_DIR / "done"

SCHEMA = [
    'Entry Date', 'CODE', 'NAME OF MADRASAH', 'Raabta Centre',
    'ESTD.', 'VILLAGE', 'Post Office', 'Police Station', 'PIN',
    'TALEEM C', 'TOTAL S', 'BOARDIN S', 'District',
    'Municipality/Block-Panchayat',
    'Nearest Railway Station', 'Train from Sealdah', 'Auto Stand', 'Main Bazaar',
    'Source Image'
]

TODAY = datetime.date.today().strftime('%d/%m/%Y')

# ── Image-to-filename mapping ───────────────────────────────────────────────
IMG_RECEIPT_1 = "WhatsApp Image 2026-07-21 at 13.48.37.jpeg"
IMG_RECEIPT_2 = "WhatsApp Image 2026-07-21 at 13.48.38.jpeg"
IMG_RECEIPT_3 = "WhatsApp Image 2026-07-21 at 13.48.38 (1).jpeg"
IMG_RECEIPT_4 = "WhatsApp Image 2026-07-21 at 13.48.38 (2).jpeg"
IMG_RECEIPT_5 = "WhatsApp Image 2026-07-21 at 13.48.39.jpeg"
IMG_TABLE_URDU = "WhatsApp Image 2026-07-21 at 13.48.39 (1).jpeg"


def make_row(code, name, estd, village, po, ps, pin, taleem, total_s, boarding_s,
             district, raabta="", municipality="", source_img=""):
    """Create a row dict matching our CSV schema."""
    return {
        'Entry Date': TODAY,
        'CODE': code,
        'NAME OF MADRASAH': name,
        'Raabta Centre': raabta,
        'ESTD.': estd,
        'VILLAGE': village,
        'Post Office': po,
        'Police Station': ps,
        'PIN': pin,
        'TALEEM C': taleem,
        'TOTAL S': total_s,
        'BOARDIN S': boarding_s,
        'District': district,
        'Municipality/Block-Panchayat': municipality,
        'Nearest Railway Station': '',
        'Train from Sealdah': '',
        'Auto Stand': '',
        'Main Bazaar': '',
        'Source Image': source_img,
    }


def extract_all_rows():
    """Extract data from all 6 images."""
    rows = []

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 1: MADRASA JAMIA HEDAYATUL ISLAM
    # Receipt header: Reg. No.- IV-190200515/2019, ESTD. 1978
    # Vill.+P.O.- KAORAKHALI, P.S.-KULTALI, DIST.- 24 PGS. (S), PIN-743338
    # Urdu header shows: موضع پوسٹ کورا خالی، تھانہ کلتلی، ضلع ۲۴ پرگنہ ساؤتھ
    # ────────────────────────────────────────────────────────────────────────
    rows.append(make_row(
        code="",
        name="MADRASA JAMIA HEDAYATUL ISLAM",
        estd="1978",
        village="KAORAKHALI",
        po="KAORAKHALI",
        ps="KULTALI",
        pin="743338",
        taleem="",
        total_s="",
        boarding_s="",
        district="South 24 Parganas",
        source_img=IMG_RECEIPT_1,
    ))

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 2: DARUL ULOOM FIDA-E-MILLAT
    # Regd. No. IV-190300406/2022
    # VILL.-PADUA(U), P.O.-MAYAHAWRI, P.S.-BAKULTALA
    # DIST.- SOUTH 24 PARGANAS W.B. INDIA
    # Bengali text: পিন-৭৪৩৩৩৭ (PIN 743337)
    # ────────────────────────────────────────────────────────────────────────
    rows.append(make_row(
        code="",
        name="DARUL ULOOM FIDA-E-MILLAT",
        estd="",
        village="PADUA",
        po="MAYAHAWRI",
        ps="BAKULTALA",
        pin="743337",
        taleem="",
        total_s="",
        boarding_s="",
        district="South 24 Parganas",
        source_img=IMG_RECEIPT_2,
    ))

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 3: TALTALA JAMIA ISLAMIA ANOWARUL ULOOM
    # REGD NO. S/1L/41919, NO. II/21022/74 (0328)-FCRA-II ESTD:1994
    # P.O.- DEWANER CHAK, PIN - 743337, P.S.- BAKULTALA (JOYNAGAR)
    # DIST - 24 PGS (S) W.B. (INDIA)
    # Code No: 3079
    # ────────────────────────────────────────────────────────────────────────
    rows.append(make_row(
        code="S24PGS 3079",
        name="TALTALA JAMIA ISLAMIA ANOWARUL ULOOM",
        estd="1994",
        village="TALTALA",
        po="DEWANER CHAK",
        ps="BAKULTALA",
        pin="743337",
        taleem="",
        total_s="",
        boarding_s="",
        district="South 24 Parganas",
        source_img=IMG_RECEIPT_3,
    ))

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 4: NATUNHAT MADRASA JAMIA ISLAMIA QASIFUL ULOOM
    # Regd. No. 80486, ESTD : 1977
    # P.O.- Bakultala Natun Hat, Via- Nimpith
    # P.S.- Jaynagar (Bakultala), 24 Pgs. (S), W.B., India, Pin- 743338
    # Bengali: নতুনহাট মাদ্রাসা জামেয়া ইসলামিয়া কাশেফুল উলূম
    # ────────────────────────────────────────────────────────────────────────
    rows.append(make_row(
        code="",
        name="NATUNHAT MADRASA JAMIA ISLAMIA QASIFUL ULOOM",
        estd="1977",
        village="NATUNHAT",
        po="BAKULTALA NATUN HAT",
        ps="JAYNAGAR",
        pin="743338",
        taleem="",
        total_s="",
        boarding_s="",
        district="South 24 Parganas",
        source_img=IMG_RECEIPT_4,
    ))

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 5: DAKSHIN DURGAPUR MADRASHA QUASIMUL ULOOM
    # Regd. No. S/2L/30584, ESTD: 1996
    # D. DURGAPUR HIGH ROAD, 743349, P.S.-KULTALI, DIST. 24 PGS. (S), W.B., (INDIA)
    # ────────────────────────────────────────────────────────────────────────
    rows.append(make_row(
        code="",
        name="DAKSHIN DURGAPUR MADRASHA QUASIMUL ULOOM",
        estd="1996",
        village="DAKSHIN DURGAPUR",
        po="DAKSHIN DURGAPUR",
        ps="KULTALI",
        pin="743349",
        taleem="",
        total_s="",
        boarding_s="",
        district="South 24 Parganas",
        source_img=IMG_RECEIPT_5,
    ))

    # ────────────────────────────────────────────────────────────────────────
    # IMAGE 6: URDU TABLE - Annual Raabta Examination
    # Header: امتحان سالانہ رابطہ مدارس اسلامیہ عربیہ مغربی بنگال
    # Centre: مدرسہ اسلام مکاں، مقام چاک باڑیا
    # District: 24 Parganas South (ضلع ۲۴ پرگنہ جنوبی)
    # Chairman: Mufti Ainuddin Sahab, Phone: 8250093099
    #
    # The table is read RIGHT-TO-LEFT (Urdu).
    # Columns: شمار (Sr#) | مقتظہ مدارس (Madrasa Name) | کوڈنمبر (Phone/Code) |
    #          مقام (Village) | تعداد طلبہ (Students) | کافی (Category) | ... | مجموعہ (Total)
    #
    # Raabta Centre for ALL these rows: MADRASA ISLAM MAKAN, CHAK BARIA
    # All rows have: حاضر رابطہ (Present at Raabta) and کل طالب (All Students)
    # ────────────────────────────────────────────────────────────────────────

    raabta_centre = "MADRASA ISLAM MAKAN, CHAK BARIA"

    # Row 1: مدرسہ دار الاسلام امبیکا (MADRASA DARUL ISLAM AMBIKA)
    # Code/Phone: 8250093099, Village: چاں ڈیا (CHAN DIA)
    # Total (مجموعہ): 35, Students numbers: 15, 9, X, 5, 3, 3
    rows.append(make_row(
        code="",
        name="MADRASA DARUL ISLAM AMBIKA",
        estd="",
        village="CHAN DIA",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="35",
        boarding_s="15",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 2: جامعہ رشیدیہ (JAMIA RASHIDIYA)
    # Code/Phone: 9064956082, Village: روپ نگر (ROOP NAGAR)
    # Total: 33, Students: 21, 10, X, 5, 5, 3 (reading from right columns)
    rows.append(make_row(
        code="",
        name="JAMIA RASHIDIYA",
        estd="",
        village="ROOP NAGAR",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="33",
        boarding_s="21",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 3: جامعۃ العلوم (JAMIA TUL ULOOM)
    # Code/Phone: 7699020062, Village: اجوڑ (AJOD)
    # Total: 11, Students: 5, 3, X, 3, X, X
    rows.append(make_row(
        code="",
        name="JAMIA TUL ULOOM",
        estd="",
        village="AJOD",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="11",
        boarding_s="5",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 4: مدرسہ اسلامیہ سریریہ (MADRASA ISLAMIA SARARIYA)
    # Code/Phone: 9733735783, Village: سکمل گھول (SUKMAL GHOL)
    # Total: 16, Students: 3, 3, X, 3, 3, 3
    rows.append(make_row(
        code="",
        name="MADRASA ISLAMIA SARARIYA",
        estd="",
        village="SUKMAL GHOL",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="16",
        boarding_s="3",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 5: مدرسہ سرسپہ دار العلوم (MADRASA SARSAPA DARUL ULOOM)
    # Code/Phone: 8013114455, Village: مٹری ٹانگ (MITRI TANG)
    # Total: 16, Students: 6, 6, X, 3, X, X
    rows.append(make_row(
        code="",
        name="MADRASA SARSAPA DARUL ULOOM",
        estd="",
        village="MITRI TANG",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="16",
        boarding_s="6",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 6: مدینۃ العلوم (MADINATUL ULOOM)
    # Code/Phone: 9733617378, Village: علی پورہ کا کا پاڑہ (ALIPUR KAKA PARA)
    # Total: 13, Students: 5, 3, X, 3, 3, X
    rows.append(make_row(
        code="",
        name="MADINATUL ULOOM",
        estd="",
        village="ALIPUR KAKA PARA",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="13",
        boarding_s="5",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 7: تعلیم القرآن دار العلوم (TALIMUL QURAN DARUL ULOOM)
    # Code/Phone: 9732764671, Village: (not clearly visible, appears blank or same area)
    # Total: 55, Students: 55, X, X, X, X, X
    rows.append(make_row(
        code="",
        name="TALIMUL QURAN DARUL ULOOM",
        estd="",
        village="",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="55",
        boarding_s="55",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 8: عین العلوم کوھال باڑیا (AIN UL ULOOM KOHAL BARIA)
    # Code/Phone: 9564548449, Village: مومریج گوڑہ (MOMREJ GORH)
    # Total: 13, Students: 13, X, X, X, X, X
    rows.append(make_row(
        code="",
        name="AIN UL ULOOM KOHAL BARIA",
        estd="",
        village="MOMREJ GORH",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="13",
        boarding_s="13",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 9: رتنا پنچ اشرف العلوم (RATNA PANCH ASHRAF UL ULOOM)
    # Code/Phone: 9733538345, Village: سرایبا (SARAIBA)
    # Total: 62, Students: 33, 8, 5, 3, 3, X (right cols)
    rows.append(make_row(
        code="",
        name="RATNA PANCH ASHRAF UL ULOOM",
        estd="",
        village="SARAIBA",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="62",
        boarding_s="33",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 10: مدرسہ بیت العلوم (MADRASA BAIT UL ULOOM)
    # Code/Phone: 9732778961, Village: آم تلا (AAM TALA)
    # Total: 30, Students: 6, 10, X, 3, X, X
    rows.append(make_row(
        code="",
        name="MADRASA BAIT UL ULOOM",
        estd="",
        village="AAM TALA",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="30",
        boarding_s="6",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 11: مدرسہ مظاہر العلوم (MADRASA MAZAHIR UL ULOOM)
    # Code/Phone: 9734487466, Village: گوبند پور (GOBINDAPUR)
    # Total: 22, Students: 11, 6, X, X, X, X
    rows.append(make_row(
        code="",
        name="MADRASA MAZAHIR UL ULOOM",
        estd="",
        village="GOBINDAPUR",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="22",
        boarding_s="11",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 12: جامعہ اسلامیہ دار العلوم (JAMIA ISLAMIA DARUL ULOOM)
    # Code/Phone: 9735354368, Village: پنچم گھران (PANCHAM GHARAN)
    # Total: 30, Students: 19, 5, X, 2, X, 5 (from right)
    rows.append(make_row(
        code="",
        name="JAMIA ISLAMIA DARUL ULOOM",
        estd="",
        village="PANCHAM GHARAN",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="30",
        boarding_s="19",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 13: مدرسہ حاتم العلوم (MADRASA KHATIM UL ULOOM)
    # Code/Phone: 7872274142, Village: دھترہ ہاٹ (DHATRA HAT)
    # Total: 35, Students: 17, 9, X, X, X, X
    rows.append(make_row(
        code="",
        name="MADRASA KHATIM UL ULOOM",
        estd="",
        village="DHATRA HAT",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="35",
        boarding_s="17",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 14: اسلامیہ ریحان العلوم (ISLAMIA REHAN UL ULOOM)
    # Code/Phone: 8913755490, Village: اتر ہند پور (UTTAR HIND PUR)
    # Total: 6, Students: 3, 5, X, X, X, X (reading numbers)
    rows.append(make_row(
        code="",
        name="ISLAMIA REHAN UL ULOOM",
        estd="",
        village="UTTAR HIND PUR",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="6",
        boarding_s="3",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 15: جامعہ شیخ دار العلوم (JAMIA SHAIKH DARUL ULOOM)
    # Code/Phone: 6299754198, Village: مشرقی چک (MASHRIQI CHAK)
    # Total: 11, Students: 3, X, X, X, X, X
    rows.append(make_row(
        code="",
        name="JAMIA SHAIKH DARUL ULOOM",
        estd="",
        village="MASHRIQI CHAK",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="11",
        boarding_s="3",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 16: مدرسہ بیت اللکتاب (MADRASA BAIT UL KITAB)
    # Code/Phone: 7044981393, Village: دھترہ ہاٹ (DHATRA HAT)
    # Total: 13, Students: 13, X, X, X, X, X
    rows.append(make_row(
        code="",
        name="MADRASA BAIT UL KITAB",
        estd="",
        village="DHATRA HAT",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="13",
        boarding_s="13",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    # Row 17: مدرسہ حاتمۃ القرآن آئمہ (MADRASA KHATMATUL QURAN AAIMA)
    # Code/Phone: 8116944282, Village: تل پور (TAL PUR)
    # Total: 9, Students: 9, X, X, X, X, X
    rows.append(make_row(
        code="",
        name="MADRASA KHATMATUL QURAN AAIMA",
        estd="",
        village="TAL PUR",
        po="",
        ps="",
        pin="",
        taleem="HIFZ",
        total_s="9",
        boarding_s="9",
        district="South 24 Parganas",
        raabta=raabta_centre,
        source_img=IMG_TABLE_URDU,
    ))

    return rows


def get_existing_keys():
    """Load existing (CODE, NAME, PIN) keys from CSV."""
    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        return set(zip(
            existing['CODE'].str.strip().str.upper(),
            existing['NAME OF MADRASAH'].str.strip().str.upper(),
            existing['PIN'].str.strip()
        ))
    return set()


def main():
    DONE_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  Receipt & Urdu Table Extractor (v5 batch)")
    print("=" * 60)

    # Extract all rows
    new_rows = extract_all_rows()
    print(f"\nExtracted {len(new_rows)} total row(s) from 6 images.\n")

    # Load existing data
    existing_keys = get_existing_keys()
    print(f"Existing database has {len(existing_keys)} unique entries.\n")

    # Prepare DataFrame
    df_new = pd.DataFrame(new_rows, columns=SCHEMA)

    # Deduplicate against existing
    valid_rows = []
    skipped = 0
    for idx, row in df_new.iterrows():
        key = (
            str(row['CODE']).strip().upper(),
            str(row['NAME OF MADRASAH']).strip().upper(),
            str(row['PIN']).strip()
        )
        if key in existing_keys:
            print(f"  SKIP (duplicate): {row['NAME OF MADRASAH']}")
            skipped += 1
        else:
            valid_rows.append(row)
            existing_keys.add(key)

    if not valid_rows:
        print("\nNo new rows to add — all already in database!")
        return

    df_valid = pd.DataFrame(valid_rows, columns=SCHEMA)

    # Append to CSV
    if CSV_PATH.exists():
        existing_df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        for col in SCHEMA:
            if col not in existing_df.columns:
                existing_df[col] = ""
        existing_df = existing_df[SCHEMA]
    else:
        existing_df = pd.DataFrame(columns=SCHEMA)

    merged = pd.concat([existing_df, df_valid], ignore_index=True)
    merged.to_csv(CSV_PATH, index=False)

    print(f"\n{'=' * 60}")
    print(f"  RESULTS")
    print(f"{'=' * 60}")
    print(f"  New rows added   : {len(valid_rows)}")
    print(f"  Duplicates skipped: {skipped}")
    print(f"  Total CSV rows now: {len(merged)}")
    print()

    # List all added entries
    print("  Added entries:")
    for i, row in df_valid.iterrows():
        print(f"    {len(valid_rows) - len(df_valid) + i + 1}. {row['NAME OF MADRASAH']}"
              f" | PIN: {row['PIN'] or 'N/A'}"
              f" | Village: {row['VILLAGE'] or 'N/A'}")
    print()

    # Move processed images to done/
    images_to_move = [
        IMG_RECEIPT_1, IMG_RECEIPT_2, IMG_RECEIPT_3,
        IMG_RECEIPT_4, IMG_RECEIPT_5, IMG_TABLE_URDU
    ]
    moved_count = 0
    for img_name in images_to_move:
        src = IMAGES_DIR / img_name
        dst = DONE_DIR / img_name
        if src.exists():
            shutil.move(str(src), str(dst))
            print(f"  Moved: {img_name} -> images/done/")
            moved_count += 1
        else:
            print(f"  WARN: {img_name} not found (already moved?)")

    print(f"\nDone. Moved {moved_count} image(s) to images/done/.")
    print(f"Next step: Run inject_csv_to_excel.py to push into forms_v5.xlsm")


if __name__ == "__main__":
    main()
