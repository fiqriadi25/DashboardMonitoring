# =======================================================================
# GOOGLE COLAB — COMBINED SCRIPT
# 1. Generate Database Rekon (Aggregated)      → Sheet: Database
# 2. Non-Rekon SPJ Tracker (Individual Rows)   → Sheet: Data_Rekon
# =======================================================================
import pandas as pd
from datetime import datetime
import gspread
from google.colab import auth, files
from google.auth import default
import io
import pytz
import re

print("=" * 70)
print("  COMBINED: MONITORING REKON + NON-REKON SPJ TRACKER")
print("=" * 70)

# =====================================================
# SECTION A — TIMEZONE & UTILITIES
# =====================================================

jakarta_tz = pytz.timezone('Asia/Jakarta')
now        = datetime.now(jakarta_tz)
now_str    = now.strftime('%Y-%m-%d %H:%M:%S')

def clean_text(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = " ".join(x.split())
    return x.lower()

def clean_spj(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.upper()
    )

# =====================================================
# SECTION B — UPLOAD FILE  [Shared]
# =====================================================

print("\n[1/10] Upload File")
uploaded = files.upload()

if not uploaded:
    raise SystemExit("❌ Tidak Ada File Diupload!")

print(f"📂 Total File Diupload: {len(uploaded)}")
df_list = []

for fname, fdata in uploaded.items():
    print(f"  📄 Membaca: {fname}")
    content = io.BytesIO(fdata)
    try:
        temp = pd.read_csv(content) if fname.lower().endswith(".csv") else pd.read_excel(content)
        temp = temp.dropna(how="all")
        temp.columns = temp.columns.str.strip()
        temp["Source_File"] = fname
        df_list.append(temp)
        print(f"     ✅ {len(temp)} Baris")
    except Exception as e:
        print(f"     ❌ Gagal: {e}")

if not df_list:
    raise SystemExit("❌ Tidak ada file CSV/Excel yang berhasil dibaca.")

df_raw = pd.concat(df_list, ignore_index=True)
print(f"Total Raw Rows: {len(df_raw)}")

# =====================================================
# SECTION C — AUTO RENAME  [Shared]
# =====================================================

print("\n[2/10] Auto Rename Kolom")

mapping = {}
for col in df_raw.columns:
    c = col.lower()
    if "kode" in c and "op" in c:
        mapping[col] = "Kode OP"
    elif "nama" in c and "op" in c:
        mapping[col] = "Nama OP"
    elif "customer" in c:
        mapping[col] = "Customer"
    elif c == "spj" or (c.startswith("spj") and len(c) <= 5):
        mapping[col] = "SPJ"
    elif "produk" in c:
        mapping[col] = "Produk"
    elif "sla" in c and "tanggal" in c:
        mapping[col] = "Tanggal SLA"
    elif "sla" in c:
        mapping[col] = "Tanggal SLA"
    elif "ujp" in c:
        mapping[col] = "UJP"
    elif "charging" in c:
        mapping[col] = "Charging"
    elif "status" in c and "rekon" in c:
        mapping[col] = "Status Rekon"

df_raw = df_raw.rename(columns=mapping)

required_base = [
    "Kode OP", "Nama OP", "Customer", "SPJ", "Produk",
    "Tanggal SLA", "UJP", "Charging", "Status Rekon"
]
missing = [c for c in required_base if c not in df_raw.columns]
if missing:
    raise SystemExit(f"❌ Kolom Tidak Ditemukan: {missing}")

# =====================================================
# SECTION D — FORMAT TANGGAL  [Shared]
# =====================================================

print("\n[3/10] Format Tanggal")

df_raw["Tanggal SLA"] = pd.to_datetime(df_raw["Tanggal SLA"], errors="coerce")
df_raw = df_raw[df_raw["Tanggal SLA"].notna()].copy()
df_raw["Tanggal"] = df_raw["Tanggal SLA"].dt.date
df_raw["Tanggal SLA_str"] = df_raw["Tanggal SLA"].dt.strftime("%Y-%m-%d")

print(f"Valid Rows: {len(df_raw)}")

# =====================================================
# SECTION E — AUTH GOOGLE SHEETS  [Shared]
# =====================================================

print("\n[4/10] Autentikasi Google Sheets")

auth.authenticate_user()
creds, _ = default()
gc = gspread.authorize(creds)

# Dua spreadsheet yang digunakan
SPREADSHEET_REKON_URL = "https://docs.google.com/spreadsheets/d/12FcleafDZHI-OkyxWEfFQM4BXcx2RBNIfG2ShnoptGo/edit?usp=sharing"
SPREADSHEET_DB_URL    = "https://docs.google.com/spreadsheets/d/1d5eFnUM65lzpaJUJXSdpZG0BmK70SCf3TF5K4tljCZw/edit?usp=sharing"
ID_REKAP              = "1y1iFq7BI3mDiNFpGTl2uQLCFdwexNYu8NfVLsVSVChY"

sh_rekon = gc.open_by_url(SPREADSHEET_REKON_URL)
sh_db    = gc.open_by_url(SPREADSHEET_DB_URL)

print("✅ Terhubung Ke Semua Spreadsheet.")

# =====================================================
# SECTION F — LOAD MASTER TOH/WOH  [Shared]
# =====================================================

print("\n[5/10] Load Master TOH/WOH")

dict_master  = {}   # untuk Script 1 (key = clean text tuple string)
lookup_dict  = {}   # untuk Script 2 (key = tuple of stripped strings)

def load_master(sh_source):
    """Load Master_OP_OH dari spreadsheet manapun, return df_master."""
    try:
        ws_m = sh_source.worksheet("Master_OP_OH")
        data_m = ws_m.get_all_values()
        if len(data_m) > 1:
            df_m = pd.DataFrame(data_m[1:], columns=data_m[0])
            df_m.columns = df_m.columns.str.strip()
            return df_m
    except Exception as e:
        print(f"  ⚠️ Master tidak ditemukan di spreadsheet: {e}")
    return pd.DataFrame()

# Coba dari sh_db dulu, fallback ke sh_rekon
df_master = load_master(sh_db)
if df_master.empty:
    df_master = load_master(sh_rekon)

if not df_master.empty:
    # Cari kolom secara dinamis
    cols = df_master.columns.tolist()
    col_kode     = next((c for c in cols if "kode op"   in c.lower()), None)
    col_nama     = next((c for c in cols if "nama op"   in c.lower() and "lama" not in c.lower()), None)
    col_customer = next((c for c in cols if c.lower()  == "customer"), None)
    col_toh      = next((c for c in cols if "toh"       in c.lower() or "woh" in c.lower()), None)

    if col_toh is None:
        # Rename fallback
        for possible in ["TOH", "Nama TOH", "Nama WOH"]:
            if possible in cols:
                col_toh = possible
                break

    if col_kode and col_nama and col_customer:
        # Build dict_master (clean_text keys) — untuk Script 1
        df_master["_key_clean"] = (
            df_master[col_kode].apply(clean_text) + "|" +
            df_master[col_nama].apply(clean_text) + "|" +
            df_master[col_customer].apply(clean_text)
        )
        for _, r in df_master.iterrows():
            toh_val = str(r[col_toh]).strip() if col_toh else ""
            dict_master[r["_key_clean"]] = toh_val

        # Build lookup_dict (stripped string keys) — untuk Script 2
        for _, r in df_master.iterrows():
            key2 = (
                str(r[col_kode]).strip(),
                str(r[col_nama]).strip(),
                str(r[col_customer]).strip()
            )
            lookup_dict[key2] = str(r[col_toh]).strip() if col_toh else ""

        print(f"  Master Loaded: {len(dict_master)} Entri")
    else:
        print("  ⚠️ Kolom Master Tidak Lengkap, Lookup Dilewati.")
else:
    print("  ⚠️ Master Kosong, Lookup Dilewati.")

# =====================================================
# SECTION G — STATUS REKON SPLIT  [Shared]
# =====================================================

print("\n[6/10] Split Status Rekon")

df_raw["Status Rekon"] = (
    df_raw["Status Rekon"].astype(str)
    .str.strip()
    .str.replace(r'\s+', ' ', regex=True)
)
status_normalized = df_raw["Status Rekon"].str.lower()

df_nonrekon_raw = df_raw[status_normalized.str.contains("non")].copy()
df_rekon_raw    = df_raw[
    ~status_normalized.str.contains("non") &
    status_normalized.str.contains("rekon")
].copy()

df_nonrekon_raw["Status Rekon"] = "Non - Rekon"
df_rekon_raw["Status Rekon"]    = "Rekon"

print(f"  Non-Rekon : {len(df_nonrekon_raw)} Baris")
print(f"  Rekon     : {len(df_rekon_raw)} Baris")

# =====================================================
# ██████████████████████████████████████████████████
#   SCRIPT 1 — DATABASE REKON (AGGREGATED)
# ██████████████████████████████████████████████████
# =====================================================

print("\n" + "=" * 70)
print("  SCRIPT 1: DATABASE REKON (AGGREGATED)")
print("=" * 70)

# ---- 1A. Setup Sheet History Log ----

print("\n[1A/10] Setup History & Upload Log")

try:
    ws_log = sh_db.worksheet("History_Log")
except Exception:
    ws_log = sh_db.add_worksheet(title="History_Log", rows=100000, cols=20)
    ws_log.append_row([
        "Timestamp", "Action", "Kode OP", "Customer",
        "Produk Lama", "Produk Baru",
        "Tanggal SLA Lama", "Tanggal SLA Baru",
        "Status Lama", "Status Baru"
    ])

try:
    ws_upload_log = sh_db.worksheet("Upload_Log")
except Exception:
    ws_upload_log = sh_db.add_worksheet(title="Upload_Log", rows=5000, cols=10)
    ws_upload_log.append_row([
        "Timestamp", "File Upload", "Rows",
        "Insert", "Update", "Delete", "Product Change"
    ])

log_rows = []

# ---- 1B. Load Database Existing ----

print("\n[1B/10] Load Database Existing")

headers_db = [
    "Kode OP", "Nama OP", "Customer", "Nama TOH/WOH", "Produk", "Tanggal SLA",
    "UJP", "Charging", "Status Rekon",
    "Jumlah Non-Rekon", "Jumlah Rekon", "Jumlah Ritase",
    "Tanggal Update"
]

try:
    ws_db = sh_db.worksheet("Database")
    data_db = ws_db.get_all_values()
    df_existing_db = (
        pd.DataFrame(data_db[1:], columns=data_db[0])
        if len(data_db) > 1
        else pd.DataFrame(columns=headers_db)
    )
except Exception:
    ws_db = sh_db.add_worksheet(title="Database", rows=200000, cols=20)
    df_existing_db = pd.DataFrame(columns=headers_db)

df_existing_db = df_existing_db.reindex(columns=headers_db, fill_value="")

# ---- 1C. Aggregation ----

print("\n[1C/10] Agregasi Data")

# Use all rows (Rekon + Non-Rekon) for accurate aggregation
df_agg_source = df_raw.copy()
df_agg_source["Tanggal SLA"] = df_agg_source["Tanggal SLA_str"]

hasil_db = []
grouped = df_agg_source.groupby(["Kode OP", "Customer", "Produk", "Tanggal"])

for (kode_op, customer, produk, tanggal), g in grouped:
    ritase = len(g)
    status_series = g["Status Rekon"].astype(str).str.strip().str.lower()
    rekon  = (status_series == "rekon").sum()
    non    = ritase - rekon

    status_final = "Rekon" if rekon == ritase else "Non - Rekon"

    ujp      = pd.to_numeric(g["UJP"],      errors="coerce").fillna(0).sum()
    charging = pd.to_numeric(g["Charging"], errors="coerce").fillna(0).sum()

    key_m = (
        clean_text(kode_op) + "|" +
        clean_text(g["Nama OP"].iloc[0]) + "|" +
        clean_text(customer)
    )
    nama_toh = dict_master.get(key_m, "")

    hasil_db.append({
        "Kode OP"          : kode_op,
        "Nama OP"          : g["Nama OP"].iloc[0],
        "Customer"         : customer,
        "Nama TOH/WOH"     : nama_toh,
        "Produk"           : produk,
        "Tanggal SLA"      : tanggal.strftime("%Y-%m-%d"),
        "UJP"              : ujp,
        "Charging"         : charging,
        "Status Rekon"     : status_final,
        "Jumlah Non-Rekon" : non,
        "Jumlah Rekon"     : rekon,
        "Jumlah Ritase"    : ritase,
        "Tanggal Update"   : now_str
    })

df_new_db = pd.DataFrame(hasil_db)

if df_new_db.empty:
    print("  ⚠️ Tidak Ada Data Setelah Agregasi, Lewati Script 1.")

else:
    # ---- 1D. Key Setup ----
    # Key          : unique per (Kode OP | Customer | Produk | Tanggal SLA)
    # Key_NoProduk : used to detect product changes for same triplet minus Produk
    for d in [df_new_db, df_existing_db]:
        d["Key"] = (
            d["Kode OP"].astype(str)    + "|" +
            d["Customer"].astype(str)   + "|" +
            d["Produk"].astype(str)     + "|" +
            d["Tanggal SLA"].astype(str)
        )
        d["Key_NoProduk"] = (
            d["Kode OP"].astype(str)    + "|" +
            d["Customer"].astype(str)   + "|" +
            d["Tanggal SLA"].astype(str)
        )

    df_new_db      = df_new_db.set_index("Key")
    df_existing_db = df_existing_db.set_index("Key")

    existing_keys_db = set(df_existing_db.index)
    new_keys_db      = set(df_new_db.index)

    # ---- 1E. Product Change Detection ----
    # Detect cases where the same (Kode OP | Customer | Tanggal SLA) now has
    # a different Produk value compared to what is stored in the database.
    product_change_count = 0
    existing_prod_map = {
        row["Key_NoProduk"]: row["Produk"]
        for _, row in df_existing_db.iterrows()
    }

    for _, row in df_new_db.iterrows():
        knp = row["Key_NoProduk"]
        if knp in existing_prod_map and existing_prod_map[knp] != row["Produk"]:
            product_change_count += 1
            log_rows.append([
                now_str, "PRODUCT_CHANGE",
                row["Kode OP"], row["Customer"],
                existing_prod_map[knp], row["Produk"],
                row["Tanggal SLA"], row["Tanggal SLA"],
                "", ""
            ])

    # ---- 1F. Update Existing Records ----
    # Update rows whose key already exists in both new and existing data.
    update_keys_db = existing_keys_db & new_keys_db
    for key in update_keys_db:
        old = df_existing_db.loc[key]
        new = df_new_db.loc[key]

        # Log only the fields that actually changed
        for action, cond in [
            ("UPDATE_STATUS",  old["Status Rekon"] != new["Status Rekon"]),
            ("UPDATE_SLA",     old["Tanggal SLA"]  != new["Tanggal SLA"]),
            ("PRODUCT_CHANGE", old["Produk"]       != new["Produk"]),
        ]:
            if cond:
                log_rows.append([
                    now_str, action,
                    new["Kode OP"], new["Customer"],
                    old["Produk"],  new["Produk"],
                    old["Tanggal SLA"], new["Tanggal SLA"],
                    old["Status Rekon"], new["Status Rekon"]
                ])

        # Overwrite the existing row with fresh aggregated data
        df_existing_db.loc[key] = new

    # ---- 1G. Insert New Records ----
    # Keys that appear in the new upload but don't yet exist in the database.
    insert_keys_db = new_keys_db - existing_keys_db
    for key in insert_keys_db:
        row = df_new_db.loc[key]
        log_rows.append([
            now_str, "INSERT",
            row["Kode OP"], row["Customer"],
            "", row["Produk"],
            "", row["Tanggal SLA"],
            "", row["Status Rekon"]
        ])
        df_existing_db.loc[key] = row

    # ---- 1H. Safe Delete ----
    # FIXED LOGIC:
    #
    # Previous (buggy) behaviour:
    #   Delete ANY existing key whose Tanggal SLA falls within the uploaded
    #   period AND is not in the new data — this incorrectly wiped records
    #   belonging to completely different Kode OP / Customer combinations
    #   that happened to share the same date range.
    #
    # Correct behaviour:
    #   Only delete a key if ALL three conditions hold:
    #     1. Its Tanggal SLA is within the uploaded period.
    #     2. Its full key is absent from the new upload.
    #     3. Its (Kode OP, Customer) pair IS present in the new upload —
    #        meaning this combination was intentionally re-uploaded and the
    #        record was deliberately removed, rather than simply not included
    #        because it belongs to a different OP/Customer.
    #
    # This ensures records from OP codes / customers that were NOT part of
    # the current upload are never touched, regardless of date overlap.

    upload_sla_dates = set(df_new_db["Tanggal SLA"])

    # Build the set of (Kode OP, Customer) pairs present in the new upload
    new_op_customer_pairs = set(
        zip(df_new_db["Kode OP"], df_new_db["Customer"])
    )

    delete_keys_db = []
    for key in existing_keys_db:
        row  = df_existing_db.loc[key]
        pair = (row["Kode OP"], row["Customer"])

        sla_in_upload_period = row["Tanggal SLA"] in upload_sla_dates
        key_not_in_new_data  = key not in new_keys_db
        op_customer_uploaded = pair in new_op_customer_pairs  # ← THE CRITICAL FIX

        if sla_in_upload_period and key_not_in_new_data and op_customer_uploaded:
            delete_keys_db.append(key)

    # Log and remove each deleted key
    for key in delete_keys_db:
        row = df_existing_db.loc[key]
        log_rows.append([
            now_str, "DELETE",
            row["Kode OP"], row["Customer"],
            row["Produk"], "",
            row["Tanggal SLA"], "",
            row["Status Rekon"], ""
        ])

    if delete_keys_db:
        df_existing_db = df_existing_db.drop(index=delete_keys_db)

    # ---- 1I. Sort & Write to Sheet ----
    df_final_db = df_existing_db.reset_index(drop=True)
    df_final_db["_sort_dt"] = pd.to_datetime(df_final_db["Tanggal SLA"], errors="coerce")
    df_final_db = df_final_db.sort_values(["Kode OP", "Customer", "Produk", "_sort_dt"])
    df_final_db = df_final_db.drop(
        columns=["_sort_dt", "Key_NoProduk"], errors="ignore"
    )

    print("\n[1D/10] Write Database")
    data_db_upload = (
        [headers_db]
        + df_final_db[headers_db].fillna("").astype(str).values.tolist()
    )
    ws_db.clear()
    ws_db.update("A1", data_db_upload)

    # ---- Write History & Upload Logs ----
    if log_rows:
        ws_log.append_rows(log_rows)

    ws_upload_log.append_row([
        now_str,
        ",".join(uploaded.keys()),
        len(df_raw),
        len(insert_keys_db),
        len(update_keys_db),
        len(delete_keys_db),
        product_change_count
    ])

    print("\n  ===== SUMMARY SCRIPT 1 =====")
    print(f"  Insert         : {len(insert_keys_db)}")
    print(f"  Update         : {len(update_keys_db)}")
    print(f"  Delete         : {len(delete_keys_db)}")
    print(f"  Product Change : {product_change_count}")
    print(f"  Total Database : {len(df_final_db)}")
    print(f"  History Logged : {len(log_rows)}")
    print("  ✅ Database Rekon Selesai!")

# =====================================================
# ██████████████████████████████████████████████████
#   SCRIPT 2 — NON-REKON SPJ TRACKER (INDIVIDUAL)
# ██████████████████████████████████████████████████
# =====================================================

print("\n" + "=" * 70)
print("  SCRIPT 2: NON-REKON SPJ TRACKER (INDIVIDUAL)")
print("=" * 70)

# ---- 2A. Siapkan data Non-Rekon & Rekon dari file ----

print("\n[8/10] Proses Non-Rekon SPJ")

df_nonrekon_baru = df_nonrekon_raw.copy()
df_rekon_baru    = df_rekon_raw.copy()

df_nonrekon_baru["Tanggal SLA"] = df_nonrekon_baru["Tanggal SLA_str"]
df_rekon_baru["Tanggal SLA"]    = df_rekon_baru["Tanggal SLA_str"]

# Bersihkan SPJ & buat Key = SPJ saja
df_nonrekon_baru["SPJ"] = clean_spj(df_nonrekon_baru["SPJ"])
df_rekon_baru["SPJ"]    = clean_spj(df_rekon_baru["SPJ"])
df_nonrekon_baru["Key"] = df_nonrekon_baru["SPJ"]
df_rekon_baru["Key"]    = df_rekon_baru["SPJ"]

# ---- 2B. Ambil Data Existing dari Sheet Data_Rekon ----

required_tracker = [
    "Kode OP", "Nama OP", "Customer", "SPJ", "Produk",
    "Tanggal SLA", "UJP", "Charging", "Status Rekon"
]

try:
    ws_tracker = sh_rekon.worksheet("Data_Rekon")
    data_existing_tracker = ws_tracker.get_all_values()
    if len(data_existing_tracker) > 1:
        df_existing_tracker = pd.DataFrame(
            data_existing_tracker[1:], columns=data_existing_tracker[0]
        )
    else:
        df_existing_tracker = pd.DataFrame(columns=required_tracker)
except:
    ws_tracker = sh_rekon.add_worksheet(title="Data_Rekon", rows=1000, cols=20)
    df_existing_tracker = pd.DataFrame(columns=required_tracker)

# ---- 2C. Merge: Hapus Rekon, Tambah Non-Rekon Terbaru ----

if not df_existing_tracker.empty:
    df_existing_tracker["SPJ"] = clean_spj(df_existing_tracker["SPJ"])
    df_existing_tracker["Key"] = df_existing_tracker["SPJ"]
else:
    df_existing_tracker = pd.DataFrame(columns=required_tracker + ["Key"])

before_count = len(df_existing_tracker)
df_existing_tracker = df_existing_tracker[~df_existing_tracker["Key"].isin(df_rekon_baru["Key"])]
removed_count = before_count - len(df_existing_tracker)
print(f"  🗑️ Baris Dihapus (Sudah Rekon)  : {removed_count}")

df_existing_tracker = df_existing_tracker[~df_existing_tracker["Key"].isin(df_nonrekon_baru["Key"])]
df_final_tracker    = pd.concat([df_existing_tracker, df_nonrekon_baru], ignore_index=True)
print(f"  ➕ Total Baris Setelah Merge     : {len(df_final_tracker)}")

# Hapus kolom bantu
for col in ["Key", "Source_File", "Tanggal SLA_str", "Tanggal"]:
    if col in df_final_tracker.columns:
        df_final_tracker = df_final_tracker.drop(columns=[col])

# ---- 2D. Lookup Nama TOH/WOH ----

def get_nama_toh(row):
    key = (
        str(row.get("Kode OP",  "")).strip(),
        str(row.get("Nama OP",  "")).strip(),
        str(row.get("Customer", "")).strip()
    )
    return lookup_dict.get(key, "")

df_final_tracker["Nama TOH/WOH"] = df_final_tracker.apply(get_nama_toh, axis=1)

# ---- 2E. Lookup Status Adjustment dari DATA_REKAP ----

print("\n[9/10] Lookup Status Adjustment")

rekap_map = {}
try:
    sheet_rekap = gc.open_by_key(ID_REKAP).worksheet("DATA_REKAP")
    data_rekap  = sheet_rekap.get_all_values()

    if len(data_rekap) > 1:
        df_rekap      = pd.DataFrame(data_rekap[1:], columns=data_rekap[0])
        spj_col       = df_rekap.columns[6]    # Kolom G
        status_col    = df_rekap.columns[15]   # Kolom P

        df_rekap[spj_col] = (
            df_rekap[spj_col].astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True).str.upper()
        )
        df_rekap[status_col] = (
            df_rekap[status_col].astype(str).str.strip().str.upper()
        )
        rekap_map = dict(zip(df_rekap[spj_col], df_rekap[status_col]))
        print(f"  ✅ Total SPJ Termapping: {len(rekap_map)}")
    else:
        print("  ⚠️ DATA_REKAP Kosong")

except Exception as e:
    print(f"  ❌ Error Membaca DATA_REKAP: {e}")

def map_status_adjustment(spj):
    s = re.sub(r"\.0$", "", str(spj).strip()).upper()
    status = rekap_map.get(s, "")
    if status == "DONE":
        return "Done"
    elif status in ["PROCESS", "PROCESSING"]:
        return "Process"
    else:
        return "No Adjustment"

# ---- 2F. Final Assembly & Sort ----

headers_tracker = required_tracker + ["Nama TOH/WOH", "Status Adjustment"]

for col in headers_tracker:
    if col not in df_final_tracker.columns:
        df_final_tracker[col] = ""

df_final_tracker = df_final_tracker[headers_tracker]

# Apply Status Adjustment
df_final_tracker["SPJ"] = clean_spj(df_final_tracker["SPJ"])
df_final_tracker["Status Adjustment"] = df_final_tracker["SPJ"].apply(map_status_adjustment)

df_final_tracker = df_final_tracker.sort_values(
    ["Kode OP", "Customer", "Tanggal SLA", "SPJ"]
)

# ---- 2G. Write Data_Rekon ----

print("\n[10/10] Write Data_Rekon & Config")

data_tracker_upload = (
    [headers_tracker] +
    df_final_tracker.fillna("").astype(str).values.tolist()
)
ws_tracker.clear()
ws_tracker.update("A1", data_tracker_upload)
print(f"  ✅ Data_Rekon Diperbarui. Total Baris: {len(df_final_tracker)}")

# ---- 2H. Catat Waktu Update ke Config ----

try:
    ws_config = sh_rekon.worksheet("Config")
except:
    ws_config = sh_rekon.add_worksheet(title="Config", rows=1, cols=1)

ws_config.update("A1", [[now_str]])
print(f"  ✅ Config Diperbarui: {now_str}")

# ---- 2I. Summary ----

print("\n  ===== SUMMARY SCRIPT 2 =====")
print(f"  Non-Rekon Di File  : {len(df_nonrekon_baru)}")
print(f"  Rekon Dihapus      : {removed_count}")
print(f"  Total Baris Tracker: {len(df_final_tracker)}")
print("\n  Status Adjustment  :")
print(df_final_tracker["Status Adjustment"].value_counts().to_string())
print("  ✅ Non-Rekon SPJ Tracker Selesai!")

# =====================================================
# FINAL SUMMARY
# =====================================================

print("\n" + "=" * 70)
print("  ✅ SEMUA PROSES SELESAI")
print(f"  Waktu Selesai: {now_str}")
print("=" * 70)
