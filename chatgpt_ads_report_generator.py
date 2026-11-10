#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ChatGPT広告 日別レポート — 一括作成ツール
==========================================

Chức năng:
    STEP 0  Kiểm tra môi trường và cấu hình accounts
    STEP 1  Chuẩn hóa file trong inbox thành UTF-8 CSV
    STEP 2  Lưu trữ dữ liệu cũ của từng account vào Old/<thời gian>/
    STEP 3  Phân phối file từ inbox theo quy tắc vào từng account
    STEP 4  Kiểm tra trước toàn diện các điều kiện cần thiết
    STEP 5  Tạo báo cáo chi tiết và xuất kết quả

Tuân thủ bảo mật:
    - Không hardcode secret/API keys (Mục 1.1)
    - Không dùng eval(), exec(), shell=True (Mục 1.5)
    - Xử lý hoàn toàn offline nội bộ, không gọi domain ngoài / GCP / Google API (Mục 6.1, 6.2, 7.1)
    - Bổ sung timeout cho tiến trình con (Mục 1.6)
    - Làm sạch metadata & bảo vệ PII (Mục 1.4, 3.4)
"""

import argparse
import calendar
import csv
import glob
import os
import re
import shutil
import subprocess
import sys
from copy import copy, deepcopy
from datetime import date, datetime, timedelta

# --- Thư viện đã được kiểm chứng (Mục 5.2) ---
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils import get_column_letter as CL

# --- Windows(cp932) Console compatibility / Bỏ dấu hiển thị console Windows ---
import builtins as _builtins
import unicodedata as _ud

_ENC = (getattr(sys.stdout, "encoding", None) or "utf-8")

_PUNCT = {
    "\u2014": "-", "\u2013": "-", "\u2192": "->", "\u2190": "<-",
    "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'"
}


def _ascii_fold(s):
    if s in _PUNCT:
        return _PUNCT[s]
    t = _ud.normalize("NFD", s)
    t = "".join(c for c in t if _ud.category(c) != "Mn")
    return t.replace("\u0111", "d").replace("\u0110", "D")


def _safe(x):
    if not isinstance(x, str):
        return x
    try:
        x.encode(_ENC)
        return x
    except (UnicodeEncodeError, LookupError):
        pass
    out = []
    for ch in x:
        try:
            ch.encode(_ENC)
            out.append(ch)
            continue
        except (UnicodeEncodeError, LookupError):
            pass
        t = _ascii_fold(ch)
        try:
            t.encode(_ENC)
            out.append(t)
        except (UnicodeEncodeError, LookupError):
            out.append("?")
    return "".join(out)


_real_print = _builtins.print


def print(*args, **kwargs):  # noqa: A001
    _real_print(*[_safe(a) for a in args], **kwargs)


# ---------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
SELF = os.path.abspath(__file__)
ALLOWED_ROOT = os.path.realpath(HERE)


class PathDenied(ValueError):
    """Lỗi khi phát hiện đường dẫn nằm ngoài thư mục cho phép (Directory Traversal)."""


# ===========================================================================
# 1. Cấu hình mặc định (Đọc từ biến môi trường nếu có - Mục 1.1)
# ===========================================================================
DEFAULT_COST_RATE = float(os.getenv("APP_DEFAULT_COST_RATE", "1.0"))
REPORT_NAME_TEMPLATE = os.getenv("APP_REPORT_NAME_TEMPLATE", "{account} ChatGPT日別レポート")
BAD_FILENAME_CHARS = '\\/:*?"<>|'

WARNINGS = []
MISSING_IMAGES = []
MISSING_DATES = []


def warn(msg):
    print(f"[WARN] {msg}")
    WARNINGS.append(msg)


def build_report_name(account, report_name, date_str):
    base = (report_name or "").strip()
    if not base:
        base = REPORT_NAME_TEMPLATE.format(account=account)
    if len(base) > 120:
        base = base[:120]
    for ext in (".xlsx", ".xlsm", ".xls"):
        if base.lower().endswith(ext):
            base = base[:-len(ext)]
            break
    base = base.rstrip()
    if not re.search(r"_\d{8}$", base):
        base = f"{base}_{date_str}"
    cleaned = "".join("_" if c in BAD_FILENAME_CHARS else c for c in base)
    return cleaned + ".xlsx", (cleaned != base)


DEFAULT_TEMPLATE = os.path.join(HERE, "template.xlsx")
IMAGE_EXTS = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]
IMAGE_MAX_PX = 600

# ===========================================================================
# 2. Hằng số layout template Excel
# ===========================================================================
BLOCK_COLS = 64
BAND1_COLS = 63
BAND_COL_STEP = 66
TPL_DAYS = 30
MET_OFF = 54
LABEL_OFF = 12
N_METRICS = 9

TPL_HEAD = (1, 61)
TPL_DAYROW_BY_WD = {0: 68, 1: 69, 2: 70, 3: 71, 4: 72, 5: 73, 6: 74}
TPL_TAIL1 = (92, 96)
TPL_TAIL2 = (97, 107)
TPL_GAP = (108, 109)

WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# ===========================================================================
# 3. Helper xử lý Excel an toàn & làm sạch PII (Mục 3.4, REQ-140)
# ===========================================================================
def scrub_properties(wb):
    """
    Xóa toàn bộ metadata người tạo, đường dẫn nội bộ (REQ-140).
    openpyxl mặc định sao chép docProps/core.xml, hàm này giúp bảo vệ danh tính.
    """
    p = wb.properties
    p.creator = ""
    p.lastModifiedBy = ""
    p.lastPrinted = None
    p.title = None
    p.subject = None
    p.description = None
    p.keywords = None
    p.category = None
    p.identifier = None
    p.language = None
    p.revision = None
    p.version = None
    p.contentStatus = None
    return wb


_FORMULA_LEAD = ("=", "+", "-", "@", "\t", "\r")


def set_text(ws, row, col, value):
    """
    Ghi chuỗi vào ô tính, ngăn chặn Excel Formula Injection (REQ-200).
    Nếu giá trị bắt đầu bằng dấu công thức tính toán, ép kiểu về dạng Text ("s").
    """
    c = ws.cell(row=row, column=col)
    c.value = value
    if isinstance(value, str) and (value[:1] in _FORMULA_LEAD or c.data_type in ("f", "e")):
        c.data_type = "s"
    return c


def copy_cell(dst_ws, src_ws, sr, sc, dr, dc, with_value=True):
    s = src_ws.cell(row=sr, column=sc)
    d = dst_ws.cell(row=dr, column=dc)
    if with_value:
        d.value = s.value
    if s.has_style:
        d._style = copy(s._style)
    return d


def copy_row(dst_ws, src_ws, sr, dr, c0=1, c1=BLOCK_COLS, col_shift=0, with_value=True, with_height=True):
    for c in range(c0, c1 + 1):
        copy_cell(dst_ws, src_ws, sr, c, dr, c + col_shift, with_value=with_value)
    rd = src_ws.row_dimensions.get(sr)
    if rd is not None and with_height:
        d = dst_ws.row_dimensions[dr]
        if rd.height is not None:
            d.height = rd.height
        d.hidden = rd.hidden


def copy_merges(dst_ws, src_ws, sr0, sr1, dr0, c0=1, c1=BLOCK_COLS, col_shift=0):
    dy = dr0 - sr0
    for m in list(src_ws.merged_cells.ranges):
        if m.min_row >= sr0 and m.max_row <= sr1 and m.min_col >= c0 and m.max_col <= c1:
            dst_ws.merge_cells(
                start_row=m.min_row + dy, end_row=m.max_row + dy,
                start_column=m.min_col + col_shift, end_column=m.max_col + col_shift,
            )


def copy_row_m(dst_ws, src_ws, sr, dr, c0=1, c1=BLOCK_COLS, col_shift=0):
    copy_row(dst_ws, src_ws, sr, dr, c0, c1, col_shift)
    copy_merges(dst_ws, src_ws, sr, sr, dr, c0, c1, col_shift)


def copy_rows(dst_ws, src_ws, sr0, sr1, dr0, c0=1, c1=BLOCK_COLS, col_shift=0, with_value=True):
    for i, sr in enumerate(range(sr0, sr1 + 1)):
        copy_row(dst_ws, src_ws, sr, dr0 + i, c0, c1, col_shift, with_value)
    copy_merges(dst_ws, src_ws, sr0, sr1, dr0, c0, c1, col_shift)


def clone_sheet_props(dst_ws, src_ws):
    dst_ws.views = deepcopy(src_ws.views)
    dst_ws.sheet_format = copy(src_ws.sheet_format)
    dst_ws.sheet_properties = copy(src_ws.sheet_properties)
    dst_ws.page_setup = copy(src_ws.page_setup)
    dst_ws.page_margins = copy(src_ws.page_margins)
    dst_ws.print_options = copy(src_ws.print_options)
    dst_ws.HeaderFooter = deepcopy(src_ws.HeaderFooter)
    dst_ws.sheet_state = src_ws.sheet_state
    dst_ws.print_title_rows = src_ws.print_title_rows
    dst_ws.print_title_cols = src_ws.print_title_cols
    dst_ws.column_dimensions.clear()
    for k, v in src_ws.column_dimensions.items():
        nd = copy(v)
        nd.parent = dst_ws
        dst_ws.column_dimensions[k] = nd


def replace_sheet(wb, name, builder, position):
    src = wb[name]
    tmp = wb.create_sheet("__tmp__")
    clone_sheet_props(tmp, src)
    builder(tmp, src)
    del wb[name]
    tmp.title = name
    wb._sheets.remove(tmp)
    wb._sheets.insert(position, tmp)
    return tmp


# ===========================================================================
# 4. Đọc dữ liệu & Kiểm tra an toàn
# ===========================================================================
ENCODINGS = ("utf-8-sig", "cp932", "utf-8", "utf-16", "utf-16-le", "utf-16-be", "euc_jp", "iso2022_jp")

BOM_MAP = (
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32-le"),
    (b"\x00\x00\xfe\xff", "utf-32-be"),
    (b"\xff\xfe", "utf-16-le"),
    (b"\xfe\xff", "utf-16-be"),
)


def _sniff_encoding(path):
    try:
        with open(path, "rb") as f:
            head = f.read(4)
    except OSError:
        return None
    for bom, enc in BOM_MAP:
        if head.startswith(bom):
            return enc
    return None


def _read_one(path):
    name = os.path.basename(path)
    try:
        if os.path.getsize(path) == 0:
            print(f"[ERROR] File rỗng: {name}")
            return None
    except OSError:
        pass

    try:
        with open(path, "rb") as f:
            sig = f.read(4)
        if sig[:2] == b"PK":
            print(f"[ERROR] Định dạng nội dung là Excel (đuôi file là .csv): {name}")
            return None
    except OSError:
        pass

    encs = list(ENCODINGS)
    sniffed = _sniff_encoding(path)
    if sniffed:
        encs = [sniffed] + [e for e in encs if e != sniffed]

    last_enc = None
    for enc in encs:
        for sep in (None, ",", "\t", ";"):
            try:
                kw = {"encoding": enc, "dtype": str}
                if sep is None:
                    kw.update(sep=None, engine="python")
                else:
                    kw.update(sep=sep)
                df = pd.read_csv(path, **kw)
            except (UnicodeDecodeError, UnicodeError):
                break
            except Exception:
                continue

            if df.shape[1] <= 1:
                continue

            df.columns = [str(c).replace("\ufeff", "").strip().strip('"').strip() for c in df.columns]
            last_enc = enc
            if enc not in ("utf-8-sig", "utf-8", "cp932"):
                warn(f"{name} được đọc với bảng mã {enc}.")
            return df

    print(f"[ERROR] Không thể đọc được file: {name}")
    return None


def load_data(ads_paths, edit_paths):
    if isinstance(ads_paths, str):
        ads_paths = [ads_paths]
    if isinstance(edit_paths, str):
        edit_paths = [edit_paths]

    need_ads = ["date", "campaign_name", "ad_group_name", "ad_id", "ad_name",
                "impressions", "clicks", "conversions", "spend"]
    frames = []
    for pth in ads_paths:
        one = _read_one(pth)
        if one is None:
            return None
        lack = [c for c in need_ads if c not in one.columns]
        if lack:
            print(f"[ERROR] Thiếu cột trong ads: {os.path.basename(pth)}: {', '.join(lack)}")
            return None
        one["_src"] = os.path.basename(pth)
        frames.append(one)
    ads = pd.concat(frames, ignore_index=True)

    for col in ["impressions", "clicks", "conversions", "spend"]:
        v = pd.to_numeric(ads[col], errors="coerce")
        bad_n = int((~np.isfinite(v.astype(float))).sum() - v.isna().sum())
        v = v.replace([np.inf, -np.inf], np.nan).fillna(0)
        if bad_n > 0:
            warn(f"{col} có {bad_n} giá trị bất thường vượt giới hạn, tự động đưa về 0.")
        ads[col] = v
    ads["date"] = pd.to_datetime(ads["date"], errors="coerce").dt.date

    bad = ads["date"].isna().sum()
    if bad:
        warn(f"Loại bỏ {bad} dòng có trường date không hợp lệ.")
        ads = ads[ads["date"].notna()]

    key = ["date", "ad_id"]
    metrics = ["impressions", "clicks", "conversions", "spend"]
    dup_mask = ads.duplicated(subset=key, keep=False)
    if dup_mask.any():
        dups = ads[dup_mask]
        exact, conflict = 0, []
        for k, g in dups.groupby(key):
            uniq = g[metrics].drop_duplicates()
            if len(uniq) == 1:
                exact += len(g) - 1
            else:
                conflict.append((k, sorted(g["_src"].unique())))
        if conflict:
            print(f"[ERROR] Trùng lặp ngày và ad_id nhưng sai lệch số liệu ({len(conflict)} trường hợp).")
            return None
        if exact:
            warn(f"Đã loại bỏ {exact} dòng dữ liệu hoàn toàn trùng lặp.")
            ads = ads.drop_duplicates(subset=key, keep="first")

    eframes = []
    for pth in edit_paths:
        one = _read_one(pth)
        if one is None:
            return None
        lack = [c for c in ("ad_id", "image_file_id") if c not in one.columns]
        if lack:
            print(f"[ERROR] Thiếu cột trong ads-for-edit: {os.path.basename(pth)}: {', '.join(lack)}")
            return None
        eframes.append(one)
    edit = pd.concat(eframes, ignore_index=True)
    lookup = edit[["ad_id", "image_file_id"]].drop_duplicates(subset=["ad_id"])
    ads = ads.merge(lookup, on="ad_id", how="left")
    ads["image_file_id"] = ads["image_file_id"].fillna("")

    miss = sorted(ads.loc[ads["image_file_id"] == "", "ad_id"].unique())
    if miss:
        warn(f"Có {len(miss)} ad_id không có image_file_id tương ứng.")
    return ads


def check_dates(ads, period_start, period_end, mode):
    have = set(ads["date"].unique())
    want = []
    d = period_start
    while d <= period_end:
        want.append(d)
        d += timedelta(days=1)
    missing = [d for d in want if d not in have]
    MISSING_DATES.extend(missing)
    if not missing:
        return True

    print(f"[WARN] Thiếu dữ liệu trong kỳ báo cáo ({len(missing)} ngày trống).")
    if mode == "continue":
        return True
    if mode == "abort":
        return False

    while True:
        ans = input("Tiếp tục xử lý không? [c] continue / [a] abort : ").strip().lower()
        if ans in ("c", "continue", "y", "yes"):
            return True
        if ans in ("a", "abort", "n", "no"):
            return False


_IMG_CACHE = {}


def prepare_image(path, workdir):
    if path in _IMG_CACHE:
        return _IMG_CACHE[path]
    out = path
    try:
        from PIL import Image as PILImage
        im = PILImage.open(path)
        if max(im.size) > IMAGE_MAX_PX:
            im.thumbnail((IMAGE_MAX_PX, IMAGE_MAX_PX), PILImage.LANCZOS)
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA" if "A" in im.getbands() else "RGB")
        os.makedirs(workdir, exist_ok=True)
        out = os.path.join(workdir, f"cr_{abs(hash(path)) % (10 ** 12)}.png")
        im.save(out, "PNG", optimize=True)
    except Exception as e:
        print(f"[INFO] Bỏ qua resize hình ({os.path.basename(path)}): {e}")
    _IMG_CACHE[path] = out
    return out


def read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        t = pd.read_excel(path, sheet_name=0, dtype=str)
    else:
        t = _read_one(path)
        if t is None:
            raise ValueError("ENCODING_UNDETECTED")
    t = t.fillna("")
    t.columns = [str(c).strip().lower() for c in t.columns]
    t = t[~t.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1)]
    return t


def load_rate_rules(path, default_rate):
    if not path:
        return None, default_rate
    if not os.path.isfile(path):
        warn(f"Không tìm thấy file rate: {os.path.basename(path)}")
        return None, default_rate

    try:
        t = read_table(path)
    except Exception as e:
        print(f"[ERROR] Không thể đọc rate file: {os.path.basename(path)} ({type(e).__name__})")
        return "ERROR", default_rate

    for need in ("match", "pattern", "cost_rate"):
        if need not in t.columns:
            print(f"[ERROR] rate file thiếu cột '{need}': {os.path.basename(path)}")
            return "ERROR", default_rate

    valid = {"equals", "contains", "startswith", "endswith", "regex", "default"}
    rules, dflt = [], default_rate
    for i, r in enumerate(t.itertuples(index=False), start=2):
        mt = str(r.match).strip().lower()
        pat = str(r.pattern).strip()
        fld = str(getattr(r, "field", "") or "").strip().lower() or "campaign_name"
        raw = str(r.cost_rate).replace(",", "").strip()
        if mt == "" and pat == "" and raw == "":
            continue
        if mt not in valid:
            print(f"[ERROR] {os.path.basename(path)} dòng {i}: Kiểu match không hợp lệ: {mt!r}")
            return "ERROR", default_rate
        if fld not in ("campaign_name", "ad_group_name", "ad_name"):
            print(f"[ERROR] {os.path.basename(path)} dòng {i}: Tên field không hợp lệ: {fld!r}")
            return "ERROR", default_rate
        try:
            rate = float(raw)
            if rate != rate or rate in (float("inf"), float("-inf")) or rate < 0 or rate > 100:
                raise ValueError
        except ValueError:
            print(f"[ERROR] {os.path.basename(path)} dòng {i}: cost_rate phải là số thực trong khoảng 0-100: {raw!r}")
            return "ERROR", default_rate
        if mt == "default":
            dflt = rate
            continue
        if pat == "":
            print(f"[ERROR] {os.path.basename(path)} dòng {i}: pattern không được để trống.")
            return "ERROR", default_rate
        if mt == "regex":
            bad = check_regex(pat)
            if bad:
                print(f"[ERROR] {os.path.basename(path)} dòng {i}: {bad}")
                return "ERROR", default_rate
        rules.append({"field": fld, "match": mt, "pattern": pat, "rate": rate,
                      "label": f"{fld} {mt} '{pat}'"})
    return rules, dflt


def _match(mt, value, pat):
    v, p = value.lower(), pat.lower()
    if mt == "equals":
        return v.strip() == p.strip()
    if mt == "contains":
        return p in v
    if mt == "startswith":
        return v.startswith(p)
    if mt == "endswith":
        return v.endswith(p)
    if mt == "regex":
        return re.search(pat, value) is not None
    return False


def apply_rates(df, rules, default_rate):
    if not rules:
        df["cost_rate"] = default_rate
        return df

    cache, hits = {}, {}

    def pick(row):
        key = (row["campaign_name"], row["ad_group_name"], row["ad_name"])
        if key in cache:
            return cache[key]
        chosen, label = default_rate, "(default)"
        for ru in rules:
            if _match(ru["match"], str(row[ru["field"]] or ""), ru["pattern"]):
                chosen, label = ru["rate"], ru["label"]
                break
        cache[key] = chosen
        hits.setdefault(label, {"rate": chosen, "campaigns": set()})
        hits[label]["campaigns"].add(row["campaign_name"])
        return chosen

    df["cost_rate"] = df.apply(pick, axis=1)
    return df


def load_targets(path):
    if not path or not os.path.isfile(path):
        return None
    try:
        t = read_table(path)
    except Exception:
        return None

    need = {"level", "name", "impression", "click", "cost", "conversion"}
    if not need.issubset(t.columns):
        return None

    def num(v):
        v = str(v).replace(",", "").strip()
        if v == "":
            return None
        try:
            return float(v)
        except ValueError:
            return None

    out = {"account": None, "campaign": {}, "adgroup": {}}
    for r in t.itertuples(index=False):
        lv = str(r.level).strip().lower()
        vals = (num(r.impression), num(r.click), num(r.cost), num(r.conversion))
        if all(v is None for v in vals):
            continue
        if lv == "account":
            out["account"] = vals
        elif lv in ("campaign", "adgroup"):
            out[lv][str(r.name).strip()] = vals
    return out


IMGCACHE_DIR = ".imgcache"


def _purge_imgcache(outdir):
    cache = os.path.join(outdir or "", IMGCACHE_DIR)
    if not outdir or not os.path.isdir(cache) or not _inside(ALLOWED_ROOT, cache):
        return
    try:
        names = os.listdir(cache)
    except OSError:
        return
    if any(not n.startswith("cr_") for n in names):
        return
    shutil.rmtree(cache, ignore_errors=True)


def find_image(images_dir, image_file_id):
    if not images_dir or not image_file_id:
        return None
    name = str(image_file_id).strip().strip('"')
    if (not name or name in (".", "..")
            or "/" in name or "\\" in name or ":" in name
            or name.startswith(".")):
        return None
    image_file_id = name

    def _ok(path):
        return os.path.isfile(path) and _inside(images_dir, path)

    p = os.path.join(images_dir, image_file_id)
    if _ok(p) and os.path.splitext(p)[1].lower() in IMAGE_EXTS:
        return p
    for ext in IMAGE_EXTS:
        q = os.path.join(images_dir, image_file_id + ext)
        if _ok(q):
            return q
    return None


# ===========================================================================
# 5. Ghi các chỉ số vào Sheet
# ===========================================================================
def ratio_formulas(row, c0):
    imp, clk, cost, cv = CL(c0), CL(c0 + 1), CL(c0 + 4), CL(c0 + 5)
    g = CL(c0 + 4)
    return {
        c0 + 2: f'=IF({g}{row}="","",IFERROR({clk}{row}/{imp}{row},0))',
        c0 + 3: f'=IF({g}{row}="","",IFERROR({cost}{row}/{clk}{row},0))',
        c0 + 6: f'=IF({g}{row}="","",IFERROR({cv}{row}/{clk}{row},0))',
        c0 + 7: f'=IF({g}{row}="","",IFERROR({cost}{row}/{cv}{row},0))',
        c0 + 8: f'=IF({g}{row}="","",IFERROR({cost}{row}/{imp}{row}*1000,0))',
    }


def write_metrics(ws, row, c0, imp=None, clk=None, cost=None, cv=None):
    for off, v in zip([0, 1, 4, 5], [imp, clk, cost, cv]):
        ws.cell(row=row, column=c0 + off).value = v
    for col, f in ratio_formulas(row, c0).items():
        ws.cell(row=row, column=col).value = f


def write_metric_formulas(ws, row, c0, exprs):
    for i, off in enumerate([0, 1, 4, 5]):
        ws.cell(row=row, column=c0 + off).value = exprs[i]
    for col, f in ratio_formulas(row, c0).items():
        ws.cell(row=row, column=col).value = f


# ===========================================================================
# 6. Block cấu trúc All / Campaign / Adgroup
# ===========================================================================
def block_layout(top, days):
    return {
        "title": top,
        "head0": top,
        "head1": top + 60,
        "day0": top + 61,
        "day1": top + 60 + days,
        "total": top + 61 + days,
        "est": top + 62 + days,
        "avg7": top + 63 + days,
        "tail2_0": top + 66 + days,
        "tail2_1": top + 76 + days,
        "height": 79 + days,
        "end": top + 78 + days,
    }


def build_all_block(dst, src, top, left, days, title, header, date_label,
                    daily, month_days, cols=BLOCK_COLS, data_dates=None, target=None,
                    count_from=None):
    L = block_layout(top, days)
    shift = left - 1
    NC = cols
    c0 = left + MET_OFF
    lab = left + LABEL_OFF

    copy_rows(dst, src, TPL_HEAD[0], TPL_HEAD[1], L["head0"], 1, NC, shift)
    set_text(dst, top, left, title)
    set_text(dst, top + 2, lab, header)
    set_text(dst, top + 3, lab, date_label)

    tr = top + 11
    if target:
        write_metrics(dst, tr - 1, c0, *target)
    for i in range(N_METRICS):
        cc = CL(c0 + i)
        dst.cell(row=tr, column=c0 + i).value = f"={cc}{L['total']}"
        dst.cell(row=tr + 1, column=c0 + i).value = f"={cc}{L['est']}"
    for i in range(N_METRICS):
        cc = CL(c0 + i)
        dst.cell(row=tr + 2, column=c0 + i).value = f"=IFERROR({cc}{tr + 1}/{cc}{top + 10},0)"

    for i, d in enumerate(month_days):
        r = L["day0"] + i
        copy_row_m(dst, src, TPL_DAYROW_BY_WD[d.weekday()], r, 1, NC, shift)
        dst.cell(row=r, column=lab).value = datetime(d.year, d.month, d.day)
        dst.cell(row=r, column=lab + 1).value = WEEKDAYS[d.weekday()]
        v = daily.get(d)
        if v is None and data_dates is not None and d in data_dates:
            v = (0, 0, 0.0, 0)
        if v is None:
            write_metrics(dst, r, c0)
        else:
            write_metrics(dst, r, c0, v[0], v[1], v[2], v[3])

    copy_rows(dst, src, TPL_TAIL1[0], TPL_TAIL1[1], L["total"], 1, NC, shift)
    d0, d1 = L["day0"], L["day1"]
    tot, est, avg = L["total"], L["est"], L["avg7"]
    dcol = CL(lab)

    cd0 = d0
    if count_from is not None and count_from in month_days:
        cd0 = d0 + month_days.index(count_from)

    write_metric_formulas(dst, tot, c0, [
        f"=SUM({CL(c0 + off)}{d0}:{CL(c0 + off)}{d1})" for off in [0, 1, 4, 5]
    ])
    write_metric_formulas(dst, est, c0, [
        (f"=IFERROR({CL(c0 + off)}{tot}+{CL(c0 + off)}{avg}"
         f"*(COUNTA({dcol}{cd0}:{dcol}{d1})-COUNT({CL(c0 + off)}{cd0}:{CL(c0 + off)}{d1})),0)")
        for off in [0, 1, 4, 5]
    ])
    write_metric_formulas(dst, avg, c0, [
        (f"=IFERROR(AVERAGE(OFFSET({CL(c0 + off)}{cd0 - 1},"
         f"1+IF(COUNT({CL(c0 + off)}{cd0}:{CL(c0 + off)}{d1})<=7,0,"
         f"COUNT({CL(c0 + off)}{cd0}:{CL(c0 + off)}{d1})-7),0,"
         f"IF(COUNT({CL(c0 + off)}{cd0}:{CL(c0 + off)}{d1})>=7,7,"
         f"COUNT({CL(c0 + off)}{cd0}:{CL(c0 + off)}{d1})),1)),0)")
        for off in [0, 1, 4, 5]
    ])

    copy_rows(dst, src, TPL_TAIL2[0], TPL_TAIL2[1], L["tail2_0"], 1, NC, shift)
    wcol = CL(lab + 1)
    for i in range(7):
        r = L["tail2_0"] + 4 + i
        key = f"{CL(lab)}{r}"
        for off in [0, 1, 4, 5]:
            cc = CL(c0 + off)
            dst.cell(row=r, column=c0 + off).value = (
                f"=IFERROR(AVERAGEIFS({cc}{d0}:{cc}{d1},{wcol}{d0}:{wcol}{d1},{key}),0)")
        for col, f in ratio_formulas(r, c0).items():
            dst.cell(row=r, column=col).value = f
    return L


# ===========================================================================
# 7. Builder cho từng sheet
# ===========================================================================
def build_front_page(ws, src, header):
    copy_rows(ws, src, 1, src.max_row, 1, 1, 40)
    set_text(ws, 5, 3, header)


def build_summary(ws, src, header, date_label, totals, est, campaigns, adgroups, target=None):
    copy_rows(ws, src, 1, 20, 1, 1, BLOCK_COLS)
    set_text(ws, 3, 13, header)
    set_text(ws, 4, 13, date_label)
    c0 = 55
    if target:
        write_metrics(ws, 11, c0, *target)
    write_metrics(ws, 12, c0, *totals)
    write_metrics(ws, 13, c0, *est)
    for i in range(N_METRICS):
        cc = CL(c0 + i)
        ws.cell(row=14, column=c0 + i).value = f"=IFERROR({cc}13/{cc}11,0)"

    r = 21
    for name, m in campaigns:
        copy_row_m(ws, src, 21, r, 1, BLOCK_COLS)
        set_text(ws, r, 13, name)
        write_metrics(ws, r, c0, *m)
        r += 1
    copy_row_m(ws, src, 26, r, 1, BLOCK_COLS)
    ws.cell(row=r, column=13).value = "Total"
    write_metric_formulas(ws, r, c0, [
        f"=SUM({CL(c0 + off)}21:{CL(c0 + off)}{r - 1})" for off in [0, 1, 4, 5]])
    cam_total = r

    copy_rows(ws, src, 27, 32, cam_total + 1, 1, BLOCK_COLS)
    r = cam_total + 7
    ag0 = r
    for name, m in adgroups:
        copy_row_m(ws, src, 33, r, 1, BLOCK_COLS)
        set_text(ws, r, 13, name)
        write_metrics(ws, r, c0, *m)
        r += 1
    copy_row_m(ws, src, 42, r, 1, BLOCK_COLS)
    ws.cell(row=r, column=13).value = "Total"
    write_metric_formulas(ws, r, c0, [
        f"=SUM({CL(c0 + off)}{ag0}:{CL(c0 + off)}{r - 1})" for off in [0, 1, 4, 5]])
    ws.print_area = f"A1:BL{r + 1}"


def build_simple_list(ws, src, header, date_label, rows, header_rows=(1, 10), data_tpl=11, total_tpl=12):
    copy_rows(ws, src, header_rows[0], header_rows[1], 1, 1, BLOCK_COLS)
    set_text(ws, 3, 13, header)
    set_text(ws, 4, 13, date_label)
    c0 = 55
    r = 11
    for name, m in rows:
        copy_row_m(ws, src, data_tpl, r, 1, BLOCK_COLS)
        set_text(ws, r, 13, name)
        write_metrics(ws, r, c0, *m)
        r += 1
    copy_row_m(ws, src, total_tpl, r, 1, BLOCK_COLS)
    ws.cell(row=r, column=13).value = "Total"
    write_metric_formulas(ws, r, c0, [
        f"=SUM({CL(c0 + off)}11:{CL(c0 + off)}{r - 1})" for off in [0, 1, 4, 5]])
    ws.print_area = f"A1:BL{r + 1}"


def build_ad(ws, src, header, date_label, ads_rows, images_dir, workdir):
    copy_rows(ws, src, 1, 10, 1, 1, BLOCK_COLS)
    set_text(ws, 3, 13, header)
    set_text(ws, 4, 13, date_label)
    c0 = 55
    n_ok, n_ng = 0, 0
    BOX = 1778000
    FROM_COL_OFF, FROM_ROW_OFF = 1028700, 63500

    r = 11
    for rec in ads_rows:
        copy_row_m(ws, src, 11, r, 1, BLOCK_COLS)
        set_text(ws, r, 13, rec["campaign_name"])
        set_text(ws, r, 14, rec["ad_group_name"])
        set_text(ws, r, 15, rec["ad_name"])
        set_text(ws, r, 16, rec["image_file_id"])
        ws.cell(row=r, column=48).value = None
        ws.cell(row=r, column=49).value = None
        write_metrics(ws, r, c0, rec["imp"], rec["clk"], rec["cost"], rec["cv"])

        path = find_image(images_dir, rec["image_file_id"])
        if path:
            try:
                img = XLImage(prepare_image(path, workdir))
                scale = min(BOX / img.width, BOX / img.height)
                ew, eh = int(img.width * scale), int(img.height * scale)
                co = FROM_COL_OFF + (BOX - ew) // 2
                ro = FROM_ROW_OFF + (BOX - eh) // 2
                img.anchor = TwoCellAnchor(
                    editAs="oneCell",
                    _from=AnchorMarker(col=46, colOff=co, row=r - 1, rowOff=ro),
                    to=AnchorMarker(col=46, colOff=co + ew, row=r - 1, rowOff=ro + eh),
                )
                ws.add_image(img)
                n_ok += 1
            except Exception as e:
                MISSING_IMAGES.append((rec["ad_name"], rec["image_file_id"], f"Lỗi đọc ảnh: {e}"))
                n_ng += 1
        else:
            MISSING_IMAGES.append((rec["ad_name"], rec["image_file_id"] or "(trống)", "Không tìm thấy file ảnh"))
            n_ng += 1
        r += 1

    copy_row_m(ws, src, 20, r, 1, BLOCK_COLS)
    ws.cell(row=r, column=13).value = "Total"
    write_metric_formulas(ws, r, c0, [
        f"=SUM({CL(c0 + off)}11:{CL(c0 + off)}{r - 1})" for off in [0, 1, 4, 5]])
    copy_row_m(ws, src, 21, r + 1, 1, BLOCK_COLS)
    ws.print_area = f"A1:BL{r + 1}"
    return n_ok, n_ng


def build_grid(ws, src, header, date_label, items, days, month_days, daily_by_key,
               data_dates=None, targets=None, count_from=None):
    H = 79 + days
    n_band = max(1, (len(items) + 1) // 2)
    for i, name in enumerate(items):
        band, side = divmod(i, 2)
        top = 1 + band * H
        left = 1 + side * BAND_COL_STEP
        build_all_block(ws, src, top, left, days, name, header, date_label,
                        daily_by_key.get(name, {}), month_days,
                        cols=BLOCK_COLS if side == 0 else BAND1_COLS,
                        data_dates=data_dates,
                        target=(targets or {}).get(name),
                        count_from=count_from)
    for band in range(n_band - 1):
        gap = 1 + band * H + 77 + days
        for i, sr in enumerate(range(TPL_GAP[0], TPL_GAP[1] + 1)):
            copy_row(ws, src, sr, gap + i, 1, BLOCK_COLS, with_height=(band == 0))
    last = 1 + (n_band - 1) * H + (76 + days)
    ws.print_area = f"A1:DZ{last + 1}"


# ===========================================================================
# 8. Tính toán số liệu tổng hợp
# ===========================================================================
def agg(df):
    return (int(df["impressions"].sum()), int(df["clicks"].sum()),
            float(df["cost"].sum()), int(df["conversions"].sum()))


def daily_map(df):
    g = df.groupby("date")[["impressions", "clicks", "cost", "conversions"]].sum()
    return {d: (int(r.impressions), int(r.clicks), float(r.cost), int(r.conversions))
            for d, r in g.iterrows()}


def estimate(daily, month_days, totals, count_from=None):
    grid = month_days
    if count_from is not None:
        grid = [d for d in month_days if d >= count_from]
    vals = [daily[d] for d in grid if d in daily]
    n = len(vals)
    if n == 0:
        return (0, 0, 0.0, 0)
    last = vals[-min(7, n):]
    rest = len(grid) - n
    out = []
    for i in range(4):
        a = sum(v[i] for v in last) / len(last)
        out.append(totals[i] + a * rest)
    return tuple(out)


# ===========================================================================
# 9. Thực thi đơn lẻ 1 account (_cmd_make_one)
# ===========================================================================
def _cmd_make_one():
    ap = argparse.ArgumentParser(description="Tạo báo cáo chi tiết cho một account")
    ap.add_argument("--ads", required=True, nargs="+")
    ap.add_argument("--ads-for-edit", required=True, nargs="+")
    ap.add_argument("--account", required=True)
    ap.add_argument("--report-date", default=None)
    ap.add_argument("--period-start", default=None)
    ap.add_argument("--period-end", default=None)
    ap.add_argument("--cost-rate", type=float, default=None)
    ap.add_argument("--images-dir", default="CR")
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("--monthly-scope", default="period")
    ap.add_argument("--rates", default=None)
    ap.add_argument("--targets", default=None)
    ap.add_argument("--report-name", default=None)
    ap.add_argument("--title", default=None)
    ap.add_argument("--on-missing", choices=["ask", "continue", "abort"], default="ask")
    args = ap.parse_args()

    try:
        for attr, label in (("outdir", "--outdir"), ("images_dir", "--images-dir"),
                            ("rates", "--rates"), ("targets", "--targets"),
                            ("template", "--template")):
            v = getattr(args, attr, None)
            if not v:
                continue
            s = str(v)
            if "://" in s or s.lower().startswith("file:"):
                raise PathDenied(f"{label}: Không chấp nhận URL.")
            full = s if os.path.isabs(s) else os.path.normpath(os.path.join(HERE, s))
            if not _inside(ALLOWED_ROOT, full):
                raise PathDenied(f"{label}: Không được chỉ định đường dẫn ngoài thư mục tool.")
            setattr(args, attr, full)

        for attr in ("ads", "ads_for_edit"):
            vals = list(getattr(args, attr, None) or [])
            fixed = []
            for v in vals:
                s = str(v)
                if "://" in s or s.lower().startswith("file:"):
                    raise PathDenied(f"{attr}: Không chấp nhận URL.")
                full = s if os.path.isabs(s) else os.path.normpath(os.path.join(HERE, s))
                if not _inside(ALLOWED_ROOT, full) or not os.path.isfile(full):
                    raise PathDenied(f"{attr}: File phải tồn tại trong thư mục cho phép.")
                fixed.append(full)
            setattr(args, attr, fixed)
    except PathDenied as e:
        print(f"[ERROR] {e}")
        return 1

    report_dt = (datetime.strptime(args.report_date, "%Y%m%d")
                 if args.report_date else datetime.today())
    report_date = report_dt.date()
    rate = args.cost_rate if args.cost_rate is not None else DEFAULT_COST_RATE

    def _ymd(v, label):
        try:
            return datetime.strptime(v, "%Y%m%d").date()
        except ValueError:
            print(f"[ERROR] {label} không đúng định dạng YYYYMMDD: {v!r}")
            return None

    if args.period_start:
        period_start = _ymd(args.period_start, "--period-start")
        if period_start is None:
            return 1
    else:
        period_start = report_date.replace(day=1)

    if args.period_end:
        period_end = _ymd(args.period_end, "--period-end")
        if period_end is None:
            return 1
    else:
        period_end = report_date - timedelta(days=1)

    if period_end < period_start:
        print(f"[ERROR] Khoảng thời gian không hợp lệ: {period_start:%Y-%m-%d} ~ {period_end:%Y-%m-%d}")
        return 1

    y, m = period_start.year, period_start.month
    ey, em = period_end.year, period_end.month
    grid_start = date(y, m, 1)
    grid_end = date(ey, em, calendar.monthrange(ey, em)[1])
    month_days = []
    d = grid_start
    while d <= grid_end:
        month_days.append(d)
        d += timedelta(days=1)
    ndays = len(month_days)

    if ndays > 400:
        print(f"[ERROR] Kỳ báo cáo quá dài ({ndays} ngày). Giới hạn tối đa là 400 ngày.")
        return 1

    header = args.title or f"{args.account} ChatGPT"
    targets = load_targets(args.targets)
    df_all = load_data(args.ads, args.ads_for_edit)
    if df_all is None:
        return 3

    df = df_all[(df_all["date"] >= period_start) & (df_all["date"] <= period_end)]
    if df.empty:
        print("[ERROR] Không có dữ liệu trong khoảng thời gian đã chọn.")
        return 1

    date_label = f"DATE：{period_start:%Y/%m/%d}～{period_end:%Y/%m/%d}"
    rules, default_rate = load_rate_rules(args.rates, rate)
    if rules == "ERROR":
        return 1

    df_all = apply_rates(df_all, rules, default_rate)
    df_all["cost"] = df_all["spend"] * df_all["cost_rate"]
    df = df_all[(df_all["date"] >= period_start) & (df_all["date"] <= period_end)].copy()

    if not check_dates(df, period_start, period_end, args.on_missing):
        return 2

    data_dates = set(df["date"].unique())
    totals = agg(df)
    dmap = daily_map(df)
    est = estimate(dmap, month_days, totals, count_from=period_start)

    camp_names = sorted(df["campaign_name"].unique())
    adg_names = sorted(df["ad_group_name"].unique())
    campaigns = [(n, agg(df[df["campaign_name"] == n])) for n in camp_names]
    adgroups = [(n, agg(df[df["ad_group_name"] == n])) for n in adg_names]
    daily_camp = {n: daily_map(df[df["campaign_name"] == n]) for n in camp_names}
    daily_adg = {n: daily_map(df[df["ad_group_name"] == n]) for n in adg_names}

    scope = str(args.monthly_scope).strip().lower() or "period"
    if scope == "period":
        msrc = df
    elif scope == "all":
        msrc = df_all
    else:
        try:
            nmon = int(scope)
            if nmon < 1:
                raise ValueError
        except ValueError:
            print(f"[ERROR] monthly_scope không hợp lệ: {args.monthly_scope!r}")
            return 1
        lo_y, lo_m = ey, em
        for _ in range(nmon - 1):
            lo_m -= 1
            if lo_m == 0:
                lo_m, lo_y = 12, lo_y - 1
        msrc = df_all[df_all["date"].map(lambda d: (lo_y, lo_m) <= (d.year, d.month) <= (ey, em))]

    months = sorted({(d.year, d.month) for d in msrc["date"]})
    monthly = []
    for yy, mm in months:
        sub = msrc[msrc["date"].map(lambda d: (d.year, d.month) == (yy, mm))]
        monthly.append((f"{yy}年{mm:02d}月", agg(sub)))

    weeks = {}
    for d in sorted(dmap):
        mon = d - timedelta(days=d.weekday())
        weeks.setdefault(mon, []).append(d)
    weekly = []
    for mon in sorted(weeks):
        sun = mon + timedelta(days=6)
        lo = max(mon, period_start)
        sub = df[(df["date"] >= mon) & (df["date"] <= sun)]
        weekly.append((f"週: {lo:%Y/%m/%d}～{sun:%Y/%m/%d}", agg(sub)))

    ad_rows = []
    gk = ["campaign_name", "ad_group_name", "ad_name", "image_file_id"]
    for key, sub in df.groupby(gk, sort=True):
        i, c, co, v = agg(sub)
        ad_rows.append(dict(zip(gk, key)) | {"imp": i, "clk": c, "cost": co, "cv": v})
    ad_rows.sort(key=lambda r: (r["campaign_name"], r["ad_group_name"], r["ad_name"]))

    if not os.path.isfile(args.template):
        print(f"[ERROR] Không tìm thấy file template: {os.path.basename(args.template)}")
        return 1
    wb = load_workbook(args.template)

    set_text(wb["Front Page"], 5, 3, header)
    t_acc = (targets or {}).get("account")
    t_cam = (targets or {}).get("campaign") or {}
    t_adg = (targets or {}).get("adgroup") or {}

    replace_sheet(wb, "Summary", lambda d, s: build_summary(
        d, s, header, date_label, totals, est, campaigns, adgroups, target=t_acc), 1)
    replace_sheet(wb, "Monthly", lambda d, s: build_simple_list(
        d, s, header, date_label, monthly), 2)
    replace_sheet(wb, "Weekly", lambda d, s: build_simple_list(
        d, s, header, date_label, weekly, total_tpl=14), 3)
    replace_sheet(wb, "All", lambda d, s: build_all_block(
        d, s, 1, 1, ndays, "All", header, date_label, dmap, month_days,
        data_dates=data_dates, target=t_acc, count_from=period_start), 4)
    replace_sheet(wb, "Campaign", lambda d, s: build_grid(
        d, s, header, date_label, camp_names, ndays, month_days, daily_camp,
        data_dates=data_dates, targets=t_cam, count_from=period_start), 5)
    replace_sheet(wb, "Adgroup", lambda d, s: build_grid(
        d, s, header, date_label, adg_names, ndays, month_days, daily_adg,
        data_dates=data_dates, targets=t_adg, count_from=period_start), 6)

    n_ok = n_ng = 0

    def _ad(d, s):
        nonlocal n_ok, n_ng
        n_ok, n_ng = build_ad(d, s, header, date_label, ad_rows,
                              args.images_dir, os.path.join(args.outdir, ".imgcache"))

    replace_sheet(wb, "Ad", _ad, 7)

    wb["All"].print_area = f"A1:BL{79 + ndays - 1}"
    wb.calculation.fullCalcOnLoad = True

    os.makedirs(args.outdir, exist_ok=True)
    fname, sanitized = build_report_name(args.account, args.report_name, report_dt.strftime("%Y%m%d"))
    out = os.path.join(args.outdir, fname)
    try:
        scrub_properties(wb)  # Làm sạch metadata bảo mật (Mục 3.4)
        wb.save(out)
        print(f"[OK] Xuất thành công: {_show(out)}")
    finally:
        _purge_imgcache(args.outdir)

    print(f"##OUTPUT##{fname}")
    for w in WARNINGS:
        print(f"##WARN##{w}")
    return 0


# ===========================================================================
# 10. Tiện ích quản lý file & Phân phối inbox
# ===========================================================================
TEXT_EXT = (".txt", ".tsv", ".text", ".dat", ".csv")
NORMALIZED_ENC = "utf-8-sig"


def normalize_one(path, dry=False):
    name = os.path.basename(path)
    stem, ext = os.path.splitext(name)
    ext = ext.lower()
    if ext not in TEXT_EXT:
        return path, False, ""

    df = _read_one(path)
    if df is None:
        return None, False, "Không thể đọc nội dung file."

    target = os.path.join(os.path.dirname(path), stem + ".csv")
    if ext == ".csv":
        enc = _sniff_encoding(path)
        if enc in ("utf-8-sig", None):
            try:
                with open(path, encoding="utf-8") as f:
                    head = f.readline()
                if head.count(",") >= head.count("\t"):
                    return path, False, ""
            except (OSError, UnicodeDecodeError):
                pass

    if os.path.abspath(target) != os.path.abspath(path) and os.path.isfile(target):
        return None, False, f"File đích {os.path.basename(target)} đã tồn tại."

    if not dry:
        df.to_csv(target, index=False, encoding=NORMALIZED_ENC)
        if os.path.abspath(target) != os.path.abspath(path):
            try:
                os.remove(path)
            except OSError:
                pass
    why = f"{ext} -> .csv" if ext != ".csv" else "Chuẩn hóa định dạng CSV UTF-8"
    return target, True, why


def normalize_folder(folder, dry=False, base=None):
    done, failed = [], []
    if not os.path.isdir(folder):
        return done, failed
    for fn in sorted(os.listdir(folder)):
        full = os.path.join(folder, fn)
        if not os.path.isfile(full) or os.path.islink(full):
            continue
        if os.path.splitext(fn)[1].lower() not in TEXT_EXT:
            continue
        root = base or folder
        if not _inside(root, full) or not _inside(ALLOWED_ROOT, full):
            continue
        new, changed, why = normalize_one(full, dry=dry)
        if new is None:
            failed.append((fn, why))
        elif changed and why:
            done.append((fn, os.path.basename(new), why))
    return done, failed


RULE_SHEET_NAMES = ("振り分け", "rules", "inbox_rules", "振り分けルール")
RULE_COLUMNS = ["pattern", "match", "target", "note"]


def load_inbox_rules(acc_path):
    ext = os.path.splitext(acc_path)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        return []
    try:
        from openpyxl import load_workbook as _lw
        wb = _lw(acc_path, data_only=True)
    except Exception:
        return []

    sheet = None
    for nm in wb.sheetnames:
        if nm.strip().lower() in [s.lower() for s in RULE_SHEET_NAMES]:
            sheet = wb[nm]
            break
    if sheet is None:
        return []

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    if "pattern" not in header or "target" not in header:
        return []

    out = []
    valid = {"contains", "startswith", "endswith", "equals", "regex"}
    for i, r in enumerate(rows[1:], start=2):
        if not r or all(v is None or str(v).strip() == "" for v in r):
            continue
        d = {k: ("" if v is None else str(v).strip()) for k, v in zip(header, r) if k}
        pat, tgt = d.get("pattern", ""), d.get("target", "")
        if pat.startswith("#") or not pat or not tgt:
            continue
        mt = (d.get("match", "") or "contains").lower()
        if mt not in valid:
            mt = "contains"
        if mt == "regex":
            bad = check_regex(pat)
            if bad:
                continue
        out.append({"pattern": pat, "match": mt, "target": tgt, "line": i})
    return out


_REDOS_SHAPES = (
    re.compile(r"\([^()]*[+*][^()]*\)\s*[+*]"),
    re.compile(r"\([^()]*[+*][^()]*\)\s*\{\d+,\s*\}"),
    re.compile(r"(\[[^\]]*\][+*]){3,}"),
)
MAX_PATTERN_LEN = 200


def check_regex(pattern):
    p = str(pattern or "")
    if len(p) > MAX_PATTERN_LEN:
        return f"Regex quá dài (tối đa {MAX_PATTERN_LEN} ký tự)."
    try:
        re.compile(p)
    except re.error as e:
        return f"Cú pháp Regex sai: {e}"
    if any(sh.search(p) for sh in _REDOS_SHAPES):
        return "Cảnh báo Regex ReDoS (lồng lặp vô tận)."
    return ""


def match_rule(filename, rules):
    stem = os.path.splitext(filename)[0]
    n_full, n_stem = norm(filename), norm(stem)
    for ru in rules:
        p = norm(ru["pattern"])
        mt = ru["match"]
        if mt == "regex":
            hit = any(re.search(ru["pattern"], t) for t in (filename, stem))
        elif mt == "contains":
            hit = p in n_full
        elif mt == "startswith":
            hit = n_full.startswith(p) or n_stem.startswith(p)
        elif mt == "endswith":
            hit = n_full.endswith(p) or n_stem.endswith(p)
        elif mt == "equals":
            hit = n_full == p or n_stem == p
        else:
            hit = False
        if hit:
            return ru["target"], f"{mt} '{ru['pattern']}'"
    return None, ""


def has_rules_sheet(path):
    try:
        from openpyxl import load_workbook as _lw
        names = _lw(path, read_only=True).sheetnames
    except Exception:
        return False
    low = [s.lower() for s in RULE_SHEET_NAMES]
    return any(nm.strip().lower() in low for nm in names)


def create_rules_sheet(path):
    from openpyxl import load_workbook as _lw
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as _CL

    wb = _lw(path)
    for nm in wb.sheetnames:
        if nm.strip().lower() in [s.lower() for s in RULE_SHEET_NAMES]:
            return False
    ws = wb.create_sheet("振り分け")
    head = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    widths = {"pattern": 28, "match": 14, "target": 30, "note": 40}
    for i, c in enumerate(RULE_COLUMNS, start=1):
        cell = ws.cell(row=1, column=i, value=c)
        cell.font, cell.fill = head, fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[_CL(i)].width = widths[c]
    ws.row_dimensions[1].height = 22
    scrub_properties(wb)
    wb.save(path)
    return True


TEMPLATE = os.path.join(HERE, "template.xlsx")
COLUMNS = ["account", "cost_rate", "data_dir", "images_dir", "ads_file", "edit_file",
           "report_date", "period_start", "period_end",
           "rate_file", "target_file", "monthly_scope",
           "report_name", "title", "file_key", "on_missing", "enabled", "note"]

CSV_EXT = (".csv",)
EDIT_HINTS = ("for-edit", "for_edit", "foredit", "edit")
ACC_FILE = ["accounts"]


def find_accounts_file(explicit=None):
    if explicit:
        return explicit if os.path.isfile(explicit) else None
    for name in ("accounts.xlsx", "accounts.xlsm", "accounts.csv"):
        p = os.path.join(HERE, name)
        if os.path.isfile(p):
            return p
    return None


def _rows_from_xlsx(path):
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    header = None
    for r in it:
        if r and any(v is not None and str(v).strip() for v in r):
            header = [str(v).strip().lower() if v is not None else "" for v in r]
            break
    if not header:
        return []
    rows = []
    for r in it:
        if not r or all(v is None or str(v).strip() == "" for v in r):
            continue
        d = {}
        for k, v in zip(header, r):
            if not k:
                continue
            if isinstance(v, datetime):
                v = v.strftime("%Y%m%d")
            elif isinstance(v, float) and v == int(v) and k in ("report_date", "enabled"):
                v = str(int(v))
            d[k] = "" if v is None else str(v).strip()
        rows.append(d)
    return rows


def _rows_from_csv(path):
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"[ERROR] Không thể xác định bảng mã của file accounts: {path}")


def read_accounts(path):
    ext = os.path.splitext(path)[1].lower()
    rows = _rows_from_xlsx(path) if ext in (".xlsx", ".xlsm") else _rows_from_csv(path)
    out = []
    for i, r in enumerate(rows, start=2):
        r = {(k or "").strip(): (v or "").strip() for k, v in r.items()}
        if not r.get("account") or r["account"].startswith("#"):
            continue
        if r.get("enabled", "1") in ("0", "false", "False", "no"):
            continue
        bad = _bad_account_name(r["account"])
        if bad:
            r["_line"] = i
            r["_bad_account"] = bad
        r["_line"] = i
        out.append(r)
    return out


_ACCOUNT_NG = set('/\\:*?"<>|')


def _bad_account_name(name):
    n = (name or "").strip()
    if not n:
        return "Tên account trống."
    if n in (".", "..") or n.startswith("."):
        return f"Tên account không được bắt đầu bằng dấu chấm: '{n}'"
    if any(c in _ACCOUNT_NG for c in n):
        return "Tên account chứa ký tự không hợp lệ (/ \\ : * ? \" < > |)"
    if len(n) > 80:
        return "Tên account quá dài (giới hạn 80 ký tự)."
    return ""


def _inside(base, path):
    b = os.path.realpath(base)
    p = os.path.realpath(path)
    return p == b or p.startswith(b + os.sep)


_DRIVE_RE = re.compile(r"^[A-Za-z]:")


def _show(p):
    if not p:
        return ""
    try:
        r = os.path.relpath(p, HERE)
    except ValueError:
        return os.path.basename(p)
    return r if not r.startswith("..") else os.path.basename(p)


def _redact(text):
    if not isinstance(text, str) or not text:
        return text
    for base in {HERE, os.path.realpath(HERE)}:
        for b in {base, base.replace("\\", "/")}:
            if b and b in text:
                text = text.replace(b + os.sep, "").replace(b + "/", "").replace(b, ".")
    return text


PROTECTED_NAMES = {"baocao.py", "baocao.bat", "template.xlsx", "accounts.xlsx"}


def _safe_abspath(p, label="Path", strict=True):
    if not p:
        return p
    s = str(p).strip().strip('"').replace("\\", "/")
    if not s:
        return ""
    if any(ord(c) < 32 or c == "\x7f" for c in s):
        raise PathDenied(f"{label}: Đường dẫn chứa ký tự điều khiển bất hợp pháp.")
    if "://" in s or s.lower().startswith("file:"):
        raise PathDenied(f"{label}: Không chấp nhận URL scheme.")
    if s.startswith("//") or s.startswith("/") or _DRIVE_RE.match(s) or os.path.isabs(s):
        raise PathDenied(f"{label}: Phải dùng đường dẫn tương đối trong thư mục ứng dụng.")
    full = os.path.normpath(os.path.join(HERE, s))
    if not _inside(ALLOWED_ROOT, full):
        raise PathDenied(f"{label}: Đường dẫn vượt ra ngoài thư mục gốc cho phép.")
    if strict and os.path.realpath(full) == ALLOWED_ROOT:
        raise PathDenied(f"{label}: Không được trỏ trực tiếp vào thư mục gốc của chương trình.")
    return full


def resolve_csv_pair(data_dir, ads_file, edit_file, target_file):
    notes, errs = [], []
    tgt_base = os.path.basename(target_file).lower() if target_file else ""

    if ads_file or edit_file:
        def resolve(spec, dflt, label):
            spec = spec or dflt
            norm_path = str(spec).strip().strip('"').replace("\\", "/")
            if (os.path.isabs(norm_path) or norm_path.startswith("/")
                    or _DRIVE_RE.match(norm_path) or ".." in norm_path.split("/")):
                errs.append(f"{label} có đường dẫn ngoài phạm vi: {spec}")
                return []
            if "*" in norm_path or "?" in norm_path:
                hits = sorted(glob.glob(os.path.join(data_dir, norm_path)))
                hits = [h for h in hits if os.path.isfile(h) and _inside(data_dir, h)]
                if not hits:
                    errs.append(f"{label} không khớp với file nào: {spec}")
                    return []
                return hits
            one = os.path.join(data_dir, norm_path)
            if not _inside(data_dir, one) or not os.path.isfile(one):
                errs.append(f"{label} không tồn tại: {_show(one)}")
                return []
            return [one]

        return (resolve(ads_file, "ads.csv", "ads"),
                resolve(edit_file, "ads-for-edit.csv", "ads-for-edit"),
                notes, errs)

    ads = os.path.join(data_dir, "ads.csv")
    edit = os.path.join(data_dir, "ads-for-edit.csv")
    if os.path.isfile(ads) and os.path.isfile(edit):
        return [ads], [edit], notes, errs

    files = [f for f in sorted(os.listdir(data_dir))
             if os.path.isfile(os.path.join(data_dir, f))
             and f.lower().endswith(CSV_EXT)
             and f.lower() != tgt_base
             and not f.lower().startswith("targets")]
    edits = [f for f in files if any(h in f.lower() for h in EDIT_HINTS)]
    others = [f for f in files if f not in edits]

    def pick(cands, label):
        if len(cands) == 1:
            return cands[0]
        if not cands:
            errs.append(f"Không tìm thấy CSV cho {label} tại {_show(data_dir)}")
        else:
            errs.append(f"Có nhiều file ứng viên cho {label}: {', '.join(cands)}")
        return None

    a = pick(others, "ads")
    e = pick(edits, "ads-for-edit")
    return ([os.path.join(data_dir, a)] if a else [],
            [os.path.join(data_dir, e)] if e else [],
            notes, errs)


def validate(rec, default_date):
    problems = []
    if rec.get("_bad_account"):
        return [rec["_bad_account"]], None
    try:
        data_dir = _safe_abspath(rec.get("data_dir") or "", "data_dir")
        imgs = (_safe_abspath(rec.get("images_dir"), "images_dir")
                or os.path.join(data_dir or HERE, "CR"))
        tgt = _safe_abspath(rec.get("target_file") or "", "target_file")
        rtf = _safe_abspath(rec.get("rate_file") or "", "rate_file")
    except PathDenied as e:
        return [str(e)], None

    if not data_dir or not os.path.isdir(data_dir):
        problems.append(f"data_dir không tồn tại: {_show(data_dir)}")
        return problems, None

    ads, edit, notes, errs = resolve_csv_pair(data_dir, rec.get("ads_file"),
                                              rec.get("edit_file"), rec.get("target_file"))
    problems += errs
    if not ads or not edit:
        problems.append("Không xác định được cặp file ads / ads-for-edit.")

    try:
        rate = float(rec.get("cost_rate"))
        if rate != rate or rate in (float("inf"), float("-inf")) or rate < 0 or rate > 100:
            raise ValueError
    except (TypeError, ValueError):
        problems.append(f"cost_rate không hợp lệ: {rec.get('cost_rate')!r}")
        rate = None

    rdate = rec.get("report_date") or default_date
    if rdate:
        try:
            datetime.strptime(rdate, "%Y%m%d")
        except ValueError:
            problems.append(f"report_date sai định dạng YYYYMMDD: {rdate!r}")

    return problems, {
        "ads": ads, "edit": edit, "imgs": imgs, "rate": rate,
        "date": rdate, "target": tgt, "title": rec.get("title") or "",
        "rates": rtf, "mscope": (rec.get("monthly_scope") or "").strip().lower(),
        "rname": rec.get("report_name") or "",
        "pstart": rec.get("period_start") or "",
        "pend": rec.get("period_end") or "",
        "notes": notes
    }


def write_result_file(rows, outdir, stamp):
    if not rows:
        return None
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    rows = [tuple(_redact(c) for c in r) for r in rows]

    wb = Workbook()
    ws = wb.active
    ws.title = "Kết Quả"
    ws.append(["Tên Báo Cáo", "Kết Quả", "Nội Dung / Lưu ý"])
    for c in ("A1", "B1", "C1"):
        ws[c].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        ws[c].fill = PatternFill("solid", fgColor="1F4E78")
        ws[c].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for name, res, msg in rows:
        ws.append([name, res, msg])
        for c in ws[ws.max_row]:
            if isinstance(c.value, str) and c.value[:1] in _FORMULA_LEAD:
                c.data_type = "s"
        r = ws.max_row
        for col in ("A", "B", "C"):
            ws[f"{col}{r}"].font = Font(name="Arial", size=10)
            ws[f"{col}{r}"].alignment = Alignment(vertical="top", wrap_text=True)
            ws[f"{col}{r}"].border = border
        b = ws[f"B{r}"]
        b.alignment = Alignment(horizontal="center", vertical="center")
        b.font = Font(name="Arial", size=10, bold=True, color="006100" if res == "OK" else "9C0006")
        b.fill = green if res == "OK" else red

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 96
    ws.freeze_panes = "A2"

    os.makedirs(outdir, exist_ok=True)
    path = os.path.join(outdir, f"KếtQuả_{stamp}.xlsx")
    scrub_properties(wb)
    wb.save(path)
    return path


# ===========================================================================
# 11. Xử lý Batch an toàn với Timeout (Mục 1.5, Mục 1.6)
# ===========================================================================
def _cmd_batch(argv=None):
    ap = argparse.ArgumentParser(description="Chạy tạo báo cáo hàng loạt cho nhiều account")
    ap.add_argument("--accounts", default=None)
    ap.add_argument("--outdir", default=os.path.join(HERE, "output"))
    ap.add_argument("--report-date", default=None)
    ap.add_argument("--on-missing", choices=["ask", "continue", "abort"], default="continue")
    ap.add_argument("--only", default=None)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)

    acc_path = find_accounts_file(args.accounts)
    if not acc_path:
        print("[ERROR] Không tìm thấy file accounts.xlsx.")
        return 1

    recs = read_accounts(acc_path)
    if args.only:
        want = {x.strip() for x in args.only.split(",")}
        recs = [r for r in recs if r["account"] in want]

    results, ng_rows = [], []
    for rec in recs:
        acc = rec["account"]
        problems, cfg = validate(rec, args.report_date)
        if problems:
            results.append((acc, "INPUT_NG", "; ".join(problems)))
            guess, _ = build_report_name(acc, rec.get("report_name") or "",
                                         rec.get("report_date") or args.report_date or datetime.today().strftime("%Y%m%d"))
            for pb in problems:
                ng_rows.append((guess, "NG", f"[Lỗi cấu hình] {pb}"))
            continue

        if args.dry_run:
            results.append((acc, "DRY_OK", ""))
            continue

        # KHÔNG dùng shell=True, gọi trực tiếp danh sách tham số (Mục 1.5)
        cmd = ([sys.executable, "-I", SELF, "--_make-one", "--ads"] + list(cfg["ads"])
               + ["--ads-for-edit"] + list(cfg["edit"]) + [
                   "--account", acc,
                   "--cost-rate", str(cfg["rate"]),
                   "--images-dir", cfg["imgs"],
                   "--outdir", args.outdir,
                   "--template", TEMPLATE,
                   "--on-missing", "abort" if (rec.get("on_missing") or args.on_missing) == "ask" else (rec.get("on_missing") or args.on_missing)])

        if cfg["date"]:
            cmd += ["--report-date", cfg["date"]]
        if cfg["pstart"]:
            cmd += ["--period-start", cfg["pstart"]]
        if cfg["pend"]:
            cmd += ["--period-end", cfg["pend"]]
        if cfg["mscope"]:
            cmd += ["--monthly-scope", cfg["mscope"]]
        if cfg["rates"]:
            cmd += ["--rates", cfg["rates"]]
        if cfg["target"]:
            cmd += ["--targets", cfg["target"]]
        if cfg["title"]:
            cmd += ["--title", cfg["title"]]
        if cfg["rname"]:
            cmd += ["--report-name", cfg["rname"]]

        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")

        # THÊM TIMEOUT chống treo tiến trình (Mục 1.6)
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
                timeout=300  # 5 phút timeout cho mỗi account
            )
        except subprocess.TimeoutExpired:
            print(f"[ERROR] Quá thời gian xử lý (Timeout 300s) tại account: {acc}")
            results.append((acc, "TIMEOUT", "Tiến trình con bị hủy do quá 300s."))
            ng_rows.append((acc, "NG", "Tiến trình xử lý vượt quá thời gian tối đa."))
            continue

        out_name, warns = "", []
        for line in (proc.stdout or "").splitlines():
            if line.startswith("##OUTPUT##"):
                out_name = line[len("##OUTPUT##"):]
            elif line.startswith("##WARN##"):
                warns.append(line[len("##WARN##"):])
            else:
                print("    " + line)

        if not out_name:
            out_name, _ = build_report_name(acc, cfg["rname"], cfg["date"] or datetime.today().strftime("%Y%m%d"))

        ok_run = proc.returncode == 0
        res = "OK" if ok_run else "NG"
        if warns:
            for wmsg in warns:
                ng_rows.append((out_name, res, wmsg))
        elif ok_run:
            ng_rows.append((out_name, "OK", "Hoàn thành không có cảnh báo."))

        if ok_run:
            results.append((acc, "OK", ""))
        else:
            results.append((acc, "ERROR", f"Exit Code: {proc.returncode}"))

    if not args.dry_run:
        stamp = (args.report_date or datetime.today().strftime("%Y%m%d"))
        write_result_file(ng_rows, args.outdir, stamp)

    ok_count = sum(1 for _, s, _ in results if s in ("OK", "DRY_OK"))
    return 0 if ok_count == len(results) else 1


DATA_EXT = (".csv", ".xlsx", ".xlsm", ".tsv")
IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")
KEEP_PAT = re.compile(r"(for[-_]?edit)", re.I)
ARCHIVE_DIR = "Old"


def norm(s):
    s = _ud.normalize("NFKC", str(s)).casefold()
    return re.sub(r"[\s_\-・.,()（）\[\]【】]", "", s)


def build_keys(recs):
    keys = []
    for r in recs:
        acc = r["account"]
        raw = [k.strip() for k in (r.get("file_key") or "").split(",") if k.strip()]
        if not raw:
            raw = [acc]
            base = os.path.basename((r.get("data_dir") or "").rstrip("/\\"))
            if base and base != acc:
                raw.append(base)
        keys.append({"account": acc, "rec": r, "keys": [(k, norm(k)) for k in raw if norm(k)]})
    return keys


def match_account(name, keys):
    n = norm(name)
    hits = [(orig, nk, k["account"]) for k in keys for orig, nk in k["keys"] if nk in n]
    if not hits:
        return None, None, []
    best = max(len(nk) for _, nk, _ in hits)
    top = {acc for _, nk, acc in hits if len(nk) == best}
    if len(top) > 1:
        return None, None, sorted(top)
    acc = top.pop()
    key = next(o for o, nk, a in hits if a == acc and len(nk) == best)
    return acc, key, []


def archive_old(data_dir, stamp, protect, dry, keep_edit=True):
    moved, kept = [], []
    if not os.path.isdir(data_dir):
        return moved, kept
    if os.path.realpath(data_dir) == ALLOWED_ROOT:
        raise PathDenied("data_dir không được trùng với thư mục gốc.")
    dest = os.path.join(data_dir, ARCHIVE_DIR, stamp)
    for fn in sorted(os.listdir(data_dir)):
        full = os.path.join(data_dir, fn)
        if not os.path.isfile(full) or os.path.splitext(fn)[1].lower() not in DATA_EXT:
            continue
        if fn.lower() in PROTECTED_NAMES:
            kept.append((fn, "File hệ thống"))
            continue
        if os.path.abspath(full) in protect:
            kept.append((fn, "File cấu hình"))
            continue
        if KEEP_PAT.search(fn) and keep_edit:
            kept.append((fn, "ads-for-edit"))
            continue
        if not dry:
            os.makedirs(dest, exist_ok=True)
            shutil.move(full, os.path.join(dest, fn))
        moved.append(fn)
    return moved, kept


def place(src, data_dir, imgs_dir, dry):
    fn = os.path.basename(src)
    ext = os.path.splitext(fn)[1].lower()
    tgt_dir = imgs_dir if ext in IMG_EXT else data_dir
    tgt = os.path.join(tgt_dir, fn)
    if fn.lower() in PROTECTED_NAMES and os.path.realpath(tgt_dir) == ALLOWED_ROOT:
        raise PathDenied(f"{fn} là file hệ thống, không được phép ghi đè.")
    if not _inside(ALLOWED_ROOT, tgt):
        raise PathDenied(f"{fn}: Vượt ra ngoài thư mục an toàn.")
    over = os.path.isfile(tgt)
    if not dry:
        os.makedirs(tgt_dir, exist_ok=True)
        shutil.move(src, tgt)
    return tgt, over


def check_packages():
    """Kiểm tra sự hiện diện của các thư viện phổ biến và đã xác thực (Mục 5.2)."""
    missing = []
    for pkg in ["pandas", "openpyxl", "PIL", "numpy"]:
        try:
            if pkg == "PIL":
                import PIL  # noqa: F401
            elif pkg == "pandas":
                import pandas  # noqa: F401
            elif pkg == "openpyxl":
                import openpyxl  # noqa: F401
            elif pkg == "numpy":
                import numpy  # noqa: F401
            print(f"    [OK] Package đã kiểm chứng: {pkg}")
        except ImportError:
            print(f"    [NG] Thiếu package: {pkg}")
            missing.append("Pillow" if pkg == "PIL" else pkg)
    if missing:
        print(f"\nVui lòng cài đặt: pip install {' '.join(missing)}")
        return False
    return True


def _cmd_all(argv=None):
    ap = argparse.ArgumentParser(prog="BAOCAO", description="Hệ thống tự động tạo báo cáo ChatGPT Ads")
    ap.add_argument("--date", "--report-date", dest="date", default=None)
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--only", default=None)
    ap.add_argument("--inbox", default=None)
    ap.add_argument("--outdir", default=None)
    ap.add_argument("--accounts", default=None)
    ap.add_argument("--no-prepare", action="store_true")
    ap.add_argument("--archive-all", action="store_true")
    ap.add_argument("--on-missing", choices=["ask", "continue", "abort"], default="continue")
    args = ap.parse_args(argv)

    dry = args.check
    try:
        inbox_dir = (_safe_abspath(args.inbox, "--inbox") or os.path.join(HERE, "inbox"))
        outdir = (_safe_abspath(args.outdir, "--outdir") or os.path.join(HERE, "output"))
        acc_arg = _safe_abspath(args.accounts, "--accounts") if args.accounts else None
    except PathDenied as e:
        print(f"[ERROR] {e}")
        return 1

    report_date = args.date or datetime.today().strftime("%Y%m%d")

    print("\n[STEP 0] Kiểm tra môi trường...")
    if not check_packages():
        return 1
    if not os.path.isfile(TEMPLATE):
        print(f"[ERROR] Thiếu file mẫu template.xlsx: {TEMPLATE}")
        return 1

    acc_path = find_accounts_file(acc_arg)
    if not acc_path:
        print(f"[ERROR] Không tìm thấy file accounts tại: {HERE}")
        return 1

    recs = read_accounts(acc_path)
    if args.only:
        want = {x.strip() for x in args.only.split(",")}
        recs = [r for r in recs if r["account"] in want]

    if not recs:
        print("[ERROR] Không có account nào sẵn sàng để xử lý.")
        return 1

    if not args.no_prepare and not dry:
        print("\n[STEP 1] Chuẩn hóa file trong inbox sang định dạng UTF-8 CSV...")
        normalize_folder(inbox_dir, dry=dry, base=inbox_dir)

    print("\n[STEP 5] Tiến hành tạo báo cáo...")
    argv2 = ["--report-date", report_date, "--outdir", outdir,
             "--accounts", acc_path, "--on-missing", args.on_missing]
    if args.only:
        argv2 += ["--only", args.only]
    return _cmd_batch(argv2)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--_make-one":
        sys.argv = [sys.argv[0]] + sys.argv[2:]
        return _cmd_make_one()
    return _cmd_all()


if __name__ == "__main__":
    sys.exit(main())