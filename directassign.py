import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import math
from io import BytesIO
import re

st.set_page_config(page_title="Unread Summary", layout="wide")
st.title("Direct Assignment Dashboard (Auto)")

# ---------- Sidebar settings ----------
with st.sidebar:
    st.header("تنظیمات")
    minutes_per_msg = st.number_input(
        "زمان هر پیام (دقیقه)",
        min_value=0.1, value=1.0, step=0.1
    )
    sla_hours = st.number_input(
        "SLA / بک‌لاگ هدف (ساعت)",
        min_value=0.5, value=3.5, step=0.5
    )
    efficiency = st.number_input(
        "راندمان (0 تا 1)",
        min_value=0.1, max_value=1.0, value=0.85, step=0.05
    )
    experts_count = st.number_input(
        "تعداد کارشناسان موجود",
        min_value=1, value=3, step=1
    )

uploaded_file = st.file_uploader("فایل Excel را آپلود کنید", type=["xlsx", "xls"])

# --------- helpers ---------
PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
ARABIC_DIGITS  = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def normalize_text(x):
    s = "" if pd.isna(x) else str(x)
    s = s.strip().translate(PERSIAN_DIGITS).translate(ARABIC_DIGITS)
    s = s.replace("\u200c", "")  # ZWNJ
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_colname(c):
    c = normalize_text(c).lower()
    c = c.replace(":", "").replace("-", " ").replace("_", " ")
    c = re.sub(r"\s+", " ", c).strip()
    return c

def parse_custom(val):
    """پارس تاریخ از حالت‌های متنوع و اعداد فارسی"""
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val

    s = normalize_text(val)
    s = s.replace("-", "/").replace(".", "/")
    s = re.sub(r"\s+", " ", s)

    fmts = [
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except Exception:
            continue

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.to_pydatetime()
    except Exception:
        return None

def is_unread(val):
    s = normalize_text(val).lower()
    return any(k in s for k in [
        "خوانده نشده", "خوانده‌نشده", "unread", "not read", "new"
    ]) or s in ["0", "false", "no"]

def guess_header_row(raw_df):
    """هدر را از بین چند ردیف اول حدس می‌زند"""
    candidates = [
        "account", "اکانت", "نام اکانت", "user", "page",
        "date", "تاریخ", "زمان", "datetime", "time",
        "status", "وضعیت", "state", "read", "unread"
    ]
    for i in range(min(15, len(raw_df))):
        row = raw_df.iloc[i].astype(str).str.lower().tolist()
        if any(any(cand in cell for cand in candidates) for cell in row):
            return i
    return 0

def auto_map_columns(df):
    """
    ستون‌ها را به account/date/status مپ می‌کند.
    اگر نتواند یک ستون را پیدا کند، خطا می‌دهد.
    """
    cols = list(df.columns)
    norm_cols = [normalize_colname(c) for c in cols]

    synonyms = {
        "account": ["account", "acc", "page", "user", "اکانت", "نام اکانت", "کانال", "پیج", "ادمین"],
        "date": ["date", "datetime", "time", "timestamp", "created", "تاریخ", "زمان", "ساعت", "تایم"],
        "status": ["status", "state", "read status", "delivery", "وضعیت", "خوانده", "unread", "seen"],
    }

    def best_match(std_key):
        keys = synonyms[std_key]
        scores = []
        for idx, c in enumerate(norm_cols):
            score = 0
            for k in keys:
                if c == k:
                    score += 5
                if k in c:
                    score += 2
            score -= len(c) * 0.01
            scores.append((score, idx))
        scores.sort(reverse=True)
        best_score, best_idx = scores[0]
        return best_score, cols[best_idx]

    mapped = {}
    for std in ["account", "date", "status"]:
        score, col = best_match(std)
        if score <= 0:
            raise ValueError(f"ستون '{std}' اتومات پیدا نشد. ستون‌های فعلی: {cols}")
        mapped[std] = col

    return mapped

def process_file(df, upload_time, mapped_cols):
    account_col = mapped_cols["account"]
    date_col    = mapped_cols["date"]
    status_col  = mapped_cols["status"]

    df = df.copy()
    df["__account__"] = df[account_col]
    df["__date__"] = df[date_col]
    df["__status__"] = df[status_col]

    df["dt"] = df["__date__"].apply(parse_custom)

    unread = df[df["__status__"].apply(is_unread)].copy()
    if unread.empty:
        return None, "هیچ پیام 'خوانده نشده' پیدا نشد. مقدارهای Status را چک کنید."

    valid_dts = [d for d in df["dt"] if d is not None]
    if not valid_dts:
        return None, "هیچ تاریخ/ساعتی پارس نشد. ستون Date یا فرمتش مشکل دارد."

    global_max_dt = max(valid_dts)

    rows = []
    for account, g in unread.groupby("__account__"):
        dts = [d for d in g["dt"] if d is not None]
        if not dts:
            continue

        oldest_dt = min(dts)
        newest_dt = max(dts)
        count_unread = len(g)

        oldest_row = g.loc[g["dt"] == oldest_dt].iloc[0]
        oldest_date_str = str(oldest_row["__date__"])

        # اختلاف ساعت از آخرین unread تا جدیدترین پیام کل فایل
        delta_hours = (global_max_dt - newest_dt).total_seconds() / 3600.0
        delta_hours_rounded = round(delta_hours, 1)

        # ---------- محاسبات نیرو (خام و درست) ----------
        work_hours_raw = (count_unread * minutes_per_msg) / 60.0
        effective_capacity_per_staff = sla_hours * efficiency

        needed_staff = math.ceil(work_hours_raw / effective_capacity_per_staff) if count_unread > 0 else 0
        needed_staff = max(needed_staff, 1) if count_unread > 0 else 0

        duration_hours = (work_hours_raw / (needed_staff * efficiency)) if needed_staff > 0 else 0.0
        finish_time = upload_time + timedelta(hours=duration_hours)

        rows.append({
            "Account": account,
            "OldestUnreadDate": oldest_date_str,
            "UnreadCount": count_unread,
            "HoursSinceLastUnread": delta_hours_rounded,

            # ✅ ورک‌لود رو نگه می‌داریم
            "WorkHours(1msg=1min)": round(work_hours_raw, 2),

            "NeededStaff(for_SLA)": needed_staff,
            "FinishBy(from_upload_time)": finish_time.strftime("%Y/%m/%d %H:%M"),

            # داخلی برای تقسیم کار
            "WorkHoursRaw": work_hours_raw,
            "OldestUnreadDT": oldest_dt,
        })

    result_df = pd.DataFrame(rows).sort_values("HoursSinceLastUnread", ascending=False)
    return result_df, None

def allocate_accounts(result_df, experts_count, sla_hours, efficiency, upload_time):
    """
    تقسیم اکانت‌ها بین کارشناسا:
    - اولویت: قدیمی‌ترین unread بعد تعداد unread
    - سپس گرِیدی روی کمترین لود
    - پایان کل = کندترین کارشناس
    """
    work_df = result_df.copy().sort_values(
        by=["OldestUnreadDT", "UnreadCount"],
        ascending=[True, False]
    ).reset_index(drop=True)

    loads = [0.0 for _ in range(experts_count)]
    assigns = [[] for _ in range(experts_count)]

    for _, row in work_df.iterrows():
        idx = loads.index(min(loads))
        assigns[idx].append(row["Account"])
        loads[idx] += float(row["WorkHoursRaw"])  # خام

    total_work_raw = work_df["WorkHoursRaw"].sum()
    feasible = total_work_raw <= experts_count * sla_hours * efficiency

    out_rows = []
    durations = []
    for i in range(experts_count):
        expert_hours_raw = loads[i]
        duration_hours = expert_hours_raw / efficiency if efficiency > 0 else 0.0
        durations.append(duration_hours)

        finish_time = upload_time + timedelta(hours=duration_hours)

        out_rows.append({
            "Expert": f"کارشناس {i+1}",
            "AssignedAccounts": " , ".join(assigns[i]) if assigns[i] else "-",
            "AssignedAccountCount": len(assigns[i]),

            # ✅ ورک‌لود هر کارشناس
            "WorkHours": round(expert_hours_raw, 2),

            "FinishBy": finish_time.strftime("%Y/%m/%d %H:%M"),
        })

    alloc_df = pd.DataFrame(out_rows)
    overall_finish = upload_time + timedelta(hours=max(durations) if durations else 0.0)

    return alloc_df, feasible, overall_finish, round(total_work_raw, 2)


# ---------- UI FLOW ----------
if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        sheet = "Message Report" if "Message Report" in sheet_names else sheet_names[0]
        st.caption(f"شیت انتخاب‌شده (اتومات): {sheet}")

        raw_preview = pd.read_excel(xl, sheet_name=sheet, header=None)
        header_row = guess_header_row(raw_preview)

        df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
        st.caption(f"هدر (اتومات) از רدیف {header_row+1} تشخیص داده شد.")

        mapped_cols = auto_map_columns(df)
        st.info(
            f"مپ اتومات ستون‌ها: "
            f"Account ← `{mapped_cols['account']}` | "
            f"Date ← `{mapped_cols['date']}` | "
            f"Status ← `{mapped_cols['status']}`"
        )

        st.subheader("پیش‌نمایش داده‌ها")
        st.dataframe(df.head(20), use_container_width=True)

        upload_time = datetime.now()
        st.caption(f"زمان آپلود/شروع محاسبه: {upload_time.strftime('%Y/%m/%d %H:%M')}")

        result_df, err = process_file(df, upload_time, mapped_cols)

        if err:
            st.error(err)
        else:
            st.subheader("خلاصه خوانده‌نشده‌ها (به‌ازای هر اکانت)")

            show_summary = result_df.drop(columns=["WorkHoursRaw", "OldestUnreadDT"])
            st.dataframe(show_summary, use_container_width=True)

            st.subheader("تقسیم اکانت‌ها بین کارشناسا")
            alloc_df, feasible, overall_finish, total_work = allocate_accounts(
                result_df, experts_count, sla_hours, efficiency, upload_time
            )

            if feasible:
                st.success(
                    f"کل ورک‌لود: {total_work} ساعت | "
                    f"با {experts_count} کارشناس، رسیدن به SLA={sla_hours} ساعت ممکنه."
                )
            else:
                st.warning(
                    f"کل ورک‌لود: {total_work} ساعت | "
                    f"با {experts_count} کارشناس، تا SLA={sla_hours} ساعت تموم نمی‌شه."
                )

            st.dataframe(alloc_df, use_container_width=True)
            st.caption(f"پایان کل بک‌لاگ: {overall_finish.strftime('%Y/%m/%d %H:%M')}")

            # دانلود خروجی‌ها
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                show_summary.to_excel(writer, index=False, sheet_name="UnreadSummary")
                alloc_df.to_excel(writer, index=False, sheet_name="Allocation")
            output.seek(0)

            st.download_button(
                "دانلود خروجی Excel (Summary + Allocation)",
                data=output,
                file_name="unread_summary_with_allocation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"خطا در خواندن/پردازش فایل: {e}")

else:
    st.info("یک فایل Excel آپلود کنید تا پردازش انجام شود.")
