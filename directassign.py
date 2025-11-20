import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import math
from io import BytesIO
import re

st.set_page_config(page_title="Unread Summary", layout="wide")
st.title("دشبرد اساین دایرکت به کارشناسان")

# ---------- Sidebar settings ----------
with st.sidebar:
    st.header("تنظیمات")
    minutes_per_msg = st.number_input("زمان هر پیام (دقیقه)", min_value=0.1, value=1.0, step=0.1)

    sla_hours = st.number_input("SLA / بک‌لاگ هدف (ساعت)", min_value=0.5, value=3.5, step=0.5)
    efficiency = st.number_input("راندمان (0 تا 1)", min_value=0.1, max_value=1.0, value=0.85, step=0.05)

    experts_count = st.number_input("تعداد کارشناسان موجود", min_value=1, value=3, step=1)

uploaded_file = st.file_uploader("فایل Excel را آپلود کنید", type=["xlsx", "xls"])

# --------- helpers ---------
PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
ARABIC_DIGITS  = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def normalize_text(x):
    s = "" if pd.isna(x) else str(x)
    s = s.strip().translate(PERSIAN_DIGITS).translate(ARABIC_DIGITS)
    s = s.replace("\u200c", "")  # ZWNJ
    return s

def parse_custom(val):
    """پارس تاریخ از حالت‌های متنوع"""
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

def best_guess(cols, keys):
    cols_low = [normalize_text(c).lower() for c in cols]
    for i, c in enumerate(cols_low):
        if any(k in c for k in keys):
            return i
    return 0

def process_file(df, upload_time, account_col, date_col, status_col):
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

        delta_hours = (global_max_dt - newest_dt).total_seconds() / 3600.0
        delta_hours_rounded = round(delta_hours, 1)

        work_hours = (count_unread * minutes_per_msg) / 60.0
        effective_capacity_per_staff = sla_hours * efficiency
        needed_staff = math.ceil(work_hours / effective_capacity_per_staff) if count_unread > 0 else 0

        duration_hours = (work_hours / (needed_staff * efficiency)) if needed_staff > 0 else 0
        finish_time = upload_time + timedelta(hours=duration_hours)

        rows.append({
            "Account": account,
            "OldestUnreadDate": oldest_date_str,
            "OldestUnreadDT": oldest_dt,  # برای اولویت‌بندی
            "UnreadCount": count_unread,
            "HoursSinceLastUnread": delta_hours_rounded,
            "WorkHours(1msg=1min)": round(work_hours, 2),
            "NeededStaff(for_SLA)": needed_staff,
            "FinishBy(from_upload_time)": finish_time.strftime("%Y/%m/%d %H:%M:%S"),
        })

    result_df = pd.DataFrame(rows).sort_values("HoursSinceLastUnread", ascending=False)
    return result_df, None

def allocate_accounts(result_df, experts_count, sla_hours, efficiency, upload_time):
    """
    تقسیم اکانت‌ها بین کارشناسا:
    - ابتدا بر اساس قدیمی‌ترین unread سپس تعداد unread اولویت‌بندی
    - سپس گرِیدی روی کمترین لود
    """
    work_df = result_df.copy()

    # priority sort: older first, if tie higher unread first
    work_df = work_df.sort_values(
        by=["OldestUnreadDT", "UnreadCount"],
        ascending=[True, False]
    ).reset_index(drop=True)

    loads = [0.0 for _ in range(experts_count)]
    assigns = [[] for _ in range(experts_count)]

    for _, row in work_df.iterrows():
        idx = loads.index(min(loads))  # کم‌کارترین کارشناس
        assigns[idx].append(row["Account"])
        loads[idx] += float(row["WorkHours(1msg=1min)"])

    total_work = work_df["WorkHours(1msg=1min)"].sum()
    feasible = total_work <= experts_count * sla_hours * efficiency

    # ساخت جدول خروجی
    out_rows = []
    for i in range(experts_count):
        expert_hours = loads[i]
        duration_hours = expert_hours / efficiency if efficiency > 0 else 0
        finish_time = upload_time + timedelta(hours=duration_hours)

        out_rows.append({
            "Expert": f"کارشناس {i+1}",
            "AssignedAccounts": " , ".join(assigns[i]) if assigns[i] else "-",
            "AssignedAccountCount": len(assigns[i]),
            "WorkHours": round(expert_hours, 2),
            "FinishBy": finish_time.strftime("%Y/%m/%d %H:%M:%S")
        })

    alloc_df = pd.DataFrame(out_rows)
    return alloc_df, feasible, total_work


# ---------- UI FLOW ----------
if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        sheet = st.selectbox("شیت را انتخاب کنید", sheet_names)

        raw_preview = pd.read_excel(xl, sheet_name=sheet, header=None)
        st.subheader("پیش‌نمایش خام")
        st.dataframe(raw_preview.head(15), use_container_width=True)

        header_row = st.number_input(
            "شماره ردیف هدر (۱ یعنی اولین ردیف جدول)",
            min_value=1, max_value=50, value=1, step=1
        ) - 1

        df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
        st.caption(f"هدر از ردیف {header_row+1} خوانده شد.")
        st.subheader("پیش‌نمایش داده‌ها بعد از انتخاب هدر")
        st.dataframe(df.head(20), use_container_width=True)

        cols = list(df.columns)

        st.subheader("مپ ستون‌ها")
        account_col = st.selectbox("ستون Account", cols, index=best_guess(cols, ["account", "اکانت", "نام اکانت"]))
        date_col    = st.selectbox("ستون Date", cols, index=best_guess(cols, ["date", "تاریخ", "زمان"]))
        status_col  = st.selectbox("ستون Status", cols, index=best_guess(cols, ["status", "وضعیت", "state"]))

        upload_time = datetime.now()
        st.caption(f"زمان آپلود/شروع محاسبه: {upload_time.strftime('%Y/%m/%d %H:%M:%S')}")

        result_df, err = process_file(df, upload_time, account_col, date_col, status_col)

        if err:
            st.error(err)
        else:
            st.subheader("خلاصه خوانده‌نشده‌ها (به‌ازای هر اکانت)")
            st.dataframe(result_df.drop(columns=["OldestUnreadDT"]), use_container_width=True)

            # -------- Allocation --------
            st.subheader("تقسیم اکانت‌ها بین کارشناسان")
            alloc_df, feasible, total_work = allocate_accounts(
                result_df, experts_count, sla_hours, efficiency, upload_time
            )

            if feasible:
                st.success(
                    f"کل کار = {total_work:.2f} ساعت | "
                    f"با {experts_count} کارشناس، رسیدن به SLA={sla_hours} ساعت ممکنه."
                )
            else:
                st.warning(
                    f"کل کار = {total_work:.2f} ساعت | "
                    f"با {experts_count} کارشناس، تا SLA={sla_hours} ساعت تموم نمی‌شه."
                )

            st.dataframe(alloc_df, use_container_width=True)

            # دانلود خروجی‌ها
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_df.drop(columns=["OldestUnreadDT"]).to_excel(writer, index=False, sheet_name="UnreadSummary")
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


