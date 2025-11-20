import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import math
from io import BytesIO

st.set_page_config(page_title="Unread Summary", layout="wide")
st.title("خلاصه پیام‌های خوانده‌نشده + تخمین نیرو")

# ---------- Sidebar settings ----------
with st.sidebar:
    st.header("تنظیمات محاسبه نیرو")
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

uploaded_file = st.file_uploader("فایل Excel را آپلود کنید", type=["xlsx", "xls"])


def parse_custom(s):
    """
    انتظار فرمت: 1404/08/24 07:21:15
    فقط برای اختلاف‌ها استفاده می‌شود (بدون تبدیل شمسی به میلادی)
    """
    try:
        return datetime.strptime(str(s), "%Y/%m/%d %H:%M:%S")
    except Exception:
        return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """ستون‌ها را از نظر فاصله/حروف یکدست می‌کند."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def process_file(df: pd.DataFrame, upload_time: datetime):
    df = normalize_columns(df)

    required_cols = {"Date", "Status", "Account"}
    if not required_cols.issubset(set(df.columns)):
        missing = required_cols - set(df.columns)
        return None, f"ستون‌های ضروری پیدا نشدن: {', '.join(missing)}"

    df["dt"] = df["Date"].apply(parse_custom)

    # فقط پیام‌های خوانده نشده
    unread = df[df["Status"].astype(str).str.strip() == "خوانده نشده"].copy()
    if unread.empty:
        return None, "هیچ پیام 'خوانده نشده' در فایل پیدا نشد."

    valid_dts = [d for d in df["dt"] if d is not None]
    if not valid_dts:
        return None, "هیچ تاریخ/ساعتی به‌درستی پارس نشد. فرمت Date را چک کنید."

    global_max_dt = max(valid_dts)

    rows = []
    for account, g in unread.groupby("Account"):
        dts = [d for d in g["dt"] if d is not None]
        if not dts:
            continue

        oldest_dt = min(dts)
        newest_dt = max(dts)
        count_unread = len(g)

        oldest_row = g.loc[g["dt"] == oldest_dt].iloc[0]
        oldest_date_str = str(oldest_row["Date"])

        # اختلاف ساعت از آخرین unread تا جدیدترین پیام کل فایل
        delta_hours = (global_max_dt - newest_dt).total_seconds() / 3600.0
        delta_hours_rounded = round(delta_hours, 1)

        # ---------- محاسبات نیرو ----------
        work_hours = (count_unread * minutes_per_msg) / 60.0
        effective_capacity_per_staff = sla_hours * efficiency

        needed_staff = math.ceil(work_hours / effective_capacity_per_staff) if count_unread > 0 else 0
        duration_hours = (work_hours / (needed_staff * efficiency)) if needed_staff > 0 else 0
        finish_time = upload_time + timedelta(hours=duration_hours)

        rows.append({
            "Account": account,
            "OldestUnreadDate": oldest_date_str,
            "UnreadCount": count_unread,
            "HoursSinceLastUnread": delta_hours_rounded,
            "WorkHours(1msg=1min)": round(work_hours, 2),
            "NeededStaff(for_3.5h_backlog)": needed_staff,
            "FinishBy(from_upload_time)": finish_time.strftime("%Y/%m/%d %H:%M:%S"),
        })

    result_df = pd.DataFrame(rows).sort_values(
        "HoursSinceLastUnread", ascending=False
    )

    return result_df, None


if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        default_sheet = "Message Report" if "Message Report" in sheet_names else sheet_names[0]
        sheet = st.selectbox("شیت را انتخاب کنید", sheet_names, index=sheet_names.index(default_sheet))

        df = pd.read_excel(xl, sheet_name=sheet)

        st.subheader("پیش‌نمایش داده‌ها")
        st.dataframe(df.head(20), use_container_width=True)

        upload_time = datetime.now()
        st.caption(f"زمان آپلود/شروع محاسبه: {upload_time.strftime('%Y/%m/%d %H:%M:%S')}")

        result_df, err = process_file(df, upload_time)
        if err:
            st.warning(err)
        else:
            st.subheader("خلاصه خوانده‌نشده‌ها")
            st.dataframe(result_df, use_container_width=True)

            # دانلود خروجی
            output = BytesIO()
            result_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "دانلود خروجی Excel",
                data=output,
                file_name="unread_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"خطا در خواندن/پردازش فایل: {e}")
else:
    st.info("یک فایل Excel آپلود کنید تا پردازش انجام شود.")
