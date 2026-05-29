"""Biometric Payroll Streamlit application.

Run with:
    streamlit run app.py
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from io import BytesIO
from typing import Iterable

import pandas as pd
import streamlit as st


APP_TITLE = "Biometric Payroll Pro"
DEFAULT_SHIFT_START = time(9, 0)
DEFAULT_SHIFT_END = time(18, 0)
DEFAULT_GRACE_MINUTES = 15
DEFAULT_MONTHLY_SALARY = 25_000.0
DEFAULT_WORKING_DAYS = 26


@dataclass(frozen=True)
class PayrollSettings:
    shift_start: time
    shift_end: time
    grace_minutes: int
    half_day_hours: float
    full_day_hours: float
    overtime_after_hours: float
    monthly_salary: float
    working_days: int
    overtime_multiplier: float


def page_config() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🧾",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def inject_css() -> None:
    st.markdown(
        """
        <style>
            .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
            .hero {
                background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 55%, #06b6d4 100%);
                padding: 1.6rem 1.8rem;
                border-radius: 1.2rem;
                color: white;
                box-shadow: 0 18px 45px rgba(15, 23, 42, 0.16);
                margin-bottom: 1.2rem;
            }
            .hero h1 {margin: 0; font-size: 2.15rem; letter-spacing: -0.04em;}
            .hero p {margin: 0.45rem 0 0; color: #dbeafe; font-size: 1.02rem;}
            .metric-card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 1rem;
                padding: 1rem;
                box-shadow: 0 10px 25px rgba(15, 23, 42, 0.06);
            }
            .section-title {font-weight: 700; color: #0f172a; margin-top: .5rem;}
            div[data-testid="stDataFrame"] {border: 1px solid #e5e7eb; border-radius: .8rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        f"""
        <div class="hero">
            <h1>🧾 {APP_TITLE}</h1>
            <p>Upload biometric attendance from Excel, calculate work hours, overtime, deductions, and export a payroll-ready report.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def parse_time_value(value: object) -> time | None:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, (float, int)):
        # Excel can store time as a fraction of one day.
        if 0 <= float(value) < 1:
            seconds = int(round(float(value) * 24 * 60 * 60))
            return (datetime.min + timedelta(seconds=seconds)).time()
    parsed = pd.to_datetime(str(value).strip(), errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.time().replace(microsecond=0)


def parse_date_value(value: object) -> date | None:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(str(value).strip(), dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def load_attendance(uploaded_file: BytesIO, has_header: bool) -> pd.DataFrame:
    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]
    header = 0 if has_header else None
    if suffix in {"xlsx", "xls"}:
        raw = pd.read_excel(uploaded_file, header=header)
    elif suffix == "csv":
        raw = pd.read_csv(uploaded_file, header=header)
    else:
        raise ValueError("Please upload an Excel (.xlsx/.xls) or CSV file.")

    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.shape[1] < 3:
        raise ValueError("The attendance file must contain at least Employee ID, Date, and Punch Time columns.")

    # The screenshot shows the first three useful columns: ID, date, time.
    attendance = raw.iloc[:, :3].copy()
    attendance.columns = ["employee_id", "punch_date", "punch_time"]
    attendance["employee_id"] = attendance["employee_id"].astype(str).str.strip()
    attendance["punch_date"] = attendance["punch_date"].apply(parse_date_value)
    attendance["punch_time"] = attendance["punch_time"].apply(parse_time_value)
    attendance = attendance.dropna(subset=["employee_id", "punch_date", "punch_time"])
    attendance = attendance[attendance["employee_id"].str.lower().ne("nan")]

    if attendance.empty:
        raise ValueError("No valid punch rows were found. Check that the first three columns match ID, date, and time.")

    attendance["punch_datetime"] = attendance.apply(
        lambda row: datetime.combine(row["punch_date"], row["punch_time"]), axis=1
    )
    return attendance.sort_values(["employee_id", "punch_datetime"]).reset_index(drop=True)


def load_employee_master(uploaded_file: BytesIO | None) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(columns=["employee_id", "employee_name", "monthly_salary"])

    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]
    if suffix in {"xlsx", "xls"}:
        master = pd.read_excel(uploaded_file)
    elif suffix == "csv":
        master = pd.read_csv(uploaded_file)
    else:
        raise ValueError("Please upload an Excel (.xlsx/.xls) or CSV employee master file.")

    normalized = {str(column).strip().lower().replace(" ", "_"): column for column in master.columns}
    id_column = normalized.get("employee_id") or normalized.get("emp_id") or normalized.get("id")
    name_column = normalized.get("employee_name") or normalized.get("name")
    salary_column = normalized.get("monthly_salary") or normalized.get("salary") or normalized.get("gross_salary")

    if not id_column:
        raise ValueError("Employee master must include an employee_id, emp_id, or id column.")

    result = pd.DataFrame()
    result["employee_id"] = master[id_column].astype(str).str.strip()
    result["employee_name"] = master[name_column].astype(str).str.strip() if name_column else result["employee_id"]
    result["monthly_salary"] = pd.to_numeric(master[salary_column], errors="coerce") if salary_column else pd.NA
    return result.dropna(subset=["employee_id"]).drop_duplicates("employee_id")


def build_daily_attendance(attendance: pd.DataFrame, settings: PayrollSettings) -> pd.DataFrame:
    grouped = attendance.groupby(["employee_id", "punch_date"], as_index=False).agg(
        first_in=("punch_datetime", "min"),
        last_out=("punch_datetime", "max"),
        punch_count=("punch_datetime", "count"),
    )
    grouped["work_hours"] = (grouped["last_out"] - grouped["first_in"]).dt.total_seconds().div(3600).round(2)
    grouped.loc[grouped["punch_count"].eq(1), "work_hours"] = 0.0

    shift_start_dt = grouped["punch_date"].apply(lambda d: datetime.combine(d, settings.shift_start))
    grace_cutoff = shift_start_dt + pd.to_timedelta(settings.grace_minutes, unit="m")
    grouped["late_minutes"] = ((grouped["first_in"] - grace_cutoff).dt.total_seconds().div(60)).clip(lower=0).round(0)
    grouped["overtime_hours"] = (grouped["work_hours"] - settings.overtime_after_hours).clip(lower=0).round(2)

    grouped["attendance_status"] = "Absent"
    grouped.loc[grouped["work_hours"].ge(settings.half_day_hours), "attendance_status"] = "Half Day"
    grouped.loc[grouped["work_hours"].ge(settings.full_day_hours), "attendance_status"] = "Present"
    grouped.loc[grouped["punch_count"].eq(1), "attendance_status"] = "Single Punch"

    grouped["payable_days"] = 0.0
    grouped.loc[grouped["attendance_status"].eq("Half Day"), "payable_days"] = 0.5
    grouped.loc[grouped["attendance_status"].eq("Present"), "payable_days"] = 1.0
    grouped["date"] = grouped["punch_date"].astype(str)
    grouped["first_in"] = grouped["first_in"].dt.strftime("%H:%M:%S")
    grouped["last_out"] = grouped["last_out"].dt.strftime("%H:%M:%S")
    return grouped[
        [
            "employee_id",
            "date",
            "first_in",
            "last_out",
            "punch_count",
            "work_hours",
            "late_minutes",
            "overtime_hours",
            "attendance_status",
            "payable_days",
        ]
    ]


def build_payroll(daily: pd.DataFrame, master: pd.DataFrame, settings: PayrollSettings) -> pd.DataFrame:
    payroll = daily.groupby("employee_id", as_index=False).agg(
        present_days=("payable_days", "sum"),
        calendar_days=("date", "nunique"),
        total_hours=("work_hours", "sum"),
        overtime_hours=("overtime_hours", "sum"),
        late_instances=("late_minutes", lambda values: int((values > 0).sum())),
        single_punch_days=("attendance_status", lambda values: int((values == "Single Punch").sum())),
    )
    if not master.empty:
        payroll = payroll.merge(master, on="employee_id", how="left")
    else:
        payroll["employee_name"] = payroll["employee_id"]
        payroll["monthly_salary"] = pd.NA

    payroll["employee_name"] = payroll["employee_name"].fillna(payroll["employee_id"])
    payroll["monthly_salary"] = pd.to_numeric(payroll["monthly_salary"], errors="coerce").fillna(settings.monthly_salary)
    payroll["per_day_salary"] = payroll["monthly_salary"] / settings.working_days
    payroll["basic_pay"] = payroll["present_days"] * payroll["per_day_salary"]
    payroll["overtime_pay"] = (
        payroll["overtime_hours"] * (payroll["per_day_salary"] / settings.full_day_hours) * settings.overtime_multiplier
    )
    payroll["gross_pay"] = payroll["basic_pay"] + payroll["overtime_pay"]
    payroll["attendance_deduction"] = payroll["monthly_salary"] - payroll["basic_pay"]
    money_columns = ["monthly_salary", "per_day_salary", "basic_pay", "overtime_pay", "gross_pay", "attendance_deduction"]
    payroll[money_columns] = payroll[money_columns].round(2)
    return payroll[
        [
            "employee_id",
            "employee_name",
            "calendar_days",
            "present_days",
            "total_hours",
            "overtime_hours",
            "late_instances",
            "single_punch_days",
            "monthly_salary",
            "basic_pay",
            "overtime_pay",
            "attendance_deduction",
            "gross_pay",
        ]
    ].sort_values("employee_id")


def to_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def render_sidebar() -> tuple[PayrollSettings, BytesIO | None, BytesIO | None, bool]:
    with st.sidebar:
        st.header("⚙️ Payroll Setup")
        st.caption("Configure salary and shift rules before processing attendance.")
        attendance_file = st.file_uploader("Attendance Excel/CSV", type=["xlsx", "xls", "csv"])
        has_header = st.checkbox("File has header row", value=False)
        employee_file = st.file_uploader("Employee master (optional)", type=["xlsx", "xls", "csv"])

        st.divider()
        shift_start = st.time_input("Shift start", value=DEFAULT_SHIFT_START)
        shift_end = st.time_input("Shift end", value=DEFAULT_SHIFT_END)
        grace_minutes = st.number_input("Grace period (minutes)", min_value=0, max_value=120, value=DEFAULT_GRACE_MINUTES)
        half_day_hours = st.number_input("Minimum hours for half day", min_value=1.0, max_value=12.0, value=4.0, step=0.5)
        full_day_hours = st.number_input("Minimum hours for present day", min_value=1.0, max_value=16.0, value=8.0, step=0.5)
        overtime_after_hours = st.number_input("Overtime starts after hours", min_value=1.0, max_value=16.0, value=8.0, step=0.5)

        st.divider()
        monthly_salary = st.number_input("Default monthly salary", min_value=0.0, value=DEFAULT_MONTHLY_SALARY, step=500.0)
        working_days = st.number_input("Salary working days", min_value=1, max_value=31, value=DEFAULT_WORKING_DAYS)
        overtime_multiplier = st.number_input("Overtime multiplier", min_value=0.0, max_value=5.0, value=1.5, step=0.25)

    settings = PayrollSettings(
        shift_start=shift_start,
        shift_end=shift_end,
        grace_minutes=int(grace_minutes),
        half_day_hours=float(half_day_hours),
        full_day_hours=float(full_day_hours),
        overtime_after_hours=float(overtime_after_hours),
        monthly_salary=float(monthly_salary),
        working_days=int(working_days),
        overtime_multiplier=float(overtime_multiplier),
    )
    return settings, attendance_file, employee_file, has_header


def render_empty_state() -> None:
    st.info("Upload an attendance file from the sidebar to generate payroll.")
    st.markdown("""
    **Expected attendance format** (like your screenshot):

    | Employee ID | Punch Date | Punch Time |
    | --- | --- | --- |
    | SKVFT027 | 01-03-2026 | 07:07:47 |
    | SKVFT048 | 01-03-2026 | 07:38:46 |

    The app reads the first three non-empty columns, so files without headers are supported.
    """)


def render_metrics(daily: pd.DataFrame, payroll: pd.DataFrame) -> None:
    total_employees = payroll["employee_id"].nunique()
    total_present_days = payroll["present_days"].sum()
    total_gross_pay = payroll["gross_pay"].sum()
    total_overtime = payroll["overtime_hours"].sum()
    cols = st.columns(4)
    metrics: Iterable[tuple[str, str]] = [
        ("Employees", f"{total_employees:,}"),
        ("Payable Days", f"{total_present_days:,.1f}"),
        ("Gross Payroll", f"₹{total_gross_pay:,.2f}"),
        ("Overtime Hours", f"{total_overtime:,.2f}"),
    ]
    for column, (label, value) in zip(cols, metrics):
        with column:
            st.metric(label, value)


def main() -> None:
    page_config()
    inject_css()
    render_header()
    settings, attendance_file, employee_file, has_header = render_sidebar()

    if attendance_file is None:
        render_empty_state()
        return

    try:
        attendance = load_attendance(attendance_file, has_header)
        master = load_employee_master(employee_file)
        daily = build_daily_attendance(attendance, settings)
        payroll = build_payroll(daily, master, settings)
    except Exception as exc:  # Streamlit should show a friendly upload/configuration error.
        st.error(f"Unable to process the uploaded file: {exc}")
        return

    render_metrics(daily, payroll)

    tabs = st.tabs(["Payroll Summary", "Daily Attendance", "Raw Punches", "Export"])
    with tabs[0]:
        st.subheader("Payroll Summary")
        st.dataframe(payroll, use_container_width=True, hide_index=True)
    with tabs[1]:
        st.subheader("Daily Attendance Register")
        status_filter = st.multiselect(
            "Filter by status",
            options=sorted(daily["attendance_status"].unique()),
            default=sorted(daily["attendance_status"].unique()),
        )
        filtered_daily = daily[daily["attendance_status"].isin(status_filter)] if status_filter else daily
        st.dataframe(filtered_daily, use_container_width=True, hide_index=True)
    with tabs[2]:
        st.subheader("Imported Biometric Punches")
        st.dataframe(
            attendance.assign(punch_datetime=attendance["punch_datetime"].dt.strftime("%Y-%m-%d %H:%M:%S")),
            use_container_width=True,
            hide_index=True,
        )
    with tabs[3]:
        st.subheader("Download Reports")
        report = to_excel({"Payroll Summary": payroll, "Daily Attendance": daily, "Raw Punches": attendance})
        st.download_button(
            "⬇️ Download payroll workbook",
            data=report,
            file_name=f"biometric_payroll_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Workbook includes payroll summary, daily attendance, and raw punch sheets.")


if __name__ == "__main__":
    main()
