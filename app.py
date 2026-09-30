import io
import sqlite3
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


DATABASE_PATH = Path(__file__).with_name("overtime_submissions.db")
INTERNAL_PASSWORD = "wxgcc&lxh"
RATE_COLUMNS = {
    150: "平时工时(150%)",
    200: "周末工时(200%)",
    300: "法定节假日工时(300%)",
}

st.set_page_config(page_title="加班登记", layout="wide")


@st.cache_data(ttl=86400)
def get_holiday_safe(date_str):
    url = f"https://timor.tech/api/holiday/info/{date_str}"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }
    try:
        response = requests.get(url, headers=headers, timeout=8)
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, ValueError) as error:
        raise RuntimeError(
            f"无法查询 {date_str} 的法定节假日信息：{error}"
        ) from error

    if not isinstance(data, dict) or data.get("code") != 0:
        raise RuntimeError(f"节假日接口未能查询 {date_str}（接口返回：{data}）")
    return data


def get_holiday_type_accurate(date_str):
    data = get_holiday_safe(date_str)
    type_info = data.get("type") or {}
    holiday_info = data.get("holiday") or {}

    if holiday_info.get("wage") == 3:
        return "法定节假日", 300

    if type_info.get("type") == 3 or (
        holiday_info and not holiday_info.get("holiday", True)
    ):
        return "工作日", 150

    if holiday_info.get("holiday") is True or type_info.get("type") == 1:
        return "休息日", 200

    return "工作日", 150


def calculate_overtime(start, end):
    if end <= start:
        raise ValueError("加班结束时间必须晚于起始时间。")

    hours_by_rate = {rate: 0.0 for rate in RATE_COLUMNS}
    current = start
    while current < end:
        next_midnight = datetime.combine(
            current.date() + timedelta(days=1), time.min
        )
        segment_end = min(end, next_midnight)
        _, rate = get_holiday_type_accurate(current.date().isoformat())
        hours_by_rate[rate] += (segment_end - current).total_seconds() / 3600
        current = segment_end

    return {
        column: round(hours_by_rate[rate], 2)
        for rate, column in RATE_COLUMNS.items()
    }


def initialize_database():
    with sqlite3.connect(DATABASE_PATH) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS overtime_submissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                department TEXT NOT NULL,
                start_datetime TEXT NOT NULL,
                end_datetime TEXT NOT NULL,
                content TEXT NOT NULL,
                weekday_hours REAL NOT NULL,
                weekend_hours REAL NOT NULL,
                holiday_hours REAL NOT NULL,
                total_hours REAL NOT NULL,
                submitted_at TEXT NOT NULL
            )
            """
        )
        columns = [
            row[1]
            for row in connection.execute("PRAGMA table_info(overtime_submissions)").fetchall()
        ]
        if "department" not in columns:
            connection.execute(
                "ALTER TABLE overtime_submissions ADD COLUMN department TEXT NOT NULL DEFAULT ''"
            )


def save_submission(name, department, start, end, content, hours):
    total_hours = round(sum(hours.values()), 2)
    with sqlite3.connect(DATABASE_PATH) as connection:
        connection.execute(
            """
            INSERT INTO overtime_submissions (
                name, department, start_datetime, end_datetime, content,
                weekday_hours, weekend_hours, holiday_hours,
                total_hours, submitted_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                department,
                start.isoformat(sep=" "),
                end.isoformat(sep=" "),
                content,
                hours[RATE_COLUMNS[150]],
                hours[RATE_COLUMNS[200]],
                hours[RATE_COLUMNS[300]],
                total_hours,
                datetime.now().isoformat(sep=" ", timespec="seconds"),
            ),
        )


def load_submissions():
    with sqlite3.connect(DATABASE_PATH) as connection:
        return pd.read_sql_query(
            """
            SELECT
                id AS 编号,
                name AS 姓名,
                department AS 部门,
                start_datetime AS 加班起始时间,
                end_datetime AS 加班结束时间,
                content AS 加班内容,
                weekday_hours AS 平时工时_150,
                weekend_hours AS 周末工时_200,
                holiday_hours AS 法定节假日工时_300,
                total_hours AS 总时长_小时,
                submitted_at AS 提交时间
            FROM overtime_submissions
            ORDER BY id DESC
            """,
            connection,
        )


def make_excel_file(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="加班登记")
    return output.getvalue()


initialize_database()

st.title("加班登记")
st.caption("请填写加班信息并提交。")

page = st.sidebar.radio("页面", ["客户填写", "内部管理"])

if page == "客户填写":
    st.subheader("加班信息")
    default_start = datetime.now().replace(minute=0, second=0, microsecond=0)
    default_end = default_start + timedelta(hours=1)

    with st.form("overtime_submission_form", clear_on_submit=True):
        name = st.text_input("姓名", placeholder="李小明 / 张三")
        department = st.text_input("部门", placeholder="维修一队 / 综合办公室")
        start_col, end_col = st.columns(2)
        with start_col:
            st.markdown("**加班起始时间**")
            start_date = st.date_input("起始日期", value=default_start.date())
            start_hour = st.selectbox(
                "起始整点",
                options=list(range(24)),
                index=default_start.hour,
                format_func=lambda hour: f"{hour:02d}:00",
            )
        with end_col:
            st.markdown("**加班结束时间**")
            end_date = st.date_input("结束日期", value=default_end.date())
            end_hour = st.selectbox(
                "结束整点",
                options=list(range(24)),
                index=default_end.hour,
                format_func=lambda hour: f"{hour:02d}:00",
            )
        content = st.text_area("加班内容", placeholder="东湖路DN400闸门更换")
        submitted = st.form_submit_button("提交")

    if submitted:
        if not name.strip():
            st.error("请填写姓名。")
        elif not department.strip():
            st.error("请填写部门。")
        elif not content.strip():
            st.error("请填写加班内容。")
        else:
            start = datetime.combine(start_date, time(start_hour))
            end = datetime.combine(end_date, time(end_hour))
            if end <= start:
                st.error("加班结束时间必须晚于起始时间。")
            else:
                try:
                    calculated_hours = calculate_overtime(start, end)
                    save_submission(
                        name.strip(),
                        department.strip(),
                        start,
                        end,
                        content.strip(),
                        calculated_hours,
                    )
                except RuntimeError as error:
                    st.error(str(error))
                except sqlite3.Error as error:
                    st.error(f"保存加班登记失败：{error}")
                else:
                    st.success("提交成功，感谢填写。")
else:
    st.subheader("内部管理页")
    internal_password = st.sidebar.text_input("内部访问口令", type="password")
    if internal_password != INTERNAL_PASSWORD:
        st.warning("这是内部页面，仅授权人员可访问。")
        st.stop()

    submissions = load_submissions()
    if submissions.empty:
        st.info("暂时没有已提交的加班登记。")
    else:
        selection = submissions.copy()
        selection.insert(0, "导出", False)
        edited_selection = st.data_editor(
            selection,
            hide_index=True,
            width="stretch",
            disabled=[column for column in submissions.columns],
            column_config={
                "导出": st.column_config.CheckboxColumn(
                    "导出", help="勾选需要导出的记录"
                )
            },
            key="submission_export_selection",
        )
        selected_ids = edited_selection.loc[
            edited_selection["导出"], "编号"
        ].tolist()
        selected_rows = submissions[submissions["编号"].isin(selected_ids)]
        selected_col, all_col = st.columns(2)
        with selected_col:
            if not selected_rows.empty:
                st.download_button(
                    "导出勾选记录",
                    data=make_excel_file(selected_rows),
                    file_name=f"加班登记_勾选_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )
            else:
                st.button("导出勾选记录", disabled=True)
        with all_col:
            st.download_button(
                "一键导出全部",
                data=make_excel_file(submissions),
                file_name=f"加班登记_全部_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
