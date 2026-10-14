import io
import os
import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="План семейного бюджета",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main {padding-top: 1rem;}

    /* Карточки метрик: фон и текст берут цвета из активной темы,
       поэтому числа видно и в светлой, и в тёмной теме. */
    div[data-testid="stMetric"] {
        background: var(--secondary-background-color);
        background: linear-gradient(
            135deg,
            color-mix(in srgb, var(--primary-color) 16%, var(--secondary-background-color)) 0%,
            var(--secondary-background-color) 100%
        );
        border: 1px solid color-mix(in srgb, var(--text-color) 14%, transparent);
        border-radius: 12px;
        padding: 0.9rem 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetricLabel"],
    div[data-testid="stMetricLabel"] * {
        color: var(--text-color) !important;
        opacity: 0.75;
    }
    div[data-testid="stMetricValue"],
    div[data-testid="stMetricValue"] > div,
    div[data-testid="stMetricValue"] * {
        color: var(--text-color) !important;
    }
    div[data-testid="stMetricDelta"] { color: var(--text-color) !important; }

    .stDataFrame {border-radius: 10px; overflow: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "план_семейного_бюджета.xlsx")
EDITED_SHEET = "План_по_датам_отредактированный"

MONTHS_RU = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
    7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}

MONEY_COLS = [
    "Поступление", "На ипотеку", "На ЖКУ", "На еду", "На билеты", "На отпуск",
    "Свободный резерв", "План отложить", "Факт отложено", "Факт на билеты", "Факт на отпуск",
]
SUMMARY_MONEY_COLS = [
    "Доход", "Ипотека", "ЖКУ", "Еда", "Билеты", "Отпуск",
    "Свободный резерв", "Накоплено на билеты", "Накоплено на отпуск",
]


@st.cache_data
def load_excel(path):
    """Читаем все листы без приведения типов, чтобы даты остались датами."""
    try:
        wb = load_workbook(path, data_only=True)
        sheets = {}
        for s in wb.sheetnames:
            df = pd.read_excel(path, sheet_name=s, header=None, dtype=object, keep_default_na=False)
            sheets[s] = df
        return sheets, wb.sheetnames
    except Exception as e:
        st.error(f"Ошибка загрузки файла: {e}")
        return None, []


def to_float(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or x == "":
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("\xa0", "").replace(" ", "").replace("₽", "").replace("р.", "").replace("р", "").replace("Р", "")
    s = s.replace(",", ".")
    m = re.findall(r"-?\d+\.?\d*", s)
    return float(m[0]) if m else 0.0


def find_header(df, required):
    """Ищем строку-заголовок, содержащую все нужные подписи."""
    for i in range(min(len(df), 10)):
        vals = [str(v) for v in df.iloc[i].values if str(v).strip() not in ("", "nan", "None")]
        joined = " ".join(vals)
        if all(word in joined for word in required):
            return i
    return None


def extract_table(df, required):
    """Возвращает (df_данных, заголовки) начиная со строки-заголовка."""
    h = find_header(df, required)
    if h is None:
        return None, None
    header = [str(c).strip() if str(c).strip() not in ("", "nan", "None") else f"col_{j}"
              for j, c in enumerate(df.iloc[h].values)]
    body = df.iloc[h + 1:].copy()
    body.columns = header
    # убираем полностью пустые строки
    body = body[body.apply(lambda r: any(str(v).strip() not in ("", "nan", "None") for v in r.values), axis=1)]
    return body.reset_index(drop=True), header


def fmt_money(v):
    return f"{v:,.0f} ₽".replace(",", " ")


def fmt_date(d):
    if pd.isna(d):
        return ""
    return f"{d.day:02d}.{d.month:02d}.{d.year}"


def _add_sheet(wb, sheet_name, df):
    """Создаёт (или заменяет) лист с данными и базовым оформлением."""
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, PatternFill
    from datetime import datetime as _dt

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    rows = dataframe_to_rows(df, index=False, header=True)
    for r_idx, row in enumerate(rows, 1):
        for c_idx, val in enumerate(row, 1):
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4472C4")
            elif isinstance(val, _dt):
                cell.number_format = "DD.MM.YYYY"
            elif isinstance(val, (int, float)):
                cell.number_format = '#,##0" ₽"'
    return ws


def can_write(path):
    """Можно ли писать рядом с файлом? В облаке файловая система только для чтения."""
    try:
        with open(path, "a+b"):
            pass
        return True
    except OSError:
        return False


def edited_workbook_bytes(path, sheet_name, df):
    """Собирает Excel-копию с новым листом в памяти, диск не трогает."""
    try:
        wb = load_workbook(path)
        _add_sheet(wb, sheet_name, df)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    except Exception as e:
        st.error(f"Ошибка подготовки файла: {e}")
        return None


def save_sheet_to_excel(path, sheet_name, df):
    """Сохраняет лист прямо в файл. Работает только там, где есть запись."""
    try:
        wb = load_workbook(path)
        _add_sheet(wb, sheet_name, df)
        wb.save(path)
        return True
    except Exception as e:
        st.error(f"Ошибка сохранения: {e}")
        return False


st.title("💸 План семейного бюджета")
st.caption("Удобная обёртка для файла «план_семейного_бюджета.xlsx»")

sheets_data, sheetnames = load_excel(EXCEL_FILE)
if not sheets_data:
    st.stop()

# --- готовим таблицы заранее ---
plan_df, _ = extract_table(sheets_data.get("План по датам", pd.DataFrame()), ["Дата", "Поступление"])
if plan_df is not None:
    plan_df["Дата"] = pd.to_datetime(plan_df["Дата"], errors="coerce")
    plan_df = plan_df.dropna(subset=["Дата"]).sort_values("Дата").reset_index(drop=True)
    for c in MONEY_COLS + ["Выполнение"]:
        if c in plan_df.columns:
            plan_df[c] = plan_df[c].apply(to_float)

summary_df, _ = extract_table(sheets_data.get("Сводка", pd.DataFrame()), ["Месяц", "Доход"])
if summary_df is not None:
    # оставляем только таблицу по месяцам, отбрасываем служебные колонки
    keep = ["Месяц"] + [c for c in SUMMARY_MONEY_COLS if c in summary_df.columns]
    summary_df = summary_df[keep].copy()
    summary_df = summary_df[summary_df["Месяц"].astype(str).str.strip().ne("")].reset_index(drop=True)
    for c in SUMMARY_MONEY_COLS:
        if c in summary_df.columns:
            summary_df[c] = summary_df[c].apply(to_float)

tabs = st.tabs(["📊 Дашборд", "📅 План по датам", "📈 Сводка", "⚙️ Параметры", "ℹ️ О файле"])

# ================= TAB 1: DASHBOARD =================
with tabs[0]:
    st.subheader("Общая картина")
    if plan_df is None:
        st.warning("Не удалось найти данные на листе «План по датам».")
    else:
        months_all = plan_df["Дата"].dt.to_period("M").astype(str).unique().tolist()
        col1, col2 = st.columns([1, 3])
        with col1:
            sel_month = st.selectbox("Месяц", ["Все месяцы"] + months_all)
        with col2:
            st.write("")

        df_f = plan_df if sel_month == "Все месяцы" else plan_df[plan_df["Дата"].dt.to_period("M").astype(str) == sel_month]

        if len(df_f) == 0:
            st.info("Нет записей за выбранный период.")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Поступлений", f"{len(df_f)}")
            c2.metric("Сумма поступлений", fmt_money(df_f["Поступление"].sum()))
            c3.metric("План отложить", fmt_money(df_f["План отложить"].sum()))
            c4.metric("Факт отложено", fmt_money(df_f["Факт отложено"].sum()))

            c5, c6, c7, c8 = st.columns(4)
            c5.metric("На билеты (план)", fmt_money(df_f["На билеты"].sum()))
            c6.metric("Факт на билеты", fmt_money(df_f["Факт на билеты"].sum()))
            c7.metric("На отпуск (план)", fmt_money(df_f["На отпуск"].sum()))
            c8.metric("Факт на отпуск", fmt_money(df_f["Факт на отпуск"].sum()))

            st.divider()
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Распределение по назначению (план)**")
                agg = pd.DataFrame({
                    "Категория": ["Ипотека", "ЖКУ", "Еда", "Билеты", "Отпуск", "Свободный резерв"],
                    "Сумма": [df_f[c].sum() for c in ["На ипотеку", "На ЖКУ", "На еду", "На билеты", "На отпуск", "Свободный резерв"]],
                })
                if agg["Сумма"].sum() > 0:
                    fig = px.pie(agg, values="Сумма", names="Категория", hole=0.45,
                                 color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig.update_traces(textposition="inside", textinfo="percent+label")
                    st.plotly_chart(fig, width="stretch")
                else:
                    st.info("Нет данных для диаграммы.")

            with col_b:
                st.markdown("**Факт vs План: билеты и отпуск**")
                fact_agg = pd.DataFrame({
                    "Категория": ["Билеты", "Отпуск"],
                    "План": [df_f["На билеты"].sum(), df_f["На отпуск"].sum()],
                    "Факт": [df_f["Факт на билеты"].sum(), df_f["Факт на отпуск"].sum()],
                })
                fig2 = go.Figure(data=[
                    go.Bar(name="План", x=fact_agg["Категория"], y=fact_agg["План"], marker_color="#667eea"),
                    go.Bar(name="Факт", x=fact_agg["Категория"], y=fact_agg["Факт"], marker_color="#43e97b"),
                ])
                fig2.update_layout(barmode="group", xaxis_title=None, yaxis_title="Сумма, ₽")
                st.plotly_chart(fig2, width="stretch")

            st.divider()
            st.markdown("**График поступлений**")
            fig3 = px.bar(df_f, x="Дата", y="Поступление", color="Поступление",
                          color_continuous_scale="Blues", text="Поступление")
            fig3.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig3.update_layout(xaxis_title=None, yaxis_title="₽", coloraxis_showscale=False, xaxis_tickformat="%d.%m.%Y")
            st.plotly_chart(fig3, width="stretch")

# ================= TAB 2: PLAN BY DATE =================
with tabs[1]:
    st.subheader("План по датам")
    if plan_df is None:
        st.warning("Данные не найдены.")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            statuses = ["Все"]
            if "Статус" in plan_df.columns:
                statuses += sorted({str(x) for x in plan_df["Статус"].unique() if str(x).strip() not in ("", "nan", "None")})
            sel_status = st.selectbox("Статус", statuses)
        with col2:
            months_m = plan_df["Дата"].dt.to_period("M").astype(str).unique().tolist()
            sel_m = st.selectbox("Месяц", ["Все месяцы"] + months_m, key="plan_month")
        with col3:
            st.write("")

        df_f2 = plan_df.copy()
        if sel_m != "Все месяцы":
            df_f2 = df_f2[df_f2["Дата"].dt.to_period("M").astype(str) == sel_m]
        if sel_status != "Все" and "Статус" in df_f2.columns:
            df_f2 = df_f2[df_f2["Статус"] == sel_status]
        df_f2 = df_f2.reset_index(drop=True)

        st.info(
            "💡 Таблицу можно редактировать прямо здесь. Нажмите «Сохранить изменения»: "
            "при локальном запуске правки допишутся в файл Excel, а в облаке (где диск только "
            "для чтения) вы получите копию файла на скачивание."
        )

        col_cfg = {}
        for c in df_f2.columns:
            if c == "Дата":
                col_cfg[c] = st.column_config.DateColumn("Дата", format="DD.MM.YYYY", width="small")
            elif c == "Выполнение":
                col_cfg[c] = st.column_config.NumberColumn("Выполнение", format="%.2f", width="small")
            elif c in MONEY_COLS:
                col_cfg[c] = st.column_config.NumberColumn(c, format="%.0f ₽")
            elif c == "Комментарий":
                col_cfg[c] = st.column_config.TextColumn("Комментарий", width="large")

        edited = st.data_editor(df_f2, width="stretch", num_rows="dynamic", hide_index=True,
                                column_config=col_cfg, key="plan_editor")

        csum1, csum2, csum3 = st.columns([1, 1, 2])
        csum1.metric("Поступлений", fmt_money(edited["Поступление"].sum()))
        csum2.metric("План отложить", fmt_money(edited["План отложить"].sum()))
        csum3.metric("Факт отложено", fmt_money(edited["Факт отложено"].sum()))

        if st.button("💾 Сохранить изменения", type="primary"):
            if can_write(EXCEL_FILE):
                if save_sheet_to_excel(EXCEL_FILE, EDITED_SHEET, edited):
                    st.cache_data.clear()
                    st.success(f"Сохранено в файл → лист «{EDITED_SHEET}».")
                    st.session_state.pop("edited_xlsx", None)
            else:
                data = edited_workbook_bytes(EXCEL_FILE, EDITED_SHEET, edited)
                if data is not None:
                    st.session_state["edited_xlsx"] = data

        payload = st.session_state.get("edited_xlsx")
        if payload:
            st.warning(
                "Диск доступен только для чтения (это норма для облачного хостинга). "
                "Скачайте готовый файл с листом изменений:"
            )
            st.download_button(
                "⬇️ Скачать Excel с изменениями",
                data=payload,
                file_name="план_семейного_бюджета_изменённый.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ================= TAB 3: SUMMARY =================
with tabs[2]:
    st.subheader("Сводка бюджета")
    if summary_df is None:
        st.warning("Не удалось найти данные на листе «Сводка».")
    else:
        col_cfg_s = {}
        for c in summary_df.columns:
            if c in SUMMARY_MONEY_COLS:
                col_cfg_s[c] = st.column_config.NumberColumn(c, format="%.0f ₽")
        st.dataframe(summary_df, width="stretch", hide_index=True, column_config=col_cfg_s)

        plot_cols = [c for c in ["Доход", "Ипотека", "ЖКУ", "Еда", "Билеты", "Отпуск"] if c in summary_df.columns]
        if "Месяц" in summary_df.columns and plot_cols:
            dp = summary_df[summary_df["Месяц"].astype(str).str.strip() != ""].copy()
            fig_s = px.bar(dp, x="Месяц", y=plot_cols, barmode="group", title="Доходы и расходы по месяцам")
            fig_s.update_layout(xaxis_title=None, yaxis_title="₽", legend_title=None)
            st.plotly_chart(fig_s, width="stretch")

# ================= TAB 4: PARAMS =================
with tabs[3]:
    st.subheader("Параметры семейного бюджета")
    raw = sheets_data.get("Параметры", pd.DataFrame())
    if len(raw) == 0:
        st.warning("Нет данных.")
    else:
        raw_disp = raw.map(lambda v: "" if str(v).strip() in ("", "nan", "None") else str(v))
        st.dataframe(raw_disp, width="stretch", hide_index=True)
        st.caption("Это исходный лист «Параметры» как он есть в Excel.")

# ================= TAB 5: ABOUT =================
with tabs[4]:
    st.subheader("О файле")
    st.write(f"**Файл:** `{os.path.basename(EXCEL_FILE)}`")
    st.write(f"**Размер:** {os.path.getsize(EXCEL_FILE) / 1024:.1f} КБ")
    st.write(f"**Листы:** {', '.join(sheetnames)}")
    st.write("Обёртка создана для удобного просмотра, редактирования и анализа семейного бюджета.")
    st.markdown("**Запуск локально**")
    st.code("""
# Запуск
cd ~/code/family-budget-app
./run.sh

# или напрямую
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
""")
    st.markdown("**Запуск в облаке**")
    st.caption(
        "Приложение работает и на Streamlit Community Cloud. Там диск доступен только "
        "для чтения, поэтому изменения не дописываются в исходный Excel, а скачиваются "
        "отдельным файлом."
    )
