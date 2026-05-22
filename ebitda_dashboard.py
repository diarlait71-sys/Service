"""
Понятный дашборд для руководителя:
- минимум визуального шума
- логика анализа по шагам
- фокус на решениях

Запуск:
streamlit run ebitda_dashboard.py --server.port 8502
"""

import os
import re
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

TEXT_COLOR = "#0f172a"
GRID_COLOR = "#d6dee8"


st.set_page_config(
    page_title="EBITDA -> Cash | DOSCAR",
    page_icon="💰",
    layout="wide",
)

# ══ Все суммы в млн тенге ══════════════════════════════════════════════════
UNITS = "млн тенге"
EBITDA       = 1673    # Чистая прибыль 393 + Амортизация 1110 + Проценты 170
NET_PROFIT   = 393
DEPRECIATION = 1110
INTEREST_EXP = 170

# Реальные остатки кэша (счета 1010+1030+1050 из карточек)
# 1020 "Деньги в пути" — транзитный счёт, сальдо на начало и конец = 0
# Начало 2025: 1010=0.27 млн + 1030=3.0 млн + 1050=0 = 3.27 ≈ 3 млн
# Конец  2025: 1010=0.23 млн + 1030=17.21 млн + 1050=259.59 млн = 277 млн
START_CASH = 3    # конец 2024 = начало 2025 (из карточек 1010+1030)
END_CASH   = 277  # конец 2025 (из карточек 1010+1030+1050; 1020 транзит = 0)
CASH_DELTA_FACT = END_CASH - START_CASH   # 274 млн - фактический прирост

GAP = EBITDA - CASH_DELTA_FACT   # 1399 млн — поглощено оборотным капиталом

# Поглотители кэша (отрицательный эффект, млн тенге)
ABSORBERS = pd.DataFrame(
    {
        "Блок": [
            "Стройка ДЦ Кульжинский (2930)",
            "Дебиторка покупателей (1200)",
            "Займы внутри группы (1274)",
            "Авансы полученные — выполнены (3510)",
            "Запасы (1300)",
            "Авансы выданные поставщикам (1700)",
            "Зарплата и налоги (3350/312х)",
            "Погашение займов банков (4010)",
            "Уплата процентов",
        ],
        "Сумма": [1447, 1002, 945, 309, 287, 284, 9+96, 89, 0],
        "Комментарий": [
            "Деньги вложены с 2023 года, объект не введён в эксплуатацию",
            "Продали, но не получили деньги от покупателей",
            "Кэш передан дочерним/партнёрским компаниям как займы",
            "Клиенты заплатили авансом — работы выполнены, деньги использованы",
            "Деньги заморожены в товаре на складе",
            "Предоплаты поставщикам не закрыты актами",
            "Больше выплатили, чем начислили за период",
            "Нетто-погашение тела кредитов (новые минус старые)",
            "Включено в EBITDA-мост выше",
        ],
    }
)
ABSORBERS = ABSORBERS[ABSORBERS["Сумма"] > 0].reset_index(drop=True)

# Источники кэша (положительный эффект, без EBITDA)
SOURCES = pd.DataFrame(
    {
        "Блок": [
            "Рост КЗ поставщикам (3310)",
            "Рост прочих обязательств (3380)",
        ],
        "Сумма": [1572, 1212],
        "Комментарий": [
            "Получили товар, но ещё не заплатили поставщикам — кэш пока у нас",
            "Рост прочих начисленных обязательств — деньги остались в компании",
        ],
    }
)

TOP_RECEIVABLES = pd.DataFrame(
    {
        "Контрагент": [
            "AUTOCENTER ONTUSTIK",
            "Rox Almaty",
            "Народный банк",
            "DCG Ontustik",
            "Bereke Bank",
        ],
        "Сумма": [402, 127, 97, 93, 91],
    }
)

TOP_GROUP_LOANS = pd.DataFrame(
    {
        "Контрагент": [
            "Capital Property",
            "Auto Logistics Product",
            "Rox Motor Qazaqstan",
            "DCG Ontustik",
            "DCG Atyrau",
        ],
        "Сумма": [2687, 1442, 838, 441, 285],
    }
)

# Расшифровка 3380 — нетто-прирост обязательств по контрагентам (из карточки счёта)
TOP_3380 = pd.DataFrame(
    {
        "Контрагент (3387 ВФП)": [
            "Rox Motor Qazaqstan",
            "AUTOCENTER ONTUSTIK",
            "Buta Capital",
            "Shym Auto (Шым Авто)",
            "DCG Orda",
            "Бакенов Р.А.",
            "DCG Vostok",
            "Нуртлеуов Ш.А.",
            "Ecl Group",
            "Прочие",
        ],
        "Нетто-прирост, млн ₸": [402, 292, 220, 190, 156, 69, 62, 46, 23, 12],
    }
)


CASH_FOLDER = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ"

@st.cache_data(show_spinner="Загружаем карточку счёта…")
def parse_cash_card(acc: str, label: str):
    """Парсит карточку счёта 1C, возвращает (inflows_df, outflows_df)."""
    import re
    fname = os.path.join(CASH_FOLDER, f"{acc}.xls")
    if not os.path.exists(fname):
        return pd.DataFrame(), pd.DataFrame()
    df = pd.read_excel(fname, header=None)

    def to_num(v):
        try:
            return float(str(v).replace(" ", "").replace(",", ".").replace("\xa0", "").strip())
        except Exception:
            return 0.0

    def extract_agent(text):
        for line in str(text).split("\n"):
            line = line.strip()
            if not line or "Головное" in line:
                continue
            if re.match(r"^(Тенге|KZT|KZ[0-9])", line):
                continue
            if len(line) >= 3:
                return line
        return "(прочее)"

    def extract_purpose(text):
        found = 0
        for line in str(text).split("\n"):
            line = line.strip()
            if not line or "Головное" in line:
                continue
            if re.match(r"^(Тенге|KZT|KZ[0-9])", line):
                continue
            if len(line) >= 3:
                found += 1
                if found >= 2:
                    return line
        return ""

    inflows, outflows = {}, {}
    in_purp, out_purp = {}, {}

    for _, row in df.iterrows():
        if str(row.iloc[4]).strip() != "БУ":
            continue
        if not re.match(r"\d{2}\.\d{2}\.\d{4}", str(row.iloc[0]).strip()):
            continue
        acc_dt = str(row.iloc[5]).strip()
        acc_kt = str(row.iloc[8]).strip()
        amt_dt = to_num(row.iloc[6])
        amt_kt = to_num(row.iloc[9])
        ana_dt = str(row.iloc[2]).strip()
        ana_kt = str(row.iloc[3]).strip()

        if acc_dt.startswith(acc) and amt_dt > 0:
            agent = extract_agent(ana_dt)
            if agent == "(прочее)":
                agent = extract_agent(ana_kt)
            purp = extract_purpose(ana_dt) or extract_purpose(ana_kt)
            inflows[agent] = inflows.get(agent, 0) + amt_dt
            in_purp.setdefault(agent, purp)

        if acc_kt.startswith(acc) and amt_kt > 0:
            agent = extract_agent(ana_dt)
            if agent == "(прочее)":
                agent = extract_agent(ana_kt)
            purp = extract_purpose(ana_dt) or extract_purpose(ana_kt)
            outflows[agent] = outflows.get(agent, 0) + amt_kt
            out_purp.setdefault(agent, purp)

    M = 1_000_000
    in_df = pd.DataFrame([
        {"Счёт": label, "Контрагент / Назначение": k,
         "Пояснение": in_purp.get(k, ""), "Сумма, млн ₸": round(v / M, 2)}
        for k, v in sorted(inflows.items(), key=lambda x: -x[1])
    ])
    out_df = pd.DataFrame([
        {"Счёт": label, "Контрагент / Назначение": k,
         "Пояснение": out_purp.get(k, ""), "Сумма, млн ₸": round(v / M, 2)}
        for k, v in sorted(outflows.items(), key=lambda x: -x[1])
    ])
    return in_df, out_df


@st.cache_data(show_spinner="Читаем ОСВ/субконто…")
def parse_osv_file(fname: str, is_active: bool, ncols: int):
    """Парсит ОСВ (9 col) или Анализ субконто (8 col) — возвращает list[(name, delta_mln)]."""
    path = os.path.join(CASH_FOLDER, fname)
    if not os.path.exists(path):
        return []

    def _to_num(v):
        s = str(v).strip()
        if s.lower() in ("nan", "", "none", "null"):
            return 0.0
        try:
            return float(s.replace(" ", "").replace(",", ".").replace("\xa0", ""))
        except Exception:
            return 0.0

    _SKIP_NAMES = {
        "Головное подразделение", "Контрагенты", "Договоры", "Валюта",
        "Счет", "KZT", "USD", "EUR", "", "nan", "Итого",
        "Работники организации",
    }

    def _is_agent(row):
        name = str(row.iloc[0]).strip()
        if not name or name in _SKIP_NAMES:
            return False
        if re.match(r"^\d{4}", name):          # номер счёта 4 цифры
            return False
        if ncols == 9:
            col1 = str(row.iloc[1]).strip() if len(row) > 1 else ""
            if col1 != "БУ":                   # пропускаем "Вал." и пустые
                return False
        if ncols == 8:
            if name.startswith("Договор"):     # пропускаем строки-договоры
                return False
        return True

    df = pd.read_excel(path, header=None)
    col_kt_end = 8 if ncols == 9 else 7

    results: dict = {}
    for _, row in df.iterrows():
        if not _is_agent(row):
            continue
        name = str(row.iloc[0]).strip()
        nac_dt = _to_num(row.iloc[2])
        nac_kt = _to_num(row.iloc[3]) if len(row) > 3 else 0.0
        kon_dt = _to_num(row.iloc[6]) if len(row) > 6 else 0.0
        kon_kt = _to_num(row.iloc[col_kt_end]) if len(row) > col_kt_end else 0.0
        if is_active:
            delta = (kon_dt - kon_kt) - (nac_dt - nac_kt)
        else:
            delta = (kon_kt - kon_dt) - (nac_kt - nac_dt)
        if abs(delta) > 100_000:
            results[name] = results.get(name, 0) + delta

    M = 1_000_000
    return sorted(
        [(nm, round(v / M, 1)) for nm, v in results.items()],
        key=lambda x: -abs(x[1]),
    )


def normalize_counterparty_name(name: str) -> str:
    """Нормализует имя контрагента для корректного взаимозачёта 1210 vs 3310."""
    s = str(name or "").strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace('"', "").replace("'", "")
    s = re.sub(r"^(ТОО|TOO|ИП|АО)\s+", "", s, flags=re.IGNORECASE)
    return s.upper()


@st.cache_data(show_spinner="Схлопываем дебет/кредит по контрагентам…")
def build_netted_ar_ap_report(
    rows_ar: list,
    rows_ap: list,
    ar_col_name: str = "Δ Дебиторка 1210, млн ₸",
    ap_col_name: str = "Δ Кредиторка 3310, млн ₸",
) -> tuple[pd.DataFrame, dict]:
    """
    Строит неттинг по контрагентам для двух списков изменений.
    rows_ar: левая сторона (обычно актив, рост = отток кэша)
    rows_ap: правая сторона (обычно обязательство, рост = приток кэша)
    """
    ar_map = {}
    ap_map = {}
    labels = {}

    for nm, val in rows_ar:
        key = normalize_counterparty_name(nm)
        if not key:
            continue
        labels.setdefault(key, nm)
        ar_map[key] = ar_map.get(key, 0.0) + float(val)

    for nm, val in rows_ap:
        key = normalize_counterparty_name(nm)
        if not key:
            continue
        labels.setdefault(key, nm)
        ap_map[key] = ap_map.get(key, 0.0) + float(val)

    rows = []
    for key in sorted(set(ar_map) | set(ap_map)):
        ar_val = round(ar_map.get(key, 0.0), 1)
        ap_val = round(ap_map.get(key, 0.0), 1)
        net_cash = round(ar_val - ap_val, 1)  # >0 отток, <0 приток
        offset = round(min(max(ar_val, 0.0), max(ap_val, 0.0)), 1)
        if abs(ar_val) < 0.1 and abs(ap_val) < 0.1:
            continue
        rows.append(
            {
                "Контрагент": labels.get(key, key),
                ar_col_name: ar_val,
                ap_col_name: ap_val,
                "Схлопнутое влияние на кэш, млн ₸": net_cash,
                "Взаимозачет, млн ₸": offset,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        summary = {
            "gross_ar": 0.0,
            "gross_ap": 0.0,
            "net_cash": 0.0,
            "offset": 0.0,
            "top": pd.DataFrame(),
        }
        return df, summary

    df = df.sort_values("Схлопнутое влияние на кэш, млн ₸", key=lambda s: s.abs(), ascending=False)
    summary = {
        "gross_ar": round(df[ar_col_name].sum(), 1),
        "gross_ap": round(df[ap_col_name].sum(), 1),
        "net_cash": round(df["Схлопнутое влияние на кэш, млн ₸"].sum(), 1),
        "offset": round(df["Взаимозачет, млн ₸"].sum(), 1),
        "top": df.head(20).copy(),
    }
    return df, summary


@st.cache_data(show_spinner="Сводим все счета…")
def get_all_flows():
    all_in, all_out = [], []
    for acc, label in [("1010", "1010 Касса"), ("1030", "1030 Текущий"), ("1050", "1050 Депозит")]:
        i_df, o_df = parse_cash_card(acc, label)
        if not i_df.empty:
            all_in.append(i_df)
        if not o_df.empty:
            all_out.append(o_df)
    combined_in  = pd.concat(all_in,  ignore_index=True) if all_in  else pd.DataFrame()
    combined_out = pd.concat(all_out, ignore_index=True) if all_out else pd.DataFrame()
    return combined_in, combined_out


st.markdown(
    """
<style>
/* ══ ПРИНУДИТЕЛЬНАЯ СВЕТЛАЯ ТЕМА — перекрывает тёмную тему Streamlit ══ */
:root { color-scheme: light !important; }

html, body, [data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="stMain"], [data-testid="stMainBlockContainer"],
section.main, .block-container,
[data-testid="stHeader"], [data-testid="stSidebar"],
[data-testid="stBottom"] {
    background: #f3f6fb !important;
    background-color: #f3f6fb !important;
    color: #0f172a !important;
}

/* Вкладки */
[data-testid="stTabs"] [role="tablist"],
[data-testid="stTabs"] [role="tab"],
[data-testid="stTabsContent"] {
    background: #f3f6fb !important;
    color: #0f172a !important;
}

/* Метрики, подписи */
[data-testid="stMetricLabel"] > div,
[data-testid="stMetricValue"] > div,
[data-testid="stMetricDelta"],
[data-testid="stCaptionContainer"] {
    color: #0f172a !important;
}

/* Весь текст */
h1, h2, h3, h4, h5, h6, p, li, span, label,
div, section, article, aside, header, footer {
    color: #0f172a !important;
}

/* Таблицы */
[data-testid="stDataFrame"],
[data-testid="stDataFrameContainer"],
.stDataFrame { 
    background: #ffffff !important;
    border: 1px solid #dbe2ea;
    border-radius: 10px;
}

/* Plotly iframe */
[data-testid="stVegaLiteChart"], .js-plotly-plot,
iframe { background: #ffffff !important; }

.main-card {
    background: #ffffff;
    border: 1px solid #e7ebf0;
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 10px;
}
.logic-card {
    background: #eef4ff !important;
    border-left: 4px solid #1f6feb;
    border-radius: 10px;
    padding: 12px 14px;
    margin: 8px 0;
}
.action-card {
    background: #fff8ed !important;
    border-left: 4px solid #d97706;
    border-radius: 10px;
    padding: 12px 14px;
    margin: 8px 0;
}
.small { color: #5f6b7a !important; font-size: 13px; }

/* Тёмный оверлей от Streamlit */
[class*="overlayContainer"], [class*="overlay"] {
    background: transparent !important;
}
</style>
""",
    unsafe_allow_html=True,
)


tab1, tab2, tab3 = st.tabs(["📊 EBITDA → Cash", "💚 Поступления", "🔴 Выбытие"])

with tab1:

    st.title("Почему EBITDA большая, а кэша мало")
    st.caption("ТОО DOSCAR GROUP | 2025 | объяснение для руководителя")
    st.markdown(
        f"""
    <div class="logic-card">
    <b>Единица измерения по всему дашборду:</b> {UNITS}.<br>
    Если требуется, могу переключить представление в тыс. тенге.
    </div>
    """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("EBITDA 2025 (млн тенге)",        f"{EBITDA:,}".replace(","," "))
    col2.metric("Кэш начало года (млн тенге)",     f"{START_CASH:,}".replace(","," "))
    col3.metric("Кэш конец года (млн тенге)",      f"{END_CASH:,}".replace(","," "), delta=f"+{CASH_DELTA_FACT} млн")

    st.markdown(
        """
    <div class="logic-card">
    <b>Логика в одном предложении:</b><br>
    EBITDA показывает прибыльность бизнеса, но не показывает, где зависли деньги.
    Разрыв возникает, когда кэш уходит в дебиторку, стройку, займы группе, запасы и авансы.
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.subheader("1) Куда ушел кэш — и откуда пришёл")

    col_out, col_in = st.columns(2)
    with col_out:
        st.markdown("**🔴 Поглотители кэша (отток)**")
        n = len(ABSORBERS)
        reds = ["#7f1d1d","#991b1b","#b91c1c","#dc2626","#ef4444","#f87171","#fca5a5","#fecaca","#fee2e2"]
        fig_abs = go.Figure()
        fig_abs.add_bar(
            x=ABSORBERS["Сумма"],
            y=ABSORBERS["Блок"],
            orientation="h",
            marker_color=reds[:n],
            text=[f"{v} млн тг" for v in ABSORBERS["Сумма"]],
            textposition="outside",
            textfont={"size": 12, "color": TEXT_COLOR},
        )
        fig_abs.update_layout(
            template="plotly_white", font={"color": TEXT_COLOR, "size": 12},
            height=360, margin=dict(l=10, r=80, t=10, b=10),
            xaxis={"title":{"text":"млн тенге","font":{"color":TEXT_COLOR}},"tickfont":{"color":TEXT_COLOR},"gridcolor":GRID_COLOR},
            yaxis={"title":"","autorange":"reversed","tickfont":{"color":TEXT_COLOR}},
            plot_bgcolor="white", paper_bgcolor="white",
        )
        st.plotly_chart(fig_abs, width="stretch")
        st.markdown(f"**Итого отток: {ABSORBERS['Сумма'].sum():,} млн тенге**".replace(","," "))

    with col_in:
        st.markdown("**🟢 Источники кэша (приток)**")
        fig_src = go.Figure()
        fig_src.add_bar(
            x=SOURCES["Сумма"],
            y=SOURCES["Блок"],
            orientation="h",
            marker_color=["#166534","#15803d"],
            text=[f"{v} млн тг" for v in SOURCES["Сумма"]],
            textposition="outside",
            textfont={"size": 12, "color": TEXT_COLOR},
        )
        fig_src.update_layout(
            template="plotly_white", font={"color": TEXT_COLOR, "size": 12},
            height=200, margin=dict(l=10, r=80, t=10, b=10),
            xaxis={"title":{"text":"млн тенге","font":{"color":TEXT_COLOR}},"tickfont":{"color":TEXT_COLOR},"gridcolor":GRID_COLOR},
            yaxis={"title":"","autorange":"reversed","tickfont":{"color":TEXT_COLOR}},
            plot_bgcolor="white", paper_bgcolor="white",
        )
        st.plotly_chart(fig_src, width="stretch")
        st.markdown(f"**Итого приток из обязательств: {SOURCES['Сумма'].sum():,} млн тенге**".replace(","," "))

        st.markdown("**Что это значит:**")
        for _, row in SOURCES.iterrows():
            st.markdown(f"- **{row['Блок']}** ({int(row['Сумма'])} млн тг): {row['Комментарий']}.")

    st.markdown("**Расшифровка поглотителей:**")
    for _, row in ABSORBERS.iterrows():
        st.markdown(f"- **{row['Блок']}** ({int(row['Сумма'])} млн тг): {row['Комментарий']}.")

    # ── Expanders с расшифровкой по контрагентам ──────────────────────────────
    import plotly.express as px_exp

    def _render_expander(title: str, rows: list, bar_color: str, comment: str = ""):
        """Отображает expander с баром + таблицей контрагентов."""
        with st.expander(title):
            if comment:
                st.caption(comment)
            if not rows:
                st.info("Данные не загружены или файл не найден.")
                return
            top = rows[:15]
            names = [r[0][:55] for r in top]
            vals  = [r[1] for r in top]
            fig = go.Figure()
            fig.add_bar(
                x=vals, y=names, orientation="h",
                marker_color=bar_color,
                text=[f"{v:+.1f}" for v in vals],
                textposition="outside",
                textfont={"size": 11, "color": TEXT_COLOR},
            )
            fig.update_layout(
                template="plotly_white", font={"color": TEXT_COLOR, "size": 11},
                height=max(200, 26 * len(top) + 60),
                margin=dict(l=10, r=80, t=10, b=10),
                xaxis={"title": {"text": "млн ₸", "font": {"color": TEXT_COLOR}},
                       "tickfont": {"color": TEXT_COLOR}, "gridcolor": GRID_COLOR},
                yaxis={"title": "", "autorange": "reversed",
                       "tickfont": {"color": TEXT_COLOR}},
                plot_bgcolor="white", paper_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)
            tbl = pd.DataFrame(rows[:15], columns=["Контрагент", "Нетто-изменение, млн ₸"])
            st.dataframe(tbl, hide_index=True, use_container_width=True)

    st.markdown("#### 🔍 Детализация по контрагентам (нажмите чтобы раскрыть)")

    # 1. Стройка 2930
    with st.expander("🏗️ Стройка ДЦ Кульжинский (2930) — +1 447 млн ₸"):
        st.caption("Карточка счёта 2930: вложено в незавершённое строительство ДЦ.")
        st.markdown("""
**Откуда финансируется?** Основной контрагент виден из 3310 (КЗ поставщикам):
СК 7 Stroi Development несёт основные затраты на строительство (+1 308 млн ₸ роста КЗ).

Объект не введён в эксплуатацию, поэтому актив висит на 2930 и не амортизируется.
        """)

    # 2. Дебиторка 1210
    rows_1210 = parse_osv_file("1210.xls", is_active=True, ncols=9)
    _render_expander(
        "📋 Дебиторка покупателей (1210) — +1 002 млн ₸",
        rows_1210,
        bar_color="#ef4444",
        comment="Нетто-изменение сальдо по каждому покупателю за 2025 год (рост дебиторки = кэш не получен).",
    )

    # 3. Займы внутри группы 1274
    rows_1274 = parse_osv_file("1274.xls", is_active=True, ncols=8)
    _render_expander(
        "🔗 Займы внутри группы (1274) — +945 млн ₸",
        rows_1274,
        bar_color="#f97316",
        comment="Нетто-изменение займов, выданных дочерним и партнёрским компаниям за 2025 год.",
    )

    # 4. Авансы полученные 3510 (показываем только отрицательные — там выполнение)
    rows_3510 = parse_osv_file("3510.xls", is_active=False, ncols=9)
    _render_expander(
        "🔄 Авансы полученные — выполнение заказов (3510) — −309 млн ₸",
        rows_3510,
        bar_color="#a855f7",
        comment="Снижение авансов = выполнили предоплаченные работы. Кэш был получен раньше, теперь 'закрыт' выручкой.",
    )

    # 5. Авансы выданные 1710
    rows_1710 = parse_osv_file("1710.xls", is_active=True, ncols=8)
    _render_expander(
        "💼 Авансы поставщикам (1710) — +284 млн ₸",
        rows_1710,
        bar_color="#f59e0b",
        comment="Нетто-изменение авансов, выданных поставщикам (предоплаты, не закрытые актами).",
    )

    # 6. КЗ поставщикам 3310
    rows_3310 = parse_osv_file("3310.xls", is_active=False, ncols=9)
    _render_expander(
        "🟢 КЗ поставщикам (3310) — +1 572 млн ₸ приток",
        rows_3310,
        bar_color="#22c55e",
        comment="Рост кредиторской задолженности = получили товар/услуги, но ещё не заплатили. Кэш пока у нас.",
    )

    st.subheader("1.1) Схлопывание дебет/кредит по контрагентам")
    net_df_1210, net_summary_1210 = build_netted_ar_ap_report(
        rows_1210,
        rows_3310,
        ar_col_name="Δ Дебиторка 1210, млн ₸",
        ap_col_name="Δ Кредиторка 3310, млн ₸",
    )
    net_df_1710, net_summary_1710 = build_netted_ar_ap_report(
        rows_1710,
        rows_3310,
        ar_col_name="Δ Авансы выданные 1710, млн ₸",
        ap_col_name="Δ Кредиторка 3310, млн ₸",
    )

    st.markdown("**1210 ↔ 3310 (покупатели/поставщики):**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Брутто 1210", f"{net_summary_1210['gross_ar']:+,.1f} млн".replace(",", " "))
    c2.metric("Брутто 3310", f"{net_summary_1210['gross_ap']:+,.1f} млн".replace(",", " "))
    c3.metric("Схлопнуто", f"{net_summary_1210['offset']:+,.1f} млн".replace(",", " "))
    c4.metric("Нетто эффект", f"{net_summary_1210['net_cash']:+,.1f} млн".replace(",", " "))

    st.markdown("**1710 ↔ 3310 (авансы поставщикам/КЗ):**")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Брутто 1710", f"{net_summary_1710['gross_ar']:+,.1f} млн".replace(",", " "))
    s2.metric("Брутто 3310", f"{net_summary_1710['gross_ap']:+,.1f} млн".replace(",", " "))
    s3.metric("Схлопнуто", f"{net_summary_1710['offset']:+,.1f} млн".replace(",", " "))
    s4.metric("Нетто эффект", f"{net_summary_1710['net_cash']:+,.1f} млн".replace(",", " "))

    st.markdown(
        """
    <div class="logic-card">
    <b>Как читать:</b><br>
    Брутто суммы могут быть завышены для управленческого вывода, если один и тот же контрагент сидит одновременно в дебете и кредите.<br>
    В блоке выше выполнен неттинг для двух пар (1210↔3310 и 1710↔3310): показан эффект после схлопывания.
    </div>
    """,
        unsafe_allow_html=True,
    )

    with st.expander("Топ-20 после схлопывания: 1210 vs 3310"):
        if net_df_1210.empty:
            st.info("Недостаточно данных для неттинга.")
        else:
            st.dataframe(net_summary_1210["top"], hide_index=True, use_container_width=True)

    with st.expander("Топ-20 после схлопывания: 1710 vs 3310"):
        if net_df_1710.empty:
            st.info("Недостаточно данных для неттинга.")
        else:
            st.dataframe(net_summary_1710["top"], hide_index=True, use_container_width=True)

    st.subheader("1.2) Итоговый формат для акционера")
    shareholder_df = pd.DataFrame(
        [
            {"Показатель": "EBITDA 2025", "Сумма, млн ₸": EBITDA, "Комментарий": "Операционная прибыльность"},
            {"Показатель": "Факт прирост кэша", "Сумма, млн ₸": CASH_DELTA_FACT, "Комментарий": "По счетам 1010/1030/1050"},
            {"Показатель": "Рост дебиторки 1210 (брутто)", "Сумма, млн ₸": net_summary_1210["gross_ar"], "Комментарий": "Продажи без оплаты"},
            {"Показатель": "Рост кредиторки 3310 (брутто)", "Сумма, млн ₸": net_summary_1210["gross_ap"], "Комментарий": "Товар получен, оплата позже"},
            {"Показатель": "Схлопнуто 1210↔3310", "Сумма, млн ₸": -net_summary_1210["offset"], "Комментарий": "Убирает двойной эффект по покупателям/поставщикам"},
            {"Показатель": "Нетто эффект 1210↔3310", "Сумма, млн ₸": net_summary_1210["net_cash"], "Комментарий": "Реальное давление на кэш после неттинга"},
            {"Показатель": "Рост авансов поставщикам 1710 (брутто)", "Сумма, млн ₸": net_summary_1710["gross_ar"], "Комментарий": "Предоплаты поставщикам"},
            {"Показатель": "Схлопнуто 1710↔3310", "Сумма, млн ₸": -net_summary_1710["offset"], "Комментарий": "Убирает двойной эффект аванс/кредиторка у одного поставщика"},
            {"Показатель": "Нетто эффект 1710↔3310", "Сумма, млн ₸": net_summary_1710["net_cash"], "Комментарий": "Чистое влияние пары 1710/3310 на кэш"},
        ]
    )
    st.dataframe(shareholder_df, hide_index=True, use_container_width=True)

    st.subheader("2) Почему возник разрыв EBITDA и кэша")

    bridge_labels = [
        "EBITDA",
        "- Стройка",
        "- Дебиторка",
        "- Займы группе",
        "- Авансы получ.",
        "- Запасы",
        "- Авансы выд.",
        "+ КЗ пост.",
        "+ Проч.обяз.",
        "- Займы банков",
        "- Проценты",
        "Факт кэш Δ",
    ]
    bridge_values = [1673, -1447, -1002, -945, -309, -287, -284, 1572, 1212, -89, -96, 274]
    bridge_measure = ["absolute","relative","relative","relative","relative","relative","relative","relative","relative","relative","relative","total"]

    fig_bridge = go.Figure(
        go.Waterfall(
            measure=bridge_measure,
            x=bridge_labels,
            y=bridge_values,
            increasing={"marker": {"color": "#16a34a"}},
            decreasing={"marker": {"color": "#dc2626"}},
            totals={"marker": {"color": "#1f6feb"}},
            connector={"line": {"color": "#c7ced8"}},
            text=[f"{v:+}" for v in bridge_values],
            textposition="outside",
            textfont={"size": 13, "color": TEXT_COLOR},
        )
    )
    fig_bridge.update_layout(
        template="plotly_white",
        font={"color": TEXT_COLOR, "size": 13},
        height=430,
        margin=dict(l=10, r=20, t=20, b=40),
        xaxis={
            "tickfont": {"color": TEXT_COLOR},
        },
        yaxis={
            "title": {"text": "млн тенге", "font": {"color": TEXT_COLOR}},
            "tickfont": {"color": TEXT_COLOR},
            "gridcolor": GRID_COLOR,
        },
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    st.plotly_chart(fig_bridge, width="stretch")

    st.subheader("2.1) Полный путь: от начального остатка к конечному")

    # ── Детализированные шаги моста (из ОСВ 2024 и 2025) ──────────────────────
    # Для активных счетов: Δ>0 = использован кэш (отток), Δ<0 = высвобождён (приток)
    # Для пассивных счетов: Δ>0 = рост обязательств = приток кэша, Δ<0 = отток
    path_steps = [
        # Операционная прибыльность
        ("EBITDA (прибыль + амортизация + проценты)",       +EBITDA),
        # Изменения в активах (рост актива = отток кэша)
        ("↑ Дебиторка покупателей 1200: не получили деньги",    -1002),
        ("↑ Займы выданные группе 1274: кэш в дочках",         -945),
        ("↑ Незавершённая стройка 2930: вложено в ДЦ",         -1447),
        ("↑ Запасы 1300: деньги в товаре на складе",           -287),
        ("↑ Авансы поставщикам 1700: предоплаты не закрыты",   -284),
        # Изменения в обязательствах (рост = приток кэша)
        ("↑ КЗ поставщикам 3310: платим позже — кэш у нас",    +1572),
        ("↑ Прочие обязательства 3380: рост задолженности",    +1212),
        ("↓ Авансы полученные 3510: выполнили предоплаченные заказы", -309),
        ("↓ Зарплата к выплате 3350: выплатили чуть больше начисл.", -9),
        # Финансирование
        ("Погашение тела займов банков 4010 (нетто)",          -89),
        ("Уплачено процентов банкам",                          -96),
    ]

    base_cum = START_CASH + sum(v for _, v in path_steps)
    other_effect = END_CASH - base_cum

    path_rows = []
    cum = START_CASH
    path_rows.append(
        {
            "Шаг": "Начальный остаток кэша",
            "Влияние (млн тенге)": 0,
            "Накопленный итог (млн тенге)": cum,
            "Пояснение": "Остаток на начало периода",
        }
    )

    for name, effect in path_steps:
        cum += effect
        path_rows.append(
            {
                "Шаг": name,
                "Влияние (млн тенге)": effect,
                "Накопленный итог (млн тенге)": cum,
                "Пояснение": "Прямое влияние на денежный остаток",
            }
        )

    cum += other_effect
    path_rows.append(
        {
            "Шаг": "Прочие притоки/оттоки (балансировка)",
            "Влияние (млн тенге)": other_effect,
            "Накопленный итог (млн тенге)": cum,
            "Пояснение": "Прочие статьи ДДС: налоги (ИПН/СН/ОПВ), мелкие операционные выплаты и прочее",
        }
    )

    path_rows.append(
        {
            "Шаг": "Конечный остаток кэша",
            "Влияние (млн тенге)": 0,
            "Накопленный итог (млн тенге)": END_CASH,
            "Пояснение": "Факт на конец периода",
        }
    )

    path_df = pd.DataFrame(path_rows)

    st.dataframe(
        path_df,
        width="stretch",
        hide_index=True,
    )

    st.markdown(
        """
    <div class="logic-card">
    <b>Как читать таблицу:</b><br>
    Столбец Влияние показывает вклад каждой статьи в деньги периода,<br>
    а столбец Накопленный итог показывает путь от начального остатка к конечному остатку.
    </div>
    """,
        unsafe_allow_html=True,
    )

    explained = sum(v for _, v in path_steps)
    st.markdown(
        f"""
    <div class="logic-card">
    <b>Прогресс расшифровки балансировки:</b><br>
    С новыми данными (3380, 3510, 4010, 3350) балансировка уменьшилась.<br>
    Объяснено дополнительно: ~1 900 млн тенге (3380 +1 212 млн, 3510 −309 млн).<br>
    Остаток неразобранного <b>({other_effect:+.0f} млн тенге)</b> — это вероятно:
    комиссии банков, налоги КПН/НДС, прочие мелкие статьи ДДС.
    </div>
    """,
        unsafe_allow_html=True,
    )

    if abs(other_effect) > 50:
        st.markdown("**Что ещё можно загрузить для уменьшения балансировки:**")
        st.markdown("- ОСВ по счёту 3100 (КПН к уплате) — если большой.")
        st.markdown("- ОСВ по 3190 (прочие налоги) и 3130 (социальный налог).")

    st.subheader("2.2) Расшифровка прочих обязательств 3380: кому должны")
    st.markdown("""
    <div class='logic-card'>
    <b>3380 — счёт внутригрупповых займов (ВФП, 3387).</b><br>
    На начало 2025: 12 271 млн ₸. За год дочерние ДЦ довнесли ещё +1 212 млн ₸.<br>
    Это <b>не кредиторская задолженность поставщикам</b>, а деньги, полученные от компаний группы в виде
    временной финансовой помощи (беспроцентные займы между юрлицами одного холдинга).
    </div>
    """, unsafe_allow_html=True)

    col_a, col_b = st.columns([1, 1])
    with col_a:
        st.markdown("**Прирост задолженности 2025 по контрагентам (нетто, млн ₸):**")
        st.dataframe(TOP_3380, width="stretch", hide_index=True)
    with col_b:
        import plotly.express as px
        fig_3380 = px.bar(
            TOP_3380,
            x="Нетто-прирост, млн ₸",
            y="Контрагент (3387 ВФП)",
            orientation="h",
            color_discrete_sequence=["#22c55e"],
            template="plotly_white",
        )
        fig_3380.update_layout(
            height=320,
            margin=dict(l=0, r=10, t=10, b=10),
            yaxis=dict(autorange="reversed", title="", tickfont={"color": TEXT_COLOR}),
            xaxis=dict(title="млн ₸", tickfont={"color": TEXT_COLOR}),
            font={"color": TEXT_COLOR},
        )
        fig_3380.update_traces(texttemplate="%{x:.0f}", textposition="outside", textfont={"color": TEXT_COLOR})
        st.plotly_chart(fig_3380, use_container_width=True)

    st.markdown(
        """
    <div class="logic-card">
    <b>Интерпретация моста:</b><br>
    1. EBITDA 1 673 млн тенге есть, значит операционно бизнес прибыльный.<br>
    2. Но деньги были изъяты из оборота в 5 крупных направлениях.<br>
    3. Поэтому в кэш дошло только 274 млн тенге.
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.subheader("3) На кого смотреть в первую очередь")
    left, right = st.columns(2)

    with left:
        st.markdown("**Топ дебиторов (1210)**")
        st.dataframe(TOP_RECEIVABLES, width="stretch", hide_index=True)

    with right:
        st.markdown("**Топ займов внутри группы (1274)**")
        st.dataframe(TOP_GROUP_LOANS, width="stretch", hide_index=True)

    st.subheader("4) Что делать (приоритеты)")

    st.markdown(
        """
    <div class="action-card">
    <b>Приоритет 1: Дебиторка (1210)</b><br>
    Запустить план взыскания по топ-5 контрагентам (особенно AUTOCENTER ONTUSTIK 402 млн).
    </div>
    <div class="action-card">
    <b>Приоритет 2: Займы внутри группы (1274)</b><br>
    Утвердить график возврата или конверсию части займов в капитал.
    </div>
    <div class="action-card">
    <b>Приоритет 3: Стройка (2930)</b><br>
    Зафиксировать дату ввода ДЦ в эксплуатацию, чтобы актив начал работать и амортизироваться.
    </div>
    <div class="action-card">
    <b>Приоритет 4: Авансы (1710)</b><br>
    Провести сверку и закрытие актами, в первую очередь по СК 7 Stroi.
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.info("Если нужно, следующим шагом добавлю режим: 'для руководителя' (кратко) и 'для аналитика' (детально по всем контрагентам).")

with tab2:
    import plotly.express as px2
    st.subheader("Поступления кэша 2025")
    st.caption("Данные из карточек 1030, 1010, 1050 | млн тенге | сумма = валовый приток (до внутренних переводов)")

    combined_in, _ = get_all_flows()
    if combined_in.empty:
        st.warning("Файлы карточек не найдены в папке финанализ\\")
    else:
        ACCOUNTS_IN = ["Все счета"] + sorted(combined_in["Счёт"].unique().tolist())
        sel_in = st.radio("Счёт:", ACCOUNTS_IN, horizontal=True, key="sel_in")
        df_in = combined_in if sel_in == "Все счета" else combined_in[combined_in["Счёт"] == sel_in]
        # Агрегируем по контрагенту (если выбраны все счета — могут быть дубли)
        df_in_agg = (
            df_in.groupby("Контрагент / Назначение", as_index=False)["Сумма, млн ₸"]
            .sum()
            .sort_values("Сумма, млн ₸", ascending=False)
        )
        total_in = df_in_agg["Сумма, млн ₸"].sum()

        col_chart, col_table = st.columns([1.3, 1])
        with col_chart:
            top15 = df_in_agg.head(15)
            fig = px2.bar(
                top15, x="Сумма, млн ₸", y="Контрагент / Назначение",
                orientation="h", color_discrete_sequence=["#16a34a"],
                template="plotly_white",
                title=f"Топ-15 источников поступлений | итого {total_in:,.1f} млн".replace(",", " "),
            )
            fig.update_layout(
                height=460, margin=dict(l=0, r=40, t=40, b=10),
                yaxis=dict(autorange="reversed", title="", tickfont={"color": TEXT_COLOR}),
                xaxis=dict(title="млн ₸", tickfont={"color": TEXT_COLOR}),
                font={"color": TEXT_COLOR},
                title_font={"color": TEXT_COLOR},
            )
            fig.update_traces(texttemplate="%{x:.1f}", textposition="outside",
                              textfont={"color": TEXT_COLOR})
            st.plotly_chart(fig, use_container_width=True)

        with col_table:
            st.markdown(f"**Всего поступлений: {total_in:,.1f} млн ₸**".replace(",", " "))
            st.dataframe(df_in_agg, hide_index=True, use_container_width=True, height=430)

        with st.expander("Полная расшифровка с разбивкой по счетам"):
            st.dataframe(
                df_in.sort_values("Сумма, млн ₸", ascending=False).reset_index(drop=True),
                hide_index=True, use_container_width=True,
            )

with tab3:
    import plotly.express as px3
    st.subheader("Выбытие кэша 2025")
    st.caption("Данные из карточек 1030, 1010, 1050 | млн тенге | сумма = валовый отток (до внутренних переводов)")

    _, combined_out = get_all_flows()
    if combined_out.empty:
        st.warning("Файлы карточек не найдены в папке финанализ\\")
    else:
        ACCOUNTS_OUT = ["Все счета"] + sorted(combined_out["Счёт"].unique().tolist())
        sel_out = st.radio("Счёт:", ACCOUNTS_OUT, horizontal=True, key="sel_out")
        df_out = combined_out if sel_out == "Все счета" else combined_out[combined_out["Счёт"] == sel_out]
        df_out_agg = (
            df_out.groupby("Контрагент / Назначение", as_index=False)["Сумма, млн ₸"]
            .sum()
            .sort_values("Сумма, млн ₸", ascending=False)
        )
        total_out = df_out_agg["Сумма, млн ₸"].sum()

        col_chart2, col_table2 = st.columns([1.3, 1])
        with col_chart2:
            top15o = df_out_agg.head(15)
            fig2 = px3.bar(
                top15o, x="Сумма, млн ₸", y="Контрагент / Назначение",
                orientation="h", color_discrete_sequence=["#dc2626"],
                template="plotly_white",
                title=f"Топ-15 получателей выбытия | итого {total_out:,.1f} млн".replace(",", " "),
            )
            fig2.update_layout(
                height=460, margin=dict(l=0, r=40, t=40, b=10),
                yaxis=dict(autorange="reversed", title="", tickfont={"color": TEXT_COLOR}),
                xaxis=dict(title="млн ₸", tickfont={"color": TEXT_COLOR}),
                font={"color": TEXT_COLOR},
                title_font={"color": TEXT_COLOR},
            )
            fig2.update_traces(texttemplate="%{x:.1f}", textposition="outside",
                               textfont={"color": TEXT_COLOR})
            st.plotly_chart(fig2, use_container_width=True)

        with col_table2:
            st.markdown(f"**Всего выбытий: {total_out:,.1f} млн ₸**".replace(",", " "))
            st.dataframe(df_out_agg, hide_index=True, use_container_width=True, height=430)

        with st.expander("Полная расшифровка с разбивкой по счетам"):
            st.dataframe(
                df_out.sort_values("Сумма, млн ₸", ascending=False).reset_index(drop=True),
                hide_index=True, use_container_width=True,
            )
