from pathlib import Path
import os

import pandas as pd
import streamlit as st
from openai import OpenAI
from streamlit.errors import StreamlitSecretNotFoundError


st.set_page_config(page_title="Financial Sample Explorer", page_icon="$", layout="wide")


st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

    :root {
        --ink: #102a43;
        --muted: #627d98;
        --paper: #f4f7f8;
        --panel: #ffffff;
        --teal: #168c8c;
        --coral: #e9795e;
        --line: #d9e2ec;
    }

    .stApp {
        background: radial-gradient(circle at 90% 0%, #d8f1ed 0, transparent 28rem), var(--paper);
        color: var(--ink);
        font-family: 'DM Sans', sans-serif;
    }

    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stToolbar"] { right: 1rem; }
    .block-container { max-width: 1380px; padding: 2.6rem 3.2rem 4rem; }

    h1, h2, h3, [data-testid="stMetricLabel"] {
        font-family: 'Space Grotesk', sans-serif;
        color: var(--ink);
    }
    h1 { font-size: clamp(2.1rem, 4vw, 3.8rem); letter-spacing: 0; line-height: 1; margin-bottom: .45rem; }
    h2 { font-size: 1.55rem; margin-top: 1.8rem; }
    h3 { font-size: 1.1rem; }
    p, label, [data-testid="stCaptionContainer"] { color: var(--muted); }

    .hero {
        border-bottom: 1px solid var(--line);
        margin-bottom: 1.8rem;
        padding-bottom: 1.8rem;
        position: relative;
    }
    .hero::after {
        background: var(--coral);
        border-radius: 99px;
        content: '';
        height: 5px;
        left: 0;
        position: absolute;
        top: -1rem;
        width: 56px;
    }
    .eyebrow {
        color: var(--teal);
        font-size: .72rem;
        font-weight: 700;
        letter-spacing: .16em;
        text-transform: uppercase;
    }
    .hero p { font-size: 1rem; margin: 0; max-width: 660px; }

    [data-testid="stMetric"] {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 8px;
        box-shadow: 0 8px 24px rgba(16, 42, 67, .06);
        padding: 1.05rem 1.2rem;
    }
    [data-testid="stMetricLabel"] { font-size: .75rem; letter-spacing: .08em; text-transform: uppercase; }
    [data-testid="stMetricValue"] { color: var(--teal); font-family: 'Space Grotesk', sans-serif; }

    [data-testid="stSidebar"] { background: var(--ink); }
    [data-testid="stSidebar"] * { color: #d9e2ec; }
    [data-testid="stSidebar"] h2 { color: #ffffff; font-size: 1.15rem; }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background: rgba(255,255,255,.08); border-color: rgba(255,255,255,.25); }
    [data-testid="stSidebar"] [data-baseweb="select"] > div { background: rgba(255,255,255,.08); border-color: rgba(255,255,255,.2); }

    .section-kicker { color: var(--coral); font-size: .75rem; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; }
    .stTabs [data-baseweb="tab-list"] { gap: 1.4rem; border-bottom: 1px solid var(--line); }
    .stTabs [data-baseweb="tab"] { color: var(--muted); font-weight: 600; padding: .8rem 0; }
    .stTabs [aria-selected="true"] { color: var(--teal); }
    .stTabs [data-baseweb="tab-highlight"] { background: var(--coral); height: 3px; }
    [data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
    .stDownloadButton button { background: var(--ink); border: 0; color: #fff; }
    .stDownloadButton button:hover { background: var(--teal); color: #fff; }
    </style>
    """,
    unsafe_allow_html=True,
)


DEFAULT_FILE = Path.home() / "Downloads" / "Financial Sample.xlsx"


@st.cache_data
def load_workbook(source):
    return pd.read_excel(source, sheet_name=None)


def build_data_context(frame):
    numeric_fields = frame.select_dtypes(include="number").columns.tolist()
    profile = pd.DataFrame(
        {
            "column": frame.columns,
            "type": [str(frame[column].dtype) for column in frame.columns],
            "missing": [int(frame[column].isna().sum()) for column in frame.columns],
        }
    )
    context_parts = [
        f"Rows: {len(frame)}",
        f"Columns: {len(frame.columns)}",
        "Column profile:\n" + profile.to_string(index=False),
        "First 8 rows:\n" + frame.head(8).to_csv(index=False),
    ]
    if numeric_fields:
        context_parts.append(
            "Numeric summary:\n"
            + frame[numeric_fields].describe().round(2).to_string()
        )
    return "\n\n".join(context_parts)


def format_value(value):
    if pd.isna(value):
        return "-"
    if isinstance(value, (int, float)):
        return f"{value:,.2f}"
    return str(value)


st.markdown(
    """
    <div class="hero">
        <div class="eyebrow">Finance / Workbook intelligence / Aga's App</div>
        <h1>Financial Sample Explorer</h1>
        <p>A clear, considered view of the numbers hiding in your Excel workbook.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.sidebar.file_uploader("Upload an Excel workbook", type=["xlsx", "xls"])
source = uploaded_file if uploaded_file is not None else DEFAULT_FILE

st.sidebar.markdown("## Workbook desk")
st.sidebar.caption("Choose a worksheet or bring in a fresh workbook.")

if uploaded_file is None and not DEFAULT_FILE.exists():
    st.info("Upload an Excel workbook from the sidebar to get started.")
    st.stop()

try:
    sheets = load_workbook(source)
except Exception as error:
    st.error(f"Could not read the workbook: {error}")
    st.stop()

sheet_name = st.sidebar.selectbox("Worksheet", list(sheets))
data = sheets[sheet_name].copy()
data.columns = [str(column).strip() for column in data.columns]

if data.empty:
    st.warning("The selected worksheet is empty.")
    st.stop()

numeric_columns = data.select_dtypes(include="number").columns.tolist()
date_columns = data.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()
category_columns = [column for column in data.columns if column not in numeric_columns + date_columns]

st.markdown(f'<div class="section-kicker">Active worksheet</div><h2>{sheet_name}</h2>', unsafe_allow_html=True)
metric_columns = st.columns(4)
metric_columns[0].metric("Rows", f"{len(data):,}")
metric_columns[1].metric("Columns", f"{len(data.columns):,}")
metric_columns[2].metric("Numeric fields", f"{len(numeric_columns):,}")
metric_columns[3].metric("Missing values", f"{int(data.isna().sum().sum()):,}")

tab_summary, tab_data, tab_quality, tab_chat = st.tabs(
    ["Summary", "Data table", "Data quality", "AI chat"]
)

with tab_summary:
    if numeric_columns:
        st.markdown('<div class="section-kicker">At a glance</div><h3>Numeric summary</h3>', unsafe_allow_html=True)
        summary = data[numeric_columns].describe().T
        summary.insert(0, "Field", summary.index)
        st.dataframe(summary.reset_index(drop=True), use_container_width=True, hide_index=True)

        st.markdown('<div class="section-kicker">Signal finder</div><h3>Explore a measure</h3>', unsafe_allow_html=True)
        selected_measure = st.selectbox("Measure", numeric_columns)
        chart_data = data[[selected_measure]].dropna().reset_index(drop=True)
        st.line_chart(chart_data, height=280)
    else:
        st.info("No numeric columns were found in this worksheet.")

    if category_columns and numeric_columns:
        st.markdown('<div class="section-kicker">Breakdown</div><h3>Measure by category</h3>', unsafe_allow_html=True)
        selected_category = st.selectbox("Category", category_columns)
        selected_value = st.selectbox("Value", numeric_columns)
        grouped = (
            data.groupby(selected_category, dropna=False)[selected_value]
            .sum()
            .sort_values(ascending=False)
            .head(15)
        )
        grouped.index = grouped.index.astype(str)
        st.bar_chart(grouped, height=320)

with tab_data:
    st.dataframe(data, use_container_width=True, height=500)
    st.download_button(
        "Download current worksheet as CSV",
        data.to_csv(index=False).encode("utf-8"),
        file_name=f"{sheet_name}.csv",
        mime="text/csv",
    )

with tab_quality:
    quality = pd.DataFrame(
        {
            "Column": data.columns,
            "Type": [str(data[column].dtype) for column in data.columns],
            "Missing": [int(data[column].isna().sum()) for column in data.columns],
            "Unique values": [int(data[column].nunique(dropna=True)) for column in data.columns],
        }
    )
    st.dataframe(quality, use_container_width=True, hide_index=True)

with tab_chat:
    st.markdown('<div class="section-kicker">Ask the workbook</div><h3>AI data guide</h3>', unsafe_allow_html=True)
    st.caption("Ask about totals, trends, columns, or simple comparisons in the active worksheet.")

    api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        try:
            api_key = st.secrets.get("OPENAI_API_KEY", "")
        except StreamlitSecretNotFoundError:
            api_key = ""
    if not api_key:
        api_key = st.sidebar.text_input(
            "OpenAI API key",
            type="password",
            help="Used only for this session and never written to the workbook or source code.",
        )
    if not api_key:
        st.info("Enter your OpenAI API key in the sidebar to enable AI answers.")
    else:
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = []

        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        question = st.chat_input("Ask a question about this worksheet")
        if question:
            st.session_state.chat_messages.append({"role": "user", "content": question})
            with st.chat_message("user"):
                st.markdown(question)

            system_prompt = (
                "You are a careful financial data assistant. Answer only from the worksheet context "
                "below. If the context cannot answer the question, say so clearly. Show the calculation "
                "briefly when giving a number. Do not invent rows, values, or trends. Keep answers concise.\n\n"
                f"Worksheet: {sheet_name}\n{build_data_context(data)}"
            )
            try:
                client = OpenAI(api_key=api_key)
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        *st.session_state.chat_messages[-8:],
                    ],
                    temperature=0,
                )
                answer = response.choices[0].message.content
            except Exception as error:
                answer = f"I could not reach the AI service: {error}"

            st.session_state.chat_messages.append({"role": "assistant", "content": answer})
            with st.chat_message("assistant"):
                st.markdown(answer)