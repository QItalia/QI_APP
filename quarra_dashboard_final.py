import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import timedelta

st.set_page_config(page_title="Quarra Dashboard", layout="centered")
st.image("logo_quarra_italia.png", width=200)
st.markdown("""<h2 style='text-align: center;'>📊 Quarra Italia - Weekly Dashboard</h2>""", unsafe_allow_html=True)

xls = pd.ExcelFile("dati_quarra.xlsx")
df_produzione = pd.read_excel(xls, sheet_name="Produzione")
df_entrate = pd.read_excel(xls, sheet_name="Entrate")
df_uscite = pd.read_excel(xls, sheet_name="Uscite")
df_saldo = pd.read_excel(xls, sheet_name="Saldo")

for df in [df_produzione, df_entrate, df_uscite, df_saldo]:
    df["Data"] = pd.to_datetime(df["Data"])

def label_week(data):
    start = data - timedelta(days=6)
    end = data
    return f"{start.day:02d}-{start.strftime('%b')} → {end.day:02d}-{end.strftime('%b')}"

df_prod_weekly = df_produzione.groupby(pd.Grouper(key="Data", freq="W-FRI")).sum().reset_index()
df_prod_weekly["Week"] = df_prod_weekly["Data"].apply(label_week)

df_entr_weekly = df_entrate.groupby(pd.Grouper(key="Data", freq="W-FRI")).sum().reset_index()
df_entr_weekly["Week"] = df_entr_weekly["Data"].apply(label_week)

df_usc_weekly = df_uscite.groupby(pd.Grouper(key="Data", freq="W-FRI")).sum().reset_index()
df_usc_weekly["Week"] = df_usc_weekly["Data"].apply(label_week)

df_saldo_weekly = df_saldo.groupby(pd.Grouper(key="Data", freq="W-FRI")).sum().reset_index()
df_saldo_weekly["Week"] = df_saldo_weekly["Data"].apply(label_week)

latest_date = df_saldo_weekly["Data"].max()
latest_label = label_week(latest_date)
st.markdown(f"### 📅 Weekly Report ending on: **{latest_label}**")

def draw_gauge(title, value, color, description):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={"suffix": " €"},
        title={"text": f"{title}<br><span style='font-size:0.8em'>{description}</span>"},
        gauge={"axis": {"range": [None, value * 1.5 if value > 0 else 100]},
               "bar": {"color": color}, "bgcolor": "white",
               "steps": [{"range": [0, value], "color": color}]},
        domain={"x": [0, 1], "y": [0, 1]}
    ))
    fig.update_layout(margin=dict(l=10, r=10, t=80, b=40), height=300)
    return fig

st.subheader("📈 Weekly Indicators")
if not df_prod_weekly.empty:
    st.plotly_chart(draw_gauge("Weekly Production", df_prod_weekly.iloc[-1, 1], "#004C99", "Value of goods produced this week"), use_container_width=True)
if not df_entr_weekly.empty:
    st.plotly_chart(draw_gauge("Weekly Bank Income", df_entr_weekly.iloc[-1, 1], "#2E8B57", "Cash income registered this week"), use_container_width=True)
if not df_usc_weekly.empty:
    st.plotly_chart(draw_gauge("Weekly Bank Expenses", df_usc_weekly.iloc[-1, 1], "#B22222", "Cash outflows this week"), use_container_width=True)
if not df_saldo_weekly.empty:
    st.plotly_chart(draw_gauge("Weekly Balance", df_saldo_weekly.iloc[-1, 1], "#FFD700", "Balance status at end of this week"), use_container_width=True)

st.markdown("---")
st.subheader("📅 Select a month to view trends")
today = pd.Timestamp.today()
df_saldo_weekly["Month"] = df_saldo_weekly["Data"].dt.to_period("M").dt.to_timestamp()
available_months = sorted(df_saldo_weekly["Month"].unique())
month_labels = [d.strftime("%B %Y") for d in available_months]
current_month_idx = next((i for i, d in enumerate(available_months) if d.month == today.month and d.year == today.year), len(available_months)-1)
month_idx = st.selectbox("Select month", options=list(range(len(available_months))), format_func=lambda i: month_labels[i], index=current_month_idx)
selected_month = available_months[month_idx]

f_prod = df_prod_weekly[df_prod_weekly["Data"].dt.to_period("M") == selected_month.to_period("M")]
f_entr = df_entr_weekly[df_entr_weekly["Data"].dt.to_period("M") == selected_month.to_period("M")]
f_usc = df_usc_weekly[df_usc_weekly["Data"].dt.to_period("M") == selected_month.to_period("M")]
f_saldo = df_saldo_weekly[df_saldo_weekly["Data"].dt.to_period("M") == selected_month.to_period("M")]

st.markdown("---")
st.subheader("📊 Monthly Trends")
config = {"staticPlot": True}

if not f_prod.empty:
    fig = px.area(f_prod, x="Week", y=f_prod.columns[1], title="Monthly Production", line_shape="linear")
    fig.update_traces(line_color="#004C99", fillcolor="#B0C4DE")
    st.plotly_chart(fig, use_container_width=True, config=config)

if not f_entr.empty:
    fig = px.area(f_entr, x="Week", y=f_entr.columns[1], title="Monthly Bank Income", line_shape="linear")
    fig.update_traces(line_color="#2E8B57", fillcolor="#C1E1C1")
    st.plotly_chart(fig, use_container_width=True, config=config)

if not f_usc.empty:
    fig = px.area(f_usc, x="Week", y=f_usc.columns[1], title="Monthly Bank Expenses", line_shape="linear")
    fig.update_traces(line_color="#B22222", fillcolor="#F08080")
    st.plotly_chart(fig, use_container_width=True, config=config)

if not f_saldo.empty:
    fig = px.area(f_saldo, x="Week", y=f_saldo.columns[1], title="Monthly Balance", line_shape="linear")
    fig.update_traces(line_color="#FFD700", fillcolor="#FFFACD")
    st.plotly_chart(fig, use_container_width=True, config=config)

st.markdown("---")
st.subheader("⬇️ Export Weekly Data (Selected Month)")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    f_prod.to_excel(writer, sheet_name='Weekly Production', index=False)
    f_entr.to_excel(writer, sheet_name='Weekly Bank Income', index=False)
    f_usc.to_excel(writer, sheet_name='Weekly Bank Expenses', index=False)
    f_saldo.to_excel(writer, sheet_name='Weekly Balance', index=False)
st.download_button("Download Excel Report", data=output.getvalue(), file_name="weekly_summary.xlsx")
