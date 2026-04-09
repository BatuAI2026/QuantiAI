import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
from prophet import Prophet
from sklearn.ensemble import IsolationForest
from io import StringIO

st.set_page_config(page_title="QuantifyAI v0.2.2", layout="wide")
st.title("🚀 QuantifyAI v0.2.2")
st.caption("AI-Enhanced Quantification Tool for Malaria Commodities | Focused on ACTs (Malawi-style)")

# ------------------- DEFAULT DATA -------------------
@st.cache_data
def load_default_data():
    # Small sample - will be replaced by upload
    data = """date,district,product_code,product_name,consumption_qty,stock_on_hand,shipments_received,adjustments,rainfall_mm,reported_cases
2023-01-01,Lilongwe,MRDT,mRDT,40424,114189,58599,-41,294.1,47227
2023-01-01,Lilongwe,LA1,LA 6x1,12769,26262,26043,-157,193.1,39277"""
    df = pd.read_csv(StringIO(data))
    df['date'] = pd.to_datetime(df['date'])
    return df

df = load_default_data()

# ------------------- UPLOAD WITH ROBUST COLUMN FIX -------------------
st.sidebar.header("📤 Upload Your 36-month Data")
uploaded_file = st.sidebar.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Robust column name fixing
        df.columns = [col.strip().lower() for col in df.columns]
        if 'date' not in df.columns:
            st.sidebar.error("Column 'date' not found. Please make sure your file has a column named 'date' (lowercase).")
        else:
            df['date'] = pd.to_datetime(df['date'], errors='coerce')
            st.sidebar.success(f"✅ Loaded {uploaded_file.name} — {len(df)} rows from {df['date'].dt.year.min()} to {df['date'].dt.year.max()}")
    except Exception as e:
        st.sidebar.error(f"Upload error: {str(e)}")

# ------------------- FILTERS -------------------
products = df['product_name'].unique() if 'product_name' in df.columns else []
selected_products = st.sidebar.multiselect("Select Products", products, default=["LA 6x4"] if "LA 6x4" in products else products[:2])

view_level = st.sidebar.radio("View Level", ["National (Aggregated)", "By District"])

# ------------------- DATA PREPARATION -------------------
if 'date' not in df.columns:
    st.error("Please upload a file with a 'date' column.")
    st.stop()

if view_level == "National (Aggregated)":
    working_df = df.groupby(['date', 'product_name']).agg({
        'consumption_qty': 'sum',
        'stock_on_hand': 'sum',
        'shipments_received': 'sum',
        'adjustments': 'sum',
        'rainfall_mm': 'mean',
        'reported_cases': 'sum'
    }).reset_index()
else:
    working_df = df.copy()

# Rest of the app (forecasting, etc.) - same as previous version but safer
st.header("1. Data Overview")
st.dataframe(working_df.head(10), use_container_width=True)

st.header("2. AI Forecasting Engine")
horizon = st.slider("Forecast Horizon (months)", 6, 36, 24)

tab1, tab2 = st.tabs(["Consumption Forecast", "Stock Status Matrix"])

with tab1:
    for prod in selected_products:
        st.subheader(f"Forecast for {prod}")
        sub_df = working_df[working_df['product_name'] == prod].copy()
        if len(sub_df) < 8:
            st.warning("Not enough data yet.")
            continue
        
        sub = sub_df.groupby('date')['consumption_qty'].sum().asfreq('MS').fillna(method='ffill')
        
        try:
            arima_model = ARIMA(sub, order=(1,1,1)).fit()
            arima_fc = arima_model.forecast(horizon)
        except:
            arima_fc = pd.Series([sub.iloc[-1]] * horizon, index=pd.date_range(sub.index[-1], periods=horizon, freq='MS'))
        
        # Prophet (simplified for now)
        prophet_data = sub_df.groupby('date').agg({'consumption_qty':'sum', 'rainfall_mm':'mean', 'reported_cases':'sum'}).reset_index()
        prophet_data = prophet_data.rename(columns={'date': 'ds', 'consumption_qty': 'y'})
        m = Prophet(yearly_seasonality=True)
        if 'rainfall_mm' in prophet_data.columns:
            m.add_regressor('rainfall_mm')
        if 'reported_cases' in prophet_data.columns:
            m.add_regressor('reported_cases')
        m.fit(prophet_data)
        future = m.make_future_dataframe(periods=horizon, freq='MS')
        future['rainfall_mm'] = prophet_data['rainfall_mm'].mean()
        future['reported_cases'] = prophet_data['reported_cases'].mean() * 1.05
        prophet_fc = m.predict(future)
        
        ensemble = (arima_fc.values + prophet_fc['yhat'].iloc[-horizon:].values) / 2
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=sub.index, y=sub.values, name="Historical"))
        fig.add_trace(go.Scatter(x=future['ds'].iloc[-horizon:], y=ensemble, name="AI Ensemble", line=dict(color="red", width=3)))
        fig.update_layout(title=f"{prod} Forecast", height=400)
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Stock Status Matrix (QAT-style)")
    latest = working_df.groupby('product_name').last().reset_index()
    matrix = latest[['product_name', 'stock_on_hand', 'consumption_qty']].copy()
    matrix['AMC'] = matrix['consumption_qty']
    matrix['MOS'] = (matrix['stock_on_hand'] / matrix['AMC']).round(1)
    st.dataframe(matrix, use_container_width=True)

st.caption("QuantifyAI v0.2.2 | Column name tolerant version")
