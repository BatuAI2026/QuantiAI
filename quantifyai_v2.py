import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
from prophet import Prophet
from sklearn.ensemble import IsolationForest
from io import StringIO
import anthropic

st.set_page_config(page_title="QuantifyAI v0.2.3", layout="wide")
st.title("🚀 QuantifyAI v0.2.3")
st.caption("AI-Enhanced Quantification Tool for Malaria Commodities | Focused on ACTs (Malawi-style)")

# ------------------- UPLOAD -------------------
st.sidebar.header("📤 Upload Your 36-month Data")
uploaded_file = st.sidebar.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Robust cleaning
        df.columns = [col.strip().lower() for col in df.columns]
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        df = df.dropna(subset=['date'])
        
        st.sidebar.success(f"✅ Loaded {uploaded_file.name} — {len(df)} rows ({df['date'].dt.year.min()}-{df['date'].dt.year.max()})")
    except Exception as e:
        st.sidebar.error(f"Error: {str(e)}")
else:
    st.info("Please upload your 36-month data file.")
    st.stop()

# ------------------- FILTERS -------------------
products = sorted(df['product_name'].unique())
selected_products = st.sidebar.multiselect("Select Products", products, default=["LA 6x4"])

view_level = st.sidebar.radio("View Level", ["National (Aggregated)", "By District"])

# ------------------- DATA PREPARATION -------------------
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

# ------------------- FORECASTING -------------------
st.header("2. AI Forecasting Engine (QAT + AI Upgrade)")
horizon = st.slider("Forecast Horizon (months)", 6, 36, 24)

tab1, tab2, tab3 = st.tabs(["Consumption Forecast", "Stock Status Matrix (QAT-style)", "AI Advisor"])

with tab1:
    for prod in selected_products:
        st.subheader(f"Forecast for {prod}")
        
        sub_df = working_df[working_df['product_name'] == prod].copy()
        if len(sub_df) < 12:
            st.warning("Not enough data for reliable forecast.")
            continue
        
        # Safe monthly aggregation
        monthly = sub_df.groupby('date')['consumption_qty'].sum().reset_index()
        monthly = monthly.set_index('date').asfreq('MS')
        cons_series = monthly['consumption_qty'].ffill().bfill()
        
        # ARIMA
        try:
            arima_model = ARIMA(cons_series, order=(1,1,1)).fit()
            arima_fc = arima_model.forecast(horizon)
        except:
            arima_fc = pd.Series([cons_series.iloc[-1]] * horizon, 
                                 index=pd.date_range(cons_series.index[-1], periods=horizon, freq='MS'))
        
        # Prophet
        prophet_data = sub_df.groupby('date').agg({
            'consumption_qty': 'sum', 
            'rainfall_mm': 'mean', 
            'reported_cases': 'sum'
        }).reset_index()
        prophet_data = prophet_data.rename(columns={'date': 'ds', 'consumption_qty': 'y'})
        
        m = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False)
        if 'rainfall_mm' in prophet_data.columns and prophet_data['rainfall_mm'].notna().any():
            m.add_regressor('rainfall_mm')
        if 'reported_cases' in prophet_data.columns and prophet_data['reported_cases'].notna().any():
            m.add_regressor('reported_cases')
        
        m.fit(prophet_data)
        future = m.make_future_dataframe(periods=horizon, freq='MS')
        future['rainfall_mm'] = prophet_data['rainfall_mm'].mean()
        future['reported_cases'] = prophet_data['reported_cases'].mean() * 1.05
        prophet_fc = m.predict(future)
        
        ensemble_fc = (arima_fc.values + prophet_fc['yhat'].iloc[-horizon:].values) / 2
        
        # Plot
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=cons_series.index, y=cons_series.values, name="Historical Consumption", line=dict(color="blue")))
        fig.add_trace(go.Scatter(x=arima_fc.index, y=arima_fc.values, name="ARIMA", line=dict(dash="dash")))
        fig.add_trace(go.Scatter(x=future['ds'].iloc[-horizon:], y=prophet_fc['yhat'].iloc[-horizon:], name="AI Prophet", line=dict(color="green")))
        fig.add_trace(go.Scatter(x=future['ds'].iloc[-horizon:], y=ensemble_fc, name="AI Ensemble (Recommended)", line=dict(color="red", width=3)))
        
        fig.update_layout(title=f"{prod} - {view_level} Forecast", xaxis_title="Date", yaxis_title="Quantity", height=450)
        st.plotly_chart(fig, use_container_width=True)
        
        # Create forecast table
        forecast_table = pd.DataFrame({
            'Month': future['ds'].iloc[-horizon:].dt.strftime('%Y-%m'),
            'ARIMA': arima_fc.values.round(0).astype(int),
            'Prophet': prophet_fc['yhat'].iloc[-horizon:].values.round(0).astype(int),
            'Ensemble (Recommended)': ensemble_fc.round(0).astype(int)
        })
        
        st.subheader(f"Monthly Forecast Table - {prod}")
        st.dataframe(forecast_table, use_container_width=True)
        
        st.metric(f"Projected Need next 12 months", f"{int(ensemble_fc[:12].sum()):,}")

with tab2:
    st.subheader("Stock Status Matrix")
    latest = working_df.groupby('product_name').last().reset_index()
    matrix = latest[['product_name', 'stock_on_hand', 'consumption_qty']].copy()
    matrix['AMC'] = matrix['consumption_qty']
    matrix['MOS'] = (matrix['stock_on_hand'] / matrix['AMC']).round(1)
    st.dataframe(matrix, use_container_width=True)

with tab3:
    st.subheader("🤖 AI Advisor - Ask Questions About Your Data")
    
    # Initialize session state for chat history
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    
    # Display chat history
    for message in st.session_state.chat_history:
        if message['role'] == 'user':
            st.write(f"**You:** {message['content']}")
        else:
            st.write(f"**AI Advisor:** {message['content']}")
    
    # User input
    user_question = st.text_input("Ask the AI Advisor anything about your forecast and stock levels...")
    
    if user_question:
        # Prepare context from current data
        latest = working_df.groupby('product_name').last().reset_index()
        
        context = f"""
        You are an expert AI advisor for malaria commodity quantification and supply chain management.
        
        Current Data Summary:
        - Products: {', '.join(selected_products)}
        - View Level: {view_level}
        - Forecast Horizon: {horizon} months
        - Data period: {working_df['date'].min().date()} to {working_df['date'].max().date()}
        
        Current Stock Status:
        {latest[['product_name', 'stock_on_hand', 'consumption_qty']].to_string()}
        
        Provide actionable insights for supply chain optimization, procurement recommendations, 
        and risk alerts based on the forecast data. Be concise and specific.
        """
        
        try:
            client = anthropic.Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY"))
            
            # Include chat history for context
            messages = []
            for msg in st.session_state.chat_history:
                messages.append({"role": msg['role'], "content": msg['content']})
            messages.append({"role": "user", "content": user_question})
            
            response = client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1024,
                system=context,
                messages=messages
            )
            
            ai_response = response.content[0].text
            
            # Store in session history
            st.session_state.chat_history.append({"role": "user", "content": user_question})
            st.session_state.chat_history.append({"role": "assistant", "content": ai_response})
            
            # Rerun to display new message
            st.rerun()
            
        except Exception as e:
            st.error(f"Error connecting to AI Advisor: {str(e)}")
    
    # Auto-generated insights
    if len(st.session_state.chat_history) == 0:
        st.info("💡 **Quick Tips:**\n- Ask about procurement recommendations\n- Request stock warnings\n- Ask for monthly ordering suggestions\n- Query seasonal patterns")

st.caption("QuantifyAI v0.2.3 | With Monthly Forecasts & AI Advisor")
pip install anthropic
