import streamlit as st
import pandas as pd
import numpy as np
import joblib

st.set_page_config(page_title="Property Price Predictor", page_icon="🏠", layout="wide")

@st.cache_resource
def load_model():
    return joblib.load("property_price_prediction_model.pkl")

@st.cache_data
def load_data():
    return pd.read_csv("Property_data (1) (1).csv")

model = load_model()
reference = load_data()

TARGET = "PropPrice"
ID_COL = "PropertyID"
feature_cols = [c for c in reference.columns if c not in [TARGET, ID_COL]]

st.title("🏠 Property Price Prediction")
st.caption("Machine Learning Capstone Project")
st.write("Enter the property characteristics below to estimate its sale price.")

with st.expander("About this application"):
    st.write("The application uses the trained preprocessing and machine-learning pipeline from the capstone project. Missing values are handled automatically.")

def make_widget(col):
    series = reference[col]
    label = col.replace("_", " ")

    if pd.api.types.is_numeric_dtype(series):
        values = pd.to_numeric(series, errors="coerce").dropna()
        if values.empty:
            return 0.0
        lo, hi, med = float(values.min()), float(values.max()), float(values.median())
        if pd.api.types.is_integer_dtype(series):
            return st.number_input(label, min_value=int(lo), max_value=int(hi),
                                   value=int(round(med)), step=1)
        step = max((hi - lo) / 100, 0.01)
        return st.number_input(label, min_value=lo, max_value=hi,
                               value=med, step=step, format="%.2f")

    values = sorted(series.dropna().astype(str).unique().tolist())
    return st.selectbox(label, values) if values else ""

groups = {
    "🏡 Property & Location": ["PropertyClass","PropertyZone","PropertyFrontage","PropertySize","Street","Alley","PropertyShape","Elevation","Amenities","LotOrientation","Grade","Neighborhood","Condition1","Condition2"],
    "🏗️ Building & Quality": ["BldgType","PropertyStyle","OverallQual","OverallCond","YearBuilt","YearRemodAdd","RoofStyle","RoofMatl","Roof1Material","Roof2Material","ExteriorCladdingType","ExteriorCladdingArea","ExterQual","ExterCond","PropertyFooting"],
    "🛋️ Basement & Rooms": ["BsmntFinish","BsmntMaintenance","BsmntVisibility","BsmntFinRat1","BsmntFinSty1","BsmntFinQual1","BsmtFinSF2","BsmtUnfSF","BsmntSqFtage","Heating","HeatingEfficiency","CentralAir","Electrical","1stFlrSF","2ndFlrSF","LowQualFinSF","GrLivArea","BsmtBath1","BsmtBath2","Bath1","Bath2","BedroomUpLev","Kitchen","KitchenQual","CntRmsUpLev"],
    "🚗 Garage & Outdoor": ["Functional","CntFireplaces","QualFireplace","BasementType","BasementYrBlt","BasementFinish","BasementCars","SquareFootage","BasementQual","BasementSqFootage","PavedDrive","WoodDeckSF","OpenPorchSF","EnclosedPorch","3SsnPorch","ScreenPorch","PoolArea","PoolQC","BoundaryFeatures","AddFeatures","AddVal"],
    "💰 Sale Information": ["SaleMon","YrSold","SaleType","SaleCondn"]
}

inputs = {}

for section, columns in groups.items():
    columns = [c for c in columns if c in feature_cols]
    if not columns:
        continue
    st.subheader(section)
    ui_cols = st.columns(2)
    for i, col in enumerate(columns):
        with ui_cols[i % 2]:
            inputs[col] = make_widget(col)

remaining = [c for c in feature_cols if c not in inputs]
if remaining:
    with st.expander("Other Dataset Features"):
        ui_cols = st.columns(2)
        for i, col in enumerate(remaining):
            with ui_cols[i % 2]:
                inputs[col] = make_widget(col)

st.divider()

if st.button("🔮 Predict Property Price", type="primary", use_container_width=True):
    try:
        input_df = pd.DataFrame([inputs], columns=feature_cols)
        prediction = float(model.predict(input_df)[0])
        st.success("Prediction generated successfully!")
        st.metric("Estimated Property Price", f"${prediction:,.0f}")
        st.info("This is a machine-learning estimate, not a guaranteed market value.")
    except Exception as exc:
        st.error(f"Prediction failed: {exc}")

st.sidebar.title("Model Information")
st.sidebar.write("**Selected model:** Gradient Boosting")
st.sidebar.write("**Target:** PropPrice")
st.sidebar.write("**Test R²:** 0.9035")
st.sidebar.write("**Test MAE:** $16,352.24")
st.sidebar.write("**Test RMSE:** $27,204.06")
