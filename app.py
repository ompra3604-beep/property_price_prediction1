import streamlit as st
import pandas as pd
# import joblib


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Property Price Prediction",
    page_icon="🏠",
    layout="wide"
)


# ============================================================
# LOAD MODEL
# ============================================================

@st.cache_resource
def load_model():
    return joblib.load("property_price_model.pkl")


model = load_model()


# ============================================================
# TITLE
# ============================================================

st.title("🏠 Property Price Prediction")

st.markdown(
    """
    ### Machine Learning Property Price Prediction System

    Enter the property characteristics below to estimate the
    expected property price.
    """
)

st.divider()


# ============================================================
# INPUT SECTION
# ============================================================

st.header("🏡 Property Information")

col1, col2, col3 = st.columns(3)


with col1:

    overall_qual = st.slider(
        "Overall Quality",
        min_value=1,
        max_value=10,
        value=7
    )

    gr_liv_area = st.number_input(
        "Living Area (sq ft)",
        min_value=300,
        max_value=5000,
        value=1500
    )

    property_size = st.number_input(
        "Property Size",
        min_value=500,
        max_value=30000,
        value=8000
    )

    year_built = st.number_input(
        "Year Built",
        min_value=1900,
        max_value=2026,
        value=2000
    )


with col2:

    neighborhood = st.selectbox(
        "Neighborhood",
        [
            "North",
            "South",
            "East",
            "West",
            "Central"
        ]
    )

    property_zone = st.selectbox(
        "Property Zone",
        [
            "Residential",
            "Residential-High",
            "Mixed"
        ]
    )

    property_class = st.selectbox(
        "Property Class",
        [20, 30, 40, 50, 60]
    )

    amenities = st.selectbox(
        "Amenities",
        [
            "None",
            "Basic",
            "Good",
            "Premium"
        ]
    )


with col3:

    bedrooms = st.number_input(
        "Bedrooms",
        min_value=1,
        max_value=8,
        value=3
    )

    bath1 = st.number_input(
        "Full Bathrooms",
        min_value=1,
        max_value=5,
        value=2
    )

    basement_cars = st.number_input(
        "Garage Capacity",
        min_value=0,
        max_value=5,
        value=1
    )

    overall_cond = st.slider(
        "Overall Condition",
        min_value=1,
        max_value=10,
        value=6
    )


# ============================================================
# PREDICTION
# ============================================================

st.divider()

if st.button(
    "🔮 Predict Property Price",
    use_container_width=True
):

    # Create input dataframe
    input_data = pd.DataFrame({
        "PropertyClass": [property_class],
        "PropertyZone": [property_zone],
        "PropertyFrontage": [100],
        "PropertySize": [property_size],
        "Street": ["Paved"],
        "PropertyShape": ["Regular"],
        "Elevation": ["Flat"],
        "Amenities": [amenities],
        "Neighborhood": [neighborhood],
        "BldgType": ["Single"],
        "PropertyStyle": ["1Story"],
        "OverallQual": [overall_qual],
        "OverallCond": [overall_cond],
        "YearBuilt": [year_built],
        "YearRemodAdd": [year_built],
        "ExteriorCladdingType": ["Brick"],
        "ExteriorCladdingArea": [200],
        "ExterQual": ["Good"],
        "ExterCond": ["Good"],
        "Heating": ["Gas"],
        "HeatingEfficiency": ["Good"],
        "CentralAir": ["Y"],
        "1stFlrSF": [gr_liv_area],
        "2ndFlrSF": [0],
        "GrLivArea": [gr_liv_area],
        "Bath1": [bath1],
        "Bath2": [0],
        "BedroomUpLev": [bedrooms],
        "KitchenQual": ["Good"],
        "CntRmsUpLev": [7],
        "Functional": ["Good"],
        "CntFireplaces": [1],
        "QualFireplace": ["Good"],
        "BasementCars": [basement_cars],
        "SquareFootage": [500],
        "BasementQual": ["Good"],
        "PavedDrive": ["Y"],
        "WoodDeckSF": [100],
        "OpenPorchSF": [50],
        "EnclosedPorch": [0],
        "ScreenPorch": [0],
        "PoolArea": [0],
        "AddVal": [0],
        "SaleMon": [6],
        "YrSold": [2026],
        "SaleType": ["Normal"],
        "SaleCondn": ["Normal"]
    })

    try:

        prediction = model.predict(input_data)[0]

        st.success("Prediction generated successfully!")

        st.metric(
            label="🏠 Estimated Property Price",
            value=f"${prediction:,.0f}"
        )

        st.info(
            "The predicted price is generated using the trained "
            "machine-learning model."
        )

    except Exception as e:

        st.error("Prediction could not be generated.")

        st.exception(e)