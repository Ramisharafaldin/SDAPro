import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from sklearn.linear_model import LinearRegression
import openpyxl

def convert_sales_date(df):
    if 'sales date' not in df.columns:
        st.warning("'sales date' column not found in the Excel file.")
        return None
    df['sales date'] = pd.to_datetime(df['sales date'], errors='coerce')
    if df['sales date'].isnull().any():
        st.warning("Some 'sales date' entries could not be converted to datetime.")
        return None
    return df

def calculate_days_of_sales(df, forecast_date):
    df = convert_sales_date(df)
    if df is None:
        return None
    if 'code' not in df.columns:
        st.warning("'code' column not found in the Excel file.")
        return None
    if 'item name' not in df.columns:
        st.warning("'item name' column not found in the Excel file.")
        return None
    if 'quy sale' not in df.columns:
        st.warning("'quy sale' column not found in the Excel file.")
        return None

    grouped = df.groupby('code')
    result = {}
    forecast_results = {}
    item_names = {}
    total_sales = {}
    for code, group in grouped:
        unique_dates = group['sales date'].dt.date.unique()
        result[code] = len(unique_dates)
        item_names[code] = group['item name'].iloc[0]
        total_sales[code] = group['quy sale'].sum()

        sales_counts = group.groupby('sales date')['quy sale'].sum().reset_index()
        if sales_counts.empty:
            st.warning(f"No sales data found for code: {code}")
            continue

        sales_counts['sales date'] = sales_counts['sales date'].map(datetime.toordinal)
        X = sales_counts['sales date'].values.reshape(-1, 1)
        y = sales_counts['quy sale'].values
        model = LinearRegression()
        model.fit(X, y)

        forecast_value = model.predict([[forecast_date.toordinal()]])[0]
        forecast_results[code] = max(0, forecast_value)

    days_of_sales_df = pd.DataFrame(list(result.items()), columns=['code', 'days_of_sales'])
    forecast_df = pd.DataFrame(list(forecast_results.items()), columns=['code', 'forecast'])
    item_names_df = pd.DataFrame(list(item_names.items()), columns=['code', 'item name'])
    total_sales_df = pd.DataFrame(list(total_sales.items()), columns=['code', 'total sales'])

    merged_df = days_of_sales_df.merge(item_names_df, on='code').merge(total_sales_df, on='code').merge(forecast_df, on='code')
    return merged_df[['code', 'item name', 'days_of_sales', 'total sales', 'forecast']]

def calculate_required_quantity(df, sales_duration, storage_duration):
    if 'quy sale' not in df.columns or 'quy' not in df.columns or 'nds' not in df.columns:
        raise KeyError("'quy sale', 'quy', and 'nds' columns are required in the input data")

    daily_sales = df['quy sale'] / sales_duration
    required_quantity = daily_sales * storage_duration - df['quy']
    df['Required Quantity'] = np.where(required_quantity > 0, required_quantity, np.nan)

    return df

def calculate_SDS(df, sales_duration, storage_duration):
    required_columns = ['quy', 'quy sale', 'nds', 'total sales']
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"'{col}' column is missing in the input data")

    df['Turnover'] = np.where(df['quy'] != 0, (df['quy sale'] / df['quy']).round(1), 'o.s')
    df['Daily Sales'] = np.where(df['quy sale'] != 0, (df['quy sale'] / df['nds']).round(1), 0)
    df['Days of Inventory'] = np.where(df['Daily Sales'] != 0, (df['quy'] / df['Daily Sales']).apply(lambda x: f"{int(x)}D" if not np.isinf(x) and not np.isnan(x) else 'inf'), '0D')
    df['Days of Inventory'] = df['Days of Inventory'].replace('inf', '9999999999.1D').replace('0D', '0.1D')

    df = calculate_required_quantity(df, sales_duration, storage_duration)

    total_sales = df['total sales'].sum()
    df['total sale percent'] = ((df['total sales'] / total_sales) * 100).apply(lambda x: f"{x:.2f}%")
    df = df.sort_values('total sales', ascending=False)
    df['total cum percent'] = (df['total sale percent'].str.rstrip('%').astype('float').cumsum()).apply(lambda x: f"{x:.2f}%")
    df['abc'] = ['A' if float(perc.rstrip('%')) <= 80 else 'B' if float(perc.rstrip('%')) <= 95 else 'C' for perc in df['total cum percent']]

    return df

def main():
    st.title("SDAPro System")
    selected_option = st.selectbox("Select Option", ("NDS Forcast", "SDS Stok"))

    file = st.file_uploader("Upload Excel file", type=['xlsx'])

    if file is not None:
        try:
            df = pd.read_excel(file)
        except Exception as e:
            st.warning(f"An error occurred while loading the Excel file: {str(e)}")
            return

        if selected_option == "NDS Forcast":
            try:
                forecast_date = st.date_input("Enter Forecast Date")
                df_processed = calculate_days_of_sales(df, forecast_date)
                if df_processed is not None:
                    st.write(df_processed)
            except KeyError as e:
                st.warning(str(e))

        elif selected_option == "SDS Stok":
            sales_duration = st.number_input("Enter Sales Duration (days):")
            storage_duration = st.number_input("Enter Storage Duration (days):")

            if st.button("Process SDS"):
                try:
                    df_processed = calculate_SDS(df, sales_duration, storage_duration)
                    st.write(df_processed)
                except KeyError as e:
                    st.warning(str(e))

if __name__ == "__main__":
    main()
