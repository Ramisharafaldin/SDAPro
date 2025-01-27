import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from sklearn.linear_model import LinearRegression
from io import BytesIO

# Function to process sales data from all sheets
def process_all_sheets(file):
    # Read all sheets from the file
    dataframes = pd.read_excel(file, sheet_name=None)
    
    # Combine all sheets into one DataFrame
    all_data = pd.DataFrame()
    for sheet_name, df in dataframes.items():
        df['Sheet Name'] = sheet_name  # Add a column to identify the sheet
        all_data = pd.concat([all_data, df], ignore_index=True)

    # Convert the sales date column to datetime format
    all_data['sales date'] = pd.to_datetime(all_data['sales date'], format='%d-%m-%y', errors='coerce')

    # Check for missing columns
    required_columns = ['sales date', 'quy sales', 'code', 'item name']
    for col in required_columns:
        if col not in all_data.columns:
            raise KeyError(f"Missing required column: {col}")

    # Remove rows with invalid dates
    invalid_dates_count = all_data['sales date'].isna().sum()
    all_data = all_data.dropna(subset=['sales date'])

    if invalid_dates_count > 0:
        st.warning(f"{invalid_dates_count} rows were removed due to invalid dates.")

    # Calculate general metrics
    unique_sales_days = all_data['sales date'].nunique()
    number_of_sales_transactions = all_data.shape[0]
    total_quantity_sold = all_data['quy sales'].sum()

    # Create a DataFrame for general results
    results = {
        'Sales Days': unique_sales_days,
        'Sales Transactions': number_of_sales_transactions,
        'Total Quantity Sold': total_quantity_sold
    }
    results_df = pd.DataFrame([results])

    # Aggregate data by item
    results_by_item = all_data.groupby(['code', 'item name'])['quy sales'].sum().reset_index()
    results_by_item['Sales Transactions'] = all_data.groupby(['code', 'item name'])['sales date'].count().values
    results_by_item['Sales Days'] = all_data.groupby(['code', 'item name'])['sales date'].nunique().values

    return results_df, results_by_item, all_data

# Function to calculate required quantity
def calculate_required_quantity(df, sales_duration, storage_duration):
    if 'quy sales' not in df.columns or 'quy' not in df.columns or 'nds' not in df.columns:
        raise KeyError("'quy sales', 'quy', and 'nds' columns are required in the input data")

    daily_sales = df['quy sales'] / sales_duration
    required_quantity = daily_sales * storage_duration - df['quy']
    df['Required Quantity'] = np.where(required_quantity > 0, required_quantity, np.nan)

    return df

# Function to calculate SDS
def calculate_SDS(df, sales_duration, storage_duration):
    required_columns = ['quy', 'quy sales', 'nds', 'total sales']
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"'{col}' column is missing in the input data")

    df['Turnover'] = np.where(df['quy'] != 0, (df['quy sales'] / df['quy']).round(1), 'o.s')
    df['Daily Sales'] = np.where(df['quy sales'] != 0, (df['quy sales'] / df['nds']).round(1), 0)
    df['Days of Inventory'] = np.where(df['Daily Sales'] != 0, (df['quy'] / df['Daily Sales']).apply(lambda x: f"{int(x)}D" if not np.isinf(x) and not np.isnan(x) else 'inf'), '0D')
    df['Days of Inventory'] = df['Days of Inventory'].replace('inf', '9999999999.1D').replace('0D', '0.1D')

    df = calculate_required_quantity(df, sales_duration, storage_duration)

    total_sales = df['total sales'].sum()
    df['total sale percent'] = ((df['total sales'] / total_sales) * 100).apply(lambda x: f"{x:.2f}%")
    df = df.sort_values('total sales', ascending=False)
    df['total cum percent'] = (df['total sale percent'].str.rstrip('%').astype('float').cumsum()).apply(lambda x: f"{x:.2f}%")
    df['abc'] = ['A' if float(perc.rstrip('%')) <= 80 else 'B' if float(perc.rstrip('%')) <= 95 else 'C' for perc in df['total cum percent']]

    return df

# Streamlit interface
def main():
    st.title("SDAPro System")
    
    # تعديل القائمة المنسدلة
    selected_option = st.selectbox("Select Option", ("Sales Data Analysis", "Stock Control"))

    file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if file:
        try:
            # Save the uploaded file to a path
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(file.getbuffer())
            
            file_path = "uploaded_file.xlsx"
            
            if selected_option == "Sales Data Analysis":
                if st.button("RUN"):  # تغيير اسم الزر إلى RUN
                    try:
                        # Process data using the new function
                        results_df, results_by_item, all_data = process_all_sheets(file_path)

                        # Display general results
                        st.subheader("General Metrics")
                        st.dataframe(results_df)

                        # Display detailed results by item
                        st.subheader("Results by Item")
                        st.dataframe(results_by_item)

                        # Display combined data
                        st.subheader("Combined Data")
                        st.dataframe(all_data)

                        # Save results to an Excel file
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            results_df.to_excel(writer, sheet_name="Summary", index=False)
                            results_by_item.to_excel(writer, sheet_name="Item Analysis", index=False)
                            all_data.to_excel(writer, sheet_name="All Data", index=False)
                        output.seek(0)

                        st.download_button(
                            label="Download Results (Excel)",
                            data=output,
                            file_name="analysis_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"An error occurred while processing the file: {e}")

            elif selected_option == "Stock Control":
                sales_duration = st.number_input("Enter Sales Duration (days):")
                storage_duration = st.number_input("Enter Storage Duration (days):")
                if st.button("RUN"):  # تغيير اسم الزر إلى RUN
                    try:
                        df = pd.read_excel(file_path)
                        df_processed = calculate_SDS(df, sales_duration, storage_duration)
                        st.write(df_processed)
                    except KeyError as e:
                        st.warning(str(e))

        except Exception as e:
            st.warning(f"An error occurred while loading the Excel file: {str(e)}")
            return

if __name__ == "__main__":
    main()
