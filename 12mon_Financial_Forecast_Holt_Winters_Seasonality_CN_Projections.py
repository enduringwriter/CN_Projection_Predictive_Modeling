import os
import pandas as pd
from statsmodels.tsa.api import ExponentialSmoothing

# Specify the directory and file name
file_path = r"/Users/keithstateson/Projects_on_MBA/DACI CN Projections Tool/MARO_Canned_Report.xlsx"

# Specify the sheet name
sheet_name = 'MARO_Forecast'  # Provide a valid sheet name

# Create a new Excel writer with the 'xlsxwriter' engine
excel_writer = pd.ExcelWriter('CN_Forecast_HW_Extract_v4_analyze_more_rows.xlsx', engine='xlsxwriter')

# Get the data from the Excel file and put it in a dataframe
# The 'usecols' parameter specifies the range of columns to read (C to Z)
# The 'skiprows' parameter skips the first 11 rows (to start reading from row 12)
df = pd.read_excel(file_path, sheet_name=0, usecols="C:Z", skiprows=11, header=None)

# Analyze rows 12 to 16
for row_num in range(12, 17):
    # Get the data for the current row
    row_data = df.iloc[row_num - 12]

    # Create a DataFrame for Time Series Analysis (TSA) for Holt-Winters
    df_temp = pd.DataFrame({'Value': row_data.astype(float)})
    df_temp['Date'] = pd.date_range(start='2016-10-01', periods=len(row_data), freq='M')  # update "start" to the earliest year
    ts_data = df_temp.set_index('Date')['Value'].asfreq('M')

    # Holt-Winters Exponential Smoothing Model
    model_hw = ExponentialSmoothing(ts_data, trend='add', seasonal='add', seasonal_periods=12)
    model_hw_fit = model_hw.fit()
    forecast_data_hw = model_hw_fit.forecast(steps=12)

    # Replace negative forecasts with zero
    forecast_data_hw[forecast_data_hw < 0] = 0

    # Create a DataFrame for forecasts
    forecast_data_hw_df = pd.DataFrame({
        'Holt-Winters': forecast_data_hw
    })

    # Transpose the DataFrame to have the data in a horizontal format
    forecast_data_hw_df = forecast_data_hw_df.T

    # Write the DataFrame to the existing sheet, starting from C12 for each row
    forecast_data_hw_df.to_excel(excel_writer, sheet_name=sheet_name, startcol=2, startrow=row_num - 1, index=False, header=False)

# Save the Excel file
excel_writer.close()
