# =====================================================================# MONTHLY FILE VOLUME FORECASTING SYSTEM# RANDOM FOREST + EXTRA TREES + SEASONAL NAIVE# AUTOMATIC MODEL SELECTION# DUPLICATE SUBMISSION IDS INCLUDED# MANAGEMENT READY REPORT# =====================================================================
!pip -q install openpyxl scikit-learn



import pandas as pd

import numpy as np

import re

import warnings



from google.colab import files

from IPython.display import display



from sklearn.ensemble import (

    RandomForestRegressor,

    ExtraTreesRegressor

)



from sklearn.metrics import mean_absolute_error



warnings.filterwarnings("ignore")
# =====================================================================# 1. UPLOAD EXCEL FILES# =====================================================================
print("=" * 70)

print("MONTHLY FILE VOLUME - MACHINE LEARNING FORECAST")

print("=" * 70)



uploaded = files.upload()



if not uploaded:

    raise ValueError("Please upload Excel files.")
# =====================================================================# 2. COLUMN IDENTIFICATION# =====================================================================
def normalize_column(name):

    return re.sub(r"[^a-z0-9]", "", str(name).lower())





DATE_NAMES = [

    "receiveddatetime",

    "receiveddate",

    "receiveddatetimeutc",

    "receiveddateutc",

    "receiveddatetimeist",

    "receiveddateandtime"

]



ID_NAMES = [

    "submissionid",

    "submissionnumber",

    "submission"

]



STATUS_NAMES = [

    "status",

    "submissionstatus",

    "filestatus",

    "processingstatus"

]





def find_column(df, possible_names):



    columns = {

        normalize_column(c): c

        for c in df.columns

    }



    for name in possible_names:

        if name in columns:

            return columns[name]



    return None
# =====================================================================# 3. DATE CONVERSION# =====================================================================
def convert_dates(series):



    if pd.api.types.is_datetime64_any_dtype(series):

        return pd.to_datetime(series, errors="coerce")



    numeric = pd.to_numeric(series, errors="coerce")



    if (

        numeric.notna().mean() > 0.80

        and numeric.max() < 100000

        and numeric.min() >= 0

    ):



        return pd.to_datetime(

            numeric,

            unit="D",

            origin="1899-12-30",

            errors="coerce"

        )



    try:



        return pd.to_datetime(

            series,

            dayfirst=True,

            format="mixed",

            errors="coerce"

        )



    except:



        return pd.to_datetime(

            series,

            dayfirst=True,

            errors="coerce"

        )
# =====================================================================# 4. READ ALL FILES AND SHEETS# =====================================================================
all_data = []

processing_summary = []



allowed_status = [

    "complete",

    "completed",

    "qa uploaded"

]



for filename in uploaded.keys():



    if not filename.lower().endswith(

        (".xlsx", ".xlsm", ".xls")

    ):

        continue



    print("\nReading:", filename)



    try:



        excel = pd.ExcelFile(filename)



        for sheet in excel.sheet_names:



            df = pd.read_excel(

                filename,

                sheet_name=sheet

            )



            if df.empty:

                continue



            date_col = find_column(df, DATE_NAMES)

            id_col = find_column(df, ID_NAMES)

            status_col = find_column(df, STATUS_NAMES)



            if date_col is None or id_col is None:



                print(

                    "Skipped sheet:",

                    sheet,

                    "- Date or Submission ID missing"

                )



                continue
# Count all nonblank Submission ID rows.# Duplicate IDs are intentionally included.
            df = df[

                df[id_col].notna()

                & df[id_col].astype(str).str.strip().ne("")

            ].copy()



            rows_before_status = len(df)
# Apply status filter only if Status column exists.
            if status_col is not None:



                df = df[

                    df[status_col]

                    .astype(str)

                    .str.strip()

                    .str.lower()

                    .isin(allowed_status)

                ].copy()



            rows_after_status = len(df)



            df["Received_DateTime"] = convert_dates(

                df[date_col]

            )



            invalid_dates = int(

                df["Received_DateTime"].isna().sum()

            )



            df = df[

                df["Received_DateTime"].notna()

            ].copy()



            if df.empty:

                continue



            df["Received_Date"] = (

                df["Received_DateTime"].dt.normalize()

            )



            all_data.append(

                df[

                    [

                        "Received_Date",

                        "Received_DateTime"

                    ]

                ].copy()

            )



            processing_summary.append({



                "File": filename,

                "Sheet": sheet,

                "Rows Before Status Filter": rows_before_status,

                "Rows After Status Filter": rows_after_status,

                "Invalid Dates": invalid_dates,

                "Final Counted Rows": len(df)



            })



            print(

                sheet,

                "=>",

                f"{len(df):,}",

                "rows counted"

            )



    except Exception as e:



        print("File error:", filename, e)





if not all_data:



    raise ValueError(

        "No valid data found. Check the Received Date Time "

        "and Submission ID columns."

    )





data = pd.concat(

    all_data,

    ignore_index=True

)



data = data.sort_values(

    "Received_Date"

).reset_index(drop=True)



print("\nTotal eligible rows:", f"{len(data):,}")

print("Duplicate Submission IDs: INCLUDED")
# =====================================================================# 5. MONTHLY ACTUAL FILE VOLUME# =====================================================================
data["Month"] = (

    data["Received_Date"].dt.to_period("M")

)



monthly = (

    data.groupby("Month")

    .size()

    .reset_index(name="Total Files")

)



monthly["Month"] = monthly["Month"].astype(str)



print("\n" + "=" * 70)

print("MONTH-WISE ACTUAL FILE VOLUME")

print("=" * 70)



display(monthly)
# =====================================================================# 6. DAILY ACTUAL FILE VOLUME# =====================================================================
daily_actual = (

    data.groupby("Received_Date")

    .size()

    .reset_index(name="File Count")

)



daily_actual["Weekday"] = (

    daily_actual["Received_Date"].dt.day_name()

)



daily_actual["Month"] = (

    daily_actual["Received_Date"].dt.strftime("%B %Y")

)



print("\nDAILY ACTUAL FILE VOLUME")



display(daily_actual)
# =====================================================================# 7. WEEKDAY ANALYSIS# =====================================================================
weekday_order = [

    "Monday",

    "Tuesday",

    "Wednesday",

    "Thursday",

    "Friday",

    "Saturday",

    "Sunday"

]



weekday_summary = (

    daily_actual.groupby("Weekday")["File Count"]

    .agg(["sum", "mean", "count"])

    .reindex(weekday_order)

    .reset_index()

)



weekday_summary.columns = [

    "Weekday",

    "Total Files",

    "Average Files Per Day",

    "Number of Days"

]



weekday_summary["Average Files Per Day"] = (

    weekday_summary["Average Files Per Day"].round(0)

)



print("\nWEEKDAY ANALYSIS")



display(weekday_summary)
# =====================================================================# 8. BEGINNING / MIDDLE / END OF MONTH# =====================================================================
data["Day"] = data["Received_Date"].dt.day



data["Month Period"] = np.select(

    [

        data["Day"] <= 10,

        data["Day"] <= 20

    ],

    [

        "Beginning (1-10)",

        "Middle (11-20)"

    ],

    default="End (21-Month End)"

)



period_summary = (

    data.groupby("Month Period")

    .size()

    .reindex([

        "Beginning (1-10)",

        "Middle (11-20)",

        "End (21-Month End)"

    ])

    .reset_index(name="Total Files")

)



print("\nMONTH PERIOD ANALYSIS")



display(period_summary)
# =====================================================================# 9. HISTORICAL TOP 10 BUSIEST 3-DAY PERIODS# =====================================================================
historical_windows = []



for month_period in sorted(

    data["Received_Date"].dt.to_period("M").unique()

):



    month_data = data[

        data["Received_Date"].dt.to_period("M")

        == month_period

    ]



    dates = pd.date_range(

        month_period.start_time.normalize(),

        month_period.end_time.normalize(),

        freq="D"

    )



    counts = (

        month_data.groupby("Received_Date")

        .size()

        .reindex(dates, fill_value=0)

    )



    for i in range(len(dates) - 2):



        historical_windows.append({



            "Start Date": dates[i],

            "End Date": dates[i + 2],

            "Month": str(month_period),

            "3-Day File Count": int(

                counts.iloc[i:i + 3].sum()

            )



        })





historical_top = pd.DataFrame(historical_windows)



historical_top = historical_top.sort_values(

    "3-Day File Count",

    ascending=False

).head(10)



print("\nHISTORICAL TOP 10 BUSIEST 3-DAY PERIODS")



display(historical_top)
# =====================================================================# 10. CONTINUOUS DAILY SERIES# =====================================================================
daily_series = (

    data.groupby("Received_Date")

    .size()

    .sort_index()

)



all_dates = pd.date_range(

    daily_series.index.min(),

    daily_series.index.max(),

    freq="D"

)



daily_series = daily_series.reindex(

    all_dates,

    fill_value=0

)



daily_series.index.name = "Date"
# =====================================================================# 11. MACHINE LEARNING FEATURES# =====================================================================
def create_features(date, past_values):



    date = pd.Timestamp(date)



    return {



        "Weekday": date.weekday(),



        "Day": date.day,



        "Month": date.month,



        "DayOfYear": date.dayofyear,



        "WeekOfYear": int(

            date.isocalendar().week

        ),



        "IsWeekend": int(

            date.weekday() >= 5

        ),



        "MonthSin": np.sin(

            2 * np.pi * date.month / 12

        ),



        "MonthCos": np.cos(

            2 * np.pi * date.month / 12

        ),



        "DaySin": np.sin(

            2 * np.pi * date.day / 31

        ),



        "DayCos": np.cos(

            2 * np.pi * date.day / 31

        ),



        "Lag1": past_values[-1],



        "Lag7": past_values[-7],



        "Lag14": past_values[-14],



        "Rolling7": np.mean(

            past_values[-7:]

        ),



        "Rolling14": np.mean(

            past_values[-14:]

        )



    }





values = daily_series.values.astype(float)



X_list = []

y_list = []

date_list = []



for i in range(14, len(values)):



    X_list.append(

        create_features(

            daily_series.index[i],

            values[:i].tolist()

        )

    )



    y_list.append(values[i])



    date_list.append(

        daily_series.index[i]

    )





X = pd.DataFrame(X_list)



y = pd.Series(y_list, dtype=float)



training_dates = pd.Series(

    pd.to_datetime(date_list)

)



if len(X) < 30:



    raise ValueError(

        "Not enough historical daily data. "

        "Please upload more complete monthly reports."

    )
# =====================================================================# 12. MODEL CONFIGURATION# =====================================================================
def build_random_forest():



    return RandomForestRegressor(



        n_estimators=400,



        max_depth=8,



        min_samples_leaf=3,



        max_features=0.8,



        random_state=42,



        n_jobs=-1



    )





def build_extra_trees():



    return ExtraTreesRegressor(



        n_estimators=400,



        max_depth=8,



        min_samples_leaf=3,



        max_features=0.8,



        random_state=42,



        n_jobs=-1



    )
# =====================================================================# 13. TEMPORAL BACKTESTING# =====================================================================
last_month = daily_series.index.max().to_period("M")



validation_start = last_month.start_time



train_mask = (

    training_dates < validation_start

)



X_train = X.loc[train_mask].copy()



y_train = y.loc[train_mask].copy()



validation_dates = pd.date_range(

    validation_start,

    daily_series.index.max(),

    freq="D"

)



actual_validation = daily_series.reindex(

    validation_dates,

    fill_value=0

).values.astype(float)
# Recursive prediction helper
def recursive_predict(model, dates, initial_history):



    past = list(initial_history)



    predictions = []



    for date in dates:



        feature_row = pd.DataFrame([

            create_features(date, past)

        ])



        prediction = float(

            model.predict(feature_row)[0]

        )



        prediction = max(0, prediction)



        predictions.append(prediction)



        past.append(prediction)



    return np.array(predictions)
# WAPE helper
def calculate_wape(actual, predicted):



    denominator = np.sum(np.abs(actual))



    if denominator == 0:

        return np.nan



    return (

        np.sum(np.abs(actual - predicted))

        / denominator

    ) * 100
# =====================================================================# 14. COMPARE MODELS# =====================================================================
model_results = []



candidate_models = {



    "Random Forest": build_random_forest(),



    "Extra Trees": build_extra_trees()



}



initial_history = daily_series[

    daily_series.index < validation_start

].tolist()



if len(X_train) < 30:



    raise ValueError(

        "Insufficient training data before validation month."

    )



print("\n" + "=" * 70)

print("MACHINE LEARNING MODEL COMPARISON")

print("=" * 70)



for model_name, test_model in candidate_models.items():



    print("\nTesting:", model_name)



    test_model.fit(

        X_train,

        y_train

    )



    predictions = recursive_predict(

        test_model,

        validation_dates,

        initial_history

    )



    model_mae = mean_absolute_error(

        actual_validation,

        predictions

    )



    model_wape = calculate_wape(

        actual_validation,

        predictions

    )



    model_score = (

        max(0, 100 - model_wape)

        if not np.isnan(model_wape)

        else np.nan

    )



    model_results.append({



        "Model": model_name,



        "MAE": model_mae,



        "WAPE": model_wape,



        "Performance Indicator": model_score



    })
# Seasonal Naive baseline: same weekday as previous week
baseline_predictions = []



for date in validation_dates:



    previous_date = date - pd.Timedelta(days=7)



    if previous_date in daily_series.index:



        baseline_predictions.append(

            float(daily_series.loc[previous_date])

        )



    else:



        baseline_predictions.append(0.0)





baseline_predictions = np.array(

    baseline_predictions

)



baseline_mae = mean_absolute_error(

    actual_validation,

    baseline_predictions

)



baseline_wape = calculate_wape(

    actual_validation,

    baseline_predictions

)



model_results.append({



    "Model": "Seasonal Naive",



    "MAE": baseline_mae,



    "WAPE": baseline_wape,



    "Performance Indicator": max(

        0, 100 - baseline_wape

    )



})



model_comparison = pd.DataFrame(

    model_results

)



model_comparison = model_comparison.sort_values(

    "WAPE",

    ascending=True

)



print("\nMODEL COMPARISON RESULTS")



display(

    model_comparison.round(2)

)
# =====================================================================# 15. SELECT BEST MACHINE LEARNING MODEL# =====================================================================
ml_results = model_comparison[

    model_comparison["Model"].isin(

        ["Random Forest", "Extra Trees"]

    )

].copy()



best_model_name = ml_results.sort_values(

    "WAPE"

).iloc[0]["Model"]



print("\nSelected Machine Learning Model:", best_model_name)



if best_model_name == "Random Forest":



    final_model = build_random_forest()



else:



    final_model = build_extra_trees()
# =====================================================================# 16. FINAL MODEL TRAINING# =====================================================================
final_model.fit(

    X,

    y

)



print("\nFinal model trained successfully.")
# =====================================================================# 17. MODEL PERFORMANCE REPORT# =====================================================================
best_result = ml_results[

    ml_results["Model"] == best_model_name

].iloc[0]



best_wape = float(best_result["WAPE"])



best_mae = float(best_result["MAE"])



performance_score = max(

    0,

    100 - best_wape

)



print("\n" + "=" * 70)

print("MODEL PERFORMANCE REPORT")

print("=" * 70)



print("Selected Model:", best_model_name)



print("Validation Month:", last_month.strftime("%B %Y"))



print("MAE:", round(best_mae, 2), "files per day")



print("WAPE:", round(best_wape, 2), "%")



print(

    "WAPE-Based Performance Indicator:",

    round(performance_score, 2),

    "%"

)



if performance_score >= 80:



    print("Performance Status: Strong historical fit")



elif performance_score >= 65:



    print("Performance Status: Good historical fit")



elif performance_score >= 50:



    print("Performance Status: Moderate historical fit")



else:



    print("Performance Status: Needs improvement")
# =====================================================================# 18. SELECT PREDICTION MONTH# =====================================================================
default_month = (

    last_month + 1

).strftime("%Y-%m")



target_input = input(

    "\nEnter prediction month YYYY-MM "

    f"(Press Enter for {default_month}): "

).strip()



target_month = (

    target_input if target_input else default_month

)



try:



    target_start = pd.to_datetime(

        target_month,

        format="%Y-%m"

    )



    target_end = (

        target_start + pd.offsets.MonthEnd(0)

    )



except ValueError:



    raise ValueError(

        "Enter prediction month in YYYY-MM format."

    )





if target_start.to_period("M") <= last_month:



    raise ValueError(

        "Select a month after the latest historical month."

    )



print(

    "\nForecast Month:",

    target_start.strftime("%B %Y")

)
# =====================================================================# 19. PREDICT DAILY FILE VOLUME# =====================================================================
future_dates = pd.date_range(



    daily_series.index.max() + pd.Timedelta(days=1),



    target_end,



    freq="D"



)



future_predictions = recursive_predict(



    final_model,



    future_dates,



    daily_series.tolist()



)



all_forecast = pd.DataFrame({



    "Date": future_dates,



    "Predicted Files": future_predictions



})



forecast = all_forecast[



    (all_forecast["Date"] >= target_start)



    &



    (all_forecast["Date"] <= target_end)



].copy()





forecast["Predicted Files"] = (



    forecast["Predicted Files"]



    .round()



    .astype(int)



)



forecast["Weekday"] = (

    forecast["Date"].dt.day_name()

)



forecast["Period"] = np.select(



    [



        forecast["Date"].dt.day <= 10,



        forecast["Date"].dt.day <= 20



    ],



    [



        "Beginning",



        "Middle"



    ],



    default="End"



)
# =====================================================================# 20. MONTHLY PREDICTED VOLUME# =====================================================================
estimated_monthly_total = int(

    forecast["Predicted Files"].sum()

)



print("\n" + "=" * 70)

print("ESTIMATED MONTHLY FILE VOLUME")

print("=" * 70)



print(

    f"{target_start.strftime('%B %Y')}: "

    f"{estimated_monthly_total:,} files"

)
# =====================================================================# 21. TOP FOUR NON-OVERLAPPING BUSY 3-DAY PERIODS# =====================================================================
candidates = []



for i in range(len(forecast) - 2):



    window = forecast.iloc[i:i+3]



    candidates.append({



        "Start": window.iloc[0]["Date"],



        "End": window.iloc[2]["Date"],



        "Files": int(

            window["Predicted Files"].sum()

        ),



        "Dates": set(

            window["Date"].dt.date

        )



    })





candidates = sorted(



    candidates,



    key=lambda x: x["Files"],



    reverse=True



)



selected = []



for candidate in candidates:



    if all(



        candidate["Dates"].isdisjoint(

            item["Dates"]

        )



        for item in selected



    ):



        selected.append(candidate)



    if len(selected) == 4:

        break





prediction_table = pd.DataFrame([



    {



        "Priority": i + 1,



        "Predicted Date Range": (



            item["Start"].strftime("%d %b")



            + " - "



            + item["End"].strftime("%d %b %Y")



        ),



        "Expected File Count": item["Files"]



    }



    for i, item in enumerate(selected)



])
# =====================================================================# 22. DAILY PREDICTION# =====================================================================
print("\nDAILY FILE VOLUME PREDICTION")



display(



    forecast[



        [



            "Date",



            "Weekday",



            "Period",



            "Predicted Files"



        ]



    ]



)
# =====================================================================# 23. TOP FOUR PREDICTED BUSY DATE RANGES# =====================================================================
print("\n" + "=" * 70)



print("TOP 4 PREDICTED BUSY DATE RANGES")



print("=" * 70)



display(prediction_table)
# =====================================================================# 24. FEATURE IMPORTANCE# =====================================================================
importance = pd.DataFrame({



    "Feature": X.columns,



    "Importance": final_model.feature_importances_



}).sort_values(



    "Importance",



    ascending=False



)



print("\nTOP 10 FACTORS AFFECTING PREDICTION")



display(

    importance.head(10).round(4)

)
# =====================================================================# 25. MANAGEMENT READY SUMMARY# =====================================================================
print("\n" + "=" * 70)



print("MANAGEMENT FILE VOLUME FORECAST SUMMARY")



print("=" * 70)



print("Forecast Month:", target_start.strftime("%B %Y"))



print("Selected Model:", best_model_name)



print(

    "Estimated Total Files:",

    f"{estimated_monthly_total:,}"

)



print(

    "Historical Validation WAPE:",

    f"{best_wape:.2f}%"

)



print(

    "WAPE-Based Performance Indicator:",

    f"{performance_score:.2f}%"

)



print(

    "Average Daily Error:",

    f"{best_mae:,.2f} files"

)



print("\nTOP 4 BUSIEST PREDICTED DATE RANGES")



display(prediction_table)
# =====================================================================# 26. DATA PROCESSING SUMMARY# =====================================================================
print("\nDATA PROCESSING SUMMARY")



print("Total eligible rows:", f"{len(data):,}")



print("Duplicate Submission IDs: INCLUDED")



print("Blank Submission IDs: EXCLUDED")



print("Invalid Received Dates: EXCLUDED")



print(

    "Status Filter:",

    "Complete / Completed / QA Uploaded"

)



print(

    "Status Filter Applied:",

    "Only when Status column exists"

)



print("\nSHEET-WISE PROCESSING SUMMARY")



display(

    pd.DataFrame(processing_summary)

)



print("\n" + "=" * 70)



print("FORECAST COMPLETED SUCCESSFULLY")



print("=" * 70)
