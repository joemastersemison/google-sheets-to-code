# Data Analysis Google Sheets Template

## Overview
This template demonstrates statistical analysis with functions like AVERAGE, STDEV, COUNTIF, RANK, and data aggregation.

## Google Sheets Setup Instructions

### Step 1: Create the Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Name it "Data Analysis Example"
4. Create the following sheets (tabs):
   - RawData
   - Parameters
   - Analysis
   - Summary

### Step 2: Set up "RawData" Sheet

| Cell | Column A (Sales) | Column B (Region) | Column C (Product) | Column D (Date) | Column E (Customer) |
|------|-----------------|-------------------|-------------------|-----------------|---------------------|
| 1 | Sales Amount | Region | Product | Date | Customer ID |
| 2 | 1250 | North | Widget A | 1/1/2024 | C001 |
| 3 | 2340 | South | Widget B | 1/2/2024 | C002 |
| 4 | 890 | East | Widget A | 1/3/2024 | C003 |
| 5 | 3450 | West | Widget C | 1/4/2024 | C004 |
| 6 | 1560 | North | Widget B | 1/5/2024 | C005 |
| 7 | 2780 | South | Widget A | 1/6/2024 | C006 |
| 8 | 4200 | East | Widget C | 1/7/2024 | C007 |
| 9 | 990 | West | Widget A | 1/8/2024 | C008 |
| 10 | 3100 | North | Widget B | 1/9/2024 | C009 |
| 11 | 1780 | South | Widget C | 1/10/2024 | C010 |
| 12 | 2900 | East | Widget A | 1/11/2024 | C011 |
| 13 | 3600 | West | Widget B | 1/12/2024 | C012 |
| 14 | 1450 | North | Widget C | 1/13/2024 | C013 |
| 15 | 2200 | South | Widget A | 1/14/2024 | C014 |
| 16 | 3800 | East | Widget B | 1/15/2024 | C015 |
| 17 | 920 | West | Widget C | 1/16/2024 | C016 |
| 18 | 2650 | North | Widget A | 1/17/2024 | C017 |
| 19 | 3300 | South | Widget B | 1/18/2024 | C018 |
| 20 | 1900 | East | Widget C | 1/19/2024 | C019 |
| 21 | 4100 | West | Widget A | 1/20/2024 | C020 |

*Continue with more sample data up to row 100*

### Step 3: Set up "Parameters" Sheet

| Cell | Label | Value | Description |
|------|-------|-------|-------------|
| A1 | **Analysis Parameters** | | Header |
| A2 | Confidence Level | | Label |
| B2 | 0.95 | | 95% confidence |
| A3 | Significance Threshold | | Label |
| B3 | 0.05 | | 5% significance |
| A4 | Moving Average Window | | Label |
| B4 | 7 | | 7-day window |
| A5 | Outlier Threshold (σ) | | Label |
| B5 | 3 | | 3 standard deviations |
| A6 | Minimum Sample Size | | Label |
| B6 | 30 | | Min 30 samples |
| A7 | Top N Items | | Label |
| B7 | 10 | | Top 10 analysis |
| A8 | Sales Target | | Label |
| B8 | 2500 | | $2,500 target |

### Step 4: Set up "Analysis" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | **Basic Statistics** | | Header |
| A2 | Mean Sales | | Label |
| B2 | `=AVERAGE(RawData!A2:A100)` | Average sales |
| A3 | Median Sales | | Label |
| B3 | `=MEDIAN(RawData!A2:A100)` | Median value |
| A4 | Standard Deviation | | Label |
| B4 | `=STDEV(RawData!A2:A100)` | Standard deviation |
| A5 | Variance | | Label |
| B5 | `=VAR(RawData!A2:A100)` | Variance |
| A6 | Min Sales | | Label |
| B6 | `=MIN(RawData!A2:A100)` | Minimum value |
| A7 | Max Sales | | Label |
| B7 | `=MAX(RawData!A2:A100)` | Maximum value |
| A8 | Range | | Label |
| B8 | `=B7-B6` | Data range |
| A9 | Count | | Label |
| B9 | `=COUNT(RawData!A2:A100)` | Total count |
| A10 | Count Above Target | | Label |
| B10 | `=COUNTIF(RawData!A2:A100, ">"&Parameters!B8)` | Sales above target |
| | | | |
| A12 | **Percentile Analysis** | | Header |
| A13 | 25th Percentile | | Label |
| B13 | `=PERCENTILE(RawData!A2:A100, 0.25)` | Q1 |
| A14 | 50th Percentile | | Label |
| B14 | `=PERCENTILE(RawData!A2:A100, 0.50)` | Q2 (Median) |
| A15 | 75th Percentile | | Label |
| B15 | `=PERCENTILE(RawData!A2:A100, 0.75)` | Q3 |
| A16 | 90th Percentile | | Label |
| B16 | `=PERCENTILE(RawData!A2:A100, 0.90)` | 90th percentile |
| A17 | Interquartile Range | | Label |
| B17 | `=B15-B13` | IQR |
| | | | |
| A19 | **Advanced Metrics** | | Header |
| A20 | Coefficient of Variation | | Label |
| B20 | `=B4/B2` | CV calculation |
| A21 | Skewness Indicator | | Label |
| B21 | `=IF(B2>B3, "Right Skewed", IF(B2<B3, "Left Skewed", "Symmetric"))` | Distribution shape |
| A22 | Outlier Count | | Label |
| B22 | `=SUMPRODUCT((ABS(RawData!A2:A100-B2)/B4>Parameters!B5)*1)` | Count of outliers |
| | | | |
| A24 | **Regional Analysis** | | Header |
| A25 | Region | Count | Average | Total |
| A26 | North | `=COUNTIF(RawData!B2:B100, "North")` | `=AVERAGEIF(RawData!B2:B100, "North", RawData!A2:A100)` | `=SUMIF(RawData!B2:B100, "North", RawData!A2:A100)` |
| A27 | South | `=COUNTIF(RawData!B2:B100, "South")` | `=AVERAGEIF(RawData!B2:B100, "South", RawData!A2:A100)` | `=SUMIF(RawData!B2:B100, "South", RawData!A2:A100)` |
| A28 | East | `=COUNTIF(RawData!B2:B100, "East")` | `=AVERAGEIF(RawData!B2:B100, "East", RawData!A2:A100)` | `=SUMIF(RawData!B2:B100, "East", RawData!A2:A100)` |
| A29 | West | `=COUNTIF(RawData!B2:B100, "West")` | `=AVERAGEIF(RawData!B2:B100, "West", RawData!A2:A100)` | `=SUMIF(RawData!B2:B100, "West", RawData!A2:A100)` |
| | | | |
| A31 | **Product Analysis** | | Header |
| A32 | Product | Count | Average | Rank |
| A33 | Widget A | `=COUNTIF(RawData!C2:C100, "Widget A")` | `=AVERAGEIF(RawData!C2:C100, "Widget A", RawData!A2:A100)` | `=RANK(C33, C33:C35)` |
| A34 | Widget B | `=COUNTIF(RawData!C2:C100, "Widget B")` | `=AVERAGEIF(RawData!C2:C100, "Widget B", RawData!A2:A100)` | `=RANK(C34, C33:C35)` |
| A35 | Widget C | `=COUNTIF(RawData!C2:C100, "Widget C")` | `=AVERAGEIF(RawData!C2:C100, "Widget C", RawData!A2:A100)` | `=RANK(C35, C33:C35)` |
| | | | |
| A37 | **Time Series Analysis** | | Header |
| A38 | Moving Average (7-day) | | Label |
| B38 | `=AVERAGE(OFFSET(RawData!A2, COUNT(RawData!A2:A100)-Parameters!B4, 0, Parameters!B4, 1))` | Recent moving avg |
| A39 | Trend Direction | | Label |
| B39 | `=IF(B38>B2, "Upward", IF(B38<B2, "Downward", "Stable"))` | Trend analysis |
| A40 | Growth Rate | | Label |
| B40 | `=(OFFSET(RawData!A2, COUNT(RawData!A2:A100)-1, 0)/RawData!A2)^(1/COUNT(RawData!A2:A100))-1` | CAGR |

### Step 5: Set up "Summary" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | **Executive Summary** | | Header |
| A2 | Total Sales | | Label |
| B2 | `=SUM(RawData!A2:A100)` | Total sales |
| A3 | Average Daily Sales | | Label |
| B3 | `=Analysis!B2` | Average |
| A4 | Best Performing Region | | Label |
| B4 | `=INDEX(Analysis!A26:A29, MATCH(MAX(Analysis!D26:D29), Analysis!D26:D29, 0))` | Top region |
| A5 | Best Performing Product | | Label |
| B5 | `=INDEX(Analysis!A33:A35, MATCH(MIN(Analysis!D33:D35), Analysis!D33:D35, 0))` | Top product |
| A6 | Sales Above Target (%) | | Label |
| B6 | `=Analysis!B10/Analysis!B9*100` | Percentage above target |
| | | | |
| A8 | **Data Quality Report** | | Header |
| A9 | Total Records | | Label |
| B9 | `=Analysis!B9` | Record count |
| A10 | Data Completeness (%) | | Label |
| B10 | `=COUNTA(RawData!A2:A100)/99*100` | Completeness |
| A11 | Outliers Detected | | Label |
| B11 | `=Analysis!B22` | Outlier count |
| A12 | Data Range (Days) | | Label |
| B12 | `=MAX(RawData!D2:D100)-MIN(RawData!D2:D100)` | Date range |
| | | | |
| A14 | **Statistical Tests** | | Header |
| A15 | Z-Score (Latest) | | Label |
| B15 | `=(OFFSET(RawData!A2, COUNT(RawData!A2:A100)-1, 0)-Analysis!B2)/Analysis!B4` | Z-score |
| A16 | Normal Distribution Test | | Label |
| B16 | `=IF(ABS(B15)<Parameters!B5, "Normal", "Outlier")` | Distribution test |
| A17 | Sample Size Adequate | | Label |
| B17 | `=IF(Analysis!B9>=Parameters!B6, "Yes", "No")` | Sample size check |
| | | | |
| A19 | **Top Performers** | | Header |
| A20 | Rank | Sales Value | Customer | Date |
| A21 | 1 | `=LARGE(RawData!A2:A100, 1)` | `=INDEX(RawData!E2:E100, MATCH(B21, RawData!A2:A100, 0))` | `=INDEX(RawData!D2:D100, MATCH(B21, RawData!A2:A100, 0))` |
| A22 | 2 | `=LARGE(RawData!A2:A100, 2)` | `=INDEX(RawData!E2:E100, MATCH(B22, RawData!A2:A100, 0))` | `=INDEX(RawData!D2:D100, MATCH(B22, RawData!A2:A100, 0))` |
| A23 | 3 | `=LARGE(RawData!A2:A100, 3)` | `=INDEX(RawData!E2:E100, MATCH(B23, RawData!A2:A100, 0))` | `=INDEX(RawData!D2:D100, MATCH(B23, RawData!A2:A100, 0))` |
| A24 | 4 | `=LARGE(RawData!A2:A100, 4)` | `=INDEX(RawData!E2:E100, MATCH(B24, RawData!A2:A100, 0))` | `=INDEX(RawData!D2:D100, MATCH(B24, RawData!A2:A100, 0))` |
| A25 | 5 | `=LARGE(RawData!A2:A100, 5)` | `=INDEX(RawData!E2:E100, MATCH(B25, RawData!A2:A100, 0))` | `=INDEX(RawData!D2:D100, MATCH(B25, RawData!A2:A100, 0))` |

## How to Use This Template

1. **Create the Google Sheet** with the structure above
2. **Add sample data** to the RawData sheet (at least 30-100 rows for meaningful analysis)
3. **Share the sheet** (read-only) to get a public URL
4. **Update the example JSON** file with your sheet's URL
5. **Run the converter**:
   ```bash
   npm run cli -- convert --config examples/data-analysis.json
   ```

## Expected Output

The converter will generate TypeScript or Python code that:
- Calculates all statistical measures (mean, median, stdev, percentiles)
- Performs conditional aggregations (COUNTIF, SUMIF, AVERAGEIF)
- Implements ranking and sorting logic
- Handles outlier detection and data quality checks
- Maintains cross-sheet references and dependencies

## Key Features Demonstrated

1. **Statistical Functions**: AVERAGE, MEDIAN, STDEV, VAR, PERCENTILE
2. **Conditional Aggregation**: COUNTIF, SUMIF, AVERAGEIF, COUNTIFS
3. **Ranking & Sorting**: RANK, LARGE, SMALL, INDEX/MATCH
4. **Data Analysis**: Z-scores, outlier detection, trend analysis
5. **Cross-sheet References**: Complex formulas referencing multiple sheets

## Sample Public Sheet URL

Once you create your sheet, it will have a URL like:
```
https://docs.google.com/spreadsheets/d/1XyZ123AbC456DeF789GhI/edit#gid=0
```

Make sure to:
1. Click "Share" → "Anyone with the link can view"
2. Copy the URL
3. Update the `data-analysis.json` file with this URL