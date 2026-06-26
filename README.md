# E-Commerce Behavioral & Sales Performance Dashboard

### ⚠️ Important Notice Regarding the Data File

The raw dataset (`RawData.csv`) is around 300MB and couldn't be uploaded to GitHub due to file size limits. To run the project:

1. Download the file from Google Drive [by clicking here](https://drive.google.com/file/d/1-MrmZYw3Cii04jcRqebvvHsUAkzy5SIe/view?usp=drive_link).
2. Place it inside the `Resources > Data` folder.
3. Keep the file named exactly: `RawData.csv`

###  How to Connect the Data and Refresh

Once the file is in place:

1. Open `E-Commerce Behavioral & Sales Performance Dashboard.xlsx`
2. Go to **Data** → **Queries & Connections**
3. Double-click the **Data** query to open Power Query Editor
4. Go to **Home** → **Data Source Settings** → **Change Source**
5. Select the folder where you saved the dataset
6. Click **OK** then **Close & Apply**

---

## The Problem

Working with 2.6M+ rows in Excel isn't just slow — it breaks things. Three issues came up immediately:

- **Scale:** That volume crashes standard Excel before you even start
- **Dirty Data:** The raw file had large gaps of nulls, blank fields, and inconsistent text formatting that made it unusable as-is
- **Latency:** Heavy formulas on unoptimized data made any filtering feel like waiting for a page to load

## The Solution

I treated Excel like a small data warehouse rather than a spreadsheet:

- Ran a full ETL process in Power Query to clean the raw data before it ever touched the model
- Built a Star Schema in Power Pivot to cut memory usage and keep the data structured
- Replaced cell formulas with DAX measures to keep calculations fast under filtering
- Locked Pivot Table layout settings to stop the dashboard from collapsing when slicers change
- End result: slicers cross-filter 2.6M+ rows in under 3 seconds

---

## Step 1: Power Query & Data Pipeline

I used Power Query to build a cleaning pipeline for the raw data. The order of steps here actually matters — doing things in the wrong sequence causes Excel to run out of memory halfway through.

### 1. Connecting to the Data Folder

Instead of importing a single file, I pointed Excel at the entire data folder using the Folder Connector. This means if a new file gets dropped into that folder later, the pipeline picks it up and processes it automatically. I also added a filter to ignore hidden system files that were causing errors.

### 2. Splitting Date and Time

The `event_time` column originally combined date and time in one field. I split them into two separate columns. A combined datetime column creates millions of unique values which tanks Excel's compression. Splitting them shrinks the file and lets the date column link cleanly to a Calendar table later. I also filtered out rows with the default date `1970-01-01` since those were corrupted records.

### 3. Early Deduplication

I ran `Table.Distinct` early in the pipeline, right after the split and before any text cleaning. The logic is simple: fewer rows means every transformation that follows runs faster.

### 4. Text Cleaning

Applied `Trim`, `Clean`, and `Proper Case` to the text columns like `brand` and `category_code`. This removes hidden spaces left over from web scraping and fixes capitalization inconsistencies so `"samsung"` and `"Samsung"` stop being treated as two different brands.

### 5. Null Handling

For text and category columns, I replaced nulls and blanks with `"MissingInfo"` so they show up visibly in slicers instead of just disappearing. For numerical columns like `price` and time columns, I left them blank on purpose — putting a fake zero in a price field would break every average and total in the dashboard.

### 6. ID Prefix Fix

I added an `id-` prefix to all ID columns to force Excel to read them as text and stop it from trying to sum them. One small bug came out of this: blanks that had been replaced with `"MissingInfo"` became `"id-MissingInfo"`. I added a final step to catch and fix those back to plain `"MissingInfo"`.

### Complete Power Query M Code

```powerquery
let
    source = Folder.Files("C:\Projects\E-Commerce Behavioral & Sales Performance Dashboard\E-Commerce Behavioral & Sales Performance Dashboard\Resources\Data"),
    #"Filtered Hidden Files1" = Table.SelectRows(source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", { "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Change type (defult steps)" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"event_time", type text}, {"order_id", Int64.Type}, {"product_id", Int64.Type}, {"category_id", Int64.Type}, {"category_code", type text}, {"brand", type text}, {"price", type number}, {"user_id", Int64.Type}}),
    #"event column split" = Table.SplitColumn(#"Change type (defult steps)", "event_time", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"event_date", "event_time"}),
    #"Reordered Columns" = Table.ReorderColumns(#"event column split",{"order_id", "product_id", "category_id", "user_id", "category_code", "brand", "event_date", "event_time", "price"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"order_id", type text}, {"product_id", type text}, {"category_id", type text}, {"user_id", type text}, {"category_code", type text}, {"brand", type text}, {"event_date", type date}, {"event_time", type time}, {"price", type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([event_date] <> #date(1970, 1, 1))),
    #"Removed Duplicates" = Table.Distinct(#"Filtered Rows"),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Duplicates",{{"order_id", Text.Trim, type text}, {"product_id", Text.Trim, type text}, {"category_id", Text.Trim, type text}, {"user_id", Text.Trim, type text}, {"category_code", Text.Trim, type text}, {"brand", Text.Trim, type text}}),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"order_id", Text.Clean, type text}, {"product_id", Text.Clean, type text}, {"category_id", Text.Clean, type text}, {"user_id", Text.Clean, type text}, {"category_code", Text.Clean, type text}, {"brand", Text.Clean, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Cleaned Text",{{"order_id", Text.Proper, type text}, {"product_id", Text.Proper, type text}, {"category_id", Text.Proper, type text}, {"user_id", Text.Proper, type text}, {"category_code", Text.Proper, type text}, {"brand", Text.Proper, type text}}),
    #"Remove blanks" = Table.ReplaceValue(#"Capitalized Each Word","","MissingInfo",Replacer.ReplaceValue,{"order_id", "product_id", "category_id", "user_id", "category_code", "brand"}),
    #"Remove Nulls" = Table.ReplaceValue(#"Remove blanks",null,"MissingInfo",Replacer.ReplaceValue,{"order_id", "product_id", "category_id", "user_id", "category_code", "brand"}),
    #"Added Prefix" = Table.TransformColumns(#"Remove Nulls", {{"order_id", each "id-" & _, type text}}),
    #"Added Prefix1" = Table.TransformColumns(#"Added Prefix", {{"product_id", each "id-" & _, type text}}),
    #"Added Prefix2" = Table.TransformColumns(#"Added Prefix1", {{"category_id", each "id-" & _, type text}}),
    #"Added Prefix3" = Table.TransformColumns(#"Added Prefix2", {{"user_id", each "id-" & _, type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Added Prefix3","id-MissingInfo","MissingInfo",Replacer.ReplaceText,{"order_id", "product_id", "category_id", "user_id"})
in
    #"Replaced Value"
```

---

## Step 2: Data Modeling (Power Pivot & Star Schema)

### Why a Star Schema?

Keeping 2.6M rows in one flat table would crash Excel. The bigger problem is that a flat table repeats long text strings like brand names and category codes millions of times, which destroys memory. Breaking the data into a Fact table and smaller Dimension tables solves both problems.

### How I Split the Tables

I used Power Query's Reference feature to create three separate tables from the same clean master query:

- **Dim_Products:** A unique product catalog with `product_id`, `category_id`, `category_code`, and `brand`. Removing duplicates on `product_id` dropped the table from 2.6M rows down to ~23,000 unique products.
- **Fact_Sales:** The transaction log with `order_id`, `product_id`, `user_id`, `event_date`, `event_time`, and `price`. All descriptive text columns stripped out — only keys and numbers stay here.
- **Dim_Date:** A continuous calendar table with Year, Month, and Day columns. Required for DAX time-intelligence functions to work correctly.

### Relationships

Connected in Power Pivot using 1-to-Many relationships:
- `Dim_Products[product_id]` → `Fact_Sales[product_id]`
- `Dim_Date[event_date]` → `Fact_Sales[event_date]`

### Why This Made the Dashboard Fast

Storing text in `Dim_Products` lets Excel's VertiPaq engine use dictionary encoding, which cuts RAM usage significantly. It also means if a brand name needs to change, you update one row and it reflects across all 2.6M records automatically.

<p align="center">
  <img src="Resources/PowerPivot relations diagram.png" alt="PowerPivot relations diagram" width="100%">
</p>
---

## Step 3: Dashboard Structure & Analytical Hubs

I split the dashboard into 4 pages instead of cramming everything onto one screen. Each page focuses on one angle of the business.

---

### Hub 1: Financial Overview

The main revenue and sales page. Built for management-level decisions around overall performance, top suppliers, and seasonal trends.


<img width="1889" height="911" alt="Financial overview" src="https://github.com/user-attachments/assets/130ee7ad-a5f0-4e63-83ad-fae9f7e8bc8b" />


**KPI Cards:**

<img width="439" height="926" alt="Financial overview cards" src="https://github.com/user-attachments/assets/ec9e4aed-c7cb-40b4-b27f-0b7561b06f63" />


| Metric | Value |
|--------|-------|
| Total Revenue | $336,963,381 |
| Total Units Sold | 2,613,215 |
| Total Orders | 1,426,253 |
| Average Order Value (AOV) | $236 |

**Charts:**

**Top 10 Brands by Revenue**
<img width="529" height="401" alt="Top 10 Brands Column Chart" src="https://github.com/user-attachments/assets/ddfe1330-55f6-458b-9ffa-c6936ca65114" />

**Top 10 Categories by Revenue**
<img width="740" height="402" alt="Top 10 Categories Bar Chart" src="https://github.com/user-attachments/assets/79267cf3-cee6-4c80-b6d6-06f6bdea85df" />

**Revenue Trend by Month**
<img width="1265" height="474" alt="Revenue Trend Line Chart" src="https://github.com/user-attachments/assets/cb888bd5-6214-4823-b3c5-178f34372eaa" />

**Slicers:** Month and Category — both cross-filter the entire page.

---

### Hub 2: Customer Behavior

Tracks user habits and platform friction. The 73.5% guest checkout rate is the most interesting number on this page — it points to a friction problem in the registration flow.

<img width="1859" height="887" alt="customer behavior" src="https://github.com/user-attachments/assets/e580a158-3bec-411c-bb9e-29134299ca72" />


**KPI Cards:**

<img width="353" height="558" alt="Customer Behavior Cards" src="https://github.com/user-attachments/assets/0e2ca810-a0dc-44fb-b255-0c683f5dd61a" />

| Metric | Value |
|--------|-------|
| Total Unique Customers | 233,576 |
| Purchase Frequency | 6.11 |
| Guest Checkout Rate | 73.5% |

**Charts:**

**Registered vs. Guest Revenue**
<img width="832" height="690" alt="Registered vs Guest Pie Chart" src="https://github.com/user-attachments/assets/a9c4d49b-0501-4719-84dd-7e62521ab987" />

**Top 10 Customers by Spend**
<img width="586" height="674" alt="Top 10 Customers Table" src="https://github.com/user-attachments/assets/a4651721-9b65-4508-9e6a-32bb66ec6567" />

---

### Hub 3: Product Intelligence

Focuses on inventory distribution, price ranges, and which products actually move volume.

<img width="1873" height="891" alt="Product intelligence" src="https://github.com/user-attachments/assets/90b6ca6a-5ca2-4fb2-b5c8-8dd5666d48ca" />

**KPI Cards:**

<img width="1001" height="150" alt="Product Intelligence Cards" src="https://github.com/user-attachments/assets/6dd0d7bf-188b-4375-b837-41c8d2ee1e5c" />

| Metric | Value |
|--------|-------|
| Top Category | Computers.Notebook |
| Top Brand | Samsung |
| Average Item Price | $154.19 |

**Charts:**

**Price vs. Volume Scatter Plot** — shows price elasticity visually
<img width="649" height="417" alt="Price vs Volume Scatter Plot" src="https://github.com/user-attachments/assets/9c57bd23-9e75-4ca5-b7f7-45610d4823b3" />

**Top 50 Categories Treemap** — product assortment density at a glance
<img width="1009" height="734" alt="Category Breakdown Treemap" src="https://github.com/user-attachments/assets/e92969b4-51b9-4c12-8fac-42a1ee8b167d" />

**Top 10 Brands by Unit Volume**
<img width="645" height="466" alt="Top 10 Brands by Unit Volume" src="https://github.com/user-attachments/assets/611b2ccd-9760-432e-a490-5af2cb56f80e" />

---

### Hub 4: Temporal Analytics

Time-based patterns: when do people buy, which hours are busiest, and how revenue shifts across the week and year.

<img width="1868" height="894" alt="temporal analysis" src="https://github.com/user-attachments/assets/2a5b6bad-4e00-466f-ad13-ab1807b0ad85" />


**KPI Cards:**
<img width="898" height="139" alt="Temporal Analytics cards" src="https://github.com/user-attachments/assets/5eef5e5b-78c3-4765-8bed-0a8d9f68a055" />

| Metric | Value |
|--------|-------|
| Peak Hour Daily | 10:00 |
| Busiest Month | September |
| Busiest Day | 7 |

**Charts:**

**Peak Hours / Days Heatmap (24×7 Revenue Density Matrix)**

Revenue concentrates between 08:00–13:00, peaking at 10:00–11:00, especially on Fridays and Saturdays. Late night (20:00–04:00) is consistently the lowest across all days.

<img width="914" height="742" alt="HeatMap" src="https://github.com/user-attachments/assets/0f358484-5eef-4d44-b238-87a7ebe117bd" />


**Sales by Day of Week**

Thursday leads (~$51.6M), Wednesday is the lowest (~$44.1M).

<img width="960" height="402" alt="Sales by week" src="https://github.com/user-attachments/assets/f69e5114-75b2-438b-9070-932550012ac6" />


**Monthly Revenue + MoM Growth (Combo Chart)**

August is the highest month (~$53M). May shows the sharpest growth acceleration.

<img width="969" height="359" alt="Combo chart" src="https://github.com/user-attachments/assets/380aec55-ee0d-46f5-832c-a6b2c249ff42" />


---

## What's Next

This is my second project. The pipeline and data model are in good shape, but the visuals and some DAX measures still have room to improve. Next steps are going deeper into time-intelligence DAX patterns and eventually moving the cleaning layer to Python.
