# Installation & Setup Instructions

Use the following steps to install and configure the Google Ads Anomaly Detection Script in your Google Ads account and link it to a Google Spreadsheet.

---

## 1. Create a Google Spreadsheet

1. **Create a new Google Spreadsheet** in your Google Drive (or pick an existing one).
2. **Rename the spreadsheet** to something meaningful, e.g. “Ads Anomaly Detection.”
3. Create (or confirm existence of) the following sheets inside this spreadsheet:
   - **Config**  
     - Row 1 could have headers, but the script reads data starting from row 2.  
     - Below is an example of how you might structure it:

       | Key                   | Value                         |
       |-----------------------|-------------------------------|
       | CTR_DROP             | 0.20                          |
       | IMPRESSIONS_DROP     | 0.25                          |
       | CONVERSION_RATE_DROP | 0.30                          |
       | SPEND_CHANGE         | 0.50                          |
       | NO_SPEND_DAYS        | 3                             |
       | LOW_IMPRESSION_SHARE | 40                            |
       | HIGH_BID_LIMIT       | 2.50                          |
       | LOW_QUALITY_SCORE    | 5                             |
       | FIRST_PAGE_BID_RATIO | 0.80                          |
       | LOW_DEVICE_PERFORMANCE | 0.30                        |
       | MAX_EXECUTION_TIME   | 10                            |
       | LOW_SEARCH_VOLUME    | 50                            |
       | LOOKBACK_DAYS        | 7                             |
       | MAX_TOTAL_ANOMALIES  | 1000                          |
       | EMAIL_RECIPIENT      | your.email@example.com        |

     - Feel free to adjust the thresholds and add other config values as needed.

   - **Exclusions**  
     - Row 1 could have headers, but the script reads from row 2 onward.  
     - Example layout:

       | EntityType | NameOfEntity           |
       |-----------|-------------------------|
       | campaign  | Brand - Exact          |
       | adGroup   | Branded Competitor Ads |
       | keyword   | free                   |

     - This allows you to exclude specific campaigns, ad groups, or keywords from checks.

   - **Log** (optional)  
     - The script will write logs into a sheet named `Log` automatically.

   - **Historical Data** (optional)  
     - The script has helper functions to store data. If you want to enable that, create a `Historical Data` sheet. The script will manage rows and columns dynamically.

   - **Any other** (the script automatically creates other named sheets like `Impression Anomalies`, `Click Anomalies`, etc.).

4. Copy the **Spreadsheet URL** (from your browser’s address bar) and note it for use in the script configuration variable `SPREADSHEET_URL`.

---

## 2. Add the Script to Google Ads

1. **Log in** to your [Google Ads account](https://ads.google.com/).
2. Navigate to **Tools & Settings** (top menu) → **Bulk Actions** → **Scripts**.
3. **Click the “+”** button (blue plus icon) to create a new script.
4. Give your script a name, e.g. **Google Ads Anomaly Detection**.
5. **Copy and paste** the entire script (from your local file or this GitHub repository) into the Google Ads Script Editor.
6. **Set the `SPREADSHEET_URL`** at the top of the script to reference your newly created Google Spreadsheet link.

   ```js
   var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/<YOUR_SPREADSHEET_ID_HERE>';
