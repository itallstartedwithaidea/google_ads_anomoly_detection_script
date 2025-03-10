# Google Ads Anomaly Detection Script

This repository contains a Google Ads Script that monitors and identifies potential anomalies within your Google Ads campaigns. The script uses a linked Google Spreadsheet to store configuration settings, exclusion lists, and logs. It also generates anomaly reports for various performance metrics such as impressions, clicks, conversions, spend, bids, budgets, ad performance, keywords, and more.

---

## Features

1. **Configurable Thresholds**  
   - Easily set thresholds for CTR drop, impression drop, conversion rate drop, spend changes, and other key metrics in the `Config` sheet of the linked Google Spreadsheet.

2. **Exclusion Lists**  
   - Exclude specific campaigns, ad groups, or keywords from anomaly detection using the `Exclusions` sheet.

3. **Automated Reporting**  
   - When executed, the script:
     - Identifies anomalies across multiple dimensions (impressions, clicks, conversions, spend, bids, budgets, impression share, quality score, device performance, etc.).
     - Logs each run to a `Log` sheet.
     - Creates separate sheets for each anomaly type (e.g., `Impression Anomalies`, `Click Anomalies`, etc.).
     - Builds a `Summary` sheet showing a quick overview of anomalies and severity.
     - Generates a `Charts` sheet with visualizations of anomaly counts.
     - Provides `Trends` sheet information to track improving/worsening trends.

4. **Email Notifications**  
   - Sends an email summary to a configurable recipient with an overview of the anomalies detected.
   - Sends an error email if the script fails.

5. **Configurable Time Constraints**  
   - You can specify a maximum execution time. The script checks its runtime throughout to prevent exceeding limits.

6. **Historical Data Storage (Optional)**  
   - Demonstrates how to store and retrieve historical data for deeper trend analysis if you choose to enable those helper functions.

---

## Script Outline

- **`main()`**  
  - Orchestrates the overall anomaly detection process.
  - Loads configuration settings from your spreadsheet.
  - Executes specific detection functions for each metric type.
  - Creates dashboards and sends notifications.

- **`loadConfiguration(spreadsheet)`**  
  - Reads in thresholds and email recipient from the `Config` sheet.

- **`loadExclusions(spreadsheet)`**  
  - Reads in campaigns, ad groups, and keywords to exclude from the `Exclusions` sheet.

- **`log(spreadsheet, message)`**  
  - Appends logs to the `Log` sheet and Google Apps Script logger.

- **`detect*Anomalies(...)`**  
  - A suite of functions for each anomaly type (impression, click, conversion, spend, etc.).

- **`createSummaryDashboard(spreadsheet, anomalies)`**  
  - Summarizes anomaly counts and assigns a “severity” score to provide a quick overview.

- **`createDataVisualizations(spreadsheet, anomalies)`**  
  - Generates a simple bar chart in the `Charts` sheet for the anomaly counts by type.

- **`performTrendAnalysis(spreadsheet, anomalies, THRESHOLDS)`**  
  - Demonstrates how you might track recurring anomalies or analyze improving vs. worsening trends.

- **`checkCustomAlerts(spreadsheet, config, anomalies)`**  
  - Allows you to add custom alerts (e.g., if total anomalies exceed a certain threshold).

- **`sendEmailSummary(recipient, anomalies)`** and **`sendErrorEmail(recipient, errorMessage)`**  
  - Sends email notifications on successes or failures.

---

## Contributing

Feel free to open issues for any bug reports, suggestions, or feature requests. Contributions are welcome—fork the repo and submit a pull request!

---

## License

You can specify the license of your choice (e.g., MIT, Apache-2.0, etc.). If no license is provided here, the default assumption is “All Rights Reserved.”

