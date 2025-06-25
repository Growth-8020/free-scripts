# Google Ads Automation Scripts 📊

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Ads Scripts](https://img.shields.io/badge/Google%20Ads-Scripts-4285F4?logo=google-ads)](https://developers.google.com/google-ads/scripts)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](http://makeapullrequest.com)

Free, production-ready Google Ads scripts to automate your reporting and data management. Save hours of manual work with automated Google Sheets exports and email reports.

**Built with ❤️ by Growth 8020** | [Website](https://growth8020.com) | [Contact Us](https://growth8020.com/contact)

## 🚀 What These Scripts Do

Our scripts help you automate the most time-consuming parts of Google Ads management:

- **📈 Automated Data Exports** - Push campaign, keyword, and ad performance data directly to Google Sheets
- **📧 Email Reports** - Send beautifully formatted performance reports to clients or stakeholders
- **⏰ Scheduled Monitoring** - Set up alerts for performance changes, budget pacing, and anomalies
- **🎯 Custom Dashboards** - Build live dashboards in Google Sheets that update automatically

## 📁 Available Scripts

### 1. **Performance Dashboard Exporter**
Exports comprehensive account performance data to Google Sheets with automatic formatting.
- Account, campaign, and ad group level metrics
- Customizable date ranges
- Automatic chart generation
- [View Script →](./performance-dashboard-exporter.js)

### 2. **Weekly Performance Email Report**
Sends formatted HTML email reports with key performance metrics and insights.
- Week-over-week comparisons
- Top performers and underperformers
- Budget pacing alerts
- [View Script →](./scripts/weekly-email-report.js)

### 3. **Budget Monitor & Alerts**
Monitors daily spend and sends alerts when accounts are over/under pacing.
- Real-time budget tracking
- Customizable alert thresholds
- Multi-account support
- [View Script →](./scripts/budget-monitor.js)

### 4. **Search Query Report Builder**
Automates search query analysis and exports to Google Sheets.
- New query identification
- Performance metrics by query
- Negative keyword suggestions
- [View Script →](./scripts/search-query-report.js)

### 5. **Multi-Account Performance Aggregator**
Consolidates data from multiple accounts into a single dashboard.
- MCC-level reporting
- Account comparison metrics
- Unified performance view
- [View Script →](./scripts/multi-account-aggregator.js)

## 🛠️ Quick Start Guide

### Prerequisites
- Google Ads account with Scripts access
- Google Sheets (for data export scripts)
- Basic JavaScript knowledge (helpful but not required)

### Installation

1. **Copy the Script**
   - Navigate to the script you want to use
   - Copy the entire code

2. **Add to Google Ads**
   - Go to **Tools & Settings** > **Bulk Actions** > **Scripts**
   - Click the **+** button to create a new script
   - Paste the code
   - Name your script

3. **Configure Settings**
   - Update the configuration variables at the top of each script:
   ```javascript
   // Example configuration
   var SPREADSHEET_URL = 'YOUR_SHEET_URL_HERE';
   var EMAIL_RECIPIENTS = 'email@example.com';
   var DATE_RANGE = 'LAST_7_DAYS';
   ```

4. **Authorize & Test**
   - Click **Preview** to test the script
   - Authorize access to Google Sheets/Gmail when prompted
   - Review the logs for any errors

5. **Schedule (Optional)**
   - Click **Frequency** to set up automatic runs
   - Recommended: Daily at 6 AM for reports, hourly for monitoring

## 📋 Script Configuration

Each script includes detailed configuration options. Here's what you can customize:

### Common Settings
```javascript
// Spreadsheet settings
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID';
var SHEET_NAME = 'Performance Data';

// Email settings
var EMAIL_RECIPIENTS = 'client@example.com, team@example.com';
var EMAIL_SUBJECT = 'Weekly Google Ads Performance Report';

// Date ranges (options: TODAY, YESTERDAY, LAST_7_DAYS, LAST_30_DAYS, THIS_MONTH, LAST_MONTH)
var DATE_RANGE = 'LAST_7_DAYS';

// Metrics to include
var METRICS = ['Impressions', 'Clicks', 'Cost', 'Conversions', 'CostPerConversion'];
```

## 📊 Example Output

### Google Sheets Export
![Sheet Example](https://via.placeholder.com/800x400?text=Example+Google+Sheets+Dashboard)

### Email Report Preview
```html
Weekly Performance Summary
━━━━━━━━━━━━━━━━━━━━━━━━
📈 Total Spend: $12,456.78 (↑ 15.3%)
🎯 Conversions: 234 (↑ 22.1%)
💰 CPA: $53.23 (↓ 5.8%)
🔥 CTR: 3.45% (↑ 0.23%)
```

## 🔧 Troubleshooting

### Common Issues

**"Authorization required" error**
- Click Authorize and follow the prompts
- Make sure to allow access to Google Sheets/Gmail

**"Spreadsheet not found" error**
- Check that the spreadsheet URL is correct
- Ensure the Google Ads account has access to the sheet

**No data appearing**
- Verify the account has data for the selected date range
- Check that campaigns are not filtered out

### Getting Help
- 📖 Check our [FAQ](./FAQ.md)
- 📧 Email us at hello@growth8020.com

## 📈 Success Stories

> "These scripts saved our team 10+ hours per week on reporting. Game changer!"
> — *Marketing Director, E-commerce Company*

> "Finally, automated reports that actually look professional. Our clients love them."
> — *Agency Owner*

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Attribution Required**: When using these scripts, please maintain the attribution comments in the code.

## 🌟 About Growth 8020

We're a performance marketing agency specializing in Google Ads, data automation, and conversion optimization. We've managed over $XX million in ad spend and love sharing our tools with the community.

### Our Services
- ✅ Google Ads Management
- ✅ Marketing Automation
- ✅ Custom Script Development
- ✅ Performance Consulting

**Ready to scale your Google Ads?** [Let's talk!](https://growth8020.com/contact)

---

### 📬 Stay Updated

Get notified about new scripts and updates:
- ⭐ Star this repository
- 👁️ Watch for releases
- 🐦 Follow us on [LinkedIn](https://www.linkedin.com/company/growth-8020)

---

Made with ☕ and late nights by the Growth 8020 team
