/**
 * Google Ads Script - Export Performance Metrics to Google Sheets
 * This script exports account, campaign, ad group, search query, country, and landing page performance data to a Google Sheet.
 *
 * Features:
 * - Includes tabs for Account Daily, Top Campaigns, Ad Groups, Search Queries, Countries, and Landing Pages.
 * - Configurable date ranges.
 * - Summary rows with totals and calculated metrics.
 * - Optional email notifications with a detailed period-over-period performance summary.
 *
 * Setup Instructions:
 * 1. Replace 'YOUR_SPREADSHEET_URL_HERE' with your Google Sheet URL.
 * 2. Adjust the DATE_RANGE if needed (default is 'LAST_30_DAYS').
 * 3. (Optional) Set SEND_EMAIL_ON_COMPLETE to true and update EMAIL_RECIPIENTS.
 * 4. Schedule the script to run as needed (e.g., daily, weekly).
 */

// ==================== CONFIGURATION ====================
const CONFIG = {
  // Replace with your Google Sheet URL
  SPREADSHEET_URL: 'YOUR_GOOGLE_SHEET_URL_HERE',
  
  // Date Range Options: TODAY, YESTERDAY, LAST_7_DAYS, LAST_14_DAYS, LAST_30_DAYS, THIS_MONTH, LAST_MONTH
  DATE_RANGE: 'LAST_30_DAYS',
  
  // Sheet names
  SHEET_NAMES: {
    ACCOUNT_DAILY: 'Account Daily',
    TOP_CAMPAIGNS: 'Top Campaigns',
    TOP_AD_GROUPS: 'Top Ad Groups',
    TOP_SEARCH_QUERIES: 'Top Search Queries',
    TOP_LANDING_PAGES: 'Top Landing Pages',
    TOP_COUNTRIES: 'Top Countries'
  },
  
  // Notification settings
  SEND_EMAIL_ON_COMPLETE: true,
  EMAIL_RECIPIENTS: 'your-email@example.com' // Comma-separated for multiple recipients
};

// ==================== MAIN FUNCTION ====================
function main() {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const summaryData = {};
    
    // --- Gather all data for the email summary ---
    const dateRange = getPriorPeriodDateRange();
    summaryData.priorTotalCost = getPriorPeriodCost(dateRange);
    summaryData.priorCampaignCount = getPriorPeriodEntityCount('campaign', 'campaign.resource_name', dateRange);
    summaryData.priorAdGroupCount = getPriorPeriodEntityCount('ad_group', 'ad_group.resource_name', dateRange);
    summaryData.priorSearchQueryCount = getPriorPeriodEntityCount('search_term_view', 'search_term_view.search_term', dateRange);
    summaryData.priorLandingPageCount = getPriorPeriodEntityCount('landing_page_view', 'landing_page_view.unexpanded_final_url', dateRange); // <-- Corrected Field
    summaryData.priorCountryCount = getPriorPeriodCountryCount(dateRange);

    clearSheets(spreadsheet);
    
    // Get current period metrics
    exportAccountDailyData(spreadsheet, summaryData);
    exportTopCampaignsData(spreadsheet, summaryData);
    exportTopAdGroupsData(spreadsheet, summaryData);
    exportTopSearchQueriesData(spreadsheet, summaryData);
    exportTopLandingPagesData(spreadsheet, summaryData);
    exportTopCountriesData(spreadsheet, summaryData);
    
    if (CONFIG.SEND_EMAIL_ON_COMPLETE) {
      sendCompletionEmail(spreadsheet.getUrl(), summaryData);
    }
    
    Logger.log('Export completed successfully!');
    
  } catch (error) {
    Logger.log('Error in main function: ' + error.toString());
    throw error;
  }
}

// ==================== PRIOR PERIOD HELPER FUNCTIONS ====================
// Functions for getPriorPeriodDateRange, getPriorPeriodCost, getPriorPeriodEntityCount,
// and getPriorPeriodCountryCount remain here. They are correct and unchanged from the previous version.
// For brevity, these functions are omitted but are assumed to be present.
function getPriorPeriodDateRange() {
    const timeZone = AdsApp.currentAccount().getTimeZone();
    const today = new Date();
    const mainPeriod_startDate = new Date();
    mainPeriod_startDate.setDate(today.getDate() - 30);
    const priorPeriod_endDate = new Date();
    priorPeriod_endDate.setTime(mainPeriod_startDate.getTime());
    priorPeriod_endDate.setDate(mainPeriod_startDate.getDate() - 1);
    const priorPeriod_startDate = new Date();
    priorPeriod_startDate.setTime(priorPeriod_endDate.getTime());
    priorPeriod_startDate.setDate(priorPeriod_endDate.getDate() - 29);
    return {
        startDate: Utilities.formatDate(priorPeriod_startDate, timeZone, 'yyyy-MM-dd').replace(/-/g, ''),
        endDate: Utilities.formatDate(priorPeriod_endDate, timeZone, 'yyyy-MM-dd').replace(/-/g, '')
    };
}

function getPriorPeriodCost(dateRange) {
    const query = `SELECT metrics.cost_micros FROM customer WHERE segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}'`;
    try {
        const report = AdsApp.report(query);
        const rows = report.rows();
        if (rows.hasNext()) {
            return (rows.next()['metrics.cost_micros'] || 0) / 1000000;
        }
    } catch(e) { Logger.log(`Could not retrieve prior period cost. Error: ${e}`); }
    return 0;
}

function getPriorPeriodEntityCount(view, selectField, dateRange) {
    const query = `SELECT ${selectField} FROM ${view} WHERE segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}' AND metrics.cost_micros > 0`;
    try {
        const report = AdsApp.report(query);
        const iterator = report.rows();
        let count = 0;
        while(iterator.hasNext()) {
            iterator.next();
            count++;
        }
        return count;
    } catch (e) {
        Logger.log(`Could not retrieve prior count for ${view}. Error: ${e}`);
    }
    return 0;
}

function getPriorPeriodCountryCount(dateRange) {
    const query = `SELECT geographic_view.country_criterion_id FROM geographic_view WHERE segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}' AND metrics.cost_micros > 0 AND geographic_view.location_type = 'LOCATION_OF_PRESENCE'`;
    try {
        const report = AdsApp.report(query);
        const countries = new Set();
        for (const row of report.rows()) {
            countries.add(row['geographic_view.country_criterion_id']);
        }
        return countries.size;
    } catch (e) { Logger.log(`Could not retrieve prior country count. Error: ${e}`); }
    return 0;
}


// ==================== SPREADSHEET FUNCTIONS ====================
// These functions are unchanged.
function getOrCreateSpreadsheet() {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  } catch (e) {
    throw new Error('Unable to open spreadsheet. Please check the URL in CONFIG.SPREADSHEET_URL');
  }
  Object.values(CONFIG.SHEET_NAMES).forEach(name => ensureSheetExists(spreadsheet, name));
  return spreadsheet;
}

function ensureSheetExists(spreadsheet, sheetName) {
  if (!spreadsheet.getSheetByName(sheetName)) {
    spreadsheet.insertSheet(sheetName);
  }
}

function clearSheets(spreadsheet) {
  Object.values(CONFIG.SHEET_NAMES).forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) sheet.clear();
  });
}

// ==================== DATA EXPORT FUNCTIONS ====================
// The existing export functions are unchanged.
// For brevity, these functions are omitted but are assumed to be present.
function exportAccountDailyData(spreadsheet, summaryData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.ACCOUNT_DAILY);
  const headers = ['Date', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const query = `SELECT segments.date, metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.conversions_from_interactions_rate, metrics.all_conversions FROM customer WHERE segments.date DURING ${CONFIG.DATE_RANGE} ORDER BY segments.date DESC`;
  
  const report = AdsApp.report(query);
  const rows = [];
  let totalCost = 0;
  
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = (row['metrics.cost_micros'] || 0) / 1000000;
    totalCost += cost;
    const convValue = row['metrics.conversions_value'] || 0;
    const conversions = row['metrics.conversions'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    const costPerConv = conversions > 0 ? cost / conversions : 0;
    
    rows.push([row['segments.date'], row['metrics.clicks'], row['metrics.impressions'], row['metrics.ctr'], row['metrics.clicks'] > 0 ? (row['metrics.average_cpc'] || 0) / 1000000 : 0, cost, convValue, roas, conversions, costPerConv, row['metrics.conversions_from_interactions_rate'], row['metrics.all_conversions']]);
  }
  
  summaryData.totalCost = totalCost;
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`);
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`);
    sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`);
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`);
    sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`);
    sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(F${summaryRow}/I${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 12).setFormula(`=SUM(L2:L${rows.length + 1})`);
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
  }
  formatSheet(sheet, headers.length, rows.length + 2);
}

function exportTopCampaignsData(spreadsheet, summaryData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_CAMPAIGNS);
  const headers = ['Campaign', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const query = `SELECT campaign.name, metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.conversions_from_interactions_rate, metrics.all_conversions FROM campaign WHERE segments.date DURING ${CONFIG.DATE_RANGE} AND campaign.status != 'REMOVED' AND metrics.cost_micros > 0 ORDER BY metrics.cost_micros DESC`;
  const report = AdsApp.report(query);
  const rows = [];
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = (row['metrics.cost_micros'] || 0) / 1000000;
    const convValue = row['metrics.conversions_value'] || 0;
    const conversions = row['metrics.conversions'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    const costPerConv = conversions > 0 ? cost / conversions : 0;
    rows.push([row['campaign.name'], row['metrics.clicks'], row['metrics.impressions'], row['metrics.ctr'], row['metrics.clicks'] > 0 ? (row['metrics.average_cpc'] || 0) / 1000000 : 0, cost, convValue, roas, conversions, costPerConv, row['metrics.conversions_from_interactions_rate'], row['metrics.all_conversions']]);
  }
  
  summaryData.campaignCount = rows.length;

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`);
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`);
    sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`);
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`);
    sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`);
    sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(F${summaryRow}/I${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 12).setFormula(`=SUM(L2:L${rows.length + 1})`);
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
  }
  formatSheet(sheet, headers.length, rows.length + 2);
}

function exportTopAdGroupsData(spreadsheet, summaryData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_AD_GROUPS);
  const headers = ['Campaign', 'Ad Group', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const query = `SELECT campaign.name, ad_group.name, metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.conversions_from_interactions_rate, metrics.all_conversions FROM ad_group WHERE segments.date DURING ${CONFIG.DATE_RANGE} AND ad_group.status != 'REMOVED' AND metrics.cost_micros > 0 ORDER BY metrics.cost_micros DESC`;
  const report = AdsApp.report(query);
  const rows = [];
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = (row['metrics.cost_micros'] || 0) / 1000000;
    const convValue = row['metrics.conversions_value'] || 0;
    const conversions = row['metrics.conversions'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    const costPerConv = conversions > 0 ? cost / conversions : 0;
    rows.push([row['campaign.name'], row['ad_group.name'], row['metrics.clicks'], row['metrics.impressions'], row['metrics.ctr'], row['metrics.clicks'] > 0 ? (row['metrics.average_cpc'] || 0) / 1000000 : 0, cost, convValue, roas, conversions, costPerConv, row['metrics.conversions_from_interactions_rate'], row['metrics.all_conversions']]);
  }

  summaryData.adGroupCount = rows.length;

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`);
    sheet.getRange(summaryRow, 4).setFormula(`=SUM(D2:D${rows.length + 1})`);
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(C${summaryRow}/D${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 6).setFormula(`=IFERROR(G${summaryRow}/C${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`);
    sheet.getRange(summaryRow, 8).setFormula(`=SUM(H2:H${rows.length + 1})`);
    sheet.getRange(summaryRow, 9).setFormula(`=IFERROR(H${summaryRow}/G${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 10).setFormula(`=SUM(J2:J${rows.length + 1})`);
    sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(G${summaryRow}/J${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 12).setFormula(`=IFERROR(J${summaryRow}/C${summaryRow}, 0)`);
    sheet.getRange(summaryRow, 13).setFormula(`=SUM(M2:M${rows.length + 1})`);
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
  }
  formatSheet(sheet, headers.length, rows.length + 2);
}

function exportTopSearchQueriesData(spreadsheet, summaryData) {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_SEARCH_QUERIES);
    const headers = ['Search Query', 'Campaign', 'Ad Group', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const query = `SELECT search_term_view.search_term, campaign.name, ad_group.name, metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.conversions_from_interactions_rate, metrics.all_conversions FROM search_term_view WHERE segments.date DURING ${CONFIG.DATE_RANGE} AND ad_group.status != 'REMOVED' AND metrics.cost_micros > 0 ORDER BY metrics.cost_micros DESC`;
    const report = AdsApp.report(query);
    const rows = [];
    const iterator = report.rows();
    while (iterator.hasNext()) {
        const row = iterator.next();
        const cost = (row['metrics.cost_micros'] || 0) / 1000000;
        const convValue = row['metrics.conversions_value'] || 0;
        const conversions = row['metrics.conversions'] || 0;
        const roas = cost > 0 ? convValue / cost : 0;
        const costPerConv = conversions > 0 ? cost / conversions : 0;
        rows.push([row['search_term_view.search_term'], row['campaign.name'], row['ad_group.name'], row['metrics.clicks'], row['metrics.impressions'], row['metrics.ctr'], row['metrics.clicks'] > 0 ? (row['metrics.average_cpc'] || 0) / 1000000 : 0, cost, convValue, roas, conversions, costPerConv, row['metrics.conversions_from_interactions_rate'], row['metrics.all_conversions']]);
    }

    summaryData.searchQueryCount = rows.length;

    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        const summaryRow = rows.length + 2;
        sheet.getRange(summaryRow, 1).setValue('Total');
        sheet.getRange(summaryRow, 4).setFormula(`=SUM(D2:D${rows.length + 1})`);
        sheet.getRange(summaryRow, 5).setFormula(`=SUM(E2:E${rows.length + 1})`);
        sheet.getRange(summaryRow, 6).setFormula(`=IFERROR(D${summaryRow}/E${summaryRow}, 0)`);
        sheet.getRange(summaryRow, 7).setFormula(`=IFERROR(H${summaryRow}/D${summaryRow}, 0)`);
        sheet.getRange(summaryRow, 8).setFormula(`=SUM(H2:H${rows.length + 1})`);
        sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`);
        sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(I${summaryRow}/H${summaryRow}, 0)`);
        sheet.getRange(summaryRow, 11).setFormula(`=SUM(K2:K${rows.length + 1})`);
        sheet.getRange(summaryRow, 12).setFormula(`=IFERROR(H${summaryRow}/K${summaryRow}, 0)`);
        sheet.getRange(summaryRow, 13).setFormula(`=IFERROR(K${summaryRow}/D${summaryRow}, 0)`);
        sheet.getRange(summaryRow, 14).setFormula(`=SUM(N2:N${rows.length + 1})`);
        sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
    }
    formatSheet(sheet, headers.length, rows.length + 2);
}

// ==================== CORRECTED LANDING PAGE FUNCTION ====================
function exportTopLandingPagesData(spreadsheet, summaryData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_LANDING_PAGES);
  const headers = ['Landing Page', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const rows = [];
  summaryData.landingPageCount = 0;

  try {
    const query = `SELECT landing_page_view.unexpanded_final_url, metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.conversions_from_interactions_rate, metrics.all_conversions FROM landing_page_view WHERE segments.date DURING ${CONFIG.DATE_RANGE} AND metrics.cost_micros > 0 ORDER BY metrics.cost_micros DESC`;
    const report = AdsApp.report(query);
    const iterator = report.rows();
    while (iterator.hasNext()) {
      const row = iterator.next();
      const cost = (row['metrics.cost_micros'] || 0) / 1000000;
      const convValue = row['metrics.conversions_value'] || 0;
      const conversions = row['metrics.conversions'] || 0;
      const roas = cost > 0 ? convValue / cost : 0;
      const costPerConv = conversions > 0 ? cost / conversions : 0;
      rows.push([row['landing_page_view.unexpanded_final_url'], row['metrics.clicks'], row['metrics.impressions'], row['metrics.ctr'], row['metrics.clicks'] > 0 ? (row['metrics.average_cpc'] || 0) / 1000000 : 0, cost, convValue, roas, conversions, costPerConv, row['metrics.conversions_from_interactions_rate'], row['metrics.all_conversions']]);
    }

    summaryData.landingPageCount = rows.length;

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      const summaryRow = rows.length + 2;
      sheet.getRange(summaryRow, 1).setValue('Total');
      sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`);
      sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`);
      sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`);
      sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`);
      sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`);
      sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(F${summaryRow}/I${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 12).setFormula(`=SUM(L2:L${rows.length + 1})`);
      sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
    } else {
        sheet.getRange(2, 1).setValue('No landing page data with spend found for the selected date range.');
    }
  } catch (e) {
      Logger.log(`Could not retrieve landing page data. This can happen if the account uses shared budgets or only has campaign types incompatible with this report. Error: ${e}`);
      sheet.getRange(2, 1).setValue(`Could not retrieve landing page data. See script logs for details.`);
  }
  
  formatSheet(sheet, headers.length, Math.max(2, rows.length + 2));
}

function exportTopCountriesData(spreadsheet, summaryData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_COUNTRIES);
  const headers = ['Country', 'Clicks', 'Impr.', 'CTR', 'Avg. CPC', 'Cost', 'Total Conv. Value', 'Conv. Value / Cost', 'Conv.', 'Cost / Conv.', 'Conv. Rate', 'All Conv.'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const rows = [];
  try {
    const countryNames = getCountryNames();
    const query = `SELECT geographic_view.country_criterion_id, metrics.clicks, metrics.impressions, metrics.cost_micros, metrics.conversions_value, metrics.conversions, metrics.all_conversions FROM geographic_view WHERE segments.date DURING ${CONFIG.DATE_RANGE} AND geographic_view.location_type = 'LOCATION_OF_PRESENCE' AND metrics.cost_micros > 0`;
    const report = AdsApp.report(query);
    const countryData = {};
    const iterator = report.rows();
    while (iterator.hasNext()) {
      const row = iterator.next();
      const countryCriterionId = row['geographic_view.country_criterion_id'];
      const countryName = countryNames[countryCriterionId] || `Unknown (${countryCriterionId})`;
      if (!countryData[countryName]) {
        countryData[countryName] = { clicks: 0, impressions: 0, cost: 0, convValue: 0, conversions: 0, allConversions: 0 };
      }
      countryData[countryName].clicks += parseInt(row['metrics.clicks']) || 0;
      countryData[countryName].impressions += parseInt(row['metrics.impressions']) || 0;
      countryData[countryName].cost += (parseFloat(row['metrics.cost_micros']) || 0) / 1000000;
      countryData[countryName].convValue += parseFloat(row['metrics.conversions_value']) || 0;
      countryData[countryName].conversions += parseFloat(row['metrics.conversions']) || 0;
      countryData[countryName].allConversions += parseFloat(row['metrics.all_conversions']) || 0;
    }
    
    for (const country in countryData) {
      const data = countryData[country];
      const ctr = data.impressions > 0 ? data.clicks / data.impressions : 0;
      const avgCpc = data.clicks > 0 ? data.cost / data.clicks : 0;
      const roas = data.cost > 0 ? data.convValue / data.cost : 0;
      const convRate = data.clicks > 0 ? data.conversions / data.clicks : 0;
      const costPerConv = data.conversions > 0 ? data.cost / data.conversions : 0;
      rows.push([country, data.clicks, data.impressions, ctr, avgCpc, data.cost, data.convValue, roas, data.conversions, costPerConv, convRate, data.allConversions]);
    }
    rows.sort((a, b) => b[5] - a[5]);

    summaryData.countryCount = rows.length;

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      const summaryRow = rows.length + 2;
      sheet.getRange(summaryRow, 1).setValue('Total');
      sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`);
      sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`);
      sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`);
      sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`);
      sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`);
      sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(F${summaryRow}/I${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`);
      sheet.getRange(summaryRow, 12).setFormula(`=SUM(L2:L${rows.length + 1})`);
      sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8e8e8');
    } else {
      sheet.getRange(2, 1).setValue('No country-level data available for the selected date range.');
    }
  } catch (e) {
    Logger.log('Error in exportTopCountriesData: ' + e.toString());
  }
  formatSheet(sheet, headers.length, Math.max(2, rows.length + 2));
}

// ==================== HELPER AND FORMATTING FUNCTIONS ====================
// Functions for getCountryNames, sendCompletionEmail, and formatSheet are unchanged.
// For brevity, they are omitted but are assumed to be present and correct.
function getCountryNames() {
  const query = `SELECT geo_target_constant.id, geo_target_constant.name FROM geo_target_constant WHERE geo_target_constant.target_type = 'Country'`;
  const countryNames = {};
  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    while (rows.hasNext()) {
      const row = rows.next();
      countryNames[row['geo_target_constant.id']] = row['geo_target_constant.name'];
    }
  } catch (e) {
    Logger.log('Error fetching country names: ' + e.toString());
  }
  return countryNames;
}

function sendCompletionEmail(spreadsheetUrl, summaryData) {
  if (!CONFIG.EMAIL_RECIPIENTS || CONFIG.EMAIL_RECIPIENTS.trim() === '' || CONFIG.EMAIL_RECIPIENTS === 'your-email@example.com') {
    Logger.log('Email notification is enabled but no valid recipient email address is configured.');
    return;
  }
  
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const timeZone = AdsApp.currentAccount().getTimeZone();
  const currentTime = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss z');
  const subject = `[${accountName}] Google Ads Performance Dashboard Updated`;

  const formatCurrency = (num) => '$' + num.toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const getChangeHtml = (current, prior, isCurrency = false) => {
      const change = current - prior;
      const percent = prior > 0 ? (change / prior) : (current > 0 ? 1 : 0);
      let color;

      if (change === 0) {
        color = '#333';
      } else if (isCurrency) {
        color = change > 0 ? '#c0392b' : '#27ae60'; 
      } else {
        color = change > 0 ? '#27ae60' : '#c0392b';
      }

      const arrow = change === 0 ? '' : (change > 0 ? '▲' : '▼');
      if (prior === 0 && current > 0) {
        return `<span style="color: #2980b9; font-weight: bold;">(New)</span>`;
      }
      const changeStr = isCurrency ? formatCurrency(Math.abs(change)) : Math.abs(change).toLocaleString();
      return `<span style="color: ${color}; font-weight: bold;">${arrow} ${changeStr} (${(percent * 100).toFixed(1)}%)</span>`;
  };

  const costSummaryHtml = `
      <tr><td style="padding: 8px 4px; border-bottom: 1px solid #eee;">Total Cost</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${formatCurrency(summaryData.totalCost || 0)}</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${formatCurrency(summaryData.priorTotalCost || 0)}</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${getChangeHtml(summaryData.totalCost || 0, summaryData.priorTotalCost || 0, true)}</td>
      </tr>`;
      
  const buildEntityRowHtml = (label, current, prior) => `
      <tr><td style="padding: 8px 4px; border-bottom: 1px solid #eee;">${label}</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${(current || 0).toLocaleString()}</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${(prior || 0).toLocaleString()}</td>
          <td style="text-align: right; padding: 8px 4px; border-bottom: 1px solid #eee;">${getChangeHtml(current || 0, prior || 0)}</td>
      </tr>`;
      
  const summaryHtml = `
    <h3 style="color: #333; border-bottom: 1px solid #ccc; padding-bottom: 5px;">Performance Summary</h3>
    <p style="font-size: 12px; color: #666;">${CONFIG.DATE_RANGE.replace(/_/g, ' ')} vs. Prior Period</p>
    <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 10px;">
      <tr style="font-weight: bold; color: #555;">
          <td style="text-align: left; padding: 8px 4px;">Metric</td>
          <td style="text-align: right; padding: 8px 4px;">Current</td>
          <td style="text-align: right; padding: 8px 4px;">Prior</td>
          <td style="text-align: right; padding: 8px 4px;">Change</td>
      </tr>
      ${costSummaryHtml}
      ${buildEntityRowHtml('Active Campaigns', summaryData.campaignCount, summaryData.priorCampaignCount)}
      ${buildEntityRowHtml('Active Ad Groups', summaryData.adGroupCount, summaryData.priorAdGroupCount)}
      ${buildEntityRowHtml('Active Search Queries', summaryData.searchQueryCount, summaryData.priorSearchQueryCount)}
      ${buildEntityRowHtml('Active Landing Pages', summaryData.landingPageCount, summaryData.priorLandingPageCount)}
      ${buildEntityRowHtml('Active Countries', summaryData.countryCount, summaryData.priorCountryCount)}
    </table>`;
  
  const getChangePlainText = (current, prior, isCurrency = false) => {
      if (prior === 0 && current > 0) return '(New)';
      const change = current - prior;
      if (change === 0) return '0 (0.0%)';
      const percent = prior > 0 ? (change / prior) : 1;
      const sign = change >= 0 ? '+' : '-';
      const changeStr = isCurrency ? formatCurrency(Math.abs(change)) : Math.abs(change).toLocaleString();
      return `${sign}${changeStr} (${(percent * 100).toFixed(1)}%)`;
  };
  const summaryPlainText = `
Performance Summary (${CONFIG.DATE_RANGE.replace(/_/g, ' ')} vs. Prior Period)
----------------------------------------------------------------
Metric                  Current        Prior          Change
----------------------------------------------------------------
Total Cost              ${formatCurrency(summaryData.totalCost || 0).padEnd(14)} ${formatCurrency(summaryData.priorTotalCost || 0).padEnd(14)} ${getChangePlainText(summaryData.totalCost || 0, summaryData.priorTotalCost || 0, true)}
Active Campaigns        ${((summaryData.campaignCount || 0).toLocaleString()).padEnd(14)} ${((summaryData.priorCampaignCount || 0).toLocaleString()).padEnd(14)} ${getChangePlainText(summaryData.campaignCount || 0, summaryData.priorCampaignCount || 0)}
Active Ad Groups        ${((summaryData.adGroupCount || 0).toLocaleString()).padEnd(14)} ${((summaryData.priorAdGroupCount || 0).toLocaleString()).padEnd(14)} ${getChangePlainText(summaryData.adGroupCount || 0, summaryData.priorAdGroupCount || 0)}
Active Search Queries   ${((summaryData.searchQueryCount || 0).toLocaleString()).padEnd(14)} ${((summaryData.priorSearchQueryCount || 0).toLocaleString()).padEnd(14)} ${getChangePlainText(summaryData.searchQueryCount || 0, summaryData.priorSearchQueryCount || 0)}
Active Landing Pages    ${((summaryData.landingPageCount || 0).toLocaleString()).padEnd(14)} ${((summaryData.priorLandingPageCount || 0).toLocaleString()).padEnd(14)} ${getChangePlainText(summaryData.landingPageCount || 0, summaryData.priorLandingPageCount || 0)}
Active Countries        ${((summaryData.countryCount || 0).toLocaleString()).padEnd(14)} ${((summaryData.priorCountryCount || 0).toLocaleString()).padEnd(14)} ${getChangePlainText(summaryData.countryCount || 0, summaryData.priorCountryCount || 0)}
  `;

  const htmlBody = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #333;"> <h2 style="color: #4285f4;">Google Ads Performance Dashboard Updated</h2> <p>Your Google Ads Performance Dashboard for <strong>${accountName} (${accountId})</strong> has been updated successfully.</p> ${summaryHtml} <p style="margin: 25px 0; text-align: center;"> <a href="${spreadsheetUrl}" style="background-color: #4285f4; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; display: inline-block; font-size: 16px;">View Full Dashboard</a> </p> <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;"> <p style="color: #666; font-size: 12px;">This report was generated automatically by Google Ads Scripts on ${currentTime}.</p> </div>`;
  const plainTextBody = `Your Google Ads Performance Dashboard has been updated successfully.\n\nAccount: ${accountName} (${accountId})\n${summaryPlainText}\n\nView the full dashboard: ${spreadsheetUrl}\n\nThis report was generated automatically by Google Ads Scripts on ${currentTime}.`;
  
  try {
    MailApp.sendEmail({ to: CONFIG.EMAIL_RECIPIENTS, subject: subject, body: plainTextBody, htmlBody: htmlBody });
    Logger.log(`Completion email sent to ${CONFIG.EMAIL_RECIPIENTS}`);
  } catch (e) {
    Logger.log(`Failed to send email: ${e.toString()}`);
  }
}

function formatSheet(sheet, numColumns, numRows) {
  const sheetName = sheet.getName();
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment(sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY ? 'right' : 'center');
  headerRange.setBackground('#f3f3f3');
  headerRange.setWrap(true);
  
  if (sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY) {
    sheet.setColumnWidth(1, 100);
    for (let i = 2; i <= numColumns; i++) sheet.setColumnWidth(i, 65);
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_CAMPAIGNS || sheetName === CONFIG.SHEET_NAMES.TOP_COUNTRIES || sheetName === CONFIG.SHEET_NAMES.TOP_LANDING_PAGES) {
    for (let i = 2; i <= numColumns; i++) sheet.setColumnWidth(i, 65);
    sheet.autoResizeColumn(1);
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_AD_GROUPS) {
    for (let i = 3; i <= numColumns; i++) sheet.setColumnWidth(i, 65);
    sheet.autoResizeColumn(1); sheet.autoResizeColumn(2);
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_SEARCH_QUERIES) {
    for (let i = 4; i <= numColumns; i++) sheet.setColumnWidth(i, 65);
    sheet.autoResizeColumn(1); sheet.autoResizeColumn(2); sheet.autoResizeColumn(3);
  }
  
  if (numRows > 1) {
    const dataRows = numRows - 1;
    const formatNoDecimals = '$#,##0';
    const formatOneDecimal = '#,##0.0';

    if (sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY || sheetName === CONFIG.SHEET_NAMES.TOP_CAMPAIGNS || sheetName === CONFIG.SHEET_NAMES.TOP_COUNTRIES || sheetName === CONFIG.SHEET_NAMES.TOP_LANDING_PAGES) {
      sheet.getRange(2, 2, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('0.00');
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat(formatOneDecimal);
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 12, dataRows, 1).setNumberFormat(formatOneDecimal);
    } else if (sheetName === CONFIG.SHEET_NAMES.TOP_AD_GROUPS) {
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat('0.00');
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat(formatOneDecimal);
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 12, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 13, dataRows, 1).setNumberFormat(formatOneDecimal);
    } else if (sheetName === CONFIG.SHEET_NAMES.TOP_SEARCH_QUERIES) {
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat('0.00');
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat(formatOneDecimal);
      sheet.getRange(2, 12, dataRows, 1).setNumberFormat(formatNoDecimals);
      sheet.getRange(2, 13, dataRows, 1).setNumberFormat('0.00%');
      sheet.getRange(2, 14, dataRows, 1).setNumberFormat(formatOneDecimal);
    }
  }
  sheet.setFrozenRows(1);
}