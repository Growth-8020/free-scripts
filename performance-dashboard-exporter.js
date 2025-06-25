/**
 * Google Ads Script - Export Performance Metrics to Google Sheets
 * This script exports account, campaign, and ad group performance data to a Google Sheet
 * 
 * Features:
 * - Account Daily: Daily performance metrics for the entire account
 * - Top Campaigns: All campaigns sorted by cost (only campaigns with cost > $0)
 * - Top Ad Groups: All ad groups sorted by cost where impressions > 0
 * - Top Countries: Country-level performance based on user location (shows where users were physically located)
 * - Configurable date ranges
 * - Summary rows with totals and calculated metrics
 * - Auto-fitted column widths for text columns
 * - Optional email notifications when export completes
 * 
 * Setup Instructions:
 * 1. Replace 'YOUR_SPREADSHEET_URL_HERE' with your Google Sheet URL
 * 2. Adjust the DATE_RANGE if needed (default is 'LAST_30_DAYS')
 * 3. (Optional) Set SEND_EMAIL_ON_COMPLETE to true and update EMAIL_RECIPIENTS
 * 4. Schedule the script to run as needed (e.g., daily, weekly)
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
    TOP_COUNTRIES: 'Top Countries'
  },
  
  // Notification settings
  SEND_EMAIL_ON_COMPLETE: false,
  EMAIL_RECIPIENTS: 'your-email@example.com' // Comma-separated for multiple recipients
};

// ==================== MAIN FUNCTION ====================
function main() {
  try {
    // Open or create the spreadsheet
    const spreadsheet = getOrCreateSpreadsheet();
    
    // Clear existing data
    clearSheets(spreadsheet);
    
    // Export data to each sheet
    exportAccountDailyData(spreadsheet);
    exportTopCampaignsData(spreadsheet);
    exportTopAdGroupsData(spreadsheet);
    exportTopCountriesData(spreadsheet);
    
    // Send email notification if enabled
    if (CONFIG.SEND_EMAIL_ON_COMPLETE) {
      sendCompletionEmail(spreadsheet.getUrl());
    }
    
    Logger.log('Export completed successfully!');
    
  } catch (error) {
    Logger.log('Error in main function: ' + error.toString());
    throw error;
  }
}

// ==================== SPREADSHEET FUNCTIONS ====================
function getOrCreateSpreadsheet() {
  let spreadsheet;
  
  try {
    spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  } catch (e) {
    throw new Error('Unable to open spreadsheet. Please check the URL in CONFIG.SPREADSHEET_URL');
  }
  
  // Create sheets if they don't exist
  ensureSheetExists(spreadsheet, CONFIG.SHEET_NAMES.ACCOUNT_DAILY);
  ensureSheetExists(spreadsheet, CONFIG.SHEET_NAMES.TOP_CAMPAIGNS);
  ensureSheetExists(spreadsheet, CONFIG.SHEET_NAMES.TOP_AD_GROUPS);
  ensureSheetExists(spreadsheet, CONFIG.SHEET_NAMES.TOP_COUNTRIES);
  
  return spreadsheet;
}

function ensureSheetExists(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function clearSheets(spreadsheet) {
  const sheets = [
    CONFIG.SHEET_NAMES.ACCOUNT_DAILY,
    CONFIG.SHEET_NAMES.TOP_CAMPAIGNS,
    CONFIG.SHEET_NAMES.TOP_AD_GROUPS,
    CONFIG.SHEET_NAMES.TOP_COUNTRIES
  ];
  
  sheets.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    }
  });
}

// ==================== DATA EXPORT FUNCTIONS ====================
function exportAccountDailyData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.ACCOUNT_DAILY);
  
  // Set headers
  const headers = [
    'Date',
    'Clicks',
    'Impr.',
    'CTR',
    'Avg. CPC',
    'Cost',
    'Total Conv. Value',
    'Conv. Value / Cost',
    'Conv.',
    'Conv. Rate',
    'All Conv.'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get account performance data
  const query = `
    SELECT 
      segments.date,
      metrics.clicks,
      metrics.impressions,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.conversions_value,
      metrics.conversions,
      metrics.conversions_from_interactions_rate,
      metrics.all_conversions
    FROM customer
    WHERE segments.date DURING ${CONFIG.DATE_RANGE}
    ORDER BY segments.date DESC
  `;
  
  const report = AdsApp.report(query);
  const rows = [];
  
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = row['metrics.cost_micros'] / 1000000;
    const convValue = row['metrics.conversions_value'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    
    rows.push([
      row['segments.date'],
      row['metrics.clicks'],
      row['metrics.impressions'],
      row['metrics.ctr'], // CTR as decimal (0.05 = 5%)
      row['metrics.clicks'] > 0 ? row['metrics.average_cpc'] / 1000000 : 0, // Avg CPC in dollars (0 if no clicks)
      cost, // Cost in dollars
      convValue, // Conv value in dollars
      roas, // ROAS as number
      row['metrics.conversions'],
      row['metrics.conversions_from_interactions_rate'], // Conv rate as decimal (0.05 = 5%)
      row['metrics.all_conversions']
    ]);
  }
  
  // Write data to sheet
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Add summary row
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    
    // Add formulas for summary row
    sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`); // Clicks
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`); // Impressions
    sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`); // CTR
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`); // Avg CPC
    sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`); // Cost
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`); // Total Conv Value
    sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`); // Conv Value / Cost
    sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`); // Conversions
    sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`); // Conv Rate
    sheet.getRange(summaryRow, 11).setFormula(`=SUM(K2:K${rows.length + 1})`); // All Conv
    
    // Bold the summary row
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(summaryRow, 1, 1, headers.length).setBackground('#e8e8e8');
  }
  
  // Format the sheet
  formatSheet(sheet, headers.length, rows.length + 2);
}

function exportTopCampaignsData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_CAMPAIGNS);
  
  // Set headers
  const headers = [
    'Campaign',
    'Clicks',
    'Impr.',
    'CTR',
    'Avg. CPC',
    'Cost',
    'Total Conv. Value',
    'Conv. Value / Cost',
    'Conv.',
    'Conv. Rate',
    'All Conv.'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get all campaign performance data (only campaigns with cost > $0)
  const query = `
    SELECT 
      campaign.name,
      metrics.clicks,
      metrics.impressions,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.conversions_value,
      metrics.conversions,
      metrics.conversions_from_interactions_rate,
      metrics.all_conversions
    FROM campaign
    WHERE segments.date DURING ${CONFIG.DATE_RANGE}
      AND campaign.status != 'REMOVED'
      AND metrics.cost_micros > 0
    ORDER BY metrics.cost_micros DESC
  `;
  
  const report = AdsApp.report(query);
  const rows = [];
  
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = row['metrics.cost_micros'] / 1000000;
    const convValue = row['metrics.conversions_value'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    
    rows.push([
      row['campaign.name'],
      row['metrics.clicks'],
      row['metrics.impressions'],
      row['metrics.ctr'], // CTR as decimal (0.05 = 5%)
      row['metrics.clicks'] > 0 ? row['metrics.average_cpc'] / 1000000 : 0, // Avg CPC in dollars (0 if no clicks)
      cost, // Cost in dollars
      convValue, // Conv value in dollars
      roas, // ROAS as number
      row['metrics.conversions'],
      row['metrics.conversions_from_interactions_rate'], // Conv rate as decimal (0.05 = 5%)
      row['metrics.all_conversions']
    ]);
  }
  
  // Write data to sheet
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Add summary row
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    
    // Add formulas for summary row
    sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`); // Clicks
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`); // Impressions
    sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`); // CTR
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`); // Avg CPC
    sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`); // Cost
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`); // Total Conv Value
    sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`); // Conv Value / Cost
    sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`); // Conversions
    sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`); // Conv Rate
    sheet.getRange(summaryRow, 11).setFormula(`=SUM(K2:K${rows.length + 1})`); // All Conv
    
    // Bold the summary row
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(summaryRow, 1, 1, headers.length).setBackground('#e8e8e8');
  }
  
  // Format the sheet
  formatSheet(sheet, headers.length, rows.length + 2);
}

function exportTopAdGroupsData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_AD_GROUPS);
  
  // Set headers
  const headers = [
    'Campaign',
    'Ad Group',
    'Clicks',
    'Impr.',
    'CTR',
    'Avg. CPC',
    'Cost',
    'Total Conv. Value',
    'Conv. Value / Cost',
    'Conv.',
    'Conv. Rate',
    'All Conv.'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get all ad group performance data (no limit, but only where impressions > 0)
  const query = `
    SELECT 
      campaign.name,
      ad_group.name,
      metrics.clicks,
      metrics.impressions,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.conversions_value,
      metrics.conversions,
      metrics.conversions_from_interactions_rate,
      metrics.all_conversions
    FROM ad_group
    WHERE segments.date DURING ${CONFIG.DATE_RANGE}
      AND ad_group.status != 'REMOVED'
      AND metrics.impressions > 0
    ORDER BY metrics.cost_micros DESC
  `;
  
  const report = AdsApp.report(query);
  const rows = [];
  
  const iterator = report.rows();
  while (iterator.hasNext()) {
    const row = iterator.next();
    const cost = row['metrics.cost_micros'] / 1000000;
    const convValue = row['metrics.conversions_value'] || 0;
    const roas = cost > 0 ? convValue / cost : 0;
    
    rows.push([
      row['campaign.name'],
      row['ad_group.name'],
      row['metrics.clicks'],
      row['metrics.impressions'],
      row['metrics.ctr'], // CTR as decimal (0.05 = 5%)
      row['metrics.clicks'] > 0 ? row['metrics.average_cpc'] / 1000000 : 0, // Avg CPC in dollars (0 if no clicks)
      cost, // Cost in dollars
      convValue, // Conv value in dollars
      roas, // ROAS as number
      row['metrics.conversions'],
      row['metrics.conversions_from_interactions_rate'], // Conv rate as decimal (0.05 = 5%)
      row['metrics.all_conversions']
    ]);
  }
  
  // Write data to sheet
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Add summary row
    const summaryRow = rows.length + 2;
    sheet.getRange(summaryRow, 1).setValue('Total');
    
    // Add formulas for summary row (columns shifted by 1 due to Campaign column)
    sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`); // Clicks
    sheet.getRange(summaryRow, 4).setFormula(`=SUM(D2:D${rows.length + 1})`); // Impressions
    sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(C${summaryRow}/D${summaryRow}, 0)`); // CTR
    sheet.getRange(summaryRow, 6).setFormula(`=IFERROR(G${summaryRow}/C${summaryRow}, 0)`); // Avg CPC
    sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`); // Cost
    sheet.getRange(summaryRow, 8).setFormula(`=SUM(H2:H${rows.length + 1})`); // Total Conv Value
    sheet.getRange(summaryRow, 9).setFormula(`=IFERROR(H${summaryRow}/G${summaryRow}, 0)`); // Conv Value / Cost
    sheet.getRange(summaryRow, 10).setFormula(`=SUM(J2:J${rows.length + 1})`); // Conversions
    sheet.getRange(summaryRow, 11).setFormula(`=IFERROR(J${summaryRow}/C${summaryRow}, 0)`); // Conv Rate
    sheet.getRange(summaryRow, 12).setFormula(`=SUM(L2:L${rows.length + 1})`); // All Conv
    
    // Bold the summary row
    sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(summaryRow, 1, 1, headers.length).setBackground('#e8e8e8');
  }
  
  // Format the sheet
  formatSheet(sheet, headers.length, rows.length + 2);
}

// ==================== HELPER FUNCTIONS ====================
// Function to get human-readable date range string
function getDateRangeString() {
  const dateRange = CONFIG.DATE_RANGE;
  
  // For predefined date ranges, return as-is
  if (dateRange.includes('_')) {
    return dateRange.replace(/_/g, ' ').toLowerCase()
      .replace(/\b\w/g, char => char.toUpperCase());
  }
  
  // For custom date ranges, you might have specific dates
  return dateRange;
}

// Function to send completion email
function sendCompletionEmail(spreadsheetUrl) {
  // Validate email configuration
  if (!CONFIG.EMAIL_RECIPIENTS || CONFIG.EMAIL_RECIPIENTS === 'your-email@example.com') {
    Logger.log('Email notification is enabled but no valid recipient email address is configured.');
    return;
  }
  
  const subject = 'Google Ads Performance Dashboard Updated';
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const timeZone = AdsApp.currentAccount().getTimeZone();
  const currentTime = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss z');
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #4285f4;">Google Ads Performance Dashboard Updated</h2>
      
      <p>Your Google Ads Performance Dashboard has been updated successfully.</p>
      
      <table style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; width: 100%;">
        <tr>
          <td><strong>Account:</strong></td>
          <td>${accountName} (${accountId})</td>
        </tr>
        <tr>
          <td><strong>Date Range:</strong></td>
          <td>${getDateRangeString()}</td>
        </tr>
        <tr>
          <td><strong>Updated:</strong></td>
          <td>${currentTime}</td>
        </tr>
      </table>
      
      <p style="margin: 20px 0;">
        <a href="${spreadsheetUrl}" style="background-color: #4285f4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">
          View Dashboard
        </a>
      </p>
      
      <p><strong>This automated report includes:</strong></p>
      <ul>
        <li>Daily account performance metrics</li>
        <li>Top campaigns by cost</li>
        <li>Top ad groups with impressions</li>
        <li>Country-level performance data</li>
      </ul>
      
      <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
      
      <p style="color: #666; font-size: 12px;">
        This report was generated automatically by Google Ads Scripts.
      </p>
    </div>
  `;
  
  const plainTextBody = `Your Google Ads Performance Dashboard has been updated successfully.

Account: ${accountName} (${accountId})
View the dashboard: ${spreadsheetUrl}

Date Range: ${getDateRangeString()}
Updated: ${currentTime}

This automated report includes:
- Daily account performance metrics
- Top campaigns by cost
- Top ad groups with impressions
- Country-level performance data

This report was generated automatically by Google Ads Scripts.`;
  
  try {
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENTS,
      subject: subject,
      body: plainTextBody,
      htmlBody: htmlBody
    });
    Logger.log(`Completion email sent to ${CONFIG.EMAIL_RECIPIENTS}`);
  } catch (e) {
    Logger.log(`Failed to send email: ${e.toString()}`);
  }
}

// Helper function to get country names from criterion IDs
function getCountryNames() {
  const query = `
    SELECT 
      geo_target_constant.id,
      geo_target_constant.name,
      geo_target_constant.country_code
    FROM geo_target_constant
    WHERE geo_target_constant.target_type = 'Country'
  `;
  
  const countryNames = {};
  
  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    while (rows.hasNext()) {
      const row = rows.next();
      const id = row['geo_target_constant.id'];
      const name = row['geo_target_constant.name'];
      countryNames[id] = name;
    }
  } catch (e) {
    Logger.log('Error fetching country names: ' + e.toString());
    // Continue with criterion IDs if names can't be fetched
  }
  
  return countryNames;
}

function exportTopCountriesData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.TOP_COUNTRIES);
  
  // Set headers
  const headers = [
    'Country',
    'Clicks',
    'Impr.',
    'CTR',
    'Avg. CPC',
    'Cost',
    'Total Conv. Value',
    'Conv. Value / Cost',
    'Conv.',
    'Conv. Rate',
    'All Conv.'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const rows = [];
  
  try {
    // First, get country names mapping
    const countryNames = getCountryNames();
    
    // Get country performance data using geographic_view
    const query = `
      SELECT 
        geographic_view.country_criterion_id,
        metrics.clicks,
        metrics.impressions,
        metrics.ctr,
        metrics.average_cpc,
        metrics.cost_micros,
        metrics.conversions_value,
        metrics.conversions,
        metrics.conversions_from_interactions_rate,
        metrics.all_conversions
      FROM geographic_view
      WHERE segments.date DURING ${CONFIG.DATE_RANGE}
        AND geographic_view.location_type = 'LOCATION_OF_PRESENCE'
    `;
    
    const report = AdsApp.report(query);
    const countryData = {};
    
    const iterator = report.rows();
    while (iterator.hasNext()) {
      const row = iterator.next();
      const countryCriterionId = row['geographic_view.country_criterion_id'];
      const countryName = countryNames[countryCriterionId] || `Unknown (${countryCriterionId})`;
      
      // Aggregate data by country
      if (!countryData[countryName]) {
        countryData[countryName] = {
          clicks: 0,
          impressions: 0,
          cost: 0,
          convValue: 0,
          conversions: 0,
          allConversions: 0
        };
      }
      
      countryData[countryName].clicks += parseInt(row['metrics.clicks']) || 0;
      countryData[countryName].impressions += parseInt(row['metrics.impressions']) || 0;
      countryData[countryName].cost += (parseFloat(row['metrics.cost_micros']) || 0) / 1000000;
      countryData[countryName].convValue += parseFloat(row['metrics.conversions_value']) || 0;
      countryData[countryName].conversions += parseFloat(row['metrics.conversions']) || 0;
      countryData[countryName].allConversions += parseFloat(row['metrics.all_conversions']) || 0;
    }
    
    // Convert to rows
    for (const country in countryData) {
      const data = countryData[country];
      if (data.impressions > 0) {
        const ctr = data.clicks / data.impressions;
        const avgCpc = data.clicks > 0 ? data.cost / data.clicks : 0;
        const roas = data.cost > 0 ? data.convValue / data.cost : 0;
        const convRate = data.clicks > 0 ? data.conversions / data.clicks : 0;
        
        rows.push([
          country,
          data.clicks,
          data.impressions,
          ctr,
          avgCpc,
          data.cost,
          data.convValue,
          roas,
          data.conversions,
          convRate,
          data.allConversions
        ]);
      }
    }
    
    // Sort by cost descending
    rows.sort((a, b) => b[5] - a[5]);
    
    // Write data to sheet
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      
      // Add summary row
      const summaryRow = rows.length + 2;
      sheet.getRange(summaryRow, 1).setValue('Total');
      
      // Add formulas for summary row
      sheet.getRange(summaryRow, 2).setFormula(`=SUM(B2:B${rows.length + 1})`); // Clicks
      sheet.getRange(summaryRow, 3).setFormula(`=SUM(C2:C${rows.length + 1})`); // Impressions
      sheet.getRange(summaryRow, 4).setFormula(`=IFERROR(B${summaryRow}/C${summaryRow}, 0)`); // CTR
      sheet.getRange(summaryRow, 5).setFormula(`=IFERROR(F${summaryRow}/B${summaryRow}, 0)`); // Avg CPC
      sheet.getRange(summaryRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`); // Cost
      sheet.getRange(summaryRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`); // Total Conv Value
      sheet.getRange(summaryRow, 8).setFormula(`=IFERROR(G${summaryRow}/F${summaryRow}, 0)`); // Conv Value / Cost
      sheet.getRange(summaryRow, 9).setFormula(`=SUM(I2:I${rows.length + 1})`); // Conversions
      sheet.getRange(summaryRow, 10).setFormula(`=IFERROR(I${summaryRow}/B${summaryRow}, 0)`); // Conv Rate
      sheet.getRange(summaryRow, 11).setFormula(`=SUM(K2:K${rows.length + 1})`); // All Conv
      
      // Bold the summary row
      sheet.getRange(summaryRow, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(summaryRow, 1, 1, headers.length).setBackground('#e8e8e8');
    } else {
      // Add note if no country data available
      sheet.getRange(2, 1).setValue('No country-level data available for the selected date range.');
    }
    
  } catch (e) {
    Logger.log('Error in exportTopCountriesData: ' + e.toString());
    sheet.getRange(2, 1).setValue('Error retrieving country data: ' + e.toString());
  }
  
  // Format the sheet
  formatSheet(sheet, headers.length, Math.max(2, rows.length + 2));
}

// ==================== FORMATTING FUNCTIONS ====================
function formatSheet(sheet, numColumns, numRows) {
  const sheetName = sheet.getName();
  
  // Bold and format the header row
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment(sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY ? 'right' : 'center');
  headerRange.setBackground('#f3f3f3');
  headerRange.setWrap(true); // Apply word wrap to header row
  
  // Set column widths based on sheet type
  // Note: Auto-resize happens after data is written for accurate sizing
  if (sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY) {
    // Account Daily: 100px for date, 60px for the rest
    sheet.setColumnWidth(1, 100); // Date column
    for (let i = 2; i <= numColumns; i++) {
      sheet.setColumnWidth(i, 60);
    }
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_CAMPAIGNS) {
    // Top Campaigns: auto-fit for first column, 60px for the rest
    for (let i = 2; i <= numColumns; i++) {
      sheet.setColumnWidth(i, 60);
    }
    sheet.autoResizeColumn(1); // Campaign column - auto-fit (do this last)
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_AD_GROUPS) {
    // Top Ad Groups: auto-fit for first two columns, 60px for the rest
    for (let i = 3; i <= numColumns; i++) {
      sheet.setColumnWidth(i, 60);
    }
    sheet.autoResizeColumn(1); // Campaign column - auto-fit (do this last)
    sheet.autoResizeColumn(2); // Ad Group column - auto-fit (do this last)
  } else if (sheetName === CONFIG.SHEET_NAMES.TOP_COUNTRIES) {
    // Top Countries: auto-fit for first column, 60px for the rest
    for (let i = 2; i <= numColumns; i++) {
      sheet.setColumnWidth(i, 60);
    }
    sheet.autoResizeColumn(1); // Country column - auto-fit (do this last)
  }
  
  // Apply number formatting based on sheet type
  if (numRows > 1) {
    // Adjust range to include summary row
    const dataRows = numRows - 1; // Total rows including summary
    
    if (sheetName === CONFIG.SHEET_NAMES.ACCOUNT_DAILY) {
      // Format Account Daily sheet
      sheet.getRange(2, 2, dataRows, 1).setNumberFormat('#,##0'); // Clicks
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0'); // Impressions
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0.00%'); // CTR
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('$#,##0.00'); // Avg. CPC
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('$#,##0'); // Cost
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0'); // Total Conv. Value
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('#,##0.00'); // Conv. Value / Cost
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat('#,##0.0'); // Conversions
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0.00%'); // Conv. Rate
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat('#,##0.0'); // All Conv.
    } else if (sheetName === CONFIG.SHEET_NAMES.TOP_CAMPAIGNS) {
      // Format Top Campaigns sheet
      sheet.getRange(2, 2, dataRows, 1).setNumberFormat('#,##0'); // Clicks
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0'); // Impressions
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0.00%'); // CTR
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('$#,##0.00'); // Avg. CPC
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('$#,##0'); // Cost
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0'); // Total Conv. Value
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('#,##0.00'); // Conv. Value / Cost
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat('#,##0.0'); // Conversions
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0.00%'); // Conv. Rate
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat('#,##0.0'); // All Conv.
    } else if (sheetName === CONFIG.SHEET_NAMES.TOP_AD_GROUPS) {
      // Format Top Ad Groups sheet (columns shifted by 1 due to Campaign column)
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0'); // Clicks
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0'); // Impressions
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('#,##0.00%'); // CTR
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('$#,##0.00'); // Avg. CPC
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0'); // Cost
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('$#,##0'); // Total Conv. Value
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat('#,##0.00'); // Conv. Value / Cost
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0.0'); // Conversions
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat('#,##0.00%'); // Conv. Rate
      sheet.getRange(2, 12, dataRows, 1).setNumberFormat('#,##0.0'); // All Conv.
    } else if (sheetName === CONFIG.SHEET_NAMES.TOP_COUNTRIES) {
      // Format Top Countries sheet
      sheet.getRange(2, 2, dataRows, 1).setNumberFormat('#,##0'); // Clicks
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0'); // Impressions
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0.00%'); // CTR
      sheet.getRange(2, 5, dataRows, 1).setNumberFormat('$#,##0.00'); // Avg. CPC
      sheet.getRange(2, 6, dataRows, 1).setNumberFormat('$#,##0'); // Cost
      sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0'); // Total Conv. Value
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('#,##0.00'); // Conv. Value / Cost
      sheet.getRange(2, 9, dataRows, 1).setNumberFormat('#,##0.0'); // Conversions
      sheet.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0.00%'); // Conv. Rate
      sheet.getRange(2, 11, dataRows, 1).setNumberFormat('#,##0.0'); // All Conv.
    }
  }
  
  
  // Freeze the header row
  sheet.setFrozenRows(1);
}