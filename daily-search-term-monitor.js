/**
 * Google Ads Script to Identify New Search Queries
 * 
 * This script fetches search terms from yesterday and compares them to those from the previous 180 days.
 * It identifies new queries (those not seen in the prior 180 days), includes campaign and ad group details,
 * sorts by clicks descending, outputs to a Google Sheet with additional columns, and emails results grouped by campaign.
 * Email results are limited to search terms with > 0 clicks and formatted in HTML tables for better readability.
 * 
 * Instructions:
 * 1. Optionally set SHEET_URL to your existing Google Sheet URL. If not set or placeholder, a new spreadsheet will be created.
 * 2. Replace RECIPIENT_EMAIL with the email address to send results to.
 * 3. Schedule the script to run daily in Google Ads.
 */

// Configuration
var SHEET_URL = 'YOUR_GOOGLE_SHEET_URL_HERE'; // Optional: Replace with your Sheet URL or leave to create new
var RECIPIENT_EMAIL = 'your-email@example.com'; // Replace with your email

// Main function
function main() {
  var timeZone = AdsApp.currentAccount().getTimeZone();
  
  // Calculate dates
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  var yesterdayStr = Utilities.formatDate(yesterday, timeZone, 'yyyy-MM-dd');
  
  var dayBeforeYesterday = new Date(yesterday);
  dayBeforeYesterday.setDate(dayBeforeYesterday.getDate() - 1);
  var dayBeforeYesterdayStr = Utilities.formatDate(dayBeforeYesterday, timeZone, 'yyyy-MM-dd');
  
  var startPrior = new Date(yesterday);
  startPrior.setDate(startPrior.getDate() - 180);
  var startPriorStr = Utilities.formatDate(startPrior, timeZone, 'yyyy-MM-dd');
  
  // Fetch prior search terms (set for uniqueness)
  var priorQueries = new Set();
  var priorQuery = 
    'SELECT search_term_view.search_term ' +
    'FROM search_term_view ' +
    'WHERE segments.date BETWEEN "' + startPriorStr + '" AND "' + dayBeforeYesterdayStr + '"';
  var priorReport = AdsApp.report(priorQuery);
  var priorRows = priorReport.rows();
  while (priorRows.hasNext()) {
    var row = priorRows.next();
    var query = row['search_term_view.search_term'].toLowerCase(); // Case-insensitive comparison
    priorQueries.add(query);
  }
  
  // Fetch yesterday's search terms with metrics and structure
  var yesterdayData = [];
  var yesterdayQuery = 
    'SELECT search_term_view.search_term, metrics.impressions, metrics.clicks, metrics.cost_micros, ' +
    'campaign.id, campaign.name, ad_group.id, ad_group.name ' +
    'FROM search_term_view ' +
    'WHERE segments.date = "' + yesterdayStr + '"';
  var yesterdayReport = AdsApp.report(yesterdayQuery);
  var yesterdayRows = yesterdayReport.rows();
  while (yesterdayRows.hasNext()) {
    var row = yesterdayRows.next();
    var origQuery = row['search_term_view.search_term'];
    var query = origQuery.toLowerCase();
    var impressions = parseInt(row['metrics.impressions']) || 0;
    var clicks = parseInt(row['metrics.clicks']) || 0;
    var cost = (parseInt(row['metrics.cost_micros']) || 0) / 1000000;
    var campaignId = row['campaign.id'];
    var campaignName = row['campaign.name'];
    var adGroupId = row['ad_group.id'];
    var adGroupName = row['ad_group.name'];
    yesterdayData.push({
      query: query,
      origQuery: origQuery,
      impressions: impressions,
      clicks: clicks,
      cost: cost.toFixed(2),
      campaignId: campaignId,
      campaignName: campaignName,
      adGroupId: adGroupId,
      adGroupName: adGroupName
    });
  }
  
  // Identify new queries
  var newData = yesterdayData.filter(function(item) {
    return !priorQueries.has(item.query);
  });
  
  // If no new queries, log and exit
  if (newData.length === 0) {
    Logger.log('No new search queries found.');
    return;
  }
  
  // Sort new data by clicks descending
  newData.sort(function(a, b) {
    return b.clicks - a.clicks;
  });
  
  // Get account details
  var accountName = AdsApp.currentAccount().getName();
  var accountId = AdsApp.currentAccount().getCustomerId();
  var last4Id = accountId.slice(-4);
  
  // Dynamic sheet name: Account Name - Last 4 of ID - YYYY-MM-DD
  var SHEET_NAME = accountName + ' - ' + last4Id + ' - ' + yesterdayStr;
  
  // Handle spreadsheet: create new if SHEET_URL not provided or placeholder
  var spreadsheet;
  if (!SHEET_URL || SHEET_URL.includes('YOUR_SPREADSHEET_ID')) {
    spreadsheet = SpreadsheetApp.create('Google Ads New Search Queries - ' + accountName + ' - ' + last4Id);
    SHEET_URL = spreadsheet.getUrl();
    Logger.log('Created new spreadsheet: ' + SHEET_URL);
  } else {
    spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  }
  
  // Get or create sheet
  var sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.appendRow(['Campaign Name', 'Campaign ID', 'Ad Group Name', 'Ad Group ID', 'Search Term', 'Impressions', 'Clicks', 'Cost']);
  newData.forEach(function(item) {
    sheet.appendRow([item.campaignName, item.campaignId, item.adGroupName, item.adGroupId, item.origQuery, item.impressions, item.clicks, item.cost]);
  });
  
  // Auto-size columns and freeze first row
  sheet.autoResizeColumns(1, 8);
  sheet.setFrozenRows(1);
  
  // Filter for email: only items with clicks > 0
  var emailData = newData.filter(function(item) {
    return item.clicks > 0;
  });
  
  // Prepare subject line
  var emailSubject = accountName + ' - New Search Queries - ' + yesterdayStr;
  
  // If no queries with >0 clicks for email, send a note
  if (emailData.length === 0) {
    var emailBody = '<html><body><p>No new search queries with >0 clicks from ' + yesterdayStr + '.</p>' +
                    '<p>View details (including zero-click queries) in the sheet: <a href="' + SHEET_URL + '">' + SHEET_URL + '</a></p></body></html>';
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL,
      subject: emailSubject,
      htmlBody: emailBody
    });
    Logger.log('Report generated and sent. No queries with >0 clicks.');
    return;
  }
  
  // Prepare email: group by campaign name, sort campaigns by total clicks desc, then alpha
  var campaignToItems = new Map();
  var campaignTotals = new Map();
  emailData.forEach(function(item) {
    var key = item.campaignName;
    if (!campaignToItems.has(key)) {
      campaignToItems.set(key, []);
      campaignTotals.set(key, 0);
    }
    campaignToItems.get(key).push(item);
    campaignTotals.set(key, campaignTotals.get(key) + item.clicks);
  });
  
  // Sort campaigns by total clicks desc, then name asc
  var sortedCampaigns = Array.from(campaignToItems.keys()).sort(function(a, b) {
    var diff = campaignTotals.get(b) - campaignTotals.get(a);
    return diff !== 0 ? diff : a.localeCompare(b);
  });
  
  // Build HTML email body
  var emailBody = '<html><body>' +
                  '<h2>New search queries with clicks from ' + yesterdayStr + ' grouped by campaign:</h2>';
  sortedCampaigns.forEach(function(camp) {
    emailBody += '<h3>Campaign: ' + camp + '</h3>' +
                 '<table border="1" style="border-collapse: collapse;">' +
                 '<thead><tr><th>Search Term</th><th>Impressions</th><th>Clicks</th><th>Cost</th></tr></thead>' +
                 '<tbody>';
    var items = campaignToItems.get(camp);
    // Sort items within campaign by clicks desc
    items.sort(function(a, b) {
      return b.clicks - a.clicks;
    });
    items.forEach(function(item) {
      emailBody += '<tr><td>' + item.origQuery + '</td>' +
                   '<td>' + item.impressions + '</td>' +
                   '<td>' + item.clicks + '</td>' +
                   '<td>' + item.cost + '</td></tr>';
    });
    emailBody += '</tbody></table><br>';
  });
  emailBody += '<p>View details in the sheet: <a href="' + SHEET_URL + '">' + SHEET_URL + '</a></p>' +
               '</body></html>';
  
  // Send email
  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: emailSubject,
    htmlBody: emailBody
  });
  
  Logger.log('Report generated and sent. New queries: ' + newData.length + '; Email queries (>0 clicks): ' + emailData.length);
}