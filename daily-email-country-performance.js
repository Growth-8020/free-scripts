/**
 * Google Ads Script - Daily Email Country Performance
 * This script sends daily country level performance to a specified email
 * 
 * Features:
 * - Top Countries: Country-level performance based on user location (shows where users were physically located)
 * - Summary rows with totals and calculated metrics
 * 
 * Setup Instructions:
 * 1. Replace 'your-email@example.com' with your email
 * 2. Schedule the script to run as needed (e.g., daily, weekly)
 */

function main() {
  // Configuration
  const EMAIL_RECIPIENT = 'your-email@example.com'; // Comma-separated for multiple recipients
  const EMAIL_SUBJECT = 'Google Ads - Country Spend Report (Previous Day)';
  
  // Get yesterday's date
  const YESTERDAY = getDateString(new Date(new Date().getTime() - 24 * 3600 * 1000));
  
  // Initialize report data
  const reportData = getCountryData();
  
  // Generate and send email
  sendEmailReport(reportData, YESTERDAY, EMAIL_RECIPIENT, EMAIL_SUBJECT);
}

function getCountryData() {
  try {
    const countryData = new Map();
    const totals = {
      spend: 0,
      impressions: 0,
      clicks: 0,
      conversions: 0
    };

    // Get location criteria report to map CountryCriteriaId to CountryName
    const locationCriteriaReport = AdsApp.report(
      "SELECT CriteriaId, CountryName " +
      "FROM LOCATION_CRITERIA_REPORT");
    const locationRows = locationCriteriaReport.rows();
    const locationMap = new Map();

    while (locationRows.hasNext()) {
      const locationRow = locationRows.next();
      const criteriaId = locationRow['CriteriaId'];
      const countryName = locationRow['CountryName'];
      locationMap.set(criteriaId, countryName);
    }

    // Get campaign location target report
    const report = AdsApp.report(
      "SELECT CriteriaId, Clicks, Impressions, Cost, Conversions " +
      "FROM CAMPAIGN_LOCATION_TARGET_REPORT " +
      "WHERE Impressions > 0 " +
      "DURING YESTERDAY");
  
    const rows = report.rows();
    while (rows.hasNext()) {
      const row = rows.next();
      const criteriaId = row['CriteriaId'];
      const countryName = locationMap.get(criteriaId) || 'Unknown';
      const spend = parseFloat(row['Cost'].replace(/,/g, ''));
      
      if (spend > 0) {
        const impressions = parseInt(row['Impressions'].replace(/,/g, ''), 10);
        const clicks = parseInt(row['Clicks'].replace(/,/g, ''), 10);
        const conversions = parseFloat(row['Conversions'].replace(/,/g, ''));

        // Aggregate data by country
        if (!countryData.has(countryName)) {
          countryData.set(countryName, {
            name: countryName,
            spend: 0,
            impressions: 0,
            clicks: 0,
            conversions: 0
          });
        }
        
        const countryStats = countryData.get(countryName);
        countryStats.spend += spend;
        countryStats.impressions += impressions;
        countryStats.clicks += clicks;
        countryStats.conversions += conversions;
        
        // Update totals
        totals.spend += spend;
        totals.impressions += impressions;
        totals.clicks += clicks;
        totals.conversions += conversions;
      }
    }
    
    // Convert Map to array and calculate metrics
    const countryReport = Array.from(countryData.values()).map(country => ({
      ...country,
      ctr: country.impressions > 0 ? country.clicks / country.impressions : 0,
      cpc: country.clicks > 0 ? country.spend / country.clicks : 0
    }));
    
    // Sort by spend in descending order
    countryReport.sort((a, b) => b.spend - a.spend);

    return {
      countries: countryReport,
      totals: totals
    };
  } catch (e) {
    console.error('Error retrieving or processing data: ' + e);
    return {
      countries: [],
      totals: {
        spend: 0,
        impressions: 0,
        clicks: 0,
        conversions: 0
      }
    };
  }
}

function sendEmailReport(reportData, date, recipient, subject) {
  const htmlBody = generateEmailHtml(reportData, date);
  
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (e) {
    console.error('Error sending email: ' + e);
  }
}

function generateEmailHtml(reportData, date) {
  const { countries, totals } = reportData;
  
  let html = `
    <html>
      <head>
        <style>
          table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
          }
          th, td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
          }
          th {
            background-color: #f2f2f2;
          }
          tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          .total-row {
            font-weight: bold;
            background-color: #e6e6e6;
          }
          .header {
            margin-bottom: 20px;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>Country Spend Report for ${date}</h2>
          <p><strong>Total Account Spend: $${formatAmount(totals.spend)}</strong></p>
        </div>
        
        <table>
          <tr>
            <th>Country</th>
            <th>Spend</th>
            <th>Impressions</th>
            <th>Clicks</th>
            <th>CTR</th>
            <th>Avg. CPC</th>
            <th>Conversions</th>
          </tr>`;

  // Add country rows
  countries.forEach(country => {
    html += `
          <tr>
            <td>${country.name}</td>
            <td>$${formatAmount(country.spend)}</td>
            <td>${formatNumber(country.impressions)}</td>
            <td>${formatNumber(country.clicks)}</td>
            <td>${formatPercentage(country.ctr)}%</td>
            <td>$${formatAmount(country.cpc)}</td>
            <td>${formatNumber(country.conversions)}</td>
          </tr>`;
  });

  // Add total row
  html += `
          <tr class="total-row">
            <td>TOTAL</td>
            <td>$${formatAmount(totals.spend)}</td>
            <td>${formatNumber(totals.impressions)}</td>
            <td>${formatNumber(totals.clicks)}</td>
            <td>${formatPercentage(totals.clicks / totals.impressions)}%</td>
            <td>$${formatAmount(totals.clicks > 0 ? totals.spend / totals.clicks : 0)}</td>
            <td>${formatNumber(totals.conversions)}</td>
          </tr>
        </table>
      </body>
    </html>`;

  return html;
}

function getDateString(date) {
  return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
}

function formatAmount(amount) {
  return amount.toFixed(2);
}

function formatNumber(number) {
  return number.toLocaleString();
}

function formatPercentage(number) {
  return (number * 100).toFixed(2);
}