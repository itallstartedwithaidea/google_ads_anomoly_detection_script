// Google Ads Anomaly Detection Script by John Williams with It All Started With A Idea

var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/your_spreadsheet_id_here';

function main() {
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  try {
    var config = loadConfiguration(spreadsheet);
    log(spreadsheet, "Script execution started");

    var THRESHOLDS = {
      CTR_DROP: parseFloat(config['CTR_DROP']),
      IMPRESSIONS_DROP: parseFloat(config['IMPRESSIONS_DROP']),
      CONVERSION_RATE_DROP: parseFloat(config['CONVERSION_RATE_DROP']),
      SPEND_CHANGE: parseFloat(config['SPEND_CHANGE']),
      NO_SPEND_DAYS: parseInt(config['NO_SPEND_DAYS']),
      LOW_IMPRESSION_SHARE: parseFloat(config['LOW_IMPRESSION_SHARE']),
      HIGH_BID_LIMIT: parseFloat(config['HIGH_BID_LIMIT']),
      LOW_QUALITY_SCORE: parseInt(config['LOW_QUALITY_SCORE']),
      FIRST_PAGE_BID_RATIO: parseFloat(config['FIRST_PAGE_BID_RATIO']),
      LOW_DEVICE_PERFORMANCE: parseFloat(config['LOW_DEVICE_PERFORMANCE']),
      MAX_EXECUTION_TIME: parseInt(config['MAX_EXECUTION_TIME']) * 60 * 1000,
      LOW_SEARCH_VOLUME: parseInt(config['LOW_SEARCH_VOLUME']),
      LOOKBACK_DAYS: parseInt(config['LOOKBACK_DAYS']),
      MAX_TOTAL_ANOMALIES: parseInt(config['MAX_TOTAL_ANOMALIES'] || '1000'),
      EMAIL_RECIPIENT: config['EMAIL_RECIPIENT']
    };

    var exclusions = loadExclusions(spreadsheet);
    var startTime = new Date().getTime();

    var anomalies = {
      impression: detectImpressionAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      click: detectClickAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      conversion: detectConversionAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      spend: detectSpendAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      bid: detectBidAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      budget: detectBudgetAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      adPerformance: detectAdPerformanceAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      keyword: detectKeywordAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      impressionShare: detectImpressionShareAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      qualityScore: detectQualityScoreAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      devicePerformance: detectDevicePerformanceAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions),
      adStatus: detectAdStatusAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions)
    };

    createSummaryDashboard(spreadsheet, anomalies);
    createDataVisualizations(spreadsheet, anomalies);
    performTrendAnalysis(spreadsheet, anomalies, THRESHOLDS);
    checkCustomAlerts(spreadsheet, config, anomalies);

    log(spreadsheet, "Script execution completed successfully");
    sendEmailSummary(THRESHOLDS.EMAIL_RECIPIENT, anomalies);
  } catch (e) {
    log(spreadsheet, "Error in script execution: " + e.message);
    sendErrorEmail(THRESHOLDS.EMAIL_RECIPIENT, e.message);
  }
}

function loadConfiguration(spreadsheet) {
  var configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    throw new Error('Config sheet not found. Please create a sheet named "Config" with appropriate settings.');
  }

  var config = {};
  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }

  return config;
}

function loadExclusions(spreadsheet) {
  var exclusionSheet = spreadsheet.getSheetByName('Exclusions');
  if (!exclusionSheet) {
    return { campaigns: [], adGroups: [], keywords: [] };
  }

  var exclusions = { campaigns: [], adGroups: [], keywords: [] };
  var data = exclusionSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var type = data[i][0].toLowerCase();
    var value = data[i][1];
    if (exclusions[type + 's']) {
      exclusions[type + 's'].push(value);
    }
  }

  return exclusions;
}

function log(spreadsheet, message) {
  var logSheet = spreadsheet.getSheetByName('Log');
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('Log');
    logSheet.appendRow(['Timestamp', 'Message']);
  }

  logSheet.appendRow([new Date(), message]);
  Logger.log(message);
}

function checkExecutionTime(startTime, THRESHOLDS) {
  var currentTime = new Date().getTime();
  if (currentTime - startTime > THRESHOLDS.MAX_EXECUTION_TIME) {
    throw new Error('Script exceeded maximum execution time limit.');
  }
}

function isExcluded(entityName, exclusions, type) {
  return exclusions[type + 's'].indexOf(entityName) !== -1;
}

function getDateRange(days) {
  var endDate = new Date();
  var startDate = new Date(endDate.getTime() - ((days - 1) * 24 * 60 * 60 * 1000));
  return formatDate(startDate) + ',' + formatDate(endDate);
}

function formatDate(date) {
  return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
}

// Detect Impression Anomalies Function
function detectImpressionAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Impression Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS * 2);
  var query = "SELECT CampaignName, Date, Impressions " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var campaignData = {};

    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;

      var date = row['Date'];
      var impressions = parseInt(row['Impressions']);

      if (!campaignData[campaignName]) {
        campaignData[campaignName] = { current: 0, previous: 0 };
      }

      var dateObj = new Date(date.substring(0, 4), parseInt(date.substring(4, 6)) - 1, date.substring(6, 8));
      var daysDiff = (new Date() - dateObj) / (1000 * 60 * 60 * 24);

      if (daysDiff < THRESHOLDS.LOOKBACK_DAYS) {
        campaignData[campaignName].current += impressions;
      } else {
        campaignData[campaignName].previous += impressions;
      }
    }

    var data = [];
    for (var campaign in campaignData) {
      var current = campaignData[campaign].current;
      var previous = campaignData[campaign].previous;
      var changePercentage = previous > 0 ? ((current - previous) / previous * 100).toFixed(2) : 'N/A';

      if (current < THRESHOLDS.LOW_SEARCH_VOLUME || (previous > 0 && current / previous < (1 - THRESHOLDS.IMPRESSIONS_DROP))) {
        var anomalyLabel = current < THRESHOLDS.LOW_SEARCH_VOLUME ? 'Low volume' : 'Significant drop';
        data.push([campaign, current, previous, changePercentage + '%', anomalyLabel]);
        anomalies.push({ campaign: campaign, current: current, previous: previous, change: changePercentage, type: anomalyLabel });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Current Impressions', 'Previous Impressions', 'Change %', 'Anomaly Type']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectImpressionAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Click Anomalies Function
function detectClickAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Click Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS * 2);
  var query = "SELECT CampaignName, Date, Clicks, Impressions " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var campaignData = {};

    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;

      var date = row['Date'];
      var clicks = parseInt(row['Clicks']);
      var impressions = parseInt(row['Impressions']);

      if (!campaignData[campaignName]) {
        campaignData[campaignName] = {
          currentClicks: 0,
          previousClicks: 0,
          currentImpressions: 0,
          previousImpressions: 0
        };
      }

      var dateObj = new Date(date.substring(0, 4), parseInt(date.substring(4, 6)) - 1, date.substring(6, 8));
      var daysDiff = (new Date() - dateObj) / (1000 * 60 * 60 * 24);

      if (daysDiff < THRESHOLDS.LOOKBACK_DAYS) {
        campaignData[campaignName].currentClicks += clicks;
        campaignData[campaignName].currentImpressions += impressions;
      } else {
        campaignData[campaignName].previousClicks += clicks;
        campaignData[campaignName].previousImpressions += impressions;
      }
    }

    var data = [];
    for (var campaign in campaignData) {
      var current = campaignData[campaign];
      var currentCTR = current.currentImpressions > 0 ? current.currentClicks / current.currentImpressions : 0;
      var previousCTR = current.previousImpressions > 0 ? current.previousClicks / current.previousImpressions : 0;
      var clicksChangePercentage = current.previousClicks > 0 ? ((current.currentClicks - current.previousClicks) / current.previousClicks * 100).toFixed(2) : 'N/A';
      var ctrChangePercentage = previousCTR > 0 ? ((currentCTR - previousCTR) / previousCTR * 100).toFixed(2) : 'N/A';

      var anomalyType = [];
      if (clicksChangePercentage !== 'N/A' && Math.abs(parseFloat(clicksChangePercentage)) >= THRESHOLDS.CTR_DROP * 100) {
        anomalyType.push('Clicks');
      }
      if (ctrChangePercentage !== 'N/A' && Math.abs(parseFloat(ctrChangePercentage)) >= THRESHOLDS.CTR_DROP * 100) {
        anomalyType.push('CTR');
      }

      if (anomalyType.length > 0) {
        data.push([
          campaign,
          current.currentClicks,
          (currentCTR * 100).toFixed(2) + '%',
          clicksChangePercentage + '%',
          ctrChangePercentage + '%',
          anomalyType.join(' & ')
        ]);
        anomalies.push({
          campaign: campaign,
          currentClicks: current.currentClicks,
          currentCTR: currentCTR,
          clicksChange: clicksChangePercentage,
          ctrChange: ctrChangePercentage,
          type: anomalyType.join(' & ')
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Clicks', 'CTR', 'Clicks Change %', 'CTR Change %', 'Anomaly Type']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectClickAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Conversion Anomalies Function
function detectConversionAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Conversion Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS * 2);
  var query = "SELECT CampaignName, Date, Conversions, Clicks " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var campaignData = {};

    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;

      var date = row['Date'];
      var conversions = parseFloat(row['Conversions']);
      var clicks = parseInt(row['Clicks']);

      if (!campaignData[campaignName]) {
        campaignData[campaignName] = {
          currentConv: 0,
          previousConv: 0,
          currentClicks: 0,
          previousClicks: 0
        };
      }

      var dateObj = new Date(date.substring(0, 4), parseInt(date.substring(4, 6)) - 1, date.substring(6, 8));
      var daysDiff = (new Date() - dateObj) / (1000 * 60 * 60 * 24);

      if (daysDiff < THRESHOLDS.LOOKBACK_DAYS) {
        campaignData[campaignName].currentConv += conversions;
        campaignData[campaignName].currentClicks += clicks;
      } else {
        campaignData[campaignName].previousConv += conversions;
        campaignData[campaignName].previousClicks += clicks;
      }
    }

    var data = [];
    for (var campaign in campaignData) {
      var current = campaignData[campaign];
      var currentRate = current.currentClicks > 0 ? current.currentConv / current.currentClicks : 0;
      var previousRate = current.previousClicks > 0 ? current.previousConv / current.previousClicks : 0;
      var convChangePercentage = current.previousConv > 0 ? ((current.currentConv - current.previousConv) / current.previousConv * 100).toFixed(2) : 'N/A';
      var rateChangePercentage = previousRate > 0 ? ((currentRate - previousRate) / previousRate * 100).toFixed(2) : 'N/A';

      if ((convChangePercentage !== 'N/A' && Math.abs(parseFloat(convChangePercentage)) >= THRESHOLDS.CONVERSION_RATE_DROP * 100) ||
          (rateChangePercentage !== 'N/A' && Math.abs(parseFloat(rateChangePercentage)) >= THRESHOLDS.CONVERSION_RATE_DROP * 100)) {
        data.push([
          campaign,
          current.currentConv.toFixed(2),
          (currentRate * 100).toFixed(2) + '%',
          convChangePercentage + '%',
          rateChangePercentage + '%'
        ]);
        anomalies.push({
          campaign: campaign,
          currentConversions: current.currentConv,
          currentConversionRate: currentRate,
          conversionsChange: convChangePercentage,
          conversionRateChange: rateChangePercentage
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Conversions', 'Conversion Rate', 'Conversions Change %', 'Conv. Rate Change %']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectConversionAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Spend Anomalies Function
function detectSpendAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Spend Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS * 2);
  var query = "SELECT CampaignName, Date, Cost " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var campaignData = {};

    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;

      var date = row['Date'];
      var cost = parseFloat(row['Cost']);

      if (!campaignData[campaignName]) {
        campaignData[campaignName] = { current: 0, previous: 0 };
      }

      var dateObj = new Date(date.substring(0, 4), parseInt(date.substring(4, 6)) - 1, date.substring(6, 8));
      var daysDiff = (new Date() - dateObj) / (1000 * 60 * 60 * 24);

      if (daysDiff < THRESHOLDS.LOOKBACK_DAYS) {
        campaignData[campaignName].current += cost;
      } else {
        campaignData[campaignName].previous += cost;
      }
    }

    var data = [];
    for (var campaign in campaignData) {
      var current = campaignData[campaign].current;
      var previous = campaignData[campaign].previous;
      var changePercentage = previous > 0 ? ((current - previous) / previous * 100).toFixed(2) : 'N/A';
      var anomalyType = '';

      if (current === 0) {
        anomalyType = 'No spend';
      } else if (previous === 0) {
        anomalyType = 'New spend';
      } else if (Math.abs(parseFloat(changePercentage)) >= THRESHOLDS.SPEND_CHANGE * 100) {
        anomalyType = 'Significant change';
      }

      if (anomalyType !== '') {
        data.push([campaign, current.toFixed(2), previous.toFixed(2), changePercentage + '%', anomalyType]);
        anomalies.push({
          campaign: campaign,
          currentSpend: current,
          previousSpend: previous,
          change: changePercentage,
          type: anomalyType
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Current Spend', 'Previous Spend', 'Change %', 'Anomaly Type']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectSpendAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Bid Anomalies Function
function detectBidAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Bid Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS);
  var query = "SELECT CampaignName, AdGroupName, Criteria, CpcBid, Clicks, Ctr, AverageCpc " +
              "FROM KEYWORDS_PERFORMANCE_REPORT " +
              "WHERE Status = ENABLED " +
              "AND CpcBid > " + THRESHOLDS.HIGH_BID_LIMIT + " " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];

      if (isExcluded(campaignName, exclusions, 'campaign') || isExcluded(adGroupName, exclusions, 'adGroup')) continue;

      var currentBid = parseFloat(row['CpcBid']);
      var clicks = parseInt(row['Clicks']);
      var ctr = parseFloat(row['Ctr'].replace('%', '')) / 100;
      var avgCpc = parseFloat(row['AverageCpc']);

      var potentialClicks = clicks > 0 ? Math.round(clicks / ctr * (THRESHOLDS.HIGH_BID_LIMIT / currentBid)) : 0;
      var impactLabel = potentialClicks > clicks ? 'Potential ' + (potentialClicks - clicks) + ' more clicks' : 'Overbidding';

      data.push([
        campaignName,
        adGroupName,
        row['Criteria'],
        currentBid.toFixed(2),
        THRESHOLDS.HIGH_BID_LIMIT.toFixed(2),
        impactLabel
      ]);
      anomalies.push({
        campaign: campaignName,
        adGroup: adGroupName,
        keyword: row['Criteria'],
        currentBid: currentBid,
        recommendedMaxBid: THRESHOLDS.HIGH_BID_LIMIT,
        potentialImpact: impactLabel
      });
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Ad Group Name', 'Keyword', 'Current Bid', 'Recommended Max Bid', 'Potential Impact']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectBidAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Budget Anomalies Function
function detectBudgetAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Budget Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var dateRange = getDateRange(THRESHOLDS.LOOKBACK_DAYS);
  var query = "SELECT CampaignName, Amount, Cost " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;
      
      var budget = parseFloat(row['Amount']);
      var cost = parseFloat(row['Cost']);
      var utilizationPercentage = (cost / budget * 100).toFixed(2);
      var changePercentage = ((budget - cost) / budget * 100).toFixed(2);
      if (cost < budget * 0.5) {
        data.push([
          campaignName,
          budget.toFixed(2),
          cost.toFixed(2),
          utilizationPercentage + '%',
          changePercentage + '%'
        ]);
        anomalies.push({
          campaign: campaignName,
          budget: budget,
          cost: cost,
          utilization: utilizationPercentage,
          change: changePercentage
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Budget', 'Cost', 'Utilization %', 'Difference %']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectBudgetAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Ad Performance Anomalies Function
function detectAdPerformanceAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Ad Performance Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var endDate = new Date();
  var startDate = new Date(endDate.getTime() - (THRESHOLDS.LOOKBACK_DAYS * 24 * 60 * 60 * 1000));
  var dateRange = formatDate(startDate) + ',' + formatDate(endDate);

  var query = "SELECT CampaignName, AdGroupName, Id, Status, AdType, Impressions, Clicks, Ctr " +
              "FROM AD_PERFORMANCE_REPORT " +
              "WHERE Impressions < " + THRESHOLDS.LOW_SEARCH_VOLUME + " OR Clicks = 0 " +
              "DURING " + dateRange;

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      
      if (isExcluded(campaignName, exclusions, 'campaign') || isExcluded(adGroupName, exclusions, 'adGroup')) continue;
      
      var impressions = parseInt(row['Impressions']);
      var clicks = parseInt(row['Clicks']);
      var ctr = parseFloat(row['Ctr'].replace('%', ''));
      
      var anomalyTypes = [];
      if (impressions < THRESHOLDS.LOW_SEARCH_VOLUME) anomalyTypes.push('Low Impressions');
      if (clicks === 0) anomalyTypes.push('No Clicks');
      if (ctr < THRESHOLDS.CTR_DROP * 100) anomalyTypes.push('Low CTR');
      
      var changePercentage = 'N/A';
      if (anomalyTypes.includes('Low CTR')) {
        changePercentage = ((THRESHOLDS.CTR_DROP * 100 - ctr) / (THRESHOLDS.CTR_DROP * 100) * 100).toFixed(2) + '%';
      }
      
      if (anomalyTypes.length > 0) {
        data.push([
          campaignName,
          adGroupName,
          row['Id'],
          row['Status'],
          row['AdType'],
          impressions,
          clicks,
          ctr + '%',
          anomalyTypes.join(', '),
          changePercentage
        ]);
        anomalies.push({
          campaign: campaignName,
          adGroup: adGroupName,
          adId: row['Id'],
          status: row['Status'],
          adType: row['AdType'],
          impressions: impressions,
          clicks: clicks,
          ctr: ctr,
          anomalyTypes: anomalyTypes,
          change: changePercentage
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Ad Group Name', 'Ad ID', 'Status', 'Ad Type', 'Impressions', 'Clicks', 'CTR', 'Anomaly Type', 'Change %']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectAdPerformanceAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Keyword Anomalies Function
function detectKeywordAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Keyword Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var query = "SELECT CampaignName, AdGroupName, Criteria, CpcBid, FirstPageCpc, Clicks, Ctr, AverageCpc " +
              "FROM KEYWORDS_PERFORMANCE_REPORT " +
              "WHERE Status = ENABLED " +
              "AND FirstPageCpc > 0 " +
              "DURING LAST_" + THRESHOLDS.LOOKBACK_DAYS + "_DAYS";

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      var keyword = row['Criteria'];
      
      if (isExcluded(campaignName, exclusions, 'campaign') || isExcluded(adGroupName, exclusions, 'adGroup') || isExcluded(keyword, exclusions, 'keyword')) continue;
      
      var currentBid = parseFloat(row['CpcBid']);
      var firstPageBid = parseFloat(row['FirstPageCpc']);
      var clicks = parseInt(row['Clicks']);
      var ctr = parseFloat(row['Ctr'].replace('%', '')) / 100;
      var avgCpc = parseFloat(row['AverageCpc']);
      
      if (currentBid < firstPageBid * THRESHOLDS.FIRST_PAGE_BID_RATIO) {
        var bidDifference = firstPageBid - currentBid;
        var expectedCost = bidDifference * clicks;
        var expectedClicks = Math.round((currentBid / avgCpc) * clicks * (1 - THRESHOLDS.FIRST_PAGE_BID_RATIO));
        
        data.push([
          campaignName,
          adGroupName,
          keyword,
          currentBid.toFixed(2),
          firstPageBid.toFixed(2),
          bidDifference.toFixed(2),
          expectedCost.toFixed(2),
          expectedClicks
        ]);
        anomalies.push({
          campaign: campaignName,
          adGroup: adGroupName,
          keyword: keyword,
          currentBid: currentBid,
          firstPageBid: firstPageBid,
          bidDifference: bidDifference,
          expectedCostIncrease: expectedCost,
          expectedMissedClicks: expectedClicks
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Ad Group Name', 'Keyword', 'Current Bid', 'First Page Bid', 'Bid Difference', 'Expected Cost Increase', 'Expected Missed Clicks']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectKeywordAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Impression Share Anomalies Function
function detectImpressionShareAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Impression Share Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var query = "SELECT CampaignName, SearchImpressionShare " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING LAST_" + THRESHOLDS.LOOKBACK_DAYS + "_DAYS";

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;
      
      var impressionShare = parseFloat(row['SearchImpressionShare'].replace('%', ''));
      var changePercentage = ((100 - impressionShare) / 100 * 100).toFixed(2);
      if (impressionShare < THRESHOLDS.LOW_IMPRESSION_SHARE) {
        data.push([
          campaignName,
          impressionShare.toFixed(2) + '%',
          changePercentage + '%'
        ]);
        anomalies.push({
          campaign: campaignName,
          impressionShare: impressionShare,
          change: changePercentage
        });
      }
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Search Impression Share', 'Change %']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectImpressionShareAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Quality Score Anomalies Function
function detectQualityScoreAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Quality Score Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var query = "SELECT CampaignName, AdGroupName, Criteria, QualityScore " +
              "FROM KEYWORDS_PERFORMANCE_REPORT " +
              "WHERE Status = ENABLED " +
              "AND QualityScore < " + THRESHOLDS.LOW_QUALITY_SCORE + " " +
              "DURING LAST_" + THRESHOLDS.LOOKBACK_DAYS + "_DAYS";

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      var keyword = row['Criteria'];
      
      if (isExcluded(campaignName, exclusions, 'campaign') || isExcluded(adGroupName, exclusions, 'adGroup') || isExcluded(keyword, exclusions, 'keyword')) continue;
      
      var qualityScore = parseInt(row['QualityScore']);
      data.push([
        campaignName,
        adGroupName,
        keyword,
        qualityScore
      ]);
      anomalies.push({
        campaign: campaignName,
        adGroup: adGroupName,
        keyword: keyword,
        qualityScore: qualityScore
      });
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Ad Group Name', 'Keyword', 'Quality Score']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectQualityScoreAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Device Performance Anomalies Function
function detectDevicePerformanceAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Device Performance Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var query = "SELECT CampaignName, Device, Conversions, ConversionRate " +
              "FROM CAMPAIGN_PERFORMANCE_REPORT " +
              "WHERE CampaignStatus = ENABLED " +
              "DURING LAST_" + THRESHOLDS.LOOKBACK_DAYS + "_DAYS";

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = {};
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      
      if (isExcluded(campaignName, exclusions, 'campaign')) continue;
      
      var device = row['Device'];
      var convRate = parseFloat(row['ConversionRate'].replace('%', ''));
      
      if (!data[campaignName]) {
        data[campaignName] = {};
      }
      data[campaignName][device] = convRate;
    }

    var anomalies = [];
    for (var campaign in data) {
      var desktop = data[campaign]['Computers'] || 0;
      var mobile = data[campaign]['Mobile devices with full browsers'] || 0;
      if (desktop > 0 && mobile / desktop < (1 - THRESHOLDS.LOW_DEVICE_PERFORMANCE)) {
        var changePercentage = ((desktop - mobile) / desktop * 100).toFixed(2);
        anomalies.push([campaign, 'Mobile', mobile.toFixed(2) + '%', desktop.toFixed(2) + '%', changePercentage + '%']);
        anomalies.push({
          campaign: campaign,
          underperformingDevice: 'Mobile',
          mobileConvRate: mobile,
          desktopConvRate: desktop,
          change: changePercentage
        });
      }
    }

    if (anomalies.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Underperforming Device', 'Device Conv Rate', 'Desktop Conv Rate', 'Difference %']);
      sheet.getRange(2, 1, anomalies.length, anomalies[0].length).setValues(anomalies);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectDevicePerformanceAnomalies: " + e.message);
  }

  return anomalies;
}

// Detect Ad Status Anomalies Function
function detectAdStatusAnomalies(spreadsheet, THRESHOLDS, startTime, exclusions) {
  checkExecutionTime(startTime, THRESHOLDS);
  var sheetName = 'Ad Status Anomalies';
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  var anomalies = [];

  var query = "SELECT CampaignName, AdGroupName, Id, Status, AdType, CreativeFinalUrls " +
              "FROM AD_PERFORMANCE_REPORT " +
              "WHERE Status != 'ENABLED' " +
              "DURING LAST_" + THRESHOLDS.LOOKBACK_DAYS + "_DAYS";

  try {
    var report = AdsApp.report(query);
    var rows = report.rows();
    var data = [];
    while (rows.hasNext()) {
      var row = rows.next();
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      
      if (isExcluded(campaignName, exclusions, 'campaign') || isExcluded(adGroupName, exclusions, 'adGroup')) continue;
      
      var status = row['Status'];
      var confirmationLabel = (status === 'PAUSED' || status === 'DISABLED') ? 'Are you sure?' : '';
      data.push([
        campaignName,
        adGroupName,
        row['Id'],
        status,
        row['AdType'],
        row['CreativeFinalUrls'],
        confirmationLabel
      ]);
      anomalies.push({
        campaign: campaignName,
        adGroup: adGroupName,
        adId: row['Id'],
        status: status,
        adType: row['AdType'],
        finalUrl: row['CreativeFinalUrls'],
        confirmation: confirmationLabel
      });
    }

    if (data.length > 0) {
      sheet.clear();
      sheet.appendRow(['Campaign Name', 'Ad Group Name', 'Ad ID', 'Status', 'Ad Type', 'Final URL', 'Confirmation']);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  } catch (e) {
    log(spreadsheet, "Error in detectAdStatusAnomalies: " + e.message);
  }

  return anomalies;
}

// Summary Dashboard Creation Function
function createSummaryDashboard(spreadsheet, anomalies) {
  var summarySheet = spreadsheet.getSheetByName('Summary') || spreadsheet.insertSheet('Summary');
  summarySheet.clear();

  summarySheet.appendRow(['Anomaly Type', 'Count', 'Severity']);
  var totalSeverity = 0;
  for (var type in anomalies) {
    var count = anomalies[type].length;
    var severity = calculateSeverity(type, count);
    summarySheet.appendRow([type, count, severity]);
    totalSeverity += severity;
  }

  summarySheet.appendRow(['Total', '', totalSeverity]);
}

// Severity Calculation Function
function calculateSeverity(type, count) {
  var baseScore = {
    'impression': 3,
    'click': 4,
    'conversion': 5,
    'spend': 5,
    'bid': 2,
    'budget': 4,
    'adPerformance': 3,
    'keyword': 2,
    'impressionShare': 3,
    'qualityScore': 2,
    'devicePerformance': 3,
    'adStatus': 4
  }[type] || 1;

  return baseScore * Math.log(count + 1);
}

// Data Visualization Creation Function
function createDataVisualizations(spreadsheet, anomalies) {
  var chartSheet = spreadsheet.getSheetByName('Charts') || spreadsheet.insertSheet('Charts');
  chartSheet.clear();

  // Create a column chart of anomaly counts
  var chartBuilder = chartSheet.newChart();
  var dataRange = chartSheet.getRange(1, 1, Object.keys(anomalies).length + 1, 2);
  dataRange.setValues([['Anomaly Type', 'Count']].concat(Object.keys(anomalies).map(function(type) {
    return [type, anomalies[type].length];
  })));

  chartBuilder.addRange(dataRange)
    .setChartType(Charts.ChartType.COLUMN)
    .setOption('title', 'Anomaly Counts by Type')
    .setOption('legend', 'none')
    .setPosition(1, 4, 0, 0);

  chartSheet.insertChart(chartBuilder.build());
}

// Trend Analysis Function
function performTrendAnalysis(spreadsheet, anomalies, THRESHOLDS) {
  var trendSheet = spreadsheet.getSheetByName('Trends') || spreadsheet.insertSheet('Trends');
  trendSheet.clear();

  trendSheet.appendRow(['Anomaly Type', 'Persistent Issues', 'Improving', 'Worsening']);

  for (var type in anomalies) {
    var persistentIssues = anomalies[type].filter(function(anomaly) {
      return anomaly.consecutiveDays >= 3;
    }).length;

    var improving = anomalies[type].filter(function(anomaly) {
      return parseFloat(anomaly.change) > 0;
    }).length;

    var worsening = anomalies[type].filter(function(anomaly) {
      return parseFloat(anomaly.change) < 0;
    }).length;

    trendSheet.appendRow([type, persistentIssues, improving, worsening]);
  }
}

// Custom Alerts Check Function
function checkCustomAlerts(spreadsheet, config, anomalies) {
  var alertSheet = spreadsheet.getSheetByName('Custom Alerts') || spreadsheet.insertSheet('Custom Alerts');
  alertSheet.clear();

  alertSheet.appendRow(['Alert Type', 'Description', 'Severity']);

  // Example custom alert: Check if total anomalies exceed a threshold
  var totalAnomalies = Object.values(anomalies).reduce((sum, arr) => sum + arr.length, 0);
  if (totalAnomalies > config['MAX_TOTAL_ANOMALIES']) {
    alertSheet.appendRow(['High Anomaly Count', 'Total anomalies (' + totalAnomalies + ') exceed threshold', 'High']);
  }

  // Add more custom alerts here based on your specific needs
}

// Email Summary Function
function sendEmailSummary(recipient, anomalies) {
  var subject = 'Google Ads Anomaly Detection Report';
  var body = 'Hello,\n\n';
  body += 'The anomaly detection script has completed its latest run.\n\n';
  body += 'Summary of findings:\n';
  
  var totalAnomalies = 0;
  for (var type in anomalies) {
    var count = anomalies[type].length;
    body += type + ' anomalies: ' + count + '\n';
    totalAnomalies += count;
  }
  
  body += '\nTotal anomalies detected: ' + totalAnomalies + '\n\n';
  body += 'Please review the Google Spreadsheet for detailed information on each anomaly.\n\n';
  body += 'Best regards,\n';
  body += 'Your Google Ads Monitoring Script';

  MailApp.sendEmail(recipient, subject, body);
}

// Error Email Function
function sendErrorEmail(recipient, errorMessage) {
  var subject = 'Error in Google Ads Anomaly Detection Script';
  var body = 'Hello,\n\n';
  body += 'An error occurred during the execution of the Google Ads Anomaly Detection Script.\n\n';
  body += 'Error message: ' + errorMessage + '\n\n';
  body += 'Please check the script and the Google Ads account for any issues.\n\n';
  body += 'Best regards,\n';
  body += 'Your Google Ads Monitoring Script';

  MailApp.sendEmail(recipient, subject, body);
}

// Helper Function: Get or Create Sheet
function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

// Helper Function: Store Historical Data
function storeHistoricalData(spreadsheet, metric, value) {
  var historySheet = getOrCreateSheet(spreadsheet, 'Historical Data');
  var today = formatDate(new Date());
  
  // Find or create row for today
  var dateColumn = historySheet.getRange('A:A');
  var dateValues = dateColumn.getValues();
  var rowIndex = dateValues.findIndex(row => row[0] === today);
  if (rowIndex === -1) {
    rowIndex = dateValues.length;
    historySheet.getRange(rowIndex + 1, 1).setValue(today);
  }
  
  // Find or create column for metric
  var headerRow = historySheet.getRange('1:1');
  var headers = headerRow.getValues()[0];
  var colIndex = headers.indexOf(metric);
  if (colIndex === -1) {
    colIndex = headers.length;
    historySheet.getRange(1, colIndex + 1).setValue(metric);
  }
  
  // Store the value
  historySheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
}

// Helper Function: Get Historical Data
function getHistoricalData(spreadsheet, metric, days) {
  var historySheet = getOrCreateSheet(spreadsheet, 'Historical Data');
  var headers = historySheet.getRange('1:1').getValues()[0];
  var colIndex = headers.indexOf(metric);
  
  if (colIndex === -1) return [];
  
  var data = historySheet.getRange(2, colIndex + 1, historySheet.getLastRow() - 1, 1).getValues();
  return data.slice(-days).map(row => row[0]).filter(val => val !== '');
}
