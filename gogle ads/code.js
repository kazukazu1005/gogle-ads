function escapeSingleQuotes(str) {
  return str.replace(/'/g, "\\'");
}

function main() {
  var spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1HVvWBwGQuaY34YcMMmvWuT8bAGax6okv8gip-dgzUP0/edit?gid=0#gid=0';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);

  var campaignSheet = spreadsheet.getSheetByName('Campaign Settings');
  if (!campaignSheet) {
    campaignSheet = spreadsheet.insertSheet('Campaign Settings');
    campaignSheet.getRange('A1').setValue('Target Campaign Name');
  }
  var targetCampaign = campaignSheet.getRange('A1').getValue();
  if (!targetCampaign) {
    Logger.log('ターゲットキャンペーン名が設定されていません。全キャンペーンのレポートを作成します。');
  } else {
    Logger.log('ターゲットキャンペーン: ' + targetCampaign);
  }

  // 1. 日別レポート
  var dateSheet = spreadsheet.getSheetByName('Date Report');
  if (!dateSheet) {
    dateSheet = spreadsheet.insertSheet('Date Report');
  } else {
    dateSheet.clear();
  }
  dateSheet.appendRow(['Date', 'Campaign Name', 'Impressions', 'Clicks', 'Cost', 'Views']);

  var dateQuery = "SELECT Date, CampaignName, Impressions, Clicks, Cost, VideoViews " +
                "FROM CAMPAIGN_PERFORMANCE_REPORT " +
                "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                "DURING LAST_30_DAYS";

  var dateReport = AdsApp.report(dateQuery);
  var dateRows = dateReport.rows();
  while (dateRows.hasNext()) {
    var row = dateRows.next();
    dateSheet.appendRow([row['Date'], row['CampaignName'], row['Impressions'], row['Clicks'], row['Cost'], row['VideoViews']]);
  }

  // 2. 時間別レポート
var hourSheet = spreadsheet.getSheetByName('Hourly Report');
if (!hourSheet) {
  hourSheet = spreadsheet.insertSheet('Hourly Report');
} else {
  hourSheet.clear();
}
hourSheet.appendRow(['Hour of Day', 'Campaign Name', 'Impressions', 'Clicks', 'Cost', 'Views']);

// 1日24時間の初期データを作成
var hoursOfDay = {};
for (var i = 0; i < 24; i++) {
  hoursOfDay[i] = {
    impressions: 0,
    clicks: 0,
    cost: 0,
    videoViews: 0
  };
}

var hourQuery = "SELECT HourOfDay, CampaignName, Impressions, Clicks, Cost, VideoViews " +
                "FROM CAMPAIGN_PERFORMANCE_REPORT " +
                "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                "DURING LAST_30_DAYS";
var hourReport = AdsApp.report(hourQuery);
var hourRows = hourReport.rows();

// 実際のデータを更新
while (hourRows.hasNext()) {
  var row = hourRows.next();
  var hour = parseInt(row['HourOfDay']);
  if (hoursOfDay.hasOwnProperty(hour)) {
    hoursOfDay[hour].impressions = row['Impressions'];
    hoursOfDay[hour].clicks = row['Clicks'];
    hoursOfDay[hour].cost = row['Cost'];
    hoursOfDay[hour].videoViews = row['VideoViews'];
  }
}

// 24時間分のデータをシートに出力
for (var hour in hoursOfDay) {
  hourSheet.appendRow([hour, targetCampaign, hoursOfDay[hour].impressions, hoursOfDay[hour].clicks, hoursOfDay[hour].cost, hoursOfDay[hour].videoViews]);
}

  // 3. エリア別レポート
  var geoSheet = spreadsheet.getSheetByName('Geo Report');
  if (!geoSheet) {
    geoSheet = spreadsheet.insertSheet('Geo Report');
  } else {
    geoSheet.clear();
  }
  geoSheet.appendRow(['Country/Territory', 'Region', 'Impressions', 'Clicks', 'Cost', 'Views']);

  var geoQuery = "SELECT CountryCriteriaId, RegionCriteriaId, Impressions, Clicks, Cost, VideoViews " +
               "FROM GEO_PERFORMANCE_REPORT  " +
               "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
               "DURING LAST_30_DAYS";
  var geoReport = AdsApp.report(geoQuery);
  var geoRows = geoReport.rows();
  while (geoRows.hasNext()) {
    var row = geoRows.next();
    geoSheet.appendRow([row['CountryCriteriaId'], row['RegionCriteriaId'], row['Impressions'], row['Clicks'], row['Cost'], row['VideoViews']]);
  }

  // 4. 年齢、デバイス、性別レポート
// 4. 年齢、デバイス、性別レポート
var demographicsSheet = spreadsheet.getSheetByName('Demographics Report');
if (!demographicsSheet) {
  demographicsSheet = spreadsheet.insertSheet('Demographics Report');
} else {
  demographicsSheet.clear();
}
demographicsSheet.appendRow(['Type', 'Criteria', 'Impressions', 'Clicks', 'Cost', 'Views']);

// 年齢別レポート
var ageMapping = {
  'AGE_RANGE_18_24': '18-24',
  'AGE_RANGE_25_34': '25-34',
  'AGE_RANGE_35_44': '35-44',
  'AGE_RANGE_45_54': '45-54',
  'AGE_RANGE_55_64': '55-64',
  'AGE_RANGE_65_UP': '65+',
  'AGE_RANGE_UNDETERMINED': 'Unknown'
};

var ageQuery = "SELECT Criteria, Impressions, Clicks, Cost, VideoViews " +
               "FROM AGE_RANGE_PERFORMANCE_REPORT " +
               "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
               "DURING LAST_30_DAYS";
var ageReport = AdsApp.report(ageQuery);
var ageRows = ageReport.rows();

while (ageRows.hasNext()) {
  var row = ageRows.next();
  var ageRange = ageMapping[row['Criteria']] || row['Criteria']; // マッピングがない場合は元の値を使用
  demographicsSheet.appendRow(['Age', ageRange, row['Impressions'], row['Clicks'], row['Cost'], row['VideoViews']]);
}


// デバイス別レポート
var deviceMapping = {
  'Computers': 'Computers',
  'Mobile devices with full browsers': 'Mobile devices',
  'Tablets with full browsers': 'Tablets',
  'TVs': 'TV' // テレビを追加
};

var deviceQuery = "SELECT Device, Impressions, Clicks, Cost, VideoViews " +
                  "FROM CAMPAIGN_PERFORMANCE_REPORT " +
                  "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                  "DURING LAST_30_DAYS";
var deviceReport = AdsApp.report(deviceQuery);
var deviceRows = deviceReport.rows();

while (deviceRows.hasNext()) {
  var row = deviceRows.next();
  var deviceType = deviceMapping[row['Device']] || row['Device']; // マッピングがない場合は元の値を使用
  demographicsSheet.appendRow(['Device', deviceType, row['Impressions'], row['Clicks'], row['Cost'], row['VideoViews']]);
}


// 性別別レポート
var genderMapping = {
  'MALE': 'Male',
  'FEMALE': 'Female',
  'UNDETERMINED': 'Unknown'
};

var genderQuery = "SELECT Criteria, Impressions, Clicks, Cost, VideoViews " +
                  "FROM GENDER_PERFORMANCE_REPORT " +
                  "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                  "DURING LAST_30_DAYS";
var genderReport = AdsApp.report(genderQuery);
var genderRows = genderReport.rows();

while (genderRows.hasNext()) {
  var row = genderRows.next();
  var genderType = genderMapping[row['Criteria']] || row['Criteria']; // マッピングがない場合は元の値を使用
  demographicsSheet.appendRow(['Gender', genderType, row['Impressions'], row['Clicks'], row['Cost'], row['VideoViews']]);
}

  // 5. プレイスメントレポート
   var placementSheet = spreadsheet.getSheetByName('Placement Report');
  if (!placementSheet) {
    placementSheet = spreadsheet.insertSheet('Placement Report');
  } else {
    placementSheet.clear();
  }
  placementSheet.appendRow(['Placement', 'Impressions', 'Clicks', 'Cost', 'VideoViews']);

  var placementQuery = "SELECT Domain, Impressions, Clicks, Cost, VideoViews " +
                       "FROM AUTOMATIC_PLACEMENTS_PERFORMANCE_REPORT " +
                       "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                       "DURING LAST_30_DAYS"; // ORDER BY を削除
 
  var placementReport = AdsApp.report(placementQuery);  
  var placementRows = placementReport.rows();
 
  var placementData = [];
 
  // 全ての結果を一旦保存
  while (placementRows.hasNext()) {
    var row = placementRows.next();
    placementData.push({
      placement: row['Domain'],
      impressions: parseInt(row['Impressions']),
      clicks: row['Clicks'],
      cost: row['Cost'],
      videoViews: row['VideoViews']
    });
  }

  // Impressionsでソートして上位50件を取得
  placementData.sort(function(a, b) {
    return b.impressions - a.impressions; // 降順にソート
  });
 
  var top50Placements = placementData.slice(0, 50); // 上位50件を取得

  // シートに書き込む
  for (var i = 0; i < top50Placements.length; i++) {
    var row = top50Placements[i];
    placementSheet.appendRow([row.placement, row.impressions, row.clicks, row.cost, row.videoViews]);
  }
//6 視聴率のレポート  
 var viewRateSheet = spreadsheet.getSheetByName('View Rate Report');
  if (!viewRateSheet) {
    viewRateSheet = spreadsheet.insertSheet('View Rate Report');
  } else {
    viewRateSheet.clear();
  }

  // シートのヘッダーを設定
  viewRateSheet.appendRow(['Campaign Name', '25% View Rate', '50% View Rate', '75% View Rate', '100% View Rate']);

  // VIDEO_PERFORMANCE_REPORTからデータを取得
  var viewRateQuery = "SELECT CampaignName, VideoQuartile25Rate, VideoQuartile50Rate, VideoQuartile75Rate, VideoQuartile100Rate " +
                      "FROM VIDEO_PERFORMANCE_REPORT " +
                      "WHERE CampaignName = '" + escapeSingleQuotes(targetCampaign) + "' " +
                      "DURING LAST_30_DAYS";

  var viewRateReport = AdsApp.report(viewRateQuery);
  var viewRateRows = viewRateReport.rows();

  // データをシートに書き込む
  while (viewRateRows.hasNext()) {
    var row = viewRateRows.next();
    viewRateSheet.appendRow([row['CampaignName'], row['VideoQuartile25Rate'], row['VideoQuartile50Rate'], row['VideoQuartile75Rate'], row['VideoQuartile100Rate']]);
  }

  Logger.log('全てのレポートがスプレッドシートに出力されました。');
}
