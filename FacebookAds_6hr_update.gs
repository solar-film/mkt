// ============================================================
//  Facebook Ads → Google Sheets  (Google Apps Script)
//  รองรับ 3 บัญชี: GFS / MHL / CAR
//  ✅ ดึงข้อมูลทั้งปี + เพิ่มรายวันอัตโนมัติ
// ============================================================

var CONFIG = {
  API_VERSION : "v21.0",
  TIMEZONE    : "Asia/Bangkok",

  ACCOUNTS: [
    { label: "GFS", id: "act_128825508332759"  },
    { label: "MHL", id: "act_617988735573201"  },
    { label: "CAR", id: "act_1382939935268827" },
  ],

  SHEET_OVERVIEW : "📊 Overview",
  SHEET_DAILY    : "📅 Daily All",
  SHEET_CAMPAIGN : "📋 By Campaign",
};

// ============================================================
//  🔧 SETUP TOKEN — รันครั้งแรก
// ============================================================
function setupToken() {
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt("🔑 Facebook Access Token", "วาง Token ของคุณ:", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var token = result.getResponseText().trim();
  if (!token) { ui.alert("❌ ไม่ได้ใส่ Token"); return; }
  PropertiesService.getScriptProperties().setProperty("FB_ACCESS_TOKEN", token);
  ui.alert("✅ บันทึก Token สำเร็จ!\n\nขั้นตอนต่อไป:\n1. รัน fetchYearToDate() เพื่อดึงข้อมูลทั้งปี\n2. รัน setupDailyTrigger() เพื่อตั้งอัปเดตอัตโนมัติทุก 6 ชั่วโมง");
}

function checkProperties() {
  var token   = PropertiesService.getScriptProperties().getProperty("FB_ACCESS_TOKEN") || "(ยังไม่ได้ตั้งค่า)";
  var preview = token.length > 10 ? token.substring(0, 10) + "..." : token;
  SpreadsheetApp.getUi().alert("📋 Script Properties", "FB_ACCESS_TOKEN: " + preview, SpreadsheetApp.getUi().ButtonSet.OK);
}

function clearToken() {
  PropertiesService.getScriptProperties().deleteProperty("FB_ACCESS_TOKEN");
  SpreadsheetApp.getUi().alert("🗑️ ลบ Token เรียบร้อยแล้ว");
}

function getFacebookAccessToken() {
  return PropertiesService.getScriptProperties().getProperty("FB_ACCESS_TOKEN");
}

// ============================================================
//  📅 ฟังก์ชัน 1: ดึงข้อมูลทั้งปี (รันครั้งแรก / รีเฟรชทั้งปี)
//  ช่วงวันที่: 1 ม.ค. ของปีนี้ → เมื่อวาน
// ============================================================
function fetchYearToDate() {
  if (!checkToken()) return;
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dateRange = getYearToDateRange();

  Logger.log("📅 ดึงข้อมูลทั้งปี: " + dateRange.since + " ถึง " + dateRange.until);
  ss.toast("⏳ กำลังดึงข้อมูลทั้งปี " + dateRange.since + " → " + dateRange.until + " ...", "FB Ads", 30);

  try {
    buildOverviewSheet(ss, dateRange);
    buildDailyAllSheet(ss, dateRange, false);   // false = เขียนใหม่ทั้งหมด
    buildCampaignSheet(ss, dateRange, false);

    var now = new Date().toLocaleString("th-TH", { timeZone: CONFIG.TIMEZONE });
    ss.toast("✅ ดึงข้อมูลทั้งปีสำเร็จ! " + now, "FB Ads Sync", 10);
    Logger.log("✅ fetchYearToDate เสร็จสมบูรณ์!");
  } catch (e) {
    Logger.log("❌ Error: " + e.message);
    ss.toast("❌ Error: " + e.message, "FB Ads Error", 15);
  }
}

// ============================================================
//  ⏰ ฟังก์ชัน 2: ดึงเฉพาะเมื่อวาน แล้วเพิ่มต่อท้าย (Auto Trigger)
// ============================================================
function fetchDailyUpdate() {
  if (!checkToken()) return;
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dateRange = getAutoUpdateRange();

  Logger.log("⏰ Daily Update: " + dateRange.since + " ถึง " + dateRange.until);

  try {
    appendDailyData(ss, dateRange);
    appendCampaignData(ss, dateRange);
    buildOverviewSheet(ss, getYearToDateRange());   

    var now = new Date().toLocaleString("th-TH", { timeZone: CONFIG.TIMEZONE });
    ss.toast("✅ อัปเดตข้อมูล " + formatDate(dateRange.since) + " ถึง " + formatDate(dateRange.until) + " สำเร็จ!", "FB Ads Daily", 6);
    Logger.log("✅ fetchDailyUpdate เสร็จสมบูรณ์!");
  } catch (e) {
    Logger.log("❌ Error: " + e.message);
    ss.toast("❌ Error: " + e.message, "FB Ads Error", 15);
  }
}

// ============================================================
//  SHEET — Overview (Year to Date)
// ============================================================
function buildOverviewSheet(ss, dateRange) {
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_OVERVIEW);
  var dates = generateDateList(dateRange.since, dateRange.until);

  var accountData = {};
  for (var i = 0; i < CONFIG.ACCOUNTS.length; i++) {
    var acc = CONFIG.ACCOUNTS[i];
    Logger.log("🔄 Overview: ดึง " + acc.label + "...");
    accountData[acc.label] = indexByDate(fetchInsights(acc.id, "account", dateRange));
  }

  var rows = dates.map(function(date) {
    var gfs = accountData["GFS"][date] || {};
    var mhl = accountData["MHL"][date] || {};
    var car = accountData["CAR"][date] || {};
    var gfsS = parseFloat(gfs.spend || 0);
    var mhlS = parseFloat(mhl.spend || 0);
    var carS = parseFloat(car.spend || 0);
    return [
      formatDate(date),
      gfsS.toFixed(2), parseInt(gfs.clicks||0), parseInt(gfs.impressions||0),
      mhlS.toFixed(2), parseInt(mhl.clicks||0), parseInt(mhl.impressions||0),
      carS.toFixed(2), parseInt(car.clicks||0), parseInt(car.impressions||0),
      (gfsS + mhlS + carS).toFixed(2)
    ];
  });

  // แถวรวม
  var sumRow = ["รวมทั้งปี","","","","","","","","","",""];
  [1,4,7,10].forEach(function(i) {
    sumRow[i] = rows.reduce(function(a,r){ return a + parseFloat(r[i]||0); }, 0).toFixed(2);
  });
  [2,3,5,6,8,9].forEach(function(i) {
    sumRow[i] = rows.reduce(function(a,r){ return a + parseInt(r[i]||0); }, 0);
  });
  rows.push(sumRow);

  sheet.clearContents().clearFormats();

  // Header
  var h1 = sheet.getRange(1,1,1,11);
  h1.setValues([["วันที่","GFS","","","MHL","","","CAR","","","รวม 3 บัญชี"]]);
  h1.setBackground("#1877F2").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  try { sheet.getRange(1,2,1,3).merge(); sheet.getRange(1,5,1,3).merge(); sheet.getRange(1,8,1,3).merge(); } catch(e){}

  var h2 = sheet.getRange(2,1,1,11);
  h2.setValues([["","Spend (฿)","Clicks","Impr.","Spend (฿)","Clicks","Impr.","Spend (฿)","Clicks","Impr.","Spend (฿)"]]);
  h2.setBackground("#4A90D9").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange(3,1,rows.length,11).setValues(rows);

  // Batch สี
  var bgColors = rows.slice(0, dates.length).map(function(_, i) {
    return Array(11).fill(i % 2 === 0 ? "#F0F7FF" : "#FFFFFF");
  });
  if (bgColors.length > 0) sheet.getRange(3,1,bgColors.length,11).setBackgrounds(bgColors);

  // แถวรวม
  sheet.getRange(3 + dates.length, 1, 1, 11)
    .setBackground("#1877F2").setFontColor("#FFFFFF").setFontWeight("bold");

  sheet.getRange(3,2,rows.length,10).setHorizontalAlignment("right");
  sheet.autoResizeColumns(1,11);
  sheet.setFrozenRows(2);
  Logger.log("✅ Overview: " + dates.length + " วัน");
}

// ============================================================
//  SHEET — Daily All (เขียนใหม่ทั้งหมด)
// ============================================================
function buildDailyAllSheet(ss, dateRange, appendMode) {
  var sheet   = getOrCreateSheet(ss, CONFIG.SHEET_DAILY);
  var headers = ["วันที่","บัญชี","ค่าใช้จ่าย (฿)","Impressions","Reach","Clicks","CPC (฿)","CPM (฿)","CTR (%)","Conversions","อัปเดตล่าสุด"];
  var now     = new Date().toLocaleString("th-TH", { timeZone: CONFIG.TIMEZONE });
  var allRows = [];

  for (var i = 0; i < CONFIG.ACCOUNTS.length; i++) {
    var acc = CONFIG.ACCOUNTS[i];
    Logger.log("🔄 Daily: ดึง " + acc.label + "...");
    fetchInsights(acc.id, "account", dateRange).forEach(function(d) {
      allRows.push([
        formatDate(d.date_start), acc.label,
        parseFloat(d.spend||0).toFixed(2),
        parseInt(d.impressions||0), parseInt(d.reach||0), parseInt(d.clicks||0),
        parseFloat(d.cpc||0).toFixed(2), parseFloat(d.cpm||0).toFixed(2),
        parseFloat(d.ctr||0).toFixed(2), extractConversions(d.actions), now
      ]);
    });
  }

  allRows.sort(function(a,b){ return parseThaiDate(b[0]) - parseThaiDate(a[0]); });

  if (!appendMode) {
    // เขียนใหม่ทั้งหมด
    sheet.clearContents().clearFormats();
    sheet.getRange(1,1,1,headers.length)
      .setValues([headers]).setBackground("#1877F2")
      .setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  }

  if (allRows.length > 0) {
    var startRow = appendMode ? sheet.getLastRow() + 1 : 2;
    sheet.getRange(startRow, 1, allRows.length, headers.length).setValues(allRows);

    // Batch สี
    var colorMap = { GFS: "#E3F2FD", MHL: "#FFF3E0", CAR: "#E8F5E9" };
    var bgColors = allRows.map(function(row) {
      return Array(headers.length).fill(colorMap[row[1]] || "#FFFFFF");
    });
    sheet.getRange(startRow, 1, allRows.length, headers.length).setBackgrounds(bgColors);
  }

  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  Logger.log("✅ Daily All: " + allRows.length + " แถว");
}

// ============================================================
//  SHEET — Campaign (เขียนใหม่ทั้งหมด)
// ============================================================
function buildCampaignSheet(ss, dateRange, appendMode) {
  var sheet   = getOrCreateSheet(ss, CONFIG.SHEET_CAMPAIGN);
  var headers = ["วันที่","บัญชี","Campaign","ค่าใช้จ่าย (฿)","Impressions","Reach","Clicks","CPC (฿)","CPM (฿)","CTR (%)","Conversions","อัปเดตล่าสุด"];
  var now     = new Date().toLocaleString("th-TH", { timeZone: CONFIG.TIMEZONE });
  var allRows = [];

  for (var i = 0; i < CONFIG.ACCOUNTS.length; i++) {
    var acc = CONFIG.ACCOUNTS[i];
    Logger.log("🔄 Campaign: ดึง " + acc.label + "...");
    fetchInsights(acc.id, "campaign", dateRange).forEach(function(d) {
      allRows.push([
        formatDate(d.date_start), acc.label, d.campaign_name||"-",
        parseFloat(d.spend||0).toFixed(2),
        parseInt(d.impressions||0), parseInt(d.reach||0), parseInt(d.clicks||0),
        parseFloat(d.cpc||0).toFixed(2), parseFloat(d.cpm||0).toFixed(2),
        parseFloat(d.ctr||0).toFixed(2), extractConversions(d.actions), now
      ]);
    });
  }

  allRows.sort(function(a,b){ return parseThaiDate(b[0]) - parseThaiDate(a[0]); });

  if (!appendMode) {
    sheet.clearContents().clearFormats();
    sheet.getRange(1,1,1,headers.length)
      .setValues([headers]).setBackground("#1877F2")
      .setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  }

  if (allRows.length > 0) {
    var startRow = appendMode ? sheet.getLastRow() + 1 : 2;
    sheet.getRange(startRow, 1, allRows.length, headers.length).setValues(allRows);

    var colorMap = { GFS: "#E3F2FD", MHL: "#FFF3E0", CAR: "#E8F5E9" };
    var bgColors = allRows.map(function(row) {
      return Array(headers.length).fill(colorMap[row[1]] || "#FFFFFF");
    });
    sheet.getRange(startRow, 1, allRows.length, headers.length).setBackgrounds(bgColors);
  }

  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  Logger.log("✅ Campaign: " + allRows.length + " แถว");
}

// ============================================================
//  APPEND — อัปเดตข้อมูลวันที่เดิมก่อนเพิ่มใหม่ (ใช้ใน Trigger ทุก 6 ชั่วโมง)
// ============================================================
function appendDailyData(ss, dateRange) {
  removeExistingRowsForDateRange(getOrCreateSheet(ss, CONFIG.SHEET_DAILY), dateRange);
  buildDailyAllSheet(ss, dateRange, true);  // true = append mode
}

function appendCampaignData(ss, dateRange) {
  removeExistingRowsForDateRange(getOrCreateSheet(ss, CONFIG.SHEET_CAMPAIGN), dateRange);
  buildCampaignSheet(ss, dateRange, true);  // true = append mode
}

function removeExistingRowsForDateRange(sheet, dateRange) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var dateSet = {};
  generateDateList(dateRange.since, dateRange.until).forEach(function(date) {
    dateSet[formatDate(date)] = true;
  });

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (dateSet[values[i][0]]) {
      sheet.deleteRow(i + 2);
    }
  }
}

// ============================================================
//  TRIGGER
// ============================================================
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  // ใช้ fetchDailyUpdate ทุก 6 ชั่วโมง โดยลบข้อมูลวันที่เดิมก่อนเพิ่มใหม่เพื่อกันข้อมูลซ้ำ
  ScriptApp.newTrigger("fetchDailyUpdate").timeBased().everyHours(6).create();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast("✅ Auto Trigger: ทุก 6 ชั่วโมง (fetchDailyUpdate)", "Setup Done", 5);
  Logger.log("✅ Trigger ตั้งค่าสำเร็จ: ทุก 6 ชั่วโมง");
}

function setupSixHourTrigger() {
  setupDailyTrigger();
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  Logger.log("🗑️ ลบ Trigger ทั้งหมดแล้ว");
}

// ============================================================
//  CORE API
// ============================================================
function fetchInsights(accountId, level, dateRange) {
  var fieldsMap = {
    account : "date_start,spend,impressions,reach,clicks,cpc,cpm,ctr,actions",
    campaign: "date_start,campaign_name,spend,impressions,reach,clicks,cpc,cpm,ctr,actions",
  };
  var token  = getFacebookAccessToken();
  var params = {
    fields        : fieldsMap[level] || fieldsMap.account,
    time_range    : JSON.stringify({ since: dateRange.since, until: dateRange.until }),
    time_increment: "1",
    level         : level,
    limit         : "500",
    access_token  : token,
  };
  var base  = "https://graph.facebook.com/" + CONFIG.API_VERSION + "/" + accountId + "/insights";
  var query = Object.entries(params)
    .map(function(e){ return encodeURIComponent(e[0]) + "=" + encodeURIComponent(e[1]); })
    .join("&");
  var resp = UrlFetchApp.fetch(base + "?" + query, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  var json = JSON.parse(resp.getContentText());
  if (code !== 200 || json.error) {
    throw new Error("[" + accountId + "] " + (json.error ? json.error.message : "HTTP " + code));
  }
  var results = json.data || [];
  var next = json.paging && json.paging.next;
  while (next) {
    var r2 = UrlFetchApp.fetch(next, { muteHttpExceptions: true });
    var j2 = JSON.parse(r2.getContentText());
    results = results.concat(j2.data || []);
    next = j2.paging && j2.paging.next;
  }
  return results;
}

// ============================================================
//  UTILITIES
// ============================================================
function getYearToDateRange() {
  var tz    = CONFIG.TIMEZONE;
  var today = new Date();
  var since = new Date(today.getFullYear(), 0, 1);  // 1 ม.ค. ปีนี้
  var until = new Date(today);  // วันนี้
  return {
    since: Utilities.formatDate(since, tz, "yyyy-MM-dd"),
    until: Utilities.formatDate(until, tz, "yyyy-MM-dd"),
  };
}

function getAutoUpdateRange() {
  var tz = CONFIG.TIMEZONE;
  var yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
  var today = new Date();
  return {
    since: Utilities.formatDate(yesterday, tz, "yyyy-MM-dd"),
    until: Utilities.formatDate(today, tz, "yyyy-MM-dd"),
  };
}

function getYesterdayRange() {
  var range = getAutoUpdateRange();
  return { since: range.since, until: range.since };
}

function generateDateList(since, until) {
  var dates = [], cur = new Date(since), end = new Date(until);
  while (cur <= end) {
    dates.push(Utilities.formatDate(cur, CONFIG.TIMEZONE, "yyyy-MM-dd"));
    cur.setDate(cur.getDate() + 1);
  }
  return dates;
}

function indexByDate(rows) {
  var map = {};
  (rows || []).forEach(function(r){ map[r.date_start] = r; });
  return map;
}

function formatDate(dateStr) {
  if (!dateStr) return "-";
  var p = dateStr.split("-");
  return p[2] + "/" + p[1] + "/" + p[0];
}

function parseThaiDate(s) {
  if (!s || s === "-") return new Date(0);
  var p = s.split("/");
  return new Date(p[2] + "-" + p[1] + "-" + p[0]);
}

function extractConversions(actions) {
  if (!actions) return 0;
  var types = ["purchase","lead","complete_registration","offsite_conversion.fb_pixel_purchase"];
  var found = actions.find(function(a){ return types.indexOf(a.action_type) >= 0; });
  return found ? parseInt(found.value || 0) : 0;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function checkToken() {
  var token = getFacebookAccessToken();
  if (!token) {
    SpreadsheetApp.getUi().alert("❌ ยังไม่ได้ตั้งค่า Token!\n\nกรุณารัน setupToken() ก่อน");
    return false;
  }
  return true;
}
