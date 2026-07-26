function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
                    .setTitle('Hệ Thống Quản Trị Lương - Vinapaco')
                    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === "sync_salary_data_pro") {
      var result = handleSyncSalaryDataPro(data);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "upload_attachment") {
      var uploadResult = handleUploadAttachment(data);
      return ContentService.createTextOutput(JSON.stringify(uploadResult)).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function autoCleanAndGetFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  var targetFolder = null;
  if (folders.hasNext()) {
    targetFolder = folders.next();
    while (folders.hasNext()) { try { folders.next().setTrashed(true); } catch(e) {} }
  } else {
    targetFolder = parentFolder.createFolder(folderName);
  }
  return targetFolder;
}

function handleUploadAttachment(data) {
  var rootFolder = autoCleanAndGetFolder(DriveApp.getRootFolder(), "Vinapaco quản lý lương hàng năm");
  var yearFolder = autoCleanAndGetFolder(rootFolder, data.year || "2026");
  var monthFolder = autoCleanAndGetFolder(yearFolder, "Tháng " + (data.month || "07"));
  var catFolder = autoCleanAndGetFolder(monthFolder, data.category || "Chứng từ");

  var decoded = Utilities.base64Decode(data.fileData.split(',')[1]);
  var blob = Utilities.newBlob(decoded, data.mimeType, data.fileName);
  var file = catFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { status: "success", fileUrl: file.getUrl() };
}

function handleSyncSalaryDataPro(data) {
  var year = data.year || "2026";
  var month = data.month || "07";
  var config = data.config || {};
  var staffs = data.staffs || [];

  var rootFolder = autoCleanAndGetFolder(DriveApp.getRootFolder(), "Vinapaco quản lý lương hàng năm");
  var yearFolder = autoCleanAndGetFolder(rootFolder, year);
  var monthFolder = autoCleanAndGetFolder(yearFolder, "Tháng " + month);

  var spreadsheetTitle = "BANG_LUONG_TONG_HOP_THANG_" + month + "_" + year;
  var spreadsheet;
  
  var filesInMonth = monthFolder.getFilesByName(spreadsheetTitle);
  if (filesInMonth.hasNext()) {
    spreadsheet = SpreadsheetApp.open(filesInMonth.next());
  } else {
    spreadsheet = SpreadsheetApp.create(spreadsheetTitle);
    DriveApp.getFileById(spreadsheet.getId()).moveTo(monthFolder);
  }

  var baseVal = parseFloat(config.base) || 2650000;
  var vungVal = parseFloat(config.vung) || 4140000;
  var lcvVal = parseFloat(config.lcv) || 1991139;

  // --- TAB 1: Lương CB ---
  var sheetCB = spreadsheet.getSheetByName("Lương CB");
  if (!sheetCB) { sheetCB = spreadsheet.insertSheet("Lương CB"); }
  sheetCB.clear();
  
  sheetCB.getRange("B2").setValue(config.companyTitle || "TỔNG CÔNG TY GIẤY VIỆT NAM").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetCB.getRange("B3").setValue(config.deptTitle || "VĂN PHÒNG").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetCB.getRange("B4").setValue("BẢNG LƯƠNG CƠ BẢN THÁNG " + month + "/" + year).setFontWeight("bold").setFontSize(13).setFontFamily("Times New Roman");

  sheetCB.getRange(6, 2, 1, 15).setValues([[
    "TT", "Họ và tên", "Hệ số lương cơ bản", "Phụ cấp", "Lương ký HĐLĐ", "Lương cơ bản", 
    "Ngày công", "", "", "", "Lương thời gian", "Lương lễ, phép thực tế", "Tổng lương CB", "BH (10.5%)", "Thực lĩnh CB"
  ]]);
  sheetCB.getRange(6, 8, 1, 3).merge().setValue("Ngày công");
  sheetCB.getRange("H7").setValue("Thời gian"); sheetCB.getRange("I7").setValue("Phép"); sheetCB.getRange("J7").setValue("Lễ, VR");

  var cbRows = [];
  staffs.forEach(function(s, idx) {
    var r = 8 + idx;
    cbRows.push([
      idx + 1, s.name || "", s.lcb_heso || 0, s.lcb_phucap || 0,
      "=ROUND((D" + r + "+E" + r + ")*" + vungVal + ", 0)", 
      "=ROUND((D" + r + "+E" + r + ")*" + baseVal + ", 0)",
      s.NT || 20, s.P || 0, s.le_vr || 1, "",
      "=ROUND((G" + r + "/22)*H" + r + ", 0)", 
      "=ROUND((F" + r + "/22)*(I" + r + "+J" + r + "), 0)",
      "=K" + r + "+L" + r, "=ROUND(F" + r + "*0.105, 0)", "=M" + r + "-N" + r
    ]);
  });
  if (cbRows.length > 0) {
    sheetCB.getRange(8, 2, cbRows.length, cbRows[0].length).setValues(cbRows);
    sheetCB.getRange(8, 6, cbRows.length, 2).setNumberFormat("#,##0");
    sheetCB.getRange(8, 11, cbRows.length, 5).setNumberFormat("#,##0");
  }
  var cbHeaderRange = sheetCB.getRange("B6:P7");
  cbHeaderRange.setBackground("#cbd5e1"); cbHeaderRange.setFontWeight("bold"); cbHeaderRange.setFontFamily("Times New Roman"); cbHeaderRange.setHorizontalAlignment("CENTER"); cbHeaderRange.setVerticalAlignment("MIDDLE"); cbHeaderRange.setWrap(true);
  sheetCB.autoResizeColumns(2, 16);

  // --- TAB 2: Lương CV ---
  var sheetCV = spreadsheet.getSheetByName("Lương CV");
  if (!sheetCV) { sheetCV = spreadsheet.insertSheet("Lương CV"); }
  sheetCV.clear();

  sheetCV.getRange("B2").setValue(config.companyTitle || "TỔNG CÔNG TY GIẤY VIỆT NAM").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetCV.getRange("B3").setValue(config.deptTitle || "VĂN PHÒNG").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetCV.getRange("B4").setValue("BẢNG LƯƠNG CÔNG VIỆC THÁNG " + month + "/" + year).setFontWeight("bold").setFontSize(13).setFontFamily("Times New Roman");

  sheetCV.getRange(6, 2, 1, 19).setValues([[
    "TT", "Họ và tên", "Hệ số cơ bản", "Phụ cấp", "Hệ số LCV", "Phụ cấp LCV", 
    "Lương ký HĐLĐ", "Lương công việc", "Ngày công", "", "", "", "LCV thực tế", "Ăn ca", "Độc hại", "Tạm ứng", "PCTT", "Trừ ứng", "Thực lĩnh CV", "Tổng NET"
  ]]);
  sheetCV.getRange(6, 10, 1, 3).merge().setValue("Ngày công");
  sheetCV.getRange("J7").setValue("Thời gian"); sheetCV.getRange("K7").setValue("Phép"); sheetCV.getRange("L7").setValue("Lễ, VR");

  var cvRows = [];
  staffs.forEach(function(s, idx) {
    var r = 8 + idx;
    var hsHtcv = parseFloat(s.hs_htcv) || 1.0;
    cvRows.push([
      idx + 1, s.name || "", s.lcb_heso || 0, s.lcb_phucap || 0, s.he_so_lcv || 0, s.phu_cap_lcv || 0,
      "=ROUND((D" + r + "+E" + r + ")*" + vungVal + ", 0)", 
      "=ROUND((F" + r + "+G" + r + ")*" + lcvVal + "*" + hsHtcv + ", 0)",
      s.NT || 20, s.P || 0, s.le_vr || 1, "",
      "=ROUND((I" + r + "/22)*J" + r + ", 0)", 
      s.anca || 1000000, s.dohai || 300000, s.tam_ung || 0, s.pctt || 0, s.tru_ung || 0,
      "=M" + r + "+N" + r + "+O" + r + "-P" + r + "-Q" + r + "-R" + r,
      "=S" + r + "+'Lương CB'!P" + r
    ]);
  });
  if (cvRows.length > 0) {
    sheetCV.getRange(8, 2, cvRows.length, cvRows[0].length).setValues(cvRows);
    sheetCV.getRange(8, 7, cvRows.length, 2).setNumberFormat("#,##0");
    sheetCV.getRange(8, 13, cvRows.length, 7).setNumberFormat("#,##0");
  }
  var cvHeaderRange = sheetCV.getRange("B6:T7");
  cvHeaderRange.setBackground("#cbd5e1"); cvHeaderRange.setFontWeight("bold"); cvHeaderRange.setFontFamily("Times New Roman"); cvHeaderRange.setHorizontalAlignment("CENTER"); cvHeaderRange.setVerticalAlignment("MIDDLE"); cvHeaderRange.setWrap(true);
  sheetCV.autoResizeColumns(2, 20);

  // --- TAB 3: ĂnCa ---
  var sheetAC = spreadsheet.getSheetByName("ĂnCa");
  if (!sheetAC) { sheetAC = spreadsheet.insertSheet("ĂnCa"); }
  sheetAC.clear();
  sheetAC.getRange("A1").setValue(config.companyTitle || "TỔNG CÔNG TY GIẤY VIỆT NAM").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetAC.getRange("A2").setValue(config.deptTitle || "VĂN PHÒNG").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetAC.getRange("A3").setValue("DANH SÁCH LĨNH TIỀN ĂN CA THÁNG " + month + "/" + year).setFontWeight("bold").setFontSize(13).setFontFamily("Times New Roman");

  sheetAC.getRange(5, 1, 1, 5).setValues([["STT", "HỌ VÀ TÊN", "SỐ CÔNG", "SỐ TIỀN ĂN CA", "KÝ NHẬN"]]);
  sheetAC.getRange("A5:E5").setBackground("#cbd5e1").setFontWeight("bold").setFontFamily("Times New Roman").setHorizontalAlignment("CENTER");

  var acRows = [];
  staffs.forEach(function(s, idx) {
    var nt = s.NT || 20;
    var ancaVal = s.anca !== undefined ? s.anca : (nt * 50000);
    acRows.push([idx + 1, s.name || "", nt, ancaVal, ""]);
  });
  if (acRows.length > 0) {
    sheetAC.getRange(6, 1, acRows.length, acRows[0].length).setValues(acRows);
    sheetAC.getRange(6, 4, acRows.length, 1).setNumberFormat("#,##0");
  }
  sheetAC.autoResizeColumns(1, 5);

  // --- TAB 4: ĐộcHại ---
  var sheetDH = spreadsheet.getSheetByName("ĐộcHại");
  if (!sheetDH) { sheetDH = spreadsheet.insertSheet("ĐộcHại"); }
  sheetDH.clear();
  sheetDH.getRange("A1").setValue(config.companyTitle || "TỔNG CÔNG TY GIẤY VIỆT NAM").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetDH.getRange("A2").setValue(config.deptTitle || "VĂN PHÒNG").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetDH.getRange("A3").setValue("DANH SÁCH BỒI DƯỠNG ĐỘC HẠI THÁNG " + month + "/" + year).setFontWeight("bold").setFontSize(13).setFontFamily("Times New Roman");

  sheetDH.getRange(5, 1, 1, 5).setValues([["STT", "HỌ VÀ TÊN", "ĐỊNH MỨC", "SỐ TIỀN ĐỘC HẠI", "KÝ NHẬN"]]);
  sheetDH.getRange("A5:E5").setBackground("#cbd5e1").setFontWeight("bold").setFontFamily("Times New Roman").setHorizontalAlignment("CENTER");

  var dhRows = [];
  staffs.forEach(function(s, idx) {
    dhRows.push([idx + 1, s.name || "", "Theo chế độ", s.dohai || 300000, ""]);
  });
  if (dhRows.length > 0) {
    sheetDH.getRange(6, 1, dhRows.length, dhRows[0].length).setValues(dhRows);
    sheetDH.getRange(6, 4, dhRows.length, 1).setNumberFormat("#,##0");
  }
  sheetDH.autoResizeColumns(1, 5);

  // --- TAB 5: Ngan hang ---
  var sheetNH = spreadsheet.getSheetByName("Ngan hang");
  if (!sheetNH) { sheetNH = spreadsheet.insertSheet("Ngan hang"); }
  sheetNH.clear();
  sheetNH.getRange("A1").setValue(config.companyTitle || "TỔNG CÔNG TY GIẤY VIỆT NAM").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetNH.getRange("A2").setValue(config.deptTitle || "VĂN PHÒNG").setFontWeight("bold").setFontFamily("Times New Roman");
  sheetNH.getRange("A4").setValue("BẢNG LƯƠNG THANH TOÁN QUA NGÂN HÀNG THÁNG " + month + "/" + year).setFontWeight("bold").setFontSize(13).setFontFamily("Times New Roman");

  sheetNH.getRange(6, 1, 1, 5).setValues([["STT", "HỌ VÀ TÊN", "TÀI KHOẢN", "SỐ TIỀN NET", "KÝ NHẬN"]]);
  sheetNH.getRange("A6:E6").setBackground("#cbd5e1").setFontWeight("bold").setFontFamily("Times New Roman").setHorizontalAlignment("CENTER");

  var nhRows = [];
  staffs.forEach(function(s, idx) {
    var r = 8 + idx;
    nhRows.push([idx + 1, s.name || "", "'" + (s.account || ""), "='Lương CV'!T" + r, ""]);
  });
  if (nhRows.length > 0) {
    sheetNH.getRange(7, 1, nhRows.length, nhRows[0].length).setValues(nhRows);
    sheetNH.getRange(7, 4, nhRows.length, 1).setNumberFormat("#,##0");
  }
  sheetNH.autoResizeColumns(1, 5);

  var defaultSheet = spreadsheet.getSheetByName("Sheet1");
  if (defaultSheet && spreadsheet.getNumSheets() > 1) {
    try { spreadsheet.deleteSheet(defaultSheet); } catch(e) {}
  }

  return { status: "success", fileUrl: spreadsheet.getUrl() };
}
