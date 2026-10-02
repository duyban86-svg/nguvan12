/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: HỆ THỐNG QUẢN LÝ THI & ĐỒNG BỘ ĐIỂM NGỮ VĂN 12 (NĂM HỌC 2026-2027)
 * =========================================================================================
 * Tương thích 100% với:
 *  1. P1_BAI_1_TIEU_THUYET.html (Bài 1 - Khả năng lớn lao của tiểu thuyết)
 *  2. P1_BAI_2_THE_GIOI_THO.html (Bài 2 - Những thế giới thơ)
 *  3. P2_THE_LOAI_TRUYEN.html (Chuyên đề: Đặc trưng thể loại Truyện & Tiểu thuyết)
 *  4. P2_THE_LOAI_THO.html (Chuyên đề: Đặc trưng thể loại Thơ)
 *  5. index.html (Hệ sinh thái Ôn thi Tốt nghiệp THPT & Auth Bridge SSO)
 * 
 * TÍNH NĂNG MỚI ĐÃ NÂNG CẤP:
 *  - Quản lý Khóa/Mở ca thi từ xa: Học sinh chỉ cần xác nhận danh tính, không cần gõ mã ca thi.
 *  - Hỗ trợ Tab 'DeTuLuan': Cấp phát 2 Đề Ôn tập + 3 Đề Đánh giá chính thức linh hoạt trực tiếp từ Sheet.
 *  - Chống gian lận & Chặn trùng lặp 100% (LockService + Deduplication).
 *  - Hỗ trợ JSONP (GET) và POST không bị chặn CORS trên mọi thiết bị di động & máy tính.
 * =========================================================================================
 */

var TEACHER_PIN_DEFAULT = "1358";
var EXAM_PASSCODE_DEFAULT = "123456";

// -----------------------------------------------------------------------------------------
// 1. XỬ LÝ YÊU CẦU GET (JSONP Callback & Tra cứu cấu hình, nộp bài, SSO)
// -----------------------------------------------------------------------------------------
function doGet(e) {
  try {
    var params = e && e.parameter ? e.parameter : {};
    var action = params.action;
    var callback = params.callback;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 0. NẾU GIÁO VIÊN MỞ TRỰC TIẾP TRÊN TRÌNH DUYỆT HOẶC BẤM NÚT QUẢN TRỊ (KHÔNG CÓ CALLBACK)
    if (!callback && (!action || action === "open" || action === "admin")) {
      var sheetUrl = ss.getUrl();
      var topicId = params.topicId || "P1_BAI_1";
      var config = getExamConfig(topicId);
      
      var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Quản trị Ca thi - Google Sheets</title>' +
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
        '<meta http-equiv="refresh" content="1;url=' + sheetUrl + '">' +
        '<style>' +
        'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #f8fafc; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }' +
        '.card { background: #1e293b; border-radius: 16px; padding: 32px; max-width: 540px; width: 100%; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.4); border: 2px solid #f59e0b; text-align: center; }' +
        'h2 { color: #fde047; margin: 0 0 10px; font-size: 1.4rem; }' +
        'p { color: #94a3b8; font-size: 0.95rem; line-height: 1.5; margin-bottom: 24px; }' +
        '.btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; background: #059669; color: #ffffff; text-decoration: none; padding: 14px 28px; border-radius: 10px; font-weight: 700; font-size: 1rem; border: none; cursor: pointer; transition: all 0.2s; box-shadow: 0 4px 12px rgba(5,150,105,0.3); }' +
        '.btn:hover { background: #10b981; transform: translateY(-2px); }' +
        '.info-box { background: #0f172a; border-radius: 10px; padding: 16px; margin-top: 24px; text-align: left; font-size: 0.85rem; border: 1px solid #334155; }' +
        '.info-row { display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid #1e293b; }' +
        '.info-row:last-child { border-bottom: none; }' +
        '.label { color: #94a3b8; font-weight: 600; }' +
        '.val { color: #38bdf8; font-weight: 700; font-family: monospace; }' +
        '</style>' +
        '</head><body>' +
        '<div class="card">' +
        '  <div style="font-size: 2.5rem; margin-bottom: 12px;">📊</div>' +
        '  <h2>CHUYỂN HƯỚNG ĐẾN GOOGLE TRANG TÍNH</h2>' +
        '  <p>Hệ thống đang tự động mở file Google Sheets quản trị ca thi và sổ điểm của Thầy Cô...</p>' +
        '  <a href="' + sheetUrl + '" class="btn" target="_top">👉 Bấm vào đây nếu không tự chuyển</a>' +
        '  <div class="info-box">' +
        '    <div style="font-weight: 700; color: #fde047; margin-bottom: 8px;">⚙️ THÔNG TIN CẤU HÌNH HIỆN TẠI TRÊN SHEET:</div>' +
        '    <div class="info-row"><span class="label">Trạng thái Ca thi:</span><span class="val">' + config.examStatus + '</span></div>' +
        '    <div class="info-row"><span class="label">Thời gian làm bài:</span><span class="val">' + config.duration + ' phút</span></div>' +
        '    <div class="info-row"><span class="label">Mã PIN Giáo viên:</span><span class="val">' + (config.teacherPin || config.masterPin) + '</span></div>' +
        '    <div class="info-row"><span class="label">Số đề Tự luận cấu hình:</span><span class="val">' + (config.essayPrompts ? ((config.essayPrompts.on_tap ? config.essayPrompts.on_tap.length : 0) + ' Ôn tập + ' + (config.essayPrompts.chinh_thuc ? config.essayPrompts.chinh_thuc.length : 0) + ' Chính thức') : 'Mặc định') + '</span></div>' +
        '  </div>' +
        '</div>' +
        '</body></html>';
        
      return HtmlService.createHtmlOutput(html)
        .setTitle("Quản trị Ca thi - Google Sheets")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (!action) action = "check";

    // A. KIỂM TRA TRẠNG THÁI CA THI & CẤU HÌNH ĐỀ THI (KÈM 5 ĐỀ TỰ LUẬN)
    if (action === "check" || action === "getExamConfig") {
      var topicId = params.topicId || "P1_BAI_1";
      var config = getExamConfig(topicId);
      return sendResponse(config, callback);
    }

    // B. NỘP BÀI THI QUA GET (JSONP)
    if (action === "submit") {
      var result = saveExamSubmission(params);
      return sendResponse(result, callback);
    }

    // C. GỬI PHẢN HỒI / BÁO LỖI
    if (action === "feedback") {
      var fbResult = saveFeedback(params);
      return sendResponse(fbResult, callback);
    }

    // D. TRA CỨU TÀI KHOẢN LIÊN KẾT GOOGLE SSO
    if (action === "check_email") {
      var email = (params.email || "").trim().toLowerCase();
      var student = findAccountByEmail(email);
      return sendResponse({
        status: "success",
        bound: !!student,
        student: student || null
      }, callback);
    }

    // E. LẤY DANH SÁCH TÀI KHOẢN
    if (action === "get_accounts") {
      var accounts = getAllAccounts();
      return sendResponse({
        status: "success",
        accounts: accounts
      }, callback);
    }

    // F. LẤY NGÂN HÀNG ĐỀ TỰ LUẬN RIÊNG
    if (action === "get_essays" || action === "getEssayPrompts") {
      var essaysTopicId = params.topicId || "P1_BAI_1";
      var essayData = getEssayPromptsFromSheet(essaysTopicId);
      return sendResponse({
        status: "success",
        topicId: essaysTopicId,
        essayPrompts: essayData
      }, callback);
    }

    // Mặc định
    return sendResponse({ status: "success", message: "Hệ thống Google Apps Script đang hoạt động tốt." }, callback);

  } catch (err) {
    return sendResponse({ status: "error", message: err.toString() }, callback);
  }
}

// -----------------------------------------------------------------------------------------
// 2. XỬ LÝ YÊU CẦU POST (Nộp bài JSON/Form, Báo lỗi, Nhập điểm)
// -----------------------------------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(15000);
  } catch(e) {}

  try {
    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    var action = params.action || "submit";

    if (action === "bind_account") {
      var bindResult = bindGoogleAccount(params);
      return sendJsonResponse(bindResult);
    }

    if (action === "feedback") {
      var fbRes = saveFeedback(params);
      return sendJsonResponse(fbRes);
    }

    if (action === "submit") {
      var subRes = saveExamSubmission(params);
      return sendJsonResponse(subRes);
    }

    if (action === "score_entry" || params.students || params.scoreType) {
      var scoreRes = saveGradebookEntry(params);
      return sendJsonResponse(scoreRes);
    }

    var defRes = saveExamSubmission(params);
    return sendJsonResponse(defRes);

  } catch (err) {
    return sendJsonResponse({ status: "error", message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// -----------------------------------------------------------------------------------------
// 3. HÀM XỬ LÝ LƯU KẾT QUẢ BÀI THI (TRẮC NGHIỆM & TỰ LUẬN)
// -----------------------------------------------------------------------------------------
function saveExamSubmission(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(15000);
  } catch(e) {}

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var topicId = d.topicId || "P1_BAI_1";
    var sheetName = d.sheetTarget || ("KẾT QUẢ - " + (d.examTopic || "ĐÁNH GIÁ NGỮ VĂN 12"));
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      setupResultSheetHeaders(sheet);
    } else if (sheet.getLastRow() === 0) {
      setupResultSheetHeaders(sheet);
    }

    var timestamp = d.timestamp || Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var className = d.className || d.studentClass || "";
    var stt = d.stt || d.studentStt || "";
    var studentName = d.studentName || "";
    var googleEmail = d.googleEmail || d.email || "";
    var mode = d.mode || "Đánh giá chính thức";
    
    var score = d.score !== undefined ? d.score : (d.mcqScore || 0);
    var correctCount = d.correctCount || "0/30";
    
    var essayPromptTitle = d.essayPromptTitle || "";
    var essayWords = d.essayWords || 0;
    var essayContent = d.essayContent || "";
    
    var duration = d.duration || d.timeSpent || "00:00";
    var tabSwitches = d.tabSwitches !== undefined ? d.tabSwitches : (d.violations || 0);
    var violationStatus = d.violationStatus || (tabSwitches > 0 ? ("Vi phạm " + tabSwitches + " lần") : "HỢP LỆ");
    
    var rank = d.rank || "";
    if (!rank && score !== "") {
      var sVal = parseFloat(score);
      rank = sVal >= 8.0 ? "GIỎI" : (sVal >= 6.5 ? "KHÁ" : (sVal >= 5.0 ? "TRUNG BÌNH" : "CẦN CỐ GẮNG"));
    }
    
    var submissionId = d.submissionId || ("SUB_" + new Date().getTime());

    // CHỐNG GỬI TRÙNG LẶP (DEDUPLICATION 100%):
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var checkRowsCount = Math.min(30, lastRow - 1);
      var startRow = lastRow - checkRowsCount + 1;
      var recentData = sheet.getRange(startRow, 1, checkRowsCount, 16).getValues();
      
      for (var i = recentData.length - 1; i >= 0; i--) {
        var rowTime = String(recentData[i][0] || "").trim();
        var rowClass = String(recentData[i][1] || "").trim();
        var rowName = String(recentData[i][3] || "").trim();
        var rowSubId = String(recentData[i][15] || "").trim();
        
        var isSameSubId = (submissionId && rowSubId && submissionId === rowSubId);
        var isSameStudentRecent = (rowClass === className && rowName === studentName && rowTime === timestamp);
        
        if (isSameSubId || isSameStudentRecent) {
          return {
            status: "success",
            message: "Bài thi của học sinh " + studentName + " (" + className + ") đã được ghi nhận trước đó (bỏ qua trùng lặp).",
            row: startRow + i,
            submissionId: submissionId,
            duplicate: true
          };
        }
      }
    }

    // Thêm dòng kết quả mới
    sheet.appendRow([
      timestamp,           // Cột 1: Thời gian nộp
      className,           // Cột 2: Lớp
      stt,                 // Cột 3: STT
      studentName,         // Cột 4: Họ và tên
      googleEmail,         // Cột 5: Email Google SSO
      mode,                // Cột 6: Chế độ thi
      score,               // Cột 7: Điểm Trắc nghiệm (/10.0)
      correctCount,        // Cột 8: Số câu đúng (/30)
      essayPromptTitle,    // Cột 9: Đề Tự luận
      essayWords,          // Cột 10: Số từ Tự luận
      essayContent,        // Cột 11: Bài viết Tự luận
      duration,            // Cột 12: Thời gian làm bài
      tabSwitches,         // Cột 13: Số lần rời tab
      violationStatus,     // Cột 14: Tình trạng quy chế
      rank,                // Cột 15: Xếp loại
      submissionId         // Cột 16: Mã bài nộp
    ]);

    var newRow = sheet.getLastRow();
    formatDataRow(sheet, newRow);

    return {
      status: "success",
      message: "Đã lưu kết quả thi thành công cho học sinh: " + studentName + " (" + className + ")",
      row: newRow,
      submissionId: submissionId
    };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// -----------------------------------------------------------------------------------------
// 4. HÀM TẠO TIÊU ĐỀ BẢNG ĐIỂM CHUẨN
// -----------------------------------------------------------------------------------------
function setupResultSheetHeaders(sheet) {
  var headers = [
    "Thời gian nộp",
    "Lớp",
    "STT",
    "Họ và tên",
    "Email Google",
    "Chế độ thi",
    "Điểm TN (/10.0)",
    "Số câu đúng",
    "Đề Tự luận",
    "Số từ TL",
    "Nội dung Bài viết Tự luận",
    "Thời gian làm",
    "Số lần rời Tab",
    "Tình trạng quy chế",
    "Xếp loại",
    "Mã bài nộp"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setFontWeight("bold").setFontSize(10).setFontFamily("Arial");
  hRange.setBackground("#0f766e").setFontColor("#ffffff");
  hRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  hRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 55);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 220);
  sheet.setColumnWidth(10, 90);
  sheet.setColumnWidth(11, 400);
  sheet.setColumnWidth(12, 110);
  sheet.setColumnWidth(13, 110);
  sheet.setColumnWidth(14, 140);
  sheet.setColumnWidth(15, 110);
  sheet.setColumnWidth(16, 160);
}

function formatDataRow(sheet, rowIdx) {
  var range = sheet.getRange(rowIdx, 1, 1, 16);
  range.setFontSize(9.5).setFontFamily("Arial");
  range.setVerticalAlignment("middle");
  
  sheet.getRange(rowIdx, 1).setHorizontalAlignment("center");
  sheet.getRange(rowIdx, 2).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(rowIdx, 3).setHorizontalAlignment("center");
  sheet.getRange(rowIdx, 4).setHorizontalAlignment("left").setFontWeight("bold");
  sheet.getRange(rowIdx, 5).setHorizontalAlignment("left");
  sheet.getRange(rowIdx, 6).setHorizontalAlignment("center");
  
  var scoreCell = sheet.getRange(rowIdx, 7);
  scoreCell.setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
  var scoreVal = parseFloat(scoreCell.getValue());
  if (scoreVal >= 8.0) {
    scoreCell.setBackground("#dcfce7").setFontColor("#166534");
  } else if (scoreVal >= 5.0) {
    scoreCell.setBackground("#fef9c3").setFontColor("#854d0e");
  } else {
    scoreCell.setBackground("#fee2e2").setFontColor("#991b1b");
  }
  
  sheet.getRange(rowIdx, 8).setHorizontalAlignment("center");
  sheet.getRange(rowIdx, 9).setHorizontalAlignment("left").setFontWeight("bold");
  sheet.getRange(rowIdx, 10).setHorizontalAlignment("center");
  sheet.getRange(rowIdx, 11).setHorizontalAlignment("left").setWrap(true);
  sheet.getRange(rowIdx, 12).setHorizontalAlignment("center");
  
  var violCell = sheet.getRange(rowIdx, 14);
  violCell.setHorizontalAlignment("center");
  if (violCell.getValue() !== "HỢP LỆ") {
    violCell.setBackground("#fee2e2").setFontColor("#b91c1c").setFontWeight("bold");
  } else {
    violCell.setFontColor("#15803d");
  }
  
  sheet.getRange(rowIdx, 15).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(rowIdx, 16).setHorizontalAlignment("center").setFontColor("#64748b");
  
  sheet.setRowHeight(rowIdx, 32);
}

// -----------------------------------------------------------------------------------------
// 5. QUẢN LÝ TAB 'DeTuLuan' (2 ĐỀ ÔN TẬP + 3 ĐỀ ĐÁNH GIÁ CHÍNH THỨC)
// -----------------------------------------------------------------------------------------
function getOrCreateEssayPromptsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DeTuLuan") || ss.getSheetByName("Đề Tự Luận") || ss.getSheetByName("De_Tu_Luan");
  
  if (!sheet) {
    sheet = ss.insertSheet("DeTuLuan");
    var headers = [
      "Chuyên đề (topicId)",
      "Chế độ (ON_TAP / CHINH_THUC)",
      "Mã đề",
      "Tiêu đề đề bài",
      "Đoạn trích & Yêu cầu đề bài tự luận (200 chữ)",
      "Gợi ý Dàn ý (Dành cho GV)"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold").setFontSize(10).setFontFamily("Arial");
    hRange.setBackground("#1e293b").setFontColor("#fde047");
    hRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 450);
    sheet.setColumnWidth(6, 350);

    // Điền dữ liệu mẫu chuẩn cho 4 chuyên đề
    var initialEssays = [
      // P1_BAI_1 (Tiểu thuyết)
      ["P1_BAI_1", "ON_TAP", "OT_1", "Đề Ôn tập 1 (Xuân Tóc Đỏ cứu quốc)", "Viết đoạn văn khoảng 200 chữ phân tích nghệ thuật trào phúng sắc sảo của Vũ Trọng Phụng qua đoạn trích Xuân Tóc Đỏ cứu quốc.", "Mở đoạn: Giới thiệu tác giả, tác phẩm, nghệ thuật trào phúng. Thân đoạn: Phân tích tình huống bi hài, ngôn ngữ trào phúng, nghệ thuật phóng đại. Kết đoạn: Khái quát giá trị hiện thực."],
      ["P1_BAI_1", "ON_TAP", "OT_2", "Đề Ôn tập 2 (Nghệ thuật xây dựng nhân vật trong Số đỏ)", "Viết đoạn văn khoảng 200 chữ phân tích nghệ thuật khắc họa nhân vật Xuân Tóc Đỏ trong tiểu thuyết Số đỏ của Vũ Trọng Phụng.", "Mở đoạn: Giới thiệu Vũ Trọng Phụng và nhân vật Xuân Tóc Đỏ. Thân đoạn: Bản chất lưu manh nhưng thức thời, sự biến ảo dị hợm khi bước chân vào xã hội thượng lưu, nghệ thuật biếm họa sắc sảo. Kết đoạn: Khẳng định giá trị hiện thực và sức sống của hình tượng."],
      ["P1_BAI_1", "CHINH_THUC", "CT_1", "Đề Chính thức 1 (Khả năng bao quát đời sống)", "Viết đoạn văn khoảng 200 chữ làm rõ khả năng to lớn của thể loại tiểu thuyết trong việc phản ánh chiều sâu hiện thực tâm lí con người qua Bài 1.", "Mở đoạn: Nêu đặc trưng tiểu thuyết. Thân đoạn: Khả năng phản ánh thế giới đa chiều, đi sâu vào góc khuất tâm lí. Kết đoạn: Ý nghĩa đối với người đọc."],
      ["P1_BAI_1", "CHINH_THUC", "CT_2", "Đề Chính thức 2 (Nghệ thuật trần thuật đa điểm nhìn)", "Từ các văn bản tiểu thuyết đã học ở Bài 1, viết đoạn văn khoảng 200 chữ phân tích tác dụng của điểm nhìn trần thuật đa chiều đối với sự hấp dẫn của tác phẩm.", "Mở đoạn: Giới thiệu nghệ thuật trần thuật. Thân đoạn: Tác dụng của sự chuyển dịch điểm nhìn, mở rộng không gian cảm xúc. Kết đoạn: Khẳng định dấu ấn cách tân nghệ thuật."],
      ["P1_BAI_1", "CHINH_THUC", "CT_3", "Đề Chính thức 3 (Số phận con người & Thời cuộc)", "Viết đoạn văn khoảng 200 chữ bàn về mối quan hệ giữa số phận cá nhân và biến động lịch sử thời đại được thể hiện qua các đoạn trích tiểu thuyết ở Bài 1.", "Mở đoạn: Dẫn dắt vấn đề. Thân đoạn: Tác động của hoàn cảnh tới số phận, sự lựa chọn nhân sinh. Kết đoạn: Bài học cuộc sống."],

      // P1_BAI_2 (Thế giới thơ)
      ["P1_BAI_2", "ON_TAP", "OT_1", "Đề Ôn tập 1 (Tây Tiến - Bi tráng)", "Viết đoạn văn khoảng 200 chữ phân tích vẻ đẹp bi tráng của hình tượng người lính Tây Tiến trong bài thơ cùng tên của Quang Dũng.", "Mở đoạn: Giới thiệu Quang Dũng và vẻ đẹp bi tráng. Thân đoạn: Hiện thực gian khổ hi sinh kết hợp với khí phách hào hùng, lãng mạn. Kết đoạn: Đánh giá tượng đài bất tử."],
      ["P1_BAI_2", "ON_TAP", "OT_2", "Đề Ôn tập 2 (Đàn ghi ta của Lor-ca)", "Viết đoạn văn khoảng 200 chữ làm rõ biểu tượng 'tiếng đàn bọt nước' và khát vọng tự do trong bài thơ Đàn ghi ta của Lor-ca của Thanh Thảo.", "Mở đoạn: Giới thiệu Thanh Thảo và biểu tượng tiếng đàn. Thân đoạn: Vẻ đẹp mong manh nhưng bất diệt, khúc tưởng niệm người nghệ sĩ. Kết đoạn: Khái quát tư tưởng nghệ thuật."],
      ["P1_BAI_2", "CHINH_THUC", "CT_1", "Đề Chính thức 1 (Tây Tiến - Hào hoa & Lãng mạn)", "Viết đoạn văn khoảng 200 chữ cảm nhận về chất hào hoa, tâm hồn lãng mạn và tình cảm gắn bó tha thiết của đoàn quân Tây Tiến với mảnh đất miền Tây.", "Mở đoạn: Giới thiệu chất lãng mạn Tây Tiến. Thân đoạn: Kỷ niệm đêm hội đuốc hoa, vẻ đẹp kiều diễm của con người và cảnh sắc miền Tây. Kết đoạn: Nét độc đáo của thơ Quang Dũng."],
      ["P1_BAI_2", "CHINH_THUC", "CT_2", "Đề Chính thức 2 (Thơ tượng trưng & Siêu thực)", "Từ các thi phẩm trong Bài 2, hãy viết đoạn văn khoảng 200 chữ phân tích nét độc đáo của hệ thống hình ảnh biểu tượng trong khuynh hướng thơ hiện đại.", "Mở đoạn: Dẫn dắt khuynh hướng thơ hiện đại. Thân đoạn: Sự chuyển dịch từ tả thực sang tượng trưng, chiều sâu cảm giác. Kết đoạn: Ý nghĩa đổi mới thơ ca."],
      ["P1_BAI_2", "CHINH_THUC", "CT_3", "Đề Chính thức 3 (Khát vọng sáng tạo của người nghệ sĩ)", "Viết đoạn văn khoảng 200 chữ bàn về sứ mệnh và khát vọng sáng tạo nghệ thuật cách tân của nhà thơ được gợi ra từ Bài 2 - Những thế giới thơ.", "Mở đoạn: Đặt vấn đề sứ mệnh nghệ thuật. Thân đoạn: Tinh thần dấn thân, đổi mới thi cảm, khát vọng tự do. Kết đoạn: Liên hệ trách nhiệm tuổi trẻ hôm nay."],

      // P2_TRUYEN (Chuyên đề Truyện)
      ["P2_TRUYEN", "ON_TAP", "OT_1", "Đề Ôn tập 1 (Nghệ thuật xây dựng nhân vật)", "Viết đoạn văn khoảng 200 chữ phân tích vai trò của độc thoại nội tâm trong việc khắc họa tính cách nhân vật truyện hiện đại.", "Mở đoạn: Nêu vai trò độc thoại nội tâm. Thân đoạn: Khám phá chiều sâu tiềm thức, sự giằng xé tâm lí. Kết đoạn: Đánh giá thành công thể loại."],
      ["P2_TRUYEN", "ON_TAP", "OT_2", "Đề Ôn tập 2 (Tình huống truyện độc đáo)", "Viết đoạn văn khoảng 200 chữ phân tích ý nghĩa của tình huống truyện đối với việc bộc lộ chủ đề và tư tưởng tác phẩm.", "Mở đoạn: Khái niệm tình huống truyện. Thân đoạn: Tác động tạo bước ngoặt, bộc lộ bản chất nhân vật. Kết đoạn: Khẳng định tài năng nhà văn."],
      ["P2_TRUYEN", "CHINH_THUC", "CT_1", "Đề Chính thức 1 (Điểm nhìn trần thuật & Ngôi kể)", "Viết đoạn văn khoảng 200 chữ làm rõ tác dụng của việc dịch chuyển điểm nhìn trần thuật trong tác phẩm truyện hiện đại.", "Mở đoạn: Giới thiệu điểm nhìn trần thuật. Thân đoạn: Tạo tính khách quan, đa thanh, làm phong phú cách tiếp cận hiện thực. Kết đoạn: Giá trị cách tân."],
      ["P2_TRUYEN", "CHINH_THUC", "CT_2", "Đề Chính thức 2 (Chi tiết nghệ thuật đắt giá)", "Viết đoạn văn khoảng 200 chữ phân tích sức sống của một chi tiết nghệ thuật giàu ý nghĩa biểu tượng trong tác phẩm tự sự.", "Mở đoạn: Vị trí chi tiết nghệ thuật. Thân đoạn: Chi tiết nhỏ làm nên nhà văn lớn, sức nén cảm xúc. Kết đoạn: Ấn tượng sâu đậm."],
      ["P2_TRUYEN", "CHINH_THUC", "CT_3", "Đề Chính thức 3 (Không - thời gian nghệ thuật)", "Viết đoạn văn khoảng 200 chữ phân tích sự đan xen giữa thời gian hiện tại và thời gian hồi tưởng trong truyện hiện đại.", "Mở đoạn: Đặt vấn đề thời gian nghệ thuật. Thân đoạn: Cấu trúc phi tuyến tính, chiều sâu tâm trạng nhân vật. Kết đoạn: Hiệu quả thẩm mĩ."],

      // P2_THO (Chuyên đề Thơ)
      ["P2_THO", "ON_TAP", "OT_1", "Đề Ôn tập 1 (Cấu tứ & Hình ảnh thơ)", "Viết đoạn văn khoảng 200 chữ phân tích vai trò của cấu tứ trong việc tổ chức và phát triển mạch cảm xúc của một bài thơ trữ tình.", "Mở đoạn: Khái niệm cấu tứ. Thân đoạn: Trục liên kết cảm xúc, sáng tạo hình ảnh độc đáo. Kết đoạn: Ý nghĩa đối với tác phẩm."],
      ["P2_THO", "ON_TAP", "OT_2", "Đề Ôn tập 2 (Nhạc điệu & Ngôn ngữ thơ)", "Viết đoạn văn khoảng 200 chữ làm rõ nét độc đáo của nhịp điệu và phối thanh trong thơ hiện đại.", "Mở đoạn: Nêu vị trí nhạc điệu thơ. Thân đoạn: Ngắt nhịp linh hoạt, hòa âm cảm xúc, tăng tính gợi cảm. Kết đoạn: Chiều sâu thẩm mĩ."],
      ["P2_THO", "CHINH_THUC", "CT_1", "Đề Chính thức 1 (Hình tượng cái Tôi trữ tình)", "Viết đoạn văn khoảng 200 chữ phân tích sự vận động của cái tôi trữ tình từ thơ truyền thống sang thơ hiện đại và hậu hiện đại.", "Mở đoạn: Khái niệm cái tôi trữ tình. Thân đoạn: Sự chuyển dịch từ cảm xúc đơn tuyến sang đa chiều, tự do bộc lộ cá tính. Kết đoạn: Ý nghĩa đổi mới."],
      ["P2_THO", "CHINH_THUC", "CT_2", "Đề Chính thức 2 (Biểu tượng & Nghệ thuật ẩn dụ)", "Viết đoạn văn khoảng 200 chữ phân tích sức gợi mở của hệ thống hình ảnh mang tính tượng trưng trong thơ thế kỉ XX.", "Mở đoạn: Giới thiệu hình ảnh tượng trưng. Thân đoạn: Đa tầng ý nghĩa, kích thích liên tưởng sáng tạo của người đọc. Kết đoạn: Vẻ đẹp thi ca."],
      ["P2_THO", "CHINH_THUC", "CT_3", "Đề Chính thức 3 (Cảm thức thời gian trong thơ)", "Viết đoạn văn khoảng 200 chữ bàn về cảm thức thời gian và khát vọng bất tử hóa cái đẹp qua các thi phẩm chuyên đề Thơ.", "Mở đoạn: Nêu cảm thức thời gian. Thân đoạn: Nhận thức về sự hữu hạn của đời người và sức sống vĩnh cửu của cái đẹp. Kết đoạn: Thông điệp nhân văn."]
    ];

    sheet.getRange(2, 1, initialEssays.length, initialEssays[0].length).setValues(initialEssays);
  }

  return sheet;
}

function getEssayPromptsFromSheet(topicId) {
  var sheet = getOrCreateEssayPromptsSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { on_tap: [], chinh_thuc: [] };

  var cleanTopic = String(topicId || "").toLowerCase().trim();
  var on_tap = [];
  var chinh_thuc = [];

  for (var r = 1; r < data.length; r++) {
    var rowTopic = String(data[r][0] || "").toLowerCase().trim();
    var mode = String(data[r][1] || "").toUpperCase().trim();
    var code = String(data[r][2] || "").trim();
    var title = String(data[r][3] || "").trim();
    var prompt = String(data[r][4] || "").trim();
    var outline = String(data[r][5] || "").trim();

    // Khớp chuyên đề
    var isMatch = (rowTopic === cleanTopic || 
      (cleanTopic.indexOf("truyen") !== -1 && rowTopic.indexOf("truyen") !== -1) || 
      (cleanTopic.indexOf("tho") !== -1 && rowTopic.indexOf("tho") !== -1) || 
      (cleanTopic.indexOf("bai_1") !== -1 && rowTopic.indexOf("bai_1") !== -1) ||
      (cleanTopic.indexOf("bai_2") !== -1 && rowTopic.indexOf("bai_2") !== -1) ||
      rowTopic === "all");

    if (isMatch && title && prompt) {
      var item = {
        id: code || ("DE_" + r),
        title: title,
        prompt: prompt,
        outline: outline
      };

      if (mode.indexOf("ON_TAP") !== -1 || mode.indexOf("ÔN TẬP") !== -1 || mode.indexOf("PRACTICE") !== -1) {
        on_tap.push(item);
      } else {
        chinh_thuc.push(item);
      }
    }
  }

  return {
    on_tap: on_tap,
    chinh_thuc: chinh_thuc
  };
}

function getOrCreateConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().toLowerCase().replace(/\s+/g, '');
    if (name.indexOf("cauhinh") !== -1 || name.indexOf("cấuhình") !== -1 || name.indexOf("config") !== -1 || name.indexOf("thiếtlập") !== -1) {
      return sheets[i];
    }
  }
  return ss.getSheetByName("CẤU HÌNH") || ss.getSheetByName("CauHinh") || ss.getSheetByName("Cấu hình");
}

// -----------------------------------------------------------------------------------------
// 6. LẤY CẤU HÌNH CA THI & ĐỀ THI ĐỘNG (HỖ TRỢ BẢNG TỪNG BÀI THI)
// -----------------------------------------------------------------------------------------
function getExamConfig(topicId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var durationDefault = 40;
  
  var statusVal = "CLOSED"; // Mặc định khóa an toàn nếu chưa tìm thấy
  var codeVal = EXAM_PASSCODE_DEFAULT;
  var durVal = durationDefault;
  var annVal = "Chào mừng các em tham gia bài đánh giá năng lực Ngữ văn 12!";
  var pinVal = TEACHER_PIN_DEFAULT;

  var cleanTopic = String(topicId || "P1_BAI_1").toUpperCase().replace(/\s+/g, '').trim();

  var sheet = getOrCreateConfigSheet();
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length >= 2) {
      
      // 1. DUYỆT BẢNG NHIỀU CHỦ ĐỀ / BÀI THI (Định dạng chuẩn của Thầy):
      // Cột A: STT | Cột B: MÃ BÀI THI (ID) | Cột C: TÊN BÀI THI | Cột D: TRẠNG THÁI | Cột E: MÃ CA THI | Cột F: THỜI GIAN | Cột G: LỜI NHẮN
      var matchedTopicRow = false;
      for (var r = 1; r < data.length; r++) {
        var rowId = String(data[r][1] || "").toUpperCase().replace(/\s+/g, '').trim(); // Cột B
        var rowName = String(data[r][2] || "").toUpperCase().trim();                   // Cột C
        var rowA = String(data[r][0] || "").toUpperCase().trim();                      // Cột A
        
        // Kiểm tra dòng chứa Mã PIN Giáo Viên
        if (rowId.indexOf("PIN") !== -1 || rowName.indexOf("PIN") !== -1 || rowA.indexOf("PIN") !== -1 || rowName.indexOf("GIÁO VIÊN") !== -1) {
          var foundPin = String(data[r][3] || data[r][4] || data[r][2] || "").trim();
          if (foundPin && /^\d{3,6}$/.test(foundPin)) {
            pinVal = foundPin;
          }
        }

        // Khớp Mã bài thi (VD: P1_BAI_1, P1_BAI_2, P2_THO, P2_TRUYEN, ...)
        var isTopicMatch = false;
        if (rowId) {
          if (rowId === cleanTopic) isTopicMatch = true;
          else if (cleanTopic.indexOf(rowId) !== -1 || rowId.indexOf(cleanTopic) !== -1) isTopicMatch = true;
          else if (cleanTopic.indexOf("BAI_1") !== -1 && rowId.indexOf("BAI_1") !== -1) isTopicMatch = true;
          else if (cleanTopic.indexOf("BAI_2") !== -1 && rowId.indexOf("BAI_2") !== -1) isTopicMatch = true;
          else if (cleanTopic.indexOf("TRUYEN") !== -1 && rowId.indexOf("TRUYEN") !== -1) isTopicMatch = true;
          else if (cleanTopic.indexOf("THO") !== -1 && rowId.indexOf("THO") !== -1) isTopicMatch = true;
        }

        if (isTopicMatch && !matchedTopicRow) {
          matchedTopicRow = true;
          
          // Cột D (index 3): Trạng thái (OPEN / CLOSED)
          var rawSt = String(data[r][3] || "").toUpperCase().trim();
          if (rawSt.indexOf("OPEN") !== -1 || rawSt.indexOf("MỞ") !== -1 || rawSt.indexOf("MO") !== -1 || rawSt.indexOf("BẬT") !== -1) {
            statusVal = "OPEN";
          } else {
            statusVal = "CLOSED";
          }

          // Cột E (index 4): Mã ca thi
          if (data[r][4]) codeVal = String(data[r][4]).trim();

          // Cột F (index 5): Thời gian làm bài (Phút)
          if (data[r][5]) {
            var parsedDur = parseInt(data[r][5]);
            if (!isNaN(parsedDur) && parsedDur > 0) durVal = parsedDur;
          }

          // Cột G (index 6): Lời nhắn từ Thầy Cô
          if (data[r][6]) {
            var customAnn = String(data[r][6]).trim();
            if (customAnn) annVal = customAnn;
          }
        }
      }

      // 2. DỰ PHÒNG: Nếu không khớp theo bảng bài thi, thử quét theo bảng Key-Value dạng dọc
      if (!matchedTopicRow) {
        for (var r = 0; r < data.length; r++) {
          var k = String(data[r][0] || "").toLowerCase().trim();
          var v = String(data[r][1] || "").trim();
          if (!k) continue;

          if (k.indexOf("trạng thái") !== -1 || k.indexOf("status") !== -1) {
            var rawSt2 = v.toUpperCase().trim();
            if (rawSt2.indexOf("OPEN") !== -1 || rawSt2.indexOf("MỞ") !== -1 || rawSt2.indexOf("MO") !== -1 || rawSt2.indexOf("BẬT") !== -1) {
              statusVal = "OPEN";
            } else {
              statusVal = "CLOSED";
            }
          } else if (k.indexOf("thời gian") !== -1 || k.indexOf("duration") !== -1 || k.indexOf("phút") !== -1) {
            durVal = parseInt(v) || durationDefault;
          } else if (k.indexOf("lời dặn") !== -1 || k.indexOf("thông báo") !== -1 || k.indexOf("announcement") !== -1) {
            annVal = v;
          } else if (k.indexOf("pin") !== -1 || k.indexOf("giáo viên") !== -1) {
            pinVal = v;
          }
        }
      }

    }
  }

  // Lấy 5 đề tự luận tương ứng
  var essayPrompts = getEssayPromptsFromSheet(topicId);

  return {
    status: "success",
    topicId: topicId,
    examStatus: statusVal,
    duration: durVal,
    announcement: annVal,
    examPasscode: codeVal,
    teacherPin: pinVal,
    masterPin: pinVal,
    teacherEmail: "duyban86@gmail.com",
    essayPrompts: essayPrompts,
    serverTime: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss")
  };
}

// -----------------------------------------------------------------------------------------
// 7. QUẢN LÝ TÀI KHOẢN GOOGLE SSO (Tab 'TaiKhoan')
// -----------------------------------------------------------------------------------------
function getOrCreateAccountsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TaiKhoan");
  if (!sheet) {
    sheet = ss.insertSheet("TaiKhoan");
    var headers = ["Thời gian liên kết", "Lớp", "STT", "Họ và tên", "Email Google", "Ghi chú"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold").setFontSize(10).setFontFamily("Arial");
    hRange.setBackground("#1e293b").setFontColor("#38bdf8");
    hRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 60);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 240);
    sheet.setColumnWidth(6, 200);
  }
  return sheet;
}

function findAccountByEmail(email) {
  if (!email) return null;
  var accSheet = getOrCreateAccountsSheet();
  var lastRow = accSheet.getLastRow();
  if (lastRow < 2) return null;

  var data = accSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var cleanEmail = email.toLowerCase().trim();

  for (var i = 0; i < data.length; i++) {
    var rowEmail = String(data[i][4] || "").toLowerCase().trim();
    if (rowEmail === cleanEmail) {
      return {
        timestamp: data[i][0],
        className: data[i][1],
        stt: data[i][2],
        studentName: data[i][3],
        email: data[i][4],
        note: data[i][5]
      };
    }
  }
  return null;
}

function bindGoogleAccount(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accSheet = getOrCreateAccountsSheet();
  
  var email = (p.googleEmail || p.email || "").toLowerCase().trim();
  var className = (p.studentClass || p.className || "").trim();
  var stt = (p.studentStt || p.stt || "").trim();
  var name = (p.studentName || "").trim();
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

  if (!email || !className || !name) {
    return { status: "error", message: "Thiếu thông tin liên kết tài khoản." };
  }

  var existing = findAccountByEmail(email);
  if (existing) {
    return {
      status: "success",
      message: "Tài khoản Google này đã được liên kết trước đó.",
      student: existing
    };
  }

  accSheet.appendRow([timestamp, className, stt, name, email, "Đã xác thực SSO"]);
  
  return {
    status: "success",
    message: "Liên kết tài khoản Google thành công!",
    student: {
      timestamp: timestamp,
      className: className,
      stt: stt,
      studentName: name,
      email: email
    }
  };
}

function getAllAccounts() {
  var accSheet = getOrCreateAccountsSheet();
  var lastRow = accSheet.getLastRow();
  if (lastRow < 2) return [];

  var data = accSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    list.push({
      timestamp: data[i][0],
      className: data[i][1],
      stt: data[i][2],
      studentName: data[i][3],
      email: data[i][4],
      note: data[i][5]
    });
  }
  return list;
}

// -----------------------------------------------------------------------------------------
// 8. BÁO LỖI / PHẢN HỒI (Tab 'PhanHoi')
// -----------------------------------------------------------------------------------------
function saveFeedback(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PhanHoi");
  if (!sheet) {
    sheet = ss.insertSheet("PhanHoi");
    var headers = ["Thời gian", "Lớp", "Họ tên", "Mã câu hỏi", "Nội dung phản hồi", "Trạng thái xử lý"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  var ts = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
  sheet.appendRow([
    ts,
    p.className || p.studentClass || "",
    p.studentName || "",
    p.questionId || "",
    p.message || p.feedback || "",
    "Chưa xử lý"
  ]);

  return { status: "success", message: "Đã ghi nhận phản hồi của học sinh." };
}

// -----------------------------------------------------------------------------------------
// 9. HÀM GỬI PHẢN HỒI JSON / JSONP CHUẨN
// -----------------------------------------------------------------------------------------
function sendResponse(data, callback) {
  var jsonString = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonString + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

function sendJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
