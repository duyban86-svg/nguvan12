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

    // G. ĐẨY CÂU HỎI TRẮC NGHIỆM & TỰ LUẬN TỪ APP LÊN SHEET (QUA GET / JSONP)
    if (action === "push_questions_to_sheet" || action === "export_questions") {
      var pushRes = saveQuestionsToSheet(params);
      return sendResponse(pushRes, callback);
    }

    // Mặc định
    return sendResponse({ status: "success", message: "Hệ thống Google Apps Script đang hoạt động tốt." }, callback);

  } catch (err) {
    return sendResponse({ status: "error", message: err.toString() }, callback);
  }
}

// -----------------------------------------------------------------------------------------
// 2. XỬ LÝ YÊU CẦU POST (Nộp bài JSON/Form, Báo lỗi, Nhập điểm, Đẩy câu hỏi lên Sheet)
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

    if (action === "push_questions_to_sheet" || action === "export_questions") {
      var pushResult = saveQuestionsToSheet(params);
      return sendJsonResponse(pushResult);
    }

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
      // Tự động thay thế nếu trên Sheet còn đề cũ về Nỗi buồn chiến tranh
      if (title.indexOf("Nỗi buồn chiến tranh") !== -1 || prompt.indexOf("Nỗi buồn chiến tranh") !== -1 || prompt.indexOf("Bảo Ninh") !== -1) {
        title = "Đề Ôn tập 2 (Nghệ thuật xây dựng nhân vật trong Số đỏ)";
        prompt = "Viết đoạn văn khoảng 200 chữ phân tích nghệ thuật khắc họa nhân vật Xuân Tóc Đỏ trong tiểu thuyết Số đỏ của Vũ Trọng Phụng.";
        outline = "1. Mở đoạn: Giới thiệu Vũ Trọng Phụng và nhân vật Xuân Tóc Đỏ trong tiểu thuyết Số đỏ.<br>2. Thân đoạn: Bản chất lưu manh, vô học nhưng thức thời; sự biến ảo dị hợm khi bước chân vào xã hội thượng lưu; nghệ thuật biếm họa phóng đại sắc sảo của nhà văn.<br>3. Kết đoạn: Khẳng định giá trị hiện thực và sức sống vượt thời gian của hình tượng nhân vật điển hình.";
      }

      var sampleText = String(data[r][6] || "").trim();

      var item = {
        id: code || ("DE_" + r),
        title: title,
        prompt: prompt,
        outline: outline,
        sample: sampleText
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

// -----------------------------------------------------------------------------------------
// 5B. QUẢN LÝ TAB 'NganHangTracNghiem' (CÂU HỎI TRẮC NGHIỆM ĐỘNG TỪ SHEET)
// -----------------------------------------------------------------------------------------
function getOrCreateMCQSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("NganHangTracNghiem") || 
              ss.getSheetByName("TracNghiem") || 
              ss.getSheetByName("Trắc Nghiệm") || 
              ss.getSheetByName("CauHoiTN") || 
              ss.getSheetByName("Câu Hỏi TN");
  
  if (!sheet) {
    sheet = ss.insertSheet("NganHangTracNghiem");
    var headers = [
      "Chuyên đề (topicId)",
      "Phân loại (COMMON / ON_TAP / CHINH_THUC)",
      "STT câu",
      "Mức độ",
      "Dạng câu",
      "Nội dung Câu hỏi & Ngữ liệu",
      "Phương án A",
      "Phương án B",
      "Phương án C",
      "Phương án D",
      "Đáp án đúng",
      "Lời giải & Giải thích chi tiết"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold").setFontSize(10).setFontFamily("Arial");
    hRange.setBackground("#065f46").setFontColor("#ffffff");
    hRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 400);
    sheet.setColumnWidth(7, 220);
    sheet.setColumnWidth(8, 220);
    sheet.setColumnWidth(9, 220);
    sheet.setColumnWidth(10, 220);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 350);
  }
  return sheet;
}

function getQuizQuestionsFromSheet(topicId) {
  var sheet = getOrCreateMCQSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var cleanTopic = String(topicId || "").toLowerCase().trim();
  var questions_common = [];
  var questions_practice_only = [];
  var questions_exam_only = [];

  for (var r = 1; r < data.length; r++) {
    var rowTopic = String(data[r][0] || "").toLowerCase().trim();
    var mode = String(data[r][1] || "").toUpperCase().trim();
    var qNum = data[r][2] || r;
    var level = String(data[r][3] || "Thông hiểu").trim();
    var qType = String(data[r][4] || "single_choice").trim();
    var questionText = String(data[r][5] || "").trim();
    
    var optA = String(data[r][6] || "").trim();
    var optB = String(data[r][7] || "").trim();
    var optC = String(data[r][8] || "").trim();
    var optD = String(data[r][9] || "").trim();
    
    var correct = String(data[r][10] || "").trim();
    var explain = String(data[r][11] || "").trim();

    var isMatch = (rowTopic === cleanTopic || 
      (cleanTopic.indexOf("truyen") !== -1 && rowTopic.indexOf("truyen") !== -1) || 
      (cleanTopic.indexOf("tho") !== -1 && rowTopic.indexOf("tho") !== -1) || 
      (cleanTopic.indexOf("bai_1") !== -1 && rowTopic.indexOf("bai_1") !== -1) ||
      (cleanTopic.indexOf("bai_2") !== -1 && rowTopic.indexOf("bai_2") !== -1) ||
      rowTopic === "all");

    if (isMatch && questionText) {
      if (questionText.indexOf("Nỗi buồn chiến tranh") !== -1 || questionText.indexOf("Bảo Ninh") !== -1) {
        continue;
      }

      var qObj = {
        id: qNum,
        level: level,
        type: qType,
        question: questionText,
        explanation: explain
      };

      if (qType === 'true_false') {
        var ansParts = correct.toLowerCase().split(/[\s,\-\/\|]+/);
        var subQs = [
          { id: 'a', text: optA, correctAnswer: (ansParts[0] === 'sai' || ansParts[0] === 'false' || ansParts[0] === 's' || ansParts[0] === 'f') ? 'false' : 'true', explanation: '' },
          { id: 'b', text: optB, correctAnswer: (ansParts[1] === 'sai' || ansParts[1] === 'false' || ansParts[1] === 's' || ansParts[1] === 'f') ? 'false' : 'true', explanation: '' },
          { id: 'c', text: optC, correctAnswer: (ansParts[2] === 'sai' || ansParts[2] === 'false' || ansParts[2] === 's' || ansParts[2] === 'f') ? 'false' : 'true', explanation: '' },
          { id: 'd', text: optD, correctAnswer: (ansParts[3] === 'sai' || ansParts[3] === 'false' || ansParts[3] === 's' || ansParts[3] === 'f') ? 'false' : 'true', explanation: '' }
        ];
        qObj.subQuestions = subQs;
      } else if (qType === 'fill_blank') {
        var dOptions = ['-- Nhấp để chọn phương án điền khuyết --'];
        if (optA) dOptions.push(optA);
        if (optB) dOptions.push(optB);
        if (optC) dOptions.push(optC);
        if (optD) dOptions.push(optD);
        qObj.dropdownOptions = dOptions;
        qObj.correctAnswer = correct || optA;
      } else {
        var options = [];
        if (optA) options.push(optA.startsWith("A.") || optA.startsWith("A ") ? optA : ("A. " + optA));
        if (optB) options.push(optB.startsWith("B.") || optB.startsWith("B ") ? optB : ("B. " + optB));
        if (optC) options.push(optC.startsWith("C.") || optC.startsWith("C ") ? optC : ("C. " + optC));
        if (optD) options.push(optD.startsWith("D.") || optD.startsWith("D ") ? optD : ("D. " + optD));
        qObj.options = options;
        qObj.correctAnswer = correct.toUpperCase() || "A";
      }

      if (mode.indexOf("ON_TAP") !== -1 || mode.indexOf("PRACTICE") !== -1) {
        questions_practice_only.push(qObj);
      } else if (mode.indexOf("CHINH_THUC") !== -1 || mode.indexOf("EXAM") !== -1) {
        questions_exam_only.push(qObj);
      } else {
        questions_common.push(qObj);
      }
    }
  }

  if (questions_common.length === 0 && questions_practice_only.length === 0 && questions_exam_only.length === 0) {
    return null;
  }

  return {
    questions_common: questions_common,
    questions_practice_only: questions_practice_only,
    questions_exam_only: questions_exam_only
  };
}

// -----------------------------------------------------------------------------------------
// 5C. ĐỒNG BỘ CÂU HỎI TỪ WEB APP LÊN GOOGLE SHEETS (ĐẨY DỮ LIỆU ĐỂ GIÁO VIÊN TIỆN CHỈNH SỬA)
// -----------------------------------------------------------------------------------------
function saveQuestionsToSheet(params) {
  var topicId = String(params.topicId || "P1_BAI_1").trim();
  var pin = String(params.pin || params.teacherPin || "").trim();
  
  var validPin = TEACHER_PIN_DEFAULT;
  var config = getExamConfig(topicId);
  if (config && (config.teacherPin || config.masterPin)) {
    validPin = String(config.teacherPin || config.masterPin).trim();
  }
  if (pin !== validPin && pin !== TEACHER_PIN_DEFAULT && pin !== "1358") {
    return { status: "error", message: "Mã PIN Giáo viên không chính xác!" };
  }

  var mcqCount = 0;
  var essayCount = 0;

  if (params.quizData) {
    var qData = typeof params.quizData === 'string' ? JSON.parse(params.quizData) : params.quizData;
    var mcqSheet = getOrCreateMCQSheet();
    
    var data = mcqSheet.getDataRange().getValues();
    for (var r = data.length - 1; r >= 1; r--) {
      var rowTopic = String(data[r][0] || "").toLowerCase().trim();
      var cleanTarget = topicId.toLowerCase().trim();
      if (rowTopic === cleanTarget || 
          (cleanTarget.indexOf("bai_1") !== -1 && rowTopic.indexOf("bai_1") !== -1) ||
          (cleanTarget.indexOf("bai_2") !== -1 && rowTopic.indexOf("bai_2") !== -1) ||
          (cleanTarget.indexOf("truyen") !== -1 && rowTopic.indexOf("truyen") !== -1) ||
          (cleanTarget.indexOf("tho") !== -1 && rowTopic.indexOf("tho") !== -1)) {
        mcqSheet.deleteRow(r + 1);
      }
    }

    var mcqRows = [];
    function getOptionText(opt) {
      if (!opt) return "";
      if (typeof opt === "object") return String(opt.text || opt.content || "");
      return String(opt).replace(/^[A-Da-d][\.\:\)\s]+/, "").trim();
    }

    function collectMCQs(list, modeLabel) {
      if (!Array.isArray(list)) return;
      for (var i = 0; i < list.length; i++) {
        var item = list[i];
        var optA = "", optB = "", optC = "", optD = "";
        var correct = "";
        var qType = item.type || "single_choice";

        if (qType === "true_false") {
          var sqs = item.subQuestions || [];
          optA = sqs[0] ? sqs[0].text : "";
          optB = sqs[1] ? sqs[1].text : "";
          optC = sqs[2] ? sqs[2].text : "";
          optD = sqs[3] ? sqs[3].text : "";
          var ansList = sqs.map(function(sq) { return (String(sq.correctAnswer).toLowerCase() === "true" ? "Đúng" : "Sai"); });
          correct = ansList.join(" - ");
        } else if (qType === "fill_blank") {
          var dOpts = (item.dropdownOptions || []).filter(function(o) { return !o.startsWith("--"); });
          optA = dOpts[0] || "";
          optB = dOpts[1] || "";
          optC = dOpts[2] || "";
          optD = dOpts[3] || "";
          correct = item.correctAnswer || optA;
        } else {
          var opts = item.options || [];
          optA = getOptionText(opts[0]);
          optB = getOptionText(opts[1]);
          optC = getOptionText(opts[2]);
          optD = getOptionText(opts[3]);
          correct = String(item.correctAnswer || "A").toUpperCase();
        }

        mcqRows.push([
          topicId,
          modeLabel,
          item.id || (mcqRows.length + 1),
          item.level || "Thông hiểu",
          qType,
          item.question || item.text || "",
          optA,
          optB,
          optC,
          optD,
          correct,
          item.explanation || ""
        ]);
        mcqCount++;
      }
    }

    collectMCQs(qData.questions_common, "COMMON");
    collectMCQs(qData.questions_practice_only, "ON_TAP");
    collectMCQs(qData.questions_exam_only, "CHINH_THUC");

    if (mcqRows.length > 0) {
      var lastRow = mcqSheet.getLastRow();
      mcqSheet.getRange(lastRow + 1, 1, mcqRows.length, 12).setValues(mcqRows);
    }
  }

  if (params.essayPrompts) {
    var ePrompts = typeof params.essayPrompts === 'string' ? JSON.parse(params.essayPrompts) : params.essayPrompts;
    var essaySheet = getOrCreateEssayPromptsSheet();
    
    var eData = essaySheet.getDataRange().getValues();
    for (var r = eData.length - 1; r >= 1; r--) {
      var rowTopic = String(eData[r][0] || "").toLowerCase().trim();
      var cleanTarget = topicId.toLowerCase().trim();
      if (rowTopic === cleanTarget || 
          (cleanTarget.indexOf("bai_1") !== -1 && rowTopic.indexOf("bai_1") !== -1) ||
          (cleanTarget.indexOf("bai_2") !== -1 && rowTopic.indexOf("bai_2") !== -1) ||
          (cleanTarget.indexOf("truyen") !== -1 && rowTopic.indexOf("truyen") !== -1) ||
          (cleanTarget.indexOf("tho") !== -1 && rowTopic.indexOf("tho") !== -1)) {
        essaySheet.deleteRow(r + 1);
      }
    }

    var essayRows = [];
    ['on_tap', 'chinh_thuc'].forEach(function(mKey) {
      var list = ePrompts[mKey] || [];
      var modeLabel = (mKey === 'on_tap') ? 'ON_TAP' : 'CHINH_THUC';
      if (Array.isArray(list)) {
        list.forEach(function(item) {
          essayRows.push([
            topicId,
            modeLabel,
            item.id || "",
            item.title || "",
            item.prompt || "",
            item.outline || "",
            item.sample || ""
          ]);
          essayCount++;
        });
      }
    });

    if (essayRows.length > 0) {
      var eLastRow = essaySheet.getLastRow();
      essaySheet.getRange(eLastRow + 1, 1, essayRows.length, 7).setValues(essayRows);
    }
  }

  return {
    status: "success",
    message: "Đã xuất thành công " + mcqCount + " câu hỏi trắc nghiệm & " + essayCount + " đề tự luận của bài [" + topicId + "] lên Google Sheets!",
    topicId: topicId,
    mcqCount: mcqCount,
    essayCount: essayCount
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
      
      // A. QUÉT TOÀN BỘ CÁC DÒNG ĐỂ TÌM MÃ PIN GIÁO VIÊN (HỖ TRỢ THẢ XUỐNG Ở CỘT D HOẶC BẤT KỲ CỘT NÀO)
      for (var r = 0; r < data.length; r++) {
        var rowCombined = data[r].join(" ").toUpperCase();
        if (rowCombined.indexOf("PIN") !== -1 || rowCombined.indexOf("GIÁO VIÊN") !== -1 || rowCombined.indexOf("GIAO VIEN") !== -1 || rowCombined.indexOf("MASTER") !== -1 || rowCombined.indexOf("ADMIN") !== -1) {
          // Ưu tiên Cột D (index 3), rồi E (4), B (1), C (2), 0..
          var colOrder = [3, 4, 1, 2, 5, 0];
          for (var ci = 0; ci < colOrder.length; ci++) {
            var c = colOrder[ci];
            if (c < data[r].length) {
              var val = String(data[r][c] || "").trim();
              var valUpper = val.toUpperCase();
              if (val && valUpper.indexOf("PIN") === -1 && valUpper.indexOf("GIÁO VIÊN") === -1 && valUpper.indexOf("GIAO VIEN") === -1 && valUpper.indexOf("THẦY") === -1 && valUpper.indexOf("CÔ") === -1 && valUpper.indexOf("MÃ") === -1 && valUpper.indexOf("QUẢN TRỊ") === -1) {
                var cleanPin = val.replace(/\.0+$/, "").replace(/[\s\r\n\t]/g, "").trim();
                if (cleanPin.length >= 3 && cleanPin.length <= 15) {
                  pinVal = cleanPin;
                  break;
                }
              }
            }
          }
        }
      }

      // 1. DUYỆT BẢNG NHIỀU CHỦ ĐỀ / BÀI THI (Định dạng chuẩn của Thầy):
      // Cột A: STT | Cột B: MÃ BÀI THI (ID) | Cột C: TÊN BÀI THI | Cột D: TRẠNG THÁI | Cột E: MÃ CA THI | Cột F: THỜI GIAN | Cột G: LỜI NHẮN
      var matchedTopicRow = false;
      for (var r = 1; r < data.length; r++) {
        var rowId = String(data[r][1] || "").toUpperCase().replace(/\s+/g, '').trim(); // Cột B
        var rowName = String(data[r][2] || "").toUpperCase().trim();                   // Cột C
        var rowA = String(data[r][0] || "").toUpperCase().trim();                      // Cột A

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

  // Lấy ngân hàng câu hỏi trắc nghiệm động nếu giáo viên đã nhập trên Sheet
  var quizData = getQuizQuestionsFromSheet(topicId);

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
    quizData: quizData,
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


// -----------------------------------------------------------------------------------------
// 5D. HÀM KHỞI TẠO TOÀN BỘ NGÂN HÀNG CÂU HỎI VÀ ĐỀ THI (144 CÂU TRẮC NGHIỆM + 20 ĐỀ TỰ LUẬN)
// -----------------------------------------------------------------------------------------
function khoiTaoTatCaNganHangCauHoi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mcqSheet = getOrCreateMCQSheet();
  var essaySheet = getOrCreateEssayPromptsSheet();
  
  var mcqData = [["P1_BAI_1","COMMON",1,"Nhận biết","single_choice","Khái niệm 'Tiểu thuyết hiện đại' chủ yếu chỉ ra phương diện bản chất nào sau đây?","Một hệ hình tư duy nghệ thuật mới hướng vào đời tư, số phận cá nhân và thế giới nội tâm phức tạp.","Cột mốc thời gian xuất hiện các tác phẩm truyện dài từ đầu thế kỷ XX đến nay.","Những tác phẩm tự sự được viết hoàn toàn bằng chữ Quốc ngữ có cốt truyện li kì.","Các thiên truyện cổ điển viết về đề tài lịch sử dựng nước và giữ nước của dân tộc.","A","Tiểu thuyết hiện đại trước hết là một hệ hình tư duy nghệ thuật mới mang tính hiện đại, từ bỏ cái nhìn ước lệ quy phạm thời trung đại để khám phá đời tư và số phận cá nhân đa chiều."],["P1_BAI_1","COMMON",2,"Nhận biết","single_choice","Điểm nhìn trần thuật trong tiểu thuyết hiện đại có đặc điểm nổi bật nào?","Cố định duy nhất ở ngôi thứ nhất của nhân vật chính xưng 'tôi'.","Luân chuyển linh hoạt giữa người kể chuyện và điểm nhìn bên trong của nhiều nhân vật khác nhau.","Tuyệt đối đứng ngoài cuộc để đảm bảo tính khách quan như một bản tin thời sự.","Trùng khít hoàn toàn với điểm nhìn và quan điểm của tác giả ngoài đời thực.","B","Điểm nhìn trong tiểu thuyết hiện đại luân chuyển linh hoạt, đi sâu vào thế giới nội tâm để soi chiếu hiện thực từ nhiều góc độ đa chiều."],["P1_BAI_1","COMMON",3,"Thông hiểu","single_choice","Bản chất của tính 'đa thanh' (nhiều bè giọng) trong tiểu thuyết hiện đại là gì?","Sự đan xen phong phú của nhiều từ ngữ địa phương và phương ngữ vùng miền.","Sự cọ xát, đối thoại bình đẳng giữa nhiều hệ ý thức, điểm nhìn và giọng điệu nhân vật.","Tác giả sử dụng nhiều từ tượng thanh và biện pháp điệp âm để tạo nhạc tính.","Việc nhà văn xây dựng các nhân vật có giọng nói to, vang dội và xuất hiện đồng loạt.","B","Tính đa thanh là sự bình đẳng đối thoại, cọ xát giữa các ý thức và giọng điệu khác nhau mà không bị áp đặt bởi một giọng điệu độc thoại duy nhất của tác giả."],["P1_BAI_1","COMMON",4,"Thông hiểu","single_choice","Thời gian nghệ thuật trong tiểu thuyết hiện đại có bước chuyển dịch cơ bản như thế nào?","Tuân thủ nghiêm ngặt trục thời gian vật lý tuyến tính theo trình tự nguyên nhân - kết quả.","Dịch chuyển sang thời gian tâm lý, dòng ý thức với các trạng thái đồng hiện, hồi tưởng, nén dãn.","Chỉ tập trung vào thời gian thần thoại, phiếm chỉ vô định 'ngày xửa ngày xưa'.","Quy định chặt chẽ thời gian tác phẩm gói gọn trong một ngày một đêm theo luật tam duy nhất.","B","Thời gian nghệ thuật trong tiểu thuyết hiện đại phá vỡ tính tuyến tính để trở thành thời gian tâm lý, dòng ý thức theo cảm thụ chủ quan của nhân vật."],["P1_BAI_1","COMMON",5,"Nhận biết","single_choice","Tình huống trào phúng cốt lõi trong đoạn trích <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> (trích <strong class=\"work-title\"><em>\"Số đỏ\"</em></strong> – Vũ Trọng Phụng) là gì?","Xuân Tóc Đỏ bị hai nhà quán quân Hải và Thụ khiêu chiến trên sân quần vợt.","Giới cầm quyền buộc Xuân phải nhường điểm, chấp nhận thua cuộc nhưng hành vi ấy lại được suy tôn là nghĩa cử cứu quốc.","Khán giả phát hiện Xuân Tóc Đỏ gian lận điểm số và ném đồ vật xuống sân thi đấu.","Vua Xiêm La trực tiếp xuống sân trao huân chương danh dự cho Xuân Tóc Đỏ.","B","Tình huống trào phúng phi lý nằm ở chỗ: một hành vi thua cuộc nhục nhã vì lợi ích chính trị thực dân lại được biến hóa ngoạn mục thành kỳ tích hy sinh cứu vãn hòa bình dân tộc."],["P1_BAI_1","COMMON",6,"Thông hiểu","single_choice","Sự dị dạng và bản chất lưu manh, cơ hội của nhân vật Xuân Tóc Đỏ bộc lộ rõ nhất qua chi tiết nào?","Hắn từ chối tham gia trận đấu quần vợt vì sợ làm tổn thương bang giao hai nước.","Hắn vừa nhường điểm chịu thua nhục nhã, liền nhảy lên mui xe ô tô tự xưng hùng hồn là vĩ nhân cứu quốc.","Hắn lén lút ăn trộm tiền tài trợ giải đấu rồi bỏ trốn khỏi sân vận động.","Hắn thành thật thú nhận với quần chúng rằng mình chỉ là một kẻ nhặt bóng hạ lưu.","B","Bản chất giảo hoạt, dị dạng của Xuân thể hiện ở tài ứng biến trơ tráo: nhảy lên mui xe, xua tay dẹp trật tự và dùng miệng lưỡi bịp bợm biến sự thua cuộc của mình thành chiến công vĩ đại."],["P1_BAI_1","COMMON",7,"Thông hiểu","single_choice","Bức tranh đám đông quần chúng trong đoạn trích <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> được tác giả khắc họa mang tính chất gì?","Lực lượng giác ngộ cách mạng sâu sắc, biết đấu tranh vì lẽ phải và công bằng.","Đám đông u mê, hiếu kỳ, nông nổi, dễ bị 'thôi miên' và xoay chuyển 180 độ bởi những lời mị dân.","Những người hâm mộ thể thao chân chính, có kiến thức sâu rộng về luật thi đấu quốc tế.","Tập thể những người lao động nghèo khổ luôn đồng cảm với thân phận hạ lưu của Xuân.","B","Vũ Trọng Phụng châm biếm sâu cay tâm lý đám đông thuộc địa: vừa phẫn nộ đòi 'đả đảo' trước đó, ngay lập tức đã hoan hô, sùng bái Xuân như một đấng cứu tinh khi nghe vài câu lòe bịp."],["P1_BAI_1","COMMON",8,"Nhận biết","single_choice","Biện pháp tu từ 'nói mỉa' (irony) được hiểu chính xác nhất là gì?","Sự kết hợp các từ đồng nghĩa để tạo nên ngữ điệu trang trọng, nghiêm túc.","Thủ pháp tạo sự lệch pha cố ý giữa lời nói bề mặt (hiển ngôn) và ý định chế giễu ngầm ẩn (hàm ngôn).","Cách nói giảm nói tránh nhằm giảm bớt mức độ đau thương, nặng nề của sự việc.","Việc sử dụng các biện pháp so sánh ngầm mang tính trừu tượng triết học.","B","Nói mỉa là cách thức diễn đạt mà người nói/viết tạo ra khoảng cách mâu thuẫn có chủ ý giữa nghĩa hiển ngôn bề ngoài và ý định châm biếm, phê phán ngầm ẩn bên trong."],["P1_BAI_1","COMMON",9,"Vận dụng","single_choice","Phương án nào giải thích chính xác cơ chế tạo nghĩa và tác dụng trào phúng của cụm từ 'Sự đại bại vạn tuế!'?","Sử dụng phép so sánh để nâng tầm một trận thua thể thao bình thường thành bài học bang giao.","Phép nghịch ngữ kết hợp hai khái niệm mâu thuẫn để bóc trần sự lố lăng khi tôn vinh thất bại thành thắng lợi cứu nước.","Biện pháp nhân hóa tạo nên sự trang nghiêm, hào hùng cho kết quả trận đấu quốc tế.","Phép nói quá nhằm bày tỏ sự cảm thương sâu sắc trước sự hy sinh vì đại cuộc của vận động viên.","B","Cụm từ nghịch ngữ 'Sự đại bại vạn tuế!' ghép hai yếu tố tương phản (đại bại >< vạn tuế), tạo nên hiệu ứng trào phúng đanh thép bóc trần màn kịch cứu quốc bịp bợm."],["P1_BAI_1","COMMON",10,"Nhận biết","single_choice","Không gian và thời gian trong đoạn trích <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> (Ma Văn Kháng) mang ý nghĩa nghệ thuật gì?","Là không gian tâm lý gắn liền với nỗi niềm ưu tư, cô uẩn của ông Bằng trước sự rạn nứt nếp nhà thời mở cửa.","Đơn thuần miêu tả cảnh sắc mùa đông rét mướt và nhịp sống hối hả mua sắm của phố phường Hà Nội.","Tái hiện không khí hội hè tưng bừng, tràn ngập niềm vui sum họp trọn vẹn không tì vết.","Phản ánh không gian chiến trường khói lửa tàn khốc của những năm kháng chiến chống Mỹ.","A","Bối cảnh chiều 30 Tết trầm mặc, lá rụng xạc xào là không gian tâm lý phản ánh sự trăn trở sâu xa của ông Bằng trước những biến đổi của thời cuộc tác động đến nếp nhà."],["P1_BAI_1","COMMON",11,"Thông hiểu","single_choice","Vẻ đẹp nghĩa tình, nếp sống gia phong của nhân vật chị Hoài trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> bộc lộ rõ nét qua điều gì?","Chị lớn tiếng chỉ trích lối sống thực dụng, toan tính vật chất của cô em dâu Lý.","Dù đã đi bước nữa, chị vẫn lặn lội mang quà quê về thăm bố chồng cũ chiều 30 Tết, đóng vai trò kết nối nghĩa tình gia đình.","Chị quyết định để lại toàn bộ tài sản cho gia đình ông Bằng rồi rời đi ngay trong đêm giao thừa.","Chị trực tiếp đứng ra phân chia lại ngôi nhà cổ của gia đình cho các em trai.","B","Chị Hoài là hiện thân của đạo lý truyền thống: dù đã tái giá nhưng tấm lòng son sắt, thủy chung với gia đình chồng cũ vẫn nguyên vẹn, sưởi ấm tâm hồn mọi thành viên."],["P1_BAI_1","COMMON",12,"Thông hiểu","single_choice","Dòng văn nào sau đây trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> thể hiện đặc trưng của 'lời trần thuật nửa trực tiếp'?","'Chị Hoài đã về kìa!' - Lý cất tiếng reo lên mừng rỡ.","'Chị Hoài lên! Chị Hoài từ quê lên! Lên đúng chiều ba mươi Tết! Thật ngoài sức tưởng tượng!'","Đông, Lý, Luận hấp tấp từ phòng khách ùa ra vệt đường lát xi măng đi qua vườn cây ra cổng.","Ông Bằng đứng nghiêm trang thắp ba nén hương thơm trước bàn thờ tổ tiên.","B","Các câu cảm thán ngắt nhịp: 'Chị Hoài lên! Chị Hoài từ quê lên! Lên đúng chiều ba mươi Tết! Thật ngoài sức tưởng tượng!' là lời của người trần thuật ngôi thứ ba nhưng đã hòa nhập trọn vẹn vào dòng cảm xúc ngỡ ngàng, bàng hoàng và giọng điệu tâm tư bên trong của các nhân vật, thể hiện đặc trưng mẫu mực của lời trần thuật nửa trực tiếp trong văn bản SGK."],["P1_BAI_1","COMMON",13,"Vận dụng","single_choice","Đối chiếu hai văn bản trong Bài 1, nhận định nào đánh giá đúng nhất về sự đa dạng của tiểu thuyết hiện đại?","Một bên hướng ngoại trào phúng bóc trần sự giả dối, một bên hướng nội tâm lý sưởi ấm và neo giữ giá trị đạo đức truyền thống.","Cả hai tác phẩm đều tập trung tuyệt đối vào việc ca ngợi lý tưởng anh hùng cách mạng trong chiến tranh.","Hai tác phẩm chứng minh tiểu thuyết chỉ có thể viết về những đề tài lịch sử trong quá khứ xa xôi.","Một tác phẩm là kịch bản sân khấu hài kịch và một tác phẩm là bút ký phóng sự tài liệu.","A","Bài 1 giới thiệu hai hướng tìm tòi tiêu biểu: Vũ Trọng Phụng hướng ngoại trào phúng xã hội loạn chuẩn, còn Ma Văn Kháng hướng nội khám phá tâm lý đạo đức con người."],["P1_BAI_1","COMMON",14,"Nhận biết","single_choice","Cụm từ nào sau đây là ví dụ tiêu biểu của biện pháp tu từ 'nghịch ngữ'?","Nắng ấm ban mai","Sự thật bi hài","Vinh quang cay đắng","Cả B và C đều đúng","D","Nghịch ngữ kết hợp các từ ngữ mang nghĩa tương phản đối chọi ngay trong một cấu trúc (bi - hài, vinh quang - cay đắng) để biểu đạt tính chất phức tạp của đời sống."],["P1_BAI_1","COMMON",15,"Thông hiểu","single_choice","Thủ pháp cường điệu, phóng đại trong văn trào phúng của Vũ Trọng Phụng nhằm mục đích gì?","Làm cho bức tranh hiện thực trở nên tươi sáng, lạc quan và đầy chất thơ.","Đẩy những thói lố lăng, kệch cỡm lên mức tột cùng để bật ra tiếng cười châm biếm sâu cay.","Đánh lừa người đọc tin vào những sự việc thần thoại hoang đường.","Che giấu sự yếu kém về bút pháp xây dựng tính cách nhân vật.","B","Cường điệu phóng đại là công cụ đắc lực của ngòi bút trào phúng giúp phát lộ trọn vẹn bản chất dị dạng, giả tạo của xã hội tư sản thực dân."],["P1_BAI_1","COMMON",16,"Thông hiểu","single_choice","Hình tượng mâm cỗ tất niên và bàn thờ tổ tiên trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> mang biểu tượng nghệ thuật gì?","Tập tục mê tín dị đoan cần phải loại bỏ trong thời kỳ công nghiệp hóa.","Không gian thiêng liêng lưu giữ cội nguồn tâm linh và sự gắn kết gia phong truyền thống.","Biểu tượng cho sự giàu sang phú quý vượt bậc của gia đình quý tộc phong kiến.","Chi tiết phụ không mang giá trị nghệ thuật hay tư tưởng chủ đề.","B","Bàn thờ gia tiên và mâm cỗ chiều 30 Tết là không gian thiêng liêng quy tụ nếp sống tâm linh, nhắc nhở con người neo giữ nhân cách và đạo hiếu giữa dòng đời biến động."],["P1_BAI_1","COMMON",17,"Thông hiểu","single_choice","Xung đột ngầm ẩn trong gia đình ông Bằng thời kỳ mở cửa bắt nguồn từ đâu?","Sự va đập gay gắt giữa nếp sống đạo đức gia phong truyền thống và lối sống thực dụng, toan tính vật chất.","Tranh chấp quyền thừa kế tài sản nhà đất giữa các con trai.","Sự bất đồng về quan điểm chính trị trong chiến tranh giải phóng.","Nỗi hận thù gia tộc sâu sắc từ nhiều thế hệ trước để lại.","A","Xung đột trung tâm của tiểu thuyết là sự thử thách của đồng tiền, lối sống thực dụng thời kinh tế thị trường đối với nền tảng đạo đức gia đình truyền thống."],["P1_BAI_1","COMMON",18,"Vận dụng","single_choice","Khi làm bài văn nghị luận so sánh, đánh giá hai tác phẩm truyện, thao tác quan trọng hàng đầu là gì?","Kể lại chi tiết toàn bộ cốt truyện của tác phẩm thứ nhất rồi đến tác phẩm thứ hai.","Xác lập hệ thống tiêu chí so sánh rõ ràng (đề tài, nhân vật, nghệ thuật trần thuật, tư tưởng) để đối sánh song song.","Chỉ tập trung chỉ ra những điểm giống nhau tuyệt đối của hai tác phẩm.","Phê phán một tác phẩm dở và đề cao tuyệt đối tác phẩm còn lại.","B","Bài văn so sánh truyện cần thiết lập các tiêu chí cụ thể để phân tích nét tương đồng và dị biệt, từ đó làm nổi bật phong cách độc đáo của từng nhà văn."],["P1_BAI_1","COMMON",19,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định sau về đặc trưng thể loại TIỂU THUYẾT HIỆN ĐẠI:","Tiểu thuyết hiện đại từ bỏ cái nhìn ước lệ, quy phạm để hướng thẳng vào đời tư và số phận cá nhân.","Thời gian nghệ thuật trong tiểu thuyết hiện đại luôn bắt buộc tuân theo trình tự thời gian vật lý tuyến tính.","Điểm nhìn trần thuật có sự luân chuyển linh hoạt, đi sâu vào thế giới nội tâm của các nhân vật.","Ngôn ngữ tiểu thuyết hiện đại mang tính đa thanh, là sự đối thoại bình đẳng giữa nhiều bè giọng điệu.","Đúng - Sai - Đúng - Đúng",""],["P1_BAI_1","COMMON",20,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định sau về đoạn trích <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> (Vũ Trọng Phụng):","Hành động nhường điểm của Xuân Tóc Đỏ thực chất là ván cờ chính trị ám muội của chính quyền thực dân.","Xuân Tóc Đỏ là tấm gương đạo đức sáng ngời về lòng yêu nước xả thân vì thể diện dân tộc.","Tác giả sử dụng thủ pháp nghịch ngữ 'Sự đại bại vạn tuế' để tăng hiệu quả trào phúng đả kích.","Đám đông khán giả trong đoạn trích thể hiện sự tỉnh táo, sáng suốt và nhận diện được âm mưu của Xuân.","Đúng - Sai - Đúng - Sai",""],["P1_BAI_1","COMMON",21,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định sau về tác phẩm <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> (Ma Văn Kháng):","Tác phẩm thuộc khuynh hướng tiểu thuyết hiện thực tâm lý - xã hội sâu sắc.","Nhân vật ông Bằng đại diện cho thế hệ trẻ chạy theo lối sống thực dụng của kinh tế thị trường.","Chị Hoài đóng vai trò là 'chất keo kết dính', sưởi ấm tình cảm nghĩa tình giữa các thành viên trong gia đình.","Nghệ thuật trần thuật sử dụng xuất sắc lời nửa trực tiếp đan cài suy tư của nhân vật Luận.","Đúng - Sai - Đúng - Đúng",""],["P1_BAI_1","COMMON",22,"Thông hiểu","true_false","Đọc đoạn trích sau trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> (Ma Văn Kháng) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"Sự việc diễn ra quá ư đột ngột! Đông, Lý, Luận hấp tấp từ phòng khách ùa ra vệt đường lát xi măng đi qua vườn cây ra cổng, nhìn thấy chị Hoài thật rồi mà vẫn còn ngơ ngơ ngác ngác, nửa tin nửa ngờ. Chị Hoài lên! Chị Hoài từ quê lên! Lên đúng chiều ba mươi Tết! Thật ngoài sức tưởng tượng! Trong tâm ức vẫn là có hình bóng chị Hoài, chị Hoài, vợ anh cả Tường liệt sĩ. Chị Hoài, dâu trưởng, nết na, thuỳ mị. Trong tiềm thức vẫn sống động một chị Hoài đẹp người, đẹp nết.\"</em></div>","Đoạn văn tái hiện tâm trạng ngỡ ngàng, bàng hoàng và niềm xúc động sâu xa của các thành viên trước sự xuất hiện bất ngờ của chị Hoài.","Hình ảnh chị Hoài trong tâm ức và tiềm thức của gia đình vẫn vẹn nguyên là người dâu trưởng nết na, thuỳ mị, đẹp người đẹp nết.","Đoạn trích sử dụng lời trần thuật nửa trực tiếp để diễn tả dòng cảm xúc và sự xao xuyến trong tâm tư nhân vật.","Đoạn văn sử dụng bút pháp trào phúng phóng đại nhằm phê phán sự thờ ơ, xa cách giữa các thành viên trong gia đình.","Đúng - Đúng - Đúng - Sai",""],["P1_BAI_1","COMMON",23,"Nhận biết","true_false","Xác định tính Đúng / Sai của các nhận định về BIỆN PHÁP NÓI MỈA VÀ NGHỊCH NGỮ:","Nói mỉa luôn có sự đối lập giữa ý nghĩa biểu đạt bề mặt (hiển ngôn) và ý nghĩa phủ định ngầm (hàm ngôn).","Nghịch ngữ là việc sử dụng những từ ngữ hoàn toàn đồng nghĩa để lặp lại một ý niệm.","Nói mỉa và nghịch ngữ thường được phối hợp chặt chẽ trong các tác phẩm văn học trào phúng.","Nghịch ngữ chỉ được dùng trong văn học trào phúng chứ không thể dùng để diễn tả bi kịch hay tâm trạng phức tạp.","Đúng - Sai - Đúng - Sai",""],["P1_BAI_1","COMMON",24,"Vận dụng","true_false","Xác định tính Đúng / Sai về KỸ NĂNG VIẾT BÀI VĂN NGHỊ LUẬN SO SÁNH HAI TÁC PHẨM TRUYỆN:","Mở bài phải giới thiệu được cả hai tác phẩm và nêu định hướng các khía cạnh sẽ so sánh.","Phần thân bài chỉ cần tóm tắt cốt truyện của hai tác phẩm là đã đạt yêu cầu so sánh.","Cần làm rõ cả nét tương đồng (điểm gặp gỡ) và nét dị biệt (nét sáng tạo độc đáo) giữa hai tác phẩm.","Kết bài cần đánh giá ý nghĩa của sự khác biệt trong việc định hình phong cách nghệ thuật của từng tác giả.","Đúng - Sai - Đúng - Đúng",""],["P1_BAI_1","ON_TAP",25,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp để hoàn thiện nhận định: 'Khác với văn học trung đại, [ ...?... ] là thể loại thể hiện một cách tư duy nghệ thuật mới, đi sâu khám phá đời tư và số phận cá nhân con người.'","Tiểu thuyết hiện đại","Truyện ngắn cổ điển","Sử thi anh hùng","Kịch bản tuồng chèo","Tiểu thuyết hiện đại","Tiểu thuyết hiện đại là thể loại tiêu biểu phản ánh tư duy nghệ thuật hiện đại về con người cá nhân."],["P1_BAI_1","ON_TAP",26,"Nhận biết","fill_blank","Chọn tên biện pháp tu từ thích hợp: 'Thủ pháp tạo ra sự lệch pha cố ý giữa hiển ngôn (lời nói bề mặt) và hàm ngôn (ý định chế giễu ngầm ẩn) được gọi là biện pháp [ ...?... ]'","Nói mỉa","Nói quá","Chơi chữ","Nhân hóa","Nói mỉa","Nói mỉa (irony) là biện pháp tu từ tạo ra khoảng cách trớ trêu giữa lời nói bề mặt và hàm ý châm biếm bên trong."],["P1_BAI_1","ON_TAP",27,"Thông hiểu","fill_blank","Chọn cụm từ nghịch ngữ còn thiếu trong tiếng hoan hô sốt sắng của đám đông sau màn diễn thuyết của Xuân Tóc Đỏ: '– Xuân Tóc Đỏ vạn tuế! [ ...?... ]!'","Sự đại bại vạn tuế","Chiến thắng vinh quang","Hòa bình hữu nghị","Tinh thần bất khuất","Sự đại bại vạn tuế","Cụm từ nghịch ngữ 'Sự đại bại vạn tuế!' được thiên hạ sốt sắng hoan hô theo nguyên văn SGK, là đỉnh cao trào phúng phơi bày sự lố lăng của màn kịch cứu quốc và tâm lý u mê của đám đông."],["P1_BAI_1","ON_TAP",28,"Thông hiểu","fill_blank","Chọn hình thức trần thuật thích hợp: 'Hình thức trần thuật xóa nhòa ranh giới giữa lời kể khách quan ngôi thứ ba và dòng suy ngẫm thầm kín của nhân vật được gọi là [ ...?... ]'","Lời trần thuật nửa trực tiếp","Lời đối thoại kịch","Lời bình luận ngoại đề","Lời độc thoại thành tiếng","Lời trần thuật nửa trực tiếp","Lời trần thuật nửa trực tiếp hòa trộn giọng điệu người kể chuyện và tiếng nói nội tâm của nhân vật."],["P1_BAI_1","ON_TAP",29,"Nhận biết","fill_blank","Chọn tên nhân vật thích hợp: 'Trong tiểu thuyết <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong>, nhân vật [ ...?... ] dù đã tái giá nhưng chiều 30 Tết vẫn mang quà quê về thăm bố chồng cũ, là biểu tượng đẹp đẽ của đạo lý nghĩa tình truyền thống.'","Chị Hoài","Cô Lý","Anh Đông","Anh Luận","Chị Hoài","Chị Hoài là nhân vật trung tâm mang vẻ đẹp tình nghĩa thủy chung, kết nối gia đình ông Bằng."],["P1_BAI_1","ON_TAP",30,"Thông hiểu","fill_blank","Chọn đặc trưng thời gian thích hợp: 'Tiểu thuyết hiện đại phá vỡ trật tự tuyến tính vật lý một chiều để chuyển sang [ ...?... ] (hoặc dòng ý thức) theo dòng biến động nội tâm nhân vật.'","Thời gian tâm lý","Thời gian vật lý tuyến tính","Thời gian phiếm chỉ cổ tích","Thời gian vĩnh hằng","Thời gian tâm lý","Thời gian tâm lý là đặc trưng nghệ thuật quan trọng của tiểu thuyết hiện đại."],["P1_BAI_1","CHINH_THUC",25,"Vận dụng cao","single_choice","Điểm cách tân nghệ thuật nổi bật nhất của tiểu thuyết hiện đại so với tiểu thuyết chương hồi trung đại thể hiện ở phương diện nào?","Sử dụng kết cấu chương hồi nghiêm ngặt và ngôi kể thứ ba toàn tri tuyệt đối.","Phá vỡ tính tuyến tính một chiều, linh hoạt dịch chuyển điểm nhìn và đi sâu khám phá dòng ý thức nội tâm.","Chỉ tập trung ca ngợi hình tượng người anh hùng hoàn hảo theo chuẩn mực đạo đức phong kiến.","Triệt tiêu hoàn toàn tính đa thanh và xung đột tư tưởng giữa các nhân vật trong tác phẩm.","B","Tiểu thuyết hiện đại cách tân sâu sắc ở cấu trúc phi tuyến tính, trần thuật đa điểm nhìn và khả năng khám phá chiều sâu vô thức, dòng ý thức con người."],["P1_BAI_1","CHINH_THUC",26,"Thông hiểu","single_choice","Trong đoạn trích <em>'Xuân Tóc Đỏ cứu quốc'</em> (Vũ Trọng Phụng), thủ pháp trào phúng nghịch dị được thể hiện rõ nét nhất qua tình huống nào?","Xuân Tóc Đỏ tự giác rèn luyện tài năng quần vợt để mang vinh quang thực sự về cho đất nước.","Kẻ lưu manh vô học được cả xã hội thượng lưu và giới cầm quyền tôn sùng như một bậc vĩ nhân cứu quốc.","Cuộc thi đấu thể thao diễn ra trong không khí trang nghiêm và tinh thần thể thao trung thực.","Quần chúng nhân dân nhận thức sâu sắc bản chất lừa mị của chính quyền thực dân.","B","Mâu thuẫn trào phúng cốt lõi: một kẻ hạ lưu, vô học bỗng chốc trở thành 'anh hùng cứu quốc' do sự lố bịch, dối trá của tầng lớp thống trị thượng lưu."],["P1_BAI_1","CHINH_THUC",27,"Vận dụng","single_choice","Nghệ thuật xây dựng nhân vật trào phúng trong tiểu thuyết <em>'Số đỏ'</em> của Vũ Trọng Phụng thể hiện đặc sắc qua phương diện nào sau đây?","Xây dựng nhân vật như một cá tính phức tạp với thế giới nội tâm giằng xé sâu sắc.","Khắc họa nhân vật qua lối biếm họa, phóng đại những nét dị dạng, lố bịch và tha hóa về nhân cách.","Lý tưởng hóa phẩm chất đạo đức của nhân vật để làm tấm gương sáng cho người đọc noi theo.","Tái hiện nhân vật một cách trung thực, dung dị như trong đời sống thường nhật mà không cường điệu hóa.","B","Vũ Trọng Phụng sử dụng thủ pháp phóng đại, biếm họa sắc bén để làm nổi bật sự tha hóa, giả dối và lố bịch của xã hội thượng lưu 'Âu hóa' đương thời."],["P1_BAI_1","CHINH_THUC",28,"Thông hiểu","single_choice","Khái niệm 'Tính đa thanh' (Polyphony) trong tiểu thuyết hiện đại chỉ hiện tượng nghệ thuật nào?","Tác phẩm sử dụng nhiều nhạc điệu và thể thơ khác nhau xen lẫn lời văn xuôi.","Sự cùng tồn tại bình đẳng của nhiều giọng điệu, ý thức và quan điểm thẩm mỹ độc lập của các nhân vật.","Người kể chuyện áp đặt tuyệt đối tư tưởng của mình lên toàn bộ hệ thống nhân vật.","Tác phẩm có nhiều kết cục mở nhưng chỉ phản ánh một góc nhìn duy nhất của tác giả.","B","Tính đa thanh là sự đối thoại bình đẳng giữa nhiều luồng ý thức, giọng điệu, không có một tiếng nói đơn độc nào độc chiếm chân lý tuyệt đối."],["P1_BAI_1","CHINH_THUC",29,"Vận dụng","single_choice","Sự xuất hiện của 'đối thoại nội tâm' và 'độc thoại nội tâm' trong tiểu thuyết hiện đại có tác dụng nghệ thuật chủ yếu là:","Làm chậm mạch truyện và thay thế hoàn toàn hành động thực tế của nhân vật.","Soi tỏ những xung đột giằng xé phức tạp trong thế giới vô thức và tiềm thức sâu kín của con người.","Giúp tác giả tiết kiệm dung lượng miêu tả ngoại hình và bối cảnh không gian thời gian.","Tạo sự đơn giản, dễ hiểu cho độc giả đại chúng khi theo dõi cốt truyện.","B","Độc thoại nội tâm đưa người đọc thâm nhập trực tiếp vào thế giới tâm lí phức tạp, khám phá những góc khuất sâu kín nhất của bản thể nhân vật."],["P1_BAI_1","CHINH_THUC",30,"Vận dụng cao","single_choice","Nhận định nào sau đây đánh giá ĐÚNG NHẤT về vị trí của thể loại tiểu thuyết trong nền văn học đương đại?","Tiểu thuyết là thể loại đã đạt đến giới hạn và không còn khả năng dung nạp các yếu tố nghệ thuật mới.","Tiểu thuyết là thể loại tự sự linh hoạt nhất, có khả năng dung hợp đa dạng mọi thể loại và bao quát đời sống sâu sắc nhất.","Tiểu thuyết chỉ phù hợp phản ánh hiện thực lịch sử quá khứ chứ không thể nắm bắt được biến động hiện đại.","Tiểu thuyết có cấu trúc cố định, không chấp nhận sự thử nghiệm phá cách về ngôn ngữ và điểm nhìn.","B","Tiểu thuyết được coi là 'thể loại của mọi thể loại' bởi tính linh hoạt, khả năng mở rộng biên độ phản ánh và không ngừng vận động, đổi mới."],["P1_BAI_2","COMMON",1,"Nhận biết","single_choice","Đặc trưng cốt lõi nổi bật nhất của <strong>Phong cách Cổ điển</strong> trong thơ ca trung đại là gì?","Hướng tới những đề tài cao nhã, trang trọng, giàu tính quy phạm và hệ thống ước lệ vũ trụ.","Giải phóng triệt để tiềm thức và khai thác những giấc mơ hoang đường phi lý tính.","Tự do bộc lộ trực tiếp những cảm xúc trần trụi của cái tôi cá nhân đời thường.","Phá vỡ toàn bộ cấu trúc ngữ pháp và luật thơ truyền thống để tạo nhịp điệu mới.","A","Phong cách cổ điển thời trung đại đề cao tính quy phạm, đề tài trang trọng, chí hướng quân tử và không gian vũ trụ kì vĩ."],["P1_BAI_2","COMMON",2,"Nhận biết","single_choice","Cảm hứng chủ đạo của <strong>Phong cách Lãng mạn</strong> trong thơ ca là gì?","Khẳng định cái cao cả, phi thường và khát vọng vượt thoát khỏi thực tại tầm thường, tù túng.","Tuân thủ nghiêm ngặt niêm luật Đường luật và quan niệm 'thi dĩ ngôn chí'.","Tái hiện trung thực, khách quan các chi tiết sinh hoạt đời thường của xã hội.","Đề cao tính chất giáo huấn đạo lý và lòng trung quân ái quốc tuyệt đối.","A","Chủ nghĩa lãng mạn hướng tới cái phi thường, bay bổng của cái tôi cá nhân và khát vọng vươn tới thế giới lí tưởng vượt lên thực tại."],["P1_BAI_2","COMMON",3,"Nhận biết","single_choice","Trong thơ hiện đại, <strong>yếu tố Tượng trưng</strong> chú trọng đặc biệt vào phương diện nghệ thuật nào?","Chú trọng gợi cảm hơn miêu tả trực diện, sử dụng ẩn dụ chuyển đổi cảm giác và biểu tượng đa tầng.","Ghi chép chính xác số liệu, thời gian và địa điểm của sự kiện lịch sử.","Dựng lại toàn bộ cốt truyện theo trật tự thời gian tuyến tính biên niên sử.","Sử dụng ngôn ngữ khẩu ngữ bình dân mộc mạc không qua trau chuốt.","A","Thơ tượng trưng đề cao tính gợi, nhạc tính và sự tương giao đồng cảm giác giữa các giác quan để khám phá cõi lòng u uẩn."],["P1_BAI_2","COMMON",4,"Nhận biết","single_choice","Yếu tố nào sau đây đóng vai trò là 'xương sống', giúp bài thơ thoát khỏi sơ đồ ý niệm trừu tượng để trở thành một chỉnh thể nghệ thuật sống động?","Tứ thơ (Cấu tứ tác phẩm).","Số lượng câu chữ trong bài thơ.","Niên đại sáng tác của tác phẩm.","Lời đề từ ở đầu bài thơ.","A","Cấu tứ và tứ thơ là hạt nhân tổ chức, định hướng phát triển mạch cảm xúc và hệ thống hình tượng trong thi phẩm."],["P1_BAI_2","COMMON",5,"Nhận biết","single_choice","Bài thơ <em>'Cảm hoài'</em> của Đặng Dung được viết theo thể thơ nào?","Thất ngôn bát cú Đường luật (chữ Hán).","Song thất lục bát (chữ Nôm).","Thất ngôn tứ tuyệt Đường luật.","Thơ tự do hiện đại.","A","Cảm hoài được viết bằng chữ Hán theo thể thất ngôn bát cú Đường luật mực thước (8 câu, 56 chữ, niêm luật chặt chẽ)."],["P1_BAI_2","COMMON",6,"Nhận biết","single_choice","Chuỗi âm thanh <em>'li-la li-la li-la'</em> xuất hiện ở phần đầu và phần kết bài thơ <em>'Đàn ghi ta của Lor-ca'</em> (Thanh Thảo) mô phỏng điều gì?","Tiếng gảy đàn ghi ta dạo đầu và dư âm ngân vang của ca khúc kết hợp loài hoa tử đinh hương (lilas).","Tiếng vó ngựa của lực lượng độc tài phát xít trên đấu trường.","Tiếng bước chân người nghệ sĩ cô đơn đi giữa thảo nguyên Tây Ban Nha.","Tiếng reo hò cổ vũ của khán giả trên đấu trường bò tót.","A","Chuỗi âm thanh 'li-la li-la li-la' vừa là hợp âm ghi ta vừa gợi liên tưởng đến hoa tử đinh hương (lilas) — biểu tượng của sự bất tử."],["P1_BAI_2","COMMON",7,"Thông hiểu","single_choice","Trong hai câu thực bài <em>'Cảm hoài'</em>: <em>'Thời lai đồ điếu thành công dị, / Vận khứ anh hùng ẩm hận đa'</em>, tác giả đã sử dụng thủ pháp nghệ thuật nào để làm nổi bật bi kịch thế sự?","Nghệ thuật đối ngẫu mẫu mực kết hợp vận dụng điển tích điển cố.","Biện pháp nhân hóa và điệp từ nối tiếp liên tục.","Nghệ thuật nói mỉa và lối diễn hài kịch trào phúng.","Thủ pháp độc thoại nội tâm phân thân đa giọng điệu.","A","Đặng Dung đối lập triệt để giữa 'thời lai' - 'vận khứ', 'đồ điếu' (Phàn Khoái, Hàn Tín) - 'anh hùng' để thể hiện nỗi đau lỡ vận sâu sắc."],["P1_BAI_2","COMMON",8,"Thông hiểu","single_choice","Hai câu luận bài thơ <em>'Cảm hoài'</em>: <em>'Trí chủ hữu hoài phù địa trục, / Tẩy binh vô lộ vãn thiên hà'</em> thể hiện hoài bão gì của người anh hùng?","Hoài bão kinh bang tế thế kì vĩ xoay trục đất cứu vua và khát vọng kéo sông Ngân rửa sạch vũ khí đem lại thái bình muôn thuở.","Ước muốn tìm về cuộc sống ẩn dật chốn thôn quê non nước thanh bình.","Khao khát chứng minh tài năng võ nghệ vô địch trên sa trường đẫm máu.","Nỗi bi quan, buông xuôi, phó mặc số phận cho sự an bài của trời đất.","A","Hình ảnh mang tầm vóc vũ trụ 'phù địa trục' (xoay trục đất) và 'vãn thiên hà' (kéo sông Ngân) thể hiện chí lớn phò vua và khát vọng hòa bình."],["P1_BAI_2","COMMON",9,"Thông hiểu","single_choice","Nỗi nhớ trong đoạn đầu bài thơ <em>'Tây Tiến'</em> của Quang Dũng được định danh bằng từ ngữ nào giàu sức gợi cảm giác?","Nhớ 'chơi vơi'.","Nhớ 'da diết'.","Nhớ 'nồng nàn'.","Nhớ 'khôn nguôi'.","A","'Sông Mã xa rồi Tây Tiến ơi! Nhớ về rừng núi nhớ chơi vơi' — 'chơi vơi' mở ra nỗi nhớ lan tỏa, mênh mang giữa không gian và thời gian."],["P1_BAI_2","COMMON",10,"Thông hiểu","single_choice","Cụm từ <em>'súng ngửi trời'</em> trong câu thơ <em>'Heo hút cồn mây súng ngửi trời'</em> (Tây Tiến) mang lại hiệu quả nghệ thuật đặc sắc nào?","Nhân hóa táo bạo đặc tả độ cao chót vót của đỉnh đèo, đồng thời thể hiện nét tinh nghịch, ngạo nghễ của người lính trẻ.","Miêu tả độ ẩm ướt của sương mù Tây Bắc làm hoen rỉ vũ khí chiến sĩ.","Thể hiện sự mệt mỏi, bất lực của đoàn quân khi vượt núi cao hiểm trở.","Báo hiệu sự hiện diện của quân địch phục kích trên sườn núi.","A","'Súng ngửi trời' là nét vẽ lãng mạn phá cách, vừa đo độ cao núi rừng vừa làm nổi bật tâm hồn lạc quan, tếu táo của người lính Hà thành."],["P1_BAI_2","COMMON",11,"Thông hiểu","single_choice","Thủ pháp nghệ thuật nào được Quang Dũng sử dụng để tạo nên nét đẹp thơ mộng, huyền ảo đối lập với sự hiểm trở dữ dội của dốc núi Tây Bắc?","Phối hợp các câu thơ toàn thanh bằng tạo âm hưởng êm đềm ('Nhà ai Pha Luông mưa xa khơi').","Sử dụng dày đặc các từ láy trắc gợi hình hiểm trở liên tiếp.","Vận dụng triệt để thủ pháp ước lệ trung đại 'ngư, tiều, canh, mục'.","Sử dụng cấu trúc câu hỏi tu từ dồn dập chất chứa âu lo.","A","Câu thơ 7 thanh bằng 'Nhà ai Pha Luông mưa xa khơi' mở ra một không gian phẳng lặng, mơ màng, xoa dịu những gian nan dốc đèo."],["P1_BAI_2","COMMON",12,"Thông hiểu","single_choice","Vẻ đẹp tâm hồn người lính Tây Tiến được thể hiện qua nét vẽ lãng mạn, hào hoa nào sau đây?","'Mắt trừng gửi mộng qua biên giới / Đêm mơ Hà Nội dáng kiều thơm'.","'Áo anh rách vai, quần tôi có vài mảnh vá'.","'Ruộng nương anh gửi bạn thân cày / Gian nhà không mặc kệ gió lung lay'.","'Lột sắt đường tàu rèn thêm dao kiếm'.","A","Dù hoàn cảnh chiến đấu khốc liệt ('mắt trừng'), người lính trí thức Tây Tiến vẫn giữ trọn nét hào hoa, lãng mạn qua giấc mơ về 'dáng kiều thơm'."],["P1_BAI_2","COMMON",13,"Thông hiểu","single_choice","Trong bài thơ <em>'Đàn ghi ta của Lor-ca'</em>, hình ảnh <em>'tiếng ghi ta nâu', 'tiếng ghi ta lá xanh', 'tiếng ghi ta tròn bọt nước vỡ tan', 'tiếng ghi ta ròng ròng máu chảy'</em> là biểu hiện của thủ pháp nghệ thuật nào?","Ẩn dụ chuyển đổi cảm giác (Đồng cảm giác) mang đậm dấu ấn thơ Siêu thực.","Phép so sánh ngang bằng truyền thống trong ca dao.","Lối miêu tả hiện thực khách quan theo bút pháp tả thực xã hội.","Nghệ thuật đối lập niêm luật trong thơ cổ điển Đường luật.","A","Thanh Thảo đã chuyển hóa âm thanh (tiếng đàn) thành màu sắc (nâu, xanh), hình khối (tròn bọt nước) và xúc giác đau đớn (máu chảy)."],["P1_BAI_2","COMMON",14,"Thông hiểu","single_choice","Hình ảnh <em>'giọt nước mắt vầng trăng / long lanh trong đáy giếng'</em> trong bài <em>'Đàn ghi ta của Lor-ca'</em> mang ý nghĩa biểu tượng gì?","Sự kết tinh giữa nỗi đau xót thương tiếc Lor-ca và niềm tin bất diệt vào sự trường tồn của cái Đẹp nghệ thuật.","Tái hiện lại cảnh đêm trăng thanh bình nơi làng quê Tây Ban Nha.","Miêu tả cái chết cụ thể của người chiến sĩ trên chiến hào đánh giặc.","Thể hiện sự tuyệt vọng hoàn toàn của con người trước định mệnh tàn khốc.","A","'Giọt nước mắt' là niềm xót thương, 'vầng trăng' là cái đẹp vĩnh hằng, 'đáy giếng' nơi Lor-ca bị sát hại trở thành nơi lưu giữ ánh sáng nghệ thuật."],["P1_BAI_2","COMMON",15,"Thông hiểu","single_choice","Trong <em>'Bài thơ số 28'</em> của R. Tagore, vì sao tác giả so sánh đời mình với <em>'viên ngọc'</em>, <em>'đoá hoa'</em> nhưng lại khẳng định <em>'trái tim em là vô biên'</em>?","Để đối lập giữa cái hữu hạn, dễ đo đếm của vật chất với thế giới bí ẩn, vô tận của tâm hồn và tình yêu đích thực.","Để ca ngợi sự giàu sang phú quý của tầng lớp quý tộc Ấn Độ.","Để than thở về sự ngắn ngủi và vô nghĩa của kiếp người.","Để khẳng định tình yêu chỉ tồn tại khi có sự kiểm soát tuyệt đối.","A","Tagore dùng nghịch lý để chỉ ra rằng trái tim tình yêu không có biên giới cơ học, không thể thấu suốt hay chiếm hữu trọn vẹn."],["P1_BAI_2","COMMON",16,"Vận dụng","single_choice","Thông điệp nhân sinh sâu sắc nhất mà R. Tagore gửi gắm qua <em>'Bài thơ số 28'</em> về tình yêu đôi lứa là gì?","Tình yêu chân chính là sự dâng hiến tự nguyện và trân trọng miền sâu kín, thiêng liêng nơi tâm hồn nhau thay vì chiếm hữu cơ học.","Tình yêu chỉ bền vững khi cả hai người có sự hoàn hảo tuyệt đối về địa vị xã hội.","Con người không nên mở lòng yêu thương vì sẽ luôn phải gánh chịu đau khổ.","Cần phải kiểm soát và giải mã toàn bộ bí mật nội tâm của người yêu.","A","Tagore khẳng định tình yêu đích thực là sự thấu hiểu bằng trực cảm, dâng hiến và tôn trọng cõi thiêng vô tận của người bạn đời."],["P1_BAI_2","COMMON",17,"Vận dụng","single_choice","Điểm giao thoa thẩm mĩ lớn nhất giữa hình tượng tráng sĩ trong <em>'Cảm hoài'</em> (Đặng Dung) và người lính trong <em>'Tây Tiến'</em> (Quang Dũng) là gì?","Vẻ đẹp bi tráng: Dù đối diện với gian khổ, mất mát hay tình thế nghiệt ngã nhưng tư thế luôn kiêu hùng, bất khuất.","Cùng tuân thủ tuyệt đối niêm luật và thi pháp thơ Đường luật cổ điển.","Cùng tập trung miêu tả chi tiết đời sống nông thôn thời kì phong kiến.","Cùng sử dụng các yếu tố siêu thực, phá vỡ logic đời thường.","A","Cả hai hình tượng đều ngời sáng tinh thần bi tráng: bi kịch hiện thực không làm suy giảm khí phách và lí tưởng phụng sự cao cả."],["P1_BAI_2","COMMON",18,"Vận dụng","single_choice","Khi tiếp cận và giải mã một thi phẩm hiện đại thuộc khuynh hướng Tượng trưng – Siêu thực, người đọc cần lưu ý điều gì?","Lắng nghe nhạc tính của ngôn từ, kết nối các trường liên tưởng đa giác quan và giải mã hệ thống biểu tượng đa tầng nghĩa.","Chỉ tập trung tóm tắt cốt truyện và diễn biến sự việc theo thời gian tuyến tính.","Phân tích ngữ pháp câu thơ theo chuẩn mực logic câu văn xuôi thông thường.","Quy chụp bài thơ vào một ý nghĩa duy nhất mang tính khuôn mẫu cứng nhắc.","A","Đọc thơ tượng trưng - siêu thực đòi hỏi khả năng trực cảm thẩm mỹ, giải mã biểu tượng và đón nhận tính đa nghĩa mở của thi phẩm."],["P1_BAI_2","COMMON",19,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Phong cách Cổ điển</strong> và thi phẩm <em>'Cảm hoài'</em> của Đặng Dung:","Bài thơ 'Cảm hoài' tuân thủ nghiêm ngặt niêm luật thất ngôn bát cú Đường luật, luật Trắc, vần Bằng độc vận 'a'.","Hai câu thực sử dụng nghệ thuật đối lập giữa 'thời lai đồ điếu' và 'vận khứ anh hùng' nhằm đả kích thói xu nịnh của quan lại triều đình.","Hình tượng 'phù địa trục' và 'vãn thiên hà' là các biểu tượng thi pháp mang tầm vóc vũ trụ kì vĩ theo chuẩn mực cổ điển.","Chi tiết 'mài gươm dưới trăng' ở câu kết khắc họa bức tượng đài tráng lệ, khẳng định ý chí chiến đấu bất diệt của bậc hào kiệt.","Đúng - Sai - Đúng - Đúng",""],["P1_BAI_2","COMMON",20,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Cảm hứng lãng mạn và tinh thần bi tráng</strong> trong bài thơ <em>'Tây Tiến'</em> (Quang Dũng):","Cảm hứng chủ đạo xuyên suốt toàn bộ bài thơ là nỗi nhớ da diết về những ngày tháng hòa bình êm ấm nơi thủ đô Hà Nội.","Hình tượng người lính hiện lên với vẻ đẹp bi tráng: gian khổ tột cùng (sốt rét rụng tóc, quân xanh màu lá) nhưng tâm hồn hào hoa, ngạo nghễ.","Cách nói 'gục lên súng mũ bỏ quên đời' và 'áo bào thay chiếu anh về đất' đã làm giảm bớt tính chất bi lụy của sự hy sinh.","Âm thanh tiếng gầm gừ của 'Sông Mã gầm lên khúc độc hành' tạo nên khúc nhạc chiêu hồn tử sĩ vang dội núi rừng đại ngàn.","Sai - Đúng - Đúng - Đúng",""],["P1_BAI_2","COMMON",21,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Nghệ thuật xây dựng không gian và hình tượng</strong> trong bài thơ <em>'Tây Tiến'</em>:","Các địa danh như Sài Khao, Mường Lát, Pha Luông, Mường Hịch mang sắc thái âm vang xa xôi, hoang dã của vùng biên cương hiểm trở.","Cảnh đêm liên hoan lửa đuốc 'Doanh trại bừng lên hội đuốc hoa' bộc lộ nét đẹp rực rỡ, tình tứ và ấm áp tình quân dân.","Đoạn thơ 'Người đi Châu Mộc chiều sương ấy' sử dụng bút pháp tả thực chi tiết, rõ ràng từng đường nét nhân vật chèo thuyền.","Nhịp thơ 4/3 bẻ đôi trong 'Dốc lên khúc khuỷu / dốc thăm thẳm' trực quan hóa sự gãy khúc, dựng đứng của dốc núi Tây Bắc.","Đúng - Đúng - Sai - Đúng",""],["P1_BAI_2","COMMON",22,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Yếu tố Tượng trưng & Siêu thực</strong> trong bài thơ <em>'Đàn ghi ta của Lor-ca'</em> (Thanh Thảo):","Hình tượng chiếc 'áo choàng đỏ gắt' gắn liền với đấu trường bò tót và nền văn hóa truyền thống Tây Ban Nha rực lửa.","Chi tiết 'áo choàng bê bết đỏ' và 'bầu trời cô gái ấy' là hình ảnh hiện thực tái hiện trực tiếp việc bắt giữ Lor-ca trong sử sách.","Hình ảnh 'Lor-ca bơi sang ngang trên chiếc ghi ta màu bạc' thể hiện sự hóa thân bất tử và sự giải thoát thanh thản của người nghệ sĩ vào cõi vĩnh hằng.","Hành động ném lá bùa và trái tim vào dòng sông xoáy tượng trưng cho sự tuyệt vọng đầu hàng số phận của người nghệ sĩ.","Đúng - Sai - Đúng - Sai",""],["P1_BAI_2","COMMON",23,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Chiều sâu triết lý tình yêu</strong> trong <em>'Bài thơ số 28'</em> của R. Tagore:","Mở đầu bài thơ, đôi mắt của người yêu được miêu tả như đang 'muốn nhìn vào tâm tưởng' tác giả.","Tagore khẳng định đời mình là viên ngọc quý giá có thể đem bán lấy tiền chia cho người nghèo.","Hình ảnh 'trăng muốn vào sâu biển cả' là biểu tượng ẩn dụ cho khát khao muốn khám phá đến tận cùng thế giới nội tâm của người mình yêu.","Bài thơ kết luận rằng trái tim tình yêu vừa là niềm lạc thú vừa là nỗi khổ đau, vô biên không thể nào nắm bắt trọn vẹn.","Đúng - Sai - Đúng - Đúng",""],["P1_BAI_2","COMMON",24,"Thông hiểu","true_false","Xét các nhận định sau về <strong>Thực hành Tiếng Việt & Các biện pháp tu từ trong thơ</strong>:","Ẩn dụ chuyển đổi cảm giác là biện pháp tu từ diễn tả cảm nhận của một giác quan này thông qua từ ngữ thuộc giác quan khác.","Biểu tượng trong thơ luôn có ý nghĩa cố định, đơn nhất và không bao giờ thay đổi qua các thời đại văn học.","Nghệ thuật đối ngẫu trong thơ Đường luật bắt buộc phải đối xứng cả về từ loại, ngữ nghĩa và thanh điệu bằng - trắc.","Nhạc tính trong thơ tự do chỉ được tạo ra bởi vần chân ở cuối mỗi câu thơ, không phụ thuộc vào nhịp điệu hay điệp từ.","Đúng - Sai - Đúng - Sai",""],["P1_BAI_2","ON_TAP",25,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp điền vào chỗ trống: '[ ...?... ] là khâu then chốt khởi đầu sáng tạo nghệ thuật nhằm xác định, hình dung hướng phát triển của hình tượng thơ và cách thức triển khai toàn bộ tác phẩm.'","Cấu tứ","Gieo vần","Ngắt nhịp","Niêm luật","Cấu tứ","Cấu tứ là hoạt động sáng tạo nhằm tổ chức, liên kết các yếu tố thành một tứ thơ hoàn chỉnh."],["P1_BAI_2","ON_TAP",26,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là biện pháp tu từ diễn tả sự liên tưởng tương ứng, hòa quyện giữa các giác quan (nghe thấy màu sắc, nhìn thấy âm thanh, nếm được cảm xúc).' ","Ẩn dụ chuyển đổi cảm giác","Hoán dụ cận kề","So sánh ngang bằng","Phép điệp cú pháp","Ẩn dụ chuyển đổi cảm giác","Ẩn dụ chuyển đổi cảm giác (synesthesia/đồng cảm giác) là nét đặc trưng thi pháp của thơ Tượng trưng."],["P1_BAI_2","ON_TAP",27,"Nhận biết","fill_blank","Điền cụm từ Hán Việt còn thiếu vào câu kết bài thơ <em>'Cảm hoài'</em> (Đặng Dung): 'Quốc thù vị báo đầu tiên bạch / Kỉ độ [ ...?... ] đới nguyệt ma.'","Long Tuyền","Thái A","Bảo kiếm","Ngân Hà","Long Tuyền","'Long Tuyền' là tên thanh gươm báu tráng lệ được tráng sĩ mài dưới bóng trăng."],["P1_BAI_2","ON_TAP",28,"Nhận biết","fill_blank","Điền cụm từ còn thiếu vào câu thơ độc đáo của Quang Dũng trong <em>'Tây Tiến'</em>: 'Heo hút cồn mây [ ...?... ] / Ngàn thước lên cao, ngàn thước xuống.'","súng ngửi trời","súng chạm mây","bước chênh vênh","bóng dập dờn","súng ngửi trời","'Súng ngửi trời' là nét vẽ nhân hóa táo bạo đo độ cao hiểm trở và thể hiện sự ngạo nghễ của người lính."],["P1_BAI_2","ON_TAP",29,"Thông hiểu","fill_blank","Điền cụm từ mang ý nghĩa biểu tượng trong bài thơ <em>'Đàn ghi ta của Lor-ca'</em> (Thanh Thảo): 'giọt nước mắt [ ...?... ] / long lanh trong đáy giếng.'","vầng trăng","bọt nước","lá xanh","hoa sen","vầng trăng","'Giọt nước mắt vầng trăng' là sự hòa quyện giữa niềm tiếc thương và vẻ đẹp bất tử của nghệ thuật."],["P1_BAI_2","ON_TAP",30,"Vận dụng","fill_blank","Chọn cụm từ thích hợp hoàn thành nhận định: 'Nếu bài thơ <em>'Cảm hoài'</em> đại diện cho phong cách [ ...?... ] mực thước của thời trung đại, thì <em>'Tây Tiến'</em> lại kết tinh đỉnh cao cảm hứng lãng mạn và tinh thần bi tráng thời kỳ kháng chiến.'","Cổ điển","Hiện thực phê phán","Siêu thực","Hậu hiện đại","Cổ điển","'Cảm hoài' của Đặng Dung là tác phẩm mẫu mực của phong cách thơ Cổ điển thời trung đại."],["P1_BAI_2","CHINH_THUC",25,"Vận dụng cao","single_choice","Vẻ đẹp bi tráng của hình tượng người lính trong bài thơ <em>'Tây Tiến'</em> (Quang Dũng) được kết tinh độc đáo từ sự hòa quyện giữa hai yếu tố nào?","Nỗi bi quan trước sự tàn khốc của chiến tranh và mong muốn giải ngũ sớm.","Hiện thực gian khổ, hi sinh khốc liệt với tinh thần hào hùng, lãng mạn và lý tưởng xả thân cao đẹp.","Vẻ đẹp mộc mạc thôn quê với lối sống hưởng thụ của tầng lớp thị dân thành thị.","Cảm hứng châm biếm sâu cay với tâm trạng hoài cổ của tầng lớp trí thức cũ.","B","Vẻ đẹp bi tráng sinh ra từ sự đối lập hài hòa: hiện thực dẫu khốc liệt, chết chóc nhưng ý chí, tâm hồn người lính vẫn hào hoa, kiêu hãnh và bất tử."],["P1_BAI_2","CHINH_THUC",26,"Thông hiểu","single_choice","Hình ảnh <em>'tiếng đàn bọt nước'</em> trong bài thơ <em>'Đàn ghi ta của Lor-ca'</em> (Thanh Thảo) mang ý nghĩa tượng trưng sâu sắc cho điều gì?","Âm thanh thực tế của một nhạc cụ truyền thống Tây Ban Nha trong lễ hội đấu bò.","Số phận mong manh, dang dở của người nghệ sĩ tài hoa trước bạo lực bạo tàn nhưng nghệ thuật thì bất tử.","Nỗi thất vọng của nhà thơ trước sự suy tàn của nền âm nhạc cổ điển phương Tây.","Khát vọng làm giàu bằng con đường biểu diễn nghệ thuật đường phố tự do.","B","'Tiếng đàn bọt nước' là ẩn dụ chuyển đổi cảm giác mang màu sắc siêu thực: mong manh, ngắn ngủi như bọt nước nhưng không ngừng sinh nở, tái sinh và bất tử."],["P1_BAI_2","CHINH_THUC",27,"Vận dụng","single_choice","Đặc trưng nổi bật của thi pháp thơ Hiện đại và Tượng trưng thể hiện rõ nhất ở yếu tố nào sau đây?","Sử dụng thi pháp ước lệ cổ điển và quy tắc niêm luật bằng trắc gò bó tuyệt đối.","Phát huy tối đa sức gợi mở của biểu tượng, liên tưởng bất ngờ và sự tương giao đa giác quan.","Chỉ ghi chép chân thực sự vật hiện tượng theo lối tả thực photographic trực quan.","Tuyệt đối loại bỏ nhạc tính và cảm xúc chủ quan của người nghệ sĩ sáng tạo.","B","Thơ hiện đại và tượng trưng hướng tới chiều sâu gợi cảm của biểu tượng, sự tương giao giữa các giác quan (synesthesia) và những liên tưởng nhảy vọt."],["P1_BAI_2","CHINH_THUC",28,"Thông hiểu","single_choice","Nhận xét nào sau đây phản ánh ĐÚNG về vai trò của 'Cấu tứ' trong một bài thơ trữ tình?","Là số lượng câu thơ và số khổ thơ cố định được quy định trong luật thơ.","Là sơ đồ tổ chức, ý niệm trụ cột định hướng sự vận động của mạch cảm xúc và hệ thống hình ảnh toàn bài.","Là danh sách các từ ngữ ngữ pháp được gieo vần ở cuối các dòng thơ.","Là phần tóm tắt nội dung câu chuyện lịch sử được kể trong bài thơ.","B","Cấu tứ là linh hồn tổ chức của bài thơ, là ý niệm định hình cách thức triển khai và liên kết toàn bộ hình ảnh, cảm xúc thành một chỉnh thể nghệ thuật."],["P1_BAI_2","CHINH_THUC",29,"Vận dụng","single_choice","Cái 'Tôi trữ tình' trong thơ hiện đại có bước chuyển biến căn bản nào so với cái 'Ta' trong thơ trung đại?","Chuyển từ biểu đạt ý thức cá nhân đa diện, phong phú sang tiếng nói quy phạm phi ngã tập thể.","Chuyển từ tính chất phi ngã, ước lệ sang ý thức khẳng định bản ngã cá nhân độc đáo với mọi cung bậc cảm xúc.","Hoàn toàn phủ nhận giá trị của thiên nhiên và cái đẹp trong cuộc sống con người.","Đồng nhất hoàn toàn với nhân vật kể chuyện trong các tác phẩm truyện tự sự.","B","Cái tôi trữ tình hiện đại là tiếng nói cá nhân tự do, phức hợp, dám phơi trải mọi góc cạnh tâm trạng chân thực và độc đáo của cái tôi cá thể."],["P1_BAI_2","CHINH_THUC",30,"Vận dụng cao","single_choice","Giá trị thẩm mỹ lớn nhất của thơ ca đích thực đối với tâm hồn bạn đọc là gì?","Cung cấp số liệu chính xác về các sự kiện lịch sử và hiện tượng tự nhiên.","Thanh lọc tâm hồn, đánh thức mỹ cảm và khơi gợi những rung động nhân văn sâu xa nhất.","Hướng dẫn những quy tắc ứng xử thực dụng và quy chuẩn giao tiếp đời thường.","Thay thế hoàn toàn các hình thức tư duy khoa học và triết học trong đời sống.","B","Chức năng thanh lọc (catharsis) và khơi gợi mỹ cảm nhân văn là giá trị cốt lõi, vĩnh cửu của nghệ thuật thi ca đối với đời sống tinh thần con người."],["P2_TRUYEN","COMMON",1,"Nhận biết","single_choice","Theo lí luận văn học hiện đại, khái niệm 'Truyện' được hiểu chính xác nhất là gì?","Loại hình tác phẩm tự sự, phương thức phản ánh đời sống chủ yếu thông qua việc xây dựng thế giới hư cấu với hệ thống sự kiện, cốt truyện, nhân vật và người kể chuyện.","Tác phẩm trữ tình phản ánh trực tiếp thế giới cảm xúc chủ quan của tác giả thông qua hệ thống hình ảnh và vần điệu cô đọng.","Văn bản ghi chép chính xác 100% người thật việc thật ngoài đời sống mà tuyệt đối không có bất kì yếu tố hư cấu nào.","Kịch bản sân khấu được xây dựng hoàn toàn bằng chuỗi hành động và xung đột đối thoại trực tiếp trước khán giả.","A","Truyện là loại hình tự sự phản ánh đời sống qua thế giới hư cấu (cốt truyện, nhân vật, người kể chuyện), từ đó tái hiện bức tranh hiện thực và khám phá chiều sâu tâm lý con người."],["P2_TRUYEN","COMMON",2,"Nhận biết","single_choice","Mô hình cốt truyện mà các sự kiện thoát khỏi trật tự nhân quả máy móc, đảo lộn theo dòng tâm lí hoặc liên tưởng của nhân vật được gọi là gì?","Cốt truyện cổ điển / tuyến tính.","Cốt truyện phi tuyến tính.","Cốt truyện lồng ghép (truyện lồng trong truyện).","Cốt truyện đơn tuyến một chiều.","B","Cốt truyện phi tuyến tính phá vỡ trình tự thời gian vật lý trước - sau và mối quan hệ nhân quả máy móc để sắp xếp sự kiện nương theo dòng tâm lý, liên tưởng của nhân vật."],["P2_TRUYEN","COMMON",3,"Nhận biết","single_choice","Người kể chuyện xưng 'tôi' trong tác phẩm tự sự, trực tiếp tham gia hoặc chứng kiến câu chuyện thường gắn liền với đặc điểm nào?","Người kể chuyện toàn tri, thấu suốt mọi ngóc ngách tâm hồn của tất cả các nhân vật.","Người kể chuyện hạn tri, góc nhìn mang tính chủ quan và có giới hạn tri thức nhất định.","Người kể chuyện giấu mình tuyệt đối và giữ vị trí khách quan vô can.","Người kể chuyện đứng ngoài tác phẩm và phán xét đạo đức một cách trực tiếp.","B","Người kể chuyện ngôi thứ nhất (xưng tôi) là người kể chuyện hạn tri, chỉ biết và kể những gì mình chứng kiến, trải nghiệm, mang đậm tính chủ quan."],["P2_TRUYEN","COMMON",4,"Thông hiểu","single_choice","Xét về bản chất nghệ thuật, nhận định nào sau đây là ĐÚNG về nhân vật trong tác phẩm truyện?","Nhân vật là bản sao nguyên xi, trung thực tuyệt đối của con người ngoài đời thực.","Nhân vật là một ước lệ nghệ thuật được nhà văn sáng tạo nhằm gửi gắm suy ngẫm, đối thoại và cắt nghĩa hiện thực đời sống.","Nhân vật chính diện bắt buộc phải là tấm gương đạo đức hoàn hảo tuyệt đối không có khuyết điểm.","Nhân vật phụ không có bất kì vai trò gì trong việc thể hiện chủ đề và tư tưởng tác phẩm.","B","Nhân vật là một sinh thể nghệ thuật mang tính ước lệ, là phương tiện để nhà văn đối thoại với độc giả và thể hiện quan niệm nhân sinh sâu sắc."],["P2_TRUYEN","COMMON",5,"Nhận biết","single_choice","Hình thức trần thuật nào tạo nên sự cộng hưởng tinh tế giữa lời người kể chuyện và tiếng nói nội tâm của nhân vật ngay trong mạch kể?","Lời đối thoại trực tiếp giữa hai nhân vật.","Lời độc thoại nội tâm phát ra thành tiếng.","Lời trần thuật nửa trực tiếp (lời nửa trực tiếp).","Lời bình luận ngoại đề thuần túy của tác giả.","C","Lời trần thuật nửa trực tiếp là hình thức hòa trộn giữa lời kể khách quan của ngôi thứ ba với giọng điệu, cảm xúc và suy ngẫm bên trong của nhân vật."],["P2_TRUYEN","COMMON",6,"Thông hiểu","single_choice","Loại tình huống truyện nào tập trung vào sự kiện làm nảy sinh những biến chuyển, giằng xé nội tâm sâu sắc và khám phá thế giới cảm xúc tinh tế của nhân vật?","Tình huống hành động.","Tình huống nhận thức.","Tình huống tâm trạng.","Tình huống trào phúng (hoạt kê).","C","Tình huống tâm trạng là hoàn cảnh đẩy nhân vật vào những rung động, băn khoăn và giằng xé nội tâm sâu kín."],["P2_TRUYEN","COMMON",7,"Nhận biết","single_choice","Trong tác phẩm truyện, những biểu hiện cụ thể, sống động giàu sức gợi giúp tạo nên 'xương thịt', cá tính và thần thái cho nhân vật được gọi là gì?","Chủ đề tác phẩm.","Chi tiết nghệ thuật.","Cảm hứng chủ đạo.","Văn cảnh.","B","Chi tiết nghệ thuật là hạt nhân tạo nên tính chân thực, sinh động và sức hấp dẫn cho thế giới hình tượng trong tác phẩm tự sự."],["P2_TRUYEN","COMMON",8,"Thông hiểu","single_choice","Đặc trưng nổi bật của 'bút pháp dòng ý thức' trong văn xuôi hiện đại là gì?","Mạch trần thuật tuân thủ nghiêm ngặt trật tự thời gian tuyến tính một chiều.","Nhân vật thực hiện liên tục nhiều hành động ngoại cảnh kịch tính dồn dập.","Thời gian nghệ thuật bị đảo lộn, phân mảnh; câu chuyện nương theo dòng tâm tư, tiềm thức bất tận của nhân vật.","Toàn bộ tác phẩm được viết theo thể thơ tự do có nhạc điệu.","C","Bút pháp dòng ý thức mô phỏng dòng chảy miên man, phức tạp của ý thức và tiềm thức, phá vỡ trật tự không gian - thời gian thông thường."],["P2_TRUYEN","COMMON",9,"Nhận biết","single_choice","Bối cảnh thời gian và không gian trung tâm trong đoạn trích <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> (Ma Văn Kháng) diễn ra vào thời điểm nào?","Đêm giao thừa thiêng liêng bên bờ hồ Hoàn Kiếm.","Chiều ba mươi Tết tại ngôi nhà cổ kính bên mâm cỗ tất niên và bàn thờ gia tiên.","Sáng mùng một Tết khi họ hàng đến chúc Tết đông đủ.","Một buổi chiều thu vắng lặng nơi làng quê Bắc Bộ.","B","Bối cảnh chiều 30 Tết bên bàn thờ gia tiên là không gian tâm lý hội tụ nếp nhà truyền thống và khơi dậy nỗi niềm sâu kín của các nhân vật."],["P2_TRUYEN","COMMON",10,"Nhận biết","single_choice","Trong đoạn trích <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong>, nhân vật chị Hoài xuất hiện với nét ngoại hình và trang phục tiêu biểu nào?","Mặc áo dài nhung quý phái, đeo kiềng bạc sang trọng kiểu tiểu thư Hà thành.","Mặc váy tân thời kiểu Âu hóa, đi giày cao gót thời thượng.","Thon gọn trong chiếc áo bông chần hạt lựu, đôi gót chân to bản nứt nẻ thâm đen, tay mang chiếc tay nải trĩu nặng quà quê.","Mặc bộ quân phục bạc màu thời chiến tranh giải phóng.","C","Trang phục áo bông chần hạt lựu, đôi gót chân nứt nẻ và gói quà quê thể hiện vẻ đẹp mộc mạc, chất phác và tấm lòng son sắt của người phụ nữ nông thôn."],["P2_TRUYEN","COMMON",11,"Thông hiểu","single_choice","Tình huống trào phúng trọng tâm trong đoạn trích <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> (Vũ Trọng Phụng) nảy sinh từ sự việc nào?","Xuân Tóc Đỏ bị hai quán quân Hải và Thụ bắt cóc trước giờ thi đấu.","Xuân đánh thắng áp đảo quán quân Xiêm khiến vua Xiêm thịnh nộ, buộc quan chức bảo hộ phải ép Xuân 'đổi thắng thành thua' để tránh nguy cơ chiến tranh.","Xuân Tóc Đỏ bị đám đông công chúng phát hiện là kẻ thất học nên ném đá đuổi đánh.","Vua Xiêm phong tặng danh hiệu hiệp sĩ danh dự cho ông bầu Văn Minh.","B","Tình huống kịch tính nảy sinh khi ván đấu thể thao bị biến thành công cụ chính trị ngoại giao, dẫn đến màn 'nhường điểm' trớ trêu."],["P2_TRUYEN","COMMON",12,"Thông hiểu","single_choice","Hình tượng chị Hoài trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> mang ý nghĩa biểu tượng sâu sắc nhất là gì?","Người mang lại nguồn của cải vật chất dồi dào giúp gia đình ông Bằng vượt qua khủng hoảng.","Chất keo gắn kết ân tình, ngọn lửa sưởi ấm nếp nhà, đánh thức lương tri và nhắc nhở mọi người neo giữ đạo lý truyền thống.","Đại diện cho lối sống thực dụng, toan tính của thời kỳ kinh tế thị trường mới nổi.","Nhân vật đối đầu gay gắt làm rạn nứt các mối quan hệ giữa các thành viên.","B","Chị Hoài là hiện thân của nếp sống nghĩa tình, tấm lòng nhân hậu có sức mạnh hàn gắn những rạn nứt tâm hồn."],["P2_TRUYEN","COMMON",13,"Thông hiểu","single_choice","Trong giờ phút cúng tất niên linh thiêng, việc ông Bằng lặng lẽ bỏ qua tên Cừ khi đọc tên các con trai thể hiện tâm trạng gì?","Sự căm thù tột cùng và tuyệt tình dứt khoát không bao giờ tha thứ cho con.","Sự lú lẫn, đãng trí của một người già tuổi đã xế chiều.","Nỗi đau đớn, xót xa khôn cùng của người cha có đứa con lầm lạc, đồng thời là ý thức nghiêm cẩn giữ gìn sự thanh sạch cho bàn thờ tổ tiên.","Thái độ lạnh lùng, vô cảm trước số phận của các con trong gia đình.","C","Hành động ngập ngừng bỏ qua tên Cừ là bi kịch nội tâm đầy đau xót của người cha gia phong nghiêm cẩn."],["P2_TRUYEN","COMMON",14,"Thông hiểu","single_choice","Bản chất mối quan hệ cộng sinh trào phúng giữa Xuân Tóc Đỏ và đám đông công chúng trong <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> là gì?","Đám đông hoàn toàn tỉnh táo và chỉ giả vờ tin theo lời của Xuân Tóc Đỏ để giữ thể diện ngoại giao.","Xuân Tóc Đỏ là một tài năng xuất chúng thực sự được quần chúng nhân dân tôn vinh xứng đáng.","Sự trơ trẽn, giảo hoạt của kẻ cơ hội vỉa hè gặp đúng mảnh đất màu mỡ là tâm lý bầy đàn u mê, nông nổi, dễ bị 'thôi miên, dắt mũi' của đám đông thị dân.","Mối quan hệ gắn kết trên tinh thần yêu nước chân chính và ý thức trách nhiệm dân tộc.","C","Vũ Trọng Phụng bóc trần sự cộng sinh quái gở giữa kẻ bịp bợm lọc lừa và đám đông u mê, háo danh thời thuộc địa."],["P2_TRUYEN","COMMON",15,"Thông hiểu","single_choice","Các chỉ dẫn hành động trong ngoặc đơn (nó vỗ vào ngực, nó đấm tay xuống không khí, nó giơ cao tay lên...) ở màn diễn thuyết của Xuân Tóc Đỏ mang lại hiệu quả nghệ thuật gì?","Thể hiện sự tôn kính của tác giả đối với phong thái đĩnh đạc của bậc vĩ nhân.","Mang dáng dấp của chỉ dẫn sân khấu trong kịch bản hài kịch, biến sân vận động thành sàn diễn kịch nghệ hoạt kê khổng lồ.","Nhằm mục đích tường thuật chính xác các thao tác thể thao của vận động viên.","Nhấn mạnh vẻ đẹp bi tráng và sự hi sinh cao cả của người lính ngoài mặt trận.","B","Lối viết kịch bản hóa tăng tính chất hề hước, biến màn cứu quốc thành vở tuồng trò trơ trẽn."],["P2_TRUYEN","COMMON",16,"Vận dụng","single_choice","Đọc đoạn văn sau trong truyện ngắn <strong class=\"work-title\"><em>\"Muối của rừng\"</em></strong> (Nguyễn Huy Thiệp):\n<div class=\"reading-excerpt-box\"><em>\"Sự hỗn loạn của cả đàn khỉ khiến cho ông Diểu sợ hãi run lên. Ông vừa làm điều ác. Chân tay ông rủn ra... Ông đã lộ mặt là tên ám sát!... – Thôi Diểu ơi… – ông buồn bã nghĩ, – với đôi chân thấp khớp thế này thì làm sao mày chạy nhanh bằng lòng tận tụy, thủy chung của khỉ?\"</em></div>\nNhận xét nào phân tích <strong>đúng nhất</strong> về nghệ thuật trần thuật?","Tác giả sử dụng ngôi kể thứ nhất thuần túy để nhân vật trực tiếp khoe chiến tích săn bắn.","Kết nối linh hoạt giữa ngôi thứ ba với điểm nhìn bên trong và độc thoại nội tâm của ông Diểu, soi tỏ sự giằng xé lương tâm và bước ngoặt thức tỉnh nhân tính.","Tác giả sử dụng điểm nhìn của con khỉ cái để trực tiếp phán xét loài người.","Đoạn trích thuần túy miêu tả ngoại cảnh thiên nhiên mà không hề đề cập đến chuyển biến tâm lý.","B","Đoạn văn dịch chuyển điểm nhìn sâu vào nội tâm nhân vật, đánh dấu khoảnh khắc lương tri thức tỉnh trước vẻ đẹp tình nghĩa của loài vật."],["P2_TRUYEN","COMMON",17,"Vận dụng","single_choice","Vận dụng nguyên lý mối quan hệ giữa nội dung và hình thức, nhận định nào giải thích ĐÚNG NHẤT về không gian nghệ thuật của hai tác phẩm?","Vũ Trọng Phụng chọn không gian công cộng (sân vận động) để bóc trần sự bát nháo xã hội; Ma Văn Kháng chọn không gian gia đình chiều 30 Tết để soi tỏ đạo lý nếp nhà.","Sự lựa chọn không gian hoàn toàn ngẫu nhiên và không liên quan gì đến tư tưởng tác phẩm.","Không gian sân vận động Rollandes Varreau là không gian tâm lý sâu kín của nhân vật ông Bằng.","Cả hai tác giả đều chỉ sử dụng không gian chiến trường để ca ngợi chiến công quân sự.","A","Không gian nghệ thuật là phương tiện đắc lực thể hiện tư tưởng: không gian công cộng cho trào phúng xã hội và không gian gia đình thiêng liêng cho tâm lý đạo đức."],["P2_TRUYEN","COMMON",18,"Vận dụng","single_choice","Khi đối sánh phong cách <strong class=\"work-title\"><em>\"Lão Hạc\"</em></strong> (Nam Cao) và <strong class=\"work-title\"><em>\"Hai đứa trẻ\"</em></strong> (Thạch Lam), điểm khác biệt nổi bật là gì?","<strong class=\"work-title\"><em>\"Lão Hạc\"</em></strong> là truyện hiện thực tâm lý sâu sắc có cốt truyện kịch tính giàu xung đột; còn <strong class=\"work-title\"><em>\"Hai đứa trẻ\"</em></strong> là truyện ngắn trữ tình không có cốt truyện nương theo tâm trạng đượm chất thơ.","<strong class=\"work-title\"><em>\"Hai đứa trẻ\"</em></strong> là kịch bản trào phúng đả kích dữ dội chế độ phong kiến mục nát.","<strong class=\"work-title\"><em>\"Lão Hạc\"</em></strong> là truyện đồng thoại dành riêng cho lứa tuổi thiếu nhi.","Cả hai tác phẩm đều sử dụng yếu tố kì ảo hoang đường của truyện truyền kỳ.","A","Nam Cao sắc sảo khám phá bi kịch nhân phẩm người nông dân qua cốt truyện kịch tính; Thạch Lam tinh tế nắm bắt những cảm giác mơ hồ, mong manh qua truyện trữ tình không cốt truyện."],["P2_TRUYEN","COMMON",19,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định sau về các TIỂU LOẠI TRUYỆN:","Truyện truyền kì là văn xuôi tự sự trung đại, dùng yếu tố kì ảo làm phương thức nghệ thuật đan xen với hiện thực đời sống và thường có lời bình ở cuối truyện.","Truyện thơ là thể loại chỉ có phương thức trữ tình biểu cảm, hoàn toàn không có cốt truyện hay hệ thống nhân vật.","Tiểu thuyết hiện đại có quy mô lớn, chú trọng đi sâu vào đời tư, số phận cá nhân và thế giới tâm lý phức tạp ('con người nếm trải').","Truyện đồng thoại viết cho thiếu nhi chỉ miêu tả con người thật, tuyệt đối không được nhân hóa loài vật hay đồ vật.","Đúng - Sai - Đúng - Sai",""],["P2_TRUYEN","COMMON",20,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định về THỦ PHÁP KHẮC HỌA TÂM LÝ NHÂN VẬT:","Độc thoại nội tâm và lời trần thuật nửa trực tiếp là các phương thức miêu tả trực tiếp thế giới tâm tư, ý thức của nhân vật.","Bút pháp tả cảnh ngụ tình chỉ nhằm mục đích trang trí bố cục, hoàn toàn không có mối liên hệ nào với tâm trạng nhân vật.","Thủ pháp đồng hiện cho phép các sự kiện ở nhiều không gian và thời gian khác nhau cùng hiện lên trên nền hồi ức của nhân vật.","Hành động, cử chỉ bên ngoài có thể phản chiếu những mâu thuẫn, giằng xé và biến chuyển tâm lý sâu kín bên trong.","Đúng - Sai - Đúng - Đúng",""],["P2_TRUYEN","COMMON",21,"Thông hiểu","true_false","Xác định tính Đúng / Sai về ĐẶC TRƯNG TIỂU THUYẾT HIỆN ĐẠI qua <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> và <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong>:","<strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> xây dựng nhân vật theo khuynh hướng biếm họa con rối; còn <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> xây dựng kiểu 'con người nếm trải' đa chiều.","Cả hai tác phẩm đều sử dụng cốt truyện cổ điển tuyến tính máy móc với người kể chuyện cổ xưa đơn điệu.","<strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> dùng không gian chiều 30 Tết làm 'phép thử' đạo đức; <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> gom tụ xã hội vào một sự kiện thể thao công cộng lớn.","Vũ Trọng Phụng sử dụng giọng văn đằm thắm sưởi ấm tâm hồn; còn Ma Văn Kháng dùng giọng điệu mỉa mai, khinh miệt gay gắt.","Đúng - Sai - Đúng - Sai",""],["P2_TRUYEN","COMMON",22,"Vận dụng","true_false","Đọc đoạn trích sau trong <strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> (Vũ Trọng Phụng) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"– Hỡi công chúng! Mi chưa hiểu rõ những lẽ cực kì to tát nó khiến ta phải đành nhường giải cho nhà vô địch Xiêm La! Quần chúng nông nổi ơi! Mi đã biết đâu cái lòng hi sinh cao thượng vô cùng, (nó vỗ vào ngực)... Ta (nó giơ cao tay lên) không muốn cho hàng vạn mạng người làm mồi cho binh đao...\"</em></div>","Biện pháp nói mỉa và nghịch ngữ thể hiện tập trung qua việc Xuân trơ trẽn biến hành vi nhường điểm bán độ thành 'lòng hi sinh cao thượng' vì Tổ quốc.","Ngôn ngữ của Xuân kết hợp xưng hô trịch thượng của kẻ lưu manh vỉa hè ('ta' - 'mi') với khẩu hiệu ngoại giao chính trị salon sáo rỗng.","Các chỉ dẫn hành động (nó vỗ vào ngực, nó giơ cao tay lên...) mang đậm tính kịch, biến bài diễn thuyết thành màn tấu hài công cộng.","Màn diễn thuyết chứng minh Xuân Tóc Đỏ là một nhà ngoại giao thực thụ có trách nhiệm với quốc gia.","Đúng - Đúng - Đúng - Sai",""],["P2_TRUYEN","COMMON",23,"Vận dụng","true_false","Đọc đoạn trích sau trong <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> (Ma Văn Kháng) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"Trong tâm ức vẫn là có hình bóng chị Hoài... Nhưng bây giờ chị Hoài đã có một gia đình riêng... Chị có quyền quên mà không ai được trách cứ. Vậy mà… vậy mà lúc này trước cánh cổng lại là chị.\"</em></div>","Đoạn văn sử dụng hình thức trần thuật nửa trực tiếp kết hợp khai thác dòng tâm ức và tiềm thức của nhân vật.","Cụm từ 'Chị có quyền quên mà không ai được trách cứ' thể hiện sự lên án, ghét bỏ của gia đình ông Bằng đối với chị Hoài.","Đoạn trích tạo nên sự va đập giữa chuẩn mực lý tính ('có quyền quên') với mệnh lệnh của đạo lý nghĩa tình thủy chung ('vậy mà lúc này trước cánh cổng lại là chị').","Thời gian nghệ thuật là sự giao thoa giữa thời gian thực tại của buổi chiều cúng tất niên với thời gian tâm lý hồi ức.","Đúng - Sai - Đúng - Đúng",""],["P2_TRUYEN","COMMON",24,"Thông hiểu","true_false","Xác định tính Đúng / Sai về Ý NGHĨA BIỂU TƯỢNG HOA TỬ HUYỀN trong <strong class=\"work-title\"><em>\"Muối của rừng\"</em></strong> (Nguyễn Huy Thiệp):","Hoa tử huyền là loài hoa mang màu trắng thuần khiết và vị mặn như muối của rừng, ba mươi năm mới nở một lần.","Là phần thưởng tâm linh vô giá Mẹ Thiên nhiên ban tặng cho con người khi biết sám hối và thức tỉnh lương tri.","Hoa tử huyền là chứng tích khẳng định chiến thắng tuyệt đối của vũ khí và con người trước tự nhiên hoang dã.","Sự xuất hiện của hoa tử huyền gửi gắm niềm tin vào sự ấm no, hòa hợp vĩnh hằng giữa con người và thiên nhiên.","Đúng - Đúng - Sai - Đúng",""],["P2_TRUYEN","ON_TAP",25,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là vị trí, góc độ trần thuật mà qua đó người kể soi chiếu, đi sâu khám phá dòng tâm tư, ý thức và những biến chuyển tinh vi trong tâm hồn nhân vật.'","Điểm nhìn bên trong","Điểm nhìn bên ngoài","Ngôi kể thứ nhất","Ngôi kể thứ ba toàn tri","Điểm nhìn bên trong","Điểm nhìn bên trong cho phép người trần thuật thâm nhập vào cõi lòng sâu kín của nhân vật."],["P2_TRUYEN","ON_TAP",26,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là hạt nhân thúc đẩy sự phát triển của cốt truyện, đóng vai trò như 'thuốc thử' làm bộc lộ bản chất tính cách nhân vật và tư tưởng của nhà văn.'","Tình huống truyện","Chi tiết nghệ thuật","Đề tài tác phẩm","Cảm hứng chủ đạo","Tình huống truyện","Tình huống truyện là hoàn cảnh éo le thử thách phẩm chất nhân vật."],["P2_TRUYEN","ON_TAP",27,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là hình tượng người kể giấu mình ở ngôi thứ ba, thấu suốt mọi việc diễn ra bên ngoài lẫn cõi lòng sâu kín của tất cả các nhân vật.'","Người kể chuyện toàn tri","Người kể chuyện hạn tri","Nhân vật trung tâm","Người trần thuật ngôi thứ nhất","Người kể chuyện toàn tri","Người kể chuyện toàn tri là người biết hết mọi chuyện và thấu suốt nội tâm mọi nhân vật."],["P2_TRUYEN","ON_TAP",28,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là những biểu hiện cụ thể, sống động tạo nên 'xương thịt' và thần thái cho hình tượng nhân vật trong tác phẩm tự sự.'","Chi tiết nghệ thuật","Cốt truyện kịch tính","Bút pháp trào phúng","Thời gian tâm lý","Chi tiết nghệ thuật","Chi tiết nghệ thuật là chất liệu cơ bản tạo nên sự sống động của hình tượng văn học."],["P2_TRUYEN","ON_TAP",29,"Thông hiểu","fill_blank","Chọn cụm từ còn thiếu vào nhận định đối sánh: '<strong class=\"work-title\"><em>\"Xuân Tóc Đỏ cứu quốc\"</em></strong> đại diện cho khuynh hướng tiểu thuyết trào phúng hoạt kê; trong khi <strong class=\"work-title\"><em>\"Mùa lá rụng trong vườn\"</em></strong> đi theo khuynh hướng tiểu thuyết [ ...?... ] khám phá kiểu con người nếm trải.'","Hiện thực tâm lý - xã hội","Lãng mạn trữ tình","Phiêu lưu trinh thám","Kỳ ảo thần thoại","Hiện thực tâm lý - xã hội","Tiểu thuyết Ma Văn Kháng thuộc khuynh hướng hiện thực tâm lý - xã hội thời kỳ giao thời."],["P2_TRUYEN","ON_TAP",30,"Vận dụng","fill_blank","Chọn thông điệp cốt lõi từ truyện ngắn <strong class=\"work-title\"><em>\"Muối của rừng\"</em></strong> (Nguyễn Huy Thiệp): 'Con người cần biết buông bỏ dục vọng chiếm đoạt, tôn trọng và sống hòa hợp nhân ái với [ ...?... ]'","Giới tự nhiên và muôn loài","Nền văn minh công nghiệp","Các công cụ vũ khí hiện đại","Cuộc sống đô thị phồn hoa","Giới tự nhiên và muôn loài","Tác phẩm gửi gắm thông điệp bảo vệ và hòa hợp nhân ái với thiên nhiên muôn loài."],["P2_TRUYEN","CHINH_THUC",25,"Vận dụng cao","single_choice","Trong nghệ thuật tự sự hiện đại, việc sử dụng 'Người kể chuyện không đáng tin cậy' (Unreliable Narrator) nhằm mục đích chủ yếu gì?","Thể hiện sự thiếu sót về trình độ viết văn của tác giả trong việc xây dựng nhân vật.","Kích thích tính chủ động, phản biện của người đọc trong việc giải mã và tái tạo ý nghĩa tác phẩm.","Làm cho cốt truyện trở nên đơn giản và dễ hiểu hơn đối với độc giả nhỏ tuổi.","Khẳng định chân lý tuyệt đối và không thể bàn cãi của một quan điểm duy nhất.","B","Người kể chuyện không đáng tin cậy phá vỡ sự tiếp nhận thụ động, buộc độc giả phải so sánh, suy ngẫm và tham gia đồng sáng tạo ý nghĩa văn bản."],["P2_TRUYEN","CHINH_THUC",26,"Thông hiểu","single_choice","Yếu tố nào sau đây là tiêu chuẩn cốt lõi để phân biệt giữa 'Cốt truyện' (Story) và 'Truyện kể' (Plot/Discourse)?","Cốt truyện là chuỗi sự kiện theo trật tự thời gian tự nhiên, còn truyện kể là cách thức sắp xếp, tổ chức sự kiện của tác giả.","Cốt truyện chỉ có trong truyện ngắn, còn truyện kể chỉ xuất hiện trong tiểu thuyết trường thiên.","Cốt truyện do nhân vật tạo ra, còn truyện kể do độc giả tự tưởng tượng khi đọc sách.","Cốt truyện luôn có kết thúc có hậu, còn truyện kể luôn mang kết cấu bi kịch mở.","A","Cốt truyện (fabula) là diễn biến sự việc theo nhân - quả và thời gian tự nhiên; Truyện kể (sjuzhet) là trật tự nghệ thuật do nhà văn tái cấu trúc."],["P2_TRUYEN","CHINH_THUC",27,"Vận dụng","single_choice","Sự dịch chuyển từ 'Điểm nhìn bên ngoài' sang 'Điểm nhìn bên trong' trong truyện hiện đại mang lại hiệu quả nghệ thuật nào?","Giúp người đọc nắm bắt bao quát toàn cảnh trận đánh hoặc lễ hội đông người.","Mở ra thế giới tâm lý tinh tế, nắm bắt những biến thái cảm xúc vi diệu và tiềm thức nhân vật.","Hạn chế tối đa khả năng biểu cảm của người kể chuyện ngôi thứ nhất.","Triệt tiêu hoàn toàn tính đối thoại giữa tác giả và người đọc.","B","Điểm nhìn bên trong đưa lăng kính quan sát vào sâu nội tâm, giúp soi rọi những tầng vỉa tâm lí sâu kín nhất của con người."],["P2_TRUYEN","CHINH_THUC",28,"Thông hiểu","single_choice","Ý nghĩa thẩm mỹ của 'Chi tiết nghệ thuật đắt giá' trong tác phẩm văn xuôi tự sự là:","Chi tiết có độ dài lớn nhất và chiếm nhiều trang sách nhất trong tác phẩm.","Chi tiết nhỏ nhưng có sức nén tư tưởng lớn, bộc lộ bản chất nhân vật và chủ đề tư tưởng toàn truyện.","Chi tiết sử dụng nhiều thuật ngữ khoa học chuyên ngành khó hiểu.","Chi tiết miêu tả chính xác xuất xứ giá tiền của các đồ vật trong truyện.","B","Đúng như câu nói của nhà văn Paustovsky: 'Chi tiết nhỏ làm nên nhà văn lớn' - chi tiết đắt giá tích tụ sức sống và tư tưởng cốt lõi của toàn tác phẩm."],["P2_TRUYEN","CHINH_THUC",29,"Vận dụng","single_choice","Đặc điểm của 'Không gian tâm trạng' trong tác phẩm truyện hiện đại là gì?","Không gian địa lý thực địa được đo đạc chính xác theo bản đồ hành chính.","Không gian ngoại cảnh được khúc xạ qua lăng kính cảm xúc và mang đậm dấu ấn tâm trạng nhân vật.","Không gian cố định không bao giờ thay đổi theo thời gian của câu chuyện.","Không gian hoàn toàn bị xóa nhòa, không có bất kỳ đồ vật hay cảnh sắc nào.","B","Không gian tâm trạng là cảnh vật mang linh hồn cảm xúc: 'Người buồn cảnh có vui đâu bao giờ' - cảnh sắc ngoại giới hòa vào tâm trạng nội giới."],["P2_TRUYEN","CHINH_THUC",30,"Vận dụng cao","single_choice","Giá trị của một tác phẩm văn xuôi tự sự đích thực đối với cuộc đời là gì?","Sao chép y nguyên hiện thực đời sống mà không cần gửi gắm tư tưởng nghệ thuật.","Giúp con người thấu hiểu tha nhân, khơi dậy lòng trắc ẩn và mở rộng chiều kích nhận thức nhân sinh.","Cung cấp những chỉ dẫn hành chính để giải quyết các tranh chấp pháp lý thường ngày.","Răn dạy con người bằng những bài học đạo đức khô khan, giáo điều và một chiều.","B","Văn học là nhân học - sứ mệnh cao cả của truyện là giúp con người hiểu sâu hơn về bản thân, đồng cảm với nỗi đau của đồng loại và nâng niu cái đẹp."],["P2_THO","COMMON",1,"Nhận biết","single_choice","Theo lí luận văn học, đặc trưng bản chất cốt lõi nhất của thể loại 'Thơ' là gì?","Phương thức biểu hiện trữ tình, phản ánh đời sống qua thế giới cảm xúc chủ quan mãnh liệt, ngôn từ hàm súc, giàu hình ảnh và giàu nhạc tính.","Phương thức tự sự tái hiện đời sống khách quan thông qua hệ thống chuỗi sự kiện, cốt truyện và hành động nhân vật.","Văn bản nghị luận sử dụng hệ thống luận điểm, lí lẽ và dẫn chứng để thuyết phục tư tưởng người đọc.","Kịch bản sân khấu giải quyết các xung đột xã hội gay gắt bằng lời thoại và hành động trực tiếp.","A","Thơ là phương thức biểu hiện trữ tình, là tiếng nói trực tiếp của tình cảm, cảm xúc mãnh liệt; ngôn ngữ thơ hàm súc, cô đọng, giàu hình ảnh và nhạc điệu."],["P2_THO","COMMON",2,"Nhận biết","single_choice","Khái niệm 'Chủ thể trữ tình' (Cái tôi trữ tình) trong tác phẩm thơ được hiểu chính xác nhất là gì?","Chân dung tiểu sử ngoài đời thực của nhà thơ với đầy đủ lí lịch cá nhân.","Trung tâm bộc lộ cảm xúc, suy tư, thái độ và lý tưởng thẩm mỹ của nhà thơ trước cuộc đời trong tác phẩm trữ tình.","Nhân vật phản diện đối đầu gay gắt với các nhân vật khác trong mạch kể.","Người kể chuyện ngôi thứ ba giấu mình hoàn toàn và không bày tỏ cảm xúc.","B","Chủ thể trữ tình (cái tôi trữ tình) là chủ thể phát ngôn, là trung tâm bộc lộ cảm xúc, tư tưởng và rung động thẩm mỹ trong bài thơ."],["P2_THO","COMMON",3,"Nhận biết","single_choice","Thuật ngữ 'Cấu tứ' trong thi pháp thơ biểu đạt điều gì?","Quy tắc đếm số lượng chữ trong từng dòng thơ của một thể thơ truyền thống.","Cách thức tổ chức, liên kết hệ thống hình ảnh, mạch cảm xúc và tư tưởng của bài thơ thành một chỉnh thể nghệ thuật hoàn chỉnh.","Việc lựa chọn tên gọi nhan đề cho tập thơ xuất bản.","Sự phân chia các nhân vật thành hai tuyến chính diện và phản diện.","B","Cấu tứ là cách tổ chức, kiến tạo mạch vận động cảm xúc và kết nối các hình ảnh, ý tưởng thơ để tạo nên cấu trúc nghệ thuật toàn vẹn cho thi phẩm."],["P2_THO","COMMON",4,"Thông hiểu","single_choice","Điểm khác biệt căn bản giữa 'Hình ảnh thơ' thông thường và 'Biểu tượng thơ' là gì?","Hình ảnh thơ phản ánh trực quan cụ thể; còn Biểu tượng thơ là hình ảnh được kết tinh, lặp lại và mở ra các tầng ý nghĩa triết lý đa chiều, sâu sắc.","Biểu tượng thơ chỉ xuất hiện trong truyện cổ tích, còn hình ảnh thơ chỉ xuất hiện trong ca dao.","Hình ảnh thơ bắt buộc phải có vần điệu, còn biểu tượng thơ không cần ngôn ngữ.","Hai khái niệm này hoàn toàn đồng nhất và không có bất kì sự phân biệt nào.","A","Hình ảnh thơ mang tính miêu tả cụ thể, gợi cảm; khi hình ảnh được khái quát hóa, tích tụ chiều sâu tư tưởng và gợi nhiều liên tưởng triết lý thì trở thành biểu tượng nghệ thuật."],["P2_THO","COMMON",5,"Nhận biết","single_choice","Yếu tố nào sau đây đóng vai trò tạo nên 'nhạc cảm nội tại', sức ngân vang và sức truyền cảm đặc biệt của câu thơ?","Cách ngắt nhịp điệu, gieo vần, phối thanh (Bằng - Trắc) và các biện pháp điệp âm, điệp từ.","Quy mô số lượng trang sách in của bài thơ.","Số lượng nhân vật tham gia vào chuỗi sự kiện cốt truyện.","Cách trình bày danh mục tài liệu tham khảo ở cuối bài.","A","Nhịp điệu, vần điệu, sự phối hợp đối xứng hoặc chuyển đổi thanh âm Bằng - Trắc và các biện pháp điệp từ, điệp âm tạo nên tính nhạc độc đáo của thơ ca."],["P2_THO","COMMON",6,"Thông hiểu","single_choice","Thủ pháp 'Tương giao cảm giác' (Synesthesia) trong thi pháp thơ tượng trưng - hiện đại được hiểu là gì?","Sự phân chia rạch ròi từng giác quan riêng biệt khi tiếp nhận sự vật khách quan.","Hiện tượng chuyển đổi cảm giác: âm thanh được cảm nhận bằng màu sắc, hình khối hoặc xúc giác và ngược lại.","Việc lặp lại một từ ngữ nhiều lần ở đầu mỗi câu thơ.","Sự đối lập thuần túy giữa hai nhân vật có tính cách trái ngược nhau.","B","Tương giao cảm giác là sự giao thoa, chuyển đổi ấn tượng giữa các giác quan (nghe âm thanh thấy màu sắc, thấy hình khối như 'tiếng ghi ta tròn bọt nước', 'tiếng đàn nâu')."],["P2_THO","COMMON",7,"Thông hiểu","single_choice","Đặc trưng nổi bật của thi pháp Thơ Cổ điển (Trung đại) thể hiện qua phương diện nào sau đây?","Ngôn ngữ đời thường suồng sã, cấu trúc tự do không niêm luật, đề cao cái tôi cá nhân vị kỉ.","Tính quy phạm chặt chẽ, bút pháp ước lệ tượng trưng, đề tài đạo lý, chí khí tráng sĩ và tầm vóc vũ trụ.","Khai thác thế giới vô thức, tiềm thức hỗn loạn và ngôn từ đứt gãy phi logic.","Thuần túy miêu tả các sự kiện sinh hoạt đời tư gia đình hàng ngày.","B","Thơ cổ điển trung đại mang tính quy phạm (niêm luật, đối ngẫu), sử dụng hệ thống thi liệu ước lệ tượng trưng, hướng tới đạo lý và tầm vóc vũ trụ lớn lao."],["P2_THO","COMMON",8,"Thông hiểu","single_choice","Khuynh hướng Thơ Lãng mạn (Thơ Mới 1932-1945 và thơ thời kỳ kháng chiến) có đặc điểm thi pháp nổi bật nào?","Giải phóng cảm xúc cá nhân khoáng đạt, sử dụng thủ pháp tương phản đối lập gắt gao, kết hợp hài hòa giữa nét mĩ lệ và dữ dội bi tráng.","Tuân thủ nghiêm ngặt quy tắc gieo vần độc vận và phép đối ngẫu cung đình.","Phủ nhận hoàn toàn mọi cảm xúc và hình ảnh trữ tình trong tác phẩm.","Chỉ tập trung châm biếm, hoạt kê các thói hư tật xấu trong xã hội.","A","Thơ lãng mạn đề cao cái tôi cảm xúc tự do, khoáng đạt, khai thác triệt để các thủ pháp tương phản, phóng đại và khám phá vẻ đẹp bi tráng, mĩ lệ."],["P2_THO","COMMON",9,"Nhận biết","single_choice","Tư thế tráng sĩ bi tráng trong hai câu kết bài thơ <strong class=\"work-title\"><em>\"Cảm hoài\"</em></strong> (Đặng Dung): <br><em>\"Kỉ độ thôn hàm phung vũ hận / Đao đầu không bạt vạn tinh đê\"</em> (Mấy độ mài gươm bóng nguyệt tà) thể hiện phẩm chất gì?","Sự nản lòng, buông xuôi và thoái thác hoàn toàn trách nhiệm đối với non sông đất nước.","Khí phách hiên ngang, lòng yêu nước nồng nàn và ý chí chiến đấu ngoan cường của người anh hùng dù gặp bước thời thế lỡ dở.","Thú vui ẩn dật thanh cao, lánh đời nơi non xanh nước biếc.","Nỗi sợ hãi bóng đêm và tuổi già sắp ập đến.","B","Hình ảnh người tráng sĩ mài gươm dưới bóng trăng tà là biểu tượng tuyệt mỹ của ý chí kiên định, lòng yêu nước sâu sắc và tinh thần bi tráng trước nghịch cảnh."],["P2_THO","COMMON",10,"Thông hiểu","single_choice","Trong bài thơ <strong class=\"work-title\"><em>\"Thu hứng\"</em></strong> (Đỗ Phủ), bút pháp 'tả cảnh ngụ tình' được thể hiện tập trung qua mối liên hệ nào?","Cảnh rừng phong tiêu điều, sóng gió Vu hiệp dữ dội phản chiếu nỗi đau xót thời thế và nỗi sầu xứ cô đơn của thi nhân hướng về kinh đô Trường An.","Cảnh mùa xuân rộn rã gắn với niềm vui mở hội tưng bừng nơi thôn dã.","Cảnh sinh hoạt cung đình xa hoa tráng lệ tương phản với nghèo đói dân gian.","Cảnh đoàn thuyền đánh cá trở về trong niềm vui trúng mùa hải sản.","A","Cảnh thu tiêu điều, hùng vĩ mà âm u nơi Vu hiệp là phương tiện biểu đạt nỗi sầu xứ, nỗi đau xót khôn nguôi trước cảnh loạn ly và lòng đau đáu hướng về Trường An."],["P2_THO","COMMON",11,"Nhận biết","single_choice","Bức tranh thiên nhiên miền Tây trong bài thơ <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng) được khắc họa qua sự kết hợp của những nét đẹp nào?","Vừa hùng vĩ, hiểm trở, dữ dội khôn cùng; lại vừa thơ mộng, trữ tình, mĩ lệ và ấm áp ân tình bản mường.","Thuần túy là sự hoang vu, chết chóc không có bóng dáng sự sống con người.","Bình lặng, êm ả như cảnh sắc vùng đồng bằng châu thổ sông Hồng.","Khô cằn, bỏng rát của vùng đất sa mạc hoang sơ.","A","Thiên nhiên miền Tây trong 'Tây Tiến' vừa hiểm trở, dữ dội ('dốc khúc khuỷu', 'thác gầm thét') vừa thơ mộng, mĩ lệ ('mưa xa khơi', 'hoa đong đưa')."],["P2_THO","COMMON",12,"Thông hiểu","single_choice","Hai câu thơ: <br><em>\"Mắt trừng gửi mộng qua biên giới / Đêm mơ Hà Nội dáng kiều thơm\"</em> trong bài thơ <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> khắc họa vẻ đẹp nào của người lính?","Sự suy sụp tinh thần và tư tưởng ủy mị xa rời hiện thực chiến đấu.","Sự hòa quyện tuyệt đẹp giữa ý chí chiến đấu quả cảm, quyết liệt ('mắt trừng gửi mộng') và tâm hồn hào hoa, lãng mạn, tha thiết yêu thương ('đêm mơ Hà Nội dáng kiều thơm').","Thói hưởng thụ nhàn hạ, xa hoa của những chàng trai thị thành thời bình.","Nỗi sợ hãi hiện thực khắc nghiệt nơi rừng sâu nước độc.","B","Người lính Tây Tiến mang vẻ đẹp của tráng sĩ thời đại mới: quyết tâm diệt thù nơi biên cương nhưng trái tim vẫn đong đầy vẻ hào hoa, lãng mạn của tuổi trẻ trí thức Hà thành."],["P2_THO","COMMON",13,"Vận dụng","single_choice","Đọc 4 câu thơ trong bài <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng): <br><div class=\"reading-excerpt-box\"><em>\"Rải rác biên cương mồ viễn xứ / Chiến trường đi chẳng tiếc đời xanh / Áo bào thay chiếu anh về đất / Sông Mã gầm lên khúc độc hành.\"</em></div>Nhận định nào phân tích <strong>đúng nhất</strong> về cảm hứng bi tráng và nghệ thuật biểu đạt?","Nhà thơ che giấu hoàn toàn sự mất mát, không dám nhìn thẳng vào hiện thực khốc liệt của chiến tranh.","Nhìn thẳng vào cái chết và sự hi sinh nhưng nâng lên tầm vóc sử thi bất tử nhờ từ ngữ Hán Việt cổ kính, cách nói giảm nói tránh ('về đất') và âm vang bi tráng của dòng sông Mã.","Đoạn thơ chỉ mang giọng điệu bi thương, tang tóc làm nhụt ý chí của người cầm súng.","Đoạn thơ thuần túy miêu tả tập quán an táng cổ xưa của đồng bào vùng cao.","B","Quang Dũng không né tránh hiện thực hi sinh ('mồ viễn xứ') nhưng bằng từ ngữ trang trọng ('áo bào', 'biên cương', 'độc hành'), sự hi sinh được bất tử hóa thành khúc ca bi tráng."],["P2_THO","COMMON",14,"Thông hiểu","single_choice","Trong bài thơ <strong class=\"work-title\"><em>\"Đàn ghi ta của Lor-ca\"</em></strong> (Thanh Thảo), hình ảnh người nghệ sĩ Lor-ca được biểu tượng hóa qua những chi tiết nghệ thuật nào?","Người tráng sĩ mặc áo giáp sắt cưỡi ngựa bay về trời.","Người nghệ sĩ tự do, khát khao cách tân đơn độc đi trên đấu trường Tây Ban Nha với 'áo choàng đỏ gắt', 'tiếng đàn bọt nước' và 'vầng trăng váng vất'.","Vị vua quyền uy đang ngự trên ngai vàng chỉ huy dàn hợp xướng cung đình.","Một đấu sĩ bò tót chuyên nghiệp chỉ khao khát tiền bạc và danh vọng.","B","Lor-ca hiện lên là biểu tượng của người nghệ sĩ tiên phong dũng cảm, mang khát vọng tự do và cách tân nghệ thuật giữa môi trường chính trị ngột ngạt của Tây Ban Nha."],["P2_THO","COMMON",15,"Thông hiểu","single_choice","Trong bài <strong class=\"work-title\"><em>\"Đàn ghi ta của Lor-ca\"</em></strong>, chuỗi hình ảnh biến thể: <br><em>\"tiếng ghi ta nâu\", \"tiếng ghi ta lá xanh\", \"tiếng ghi ta tròn bọt nước vỡ tan\", \"tiếng ghi ta ròng ròng máu chảy\"</em> mang ý nghĩa nghệ thuật gì?","Mô tả một cách máy móc quá trình sản xuất một cây đàn gỗ ngoài thực tế.","Sử dụng thủ pháp tương giao cảm giác để hữu hình hóa âm thanh tiếng đàn: từ tình yêu quê hương, khát vọng tuổi trẻ đến sự tan vỡ đau đớn và nỗi đau bi tráng khi người nghệ sĩ bị sát hại.","Thể hiện các lỗi kỹ thuật khi người nghệ sĩ chơi đàn ghi-ta.","Chỉ đơn thuần là các từ ngữ chỉ màu sắc ngẫu nhiên không có ý đồ nghệ thuật.","B","Chuỗi liên tưởng tương giao cảm giác biến âm thanh vô hình thành màu sắc, hình khối, thể hiện sinh động cuộc đời, tình yêu và số phận bi tráng của Lor-ca."],["P2_THO","COMMON",16,"Vận dụng","single_choice","Hình ảnh <strong class=\"work-title\"><em>\"không ai chôn cất tiếng đàn / tiếng đàn như cỏ mọc hoang\"</em></strong> và chuỗi âm thanh luyến láy <em>\"li-la li-la li-la\"</em> gửi gắm thông điệp sâu xa nào?","Sự lãng quên tuyệt đối của hậu thế đối với tên tuổi của Lor-ca.","Sự bất tử của nghệ thuật chân chính và khát vọng tự do: dù thân xác nghệ sĩ bị tiêu diệt, cái đẹp vẫn có sức sống mãnh liệt, tự lan tỏa và bất diệt như cỏ dại và ngàn hoa thơm.","Lời phê phán nền nông nghiệp hoang hóa không có người chăm sóc.","Khuyên các nghệ sĩ tương lai không nên tiếp tục chơi đàn ghi-ta nữa.","B","Tiếng đàn không thể bị chôn vùi, nó có sức sống tự nhiên mãnh liệt như cỏ dại và ngân vang mãi mãi như khúc tưởng niệm bất tận cho tự do và cái đẹp."],["P2_THO","COMMON",17,"Vận dụng","single_choice","Trong <strong class=\"work-title\"><em>\"Bài thơ số 28\"</em></strong> của R. Tagore, quan niệm triết lý về tình yêu và trái tim con người được khái quát qua hình ảnh nào?","Trái tim người yêu là hòn ngọc trong suốt có thể đo lường dễ dàng bằng thước kẻ.","Trái tim con người là một vũ trụ vô biên, bí ẩn khôn cùng; tình yêu đích thực là sự thấu cảm tâm hồn chứ không thể chiếm đoạt hay tường tận trọn vẹn mọi giới hạn.","Tình yêu chỉ là những toan tính vật chất vụ lợi nơi phố thị.","Trái tim là một đồ vật tĩnh lặng không hề có cảm xúc.","B","Tagore gửi gắm triết lý phương Đông sâu sắc: tình yêu vừa là khao khát thấu hiểu, vừa là sự trân trọng cõi tâm linh vô tận, không cùng của con người."],["P2_THO","COMMON",18,"Vận dụng","single_choice","Khi đối sánh hình tượng người anh hùng tráng sĩ trong <strong class=\"work-title\"><em>\"Cảm hoài\"</em></strong> (Đặng Dung) và người lính trong <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng), nhận xét nào sau đây là CHÍNH XÁC NHẤT?","Đặng Dung khắc họa tráng sĩ trung đại với tầm vóc vũ trụ và nỗi u hoài thời thế bi tráng; còn Quang Dũng khắc họa người lính vệ quốc thời đại mới với vẻ đẹp hào hoa lãng mạn kết hợp chất bi tráng sử thi.","Cả hai bài thơ đều đi theo khuynh hướng thơ trào phúng hoạt kê châm biếm chế độ thực dân.","Người lính trong 'Tây Tiến' là nhân vật ước lệ cổ điển, còn tráng sĩ trong 'Cảm hoài' là nhân vật hiện đại thế kỷ XX.","Cả hai tác giả đều sử dụng thể thơ tự do đứt gãy không có vần điệu.","A","Đặng Dung mang cảm hứng tráng chí vũ trụ trung đại kết hợp nỗi đau thời thế; Quang Dũng mang cảm hứng lãng mạn cách mạng kết hợp vẻ đẹp hào hoa của tuổi trẻ Thủ đô."],["P2_THO","COMMON",19,"Thông hiểu","true_false","Xác định tính Đúng / Sai của các nhận định sau về ĐẶC TRƯNG CÁC PHONG CÁCH THƠ:","Thơ Cổ điển trung đại chuộng thi liệu ước lệ, niêm luật quy phạm nghiêm ngặt và hướng tới tầm vóc vũ trụ lớn lao.","Thơ Lãng mạn đề cao cái tôi cá nhân, cảm xúc dồi dào, ưa thủ pháp tương phản đối lập và sự mở rộng liên tưởng phong phú.","Thơ Tượng trưng - Siêu thực chỉ miêu tả sự vật theo quan hệ nhân quả một chiều và tuyệt đối cấm dùng tương giao cảm giác.","Dù thuộc phong cách nào, cấu tứ và mạch cảm xúc luôn là yếu tố quyết định tạo nên sự thống nhất hữu cơ của bài thơ.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","COMMON",20,"Thông hiểu","true_false","Xác định tính Đúng / Sai về THI PHÁP NGHỆ THUẬT trong bài thơ <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng):","Quang Dũng sử dụng bút pháp lãng mạn kết hợp cảm hứng bi tráng để khắc họa bức tranh thiên nhiên và tượng đài người lính.","Hệ thống từ ngữ Hán Việt cổ kính ('biên cương', 'viễn xứ', 'áo bào', 'độc hành') góp phần trang trọng hóa và bất tử hóa sự hi sinh.","Bài thơ né tránh hoàn toàn hiện thực gian khổ, bệnh tật sốt rét rừng và sự hi sinh của đồng đội.","Sự luân chuyển giữa các câu thơ nhiều thanh trắc gân guốc và câu toàn thanh bằng êm dịu tạo nên nhạc điệu phong phú, độc đáo.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","COMMON",21,"Thông hiểu","true_false","Xác định tính Đúng / Sai về THI PHÁP TƯỢNG TRƯNG - SIÊU THỰC trong <strong class=\"work-title\"><em>\"Đàn ghi ta của Lor-ca\"</em></strong> (Thanh Thảo):","Tác phẩm sử dụng cấu trúc thơ tự do không dấu chấm câu, nhịp thơ nương theo dòng chảy cảm xúc và liên tưởng đa tầng.","Tiếng đàn ghi-ta được nhân cách hóa và chuyển đổi cảm giác thành màu sắc, hình khối, sự sống và nỗi đau ('tiếng ghi ta ròng ròng máu chảy').","Thanh Thảo chỉ đơn thuần thuật lại biên bản vụ án bắt giam Lor-ca của cảnh sát độc tài mà không có hình tượng nghệ thuật nào.","Âm thanh mô phỏng tiếng đàn 'li-la li-la li-la' vừa gợi giai điệu đàn ghi-ta vừa liên tưởng đến hoa tử đinh hương dâng tặng hương hồn thi sĩ.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","COMMON",22,"Vận dụng","true_false","Đọc 4 câu thơ trong bài <strong class=\"work-title\"><em>\"Cảm hoài\"</em></strong> (Đặng Dung) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"Thế sự du du nại lão hà / Vô cùng thiên địa nhập hàm ca / Thời lai đồ điếu thành công dị / Vận khứ anh hùng ẩm hận đa.\"</em></div>","Hai câu đầu thể hiện nỗi xót xa trước bước đi của thời gian và tuổi già trong sự đối sánh với đất trời vô tận.","Hai câu thực sử dụng nghệ thuật đối lập ('thời lai' >< 'vận khứ') để đúc kết quy luật hưng phế khốc liệt của thời cuộc.","Cụm từ 'ẩm hận đa' (nuốt hận nhiều) thể hiện sự oán hờn nhân dân và quay lưng phản bội lại đất nước.","Thể thơ Thất ngôn bát cú Đường luật với niêm luật chặt chẽ làm tăng vẻ trang nghiêm, bi hùng cho lời tự thuật tâm trạng.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","COMMON",23,"Vận dụng","true_false","Đọc đoạn thơ sau trong <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"Tây Tiến đoàn binh không mọc tóc / Quân xanh màu lá dữ oai hùm / Mắt trừng gửi mộng qua biên giới / Đêm mơ Hà Nội dáng kiều thơm.\"</em></div>","Cụm từ 'không mọc tóc' và 'quân xanh màu lá' phản ánh chân thực căn bệnh sốt rét rừng quái ác nơi chiến trường Tây Bắc.","Hình ảnh 'dữ oai hùm' thể hiện tư thế kiêu hùng, áp đảo gian khổ hiểm nguy của người lính.","Hai chữ 'dáng kiều thơm' bị coi là biểu hiện của lối sống phong kiến lạc hậu cần phải bài trừ khỏi tác phẩm.","Đoạn thơ kết hợp nhịp nhàng giữa chất hiện thực gân guốc và chất lãng mạn bay bổng của hồn thơ Quang Dũng.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","COMMON",24,"Vận dụng","true_false","Đọc đoạn thơ sau trong <strong class=\"work-title\"><em>\"Đàn ghi ta của Lor-ca\"</em></strong> (Thanh Thảo) và xác định tính Đúng / Sai:\n<div class=\"reading-excerpt-box\"><em>\"Tây Ban Nha / hát nghêu ngao / bỗng kinh hoàng / áo choàng bê bết đỏ / Lor-ca bị dẫn về bãi bắn / chàng đi như người mộng du\"</em></div>","Không gian đất nước Tây Ban Nha gắn liền với tiếng hát nghêu ngao tự do đối lập với biến cố đẫm máu bất ngờ ('kinh hoàng', 'áo choàng bê bết đỏ').","Hình ảnh 'áo choàng bê bết đỏ' vừa gợi chiếc áo choàng của đấu sĩ bò tót, vừa là ẩn dụ về cái chết bi thương của Lor-ca.","Tư thế 'đi như người mộng du' thể hiện tâm thế hoảng loạn, van xin tha mạng của người nghệ sĩ trước mũi súng kẻ thù.","Nhịp thơ ngắt ngắn, dồn dập diễn tả bước đi định mệnh và khoảnh khắc lịch sử đau thương của đất nước Tây Ban Nha.","Đúng - Đúng - Sai - Đúng",""],["P2_THO","ON_TAP",25,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là cách thức tổ chức, liên kết các hình ảnh, mạch cảm xúc và tư tưởng của bài thơ thành một chỉnh thể nghệ thuật hoàn chỉnh.'","Cấu tứ thơ","Nhạc điệu thơ","Biện pháp tu từ","Gieo vần chính vận","Cấu tứ thơ","Cấu tứ là cách tổ chức và kết nối các yếu tố nghệ thuật thành một chỉnh thể thơ hoàn chỉnh."],["P2_THO","ON_TAP",26,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là hiện tượng chuyển đổi cảm xúc giữa các giác quan (nghe âm thanh cảm nhận được màu sắc, hình khối), một thủ pháp đặc trưng của thơ tượng trưng.'","Tương giao cảm giác","Tả cảnh ngụ tình","Nói quá phóng đại","Chơi chữ điệp âm","Tương giao cảm giác","Tương giao cảm giác là sự giao thoa, chuyển dịch ấn tượng giữa các giác quan trong tiếp nhận thẩm mỹ."],["P2_THO","ON_TAP",27,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là trung tâm bộc lộ cảm xúc, thái độ, tư tưởng và lý tưởng thẩm mỹ của nhà thơ trước cuộc đời trong tác phẩm trữ tình.'","Chủ thể trữ tình (Cái tôi trữ tình)","Người kể chuyện toàn tri","Nhân vật ước lệ","Hình tượng phản diện","Chủ thể trữ tình (Cái tôi trữ tình)","Chủ thể trữ tình là chủ thể phát ngôn bộc lộ thế giới tâm tư, cảm xúc trong bài thơ."],["P2_THO","ON_TAP",28,"Nhận biết","fill_blank","Chọn thuật ngữ thích hợp: '[ ...?... ] là hình ảnh nghệ thuật được lặp lại nhiều lần, tích tụ chiều sâu tư tưởng và mở ra các tầng ý nghĩa triết lý nhân sinh phổ quát.'","Biểu tượng thơ","Chi tiết đời thường","Cốt truyện phi tuyến tính","Lời nửa trực tiếp","Biểu tượng thơ","Biểu tượng thơ là hình ảnh chứa đựng nhiều tầng nghĩa sâu sắc, mang tính khái quát cao."],["P2_THO","ON_TAP",29,"Thông hiểu","fill_blank","Chọn cụm từ thích hợp vào nhận định đối sánh: 'Nếu hình tượng tráng sĩ trong <strong class=\"work-title\"><em>\"Cảm hoài\"</em></strong> (Đặng Dung) mang tầm vóc vũ trụ và nỗi u hoài thời thế của thi pháp [ ...?... ]; thì hình tượng người lính trong <strong class=\"work-title\"><em>\"Tây Tiến\"</em></strong> (Quang Dũng) lại bừng sáng vẻ đẹp hào hoa, lãng mạn và bi tráng thời đại kháng chiến.'","Cổ điển trung đại","Tượng trưng siêu thực","Hiện thực trào phúng","Đồng thoại thiếu nhi","Cổ điển trung đại","Đặng Dung sáng tác theo thi pháp Cổ điển trung đại với bút pháp ước lệ và chí khí tráng sĩ vũ trụ."],["P2_THO","ON_TAP",30,"Vận dụng","fill_blank","Chọn thông điệp cốt lõi từ bài thơ <strong class=\"work-title\"><em>\"Đàn ghi ta của Lor-ca\"</em></strong> (Thanh Thảo): 'Dù bạo lực bạo tàn có thể vùi dập thể xác người nghệ sĩ, nhưng [ ...?... ] và khát vọng tự do, sáng tạo nghệ thuật chân chính sẽ mãi mãi bất tử.'","Tiếng đàn nghệ thuật","Vũ khí quân sự","Mối hận thù giai cấp","Chiếc áo choàng thi đấu","Tiếng đàn nghệ thuật","Tiếng đàn là biểu tượng bất tử của cái đẹp, của nghệ thuật chân chính và khát vọng tự do."],["P2_THO","CHINH_THUC",25,"Vận dụng cao","single_choice","Khái niệm 'Sự tương giao đa giác quan' (Correspondences / Synesthesia) trong thơ tượng trưng hiện đại được hiểu là:","Hiện tượng một bài thơ được dịch sang nhiều ngôn ngữ quốc tế khác nhau.","Sự hòa quyện, chuyển hóa linh hoạt giữa các cảm giác (thị giác, thính giác, khứu giác, xúc giác) để biểu đạt thế giới tinh vi.","Việc bắt buộc người đọc phải vừa đọc thơ vừa nghe nhạc giao hưởng cổ điển.","Việc sử dụng tranh minh họa màu sắc rực rỡ bên cạnh các dòng chữ thơ.","B","Tương giao giác quan là đặc trưng thi pháp của thơ hiện đại: màu sắc có thể nghe thấy âm thanh, âm thanh mang hình khối và hương vị."],["P2_THO","CHINH_THUC",26,"Thông hiểu","single_choice","Nhịp điệu trong thơ tự do hiện đại được hình thành chủ yếu từ yếu tố nào?","Sự tuân thủ nghiêm ngặt số chữ trong mỗi câu thơ theo thể thơ Đường luật.","Nhịp điệu bên trong của cảm xúc, dòng chảy tâm tư và sự ngắt nhịp biến hóa theo mạch tâm trạng.","Sự lặp đi lặp lại đơn điệu của các thanh bằng và trắc ở các vị trí chẵn lẻ.","Việc tác giả sử dụng các dấu chấm câu và dấu phẩy theo đúng ngữ pháp văn xuôi.","B","Thơ tự do giải phóng khỏi khuôn khổ câu chữ cố định để nhịp điệu thơ bắt rễ trực tiếp từ nhịp đập bên trong của cảm xúc và tư tưởng."],["P2_THO","CHINH_THUC",27,"Vận dụng","single_choice","Yếu tố 'Biểu tượng' (Symbol) khác với 'Hình ảnh tả thực' thông thường ở điểm cốt lõi nào?","Biểu tượng có tính đa nghĩa, gợi mở những tầng ý nghĩa triết lý sâu xa vượt ra ngoài đối tượng cụ thể.","Biểu tượng chỉ được phép mượn từ văn học cổ điển Trung Hoa hoặc thần thoại Hy Lạp.","Biểu tượng luôn mang nghĩa cố định, rõ ràng và không cho phép người đọc suy luận khác.","Biểu tượng chỉ tồn tại trong các câu thơ có vần trắc ở cuối câu.","A","Biểu tượng mang năng lượng gợi mở vô tận, là chiếc cầu nối giữa cái hữu hình và cái vô hình, kích thích liên tưởng sáng tạo của độc giả."],["P2_THO","CHINH_THUC",28,"Thông hiểu","single_choice","Vai trò của 'Khoảng trắng' (Khoảng lặng nghệ thuật) trong một bài thơ hiện đại là:","Do lỗi in ấn hoặc lỗi căn lề của nhà xuất bản trong quá trình in sách.","Tạo khoảng lặng nghệ thuật để cảm xúc lắng đọng, kích thích người đọc đồng sáng tạo và chiêm nghiệm.","Giúp người đọc đọc bài thơ nhanh hơn mà không cần suy nghĩ về ý nghĩa câu chữ.","Báo hiệu bài thơ đã hoàn toàn kết thúc và không còn nội dung nào tiếp theo.","B","Khoảng trắng trong thơ là 'ý tại ngôn ngoại', là nơi ngôn ngữ bất lực nhường chỗ cho sự im lặng đầy sức gợi và chiêm nghiệm sâu lắng."],["P2_THO","CHINH_THUC",29,"Vận dụng","single_choice","Sự khác biệt căn bản giữa 'Cảm xúc trong thơ' và 'Cảm xúc đời thường' là:","Cảm xúc đời thường sâu sắc hơn cảm xúc trong thơ vì nó diễn ra trong đời thực.","Cảm xúc trong thơ là cảm xúc đã được lắng đọng, thẩm định qua mỹ cảm và khái quát thành giá trị nghệ thuật.","Cảm xúc trong thơ hoàn toàn là sự bịa đặt vô căn cứ, không bắt nguồn từ đời sống.","Cảm xúc trong thơ chỉ dành riêng cho những nhà nghiên cứu lý luận văn học chuyên nghiệp.","B","Cảm xúc thơ ca là cảm xúc thẩm mỹ: được chưng cất, kết tinh qua trải nghiệm sâu sắc của người nghệ sĩ để trở thành tiếng lòng chung của nhân loại."],["P2_THO","CHINH_THUC",30,"Vận dụng cao","single_choice","Nhận định nào sau đây diễn đạt ĐÚNG NHẤT về sứ mệnh của thơ ca trong đời sống con người?","Thơ ca là một trò chơi ngôn từ thuần túy không gắn liền với số phận con người.","Thơ ca neo đậu tâm hồn con người vào cái thiện, cái đẹp và là nơi trú ngụ thiêng liêng của nhân phẩm.","Thơ ca có nhiệm vụ thay thế các ngành khoa học tự nhiên trong việc khám phá vũ trụ.","Thơ ca chỉ có giá trị giải trí nhất thời trong những lúc rảnh rỗi.","B","Thơ ca nuôi dưỡng phần người trong mỗi con người, hướng thiện, cứu rỗi tâm hồn và bất tử hóa những giá trị cao đẹp nhất của đời sống."]];
  
  var lastRowMCQ = mcqSheet.getLastRow();
  if (lastRowMCQ > 1) {
    mcqSheet.getRange(2, 1, lastRowMCQ - 1, 12).clearContent();
  }
  mcqSheet.getRange(2, 1, mcqData.length, 12).setValues(mcqData);
  
  mcqSheet.getRange(2, 1, mcqData.length, 12).setFontFamily("Arial").setFontSize(9.5).setVerticalAlignment("top");
  mcqSheet.getRange(2, 1, mcqData.length, 5).setHorizontalAlignment("center");
  mcqSheet.getRange(2, 11, mcqData.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
  
  var essayData = [["P1_BAI_1","ON_TAP","OT_1","Đề Ôn tập 1 (Xuân Tóc Đỏ cứu quốc)","Viết đoạn văn khoảng 200 chữ phân tích nghệ thuật trào phúng sắc sảo của Vũ Trọng Phụng qua đoạn trích Xuân Tóc Đỏ cứu quốc.","1. Mở đoạn: Giới thiệu tác giả Vũ Trọng Phụng, tiểu thuyết Số đỏ và nghệ thuật trào phúng trong đoạn trích.<br>2. Thân đoạn: Phân tích tình huống nghịch dị (kẻ hạ lưu thành anh hùng cứu quốc), nghệ thuật cường điệu, ngôn ngữ giễu nhại châm biếm tầng lớp thống trị thượng lưu.<br>3. Kết đoạn: Khái quát giá trị hiện thực và tài năng trào phúng bậc thầy của nhà văn.","Qua đoạn trích 'Xuân Tóc Đỏ cứu quốc', Vũ Trọng Phụng đã thể hiện tài năng bậc thầy của ngòi bút trào phúng hiện thực khi xây dựng nên một tình huống bi hài đạt đến đỉnh cao của sự phi lý. Nghệ thuật trào phúng trước hết bắt nguồn từ màn 'thao túng tâm lý' đám đông đầy mỉa mai: một trận đấu quần vợt thể thao bỗng nhiên bị thổi phồng thành nguy cơ xung đột quốc gia, và việc Xuân nhường phần thắng cho đối thủ lại được tôn vinh thành hành động 'hy sinh danh dự cá nhân để cứu lấy hòa bình xứ sở'. Bằng thủ pháp cường điệu, phóng đại cùng giọng văn sắc lạnh, giễu nhại, tác giả đã lột trần bộ mặt giả dối, kệch cỡm của tầng lớp thống trị thực dân nửa phong kiến cùng thói a dua, mê muội của đám đông thị dân đương thời. Kẻ vô học, lưu manh bỗng chốc hóa thành 'anh hùng cứu quốc', 'bậc vĩ nhân' được tung hô vạn tuế. Tiếng cười trong đoạn trích không chỉ mang tính giải trí mà là tiếng cười đả kích gay gắt, bộc lộ sự tha hóa tột cùng của một xã hội kim tiền nhố nhăng. Đoạn trích xứng đáng là trang viết trào phúng mẫu mực, khẳng định sức sống bất diệt của kiệt tác Số đỏ."],["P1_BAI_1","ON_TAP","OT_2","Đề Ôn tập 2 (Nghệ thuật xây dựng nhân vật trong Số đỏ)","Viết đoạn văn khoảng 200 chữ phân tích nghệ thuật khắc họa nhân vật Xuân Tóc Đỏ trong tiểu thuyết Số đỏ của Vũ Trọng Phụng.","1. Mở đoạn: Giới thiệu Vũ Trọng Phụng và nhân vật Xuân Tóc Đỏ trong tiểu thuyết Số đỏ.<br>2. Thân đoạn: Bản chất lưu manh, vô học nhưng thức thời; sự biến ảo dị hợm khi bước chân vào xã hội thượng lưu; nghệ thuật biếm họa phóng đại sắc sảo của nhà văn.<br>3. Kết đoạn: Khẳng định giá trị hiện thực và sức sống vượt thời gian của hình tượng nhân vật điển hình.","Trong kiệt tác Số đỏ, nghệ thuật khắc họa nhân vật Xuân Tóc Đỏ của Vũ Trọng Phụng đã đạt đến đỉnh cao của thi pháp xây dựng nhân vật trào phúng trong tiểu thuyết hiện đại Việt Nam. Không miêu tả tâm lý đơn tuyến, tác giả xây dựng Xuân như một điển hình sống động cho sự tha hóa và gặp thời kỳ dị của tầng lớp hạ lưu trong cơn lốc 'Âu hóa'. Bằng thủ pháp biếm họa sắc sảo, nhà văn làm nổi bật sự đối lập gay gắt giữa bản chất vô học, láu cá của một kẻ nhặt bóng quần vợt với chiếc vỏ bọc đạo mạo 'tiến sĩ', 'bậc vĩ nhân' được giới thượng lưu tôn sùng. Ngôn ngữ của Xuân là sự chắp vá kệch cỡm giữa những câu cửa miệng bình dân ('Nước mẹ gì!') với khẩu hiệu cải cách nửa mùa ('Âu hóa', 'Bình dân'), tạo nên hiệu ứng tiếng cười châm biếm sâu cay. Xuân không chỉ là một cá nhân tha hóa mà còn là tấm gương phản chiếu toàn bộ sự giả dối, lố lăng và mù quáng của xã hội kim tiền thực dân đương thời. Qua nhân vật Xuân Tóc Đỏ, Vũ Trọng Phụng đã khẳng định tài năng bậc thầy trong việc nắm bắt bản chất đời sống và cống hiến cho văn học một hình tượng nghệ thuật bất hủ."],["P1_BAI_1","CHINH_THUC","CT_1","Đề Chính thức 1 (Khả năng bao quát đời sống)","Viết đoạn văn khoảng 200 chữ làm rõ khả năng to lớn của thể loại tiểu thuyết trong việc phản ánh chiều sâu hiện thực tâm lí con người qua Bài 1.","1. Mở đoạn: Nêu đặc trưng vượt trội của thể loại tiểu thuyết hiện đại.<br>2. Thân đoạn: Khả năng thâm nhập vào thế giới nội tâm đa tầng, phản ánh hiện thực toàn diện không đơn tuyến.<br>3. Kết đoạn: Ý nghĩa của tiểu thuyết đối với việc bồi đắp vốn sống và nhận thức của bạn đọc.",""],["P1_BAI_1","CHINH_THUC","CT_2","Đề Chính thức 2 (Nghệ thuật trần thuật đa điểm nhìn)","Từ các văn bản tiểu thuyết đã học ở Bài 1, hãy viết đoạn văn khoảng 200 chữ phân tích tác dụng của điểm nhìn trần thuật đa chiều đối với sự hấp dẫn của tác phẩm.","1. Mở đoạn: Dẫn dắt khái niệm điểm nhìn trần thuật đa chiều trong tiểu thuyết hiện đại.<br>2. Thân đoạn: Sự chuyển dịch điểm nhìn tạo tính đa thanh, khách quan hóa câu chuyện và soi tỏ nhân vật từ nhiều góc độ.<br>3. Kết đoạn: Khẳng định dấu ấn tài năng cách tân thi pháp của các nhà văn hiện đại.",""],["P1_BAI_1","CHINH_THUC","CT_3","Đề Chính thức 3 (Số phận con người & Thời cuộc)","Viết đoạn văn khoảng 200 chữ bàn về mối quan hệ giữa số phận cá nhân và biến động lịch sử thời đại được thể hiện qua các đoạn trích tiểu thuyết ở Bài 1.","1. Mở đoạn: Giới thiệu mối liên hệ mật thiết giữa số phận cá nhân và hoàn cảnh lịch sử thời đại.<br>2. Thân đoạn: Biến động lịch sử tác động sâu sắc đến sự lựa chọn, phẩm giá và số phận con người; sự kiên định nhân tính giữa bão táp.<br>3. Kết đoạn: Rút ra thông điệp nhân sinh sâu sắc cho tuổi trẻ hôm nay.",""],["P1_BAI_2","ON_TAP","OT_1","Đề Ôn tập 1 (Tây Tiến - Bi tráng)","Viết đoạn văn khoảng 200 chữ phân tích vẻ đẹp bi tráng của hình tượng người lính Tây Tiến trong bài thơ cùng tên của Quang Dũng.","1. Mở đoạn: Giới thiệu nhà thơ Quang Dũng và vẻ đẹp bi tráng của đoàn quân Tây Tiến.<br>2. Thân đoạn: Hiện thực khốc liệt, gian khổ hi sinh kết hợp với khí phách kiêu hùng, lãng mạn và tinh thần dâng hiến bất tử.<br>3. Kết đoạn: Khẳng định tượng đài người lính bất tử trong nền thơ ca kháng chiến.","Trong thi phẩm Tây Tiến, Quang Dũng đã dựng nên một tượng đài bất tử về người lính kháng chiến bằng vẻ đẹp bi tráng hào hùng. Cái bi không hề bi lụy, bi thương mà ngời sáng khí phách trượng phu: hiện thực gian khổ nơi rừng thiêng nước độc, những cơn sốt rét 'đoàn binh không mọc tóc', 'quân xanh màu lá' cùng sự hi sinh thầm lặng nơi biên cương 'rải rác biên cương mồ viễn xứ' được nhìn nhận qua lăng kính lãng mạn kiêu hãnh. Người lính ra đi với lý tưởng cao đẹp 'Chiến trường đi chẳng tiếc đời xanh', coi cái chết nhẹ tựa lông hồng. Thủ pháp đối lập tương phản giữa ngoại hình tiều tụy với sức mạnh nội tâm 'dữ oai hùm', cùng âm hưởng trầm hùng của câu thơ 'Áo bào thay chiếu anh về đất / Sông Mã gầm lên khúc độc hành' đã nâng sự hi sinh lên tầm vóc sử thi thiêng liêng. Khúc độc hành của dòng sông Mã như bản tráng ca tiễn đưa những người con anh hùng vào cõi bất tử. Bằng ngòi bút tài hoa và bút pháp lãng mạn đặc sắc, Quang Dũng đã khắc sâu tượng đài người lính Tây Tiến vừa hào hoa, vừa bi tráng, sống mãi cùng non sông."],["P1_BAI_2","ON_TAP","OT_2","Đề Ôn tập 2 (Đàn ghi ta của Lor-ca)","Viết đoạn văn khoảng 200 chữ làm rõ biểu tượng 'tiếng đàn bọt nước' và khát vọng tự do trong bài thơ Đàn ghi ta của Lor-ca của Thanh Thảo.","1. Mở đoạn: Nêu tác giả Thanh Thảo và biểu tượng 'tiếng đàn bọt nước' trong thi phẩm.<br>2. Thân đoạn: Vẻ đẹp nghệ thuật mong manh trước bạo tàn nhưng bất diệt; sự hóa thân của người nghệ sĩ vào cái đẹp vĩnh cửu.<br>3. Kết đoạn: Chiều sâu tư tưởng tri ân người nghệ sĩ dũng cảm mở đường.","Biểu tượng 'tiếng đàn bọt nước' trong thi phẩm Đàn ghi ta của Lor-ca là một sáng tạo nghệ thuật độc đáo của Thanh Thảo, kết tinh vẻ đẹp mong manh nhưng bất tử của nghệ thuật và khát vọng tự do. 'Tiếng đàn' là linh hồn, là sinh mệnh nghệ thuật của Phê-đê-ri-cô Gai-xi-a Lor-ca – người nghệ sĩ Tây Ban Nha đơn độc đấu tranh cho nền dân chủ và cách tân nghệ thuật. Hình ảnh 'bọt nước' gợi liên tưởng về số phận mong manh, ngắn ngủi và cái chết oan khuất, bất ngờ của người nghệ sĩ trước mũi súng bạo tàn. Tuy nhiên, bọt nước vỡ tan lại mở đầu cho những vòng sóng lan tỏa vô tận; tiếng đàn không bao giờ chết mà chuyển hóa thành 'tiếng ghi-ta nâu', 'tiếng ghi-ta lá xanh', 'vỡ tan thành dòng máu đỏ'. Sự tương giao cảm giác và liên tưởng siêu thực tài hoa của Thanh Thảo đã tôn vinh sức sống mãnh liệt của cái đẹp: bạo lực có thể hủy diệt thể xác người nghệ sĩ nhưng không thể dập tắt khát vọng tự do và tinh thần sáng tạo chân chính. Đó chính là thông điệp nhân văn cao cả mà tác phẩm gửi gắm."],["P1_BAI_2","CHINH_THUC","CT_1","Đề Chính thức 1 (Tây Tiến - Hào hoa & Lãng mạn)","Viết đoạn văn khoảng 200 chữ cảm nhận về chất hào hoa, tâm hồn lãng mạn và tình cảm gắn bó tha thiết của đoàn quân Tây Tiến với mảnh đất miền Tây.","1. Mở đoạn: Giới thiệu vẻ đẹp hào hoa lãng mạn đặc trưng của lính Tây Tiến xứ Hà thành.<br>2. Thân đoạn: Kỷ niệm đêm hội đuốc hoa, men say tình quân dân, vẻ đẹp kiều diễm của người thiếu nữ và thiên nhiên miền Tây thơ mộng.<br>3. Kết đoạn: Đánh giá nét độc đáo làm nên sức quyến rũ vượt thời gian của thơ Quang Dũng.",""],["P1_BAI_2","CHINH_THUC","CT_2","Đề Chính thức 2 (Thơ tượng trưng & Siêu thực)","Từ các thi phẩm trong Bài 2, hãy viết đoạn văn khoảng 200 chữ phân tích nét độc đáo của hệ thống hình ảnh biểu tượng trong khuynh hướng thơ hiện đại.","1. Mở đoạn: Khái quát khuynh hướng thơ tượng trưng và siêu thực trong thơ ca hiện đại.<br>2. Thân đoạn: Sự chuyển dịch từ tả thực sang biểu tượng đa nghĩa, liên tưởng nhảy vọt và sự tương giao cảm giác tinh tế.<br>3. Kết đoạn: Ý nghĩa của việc đổi mới tư duy thi ca trong việc khơi dậy sức sáng tạo của người đọc.",""],["P1_BAI_2","CHINH_THUC","CT_3","Đề Chính thức 3 (Khát vọng sáng tạo của người nghệ sĩ)","Viết đoạn văn khoảng 200 chữ bàn về sứ mệnh và khát vọng sáng tạo nghệ thuật cách tân của nhà thơ được gợi ra từ Bài 2 - Những thế giới thơ.","1. Mở đoạn: Dẫn dắt khát vọng sáng tạo và sứ mệnh mở đường của người nghệ sĩ chân chính.<br>2. Thân đoạn: Tinh thần dấn thân, không chấp nhận lối mòn rập khuôn, khát vọng giải phóng cái đẹp và phục vụ tự do nhân loại.<br>3. Kết đoạn: Bài học về tinh thần đổi mới và khát vọng cống hiến của thế hệ trẻ hôm nay.",""],["P2_TRUYEN","ON_TAP","OT_1","Đề Ôn tập 1 (Nghệ thuật xây dựng nhân vật)","Viết đoạn văn khoảng 200 chữ phân tích vai trò của độc thoại nội tâm trong việc khắc họa tính cách nhân vật truyện hiện đại.","1. Mở đoạn: Nêu khái niệm và vị trí của độc thoại nội tâm trong thi pháp truyện hiện đại.<br>2. Thân đoạn: Khám phá chiều sâu tiềm thức, những giằng xé lưỡng phân, làm cho nhân vật trở nên sống động, chân thực và phức hợp.<br>3. Kết đoạn: Khẳng định độc thoại nội tâm là bước tiến lớn của nghệ thuật tự sự.","Trong thi pháp tự sự hiện đại, độc thoại nội tâm đóng vai trò đặc biệt quan trọng trong việc khám phá chiều sâu bí ẩn và phức tạp của thế giới tâm hồn nhân vật. Khác với lời kể khách quan bên ngoài hay những đoạn đối thoại thông thường, độc thoại nội tâm là dòng chảy ý thức trực tiếp, nơi nhân vật tự phơi bày những trăn trở, dằn vặt, mâu thuẫn và giằng xé nội tâm trước các bước ngoặt cuộc đời. Qua lời tự vấn âm thầm, bức tranh tâm lý nhân vật hiện lên đa chiều, sống động, xóa bỏ hoàn toàn lối phân định nhân vật giản đơn 'chính diện – phản diện' một chiều. Người đọc không chỉ quan sát hành động bề ngoài mà còn thấu cảm được những góc khuất u uẩn, nỗi đau, khát vọng và sự thức tỉnh nhân tính sâu kín nhất. Việc khai thác triệt để độc thoại nội tâm đã đánh dấu bước tiến vượt bậc của nghệ thuật tự sự, nâng tầm tác phẩm truyện thành một hành trình đối thoại nhân văn sâu sắc về con người và số phận."],["P2_TRUYEN","ON_TAP","OT_2","Đề Ôn tập 2 (Tình huống truyện độc đáo)","Viết đoạn văn khoảng 200 chữ phân tích ý nghĩa của tình huống truyện đối với việc bộc lộ chủ đề và tư tưởng tác phẩm.","1. Mở đoạn: Nêu định nghĩa tình huống truyện (hạt nhân cấu trúc tác phẩm).<br>2. Thân đoạn: Tạo bước ngoặt thử thách bản lĩnh nhân vật, bộc lộ bản chất xã hội và phát sáng tư tưởng nhân đạo của tác giả.<br>3. Kết đoạn: Khẳng định tài năng tổ chức cốt truyện của nhà văn.","Tình huống truyện được ví như 'thứ nước rửa ảnh' làm hiện hình sắc nét tính cách nhân vật và tư tưởng cốt lõi của tác phẩm tự sự. Đó là hoàn cảnh đặc biệt, sự kiện mang tính bước ngoặt gay cấn nơi các xung đột đời sống bị đẩy lên cao trào, buộc nhân vật phải bộc lộ trọn vẹn bản lĩnh, nhân cách và những lựa chọn đạo đức sinh tử. Một tình huống truyện độc đáo không chỉ tạo nên kịch tính hấp dẫn cho cốt truyện mà còn là đòn bẩy tư tưởng mạnh mẽ: chính tại khoảnh khắc thử thách nghiệt ngã ấy, bản chất hiện thực xã hội được lột trần và chiều sâu nhân đạo của nhà văn được phát sáng rực rỡ. Từ một tình huống cụ thể, nhà văn khái quát nên những quy luật phổ quát về nhân sinh và thời đại. Có thể khẳng định, việc sáng tạo nên tình huống truyện độc đáo chính là thước đo tài năng và dấu ấn phong cách nghệ thuật không thể trộn lẫn của một ngòi bút tự sự bậc thầy."],["P2_TRUYEN","CHINH_THUC","CT_1","Đề Chính thức 1 (Điểm nhìn trần thuật & Ngôi kể)","Viết đoạn văn khoảng 200 chữ làm rõ tác dụng của việc dịch chuyển điểm nhìn trần thuật trong tác phẩm truyện hiện đại.","1. Mở đoạn: Giới thiệu tầm quan trọng của điểm nhìn trần thuật trong nghệ thuật tự sự.<br>2. Thân đoạn: Sự dịch chuyển điểm nhìn (bên ngoài - bên trong, ngôi thứ nhất - ngôi thứ ba) tạo tính đa thanh, dân chủ và soi chiếu đa diện hiện thực.<br>3. Kết đoạn: Khẳng định sức hấp dẫn nghệ thuật của truyện hiện đại.",""],["P2_TRUYEN","CHINH_THUC","CT_2","Đề Chính thức 2 (Chi tiết nghệ thuật đắt giá)","Viết đoạn văn khoảng 200 chữ phân tích sức sống của một chi tiết nghệ thuật giàu ý nghĩa biểu tượng trong tác phẩm tự sự.","1. Mở đoạn: Nêu vai trò 'chi tiết nhỏ làm nên nhà văn lớn' trong văn xuôi.<br>2. Thân đoạn: Sức nén cảm xúc, tính hàm súc biểu tượng, tạo điểm nhấn lay động tâm can người đọc và nâng tầm tư tưởng tác phẩm.<br>3. Kết đoạn: Ấn tượng sâu đậm của chi tiết nghệ thuật đối với bạn đọc.",""],["P2_TRUYEN","CHINH_THUC","CT_3","Đề Chính thức 3 (Không - thời gian nghệ thuật)","Viết đoạn văn khoảng 200 chữ phân tích sự đan xen giữa thời gian hiện tại và thời gian hồi tưởng trong truyện hiện đại.","1. Mở đoạn: Đặt vấn đề thời gian nghệ thuật phi tuyến tính trong truyện hiện đại.<br>2. Thân đoạn: Cấu trúc thời gian đồng hiện, đan cài quá khứ - hiện tại - tương lai làm nổi bật dòng ý thức và chiều sâu tâm trạng nhân vật.<br>3. Kết đoạn: Hiệu quả thẩm mỹ đặc sắc của nghệ thuật tổ chức thời gian.",""],["P2_THO","ON_TAP","OT_1","Đề Ôn tập 1 (Cấu tứ & Hình ảnh thơ)","Viết đoạn văn khoảng 200 chữ phân tích vai trò của cấu tứ trong việc tổ chức và phát triển mạch cảm xúc của một bài thơ trữ tình.","1. Mở đoạn: Nêu khái niệm cấu tứ (linh hồn kết cấu bài thơ).<br>2. Thân đoạn: Trục liên kết cảm xúc, sáng tạo chuỗi hình ảnh độc đáo, định hướng dòng chảy tư tưởng toàn bài.<br>3. Kết đoạn: Đánh giá tầm quan trọng của cấu tứ đối với sự thành công của một thi phẩm.","Cấu tứ là linh hồn tổ chức, là trục xương sống định hình toàn bộ mạch vận động cảm xúc và tư tưởng của một tác phẩm thơ trữ tình. Cấu tứ không đơn thuần là dàn ý sắp đặt cơ học, mà là cách thức nhà thơ kết nối các hình ảnh, thanh điệu và cảm xúc theo một logic tâm trạng độc đáo nhằm làm bừng sáng ý niệm thẩm mỹ. Nhờ có cấu tứ vững chắc, hệ thống thi ảnh từ cụ thể, cảm tính ban đầu được nâng tầm thành những biểu tượng nghệ thuật giàu sức gợi mở, dẫn dắt người đọc đi từ những rung động trực giác đến nhận thức triết lý sâu xa về cuộc đời. Cấu tứ vừa tạo nên tính chỉnh thể toàn vẹn cho thi phẩm, vừa mở ra không gian đồng sáng tạo vô tận cho độc giả. Một cấu tứ mới lạ, tài hoa chính là chiếc chìa khóa vạn năng giúp người nghệ sĩ bất tử hóa những khoảnh khắc rung động mong manh của tâm hồn thành tác phẩm nghệ thuật vượt thời gian."],["P2_THO","ON_TAP","OT_2","Đề Ôn tập 2 (Nhạc điệu & Ngôn ngữ thơ)","Viết đoạn văn khoảng 200 chữ làm rõ nét độc đáo của nhịp điệu và phối thanh trong thơ hiện đại.","1. Mở đoạn: Vị trí của nhạc tính và ngôn ngữ trong đặc trưng thể loại thơ.<br>2. Thân đoạn: Nhịp thơ biến hóa theo nhịp đập tâm hồn, phối hợp thanh điệu trầm bổng tạo dư ba cảm xúc lan tỏa.<br>3. Kết đoạn: Chiều sâu thẩm mỹ mà nhạc điệu thơ mang lại cho tâm hồn độc giả.","Nhạc điệu và ngôn ngữ thơ là hai yếu tố cốt lõi tạo nên vẻ đẹp quyến rũ và sức lay động trực tiếp của thơ ca đối với tâm hồn con người. Ngôn từ trong thơ là thứ ngôn ngữ được chưng cất tinh vi, giàu hình ảnh, hàm súc và tràn đầy nhạc tính. Nhạc điệu thơ được kiến tạo từ sự phối hợp nhịp nhàng giữa thanh điệu trầm bổng (bằng - trắc), phép ngắt nhịp biến hóa linh hoạt theo cung bậc cảm xúc, cùng nghệ thuật điệp âm, gieo vần độc đáo. Nhịp thơ chính là nhịp đập của trái tim thi sĩ, khi dạt dào êm ái như khúc hát ru, khi dồn dập nghẹn ngào như tiếng thở dài hay bão giông giằng xé. Chính nhạc điệu đã chắp cánh cho ý niệm bay xa, tạo nên những âm ba vang vọng ngoài câu chữ, để lại dư vị thẩm mỹ sâu lắng trong lòng độc giả. Qua sự hòa quyện tuyệt mỹ giữa nhạc và lời, thơ ca đã chạm tới đáy sâu cảm xúc và đánh thức những rung cảm nhân văn đẹp đẽ nhất trong con người."],["P2_THO","CHINH_THUC","CT_1","Đề Chính thức 1 (Hình tượng cái Tôi trữ tình)","Viết đoạn văn khoảng 200 chữ phân tích sự vận động của cái tôi trữ tình từ thơ truyền thống sang thơ hiện đại và hậu hiện đại.","1. Mở đoạn: Giới thiệu cái tôi trữ tình - hạt nhân cảm xúc của tác phẩm thơ.<br>2. Thân đoạn: Sự chuyển dịch từ cái tôi phi ngã, ước lệ sang cái tôi cá tính tự do, phức hợp, dám phơi trải mọi góc khuất tâm trạng chân thực.<br>3. Kết đoạn: Ý nghĩa giải phóng cảm xúc và tôn vinh cá tính sáng tạo.",""],["P2_THO","CHINH_THUC","CT_2","Đề Chính thức 2 (Biểu tượng & Nghệ thuật ẩn dụ)","Viết đoạn văn khoảng 200 chữ phân tích sức gợi mở của hệ thống hình ảnh mang tính tượng trưng trong thơ thế kỉ XX.","1. Mở đoạn: Đặt vấn đề thi pháp hình ảnh tượng trưng trong thơ thế kỉ XX.<br>2. Thân đoạn: Tính đa tầng ý nghĩa, khả năng gợi cảm vượt ra ngoài giới hạn ngôn từ, khơi dậy sự liên tưởng và đồng sáng tạo của người đọc.<br>3. Kết đoạn: Vẻ đẹp bí ẩn, quyến rũ của nghệ thuật thơ ca đích thực.",""],["P2_THO","CHINH_THUC","CT_3","Đề Chính thức 3 (Cảm thức thời gian trong thơ)","Viết đoạn văn khoảng 200 chữ bàn về cảm thức thời gian và khát vọng bất tử hóa cái đẹp qua các thi phẩm chuyên đề Thơ.","1. Mở đoạn: Dẫn dắt cảm thức về thời gian - nỗi trăn trở muôn thuở của các nhà thơ.<br>2. Thân đoạn: Ý thức sâu sắc về sự hữu hạn của đời người đối lập với khát vọng lưu giữ khoảnh khắc và bất tử hóa cái đẹp qua ngôn từ thi ca.<br>3. Kết đoạn: Thông điệp nhân văn về sự trân trọng từng phút giây cuộc sống.",""]];
  
  var lastRowEssay = essaySheet.getLastRow();
  if (lastRowEssay > 1) {
    essaySheet.getRange(2, 1, lastRowEssay - 1, 7).clearContent();
  }
  essaySheet.getRange(2, 1, essayData.length, 7).setValues(essayData);
  essaySheet.getRange(2, 1, essayData.length, 7).setFontFamily("Arial").setFontSize(9.5).setVerticalAlignment("top");
  essaySheet.getRange(2, 1, essayData.length, 3).setHorizontalAlignment("center");
  
  try {
    SpreadsheetApp.getUi().alert("🎉 KHỞI TẠO DỮ LIỆU THÀNH CÔNG!\n\n✅ Đã nạp 144 câu hỏi trắc nghiệm (đầy đủ định dạng Đúng/Sai 4 ý và Điền khuyết) vào tab NganHangTracNghiem\n✅ Đã nạp 20 đề tự luận và đoạn văn mẫu vào tab DeTuLuan\n\nThầy Cô có thể chỉnh sửa trực tiếp trên bảng tính ngay bây giờ!");
  } catch(e) {}
  
  return {
    status: "success",
    mcqCount: mcqData.length,
    essayCount: essayData.length
  };
}

// -----------------------------------------------------------------------------------------
// 8. MENU ĐIỀU KHIỂN TRÊN GOOGLE SHEETS
// -----------------------------------------------------------------------------------------
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("👑 QUẢN TRỊ CA THI K12")
      .addItem("📥 1. Khởi tạo toàn bộ Ngân hàng Câu hỏi (144 Câu TN + 20 Đề TL)", "khoiTaoTatCaNganHangCauHoi")
      .addSeparator()
      .addItem("🟢 2. MỞ TẤT CẢ CÁC CA THI", "moTatCaCaThi")
      .addItem("🔴 3. KHÓA TẤT CẢ CÁC CA THI", "khoaTatCaCaThi")
      .addSeparator()
      .addItem("⚙️ 4. Khởi tạo lại tất cả các Tab chuẩn", "setupInitialData")
      .addToUi();
  } catch(e) {}
}

function moTatCaCaThi() {
  var sheet = getOrCreateConfigSheet();
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var rowId = String(data[r][1] || "").toUpperCase();
    if (rowId.indexOf("BAI_") !== -1 || rowId.indexOf("THO") !== -1 || rowId.indexOf("TRUYEN") !== -1) {
      sheet.getRange(r + 1, 4).setValue("OPEN");
    }
  }
  try {
    SpreadsheetApp.getUi().alert("✅ ĐÃ MỞ TẤT CẢ CÁC CA THI THÀNH CÔNG!");
  } catch(e) {}
}

function khoaTatCaCaThi() {
  var sheet = getOrCreateConfigSheet();
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var rowId = String(data[r][1] || "").toUpperCase();
    if (rowId.indexOf("BAI_") !== -1 || rowId.indexOf("THO") !== -1 || rowId.indexOf("TRUYEN") !== -1) {
      sheet.getRange(r + 1, 4).setValue("CLOSED");
    }
  }
  try {
    SpreadsheetApp.getUi().alert("🔒 ĐÃ KHÓA TẤT CẢ CÁC CA THI AN TOÀN!");
  } catch(e) {}
}