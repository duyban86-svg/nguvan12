/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: HỆ THỐNG QUẢN LÝ THI & ĐỒNG BỘ ĐIỂM NGỮ VĂN 12 (NĂM HỌC 2026-2027)
 * =========================================================================================
 * Tương thích 100% với:
 *  1. P1_BAI_1_TIEU_THUYET.html (Bài 1 - Khả năng lớn lao của tiểu thuyết)
 *  2. P2_THE_LOAI_TRUYEN.html (Chuyên đề: Đặc trưng thể loại Truyện & Tiểu thuyết - 40 Phút)
 *  3. index.html (Hệ sinh thái Ôn thi Tốt nghiệp THPT & Auth Bridge SSO)
 *  4. Ứng dụng Sổ điểm cá nhân / Nhập điểm học sinh
 * 
 * TÍNH NĂNG MỚI ĐÃ CẬP NHẬT:
 *  - Hỗ trợ ghi nhận đầy đủ 2 phần: TRẮC NGHIỆM (7.0đ) & TỰ LUẬN ĐOẠN VĂN (3.0đ).
 *  - Phân hóa tự luận theo từng lớp (12A08, 12A13, 12A21), lưu tiêu đề đề bài, số từ và toàn văn bài viết.
 *  - Cập nhật thời gian làm bài chính thức 40 phút.
 *  - Chống gian lận: ghi nhận số lần chuyển tab / vi phạm quy chế.
 *  - Nhận diện Google SSO: lưu Email Google và Tên tài khoản Google.
 *  - Hỗ trợ đồng thời JSONP (GET) và POST (Form/JSON) bảo đảm 100% không bị chặn CORS trên mọi thiết bị.
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
      var topicId = params.topicId || "P2_TRUYEN";
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
        '    <div class="info-row"><span class="label">Mã Ca thi (Passcode):</span><span class="val">' + (config.examPasscode || config.passcode) + '</span></div>' +
        '    <div class="info-row"><span class="label">Trạng thái Ca thi:</span><span class="val">' + config.examStatus + '</span></div>' +
        '    <div class="info-row"><span class="label">Thời gian làm bài:</span><span class="val">' + config.duration + ' phút</span></div>' +
        '    <div class="info-row"><span class="label">Mã PIN Giáo viên:</span><span class="val">' + (config.teacherPin || config.masterPin) + '</span></div>' +
        '  </div>' +
        '</div>' +
        '</body></html>';
        
      return HtmlService.createHtmlOutput(html)
        .setTitle("Quản trị Ca thi - Google Sheets")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (!action) action = "check";

    // A. KIỂM TRA TRẠNG THÁI CA THI & CẤU HÌNH ĐỀ THI
    // DEBUG ACTION ĐỂ SOI CHÍNH XÁC NỘI DUNG TỪNG Ô TRÊN SHEET
    if (action === "debug") {
      var debugSheet = ss.getSheetByName("CauHinh") || ss.getSheetByName("Cấu hình") || ss.getSheetByName("Cấu Hình");
      var debugData = debugSheet ? debugSheet.getDataRange().getValues() : [];
      return sendResponse({
        status: "success",
        sheetFound: !!debugSheet,
        sheetName: debugSheet ? debugSheet.getName() : null,
        totalRows: debugData.length,
        rawRows: debugData,
        parsedConfig: getExamConfig(params.topicId || "P2_TRUYEN")
      }, callback);
    }

    if (action === "check" || action === "getExamConfig") {
      var topicId = params.topicId || "P2_TRUYEN";
      var config = getExamConfig(topicId);
      return sendResponse(config, callback);
    }

    // F. LẤY NGÂN HÀNG CÂU HỎI ĐỘNG TỪ GOOGLE SHEETS
    if (action === "get_questions" || action === "getQuestions") {
      var qTopicId = params.topicId || "P2_TRUYEN";
      var qData = getQuestionsFromSheet(qTopicId);
      return sendResponse(qData, callback);
    }

    // G. LƯU NGÂN HÀNG CÂU HỎI LÊN GOOGLE SHEETS
    if (action === "save_questions" || action === "saveQuestions") {
      var saveTopicId = params.topicId || "P2_TRUYEN";
      var qList = [];
      try {
        qList = JSON.parse(params.questionsJson || "[]");
      } catch(e) {}
      var saveRes = saveQuestionsToSheet(saveTopicId, qList);
      return sendResponse(saveRes, callback);
    }

    // B. NỘP BÀI THI QUA GET (JSONP - Phương thức chống chặn CORS tối ưu)
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
  lock.tryLock(10000);

  try {
    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        // Fallback x-www-form-urlencoded
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    var action = params.action || "submit";

    // A. LIÊN KẾT TÀI KHOẢN GOOGLE SSO (BIND ACCOUNT)
    if (action === "bind_account") {
      var bindResult = bindGoogleAccount(params);
      return sendJsonResponse(bindResult);
    }

    // B. GỬI BÁO LỖI / GÓP Ý
    if (action === "feedback") {
      var fbRes = saveFeedback(params);
      return sendJsonResponse(fbRes);
    }

    // C. NỘP BÀI THI
    if (action === "submit") {
      var subRes = saveExamSubmission(params);
      return sendJsonResponse(subRes);
    }

    // D. NHẬP ĐIỂM SỔ ĐIỂM (Hỗ trợ app Sổ điểm cá nhân)
    if (action === "score_entry" || params.students || params.scoreType) {
      var scoreRes = saveGradebookEntry(params);
      return sendJsonResponse(scoreRes);
    }

    var defRes = saveExamSubmission(params);
    return sendJsonResponse(defRes);

  } catch (err) {
    return sendJsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------------------------------
// 3. HÀM XỬ LÝ LƯU KẾT QUẢ BÀI THI (TRẮC NGHIỆM & TỰ LUẬN)
// -----------------------------------------------------------------------------------------
function saveExamSubmission(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var topicId = d.topicId || "P2_TRUYEN";
  var sheetName = d.sheetTarget || ("KẾT QUẢ - " + (d.examTopic || "CHUYÊN ĐỀ TRUYỆN"));
  
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
  var mode = d.mode || "Thi thử chính thức";
  
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

  // Đặt độ rộng các cột
  sheet.setColumnWidth(1, 150); // Thời gian
  sheet.setColumnWidth(2, 75);  // Lớp
  sheet.setColumnWidth(3, 55);  // STT
  sheet.setColumnWidth(4, 180); // Họ tên
  sheet.setColumnWidth(5, 200); // Email Google
  sheet.setColumnWidth(6, 130); // Chế độ
  sheet.setColumnWidth(7, 100); // Điểm TN
  sheet.setColumnWidth(8, 90);  // Số câu đúng
  sheet.setColumnWidth(9, 200); // Đề Tự luận
  sheet.setColumnWidth(10, 80); // Số từ TL
  sheet.setColumnWidth(11, 350); // Bài viết Tự luận
  sheet.setColumnWidth(12, 100); // Thời gian làm
  sheet.setColumnWidth(13, 90); // Rời Tab
  sheet.setColumnWidth(14, 130); // Quy chế
  sheet.setColumnWidth(15, 100); // Xếp loại
  sheet.setColumnWidth(16, 140); // Mã nộp
}

function formatDataRow(sheet, rowNum) {
  var range = sheet.getRange(rowNum, 1, 1, 16);
  range.setFontFamily("Arial").setFontSize(9.5).setVerticalAlignment("middle");
  range.setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(rowNum, 28);

  sheet.getRange(rowNum, 1).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 2).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 3).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 4).setHorizontalAlignment("left");
  sheet.getRange(rowNum, 5).setHorizontalAlignment("left").setFontColor("#64748b");
  sheet.getRange(rowNum, 6).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 7).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(rowNum, 8).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 9).setHorizontalAlignment("left");
  sheet.getRange(rowNum, 10).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 11).setHorizontalAlignment("left").setWrap(true);
  sheet.getRange(rowNum, 12).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 13).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 14).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 15).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 16).setHorizontalAlignment("center").setFontColor("#94a3b8");
}

// -----------------------------------------------------------------------------------------
// 5. QUẢN LÝ CẤU HÌNH ĐỀ THI (Tab 'CauHinh')
// -----------------------------------------------------------------------------------------
function getExamConfig(topicId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var durationDefault = (topicId === "P2_TRUYEN") ? 40 : 30;
  
  // 1. Tìm Sheet Cấu hình
  var sheet = ss.getSheetByName("CauHinh") ||
              ss.getSheetByName("Cấu hình") ||
              ss.getSheetByName("Cấu Hình") ||
              ss.getSheetByName("Cau_hinh") ||
              ss.getSheetByName("Config") ||
              ss.getSheetByName("CONFIG") ||
              ss.getSheetByName("Thiết lập") ||
              ss.getSheetByName("ThietLap");

  if (!sheet) {
    sheet = ss.insertSheet("CauHinh");
    sheet.appendRow(["Mã Chuyên Đề", "Trạng thái", "Mã ca thi", "Thời gian (phút)", "Lời dặn dò", "Mã PIN Giáo viên"]);
    sheet.appendRow(["P2_TRUYEN", "OPEN", "123456", 40, "Chúc các em làm bài thi thật tốt! Hãy phân bổ thời gian hợp lý.", "1358"]);
    sheet.appendRow(["P1_BAI_1", "OPEN", "123456", 30, "Chào mừng các em tham gia bài kiểm tra!", "1358"]);
    sheet.appendRow(["ALL", "OPEN", "123456", 40, "Hệ thống thi trực tuyến sẵn sàng.", "1358"]);
    sheet.getRange(1, 1, 1, 6).setBackground("#1e293b").setFontColor("#fde047").setFontWeight("bold");
  }

  var data = sheet.getDataRange().getValues();
  var pinVal = TEACHER_PIN_DEFAULT;
  var codeVal = EXAM_PASSCODE_DEFAULT;
  var durVal = durationDefault;
  var statusVal = "OPEN";
  var annVal = "Chúc các em làm bài thi thật tốt!";

  if (data && data.length > 0) {
    // A. KIỂM TRA DẠNG DỌC (Key-Value)
    var isVertical = false;
    for (var r = 0; r < data.length; r++) {
      var k = String(data[r][0] || "").toLowerCase().trim();
      if (k.indexOf("mã ca thi") !== -1 || k.indexOf("passcode") !== -1 || k.indexOf("mã pin") !== -1 || k.indexOf("mã gv") !== -1 || k.indexOf("giáo viên") !== -1) {
        isVertical = true;
        break;
      }
    }

    if (isVertical) {
      for (var r = 0; r < data.length; r++) {
        var k = String(data[r][0] || "").toLowerCase().trim();
        var v = String(data[r][1] !== undefined ? data[r][1] : "").trim();
        if (!v) continue;
        if (k.indexOf("passcode") !== -1 || k.indexOf("mã ca thi") !== -1 || k.indexOf("mã thi") !== -1 || k.indexOf("mật mã") !== -1 || k.indexOf("mật khẩu") !== -1) {
          codeVal = v;
        } else if (k.indexOf("pin") !== -1 || k.indexOf("mã gv") !== -1 || k.indexOf("giáo viên") !== -1 || k.indexOf("master") !== -1) {
          pinVal = v;
        } else if (k.indexOf("trạng thái") !== -1 || k.indexOf("status") !== -1 || k.indexOf("tình trạng") !== -1) {
          var vUp = v.toUpperCase();
          if (vUp === "CLOSED" || vUp === "LOCKED" || vUp === "KHÓA" || vUp === "KHOA" || vUp === "ĐÓNG" || vUp === "DONG" || vUp === "TẠM KHÓA" || vUp === "TAM KHOA") {
            statusVal = "CLOSED";
          } else {
            statusVal = "OPEN";
          }
        } else if (k.indexOf("thời gian") !== -1 || k.indexOf("duration") !== -1 || k.indexOf("phút") !== -1) {
          durVal = parseInt(v) || durationDefault;
        } else if (k.indexOf("thông báo") !== -1 || k.indexOf("announcement") !== -1 || k.indexOf("lời dặn") !== -1 || k.indexOf("dặn dò") !== -1) {
          annVal = v;
        }
      }
    } else {
      // B. KIỂM TRA DẠNG BẢNG NGANG
      var colTopic = 0, colStatus = -1, colCode = -1, colDur = -1, colAnn = -1, colPin = -1;
      
      if (data.length > 0) {
        var headers = data[0];
        for (var c = 0; c < headers.length; c++) {
          var h = String(headers[c] || "").toLowerCase().trim();
          if (h.indexOf("chuyên đề") !== -1 || h.indexOf("topic") !== -1 || h.indexOf("bài") !== -1) colTopic = c;
          else if (h.indexOf("trạng thái") !== -1 || h.indexOf("status") !== -1 || h.indexOf("tình trạng") !== -1) colStatus = c;
          else if (h.indexOf("mã ca thi") !== -1 || h.indexOf("passcode") !== -1 || h.indexOf("mã thi") !== -1 || h.indexOf("mật mã") !== -1 || h.indexOf("mật khẩu") !== -1) colCode = c;
          else if (h.indexOf("thời gian") !== -1 || h.indexOf("duration") !== -1 || h.indexOf("phút") !== -1) colDur = c;
          else if (h.indexOf("lời dặn") !== -1 || h.indexOf("thông báo") !== -1 || h.indexOf("announcement") !== -1 || h.indexOf("dặn dò") !== -1) colAnn = c;
          else if (h.indexOf("pin") !== -1 || h.indexOf("mã gv") !== -1 || h.indexOf("giáo viên") !== -1 || h.indexOf("giao vien") !== -1 || h.indexOf("master") !== -1) colPin = c;
        }
      }

      if (colStatus === -1) colStatus = 1;
      if (colCode === -1) colCode = 2;
      if (colDur === -1) colDur = 3;
      if (colAnn === -1) colAnn = 4;
      if (colPin === -1) colPin = 5;

      var targetRow = null;
      var cleanTopic = String(topicId || "").toLowerCase().trim();
      
      for (var r = 1; r < data.length; r++) {
        var t = String(data[r][colTopic] || "").toLowerCase().trim();
        if (t === cleanTopic || (cleanTopic.indexOf("truyen") !== -1 && t.indexOf("truyen") !== -1) || (cleanTopic.indexOf("bai_1") !== -1 && t.indexOf("bai_1") !== -1)) {
          targetRow = data[r];
          break;
        } else if (t === "all") {
          targetRow = data[r];
        }
      }

      if (!targetRow && data.length > 1) {
        targetRow = data[1];
      }

      if (targetRow) {
        if (targetRow[colStatus] !== undefined && String(targetRow[colStatus]).trim()) {
          var rawStatus = String(targetRow[colStatus]).trim().toUpperCase();
          if (rawStatus === "CLOSED" || rawStatus === "LOCKED" || rawStatus === "KHÓA" || rawStatus === "KHOA" || rawStatus === "ĐÓNG" || rawStatus === "DONG" || rawStatus === "TẠM KHÓA" || rawStatus === "TAM KHOA") {
            statusVal = "CLOSED";
          } else {
            statusVal = "OPEN";
          }
        }
        if (targetRow[colCode] !== undefined && String(targetRow[colCode]).trim()) {
          codeVal = String(targetRow[colCode]).trim();
        }
        if (targetRow[colDur] !== undefined && parseInt(targetRow[colDur])) {
          durVal = parseInt(targetRow[colDur]);
        }
        if (targetRow[colAnn] !== undefined && String(targetRow[colAnn]).trim()) {
          annVal = String(targetRow[colAnn]).trim();
        }
        if (targetRow[colPin] !== undefined && String(targetRow[colPin]).trim()) {
          pinVal = String(targetRow[colPin]).trim();
        }
      }

      // ƯU TIÊN MỞ THI: Nếu Thầy đổi bất kỳ dòng nào (hoặc dòng ALL) thành OPEN -> Mở thi ngay!
      for (var r2 = 1; r2 < data.length; r2++) {
        var checkSt = String(data[r2][colStatus] || "").toUpperCase().trim();
        if (checkSt === "OPEN" || checkSt === "MỞ" || checkSt === "MO" || checkSt === "ĐANG MỞ") {
          statusVal = "OPEN";
          break;
        }
      }
    }
  }

  // Bảo mật: Lời dặn không hiển thị mã PIN hoặc mã Ca thi
  if (annVal === pinVal || annVal === codeVal || (/^\d{3,6}$/.test(annVal) && annVal !== "2026")) {
    annVal = "Chúc các em làm bài thi thật tốt!";
  }

  return {
    status: "success",
    topicId: topicId,
    examStatus: statusVal,
    passcode: codeVal,
    examPasscode: codeVal,
    duration: durVal,
    announcement: annVal,
    teacherPin: pinVal,
    masterPin: pinVal,
    teacherEmail: "duyban86@gmail.com",
    serverTime: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss")
  };
}
// -----------------------------------------------------------------------------------------
// 6. QUẢN LÝ TÀI KHOẢN GOOGLE SSO (Tab 'TaiKhoan')
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
  for (var i = 0; i < data.length; i++) {
    var rowEmail = String(data[i][4] || "").trim().toLowerCase();
    if (rowEmail === email) {
      return {
        timestamp: data[i][0],
        lop: String(data[i][1]).trim(),
        stt: String(data[i][2]).trim(),
        name: String(data[i][3]).trim(),
        email: rowEmail,
        note: String(data[i][5] || "").trim()
      };
    }
  }
  return null;
}

function getAllAccounts() {
  var accSheet = getOrCreateAccountsSheet();
  var lastRow = accSheet.getLastRow();
  var accounts = [];
  if (lastRow >= 2) {
    var data = accSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][4]) {
        accounts.push({
          timestamp: data[i][0],
          lop: String(data[i][1]).trim(),
          stt: String(data[i][2]).trim(),
          name: String(data[i][3]).trim(),
          email: String(data[i][4]).trim().toLowerCase(),
          note: String(data[i][5] || "").trim()
        });
      }
    }
  }
  return accounts;
}

function bindGoogleAccount(payload) {
  var accSheet = getOrCreateAccountsSheet();
  var email = String(payload.email || "").trim().toLowerCase();
  var lop = String(payload.lop || "").trim();
  var stt = String(payload.stt || "").trim();
  var name = String(payload.name || "").trim();
  var note = String(payload.note || "Tự liên kết qua Google Sign-In").trim();
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

  if (!email || !lop || !stt || !name) {
    return { status: "error", message: "Thiếu thông tin bắt buộc (Email, Lớp, STT, Tên)!" };
  }

  var lastRowAcc = accSheet.getLastRow();
  var isUpdated = false;

  if (lastRowAcc >= 2) {
    var accData = accSheet.getRange(2, 1, lastRowAcc - 1, 6).getValues();
    for (var idx = 0; idx < accData.length; idx++) {
      var rowLop = String(accData[idx][1]).trim();
      var rowStt = String(accData[idx][2]).trim();
      var rowEmail = String(accData[idx][4]).trim().toLowerCase();

      if (rowLop === lop && rowStt === stt && rowEmail !== email && rowEmail !== "") {
        return {
          status: "error",
          message: "Học sinh [" + name + " - " + lop + "] đã được liên kết với email: " + rowEmail
        };
      }

      if (rowEmail === email) {
        var updateRow = idx + 2;
        accSheet.getRange(updateRow, 1).setValue(timestamp);
        accSheet.getRange(updateRow, 2).setValue(lop);
        accSheet.getRange(updateRow, 3).setValue(stt);
        accSheet.getRange(updateRow, 4).setValue(name);
        accSheet.getRange(updateRow, 6).setValue("Cập nhật lại: " + note);
        isUpdated = true;
        break;
      }
    }
  }

  if (!isUpdated) {
    accSheet.appendRow([timestamp, lop, stt, name, email, note]);
  }

  return {
    status: "success",
    message: "Liên kết tài khoản thành công cho: " + name + " (" + lop + ")",
    student: { lop: lop, stt: stt, name: name, email: email }
  };
}

// -----------------------------------------------------------------------------------------
// 7. LƯU BÁO LỖI / GÓP Ý (Tab 'GopY_BaoLoi')
// -----------------------------------------------------------------------------------------
function saveFeedback(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GopY_BaoLoi");
  if (!sheet) {
    sheet = ss.insertSheet("GopY_BaoLoi");
    sheet.appendRow(["Thời gian", "Lớp", "Họ và tên", "Chuyên đề / Bài", "Loại góp ý", "Nội dung phản hồi"]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#e0f2fe");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 400);
  }

  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
  sheet.appendRow([
    timestamp,
    d.studentClass || d.className || "",
    d.studentName || "",
    d.feedbackTopic || d.examTopic || "",
    d.feedbackType || "Góp ý",
    d.feedbackText || d.content || ""
  ]);

  return { status: "success", message: "Đã gửi ý kiến đóng góp thành công. Cảm ơn em!" };
}

// -----------------------------------------------------------------------------------------
// 8. HÀM TẠO NHANH CẤU TRÚC SHEETS (CHẠY 1 LẦN TRÊN TRÌNH SOẠN THẢO)
// -----------------------------------------------------------------------------------------
function setupFullSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet CauHinh
  var cfgSheet = ss.getSheetByName("CauHinh");
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet("CauHinh", 0);
    cfgSheet.appendRow(["Chuyên đề / TopicId", "Trạng thái (OPEN/CLOSED)", "Mã mở đề", "Thời gian (phút)", "Thông báo ca thi", "Mã PIN Giáo viên"]);
    cfgSheet.getRange("1:1").setFontWeight("bold").setBackground("#fef3c7");
    cfgSheet.appendRow(["P2_TRUYEN", "OPEN", "123456", 40, "Chúc các em làm bài tốt! Hãy phân bổ thời gian cho cả 30 câu trắc nghiệm và đoạn văn tự luận.", "1358"]);
    cfgSheet.appendRow(["P1_BAI_1", "OPEN", "123456", 30, "Chúc các em làm bài tốt!", "1358"]);
    cfgSheet.setFrozenRows(1);
  }

  // 2. Sheet Kết quả Chuyên đề Truyện
  var p2Sheet = ss.getSheetByName("KẾT QUẢ - CHUYÊN ĐỀ: THỂ LOẠI TRUYỆN");
  if (!p2Sheet) {
    p2Sheet = ss.insertSheet("KẾT QUẢ - CHUYÊN ĐỀ: THỂ LOẠI TRUYỆN");
    setupResultSheetHeaders(p2Sheet);
  }

  // 3. Sheet Kết quả Bài 1 Tiểu thuyết
  var p1Sheet = ss.getSheetByName("KẾT QUẢ - BÀI 1: KHẢ NĂNG LỚN LAO CỦA TIỂU THUYẾT");
  if (!p1Sheet) {
    p1Sheet = ss.insertSheet("KẾT QUẢ - BÀI 1: KHẢ NĂNG LỚN LAO CỦA TIỂU THUYẾT");
    setupResultSheetHeaders(p1Sheet);
  }

  // 4. Sheet TaiKhoan & GopY_BaoLoi
  getOrCreateAccountsSheet();
  saveFeedback({ studentName: "Hệ thống", className: "12", feedbackTopic: "Khởi tạo", feedbackText: "Khởi tạo hệ thống thành công!" });

  SpreadsheetApp.getUi().alert("✅ Đã thiết lập hoàn tất toàn bộ các bảng dữ liệu trên Google Sheets!");
}

// -----------------------------------------------------------------------------------------
// 9. HELPER GỬI RESPONSE
// -----------------------------------------------------------------------------------------
function sendResponse(data, callback) {
  var jsonStr = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonStr + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

function sendJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// -----------------------------------------------------------------------------------------
// 7. QUẢN LÝ NGÂN HÀNG CÂU HỎI TRỰC TUYẾN TỪ GOOGLE SHEETS (Tab 'CauHoi_...')
// -----------------------------------------------------------------------------------------
function getQuestionsFromSheet(topicId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "CauHoi_" + (topicId || "P2_TRUYEN");
  var sheet = ss.getSheetByName(sheetName) || ss.getSheetByName("CauHoi") || ss.getSheetByName("NganHangCauHoi");

  if (!sheet) {
    return {
      status: "fallback",
      message: "Chưa có Sheet câu hỏi, sử dụng ngân hàng mặc định",
      questions: []
    };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { status: "fallback", message: "Sheet câu hỏi trống", questions: [] };
  }

  var questions = [];
  // Cột: 0: STT, 1: Mức độ, 2: Loại câu, 3: Nội dung câu hỏi, 4: Phương án A, 5: Phương án B, 6: Phương án C, 7: Phương án D, 8: Đáp án đúng, 9: Giải thích chi tiết, 10: Tùy chọn mở rộng
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var qText = String(row[3] || "").trim();
    if (!qText) continue;

    var level = String(row[1] || "Thông hiểu").trim();
    var type = String(row[2] || "single_choice").trim();
    var optA = String(row[4] || "").trim();
    var optB = String(row[5] || "").trim();
    var optC = String(row[6] || "").trim();
    var optD = String(row[7] || "").trim();
    var ans = String(row[8] || "").trim();
    var exp = String(row[9] || "").trim();
    var extra = String(row[10] || "").trim();

    var options = [];
    if (optA) options.push({ key: "A", text: optA.replace(/^[A-D][.:\s]+/, "") });
    if (optB) options.push({ key: "B", text: optB.replace(/^[A-D][.:\s]+/, "") });
    if (optC) options.push({ key: "C", text: optC.replace(/^[A-D][.:\s]+/, "") });
    if (optD) options.push({ key: "D", text: optD.replace(/^[A-D][.:\s]+/, "") });

    var qItem = {
      id: r,
      level: level,
      type: type,
      text: qText,
      options: options,
      correctAnswer: ans,
      explanation: exp
    };

    // Nếu là dạng điền khuyết / dropdown
    if (type === "fill_blank" && extra) {
      qItem.dropdownOptions = extra.split(",").map(function(s) { return s.trim(); });
    }

    questions.push(qItem);
  }

  return {
    status: "success",
    topicId: topicId,
    total: questions.length,
    questions: questions
  };
}

function saveQuestionsToSheet(topicId, questionsList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "CauHoi_" + (topicId || "P2_TRUYEN");
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  sheet.appendRow([
    "STT", "Mức độ", "Loại câu", "Nội dung câu hỏi",
    "Phương án A", "Phương án B", "Phương án C", "Phương án D",
    "Đáp án đúng", "Giải thích chi tiết", "Tùy chọn mở rộng (Dropdown)"
  ]);

  sheet.getRange(1, 1, 1, 11).setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold");

  for (var i = 0; i < questionsList.length; i++) {
    var q = questionsList[i];
    var optA = "", optB = "", optC = "", optD = "";
    if (q.options && Array.isArray(q.options)) {
      if (q.options[0]) optA = typeof q.options[0] === 'string' ? q.options[0] : (q.options[0].text || "");
      if (q.options[1]) optB = typeof q.options[1] === 'string' ? q.options[1] : (q.options[1].text || "");
      if (q.options[2]) optC = typeof q.options[2] === 'string' ? q.options[2] : (q.options[2].text || "");
      if (q.options[3]) optD = typeof q.options[3] === 'string' ? q.options[3] : (q.options[3].text || "");
    }

    var extra = "";
    if (q.dropdownOptions && Array.isArray(q.dropdownOptions)) {
      extra = q.dropdownOptions.join(", ");
    }

    sheet.appendRow([
      (i + 1),
      q.level || "Thông hiểu",
      q.type || "single_choice",
      q.text || "",
      optA, optB, optC, optD,
      q.correctAnswer || "",
      q.explanation || "",
      extra
    ]);
  }

  sheet.setFrozenRows(1);
  return { status: "success", message: "Đã lưu " + questionsList.length + " câu hỏi lên Google Sheets thành công!" };
}
