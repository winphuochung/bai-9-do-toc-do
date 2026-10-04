/**
 * GOOGLE APPS SCRIPT CHO BÀI 9: ĐO TỐC ĐỘ (KHTN 7)
 * Hướng dẫn:
 * 1. Mở một Google Sheet mới.
 * 2. Vào Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Dán toàn bộ mã bên dưới vào Code.gs.
 * 4. Chạy hàm setupSheet() một lần để tạo tiêu đề các cột.
 * 5. Chọn Triển khai (Deploy) -> Tùy chọn triển khai mới (New deployment) -> Loại: Ứng dụng web (Web app).
 * 6. Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone).
 * 7. Sao chép URL Web App nhận được và dán vào ứng dụng web học tập.
 */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("KetQuaHocSinh");
  if (!sheet) {
    sheet = ss.insertSheet("KetQuaHocSinh");
    const headers = [
      "Thời Gian Nộp", 
      "Họ Và Tên Học Sinh", 
      "Điểm Trắc Nghiệm", 
      "Số Cá Đã Câu", 
      "Tự Luận 1 (Thí nghiệm & Sai số)", 
      "Tự Luận 2 (Bảng 9.1)", 
      "Tự Luận 3 (Cổng quang vs Bấm giây)", 
      "Tự Luận 4 (Bắn tốc độ & An toàn)"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground("#0369a1")
         .setFontColor("#ffffff")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 300);
    sheet.setColumnWidth(6, 300);
    sheet.setColumnWidth(7, 300);
    sheet.setColumnWidth(8, 300);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("KetQuaHocSinh");
    if (!sheet) {
      sheet = ss.getActiveSheet();
    }

    const data = JSON.parse(e.postData.contents);
    
    sheet.appendRow([
      data.time || new Date().toLocaleString("vi-VN"),
      data.name,
      data.score,
      data.fish,
      data.essay1,
      data.essay2,
      data.essay3,
      data.essay4
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Đã lưu kết quả thành công" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
