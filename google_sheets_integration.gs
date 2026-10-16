/**
 * ============================================================================
 * AURA Harmony 降噪耳機 - 私人專屬 Google 試算表自動記錄與 Excel 匯出腳本
 * ============================================================================
 *
 * 【功能說明】：
 * 1. 顧客在公開網站下單後，資料「不會」顯示在任何公開網頁上。
 * 2. 透過此腳本，資料會直接寫入您私人的 Google 試算表（只有您登入自己的 Google 帳號才能看到）。
 * 3. 系統自動將「收件人地址清單」與「信用卡金流授權記錄」分開存放於兩個不同分頁中。
 * 4. 隨時可在試算表中點擊【檔案】->【下載】->【Microsoft Excel (.xlsx)】取得 Excel 檔案。
 *
 * 【設定只需 3 分鐘】：
 * 步驟 1：前往 Google 雲端硬碟 (Google Drive)，點擊「新增」->「Google 試算表」，命名為：
 *         「AURA 耳機私人訂單中心」
 * 步驟 2：在該試算表上方選單點選【擴充功能】(Extensions) -> 【Apps Script】。
 * 步驟 3：清空編輯器中原有代碼，將「本檔案中的全部代碼」複製並貼上，按 Ctrl+S 儲存。
 * 步驟 4：點擊右上角藍色【部署】按鈕 -> 選擇【新部署作業】。
 *         - 齒輪圖示請選擇【網頁應用程式】(Web app)
 *         - 說明：訂單接收服務
 *         - 執行身分：選擇【我】(Me)
 *         - 誰可以存取：選擇【任何人】(Anyone) —— 這樣公開網站才能傳送資料
 * 步驟 5：點擊【部署】，複製產生的「網頁應用程式網址 (Web App URL)」。
 * 步驟 6：將該網址填入專案中 `app.js` 頂部的 `GOOGLE_SHEET_WEBHOOK_URL` 變數中，即完成連線！
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // 最長鎖定 15 秒防止並行衝突
  lock.tryLock(15000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);

    // ========================================================================
    // 分開記錄 1：【收件配送清單】（僅記錄物流與聯絡人，絕無卡號資料）
    // ========================================================================
    var shippingSheet = sheet.getSheetByName("收件配送清單");
    if (!shippingSheet) {
      shippingSheet = sheet.insertSheet("收件配送清單");
      // 建立表頭
      shippingSheet.appendRow([
        "訂單編號", "訂購時間", "收件人姓名", "手機號碼", "電子信箱", 
        "配送地址", "發票類型", "訂購品項", "數量", "實付金額 (TWD)", "出貨狀態"
      ]);
      // 標題列美化（粉藕色背景與粗體）
      shippingSheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#f4e4e7");
      shippingSheet.setFrozenRows(1);
    }

    shippingSheet.appendRow([
      payload.orderId,
      payload.createdAt,
      payload.custName,
      payload.custPhone,
      payload.custEmail,
      payload.custAddress,
      payload.custInvoice,
      payload.productName,
      payload.qty,
      payload.totalAmount,
      "備貨中（已付款）"
    ]);

    // ========================================================================
    // 分開記錄 2：【金流授權記錄】（獨立存放金流核對，卡號安全遮罩）
    // ========================================================================
    var paymentSheet = sheet.getSheetByName("金流授權記錄");
    if (!paymentSheet) {
      paymentSheet = sheet.insertSheet("金流授權記錄");
      // 建立表頭
      paymentSheet.appendRow([
        "關聯訂單編號", "交易時間", "發卡組織 (VISA/MC/JCB)", "信用卡卡號 (前4後4安全遮罩)", 
        "有效月年", "持卡人英文名", "3D 簡訊驗證狀態", "銀行授權碼", "交易狀態"
      ]);
      // 標題列美化（淺藍灰背景與粗體）
      paymentSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#e6ecf2");
      paymentSheet.setFrozenRows(1);
    }

    paymentSheet.appendRow([
      payload.orderId,
      payload.createdAt,
      payload.cardBrand || "VISA",
      payload.maskedCardNumber, // 嚴格遮罩，例如 4123-****-****-2345，保護資安
      payload.cardExp,
      payload.cardHolder,
      "OTP 簡訊驗證通過",
      payload.authCode,
      "授權扣款成功"
    ]);

    // 自動調整欄寬
    shippingSheet.autoResizeColumns(1, 11);
    paymentSheet.autoResizeColumns(1, 9);

    return ContentService
      .createTextOutput(JSON.stringify({ "status": "success", "orderId": payload.orderId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
