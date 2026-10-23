/**
 * ============================================================================
 * AURA Harmony 降噪耳機 - Google 試算表後端資料庫整合腳本 (GAS Webhook)
 * ============================================================================
 *
 * 【功能說明】：
 * 1. 完整對應前台 HTML 與 JavaScript 結帳資訊，100% 欄位統一。
 * 2. 分頁 1【訂單管理總表】：自動彙整訂單編號、時間、商品規格、金額、收件人資料、
 *    發票資訊、訂單備註、安全金流授權記錄，並預留「物流單號」與「內部註記」供店家維護。
 * 3. 分頁 2【顧客評價收集表】：前台顧客送出產品評分與心得時，自動分流至此分頁。
 * 4. 隨時可在試算表中點擊【檔案】->【下載】->【Microsoft Excel (.xlsx)】匯出管理。
 *
 * 【設定只需 3 分鐘】：
 * 步驟 1：前往 Google 雲端硬碟 (Google Drive)，點擊「新增」->「Google 試算表」，命名為：
 *         「AURA Harmony 官方訂單中心」
 * 步驟 2：在該試算表上方選單點選【擴充功能】(Extensions) -> 【Apps Script】。
 * 步驟 3：清空編輯器中原有代碼，將「本檔案中的全部代碼」複製並貼上，按 Ctrl+S 儲存。
 * 步驟 4：點擊右上角藍色【部署】按鈕 -> 選擇【新部署作業】。
 *         - 齒輪圖示請選擇【網頁應用程式】(Web app)
 *         - 說明：AURA 訂單與評價接收服務
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);

    // ========================================================================
    // 分流 A：顧客評價心得收集表
    // ========================================================================
    if (payload.type === "review") {
      var reviewSheet = ss.getSheetByName("顧客評價收集表");
      if (!reviewSheet) {
        reviewSheet = ss.insertSheet("顧客評價收集表");
        var reviewHeaders = [
          "評價時間", "顧客暱稱", "體驗版本", "評分星級", "評價標題", "心得內容", "審核狀態"
        ];
        reviewSheet.appendRow(reviewHeaders);

        // 表頭樣式美化
        var rHeaderRange = reviewSheet.getRange(1, 1, 1, reviewHeaders.length);
        rHeaderRange.setFontWeight("bold")
                    .setBackground("#4a5568")
                    .setFontColor("#ffffff")
                    .setHorizontalAlignment("center");
        reviewSheet.setFrozenRows(1);
        reviewSheet.setRowHeight(1, 35);
      }

      reviewSheet.appendRow([
        payload.createdAt,
        payload.reviewerName,
        payload.reviewedEdition,
        payload.reviewScore,
        payload.reviewTitle,
        payload.reviewComment,
        "已上架公開"
      ]);

      reviewSheet.autoResizeColumns(1, 7);

      return ContentService
        .createTextOutput(JSON.stringify({ "status": "success", "type": "review" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ========================================================================
    // 分流 B：核心【訂單管理總表】（共 21 個精準欄位）
    // ========================================================================
    var orderSheet = ss.getSheetByName("訂單管理總表");
    if (!orderSheet) {
      orderSheet = ss.insertSheet("訂單管理總表");
      var orderHeaders = [
        "訂單編號", "下單時間", "訂單狀態",
        "商品版本", "商品規格與顏色", "購買數量", "實付金額 (TWD)",
        "收件人姓名", "手機號碼", "電子信箱", "配送地址", "訂單備註",
        "發票類型", "載具或統編",
        "發卡組織", "信用卡卡號 (安全遮罩)", "持卡人姓名", "金流授權碼", "金流驗證狀態",
        "物流追蹤單號", "內部管理註記"
      ];
      orderSheet.appendRow(orderHeaders);

      // 表頭美化：深藍灰底色 + 白字粗體 + 置中 + 凍結首列
      var headerRange = orderSheet.getRange(1, 1, 1, orderHeaders.length);
      headerRange.setFontWeight("bold")
                 .setBackground("#2c3e50")
                 .setFontColor("#ffffff")
                 .setHorizontalAlignment("center");
      orderSheet.setFrozenRows(1);
      orderSheet.setRowHeight(1, 36);
    }

    // 依序寫入欄位（完全與 app.js payload 鍵值 1:1 對齊）
    orderSheet.appendRow([
      payload.orderId,
      payload.createdAt,
      payload.orderStatus || "備貨中",
      payload.productEdition || "高階版",
      payload.productSpec || "AURA Harmony",
      payload.qty || 1,
      payload.totalAmount, // 純數字，便於 SUM 統計
      payload.custName,
      payload.custPhone,
      payload.custEmail,
      payload.custAddress,
      payload.orderNotes || "",
      payload.custInvoice,
      payload.invoiceCarrier || "",
      payload.cardBrand || "VISA",
      payload.maskedCardNumber, // 前4後4遮罩，符合資安規範
      payload.cardHolder,
      payload.authCode,
      payload.paymentStatus || "3D OTP 驗證成功",
      "", // 物流追蹤單號（留空供店家手動填寫）
      ""  // 內部管理註記（留空供店家手動填寫）
    ]);

    // 自動調整各欄寬度
    orderSheet.autoResizeColumns(1, 21);

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
