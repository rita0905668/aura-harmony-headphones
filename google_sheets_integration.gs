/**
 * ============================================================================
 * AURA Harmony 降噪耳機 - Google 試算表後端資料庫整合腳本 (GAS Webhook)
 * 包含：自動寫入試算表資料庫 + 顧客訂單成立 HTML 自動確認信發送
 * ============================================================================
 *
 * 【功能說明】：
 * 1. 完整對應前台 HTML 與 JavaScript 結帳資訊，100% 欄位統一。
 * 2. 分頁 1【訂單管理總表】：自動彙整訂單編號、時間、商品規格、金額、收件人資料、
 *    發票資訊、訂單備註、安全金流授權記錄，並預留「物流單號」與「內部註記」供店家維護。
 * 3. 📧【自動寄送訂單確認信】：顧客下單成功後，系統自動寄出專屬品牌 HTML 訂購憑證信件至顧客信箱。
 * 4. 分頁 2【顧客評價收集表】：前台顧客送出產品評分與心得時，自動分流至此分頁。
 * 5. 隨時可在試算表中點擊【檔案】->【下載】->【Microsoft Excel (.xlsx)】匯出管理。
 *
 * 【如何更新已部署的 Apps Script】：
 * 步驟 1：打開您的 Google 試算表 ➜ 點選【擴充功能】(Extensions) ➜ 【Apps Script】。
 * 步驟 2：清空編輯器中現有代碼，將「本檔案中的全部代碼」全選貼上，按 Ctrl+S 儲存。
 * 步驟 3：點擊右上角藍色【部署】按鈕 ➜ 選擇【管理部署作業】(Manage deployments)。
 * 步驟 4：點擊右側鉛筆 ✏️ 圖示（編輯） ➜ 「版本」選擇【建立新版本】 ➜ 點擊【部署】。
 *         （若 Google 彈出權限審查，點選「審查權限」➜ 選擇您的 Google 帳號 ➜ 點選進階 ➜ 允許寄件即可）
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

    // ========================================================================
    // 分流 C：寄送顧客訂單成立確認信 (HTML 品牌信件)
    // ========================================================================
    try {
      sendOrderConfirmationEmail(payload);
    } catch (emailErr) {
      // 避免郵件配額或信箱格式問題影響訂單記錄主流程
      console.log("郵件發送失敗但訂單已安全寫入試算表: " + emailErr.toString());
    }

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

/**
 * 寄送美化訂單成立確認信至顧客 Email
 */
function sendOrderConfirmationEmail(payload) {
  if (!payload.custEmail || payload.custEmail.indexOf("@") === -1) {
    return;
  }

  var formattedAmount = "NT$ " + Number(payload.totalAmount || 0).toLocaleString();
  var carrierInfo = payload.invoiceCarrier ? "（" + payload.invoiceCarrier + "）" : "";
  var notesInfo = payload.orderNotes ? payload.orderNotes : "無特殊備註";

  var htmlBody = 
    '<div style="max-width: 600px; margin: 0 auto; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; color: #2d3748; line-height: 1.6; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; background: #ffffff;">' +
      '<!-- 品牌頂部 -->' +
      '<div style="background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%); padding: 32px 24px; text-align: center; color: #ffffff;">' +
        '<h1 style="margin: 0; font-size: 24px; letter-spacing: 2px; font-weight: 700;">AURA SOUND</h1>' +
        '<p style="margin: 8px 0 0 0; font-size: 14px; color: #cbd5e0; letter-spacing: 1px;">訂單成立通知 ｜ 原廠正品保證</p>' +
      '</div>' +

      '<!-- 親切問候 -->' +
      '<div style="padding: 28px 24px 20px 24px;">' +
        '<h2 style="font-size: 18px; color: #1a202c; margin: 0 0 12px 0;">親愛的 ' + escapeHtml(payload.custName) + ' 您好：</h2>' +
        '<p style="margin: 0; font-size: 15px; color: #4a5568;">' +
          '非常感謝您購買 <strong>AURA Harmony 降噪耳機</strong>！您的訂單已成功完成信用卡線上付款授權，倉庫團隊正全力為您備貨中。以下為您的訂購明細收據：' +
        '</p>' +
      '</div>' +

      '<!-- 訂單明細收據卡片 -->' +
      '<div style="margin: 0 24px 24px 24px; background: #f7fafc; border: 1px solid #edf2f7; border-radius: 8px; padding: 20px;">' +
        '<table style="width: 100%; border-collapse: collapse; font-size: 14px;">' +
          '<tr>' +
            '<td style="padding: 8px 0; color: #718096; width: 35%;">訂單編號</td>' +
            '<td style="padding: 8px 0; color: #2d3748; font-weight: 700;">' + escapeHtml(payload.orderId) + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 8px 0; color: #718096;">訂購時間</td>' +
            '<td style="padding: 8px 0; color: #2d3748;">' + escapeHtml(payload.createdAt) + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 8px 0; color: #718096;">訂購品項</td>' +
            '<td style="padding: 8px 0; color: #2b6cb0; font-weight: 600;">' + escapeHtml(payload.productSpec) + ' × ' + escapeHtml(payload.qty) + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 8px 0; color: #718096;">付款方式</td>' +
            '<td style="padding: 8px 0; color: #2d3748;">信用卡安全付款（' + escapeHtml(payload.cardBrand) + ' ' + escapeHtml(payload.maskedCardNumber) + '）</td>' +
          '</tr>' +
          '<tr style="border-top: 1px dashed #cbd5e0; border-bottom: 1px dashed #cbd5e0;">' +
            '<td style="padding: 12px 0; color: #718096; font-size: 15px; font-weight: 600;">實付總額</td>' +
            '<td style="padding: 12px 0; color: #b2707b; font-size: 18px; font-weight: 800;">' + formattedAmount + ' <span style="font-size: 12px; color: #718096; font-weight: normal;">(全台免運費)</span></td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 10px 0 6px 0; color: #718096;">收件人</td>' +
            '<td style="padding: 10px 0 6px 0; color: #2d3748;">' + escapeHtml(payload.custName) + ' (' + escapeHtml(payload.custPhone) + ')</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; color: #718096;">配送地址</td>' +
            '<td style="padding: 6px 0; color: #2d3748;">' + escapeHtml(payload.custAddress) + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; color: #718096;">發票資訊</td>' +
            '<td style="padding: 6px 0; color: #2d3748;">' + escapeHtml(payload.custInvoice) + escapeHtml(carrierInfo) + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; color: #718096;">訂單備註</td>' +
            '<td style="padding: 6px 0; color: #4a5568;">' + escapeHtml(notesInfo) + '</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +

      '<!-- 配送與權益提醒 -->' +
      '<div style="padding: 0 24px 28px 24px; font-size: 13px; color: #718096; line-height: 1.7;">' +
        '<div style="background: #fffaf0; border-left: 4px solid #dd6b20; padding: 12px 16px; border-radius: 4px; margin-bottom: 16px;">' +
          '<strong style="color: #c05621;">📦 出貨與配送時程：</strong><br>' +
          '商品預計將於 1~2 個工作天內依訂單順序出貨（遇週末或國定假日順延）。商品出貨後將會發送出貨通知與物流單號至您的手機簡訊。' +
        '</div>' +
        '<p style="margin: 0 0 8px 0;">🛡️ <strong>享有權益</strong>：本商品享有原廠正品保證、2 年原廠保固與 7 天鑑賞期。</p>' +
        '<p style="margin: 0;">若您對訂單有任何疑問或需變更配送資料，歡迎直接回覆此電子郵件，我們將有專人為您服務。</p>' +
      '</div>' +

      '<!-- 頁尾 -->' +
      '<div style="background: #edf2f7; padding: 16px 24px; text-align: center; font-size: 12px; color: #a0aec0;">' +
        '© 2026 AURA SOUND Inc. All rights reserved. ｜ 傾聽心境，粉霧純粹' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: payload.custEmail,
    subject: "【AURA SOUND】訂單成立通知 - 訂單編號：" + payload.orderId,
    htmlBody: htmlBody,
    name: "AURA SOUND 官方旗艦店"
  });
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
