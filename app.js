/**
 * AURA Harmony 降噪耳機 一頁式官方網站互動邏輯
 * 包含：動態商品規格切換、擬真信用卡 3D 預覽、3D Secure OTP 驗證、即時訂單收據與用戶評價系統
 */

// 產品規格資料設定
const PRODUCTS = {
  standard: {
    id: 'standard',
    name: 'AURA Harmony 普通版（煙燻粉）',
    price: 3990,
    desc: '標準主動降噪 / 40mm 單體 / 30h 續航 / 束口收納袋'
  },
  flagship: {
    id: 'flagship',
    name: 'AURA Harmony 高階版（煙燻粉）',
    price: 5990,
    desc: 'Ultra-ANC / LDAC 金標 / 40h 快充 / 贈磁吸硬殼盒+發燒線'
  }
};

let currentEdition = 'flagship';
let currentQty = 1;
let otpInterval = null;

document.addEventListener('DOMContentLoaded', () => {
  initNavbarAndScroll();
  initFaqAccordion();
  initReviewSystem();
  initCreditCardInputs();
});

/* ==========================================================================
   1. 導覽列滾動監聽與浮動購買列
   ========================================================================== */
function initNavbarAndScroll() {
  const floatingBar = document.getElementById('floatingBar');
  const mobileMenuBtn = document.getElementById('mobileMenuBtn');
  const navLinks = document.querySelector('.nav-links');

  window.addEventListener('scroll', () => {
    // 當滾動超過 450px 時顯示底部懸浮購買欄
    if (window.scrollY > 450) {
      floatingBar.classList.add('visible');
    } else {
      floatingBar.classList.remove('visible');
    }
  });

  // 手機版選單切換
  if (mobileMenuBtn && navLinks) {
    mobileMenuBtn.addEventListener('click', () => {
      const isVisible = navLinks.style.display === 'flex';
      navLinks.style.display = isVisible ? 'none' : 'flex';
      if (!isVisible) {
        navLinks.style.position = 'absolute';
        navLinks.style.top = '76px';
        navLinks.style.left = '0';
        navLinks.style.width = '100%';
        navLinks.style.background = '#ffffff';
        navLinks.style.flexDirection = 'column';
        navLinks.style.padding = '20px';
        navLinks.style.boxShadow = '0 10px 25px rgba(0,0,0,0.1)';
      }
    });

    // 點擊連結時自動收起手機版選單
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        if (window.innerWidth <= 768) {
          navLinks.style.display = 'none';
        }
      });
    });
  }
}

/* ==========================================================================
   2. FAQ 手風琴切換
   ========================================================================== */
function initFaqAccordion() {
  const faqItems = document.querySelectorAll('.faq-item');
  faqItems.forEach(item => {
    const questionBtn = item.querySelector('.faq-question');
    questionBtn.addEventListener('click', () => {
      const isActive = item.classList.contains('active');
      faqItems.forEach(otherItem => otherItem.classList.remove('active'));
      if (!isActive) {
        item.classList.add('active');
      }
    });
  });
}

/* ==========================================================================
   3. 用戶回饋與評價互動系統
   ========================================================================== */
function initReviewSystem() {
  // 分類標籤篩選
  const reviewTabs = document.querySelectorAll('.review-tab');
  const reviewCards = document.querySelectorAll('.review-card');

  reviewTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      reviewTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      const filter = tab.dataset.filter;
      reviewCards.forEach(card => {
        if (filter === 'all' || card.dataset.category === filter) {
          card.style.display = 'block';
        } else {
          card.style.display = 'none';
        }
      });
    });
  });

  // 評論表單展開與收起
  const openReviewBtn = document.getElementById('openReviewFormBtn');
  const closeReviewBtn = document.getElementById('closeReviewFormBtn');
  const reviewFormContainer = document.getElementById('reviewFormContainer');

  if (openReviewBtn && reviewFormContainer) {
    openReviewBtn.addEventListener('click', () => {
      reviewFormContainer.style.display = 'block';
      reviewFormContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    });
  }

  if (closeReviewBtn && reviewFormContainer) {
    closeReviewBtn.addEventListener('click', () => {
      reviewFormContainer.style.display = 'none';
    });
  }

  // 星級選擇器
  const starContainer = document.getElementById('starRatingSelect');
  const scoreText = document.getElementById('scoreText');
  const reviewScoreInput = document.getElementById('reviewScore');

  if (starContainer) {
    const stars = starContainer.querySelectorAll('.star-opt');
    const scoreDescriptions = {
      '5': '5 星（極致滿意）',
      '4': '4 星（非常滿意）',
      '3': '3 星（符合期待）',
      '2': '2 星（尚可改進）',
      '1': '1 星（不符合預期）'
    };

    stars.forEach(star => {
      star.addEventListener('click', () => {
        const score = star.dataset.score;
        reviewScoreInput.value = score;
        scoreText.textContent = scoreDescriptions[score];

        stars.forEach(s => {
          if (parseInt(s.dataset.score) <= parseInt(score)) {
            s.classList.add('active');
          } else {
            s.classList.remove('active');
          }
        });
      });
    });
  }

  // 提交新評價
  const newReviewForm = document.getElementById('newReviewForm');
  const reviewsGrid = document.getElementById('reviewsGrid');

  if (newReviewForm) {
    newReviewForm.addEventListener('submit', (e) => {
      e.preventDefault();

      const name = document.getElementById('reviewerName').value.trim();
      const edition = document.getElementById('reviewedEdition').value;
      const score = parseInt(document.getElementById('reviewScore').value) || 5;
      const title = document.getElementById('reviewTitle').value.trim();
      const comment = document.getElementById('reviewComment').value.trim();

      const isFlagship = edition.includes('高階版');
      const category = isFlagship ? 'flagship' : 'standard';
      const initial = name.charAt(0).toUpperCase();

      // 動態組裝星號字串
      let starsHtml = '';
      for (let i = 0; i < score; i++) starsHtml += '★';

      // 建立新的評價卡片並加入網格最上方
      const newCard = document.createElement('div');
      newCard.className = 'review-card';
      newCard.dataset.category = category;
      newCard.innerHTML = `
        <div class="review-user-info">
          <div class="user-avatar" style="background-color: #b2707b;">${initial}</div>
          <div class="user-meta">
            <div class="user-name-row">
              <strong>${escapeHtml(name)}</strong>
              <span class="verified-badge">✓ 已驗證購買</span>
            </div>
            <div class="edition-purchased">購買：${escapeHtml(edition)}</div>
          </div>
          <div class="review-date">剛剛發佈</div>
        </div>
        <div class="review-stars">${starsHtml}</div>
        <h4 class="review-card-title">「${escapeHtml(title)}」</h4>
        <p class="review-card-text">${escapeHtml(comment)}</p>
      `;

      reviewsGrid.prepend(newCard);

      // 重置並收起表單
      newReviewForm.reset();
      reviewFormContainer.style.display = 'none';

      // 友好提示
      alert('🎉 感謝您的熱情分享！您的評價心得已成功發佈。');
    });
  }
}

/* ==========================================================================
   4. 訂購彈窗與金額計算
   ========================================================================== */
function openOrderModal(editionKey = 'flagship') {
  currentEdition = editionKey;
  currentQty = 1;

  // 切換選取的 Radio
  const radioFlagship = document.querySelector('input[name="selectedEdition"][value="flagship"]');
  const radioStandard = document.querySelector('input[name="selectedEdition"][value="standard"]');

  if (editionKey === 'flagship') {
    radioFlagship.checked = true;
  } else {
    radioStandard.checked = true;
  }

  updateSelectedProduct();

  // 重置步驟至步驟 1
  document.getElementById('checkoutStepContainer').style.display = 'block';
  document.getElementById('otpSimulationStep').style.display = 'none';
  document.getElementById('orderSuccessStep').style.display = 'none';

  // 顯示 Modal
  const modalBackdrop = document.getElementById('orderModalBackdrop');
  modalBackdrop.classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeOrderModal() {
  const modalBackdrop = document.getElementById('orderModalBackdrop');
  modalBackdrop.classList.remove('open');
  document.body.style.overflow = '';
  if (otpInterval) clearInterval(otpInterval);
}

function updateSelectedProduct() {
  const selectedRadio = document.querySelector('input[name="selectedEdition"]:checked');
  currentEdition = selectedRadio ? selectedRadio.value : 'flagship';

  const optStandard = document.getElementById('optStandard');
  const optFlagship = document.getElementById('optFlagship');

  if (currentEdition === 'standard') {
    optStandard.classList.add('selected');
    optFlagship.classList.remove('selected');
  } else {
    optFlagship.classList.add('selected');
    optStandard.classList.remove('selected');
  }

  calculateOrderTotal();
}

function changeQty(delta) {
  currentQty = Math.max(1, Math.min(10, currentQty + delta));
  document.getElementById('orderQty').textContent = currentQty;
  calculateOrderTotal();
}

function calculateOrderTotal() {
  const prod = PRODUCTS[currentEdition];
  const total = prod.price * currentQty;
  const formattedTotal = `NT$ ${total.toLocaleString()}`;

  document.getElementById('orderTotalPreview').textContent = formattedTotal;
  document.getElementById('payBtnAmount').textContent = formattedTotal;
  document.getElementById('otpAmountDisplay').textContent = formattedTotal;
  return total;
}

/* ==========================================================================
   5. 擬真信用卡即時互動與 3D 翻轉效果
   ========================================================================== */
function initCreditCardInputs() {
  const ccNumberInput = document.getElementById('ccNumber');
  const ccExpInput = document.getElementById('ccExp');
  const ccCvcInput = document.getElementById('ccCvc');
  const ccNameInput = document.getElementById('ccName');

  const virtualCard = document.getElementById('virtualCard');
  const cardNumDisplay = document.getElementById('cardNumDisplay');
  const cardHolderDisplay = document.getElementById('cardHolderDisplay');
  const cardExpDisplay = document.getElementById('cardExpDisplay');
  const cardCvcDisplay = document.getElementById('cardCvcDisplay');
  const cardBrandLogo = document.getElementById('cardBrandLogo');
  const detectedBrandTag = document.getElementById('detectedBrandTag');

  // 卡號輸入格式化（每 4 碼空格）與發卡行辨識
  ccNumberInput.addEventListener('input', (e) => {
    let val = e.target.value.replace(/\D/g, '').substring(0, 16);
    let formatted = val.match(/.{1,4}/g)?.join(' ') || '';
    e.target.value = formatted;

    cardNumDisplay.textContent = formatted || '•••• •••• •••• ••••';

    // 辨識卡片組織
    if (/^4/.test(val)) {
      cardBrandLogo.textContent = 'VISA';
      detectedBrandTag.textContent = 'VISA';
    } else if (/^(5[1-5]|222[1-9]|22[3-9]|2[3-6]|27[01]|2720)/.test(val)) {
      cardBrandLogo.textContent = 'Mastercard';
      detectedBrandTag.textContent = 'Mastercard';
    } else if (/^35/.test(val)) {
      cardBrandLogo.textContent = 'JCB';
      detectedBrandTag.textContent = 'JCB';
    } else if (/^3[47]/.test(val)) {
      cardBrandLogo.textContent = 'AMEX';
      detectedBrandTag.textContent = 'AMEX';
    } else {
      cardBrandLogo.textContent = 'CARD';
      detectedBrandTag.textContent = 'VISA / MC';
    }
  });

  // 有效月年 MM/YY 格式化
  ccExpInput.addEventListener('input', (e) => {
    let val = e.target.value.replace(/\D/g, '').substring(0, 4);
    if (val.length >= 3) {
      val = val.substring(0, 2) + '/' + val.substring(2);
    }
    e.target.value = val;
    cardExpDisplay.textContent = val || 'MM/YY';
  });

  // 持卡人姓名
  ccNameInput.addEventListener('input', (e) => {
    const val = e.target.value.toUpperCase();
    cardHolderDisplay.textContent = val || 'YOUR NAME';
  });

  // 安全碼 CVC 及卡片 3D 翻轉效果
  ccCvcInput.addEventListener('input', (e) => {
    const val = e.target.value.replace(/\D/g, '').substring(0, 4);
    e.target.value = val;
    cardCvcDisplay.textContent = val ? '•'.repeat(val.length) : '•••';
  });

  ccCvcInput.addEventListener('focus', () => {
    virtualCard.classList.add('flipped');
  });

  ccCvcInput.addEventListener('blur', () => {
    virtualCard.classList.remove('flipped');
  });
}

/* ==========================================================================
   6. 結帳送出與 3D Secure OTP 驗證流程
   ========================================================================== */
function handleCheckoutSubmit(e) {
  e.preventDefault();

  const ccNumber = document.getElementById('ccNumber').value.replace(/\s/g, '');
  if (ccNumber.length < 15) {
    alert('請輸入正確的 16 碼信用卡號');
    document.getElementById('ccNumber').focus();
    return;
  }

  // 取得卡號末四碼遮罩
  const last4 = ccNumber.slice(-4);
  document.getElementById('otpCardMasked').textContent = `•••• •••• •••• ${last4}`;

  // 切換至 OTP 模擬畫面
  document.getElementById('checkoutStepContainer').style.display = 'none';
  document.getElementById('otpSimulationStep').style.display = 'block';

  // 啟動 60 秒倒數計時器
  let countdown = 60;
  const countdownEl = document.getElementById('otpCountdown');
  countdownEl.textContent = countdown;

  if (otpInterval) clearInterval(otpInterval);
  otpInterval = setInterval(() => {
    countdown--;
    countdownEl.textContent = countdown;
    if (countdown <= 0) {
      clearInterval(otpInterval);
      alert('驗證碼已逾時，請重新進行授權結帳。');
      cancelOtp();
    }
  }, 1000);
}

function cancelOtp() {
  if (otpInterval) clearInterval(otpInterval);
  document.getElementById('otpSimulationStep').style.display = 'none';
  document.getElementById('checkoutStepContainer').style.display = 'block';
}

function confirmOtpAndProcess() {
  const otpInput = document.getElementById('otpCodeInput').value.trim();
  if (otpInput.length < 6) {
    alert('請輸入 6 位數簡訊驗證碼（預設為 888888）');
    return;
  }

  const confirmBtn = document.getElementById('confirmOtpBtn');
  confirmBtn.disabled = true;
  confirmBtn.textContent = '金流中心授權中...';

  if (otpInterval) clearInterval(otpInterval);

  // 模擬金流銀行 1 秒驗證回傳
  setTimeout(() => {
    confirmBtn.disabled = false;
    confirmBtn.textContent = '確認送出驗證';

    // 填寫收據資料
    const randomOrderId = 'AUR-' + new Date().getFullYear() + String(Math.floor(100000 + Math.random() * 900000));
    const now = new Date();
    const orderDateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    const custName = document.getElementById('custName').value;
    const custPhone = document.getElementById('custPhone').value;
    const custAddress = document.getElementById('custAddress').value;
    const prod = PRODUCTS[currentEdition];
    const total = prod.price * currentQty;
    const ccNumber = document.getElementById('ccNumber').value.replace(/\s/g, '');
    const last4 = ccNumber.slice(-4) || '8888';

    document.getElementById('recOrderId').textContent = randomOrderId;
    document.getElementById('recOrderDate').textContent = orderDateStr;
    document.getElementById('recProductName').textContent = `${prod.name} × ${currentQty}`;
    document.getElementById('recPaymentMethod').textContent = `信用卡一次付清 (•••• ${last4})`;
    document.getElementById('recTotalAmount').textContent = `NT$ ${total.toLocaleString()}`;
    document.getElementById('recCustName').textContent = `${custName} (${custPhone})`;
    document.getElementById('recAddress').textContent = custAddress;

    // 切換顯示訂單成功步驟
    document.getElementById('otpSimulationStep').style.display = 'none';
    document.getElementById('orderSuccessStep').style.display = 'block';
  }, 900);
}

/* ==========================================================================
   工具函式
   ========================================================================== */
function escapeHtml(string) {
  const entityMap = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
    '/': '&#x2F;'
  };
  return String(string).replace(/[&<>"'/]/g, (s) => entityMap[s]);
}
