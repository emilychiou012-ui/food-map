// ===== 狀態 =====
let foods      = JSON.parse(localStorage.getItem('foodList')) || [];
let selectedId = null;
let editingId  = null;
let currentRating = 0;
let showAddRecordForm = false;

// ===== 區域選擇 =====
let selectedDistrict = '';

const DISTRICT_INFO = {
  '南港區': { tagline: '展覽商圈・南港車站' },
  '內湖區': { tagline: '內科商圈・碧湖公園' },
  '信義區': { tagline: '台北101・信義商圈' },
  '中山區': { tagline: '赤峰街・林森北路' },
};

function showLanding() {
  Object.keys(DISTRICT_INFO).forEach(district => {
    const count = foods.filter(f => f.address && f.address.includes(district)).length;
    const el = document.getElementById('count-' + district);
    if (el) el.textContent = count + ' 家餐廳';
  });
  const allEl = document.getElementById('count-all');
  if (allEl) allEl.textContent = foods.length;

  const overlay = document.getElementById('landingOverlay');
  overlay.style.display = 'flex';

  gsap.set(['.landing-logo', '.landing-title', '.landing-subtitle', '.landing-divider', '.landing-prompt', '.district-card'], { opacity: 0, y: 30 });
  const tl = gsap.timeline();
  tl.to(overlay, { opacity: 1, duration: 0.5, ease: 'power2.out' })
    .to('.landing-logo',     { opacity: 1, y: 0, duration: 0.6, ease: 'power3.out' }, '-=0.2')
    .to('.landing-title',    { opacity: 1, y: 0, duration: 0.7, ease: 'power3.out' }, '-=0.4')
    .to('.landing-subtitle', { opacity: 1, y: 0, duration: 0.5 }, '-=0.3')
    .to('.landing-divider',  { opacity: 1, y: 0, duration: 0.5, ease: 'power2.inOut' }, '-=0.2')
    .to('.landing-prompt',   { opacity: 1, y: 0, duration: 0.4 }, '-=0.2')
    .to('.district-card',    { opacity: 1, y: 0, duration: 0.5, stagger: 0.1, ease: 'power3.out' }, '-=0.2');
}

function selectDistrict(district) {
  selectedDistrict = district;
  const overlay = document.getElementById('landingOverlay');
  gsap.to(overlay, {
    opacity: 0, y: -15, duration: 0.4, ease: 'power2.in',
    onComplete: () => { overlay.style.display = 'none'; }
  });
  selectedId = null;
  document.getElementById('detailContent').innerHTML = '';
  document.getElementById('detailPlaceholder').style.display = '';
  renderList();
  renderStats();
  renderEmptyPanel();
}

function saveData() {
  localStorage.setItem('foodList', JSON.stringify(foods));
}

// ===== EXCEL 匯入 / 範本下載 =====
const EXCEL_COL_MAP = {
  '店名': 'name', '用餐時段': 'mealType', '適合場合': 'occasion',
  '飲食需求': 'dietary', '料理類型': 'category', '電話': 'phone',
  '人均消費最低': 'priceMin', '人均消費最高': 'priceMax',
  '營業時間': 'hours', '付款方式': 'payment', '地址': 'address',
  '評分': 'rating', '需要訂位': 'reservation', '取餐方式': 'visited',
  '餐廳網頁': 'restaurantUrl', '訂位網址': 'bookingUrl', '備註': 'notes'
};
const EXCEL_HEADERS = Object.keys(EXCEL_COL_MAP);

function importFromExcel(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const wb  = XLSX.read(e.target.result, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (!rows.length) { alert('❌ 試算表內沒有資料'); return; }

      const existingNames = new Set(foods.map(f => f.name));
      const now = new Date().toLocaleDateString('zh-TW');
      let added = 0, skipped = 0;

      rows.forEach((row, i) => {
        const name = String(row['店名'] || '').trim();
        if (!name) { skipped++; return; }
        if (existingNames.has(name)) { skipped++; return; }

        const entry = { id: Date.now() + i + 1, createdAt: now, nextBooking: null, bookingHistory: [] };
        for (const [zh, key] of Object.entries(EXCEL_COL_MAP)) {
          let val = row[zh] !== undefined ? String(row[zh]).trim() : '';
          if (key === 'payment' || key === 'visited') {
            entry[key] = val ? val.split(/[,，、]/).map(s => s.trim()).filter(Boolean) : [];
          } else if (key === 'rating') {
            entry[key] = val ? Math.min(5, Math.max(0, Number(val) || 0)) : 0;
          } else {
            entry[key] = val;
          }
        }
        if (!entry.reservation) entry.reservation = '不需要';
        foods.unshift(entry);
        existingNames.add(name);
        added++;
      });

      saveData();
      renderList();
      renderStats();
      renderEmptyPanel();
      showToast(`✅ 匯入完成：新增 ${added} 筆${skipped ? `，略過 ${skipped} 筆（重複或空白）` : ''}`);
    } catch (err) {
      alert('❌ 無法解析檔案，請確認為 .xlsx 或 .xls 格式\n' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function exportToExcel() {
  if (!foods.length) { alert('目前沒有餐廳資料可匯出'); return; }
  const rows = foods.map(f => ({
    '店名':       f.name || '',
    '用餐時段':   f.mealType || '',
    '適合場合':   f.occasion || '',
    '飲食需求':   f.dietary || '',
    '料理類型':   f.category || '',
    '電話':       f.phone || '',
    '人均消費最低': f.priceMin || '',
    '人均消費最高': f.priceMax || '',
    '營業時間':   f.hours || '',
    '付款方式':   Array.isArray(f.payment) ? f.payment.join('、') : (f.payment || ''),
    '地址':       f.address || '',
    '評分':       f.rating || 0,
    '需要訂位':   f.reservation || '',
    '取餐方式':   Array.isArray(f.visited) ? f.visited.join('、') : (f.visited || ''),
    '餐廳網頁':   f.restaurantUrl || '',
    '訂位網址':   f.bookingUrl || '',
    '備註':       f.notes || ''
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [6,8,8,8,8,14,10,10,26,20,24,4,8,16,30,30,30].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '餐廳清單');
  const date = new Date().toLocaleDateString('zh-TW').replace(/\//g, '-');
  XLSX.writeFile(wb, `YumMap_${date}.xlsx`);
  showToast(`✅ 已匯出 ${foods.length} 家餐廳`);
}

function downloadExcelTemplate() {
  const ws = XLSX.utils.aoa_to_sheet([
    EXCEL_HEADERS,
    ['永寶餐廳', '聚餐', '部門聚餐', '葷食為主', '台灣料理',
     '02-2885-7523', '450', '700', '週日及每月第一、三週週一公休',
     '現金', '台北市內湖區瑞光路106號', '5', '必須訂位',
     '自取', '', '', '熱炒合菜，包廂預訂制']
  ]);
  ws['!cols'] = EXCEL_HEADERS.map((h, i) =>
    ({ wch: [6,8,8,8,8,14,10,10,26,20,24,4,8,16,30,30,30][i] || 14 })
  );
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '餐廳清單');
  XLSX.writeFile(wb, 'YumMap_範本.xlsx');
}

// ===== Google Maps Takeout 匯入 =====
const DISTRICT_KEYWORDS = ['南港區','內湖區','信義區','中山區','大安區','松山區','士林區','北投區','萬華區','中正區','文山區','大同區','新店區','板橋區','三重區','新莊區'];

function detectDistrict(address) {
  if (!address) return '';
  for (const d of DISTRICT_KEYWORDS) {
    if (address.includes(d)) return d;
  }
  return '';
}

function importFromGoogleMaps(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = JSON.parse(e.target.result);

      let features = [];
      if (data.type === 'FeatureCollection' && Array.isArray(data.features)) {
        features = data.features;
      } else if (Array.isArray(data)) {
        features = data;
      } else {
        alert('❌ 無法識別格式，請確認是 Google Takeout 的 JSON 檔案');
        return;
      }

      const existingNames = new Set(foods.map(f => f.name));
      const now = new Date().toLocaleDateString('zh-TW');
      const districtSummary = {};
      let added = 0, skipped = 0, outsideTW = 0;

      features.forEach((feat, i) => {
        const props = feat.properties || feat;

        // 相容兩種格式：舊版 Location 物件 / 新版直接 name+address
        const loc     = props.Location || {};
        const name    = (loc['Business Name'] || props.Title || props.name || '').trim();
        const address = (loc.Address || props.address || '').trim();
        const phone   = (loc['Phone Number'] || props.phone || '').replace(/^\+886/, '0').trim();
        const gmapUrl = props['Google Maps URL'] || props['google_maps_url'] || props.url || '';

        if (!name) { skipped++; return; }

        // 只保留台灣地址（含「台北」「新北」「台灣」或無法判斷但有台灣行政區）
        const isTW = address.includes('台北') || address.includes('臺北') ||
                     address.includes('新北') || address.includes('台灣') ||
                     address.includes('臺灣') || detectDistrict(address) !== '';
        if (address && !isTW) { outsideTW++; return; }

        if (existingNames.has(name)) { skipped++; return; }

        const district = detectDistrict(address);

        const entry = {
          id: Date.now() + i + 1,
          name,
          address,
          phone,
          mealType: '',
          occasion: '',
          dietary: '',
          category: '',
          priceMin: '', priceMax: '',
          hours: '',
          payment: [],
          rating: 0,
          reservation: '不需要',
          visited: ['自取'],
          restaurantUrl: gmapUrl,
          bookingUrl: '',
          notes: '從 Google Maps 想去清單匯入',
          nextBooking: null,
          bookingHistory: [],
          createdAt: now
        };

        foods.push(entry);
        existingNames.add(name);
        districtSummary[district || '其他區域'] = (districtSummary[district || '其他區域'] || 0) + 1;
        added++;
      });

      saveData();
      renderList();
      renderStats();
      renderEmptyPanel();

      const summary = Object.entries(districtSummary)
        .sort((a, b) => b[1] - a[1])
        .map(([d, n]) => `${d} ${n} 家`).join('　');
      const outsideMsg = outsideTW ? `，略過非台灣地點 ${outsideTW} 筆` : '';
      const skipMsg    = skipped   ? `，重複/空白 ${skipped} 筆` : '';
      showToast(`📍 匯入完成：新增 ${added} 筆${outsideMsg}${skipMsg}　${summary}`);

    } catch (err) {
      alert('❌ 解析失敗，請確認是正確的 JSON 檔案\n' + err.message);
    }
  };
  reader.readAsText(file, 'utf-8');
}

function showToast(msg) {
  let t = document.getElementById('importToast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'importToast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.className = 'import-toast show';
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3500);
}

// ===== Modal 開關 =====
function openAddForm() {
  editingId = null;
  document.getElementById('modalTitle').textContent = '＋ 新增餐廳';
  document.getElementById('submitBtn').textContent = '加入清單';
  document.getElementById('submitBtn').classList.remove('btn-editing');
  resetForm();
  document.getElementById('modalOverlay').classList.add('show');
}

function openEditForm(id) {
  const food = foods.find(f => f.id === id);
  if (!food) return;
  editingId = id;
  document.getElementById('modalTitle').textContent = '✏️ 編輯餐廳';
  document.getElementById('submitBtn').textContent = '儲存變更';
  document.getElementById('submitBtn').classList.add('btn-editing');

  document.getElementById('name').value       = food.name || '';
  document.getElementById('mealType').value   = food.mealType || '';
  document.getElementById('occasion').value   = food.occasion || '';
  document.getElementById('dietary').value    = food.dietary || '';
  document.getElementById('category').value   = food.category || '';
  document.getElementById('phone').value      = food.phone || '';
  document.getElementById('address').value    = food.address || '';
  document.getElementById('hours').value      = food.hours || '';
  document.getElementById('priceMin').value   = food.priceMin || food.price || '';
  document.getElementById('priceMax').value   = food.priceMax || '';
  document.querySelectorAll('input[name="payment"]').forEach(cb => {
    cb.checked = (food.payment || []).includes(cb.value);
  });
  document.getElementById('restaurantUrl').value = food.restaurantUrl || '';
  document.getElementById('bookingUrl').value    = food.bookingUrl || '';
  document.getElementById('notes').value         = food.notes || '';
  document.getElementById('mrtStation').value    = food.mrtStation || '';
  document.getElementById('walkTime').value      = food.walkTime || '';

  // 結構化營業時間
  const oh = food.openingHours;
  if (oh) {
    if (oh.weekday && oh.weekday !== 'closed') {
      document.getElementById('wdOpen').value  = oh.weekday.open  || '';
      document.getElementById('wdClose').value = oh.weekday.close || '';
    } else if (oh.weekday === 'closed') {
      document.getElementById('wdClosed').checked = true;
    }
    if (oh.weekend && oh.weekend !== 'closed') {
      document.getElementById('weOpen').value  = oh.weekend.open  || '';
      document.getElementById('weClose').value = oh.weekend.close || '';
    } else if (oh.weekend === 'closed') {
      document.getElementById('weClosed').checked = true;
    }
  }

  currentRating = food.rating || 0;
  document.getElementById('rating').value = currentRating;
  highlightStars(currentRating);

  const resR = document.querySelector(`input[name="reservation"][value="${food.reservation}"]`);
  if (resR) resR.checked = true;
  const visitedArr = Array.isArray(food.visited) ? food.visited : [food.visited].filter(Boolean);
  document.querySelectorAll('input[name="visited"]').forEach(cb => {
    cb.checked = visitedArr.includes(cb.value);
  });

  document.getElementById('modalOverlay').classList.add('show');
}

function closeForm() {
  document.getElementById('modalOverlay').classList.remove('show');
  resetForm();
  editingId = null;
}

// 點 overlay 背景關閉
document.getElementById('modalOverlay').addEventListener('click', function(e) {
  if (e.target === this) closeForm();
});

// ===== 星星評分 =====
const stars      = document.querySelectorAll('.star');
const ratingInput = document.getElementById('rating');

stars.forEach(star => {
  star.addEventListener('mouseover', () => highlightStars(parseInt(star.dataset.value)));
  star.addEventListener('mouseleave', () => highlightStars(currentRating));
  star.addEventListener('click', () => {
    currentRating = parseInt(star.dataset.value);
    ratingInput.value = currentRating;
    highlightStars(currentRating);
  });
});

function highlightStars(count) {
  stars.forEach(s => s.classList.toggle('active', parseInt(s.dataset.value) <= count));
}

// ===== 新增 / 儲存 =====
document.getElementById('foodForm').addEventListener('submit', function(e) {
  e.preventDefault();

  const formData = {
    name:        document.getElementById('name').value.trim(),
    mealType:    document.getElementById('mealType').value,
    occasion:    document.getElementById('occasion').value,
    dietary:     document.getElementById('dietary').value,
    category:    document.getElementById('category').value,
    phone:       document.getElementById('phone').value.trim(),
    address:     document.getElementById('address').value.trim(),
    hours:       document.getElementById('hours').value.trim(),
    priceMin:    document.getElementById('priceMin').value,
    priceMax:    document.getElementById('priceMax').value,
    rating:      parseInt(document.getElementById('rating').value) || 0,
    reservation: document.querySelector('input[name="reservation"]:checked').value,
    visited:     [...document.querySelectorAll('input[name="visited"]:checked')].map(el => el.value),
    payment:       [...document.querySelectorAll('input[name="payment"]:checked')].map(el => el.value),
    restaurantUrl: document.getElementById('restaurantUrl').value.trim(),
    bookingUrl:    document.getElementById('bookingUrl').value.trim(),
    notes:         document.getElementById('notes').value.trim(),
    mrtStation:    document.getElementById('mrtStation').value.trim(),
    walkTime:      document.getElementById('walkTime').value ? parseInt(document.getElementById('walkTime').value) : null,
    openingHours:  (()=>{
      const wdO = document.getElementById('wdOpen').value;
      const wdC = document.getElementById('wdClose').value;
      const wdX = document.getElementById('wdClosed').checked;
      const weO = document.getElementById('weOpen').value;
      const weC = document.getElementById('weClose').value;
      const weX = document.getElementById('weClosed').checked;
      if (!wdO && !wdC && !weO && !weC) return null;
      return {
        weekday: wdX ? 'closed' : (wdO && wdC ? { open: wdO, close: wdC } : null),
        weekend: weX ? 'closed' : (weO && weC ? { open: weO, close: weC } : null),
      };
    })(),
  };

  if (editingId !== null) {
    const idx = foods.findIndex(f => f.id === editingId);
    if (idx !== -1) foods[idx] = { ...foods[idx], ...formData };
    if (selectedId === editingId) renderDetail(foods[idx]);
  } else {
    const newFood = {
      id: Date.now(), ...formData,
      nextBooking: null,
      bookingHistory: [],
      createdAt: new Date().toLocaleDateString('zh-TW')
    };
    foods.unshift(newFood);
    selectedId = newFood.id;
  }

  saveData();
  renderList();
  renderStats();
  closeForm();

  // 新增後自動顯示詳情
  if (editingId === null) {
    renderDetail(foods[0]);
  }
});

// ===== 刪除 =====
function deleteFood(id) {
  if (!confirm('確定要刪除這筆紀錄嗎？')) return;
  foods = foods.filter(f => f.id !== id);
  if (selectedId === id) {
    selectedId = null;
    document.getElementById('detailContent').innerHTML = '';
    document.getElementById('detailPlaceholder').style.display = '';
    renderEmptyPanel();
  }
  saveData();
  renderList();
  renderStats();
}

// ===== 選取餐廳（顯示詳情） =====
function selectFood(id) {
  selectedId = id;
  const food = foods.find(f => f.id === id);
  if (!food) return;
  renderList();
  renderDetail(food);
}

// ===== 詳情面板渲染 =====
function renderDetail(food) {
  document.getElementById('detailPlaceholder').style.display = 'none';
  const panel = document.getElementById('detailContent');

  const nextBookingHtml = food.nextBooking
    ? `<div class="next-booking-display">
        📅 <strong>${food.nextBooking.date}</strong>
        ${food.nextBooking.time ? ' ' + food.nextBooking.time : ''}
        ${food.nextBooking.people ? '・' + food.nextBooking.people + ' 人' : ''}
        <button class="btn-clear-booking" onclick="clearNextBooking(${food.id})">清除</button>
       </div>`
    : '';

  const historyHtml = (food.bookingHistory || []).length > 0
    ? food.bookingHistory.map(r => `
        <div class="booking-record">
          <div class="booking-record-info">
            <strong>${r.date}${r.time ? ' ' + r.time : ''}</strong>
            ${r.people ? `・${r.people} 人` : ''}
            ${r.occasion ? `・${r.occasion}` : ''}
            ${r.note ? `<br>📝 ${r.note}` : ''}
          </div>
          <button class="btn-del-record" onclick="deleteBookingRecord(${food.id}, ${r.id})">✕</button>
        </div>`).join('')
    : '<div style="font-size:0.82rem;color:#C4A882;padding:8px 0;">尚無訂位紀錄</div>';

  panel.innerHTML = `
    <div class="detail-top">
      <div class="detail-name-row">
        <div class="detail-name">${food.name}</div>
        <a class="btn-restaurant-inline" href="${food.restaurantUrl || 'https://www.google.com/search?q=' + encodeURIComponent(food.name + ' ' + (food.address || ''))}" target="_blank" rel="noopener">🔍 餐廳介紹</a>
      </div>
      <div class="detail-actions">
        <button class="btn-detail-edit" onclick="openEditForm(${food.id})">✏️ 編輯</button>
        <button class="btn-detail-delete" onclick="deleteFood(${food.id})">🗑️ 刪除</button>
      </div>
    </div>
    <div class="detail-body">

      <div class="detail-tags">
        ${food.mealType   ? `<span class="tag tag-meal">${food.mealType}</span>` : ''}
        ${food.occasion   ? `<span class="tag tag-occasion">${food.occasion}</span>` : ''}
        ${food.dietary    ? `<span class="tag ${dietaryClass(food.dietary)}">${food.dietary}</span>` : ''}
        ${food.category   ? `<span class="tag tag-category">${food.category}</span>` : ''}
        ${food.reservation !== '不需要' ? `<span class="tag tag-reservation-${food.reservation}">${food.reservation}</span>` : ''}
      </div>

      ${food.rating > 0 ? `<div class="detail-stars">${renderStars(food.rating)}</div>` : ''}

      <div class="detail-info">
        <div class="info-row">
          <span class="info-icon">📞 電話</span>
          ${food.phone
            ? `<a class="info-link" href="tel:${food.phone}">${food.phone}</a>
               <button class="btn-copy" onclick="copyText('${food.phone}', this)">複製</button>`
            : `<span class="info-empty">未填寫</span>`}
        </div>
        <div class="info-row">
          <span class="info-icon">🕐 時間</span>
          <div style="flex:1;padding:11px 14px;line-height:1.7">
            ${(()=>{
              const oh = food.openingHours;
              if (oh) {
                const fmtSlot = s => s && s !== 'closed' && s.open ? `${s.open} – ${s.close}` : '公休';
                const now = new Date(); const isWe = now.getDay() === 0 || now.getDay() === 6;
                const curSched = isWe ? oh.weekend : oh.weekday;
                const openNow = isOpenNow(food);
                const statusBadge = openNow === true  ? '<span class="badge-open" style="font-size:0.7rem">🟢 營業中</span>'
                                  : openNow === false ? '<span class="badge-closed" style="font-size:0.7rem">🔴 已打烊</span>'
                                  : '';
                return `<div style="display:flex;gap:8px;align-items:center;margin-bottom:4px">${statusBadge}</div>
                        <div class="hours-display-row"><span class="hours-display-label">平日</span><span>${fmtSlot(oh.weekday)}</span></div>
                        <div class="hours-display-row"><span class="hours-display-label">假日</span><span>${fmtSlot(oh.weekend)}</span></div>
                        ${food.hours ? `<div style="font-size:0.75rem;color:#bbb;margin-top:4px">${food.hours}</div>` : ''}`;
              }
              return food.hours ? `<span>${food.hours}</span>` : `<span class="info-empty">未填寫</span>`;
            })()}
          </div>
        </div>
        <div class="info-row">
          <span class="info-icon">💰 人均</span>
          ${formatPrice(food)
            ? `<span class="price-badge">${formatPrice(food)}</span>`
            : `<span class="info-empty">未填寫</span>`}
        </div>
        <div class="info-row">
          <span class="info-icon">💳 付款</span>
          ${food.payment && food.payment.length > 0
            ? `<div class="payment-tags">${food.payment.map(renderPaymentBadge).join('')}</div>`
            : `<span class="info-empty">未填寫</span>`}
        </div>
        <div class="info-row">
          <span class="info-icon">🛵 取餐</span>
          <div class="payment-tags">
            ${(Array.isArray(food.visited) ? food.visited : [food.visited]).filter(Boolean).map(v => `<span class="tag tag-status-${v}">${v}</span>`).join('')}
          </div>
        </div>
        <div class="info-row">
          <span class="info-icon">📍 地址</span>
          ${food.address
            ? `<span>${food.address}</span>
               <button class="btn-copy" onclick="copyText('${food.address}', this)">複製</button>`
            : `<span class="info-empty">未填寫</span>`}
        </div>
        ${(food.mrtStation || food.walkTime) ? `
        <div class="info-row">
          <span class="info-icon">🚇 捷運</span>
          <span style="padding:11px 14px;flex:1">
            ${food.mrtStation || ''}
            ${food.walkTime ? `<span class="detail-walk-chip">🚶 步行 ${food.walkTime} 分鐘</span>` : ''}
          </span>
        </div>` : ''}
        ${food.address ? `
          <div class="detail-map">
            <iframe
              src="https://maps.google.com/maps?q=${encodeURIComponent(food.address)}&output=embed&z=16"
              width="100%" height="200" frameborder="0"
              style="border:0;border-radius:8px;" allowfullscreen loading="lazy">
            </iframe>
          </div>` : ''}
      </div>

      ${food.bookingUrl ? `<a class="btn-booking-detail" href="${food.bookingUrl}" target="_blank" rel="noopener">📅 立即訂位</a>` : ''}
      ${food.notes ? `<div class="detail-notes">✨ 餐廳特色備註：${food.notes}</div>` : ''}

      <!-- 同事評論 -->
      <div class="detail-section reviews-section">
        <div class="reviews-header">
          <span class="detail-section-title">🍽️ 美食回饋區</span>
          ${(food.reviews||[]).length ? `<span class="reviews-count">${food.reviews.length} 則</span>` : ''}
        </div>

        <div class="reviews-list" id="reviewsList-${food.id}">
          ${(food.reviews||[]).length === 0 ? `<p class="reviews-empty">還沒有評論，成為第一個留言的人吧！</p>` :
            (food.reviews||[]).map(r => `
              <div class="review-card" id="review-${r.id}">
                <div class="review-avatar" style="background:${avatarColor(r.author)}">${r.author.charAt(0).toUpperCase()}</div>
                <div class="review-body">
                  <div class="review-meta">
                    <span class="review-author">${r.author}</span>
                    ${r.dept ? `<span class="review-dept">${r.dept}</span>` : ''}
                    <span class="review-stars">${'★'.repeat(r.rating || 0)}${'☆'.repeat(5 - (r.rating||0))}</span>
                    <span class="review-date">${r.date}</span>
                    <button class="review-delete" onclick="deleteReview(${food.id}, ${r.id})" title="刪除">✕</button>
                  </div>
                  <div class="review-text">${r.text.replace(/\n/g,'<br>')}</div>
                </div>
              </div>`).join('')
          }
        </div>

        <div class="review-form">
          <div class="review-form-top">
            <input class="review-input-name" id="reviewAuthor-${food.id}" placeholder="你的名字" maxlength="20" />
            <div class="review-star-pick" id="reviewStarPick-${food.id}" data-val="0">
              ${[1,2,3,4,5].map(n=>`<span class="rsp-star" data-n="${n}" onclick="setReviewStar(${food.id},${n})">☆</span>`).join('')}
            </div>
          </div>
          <textarea class="review-input-text" id="reviewText-${food.id}" placeholder="分享這家的用餐體驗、必點菜色或雷點…" rows="3"></textarea>
          <button class="btn-submit-review" onclick="addReview(${food.id})">送出評論</button>
        </div>
      </div>

    </div>
  `;
}

// ===== 評論功能 =====
const AVATAR_COLORS = ['#9A7B5A','#7A956B','#B5715A','#6B8A95','#A08B6E','#7B8F80'];
function avatarColor(name) {
  let h = 0; for (const c of name) h = (h * 31 + c.charCodeAt(0)) % AVATAR_COLORS.length;
  return AVATAR_COLORS[Math.abs(h)];
}

function setReviewStar(id, n) {
  const el = document.getElementById(`reviewStarPick-${id}`);
  if (!el) return;
  el.dataset.val = n;
  el.querySelectorAll('.rsp-star').forEach((s, i) => { s.textContent = i < n ? '★' : '☆'; });
}

function addReview(id) {
  const idx    = foods.findIndex(f => f.id === id);
  const author = (document.getElementById(`reviewAuthor-${id}`)?.value || '').trim();
  const text   = (document.getElementById(`reviewText-${id}`)?.value   || '').trim();
  const rating = parseInt(document.getElementById(`reviewStarPick-${id}`)?.dataset.val || '0');

  if (!author) { alert('請輸入你的名字'); return; }
  if (!text)   { alert('請輸入評論內容'); return; }

  if (!foods[idx].reviews) foods[idx].reviews = [];
  foods[idx].reviews.unshift({
    id: Date.now(),
    author, text, rating,
    date: new Date().toLocaleDateString('zh-TW')
  });
  saveData();
  renderDetail(foods[idx]);
}

function deleteReview(foodId, reviewId) {
  if (!confirm('確定刪除這則評論？')) return;
  const idx = foods.findIndex(f => f.id === foodId);
  foods[idx].reviews = (foods[idx].reviews || []).filter(r => r.id !== reviewId);
  saveData();
  renderDetail(foods[idx]);
}

// ===== 備註（黑名單 / 回購理由）=====
function saveRemarks(id) {
  const idx = foods.findIndex(f => f.id === id);
  foods[idx].remark = document.getElementById(`remark-${id}`)?.value.trim() || '';
  saveData();
}

// ===== 訂位紀錄 =====
function toggleRecordForm(id, show) {
  document.getElementById(`addRecordForm-${id}`).style.display = show ? 'block' : 'none';
  document.getElementById(`btnAddRecord-${id}`).style.display  = show ? 'none'  : 'block';
}

function saveBookingRecord(id) {
  const date    = document.getElementById(`rDate-${id}`).value;
  const time    = document.getElementById(`rTime-${id}`).value;
  const people  = document.getElementById(`rPeople-${id}`).value;
  const occasion= document.getElementById(`rOccasion-${id}`).value;
  const note    = document.getElementById(`rNote-${id}`).value.trim();
  if (!date) { alert('請選擇日期'); return; }

  const idx = foods.findIndex(f => f.id === id);
  if (!foods[idx].bookingHistory) foods[idx].bookingHistory = [];
  foods[idx].bookingHistory.unshift({ id: Date.now(), date, time, people, occasion, note });
  saveData();
  renderDetail(foods[idx]);
}

function deleteBookingRecord(foodId, recordId) {
  const idx = foods.findIndex(f => f.id === foodId);
  foods[idx].bookingHistory = foods[idx].bookingHistory.filter(r => r.id !== recordId);
  saveData();
  renderDetail(foods[idx]);
}

// ===== 複製到剪貼簿 =====
function copyText(text, btn) {
  navigator.clipboard.writeText(text).then(() => {
    if (btn) {
      btn.textContent = '已複製';
      btn.classList.add('copied');
      setTimeout(() => { btn.textContent = '複製'; btn.classList.remove('copied'); }, 1800);
    }
  });
}

// ===== 價位區間顯示 =====
function formatPrice(food) {
  const min = food.priceMin || food.price;
  const max = food.priceMax;
  if (!min) return null;
  return max ? `$${Number(min).toLocaleString()} - $${Number(max).toLocaleString()}` : `$${Number(min).toLocaleString()}`;
}

// ===== 飲食需求 CSS class 對應 =====
// ===== 右側空白面板：今日推薦 + 隨機抽籤 =====
const SVG_BOWL = `<svg width="80" height="80" viewBox="0 0 88 88" fill="none" xmlns="http://www.w3.org/2000/svg">
  <path d="M30 22 Q32 16 30 10" stroke="#C4A882" stroke-width="2" stroke-linecap="round" fill="none"/>
  <path d="M44 19 Q46 13 44 7"  stroke="#C4A882" stroke-width="2" stroke-linecap="round" fill="none"/>
  <path d="M58 22 Q60 16 58 10" stroke="#C4A882" stroke-width="2" stroke-linecap="round" fill="none"/>
  <path d="M16 42 Q16 70 44 70 Q72 70 72 42 Z" fill="#EDE0CE"/>
  <path d="M16 42 Q16 70 44 70 Q72 70 72 42" stroke="#C4A882" stroke-width="1.5" fill="none"/>
  <rect x="14" y="36" width="60" height="6" rx="3" fill="#9A7B5A" opacity="0.5"/>
  <ellipse cx="44" cy="74" rx="22" ry="3.5" fill="rgba(154,123,90,0.10)"/>
  <line x1="6"  y1="28" x2="6"  y2="78" stroke="#C4A882" stroke-width="2.5" stroke-linecap="round"/>
  <line x1="11" y1="28" x2="11" y2="78" stroke="#C4A882" stroke-width="2.5" stroke-linecap="round"/>
  <line x1="16" y1="28" x2="16" y2="78" stroke="#C4A882" stroke-width="2.5" stroke-linecap="round"/>
  <path d="M6 38 Q11 46 11 52 L11 78" stroke="#C4A882" stroke-width="2.5" stroke-linecap="round" fill="none"/>
  <path d="M82 28 L79 28 Q72 32 72 38 L82 38 Z" fill="#C4A882" opacity="0.65"/>
  <line x1="80" y1="38" x2="80" y2="78" stroke="#C4A882" stroke-width="2.5" stroke-linecap="round"/>
</svg>`;

function renderEmptyPanel() {
  const el = document.getElementById('detailPlaceholder');
  if (!el) return;

  const pool = foods.filter(f => !selectedDistrict || (f.address && f.address.includes(selectedDistrict)));
  const top  = pool.filter(f => f.rating === 5).slice(0, 3);

  const featuredHTML = top.length ? `
    <div style="margin-top:22px">
      <div class="featured-section-title">⭐ 今日推薦</div>
      ${top.map(f => {
        const open = isOpenNow(f);
        const badge = open === true  ? '<span class="badge-open" style="font-size:0.63rem">🟢 營業中</span>'
                    : open === false ? '<span class="badge-closed" style="font-size:0.63rem">🔴 已打烊</span>'
                    : '';
        return `
        <div class="featured-card" onclick="selectFood(${f.id})">
          <div>
            <div class="featured-name">${f.name}</div>
            <div style="display:flex;gap:5px;margin-top:5px;flex-wrap:wrap">
              ${f.mealType ? `<span class="tag tag-meal">${f.mealType}</span>` : ''}
              ${f.category ? `<span class="tag tag-category">${f.category}</span>` : ''}
            </div>
          </div>
          <div class="featured-meta">
            ${badge}
            ${formatPrice(f) ? `<span class="featured-price">${formatPrice(f)}</span>` : ''}
          </div>
        </div>`;
      }).join('')}
    </div>` : '';

  el.innerHTML = `
    <div class="empty-hero">
      <div class="placeholder-illustration">${SVG_BOWL}</div>
      <p class="placeholder-title">今天吃什麼？</p>
      <p class="placeholder-sub">點選左側餐廳查看詳情<br>或讓我們幫你決定</p>
      <button class="btn-random-pick" onclick="pickRandom()">🎲 隨機幫我選</button>
    </div>
    ${featuredHTML}
  `;
}

function pickRandom() {
  const pool = foods.filter(f => !selectedDistrict || (f.address && f.address.includes(selectedDistrict)));
  if (!pool.length) return;
  const food = pool[Math.floor(Math.random() * pool.length)];
  selectFood(food.id);
}

// ===== 台灣支付 badge =====
const PAYMENT_BADGE_MAP = {
  '現金':     { cls: 'pay-cash',     icon: '💵' },
  'Line Pay': { cls: 'pay-linepay',  icon: '💚' },
  '街口支付': { cls: 'pay-jko',      icon: '🔵' },
  'Apple Pay':{ cls: 'pay-apple',    icon: '🍎' },
  '台灣Pay':  { cls: 'pay-twpay',    icon: '🇹🇼' },
  '信用卡':   { cls: 'pay-card',     icon: '💳' },
  'Visa':     { cls: 'pay-card',     icon: '💳' },
  'Mastercard':{ cls: 'pay-card',    icon: '💳' },
  'Jcb':      { cls: 'pay-card',     icon: '💳' },
  'Amex':     { cls: 'pay-card',     icon: '💳' },
};
function renderPaymentBadge(p) {
  const info = PAYMENT_BADGE_MAP[p] || { cls: 'pay-card', icon: '💳' };
  return `<span class="pay-badge ${info.cls}">${info.icon} ${p}</span>`;
}

// ===== 模擬員工名單 =====
const MOCK_EMPLOYEES = [
  { name: 'Emily', dept: 'Design 團隊' },
  { name: 'Eric',  dept: 'Line 團隊'   },
  { name: 'Jane',  dept: 'HR 團隊'     },
  { name: 'Sam',   dept: 'RD 團隊'     },
  { name: 'Tina',  dept: 'PM 團隊'     },
  { name: 'Kevin', dept: 'BD 團隊'     },
];

// ===== 營業狀態判斷 =====
function isOpenNow(food) {
  const now   = new Date();
  const today = now.getDay();                            // 0=Sun
  const cur   = now.getHours() * 60 + now.getMinutes();

  // 優先使用結構化 openingHours
  if (food.openingHours) {
    const oh = food.openingHours;
    const isWeekend = today === 0 || today === 6;
    const schedule  = isWeekend ? oh.weekend : oh.weekday;
    if (!schedule || schedule === 'closed') return false;
    if (schedule.open && schedule.close) {
      const [oh2, om] = schedule.open.split(':').map(Number);
      const [ch,  cm] = schedule.close.split(':').map(Number);
      return cur >= (oh2 * 60 + om) && cur <= (ch * 60 + cm);
    }
    return null;
  }

  if (!food.hours) return null;
  const text  = food.hours;
  const D = { '日':0,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6 };

  function slots(str) {
    return [...str.matchAll(/(\d{1,2}):(\d{2})\s*[-–~～]\s*(\d{1,2}):(\d{2})/g)]
      .map(m => [+m[1]*60 + +m[2], +m[3]*60 + +m[4]]);
  }

  function daySet(str) {
    if (/每[日天]/.test(str)) return new Set([0,1,2,3,4,5,6]);
    const s = new Set();
    str.replace(/(?:週)?([一二三四五六日])[至\-–](?:週)?([一二三四五六日])/g, (_, a, b) => {
      if (D[a] === undefined || D[b] === undefined) return;
      const dist = (D[b] - D[a] + 7) % 7;
      for (let k = 0; k <= dist; k++) s.add((D[a] + k) % 7);
    });
    str.replace(/週([一二三四五六日]+)/g, (_, chars) => {
      for (const c of chars) if (D[c] !== undefined) s.add(D[c]);
    });
    return s;
  }

  const closed = new Set();
  for (const [, d] of text.matchAll(/週([一二三四五六日])公休/g)) closed.add(D[d]);
  if (closed.has(today)) return false;

  let applicable = null;
  for (const part of text.split(/[，,]/)) {
    if (/公休/.test(part)) continue;
    const s = slots(part);
    if (!s.length) continue;
    if (/每[日天]|週[一二三四五六日]/.test(part)) {
      if (daySet(part).has(today)) { applicable = s; break; }
    } else {
      applicable = s;
    }
  }

  if (!applicable) return null;
  return applicable.some(([s, e]) => cur >= s && cur <= e);
}

function dietaryClass(dietary) {
  const map = { '葷食為主': 'tag-dietary-meat', '素食友善': 'tag-dietary-veg', '全素 / 純素': 'tag-dietary-vegan' };
  return map[dietary] || 'tag-dietary-other';
}

// ===== 星星顯示 =====
function renderStars(rating) {
  let html = '';
  for (let i = 1; i <= 5; i++) {
    html += `<span class="${i <= rating ? 'filled' : 'empty'}">★</span>`;
  }
  return html;
}

// ===== 頂部統計 =====
function renderStats() {
  const displayFoods = selectedDistrict
    ? foods.filter(f => f.address && f.address.includes(selectedDistrict))
    : foods;
  const total    = displayFoods.length;
  const topRated = displayFoods.filter(f => f.rating === 5).length;
  const districtLabel = selectedDistrict
    ? `<div class="stat-item"><span class="stat-num" style="font-size:0.85rem;letter-spacing:0.04em;">${selectedDistrict}</span><span class="stat-label">目前區域</span></div>`
    : '';
  document.getElementById('headerStats').innerHTML = `
    ${districtLabel}
    <div class="stat-item"><span class="stat-num">${total}</span><span class="stat-label">總餐廳數</span></div>
    <div class="stat-item"><span class="stat-num">${topRated}</span><span class="stat-label">最熱門 ⭐</span></div>
  `;
}

// ===== 清單渲染 =====
function renderList() {
  const search         = document.getElementById('searchInput').value.toLowerCase();
  const mealFilter     = document.getElementById('filterMealType').value;
  const occasionFilter = document.getElementById('filterOccasion').value;
  const visitFilter    = document.getElementById('filterVisited').value;
  const dietaryFilter  = document.getElementById('filterDietary').value;
  const walkFilter     = parseInt(document.getElementById('filterWalk').value) || 0;
  const maxBudget      = parseInt(document.getElementById('filterBudget').value);
  const budgetActive   = maxBudget < 2000;

  const filtered = foods.filter(f => {
    const matchDistrict = !selectedDistrict || (f.address && f.address.includes(selectedDistrict));
    const matchSearch   = f.name.toLowerCase().includes(search);
    const matchMeal     = !mealFilter     || f.mealType === mealFilter;
    const matchOccasion = !occasionFilter || f.occasion === occasionFilter;
    const matchVisit    = !visitFilter    || (Array.isArray(f.visited) ? f.visited.includes(visitFilter) : f.visited === visitFilter);
    const matchDietary  = !dietaryFilter  || f.dietary  === dietaryFilter;
    const matchBudget   = !budgetActive   || !f.priceMin || parseInt(f.priceMin) <= maxBudget;
    const matchWalk     = !walkFilter     || (f.walkTime && parseInt(f.walkTime) <= walkFilter);
    return matchDistrict && matchSearch && matchMeal && matchOccasion && matchVisit && matchDietary && matchBudget && matchWalk;
  });

  const list  = document.getElementById('foodList');
  const empty = document.getElementById('emptyState');
  document.getElementById('foodCount').textContent = `共 ${filtered.length} 筆`;

  if (filtered.length === 0) {
    list.innerHTML = '';
    empty.classList.add('show');
    return;
  }
  empty.classList.remove('show');

  list.innerHTML = filtered.map(food => {
    const openStatus = isOpenNow(food);
    const statusBadge = openStatus === true  ? '<span class="badge-open">🟢 營業中</span>'
                      : openStatus === false ? '<span class="badge-closed">🔴 已打烊</span>'
                      : '';
    return `
    <div class="food-card ${selectedId === food.id ? 'selected' : ''}" onclick="selectFood(${food.id})">
      <div class="card-row1">
        <div class="card-name">${food.name}</div>
        <div class="card-right-top">
          ${statusBadge}
          <div class="card-stars-sm">${'★'.repeat(food.rating)}${'☆'.repeat(5 - food.rating)}</div>
        </div>
      </div>
      <div class="card-tags">
        ${food.mealType   ? `<span class="tag tag-meal">${food.mealType}</span>` : ''}
        ${food.occasion   ? `<span class="tag tag-occasion">${food.occasion}</span>` : ''}
        ${food.dietary    ? `<span class="tag ${dietaryClass(food.dietary)}">${food.dietary}</span>` : ''}
        ${food.category   ? `<span class="tag tag-category">${food.category}</span>` : ''}
        ${(Array.isArray(food.visited) ? food.visited : [food.visited]).map(v => `<span class="tag tag-status-${v}">${v}</span>`).join('')}
        ${food.reservation !== '不需要' ? `<span class="tag tag-reservation-${food.reservation}">${food.reservation}</span>` : ''}
      </div>
      <div class="card-quick">
        ${food.phone      ? `<span class="card-quick-phone">📞 ${food.phone}</span>` : ''}
        ${formatPrice(food) ? `<span>💰 ${formatPrice(food)}</span>` : ''}
      </div>
      ${(food.mrtStation || food.walkTime) ? `
        <div class="card-mrt">
          🚇 ${food.mrtStation || ''}
          ${food.walkTime ? `<span class="card-walk">🚶 ${food.walkTime} 分鐘</span>` : ''}
        </div>` : ''}
      ${food.remark ? `<div class="card-remark">📝 ${food.remark}</div>` : ''}
    </div>`;
  }).join('');
}

// ===== 一鍵匯入 =====
function handleImport() {
  if (!confirm('確定要重新同步預設餐廳清單嗎？（已存在的店名不會重複）')) return;
  importRestaurants();
  foods = JSON.parse(localStorage.getItem('foodList')) || [];
  renderList();
  renderStats();
  document.getElementById('importBanner').style.display = 'none';
  alert('✅ 匯入成功！共 13 家餐廳已加入清單。');
}

// ===== 篩選監聽 =====
document.getElementById('searchInput').addEventListener('input', renderList);
document.getElementById('filterMealType').addEventListener('change', renderList);
document.getElementById('filterOccasion').addEventListener('change', renderList);
document.getElementById('filterVisited').addEventListener('change', renderList);
document.getElementById('filterDietary').addEventListener('change', renderList);
document.getElementById('filterBudget').addEventListener('input', function () {
  const v   = parseInt(this.value);
  const pct = ((v / 2000) * 100).toFixed(1) + '%';
  this.style.background = `linear-gradient(to right, var(--gold) 0%, var(--gold) ${pct}, var(--border) ${pct})`;
  document.getElementById('budgetVal').textContent = v >= 2000 ? '不限' : `$${v} 以下`;
  renderList();
});

// ===== 篩選監聽（步行時間）=====
document.getElementById('filterWalk').addEventListener('change', renderList);

// ===== 表單重置 =====
function resetForm() {
  document.getElementById('foodForm').reset();
  currentRating = 0;
  highlightStars(0);
  ratingInput.value = 0;
  ['wdOpen','wdClose','weOpen','weClose'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  ['wdClosed','weClosed'].forEach(id => { const el = document.getElementById(id); if (el) el.checked = false; });
}

// ===== 初始化 =====
// 每次都執行，importRestaurants 內部已做防重複，只會新增缺少的店
importRestaurants();
foods = JSON.parse(localStorage.getItem('foodList')) || [];
document.getElementById('importBanner').style.display = 'none';

// 遷移外送方式：字串 → 陣列 + 更新舊標籤 + 移除廢棄值
(function migrateVisited() {
  const labelMap  = { '僅限自取': '自取', '可外送': '外送', 'Uber Eats': 'Ubereats', 'Foodpanda': 'Foodpanda' };
  const removeSet = new Set(['想去', '常去', '已去過']);
  let changed = false;
  foods = foods.map(f => {
    const arr = Array.isArray(f.visited) ? f.visited : [f.visited].filter(Boolean);
    const updated = arr
      .filter(v => !removeSet.has(v))
      .map(v => labelMap[v] || v);
    if (updated.join() !== arr.join() || !Array.isArray(f.visited)) {
      changed = true;
      return { ...f, visited: updated };
    }
    return f;
  });
  if (changed) saveData();
})();

// 遷移舊付款方式標籤
(function migratePaymentLabels() {
  const map = { 'JCB': 'Jcb', 'AE Card': 'Amex' };
  let changed = false;
  foods = foods.map(f => {
    if (!f.payment) return f;
    const updated = f.payment.map(p => map[p] || p);
    if (updated.join() !== f.payment.join()) { changed = true; return { ...f, payment: updated }; }
    return f;
  });
  if (changed) saveData();
})();

// 去除重複（依店名判斷）
(function deduplicateFoods() {
  const seen = new Set();
  const deduped = foods.filter(f => {
    if (seen.has(f.name)) return false;
    seen.add(f.name);
    return true;
  });
  if (deduped.length !== foods.length) {
    foods = deduped;
    localStorage.setItem('foodList', JSON.stringify(foods));
  }
})();

// 補齊已存入餐廳的營業時間、電話、地址
(function patchRestaurantData() {
  const patches = {
    '輪輪の町':          { hours: '週日-週四 11:00-21:30，週五六 11:00-22:00', phone: '02-2389-8118' },
    'Preserve LaLaPort 南港店': { hours: '週日-週四 11:00-21:30，週五六 11:00-22:00', phone: '02-2786-3757', address: '台北市南港區經貿二路131號1樓（南港 LaLaPort）' },
    '午葉 Leaves Cafe':  { hours: '每日 09:00-20:30', phone: '02-2659-0112', address: '台北市內湖區瑞光路168號1樓' },
    'Omnivore 雜食者':   { hours: '週一、三-日 09:00-17:00，週二公休', phone: '070-1018-2730', address: '台北市內湖區金湖路57號' },
    'Fika Fika Cafe 內湖店': { hours: '週二-週日 10:00-18:00，週一公休', phone: '02-2656-0133', address: '台北市內湖區陽光街321巷40號' },
    'MIACUCINA 內湖店':  { hours: '週日-週四 11:00-21:30，週五六 11:00-22:00', phone: '02-2659-3918', address: '台北市內湖區瑞光路601號' },
    '欣葉小聚 南港店':   { hours: '每日 11:00-14:30 / 17:00-21:00', phone: '02-2785-1819', address: '台北市南港區經貿二路166號A棟1樓' },
    '欣葉鐘菜 劍南路':   { hours: '每日 11:30-15:00 / 17:00-21:30', phone: '02-2532-7373', address: '台北市中山區樂群二路199號1樓（台北萬豪酒店，捷運劍南路站）' },
  };
  let changed = false;
  foods = foods.map(f => {
    const p = patches[f.name];
    if (!p) return f;
    const updated = { ...f };
    if (p.hours   && !f.hours)   { updated.hours   = p.hours;   changed = true; }
    if (p.phone   && !f.phone)   { updated.phone   = p.phone;   changed = true; }
    if (p.address && (!f.address || f.address.length < 10)) { updated.address = p.address; changed = true; }
    return updated;
  });
  if (changed) saveData();
})();

renderStats();
renderList();
renderEmptyPanel();

// 顯示進場 Landing Overlay（資料載入後再更新數量）
showLanding();
