let FIELDS = [];          // [{name, kind, format}]
let LECTURES = {};        // {ten_bai: {truong: giatri}}
let LECTURE_ORDER = [];
let CURRENT = null;       // tên bài giảng đang chọn
let SELECTED_FOR_EXPORT = new Set();  // các bài giảng đang tick để đưa vào "Xuất tất cả"
let TAB_FILTER = '';                  // chuỗi đang gõ để lọc nhanh tab bài giảng
const TAB_FILTER_THRESHOLD = 6;        // chỉ hiện ô tìm kiếm khi có nhiều hơn ngần này bài giảng

// Công thức kiểm tra tổng thời gian (đơn vị: phút). Nếu tổng vế phải không
// khớp vế trái thì báo cho người soạn kiểm tra lại, không chặn thao tác.
const TIME_CHECKS = [
  {
    total: 'Thời gian thực hành giảng bài',
    parts: ['Thời gian thủ tục giảng bài', 'Thời gian nội dung bài giảng', 'Thời gian kết thúc giảng bài'],
  },
  {
    total: 'Thời gian nội dung bài giảng',
    parts: ['Thời gian mở đầu', 'Thời gian phần nội dung', 'Thời gian kết luận'],
  },
];

function shouldAutoPunctuate(fieldName) {
  // Bỏ qua các trường thông tin ngắn, tên riêng, thời gian
  const noPunctWords = [
    'Tên', 'Họ', 'Cấp bậc', 'Chức vụ', 'Ngày', 'Giờ', 
    'Năm', 'Đối tượng', 'Bản số', 'Địa điểm', 'Thời gian '
  ];
  return !noPunctWords.some(w => fieldName.includes(w));
}

function autoPunctuate(text) {
  const lines = text.split('\n').map((l) => l.trim()).filter((l) => l.length > 0);
  if (!lines.length) return text;
  const multiline = lines.length > 1;
  return lines.map((line, i) => {
    const isLast = i === lines.length - 1;
    let body = line
      .replace(/^[-–—+*>]\s*/, '')     // bỏ gạch/cộng/sao đầu dòng cũ
      .replace(/[;,.:]+\s*$/, '');  // bỏ dấu kết thúc cũ
    if (multiline) body = '- ' + body;
    return body + (isLast ? '.' : ';');
  }).join('\n');
}

const $ = (sel) => document.querySelector(sel);

// ---------- Chú thích hướng dẫn (tooltip) ----------
function applyHints() {
  document.querySelectorAll('[data-tip]').forEach((el) => {
    el.title = el.dataset.tip;
  });
}
applyHints();

$('#btnHints').addEventListener('click', () => {
  const driverObj = window.driver.js.driver({
    showProgress: true,
    steps: [
      { element: '#btnDownloadTemplate', popover: { title: '1. Tải file mẫu', description: 'Tải file Excel mẫu được định dạng chuẩn, chỉ cho phép nhập dữ liệu.', side: 'bottom' } },
      { element: '#dropXlsx', popover: { title: '2. Nạp file Excel', description: 'Bấm vào đây để chọn (hoặc kéo thả) file Excel dữ liệu. Phần mềm sẽ tự động nạp ngay sau khi chọn.', side: 'right' } },
      { element: '#fieldForm', popover: { title: '3. Chỉnh sửa nội dung', description: 'Bạn có thể chỉnh sửa trực tiếp nội dung bài giảng tại đây. Dữ liệu sẽ không thay đổi ở file gốc.', side: 'right' } },
      { element: '.panel-preview', popover: { title: '4. Xem trước PDF', description: 'Bản xem trước của tài liệu sẽ hiển thị và tự động cập nhật ở đây mỗi khi bạn gõ văn bản.', side: 'left' } },
      { element: '.form-actions', popover: { title: '5. Xuất file', description: 'Tải bài giảng về máy dưới dạng file Word (.docx) hoặc xuất hàng loạt ra file nén (.zip).', side: 'top' } },
    ]
  });
  driverObj.drive();
});

function applyDataset(data, statusMsg) {
  FIELDS = data.fields;
  LECTURES = data.lectures;

  for (const lectureName in LECTURES) {
    const lectureData = LECTURES[lectureName];
    for (const fieldName in lectureData) {
      if (shouldAutoPunctuate(fieldName) && lectureData[fieldName]) {
        lectureData[fieldName] = autoPunctuate(String(lectureData[fieldName]));
      }
    }
  }

  LECTURE_ORDER = data.lecture_order;
  CURRENT = LECTURE_ORDER[0];
  SELECTED_FOR_EXPORT = new Set(LECTURE_ORDER);
  TAB_FILTER = '';
  if ($('#tabFilter')) $('#tabFilter').value = '';

  renderWarnings(data.warnings);
  renderTabs();
  renderForm();
  if ($('#btnExportOne')) $('#btnExportOne').disabled = false;
  if ($('#btnDirectEdit')) $('#btnDirectEdit').disabled = false;
  updateExportAllLabel();
  $('#statusBox').textContent = statusMsg || `Đã nạp ${LECTURE_ORDER.length} bài giảng · ${FIELDS.length} trường`;
  $('#statusBox').classList.add('ready');

  updatePreview();
}

// ---------- Toast Notification System ----------
function showToast(message, type = 'info', duration = 3000) {
  const container = $('#toastContainer') || document.body;
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  let icon = 'ℹ️';
  if (type === 'success') icon = '✅';
  if (type === 'warning') icon = '⚠️';

  toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ---------- Auto-Session Restore ----------
function saveWorkspaceSession() {
  try {
    const sessionData = {
      FIELDS,
      LECTURES,
      LECTURE_ORDER,
      CURRENT,
      SELECTED_FOR_EXPORT: Array.from(SELECTED_FOR_EXPORT)
    };
    localStorage.setItem('APPBGM_SESSION_V16', JSON.stringify(sessionData));
  } catch (e) {
    console.error('Save session error:', e);
  }
}

function restoreWorkspaceSession() {
  try {
    const raw = localStorage.getItem('APPBGM_SESSION_V16');
    if (!raw) return false;
    const data = JSON.parse(raw);
    if (!data || !data.LECTURE_ORDER || !data.LECTURE_ORDER.length || !data.LECTURES) return false;

    if (data.FIELDS) FIELDS = data.FIELDS;
    LECTURES = data.LECTURES;
    LECTURE_ORDER = data.LECTURE_ORDER;
    CURRENT = data.CURRENT && LECTURES[data.CURRENT] ? data.CURRENT : LECTURE_ORDER[0];
    SELECTED_FOR_EXPORT = new Set(data.SELECTED_FOR_EXPORT || LECTURE_ORDER);

    renderTabs();
    renderForm();
    if ($('#btnExportOne')) $('#btnExportOne').disabled = false;
    if ($('#btnDirectEdit')) $('#btnDirectEdit').disabled = false;
    updateExportAllLabel();
    $('#statusBox').textContent = `Đã khôi phục ${LECTURE_ORDER.length} bài giảng · ${FIELDS.length} trường`;
    $('#statusBox').classList.add('ready');
    updatePreview();

    showToast('Đã tự động khôi phục phiên làm việc gần nhất!', 'success');
    return true;
  } catch (e) {
    console.error('Restore session error:', e);
    return false;
  }
}

// ---------- Phím Tắt Nhanh (Keyboard Shortcuts) ----------
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && !e.altKey) {
    const key = e.key.toLowerCase();
    if (key === 's') {
      e.preventDefault();
      if ($('#btnExportData')) $('#btnExportData').click();
    } else if (key === 'n') {
      e.preventDefault();
      if ($('#btnAddLecture')) $('#btnAddLecture').click();
    } else if (key === 'd') {
      e.preventDefault();
      if ($('#btnDuplicateLecture')) $('#btnDuplicateLecture').click();
    }
  }
});

// ---------- Tự động nạp dữ liệu mặc định hoặc khôi phục phiên ----------
window.addEventListener('DOMContentLoaded', async () => {
  const restored = restoreWorkspaceSession();
  if (restored) return;

  try {
    const res = await fetch('/api/init_default_data');
    if (res.ok) {
      const data = await res.json();
      applyDataset(data, 'Đã sẵn sàng chỉnh sửa dữ liệu bài giảng');
    }
  } catch (e) {
    console.log('Init default data notice:', e);
  }
});

// ---------- Chọn file ----------
if ($('#dropXlsx')) {
  $('#dropXlsx').addEventListener('click', async () => {
    try {
      const res = await fetch('/api/open_excel_native', { method: 'POST' });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.error || 'Lỗi không xác định', 'warning');
        return;
      }

      if (!data.success) return;

      if ($('#dropXlsx')) $('#dropXlsx').classList.add('filled');
      applyDataset(data, `Đã nạp ${data.lecture_order.length} bài giảng từ file Excel`);
      showToast(`Đã nạp thành công ${data.lecture_order.length} bài giảng!`, 'success');
      saveWorkspaceSession();
    } catch (e) {
      showToast('Lỗi kết nối: ' + e, 'warning');
    }
  });
}

// ---------- Tabs bài giảng ----------
function renderTabs() {
  const wrap = $('#lectureTabs');
  if (!wrap) return;

  wrap.innerHTML = '';

  LECTURE_ORDER.forEach((name) => {
    const t = document.createElement('div');
    t.className = 'tab' + (name === CURRENT ? ' active' : '');
    t.dataset.tip = `Bấm để chuyển sang xem/sửa bài giảng "${name}"`;

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'tab-check';
    cb.checked = SELECTED_FOR_EXPORT.has(name);
    cb.title = 'Tick để đưa bài này vào "Xuất tất cả"';
    cb.addEventListener('click', (e) => {
      e.stopPropagation();
      if (cb.checked) SELECTED_FOR_EXPORT.add(name); else SELECTED_FOR_EXPORT.delete(name);
      updateExportAllLabel();
      saveWorkspaceSession();
    });
    t.appendChild(cb);

    const label = document.createElement('span');
    label.textContent = name;
    t.appendChild(label);

    t.addEventListener('click', () => {
      saveCurrentFormIntoState();
      CURRENT = name;
      renderTabs();
      renderForm();
      updatePreview();
      saveWorkspaceSession();
    });
    wrap.appendChild(t);
  });

  if (!LECTURE_ORDER.length) {
    const empty = document.createElement('div');
    empty.className = 'tab-empty';
    empty.textContent = 'Chưa có bài giảng nào.';
    wrap.appendChild(empty);
  }
  applyHints();
}

function updateExportAllLabel() {
  const btn = $('#btnExportAll');
  const total = LECTURE_ORDER.length;
  const n = SELECTED_FOR_EXPORT.size;
  if (total === 0) {
    btn.textContent = 'Xuất tất cả (.zip)';
    btn.disabled = true;
    return;
  }
  btn.textContent = (n === total) ? 'Xuất tất cả (.zip)' : `Xuất ${n} bài đã chọn (.zip)`;
  btn.disabled = (n === 0);
}

const SECTION_DEFINITIONS = [
  { title: '🏛️ I. THÔNG TIN HÀNH CHÍNH & PHÊ DUYỆT', firstField: 'Tên khoa' },
  { title: '🎯 II. MỤC ĐÍCH & YÊU CẦU BÀI GIẢNG', firstField: 'Mục đích' },
  { title: '⏱️ III. NỘI DUNG & PHÂN BỔ THỜI GIAN', firstField: 'Nội dung bài giảng' },
  { title: '🏫 IV. TỔ CHỨC, PHƯƠNG PHÁP & VẬT CHẤT', firstField: 'Tổ chức lớp học' },
  { title: '📝 V. CHI TIẾT CÁC PHẦN GIẢNG BÀI', firstField: 'Thời gian thực hành giảng bài' }
];

function isFullWidthField(name) {
  const longFields = [
    'Tên bài giảng',
    'Cấp bậc, học vị, họ tên giảng viên biên soạn',
    'Cấp bậc, học vị, họ tên người thông qua',
    'Cấp bậc, học vị, họ tên người phê duyệt',
    'Nhận xét phần nội dung bài giảng', 'Nhận xét phần thực hành giảng bài', 'Kết luận phê duyệt',
    'Mục đích', 'Yêu cầu về kiến thức', 'Yêu cầu về kỹ năng', 'Mức tự chủ và trách nhiệm',
    'Nội dung bài giảng', 'Phương pháp của giảng viên', 'Phương pháp của học viên',
    'Vật chất bảo đảm của giảng viên', 'Vật chất bảo đảm của học viên',
    'Nội dung mở đầu bài giảng', 'Nội dung kết luận bài giảng', 'Nội dung kết thúc giảng bài'
  ];
  return longFields.includes(name);
}

// ---------- Form dữ liệu ----------
function renderForm() {
  const wrap = $('#fieldForm');
  wrap.innerHTML = '';
  const values = LECTURES[CURRENT] || {};
  const rowByField = {};

  let currentSecDiv = null;
  let currentGridDiv = null;

  let currentSecIdx = -1;
  FIELDS.forEach((f) => {
    const secDefIdx = SECTION_DEFINITIONS.findIndex(s => s.firstField === f.name);
    if (secDefIdx !== -1) {
      currentSecIdx = secDefIdx;
      const secDiv = document.createElement('div');
      secDiv.className = 'form-section';

      // Đặt mặc định Nhóm I (index 0) là Thu gọn (collapsed)
      if (secDefIdx === 0) {
        secDiv.classList.add('collapsed');
      }

      const secHeader = document.createElement('div');
      secHeader.className = 'form-section-header';

      if (secDefIdx === 0) {
        // Nhóm I: Thêm 1 Nút Duy Nhất ở Header để áp dụng toàn bộ TT Hành Chính cho tất cả bài giảng!
        secHeader.innerHTML = `
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="sec-title">${SECTION_DEFINITIONS[secDefIdx].title}</span>
            <span class="sec-arrow">▼</span>
          </div>
          <button type="button" class="btn btn-xs btn-outline btn-sync-sec1" style="background:#fff; color:var(--navy); font-size:11px; padding:2px 8px; font-weight:600;" title="Sao chép toàn bộ Thông Tin Hành Chính sang tất cả các bài giảng">⏩ Đồng bộ TT Hành Chính cho tất cả bài</button>
        `;

        setTimeout(() => {
          const btnSyncSec = secHeader.querySelector('.btn-sync-sec1');
          if (btnSyncSec) {
            btnSyncSec.addEventListener('click', (e) => {
              e.preventDefault();
              e.stopPropagation();
              if (!CURRENT || !LECTURES[CURRENT]) return;

              const sec1FieldNames = [];
              for (const fieldObj of FIELDS) {
                if (fieldObj.name === SECTION_DEFINITIONS[1].firstField) break;
                sec1FieldNames.push(fieldObj.name);
              }

              for (const lName in LECTURES) {
                sec1FieldNames.forEach((fn) => {
                  if (fn !== 'Tên bài giảng') {
                    LECTURES[lName][fn] = LECTURES[CURRENT][fn] || '';
                  }
                });
              }
              alert(`Đã áp dụng toàn bộ Thông Tin Hành Chính cho tất cả ${LECTURE_ORDER.length} bài giảng!`);
              queuePreview();
            });
          }
        }, 0);

      } else {
        secHeader.innerHTML = `
          <span class="sec-title">${SECTION_DEFINITIONS[secDefIdx].title}</span>
          <span class="sec-arrow">▼</span>
        `;
      }

      // Independent Section Collapsible Toggle Handler
      secHeader.addEventListener('click', (e) => {
        if (e.target.classList.contains('btn-sync-sec1')) return;
        secDiv.classList.toggle('collapsed');
      });

      secDiv.appendChild(secHeader);

      currentGridDiv = document.createElement('div');
      currentGridDiv.className = 'field-group-grid';
      secDiv.appendChild(currentGridDiv);

      wrap.appendChild(secDiv);
      currentSecDiv = secDiv;
    }

    const row = document.createElement('div');
    const isFull = isFullWidthField(f.name);
    row.className = `field-row kind-${f.kind} ${isFull ? 'full-width' : ''}`;
    rowByField[f.name] = row;

    const labelRow = document.createElement('div');
    labelRow.className = 'field-label-row';

    const label = document.createElement('label');
    label.className = 'field-label';
    label.textContent = (f.kind === 'number' && f.name.startsWith('Thời gian')) ? `${f.name} (phút)` : f.name;
    labelRow.appendChild(label);
    row.appendChild(labelRow);

    let input;
    const raw = values[f.name] ?? '';

    if (f.kind === 'date') {
      input = document.createElement('input');
      input.type = 'text';
      input.placeholder = 'dd/mm/yyyy';
      input.value = raw;
    } else if (f.kind === 'number') {
      input = document.createElement('input');
      input.type = 'number';
      input.value = raw;
    } else if (String(raw).length > 60 || String(raw).includes('\n') || isLikelyLong(f.name)) {
      input = document.createElement('textarea');
      input.value = raw;
      input.rows = 2;
    } else {
      input = document.createElement('input');
      input.type = 'text';
      input.value = raw;
    }
    input.dataset.field = f.name;
    input.spellcheck = true;
    input.setAttribute('lang', 'vi');
    input.setAttribute('autocomplete', 'off');
    
    let isComposing = false;
    input.addEventListener('compositionstart', () => { isComposing = true; });
    input.addEventListener('compositionend', () => {
      isComposing = false;
      lastEditedField = f.name;
      lastEditedText = input.value;
      if (input.tagName === 'TEXTAREA') autoResizeTextarea(input);
      queuePreview();
    });

    const onInputOrBlur = (e) => {
      lastEditedField = f.name;
      lastEditedText = e.target.value;
    };
    
    if (shouldAutoPunctuate(f.name)) {
      input.addEventListener('blur', (e) => {
        onInputOrBlur(e);
        const fixed = autoPunctuate(input.value);
        if (fixed !== input.value) {
          input.value = fixed;
          if (input.tagName === 'TEXTAREA') autoResizeTextarea(input);
          lastEditedText = fixed;
          queuePreview();
        }
      });
    }
    input.addEventListener('input', (e) => {
      if (input.tagName === 'TEXTAREA') autoResizeTextarea(input);
      if (!isComposing) {
        onInputOrBlur(e);
        queuePreview();
      }
    });
    row.appendChild(input);

    if (currentGridDiv) {
      currentGridDiv.appendChild(row);
    } else {
      wrap.appendChild(row);
    }
    if (input.tagName === 'TEXTAREA') autoResizeTextarea(input);
  });

  // Cảnh báo tổng thời gian: chèn ngay dưới ô "tổng" của từng công thức,
  // tự cập nhật ngay khi người dùng sửa bất kỳ ô liên quan nào.
  TIME_CHECKS.forEach((check) => {
    const totalRow = rowByField[check.total];
    if (!totalRow) return;
    const warn = document.createElement('div');
    warn.className = 'field-warn';
    warn.hidden = true;
    totalRow.appendChild(warn);

    [check.total, ...check.parts].forEach((name) => {
      const el = rowByField[name] && rowByField[name].querySelector('[data-field]');
      if (el) el.addEventListener('input', () => updateTimeCheck(check, warn));
    });
    updateTimeCheck(check, warn);
  });
}

function updateTimeCheck(check, warn) {
  const readNumber = (name) => {
    const el = document.querySelector(`#fieldForm [data-field="${CSS.escape(name)}"]`);
    const v = parseFloat(el ? el.value : '');
    return Number.isNaN(v) ? null : v;
  };
  const total = readNumber(check.total);
  const parts = check.parts.map(readNumber);
  if (total === null || parts.some((v) => v === null)) {
    warn.hidden = true;
    return;
  }
  const sum = parts.reduce((a, b) => a + b, 0);
  if (sum !== total) {
    warn.hidden = false;
    warn.textContent = `⚠ "${check.total}" đang ghi ${total} phút, nhưng ${check.parts.join(' + ')} = ${sum} phút. Kiểm tra lại.`;
  } else {
    warn.hidden = true;
  }
}

// ---------- Kiểm tra thiếu dữ liệu trước khi xuất ----------
function findMissingFields(lectureNames) {
  const missing = {};
  lectureNames.forEach((name) => {
    const vals = LECTURES[name] || {};
    const empty = FIELDS.filter((f) => !String(vals[f.name] ?? '').trim()).map((f) => f.name);
    if (empty.length) missing[name] = empty;
  });
  return missing;
}

function confirmIfMissing(lectureNames) {
  const missing = findMissingFields(lectureNames);
  const names = Object.keys(missing);
  if (!names.length) return true;
  const detail = names.map((n) => {
    const fields = missing[n];
    const shown = fields.slice(0, 5).join(', ') + (fields.length > 5 ? '…' : '');
    return `- ${n}: thiếu ${fields.length} trường (${shown})`;
  }).join('\n');
  return confirm(`Một số bài giảng còn thiếu dữ liệu:\n${detail}\n\nVẫn tiếp tục xuất?`);
}

function autoResizeTextarea(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 'px';
}

function isLikelyLong(name) {
  const longFields = [
    'Nhận xét', 'Mục đích', 'Yêu cầu', 'Mức tự chủ', 'Nội dung bài giảng',
    'Phân bổ thời gian', 'Phương pháp của', 'Vật chất bảo đảm', 'Nội dung kết thúc',
    'Tổ chức lớp học', 'Tài liệu', 'Tổng số thời gian'
  ];
  return longFields.some((p) => name.startsWith(p) || name.includes(p));
}

function saveCurrentFormIntoState() {
  if (!CURRENT) return;
  const inputs = document.querySelectorAll('#fieldForm [data-field]');
  const values = {};
  inputs.forEach((el) => { values[el.dataset.field] = el.value; });
  LECTURES[CURRENT] = values;
}

let PREVIEW_TIMER = null;

function getSearchTerm() {
  if (lastEditedText && lastEditedText.trim().length > 0) {
    const lines = lastEditedText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length > 0) {
      let term = lines[0].replace(/^[-–—+]\s*/, '').replace(/[;,.:]+$/, '');
      if (term.length > 30) term = term.substring(0, 30);
      return term;
    }
  }
  if (lastEditedField) {
     const words = lastEditedField.split(' ');
     if (words.length > 2) return words.slice(0, 2).join(' ');
     return lastEditedField;
  }
  return '';
}

function updatePreview() {
  saveCurrentFormIntoState();
  if (!CURRENT) return;
  
  // Clone data and ensure everything is auto-punctuated for the preview
  const data = Object.assign({}, LECTURES[CURRENT]);
  for (const key in data) {
    if (shouldAutoPunctuate(key) && data[key]) {
      data[key] = autoPunctuate(String(data[key]));
    }
  }
  
  const spinner = $('#previewSpinner');
  if (spinner) spinner.hidden = false;
  
  fetch('/api/preview', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: data }),
  })
    .then(res => res.blob())
    .then(blob => {
      let url = URL.createObjectURL(blob);
      const frame = $('#previewFrame');
      frame.onload = () => { if(spinner) spinner.hidden = true; };
      frame.src = url;
    })
    .catch((e) => {
      console.error('Lỗi khi tạo bản xem trước:', e);
      if (spinner) spinner.hidden = true;
    });
}

function queuePreview() {
  if (PREVIEW_TIMER) clearTimeout(PREVIEW_TIMER);
  PREVIEW_TIMER = setTimeout(updatePreview, 1200);
}

// ---------- Preset Manager (Lưu/Tải Mẫu Thông Tin Cố Định) ----------
const PRESET_FIELDS = [
  'Tên khoa', 'Chức vụ người phê duyệt', 'Cấp bậc, học vị, họ tên người phê duyệt',
  'Địa điểm phê duyệt', 'Chức vụ người thông qua', 'Cấp bậc, học vị, họ tên người thông qua',
  'Phương pháp thông qua', 'Địa điểm thông qua', 'Chức vụ giảng viên biên soạn',
  'Cấp bậc, học vị, họ tên giảng viên biên soạn'
];

const btnSavePreset = $('#btnSavePreset');
if (btnSavePreset) {
  btnSavePreset.addEventListener('click', () => {
    saveCurrentFormIntoState();
    if (!CURRENT || !LECTURES[CURRENT]) {
      alert('Chưa chọn bài giảng nào để lưu mẫu.');
      return;
    }
    const presetData = {};
    PRESET_FIELDS.forEach((f) => {
      if (LECTURES[CURRENT][f] !== undefined) {
        presetData[f] = LECTURES[CURRENT][f];
      }
    });

    let txtContent = "# APPBGM PRESET DATA FILE\n# File Mau Thong Tin Hanh Chinh & Phe Duyet\n\n";
    for (const key in presetData) {
      txtContent += `${key}=${String(presetData[key]).replace(/\n/g, '\\n')}\n`;
    }

    const blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'Mau_ThongTin_BaiGiang.txt';
    link.click();
    URL.revokeObjectURL(link.href);

    alert('Đã xuất file Mẫu Thông Tin thành công: "Mau_ThongTin_BaiGiang.txt"');
  });
}

const btnLoadPreset = $('#btnLoadPreset');
if (btnLoadPreset) {
  btnLoadPreset.addEventListener('click', () => {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.txt';

    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (evt) => {
        const text = evt.target.result;
        const presetData = {};
        const lines = text.split('\n');

        lines.forEach((line) => {
          line = line.trim();
          if (!line || line.startsWith('#')) return;
          const eqIdx = line.indexOf('=');
          if (eqIdx !== -1) {
            const key = line.substring(0, eqIdx).trim();
            const val = line.substring(eqIdx + 1).replace(/\\n/g, '\n').trim();
            presetData[key] = val;
          }
        });

        if (Object.keys(presetData).length === 0) {
          alert('File .txt không chứa dữ liệu mẫu hợp lệ.');
          return;
        }

        const applyToAll = confirm('Bạn có muốn áp dụng Mẫu Thông Tin này cho TẤT CẢ các bài giảng hiện tại không?\n\n- Bấm OK (Đồng ý): Áp dụng cho TẤT CẢ bài giảng.\n- Bấm Cancel (Hủy): Chỉ áp dụng cho bài giảng hiện tại.');

        if (applyToAll) {
          for (const lName in LECTURES) {
            Object.assign(LECTURES[lName], presetData);
          }
          alert(`Đã nạp và áp dụng Mẫu Thông Tin từ file .txt cho tất cả ${LECTURE_ORDER.length} bài giảng!`);
        } else {
          if (CURRENT && LECTURES[CURRENT]) {
            Object.assign(LECTURES[CURRENT], presetData);
            alert(`Đã nạp Mẫu Thông Tin cho bài giảng: "${CURRENT}"!`);
          }
        }

        renderForm();
        queuePreview();
      };
      reader.readAsText(file, 'UTF-8');
    });

    fileInput.click();
  });
}

// ---------- Xuất dữ liệu bài giảng ra file Excel để lưu trữ ----------
const btnExportData = $('#btnExportData');
if (btnExportData) {
  btnExportData.addEventListener('click', async () => {
    saveCurrentFormIntoState();
    if (!LECTURE_ORDER.length) {
      alert('Chưa có dữ liệu bài giảng để lưu.');
      return;
    }
    btnExportData.disabled = true;
    btnExportData.textContent = 'Đang xuất dữ liệu…';
    try {
      const res = await fetch('/api/export_excel_data_native', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ lectures: LECTURES, lecture_order: LECTURE_ORDER })
      });
      const data = await res.json();
      if (!res.ok) {
        alert(data.error || 'Lỗi khi lưu dữ liệu.');
      } else if (data.success) {
        alert(data.message);
      }
    } catch (e) {
      alert('Lỗi kết nối: ' + e);
    } finally {
      btnExportData.disabled = false;
      btnExportData.textContent = '💾 Lưu file dữ liệu (.xlsx)';
    }
  });
}

// ---------- Nhân bản bài giảng ----------
const btnDuplicate = $('#btnDuplicateLecture');
if (btnDuplicate) {
  btnDuplicate.addEventListener('click', () => {
    saveCurrentFormIntoState();
    if (!CURRENT || !LECTURES[CURRENT]) {
      alert('Chưa chọn bài giảng để nhân bản.');
      return;
    }
    let copyName = `${CURRENT} (Bản sao)`;
    let n = 2;
    while (LECTURES[copyName]) {
      copyName = `${CURRENT} (Bản sao ${n})`;
      n++;
    }
    LECTURES[copyName] = JSON.parse(JSON.stringify(LECTURES[CURRENT]));
    LECTURES[copyName]['Tên bài giảng'] = copyName;
    LECTURE_ORDER.push(copyName);
    SELECTED_FOR_EXPORT.add(copyName);
    renderTabs();
    selectTab(copyName);
    alert(`Đã nhân bản bài giảng: "${copyName}"`);
  });
}

// ---------- Thêm bài giảng mới ----------
function addNewLecture() {
  saveCurrentFormIntoState();
  let baseName = "Bài giảng mới";
  let newName = baseName;
  let count = 1;
  while (LECTURES[newName]) {
    count++;
    newName = `${baseName} ${count}`;
  }

  const prevLecture = CURRENT ? LECTURES[CURRENT] : {};
  const newLectureData = {};
  FIELDS.forEach((f) => {
    if (PRESET_FIELDS.includes(f.name) || f.name.startsWith('Tên khoa') || f.name.startsWith('Tên học phần')) {
      newLectureData[f.name] = prevLecture[f.name] || '';
    } else {
      newLectureData[f.name] = '';
    }
  });

  newLectureData['Tên bài giảng'] = newName;
  LECTURES[newName] = newLectureData;
  LECTURE_ORDER.push(newName);
  SELECTED_FOR_EXPORT.add(newName);

  renderTabs();
  selectTab(newName);
  const mainWorkspace = $('#mainWorkspace') || document.querySelector('.workspace');
  if (mainWorkspace && mainWorkspace.classList.contains('matrix-mode-active')) {
    renderMatrixView();
  }
}

if ($('#btnAddLecture')) $('#btnAddLecture').addEventListener('click', addNewLecture);
if ($('#btnAddLectureMatrix')) $('#btnAddLectureMatrix').addEventListener('click', addNewLecture);

// ---------- Mode Switcher (Xem chi tiết vs Bảng Excel) ----------
const btnViewForm = $('#btnViewForm');
const btnViewMatrix = $('#btnViewMatrix');
const panelForm = $('#panelForm');
const panelMatrix = $('#panelMatrix');
const panelPreview = document.querySelector('.panel-preview');

if (btnViewForm && btnViewMatrix) {
  btnViewForm.addEventListener('click', () => {
    btnViewForm.classList.add('active');
    btnViewForm.style.background = '#fff';
    btnViewForm.style.color = 'var(--navy)';
    btnViewForm.style.fontWeight = 'bold';

    btnViewMatrix.classList.remove('active');
    btnViewMatrix.style.background = 'transparent';
    btnViewMatrix.style.color = '#fff';

    const mainWorkspace = $('#mainWorkspace') || document.querySelector('.workspace');
    if (mainWorkspace) mainWorkspace.classList.remove('matrix-mode-active');

    panelForm.style.display = 'flex';
    if (panelPreview) panelPreview.style.display = 'flex';
    panelMatrix.style.display = 'none';

    renderForm();
    updatePreview();
  });

  btnViewMatrix.addEventListener('click', () => {
    saveCurrentFormIntoState();

    btnViewMatrix.classList.add('active');
    btnViewMatrix.style.background = '#fff';
    btnViewMatrix.style.color = 'var(--navy)';
    btnViewMatrix.style.fontWeight = 'bold';

    btnViewForm.classList.remove('active');
    btnViewForm.style.background = 'transparent';
    btnViewForm.style.color = '#fff';

    const mainWorkspace = $('#mainWorkspace') || document.querySelector('.workspace');
    if (mainWorkspace) mainWorkspace.classList.add('matrix-mode-active');

    renderMatrixView();
  });
}

function renderMatrixView() {
  const wrap = $('#matrixTableWrapper');
  if (!wrap) return;
  wrap.innerHTML = '';

  const table = document.createElement('table');
  table.className = 'matrix-table';

  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  trHead.innerHTML = `<th class="field-name-col">Tên Trường Thông Tin</th>`;

  LECTURE_ORDER.forEach((lName) => {
    const th = document.createElement('th');
    th.textContent = lName;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  let currentSecIdx = -1;

  FIELDS.forEach((f) => {
    const secDefIdx = SECTION_DEFINITIONS.findIndex(s => s.firstField === f.name);
    if (secDefIdx !== -1) {
      currentSecIdx = secDefIdx;
      const secDef = SECTION_DEFINITIONS[secDefIdx];

      const trGroup = document.createElement('tr');
      trGroup.className = 'matrix-group-header';
      trGroup.dataset.secIdx = secDefIdx;

      const tdGroup = document.createElement('td');
      tdGroup.colSpan = LECTURE_ORDER.length + 1;
      tdGroup.innerHTML = `
        <span class="group-toggle-icon">▼</span>
        <span>${secDef.title}</span>
      `;
      trGroup.appendChild(tdGroup);
      tbody.appendChild(trGroup);

      trGroup.addEventListener('click', () => {
        const isCollapsed = trGroup.classList.toggle('collapsed');
        const rows = tbody.querySelectorAll(`.matrix-row-sec-${secDefIdx}`);
        rows.forEach(r => {
          r.style.display = isCollapsed ? 'none' : 'table-row';
        });
        const icon = trGroup.querySelector('.group-toggle-icon');
        if (icon) icon.textContent = isCollapsed ? '▶' : '▼';
      });
    }

    const tr = document.createElement('tr');
    if (currentSecIdx !== -1) {
      tr.classList.add(`matrix-row-sec-${currentSecIdx}`);
    }

    const tdName = document.createElement('td');
    tdName.className = 'field-name-col';
    tdName.textContent = f.name;
    tr.appendChild(tdName);

    LECTURE_ORDER.forEach((lName) => {
      const td = document.createElement('td');
      const textarea = document.createElement('textarea');
      const val = (LECTURES[lName] && LECTURES[lName][f.name]) ?? '';
      textarea.value = val;

      textarea.addEventListener('input', (e) => {
        if (!LECTURES[lName]) LECTURES[lName] = {};
        LECTURES[lName][f.name] = e.target.value;
        autoResizeTextarea(textarea);
      });
      textarea.addEventListener('blur', (e) => {
        if (shouldAutoPunctuate(f.name)) {
          const fixed = autoPunctuate(textarea.value);
          if (fixed !== textarea.value) {
            textarea.value = fixed;
            LECTURES[lName][f.name] = fixed;
            autoResizeTextarea(textarea);
          }
        }
      });

      td.appendChild(textarea);
      tr.appendChild(td);
      setTimeout(() => autoResizeTextarea(textarea), 0);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  wrap.appendChild(table);
}

// ---------- Hàm dùng chung xuất tất cả / xuất ZIP Word & PDF ----------
async function runExportAllJob(includePdf = false) {
  saveCurrentFormIntoState();
  const names = LECTURE_ORDER.filter((n) => SELECTED_FOR_EXPORT.has(n));
  if (!names.length) { alert('Chưa chọn bài giảng nào để xuất.'); return; }
  if (!confirmIfMissing(names)) return;

  const lecturesPayload = {};
  names.forEach((n) => { lecturesPayload[n] = LECTURES[n]; });

  const btnAll = $('#btnExportAll');
  const btnPdf = $('#btnExportZipPdf');
  const progressBox = $('#exportProgress');
  const progressText = $('#exportProgressText');
  
  if (btnAll) btnAll.disabled = true;
  if (btnPdf) btnPdf.disabled = true;
  progressBox.hidden = false;
  progressText.textContent = `Đang xuất 0/${names.length} bài…`;

  try {
    const startRes = await fetch('/api/export_all/start', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lectures: lecturesPayload, include_pdf: includePdf }),
    });
    const startData = await startRes.json();
    if (!startRes.ok) { alert(startData.error || 'Lỗi khi xuất file.'); return; }
    const jobId = startData.job_id;

    let job;
    while (true) {
      await new Promise((resolve) => setTimeout(resolve, 400));
      const pRes = await fetch(`/api/export_all/progress/${jobId}`);
      job = await pRes.json();
      if (!pRes.ok) { alert(job.error || 'Lỗi khi theo dõi tiến trình xuất.'); return; }
      progressText.textContent = `Đang xuất ${job.done}/${job.total} bài…`;
      if (job.status === 'done' || job.status === 'error') break;
    }

    if (job.status === 'error') {
      alert(job.error || 'Xuất file thất bại.');
      return;
    }

    const resultRes = await fetch(`/api/export_all/result_native/${jobId}`, { method: 'POST' });
    const data = await resultRes.json();
    if (!resultRes.ok) {
      alert(data.error || 'Lỗi khi tải file đã xuất.');
      return;
    }
    
    if (data.success) {
      alert(data.message);
    }

    const skippedHeader = resultRes.headers.get('X-Export-Skipped');
    if (skippedHeader) {
      const skipped = JSON.parse(skippedHeader);
      alert(`Đã xuất xong, nhưng ${skipped.length} bài bị lỗi và không có trong file zip:\n`
        + skipped.map((s) => `- ${s.lecture}: ${s.error}`).join('\n'));
    }
  } catch (e) {
    alert('Lỗi kết nối: ' + e);
  } finally {
    if (btnAll) btnAll.disabled = false;
    if (btnPdf) btnPdf.disabled = false;
    progressBox.hidden = true;
    updateExportAllLabel();
  }
}

// ---------- Nút Xuất tất cả Word (.zip) ----------
const btnExportAll = $('#btnExportAll');
if (btnExportAll) {
  btnExportAll.addEventListener('click', () => runExportAllJob(false));
}

// ---------- Nút Xuất tất cả Word & PDF (.zip) ----------
const btnExportZipPdf = $('#btnExportZipPdf');
if (btnExportZipPdf) {
  btnExportZipPdf.addEventListener('click', () => runExportAllJob(true));
}

// ---------- Xuất 1 bài ----------
$('#btnExportOne').addEventListener('click', async () => {
  saveCurrentFormIntoState();
  if (!confirmIfMissing([CURRENT])) return;
  
  const btn = $('#btnExportOne');
  btn.disabled = true;
  btn.textContent = 'Đang xuất...';
  
  try {
    const res = await fetch('/api/export_native', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: LECTURES[CURRENT] }),
    });
    const data = await res.json();
    if (!res.ok) {
      alert(data.error || 'Lỗi khi xuất file.');
      return;
    }
    if (data.success) {
      alert(data.message);
    }
  } catch (e) {
    alert('Lỗi kết nối: ' + e);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span class="icon">💾</span> Xuất 1 bài này';
  }
});

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

// ---------- Tải file Excel mẫu ----------
$('#btnDownloadTemplate').addEventListener('click', async () => {
  const btn = $('#btnDownloadTemplate');
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Đang gọi hộp thoại lưu...';
  try {
    const res = await fetch('/api/download_template_native', { method: 'POST' });
    const data = await res.json();
    if (!res.ok) {
      alert(data.error || 'Lỗi khi tạo file mẫu.');
      return;
    }
    if (data.success) {
      alert(data.message);
    }
  } catch (e) {
    alert('Lỗi kết nối: ' + e);
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
});

// ---------- Chỉnh sửa trực tiếp (LibreOffice) ----------
let CURRENT_EDIT_ID = null;

if ($('#btnDirectEdit')) {
  $('#btnDirectEdit').addEventListener('click', async () => {
    saveCurrentFormIntoState();
    if (!CURRENT) return;

    const btn = $('#btnDirectEdit');
    btn.disabled = true;
    btn.textContent = 'Đang mở LibreOffice…';

    try {
      const data = Object.assign({}, LECTURES[CURRENT]);
      for (const key in data) {
        if (shouldAutoPunctuate(key) && data[key]) {
          data[key] = autoPunctuate(String(data[key]));
        }
      }

      const res = await fetch('/api/direct_edit/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: data }),
      });
      const result = await res.json();
      if (!res.ok) {
        alert(result.error || 'Lỗi khi mở trình chỉnh sửa.');
        btn.disabled = false;
        btn.textContent = 'Chỉnh sửa trực tiếp';
        return;
      }

      CURRENT_EDIT_ID = result.edit_id;
      btn.textContent = 'Đang chỉnh sửa…';
      if ($('#btnCompleteEdit')) $('#btnCompleteEdit').hidden = false;
      alert('Đã mở LibreOffice Writer!\n\nSau khi sửa xong văn bản, hãy bấm "Ctrl + S" để lưu trên LibreOffice rồi bấm nút "Hoàn tất" ở ứng dụng để cập nhật lại bản xem trước PDF.');
    } catch (e) {
      alert('Lỗi kết nối: ' + e);
      btn.disabled = false;
      btn.textContent = 'Chỉnh sửa trực tiếp';
    }
  });
}

if ($('#btnCompleteEdit')) {
  $('#btnCompleteEdit').addEventListener('click', async () => {
    if (!CURRENT_EDIT_ID) return;

    const btnComplete = $('#btnCompleteEdit');
    const btnDirect = $('#btnDirectEdit');
    const spinner = $('#previewSpinner');
    if (spinner) spinner.hidden = false;
    btnComplete.disabled = true;
    btnComplete.textContent = 'Đang nạp lại PDF…';

    try {
      const res = await fetch('/api/direct_edit/complete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ edit_id: CURRENT_EDIT_ID }),
      });

      if (!res.ok) {
        const errData = await res.json();
        alert(errData.error || 'Lỗi khi cập nhật bản xem trước PDF.');
        btnComplete.disabled = false;
        btnComplete.textContent = 'Hoàn tất';
        return;
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const frame = $('#previewFrame');
      frame.onload = () => { if (spinner) spinner.hidden = true; };
      frame.src = url;

      alert('Đã cập nhật bản xem trước PDF từ file chỉnh sửa thành công!');
      CURRENT_EDIT_ID = null;
      btnComplete.hidden = true;
      btnComplete.disabled = false;
      btnComplete.textContent = 'Hoàn tất';
      btnDirect.disabled = false;
      btnDirect.textContent = 'Chỉnh sửa trực tiếp';
    } catch (e) {
      alert('Lỗi kết nối: ' + e);
      btnComplete.disabled = false;
      btnComplete.textContent = 'Hoàn tất';
    } finally {
      if (spinner) spinner.hidden = true;
    }
  });
}

// ---------- Custom Context Menu chuột phải cho tất cả các ô nhập liệu ----------
let activeContextInput = null;
let internalClipboard = "";
const inputUndoHistoryMap = new Map(); // Lưu giá trị cũ trước khi sửa của từng ô

const ctxMenu = $('#customContextMenu');
const ctxUndo = $('#ctxUndo');
const ctxCopy = $('#ctxCopy');
const ctxPaste = $('#ctxPaste');
const ctxSyncField = $('#ctxSyncField');
const ctxClear = $('#ctxClear');

document.addEventListener('contextmenu', (e) => {
  const target = e.target;
  if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) {
    e.preventDefault();
    activeContextInput = target;

    const x = Math.min(e.clientX, window.innerWidth - 240);
    const y = Math.min(e.clientY, window.innerHeight - 220);

    ctxMenu.style.left = `${x}px`;
    ctxMenu.style.top = `${y}px`;
    ctxMenu.hidden = false;

    // Kiểm tra xem ô này có lịch sử Undo không
    if (ctxUndo) {
      const hasHistory = inputUndoHistoryMap.has(activeContextInput) && inputUndoHistoryMap.get(activeContextInput).length > 0;
      if (hasHistory) {
        ctxUndo.style.opacity = '1';
        ctxUndo.style.pointerEvents = 'auto';
      } else {
        ctxUndo.style.opacity = '0.5';
        ctxUndo.style.pointerEvents = 'none';
      }
    }
  } else {
    if (ctxMenu) ctxMenu.hidden = true;
  }
});

document.addEventListener('click', (e) => {
  if (ctxMenu && !ctxMenu.contains(e.target)) {
    ctxMenu.hidden = true;
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && ctxMenu) {
    ctxMenu.hidden = true;
  }
});

function pushUndoState(inputEl) {
  if (!inputEl) return;
  if (!inputUndoHistoryMap.has(inputEl)) {
    inputUndoHistoryMap.set(inputEl, []);
  }
  const history = inputUndoHistoryMap.get(inputEl);
  history.push(inputEl.value);
  if (history.length > 20) history.shift();
}

if (ctxUndo) {
  ctxUndo.addEventListener('click', () => {
    if (!activeContextInput) return;
    if (inputUndoHistoryMap.has(activeContextInput)) {
      const history = inputUndoHistoryMap.get(activeContextInput);
      if (history.length > 0) {
        const prevValue = history.pop();
        activeContextInput.value = prevValue;
        activeContextInput.dispatchEvent(new Event('input', { bubbles: true }));
        if (activeContextInput.tagName === 'TEXTAREA') autoResizeTextarea(activeContextInput);
      }
    }
    ctxMenu.hidden = true;
  });
}

if (ctxCopy) {
  ctxCopy.addEventListener('click', async () => {
    if (!activeContextInput) return;
    const selectedText = window.getSelection().toString();
    const textToCopy = selectedText || activeContextInput.value || '';
    internalClipboard = textToCopy;

    try {
      await navigator.clipboard.writeText(textToCopy);
    } catch (err) {}
    ctxMenu.hidden = true;
  });
}

if (ctxPaste) {
  ctxPaste.addEventListener('click', async () => {
    if (!activeContextInput) return;
    pushUndoState(activeContextInput);

    let pasteText = internalClipboard;
    try {
      const clipText = await navigator.clipboard.readText();
      if (clipText) pasteText = clipText;
    } catch (err) {}

    if (pasteText) {
      activeContextInput.value = pasteText;
      activeContextInput.dispatchEvent(new Event('input', { bubbles: true }));
      if (activeContextInput.tagName === 'TEXTAREA') autoResizeTextarea(activeContextInput);
    }
    ctxMenu.hidden = true;
  });
}

if (ctxSyncField) {
  ctxSyncField.addEventListener('click', () => {
    if (!activeContextInput) return;
    const fieldName = activeContextInput.dataset.field;
    const val = activeContextInput.value || '';

    if (fieldName) {
      for (const lName in LECTURES) {
        LECTURES[lName][fieldName] = val;
      }
      alert(`Đã sao chép nội dung ô "${fieldName}" sang tất cả ${LECTURE_ORDER.length} bài giảng!`);
      queuePreview();
    } else {
      alert('Đã đồng bộ nội dung ô hiện tại!');
    }
    ctxMenu.hidden = true;
  });
}

if (ctxClear) {
  ctxClear.addEventListener('click', () => {
    if (!activeContextInput) return;
    pushUndoState(activeContextInput);
    activeContextInput.value = '';
    activeContextInput.dispatchEvent(new Event('input', { bubbles: true }));
    if (activeContextInput.tagName === 'TEXTAREA') autoResizeTextarea(activeContextInput);
    ctxMenu.hidden = true;
  });
}

