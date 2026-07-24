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

// ---------- Tự động nạp dữ liệu mặc định khi vừa mở App ----------
window.addEventListener('DOMContentLoaded', async () => {
  try {
    const res = await fetch('/api/init_default_data');
    if (res.ok) {
      const data = await res.json();
      applyDataset(data, 'Đã sẵn sàng chỉnh sửa dữ liệu bài giảng');
      $('#nameXlsx').textContent = 'Đã nạp dữ liệu mặc định (Bấm để chọn file khác)';
      $('#dropXlsx').classList.add('filled');
    }
  } catch (e) {
    console.log('Init default data notice:', e);
  }
});

// ---------- Chọn file ----------
$('#dropXlsx').addEventListener('click', async () => {
  const nameEl = $('#nameXlsx');
  nameEl.textContent = 'Đang mở hộp thoại...';
  try {
    const res = await fetch('/api/open_excel_native', { method: 'POST' });
    const data = await res.json();
    if (!res.ok) {
      alert(data.error || 'Lỗi không xác định');
      nameEl.textContent = 'Bấm để chọn file Excel dữ liệu…';
      return;
    }
    
    if (!data.success) {
      // User cancelled
      nameEl.textContent = 'Bấm để chọn file Excel dữ liệu…';
      return;
    }

    nameEl.textContent = 'Đã nạp file dữ liệu';
    $('#dropXlsx').classList.add('filled');
    applyDataset(data, `Đã nạp ${data.lecture_order.length} bài giảng từ file Excel`);
  } catch (e) {
    alert('Lỗi kết nối: ' + e);
    nameEl.textContent = 'Bấm để chọn file Excel dữ liệu…';
  }
});

// ---------- Nạp dữ liệu ----------
async function triggerUpload(file) {
  const fd = new FormData();
  fd.append('xlsx', file);

  const nameEl = $('#nameXlsx');
  nameEl.textContent = 'Đang nạp…';
  try {
    const res = await fetch('/api/upload_native', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) { alert(data.error || 'Lỗi không xác định'); return; }

    FIELDS = data.fields;
    LECTURES = data.lectures;
    
    // Tự động chuẩn hóa toàn bộ dữ liệu vừa nạp từ Excel
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
    $('#tabFilter').value = '';

    renderWarnings(data.warnings);
    renderTabs();
    renderForm();
    $('#btnExportOne').disabled = false;
    if ($('#btnDirectEdit')) $('#btnDirectEdit').disabled = false;
    updateExportAllLabel();
    $('#statusBox').textContent = `Đã nạp ${LECTURE_ORDER.length} bài giảng · ${FIELDS.length} trường`;
    $('#statusBox').classList.add('ready');
    updatePreview();
  } catch (e) {
    alert('Không kết nối được tới server: ' + e);
  } finally {
    nameEl.textContent = file.name;
  }
}

function renderWarnings(w) {
  const box = $('#warnBox');
  const msgs = [];
  if (w.thieu_trong_excel.length) {
    msgs.push(`<b>Excel còn thiếu ${w.thieu_trong_excel.length} trường</b> mà mẫu Word cần: ${w.thieu_trong_excel.join(', ')}.`);
  }
  if (w.thua_trong_excel.length) {
    msgs.push(`<b>Excel có ${w.thua_trong_excel.length} trường thừa</b>, không thấy trong mẫu Word: ${w.thua_trong_excel.join(', ')}.`);
  }
  if (msgs.length) {
    box.innerHTML = msgs.join('<br>');
    box.hidden = false;
  } else {
    box.hidden = true;
  }
}

// ---------- Tabs bài giảng ----------
function renderTabs() {
  const wrap = $('#lectureTabs');
  const filterRow = $('#tabFilterRow');
  filterRow.hidden = LECTURE_ORDER.length <= TAB_FILTER_THRESHOLD;

  wrap.innerHTML = '';
  const needle = TAB_FILTER.trim().toLowerCase();
  const filtered = needle
    ? LECTURE_ORDER.filter((name) => name.toLowerCase().includes(needle))
    : LECTURE_ORDER;

  filtered.forEach((name) => {
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
    });
    wrap.appendChild(t);
  });

  if (!filtered.length) {
    const empty = document.createElement('div');
    empty.className = 'tab-empty';
    empty.textContent = 'Không tìm thấy bài giảng nào khớp.';
    wrap.appendChild(empty);
  }
  applyHints();
}

$('#tabFilter').addEventListener('input', () => {
  TAB_FILTER = $('#tabFilter').value;
  renderTabs();
});

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
      currentSecDiv = document.createElement('div');
      currentSecDiv.className = 'form-section';

      const secHeader = document.createElement('div');
      secHeader.className = 'form-section-header';
      secHeader.innerHTML = `
        <span class="sec-title">${SECTION_DEFINITIONS[secDefIdx].title}</span>
        <span class="sec-arrow">▼</span>
      `;
      secHeader.addEventListener('click', () => {
        currentSecDiv.classList.toggle('collapsed');
      });
      currentSecDiv.appendChild(secHeader);

      currentGridDiv = document.createElement('div');
      currentGridDiv.className = 'field-group-grid';
      currentSecDiv.appendChild(currentGridDiv);

      wrap.appendChild(currentSecDiv);
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

    // CHỈ HIỂN THỊ NÚT ⏩ Ở NHÓM I (Thông tin hành chính & Phê duyệt)
    if (currentSecIdx === 0) {
      const btnSync = document.createElement('button');
      btnSync.type = 'button';
      btnSync.className = 'btn-sync-field';
      btnSync.textContent = '⏩ Áp dụng tất cả';
      btnSync.title = `Sao chép thông tin "${f.name}" sang tất cả bài giảng`;
      btnSync.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const inputEl = row.querySelector('[data-field]');
        const val = inputEl ? inputEl.value : (values[f.name] || '');
        for (const lName in LECTURES) {
          LECTURES[lName][f.name] = val;
        }
        alert(`Đã áp dụng thông tin "${f.name}" cho tất cả ${LECTURE_ORDER.length} bài giảng!`);
        queuePreview();
      });
      labelRow.appendChild(btnSync);
    }
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
      input.addEventListener('input', () => {
        autoResizeTextarea(input);
        queuePreview();
      });
    } else {
      input = document.createElement('input');
      input.type = 'text';
      input.value = raw;
    }
    input.dataset.field = f.name;
    input.spellcheck = false;
    input.setAttribute('lang', 'vi');
    input.setAttribute('autocomplete', 'off');
    
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
      onInputOrBlur(e);
      queuePreview();
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
    localStorage.setItem('APPBGM_PRESET', JSON.stringify(presetData));
    alert('Đã lưu mẫu thông tin đơn vị & người phê duyệt/biên soạn thành công!');
  });
}

const btnLoadPreset = $('#btnLoadPreset');
if (btnLoadPreset) {
  btnLoadPreset.addEventListener('click', () => {
    const saved = localStorage.getItem('APPBGM_PRESET');
    if (!saved) {
      alert('Chưa có mẫu thông tin nào được lưu trước đây.');
      return;
    }
    const presetData = JSON.parse(saved);
    if (!CURRENT || !LECTURES[CURRENT]) {
      alert('Chưa chọn bài giảng để áp dụng mẫu.');
      return;
    }
    Object.assign(LECTURES[CURRENT], presetData);
    renderForm(CURRENT);
    queuePreview();
    alert('Đã nạp mẫu thông tin cố định thành công!');
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
  if ($('#panelMatrix') && $('#panelMatrix').style.display !== 'none') {
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

    panelForm.style.display = 'none';
    if (panelPreview) panelPreview.style.display = 'none';
    panelMatrix.style.display = 'block';

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
  FIELDS.forEach((f) => {
    const tr = document.createElement('tr');
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

