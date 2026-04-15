/* ═══════════════════════════════════════════════════════════
   StockOps — Step 01 JavaScript
   ═══════════════════════════════════════════════════════════ */

// ROUND_ID and APPLICANT_NAMES are set inline in step01.html

// ── Pending uploads buffer (before assignment to applicant) ──
const pendingUploads = {
  application: [],
  id_copy: [],
  account_copy: [],
};

// ══════════════════════════════════════════════════════
// 체크박스 & 선택 삭제
// ══════════════════════════════════════════════════════

function toggleAllChecks(masterCb) {
  document.querySelectorAll('#applicant-tbody .row-check').forEach(cb => {
    cb.checked = masterCb.checked;
  });
  updateBulkDeleteBtn();
}

function onRowCheckChange() {
  const all  = document.querySelectorAll('#applicant-tbody .row-check');
  const checked = document.querySelectorAll('#applicant-tbody .row-check:checked');
  const master = document.getElementById('check-all');
  if (master) {
    master.indeterminate = checked.length > 0 && checked.length < all.length;
    master.checked = checked.length === all.length && all.length > 0;
  }
  updateBulkDeleteBtn();
}

function updateBulkDeleteBtn() {
  const count = document.querySelectorAll('#applicant-tbody .row-check:checked').length;
  const btn   = document.getElementById('bulk-delete-btn');
  const span  = document.getElementById('selected-count');
  if (!btn) return;
  btn.style.display = count > 0 ? '' : 'none';
  if (span) span.textContent = count;
}

async function deleteSelected() {
  const checked = Array.from(document.querySelectorAll('#applicant-tbody .row-check:checked'));
  if (!checked.length) return;

  const ids   = checked.map(cb => parseInt(cb.closest('tr').dataset.id));
  const names = checked.map(cb => cb.closest('tr').querySelector('.applicant-name').textContent.trim());

  if (!confirm(`선택된 ${ids.length}명을 삭제하시겠습니까?\n\n${names.join(', ')}\n\n업로드된 서류도 함께 삭제됩니다.`)) return;

  let failed = 0;
  for (const id of ids) {
    try {
      const res = await apiCall(`/round/${ROUND_ID}/applicants/${id}`, 'DELETE');
      if (res.success) {
        const row = document.querySelector(`tr.applicant-row[data-id="${id}"]`);
        if (row) row.remove();
        const statusRow = document.querySelector(`#status-tbody tr[data-id="${id}"]`);
        if (statusRow) statusRow.remove();
      } else {
        failed++;
      }
    } catch { failed++; }
  }

  // 빈 테이블 처리
  const tbody = document.getElementById('applicant-tbody');
  if (tbody && tbody.querySelectorAll('tr.applicant-row').length === 0) {
    tbody.innerHTML = `<tr id="empty-row"><td colspan="14" class="empty-table-cell">신청자를 추가하세요</td></tr>`;
  }

  // 순서 번호 재정렬
  document.querySelectorAll('#applicant-tbody tr.applicant-row').forEach((r, i) => {
    const cell = r.querySelector('.sort-order-cell');
    if (cell) cell.textContent = i + 1;
  });

  updateBulkDeleteBtn();
  const master = document.getElementById('check-all');
  if (master) { master.checked = false; master.indeterminate = false; }

  if (failed > 0) showToast(`${ids.length - failed}명 삭제, ${failed}명 실패`, 'warning');
  else showToast(`${ids.length}명 삭제 완료`, 'success');
}

// ══════════════════════════════════════════════════════
// Applicant Modal
// ══════════════════════════════════════════════════════

function openAddApplicantModal() {
  openModal('신청자 추가', 'add-applicant-form');
}

async function submitAddApplicant(event) {
  event.preventDefault();
  const form = event.target;
  const fd = new FormData(form);
  const payload = {
    name: fd.get('name'),
    exercise_price: fd.get('exercise_price') || null,
    quantity: fd.get('quantity') || null,
    broker: fd.get('broker'),
    account_number: fd.get('account_number'),
  };

  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/add`, 'POST', payload);
    if (res.success) {
      showToast(res.message, 'success');
      closeModal();
      addApplicantRow(res.data);
      form.reset();
      // Remove empty row if present
      const emptyRow = document.getElementById('empty-row');
      if (emptyRow) emptyRow.remove();
    } else {
      showToast(res.message, 'error');
    }
  } catch (e) {
    showToast('오류가 발생했습니다: ' + e.message, 'error');
  }
}

function addApplicantRow(ap) {
  const tbody = document.getElementById('applicant-tbody');
  const tr = document.createElement('tr');
  tr.className = 'applicant-row';
  tr.dataset.id = ap.id;
  tr.draggable = true;

  const priceStr = ap.exercise_price ? Number(ap.exercise_price).toLocaleString('ko-KR') : '-';
  const qtyStr   = ap.quantity       ? Number(ap.quantity).toLocaleString('ko-KR')       : '-';
  const linkVal  = ap.token_link || '';

  tr.innerHTML = `
    <td class="col-center"><input type="checkbox" class="row-check" onchange="onRowCheckChange()"></td>
    <td class="drag-handle" title="드래그하여 순서 변경">&#9776;</td>
    <td class="col-order sort-order-cell">${ap.sort_order}</td>
    <td class="applicant-name">${escHtml(ap.name)}</td>
    <td>${priceStr}</td>
    <td>${qtyStr}</td>
    <td>${escHtml(ap.broker || '-')}</td>
    <td>${escHtml(ap.account_number || '-')}</td>
    <td class="col-center doc-status" data-id="${ap.id}" data-type="application">
      <span class="doc-x">&#10007;</span>
    </td>
    <td class="col-center doc-status" data-id="${ap.id}" data-type="id_copy">
      <span class="doc-x">&#10007;</span>
    </td>
    <td class="col-center doc-status" data-id="${ap.id}" data-type="account_copy">
      <span class="doc-x">&#10007;</span>
    </td>
    <td>
      <div class="link-cell">
        <input type="text" class="link-input" readonly
               value="${escHtml(linkVal)}" id="link-${ap.id}">
        <button class="btn btn-outline btn-xs" onclick="copyLink('link-${ap.id}')">복사</button>
      </div>
    </td>
  `;

  tbody.appendChild(tr);
  attachDragEvents(tr);

  // Also add to status table
  const statusTbody = document.getElementById('status-tbody');
  if (statusTbody) {
    const emptyStatusRow = statusTbody.querySelector('td[colspan]');
    if (emptyStatusRow) emptyStatusRow.closest('tr').remove();

    const sr = document.createElement('tr');
    sr.dataset.id = ap.id;
    sr.innerHTML = `
      <td>${ap.sort_order}</td>
      <td>${escHtml(ap.name)}</td>
      <td class="col-center status-application-${ap.id}"><span class="doc-x">&#10007;</span></td>
      <td class="col-center status-id_copy-${ap.id}"><span class="doc-x">&#10007;</span></td>
      <td class="col-center status-account_copy-${ap.id}"><span class="doc-x">&#10007;</span></td>
      <td class="col-center"><span class="badge badge-neutral">미완료</span></td>
      <td>
        <input type="text" class="link-input link-small" readonly
               value="${escHtml(linkVal)}" id="slink-${ap.id}">
        <button class="btn btn-outline btn-xs" onclick="copyLink('slink-${ap.id}')">복사</button>
      </td>
    `;
    statusTbody.appendChild(sr);
  }
}

// ══════════════════════════════════════════════════════
// Delete applicant
// ══════════════════════════════════════════════════════

async function deleteApplicant(applicantId, name) {
  if (!confirm(`"${name}" 신청자를 삭제하시겠습니까?\n업로드된 서류도 모두 삭제됩니다.`)) return;

  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/${applicantId}`, 'DELETE');
    if (res.success) {
      showToast(res.message, 'success');
      // Remove from main table
      const row = document.querySelector(`tr.applicant-row[data-id="${applicantId}"]`);
      if (row) row.remove();
      // Remove from status table
      const statusRow = document.querySelector(`#status-tbody tr[data-id="${applicantId}"]`);
      if (statusRow) statusRow.remove();
      // Show empty row if needed
      const tbody = document.getElementById('applicant-tbody');
      if (tbody && tbody.children.length === 0) {
        tbody.innerHTML = `<tr id="empty-row"><td colspan="12" class="empty-table-cell">신청자를 추가하세요</td></tr>`;
      }
    } else {
      showToast(res.message, 'error');
    }
  } catch (e) {
    showToast('삭제 오류: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════
// Drag-to-reorder
// ══════════════════════════════════════════════════════

let dragSrc = null;

function attachDragEvents(row) {
  row.addEventListener('dragstart', onDragStart);
  row.addEventListener('dragover',  onDragOver);
  row.addEventListener('dragleave', onDragLeave);
  row.addEventListener('drop',      onDrop);
  row.addEventListener('dragend',   onDragEnd);
}

function onDragStart(e) {
  dragSrc = this;
  this.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', this.dataset.id);
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  if (this !== dragSrc) this.classList.add('drag-over');
}

function onDragLeave() {
  this.classList.remove('drag-over');
}

function onDrop(e) {
  e.preventDefault();
  this.classList.remove('drag-over');
  if (this === dragSrc) return;

  const tbody = document.getElementById('applicant-tbody');
  const rows  = Array.from(tbody.querySelectorAll('tr.applicant-row'));
  const srcIdx = rows.indexOf(dragSrc);
  const tgtIdx = rows.indexOf(this);

  if (srcIdx < tgtIdx) {
    this.after(dragSrc);
  } else {
    this.before(dragSrc);
  }

  saveNewOrder();
}

function onDragEnd() {
  this.classList.remove('dragging');
  document.querySelectorAll('.applicant-row').forEach(r => r.classList.remove('drag-over'));
}

function saveNewOrder() {
  const tbody = document.getElementById('applicant-tbody');
  const rows  = Array.from(tbody.querySelectorAll('tr.applicant-row'));
  const ids   = rows.map(r => parseInt(r.dataset.id));

  // Update displayed sort_order numbers
  rows.forEach((r, i) => {
    const cell = r.querySelector('.sort-order-cell');
    if (cell) cell.textContent = i + 1;
  });

  apiCall(`/round/${ROUND_ID}/applicants/reorder`, 'POST', { order: ids })
    .then(res => {
      if (!res.success) showToast(res.message, 'error');
    })
    .catch(() => showToast('순서 저장 오류', 'error'));
}

// ══════════════════════════════════════════════════════
// File upload — drag-drop zones
// ══════════════════════════════════════════════════════

function handleDragOver(event) {
  event.preventDefault();
  event.currentTarget.classList.add('dragover');
}

function handleDragLeave(event) {
  event.currentTarget.classList.remove('dragover');
}

function handleDrop(event, docType) {
  event.preventDefault();
  event.currentTarget.classList.remove('dragover');
  const files = event.dataTransfer.files;
  processFiles(files, docType);
}

function handleFileSelect(event, docType) {
  processFiles(event.target.files, docType);
  event.target.value = '';
}

function processFiles(files, docType) {
  if (!files || files.length === 0) return;

  const listEl = document.getElementById('files-' + docType);
  Array.from(files).forEach(async file => {
    const item = { file, docType, status: 'pending', matchedId: null, matchedName: null };

    // 1단계: 파일명으로 자동 매칭
    const names = getApplicantNames();
    for (const [id, name] of Object.entries(names)) {
      if (file.name.includes(name)) {
        item.matchedId   = id;
        item.matchedName = name;
        break;
      }
    }

    // 2단계: 파일명 매칭 실패 && PDF 파일이면 → 내용에서 이름 추출
    if (!item.matchedId && file.type === 'application/pdf' && docType === 'application') {
      item.extracting = true;  // 추출 중 표시용
      pendingUploads[docType].push(item);
      const div = renderFileItem(listEl, item, docType);

      // PDF 내용에서 이름 추출
      const extracted = await extractNameFromPDF(file);
      item.extracting = false;

      if (extracted && extracted.success) {
        item.matchedId = extracted.matched_id;
        item.matchedName = extracted.matched_name;
        item.extractedName = extracted.extracted_name;

        // UI 업데이트
        const select = div.querySelector('[data-file-item]');
        if (select) select.value = extracted.matched_id;

        const statusSpan = div.querySelector('.file-list-item-status');
        if (statusSpan) {
          statusSpan.textContent = `자동매칭: ${extracted.matched_name}`;
          statusSpan.className = 'file-list-item-status success';
        }

        // 자동 업로드
        const uploadBtn = div.querySelector('.upload-btn');
        if (uploadBtn) uploadBtn.click();
      } else {
        // 추출 실패 - 수동 선택 필요
        const statusSpan = div.querySelector('.file-list-item-status');
        if (statusSpan) {
          statusSpan.textContent = extracted?.message || '수동 선택 필요';
          statusSpan.className = 'file-list-item-status pending';
        }
      }
    } else {
      pendingUploads[docType].push(item);
      renderFileItem(listEl, item, docType);
    }
  });
}

async function extractNameFromPDF(file) {
  /**
   * PDF에서 이름 추출 API 호출
   */
  try {
    const fd = new FormData();
    fd.append('file', file);
    const res = await apiCall(`/round/${ROUND_ID}/extract_name_from_pdf`, 'POST', fd);
    return res;
  } catch (e) {
    console.error('PDF 이름 추출 오류:', e);
    return { success: false, message: e.message };
  }
}

function getApplicantNames() {
  const result = {};
  document.querySelectorAll('tr.applicant-row').forEach(row => {
    const id   = row.dataset.id;
    const name = row.querySelector('.applicant-name')?.textContent?.trim();
    if (id && name) result[id] = name;
  });
  return result;
}

function renderFileItem(listEl, item, docType) {
  const div = document.createElement('div');
  div.className = 'file-list-item';

  const docLabels = { application: '신청서', id_copy: '신분증사본', account_copy: '계좌사본' };

  const names = getApplicantNames();
  let assignSelect = `<select class="form-control" style="font-size:12px;padding:3px 6px;width:auto;min-width:100px;" data-file-item>
    <option value="">배정 미정</option>`;
  for (const [id, name] of Object.entries(names)) {
    const sel = item.matchedId == id ? 'selected' : '';
    assignSelect += `<option value="${id}" ${sel}>${name}</option>`;
  }
  assignSelect += `</select>`;

  const statusText = item.extracting ? '이름 추출 중...' : '대기';
  const statusClass = item.extracting ? 'file-list-item-status extracting' : 'file-list-item-status pending';

  div.innerHTML = `
    <div class="file-list-item-name" title="${escHtml(item.file.name)}">${escHtml(item.file.name)}</div>
    ${assignSelect}
    <button class="btn btn-primary btn-xs upload-btn">업로드</button>
    <span class="${statusClass}" id="status-${item.file.name.replace(/[^a-zA-Z0-9]/g,'_')}">${statusText}</span>
  `;

  const uploadBtn = div.querySelector('.upload-btn');
  uploadBtn.addEventListener('click', async () => {
    const select = div.querySelector('[data-file-item]');
    const applicantId = select ? select.value : '';
    if (!applicantId) {
      showToast('신청자를 선택하세요', 'warning');
      return;
    }
    uploadBtn.disabled = true;
    uploadBtn.textContent = '업로드 중...';
    const statusSpan = div.querySelector('.file-list-item-status');

    const fd = new FormData();
    fd.append('file', item.file);
    try {
      const res = await apiCall(`/round/${ROUND_ID}/upload/${applicantId}/${docType}`, 'POST', fd);
      if (res.success) {
        statusSpan.textContent = '완료';
        statusSpan.className = 'file-list-item-status success';
        uploadBtn.textContent = '완료';
        // Update doc status cells
        updateDocStatus(parseInt(applicantId), docType, true);
        showToast(`${item.file.name} 업로드 완료`, 'success');
      } else {
        statusSpan.textContent = '실패';
        statusSpan.className = 'file-list-item-status error';
        uploadBtn.textContent = '재시도';
        uploadBtn.disabled = false;
        showToast(res.message, 'error');
      }
    } catch (e) {
      statusSpan.textContent = '오류';
      statusSpan.className = 'file-list-item-status error';
      uploadBtn.textContent = '재시도';
      uploadBtn.disabled = false;
      showToast('업로드 오류: ' + e.message, 'error');
    }
  });

  listEl.appendChild(div);

  // If matched, auto-upload (extracting 중이 아닐 때만)
  if (item.matchedId && !item.extracting) {
    uploadBtn.click();
  }

  return div;  // processFiles에서 사용하기 위해 반환
}

// ══════════════════════════════════════════════════════
// Update doc status icons
// ══════════════════════════════════════════════════════

function updateDocStatus(applicantId, docType, hasDoc) {
  const checkHtml = hasDoc ? '<span class="doc-check">&#10003;</span>' : '<span class="doc-x">&#10007;</span>';
  // Main table
  const cell = document.querySelector(`.doc-status[data-id="${applicantId}"][data-type="${docType}"]`);
  if (cell) cell.innerHTML = checkHtml;
  // Status table
  const statusCell = document.querySelector(`.status-${docType}-${applicantId}`);
  if (statusCell) statusCell.innerHTML = checkHtml;

  // Update merge counts
  refreshMergeCounts();
}

function updateOcrResults(applicantId, ocrData) {
  // Update RRN (실명번호)
  if (ocrData.rrn) {
    const rrnCell = document.querySelector(`.ocr-rrn[data-id="${applicantId}"]`);
    if (rrnCell) {
      rrnCell.textContent = ocrData.rrn;
      rrnCell.style.color = '#2563eb'; // 파란색으로 강조
      setTimeout(() => { rrnCell.style.color = '#555'; }, 2000);
    }
  }
  // Update account number
  if (ocrData.ocr_account) {
    const accountCell = document.querySelector(`.ocr-account[data-id="${applicantId}"]`);
    if (accountCell) {
      accountCell.textContent = ocrData.ocr_account;
      accountCell.style.color = '#2563eb';
      setTimeout(() => { accountCell.style.color = '#555'; }, 2000);
    }
  }
}

async function refreshMergeCounts() {
  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/status`);
    if (!res.success) return;
    const data = res.data;
    const counts = { application: 0, id_copy: 0, account_copy: 0 };
    data.forEach(ap => {
      if (ap.application)   counts.application++;
      if (ap.id_copy)       counts.id_copy++;
      if (ap.account_copy)  counts.account_copy++;
    });
    for (const [dt, cnt] of Object.entries(counts)) {
      const el = document.getElementById('cnt-' + dt);
      if (el) el.textContent = cnt;
    }
  } catch (e) {}
}

// ══════════════════════════════════════════════════════
// Auto-match files
// ══════════════════════════════════════════════════════

function autoMatchFiles() {
  const names = getApplicantNames();
  let matched = 0;
  let total   = 0;

  for (const dt of ['application', 'id_copy', 'account_copy']) {
    const listEl = document.getElementById('files-' + dt);
    if (!listEl) continue;
    pendingUploads[dt].forEach(item => {
      total++;
      for (const [id, name] of Object.entries(names)) {
        if (item.file.name.includes(name)) {
          item.matchedId   = id;
          item.matchedName = name;
          matched++;
          // Update the select element in the list
          const selects = listEl.querySelectorAll('[data-file-item]');
          selects.forEach(sel => {
            // crude match by file list position, just re-set all matching filename
            const fileNameEl = sel.closest('.file-list-item')?.querySelector('.file-list-item-name');
            if (fileNameEl && fileNameEl.title === item.file.name) {
              sel.value = id;
            }
          });
          break;
        }
      }
    });
  }

  const resultEl = document.getElementById('match-results');
  if (resultEl) {
    resultEl.textContent = `자동매칭 결과: 전체 ${total}건 중 ${matched}건 매칭됨`;
  }
  showToast(`${matched}/${total}건 자동매칭 완료`, matched > 0 ? 'success' : 'warning');
}

// ══════════════════════════════════════════════════════
// Copy all links
// ══════════════════════════════════════════════════════

function copyAllLinks() {
  const inputs = document.querySelectorAll('#status-tbody .link-input');
  const links = Array.from(inputs).map(i => i.value).join('\n');
  if (!links.trim()) {
    showToast('복사할 링크가 없습니다', 'warning');
    return;
  }
  copyText(links);
}

// ══════════════════════════════════════════════════════
// Refresh submission status
// ══════════════════════════════════════════════════════

async function refreshStatus() {
  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/status`);
    if (!res.success) { showToast(res.message, 'error'); return; }

    const docTypes = ['application', 'id_copy', 'account_copy'];
    res.data.forEach(ap => {
      docTypes.forEach(dt => {
        updateDocStatus(ap.applicant_id, dt, ap[dt]);
      });
      // Update overall status badge in status table
      const row = document.querySelector(`#status-tbody tr[data-id="${ap.applicant_id}"]`);
      if (row) {
        const badgeCell = row.children[5];
        if (badgeCell) {
          badgeCell.innerHTML = ap.all_submitted
            ? '<span class="badge badge-success">완료</span>'
            : '<span class="badge badge-neutral">미완료</span>';
        }
      }
    });
    showToast('현황이 업데이트되었습니다', 'success');
  } catch (e) {
    showToast('새로고침 오류: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════
// PDF Merge
// ══════════════════════════════════════════════════════

async function runMerge(roundId) {
  const btn = document.getElementById('merge-btn');
  btn.disabled = true;
  btn.textContent = '병합 중...';

  try {
    const res = await apiCall(`/round/${roundId}/step01/merge`, 'POST');
    if (!res.success) {
      showToast(res.message, 'error');
      btn.disabled = false;
      btn.textContent = '&#128196; PDF 병합 실행';
      return;
    }

    const mergeResults = document.getElementById('merge-results');
    let html = '<h3 class="merge-results-title">병합 완료 결과물</h3><div class="download-list">';

    for (const [dt, info] of Object.entries(res.data)) {
      if (info.success) {
        const sizeStr  = formatBytes(info.size);
        const pageStr  = info.pages ? `${info.pages}페이지` : '';
        html += `
          <div class="download-item">
            <div>
              <div class="download-item-name">${escHtml(info.label)} 합본</div>
              <div class="download-item-meta">${escHtml(info.filename)} &nbsp;|&nbsp; ${pageStr} &nbsp;|&nbsp; ${sizeStr}</div>
            </div>
            <a href="${info.download_url}" class="btn btn-success btn-sm" download>
              &#11123; 다운로드
            </a>
          </div>`;
      } else {
        html += `
          <div class="download-item" style="border-color:#fecaca;background:#fef2f2;">
            <div class="download-item-name" style="color:#991b1b;">${escHtml(info.label)} 오류: ${escHtml(info.message || '')}</div>
          </div>`;
      }
    }
    html += '</div>';
    mergeResults.innerHTML = html;
    showToast('PDF 병합이 완료되었습니다', 'success');
  } catch (e) {
    showToast('병합 오류: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '&#128196; PDF 병합 실행';
  }
}

// ══════════════════════════════════════════════════════
// Excel Import
// ══════════════════════════════════════════════════════

let _importedData = [];

async function importFromExcel(event) {
  const file = event.target.files[0];
  event.target.value = '';  // reset so same file can be re-selected
  if (!file) return;

  showToast('엑셀 파일 파싱 중...', 'info', 2000);

  const fd = new FormData();
  fd.append('file', file);

  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/import-excel`, 'POST', fd);
    if (!res.success) {
      showToast(res.message, 'error');
      return;
    }

    _importedData = res.data;

    // 미리보기 테이블 채우기
    const tbody = document.getElementById('import-preview-tbody');
    tbody.innerHTML = '';
    _importedData.forEach((row, i) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${i + 1}</td>
        <td><strong>${escHtml(row.name)}</strong></td>
        <td>${escHtml(row.grant_date || '-')}</td>
        <td>${row.exercise_price ? Number(row.exercise_price).toLocaleString('ko-KR') : '-'}</td>
        <td>${row.quantity ? Number(row.quantity).toLocaleString('ko-KR') : '-'}</td>
        <td>${escHtml(row.exercise_date || '-')}</td>
        <td>${escHtml(row.broker || '-')}</td>
        <td>${escHtml(row.account_number || '-')}</td>
      `;
      tbody.appendChild(tr);
    });

    const filteredMsg = res.filtered_out > 0
      ? ` (다른 행사일 ${res.filtered_out}행 제외됨)`
      : '';
    document.getElementById('import-count-msg').textContent =
      `행사일 ${res.exercise_date} 기준 ${res.count}명 인식됨${filteredMsg}. 내용 확인 후 추가하세요.`;

    openModal('엑셀 신청자 명단 미리보기', 'excel-import-form');
  } catch (e) {
    showToast('파싱 오류: ' + e.message, 'error');
  }
}

async function confirmImport() {
  if (!_importedData.length) return;

  const modeEl = document.querySelector('input[name="import-mode"]:checked');
  const mode = modeEl ? modeEl.value : 'append';

  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicants/import-confirm`, 'POST', {
      applicants: _importedData,
      mode,
    });
    if (res.success) {
      showToast(res.message, 'success');
      closeModal();
      // 페이지 새로고침으로 명단 반영
      setTimeout(() => location.reload(), 600);
    } else {
      showToast(res.message || '저장 실패', 'error');
    }
  } catch (e) {
    showToast('저장 오류: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════
// Utility
// ══════════════════════════════════════════════════════

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ══════════════════════════════════════════════════════
// 결과물 삭제
// ══════════════════════════════════════════════════════

async function deleteOutput(outputId, filename) {
  if (!confirm(`"${filename}"을(를) 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다.`)) return;

  try {
    const res = await apiCall(`/round/${ROUND_ID}/output/${outputId}`, 'DELETE');
    if (res.success) {
      // UI에서 제거
      const item = document.querySelector(`.download-item[data-output-id="${outputId}"]`);
      if (item) item.remove();

      // 목록이 비었으면 전체 섹션 숨김
      const list = document.getElementById('output-list');
      if (list && list.children.length === 0) {
        const results = document.getElementById('merge-results');
        if (results) results.innerHTML = '';
      }

      showToast('결과물이 삭제되었습니다', 'success');
    } else {
      showToast(res.message || '삭제 실패', 'error');
    }
  } catch (e) {
    showToast('삭제 오류: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════
// 서류 관리 (조회 및 삭제)
// ══════════════════════════════════════════════════════

async function openDocumentManager(applicantId, docType, applicantName) {
  const docLabels = {
    application: '신청서',
    id_copy: '신분증사본',
    account_copy: '계좌사본'
  };

  const docLabel = docLabels[docType] || docType;
  const infoEl = document.getElementById('doc-manager-info');
  const listEl = document.getElementById('doc-manager-list');

  infoEl.textContent = `${applicantName}님의 ${docLabel}`;
  listEl.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">로딩 중...</div>';

  openModal('서류 관리', 'document-manager');

  // 서류 목록 조회
  try {
    const res = await apiCall(`/round/${ROUND_ID}/applicant/${applicantId}/documents`, 'GET');
    if (!res.success) {
      listEl.innerHTML = `<div style="text-align:center;padding:20px;color:#d32f2f;">${res.message || '조회 실패'}</div>`;
      return;
    }

    const docs = res.documents.filter(d => d.doc_type === docType);

    if (docs.length === 0) {
      listEl.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">업로드된 서류가 없습니다</div>';
      return;
    }

    // 서류 목록 렌더링
    listEl.innerHTML = '';
    docs.forEach(doc => {
      const item = document.createElement('div');
      item.style.cssText = 'display:flex;justify-content:space-between;align-items:center;padding:12px;border:1px solid #ddd;border-radius:4px;margin-bottom:8px;background:#fafafa;';

      const nameDiv = document.createElement('div');
      nameDiv.style.cssText = 'flex:1;font-size:14px;color:#333;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;';
      nameDiv.textContent = doc.original_filename || doc.filename;
      nameDiv.title = doc.original_filename || doc.filename;

      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'btn btn-danger btn-xs';
      deleteBtn.textContent = '삭제';
      deleteBtn.style.marginLeft = '12px';
      deleteBtn.onclick = async () => {
        if (!confirm(`"${doc.original_filename || doc.filename}"을(를) 삭제하시겠습니까?`)) return;

        deleteBtn.disabled = true;
        deleteBtn.textContent = '삭제 중...';

        try {
          const delRes = await apiCall(`/round/${ROUND_ID}/document/${doc.id}`, 'DELETE');
          if (delRes.success) {
            item.remove();
            // 서류 상태 업데이트
            updateDocStatus(applicantId, docType, docs.length > 1);
            showToast('서류가 삭제되었습니다', 'success');

            // 목록이 비었으면 안내 메시지 표시
            if (listEl.children.length === 0) {
              listEl.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">업로드된 서류가 없습니다</div>';
            }
          } else {
            showToast(delRes.message || '삭제 실패', 'error');
            deleteBtn.disabled = false;
            deleteBtn.textContent = '삭제';
          }
        } catch (e) {
          showToast('삭제 오류: ' + e.message, 'error');
          deleteBtn.disabled = false;
          deleteBtn.textContent = '삭제';
        }
      };

      item.appendChild(nameDiv);
      item.appendChild(deleteBtn);
      listEl.appendChild(item);
    });

  } catch (e) {
    listEl.innerHTML = `<div style="text-align:center;padding:20px;color:#d32f2f;">조회 오류: ${e.message}</div>`;
  }
}

// ══════════════════════════════════════════════════════
// Init on DOM ready
// ══════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  // Attach drag events to existing applicant rows
  document.querySelectorAll('tr.applicant-row').forEach(attachDragEvents);

  // Initial merge count refresh
  refreshMergeCounts();
});
