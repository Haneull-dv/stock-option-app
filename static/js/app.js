/* ═══════════════════════════════════════════════════════════
   StockOps — Global JavaScript
   ═══════════════════════════════════════════════════════════ */

// ── Toast ────────────────────────────────────────────────────
function showToast(message, type = 'info', duration = 3000) {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ── Modal ────────────────────────────────────────────────────
function openModal(title, contentId) {
  const overlay = document.getElementById('modal-overlay');
  const titleEl = document.getElementById('modal-title');
  const bodyEl  = document.getElementById('modal-body');
  const source  = document.getElementById(contentId);

  titleEl.textContent = title;
  bodyEl.innerHTML = '';
  if (source) {
    const clone = source.cloneNode(true);
    clone.style.display = 'block';
    bodyEl.appendChild(clone);
  }
  overlay.style.display = 'flex';
  document.body.style.overflow = 'hidden';

  // Re-attach event handlers inside modal (forms etc.)
  const form = bodyEl.querySelector('form');
  if (form && form.id === 'applicant-form') {
    form.onsubmit = submitAddApplicant;
  }
}

function closeModal() {
  const overlay = document.getElementById('modal-overlay');
  if (overlay) overlay.style.display = 'none';
  document.body.style.overflow = '';
}

// Close modal when clicking overlay background
document.addEventListener('DOMContentLoaded', () => {
  const overlay = document.getElementById('modal-overlay');
  if (overlay) {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal();
    });
  }
});

// ── Tab switching ────────────────────────────────────────────
function switchTab(tabId, btn) {
  // hide all tab contents
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  const tab = document.getElementById(tabId);
  if (tab) tab.classList.add('active');
  if (btn) btn.classList.add('active');
}

// ── Clipboard ────────────────────────────────────────────────
function copyLink(inputId) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.select();
  input.setSelectionRange(0, 99999);
  try {
    navigator.clipboard.writeText(input.value).then(() => {
      showToast('링크가 복사되었습니다', 'success');
    }).catch(() => {
      document.execCommand('copy');
      showToast('링크가 복사되었습니다', 'success');
    });
  } catch (e) {
    document.execCommand('copy');
    showToast('링크가 복사되었습니다', 'success');
  }
}

function copyText(text) {
  try {
    navigator.clipboard.writeText(text).then(() => {
      showToast('복사되었습니다', 'success');
    });
  } catch (e) {
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    showToast('복사되었습니다', 'success');
  }
}

// ── AJAX helper ──────────────────────────────────────────────
async function apiCall(url, method = 'GET', body = null) {
  const opts = {
    method,
    headers: {},
  };
  if (body !== null) {
    if (body instanceof FormData) {
      opts.body = body;
    } else {
      opts.headers['Content-Type'] = 'application/json';
      opts.body = JSON.stringify(body);
    }
  }
  const res = await fetch(url, opts);
  const data = await res.json();
  return data;
}

// ── Format helpers ───────────────────────────────────────────
function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

function formatNum(n) {
  if (!n) return '-';
  return Number(n).toLocaleString('ko-KR');
}
