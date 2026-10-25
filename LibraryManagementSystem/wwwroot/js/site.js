/* ═══════════════════════════════════════════════════════════════
   SITE.JS — Global functionality
   Sidebar toggle, toasts, notifications, CSRF, dropdowns
   ═══════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', function () {
    initSidebar();
    initDropdowns();
    initToasts();
    initNotifications();
    initConfirmDialogs();
});

/* ── Sidebar Toggle ────────────────────────────────────────────── */
function initSidebar() {
    const hamburger = document.getElementById('hamburger');
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('sidebar-overlay');

    if (hamburger) {
        hamburger.addEventListener('click', () => {
            sidebar?.classList.toggle('open');
            overlay?.classList.toggle('show');
        });
    }

    if (overlay) {
        overlay.addEventListener('click', () => {
            sidebar?.classList.remove('open');
            overlay.classList.remove('show');
        });
    }
}

/* ── User Dropdown ─────────────────────────────────────────────── */
function initDropdowns() {
    document.querySelectorAll('.user-dropdown').forEach(trigger => {
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            const menu = trigger.querySelector('.dropdown-menu');
            document.querySelectorAll('.dropdown-menu.show').forEach(m => {
                if (m !== menu) m.classList.remove('show');
            });
            menu?.classList.toggle('show');
        });
    });

    document.addEventListener('click', () => {
        document.querySelectorAll('.dropdown-menu.show').forEach(m => m.classList.remove('show'));
    });
}

/* ── Toast Notifications ───────────────────────────────────────── */
function initToasts() {
    const container = document.getElementById('toast-container');
    if (!container) return;

    // Auto-show toasts from TempData
    document.querySelectorAll('.toast').forEach(toast => {
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100%)';
            setTimeout(() => toast.remove(), 300);
        }, 5000);
    });
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <span>${message}</span>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(100%)';
        toast.style.transition = 'all 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}

/* ── Notification Polling ──────────────────────────────────────── */
function initNotifications() {
    const badge = document.getElementById('notification-badge');
    if (!badge) return;

    function fetchCount() {
        fetch('/Notifications/UnreadCount')
            .then(r => r.json())
            .then(data => {
                if (data.count > 0) {
                    badge.textContent = data.count > 9 ? '9+' : data.count;
                    badge.style.display = 'flex';
                } else {
                    badge.style.display = 'none';
                }
            })
            .catch(() => {});
    }

    fetchCount();
    setInterval(fetchCount, 60000);
}

/* ── Confirmation Dialogs ──────────────────────────────────────── */
function initConfirmDialogs() {
    document.querySelectorAll('[data-confirm]').forEach(el => {
        el.addEventListener('click', (e) => {
            const message = el.getAttribute('data-confirm') || 'Are you sure?';
            if (!confirm(message)) {
                e.preventDefault();
                e.stopPropagation();
            }
        });
    });
}

/* ── CSRF Token Helper ─────────────────────────────────────────── */
function getCsrfToken() {
    const token = document.querySelector('input[name="__RequestVerificationToken"]');
    return token ? token.value : '';
}

function fetchWithCsrf(url, options = {}) {
    options.headers = options.headers || {};
    options.headers['RequestVerificationToken'] = getCsrfToken();
    return fetch(url, options);
}
