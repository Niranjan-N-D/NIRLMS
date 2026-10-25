/* ═══════════════════════════════════════════════════════════════
   BORROWING.JS — Issue/Return form autocomplete & lookup
   ═══════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', function () {
    initUserSearch();
    initCopySearch();
    initReturnLookup();
});

/* ── User Search Autocomplete ──────────────────────────────────── */
function initUserSearch() {
    const input = document.getElementById('user-search');
    const results = document.getElementById('user-results');
    const hiddenInput = document.getElementById('UserId');
    if (!input || !results) return;

    let debounce;
    input.addEventListener('input', () => {
        clearTimeout(debounce);
        debounce = setTimeout(() => {
            const q = input.value.trim();
            if (q.length < 2) { results.classList.remove('show'); return; }

            fetch(`/Borrowings/SearchUsers?q=${encodeURIComponent(q)}`)
                .then(r => r.json())
                .then(data => {
                    results.innerHTML = data.map(u => `
                        <div class="autocomplete-item" data-id="${u.id}">
                            <div class="primary-text">${u.fullName}</div>
                            <div class="secondary-text">${u.email}</div>
                        </div>
                    `).join('') || '<div class="autocomplete-item">No users found</div>';
                    results.classList.add('show');

                    results.querySelectorAll('[data-id]').forEach(item => {
                        item.addEventListener('click', () => {
                            hiddenInput.value = item.dataset.id;
                            input.value = item.querySelector('.primary-text').textContent;
                            results.classList.remove('show');
                        });
                    });
                });
        }, 300);
    });

    document.addEventListener('click', (e) => {
        if (!results.contains(e.target) && e.target !== input)
            results.classList.remove('show');
    });
}

/* ── Book Copy Search Autocomplete ─────────────────────────────── */
function initCopySearch() {
    const input = document.getElementById('copy-search');
    const results = document.getElementById('copy-results');
    const hiddenInput = document.getElementById('BookCopyId');
    if (!input || !results) return;

    let debounce;
    input.addEventListener('input', () => {
        clearTimeout(debounce);
        debounce = setTimeout(() => {
            const q = input.value.trim();
            if (q.length < 2) { results.classList.remove('show'); return; }

            fetch(`/Borrowings/SearchCopies?q=${encodeURIComponent(q)}`)
                .then(r => r.json())
                .then(data => {
                    results.innerHTML = data.map(c => `
                        <div class="autocomplete-item" data-id="${c.bookCopyId}">
                            <div class="primary-text">${c.bookTitle}</div>
                            <div class="secondary-text">${c.accessionNumber} | ${c.barcode}</div>
                        </div>
                    `).join('') || '<div class="autocomplete-item">No copies found</div>';
                    results.classList.add('show');

                    results.querySelectorAll('[data-id]').forEach(item => {
                        item.addEventListener('click', () => {
                            hiddenInput.value = item.dataset.id;
                            input.value = item.querySelector('.primary-text').textContent +
                                ' — ' + item.querySelector('.secondary-text').textContent;
                            results.classList.remove('show');
                        });
                    });
                });
        }, 300);
    });

    document.addEventListener('click', (e) => {
        if (!results.contains(e.target) && e.target !== input)
            results.classList.remove('show');
    });
}

/* ── Return Book Lookup ────────────────────────────────────────── */
function initReturnLookup() {
    const input = document.getElementById('return-lookup');
    const btn = document.getElementById('return-lookup-btn');
    if (!input || !btn) return;

    btn.addEventListener('click', () => {
        const q = input.value.trim();
        if (!q) return;

        fetch(`/Borrowings/LookupBorrowing?q=${encodeURIComponent(q)}`)
            .then(r => r.json())
            .then(data => {
                if (data && data.borrowingId) {
                    window.location.href = `/Borrowings/Return/${data.borrowingId}`;
                } else {
                    showToast('No active borrowing found for this barcode/accession number.', 'error');
                }
            });
    });

    input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); btn.click(); }
    });
}
