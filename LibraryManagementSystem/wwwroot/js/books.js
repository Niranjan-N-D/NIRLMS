/* ═══════════════════════════════════════════════════════════════
   BOOKS.JS — AJAX catalog search & filter
   ═══════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', function () {
    initBookSearch();
});

function initBookSearch() {
    const searchForm = document.getElementById('book-search-form');
    if (!searchForm) return;

    // Debounced live search
    const searchInput = searchForm.querySelector('input[name="searchTerm"]');
    if (searchInput) {
        let debounce;
        searchInput.addEventListener('input', () => {
            clearTimeout(debounce);
            debounce = setTimeout(() => searchForm.submit(), 500);
        });
    }
}
