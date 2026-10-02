/**
 * Farmacia ni Dok - Dynamic Conditional Table Pagination
 * Rule: Only display pagination controls if total matching rows > PAGE_SIZE (default: 10).
 * If total rows <= PAGE_SIZE, pagination is completely hidden.
 * Delegates to window.paginateTable for consistent pagination across all pages.
 */

(function () {
    window.initTablePagination = function() {
        if (typeof window.paginateTable === 'function') {
            const tables = document.querySelectorAll('table');
            tables.forEach(table => {
                if (!table.closest('.modal, .modal-content, .dispose-modal, [class*="modal"]') && table.querySelector('tbody')) {
                    const rowsPerPage = parseInt(table.dataset.rowsPerPage || '10', 10);
                    window.paginateTable(table, rowsPerPage);
                }
            });
        }
    };
})();
