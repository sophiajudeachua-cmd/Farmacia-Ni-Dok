/**
 * Farmacia ni Dok - Dynamic Conditional Table Pagination
 * Rule: Only display pagination controls if total matching rows > PAGE_SIZE (default: 10).
 * If total rows <= PAGE_SIZE, pagination is completely hidden.
 */

(function () {
    const DEFAULT_PAGE_SIZE = 10;

    function setupTablePagination(table, pageSize = DEFAULT_PAGE_SIZE) {
        if (!table) return;

        const tbody = table.querySelector('tbody');
        if (!tbody) return;

        // Ensure outer card box is found so pagination is placed OUTSIDE below it!
        let outerBox = table.closest('.panel, .card, .table-card, .table-responsive, .inventory-card, .suppliers-card, .po-card, .reports-card') || table.parentElement;
        if (!outerBox) return;

        // Check or create pagination container element outside below outerBox
        let paginationWrapper = outerBox.nextElementSibling && outerBox.nextElementSibling.classList && outerBox.nextElementSibling.classList.contains('pagination-container')
            ? outerBox.nextElementSibling
            : null;

        if (!paginationWrapper && outerBox.parentNode) {
            paginationWrapper = document.createElement('div');
            paginationWrapper.className = 'pagination-container';
            paginationWrapper.style.display = 'none';
            outerBox.parentNode.insertBefore(paginationWrapper, outerBox.nextSibling);
        }

        let currentPage = 1;

        function render() {
            // Get all rows that are not explicitly hidden by external search filtering
            const allRows = Array.from(tbody.querySelectorAll('tr'));
            
            // Filter out rows marked as filtered-out by search if any
            const matchingRows = allRows.filter(row => {
                if (row.dataset.filteredHidden === 'true') {
                    row.style.display = 'none';
                    delete row.dataset.paginatedFiltered;
                    return false;
                }
                return row.style.display !== 'none' || row.dataset.paginatedFiltered === 'true';
            });

            // Reset dataset flag for matching rows
            matchingRows.forEach(r => delete r.dataset.paginatedFiltered);

            const totalItems = matchingRows.length;

            // CONDITIONAL RULE: Hide pagination if total items <= pageSize
            if (totalItems <= pageSize) {
                // Show all matching rows
                matchingRows.forEach(row => {
                    row.style.display = '';
                });
                paginationWrapper.style.display = 'none';
                paginationWrapper.innerHTML = '';
                return;
            }

            // Total items > pageSize: Show pagination!
            paginationWrapper.style.display = 'flex';
            paginationWrapper.style.justifyContent = 'center';
            paginationWrapper.style.alignItems = 'center';

            const totalPages = Math.ceil(totalItems / pageSize);

            if (currentPage > totalPages) currentPage = totalPages;
            if (currentPage < 1) currentPage = 1;

            const startIndex = (currentPage - 1) * pageSize;
            const endIndex = Math.min(startIndex + pageSize, totalItems);

            // Display current page items, hide others
            matchingRows.forEach((row, index) => {
                if (index >= startIndex && index < endIndex) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                    row.dataset.paginatedFiltered = 'true';
                }
            });

            // Render Pagination Controls HTML (Centered without info text)
            paginationWrapper.innerHTML = `
                <div class="pagination-controls">
                    <button class="pagination-btn prev-btn" ${currentPage === 1 ? 'disabled' : ''}>
                        <i class="fa-solid fa-chevron-left"></i> Prev
                    </button>
                    <div class="page-numbers" style="display: flex; gap: 4px;">
                        ${generatePageButtons(currentPage, totalPages)}
                    </div>
                    <button class="pagination-btn next-btn" ${currentPage === totalPages ? 'disabled' : ''}>
                        Next <i class="fa-solid fa-chevron-right"></i>
                    </button>
                </div>
            `;

            // Attach event listeners
            const prevBtn = paginationWrapper.querySelector('.prev-btn');
            if (prevBtn && !prevBtn.disabled) {
                prevBtn.addEventListener('click', () => {
                    currentPage--;
                    render();
                });
            }

            const nextBtn = paginationWrapper.querySelector('.next-btn');
            if (nextBtn && !nextBtn.disabled) {
                nextBtn.addEventListener('click', () => {
                    currentPage++;
                    render();
                });
            }

            const numBtns = paginationWrapper.querySelectorAll('.page-num-btn');
            numBtns.forEach(btn => {
                btn.addEventListener('click', () => {
                    currentPage = parseInt(btn.dataset.page, 10);
                    render();
                });
            });
        }

        function generatePageButtons(current, total) {
            let html = '';
            const maxVisible = 5;
            let start = Math.max(1, current - Math.floor(maxVisible / 2));
            let end = Math.min(total, start + maxVisible - 1);

            if (end - start + 1 < maxVisible) {
                start = Math.max(1, end - maxVisible + 1);
            }

            for (let i = start; i <= end; i++) {
                html += `<button class="pagination-btn page-num-btn ${i === current ? 'active' : ''}" data-page="${i}">${i}</button>`;
            }
            return html;
        }

        // Attach listener for search inputs in parent panel or topbar
        const searchInputs = document.querySelectorAll('input[type="text"], input[type="search"], select');
        searchInputs.forEach(input => {
            input.addEventListener('input', () => {
                currentPage = 1;
                setTimeout(render, 50);
            });
            input.addEventListener('change', () => {
                currentPage = 1;
                setTimeout(render, 50);
            });
        });

        // Expose rePaginate method on table
        table.rePaginate = function() {
            currentPage = 1;
            render();
        };

        // Initial render
        render();
    }

    function initAllTablePagination() {
        const tables = document.querySelectorAll('table');
        tables.forEach(table => {
            // Ignore modal internal tables or small pickers if any
            if (!table.closest('.modal-content')) {
                setupTablePagination(table, DEFAULT_PAGE_SIZE);
            }
        });
    }

    window.initTablePagination = initAllTablePagination;

    document.addEventListener('DOMContentLoaded', () => {
        initAllTablePagination();
    });
})();
