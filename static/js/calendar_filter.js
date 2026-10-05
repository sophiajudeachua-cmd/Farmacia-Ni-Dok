/**
 * Farmacia ni Dok - Interactive Calendar Month & Year Filter
 */

const CAL_MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const CAL_MONTHS_FULL = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];

class CalendarFilter {
    constructor(wrapperEl, onFilterCallback) {
        this.wrapper = typeof wrapperEl === 'string' ? document.getElementById(wrapperEl) : wrapperEl;
        if (!this.wrapper) return;

        this.onFilter = onFilterCallback || (() => {});
        this.now = new Date();
        this.currentYear = this.now.getFullYear();
        this.viewYear = this.currentYear;
        this.selectedMonth = null; // '01' - '12'
        this.selectedYear = null;  // '2026'

        this.render();
        this.bindEvents();
    }

    render() {
        this.wrapper.innerHTML = `
            <button type="button" class="cal-pill-btn" id="cal-pill-trigger">
                <i class="fa-solid fa-calendar-days" style="color: var(--accent, #c0392b);"></i>
                <span class="cal-pill-text">Filter by Month</span>
                <i class="fa-solid fa-chevron-down cal-arrow-icon" style="font-size: 10px; color: var(--text-muted, #94a3b8); margin-left: 2px;"></i>
            </button>

            <div class="cal-popover" id="cal-popover-menu">
                <div class="cal-popover-header">
                    <button type="button" class="cal-nav-btn" id="cal-prev-year" title="Previous Year">
                        <i class="fa-solid fa-chevron-left"></i>
                    </button>
                    <select class="cal-year-select" id="cal-year-dropdown">
                        ${this.generateYearOptions()}
                    </select>
                    <button type="button" class="cal-nav-btn" id="cal-next-year" title="Next Year">
                        <i class="fa-solid fa-chevron-right"></i>
                    </button>
                </div>

                <div class="cal-months-grid" id="cal-grid">
                    ${this.generateMonthButtons()}
                </div>

                <div class="cal-popover-footer">
                    <button type="button" class="cal-action-btn cal-btn-clear" id="cal-clear-btn">
                        <i class="fa-solid fa-rotate-left"></i> Reset
                    </button>
                    <button type="button" class="cal-action-btn cal-btn-current" id="cal-current-month-btn">
                        This Month
                    </button>
                </div>
            </div>
        `;

        this.pillBtn = this.wrapper.querySelector('#cal-pill-trigger');
        this.pillText = this.wrapper.querySelector('.cal-pill-text');
        this.popover = this.wrapper.querySelector('#cal-popover-menu');
        this.yearSelect = this.wrapper.querySelector('#cal-year-dropdown');
        this.grid = this.wrapper.querySelector('#cal-grid');
    }

    generateYearOptions() {
        let html = '';
        const startYear = this.currentYear - 4;
        const endYear = this.currentYear + 2;
        for (let y = endYear; y >= startYear; y--) {
            html += `<option value="${y}" ${y === this.viewYear ? 'selected' : ''}>${y}</option>`;
        }
        return html;
    }

    generateMonthButtons() {
        let html = '';
        CAL_MONTHS_SHORT.forEach((mName, idx) => {
            const mVal = String(idx + 1).padStart(2, '0');
            const isSelected = (this.selectedYear === String(this.viewYear) && this.selectedMonth === mVal);
            html += `<button type="button" class="cal-month-btn ${isSelected ? 'active' : ''}" data-month="${mVal}">${mName}</button>`;
        });
        return html;
    }

    updateGrid() {
        if (this.grid) {
            this.grid.innerHTML = this.generateMonthButtons();
        }
        if (this.yearSelect) {
            this.yearSelect.value = this.viewYear;
        }
    }

    updatePillDisplay() {
        if (this.selectedMonth && this.selectedYear) {
            const mIdx = parseInt(this.selectedMonth, 10) - 1;
            const fullMonth = CAL_MONTHS_FULL[mIdx] || '';
            this.pillText.textContent = `${fullMonth} ${this.selectedYear}`;
            this.pillBtn.classList.add('has-value');
        } else if (this.selectedYear) {
            this.pillText.textContent = `Year ${this.selectedYear}`;
            this.pillBtn.classList.add('has-value');
        } else {
            this.pillText.textContent = `Filter by Month`;
            this.pillBtn.classList.remove('has-value');
        }
    }

    bindEvents() {
        // Toggle popover
        this.pillBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const isOpen = this.popover.style.display === 'block';
            document.querySelectorAll('.cal-popover').forEach(p => p.style.display = 'none');
            this.popover.style.display = isOpen ? 'none' : 'block';
        });

        // Close on click outside
        document.addEventListener('click', (e) => {
            if (!this.wrapper.contains(e.target)) {
                this.popover.style.display = 'none';
            }
        });

        // Prev Year
        this.wrapper.querySelector('#cal-prev-year').addEventListener('click', (e) => {
            e.stopPropagation();
            this.viewYear--;
            this.updateGrid();
        });

        // Next Year
        this.wrapper.querySelector('#cal-next-year').addEventListener('click', (e) => {
            e.stopPropagation();
            this.viewYear++;
            this.updateGrid();
        });

        // Year select change
        this.yearSelect.addEventListener('change', (e) => {
            this.viewYear = parseInt(e.target.value, 10);
            this.updateGrid();
        });

        // Month click
        this.grid.addEventListener('click', (e) => {
            const btn = e.target.closest('.cal-month-btn');
            if (!btn) return;
            e.stopPropagation();

            const monthVal = btn.dataset.month;
            this.selectedMonth = monthVal;
            this.selectedYear = String(this.viewYear);

            this.updateGrid();
            this.updatePillDisplay();
            this.popover.style.display = 'none';

            this.onFilter(this.selectedMonth, this.selectedYear);
        });

        // Reset / Clear
        this.wrapper.querySelector('#cal-clear-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            this.selectedMonth = null;
            this.selectedYear = null;
            this.viewYear = this.currentYear;

            this.updateGrid();
            this.updatePillDisplay();
            this.popover.style.display = 'none';

            this.onFilter(null, null);
        });

        // This Month shortcut
        this.wrapper.querySelector('#cal-current-month-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            const now = new Date();
            this.selectedMonth = String(now.getMonth() + 1).padStart(2, '0');
            this.selectedYear = String(now.getFullYear());
            this.viewYear = now.getFullYear();

            this.updateGrid();
            this.updatePillDisplay();
            this.popover.style.display = 'none';

            this.onFilter(this.selectedMonth, this.selectedYear);
        });
    }

    getFilterValues() {
        return {
            month: this.selectedMonth,
            year: this.selectedYear
        };
    }
}

window.CalendarFilter = CalendarFilter;
