import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getFirestore, doc, onSnapshot, setDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";
// === Firebase Init ===
console.log("App Version: 3.1.0 - Diamond Canvas");
let app, db;
let unsubscribeSnapshot = null;
try {
    const firebaseConfig = {
      apiKey: "AIzaSyA15XcKQNF8vg72GoFflNOPqv4PthJF_EY",
      authDomain: "diamondcanvasfinance.firebaseapp.com",
      projectId: "diamondcanvasfinance",
      storageBucket: "diamondcanvasfinance.firebasestorage.app",
      messagingSenderId: "343229435947",
      appId: "1:343229435947:web:0ed9c0c5660ba5e3fd9a6e",
      measurementId: "G-3GNG09F2X0"
    };
    app = initializeApp(firebaseConfig);
    db = getFirestore(app);
} catch (e) {
    console.error("Firebase init fallback error:", e);
}
// === State & LocalStorage ===
const STORE_PREFIX = 'diamond_canvas';
const DEFAULT_CATEGORIES = [
    { id: 'cat_1', name: 'Аренда', items: [] },
    { id: 'cat_2', name: 'Зарплата', items: [] },
    { id: 'cat_3', name: 'Хоз. товары', items: [] },
    { id: 'cat_4', name: 'Продукты', items: [] },
    { id: 'cat_5', name: 'Налоги', items: [] },
];
let state = {
    currentDate: new Date(),
    categories: [],
    income: {}, // { 'YYYY-MM-DD': { soc, wb, ozon, yandex } }
    expenses: {} // { 'YYYY-MM-DD': [{ id, name, categoryId, amount }] }
};
function loadState() {
    if (unsubscribeSnapshot) unsubscribeSnapshot();
    
    unsubscribeSnapshot = onSnapshot(doc(db, 'dc_finance', 'main_state'), (docSnap) => {
        if (docSnap.exists()) {
            const data = docSnap.data();
            state.categories = data.categories || [...DEFAULT_CATEGORIES];
            state.income = data.income || {};
            state.expenses = data.expenses || {};
            
            // Skip destructive re-renders while user is actively editing
            const isEditing = isUserEditing();
            
            if (document.getElementById('view-data') && document.getElementById('view-data').classList.contains('active')) {
                // Always safe — updateValuesOnly skips focused elements
                updateValuesOnly();
            }
            // Settings and analytics use innerHTML, skip if user is in an input
            if (!isEditing) {
                if (document.getElementById('view-settings') && document.getElementById('view-settings').classList.contains('active')) renderSettings();
                if (document.getElementById('view-analytics') && document.getElementById('view-analytics').classList.contains('active')) renderAnalytics();
            }
        } else {
            state.categories = [...DEFAULT_CATEGORIES];
            state.income = {};
            state.expenses = {};
            saveState(); // push defaults to cloud
        }
    }, (error) => {
        console.error("Error syncing data:", error);
    });
}
function saveState(key) {
    setDoc(doc(db, 'dc_finance', 'main_state'), {
        categories: state.categories,
        income: state.income,
        expenses: state.expenses
    }, { merge: true })
    .catch((error) => console.error("Error saving data:", error));
}
// === Utils ===
/** Check if user is actively editing (input/select/textarea focused) */
function isUserEditing() {
    const el = document.activeElement;
    if (!el) return false;
    const tag = el.tagName;
    return tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
}
/** Debounced save — batches rapid keystrokes into a single Firebase write */
let _saveTimer = null;
function debouncedSave(delay = 600) {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(() => saveState(), delay);
}
function formatDateStr(date) {
    const d = new Date(date);
    let month = '' + (d.getMonth() + 1);
    let day = '' + d.getDate();
    let year = d.getFullYear();
    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;
    return [year, month, day].join('-');
}
function parseLocalDate(dateStr) {
    if (!dateStr) return new Date();
    const parts = dateStr.split('-');
    if (parts.length !== 3) return new Date(dateStr);
    return new Date(parts[0], parts[1] - 1, parts[2]);
}
function formatNumber(num) {
    if (!num) return '';
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}
function parseNumber(str) {
    if (!str) return 0;
    if (typeof str === 'number') return str;
    return parseInt(str.toString().replace(/\s/g, ''), 10) || 0;
}
function getDailyIncome(dateStr) {
    const inc = state.income[dateStr];
    if (!inc) return 0;
    if (typeof inc === 'object') {
        return (parseNumber(inc.soc) || 0) + (parseNumber(inc.wb) || 0) + (parseNumber(inc.ozon) || 0) + (parseNumber(inc.yandex) || 0);
    }
    return parseNumber(inc) || 0;
}
// === Navigation ===
function initNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    const views = document.querySelectorAll('.view');
    navItems.forEach(item => {
        item.addEventListener('click', () => {
            const targetId = item.getAttribute('data-target');
            
            navItems.forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            views.forEach(v => {
                if (v.id === targetId) {
                    v.classList.add('active');
                } else {
                    v.classList.remove('active');
                }
            });
            // Refresh specific views on enter
            if (targetId === 'view-data') {
                updateValuesOnly();
            } else if (targetId === 'view-analytics') {
                renderAnalytics();
            } else if (targetId === 'view-settings') {
                renderSettings();
            }
        });
    });
}
// === Calendar ===
function initCalendar() {
    const btnPrev = document.getElementById('cal-prev');
    const btnNext = document.getElementById('cal-next');
    
    btnPrev.addEventListener('click', () => {
        state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() - 1, 1);
        renderCalendar();
        renderDataEntry();
        updateValuesOnly();
    });
    
    btnNext.addEventListener('click', () => {
        state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() + 1, 1);
        renderCalendar();
        renderDataEntry();
        updateValuesOnly();
    });
    renderCalendar();
}
function renderCalendar() {
    const monthYearStr = state.currentDate.toLocaleString('ru-RU', { month: 'long', year: 'numeric' });
    document.getElementById('cal-month-year').textContent = monthYearStr.charAt(0).toUpperCase() + monthYearStr.slice(1);
    const grid = document.getElementById('cal-days-grid');
    grid.innerHTML = '';
    const year = state.currentDate.getFullYear();
    const month = state.currentDate.getMonth();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const weekdays = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
    let activeElement = null;
    for (let i = 1; i <= daysInMonth; i++) {
        const d = new Date(year, month, i);
        const dateStr = formatDateStr(d);
        
        const el = document.createElement('div');
        el.className = 'cal-day';
        
        if (dateStr === formatDateStr(state.currentDate)) {
            el.classList.add('active');
            activeElement = el;
        }
        
        // Has data marker
        if (getDailyIncome(dateStr) > 0 || (state.expenses[dateStr] && state.expenses[dateStr].some(e => e.amount > 0))) {
            el.classList.add('has-data');
        }
        const weekdayIndex = d.getDay();
        
        el.innerHTML = `
            <span class="weekday">${weekdays[weekdayIndex]}</span>
            <span class="day-num">${d.getDate()}</span>
        `;
        
        el.addEventListener('click', () => {
            state.currentDate = new Date(d);
            renderCalendar();
            renderDataEntry();
            updateValuesOnly();
        });
        
        grid.appendChild(el);
    }
    
    if (activeElement) {
        requestAnimationFrame(() => {
            activeElement.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
        });
    }
}
// === Data Entry ===
function setupInputFormatting(inputEl, callback) {
    inputEl.addEventListener('input', (e) => {
        let val = e.target.value.replace(/[^\d]/g, ''); // Remove non-digits
        
        // Format with spaces
        if (val !== '') {
            e.target.value = formatNumber(val);
        } else {
            e.target.value = '';
        }
        
        if (callback) callback(parseNumber(e.target.value));
    });
}
function renderDataEntry() {
    const sources = ['soc', 'wb', 'ozon', 'yandex'];
    sources.forEach(src => {
        const inputEl = document.getElementById(`inc-${src}`);
        if (inputEl && !inputEl.dataset.listenerAttached) {
            inputEl.dataset.listenerAttached = 'true';
            inputEl.addEventListener('focus', function() { this.select(); });
            setupInputFormatting(inputEl, (val) => {
                const dateStr = formatDateStr(state.currentDate);
                if (typeof state.income[dateStr] !== 'object' || state.income[dateStr] === null) {
                    const oldVal = state.income[dateStr];
                    state.income[dateStr] = { soc: oldVal || 0, wb: 0, ozon: 0, yandex: 0 };
                }
                state.income[dateStr][src] = val;
                // Debounce to avoid flooding Firebase on every keystroke
                debouncedSave();
                updateValuesOnly();
                renderCalendar(); 
            });
        }
    });
    const rowsContainer = document.getElementById('expenses-rows');
    if (!rowsContainer) return;
    // Skip full rebuild if rows already exist — prevents focus/scroll loss
    if (rowsContainer.children.length > 0) return;
    rowsContainer.innerHTML = '';
    
    for (let idx = 0; idx < 15; idx++) {
        const row = document.createElement('div');
        row.className = 'table-row';
        row.setAttribute('data-idx', idx);
        
        row.innerHTML = `
            <div class="row-num">${idx + 1}</div>
            <div>
                <select class="table-select exp-cat"></select>
            </div>
            <div>
                <select class="table-select exp-name"></select>
            </div>
            <div class="amount-wrapper">
                <input type="text" class="table-input amount-input exp-amount" placeholder="0" inputmode="numeric">
            </div>
            <div>
                <button class="icon-btn danger btn-clear-row" title="Очистить строку">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6 6 18M6 6l12 12"/></svg>
                </button>
            </div>
        `;
        
        const nameSelect = row.querySelector('.exp-name');
        const catSelect = row.querySelector('.exp-cat');
        const amountInput = row.querySelector('.exp-amount');
        const clearBtn = row.querySelector('.btn-clear-row');
        amountInput.addEventListener('focus', function() { this.select(); });
        catSelect.addEventListener('change', (e) => {
            const dateStr = formatDateStr(state.currentDate);
            updateExpense(dateStr, idx, 'categoryId', e.target.value);
            updateExpense(dateStr, idx, 'name', '');
            updateValuesOnly(); 
        });
        nameSelect.addEventListener('change', (e) => {
            const dateStr = formatDateStr(state.currentDate);
            if (e.target.value === '__add__') {
                const exps = state.expenses[dateStr] || [];
                e.target.value = (exps[idx] && exps[idx].name) || '';
                // Read categoryId from state, not DOM — Safari/iOS can return empty catSelect.value after innerHTML update
                const catId = (exps[idx] && exps[idx].categoryId) || catSelect.value;
                openModal(catId, idx, dateStr);
            } else {
                updateExpense(dateStr, idx, 'name', e.target.value);
            }
        });
        setupInputFormatting(amountInput, (val) => {
            const dateStr = formatDateStr(state.currentDate);
            updateExpenseLocal(dateStr, idx, 'amount', val);
        });
        clearBtn.addEventListener('click', () => {
            const dateStr = formatDateStr(state.currentDate);
            if (!state.expenses[dateStr]) return;
            state.expenses[dateStr][idx] = { name: '', categoryId: '', amount: 0 };
            saveState('expenses');
            updateValuesOnly();
        });
        rowsContainer.appendChild(row);
    }
}
function updateValuesOnly() {
    const dateStr = formatDateStr(state.currentDate);
    
    // Income
    const incData = state.income[dateStr];
    const isObj = typeof incData === 'object' && incData !== null;
    
    ['soc', 'wb', 'ozon', 'yandex'].forEach(src => {
        const inputEl = document.getElementById(`inc-${src}`);
        let val = '';
        if (incData) {
             val = isObj ? (incData[src] || '') : (src === 'soc' ? incData : '');
        }
        if (inputEl && document.activeElement !== inputEl) {
            inputEl.value = formatNumber(val);
        }
    });
    const totalEl = document.getElementById('daily-income-total');
    if (totalEl) {
        totalEl.textContent = `${formatNumber(getDailyIncome(dateStr))} ₽`;
    }
    
    // Expenses
    let dayExpenses = state.expenses[dateStr] || [];
    const rowsContainer = document.getElementById('expenses-rows');
    if (!rowsContainer) return;
    const rows = rowsContainer.children;
    
    for (let idx = 0; idx < 15; idx++) {
        if (idx >= rows.length) break;
        const exp = dayExpenses[idx] || { name: '', categoryId: '', amount: 0 };
        const row = rows[idx];
        const catSelect = row.querySelector('.exp-cat');
        const nameSelect = row.querySelector('.exp-name');
        const amountInput = row.querySelector('.exp-amount');
        const activeEl = document.activeElement;
        if (catSelect && activeEl !== catSelect) {
            let catOptions = `<option value="">Выбрать...</option>`;
            state.categories.forEach(cat => {
                catOptions += `<option value="${cat.id}" ${cat.id === exp.categoryId ? 'selected' : ''}>${cat.name}</option>`;
            });
            catSelect.innerHTML = catOptions;
        }
        if (nameSelect && activeEl !== nameSelect) {
            nameSelect.disabled = !exp.categoryId;
            let nameOptions = `<option value="">Выбрать...</option>`;
            let hasItems = false;
            
            if (exp.categoryId) {
                const cat = state.categories.find(c => c.id === exp.categoryId);
                if (cat && cat.items && cat.items.length > 0) {
                    hasItems = true;
                    cat.items.forEach(item => {
                        nameOptions += `<option value="${item}" ${item.toLowerCase() === (exp.name || '').toLowerCase() ? 'selected' : ''}>${item}</option>`;
                    });
                }
                
                if (!hasItems) {
                    nameOptions = `<option value="__add__" style="font-weight: 600; color: var(--accent-income);">+ Добавить наименование</option>`;
                } else {
                    nameOptions += `<option value="__add__" style="font-weight: 600; color: var(--accent-income);">+ Добавить наименование</option>`;
                }
            }
            nameSelect.innerHTML = nameOptions;
        }
        if (amountInput && activeEl !== amountInput) {
            amountInput.value = formatNumber(exp.amount) || '';
        }
    }
    updateExpenseTotal(dateStr);
}
function updateExpense(dateStr, index, field, value) {
    if (!state.expenses[dateStr]) {
        state.expenses[dateStr] = Array(15).fill().map(() => ({ name: '', categoryId: '', amount: 0 }));
    }
    state.expenses[dateStr][index][field] = value;
    saveState('expenses');
    
    if (field === 'amount') {
        updateExpenseTotal(dateStr);
        renderCalendar(); // Update dots
    }
}
/** Debounced version for amount inputs — prevents Firebase flood during typing */
function updateExpenseLocal(dateStr, index, field, value) {
    if (!state.expenses[dateStr]) {
        state.expenses[dateStr] = Array(15).fill().map(() => ({ name: '', categoryId: '', amount: 0 }));
    }
    state.expenses[dateStr][index][field] = value;
    debouncedSave();
    
    if (field === 'amount') {
        updateExpenseTotal(dateStr);
    }
}
function updateExpenseTotal(dateStr) {
    const exps = state.expenses[dateStr] || [];
    const total = exps.reduce((acc, curr) => acc + (parseInt(curr.amount) || 0), 0);
    document.getElementById('expense-total').textContent = `${formatNumber(total)} ₽`;
}
// === Modal (Names) ===
let currentModalContext = { categoryId: null, rowIndex: null, dateStr: null };
function openModal(categoryId, rowIndex = null, dateStr = null) {
    currentModalContext = { categoryId, rowIndex, dateStr };
    const modal = document.getElementById('add-name-modal');
    const input = document.getElementById('new-item-name');
    input.value = '';
    modal.classList.remove('hidden');
    setTimeout(() => input.focus(), 100);
}
function initModal() {
    const modal = document.getElementById('add-name-modal');
    const cancelBtn = document.getElementById('modal-cancel-btn');
    const saveBtn = document.getElementById('modal-save-btn');
    const input = document.getElementById('new-item-name');
    const closeModal = () => modal.classList.add('hidden');
    cancelBtn.onclick = closeModal;
    saveBtn.onclick = () => {
        const val = input.value.trim();
        if (!val || !currentModalContext.categoryId) {
            closeModal();
            return;
        }
        
        const cat = state.categories.find(c => c.id === currentModalContext.categoryId);
        if (cat) {
            const valUpper = val.charAt(0).toUpperCase() + val.slice(1);
            if (!cat.items) cat.items = [];
            
            const exists = cat.items.find(i => i.toLowerCase() === valUpper.toLowerCase());
            if (!exists) {
                cat.items.push(valUpper);
                saveState('categories');
            }
            
            if (currentModalContext.rowIndex !== null && currentModalContext.dateStr) {
                updateExpense(currentModalContext.dateStr, currentModalContext.rowIndex, 'name', valUpper);
                updateValuesOnly();
            } else {
                renderSettings();
            }
        }
        closeModal();
    };
}
// === Analytics ===
function renderAnalytics() {
    const dateFromEl = document.getElementById('date-from');
    const dateToEl = document.getElementById('date-to');
    
    // Set default range to current month if empty
    if (!dateFromEl.value) {
        const d = new Date(state.currentDate);
        d.setDate(1);
        dateFromEl.value = formatDateStr(d);
    }
    if (!dateToEl.value) {
        const d = new Date(state.currentDate);
        d.setMonth(d.getMonth() + 1);
        d.setDate(0);
        dateToEl.value = formatDateStr(d);
    }
    const calcAndRender = () => {
        const from = parseLocalDate(dateFromEl.value);
        const to = parseLocalDate(dateToEl.value);
        if (from > to) return;
        let totalInc = 0;
        let totalExp = 0;
        const catTotals = {};
        const catDetails = {};
        const dailyIncomes = [];
        let sourcesTotal = { soc: 0, wb: 0, ozon: 0, yandex: 0 };
        // Loop through all days in range
        let curr = new Date(from);
        while (curr <= to) {
            const dStr = formatDateStr(curr);
            const inc = getDailyIncome(dStr);
            totalInc += inc;
            dailyIncomes.push({ date: dStr, val: inc });
            // Accumulate per-source totals
            const incObj = state.income[dStr];
            if (incObj && typeof incObj === 'object') {
                sourcesTotal.soc += parseNumber(incObj.soc) || 0;
                sourcesTotal.wb += parseNumber(incObj.wb) || 0;
                sourcesTotal.ozon += parseNumber(incObj.ozon) || 0;
                sourcesTotal.yandex += parseNumber(incObj.yandex) || 0;
            } else if (incObj) {
                // Legacy: single number goes to soc
                sourcesTotal.soc += parseNumber(incObj) || 0;
            }
            const exps = state.expenses[dStr] || [];
            exps.forEach(e => {
                const am = parseInt(e.amount) || 0;
                totalExp += am;
                if (am > 0 && e.categoryId) {
                    catTotals[e.categoryId] = (catTotals[e.categoryId] || 0) + am;
                    
                    if (!catDetails[e.categoryId]) catDetails[e.categoryId] = {};
                    const n = (e.name || 'Без названия').trim();
                    const nLower = n.toLowerCase();
                    if (!catDetails[e.categoryId][nLower]) {
                        catDetails[e.categoryId][nLower] = { name: n.charAt(0).toUpperCase() + n.slice(1), amount: 0 };
                    }
                    catDetails[e.categoryId][nLower].amount += am;
                }
            });
            curr.setDate(curr.getDate() + 1);
        }
        // Summary Cards
        document.getElementById('summary-income').textContent = `${formatNumber(totalInc)} ₽`;
        document.getElementById('summary-expense').textContent = `${formatNumber(totalExp)} ₽`;
        // Update income details modal values
        document.getElementById('detail-soc').textContent = `${formatNumber(sourcesTotal.soc)} ₽`;
        document.getElementById('detail-wb').textContent = `${formatNumber(sourcesTotal.wb)} ₽`;
        document.getElementById('detail-ozon').textContent = `${formatNumber(sourcesTotal.ozon)} ₽`;
        document.getElementById('detail-yandex').textContent = `${formatNumber(sourcesTotal.yandex)} ₽`;
        document.getElementById('detail-total').textContent = `${formatNumber(totalInc)} ₽`;
        // Render Chart
        renderChart(dailyIncomes);
        // Sort categories by amount desc (used by expense modal)
        const sortedCats = Object.entries(catTotals).sort((a,b) => b[1] - a[1]);
        // Update expense details modal list
        const expModalList = document.getElementById('expense-modal-list');
        if (sortedCats.length > 0) {
            let expModalHtml = sortedCats.map(([catId, amount]) => {
                const cat = state.categories.find(c => c.id === catId);
                const catName = cat ? cat.name : 'Удаленная категория';
                const perc = totalExp > 0 ? ` (${Math.round(amount / totalExp * 100)}%)` : '';
                return `<div class="income-detail-row">
                    <span class="source-name">${catName}</span>
                    <span class="source-value">${formatNumber(amount)} ₽<span class="muted" style="font-weight:400; font-size:12px; margin-left:4px;">${perc}</span></span>
                </div>`;
            }).join('');
            expModalHtml += `<div class="income-detail-row total">
                <span class="source-name">Итого</span>
                <span class="source-value expense-total-value">${formatNumber(totalExp)} ₽</span>
            </div>`;
            expModalList.innerHTML = expModalHtml;
        } else {
            expModalList.innerHTML = '<p class="muted" style="padding: 16px 0;">Нет расходов за выбранный период</p>';
        }
    };
    dateFromEl.addEventListener('change', calcAndRender);
    dateToEl.addEventListener('change', calcAndRender);
    
    document.getElementById('btn-export-excel').onclick = exportToExcel;
    // Income details modal — use onclick to prevent duplicate listeners on tab switch
    const incomeModal = document.getElementById('income-details-modal');
    document.getElementById('income-card-clickable').onclick = () => {
        incomeModal.classList.remove('hidden');
    };
    document.getElementById('income-modal-close').onclick = () => {
        incomeModal.classList.add('hidden');
    };
    // Expense details modal
    const expenseModal = document.getElementById('expense-details-modal');
    document.getElementById('expense-card-clickable').onclick = () => {
        expenseModal.classList.remove('hidden');
    };
    document.getElementById('expense-modal-close').onclick = () => {
        expenseModal.classList.add('hidden');
    };
    calcAndRender();
}
async function exportToExcel() {
    if (typeof ExcelJS === 'undefined') {
        alert('Библиотека ExcelJS не загружена');
        return;
    }
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Даймонд Канвас';
    workbook.created = new Date();
    const monthsSet = new Set();
    Object.keys(state.income).forEach(d => monthsSet.add(d.substring(0, 7)));
    Object.keys(state.expenses).forEach(d => {
        if (state.expenses[d].some(e => e.amount > 0)) {
            monthsSet.add(d.substring(0, 7));
        }
    });
    const monthsArr = Array.from(monthsSet).sort();
    
    if (monthsArr.length === 0) {
        alert('Нет данных для выгрузки');
        return;
    }
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    monthsArr.forEach(monthStr => {
        const [year, month] = monthStr.split('-');
        const sheetName = `${monthNames[parseInt(month) - 1]} ${year}`;
        const ws = workbook.addWorksheet(sheetName);
        
        ws.properties.outlineProperties = { summaryBelow: false };
        
        // Define Columns widths
        // Table 1 (Доходы)
        ws.getColumn('A').width = 15; // Дата
        ws.getColumn('B').width = 15; // СоцСети
        ws.getColumn('C').width = 15; // WB
        ws.getColumn('D').width = 15; // Ozon
        ws.getColumn('E').width = 15; // Яндекс
        ws.getColumn('F').width = 20; // ИТОГО ДОХОДЫ
        
        // Spacer
        ws.getColumn('G').width = 5;  
        
        // Table 2 (Расходы по категориям)
        ws.getColumn('H').width = 25; // Категория
        ws.getColumn('I').width = 28; // Наименование
        ws.getColumn('J').width = 20; // Сумма
        
        let totalInc = 0;
        let totalExp = 0;
        
        const monthIncomes = [];
        const monthExpensesByCategory = {};
        const monthTransactions = [];
        const daysInMonth = new Date(year, month, 0).getDate();
        for (let i = 1; i <= daysInMonth; i++) {
            const dStr = `${year}-${month}-${i.toString().padStart(2, '0')}`;
            
            // Incomes
            const incData = state.income[dStr];
            let soc=0, wb=0, ozon=0, yandex=0, dTotal=0;
            if (incData) {
                if (typeof incData === 'object') {
                    soc = parseNumber(incData.soc) || 0;
                    wb = parseNumber(incData.wb) || 0;
                    ozon = parseNumber(incData.ozon) || 0;
                    yandex = parseNumber(incData.yandex) || 0;
                } else {
                    soc = parseNumber(incData) || 0;
                }
                dTotal = soc + wb + ozon + yandex;
            }
            if (dTotal > 0) {
                totalInc += dTotal;
                monthIncomes.push({ date: dStr, soc, wb, ozon, yandex, total: dTotal });
            }
            // Expenses
            const exps = state.expenses[dStr] || [];
            exps.forEach(e => {
                const am = parseInt(e.amount) || 0;
                if (am > 0 && e.categoryId) {
                    totalExp += am;
                    if (!monthExpensesByCategory[e.categoryId]) {
                        monthExpensesByCategory[e.categoryId] = { total: 0, items: [] };
                    }
                    monthExpensesByCategory[e.categoryId].total += am;
                    
                    const name = (e.name || 'Без названия').trim();
                    const nameCap = name.charAt(0).toUpperCase() + name.slice(1);
                    
                    const existing = monthExpensesByCategory[e.categoryId].items.find(item => item.name.toLowerCase() === nameCap.toLowerCase());
                    if (existing) {
                        existing.amount += am;
                    } else {
                        monthExpensesByCategory[e.categoryId].items.push({ name: nameCap, amount: am });
                    }
                    
                    const cat = state.categories.find(c => c.id === e.categoryId);
                    const catName = cat ? cat.name : 'Удаленная категория';
                    monthTransactions.push({ date: dStr, name: nameCap, category: catName, amount: am });
                }
            });
        }
        // Headers row 1
        ws.mergeCells('A1:F1');
        ws.getCell('A1').value = `ИТОГО ДОХОДЫ: ${totalInc} ₽`;
        ws.getCell('A1').font = { bold: true, color: { argb: 'FF059669' }, size: 14 };
        ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
        
        ws.mergeCells('H1:J1');
        ws.getCell('H1').value = `ИТОГО РАСХОДЫ: ${totalExp} ₽`;
        ws.getCell('H1').font = { bold: true, color: { argb: 'FFE11D48' }, size: 14 };
        ws.getCell('H1').alignment = { vertical: 'middle', horizontal: 'center' };
        // Headers row 2
        ws.getCell('A2').value = 'Дата';
        ws.getCell('B2').value = 'СоцСети';
        ws.getCell('C2').value = 'WB';
        ws.getCell('D2').value = 'Ozon';
        ws.getCell('E2').value = 'Яндекс';
        ws.getCell('F2').value = 'ИТОГО';
        
        ws.getCell('H2').value = 'Категория';
        ws.getCell('I2').value = 'Наименование';
        ws.getCell('J2').value = 'Сумма';
        ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'H2', 'I2', 'J2'].forEach(c => {
            ws.getCell(c).font = { bold: true };
            ws.getCell(c).border = { bottom: { style: 'medium' } };
        });
        // Fill Income (Left)
        let incRow = 3;
        monthIncomes.forEach(inc => {
            ws.getCell(`A${incRow}`).value = inc.date;
            ws.getCell(`B${incRow}`).value = inc.soc;
            ws.getCell(`B${incRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`C${incRow}`).value = inc.wb;
            ws.getCell(`C${incRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`D${incRow}`).value = inc.ozon;
            ws.getCell(`D${incRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`E${incRow}`).value = inc.yandex;
            ws.getCell(`E${incRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`F${incRow}`).value = inc.total;
            ws.getCell(`F${incRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`F${incRow}`).font = { bold: true };
            incRow++;
        });
        // Fill Expenses (Right)
        let expRow = 3;
        const sortedCats = Object.entries(monthExpensesByCategory).sort((a,b) => b[1].total - a[1].total);
        
        sortedCats.forEach(([catId, data]) => {
            const cat = state.categories.find(c => c.id === catId);
            const catName = cat ? cat.name : 'Удаленная категория';
            
            ws.getCell(`H${expRow}`).value = catName;
            ws.getCell(`J${expRow}`).value = data.total;
            
            ws.getCell(`H${expRow}`).font = { bold: true };
            ws.getCell(`J${expRow}`).font = { bold: true };
            ws.getCell(`J${expRow}`).numFmt = '#,##0 \\₽';
            ws.getCell(`H${expRow}`).border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
            ws.getCell(`I${expRow}`).border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
            ws.getCell(`J${expRow}`).border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
            
            ws.getRow(expRow).outlineLevel = 0;
            expRow++;
            
            data.items.sort((a,b) => b.amount - a.amount).forEach(item => {
                ws.getCell(`I${expRow}`).value = item.name;
                ws.getCell(`J${expRow}`).value = item.amount;
                ws.getCell(`J${expRow}`).numFmt = '#,##0 \\₽';
                ws.getCell(`I${expRow}`).font = { color: { argb: 'FF64748B' } }; 
                
                ws.getRow(expRow).outlineLevel = 1;
                expRow++;
            });
        });
        // Table 3 (Bottom)
        let t3Row = Math.max(incRow, expRow) + 3; // Leave a few blank lines
        
        ws.mergeCells(`A${t3Row}:D${t3Row}`);
        ws.getCell(`A${t3Row}`).value = 'ДЕТАЛЬНЫЙ ЖУРНАЛ ТРАНЗАКЦИЙ';
        ws.getCell(`A${t3Row}`).font = { bold: true, size: 12 };
        t3Row++;
        ws.getCell(`A${t3Row}`).value = 'Дата';
        ws.getCell(`B${t3Row}`).value = 'Наименование';
        ws.getCell(`C${t3Row}`).value = 'Категория';
        ws.getCell(`D${t3Row}`).value = 'Сумма';
        ['A', 'B', 'C', 'D'].forEach(c => {
            ws.getCell(`${c}${t3Row}`).font = { bold: true };
            ws.getCell(`${c}${t3Row}`).border = { bottom: { style: 'medium' } };
        });
        t3Row++;
        monthTransactions.forEach(t => {
            ws.getCell(`A${t3Row}`).value = t.date;
            ws.getCell(`B${t3Row}`).value = t.name;
            ws.getCell(`C${t3Row}`).value = t.category;
            ws.getCell(`D${t3Row}`).value = t.amount;
            ws.getCell(`D${t3Row}`).numFmt = '#,##0 \\₽';
            t3Row++;
        });
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const now = new Date();
    const dateString = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}`;
    a.download = `diamond_canvas_report_${dateString}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
}
function renderChart(data) {
    const container = document.getElementById('bar-chart');
    container.innerHTML = '';
    
    if (data.length === 0) return;
    
    // Fixed bar width for horizontal scroll — 44px per day
    const BAR_WIDTH = 44;
    const chartWidth = data.length * BAR_WIDTH;
    container.style.minWidth = Math.max(chartWidth, 600) + 'px';
    
    const maxVal = Math.max(...data.map(d => d.val), 1000); // min max for scale
    data.forEach(item => {
        const heightPct = (item.val / maxVal) * 100;
        const dObj = parseLocalDate(item.date);
        const label = `${dObj.getDate()}.${dObj.getMonth()+1}`;
        
        const wrapper = document.createElement('div');
        wrapper.className = 'chart-bar-wrapper';
        wrapper.style.width = BAR_WIDTH + 'px';
        wrapper.style.flex = 'none';
        wrapper.innerHTML = `
            <div class="chart-tooltip">${formatNumber(item.val)} ₽</div>
            <div class="chart-bar" style="height: ${Math.max(heightPct, 1)}%"></div>
            <div class="chart-label">${label}</div>
        `;
        container.appendChild(wrapper);
    });
}
// === Settings ===
function renderSettings() {
    const list = document.getElementById('categories-list');
    list.innerHTML = '';
    state.categories.forEach((cat, idx) => {
        const li = document.createElement('li');
        li.className = 'settings-item';
        
        let itemsHtml = '<div class="settings-items-list">';
        if (cat.items && cat.items.length > 0) {
            cat.items.forEach((item, itemIdx) => {
                itemsHtml += `
                    <div class="settings-subitem">
                        <span>${item}</span>
                        <button class="icon-btn danger delete-subitem" data-item="${itemIdx}">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6 6 18M6 6l12 12"/></svg>
                        </button>
                    </div>
                `;
            });
        }
        itemsHtml += `
            <button class="btn-add-item" data-catid="${cat.id}">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
                Добавить наименование
            </button>
        </div>`;
        li.innerHTML = `
            <div style="display:flex; flex-direction:column; width:100%;">
                <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border); padding-bottom:8px;">
                    <input type="text" value="${cat.name}" data-idx="${idx}" style="font-weight:600; font-size:18px; outline:none; border:none; background:transparent;">
                    <button class="icon-btn danger delete-cat">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                    </button>
                </div>
                ${itemsHtml}
            </div>
        `;
        
        const input = li.querySelector('input');
        const delBtn = li.querySelector('.delete-cat');
        input.addEventListener('blur', (e) => {
            const newName = e.target.value.trim();
            if (newName) {
                state.categories[idx].name = newName;
                saveState('categories');
            } else {
                e.target.value = state.categories[idx].name; // revert
            }
        });
        delBtn.addEventListener('click', () => {
            if (confirm(`Удалить категорию "${cat.name}" и все её наименования?`)) {
                state.categories.splice(idx, 1);
                saveState('categories');
                renderSettings();
            }
        });
        const deleteSubItemBtns = li.querySelectorAll('.delete-subitem');
        deleteSubItemBtns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const iIdx = parseInt(btn.getAttribute('data-item'));
                const itemName = state.categories[idx].items[iIdx];
                if (confirm(`Удалить наименование "${itemName}"?`)) {
                    state.categories[idx].items.splice(iIdx, 1);
                    saveState('categories');
                    renderSettings();
                }
            });
        });
        const btnAdd = li.querySelector('.btn-add-item');
        btnAdd.addEventListener('click', () => {
            openModal(cat.id);
        });
        list.appendChild(li);
    });
    const addBtn = document.getElementById('add-category-btn');
    // Replace element to clear listeners
    const newAddBtn = addBtn.cloneNode(true);
    addBtn.parentNode.replaceChild(newAddBtn, addBtn);
    
    newAddBtn.addEventListener('click', () => {
        state.categories.push({
            id: 'cat_' + Date.now(),
            name: 'Новая категория'
        });
        saveState('categories');
        renderSettings();
    });
    // Backup Export
    document.getElementById('btn-export-backup').onclick = () => {
        const dataStr = JSON.stringify(state, null, 2);
        const blob = new Blob([dataStr], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        const d = new Date();
        const dateString = `${d.getFullYear()}${(d.getMonth()+1).toString().padStart(2,'0')}${d.getDate().toString().padStart(2,'0')}`;
        a.download = `diamond_canvas_backup_${dateString}.json`;
        a.click();
        URL.revokeObjectURL(url);
    };
    // Backup Import
    const importBtn = document.getElementById('btn-import-backup');
    const importFile = document.getElementById('input-import-backup');
    
    importBtn.onclick = () => importFile.click();
    
    importFile.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        if (!confirm('Вы уверены? Загрузка резервной копии ПЕРЕЗАПИШЕТ все текущие данные. Это действие нельзя отменить.')) {
            importFile.value = ''; // reset
            return;
        }
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const importedState = JSON.parse(event.target.result);
                if (importedState && typeof importedState === 'object') {
                    if (importedState.categories) state.categories = importedState.categories;
                    if (importedState.income) state.income = importedState.income;
                    if (importedState.expenses) state.expenses = importedState.expenses;
                    
                    saveState('categories');
                    saveState('income');
                    saveState('expenses');
                    
                    alert('Данные успешно восстановлены!');
                    window.location.reload();
                } else {
                    alert('Неверный формат файла резервной копии.');
                }
            } catch (err) {
                alert('Ошибка чтения файла. Возможно, он поврежден.');
            }
        };
        reader.readAsText(file);
    };
}
// === Init ===
document.addEventListener('DOMContentLoaded', () => {
    try {
        initNavigation();
        initCalendar();
        initModal();
        renderDataEntry(); // 1. Создаем каркас таблицы и инпутов
        loadState();       // 2. Подключаем Firebase и заливаем данные
    } catch (e) {
        console.error("Critical Init Error:", e);
    }
});
