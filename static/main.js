let spreadsheetData = [];
const rows = 50;
const cols = 26;

async function init() {
    await loadData();
    renderGrid();
    setupEventListeners();
}

async function loadData() {
    try {
        const response = await fetch('/api/load');
        spreadsheetData = await response.json();
    } catch (err) {
        console.error('Failed to load data', err);
        showStatus('データの読み込みに失敗しました');
    }
}

async function saveData() {
    try {
        const response = await fetch('/api/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(spreadsheetData)
        });
        if (response.ok) {
            showStatus('保存しました');
        }
    } catch (err) {
        console.error('Failed to save data', err);
        showStatus('保存に失敗しました');
    }
}

function showStatus(msg) {
    const statusEl = document.getElementById('status-msg');
    statusEl.textContent = msg;
    setTimeout(() => { statusEl.textContent = ''; }, 3000);
}

function renderGrid() {
    const container = document.getElementById('spreadsheet-container');
    const table = document.createElement('table');
    table.className = 'spreadsheet-table';

    // Add colgroup for efficient resizing
    const colGroup = document.createElement('colgroup');
    const cornerCol = document.createElement('col');
    cornerCol.style.width = '40px';
    colGroup.appendChild(cornerCol);
    for (let c = 0; c < cols; c++) {
        const col = document.createElement('col');
        col.style.width = '120px'; // Default width
        colGroup.appendChild(col);
    }
    table.appendChild(colGroup);

    // Column Headers (A, B, C...)
    const headerRow = document.createElement('tr');
    const corner = document.createElement('th');
    corner.className = 'row-header col-header corner-header';
    headerRow.appendChild(corner);

    for (let c = 0; c < cols; c++) {
        const th = document.createElement('th');
        th.className = 'col-header';
        th.textContent = String.fromCharCode(65 + c);

        // Add resizer handle
        const resizer = document.createElement('div');
        resizer.className = 'resizer';
        th.appendChild(resizer);

        const colElement = colGroup.children[c + 1]; // +1 for corner col
        setupResizer(resizer, colElement);

        headerRow.appendChild(th);
    }
    table.appendChild(headerRow);

    // Rows
    for (let r = 0; r < rows; r++) {
        const tr = document.createElement('tr');
        const rh = document.createElement('th');
        rh.className = 'row-header';
        rh.textContent = r + 1;
        tr.appendChild(rh);

        for (let c = 0; c < cols; c++) {
            const td = document.createElement('td');
            td.className = 'cell';
            td.contentEditable = true;
            td.textContent = (spreadsheetData[r] && spreadsheetData[r][c]) || '';

            td.addEventListener('blur', (e) => {
                updateData(r, c, e.target.textContent);
            });

            // Prevent Enter from creating new divs in contentEditable
            td.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    td.blur();
                    // Move focus to the cell below
                    const nextRow = table.rows[r + 2]; // +1 for header, +1 for next
                    if (nextRow) {
                        nextRow.cells[c + 1].focus(); // +1 for row header
                    }
                }
            });

            tr.appendChild(td);
        }
        table.appendChild(tr);
    }

    container.innerHTML = '';
    container.appendChild(table);
}

function updateData(r, c, value) {
    if (!spreadsheetData[r]) {
        spreadsheetData[r] = [];
    }
    spreadsheetData[r][c] = value;
}

function setupEventListeners() {
    document.getElementById('save-btn').addEventListener('click', saveData);
    document.getElementById('load-btn').addEventListener('click', async () => {
        await loadData();
        renderGrid();
        showStatus('読み込みました');
    });

    document.addEventListener('paste', handlePaste);
}

async function handlePaste(e) {
    const activeCell = document.activeElement;
    if (!activeCell || !activeCell.classList.contains('cell')) return;

    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedText = clipboardData.getData('text');

    if (pastedText.includes('\t') || pastedText.includes('\n')) {
        e.preventDefault();
        const rows_pasted = pastedText.split(/\r?\n/);
        const startRow = Array.from(activeCell.parentNode.parentNode.children).indexOf(activeCell.parentNode) - 1;
        const startCol = Array.from(activeCell.parentNode.children).indexOf(activeCell) - 1;

        rows_pasted.forEach((rowData, rIdx) => {
            if (rowData.trim() === '' && rIdx === rows_pasted.length - 1) return;
            const cells_pasted = rowData.split('\t');
            cells_pasted.forEach((cellData, cIdx) => {
                const targetRow = startRow + rIdx;
                const targetCol = startCol + cIdx;
                if (targetRow < rows && targetCol < cols) {
                    updateData(targetRow, targetCol, cellData);
                }
            });
        });
        renderGrid();
        showStatus('貼り付けました');
    }
}

function setupResizer(resizer, col) {
    let x = 0;
    let w = 0;

    const mouseDownHandler = function (e) {
        x = e.clientX;
        const styles = window.getComputedStyle(col);
        w = parseInt(styles.width, 10);

        document.addEventListener('mousemove', mouseMoveHandler);
        document.addEventListener('mouseup', mouseUpHandler);

        resizer.classList.add('resizing');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault(); // Prevent text selection
    };

    const mouseMoveHandler = function (e) {
        const dx = e.clientX - x;
        col.style.width = `${Math.max(30, w + dx)}px`;
    };

    const mouseUpHandler = function () {
        document.removeEventListener('mousemove', mouseMoveHandler);
        document.removeEventListener('mouseup', mouseUpHandler);
        resizer.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    };

    resizer.addEventListener('mousedown', mouseDownHandler);
}

init();
