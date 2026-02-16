let spreadsheetData = [];
let rowCount = 100;
const cols = 26;
let selectedRow = null;
let selectedCol = null;
let selectionStart = null; // {r, c}
let selectionEnd = null;   // {r, c}
let isDragging = false;
let cellElements = []; // 2D array to store TD elements
let rowHeaderElements = [];
let colHeaderElements = [];
let undoStack = [];
const maxUndoSteps = 50;

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
    colHeaderElements = []; // Reset
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
        if (selectedCol === c) th.classList.add('active');
        th.textContent = String.fromCharCode(65 + c);
        colHeaderElements[c] = th;

        th.addEventListener('click', () => {
            selectColumn(c);
        });

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
    const displayRows = Math.max(rowCount, spreadsheetData.length + 10);
    for (let r = 0; r < displayRows; r++) {
        const tr = document.createElement('tr');
        const rh = document.createElement('th');
        rh.className = 'row-header';
        if (selectedRow === r) rh.classList.add('active');
        rh.textContent = r + 1;
        rowHeaderElements[r] = rh;

        rh.addEventListener('click', () => {
            selectRow(r);
        });

        tr.appendChild(rh);

        cellElements[r] = [];
        for (let c = 0; c < cols; c++) {
            const td = document.createElement('td');
            td.className = 'cell';
            if (isInSelection(r, c)) {
                td.classList.add('selected');
            }
            cellElements[r][c] = td;
            td.contentEditable = true;
            td.textContent = (spreadsheetData[r] && spreadsheetData[r][c]) || '';

            td.addEventListener('mousedown', (e) => {
                if (e.button !== 0) return; // Only left click
                isDragging = true;
                selectionStart = { r, c };
                selectionEnd = { r, c };
                clearSelection(false);
                document.body.classList.add('dragging');
                updateSelectionUI();
            });

            td.addEventListener('mouseenter', () => {
                if (isDragging) {
                    selectionEnd = { r, c };
                    updateSelectionUI();
                }
            });

            td.addEventListener('blur', (e) => {
                const newValue = e.target.textContent;
                const oldValue = (spreadsheetData[r] && spreadsheetData[r][c]) || '';
                if (newValue !== oldValue) {
                    pushToUndoStack();
                    updateData(r, c, newValue);
                }
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
        undoStack = []; // Clear undo stack on full reload
        renderGrid();
        showStatus('読み込みました');
    });
    document.getElementById('clear-btn').addEventListener('click', () => {
        if (confirm('全てのデータをクリアしますか？')) {
            pushToUndoStack();
            spreadsheetData = [];
            clearSelection();
            renderGrid();
            showStatus('クリアしました（保存はされていません）');
        }
    });
    document.getElementById('undo-btn').addEventListener('click', undo);

    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'z') {
            e.preventDefault();
            undo();
        }
    });

    document.getElementById('delete-row-btn').addEventListener('click', deleteSelectedRow);
    document.getElementById('delete-col-btn').addEventListener('click', deleteSelectedCol);

    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            document.body.classList.remove('dragging');
        }
    });

    document.addEventListener('copy', handleCopy);
    document.addEventListener('paste', handlePaste);
}

function isInSelection(r, c) {
    if (selectedRow === r || selectedCol === c) return true;
    if (selectionStart && selectionEnd) {
        const rMin = Math.min(selectionStart.r, selectionEnd.r);
        const rMax = Math.max(selectionStart.r, selectionEnd.r);
        const cMin = Math.min(selectionStart.c, selectionEnd.c);
        const cMax = Math.max(selectionStart.c, selectionEnd.c);
        return r >= rMin && r <= rMax && c >= cMin && c <= cMax;
    }
    return false;
}

function selectRow(r) {
    clearSelection(true);
    selectedRow = r;
    updateSelectionUI();
    updateDeleteButtons();
}

function selectColumn(c) {
    clearSelection(true);
    selectedCol = c;
    updateSelectionUI();
    updateDeleteButtons();
}

function updateSelectionUI() {
    let rMin = Infinity, rMax = -Infinity, cMin = Infinity, cMax = -Infinity;
    let hasRangeSelection = false;

    if (selectionStart && selectionEnd) {
        rMin = Math.min(selectionStart.r, selectionEnd.r);
        rMax = Math.max(selectionStart.r, selectionEnd.r);
        cMin = Math.min(selectionStart.c, selectionEnd.c);
        cMax = Math.max(selectionStart.c, selectionEnd.c);
        hasRangeSelection = true;
    }

    // Update cell classes
    cellElements.forEach((row, r) => {
        row.forEach((td, c) => {
            td.classList.remove('selected', 'selected-border-top', 'selected-border-bottom', 'selected-border-left', 'selected-border-right');

            if (isInSelection(r, c)) {
                td.classList.add('selected');

                if (hasRangeSelection) {
                    if (r === rMin) td.classList.add('selected-border-top');
                    if (r === rMax) td.classList.add('selected-border-bottom');
                    if (c === cMin) td.classList.add('selected-border-left');
                    if (c === cMax) td.classList.add('selected-border-right');
                } else if (selectedRow === r) {
                    // Row selection borders
                    if (r === selectedRow) {
                        td.classList.add('selected-border-top');
                        td.classList.add('selected-border-bottom');
                        if (c === 0) td.classList.add('selected-border-left');
                        if (c === cols - 1) td.classList.add('selected-border-right');
                    }
                } else if (selectedCol === c) {
                    // Column selection borders
                    if (c === selectedCol) {
                        td.classList.add('selected-border-left');
                        td.classList.add('selected-border-right');
                        if (r === 0) td.classList.add('selected-border-top');
                        // For bottom border of column, we'd need displayRows, but let's stick to visible range or just leave it open-ended
                        // Actually, displayRows is available in renderGrid but not globally. 
                        // Let's just do top for now or check cellElements length.
                        if (r === cellElements.length - 1) td.classList.add('selected-border-bottom');
                    }
                }
            }
        });
    });

    // Update header classes
    rowHeaderElements.forEach((th, r) => {
        if (selectedRow === r) th.classList.add('active');
        else th.classList.remove('active');
    });

    colHeaderElements.forEach((th, c) => {
        if (selectedCol === c) th.classList.add('active');
        else th.classList.remove('active');
    });
}

function clearSelection(resetRange = true) {
    selectedRow = null;
    selectedCol = null;
    if (resetRange) {
        selectionStart = null;
        selectionEnd = null;
    }
    updateDeleteButtons();
}

function updateDeleteButtons() {
    document.getElementById('delete-row-btn').disabled = selectedRow === null;
    document.getElementById('delete-col-btn').disabled = selectedCol === null;
}

function deleteSelectedRow() {
    if (selectedRow === null) return;
    if (confirm(`行 ${selectedRow + 1} を削除しますか？`)) {
        pushToUndoStack();
        spreadsheetData.splice(selectedRow, 1);
        clearSelection();
        renderGrid();
        showStatus('行を削除しました');
    }
}

function deleteSelectedCol() {
    if (selectedCol === null) return;
    if (confirm(`列 ${String.fromCharCode(65 + selectedCol)} を削除しますか？`)) {
        pushToUndoStack();
        spreadsheetData.forEach(row => {
            if (row && row.length > selectedCol) {
                row.splice(selectedCol, 1);
            }
        });
        clearSelection();
        renderGrid();
        showStatus('列を削除しました');
    }
}

function handleCopy(e) {
    if ((selectionStart && selectionEnd) || selectedRow !== null || selectedCol !== null) {
        let rMin, rMax, cMin, cMax;

        if (selectedRow !== null) {
            rMin = rMax = selectedRow;
            cMin = 0;
            cMax = cols - 1;
        } else if (selectedCol !== null) {
            cMin = cMax = selectedCol;
            rMin = 0;
            const displayRows = Math.max(rowCount, spreadsheetData.length + 10);
            rMax = displayRows - 1;
        } else {
            rMin = Math.min(selectionStart.r, selectionEnd.r);
            rMax = Math.max(selectionStart.r, selectionEnd.r);
            cMin = Math.min(selectionStart.c, selectionEnd.c);
            cMax = Math.max(selectionStart.c, selectionEnd.c);
        }

        let output = "";
        for (let r = rMin; r <= rMax; r++) {
            let rowData = [];
            for (let c = cMin; c <= cMax; c++) {
                rowData.push((spreadsheetData[r] && spreadsheetData[r][c]) || "");
            }
            output += rowData.join("\t") + "\n";
        }

        e.clipboardData.setData('text/plain', output);
        e.preventDefault();
        showStatus('範囲をコピーしました');
    }
}

async function handlePaste(e) {
    const activeCell = document.activeElement;
    if (!activeCell || !activeCell.classList.contains('cell')) return;

    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedText = clipboardData.getData('text');

    if (pastedText.includes('\t') || pastedText.includes('\n')) {
        e.preventDefault();
        pushToUndoStack();
        const rows_pasted = pastedText.split(/\r?\n/);
        const startRow = activeCell.parentElement.rowIndex - 1;
        const startCol = activeCell.cellIndex - 1;

        rows_pasted.forEach((rowData, rIdx) => {
            if (rowData.trim() === '' && rIdx === rows_pasted.length - 1) return;
            const cells_pasted = rowData.split('\t');
            cells_pasted.forEach((cellData, cIdx) => {
                const targetRow = startRow + rIdx;
                const targetCol = startCol + cIdx;
                if (targetCol < cols) {
                    updateData(targetRow, targetCol, cellData);
                }
            });
        });
        renderGrid();
        showStatus('貼り付けました');
    }
}

function pushToUndoStack() {
    // Deep copy spreadsheetData
    const snapshot = spreadsheetData.map(row => row ? [...row] : []);
    undoStack.push(snapshot);
    if (undoStack.length > maxUndoSteps) {
        undoStack.shift();
    }
}

function undo() {
    if (undoStack.length === 0) {
        showStatus('これ以上戻せません');
        return;
    }
    spreadsheetData = undoStack.pop();
    renderGrid();
    showStatus('元に戻しました');
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
