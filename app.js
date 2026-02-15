
/* global dayjs, Dexie */
(() => {
  const DATE_FMT = 'YYYY-MM-DD';
  const TASK_STATUS = {
    NONE: 'none',
    ACTIVE: 'active',
    DONE: 'done'
  };
  dayjs.extend(dayjs_plugin_customParseFormat);

  const state = {
    schemaVersion: 1,
    tasks: [],
    holidays: [],
    manualOrder: []
  };

  const uiState = {
    sort: { key: 'manual', dir: 'asc' },
    filter: { text: '', category: '', assignee: '' },
    selectedId: null,
    dayWidth: 24,
    ganttRange: null,
    leftCollapsed: false,
    rightCollapsed: false,
    leftPaneWidthPx: null,
    paneHeaderBase: null
  };

  const history = {
    stack: [],
    index: -1,
    locked: false,
    limit: 100
  };

  let holidaySet = new Set();
  let db = null;
  let storageMode = 'none';
  let saveTimer = null;
  let paneResizeState = null;

  const els = {};

  function qs(id) {
    return document.getElementById(id);
  }

  function initElements() {
    els.app = qs('app');
    els.main = document.querySelector('.main');
    els.addTaskBtn = qs('addTaskBtn');
    els.toggleLeftBtn = qs('toggleLeftBtn');
    els.toggleRightBtn = qs('toggleRightBtn');
    els.paneSplitter = qs('paneSplitter');
    els.undoBtn = qs('undoBtn');
    els.redoBtn = qs('redoBtn');
    els.exportJsonBtn = qs('exportJsonBtn');
    els.resetGanttBtn = qs('resetGanttBtn');
    els.importJsonBtn = qs('importJsonBtn');
    els.manageHolidayBtn = qs('manageHolidayBtn');
    els.exportMermaidBtn = qs('exportMermaidBtn');
    els.storageBadge = qs('storageBadge');

    els.filterText = qs('filterText');
    els.filterCategory = qs('filterCategory');
    els.filterAssignee = qs('filterAssignee');

    els.taskTbody = qs('taskTbody');
    els.tableWrap = qs('tableWrap');

    els.ganttHeader = qs('ganttHeader');
    els.ganttBody = qs('ganttBody');
    els.ganttBg = qs('ganttBg');
    els.ganttLinks = qs('ganttLinks');
    els.ganttRows = qs('ganttRows');
    els.todayLine = qs('todayLine');

    els.zoomRange = qs('zoomRange');
    els.rangeLabel = qs('rangeLabel');

    els.detailPanel = qs('detailPanel');
    els.closePanelBtn = qs('closePanelBtn');
    els.detailTitle = qs('detailTitle');
    els.detailCategory = qs('detailCategory');
    els.detailAssignee = qs('detailAssignee');
    els.detailStatus = qs('detailStatus');
    els.detailStart = qs('detailStart');
    els.detailEnd = qs('detailEnd');
    els.detailDuration = qs('detailDuration');
    els.detailMilestone = qs('detailMilestone');
    els.detailNotes = qs('detailNotes');
    els.dependencyList = qs('dependencyList');
    els.addDependencyBtn = qs('addDependencyBtn');
    els.deleteTaskBtn = qs('deleteTaskBtn');

    els.overlay = qs('overlay');
    els.mermaidModal = qs('mermaidModal');
    els.mermaidGroup = qs('mermaidGroup');
    els.mermaidDateMode = qs('mermaidDateMode');
    els.mermaidOutput = qs('mermaidOutput');
    els.copyMermaidBtn = qs('copyMermaidBtn');
    els.importMermaidBtn = qs('importMermaidBtn');
    els.closeMermaidBtn = qs('closeMermaidBtn');
    els.holidayModal = qs('holidayModal');
    els.closeHolidayBtn = qs('closeHolidayBtn');
    els.holidayDateInput = qs('holidayDateInput');
    els.holidayAddBtn = qs('holidayAddBtn');
    els.holidayImportBtn = qs('holidayImportBtn');
    els.holidayExportBtn = qs('holidayExportBtn');
    els.holidayList = qs('holidayList');

    els.jsonFileInput = qs('jsonFileInput');
    els.holidayFileInput = qs('holidayFileInput');
    els.mermaidFileInput = qs('mermaidFileInput');
  }

  function setupEvents() {
    els.addTaskBtn.addEventListener('click', addTask);
    if (els.toggleLeftBtn) {
      els.toggleLeftBtn.addEventListener('click', toggleLeftPane);
    }
    if (els.toggleRightBtn) {
      els.toggleRightBtn.addEventListener('click', toggleRightPane);
    }
    if (els.paneSplitter) {
      els.paneSplitter.addEventListener('pointerdown', startPaneResize);
    }
    els.undoBtn.addEventListener('click', undo);
    els.redoBtn.addEventListener('click', redo);
    els.exportJsonBtn.addEventListener('click', exportJson);
    if (els.resetGanttBtn) {
      els.resetGanttBtn.addEventListener('click', resetGantt);
    }
    els.importJsonBtn.addEventListener('click', () => els.jsonFileInput.click());
    if (els.manageHolidayBtn) {
      els.manageHolidayBtn.addEventListener('click', openHolidayModal);
    }
    els.exportMermaidBtn.addEventListener('click', () => openMermaidModal());

    els.filterText.addEventListener('input', () => {
      uiState.filter.text = els.filterText.value.trim();
      render();
    });
    els.filterCategory.addEventListener('change', () => {
      uiState.filter.category = els.filterCategory.value;
      render();
    });
    els.filterAssignee.addEventListener('change', () => {
      uiState.filter.assignee = els.filterAssignee.value;
      render();
    });

    els.zoomRange.addEventListener('input', () => {
      const oldWidth = uiState.dayWidth;
      const oldScroll = els.ganttBody ? els.ganttBody.scrollLeft : 0;
      uiState.dayWidth = Number(els.zoomRange.value);
      document.documentElement.style.setProperty('--day-width', `${uiState.dayWidth}px`);
      renderGantt();
      const newScroll = oldWidth ? (oldScroll / oldWidth) * uiState.dayWidth : oldScroll;
      if (els.ganttBody) els.ganttBody.scrollLeft = newScroll;
      if (els.ganttHeader) els.ganttHeader.scrollLeft = newScroll;
      syncLayout();
    });

    els.taskTbody.addEventListener('click', (event) => {
      const row = event.target.closest('tr');
      if (!row) return;
      openDetail(row.dataset.id);
    });

    els.taskTbody.addEventListener('dragstart', (event) => {
      const row = event.target.closest('tr');
      if (!row) return;
      uiState.sort.key = 'manual';
      event.dataTransfer.setData('text/plain', row.dataset.id);
    });

    els.taskTbody.addEventListener('dragover', (event) => {
      if (uiState.sort.key !== 'manual') return;
      event.preventDefault();
    });

    els.taskTbody.addEventListener('drop', (event) => {
      if (uiState.sort.key !== 'manual') return;
      event.preventDefault();
      const fromId = event.dataTransfer.getData('text/plain');
      const row = event.target.closest('tr');
      if (!row) return;
      const toId = row.dataset.id;
      if (fromId && toId && fromId !== toId) {
        moveManualOrder(fromId, toId);
        commitStateChange();
      }
    });

    document.querySelectorAll('.task-table thead th[data-key]').forEach((th) => {
      th.addEventListener('click', () => {
        const key = th.dataset.key;
        if (uiState.sort.key === key) {
          uiState.sort.dir = uiState.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          uiState.sort.key = key;
          uiState.sort.dir = 'asc';
        }
        render();
      });
    });

    els.tableWrap.addEventListener('scroll', () => {
      if (els.ganttBody.scrollTop !== els.tableWrap.scrollTop) {
        els.ganttBody.scrollTop = els.tableWrap.scrollTop;
      }
    });

    els.ganttBody.addEventListener('scroll', () => {
      if (els.tableWrap.scrollTop !== els.ganttBody.scrollTop) {
        els.tableWrap.scrollTop = els.ganttBody.scrollTop;
      }
      if (els.ganttHeader) {
        els.ganttHeader.scrollLeft = els.ganttBody.scrollLeft;
      }
    });

    els.closePanelBtn.addEventListener('click', closeDetail);
    els.deleteTaskBtn.addEventListener('click', deleteTask);

    els.detailTitle.addEventListener('input', () => updateDetailField('title', els.detailTitle.value));
    els.detailCategory.addEventListener('input', () => updateDetailField('category', els.detailCategory.value));
    els.detailAssignee.addEventListener('input', () => updateDetailField('assignee', els.detailAssignee.value));
    if (els.detailStatus) {
      els.detailStatus.addEventListener('change', () => updateDetailStatus(els.detailStatus.value));
    }
    els.detailStart.addEventListener('change', () => updateDetailDate('start', els.detailStart.value));
    els.detailEnd.addEventListener('change', () => updateDetailDate('end', els.detailEnd.value));
    els.detailDuration.addEventListener('change', () => updateDetailDuration(els.detailDuration.value));
    if (els.detailMilestone) {
      els.detailMilestone.addEventListener('change', () => updateDetailMilestone(els.detailMilestone.checked));
    }
    els.detailNotes.addEventListener('input', () => updateDetailField('notes', els.detailNotes.value));
    els.addDependencyBtn.addEventListener('click', addDependencyRow);

    els.overlay.addEventListener('click', closeAllModals);
    els.closeMermaidBtn.addEventListener('click', closeMermaidModal);
    els.copyMermaidBtn.addEventListener('click', copyMermaid);
    if (els.importMermaidBtn) {
      els.importMermaidBtn.addEventListener('click', () => els.mermaidFileInput.click());
    }
    els.mermaidGroup.addEventListener('change', renderMermaid);
    if (els.mermaidDateMode) {
      els.mermaidDateMode.addEventListener('change', renderMermaid);
    }
    if (els.closeHolidayBtn) {
      els.closeHolidayBtn.addEventListener('click', closeHolidayModal);
    }
    if (els.holidayAddBtn) {
      els.holidayAddBtn.addEventListener('click', addHolidayFromInput);
    }
    if (els.holidayImportBtn) {
      els.holidayImportBtn.addEventListener('click', () => els.holidayFileInput.click());
    }
    if (els.holidayDateInput) {
      els.holidayDateInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          addHolidayFromInput();
        }
      });
    }
    if (els.holidayExportBtn) {
      els.holidayExportBtn.addEventListener('click', exportHolidays);
    }

    els.jsonFileInput.addEventListener('change', handleJsonImport);
    els.holidayFileInput.addEventListener('change', handleHolidayImport);
    if (els.mermaidFileInput) {
      els.mermaidFileInput.addEventListener('change', handleMermaidImport);
    }

    window.addEventListener('keydown', (event) => {
      if (handleShortcuts(event)) return;
      if (event.ctrlKey && !event.shiftKey && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        undo();
      } else if (event.ctrlKey && (event.key.toLowerCase() === 'y' || (event.shiftKey && event.key.toLowerCase() === 'z'))) {
        event.preventDefault();
        redo();
      }
    });

    window.addEventListener('resize', () => {
      applyPaneWidth();
      syncLayout();
    });
  }

  function showToast(message, type = 'info') {
    let toast = document.getElementById('toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'toast';
      toast.style.position = 'fixed';
      toast.style.bottom = '20px';
      toast.style.left = '50%';
      toast.style.transform = 'translateX(-50%)';
      toast.style.padding = '10px 16px';
      toast.style.borderRadius = '8px';
      toast.style.background = '#1f2937';
      toast.style.color = '#fff';
      toast.style.fontSize = '13px';
      toast.style.zIndex = '40';
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 0.2s ease';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.style.background = type === 'error' ? '#b42318' : '#1f2937';
    requestAnimationFrame(() => {
      toast.style.opacity = '1';
    });
    setTimeout(() => {
      toast.style.opacity = '0';
    }, 1800);
  }

  function generateId() {
    if (crypto && crypto.randomUUID) return crypto.randomUUID();
    return `t-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
  }

  function parseDate(value) {
    const d = dayjs(value, DATE_FMT, true);
    return d.isValid() ? d.startOf('day') : dayjs().startOf('day');
  }

  function toISO(date) {
    return dayjs(date).format(DATE_FMT);
  }

  function normalizeTaskStatus(value) {
    const status = String(value || '').toLowerCase();
    if (status === TASK_STATUS.ACTIVE) return TASK_STATUS.ACTIVE;
    if (status === TASK_STATUS.DONE) return TASK_STATUS.DONE;
    return TASK_STATUS.NONE;
  }

  function getStatusLabel(status) {
    const normalized = normalizeTaskStatus(status);
    if (normalized === TASK_STATUS.ACTIVE) return '進行中';
    if (normalized === TASK_STATUS.DONE) return '完了';
    return '未設定';
  }

  function isWorkday(date) {
    const d = dayjs(date);
    const day = d.day();
    if (day === 0 || day === 6) return false;
    return !holidaySet.has(d.format(DATE_FMT));
  }

  function adjustToWorkday(date, direction = 1) {
    let d = dayjs(date).startOf('day');
    let step = direction >= 0 ? 1 : -1;
    while (!isWorkday(d)) {
      d = d.add(step, 'day');
    }
    return d;
  }

  function addWorkdays(date, count) {
    let d = dayjs(date).startOf('day');
    if (count === 0) return d;
    let remaining = Math.abs(count);
    let step = count > 0 ? 1 : -1;
    while (remaining > 0) {
      d = d.add(step, 'day');
      if (isWorkday(d)) remaining -= 1;
    }
    return d;
  }

  function workdayCount(start, end) {
    let s = dayjs(start).startOf('day');
    let e = dayjs(end).startOf('day');
    if (e.isBefore(s, 'day')) return 1;
    let count = 0;
    let cur = s;
    while (!cur.isAfter(e, 'day')) {
      if (isWorkday(cur)) count += 1;
      cur = cur.add(1, 'day');
    }
    return Math.max(count, 1);
  }

  function ensureHolidaySet() {
    holidaySet = new Set((state.holidays || []).filter(Boolean));
  }

  function normalizeTask(task) {
    if (!task.id) task.id = generateId();
    task.title = task.title || 'Untitled';
    task.category = task.category || '';
    task.assignee = task.assignee || '';
    task.status = normalizeTaskStatus(task.status);
    task.notes = task.notes || '';
    task.dependencies = Array.isArray(task.dependencies) ? task.dependencies : [];
    task.isMilestone = Boolean(task.isMilestone);
    task.duration = Number(task.duration) || 1;
    if (task.duration < 1) task.duration = 1;
    task.start = task.start || toISO(adjustToWorkday(dayjs()));
    task.start = toISO(adjustToWorkday(parseDate(task.start)));
    if (task.end) {
      task.end = toISO(adjustToWorkday(parseDate(task.end)));
      task.duration = workdayCount(task.start, task.end);
    }
    if (task.isMilestone) {
      task.duration = 1;
      task.end = task.start;
      return task;
    }
    task.end = toISO(addWorkdays(parseDate(task.start), task.duration - 1));
    return task;
  }

  function syncManualOrder() {
    const ids = new Set(state.tasks.map((t) => t.id));
    state.manualOrder = state.manualOrder.filter((id) => ids.has(id));
    state.tasks.forEach((task) => {
      if (!state.manualOrder.includes(task.id)) {
        state.manualOrder.push(task.id);
      }
    });
  }

  function getFilteredTasks() {
    let tasks = [...state.tasks];
    const text = uiState.filter.text.toLowerCase();
    if (text) {
      tasks = tasks.filter((t) => `${t.title} ${t.notes}`.toLowerCase().includes(text));
    }
    if (uiState.filter.category) {
      tasks = tasks.filter((t) => t.category === uiState.filter.category);
    }
    if (uiState.filter.assignee) {
      tasks = tasks.filter((t) => t.assignee === uiState.filter.assignee);
    }

    if (uiState.sort.key === 'manual') {
      const index = new Map(state.manualOrder.map((id, i) => [id, i]));
      tasks.sort((a, b) => (index.get(a.id) ?? 0) - (index.get(b.id) ?? 0));
    } else {
      const key = uiState.sort.key;
      const dir = uiState.sort.dir === 'asc' ? 1 : -1;
      tasks.sort((a, b) => {
        let va = a[key];
        let vb = b[key];
        if (key === 'dependencies') {
          va = a.dependencies.length;
          vb = b.dependencies.length;
        }
        if (key === 'start' || key === 'end') {
          va = a[key] || '';
          vb = b[key] || '';
        }
        if (typeof va === 'string') {
          va = va.toLowerCase();
          vb = vb.toLowerCase();
        }
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
      });
    }
    return tasks;
  }

  function renderFilters() {
    const categories = [...new Set(state.tasks.map((t) => t.category).filter(Boolean))].sort();
    const assignees = [...new Set(state.tasks.map((t) => t.assignee).filter(Boolean))].sort();
    renderSelectOptions(els.filterCategory, '分類: すべて', categories, uiState.filter.category);
    renderSelectOptions(els.filterAssignee, '担当: すべて', assignees, uiState.filter.assignee);
  }

  function renderSelectOptions(select, placeholder, values, current) {
    const prev = select.value;
    select.innerHTML = '';
    const option = document.createElement('option');
    option.value = '';
    option.textContent = placeholder;
    select.appendChild(option);
    values.forEach((value) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = value;
      select.appendChild(opt);
    });
    select.value = current || prev || '';
  }

  function renderTable() {
    const tasks = getFilteredTasks();
    const tbody = els.taskTbody;
    tbody.innerHTML = '';
    const frag = document.createDocumentFragment();
    tasks.forEach((task) => {
      const tr = document.createElement('tr');
      tr.dataset.id = task.id;
      tr.draggable = uiState.sort.key === 'manual';
      if (task.id === uiState.selectedId) tr.classList.add('selected');
      const title = task.isMilestone
        ? `<span class="milestone-dot" aria-hidden="true"></span>${escapeHtml(task.title)}`
        : escapeHtml(task.title);
      const status = normalizeTaskStatus(task.status);
      const statusLabel = getStatusLabel(status);
      tr.innerHTML = `
        <td class="col-handle"><span class="drag-handle">&#xe700;</span></td>
        <td>${title}</td>
        <td>${escapeHtml(task.category)}</td>
        <td>${escapeHtml(task.assignee)}</td>
        <td><span class="status-badge status-${status}">${statusLabel}</span></td>
        <td>${escapeHtml(task.start)}</td>
        <td>${escapeHtml(task.end)}</td>
        <td>${task.duration}</td>
        <td>${task.dependencies.length}</td>
      `;
      frag.appendChild(tr);
    });
    tbody.appendChild(frag);
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function renderGantt() {
    const tasks = getFilteredTasks();
    const { start, end, days } = computeGanttRange(tasks);
    uiState.ganttRange = { start, end, days };
    els.rangeLabel.textContent = `${toISO(start)} 〜 ${toISO(end)} (${days}日)`;

    renderGanttHeader(start, days);
    renderGanttBackground(days, tasks.length);
    renderGanttRows(tasks, start, days);
    renderDependencies(tasks, start, days);
    renderTodayLine(start, end);
    if (els.ganttHeader && els.ganttBody) {
      els.ganttHeader.scrollLeft = els.ganttBody.scrollLeft;
    }
  }

  function syncHeaderHeights() {
    // Intentionally left empty. Header height is fixed via CSS variables.
  }

  function syncPaneHeaderHeights() {
    const leftHeader = document.querySelector('.pane.left .pane-header');
    const rightHeader = document.querySelector('.pane.right .pane-header');
    if (!leftHeader || !rightHeader) return;
    if (uiState.paneHeaderBase === null) {
      const base = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--pane-header-height')) || 0;
      uiState.paneHeaderBase = base;
    }
    const leftHeight = Math.max(leftHeader.scrollHeight, leftHeader.offsetHeight);
    const rightHeight = Math.max(rightHeader.scrollHeight, rightHeader.offsetHeight);
    const current = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--pane-header-height')) || 0;
    const target = Math.max(uiState.paneHeaderBase || 0, leftHeight, rightHeight);
    if (target > 0 && Math.abs(target - current) > 0.5) {
      document.documentElement.style.setProperty('--pane-header-height', `${Math.ceil(target)}px`);
    }
  }

  function syncRowHeights() {
    const leftRow = els.taskTbody ? els.taskTbody.querySelector('tr') : null;
    if (!leftRow || !els.ganttBody) return false;
    const leftHeight = leftRow.getBoundingClientRect().height;
    if (leftHeight > 0) {
      els.ganttBody.style.setProperty('--row-height', `${leftHeight}px`);
      return true;
    }
    return false;
  }

  function syncRowAlignment() {
    // No-op: row alignment is fixed by shared row height.
  }

  function syncLayout() {
    syncPaneHeaderHeights();
    document.documentElement.style.setProperty('--gantt-body-offset', '0px');
  }

  function getPaneWidthBounds() {
    if (!els.main) {
      return { min: 280, max: 1200 };
    }
    const totalWidth = els.main.getBoundingClientRect().width;
    const splitterWidth = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--splitter-width')) || 10;
    const minLeft = 280;
    const minRight = 360;
    const maxLeft = Math.max(minLeft, totalWidth - minRight - splitterWidth);
    return { min: minLeft, max: maxLeft };
  }

  function applyPaneWidth() {
    if (!els.main || uiState.leftCollapsed || uiState.rightCollapsed) return;
    const { min, max } = getPaneWidthBounds();
    if (uiState.leftPaneWidthPx === null) {
      const leftPane = document.querySelector('.pane.left');
      if (leftPane) {
        const measured = leftPane.getBoundingClientRect().width;
        if (measured > 0) {
          uiState.leftPaneWidthPx = measured;
        }
      }
    }
    const current = uiState.leftPaneWidthPx === null ? max : uiState.leftPaneWidthPx;
    const clamped = Math.min(Math.max(current, min), max);
    uiState.leftPaneWidthPx = clamped;
    document.documentElement.style.setProperty('--left-pane-width', `${clamped}px`);
  }

  function startPaneResize(event) {
    if (uiState.leftCollapsed || uiState.rightCollapsed) return;
    const leftPane = document.querySelector('.pane.left');
    if (!leftPane) return;
    paneResizeState = {
      startX: event.clientX,
      startWidth: leftPane.getBoundingClientRect().width
    };
    document.body.classList.add('resizing');
    window.addEventListener('pointermove', onPaneResize);
    window.addEventListener('pointerup', endPaneResize);
    event.preventDefault();
  }

  function onPaneResize(event) {
    if (!paneResizeState) return;
    const delta = event.clientX - paneResizeState.startX;
    uiState.leftPaneWidthPx = paneResizeState.startWidth + delta;
    applyPaneWidth();
    syncLayout();
  }

  function endPaneResize() {
    if (!paneResizeState) return;
    paneResizeState = null;
    document.body.classList.remove('resizing');
    window.removeEventListener('pointermove', onPaneResize);
    window.removeEventListener('pointerup', endPaneResize);
  }

  function getGanttRowHeight() {
    if (els.ganttBody) {
      const localValue = getComputedStyle(els.ganttBody).getPropertyValue('--row-height');
      const parsed = parseFloat(localValue);
      if (parsed > 0) return parsed;
    }
    const rootValue = getComputedStyle(document.documentElement).getPropertyValue('--row-height');
    const fallback = parseFloat(rootValue);
    return fallback > 0 ? fallback : 30;
  }

  function computeGanttRange(tasks) {
    const today = dayjs().startOf('day');
    let min = today;
    let max = today.add(30, 'day');
    if (tasks.length) {
      const starts = tasks.map((t) => parseDate(t.start));
      const ends = tasks.map((t) => parseDate(t.end));
      min = starts.reduce((a, b) => (a.isBefore(b) ? a : b));
      max = ends.reduce((a, b) => (a.isAfter(b) ? a : b));
      min = min.subtract(7, 'day');
      max = max.add(14, 'day');
    }
    const days = max.diff(min, 'day') + 1;
    return { start: min, end: max, days };
  }

  function renderGanttHeader(start, days) {
    els.ganttHeader.innerHTML = '';
    const monthRow = document.createElement('div');
    monthRow.className = 'gantt-month-row';
    const dayRow = document.createElement('div');
    dayRow.className = 'gantt-day-row';

    let current = start.clone();
    let monthStart = current.clone();
    for (let i = 0; i < days; i += 1) {
      const dayCell = document.createElement('div');
      dayCell.className = 'gantt-day';
      if (!isWorkday(current)) dayCell.classList.add('nonwork');
      dayCell.textContent = current.date();
      dayRow.appendChild(dayCell);

      const next = current.add(1, 'day');
      if (next.month() !== current.month() || i === days - 1) {
        const spanDays = next.diff(monthStart, 'day');
        const monthCell = document.createElement('div');
        monthCell.className = 'gantt-month';
        monthCell.style.width = `${spanDays * uiState.dayWidth}px`;
        monthCell.textContent = `${monthStart.format('YYYY-MM')}`;
        monthRow.appendChild(monthCell);
        monthStart = next;
      }
      current = next;
    }
    els.ganttHeader.appendChild(monthRow);
    els.ganttHeader.appendChild(dayRow);
  }
  function renderGanttBackground(days, rowCount) {
    els.ganttBg.innerHTML = '';
    els.ganttBg.style.width = `${days * uiState.dayWidth}px`;
    const frag = document.createDocumentFragment();
    let current = uiState.ganttRange.start.clone();
    for (let i = 0; i < days; i += 1) {
      const col = document.createElement('div');
      col.className = 'day-col';
      if (!isWorkday(current)) col.classList.add('nonwork');
      frag.appendChild(col);
      current = current.add(1, 'day');
    }
    els.ganttBg.appendChild(frag);
    const rowHeight = getGanttRowHeight();
    els.ganttBg.style.height = `${rowCount * rowHeight}px`;
  }

  function renderDependencies(tasks, rangeStart, days) {
    if (!els.ganttLinks) return;
    const rowHeight = getGanttRowHeight();
    const barTop = 4;
    const barHeight = 18;
    const svgHeight = rowHeight * tasks.length;
    const svgWidth = days * uiState.dayWidth;

    const svg = els.ganttLinks;
    svg.setAttribute('width', svgWidth);
    svg.setAttribute('height', svgHeight);
    svg.setAttribute('viewBox', `0 0 ${svgWidth} ${svgHeight}`);
    svg.innerHTML = '';

    const ns = 'http://www.w3.org/2000/svg';
    const defs = document.createElementNS(ns, 'defs');
    const marker = document.createElementNS(ns, 'marker');
    marker.setAttribute('id', 'arrow');
    marker.setAttribute('markerWidth', '8');
    marker.setAttribute('markerHeight', '8');
    marker.setAttribute('refX', '7');
    marker.setAttribute('refY', '4');
    marker.setAttribute('orient', 'auto');
    const arrowPath = document.createElementNS(ns, 'path');
    arrowPath.setAttribute('d', 'M0,0 L8,4 L0,8 Z');
    arrowPath.setAttribute('fill', '#5b6b7f');
    marker.appendChild(arrowPath);
    defs.appendChild(marker);
    svg.appendChild(defs);

    const indexMap = new Map(tasks.map((t, i) => [t.id, i]));
    const taskMap = new Map(tasks.map((t) => [t.id, t]));

    tasks.forEach((task, rowIndex) => {
      task.dependencies.forEach((dep) => {
        const predIndex = indexMap.get(dep.id);
        const pred = taskMap.get(dep.id);
        if (predIndex === undefined || !pred) return;

        const predPos = computeBarPosition(pred, rangeStart);
        const succPos = computeBarPosition(task, rangeStart);

        const fromY = predIndex * rowHeight + barTop + barHeight / 2;
        const toY = rowIndex * rowHeight + barTop + barHeight / 2;
        let fromX;
        let toX;

        if (dep.type === 'SS') {
          fromX = predPos.left;
          toX = succPos.left;
        } else if (dep.type === 'FF') {
          fromX = predPos.left + predPos.width;
          toX = succPos.left + succPos.width;
        } else if (dep.type === 'SF') {
          fromX = predPos.left;
          toX = succPos.left + succPos.width;
        } else {
          fromX = predPos.left + predPos.width;
          toX = succPos.left;
        }

        const midX = fromX + Math.max(10, Math.abs(toX - fromX) * 0.3) * (fromX <= toX ? 1 : -1);
        const d = `M${fromX},${fromY} L${midX},${fromY} L${midX},${toY} L${toX},${toY}`;
        const path = document.createElementNS(ns, 'path');
        path.setAttribute('d', d);
        path.setAttribute('fill', 'none');
        path.setAttribute('stroke', '#5b6b7f');
        path.setAttribute('stroke-width', '1.2');
        path.setAttribute('marker-end', 'url(#arrow)');
        svg.appendChild(path);
      });
    });
  }

  function renderGanttRows(tasks, rangeStart, days) {
    els.ganttRows.innerHTML = '';
    els.ganttRows.style.width = `${days * uiState.dayWidth}px`;
    const frag = document.createDocumentFragment();
    const rowHeight = getGanttRowHeight();
    tasks.forEach((task) => {
      const row = document.createElement('div');
      row.className = 'gantt-row';
      row.style.height = `${rowHeight}px`;

      const bar = document.createElement('div');
      bar.className = 'gantt-bar';
      bar.dataset.id = task.id;
      if (task.id === uiState.selectedId) bar.classList.add('selected');
      bar.classList.add(`status-${normalizeTaskStatus(task.status)}`);
      const isMilestone = Boolean(task.isMilestone);
      if (isMilestone) {
        bar.classList.add('milestone');
        const label = document.createElement('span');
        label.className = 'milestone-label';
        label.textContent = task.title;
        bar.appendChild(label);
      } else {
        bar.textContent = task.title;
      }

      const { left, width } = computeBarPosition(task, rangeStart);
      bar.style.left = `${left}px`;
      bar.style.width = `${width}px`;

      if (!isMilestone) {
        const handleLeft = document.createElement('div');
        handleLeft.className = 'gantt-handle left';
        const handleRight = document.createElement('div');
        handleRight.className = 'gantt-handle right';
        bar.appendChild(handleLeft);
        bar.appendChild(handleRight);
      }

      row.appendChild(bar);
      frag.appendChild(row);
    });
    els.ganttRows.appendChild(frag);

    setupGanttDrag();
  }

  function computeBarPosition(task, rangeStart) {
    const start = parseDate(task.start);
    const end = parseDate(task.end);
    const left = Math.max(0, start.diff(rangeStart, 'day') * uiState.dayWidth);
    const width = Math.max(uiState.dayWidth, (end.diff(start, 'day') + 1) * uiState.dayWidth);
    return { left, width };
  }

  function renderTodayLine(rangeStart, rangeEnd) {
    const today = dayjs().startOf('day');
    if (today.isBefore(rangeStart, 'day') || today.isAfter(rangeEnd, 'day')) {
      els.todayLine.style.display = 'none';
      return;
    }
    els.todayLine.style.display = 'block';
    els.todayLine.style.left = `${today.diff(rangeStart, 'day') * uiState.dayWidth}px`;
    els.todayLine.style.height = `${els.ganttRows.offsetHeight}px`;
  }

  function setupGanttDrag() {
    const bars = els.ganttRows.querySelectorAll('.gantt-bar');
    bars.forEach((bar) => {
      bar.onpointerdown = (event) => startDrag(event, bar);
    });
  }

  let dragState = null;

  function startDrag(event, bar) {
    const taskId = bar.dataset.id;
    const task = state.tasks.find((t) => t.id === taskId);
    if (!task) return;

    let mode = event.target.classList.contains('left')
      ? 'resize-left'
      : event.target.classList.contains('right')
      ? 'resize-right'
      : 'move';
    if (task.isMilestone) {
      mode = 'move';
    }

    dragState = {
      id: taskId,
      mode,
      startX: event.clientX,
      origStart: task.start,
      origEnd: task.end,
      origDuration: task.duration,
      bar,
      moved: false
    };

    window.addEventListener('pointermove', onDragMove);
    window.addEventListener('pointerup', endDrag);
  }

  function onDragMove(event) {
    if (!dragState) return;
    const delta = Math.round((event.clientX - dragState.startX) / uiState.dayWidth);
    if (dragState.lastDelta === delta) return;
    dragState.lastDelta = delta;
    if (delta !== 0) dragState.moved = true;

    let start = parseDate(dragState.origStart).add(delta, 'day');
    let end = parseDate(dragState.origEnd).add(delta, 'day');
    let duration = dragState.origDuration;

    if (dragState.mode === 'move') {
      start = adjustToWorkday(start, 1);
      end = addWorkdays(start, duration - 1);
    } else if (dragState.mode === 'resize-right') {
      end = parseDate(dragState.origEnd).add(delta, 'day');
      end = adjustToWorkday(end, -1);
      if (end.isBefore(start, 'day')) end = start.clone();
      duration = workdayCount(start, end);
    } else if (dragState.mode === 'resize-left') {
      start = parseDate(dragState.origStart).add(delta, 'day');
      start = adjustToWorkday(start, 1);
      if (start.isAfter(end, 'day')) start = end.clone();
      duration = workdayCount(start, end);
    }

    dragState.preview = {
      start: toISO(start),
      end: toISO(end),
      duration
    };

    const { left, width } = computeBarPosition({ start: dragState.preview.start, end: dragState.preview.end }, uiState.ganttRange.start);
    dragState.bar.style.left = `${left}px`;
    dragState.bar.style.width = `${width}px`;
  }

  function endDrag() {
    if (!dragState) return;
    window.removeEventListener('pointermove', onDragMove);
    window.removeEventListener('pointerup', endDrag);

    if (!dragState.moved && dragState.mode === 'move') {
      openDetail(dragState.id);
      dragState = null;
      return;
    }

    if (dragState.preview) {
      const task = state.tasks.find((t) => t.id === dragState.id);
      if (task) {
        task.start = dragState.preview.start;
        task.end = dragState.preview.end;
        task.duration = dragState.preview.duration;
        scheduleAll();
        commitStateChange();
      }
    }

    dragState = null;
  }

  function openDetail(id) {
    uiState.selectedId = id;
    const task = state.tasks.find((t) => t.id === id);
    if (!task) return;
    els.detailTitle.value = task.title;
    els.detailCategory.value = task.category;
    els.detailAssignee.value = task.assignee;
    if (els.detailStatus) {
      els.detailStatus.value = normalizeTaskStatus(task.status);
    }
    els.detailStart.value = task.start;
    els.detailEnd.value = task.end;
    els.detailDuration.value = task.duration;
    if (els.detailMilestone) {
      els.detailMilestone.checked = Boolean(task.isMilestone);
    }
    els.detailNotes.value = task.notes;
    renderDependencyList(task);
    applyMilestoneControls(task);
    els.detailPanel.classList.add('open');
    render();
  }

  function closeDetail() {
    uiState.selectedId = null;
    els.detailPanel.classList.remove('open');
    render();
  }

  function applyMilestoneControls(task) {
    if (!els.detailMilestone) return;
    const isMilestone = Boolean(task && task.isMilestone);
    if (els.detailEnd) els.detailEnd.disabled = isMilestone;
    if (els.detailDuration) els.detailDuration.disabled = isMilestone;
  }

  function updateDetailField(field, value) {
    const task = getSelectedTask();
    if (!task) return;
    task[field] = value;
    commitStateChange({ skipSchedule: field === 'notes' });
  }

  function updateDetailStatus(value) {
    const task = getSelectedTask();
    if (!task) return;
    task.status = normalizeTaskStatus(value);
    commitStateChange({ skipSchedule: true });
  }

  function updateDetailDate(field, value) {
    const task = getSelectedTask();
    if (!task) return;
    if (!value) return;
    if (task.isMilestone) {
      const date = adjustToWorkday(parseDate(value), 1);
      task.start = toISO(date);
      task.end = task.start;
      task.duration = 1;
    } else {
      const date = adjustToWorkday(parseDate(value), field === 'end' ? -1 : 1);
      if (field === 'start') {
        task.start = toISO(date);
        task.end = toISO(addWorkdays(date, task.duration - 1));
      } else {
        task.end = toISO(date);
        task.duration = workdayCount(task.start, task.end);
      }
    }
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function updateDetailDuration(value) {
    const task = getSelectedTask();
    if (!task) return;
    if (task.isMilestone) {
      task.duration = 1;
      task.end = task.start;
    } else {
      const duration = Math.max(1, Number(value) || 1);
      task.duration = duration;
      task.end = toISO(addWorkdays(parseDate(task.start), duration - 1));
    }
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function updateDetailMilestone(enabled) {
    const task = getSelectedTask();
    if (!task) return;
    task.isMilestone = Boolean(enabled);
    if (task.isMilestone) {
      task.duration = 1;
      task.start = toISO(adjustToWorkday(parseDate(task.start), 1));
      task.end = task.start;
    } else {
      task.duration = Math.max(1, Number(task.duration) || 1);
      task.end = toISO(addWorkdays(parseDate(task.start), task.duration - 1));
    }
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function addDependencyRow() {
    const task = getSelectedTask();
    if (!task) return;
    const other = state.tasks.find((t) => t.id !== task.id);
    if (!other) {
      showToast('依存先のタスクが存在しません', 'error');
      return;
    }
    task.dependencies.push({ id: other.id, type: 'FS' });
    if (hasCycle(state.tasks)) {
      task.dependencies.pop();
      showToast('循環依存のため追加できません', 'error');
      return;
    }
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function renderDependencyList(task) {
    els.dependencyList.innerHTML = '';
    const options = state.tasks.filter((t) => t.id !== task.id);
    if (options.length === 0) {
      els.dependencyList.innerHTML = '<div class="muted">依存先のタスクが存在しません</div>';
      return;
    }
    task.dependencies.forEach((dep, index) => {
      const row = document.createElement('div');
      row.className = 'dep-row';

      const typeSelect = document.createElement('select');
      typeSelect.className = 'select';
      ['FS', 'SS', 'FF', 'SF'].forEach((type) => {
        const opt = document.createElement('option');
        opt.value = type;
        opt.textContent = type;
        typeSelect.appendChild(opt);
      });
      typeSelect.value = dep.type || 'FS';

      const taskSelect = document.createElement('select');
      taskSelect.className = 'select';
      options.forEach((optTask) => {
        const opt = document.createElement('option');
        opt.value = optTask.id;
        opt.textContent = optTask.title;
        taskSelect.appendChild(opt);
      });
      taskSelect.value = dep.id;

      const removeBtn = document.createElement('button');
      removeBtn.className = 'btn ghost';
      removeBtn.innerHTML = '<span class="icon">&#xE74D;</span>';

      typeSelect.addEventListener('change', () => updateDependency(index, { type: typeSelect.value, id: taskSelect.value }));
      taskSelect.addEventListener('change', () => updateDependency(index, { type: typeSelect.value, id: taskSelect.value }));
      removeBtn.addEventListener('click', () => removeDependency(index));

      row.appendChild(typeSelect);
      row.appendChild(taskSelect);
      row.appendChild(removeBtn);
      els.dependencyList.appendChild(row);
    });
  }

  function updateDependency(index, next) {
    const task = getSelectedTask();
    if (!task) return;
    const prev = task.dependencies[index];
    if (!prev) return;
    if (next.id === task.id) {
      showToast('自分自身のタスクに依存できません', 'error');
      return;
    }
    task.dependencies[index] = next;
    if (hasCycle(state.tasks)) {
      task.dependencies[index] = prev;
      showToast('循環依存のため変更できません', 'error');
      return;
    }
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function removeDependency(index) {
    const task = getSelectedTask();
    if (!task) return;
    task.dependencies.splice(index, 1);
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }
  function getSelectedTask() {
    return state.tasks.find((t) => t.id === uiState.selectedId);
  }

  function addTask() {
    const start = adjustToWorkday(dayjs(), 1);
    const task = normalizeTask({
      id: generateId(),
      title: 'New Task',
      category: '',
      assignee: '',
      status: TASK_STATUS.NONE,
      start: toISO(start),
      duration: 3,
      dependencies: [],
      notes: '',
      isMilestone: false
    });
    state.tasks.push(task);
    syncManualOrder();
    scheduleAll();
    commitStateChange();
    openDetail(task.id);
  }

  function deleteTask() {
    const task = getSelectedTask();
    if (!task) return;
    const idx = state.tasks.findIndex((t) => t.id === task.id);
    if (idx >= 0) state.tasks.splice(idx, 1);
    state.tasks.forEach((t) => {
      t.dependencies = t.dependencies.filter((d) => d.id !== task.id);
    });
    syncManualOrder();
    uiState.selectedId = null;
    commitStateChange();
    closeDetail();
  }

  function moveManualOrder(fromId, toId) {
    const list = state.manualOrder;
    const fromIdx = list.indexOf(fromId);
    const toIdx = list.indexOf(toId);
    if (fromIdx === -1 || toIdx === -1) return;
    list.splice(fromIdx, 1);
    list.splice(toIdx, 0, fromId);
  }

  function hasCycle(tasks) {
    const graph = new Map();
    tasks.forEach((task) => {
      graph.set(task.id, task.dependencies.map((d) => d.id).filter(Boolean));
    });

    const visiting = new Set();
    const visited = new Set();

    function dfs(node) {
      if (visiting.has(node)) return true;
      if (visited.has(node)) return false;
      visiting.add(node);
      const deps = graph.get(node) || [];
      for (const dep of deps) {
        if (graph.has(dep) && dfs(dep)) return true;
      }
      visiting.delete(node);
      visited.add(node);
      return false;
    }

    for (const node of graph.keys()) {
      if (dfs(node)) return true;
    }
    return false;
  }

  function scheduleAll() {
    if (hasCycle(state.tasks)) return false;
    const map = new Map(state.tasks.map((t) => [t.id, t]));
    const order = topoOrder(state.tasks);
    order.forEach((id) => {
      const task = map.get(id);
      if (!task) return;
      let start = parseDate(task.start);
      let duration = Math.max(1, Number(task.duration) || 1);
      if (task.isMilestone) duration = 1;
      let constraintStart = null;
      let constraintEnd = null;

      task.dependencies.forEach((dep) => {
        const pred = map.get(dep.id);
        if (!pred) return;
        const predStart = parseDate(pred.start);
        const predEnd = parseDate(pred.end);
        if (dep.type === 'FS') {
          const cs = addWorkdays(predEnd, 1);
          if (!constraintStart || cs.isAfter(constraintStart)) constraintStart = cs;
        } else if (dep.type === 'SS') {
          const cs = adjustToWorkday(predStart, 1);
          if (!constraintStart || cs.isAfter(constraintStart)) constraintStart = cs;
        } else if (dep.type === 'FF') {
          const ce = adjustToWorkday(predEnd, 1);
          if (!constraintEnd || ce.isAfter(constraintEnd)) constraintEnd = ce;
        } else if (dep.type === 'SF') {
          const ce = adjustToWorkday(predStart, 1);
          if (!constraintEnd || ce.isAfter(constraintEnd)) constraintEnd = ce;
        }
      });

      if (constraintStart && start.isBefore(constraintStart, 'day')) {
        start = constraintStart;
      }

      start = adjustToWorkday(start, 1);
      let end = addWorkdays(start, duration - 1);

      if (constraintEnd && end.isBefore(constraintEnd, 'day')) {
        end = constraintEnd;
        end = adjustToWorkday(end, 1);
        start = addWorkdays(end, -(duration - 1));
      }

      task.start = toISO(start);
      task.end = toISO(end);
      task.duration = duration;
    });
    return true;
  }

  function topoOrder(tasks) {
    const indeg = new Map();
    const adj = new Map();
    tasks.forEach((t) => {
      indeg.set(t.id, 0);
      adj.set(t.id, []);
    });
    tasks.forEach((t) => {
      t.dependencies.forEach((dep) => {
        if (adj.has(dep.id)) {
          adj.get(dep.id).push(t.id);
          indeg.set(t.id, (indeg.get(t.id) || 0) + 1);
        }
      });
    });
    const queue = [];
    indeg.forEach((v, k) => {
      if (v === 0) queue.push(k);
    });
    const order = [];
    while (queue.length) {
      const node = queue.shift();
      order.push(node);
      adj.get(node).forEach((next) => {
        indeg.set(next, indeg.get(next) - 1);
        if (indeg.get(next) === 0) queue.push(next);
      });
    }
    return order.length === tasks.length ? order : tasks.map((t) => t.id);
  }

  function commitStateChange(options = {}) {
    if (!options.skipSchedule) scheduleAll();
    pushHistory();
    scheduleSave();
    render();
  }

  function pushHistory() {
    if (history.locked) return;
    const snapshot = JSON.stringify({
      schemaVersion: state.schemaVersion,
      tasks: state.tasks,
      holidays: state.holidays,
      manualOrder: state.manualOrder
    });
    if (history.index >= 0 && history.stack[history.index] === snapshot) return;
    history.stack = history.stack.slice(0, history.index + 1);
    history.stack.push(snapshot);
    if (history.stack.length > history.limit) {
      history.stack.shift();
    }
    history.index = history.stack.length - 1;
  }

  function applySnapshot(snapshot) {
    const data = JSON.parse(snapshot);
    state.schemaVersion = data.schemaVersion || 1;
    state.tasks = Array.isArray(data.tasks) ? data.tasks.map(normalizeTask) : [];
    state.holidays = Array.isArray(data.holidays) ? data.holidays : [];
    state.manualOrder = Array.isArray(data.manualOrder) ? data.manualOrder : [];
    ensureHolidaySet();
    syncManualOrder();
  }

  function undo() {
    if (history.index <= 0) return;
    history.locked = true;
    history.index -= 1;
    applySnapshot(history.stack[history.index]);
    history.locked = false;
    scheduleAll();
    scheduleSave();
    render();
  }

  function redo() {
    if (history.index >= history.stack.length - 1) return;
    history.locked = true;
    history.index += 1;
    applySnapshot(history.stack[history.index]);
    history.locked = false;
    scheduleAll();
    scheduleSave();
    render();
  }

  function updateStorageBadge() {
    if (!els.storageBadge) return;
    if (storageMode === 'indexeddb') {
      els.storageBadge.textContent = '保存: IndexedDB';
    } else if (storageMode === 'localStorage') {
      els.storageBadge.textContent = '保存: localStorage (簡易)';
    } else {
      els.storageBadge.textContent = '保存なし: JSONエクスポート推奨';
    }
  }

  async function initStorage() {
    if (!('indexedDB' in window)) {
      storageMode = 'localStorage';
      return;
    }
    try {
      db = new Dexie('LocalGanttDB');
      db.version(1).stores({ app: 'id' });
      await db.open();
      storageMode = 'indexeddb';
    } catch (error) {
      console.warn('IndexedDB unavailable', error);
      storageMode = 'localStorage';
    }
  }

  async function loadState() {
    await initStorage();
    let data = null;
    if (storageMode === 'indexeddb') {
      data = await db.table('app').get('state');
      data = data ? data.value : null;
    }
    if (!data && storageMode !== 'indexeddb') {
      const raw = localStorage.getItem('ganttState');
      if (raw) data = JSON.parse(raw);
    }
    if (data) {
      state.schemaVersion = data.schemaVersion || 1;
      state.tasks = Array.isArray(data.tasks) ? data.tasks.map(normalizeTask) : [];
      state.holidays = Array.isArray(data.holidays) ? data.holidays : [];
      state.manualOrder = Array.isArray(data.manualOrder) ? data.manualOrder : [];
    }
    ensureHolidaySet();
    syncManualOrder();
    scheduleAll();
    pushHistory();
    render();
  }
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveState, 400);
  }

  async function saveState() {
    const payload = {
      schemaVersion: state.schemaVersion,
      tasks: state.tasks,
      holidays: state.holidays,
      manualOrder: state.manualOrder
    };
    if (storageMode === 'indexeddb') {
      await db.table('app').put({ id: 'state', value: payload });
    } else if (storageMode === 'localStorage') {
      try {
        localStorage.setItem('ganttState', JSON.stringify(payload));
      } catch (error) {
        storageMode = 'none';
        showToast('キャッシュが使用できませんでした。JSONエクスポートで保存してください。', 'error');
      }
    }
    updateStorageBadge();
  }

  function exportJson() {
    const payload = {
      schemaVersion: state.schemaVersion,
      tasks: state.tasks,
      holidays: state.holidays,
      manualOrder: state.manualOrder
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    downloadBlob(blob, `gantt-${toISO(dayjs())}.json`);
  }

  function resetGantt() {
    const confirmed = window.confirm('ガントチャートをリセットしますか？');
    if (!confirmed) return;
    state.schemaVersion = 1;
    state.tasks = [];
    state.holidays = [];
    state.manualOrder = [];
    ensureHolidaySet();

    uiState.selectedId = null;
    uiState.filter = { text: '', category: '', assignee: '' };
    uiState.sort = { key: 'manual', dir: 'asc' };

    if (els.filterText) els.filterText.value = '';
    if (els.filterCategory) els.filterCategory.value = '';
    if (els.filterAssignee) els.filterAssignee.value = '';

    if (els.detailPanel) els.detailPanel.classList.remove('open');
    closeMermaidModal();
    closeHolidayModal();

    history.stack = [];
    history.index = -1;
    history.locked = false;
    pushHistory();
    scheduleSave();
    render();
    showToast('ガントチャートをリセットしました');
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function parseMermaidDate(value, format) {
    if (!value) return null;
    const trimmed = value.trim();
    if (!trimmed) return null;
    let parsed = dayjs(trimmed, format || DATE_FMT, true);
    if (!parsed.isValid()) {
      parsed = dayjs(trimmed, DATE_FMT, true);
    }
    return parsed.isValid() ? parsed.startOf('day') : null;
  }

  function parseMermaidDuration(value) {
    if (!value) return null;
    const match = value.trim().match(/^(\d+)\s*([dw])$/i);
    if (!match) return null;
    const amount = Number(match[1]);
    if (!Number.isFinite(amount) || amount < 0) return null;
    const unit = match[2].toLowerCase();
    return unit === 'w' ? amount * 5 : amount;
  }

  function isMermaidTag(value) {
    const lower = String(value).toLowerCase();
    return lower === 'crit';
  }

  function parseMermaidTaskLine(line, dateFormat, groupKey, currentSection, autoIndex) {
    const colonIndex = line.indexOf(':');
    if (colonIndex === -1) return null;
    const title = line.slice(0, colonIndex).trim();
    const rest = line.slice(colonIndex + 1).trim();
    if (!rest) return null;

    const parts = rest
      .split(',')
      .map((part) => part.trim())
      .filter(Boolean);

    let mermaidId = null;
    let afterIds = [];
    const dateTokens = [];
    let duration = null;
    let isMilestone = false;
    let status = TASK_STATUS.NONE;

    parts.forEach((part) => {
      if (!part) return;
      const lower = part.toLowerCase();
      if (lower.startsWith('after ')) {
        const refs = part.slice(6).trim().split(/\s+/).filter(Boolean);
        afterIds = afterIds.concat(refs);
        return;
      }
      if (lower === 'milestone') {
        isMilestone = true;
        return;
      }
      if (lower === 'done') {
        status = TASK_STATUS.DONE;
        return;
      }
      if (lower === 'active' && status !== TASK_STATUS.DONE) {
        status = TASK_STATUS.ACTIVE;
        return;
      }
      if (isMermaidTag(part)) return;
      const dur = parseMermaidDuration(part);
      if (dur !== null) {
        if (duration === null) duration = dur;
        return;
      }
      const date = parseMermaidDate(part, dateFormat);
      if (date) {
        dateTokens.push(date);
        return;
      }
      if (!mermaidId) {
        mermaidId = part;
      }
    });

    if (!mermaidId) {
      mermaidId = `auto${autoIndex}`;
      autoIndex += 1;
    }

    const task = {
      id: generateId(),
      title: title || 'Untitled',
      category: '',
      assignee: '',
      notes: '',
      dependencies: [],
      isMilestone,
      status
    };

    if (currentSection) {
      task[groupKey] = currentSection;
    }

    if (dateTokens[0]) {
      task.start = toISO(dateTokens[0]);
    }
    if (dateTokens[1]) {
      task.end = toISO(dateTokens[1]);
    }
    if (duration !== null) {
      task.duration = duration;
    }
    if (task.isMilestone) {
      task.duration = 1;
    }
    if (!task.start) {
      task.start = toISO(adjustToWorkday(dayjs(), 1));
    }
    if (!task.duration && task.end) {
      task.duration = workdayCount(task.start, task.end);
    } else if (!task.duration) {
      task.duration = 1;
    }

    return { task, mermaidId, afterIds, nextIndex: autoIndex };
  }

  function parseMermaidGantt(text) {
    const lines = text.split(/\r?\n/);
    const hasGantt = lines.some((line) => line.trim().toLowerCase().startsWith('gantt'));
    let inGantt = !hasGantt;
    let dateFormat = DATE_FMT;
    let currentSection = '';
    let autoIndex = 1;
    const groupKey = (els.mermaidGroup && els.mermaidGroup.value) || 'category';
    const tasks = [];
    const holidays = new Set();
    const mermaidIdMap = new Map();
    const pending = [];

    lines.forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line) return;
      if (line.startsWith('```') || line.startsWith('%%')) return;
      if (line.toLowerCase().startsWith('gantt')) {
        inGantt = true;
        return;
      }
      if (!inGantt) return;

      const lower = line.toLowerCase();
      if (lower.startsWith('dateformat')) {
        const fmt = line.replace(/dateformat/i, '').trim();
        if (fmt) dateFormat = fmt;
        return;
      }
      if (lower.startsWith('excludes')) {
        const rest = line.replace(/excludes/i, '').trim();
        if (rest) {
          rest.split(',').forEach((token) => {
            const value = token.trim();
            if (!value) return;
            if (/weekends?/i.test(value)) return;
            const parsed = parseMermaidDate(value, dateFormat);
            if (parsed) holidays.add(toISO(parsed));
          });
        }
        return;
      }
      if (lower.startsWith('section')) {
        currentSection = line.replace(/section/i, '').trim();
        return;
      }

      const parsed = parseMermaidTaskLine(line, dateFormat, groupKey, currentSection, autoIndex);
      if (!parsed) return;
      autoIndex = parsed.nextIndex;
      tasks.push(parsed.task);
      if (parsed.mermaidId) {
        mermaidIdMap.set(parsed.mermaidId, parsed.task);
      }
      pending.push({ task: parsed.task, afterIds: parsed.afterIds });
    });

    pending.forEach((item) => {
      const deps = [];
      item.afterIds.forEach((depId) => {
        const depTask = mermaidIdMap.get(depId);
        if (depTask) deps.push({ id: depTask.id, type: 'FS' });
      });
      item.task.dependencies = deps;
    });

    return { tasks, holidays: Array.from(holidays).sort() };
  }

  async function handleJsonImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    const text = await file.text();
    try {
      const data = JSON.parse(text);
      if (!data || typeof data !== 'object') throw new Error('invalid');
      state.schemaVersion = data.schemaVersion || 1;
      state.tasks = Array.isArray(data.tasks) ? data.tasks.map(normalizeTask) : [];
      state.holidays = Array.isArray(data.holidays) ? data.holidays : [];
      state.manualOrder = Array.isArray(data.manualOrder) ? data.manualOrder : [];
      ensureHolidaySet();
      syncManualOrder();
      scheduleAll();
      commitStateChange();
      showToast('JSONを読み込んでガントチャートを開きました');
    } catch (error) {
      showToast('JSONの読み込みに失敗しました', 'error');
    }
    event.target.value = '';
  }

  async function handleMermaidImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    const text = await file.text();
    try {
      const result = parseMermaidGantt(text);
      if (!result.tasks.length) throw new Error('empty');
      if (hasCycle(result.tasks)) {
        showToast('循環依存のため読み込みできません', 'error');
        event.target.value = '';
        return;
      }
      state.schemaVersion = 1;
      state.tasks = result.tasks.map(normalizeTask);
      state.holidays = Array.isArray(result.holidays) ? result.holidays : [];
      state.manualOrder = state.tasks.map((task) => task.id);
      ensureHolidaySet();
      syncManualOrder();
      scheduleAll();
      commitStateChange({ skipSchedule: true });
      closeMermaidModal();
      showToast('Mermaid を読み込んでガントチャートを開きました');
    } catch (error) {
      showToast('Mermaid の読み込みに失敗しました', 'error');
    }
    event.target.value = '';
  }

  async function handleHolidayImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    const text = await file.text();
    try {
      const data = JSON.parse(text);
      let holidays = [];
      if (Array.isArray(data)) {
        holidays = data;
      } else if (Array.isArray(data.holidays)) {
        holidays = data.holidays;
      }
      holidays = holidays
        .map((d) => toISO(parseDate(d)))
        .filter((d) => d && d.length === 10);
      state.holidays = Array.from(new Set(holidays)).sort();
      ensureHolidaySet();
      scheduleAll();
      commitStateChange();
      renderHolidayList();
      showToast('祝日を更新しました');
    } catch (error) {
      showToast('JSONファイルの読み込みに失敗しました', 'error');
    }
    event.target.value = '';
  }

  function renderHolidayList() {
    if (!els.holidayList) return;
    const holidays = [...state.holidays].sort();
    if (holidays.length === 0) {
      els.holidayList.innerHTML = '<div class="muted">祝日が未設定です</div>';
      return;
    }
    const frag = document.createDocumentFragment();
    holidays.forEach((date) => {
      const row = document.createElement('div');
      row.className = 'holiday-row';
      const label = document.createElement('div');
      label.textContent = date;
      const removeBtn = document.createElement('button');
      removeBtn.className = 'btn ghost';
      removeBtn.innerHTML = '<span class="icon">&#xE74D;</span>';
      removeBtn.addEventListener('click', () => removeHoliday(date));
      row.appendChild(label);
      row.appendChild(removeBtn);
      frag.appendChild(row);
    });
    els.holidayList.innerHTML = '';
    els.holidayList.appendChild(frag);
  }

  function addHolidayFromInput() {
    if (!els.holidayDateInput) return;
    if (!els.holidayDateInput.value) {
      showToast('日付を入力してください', 'error');
      return;
    }
    const date = toISO(parseDate(els.holidayDateInput.value));
    if (!state.holidays.includes(date)) {
      state.holidays.push(date);
      state.holidays.sort();
      ensureHolidaySet();
      scheduleAll();
      commitStateChange();
      renderHolidayList();
    }
    els.holidayDateInput.value = '';
  }

  function removeHoliday(date) {
    const next = state.holidays.filter((d) => d !== date);
    if (next.length === state.holidays.length) return;
    state.holidays = next;
    ensureHolidaySet();
    scheduleAll();
    commitStateChange();
    renderHolidayList();
  }

  function exportHolidays() {
    const payload = {
      schemaVersion: state.schemaVersion,
      holidays: [...state.holidays].sort()
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    downloadBlob(blob, `holidays-${toISO(dayjs())}.json`);
  }

  function openMermaidModal() {
    renderMermaid();
    els.overlay.classList.remove('hidden');
    els.mermaidModal.classList.remove('hidden');
  }

  function closeMermaidModal() {
    els.overlay.classList.add('hidden');
    els.mermaidModal.classList.add('hidden');
  }

  function openHolidayModal() {
    if (!els.holidayModal) return;
    renderHolidayList();
    els.overlay.classList.remove('hidden');
    els.holidayModal.classList.remove('hidden');
  }

  function closeHolidayModal() {
    if (!els.holidayModal) return;
    els.overlay.classList.add('hidden');
    els.holidayModal.classList.add('hidden');
  }

  function closeAllModals() {
    closeMermaidModal();
    closeHolidayModal();
  }

  function renderMermaid() {
    const groupKey = els.mermaidGroup.value;
    const dateMode = els.mermaidDateMode ? els.mermaidDateMode.value : 'duration';
    const grouped = {};
    state.tasks.forEach((task) => {
      const key = task[groupKey] || '未分類';
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(task);
    });

    const idMap = new Map();
    state.tasks.forEach((task, index) => {
      idMap.set(task.id, `t${index + 1}`);
    });

    const lines = [];
    lines.push('gantt');
    lines.push('  title Project');
    lines.push('  dateFormat  YYYY-MM-DD');

    const excludes = ['weekends'];
    state.holidays.forEach((d) => excludes.push(d));
    lines.push(`  excludes ${excludes.join(',')}`);

    Object.keys(grouped).forEach((group) => {
      lines.push(`  section ${group}`);
      grouped[group].forEach((task) => {
        const deps = task.dependencies
          .filter((d) => d.type === 'FS')
          .map((d) => idMap.get(d.id))
          .filter(Boolean);
        const afterText = deps.length ? `after ${deps.join(' ')}` : '';
        let timing = '';
        const isMilestone = Boolean(task.isMilestone);
        const status = normalizeTaskStatus(task.status);
        if (deps.length) {
          if (isMilestone) {
            timing = dateMode === 'end' ? task.end : '1d';
          } else {
            timing = dateMode === 'end' ? task.end : `${task.duration}d`;
          }
        } else if (isMilestone) {
          timing = dateMode === 'end' ? `${task.start}, ${task.end}` : `${task.start}, 1d`;
        } else {
          timing = dateMode === 'end' ? `${task.start}, ${task.end}` : `${task.start}, ${task.duration}d`;
        }
        const tags = [];
        if (status === TASK_STATUS.ACTIVE) tags.push('active');
        if (status === TASK_STATUS.DONE) tags.push('done');
        if (isMilestone) tags.push('milestone');
        const headParts = [...tags, idMap.get(task.id)].filter(Boolean);
        const parts = [`  ${task.title} :${headParts.join(', ')}`];
        if (afterText) parts.push(afterText);
        if (timing) parts.push(timing);
        lines.push(parts.join(', '));
      });
    });

    els.mermaidOutput.value = lines.join('\n');
  }

  function copyMermaid() {
    els.mermaidOutput.select();
    document.execCommand('copy');
    showToast('コピーしました。');
  }

  function render() {
    ensureHolidaySet();
    renderFilters();
    renderTable();
    syncRowHeights();
    renderGantt();
    updateStorageBadge();
    syncLayout();
  }

  function toggleLeftPane() {
    const next = !uiState.leftCollapsed;
    if (next && uiState.rightCollapsed) {
      showToast('左右ペインを同時に非表示にはできません', 'error');
      return;
    }
    uiState.leftCollapsed = next;
    applyPaneState();
  }

  function toggleRightPane() {
    const next = !uiState.rightCollapsed;
    if (next && uiState.leftCollapsed) {
      showToast('左右ペインを同時に非表示にはできません', 'error');
      return;
    }
    uiState.rightCollapsed = next;
    applyPaneState();
  }

  function applyPaneState() {
    if (!els.app) return;
    els.app.classList.toggle('left-collapsed', uiState.leftCollapsed);
    els.app.classList.toggle('right-collapsed', uiState.rightCollapsed);
    if (els.toggleLeftBtn) {
      els.toggleLeftBtn.setAttribute('aria-pressed', uiState.leftCollapsed ? 'true' : 'false');
      els.toggleLeftBtn.disabled = uiState.rightCollapsed;
    }
    if (els.toggleRightBtn) {
      els.toggleRightBtn.setAttribute('aria-pressed', uiState.rightCollapsed ? 'true' : 'false');
      els.toggleRightBtn.disabled = uiState.leftCollapsed;
    }
    if (els.paneSplitter) {
      els.paneSplitter.setAttribute('aria-disabled', uiState.leftCollapsed || uiState.rightCollapsed ? 'true' : 'false');
    }
    applyPaneWidth();
    syncLayout();
  }

  function isEditableTarget(target) {
    if (!target) return false;
    const tag = target.tagName ? target.tagName.toLowerCase() : '';
    return tag === 'input' || tag === 'textarea' || tag === 'select' || target.isContentEditable;
  }

  function handleShortcuts(event) {
    if (isEditableTarget(event.target)) return false;
    const key = event.key.toLowerCase();
    if (event.ctrlKey && !event.shiftKey && key === 'n') {
      event.preventDefault();
      addTask();
      return true;
    }
    if (event.ctrlKey && !event.shiftKey && key === 'f') {
      event.preventDefault();
      els.filterText.focus();
      return true;
    }
    if (event.ctrlKey && !event.shiftKey && key === 'm') {
      event.preventDefault();
      openMermaidModal();
      return true;
    }
    if (event.ctrlKey && !event.shiftKey && key === 's') {
      event.preventDefault();
      exportJson();
      return true;
    }
    if (key === 'delete' || key === 'backspace') {
      if (uiState.selectedId) {
        event.preventDefault();
        deleteTask();
        return true;
      }
    }
    if (key === 'escape') {
      if (els.mermaidModal && !els.mermaidModal.classList.contains('hidden')) {
        closeMermaidModal();
        return true;
      }
      if (els.holidayModal && !els.holidayModal.classList.contains('hidden')) {
        closeHolidayModal();
        return true;
      }
      if (els.detailPanel && els.detailPanel.classList.contains('open')) {
        closeDetail();
        return true;
      }
    }
    return false;
  }

  document.addEventListener('DOMContentLoaded', () => {
    initElements();
    setupEvents();
    applyPaneState();
    loadState();
    syncLayout();
  });
})();
