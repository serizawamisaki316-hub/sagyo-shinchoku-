// =========================================================
// 倉庫作業進捗サイネージ - フロントエンド制御 (app.js)
// 【時間テキスト発光点滅警告 ＆ クリック一時停止 ＆ 5秒自動同期】
// =========================================================

(function () {
  'use strict';

  let config = {
    excel_path: '',
    poll_interval_sec: 5,
    scroll_speed_px_per_sec: 35,
    bottom_pause_sec: 4,
    top_pause_sec: 2,
    font_size_scale: 1.0
  };

  let isScrolling = true;
  let currentScrollY = 0;
  let lastTimestamp = 0;
  let pauseTimer = null;
  let lastRenderedDataHash = '';
  let latestSignageData = null;

  // DOM Elements
  const viewport = document.getElementById('scroll-viewport');
  const scrollContent = document.getElementById('scroll-content');
  const courseCardsContainer = document.getElementById('course-cards-container');
  const sheetTitle = document.getElementById('sheet-title');
  const courseCountBadge = document.getElementById('course-count-badge');
  const clockDate = document.getElementById('clock-date');
  const clockTime = document.getElementById('clock-time');
  const syncStatusText = document.getElementById('sync-status-text');
  const syncStatusIcon = document.getElementById('sync-status-icon');
  const btnToggleScroll = document.getElementById('btn-toggle-scroll');
  const scrollIcon = document.getElementById('scroll-icon');
  const scrollBtnText = document.getElementById('scroll-btn-text');
  const btnSpeed = document.getElementById('btn-speed');
  const speedLabel = document.getElementById('speed-label');
  const btnFullscreen = document.getElementById('btn-fullscreen');
  const btnSettings = document.getElementById('btn-settings');
  const errorBanner = document.getElementById('error-banner');
  const errorMsg = document.getElementById('error-msg');
  const loopNoticeBar = document.getElementById('loop-notice-bar');

  // Modal Elements
  const settingsModal = document.getElementById('settings-modal');
  const btnCloseModal = document.getElementById('btn-close-modal');
  const btnCancelSettings = document.getElementById('btn-cancel-settings');
  const btnSaveSettings = document.getElementById('btn-save-settings');

  // Table Header Course Columns
  const thCourseMain = document.getElementById('th-course-main');
  const thCourseSub = document.getElementById('th-course-sub');
  const inputExcelPath = document.getElementById('input-excel-path');
  const inputPollInterval = document.getElementById('input-poll-interval');
  const inputScrollSpeed = document.getElementById('input-scroll-speed');
  const inputBottomPause = document.getElementById('input-bottom-pause');
  const inputTopPause = document.getElementById('input-top-pause');
  const inputFontScale = document.getElementById('input-font-scale');

  // Floating Pause Toast
  const pauseToast = document.createElement('div');
  pauseToast.className = 'pause-toast';
  pauseToast.id = 'pause-toast';
  pauseToast.style.display = 'none';
  pauseToast.innerHTML = '<span>⏸️ 一時停止中（マウスで自由にスクロール可・画面クリックで再開）</span>';
  document.body.appendChild(pauseToast);

  // 1. Digital Clock & Real-time Warning Evaluation
  const DAYS_JP = ['日', '月', '火', '水', '木', '金', '土'];
  function updateClock() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    const day = DAYS_JP[now.getDay()];
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');

    clockDate.textContent = `${y}/${m}/${d} (${day})`;
    clockTime.textContent = `${hh}:${mm}:${ss}`;

    if (latestSignageData) {
      updateTimeWarningClasses();
    }
  }
  setInterval(updateClock, 500);
  updateClock();

  // Helper: Check if within 10 minutes or past delivery time
  function isWithin10MinOrPast(timeStr) {
    if (!timeStr || !timeStr.includes(':')) return false;
    const parts = timeStr.split(':');
    const targetH = parseInt(parts[0], 10);
    const targetM = parseInt(parts[1], 10);
    if (isNaN(targetH) || isNaN(targetM)) return false;

    const now = new Date();
    const currentTotalMin = now.getHours() * 60 + now.getMinutes();
    const targetTotalMin = targetH * 60 + targetM;

    const diffMinutes = targetTotalMin - currentTotalMin;
    return diffMinutes <= 10;
  }

  function updateTimeWarningClasses() {
    const timeCells = document.querySelectorAll('[data-time-val]');
    timeCells.forEach(cell => {
      const timeStr = cell.getAttribute('data-time-val');
      const isCompleted = cell.getAttribute('data-is-completed') === 'true';
      const span = cell.querySelector('.time-val');

      const shouldWarn = !isCompleted && isWithin10MinOrPast(timeStr);
      if (span) {
        if (shouldWarn) {
          span.classList.add('time-val-warning');
        } else {
          span.classList.remove('time-val-warning');
        }
      }
    });
  }

  // Day of week management & URL sync
  function getInitialDay() {
    const params = new URLSearchParams(window.location.search);
    const dayParam = params.get('day');
    if (dayParam) {
      const clean = dayParam.trim().toLowerCase();
      if (['平日', 'heijitsu', 'weekday', '水', '木', '金', '土'].includes(clean)) return '平日';
      if (['月曜', '月', 'mon', 'monday'].includes(clean)) return '月曜';
      if (['火曜', '火', 'tue', 'tuesday'].includes(clean)) return '火曜';
      if (['日・祝', '日祝', '日', '祝', 'sun', 'sunday', 'holiday'].includes(clean)) return '日・祝';
    }
    const dayNum = new Date().getDay(); // 0:Sun, 1:Mon, 2:Tue, 3:Wed, 4:Thu, 5:Fri, 6:Sat
    if (dayNum === 0) return '日・祝';
    if (dayNum === 1) return '月曜';
    if (dayNum === 2) return '火曜';
    return '平日';
  }

  // 閲覧専用（見るだけ）モードの判定 (?view=1 または ?readonly=1 または ?viewonly=1 または file:プロトコル)
  const isFileProtocol = window.location.protocol === 'file:' || !window.location.protocol.startsWith('http');
  const urlParams = new URLSearchParams(window.location.search);
  const isViewOnly = isFileProtocol || urlParams.has('view') || urlParams.has('readonly') || urlParams.has('viewonly');
  if (isViewOnly && btnSettings) {
    btnSettings.style.display = 'none';
  }

  let useStaticScriptMode = isFileProtocol || (typeof window.__ALL_SIGNAGE_DATA__ !== 'undefined');
  let currentSelectedDay = getInitialDay();

  function updateHeaderTitles(day) {
    if (!thCourseMain || !thCourseSub) return;
    const clean = (day || currentSelectedDay || '平日').replace(/[（\(\)）]/g, '').trim();
    if (clean === '月曜' || clean === '月') {
      thCourseMain.textContent = '月曜コース';
      thCourseSub.textContent = '平日コース';
    } else if (clean === '火曜' || clean === '火') {
      thCourseMain.textContent = '火曜コース';
      thCourseSub.textContent = '平日コース';
    } else if (clean === '日・祝' || clean === '日祝' || clean === '日') {
      thCourseMain.textContent = '日祝コース';
      thCourseSub.textContent = '平日コース';
    } else {
      thCourseMain.textContent = '平日コース';
      thCourseSub.textContent = '平日コース';
    }
  }

  function updateDayTabsUi(activeDay) {
    const tabs = document.querySelectorAll('.day-tab');
    tabs.forEach(t => {
      if (t.getAttribute('data-day') === activeDay) {
        t.classList.add('active');
      } else {
        t.classList.remove('active');
      }
    });
    updateHeaderTitles(activeDay);
  }

  // Initial tab highlight and header titles
  updateDayTabsUi(currentSelectedDay);

  // Helper: Reload signage data dynamically across all environments (file://, http://, X:, OneDrive)
  async function loadDataFromSharedStorage() {
    // 1. First Priority: Try Fetching JSON with Cache-Busting (fastest & cleanest)
    try {
      const url = 'signage_data.json' + (isFileProtocol ? '' : ('?_t=' + Date.now()));
      const resp = await fetch(url, { cache: 'no-store' });
      if (resp.ok) {
        const json = await resp.json();
        if (json && json.days && Object.keys(json.days).length > 0) {
          window.__ALL_SIGNAGE_DATA__ = json;
          return json;
        }
      }
    } catch (e) {}

    // 2. Second Priority: Try XMLHttpRequest for JSON (Supported locally in file:/// on Chromium)
    try {
      const xhrJson = await new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', 'signage_data.json', true);
        xhr.onload = () => {
          if (xhr.status === 200 || xhr.status === 0) {
            try {
              const res = JSON.parse(xhr.responseText);
              resolve(res);
            } catch (err) { reject(err); }
          } else {
            reject(new Error('XHR status ' + xhr.status));
          }
        };
        xhr.onerror = reject;
        xhr.send();
      });
      if (xhrJson && xhrJson.days && Object.keys(xhrJson.days).length > 0) {
        window.__ALL_SIGNAGE_DATA__ = xhrJson;
        return xhrJson;
      }
    } catch (e) {}

    // 3. Third Priority: Dynamic Script Element Loading with cache buster on HTTP or direct script on file://
    return new Promise((resolve) => {
      const scriptId = 'dynamic_signage_data_script';
      const existing = document.getElementById(scriptId);
      if (existing) {
        existing.remove();
      }
      const script = document.createElement('script');
      script.id = scriptId;
      // On file:// do not append query string to prevent ERR_FILE_NOT_FOUND on Windows
      script.src = isFileProtocol ? 'signage_data.js' : ('signage_data.js?_t=' + Date.now());

      let timer = setTimeout(() => {
        resolve(window.__ALL_SIGNAGE_DATA__ || null);
      }, 1200);

      script.onload = () => {
        clearTimeout(timer);
        resolve(window.__ALL_SIGNAGE_DATA__ || null);
      };
      script.onerror = () => {
        clearTimeout(timer);
        resolve(window.__ALL_SIGNAGE_DATA__ || null);
      };
      document.head.appendChild(script);
    });
  }

  // 2. Guaranteed Real-Time Polling Engine (Dual Mode: Server HTTP / Shared File Script)
  async function fetchSignageData() {
    try {
      const nowTimeStr = new Date().toTimeString().split(' ')[0];
      let data = null;

      let allData = window.__ALL_SIGNAGE_DATA__;

      if (!useStaticScriptMode) {
        try {
          const resp = await fetch('/api/data?day=' + encodeURIComponent(currentSelectedDay) + '&_t=' + Date.now(), {
            cache: 'no-store',
            headers: { 'Pragma': 'no-cache', 'Cache-Control': 'no-cache' }
          });
          if (!resp.ok) throw new Error(`HTTP Error ${resp.status}`);
          data = await resp.json();
        } catch (fetchErr) {
          console.warn('Web server fetch failed, switching to shared script mode...', fetchErr);
          useStaticScriptMode = true;
        }
      }

      if (useStaticScriptMode) {
        // ALWAYS await loadDataFromSharedStorage to get freshest data on every poll
        try {
          const reloaded = await loadDataFromSharedStorage();
          if (reloaded) {
            allData = reloaded;
          }
        } catch (e) {}

        if (!allData || !allData.days || Object.keys(allData.days).length === 0) {
          throw new Error('進捗データを同期中（データ準備待ち）...');
        }
        const daysMap = allData.days || {};
        data = daysMap[currentSelectedDay] || daysMap['平日'] || Object.values(daysMap)[0];
        if (!data) {
          throw new Error('進捗データが空です');
        }
        if (allData.config) {
          data.config = allData.config;
        }
        data.sharedTimestamp = allData.timestamp;
      }

      if (data.day && data.day !== currentSelectedDay && !useStaticScriptMode) {
        currentSelectedDay = data.day;
        updateDayTabsUi(currentSelectedDay);
      }

      if (!data.success) {
        showError(data.error || 'データ読込エラー');
        syncStatusIcon.textContent = '🔴';
        syncStatusText.textContent = `読込エラー (${nowTimeStr})`;
      } else {
        hideError();
        syncStatusIcon.textContent = '🟢';
        const fileTime = data.last_modified ? data.last_modified.split(' ')[1] : (data.sharedTimestamp ? data.sharedTimestamp.split(' ')[1] : '--:--:--');
        const modeLabel = useStaticScriptMode ? '共有同期' : '同期';
        syncStatusText.textContent = `${modeLabel}: ${nowTimeStr} (${data.excel_file || 'Excel'}: ${fileTime})`;

        if (data.config && !config.excel_path) {
          config = Object.assign({}, config, data.config);
          document.documentElement.style.setProperty('--font-scale', config.font_size_scale || 1.0);
        }

        latestSignageData = data;

        // Smart Render
        const currentDataHash = JSON.stringify(data.courses);
        if (currentDataHash !== lastRenderedDataHash) {
          lastRenderedDataHash = currentDataHash;
          renderGroupedCards(data);
        } else {
          updateTimeWarningClasses();
        }
      }
    } catch (err) {
      console.error('[FETCH ERROR]', err);
      showError(`データ同期エラー: ${err.message}`);
      syncStatusIcon.textContent = '🔴';
      syncStatusText.textContent = '同期エラー';
    }
  }

  async function startPollingLoop() {
    while (true) {
      await fetchSignageData();
      const intervalSec = Math.max(2, config.poll_interval_sec || 5);
      await new Promise(res => setTimeout(res, intervalSec * 1000));
    }
  }

  function renderGroupedCards(data) {
    if (sheetTitle && data.title) {
      const cleanTitle = data.title.replace(/[（\(\)）]/g, '').replace(/表示/g, '').trim();
      sheetTitle.textContent = cleanTitle || '平日';
    }
    if (data.day || data.title) {
      updateHeaderTitles(data.day || data.title);
    }
    if (courseCountBadge && data.count !== undefined) {
      courseCountBadge.textContent = `${data.count} コース`;
    }

    const courses = data.courses || [];
    if (courses.length === 0) {
      courseCardsContainer.innerHTML = `<div style="padding: 40px; text-align: center; color: #64748B;">データがありません</div>`;
      return;
    }

    function cleanNumericText(val) {
      if (val === null || val === undefined) return '';
      let s = String(val).trim();
      if (s.endsWith('.0')) {
        const num = parseFloat(s);
        if (!isNaN(num) && Number.isInteger(num)) {
          s = String(parseInt(s, 10));
        }
      }
      return s;
    }

    // Continuous vehicle grouping (Works for all days: 平日, 月曜, 火曜, 日・祝)
    const groups = [];
    let currentGroup = null;

    courses.forEach(c => {
      const v = cleanNumericText(c.vehicle || '');
      c.vehicle = v;
      c.course = cleanNumericText(c.course || '');

      if (v !== '' && currentGroup && currentGroup.vehicleName === v) {
        currentGroup.courses.push(c);
      } else {
        currentGroup = {
          vehicleName: v,
          courses: [c]
        };
        groups.push(currentGroup);
      }
    });

    let html = '';

    groups.forEach((group) => {
      const numCourses = group.courses.length;
      const totalRows = numCourses * 2;
      const isMultiCourse = numCourses > 1;

      // Group representative time & completion metrics
      let groupTime = '';
      let groupCompletedTime = '';
      let groupDiffMinutes = null;

      for (let i = 0; i < group.courses.length; i++) {
        if (group.courses[i].time && group.courses[i].time.trim()) {
          groupTime = group.courses[i].time.trim();
          break;
        }
      }

      for (let i = 0; i < group.courses.length; i++) {
        if (group.courses[i].group_completed_time) {
          groupCompletedTime = group.courses[i].group_completed_time;
        }
        if (group.courses[i].group_diff_minutes !== null && group.courses[i].group_diff_minutes !== undefined) {
          groupDiffMinutes = group.courses[i].group_diff_minutes;
        }
      }
      if (!groupCompletedTime && group.courses.length > 0) {
        groupCompletedTime = group.courses[0].course_completed_time || '';
      }
      if (groupDiffMinutes === null && group.courses.length > 0) {
        groupDiffMinutes = group.courses[0].course_diff_minutes;
      }

      // Group completion check (どれか1つでも未完了(F列=0)なら警告対象)
      const isGroupAllCompleted = group.courses.every(c => c.is_completed === true);
      const isGroupWarning = !isGroupAllCompleted && isWithin10MinOrPast(groupTime);
      const groupWarningValClass = isGroupWarning ? 'time-val-warning' : '';

      // Helper to build time cell contents (2段ボックス表示)
      const buildTimeCellHtml = (timeStr, isCompleted, compTime, diffMin, warnClass) => {
        let diffHtml = '';
        if (isCompleted && (compTime || diffMin !== null && diffMin !== undefined)) {
          const timeLabel = compTime ? `完了 ${compTime}` : '完了';
          let diffBoxClass = 'diff-box-ontime';
          let diffValText = '定刻 (±0)';

          if (diffMin !== null && diffMin !== undefined) {
            if (diffMin < 0) {
              diffBoxClass = 'diff-box-early';
              diffValText = `${diffMin}分 早着`;
            } else if (diffMin > 0) {
              diffBoxClass = 'diff-box-late';
              diffValText = `+${diffMin}分 遅延`;
            } else {
              diffBoxClass = 'diff-box-ontime';
              diffValText = '定刻 (±0)';
            }
          }

          diffHtml = `
            <div class="time-diff-2tier ${diffBoxClass}">
              <span class="diff-tier-time">${timeLabel}</span>
              <span class="diff-tier-val">${diffValText}</span>
            </div>
          `;
        }

        return `
          <div class="time-cell-content">
            <span class="time-val ${warnClass}">${timeStr || ''}</span>
            ${diffHtml}
          </div>
        `;
      };

      html += `
        <table class="signage-table vehicle-group-card">
          <colgroup>
            <col class="col-vehicle">
            <col class="col-course-sub">
            <col class="col-time">
            <col class="col-line">
            <col class="col-num" span="25">
            <col class="col-slip">
          </colgroup>
          <tbody>
      `;

      group.courses.forEach((c, courseIdx) => {
        const isFirstCourseInGroup = courseIdx === 0;
        const isLastCourseInGroup = courseIdx === numCourses - 1;
        const sepClass = (!isLastCourseInGroup) ? 'course-separator-row' : '';

        // Helper to format tile
        const isNoDel = c.is_no_delivery;

        const buildTileHtml = (item) => {
          let tileClass = 'tile-unstarted';
          let displayNum = '';

          if (isNoDel === true) {
            tileClass = 'tile-green-nodelivery'; // 配送なし: 振出25件すべてが0 かつ AF列が99
          } else if (item.status === 99 || item.status >= 99) {
            tileClass = 'tile-blue-done';        // 完了済み: 99 (青色)
            displayNum = item.num;
          } else if (item.status === 1) {
            tileClass = 'tile-grey-active';      // 作業中: 1 (グレー)
            displayNum = item.num;
          } else {
            tileClass = 'tile-unstarted';        // 未着手: 0 (白枠)
          }

          return `<td><div class="cell-tile-container"><div class="progress-tile ${tileClass}">${displayNum}</div></div></td>`;
        };

        // 1..25 Tiles for 振出
        const furidashiTiles = (c.furidashi.items || []).map(buildTileHtml).join('');

        // 1..25 Tiles for 査照
        const sagyoTiles = (c.sagyo.items || []).map(buildTileHtml).join('');

        // Slip Checkbox (伝票: 配送なし優先で薄緑、完了済みブルー背景 + チェック、未着手は白枠)
        const slipChecked = c.slip && c.slip.is_done;
        let slipBoxHtml = '';
        if (isNoDel) {
          slipBoxHtml = '<div class="slip-checkbox slip-nodelivery-green"></div>';
        } else if (slipChecked) {
          slipBoxHtml = '<div class="slip-checkbox slip-done-blue">✓</div>';
        } else {
          slipBoxHtml = '<div class="slip-checkbox slip-pending-empty"></div>';
        }

        // Course Columns & Time Column HTML
        let leftColsHtml = '';

        const groupCompletedClass = isGroupAllCompleted ? 'group-completed-cell' : '';

        if (isMultiCourse) {
          if (isFirstCourseInGroup) {
            const timeInnerHtml = buildTimeCellHtml(groupTime, isGroupAllCompleted, groupCompletedTime, groupDiffMinutes, groupWarningValClass);
            // Vehicle Plate & Time Cell
            leftColsHtml += `
              <td class="cell-vehicle-tall ${groupCompletedClass}" rowspan="${totalRows}">
                <div class="badge-vehicle-tall ${isGroupAllCompleted ? 'badge-completed' : ''}"><span class="badge-text-inner">${group.vehicleName}</span></div>
              </td>
              <td class="cell-course-sub ${c.is_completed ? 'group-completed-cell' : ''}" rowspan="2">
                <div class="badge-course-sub ${c.is_completed ? 'badge-completed' : ''}"><span class="badge-text-inner">${c.course || '-'}</span></div>
              </td>
              <td class="cell-time-tall ${groupCompletedClass}" rowspan="${totalRows}" data-time-val="${groupTime}" data-is-completed="${isGroupAllCompleted}">
                ${timeInnerHtml}
              </td>
            `;
          } else {
            leftColsHtml += `
              <td class="cell-course-sub ${c.is_completed ? 'group-completed-cell' : ''}" rowspan="2">
                <div class="badge-course-sub ${c.is_completed ? 'badge-completed' : ''}"><span class="badge-text-inner">${c.course || '-'}</span></div>
              </td>
            `;
          }
        } else {
          // Single course
          const isSingleCompleted = c.is_completed === true;
          const isSingleWarning = !isSingleCompleted && isWithin10MinOrPast(c.time);
          const singleWarningValClass = isSingleWarning ? 'time-val-warning' : '';
          const singleCompletedClass = isSingleCompleted ? 'group-completed-cell' : '';

          const singleCompletedTime = c.group_completed_time || c.course_completed_time || '';
          const singleDiffMinutes = (c.group_diff_minutes !== null && c.group_diff_minutes !== undefined)
            ? c.group_diff_minutes
            : c.course_diff_minutes;

          const singleTimeInnerHtml = buildTimeCellHtml(c.time, isSingleCompleted, singleCompletedTime, singleDiffMinutes, singleWarningValClass);

          function isHyphenOrEmpty(str) {
            if (str === null || str === undefined) return true;
            const s = String(str).trim();
            return s === '' || s === '-' || s === 'ー' || s === '―' || s === '‐' || s === '－' || s === 'ｰ';
          }

          const isCourseEmpty = isHyphenOrEmpty(c.course);
          const isVehicleEmpty = isHyphenOrEmpty(group.vehicleName);
          const isSame = !isCourseEmpty && (c.course === group.vehicleName);
          const isMerged = isCourseEmpty || isVehicleEmpty || isSame;

          if (isMerged) {
            const mergedLabel = (!isVehicleEmpty) ? group.vehicleName : ((!isCourseEmpty) ? c.course : '-');
            leftColsHtml += `
              <td class="cell-course-full ${singleCompletedClass}" colspan="2" rowspan="2">
                <div class="badge-course-full ${isSingleCompleted ? 'badge-completed' : ''}"><span class="badge-text-inner">${mergedLabel}</span></div>
              </td>
            `;
          } else {
            leftColsHtml += `
              <td class="cell-vehicle-single ${singleCompletedClass}" rowspan="2">
                <div class="badge-vehicle-single ${isSingleCompleted ? 'badge-completed' : ''}"><span class="badge-text-inner">${group.vehicleName || '-'}</span></div>
              </td>
              <td class="cell-course-sub ${singleCompletedClass}" rowspan="2">
                <div class="badge-course-sub ${isSingleCompleted ? 'badge-completed' : ''}"><span class="badge-text-inner">${c.course || '-'}</span></div>
              </td>
            `;
          }

          leftColsHtml += `
            <td class="cell-time ${singleCompletedClass}" rowspan="2" data-time-val="${c.time || ''}" data-is-completed="${isSingleCompleted}">
              ${singleTimeInnerHtml}
            </td>
          `;
        }

        // Row 1: 振出
        const completedClass = c.is_completed ? 'course-completed-row' : '';
        html += `
          <tr class="course-row-1 ${completedClass}" id="${c.id}_r1">
            ${leftColsHtml}
            <td class="cell-line-furidashi">${c.furidashi.label || '振出'}</td>
            ${furidashiTiles}
            <td class="cell-slip" rowspan="2">
              <div class="slip-checkbox-container">${slipBoxHtml}</div>
            </td>
          </tr>
        `;

        // Row 2: 査照
        html += `
          <tr class="course-row-2 ${sepClass} ${completedClass}" id="${c.id}_r2">
            <td class="cell-line-sagyo">${c.sagyo.label || '査照'}</td>
            ${sagyoTiles}
          </tr>
        `;
      });

      html += `
          </tbody>
        </table>
      `;
    });

    const prevScrollTop = viewport ? viewport.scrollTop : 0;
    const prevIsScrolling = isScrolling;

    courseCardsContainer.innerHTML = html;
    fitBadgeFontSizes();

    // 1. Synchronous immediate scroll restoration
    if (viewport && prevScrollTop > 0) {
      viewport.scrollTop = prevScrollTop;
      currentScrollY = prevScrollTop;
    }

    // 2. Secondary restoration on animation frame (protects against browser layout reflow)
    requestAnimationFrame(() => {
      if (viewport && prevScrollTop > 0) {
        viewport.scrollTop = prevScrollTop;
        currentScrollY = prevScrollTop;
      }
      if (!prevIsScrolling) {
        isScrolling = false;
        if (pauseToast) pauseToast.style.display = 'flex';
        if (scrollIcon) scrollIcon.textContent = '▶️';
        if (scrollBtnText) scrollBtnText.textContent = '停止中';
      }
    });
  }

  function fitBadgeFontSizes() {
    requestAnimationFrame(() => {
      const badges = document.querySelectorAll('.badge-vehicle-tall, .badge-course-sub, .badge-vehicle-single, .badge-course-full');
      badges.forEach(badge => {
        const inner = badge.querySelector('.badge-text-inner') || badge;
        inner.style.transform = 'none';
        inner.style.whiteSpace = 'nowrap';
        inner.style.display = 'inline-block';
        inner.style.width = 'max-content';
        inner.style.maxWidth = 'none';

        // 利用可能幅（左右パディングの余白を引いた幅）
        const paddingOffset = 10;
        const availableW = badge.clientWidth - paddingOffset;
        const textW = inner.offsetWidth || inner.scrollWidth;

        if (availableW > 0 && textW > availableW) {
          const ratio = Math.max(0.35, availableW / textW);
          inner.style.transformOrigin = 'center center';
          inner.style.transform = `scale(${ratio.toFixed(4)})`;
        } else {
          inner.style.transform = 'none';
        }
      });
    });
  }

  window.addEventListener('resize', fitBadgeFontSizes);
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(fitBadgeFontSizes);
  }

  function showError(msg) {
    errorMsg.textContent = msg;
    errorBanner.style.display = 'flex';
  }

  function hideError() {
    errorBanner.style.display = 'none';
  }

  // 3. Smooth Auto-Scroll Engine with Click Toggle & Free Mouse Scroll
  function getScrollMetrics() {
    const scrollH = viewport.scrollHeight;
    const viewH = viewport.clientHeight;
    const maxScroll = Math.max(0, scrollH - viewH);
    return { scrollH, viewH, maxScroll };
  }

  function scrollStep(timestamp) {
    if (!lastTimestamp) lastTimestamp = timestamp;
    const delta = (timestamp - lastTimestamp) / 1000;
    lastTimestamp = timestamp;

    if (isScrolling) {
      const { maxScroll } = getScrollMetrics();

      if (maxScroll > 0) {
        const speed = config.scroll_speed_px_per_sec || 35;
        currentScrollY += speed * delta;

        if (currentScrollY >= maxScroll) {
          currentScrollY = maxScroll;
          viewport.scrollTop = maxScroll;
          isScrolling = false;

          pauseTimer = setTimeout(() => {
            currentScrollY = 0;
            viewport.scrollTo({ top: 0, behavior: 'smooth' });

            pauseTimer = setTimeout(() => {
              isScrolling = true;
              lastTimestamp = performance.now();
              requestAnimationFrame(scrollStep);
            }, (config.top_pause_sec || 2) * 1000);

          }, (config.bottom_pause_sec || 4) * 1000);

          return;
        }

        viewport.scrollTop = currentScrollY;
      }
    }

    if (isScrolling) {
      requestAnimationFrame(scrollStep);
    }
  }

  function startScrolling() {
    if (pauseTimer) clearTimeout(pauseTimer);
    isScrolling = true;
    currentScrollY = viewport.scrollTop;
    scrollIcon.textContent = '⏸️';
    scrollBtnText.textContent = 'スクロール中';
    pauseToast.style.display = 'none';
    lastTimestamp = performance.now();
    requestAnimationFrame(scrollStep);
  }

  function pauseScrolling() {
    if (pauseTimer) clearTimeout(pauseTimer);
    isScrolling = false;
    currentScrollY = viewport.scrollTop;
    scrollIcon.textContent = '▶️';
    scrollBtnText.textContent = '停止中';
    pauseToast.style.display = 'flex';
  }

  function toggleScrolling() {
    if (isScrolling) {
      pauseScrolling();
    } else {
      startScrolling();
    }
  }

  viewport.addEventListener('click', (e) => {
    if (e.target.closest('button') || e.target.closest('.modal-overlay') || e.target.closest('input')) {
      return;
    }
    toggleScrolling();
  });

  viewport.addEventListener('wheel', () => {
    currentScrollY = viewport.scrollTop;
  }, { passive: true });

  viewport.addEventListener('scroll', () => {
    if (!isScrolling) {
      currentScrollY = viewport.scrollTop;
    }
  }, { passive: true });

  // 4. Speed & Controls
  const SPEEDS = [
    { label: '低速', speed: 20 },
    { label: '標準', speed: 35 },
    { label: '高速', speed: 55 }
  ];
  let currentSpeedIndex = 1;

  function cycleSpeed() {
    currentSpeedIndex = (currentSpeedIndex + 1) % SPEEDS.length;
    const selected = SPEEDS[currentSpeedIndex];
    config.scroll_speed_px_per_sec = selected.speed;
    speedLabel.textContent = selected.label;
  }

  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => {
        console.warn('Fullscreen error:', err);
      });
      btnFullscreen.textContent = '🗗 解除';
    } else {
      document.exitFullscreen();
      btnFullscreen.textContent = '⛶ 全画面';
    }
  }

  // 5. Config & Settings Modal
  function openSettings() {
    inputExcelPath.value = config.excel_path || '';
    inputPollInterval.value = config.poll_interval_sec || 5;
    inputScrollSpeed.value = config.scroll_speed_px_per_sec || 35;
    inputBottomPause.value = config.bottom_pause_sec || 4;
    inputTopPause.value = config.top_pause_sec || 2;
    inputFontScale.value = String(config.font_size_scale || 1.0);
    settingsModal.style.display = 'flex';
  }

  function closeSettings() {
    settingsModal.style.display = 'none';
  }

  async function saveSettings() {
    const updated = {
      excel_path: inputExcelPath.value.trim(),
      poll_interval_sec: parseInt(inputPollInterval.value, 10) || 5,
      scroll_speed_px_per_sec: parseInt(inputScrollSpeed.value, 10) || 35,
      bottom_pause_sec: parseInt(inputBottomPause.value, 10) || 4,
      top_pause_sec: parseInt(inputTopPause.value, 10) || 2,
      font_size_scale: parseFloat(inputFontScale.value) || 1.0
    };

    const isHttp = window.location.protocol.startsWith('http');

    if (isHttp) {
      try {
        const resp = await fetch('/api/config', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(updated)
        });
        if (resp.ok) {
          config = Object.assign({}, config, updated);
          document.documentElement.style.setProperty('--font-scale', config.font_size_scale || 1.0);
          closeSettings();
          fetchSignageData();
          return;
        }
      } catch (e) {
        console.warn('HTTP config save failed:', e);
      }
    }

    // Fallback for file:/// mode: apply and save locally
    try {
      localStorage.setItem('local_signage_config', JSON.stringify(updated));
      config = Object.assign({}, config, updated);
      document.documentElement.style.setProperty('--font-scale', config.font_size_scale || 1.0);
      closeSettings();
    } catch (e) {
      console.warn('LocalStorage config save failed:', e);
    }
  }

  // Event Listeners
  const dayTabsGroup = document.getElementById('day-tabs-group');
  if (dayTabsGroup) {
    dayTabsGroup.addEventListener('click', (e) => {
      const tabBtn = e.target.closest('.day-tab');
      if (!tabBtn) return;
      e.stopPropagation();
      const chosenDay = tabBtn.getAttribute('data-day');
      if (chosenDay && chosenDay !== currentSelectedDay) {
        currentSelectedDay = chosenDay;
        updateDayTabsUi(currentSelectedDay);
        try {
          const newUrl = new URL(window.location);
          newUrl.searchParams.set('day', currentSelectedDay);
          window.history.replaceState({}, '', newUrl);
        } catch (historyErr) {}
        lastRenderedDataHash = '';
        viewport.scrollTop = 0;
        currentScrollY = 0;
        fetchSignageData();
      }
    });
  }

  if (btnToggleScroll) {
    btnToggleScroll.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleScrolling();
    });
  }
  if (btnSpeed) {
    btnSpeed.addEventListener('click', (e) => {
      e.stopPropagation();
      cycleSpeed();
    });
  }
  if (btnFullscreen) {
    btnFullscreen.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleFullscreen();
    });
  }
  if (btnSettings) {
    btnSettings.addEventListener('click', (e) => {
      e.stopPropagation();
      openSettings();
    });
  }
  if (btnCloseModal) btnCloseModal.addEventListener('click', closeSettings);
  if (btnCancelSettings) btnCancelSettings.addEventListener('click', closeSettings);
  if (btnSaveSettings) btnSaveSettings.addEventListener('click', saveSettings);

  window.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
    if (e.code === 'Space') {
      e.preventDefault();
      toggleScrolling();
    } else if (e.code === 'KeyF') {
      toggleFullscreen();
    }
  });

  // Start Real-Time 5s Polling
  startPollingLoop();

  // Restore saved scroll position & state on auto-refresh
  let initialScrollRestored = false;
  try {
    const savedScroll = sessionStorage.getItem('signage_scroll_top');
    const savedPaused = sessionStorage.getItem('signage_is_paused');
    if (savedScroll !== null) {
      const sY = parseFloat(savedScroll);
      if (!isNaN(sY) && sY > 0) {
        viewport.scrollTop = sY;
        currentScrollY = sY;
        initialScrollRestored = true;
      }
    }
    if (savedPaused === 'true') {
      pauseScrolling();
    } else {
      setTimeout(startScrolling, 1000);
    }
  } catch (e) {
    setTimeout(startScrolling, 1000);
  }

})();
