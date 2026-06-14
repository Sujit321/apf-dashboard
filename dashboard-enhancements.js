// =====================================================================
// DASHBOARD ENHANCEMENTS — APF Dashboard
// Extends app.js without modifying it.
// Loaded AFTER app.js, excel-analytics.js, school-map.js
// =====================================================================

// ── Task 17: Append new encrypted data keys ──────────────────────────
(function _appendEncryptedKeys() {
  if (typeof ENCRYPTED_DATA_KEYS === 'undefined' || !Array.isArray(ENCRYPTED_DATA_KEYS)) return;
  ['reportDrafts', 'parentContacts'].forEach(function (k) {
    if (!ENCRYPTED_DATA_KEYS.includes(k)) ENCRYPTED_DATA_KEYS.push(k);
  });
})();

// ─────────────────────────────────────────────────────────────────────
// TASKS 1–9 · WidgetBuilder Extensions
// ─────────────────────────────────────────────────────────────────────
(function _extendWidgetBuilder() {
  if (typeof WidgetBuilder === 'undefined') {
    console.warn('[APF Enhancements] WidgetBuilder not found – widget extensions skipped.');
    return;
  }

  // ── Task 1: 5 new registry entries ───────────────────────────────
  var newWidgets = [
    {
      id: 'today-schedule',
      title: "Today's Schedule",
      icon: 'fa-calendar-day',
      size: 'medium',
      category: 'Insights',
      desc: 'Visits, trainings & tasks planned for today'
    },
    {
      id: 'motivational-quote',
      title: 'Daily Motivation',
      icon: 'fa-quote-left',
      size: 'small',
      category: 'Insights',
      desc: 'An inspiring quote refreshed every day'
    },
    {
      id: 'weather',
      title: 'Field Weather',
      icon: 'fa-cloud-sun',
      size: 'small',
      category: 'Insights',
      desc: 'Live weather at your home-base location'
    },
    {
      id: 'chart-teacher-visit-freq',
      title: 'Teacher Observation Frequency',
      icon: 'fa-user-clock',
      size: 'medium',
      category: 'Charts',
      desc: 'Teachers most observed (from observations data)'
    },
    {
      id: 'chart-engagement-trend',
      title: 'Engagement Rate Trend',
      icon: 'fa-chart-area',
      size: 'medium',
      category: 'Charts',
      desc: 'Monthly classroom engagement % trend'
    }
  ];

  newWidgets.forEach(function (w) {
    var exists = WidgetBuilder.registry.some(function (r) { return r.id === w.id; });
    if (!exists) WidgetBuilder.registry.push(w);
  });

  // Append to defaultOrder (hidden by default via getLayout logic)
  newWidgets.forEach(function (w) {
    if (!WidgetBuilder.defaultOrder.includes(w.id)) {
      WidgetBuilder.defaultOrder.push(w.id);
    }
  });

  // ── Task 9: Extend _getWidgetLink ────────────────────────────────
  var _origGetLink = WidgetBuilder._getWidgetLink.bind(WidgetBuilder);
  WidgetBuilder._getWidgetLink = function (w) {
    var extra = {
      'today-schedule':          '<a href="#" class="widget-link" onclick="navigateTo(\'planner\')">Planner</a>',
      'motivational-quote':      '',
      'weather':                 '',
      'chart-teacher-visit-freq':'<a href="#" class="widget-link" onclick="navigateTo(\'observations\')">Observations</a>',
      'chart-engagement-trend':  '<a href="#" class="widget-link" onclick="navigateTo(\'analytics\')">Analytics</a>'
    };
    if (w.id in extra) return extra[w.id];
    return _origGetLink(w);
  };

  // ── Task 7: Extend _renderWidgetContent switch ────────────────────
  var _origContent = WidgetBuilder._renderWidgetContent.bind(WidgetBuilder);
  WidgetBuilder._renderWidgetContent = function (w) {
    switch (w.id) {
      case 'today-schedule':          return this._renderTodaySchedule();
      case 'motivational-quote':      return this._renderMotivationalQuote();
      case 'weather':                 return this._renderWeatherWidget();
      case 'chart-teacher-visit-freq':
        return '<canvas id="wbChartTeacherFreq" style="max-height:240px;"></canvas>';
      case 'chart-engagement-trend':
        return '<canvas id="wbChartEngageTrend" style="max-height:240px;"></canvas>';
      default: return _origContent(w);
    }
  };

  // ── Task 8: Extend render() rAF loop ─────────────────────────────
  var _origRender = WidgetBuilder.render.bind(WidgetBuilder);
  WidgetBuilder.render = function () {
    _origRender();
    // Re-run the rAF (original already ran one, this adds ours after DOM settles)
    var self = this;
    requestAnimationFrame(function () {
      var layout  = self.getLayout();
      var visible = layout.order.filter(function (id) { return !layout.hidden.includes(id); });
      if (visible.includes('chart-teacher-visit-freq')) self._renderChartTeacherFreq();
      if (visible.includes('chart-engagement-trend'))   self._renderChartEngageTrend();
    });
  };

  // ── Task 2: Today's Schedule widget ──────────────────────────────
  WidgetBuilder._renderTodaySchedule = function () {
    var today      = new Date().toISOString().slice(0, 10);
    var visits     = (typeof DB !== 'undefined' ? DB.get('visits')       : null) || [];
    var trainings  = (typeof DB !== 'undefined' ? DB.get('trainings')    : null) || [];
    var tasks      = (typeof DB !== 'undefined' ? DB.get('plannerTasks') : null) || [];

    var items = [];
    visits.filter(function (v) { return v.date === today; }).forEach(function (v) {
      items.push({
        time: v.time || '—',
        icon: 'fa-school',
        cls: 'visit',
        text: (typeof escapeHtml === 'function' ? escapeHtml(v.school) : v.school) +
              (v.purpose ? ' <span class="sched-tag">' + (typeof escapeHtml === 'function' ? escapeHtml(v.purpose) : v.purpose) + '</span>' : '')
      });
    });
    trainings.filter(function (t) { return t.date === today; }).forEach(function (t) {
      items.push({
        time: t.time || '—',
        icon: 'fa-chalkboard-teacher',
        cls: 'training',
        text: (typeof escapeHtml === 'function' ? escapeHtml(t.title || 'Training') : t.title || 'Training')
      });
    });
    tasks.filter(function (t) { return t.date === today && !t.done; }).forEach(function (t) {
      items.push({
        time: '',
        icon: 'fa-tasks',
        cls: 'task',
        text: (typeof escapeHtml === 'function' ? escapeHtml(t.task || t.text || 'Task') : t.task || t.text || 'Task')
      });
    });

    if (items.length === 0) {
      var dayName = new Date().toLocaleDateString('en', { weekday: 'long' });
      return '<div class="empty-state small">' +
             '<i class="fas fa-calendar-check"></i>' +
             '<p>No activities scheduled for ' + dayName + '</p>' +
             '<a href="#" class="widget-link" onclick="navigateTo(\'planner\')" ' +
             'style="margin-top:8px;display:inline-block">Add to Planner</a></div>';
    }

    return '<div class="widget-schedule-list">' +
      items.map(function (it) {
        return '<div class="schedule-item">' +
               '<div class="schedule-time">' + it.time + '</div>' +
               '<div class="schedule-icon ' + it.cls + '"><i class="fas ' + it.icon + '"></i></div>' +
               '<div class="schedule-text">' + it.text + '</div>' +
               '</div>';
      }).join('') +
    '</div>';
  };

  // ── Task 3: Motivational Quote — Sarvam AI (live) + local fallback ─
  // Local fallback quotes (used when offline or no API key)
  var _QUOTES = [
    { text: "Education is the most powerful weapon which you can use to change the world.", author: "Nelson Mandela" },
    { text: "The best teachers are those who show you where to look but don't tell you what to see.", author: "Alexandra K. Trenfor" },
    { text: "Every child deserves a champion — an adult who will never give up on them.", author: "Rita Pierson" },
    { text: "Teaching is the one profession that creates all other professions.", author: "Unknown" },
    { text: "The art of teaching is the art of assisting discovery.", author: "Mark Van Doren" },
    { text: "One book, one pen, one child, and one teacher can change the world.", author: "Malala Yousafzai" },
    { text: "The mediocre teacher tells. The good teacher explains. The great teacher inspires.", author: "William Arthur Ward" },
    { text: "Small progress is still progress.", author: "Unknown" },
    { text: "The roots of education are bitter, but the fruit is sweet.", author: "Aristotle" },
    { text: "To teach is to learn twice.", author: "Joseph Joubert" },
    { text: "Change is the end result of all true learning.", author: "Leo Buscaglia" },
    { text: "It is the supreme art of the teacher to awaken joy in creative expression and knowledge.", author: "Albert Einstein" },
    { text: "Believe you can and you're halfway there.", author: "Theodore Roosevelt" },
    { text: "The beautiful thing about learning is nobody can take it away from you.", author: "B.B. King" },
    { text: "Keep going. Everything you need will come to you at the perfect time.", author: "Unknown" }
  ];

  // Render: show loading spinner, then fill via async Sarvam AI call
  WidgetBuilder._renderMotivationalQuote = function () {
    var containerId = 'wbMotivationContainer';
    setTimeout(function () { _fetchSarvamQuote(containerId); }, 80);
    return '<div id="' + containerId + '" class="widget-quote quote-loading">' +
           '<i class="fas fa-spinner fa-spin quote-icon" style="font-size:18px;opacity:0.5"></i>' +
           '<p class="quote-text" style="opacity:0.4">Generating today\'s motivation...</p>' +
           '</div>';
  };

  // ── Task 4: Weather widget ────────────────────────────────────────
  WidgetBuilder._renderWeatherWidget = function () {
    var containerId = 'wbWeatherContainer';
    // Return placeholder; fetch weather asynchronously
    setTimeout(function () { _fetchWeatherWidget(containerId); }, 100);
    return '<div id="' + containerId + '" class="widget-weather">' +
           '<div class="weather-loading"><i class="fas fa-spinner fa-spin"></i> Detecting location...</div>' +
           '</div>';
  };

  // ── Task 5: Teacher Visit Frequency chart ────────────────────────
  WidgetBuilder._renderChartTeacherFreq = function () {
    var canvas = document.getElementById('wbChartTeacherFreq');
    if (!canvas) return;
    var observations = (typeof DB !== 'undefined' ? DB.get('observations') : null) || [];
    var freq = {};
    observations.forEach(function (o) {
      // 'teacher' is the actual field name used by DMT observation model (app.js line 2309)
      var name = (o.teacher || '').trim();
      if (name) freq[name] = (freq[name] || 0) + 1;
    });
    var sorted = Object.entries(freq).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 8);
    if (sorted.length === 0) {
      canvas.style.display = 'none';
      canvas.insertAdjacentHTML('afterend',
        '<div class="empty-state small"><i class="fas fa-user-clock"></i>' +
        '<p>Add observations with teacher names to see frequency</p></div>');
      return;
    }
    var colors = ['#6366f1','#8b5cf6','#3b82f6','#10b981','#f59e0b','#ef4444','#ec4899','#14b8a6'];
    this._charts['teacherFreq'] = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: sorted.map(function (s) { return s[0]; }),
        datasets: [{
          label: 'Observations',
          data: sorted.map(function (s) { return s[1]; }),
          backgroundColor: sorted.map(function (_, i) { return colors[i % colors.length]; }),
          borderWidth: 0,
          borderRadius: 6
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { beginAtZero: true, ticks: { color: '#6b7280', precision: 0 }, grid: { color: 'rgba(255,255,255,0.04)' } },
          y: { ticks: { color: '#9ca3b8', font: { size: 11 } }, grid: { display: false } }
        }
      }
    });
  };

  // ── Task 6: Engagement Rate Trend chart ──────────────────────────
  WidgetBuilder._renderChartEngageTrend = function () {
    var canvas = document.getElementById('wbChartEngageTrend');
    if (!canvas) return;
    var observations = (typeof DB !== 'undefined' ? DB.get('observations') : null) || [];
    var now = new Date();
    var labels = [], data = [];

    for (var i = -5; i <= 0; i++) {
      var d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      labels.push(d.toLocaleDateString('en', { month: 'short' }));
      var y = d.getFullYear(), m = d.getMonth();
      var monthObs = observations.filter(function (o) {
        if (!o.date) return false;
        var od = new Date(o.date);
        return od.getFullYear() === y && od.getMonth() === m;
      });
      var engaged = monthObs.filter(function (o) {
        return (parseInt(o.engagementScore || o.engagement || 0) >= 70) ||
               (o.classroomEngagement === 'High') ||
               (o.engagement === 'High');
      }).length;
      data.push(monthObs.length > 0 ? Math.round((engaged / monthObs.length) * 100) : 0);
    }

    this._charts['engageTrend'] = new Chart(canvas, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: 'Engagement %',
          data: data,
          borderColor: '#f59e0b',
          backgroundColor: 'rgba(245,158,11,0.08)',
          fill: true,
          tension: 0.4,
          pointRadius: 5,
          pointBackgroundColor: '#f59e0b'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#9ca3b8', font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: function (ctx) { return ' ' + ctx.parsed.y + '% engaged'; }
            }
          }
        },
        scales: {
          x: { ticks: { color: '#6b7280' }, grid: { color: 'rgba(255,255,255,0.04)' } },
          y: {
            beginAtZero: true, max: 100,
            ticks: { color: '#6b7280', callback: function (v) { return v + '%'; } },
            grid: { color: 'rgba(255,255,255,0.04)' }
          }
        }
      }
    });
  };

})(); // end WidgetBuilder extension IIFE

// ─────────────────────────────────────────────────────────────────────
// Weather fetch helper (Task 4 continued)
// ─────────────────────────────────────────────────────────────────────
function _fetchWeatherWidget(containerId) {
  var el = document.getElementById(containerId);
  if (!el) return;

  if (!navigator.geolocation) {
    el.innerHTML = '<div class="weather-error"><i class="fas fa-map-marker-slash"></i> Location unavailable</div>';
    return;
  }

  navigator.geolocation.getCurrentPosition(
    function (pos) {
      var lat = pos.coords.latitude.toFixed(4);
      var lon = pos.coords.longitude.toFixed(4);
      var url = 'https://api.open-meteo.com/v1/forecast?latitude=' + lat +
                '&longitude=' + lon +
                '&current_weather=true&hourly=relativehumidity_2m&timezone=auto';
      fetch(url)
        .then(function (r) { return r.json(); })
        .then(function (data) {
          var cw = data.current_weather;
          var temp = Math.round(cw.temperature);
          var wind = Math.round(cw.windspeed);
          var wcode = cw.weathercode;
          var icon = _weatherIcon(wcode);
          var desc = _weatherDesc(wcode);
          el.innerHTML =
            '<div class="weather-main">' +
            '<div class="weather-icon-big"><i class="fas ' + icon + '"></i></div>' +
            '<div class="weather-info">' +
            '<div class="weather-temp">' + temp + '°C</div>' +
            '<div class="weather-desc">' + desc + '</div>' +
            '<div class="weather-meta"><i class="fas fa-wind"></i> ' + wind + ' km/h</div>' +
            '</div></div>' +
            '<div class="weather-coords"><i class="fas fa-map-marker-alt"></i> ' + lat + ', ' + lon + '</div>';
        })
        .catch(function () {
          el.innerHTML = '<div class="weather-error"><i class="fas fa-wifi-slash"></i> Weather unavailable</div>';
        });
    },
    function () {
      el.innerHTML = '<div class="weather-error"><i class="fas fa-ban"></i> Location access denied</div>';
    },
    { timeout: 8000 }
  );
}

function _weatherIcon(code) {
  if (code === 0) return 'fa-sun';
  if (code <= 2) return 'fa-cloud-sun';
  if (code <= 45) return 'fa-cloud';
  if (code <= 67) return 'fa-cloud-rain';
  if (code <= 77) return 'fa-snowflake';
  if (code <= 82) return 'fa-cloud-showers-heavy';
  if (code <= 99) return 'fa-bolt';
  return 'fa-cloud';
}

function _weatherDesc(code) {
  var map = {
    0:'Clear sky', 1:'Mainly clear', 2:'Partly cloudy', 3:'Overcast',
    45:'Fog', 48:'Icy fog', 51:'Light drizzle', 53:'Drizzle',
    55:'Heavy drizzle', 61:'Slight rain', 63:'Rain', 65:'Heavy rain',
    71:'Light snow', 73:'Snow', 75:'Heavy snow', 80:'Rain showers',
    81:'Rain showers', 82:'Heavy showers', 95:'Thunderstorm',
    96:'Thunderstorm w/ hail', 99:'Thunderstorm w/ hail'
  };
  return map[code] || 'Weather';
}

// ─────────────────────────────────────────────────────────────────────
// Sarvam AI Quote Fetcher
// API: https://api.sarvam.ai/v1/chat/completions  (model: sarvam-m)
// Auth header: api-subscription-key
//
// First-time setup — run once in the browser console:
//   setSarvamApiKey('YOUR_SARVAM_API_KEY')
//
// Key is stored in localStorage so it persists across sessions.
// Quotes are cached in sessionStorage (once per day) to avoid
// burning API credits on every widget re-render.
// ─────────────────────────────────────────────────────────────────────

var _SARVAM_KEY_STORE  = 'apf_sarvam_api_key';
var _SARVAM_CACHE_KEY  = 'apf_sarvam_quote_cache';
var _SARVAM_CACHE_DATE = 'apf_sarvam_quote_date';

/** Call this once from the browser console to save your Sarvam AI key */
function setSarvamApiKey(key) {
  if (!key || typeof key !== 'string') { console.warn('[APF] Usage: setSarvamApiKey("your-key-here")'); return; }
  localStorage.setItem(_SARVAM_KEY_STORE, key.trim());
  // Bust the quote cache so the new key is used immediately
  sessionStorage.removeItem(_SARVAM_CACHE_KEY);
  sessionStorage.removeItem(_SARVAM_CACHE_DATE);
  if (typeof showToast === 'function') showToast('Sarvam AI key saved! Refresh the widget to see a live quote.', 'success');
  console.log('[APF] Sarvam AI API key saved. Run WidgetBuilder.render() to refresh the dashboard.');
}

/** Returns the stored Sarvam AI key (or empty string) */
function getSarvamApiKey() {
  return localStorage.getItem(_SARVAM_KEY_STORE) || '';
}

/** Async fetch: Sarvam AI → sessionStorage cache → local fallback */
function _fetchSarvamQuote(containerId) {
  var el = document.getElementById(containerId);
  if (!el) return;

  var today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD

  // ── 1. Check sessionStorage cache (valid for today) ──
  try {
    var cachedDate  = sessionStorage.getItem(_SARVAM_CACHE_DATE);
    var cachedQuote = sessionStorage.getItem(_SARVAM_CACHE_KEY);
    if (cachedDate === today && cachedQuote) {
      var parsed = JSON.parse(cachedQuote);
      _renderQuoteInContainer(el, parsed.text, parsed.author, 'sarvam');
      return;
    }
  } catch (e) { /* ignore cache errors */ }

  // ── 2. Try Sarvam AI if key is set ──
  var apiKey = getSarvamApiKey();
  if (!apiKey) {
    // No key: use local fallback + show hint
    _renderQuoteFallback(el, true);
    return;
  }

  var prompt = 'Give me one short motivational quote for a teacher or education professional working in rural India. ' +
               'The quote should be inspiring, original and practical. ' +
               'Reply ONLY in this exact JSON format (no markdown): ' +
               '{"text":"<the quote>","author":"<author name or Sarvam AI>"}';

  fetch('https://api.sarvam.ai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-subscription-key': apiKey
    },
    body: JSON.stringify({
      model: 'sarvam-m',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 200,
      temperature: 0.9
    })
  })
  .then(function (r) {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
  })
  .then(function (data) {
    var raw = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    // Strip markdown code fences if present
    raw = raw.replace(/```json|```/g, '').trim();
    var quote;
    try { quote = JSON.parse(raw); } catch (e) {
      // Try extracting with regex if JSON parse fails
      var tm = raw.match(/"text"\s*:\s*"([^"]+)"/);
      var am = raw.match(/"author"\s*:\s*"([^"]+)"/);
      quote = tm ? { text: tm[1], author: am ? am[1] : 'Sarvam AI' } : null;
    }
    if (!quote || !quote.text) { _renderQuoteFallback(el, false); return; }

    // Cache for today
    try {
      sessionStorage.setItem(_SARVAM_CACHE_DATE, today);
      sessionStorage.setItem(_SARVAM_CACHE_KEY, JSON.stringify(quote));
    } catch (e) { /* ignore */ }

    _renderQuoteInContainer(el, quote.text, quote.author || 'Sarvam AI', 'sarvam');
  })
  .catch(function (err) {
    console.warn('[APF] Sarvam AI quote fetch failed:', err.message);
    _renderQuoteFallback(el, false);
  });
}

function _renderQuoteInContainer(el, text, author, source) {
  var badge = source === 'sarvam'
    ? '<span class="quote-ai-badge"><i class="fas fa-robot"></i> Sarvam AI</span>'
    : '';
  el.classList.remove('quote-loading');
  el.innerHTML =
    '<i class="fas fa-quote-left quote-icon"></i>' +
    '<p class="quote-text">' + (typeof escapeHtml === 'function' ? escapeHtml(text) : text) + '</p>' +
    '<p class="quote-author">\u2014 ' + (typeof escapeHtml === 'function' ? escapeHtml(author) : author) + '</p>' +
    badge;
}

function _renderQuoteFallback(el, showSetupHint) {
  // Pick from local array based on day-of-year
  var _FB = [
    { text: 'Education is the most powerful weapon which you can use to change the world.', author: 'Nelson Mandela' },
    { text: 'Every child deserves a champion \u2014 an adult who will never give up on them.', author: 'Rita Pierson' },
    { text: 'The art of teaching is the art of assisting discovery.', author: 'Mark Van Doren' },
    { text: 'One book, one pen, one child, and one teacher can change the world.', author: 'Malala Yousafzai' },
    { text: 'The mediocre teacher tells. The good teacher explains. The great teacher inspires.', author: 'William Arthur Ward' },
    { text: 'Small progress is still progress.', author: 'Unknown' },
    { text: 'The roots of education are bitter, but the fruit is sweet.', author: 'Aristotle' },
    { text: 'To teach is to learn twice.', author: 'Joseph Joubert' },
    { text: 'It is the supreme art of the teacher to awaken joy in creative expression.', author: 'Albert Einstein' },
    { text: 'Believe you can and you\u2019re halfway there.', author: 'Theodore Roosevelt' }
  ];
  var dayOfYear = Math.floor((new Date() - new Date(new Date().getFullYear(), 0, 0)) / 86400000);
  var q = _FB[dayOfYear % _FB.length];
  _renderQuoteInContainer(el, q.text, q.author, 'local');
}

// ─────────────────────────────────────────────────────────────────────
// TASK 13–14 · Quick Action FAB
// ─────────────────────────────────────────────────────────────────────
function openQuickFab() {
  var menu = document.getElementById('fabMenu');
  var btn  = document.getElementById('fabMainBtn');
  if (!menu) return;
  var isOpen = menu.classList.contains('fab-menu--open');
  if (isOpen) {
    menu.classList.remove('fab-menu--open');
    if (btn) btn.classList.remove('fab-main--open');
  } else {
    menu.classList.add('fab-menu--open');
    if (btn) btn.classList.add('fab-main--open');
  }
}

function closeQuickFab() {
  var menu = document.getElementById('fabMenu');
  var btn  = document.getElementById('fabMainBtn');
  if (menu) menu.classList.remove('fab-menu--open');
  if (btn)  btn.classList.remove('fab-main--open');
}

// Close FAB on outside click
document.addEventListener('click', function (e) {
  var fab = document.getElementById('fabContainer');
  if (fab && !fab.contains(e.target)) closeQuickFab();
});

// ─────────────────────────────────────────────────────────────────────
// TASK 15 · Report Writing Module
// ─────────────────────────────────────────────────────────────────────

var _pageStateReports = 1;

function openReportModal(id) {
  var report = id ? (DB.get('reportDrafts') || []).find(function (r) { return r.id === id; }) : null;
  document.getElementById('reportEditId').value     = report ? report.id : '';
  document.getElementById('reportTitle').value      = report ? report.title : '';
  document.getElementById('reportType').value       = report ? (report.type || 'Visit Report') : 'Visit Report';
  document.getElementById('reportSchool').value     = report ? (report.school || '') : '';
  document.getElementById('reportDate').value       = report ? (report.date || new Date().toISOString().slice(0,10)) : new Date().toISOString().slice(0,10);
  document.getElementById('reportContent').value    = report ? (report.content || '') : '';
  document.getElementById('reportStatus').value     = report ? (report.status || 'draft') : 'draft';
  document.getElementById('reportModalTitle').textContent = (report ? 'Edit' : 'New') + ' Report';
  openModal('reportWritingModal');
}

function saveReport() {
  var id      = document.getElementById('reportEditId').value;
  var title   = document.getElementById('reportTitle').value.trim();
  var type    = document.getElementById('reportType').value;
  var school  = document.getElementById('reportSchool').value.trim();
  var date    = document.getElementById('reportDate').value;
  var content = document.getElementById('reportContent').value.trim();
  var status  = document.getElementById('reportStatus').value;

  if (!title) { if (typeof showToast === 'function') showToast('Title is required', 'error'); return; }
  if (!content) { if (typeof showToast === 'function') showToast('Report content is required', 'error'); return; }

  var reports = DB.get('reportDrafts') || [];
  if (id) {
    var idx = reports.findIndex(function (r) { return r.id === id; });
    if (idx !== -1) {
      reports[idx] = Object.assign(reports[idx], { title, type, school, date, content, status, updatedAt: new Date().toISOString() });
    }
  } else {
    reports.unshift({
      id: (typeof DB !== 'undefined' ? DB.generateId() : Date.now().toString(36)),
      title, type, school, date, content, status,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  DB.set('reportDrafts', reports);
  if (typeof markUnsavedChanges === 'function') markUnsavedChanges();
  closeModal('reportWritingModal');
  renderReports();
  if (typeof showToast === 'function') showToast('Report saved! ✅', 'success');
}

function deleteReport(id) {
  if (!confirm('Delete this report draft?')) return;
  var reports = (DB.get('reportDrafts') || []).filter(function (r) { return r.id !== id; });
  DB.set('reportDrafts', reports);
  if (typeof markUnsavedChanges === 'function') markUnsavedChanges();
  renderReports();
  if (typeof showToast === 'function') showToast('Report deleted', 'info');
}

function renderReports() {
  var container = document.getElementById('reportsContainer');
  if (!container) return;

  var reports  = DB.get('reportDrafts') || [];
  var search   = (document.getElementById('reportSearchInput')?.value || '').toLowerCase();
  var typeF    = document.getElementById('reportTypeFilter')?.value || 'all';
  var statusF  = document.getElementById('reportStatusFilter')?.value || 'all';

  var filtered = reports.filter(function (r) {
    if (typeF   !== 'all' && r.type   !== typeF)   return false;
    if (statusF !== 'all' && r.status !== statusF)  return false;
    if (search && !(
      (r.title   || '').toLowerCase().includes(search) ||
      (r.school  || '').toLowerCase().includes(search) ||
      (r.content || '').toLowerCase().includes(search)
    )) return false;
    return true;
  });

  // Update stats
  var el = function (id) { return document.getElementById(id); };
  if (el('reportStatTotal'))  el('reportStatTotal').textContent  = reports.length;
  if (el('reportStatDraft'))  el('reportStatDraft').textContent  = reports.filter(function(r){return r.status==='draft';}).length;
  if (el('reportStatFinal'))  el('reportStatFinal').textContent  = reports.filter(function(r){return r.status==='final';}).length;

  if (filtered.length === 0) {
    container.innerHTML = '<div class="empty-state">' +
      '<i class="fas fa-file-alt"></i>' +
      '<h3>No reports found</h3>' +
      '<p>Click "New Report" to create your first report draft</p></div>';
    return;
  }

  var statusBadge = { draft: 'badge-warning', final: 'badge-success', submitted: 'badge-info' };
  var typeColor   = {
    'Visit Report':        '#6366f1',
    'Training Report':     '#8b5cf6',
    'Monthly Report':      '#3b82f6',
    'Observation Report':  '#10b981',
    'Field Note':          '#f59e0b',
    'Other':               '#6b7280'
  };

  container.innerHTML = filtered.map(function (r) {
    var color = typeColor[r.type] || '#6b7280';
    var date  = r.date ? new Date(r.date).toLocaleDateString('en', { day:'numeric', month:'short', year:'numeric' }) : '—';
    var preview = (r.content || '').substring(0, 140);
    if ((r.content || '').length > 140) preview += '…';
    return '<div class="report-card">' +
      '<div class="report-card-header">' +
      '<div class="report-type-bar" style="background:' + color + '"></div>' +
      '<div class="report-card-meta">' +
      '<span class="report-tag" style="background:' + color + '20;color:' + color + '">' + (r.type || 'Report') + '</span>' +
      '<span class="report-status badge ' + (statusBadge[r.status] || 'badge-warning') + '">' + (r.status || 'draft') + '</span>' +
      '</div>' +
      '<div class="report-card-actions">' +
      '<button class="btn btn-ghost btn-sm" onclick="exportReportPDF(\'' + r.id + '\')" title="Export PDF"><i class="fas fa-file-pdf"></i></button>' +
      '<button class="btn btn-ghost btn-sm" onclick="openReportModal(\'' + r.id + '\')" title="Edit"><i class="fas fa-edit"></i></button>' +
      '<button class="btn btn-ghost btn-sm" onclick="deleteReport(\'' + r.id + '\')" title="Delete" style="color:#ef4444"><i class="fas fa-trash"></i></button>' +
      '</div></div>' +
      '<h3 class="report-card-title">' + (typeof escapeHtml === 'function' ? escapeHtml(r.title) : r.title) + '</h3>' +
      (r.school ? '<div class="report-card-school"><i class="fas fa-school"></i> ' + (typeof escapeHtml === 'function' ? escapeHtml(r.school) : r.school) + '</div>' : '') +
      '<p class="report-card-preview">' + (typeof escapeHtml === 'function' ? escapeHtml(preview) : preview) + '</p>' +
      '<div class="report-card-footer"><i class="fas fa-calendar-alt"></i> ' + date + '</div>' +
      '</div>';
  }).join('');
}

function exportReportPDF(id) {
  var report = (DB.get('reportDrafts') || []).find(function (r) { return r.id === id; });
  if (!report) return;

  if (typeof html2pdf === 'undefined') {
    if (typeof showToast === 'function') showToast('PDF library not loaded', 'error');
    return;
  }

  var content = '<div style="font-family:Inter,sans-serif;padding:32px;color:#111;">' +
    '<h1 style="color:#4f46e5;margin-bottom:4px">' + report.title + '</h1>' +
    '<p style="color:#6b7280;font-size:13px;margin-bottom:24px">' +
    (report.type || '') + (report.school ? ' · ' + report.school : '') +
    (report.date ? ' · ' + report.date : '') +
    '</p>' +
    '<hr style="border-color:#e5e7eb;margin-bottom:20px">' +
    '<div style="line-height:1.8;font-size:14px;white-space:pre-wrap">' + report.content + '</div>' +
    '</div>';

  var el = document.createElement('div');
  el.innerHTML = content;
  document.body.appendChild(el);

  html2pdf().set({
    margin: 15,
    filename: (report.title.replace(/[^a-z0-9]/gi,'_') || 'report') + '.pdf',
    html2canvas: { scale: 2 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  }).from(el).save().then(function () {
    document.body.removeChild(el);
    if (typeof showToast === 'function') showToast('PDF exported!', 'success');
  });
}

// ─────────────────────────────────────────────────────────────────────
// TASK 16 · Parent Contacts Module
// ─────────────────────────────────────────────────────────────────────

function openParentContactModal(id) {
  var contacts = DB.get('parentContacts') || [];
  var c = id ? contacts.find(function (x) { return x.id === id; }) : null;
  document.getElementById('pcEditId').value       = c ? c.id : '';
  document.getElementById('pcParentName').value   = c ? c.parentName : '';
  document.getElementById('pcPhone').value        = c ? c.phone : '';
  document.getElementById('pcEmail').value        = c ? (c.email || '') : '';
  document.getElementById('pcStudentName').value  = c ? (c.studentName || '') : '';
  document.getElementById('pcClass').value        = c ? (c.studentClass || '') : '';
  document.getElementById('pcSchool').value       = c ? (c.school || '') : '';
  document.getElementById('pcRelation').value     = c ? (c.relation || 'Father') : 'Father';
  document.getElementById('pcNotes').value        = c ? (c.notes || '') : '';
  document.getElementById('pcModalTitle').textContent = (c ? 'Edit' : 'New') + ' Parent Contact';
  openModal('parentContactModal');
}

function saveParentContact() {
  var id          = document.getElementById('pcEditId').value;
  var parentName  = document.getElementById('pcParentName').value.trim();
  var phone       = document.getElementById('pcPhone').value.trim();
  var email       = document.getElementById('pcEmail').value.trim();
  var studentName = document.getElementById('pcStudentName').value.trim();
  var studentClass= document.getElementById('pcClass').value.trim();
  var school      = document.getElementById('pcSchool').value.trim();
  var relation    = document.getElementById('pcRelation').value;
  var notes       = document.getElementById('pcNotes').value.trim();

  if (!parentName) { if (typeof showToast === 'function') showToast('Parent name is required', 'error'); return; }
  if (!phone) { if (typeof showToast === 'function') showToast('Phone number is required', 'error'); return; }

  var contacts = DB.get('parentContacts') || [];
  if (id) {
    var idx = contacts.findIndex(function (c) { return c.id === id; });
    if (idx !== -1) {
      contacts[idx] = Object.assign(contacts[idx], { parentName, phone, email, studentName, studentClass, school, relation, notes, updatedAt: new Date().toISOString() });
    }
  } else {
    contacts.push({
      id: DB.generateId(),
      parentName, phone, email, studentName, studentClass, school, relation, notes,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  DB.set('parentContacts', contacts);
  if (typeof markUnsavedChanges === 'function') markUnsavedChanges();
  closeModal('parentContactModal');
  renderParentContacts();
  if (typeof showToast === 'function') showToast('Contact saved! ✅', 'success');
}

function deleteParentContact(id) {
  if (!confirm('Delete this parent contact?')) return;
  var contacts = (DB.get('parentContacts') || []).filter(function (c) { return c.id !== id; });
  DB.set('parentContacts', contacts);
  if (typeof markUnsavedChanges === 'function') markUnsavedChanges();
  renderParentContacts();
  if (typeof showToast === 'function') showToast('Contact deleted', 'info');
}

function renderParentContacts() {
  var container = document.getElementById('parentContactsContainer');
  if (!container) return;

  var contacts = DB.get('parentContacts') || [];
  var search   = (document.getElementById('pcSearchInput')?.value || '').toLowerCase();
  var schoolF  = (document.getElementById('pcSchoolFilter')?.value || '').toLowerCase();

  var filtered = contacts.filter(function (c) {
    if (schoolF && !(c.school || '').toLowerCase().includes(schoolF)) return false;
    if (search && !(
      (c.parentName  || '').toLowerCase().includes(search) ||
      (c.studentName || '').toLowerCase().includes(search) ||
      (c.phone       || '').toLowerCase().includes(search) ||
      (c.school      || '').toLowerCase().includes(search)
    )) return false;
    return true;
  });

  // Stats
  var el = function (id) { return document.getElementById(id); };
  if (el('pcStatTotal'))   el('pcStatTotal').textContent   = contacts.length;
  if (el('pcStatSchools')) el('pcStatSchools').textContent = new Set(contacts.map(function(c){return c.school||'';}).filter(Boolean)).size;

  if (filtered.length === 0) {
    container.innerHTML = '<div class="empty-state">' +
      '<i class="fas fa-address-card"></i>' +
      '<h3>No parent contacts found</h3>' +
      '<p>Click "Add Contact" to register a parent contact</p></div>';
    return;
  }

  var relationColors = {
    'Father':'#3b82f6','Mother':'#ec4899','Guardian':'#10b981','Other':'#6b7280'
  };

  container.innerHTML = filtered.map(function (c) {
    var initials = (c.parentName || '?').split(' ').map(function(p){return p[0]||'';}).join('').substring(0,2).toUpperCase();
    var color    = relationColors[c.relation] || '#6b7280';
    return '<div class="parent-contact-card">' +
      '<div class="contact-avatar" style="background:' + color + '20;color:' + color + '">' + initials + '</div>' +
      '<div class="contact-body">' +
      '<div class="contact-name">' + (typeof escapeHtml === 'function' ? escapeHtml(c.parentName) : c.parentName) +
      ' <span class="contact-relation" style="background:' + color + '20;color:' + color + '">' + (c.relation || '') + '</span></div>' +
      (c.studentName ? '<div class="contact-meta"><i class="fas fa-child"></i> ' + (typeof escapeHtml === 'function' ? escapeHtml(c.studentName) : c.studentName) + (c.studentClass ? ' (Class ' + c.studentClass + ')' : '') + '</div>' : '') +
      (c.school ? '<div class="contact-meta"><i class="fas fa-school"></i> ' + (typeof escapeHtml === 'function' ? escapeHtml(c.school) : c.school) + '</div>' : '') +
      '<div class="contact-actions">' +
      '<a href="tel:' + c.phone + '" class="btn btn-outline btn-sm"><i class="fas fa-phone"></i> ' + c.phone + '</a>' +
      (c.email ? '<a href="mailto:' + c.email + '" class="btn btn-ghost btn-sm"><i class="fas fa-envelope"></i></a>' : '') +
      '<button class="btn btn-ghost btn-sm" onclick="openParentContactModal(\'' + c.id + '\')"><i class="fas fa-edit"></i></button>' +
      '<button class="btn btn-ghost btn-sm" onclick="deleteParentContact(\'' + c.id + '\')" style="color:#ef4444"><i class="fas fa-trash"></i></button>' +
      '</div></div></div>';
  }).join('');
}

function exportParentContactsExcel() {
  var contacts = DB.get('parentContacts') || [];
  if (!contacts.length) { if (typeof showToast === 'function') showToast('No contacts to export', 'info'); return; }
  if (typeof XLSX === 'undefined') { if (typeof showToast === 'function') showToast('Excel library not loaded', 'error'); return; }

  var rows = contacts.map(function (c) {
    return {
      'Parent Name':   c.parentName   || '',
      'Relation':      c.relation     || '',
      'Phone':         c.phone        || '',
      'Email':         c.email        || '',
      'Student Name':  c.studentName  || '',
      'Class':         c.studentClass || '',
      'School':        c.school       || '',
      'Notes':         c.notes        || '',
      'Added On':      c.createdAt ? new Date(c.createdAt).toLocaleDateString() : ''
    };
  });

  var ws = XLSX.utils.json_to_sheet(rows);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Parent Contacts');
  XLSX.writeFile(wb, 'ParentContacts_' + new Date().toISOString().slice(0,10) + '.xlsx');
  if (typeof showToast === 'function') showToast('Exported to Excel!', 'success');
}

// =====================================================================
// REMINDER ENGINE FOR QUICK NOTES
// Monkey-patches existing saveNote / renderNotes from app.js.
// Channels: Browser Push · Telegram Bot · Email (Google Apps Script)
// =====================================================================

(function _initNoteReminders() {

  // ── Settings store keys ───────────────────────────────────────────
  var RS = {
    emailGasUrl: 'apf_rem_gas_url',
    emailTo:     'apf_rem_email_to'
  };

  function rsGet(k) { return localStorage.getItem(RS[k]) || ''; }
  function rsSet(k, v) { localStorage.setItem(RS[k], v); }

  // ── Wait for app.js globals to be ready ──────────────────────────
  function _whenReady(fn) {
    if (typeof saveNote === 'function' && typeof renderNotes === 'function' && typeof DB !== 'undefined') {
      fn();
    } else {
      setTimeout(function () { _whenReady(fn); }, 150);
    }
  }

  _whenReady(function () {

    // 1. Monkey-patch saveNote
    var _origSaveNote = window.saveNote;
    window.saveNote = function (id) {
      var row      = document.getElementById('noteReminderRow-' + id);
      var remindAt = row ? (row.dataset.remindAt || '') : '';
      var pushEl   = document.getElementById('noteRemindPush-' + id);
      var tgEl     = document.getElementById('noteRemindTelegram-' + id);
      var emailEl  = document.getElementById('noteRemindEmail-' + id);
      var channels = {
        push:     pushEl  ? pushEl.checked  : false,
        telegram: tgEl    ? tgEl.checked    : false,
        email:    emailEl ? emailEl.checked : false
      };
      _origSaveNote(id);
      if (!remindAt) return;
      var reminders = _getReminders();
      reminders[id] = { remindAt: remindAt, channels: channels, fired: false, noteId: id };
      _setReminders(reminders);
      var dt = new Date(remindAt);
      if (typeof showToast === 'function')
        showToast('\u23f0 Reminder set for ' + dt.toLocaleString('en-IN', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' }), 'success');
    };

    // 2. Monkey-patch renderNotes
    var _origRenderNotes = window.renderNotes;
    window.renderNotes = function () {
      _origRenderNotes();
      setTimeout(_injectReminderFields, 50);
      setTimeout(function () {
        if (typeof window._injectNoteCountdowns === 'function') window._injectNoteCountdowns();
      }, 100);
    };
    setTimeout(_injectReminderFields, 300);
    setTimeout(function () {
      if (typeof window._injectNoteCountdowns === 'function') window._injectNoteCountdowns();
    }, 400);

    // 3. Reminder loop
    _checkReminders();
    if (typeof window._checkMorningDigest === 'function') window._checkMorningDigest();
    setInterval(function () {
      _checkReminders();
      if (typeof window._checkMorningDigest === 'function') window._checkMorningDigest();
      if (typeof window._refreshCountdownBadges === 'function') window._refreshCountdownBadges();
    }, 30000);

  }); // end _whenReady

  // -- VDR date picker + time injected into editing note cards
  function _injectReminderFields() {
    var editCards = document.querySelectorAll('.note-card .note-input');
    editCards.forEach(function (input) {
      var id = input.id.replace('noteTitle-', '');
      if (!id || document.getElementById('noteReminderRow-' + id)) return;

      var reminders  = _getReminders();
      var existing   = reminders[id] || {};
      var remindAt   = existing.remindAt || '';
      var ch         = existing.channels || {};
      var storedDate = remindAt ? remindAt.slice(0, 10) : '';
      var storedTime = remindAt ? remindAt.slice(11, 16) : '';

      function _fmtBtn(iso) {
        if (!iso) return '<i class="fas fa-calendar-alt"></i> Pick date';
        var d = new Date(iso + 'T00:00');
        return '<i class="fas fa-calendar-check"></i> ' +
               d.toLocaleDateString('en-IN', { day:'numeric', month:'short', year:'numeric' });
      }

      var row = document.createElement('div');
      row.className        = 'note-reminder-row note-reminder-row--stacked';
      row.id               = 'noteReminderRow-' + id;
      row.dataset.remindAt = remindAt;

      row.innerHTML =
        '<div class="note-reminder-label"><i class="fas fa-bell"></i> Remind me at</div>' +
        '<div class="note-remind-dt-wrap">' +
          '<button type="button" id="noteRemindDateBtn-' + id + '" class="note-remind-date-btn" data-date="' + storedDate + '">' +
            _fmtBtn(storedDate) +
          '</button>' +
          '<input type="time" id="noteRemindTime-' + id + '" class="note-remind-time" value="' + storedTime + '">' +
        '</div>' +
        '<div class="note-reminder-channels">' +
          '<label class="rem-ch-label" title="Push"><input type="checkbox" id="noteRemindPush-' + id + '" ' + (ch.push ? 'checked' : '') + '><i class="fas fa-bell"></i></label>' +
          '<label class="rem-ch-label" title="Telegram"><input type="checkbox" id="noteRemindTelegram-' + id + '" ' + (ch.telegram ? 'checked' : '') + '><i class="fab fa-telegram-plane"></i></label>' +
          '<label class="rem-ch-label" title="Email"><input type="checkbox" id="noteRemindEmail-' + id + '" ' + (ch.email ? 'checked' : '') + '><i class="fas fa-envelope"></i></label>' +
        '</div>';

      var card    = input.closest('.note-card');
      var actions = card ? card.querySelector('.note-card-actions') : null;
      if (actions) card.insertBefore(row, actions);

      var timeInput = document.getElementById('noteRemindTime-' + id);
      var dateBtn   = document.getElementById('noteRemindDateBtn-' + id);

      function _sync() {
        var d = dateBtn ? (dateBtn.dataset.date || '') : '';
        var t = timeInput ? (timeInput.value || '00:00') : '00:00';
        row.dataset.remindAt = d ? (d + 'T' + t) : '';
      }
      if (timeInput) timeInput.addEventListener('change', _sync);

      if (dateBtn) {
        dateBtn.addEventListener('click', function (e) {
          e.stopPropagation();
          if (typeof VDR === 'undefined' || typeof VDR.open !== 'function') {
            var p = prompt('Date (YYYY-MM-DD):', dateBtn.dataset.date || new Date().toISOString().slice(0,10));
            if (p && /^\d{4}-\d{2}-\d{2}$/.test(p)) {
              dateBtn.dataset.date = p; dateBtn.innerHTML = _fmtBtn(p); _sync();
            }
            return;
          }
          if (typeof VDR.syncFrom === 'function') VDR.syncFrom(dateBtn.dataset.date || '', '');
          VDR.open('from', dateBtn, function (picked) {
            if (!picked) return;
            dateBtn.dataset.date = picked;
            dateBtn.innerHTML    = _fmtBtn(picked);
            _sync();
          }, { single: true });
        });
      }
    });
  }
  function _getReminders() {
    try { return JSON.parse(localStorage.getItem('apf_note_reminders') || '{}'); } catch (e) { return {}; }
  }
  function _setReminders(obj) {
    localStorage.setItem('apf_note_reminders', JSON.stringify(obj));
  }

  // ── Check & fire due reminders ────────────────────────────────────
  function _checkReminders() {
    var reminders = _getReminders();
    var now = new Date();
    var changed = false;

    Object.keys(reminders).forEach(function (noteId) {
      var rem = reminders[noteId];
      if (rem.fired) return;
      var remTime = new Date(rem.remindAt);
      // Fire if we're within the current minute window
      if (remTime <= now && (now - remTime) < 90000) {
        _fireReminder(noteId, rem);
        reminders[noteId].fired = true;
        changed = true;
      }
    });

    if (changed) _setReminders(reminders);
  }

  function _fireReminder(noteId, rem) {
    var notes = typeof DB !== 'undefined' ? DB.get('notes') : [];
    var note  = notes.find(function (n) { return n.id === noteId; });
    var title = note ? (note.title || 'Quick Note') : 'Quick Note';
    var body  = note ? (note.content || '') : '';

    if (rem.channels.push)     _sendPushNotification(title, body);
    if (rem.channels.telegram) _sendTelegramReminder(title, body, rem.remindAt);
    if (rem.channels.email)    _sendEmailReminder(title, body, rem.remindAt);
  }

  // ── Channel: Browser Push ─────────────────────────────────────────
  function _sendPushNotification(title, body) {
    if (Notification.permission !== 'granted') { return; }
    try {
      var n = new Notification('⏰ APF Reminder: ' + title, {
        body: body.substring(0, 200) || 'Your scheduled note reminder.',
        icon: './icon-192.png',
        badge: './icon-192.png',
        tag: 'apf-note-' + Date.now(),
        requireInteraction: true
      });
      n.onclick = function () { window.focus(); navigateTo('notes'); };
    } catch (e) { console.warn('[APF Reminder] Push failed:', e); }
  }

  // ── Channel: Telegram ─────────────────────────────────────────────
  // ── Channel: Telegram ─────────────────────────────────────────────
  // Auto-reads from app.js's existing 'apf_telegram_config' — no re-entry needed
  function _getTgCfg() {
    try { return JSON.parse(localStorage.getItem('apf_telegram_config') || 'null'); } catch (e) { return null; }
  }

  function _sendTelegramReminder(title, body, remindAt) {
    var cfg = _getTgCfg();
    if (!cfg || !cfg.token || !cfg.chatId) {
      console.warn('[APF Reminder] Telegram not configured in app settings');
      return;
    }
    var dt  = new Date(remindAt).toLocaleString('en-IN', { dateStyle:'medium', timeStyle:'short' });
    var msg = '\u23f0 *APF Reminder*\n\n' +
              '\uD83D\uDCCC *' + title.replace(/[_*[\]()~`>#+=|{}.!-]/g, '\\$&') + '*\n' +
              (body ? body.substring(0, 300) + '\n' : '') +
              '\n\uD83D\uDCC5 Scheduled: ' + dt;

    fetch('https://api.telegram.org/bot' + cfg.token + '/sendMessage', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: cfg.chatId, text: msg, parse_mode: 'Markdown' })
    }).catch(function (e) { console.warn('[APF Reminder] Telegram error:', e); });
  }

  // ── Channel: Email via Google Apps Script ─────────────────────────
  function _sendEmailReminder(title, body, remindAt) {
    var gasUrl  = rsGet('emailGasUrl');
    var emailTo = rsGet('emailTo');
    if (!gasUrl || !emailTo) { console.warn('[APF Reminder] Email (GAS) not configured'); return; }

    var dt = new Date(remindAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' });
    var payload = JSON.stringify({
      to:      emailTo,
      subject: '\u23f0 APF Reminder: ' + title,
      body:    (body ? body + '\n\n' : '') + 'Scheduled for: ' + dt + '\n\nSent from APF Dashboard'
    });

    // mode: no-cors — fire and forget (GAS doesn't return CORS headers)
    fetch(gasUrl, {
      method:  'POST',
      mode:    'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body:    payload
    }).catch(function (e) { console.warn('[APF Reminder] GAS email error:', e); });
  }

  // ── Push permission helper ────────────────────────────────────────
  window.requestPushPermission = function () {
    Notification.requestPermission().then(function (result) {
      _updatePushStatusBadge();
      if (result === 'granted' && typeof showToast === 'function') {
        showToast('Push notifications enabled! ✅', 'success');
      }
    });
  };

  function _updatePushStatusBadge() {
    var el = document.getElementById('pushPermStatus');
    var btn = document.getElementById('pushPermBtn');
    if (!el) return;
    var p = Notification.permission;
    var colors = { granted: '#10b981', denied: '#ef4444', default: '#f59e0b' };
    var labels = { granted: 'Enabled ✅', denied: 'Blocked ❌ (reset in browser settings)', default: 'Not set — click Enable' };
    el.innerHTML = '<span style="color:' + (colors[p]||'#6b7280') + ';font-size:12px;font-weight:600">' + (labels[p]||p) + '</span>';
    if (btn) btn.style.display = (p === 'granted') ? 'none' : '';
  }

  // ── Reminder Settings modal save ──────────────────────────────────
  window.saveReminderSettings = function () {
    rsSet('emailGasUrl', document.getElementById('rsGasUrl').value.trim());
    rsSet('emailTo',     document.getElementById('rsEmailTo').value.trim());
    closeModal('reminderSettingsModal');
    if (typeof showToast === 'function') showToast('Reminder settings saved! ✅', 'success');
  };

  // ── Test helpers ──────────────────────────────────────────────────
  window.testTelegramReminder = function () {
    var cfg = _getTgCfg();
    if (!cfg || !cfg.token || !cfg.chatId) {
      if (typeof showToast === 'function') showToast('Telegram not configured in APF Settings ❌', 'error');
      return;
    }
    _sendTelegramReminder('Test Reminder 🔔', 'This is a test from your APF Dashboard!', new Date().toISOString());
    if (typeof showToast === 'function') showToast('Test message sent to Telegram! ✅', 'success');
  };

  window.testEmailReminder = function () {
    rsSet('emailGasUrl', document.getElementById('rsGasUrl').value.trim());
    rsSet('emailTo',     document.getElementById('rsEmailTo').value.trim());
    _sendEmailReminder('Test Reminder', 'This is a test from your APF Dashboard!', new Date().toISOString());
    if (typeof showToast === 'function') showToast('Test email sent via Google Apps Script!', 'info');
  };

  // ── Global opener (called from button in HTML) ────────────────────
  window.openReminderSettings = function () {
    // Telegram — show live status from app config
    var tgEl  = document.getElementById('tgReminderStatus');
    var cfg   = _getTgCfg();
    if (tgEl) {
      tgEl.innerHTML = cfg && cfg.token && cfg.chatId
        ? '<span class="rsb-connected"><i class="fas fa-check-circle"></i> Connected (bot configured in APF Settings)</span>'
        : '<span class="rsb-disconnected"><i class="fas fa-exclamation-circle"></i> Not configured — go to APF Settings → Telegram</span>';
    }
    // Email (GAS)
    document.getElementById('rsGasUrl').value  = rsGet('emailGasUrl');
    document.getElementById('rsEmailTo').value = rsGet('emailTo');
    _updatePushStatusBadge();
    if (typeof openModal === 'function') openModal('reminderSettingsModal');
  };

})(); // end reminder engine IIFE

// =====================================================================
// NOTE COUNTDOWN TIMERS · DONE BUTTON · MORNING DIGEST
// =====================================================================

(function _initCountdownAndDigest() {

  // ── Helpers (mirrors keys in the reminder IIFE) ───────────────────
  function _getReminders() {
    try { return JSON.parse(localStorage.getItem('apf_note_reminders') || '{}'); } catch (e) { return {}; }
  }
  function _setReminders(obj) {
    localStorage.setItem('apf_note_reminders', JSON.stringify(obj));
  }
  function _getTgCfg() {
    try { return JSON.parse(localStorage.getItem('apf_telegram_config') || 'null'); } catch (e) { return null; }
  }
  function _rsGet(k) {
    var keys = { emailGasUrl: 'apf_rem_gas_url', emailTo: 'apf_rem_email_to' };
    return localStorage.getItem(keys[k]) || '';
  }

  // ── Format milliseconds → "2h 15m" / "3d 4h" / "45m" ────────────
  function _fmtDiff(ms) {
    var totalMins = Math.floor(Math.abs(ms) / 60000);
    var totalHrs  = Math.floor(totalMins / 60);
    var days      = Math.floor(totalHrs / 24);
    var hrs       = totalHrs % 24;
    var mins      = totalMins % 60;
    if (days > 0)      return days + 'd ' + hrs + 'h';
    if (totalHrs > 0)  return totalHrs + 'h ' + mins + 'm';
    if (totalMins > 0) return totalMins + 'm';
    return 'now';
  }

  // ── Build inner HTML for the countdown badge (NO outer wrapper) ────
  function _buildBadgeInner(rem, id) {
    if (rem.done) {
      return { cls: 'done', html: '<span class="note-cd-text done-text"><i class="fas fa-check-circle"></i> Done</span>' };
    }
    var now    = new Date();
    var target = new Date(rem.remindAt);
    var diff   = target - now;
    var cls, icon, text;

    if (diff < 0) {
      cls  = 'overdue'; icon = 'fa-exclamation-circle';
      text = 'Overdue by ' + _fmtDiff(diff);
    } else {
      cls  = 'pending'; icon = 'fa-clock';
      text = _fmtDiff(diff) + ' remaining';
    }

    var chHtml = '';
    if (rem.channels) {
      if (rem.channels.push)     chHtml += '<i class="fas fa-bell note-ch-mini" title="Push"></i>';
      if (rem.channels.telegram) chHtml += '<i class="fab fa-telegram-plane note-ch-mini" title="Telegram"></i>';
      if (rem.channels.email)    chHtml += '<i class="fas fa-envelope note-ch-mini" title="Email"></i>';
    }

    return { cls: cls,
             html: '<span class="note-cd-text"><i class="fas ' + icon + '"></i> ' + text + '</span>' +
                   chHtml +
                   '<button class="btn-note-done" onclick="markNoteDone(\'' + id + '\')">' +
                   '<i class="fas fa-check"></i> Done</button>' };
  }

  // ── Inject countdown badges on SAVED (non-editing) note cards ─────
  window._injectNoteCountdowns = function () {
    var reminders = _getReminders();
    document.querySelectorAll('.note-card').forEach(function (card) {
      if (card.querySelector('.note-input')) return; // skip editing cards

      // Get note ID from Edit button — app uses onclick="editNote('id')"
      var editBtn = card.querySelector('[onclick*="editNote"]');
      if (!editBtn) return;
      var raw   = editBtn.getAttribute('onclick') || '';
      var match = raw.match(/editNote\(['"]([^'"]+)['"]/); // handles both ' and "
      if (!match) return;
      var id = match[1];

      var rem = reminders[id];
      if (!rem || !rem.remindAt) return;

      // Update existing badge or create new one
      var el = document.getElementById('noteStatus-' + id);
      if (!el) {
        el = document.createElement('div');
        el.id = 'noteStatus-' + id;
        var actions = card.querySelector('.note-card-actions');
        if (actions) card.insertBefore(el, actions);
        else         card.appendChild(el);
      }

      // Set class + inner content
      var built = rem.done
        ? { cls: 'done', html: '<span class="note-cd-text done-text"><i class="fas fa-check-circle"></i> Done</span>' }
        : _buildBadgeInner(rem, id);
      el.className = 'note-cd-badge ' + built.cls;
      el.innerHTML = built.html;
    });
  };

  // ── Refresh all visible countdown badges (called every 30s) ───────
  window._refreshCountdownBadges = function () {
    var reminders = _getReminders();
    document.querySelectorAll('[id^="noteStatus-"]').forEach(function (el) {
      var id  = el.id.replace('noteStatus-', '');
      var rem = reminders[id];
      if (!rem) return;
      var built = rem.done
        ? { cls: 'done', html: '<span class="note-cd-text done-text"><i class="fas fa-check-circle"></i> Done</span>' }
        : _buildBadgeInner(rem, id);
      el.className = 'note-cd-badge ' + built.cls;
      el.innerHTML = built.html;
    });
  };

  // ── Mark note as done ─────────────────────────────────────────────
  window.markNoteDone = function (id) {
    var reminders = _getReminders();
    if (!reminders[id]) return;
    reminders[id].done = true;
    _setReminders(reminders);

    // Update badge in place
    var el = document.getElementById('noteStatus-' + id);
    if (el) {
      el.className = 'note-cd-badge done';
      el.innerHTML = '<span class="note-cd-text done-text"><i class="fas fa-check-circle"></i> Done</span>';
    }

    if (typeof showToast === 'function') showToast('Marked as done! \u2705', 'success');
  };

  // ── Morning Digest (fires at 8:00 AM every day) ───────────────────
  window._checkMorningDigest = function () {
    var now = new Date();
    // Only fire between 08:00 and 08:01
    if (now.getHours() !== 8 || now.getMinutes() > 1) return;

    var today    = now.toISOString().slice(0, 10);
    var sentKey  = 'apf_morning_digest_sent';
    if (localStorage.getItem(sentKey) === today) return; // already sent today
    localStorage.setItem(sentKey, today);

    _sendMorningDigest();
  };

  function _sendMorningDigest() {
    var reminders = _getReminders();
    var notes     = (typeof DB !== 'undefined' ? DB.get('notes') : null) || [];

    // Collect all unfinished notes that have a reminder set
    var unfinished = [];
    Object.keys(reminders).forEach(function (noteId) {
      var rem = reminders[noteId];
      if (rem.done) return;
      var note = notes.find(function (n) { return n.id === noteId; });
      if (!note) return;
      unfinished.push({ note: note, rem: rem });
    });

    if (unfinished.length === 0) return; // nothing to report

    var count = unfinished.length;

    // ── Push notification ──
    if (Notification.permission === 'granted') {
      try {
        var titles = unfinished.map(function (u) { return (u.note.title || 'Untitled'); }).join(', ');
        var n = new Notification('\uD83C\uDF05 Good Morning! ' + count + ' unfinished reminder' + (count > 1 ? 's' : ''), {
          body: titles.substring(0, 200),
          icon: './icon-192.png',
          requireInteraction: true,
          tag: 'apf-morning-digest'
        });
        n.onclick = function () { window.focus(); if (typeof navigateTo === 'function') navigateTo('notes'); };
      } catch (e) { console.warn('[APF Digest] Push error:', e); }
    }

    // ── Telegram ──
    var cfg = _getTgCfg();
    if (cfg && cfg.token && cfg.chatId) {
      var tgMsg = '\uD83C\uDF05 *Good Morning! Daily Reminder Digest*\n' +
                  '_You have *' + count + '* unfinished reminder' + (count > 1 ? 's' : '') + ' today_\n\n';
      unfinished.forEach(function (u, i) {
        var dt = u.rem.remindAt
          ? new Date(u.rem.remindAt).toLocaleString('en-IN', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' })
          : 'No time set';
        tgMsg += (i + 1) + '. \uD83D\uDCCC *' + (u.note.title || 'Untitled').replace(/[_*[\]()~`>#+=|{}.!-]/g, '\\$&') + '*\n';
        if (u.note.content) tgMsg += '   ' + u.note.content.substring(0, 80) + (u.note.content.length > 80 ? '\u2026' : '') + '\n';
        tgMsg += '   \u23f0 ' + dt + '\n\n';
      });
      tgMsg += '_Open APF Dashboard \u2192 Quick Notes to mark items done._';

      fetch('https://api.telegram.org/bot' + cfg.token + '/sendMessage', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ chat_id: cfg.chatId, text: tgMsg, parse_mode: 'Markdown' })
      }).catch(function (e) { console.warn('[APF Digest] Telegram error:', e); });
    }

    // ── Email (Google Apps Script) ──
    var gasUrl  = _rsGet('emailGasUrl');
    var emailTo = _rsGet('emailTo');
    if (gasUrl && emailTo) {
      var emailBody = 'Good Morning!\n\nYou have ' + count + ' unfinished reminder' + (count > 1 ? 's' : '') + ':\n\n';
      unfinished.forEach(function (u, i) {
        var dt = u.rem.remindAt ? new Date(u.rem.remindAt).toLocaleString('en-IN') : '';
        emailBody += (i + 1) + '. ' + (u.note.title || 'Untitled') + '\n';
        if (u.note.content) emailBody += '   ' + u.note.content.substring(0, 100) + '\n';
        if (dt) emailBody += '   Scheduled: ' + dt + '\n';
        emailBody += '\n';
      });
      emailBody += 'Open your APF Dashboard \u2192 Quick Notes to mark items as done.';

      fetch(gasUrl, {
        method: 'POST',
        mode:   'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          to:      emailTo,
          subject: '\uD83C\uDF05 APF Morning Digest \u2014 ' + count + ' unfinished reminder' + (count > 1 ? 's' : ''),
          body:    emailBody
        })
      }).catch(function (e) { console.warn('[APF Digest] Email error:', e); });
    }
  }

})(); // end countdown/digest IIFE
