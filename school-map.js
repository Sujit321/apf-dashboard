// =====================================================================
// SCHOOL MAP MODULE — APF Dashboard
// =====================================================================
const SchoolMap = {
  LOC_KEY: 'apf_school_locations',
  _map: null,
  _previewMap: null,
  _markers: [],
  _circles: [],
  _editingSchool: null,
  _previewMarker: null,

  // --- Persistence ---
  getAll() {
    try { return JSON.parse(localStorage.getItem(this.LOC_KEY) || '{}'); } catch (e) { return {}; }
  },
  save(locations) {
    localStorage.setItem(this.LOC_KEY, JSON.stringify(locations));
  },
  get(name) { return this.getAll()[name] || null; },
  upsert(name, data) {
    const all = this.getAll();
    all[name] = Object.assign({}, all[name], data, { school: name, updatedAt: new Date().toISOString() });
    this.save(all);
  },
  remove(name) {
    const all = this.getAll();
    delete all[name];
    this.save(all);
  },

  // --- Get all known school names from School Profiles (visits + observations) ---
  getKnownSchools() {
    // Prefer getSchoolData() which aggregates visits + observations (same as School Profiles section)
    if (typeof getSchoolData === 'function') {
      try {
        const schoolMap = getSchoolData();
        const profileSchools = Object.values(schoolMap)
          .map(function(s) { return { name: s.name, cluster: s.cluster || s.block || '' }; })
          .filter(function(s) { return s.name; })
          .sort(function(a, b) { return a.name.localeCompare(b.name); });
        // Also include any pinned schools not yet in profiles
        const locs = this.getAll();
        Object.keys(locs).forEach(function(pinName) {
          if (!profileSchools.find(function(s) { return s.name.toLowerCase() === pinName.toLowerCase(); })) {
            profileSchools.push({ name: pinName, cluster: locs[pinName].cluster || '' });
          }
        });
        return profileSchools;
      } catch (e) { /* fall through */ }
    }
    // Fallback: derive from raw visits + observations
    const visits = (typeof DB !== 'undefined' ? DB.get('visits') : null) || [];
    const obs    = (typeof DB !== 'undefined' ? DB.get('observations') : null) || [];
    const map = {};
    var addEntry = function(name, cluster) {
      if (!name) return;
      var key = name.toLowerCase();
      if (!map[key]) map[key] = { name: name, cluster: cluster || '' };
      else if (!map[key].cluster && cluster) map[key].cluster = cluster;
    };
    visits.forEach(function(v) { addEntry(v.school || v.schoolName, v.cluster || v.block || ''); });
    obs.forEach(function(o)    { addEntry(o.school, o.cluster || o.block || ''); });
    var locs = this.getAll();
    Object.keys(locs).forEach(function(n) { addEntry(n, locs[n].cluster || ''); });
    return Object.values(map).sort(function(a, b) { return a.name.localeCompare(b.name); });
  },

  // --- Visit stats for a school (reads visits + observations) ---
  visitStats(schoolName) {
    const visits = (typeof DB !== 'undefined' ? DB.get('visits') : null) || [];
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split('T')[0];
    const ln = schoolName.toLowerCase();
    const schoolVisits = visits.filter(function(v) {
      return (v.school || v.schoolName || '').toLowerCase() === ln;
    });
    const thisMonth = schoolVisits.filter(function(v) { return (v.date || '') >= monthStart; });
    const sorted = schoolVisits.slice().sort(function(a, b) { return (b.date || '') < (a.date || '') ? -1 : 1; });
    const lastVisit = sorted[0] || null;
    // Also check observations for cluster info
    const obs = (typeof DB !== 'undefined' ? DB.get('observations') : null) || [];
    const schoolObs = obs.filter(function(o) { return (o.school || '').toLowerCase() === ln; });
    var cluster = (lastVisit && (lastVisit.cluster || lastVisit.block || '')) || '';
    if (!cluster && schoolObs.length) cluster = schoolObs[0].cluster || schoolObs[0].block || '';
    return {
      total: schoolVisits.length,
      thisMonth: thisMonth.length,
      lastVisitDate: lastVisit ? (lastVisit.date || '') : null,
      cluster: cluster
    };
  },

  // --- Pin color ---
  pinColor(schoolName) {
    const s = this.visitStats(schoolName);
    if (s.thisMonth > 0) return 'green';
    if (s.total > 0) return 'yellow';
    return 'red';
  },

  // --- Custom SVG map pin icon ---
  makeIcon(color) {
    const fill   = { green: '#22c55e', yellow: '#f59e0b', red: '#ef4444' };
    const glow   = { green: 'rgba(34,197,94,0.6)', yellow: 'rgba(245,158,11,0.5)', red: 'rgba(239,68,68,0.6)' };
    const hex = fill[color] || fill.red;
    const sh  = glow[color] || glow.red;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="34" height="44" viewBox="0 0 34 44">
      <ellipse cx="17" cy="42" rx="7" ry="3" fill="rgba(0,0,0,0.22)"/>
      <path d="M17 1C10.37 1 5 6.37 5 13c0 10 12 27 12 27s12-17 12-27C29 6.37 23.63 1 17 1z"
            fill="${hex}" stroke="${sh}" stroke-width="1.5"/>
      <circle cx="17" cy="13" r="5.5" fill="rgba(255,255,255,0.9)"/>
    </svg>`;
    return L.divIcon({
      html: svg,
      className: '',
      iconSize: [34, 44],
      iconAnchor: [17, 44],
      popupAnchor: [0, -42]
    });
  },

  // --- Popup HTML ---
  popupHtml(name, stats, locData) {
    const safeName = name.replace(/'/g, "\\'");
    return `<div class="smap-popup">
      <div class="smap-popup-name">🏫 ${name}</div>
      <div class="smap-popup-meta">
        <div><b>Cluster:</b> ${locData.cluster || stats.cluster || '—'}</div>
        <div><b>Total Visits:</b> ${stats.total}</div>
        <div><b>This Month:</b> ${stats.thisMonth}</div>
        <div><b>Last Visit:</b> ${stats.lastVisitDate || 'Never'}</div>
        <div><b>Coords:</b> ${Number(locData.lat).toFixed(5)}, ${Number(locData.lng).toFixed(5)}</div>
        ${locData.plusCode ? `<div><b>Plus Code:</b> ${locData.plusCode}</div>` : ''}
      </div>
      <div class="smap-popup-actions">
        <button class="btn btn-outline btn-sm" onclick="openSchoolLocationEditor('${safeName}')">
          <i class="fas fa-edit"></i> Edit
        </button>
        <button class="btn btn-outline btn-sm" onclick="smapShowRoute('${safeName}')" style="color:#a5b4fc;border-color:rgba(99,102,241,0.4);">
          <i class="fas fa-route"></i> Route
        </button>
      </div>
    </div>`;
  }
};

// ---- Main map init ----
function initSchoolMap() {
  if (typeof L === 'undefined') {
    const el = document.getElementById('apf-school-map');
    if (el) el.innerHTML = '<div style="padding:48px;text-align:center;color:#ef4444;">' +
      '<i class="fas fa-exclamation-triangle" style="font-size:28px;display:block;margin-bottom:12px;"></i>' +
      'Leaflet.js not loaded. Please check your internet connection and refresh the page.</div>';
    return;
  }

  if (!SchoolMap._map) {
    var home = smapGetHome();
    SchoolMap._map = L.map('apf-school-map', {
      center: [home.lat, home.lng],
      zoom: home.zoom || 11,
      zoomControl: true
    });
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
      maxZoom: 19,
      crossOrigin: true
    }).addTo(SchoolMap._map);
  } else {
    setTimeout(function() { SchoolMap._map.invalidateSize(); }, 150);
  }

  _smapInitHome();
  renderSchoolMap();
  renderSchoolMapTable();
}

// ---- Render all pins ----
function renderSchoolMap() {
  const map = SchoolMap._map;
  if (!map) return;

  SchoolMap._markers.forEach(function(m) { map.removeLayer(m); });
  SchoolMap._circles.forEach(function(c) { map.removeLayer(c); });
  SchoolMap._markers = [];
  SchoolMap._circles = [];

  const locations = SchoolMap.getAll();
  const names = Object.keys(locations);
  var green = 0, yellow = 0, red = 0;

  names.forEach(function(name) {
    var loc = locations[name];
    if (!loc.lat || !loc.lng) return;
    var stats = SchoolMap.visitStats(name);
    var color = SchoolMap.pinColor(name);
    if (color === 'green') green++;
    else if (color === 'yellow') yellow++;
    else red++;

    var colorHex = color === 'green' ? '#22c55e' : color === 'yellow' ? '#f59e0b' : '#ef4444';
    var radius = Math.max(250, Math.min(stats.total * 180, 1500));

    var circle = L.circle([loc.lat, loc.lng], {
      radius: radius,
      color: colorHex,
      fillColor: colorHex,
      fillOpacity: 0.09,
      weight: 1,
      opacity: 0.35
    }).addTo(map);
    SchoolMap._circles.push(circle);

    var marker = L.marker([loc.lat, loc.lng], { icon: SchoolMap.makeIcon(color) })
      .bindPopup(SchoolMap.popupHtml(name, stats, loc), { maxWidth: 300 });
    marker.addTo(map);
    SchoolMap._markers.push(marker);
  });

  // Update stats bar
  function setEl(id, val) {
    var el = document.getElementById(id);
    if (el) el.textContent = val;
  }
  setEl('smapStatGreen', green);
  setEl('smapStatYellow', yellow);
  setEl('smapStatRed', red);
  setEl('smapStatTotal', names.length);

  // Populate region filter from all known clusters
  _schoolMapPopulateRegions();
}

// ---- Populate region dropdown ----
function _schoolMapPopulateRegions() {
  var sel = document.getElementById('smapRegionFilter');
  if (!sel) return;
  var current = sel.value;
  var knownList = SchoolMap.getKnownSchools();
  var clusters = [];
  knownList.forEach(function(s) { if (s.cluster && !clusters.includes(s.cluster)) clusters.push(s.cluster); });
  // Also read from pinned locations
  var locs = SchoolMap.getAll();
  Object.values(locs).forEach(function(l) { if (l.cluster && !clusters.includes(l.cluster)) clusters.push(l.cluster); });
  clusters.sort();
  sel.innerHTML = '<option value="all">All Regions</option>' +
    clusters.map(function(c) { return '<option value="' + c + '">' + c + '</option>'; }).join('');
  if (current && clusters.includes(current)) sel.value = current;
}

// ---- Filter map by region ----
function schoolMapFilterRegion(region) {
  var map = SchoolMap._map;
  if (!map) return;
  var locs = SchoolMap.getAll();
  var knownList = SchoolMap.getKnownSchools();

  SchoolMap._markers.forEach(function(m) { map.removeLayer(m); });
  SchoolMap._circles.forEach(function(c) { map.removeLayer(c); });
  SchoolMap._markers = [];
  SchoolMap._circles = [];

  var matchingLatLngs = [];

  Object.keys(locs).forEach(function(name) {
    var loc = locs[name];
    if (!loc.lat || !loc.lng) return;
    var clusterEntry = knownList.find(function(s) { return s.name.toLowerCase() === name.toLowerCase(); });
    var cluster = loc.cluster || (clusterEntry && clusterEntry.cluster) || '';

    if (region !== 'all' && cluster !== region) return; // Skip if filtered out

    var stats = SchoolMap.visitStats(name);
    var color = SchoolMap.pinColor(name);
    var colorHex = color === 'green' ? '#22c55e' : color === 'yellow' ? '#f59e0b' : '#ef4444';
    var radius = Math.max(250, Math.min(stats.total * 180, 1500));

    var circle = L.circle([loc.lat, loc.lng], {
      radius: radius, color: colorHex, fillColor: colorHex,
      fillOpacity: 0.09, weight: 1, opacity: 0.35
    }).addTo(map);
    SchoolMap._circles.push(circle);

    var marker = L.marker([loc.lat, loc.lng], { icon: SchoolMap.makeIcon(color) })
      .bindPopup(SchoolMap.popupHtml(name, stats, loc), { maxWidth: 300 });
    marker.addTo(map);
    SchoolMap._markers.push(marker);
    matchingLatLngs.push([loc.lat, loc.lng]);
  });

  if (region !== 'all' && matchingLatLngs.length > 0) {
    map.fitBounds(L.latLngBounds(matchingLatLngs), { padding: [60, 60] });
  } else if (region === 'all') {
    schoolMapFitAll();
  }

  if (typeof showToast === 'function') {
    showToast(region === 'all' ? 'Showing all regions' : ('Region: ' + region + ' (' + matchingLatLngs.length + ' schools)'), 'info');
  }
}

// ---- Hotspot layer ----
SchoolMap._hotspotLayers = [];
SchoolMap._hotspotOn = false;

function schoolMapToggleHotspot(on) {
  var map = SchoolMap._map;
  if (!map) return;
  SchoolMap._hotspotOn = on;

  // Remove existing hotspot circles
  SchoolMap._hotspotLayers.forEach(function(l) { map.removeLayer(l); });
  SchoolMap._hotspotLayers = [];

  var legendEl = document.getElementById('smapLegendHotspot');
  if (legendEl) legendEl.style.display = on ? '' : 'none';

  // When hotspot is ON: hide pins + normal circles, show only hotspot circles
  // When OFF: restore pins + normal circles
  if (on) {
    SchoolMap._markers.forEach(function(m) { map.removeLayer(m); });
    SchoolMap._circles.forEach(function(c) { map.removeLayer(c); });
  } else {
    // Restore pins and circles
    SchoolMap._markers.forEach(function(m) { m.addTo(map); });
    SchoolMap._circles.forEach(function(c) { c.addTo(map); });
    return;
  }

  var locs = SchoolMap.getAll();
  var maxVisits = 1;

  // Find max visit count for normalization
  Object.keys(locs).forEach(function(name) {
    var s = SchoolMap.visitStats(name);
    if (s.total > maxVisits) maxVisits = s.total;
  });

  // Draw gradient hotspot circles for each school
  Object.keys(locs).forEach(function(name) {
    var loc = locs[name];
    if (!loc.lat || !loc.lng) return;
    var stats = SchoolMap.visitStats(name);

    var intensity = maxVisits > 0 ? stats.total / maxVisits : 0; // 0-1
    // Color: blue (low/0) → green → yellow → red (high)
    var r, g, b;
    if (stats.total === 0) {
      r = 100; g = 116; b = 139; // gray for never visited
    } else if (intensity < 0.33) {
      var t = intensity / 0.33;
      r = Math.round(59 + t * (34 - 59));
      g = Math.round(130 + t * (197 - 130));
      b = Math.round(246 + t * (94 - 246));
    } else if (intensity < 0.66) {
      var t = (intensity - 0.33) / 0.33;
      r = Math.round(34 + t * (245 - 34));
      g = Math.round(197 + t * (158 - 197));
      b = Math.round(94 + t * (11 - 94));
    } else {
      var t = (intensity - 0.66) / 0.34;
      r = Math.round(245 + t * (239 - 245));
      g = Math.round(158 + t * (68 - 158));
      b = Math.round(11 + t * (68 - 11));
    }
    var hex = '#' + ('0' + r.toString(16)).slice(-2) + ('0' + g.toString(16)).slice(-2) + ('0' + b.toString(16)).slice(-2);

    // Circle radius based on visit count
    var outerR = stats.total === 0 ? 400 : Math.max(600, Math.min(stats.total * 500, 8000));
    var outer = L.circle([loc.lat, loc.lng], {
      radius: outerR,
      color: hex,
      fillColor: hex,
      fillOpacity: stats.total === 0 ? 0.05 : (0.08 + intensity * 0.12),
      weight: 0,
      opacity: 0
    }).addTo(map);

    // Inner core
    var innerR = stats.total === 0 ? 200 : Math.max(200, Math.min(stats.total * 150, 3000));
    var inner = L.circle([loc.lat, loc.lng], {
      radius: innerR,
      color: hex,
      fillColor: hex,
      fillOpacity: stats.total === 0 ? 0.12 : (0.22 + intensity * 0.18),
      weight: 1.5,
      opacity: 0.5
    }).addTo(map)
      .bindTooltip('<b>' + name + '</b><br>' +
        (stats.total === 0 ? '<span style="color:#ef4444">Never visited</span>' :
         stats.total + ' visit' + (stats.total > 1 ? 's' : '') + (stats.thisMonth > 0 ? ' (' + stats.thisMonth + ' this month)' : '')),
        { permanent: false, direction: 'top', className: 'smap-hotspot-tip' });

    SchoolMap._hotspotLayers.push(outer, inner);
  });

  if (typeof showToast === 'function') showToast('🔥 Hotspot mode: showing visit intensity circles', 'info');
}

// ---- Search autocomplete ----
function schoolMapSearchInput(query) {
  var drop = document.getElementById('smapSearchDrop');
  if (!drop) return;

  var q = (query || '').trim().toLowerCase();
  if (!q) { drop.style.display = 'none'; return; }

  var locs = SchoolMap.getAll();
  var knownList = SchoolMap.getKnownSchools();

  // Search across all known schools (not just pinned)
  var results = knownList.filter(function(s) {
    return s.name.toLowerCase().includes(q) || (s.cluster || '').toLowerCase().includes(q);
  }).slice(0, 10);

  if (!results.length) {
    drop.innerHTML = '<div class="smap-drop-empty"><i class="fas fa-search" style="opacity:.4;display:block;margin-bottom:6px;"></i>No schools found</div>';
    drop.style.display = 'block';
    return;
  }

  drop.innerHTML = results.map(function(s) {
    var loc = locs[s.name] || null;
    var hasPinned = loc && loc.lat && loc.lng;
    var color = hasPinned ? SchoolMap.pinColor(s.name) : null;
    var dotColor = color === 'green' ? '#22c55e' : color === 'yellow' ? '#f59e0b' : color === 'red' ? '#ef4444' : '#475569';
    var safeName = s.name.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    return '<div class="smap-drop-item" onclick="schoolMapSelectResult(\'' + safeName + '\')">' +
      '<span class="smap-drop-item-dot" style="background:' + dotColor + ';"></span>' +
      '<span class="smap-drop-item-name">' + s.name + '</span>' +
      '<span class="smap-drop-item-meta">' + (s.cluster ? s.cluster + ' • ' : '') + (hasPinned ? '📍 Pinned' : 'Not pinned') + '</span>' +
    '</div>';
  }).join('');
  drop.style.display = 'block';
}

function schoolMapSelectResult(schoolName) {
  var drop = document.getElementById('smapSearchDrop');
  var inp = document.getElementById('smapSearchInput');
  if (inp) inp.value = schoolName;
  if (drop) drop.style.display = 'none';

  var loc = SchoolMap.get(schoolName);
  if (!loc || !loc.lat) {
    if (typeof showToast === 'function') showToast('"' + schoolName + '" has no pin yet. Click Pin to add location.', 'info');
    return;
  }
  if (!SchoolMap._map) return;
  SchoolMap._map.flyTo([loc.lat, loc.lng], 15, { duration: 1.3 });
  var marker = SchoolMap._markers.find(function(m) {
    var ll = m.getLatLng();
    return Math.abs(ll.lat - loc.lat) < 0.0002 && Math.abs(ll.lng - loc.lng) < 0.0002;
  });
  if (marker) setTimeout(function() { marker.openPopup(); }, 1400);
}

// Keep old schoolMapSearch as alias for backwards compat
function schoolMapSearch(query) { schoolMapSearchInput(query); }

// ---- Export map as image ----
function exportSchoolMapImage() {
  var map = SchoolMap._map;
  if (!map) { if (typeof showToast === 'function') showToast('Open the map first', 'info'); return; }

  if (typeof showToast === 'function') showToast('Generating map image…', 'info');

  var mapContainer = map.getContainer();
  var mapRect = mapContainer.getBoundingClientRect();
  var w = mapRect.width;
  var h = mapRect.height;
  var dpr = 2;

  // ---- Snapshot all positions synchronously RIGHT NOW ----
  // Tiles
  var tileDraw = [];
  mapContainer.querySelectorAll('.leaflet-tile-loaded').forEach(function(tile) {
    var r = tile.getBoundingClientRect();
    tileDraw.push({
      src: tile.src,
      x: r.left - mapRect.left,
      y: r.top - mapRect.top,
      w: r.width,
      h: r.height
    });
  });

  // Markers — read positions from actual DOM elements (same coord space as tiles)
  // Icon anchor is [16, 40] so the pin-tip = element top-left + (16, 40)
  var markerDraw = [];
  SchoolMap._markers.forEach(function(marker) {
    var el = marker.getElement ? marker.getElement() : null;
    if (!el) return;
    var er = el.getBoundingClientRect();
    // Pin-tip position (the actual lat/lng point on the map)
    var px = er.left + 16 - mapRect.left;
    var py = er.top + 40 - mapRect.top;
    var ll = marker.getLatLng();
    // Find the school name from saved locations
    var locs = SchoolMap.getAll();
    var name = Object.keys(locs).find(function(n) {
      var loc = locs[n];
      return loc && Math.abs(loc.lat - ll.lat) < 0.0002 && Math.abs(loc.lng - ll.lng) < 0.0002;
    }) || '';
    var color = name ? SchoolMap.pinColor(name) : 'red';
    markerDraw.push({ px: px, py: py, color: color, name: name });
  });

  // Circles — also from DOM
  var circleDraw = [];
  SchoolMap._circles.forEach(function(circle) {
    var el = circle.getElement ? circle.getElement() : null;
    var ll = circle.getLatLng();
    var cp = { x: 0, y: 0 };
    // Match to a marker to get correct position
    var match = markerDraw.find(function(md) { return md.name; });
    // Use map projection for circles (they don't have a simple DOM rect)
    var locs = SchoolMap.getAll();
    var name = Object.keys(locs).find(function(n) {
      var loc = locs[n];
      return loc && Math.abs(loc.lat - ll.lat) < 0.0002 && Math.abs(loc.lng - ll.lng) < 0.0002;
    }) || '';
    // Use the marker's pin-tip position for circle center
    var md = markerDraw.find(function(m) { return m.name === name; });
    if (md) {
      cp.x = md.px;
      cp.y = md.py;
    }
    var color = name ? SchoolMap.pinColor(name) : 'red';
    var colorHex = color === 'green' ? '#22c55e' : color === 'yellow' ? '#f59e0b' : '#ef4444';
    var stats = name ? SchoolMap.visitStats(name) : { total: 0 };
    var radiusM = Math.max(250, Math.min(stats.total * 180, 1500));
    var zoom = map.getZoom();
    var lat = ll.lat || 20;
    var radiusPx = Math.max(8, radiusM / (156543.03 * Math.cos(lat * Math.PI / 180) / Math.pow(2, zoom)));
    circleDraw.push({ x: cp.x, y: cp.y, r: radiusPx, color: colorHex });
  });

  // ---- Now build the canvas ----
  var canvas = document.createElement('canvas');
  canvas.width = w * dpr;
  canvas.height = h * dpr;
  var ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#0f1117';
  ctx.fillRect(0, 0, w, h);

  // Step 1: Load tiles with CORS and draw
  var tilePromises = tileDraw.map(function(td) {
    return new Promise(function(resolve) {
      var img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = function() {
        try { ctx.drawImage(img, td.x, td.y, td.w, td.h); } catch(e) {}
        resolve();
      };
      img.onerror = function() { resolve(); };
      img.src = td.src;
    });
  });

  Promise.all(tilePromises).then(function() {
    // Step 2: Draw intensity circles
    circleDraw.forEach(function(cd) {
      if (!cd.x && !cd.y) return;
      ctx.beginPath();
      ctx.arc(cd.x, cd.y, cd.r, 0, Math.PI * 2);
      ctx.fillStyle = cd.color;
      ctx.globalAlpha = 0.12;
      ctx.fill();
      ctx.globalAlpha = 1.0;
    });

    // Step 3: Draw pin markers at their exact DOM positions
    markerDraw.forEach(function(md) {
      var px = md.px, py = md.py;
      var colorHex = md.color === 'green' ? '#22c55e' : md.color === 'yellow' ? '#f59e0b' : '#ef4444';

      // Pin shadow
      ctx.beginPath();
      ctx.ellipse(px, py + 1, 5, 2, 0, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(0,0,0,0.25)';
      ctx.fill();

      // Pin body (teardrop)
      ctx.beginPath();
      ctx.moveTo(px, py);
      ctx.bezierCurveTo(px - 8, py - 10, px - 10, py - 22, px, py - 26);
      ctx.bezierCurveTo(px + 10, py - 22, px + 8, py - 10, px, py);
      ctx.fillStyle = colorHex;
      ctx.fill();
      ctx.strokeStyle = '#fff';
      ctx.lineWidth = 1.5;
      ctx.stroke();

      // Inner circle
      ctx.beginPath();
      ctx.arc(px, py - 18, 5, 0, Math.PI * 2);
      ctx.fillStyle = '#fff';
      ctx.fill();

      // Plus icon
      ctx.fillStyle = colorHex;
      ctx.font = 'bold 9px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('+', px, py - 17.5);
    });

    // Step 4: Hotspot overlay if active
    if (SchoolMap._hotspotOn && SchoolMap._hotspotLayers.length) {
      var locs = SchoolMap.getAll();
      var maxV = 1;
      Object.keys(locs).forEach(function(n) {
        var s = SchoolMap.visitStats(n); if (s.total > maxV) maxV = s.total;
      });
      markerDraw.forEach(function(md) {
        if (!md.name) return;
        var stats = SchoolMap.visitStats(md.name);
        if (stats.total === 0) return;
        var intensity = stats.total / maxV;
        var zoom = map.getZoom();
        var loc = locs[md.name];
        if (!loc) return;
        var outerR = Math.max(600, Math.min(stats.total * 500, 8000));
        var radiusPx = Math.max(15, outerR / (156543.03 * Math.cos(loc.lat * Math.PI / 180) / Math.pow(2, zoom)));

        var gradient = ctx.createRadialGradient(md.px, md.py, 0, md.px, md.py, radiusPx);
        var r = intensity < 0.5 ? Math.round(59 + intensity * 2 * 186) : Math.round(245 - (intensity - 0.5) * 2 * 6);
        var g = intensity < 0.5 ? Math.round(130 + intensity * 2 * 67) : Math.round(197 - (intensity - 0.5) * 2 * 129);
        var b = intensity < 0.5 ? Math.round(246 - intensity * 2 * 152) : Math.round(94 - (intensity - 0.5) * 2 * 26);
        var hex = 'rgb(' + r + ',' + g + ',' + b + ')';
        gradient.addColorStop(0, hex);
        gradient.addColorStop(1, 'rgba(' + r + ',' + g + ',' + b + ',0)');
        ctx.beginPath();
        ctx.arc(md.px, md.py, radiusPx, 0, Math.PI * 2);
        ctx.fillStyle = gradient;
        ctx.globalAlpha = 0.3;
        ctx.fill();
        ctx.globalAlpha = 1.0;
      });
    }

    // Step 5: Legend box
    var lx = w - 175, ly = h - 90, lw = 160, lh = 78;
    ctx.fillStyle = 'rgba(15, 17, 23, 0.85)';
    ctx.beginPath();
    ctx.roundRect(lx, ly, lw, lh, 8);
    ctx.fill();
    ctx.strokeStyle = 'rgba(255,255,255,0.08)';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.font = '600 11px Inter, sans-serif';
    ctx.fillStyle = '#94a3b8';
    ctx.textAlign = 'left';
    ctx.fillText('Legend', lx + 10, ly + 16);
    var legendItems = [
      { color: '#22c55e', text: 'Visited this month' },
      { color: '#f59e0b', text: 'Visited before' },
      { color: '#ef4444', text: 'Never visited' }
    ];
    legendItems.forEach(function(item, i) {
      var iy = ly + 30 + i * 16;
      ctx.beginPath();
      ctx.arc(lx + 16, iy, 4, 0, Math.PI * 2);
      ctx.fillStyle = item.color;
      ctx.fill();
      ctx.fillStyle = '#cbd5e1';
      ctx.font = '11px Inter, sans-serif';
      ctx.fillText(item.text, lx + 26, iy + 4);
    });

    // Step 6: Watermark
    ctx.font = '11px Inter, sans-serif';
    ctx.textAlign = 'left';
    var wmText = 'APF Dashboard • School Map • ' + new Date().toLocaleDateString();
    var wmW = ctx.measureText(wmText).width + 20;
    ctx.fillStyle = 'rgba(17, 24, 39, 0.8)';
    ctx.beginPath();
    ctx.roundRect(10, h - 28, wmW, 20, 5);
    ctx.fill();
    ctx.fillStyle = '#94a3b8';
    ctx.fillText(wmText, 20, h - 14);

    // Step 7: Download
    var link = document.createElement('a');
    var now = new Date();
    link.download = 'school-map-' + now.getFullYear() + '-' +
      String(now.getMonth() + 1).padStart(2, '0') + '-' +
      String(now.getDate()).padStart(2, '0') + '.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
    if (typeof showToast === 'function') showToast('✅ Map exported as PNG', 'success');

  }).catch(function(err) {
    if (typeof showToast === 'function') showToast('Export failed: ' + err.message, 'error');
  });
}



// ---- Fit all pins in view ----
function schoolMapFitAll() {
  var map = SchoolMap._map;
  if (!map) return;
  var locs = SchoolMap.getAll();
  var latLngs = Object.values(locs).filter(function(l) { return l.lat && l.lng; }).map(function(l) { return [l.lat, l.lng]; });
  if (!latLngs.length) { if (typeof showToast === 'function') showToast('No pinned schools yet', 'info'); return; }
  map.fitBounds(L.latLngBounds(latLngs), { padding: [48, 48] });
}

// ---- Search and fly to a school ----
function schoolMapSearch(query) {
  if (!query || !query.trim()) return;
  var q = query.toLowerCase();
  var locs = SchoolMap.getAll();
  var match = Object.keys(locs).find(function(n) { return n.toLowerCase().includes(q); });
  if (!match) { if (typeof showToast === 'function') showToast('School not found on map', 'info'); return; }
  var loc = locs[match];
  if (!loc.lat) { if (typeof showToast === 'function') showToast(match + ' has no pin yet', 'info'); return; }
  SchoolMap._map.flyTo([loc.lat, loc.lng], 14, { duration: 1.3 });
  var marker = SchoolMap._markers.find(function(m) {
    var ll = m.getLatLng();
    return Math.abs(ll.lat - loc.lat) < 0.0002 && Math.abs(ll.lng - loc.lng) < 0.0002;
  });
  if (marker) setTimeout(function() { marker.openPopup(); }, 1400);
}

// ---- Fly to user GPS location ----
function schoolMapFlyToMyLocation() {
  if (!navigator.geolocation) { if (typeof showToast === 'function') showToast('GPS not available in this browser', 'error'); return; }
  if (typeof showToast === 'function') showToast('Getting your GPS location…', 'info');
  navigator.geolocation.getCurrentPosition(function(pos) {
    SchoolMap._map.flyTo([pos.coords.latitude, pos.coords.longitude], 15, { duration: 1.5 });
    L.marker([pos.coords.latitude, pos.coords.longitude], {
      icon: L.divIcon({
        html: '<div style="width:14px;height:14px;background:#6366f1;border:3px solid white;border-radius:50%;box-shadow:0 0 10px rgba(99,102,241,0.8)"></div>',
        className: '',
        iconSize: [14, 14],
        iconAnchor: [7, 7]
      })
    }).addTo(SchoolMap._map).bindPopup('<b>📍 Your Current Location</b>').openPopup();
  }, function(err) {
    if (typeof showToast === 'function') showToast('GPS error: ' + err.message, 'error');
  });
}

// ---- Home Base system ----
var SMAP_HOME_KEY = 'apf_map_home';
var SMAP_HOME_DEFAULT = {
  block: 'Magarlod', district: 'Durg', state: 'Chhattisgarh',
  lat: 21.1793, lng: 81.2833, zoom: 11
};

function smapGetHome() {
  try {
    var stored = localStorage.getItem(SMAP_HOME_KEY);
    if (stored) return Object.assign({}, SMAP_HOME_DEFAULT, JSON.parse(stored));
  } catch(e) {}
  return Object.assign({}, SMAP_HOME_DEFAULT);
}
function smapHomeUseGPS() {
  if (!navigator.geolocation) {
    if (typeof showToast === 'function') showToast('Geolocation is not supported by your browser.', 'error');
    return;
  }
  if (typeof showToast === 'function') showToast('Getting your location...', 'info');
  navigator.geolocation.getCurrentPosition(function(position) {
    var lat = position.coords.latitude;
    var lng = position.coords.longitude;
    var latInput = document.getElementById('smapHomeLat');
    var lngInput = document.getElementById('smapHomeLng');
    if (latInput) latInput.value = lat.toFixed(5);
    if (lngInput) lngInput.value = lng.toFixed(5);
    
    // Reverse Geocode to get Block, District, State
    if (typeof showToast === 'function') showToast('GPS found. Fetching region names...', 'info');
    fetch('https://nominatim.openstreetmap.org/reverse?format=json&lat=' + lat + '&lon=' + lng)
      .then(function(res) { return res.json(); })
      .then(function(data) {
        if (data && data.address) {
          var blockInput = document.getElementById('smapHomeBlock');
          var distInput  = document.getElementById('smapHomeDist');
          var stateInput = document.getElementById('smapHomeState');
          
          var addr = data.address;
          var block = addr.village || addr.suburb || addr.town || addr.city_district || addr.county || '';
          var dist = addr.state_district || addr.county || addr.city || '';
          var state = addr.state || '';
          
          if (blockInput && block) blockInput.value = block.replace(/ Tehsil/gi, '');
          if (distInput && dist) distInput.value = dist.replace(/ District/gi, '');
          if (stateInput && state) stateInput.value = state;
          
          if (typeof showToast === 'function') showToast('Region auto-filled! Click Save.', 'success');
        }
      })
      .catch(function(err) {
        console.error('Reverse geocode error:', err);
        if (typeof showToast === 'function') showToast('GPS retrieved! Click Save.', 'success');
      });
      
  }, function(error) {
    var msg = 'Could not get location.';
    if (error.code === error.PERMISSION_DENIED) msg = 'Location permission denied.';
    if (typeof showToast === 'function') showToast(msg, 'error');
  }, { enableHighAccuracy: true });
}

function smapSaveHome() {
  var block = (document.getElementById('smapHomeBlock') || {}).value || '';
  var dist  = (document.getElementById('smapHomeDist')  || {}).value || '';
  var state = (document.getElementById('smapHomeState') || {}).value || '';
  var lat   = parseFloat((document.getElementById('smapHomeLat') || {}).value);
  var lng   = parseFloat((document.getElementById('smapHomeLng') || {}).value);

  if (isNaN(lat) || isNaN(lng) || lat < -90 || lat > 90 || lng < -180 || lng > 180) {
    if (typeof showToast === 'function') showToast('Enter valid latitude and longitude', 'error');
    return;
  }
  var home = { block: block, district: dist, state: state, lat: lat, lng: lng, zoom: 11 };
  localStorage.setItem(SMAP_HOME_KEY, JSON.stringify(home));
  _smapUpdateHomeButton(home);
  _smapUpdatePlusCodeRef(home);
  if (typeof showToast === 'function') showToast('✅ Home base saved: ' + [block, dist, state].filter(Boolean).join(', '), 'success');

  // Fly to the new home
  if (SchoolMap._map) SchoolMap._map.flyTo([lat, lng], home.zoom || 11, { duration: 1.5 });
}

function _smapUpdateHomeButton(home) {
  var lbl = document.getElementById('smapHomeBtnLabel');
  if (lbl) {
    var parts = [home.block, home.district].filter(Boolean);
    lbl.textContent = parts.length ? parts.join(', ') : 'Home Area';
  }
  // Update region label overlay on map
  var rlabel = document.getElementById('smapRegionLabel');
  var rtext  = document.getElementById('smapRegionLabelText');
  if (rlabel && rtext) {
    var parts2 = [home.block, home.district, home.state].filter(Boolean);
    if (parts2.length) {
      rtext.innerHTML = '<i class="fas fa-map-marker-alt" style="color:#6366f1;"></i> <span>' + parts2.join(' · ') + '</span>';
      rlabel.style.display = 'block';
    } else {
      rlabel.style.display = 'none';
    }
  }
  // Update/place home base marker
  _smapPlaceHomeMarker(home);
}

function _smapUpdatePlusCodeRef(home) {
  var el = document.getElementById('slmPlusCodeHomeRef');
  if (!el) return;
  var loc = [home.block, home.district, home.state].filter(Boolean).join(', ');
  el.innerHTML = '<i class="fas fa-home" style="margin-right:4px;"></i>Short codes decoded relative to: <b>' + (loc || 'India center') + '</b> (' + home.lat + ', ' + home.lng + ')';
}

function smapToggleHomePanel() {
  var panel = document.getElementById('smapHomePanel');
  if (!panel) return;
  var showing = panel.style.display !== 'none';
  panel.style.display = showing ? 'none' : 'block';
  var cog = document.getElementById('smapHomeCogBtn');
  if (cog) cog.style.background = showing ? '' : 'rgba(99,102,241,0.15)';
  if (!showing) {
    // Populate panel with current saved values
    var home = smapGetHome();
    var setVal = function(id, v) { var el = document.getElementById(id); if (el) el.value = v; };
    setVal('smapHomeBlock', home.block || '');
    setVal('smapHomeDist',  home.district || '');
    setVal('smapHomeState', home.state || '');
    setVal('smapHomeLat',   home.lat || '');
    setVal('smapHomeLng',   home.lng || '');
  }
}

function schoolMapGoHome() {
  var map = SchoolMap._map;
  if (!map) { if (typeof showToast === 'function') showToast('Open the map first', 'info'); return; }
  var home = smapGetHome();
  map.flyTo([home.lat, home.lng], home.zoom || 11, { duration: 1.5 });
  var loc = [home.block, home.district, home.state].filter(Boolean).join(', ');
  if (typeof showToast === 'function') showToast('🏠 Navigating to ' + (loc || 'home base'), 'info');
}

function _smapInitHome() {
  var home = smapGetHome();
  var setVal = function(id, v) { var el = document.getElementById(id); if (el) el.value = v; };
  setVal('smapHomeBlock', home.block || '');
  setVal('smapHomeDist',  home.district || '');
  setVal('smapHomeState', home.state || '');
  setVal('smapHomeLat',   home.lat || '');
  setVal('smapHomeLng',   home.lng || '');
  _smapUpdateHomeButton(home);
  _smapUpdatePlusCodeRef(home);
}

function _smapPlaceHomeMarker(home) {
  var map = SchoolMap._map;
  if (!map) return;
  if (SchoolMap._homeMarker) {
    map.removeLayer(SchoolMap._homeMarker);
    SchoolMap._homeMarker = null;
  }
  if (!home || isNaN(home.lat) || isNaN(home.lng)) return;
  
  SchoolMap._homeMarker = L.marker([home.lat, home.lng], {
    icon: L.divIcon({
      html: '<div style="width:28px;height:28px;background:#1e293b;border:2px solid #6366f1;border-radius:8px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 12px rgba(0,0,0,0.4);"><i class="fas fa-home" style="color:#a5b4fc;font-size:14px;"></i></div>',
      className: '',
      iconSize: [28, 28],
      iconAnchor: [14, 14]
    }),
    zIndexOffset: 1000 // Always on top
  }).addTo(map).bindTooltip('<b>🏠 Home Base</b><br>' + [home.block, home.district].filter(Boolean).join(', '), { direction: 'top' });
}

// ---- Routing System ----
function smapShowRoute(schoolName) {
  var map = SchoolMap._map;
  if (!map) return;
  
  var targetLoc = SchoolMap.get(schoolName);
  if (!targetLoc || !targetLoc.lat) {
    if (typeof showToast === 'function') showToast('School location not pinned yet.', 'error');
    return;
  }
  
  var home = smapGetHome();
  if (!home || isNaN(home.lat)) {
    if (typeof showToast === 'function') showToast('Please set your Home Base first.', 'error');
    smapToggleHomePanel();
    return;
  }
  
  smapClearRoute();
  
  if (typeof showToast === 'function') showToast('Fetching route…', 'info');
  
  // Close popups
  map.closePopup();
  
  // OSRM API (longitude,latitude format)
  var url = 'https://router.project-osrm.org/route/v1/driving/' + 
            home.lng + ',' + home.lat + ';' + 
            targetLoc.lng + ',' + targetLoc.lat + 
            '?overview=full&geometries=geojson';
            
  fetch(url)
    .then(function(res) { return res.json(); })
    .then(function(data) {
      if (data.code !== 'Ok' || !data.routes || !data.routes.length) {
        throw new Error('No route found');
      }
      
      var route = data.routes[0];
      var coords = route.geometry.coordinates.map(function(c) { return [c[1], c[0]]; }); // flip to lat,lng
      
      // Draw polyline
      SchoolMap._routeLine = L.polyline(coords, {
        color: '#6366f1',
        weight: 5,
        opacity: 0.8,
        dashArray: '10, 10',
        lineJoin: 'round'
      }).addTo(map);
      
      // Animate dash array (optional css class could be added here)
      
      // Fit bounds to show the whole route
      map.fitBounds(SchoolMap._routeLine.getBounds(), { padding: [50, 50] });
      
      // Show Panel
      var distKm = (route.distance / 1000).toFixed(1);
      var mins = Math.round(route.duration / 60);
      var hrs = Math.floor(mins / 60);
      var remMins = mins % 60;
      var timeStr = hrs > 0 ? hrs + 'h ' + remMins + 'm' : mins + ' min';
      
      var homeLabel = [home.block, home.district].filter(Boolean).join(', ') || 'Home Base';
      
      document.getElementById('smapRouteTo').textContent = 'Route to ' + schoolName;
      document.getElementById('smapRouteMeta').textContent = 'From ' + homeLabel;
      document.getElementById('smapRouteDist').textContent = distKm + ' km';
      document.getElementById('smapRouteEta').innerHTML = '<i class="far fa-clock"></i> ' + timeStr;
      
      document.getElementById('smapRoutePanel').style.display = 'block';
      if (typeof showToast === 'function') showToast('✅ Route found: ' + distKm + ' km (' + timeStr + ')', 'success');
      
    })
    .catch(function(err) {
      console.error('Routing error:', err);
      if (typeof showToast === 'function') showToast('Could not calculate route. Try again later.', 'error');
    });
}

function smapClearRoute() {
  if (SchoolMap._map && SchoolMap._routeLine) {
    SchoolMap._map.removeLayer(SchoolMap._routeLine);
    SchoolMap._routeLine = null;
  }
  if (SchoolMap._map && SchoolMap._allRouteLines) {
    SchoolMap._allRouteLines.forEach(function(l) { SchoolMap._map.removeLayer(l); });
    SchoolMap._allRouteLines = [];
  }
  var panel = document.getElementById('smapRoutePanel');
  if (panel) panel.style.display = 'none';
}

function smapShowAllRoutes() {
  var map = SchoolMap._map;
  if (!map) return;
  var home = smapGetHome();
  if (!home || isNaN(home.lat)) {
    if (typeof showToast === 'function') showToast('Please set your Home Base first.', 'error');
    smapToggleHomePanel();
    return;
  }
  smapClearRoute();
  SchoolMap._allRouteLines = [];
  
  var locs = SchoolMap.getAll();
  var schoolNames = Object.keys(locs).filter(function(name) { return locs[name].lat && locs[name].lng; });
  if (!schoolNames.length) {
    if (typeof showToast === 'function') showToast('No schools pinned yet.', 'info');
    return;
  }
  
  if (typeof showToast === 'function') showToast('Fetching network routes for ' + schoolNames.length + ' schools...', 'info');
  
  var bounds = L.latLngBounds([[home.lat, home.lng]]);
  var loadedCount = 0;
  
  // Stagger requests to avoid OSRM rate limits
  schoolNames.forEach(function(name, idx) {
    var loc = locs[name];
    setTimeout(function() {
      // 1. Draw a temporary straight dashed line immediately
      var tempLine = L.polyline([[home.lat, home.lng], [loc.lat, loc.lng]], {
        color: '#818cf8', weight: 2, opacity: 0.4, dashArray: '4, 8'
      }).addTo(map);
      SchoolMap._allRouteLines.push(tempLine);
      bounds.extend([loc.lat, loc.lng]);
      
      // 2. Fetch the true driving route
      var url = 'https://router.project-osrm.org/route/v1/driving/' + 
                home.lng + ',' + home.lat + ';' + 
                loc.lng + ',' + loc.lat + '?overview=simplified&geometries=geojson';
                
      fetch(url).then(function(res) { return res.json(); }).then(function(data) {
        if (data.code === 'Ok' && data.routes && data.routes.length) {
          // Remove temp line
          map.removeLayer(tempLine);
          var tIdx = SchoolMap._allRouteLines.indexOf(tempLine);
          if (tIdx > -1) SchoolMap._allRouteLines.splice(tIdx, 1);
          
          // Add true route
          var coords = data.routes[0].geometry.coordinates.map(function(c) { return [c[1], c[0]]; });
          var realLine = L.polyline(coords, {
            color: '#6366f1', weight: 3, opacity: 0.55
          }).addTo(map);
          SchoolMap._allRouteLines.push(realLine);
        }
      }).catch(function(e) {
        // Keep the straight line on error
      }).finally(function() {
        loadedCount++;
        if (loadedCount === schoolNames.length) {
          if (typeof showToast === 'function') showToast('Network loaded successfully.', 'success');
          map.fitBounds(bounds, { padding: [30, 30] });
        }
      });
    }, idx * 120); // 120ms stagger between API calls
  });
}


function smapFlyTo(schoolName) {
  var loc = SchoolMap.get(schoolName);
  if (!loc || !loc.lat) return;
  if (!SchoolMap._map) { if (typeof showToast === 'function') showToast('Open the School Map section first', 'info'); return; }
  SchoolMap._map.flyTo([loc.lat, loc.lng], 15, { duration: 1.3 });
  var marker = SchoolMap._markers.find(function(m) {
    var ll = m.getLatLng();
    return Math.abs(ll.lat - loc.lat) < 0.0002 && Math.abs(ll.lng - loc.lng) < 0.0002;
  });
  if (marker) setTimeout(function() { marker.openPopup(); }, 1400);
}

// ---- Render location management table ----
function renderSchoolMapTable(filter) {
  var tbody = document.getElementById('smapTableBody');
  var countEl = document.getElementById('smapTableCount');
  if (!tbody) return;

  var locs = SchoolMap.getAll();
  // Use getKnownSchools() to pull from School Profiles (visits + observations)
  var knownList = SchoolMap.getKnownSchools();
  var allSchools = knownList.map(function(s) { return s.name; });

  var searchEl = document.getElementById('smapTableSearch');
  var q = ((filter !== undefined ? filter : (searchEl ? searchEl.value : '')) || '').toLowerCase();
  var filtered = q ? allSchools.filter(function(s) { return s.toLowerCase().includes(q); }) : allSchools;

  if (countEl) countEl.textContent = filtered.length + ' school' + (filtered.length !== 1 ? 's' : '');

  if (!filtered.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:48px;color:var(--text-secondary);">' +
      '<i class="fas fa-school" style="font-size:28px;display:block;margin-bottom:10px;opacity:.4;"></i>' +
      'No schools yet. Log some visits or click <b>Add School Pin</b>.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(function(name) {
    var loc = locs[name];
    var stats = SchoolMap.visitStats(name);
    var hasPinned = loc && loc.lat && loc.lng;
    var color = hasPinned ? SchoolMap.pinColor(name) : null;
    var statusClass = 'smap-status-unpinned';
    var statusText = '— Not pinned';
    if (hasPinned) {
      if (color === 'green') { statusClass = 'smap-status-green'; statusText = '🟢 This month'; }
      else if (color === 'yellow') { statusClass = 'smap-status-yellow'; statusText = '🟡 Visited'; }
      else { statusClass = 'smap-status-red'; statusText = '🔴 Never'; }
    }
    var safeName = name.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    return '<tr>' +
      '<td><b>' + name + '</b></td>' +
      '<td>' + ((loc && loc.cluster) || stats.cluster || '—') + '</td>' +
      '<td>' + (hasPinned ? Number(loc.lat).toFixed(5) : '—') + '</td>' +
      '<td>' + (hasPinned ? Number(loc.lng).toFixed(5) : '—') + '</td>' +
      '<td style="font-size:11px;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + ((loc && loc.plusCode) || '—') + '</td>' +
      '<td>' + stats.total + '</td>' +
      '<td>' + (stats.lastVisitDate || '—') + '</td>' +
      '<td><span class="smap-status-badge ' + statusClass + '">' + statusText + '</span></td>' +
      '<td>' +
        '<button class="btn btn-outline btn-sm" style="padding:4px 9px;font-size:11px;" onclick="openSchoolLocationEditor(\'' + safeName + '\')">' +
          '<i class="fas fa-' + (hasPinned ? 'edit' : 'map-pin') + '"></i> ' + (hasPinned ? 'Edit' : 'Pin') +
        '</button>' +
        (hasPinned ? '<button class="btn btn-ghost btn-sm" style="padding:4px 7px;font-size:11px;margin-left:4px;" onclick="smapFlyTo(\'' + safeName + '\')" title="Fly to on map"><i class="fas fa-search-location"></i></button>' : '') +
        (hasPinned ? '<a class="btn btn-ghost btn-sm" style="padding:4px 7px;font-size:11px;margin-left:4px;color:#10b981;" href="https://www.google.com/maps/search/?api=1&query=' + loc.lat + ',' + loc.lng + '" target="_blank" title="Open in Google Maps"><i class="fas fa-map-marked-alt"></i></a>' : '') +
      '</td>' +
    '</tr>';
  }).join('');
}

// ---- Export to CSV/Excel ----
function smapExportTableCSV() {
  var locs = SchoolMap.getAll();
  var knownList = SchoolMap.getKnownSchools();
  var allSchools = knownList.map(function(s) { return s.name; });
  
  if (!allSchools.length) {
    if (typeof showToast === 'function') showToast('No data to export', 'info');
    return;
  }
  
  var csvContent = "data:text/csv;charset=utf-8,";
  csvContent += "School Name,Cluster,Latitude,Longitude,Plus Code,Total Visits,Last Visit,Status\n";
  
  allSchools.forEach(function(name) {
    var loc = locs[name] || {};
    var stats = SchoolMap.visitStats(name);
    var hasPinned = loc.lat && loc.lng;
    
    var statusText = 'Not Pinned';
    if (hasPinned) {
      var color = SchoolMap.pinColor(name);
      if (color === 'green') statusText = 'Visited This Month';
      else if (color === 'yellow') statusText = 'Visited Before';
      else statusText = 'Never Visited';
    }
    
    var row = [
      '"' + name.replace(/"/g, '""') + '"',
      '"' + ((loc.cluster || stats.cluster || '')).replace(/"/g, '""') + '"',
      hasPinned ? loc.lat : '',
      hasPinned ? loc.lng : '',
      loc.plusCode || '',
      stats.total,
      stats.lastVisitDate || '',
      statusText
    ];
    csvContent += row.join(",") + "\n";
  });
  
  var encodedUri = encodeURI(csvContent);
  var link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", "school_locations_export.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  if (typeof showToast === 'function') showToast('Excel / CSV export downloaded!', 'success');
}

// =====================================================================
// SCHOOL LOCATION EDITOR MODAL
// =====================================================================
function openSchoolLocationEditor(schoolName) {
  SchoolMap._editingSchool = schoolName || null;

  // Pull school names from School Profiles (getSchoolData aggregates visits + observations)
  var knownList = SchoolMap.getKnownSchools();
  var knownSchools = knownList.map(function(s) { return s.name; });

  var dl = document.getElementById('slmSchoolList');
  if (dl) dl.innerHTML = knownSchools.map(function(s) { return '<option value="' + s + '">'; }).join('');

  var nameEl    = document.getElementById('slmSchoolName');
  var clusterEl = document.getElementById('slmCluster');
  var latEl     = document.getElementById('slmLat');
  var lngEl     = document.getElementById('slmLng');
  var plusEl    = document.getElementById('slmPlusCode');
  var delBtn    = document.getElementById('slmDeleteBtn');
  var title     = document.getElementById('schoolLocationModalTitle');
  var resultEl  = document.getElementById('slmPlusCodeResult');
  var previewEl = document.getElementById('slmPreviewMap');

  // Reset
  if (latEl) latEl.value = '';
  if (lngEl) lngEl.value = '';
  if (plusEl) plusEl.value = '';
  if (clusterEl) clusterEl.value = '';
  if (resultEl) resultEl.style.display = 'none';
  if (previewEl) previewEl.style.display = 'none';
  if (delBtn) delBtn.style.display = 'none';

  // Destroy previous preview map and wipe its DOM so no orphaned tiles show
  if (SchoolMap._previewMap) {
    try { SchoolMap._previewMap.remove(); } catch(e) {}
    SchoolMap._previewMap = null;
    SchoolMap._previewMarker = null;
  }
  // Always clear the preview div contents and hide it
  if (previewEl) {
    previewEl.style.display = 'none';
    previewEl.innerHTML = '';
  }

  if (schoolName) {
    if (title) title.innerHTML = '<i class="fas fa-map-pin"></i> Edit School Location';
    if (nameEl) { nameEl.value = schoolName; nameEl.readOnly = true; }
    var loc = SchoolMap.get(schoolName);
    if (loc) {
      if (latEl) latEl.value = loc.lat || '';
      if (lngEl) lngEl.value = loc.lng || '';
      if (plusEl) plusEl.value = loc.plusCode || '';
      if (clusterEl) clusterEl.value = loc.cluster || '';
      if (delBtn) delBtn.style.display = '';
      if (loc.lat && loc.lng) setTimeout(slmPreviewPin, 250);
    } else {
      // Try to auto-fill cluster from School Profiles data
      var profileEntry = knownList.find(function(s) { return s.name.toLowerCase() === schoolName.toLowerCase(); });
      if (clusterEl && profileEntry && profileEntry.cluster) clusterEl.value = profileEntry.cluster;
    }
  } else {
    if (title) title.innerHTML = '<i class="fas fa-map-pin"></i> Add School Pin';
    if (nameEl) { nameEl.value = ''; nameEl.readOnly = false; }
  }

  slmSwitchTab('latlng');

  if (typeof openModal === 'function') {
    openModal('schoolLocationModal');
  } else {
    var modal = document.getElementById('schoolLocationModal');
    if (modal) modal.classList.add('active');
  }
}

function slmSchoolChanged(name) {
  var nameEl = document.getElementById('slmSchoolName');
  if (nameEl) nameEl.readOnly = false;

  // Auto-fill cluster from School Profiles
  var clEl = document.getElementById('slmCluster');
  if (clEl && !clEl.value) {
    var knownList = SchoolMap.getKnownSchools();
    var entry = knownList.find(function(s) { return s.name.toLowerCase() === name.toLowerCase(); });
    if (entry && entry.cluster) clEl.value = entry.cluster;
  }

  // If this school already has a pin, pre-fill lat/lng
  var loc = SchoolMap.get(name);
  if (loc && loc.lat) {
    var latEl = document.getElementById('slmLat');
    var lngEl = document.getElementById('slmLng');
    if (latEl) latEl.value = loc.lat;
    if (lngEl) lngEl.value = loc.lng;
    slmPreviewPin();
  }
}

function slmSwitchTab(tab) {
  var tabLl = document.getElementById('slmTabLatLng');
  var tabPc = document.getElementById('slmTabPlusCode');
  var panLl = document.getElementById('slmPanelLatLng');
  var panPc = document.getElementById('slmPanelPlusCode');
  if (tab === 'latlng') {
    if (tabLl) tabLl.classList.add('active');
    if (tabPc) tabPc.classList.remove('active');
    if (panLl) panLl.style.display = '';
    if (panPc) panPc.style.display = 'none';
  } else {
    if (tabPc) tabPc.classList.add('active');
    if (tabLl) tabLl.classList.remove('active');
    if (panPc) panPc.style.display = '';
    if (panLl) panLl.style.display = 'none';
  }
}

function slmUseGPS() {
  if (!navigator.geolocation) { if (typeof showToast === 'function') showToast('GPS not supported in this browser', 'error'); return; }
  if (typeof showToast === 'function') showToast('Detecting GPS location…', 'info');
  navigator.geolocation.getCurrentPosition(function(pos) {
    var lat = pos.coords.latitude.toFixed(6);
    var lng = pos.coords.longitude.toFixed(6);
    var latEl = document.getElementById('slmLat');
    var lngEl = document.getElementById('slmLng');
    if (latEl) latEl.value = lat;
    if (lngEl) lngEl.value = lng;
    if (typeof showToast === 'function') showToast('\uD83D\uDCCD GPS: ' + lat + ', ' + lng, 'success');
    slmPreviewPin();
  }, function(err) {
    if (typeof showToast === 'function') showToast('GPS error: ' + err.message + '. Check browser permissions.', 'error');
  });
}

function slmDecodePlusCode() {
  var raw = (document.getElementById('slmPlusCode') ? document.getElementById('slmPlusCode').value : '').trim();
  if (!raw) { if (typeof showToast === 'function') showToast('Enter a Plus Code first', 'error'); return; }

  if (typeof OpenLocationCode === 'undefined') {
    if (typeof showToast === 'function') showToast('Plus Code library not loaded. Check internet and refresh.', 'error');
    return;
  }

  // The library uses INSTANCE methods — must create an instance first
  var olc = new OpenLocationCode();

  try {
    var code = raw.trim();
    var refLat = 20.5937, refLng = 78.9629; // India center default
    var cityRef = '';

    // Strip Google Maps URL prefix: https://plus.codes/XXXX+YY
    var urlMatch = code.match(/plus\.codes\/([A-Z0-9+]+)/i);
    if (urlMatch) code = urlMatch[1];

    // Separate the plus-code part from any city suffix
    // Plus codes always contain exactly one '+'
    var plusIdx = code.indexOf('+');
    if (plusIdx === -1) {
      if (typeof showToast === 'function') showToast('Invalid Plus Code — must contain a "+" character. Example: 7JCGXM2P+HX', 'error');
      return;
    }
    // Code is everything up to and including a few chars after '+'
    // City is anything separated by whitespace after the code
    var parts = code.split(/\s+/);
    code = parts[0].toUpperCase();
    cityRef = parts.slice(1).join(' ').trim();

    // Validate
    if (!olc.isValid(code)) {
      if (typeof showToast === 'function') showToast('Invalid Plus Code "' + code + '". Example format: 7JCGXM2P+HX or XM2P+HX Bankura', 'error');
      return;
    }

    var decoded;

    if (olc.isShort(code)) {
      // Short code needs a reference — use home base first, then city from suffix
      var home = smapGetHome();
      refLat = home.lat;
      refLng = home.lng;

      // If a city suffix was provided, try to match it (overrides home base)
      if (cityRef) {
        var cityCoords = {
          'bankura': [23.2324, 87.0749], 'kolkata': [22.5726, 88.3639],
          'delhi': [28.6139, 77.2090], 'newdelhi': [28.6139, 77.2090],
          'mumbai': [19.0760, 72.8777], 'bangalore': [12.9716, 77.5946],
          'bengaluru': [12.9716, 77.5946], 'chennai': [13.0827, 80.2707],
          'hyderabad': [17.3850, 78.4867], 'pune': [18.5204, 73.8567],
          'jaipur': [26.9124, 75.7873], 'lucknow': [26.8467, 80.9462],
          'bhopal': [23.2599, 77.4126], 'patna': [25.6093, 85.1376],
          'ranchi': [23.3441, 85.3096], 'bhubaneswar': [20.2961, 85.8245],
          'ahmedabad': [23.0225, 72.5714], 'indore': [22.7196, 75.8577],
          'nagpur': [21.1458, 79.0882], 'raipur': [21.2514, 81.6296],
          'durg': [21.1904, 81.2849], 'bhilai': [21.2090, 81.4285],
          'rajnandgaon': [21.0994, 81.0320], 'balod': [20.7323, 81.2048],
          'kawardha': [22.0121, 81.2305], 'kabirdham': [22.0121, 81.2305],
          'dehradun': [30.3165, 78.0322], 'shimla': [31.1048, 77.1734],
          'guwahati': [26.1445, 91.7362], 'kochi': [9.9312, 76.2673],
          'thiruvananthapuram': [8.5241, 76.9366], 'mysuru': [12.2958, 76.6394],
          'mysore': [12.2958, 76.6394], 'varanasi': [25.3176, 82.9739],
          'agra': [27.1767, 78.0081], 'surat': [21.1702, 72.8311],
          'coimbatore': [11.0168, 76.9558], 'vizag': [17.6868, 83.2185],
          'visakhapatnam': [17.6868, 83.2185]
        };
        var cityKey = cityRef.toLowerCase().replace(/[^a-z]/g, '');
        if (cityCoords[cityKey]) {
          refLat = cityCoords[cityKey][0];
          refLng = cityCoords[cityKey][1];
        }
      }

      var refName = cityRef || [home.block, home.district].filter(Boolean).join(', ') || 'home base';
      var fullCode = olc.recoverNearest(code, refLat, refLng);
      decoded = olc.decode(fullCode);
      if (typeof showToast === 'function') showToast('Short code expanded relative to: ' + refName, 'info', 4000);

    } else if (olc.isFull(code)) {
      decoded = olc.decode(code);

    } else {
      if (typeof showToast === 'function') showToast('Could not decode this Plus Code. Try a full 11-char code like "7JCGXM2P+HX".', 'error');
      return;
    }

    var lat = decoded.latitudeCenter;
    var lng = decoded.longitudeCenter;
    var latEl = document.getElementById('slmLat');
    var lngEl = document.getElementById('slmLng');
    if (latEl) latEl.value = lat.toFixed(6);
    if (lngEl) lngEl.value = lng.toFixed(6);

    var resultEl = document.getElementById('slmPlusCodeResult');
    if (resultEl) {
      resultEl.style.display = 'block';
      var precLatM = decoded.latitudeHeight  !== undefined ? decoded.latitudeHeight  : (decoded.latitudeHi  - decoded.latitudeLo);
      var precLngM = decoded.longitudeWidth !== undefined ? decoded.longitudeWidth : (decoded.longitudeHi - decoded.longitudeLo);
      var precLat = isFinite(precLatM) ? Math.round(precLatM * 111000) : '—';
      var precLng = isFinite(precLngM) ? Math.round(precLngM * 111000) : '—';
      resultEl.innerHTML = '✅ <b>Decoded:</b> ' + lat.toFixed(5) + ', ' + lng.toFixed(5) +
        '<br><small style="color:var(--text-secondary)">Precision: ~' + precLat + 'm \xD7 ' + precLng + 'm</small>';
    }

    if (typeof showToast === 'function') showToast('\u2705 Decoded: ' + lat.toFixed(5) + ', ' + lng.toFixed(5), 'success');
    slmPreviewPin();

  } catch(err) {
    if (typeof showToast === 'function') showToast('Decode error: ' + err.message, 'error');
  }
}


function slmPreviewPin() {
  var latEl = document.getElementById('slmLat');
  var lngEl = document.getElementById('slmLng');
  if (!latEl || !lngEl) return;
  var lat = parseFloat(latEl.value);
  var lng = parseFloat(lngEl.value);
  // Only show preview when coordinates are actually valid
  if (isNaN(lat) || isNaN(lng) || lat === 0 || lng === 0) return;
  if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return;

  var previewEl = document.getElementById('slmPreviewMap');
  if (!previewEl) return;
  previewEl.style.display = 'block';

  if (!SchoolMap._previewMap) {
    SchoolMap._previewMap = L.map('slmPreviewMap', { zoomControl: false, attributionControl: false }).setView([lat, lng], 14);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19 }).addTo(SchoolMap._previewMap);
    SchoolMap._previewMarker = L.marker([lat, lng]).addTo(SchoolMap._previewMap);
  } else {
    SchoolMap._previewMap.setView([lat, lng], 14);
    SchoolMap._previewMarker.setLatLng([lat, lng]);
    setTimeout(function() { SchoolMap._previewMap.invalidateSize(); }, 50);
  }
}

function saveSchoolLocation() {
  var nameEl    = document.getElementById('slmSchoolName');
  var latEl     = document.getElementById('slmLat');
  var lngEl     = document.getElementById('slmLng');
  var clusterEl = document.getElementById('slmCluster');
  var plusEl    = document.getElementById('slmPlusCode');

  var name    = (nameEl ? nameEl.value : '').trim();
  var lat     = parseFloat(latEl ? latEl.value : '');
  var lng     = parseFloat(lngEl ? lngEl.value : '');
  var cluster = (clusterEl ? clusterEl.value : '').trim();
  var plusCode= (plusEl ? plusEl.value : '').trim();

  if (!name) { if (typeof showToast === 'function') showToast('School name is required', 'error'); return; }
  if (isNaN(lat) || isNaN(lng)) {
    if (typeof showToast === 'function') showToast('Enter a valid latitude & longitude, or decode a Plus Code', 'error');
    return;
  }
  if (lat < -90 || lat > 90) { if (typeof showToast === 'function') showToast('Latitude must be between -90 and 90', 'error'); return; }
  if (lng < -180 || lng > 180) { if (typeof showToast === 'function') showToast('Longitude must be between -180 and 180', 'error'); return; }

  SchoolMap.upsert(name, { lat: lat, lng: lng, cluster: cluster, plusCode: plusCode });
  if (typeof showToast === 'function') showToast('\uD83D\uDCCD "' + name + '" pinned at ' + lat.toFixed(4) + ', ' + lng.toFixed(4), 'success');

  if (typeof closeModal === 'function') closeModal('schoolLocationModal');

  if (SchoolMap._map) renderSchoolMap();
  renderSchoolMapTable();
}

function deleteSchoolLocation() {
  var name = SchoolMap._editingSchool || ((document.getElementById('slmSchoolName') ? document.getElementById('slmSchoolName').value : '').trim());
  if (!name) return;
  SchoolMap.remove(name);
  if (typeof showToast === 'function') showToast('Pin removed for "' + name + '"', 'info');
  if (typeof closeModal === 'function') closeModal('schoolLocationModal');
  if (SchoolMap._map) renderSchoolMap();
  renderSchoolMapTable();
}

// =====================================================================
// AI VISIT PLANNER
// =====================================================================

var _smapAILastPlan = null; // stores last generated plan for copy/highlight

function smapToggleAIPanel() {
  var panel = document.getElementById('smapAIPanel');
  var btn   = document.getElementById('smapAIBtn');
  if (!panel) return;
  var isOpen = panel.style.display !== 'none';
  panel.style.display = isOpen ? 'none' : 'block';
  if (btn) {
    btn.style.background = isOpen ? '' : 'rgba(52,211,153,0.1)';
    btn.style.borderColor = isOpen ? '' : 'rgba(52,211,153,0.5)';
  }
}

function smapRunAIPlanner() {
  if (typeof SarvamAI === 'undefined' || !SarvamAI.isConfigured()) {
    if (typeof showToast === 'function') showToast('Configure Sarvam AI in Settings → Sarvam AI first', 'error');
    return;
  }

  var days   = parseInt((document.getElementById('smapAIDays')   || {}).value  || '5');
  var perDay = parseInt((document.getElementById('smapAIPerDay') || {}).value  || '3');

  // Gather all pinned schools with full context
  var locs = SchoolMap.getAll();
  var pinnedSchools = Object.keys(locs).filter(function(n) {
    return locs[n].lat && locs[n].lng;
  });

  if (pinnedSchools.length < 2) {
    if (typeof showToast === 'function') showToast('Pin at least 2 schools on the map first', 'warning');
    return;
  }

  var home = (typeof smapGetHome === 'function') ? smapGetHome() : { lat: 21.1793, lng: 81.2833, block: 'Magarlod', district: 'Durg' };

  // Build rich school data list
  var schoolList = pinnedSchools.map(function(name) {
    var loc   = locs[name];
    var stats = SchoolMap.visitStats(name);
    var color = SchoolMap.pinColor(name);

    // Straight-line distance from home base (km)
    var dLat = (loc.lat - home.lat) * Math.PI / 180;
    var dLng = (loc.lng - home.lng) * Math.PI / 180;
    var a = Math.sin(dLat/2)*Math.sin(dLat/2) +
            Math.cos(home.lat * Math.PI/180)*Math.cos(loc.lat * Math.PI/180)*Math.sin(dLng/2)*Math.sin(dLng/2);
    var distKm = Math.round(6371 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));

    var priority = color === 'red' ? 'HIGH (never visited)' :
                   color === 'yellow' ? 'MEDIUM (visited before)' : 'LOW (visited this month)';

    return {
      name: name,
      cluster: loc.cluster || stats.cluster || '—',
      lat: Number(loc.lat).toFixed(4),
      lng: Number(loc.lng).toFixed(4),
      totalVisits: stats.total,
      lastVisit: stats.lastVisitDate || 'Never',
      thisMonth: stats.thisMonth,
      distanceFromHomeKm: distKm,
      priority: priority
    };
  });

  // Sort by priority then distance
  var priorityOrder = { 'HIGH (never visited)': 0, 'MEDIUM (visited before)': 1, 'LOW (visited this month)': 2 };
  schoolList.sort(function(a, b) {
    var pa = priorityOrder[a.priority] || 2;
    var pb = priorityOrder[b.priority] || 2;
    if (pa !== pb) return pa - pb;
    return a.distanceFromHomeKm - b.distanceFromHomeKm;
  });

  var homeStr = [home.block, home.district].filter(Boolean).join(', ') || (home.lat + ', ' + home.lng);

  var schoolContext = schoolList.map(function(s, i) {
    return (i+1) + '. ' + s.name +
      ' | Cluster: ' + s.cluster +
      ' | Coords: (' + s.lat + ', ' + s.lng + ')' +
      ' | Distance from home: ~' + s.distanceFromHomeKm + 'km' +
      ' | Total visits: ' + s.totalVisits +
      ' | Last visit: ' + s.lastVisit +
      ' | Priority: ' + s.priority;
  }).join('\n');

  var systemPrompt = 'You are an expert field visit planner for Azim Premji Foundation Resource Persons in India. ' +
    'You optimize school visit schedules based on geographic proximity, visit frequency, and educational priority. ' +
    'You create practical, realistic daily visit plans that minimize travel time while maximizing coverage of neglected schools. ' +
    'CRITICAL: Output ONLY the visit plan using the exact format given. Do NOT include any reasoning, thinking, explanation, or preamble. ' +
    'Start your response DIRECTLY with ## 📅 Day 1. Never repeat any section twice.';


  var userPrompt = 'Create an optimized ' + days + '-day school visit plan for an APF Resource Person.\n\n' +
    'HOME BASE: ' + homeStr + ' (Lat: ' + home.lat + ', Lng: ' + home.lng + ')\n' +
    'PLAN: ' + perDay + ' schools per day for ' + days + ' day(s)\n' +
    'TOTAL schools to plan: ' + Math.min(days * perDay, schoolList.length) + ' (from ' + schoolList.length + ' pinned)\n\n' +
    'SCHOOLS (sorted by priority then distance from home):\n' + schoolContext + '\n\n' +
    'RULES:\n' +
    '1. Prioritize HIGH priority (never visited) schools first\n' +
    '2. Group geographically close schools on the same day to minimize travel\n' +
    '3. Start each day from home base and suggest a logical visit sequence\n' +
    '4. Include estimated travel distance for each day\n' +
    '5. If a cluster has multiple schools, try to group them together\n' +
    '6. Suggest the best time of day for each visit (morning/afternoon)\n\n' +
    'OUTPUT FORMAT (use exactly this structure):\n' +
    '## 📅 Day 1 — [Date suggestion or "Monday"]\n' +
    '**Route:** Home → School A → School B → School C → Home\n' +
    '**Est. distance:** ~XX km\n' +
    '- 🏫 **[School Name]** (Cluster: X) — [Priority] — [Why visit today: e.g., never visited, closest to School B]\n' +
    '  - ⏰ Suggested time: 9:30 AM\n' +
    '  - 💡 Focus: [specific suggestion e.g., check attendance records, observe Math class]\n' +
    '[repeat for each school]\n\n' +
    '## 📅 Day 2 — ...\n' +
    '[continue for all days]\n\n' +
    '## 🗺️ Coverage Summary\n' +
    '- Total schools planned: X\n' +
    '- Never-visited schools covered: X\n' +
    '- Estimated total travel: ~XX km\n' +
    '- Schools not yet planned (visit next cycle): [list remaining]';

  // Show loading
  var loadingEl = document.getElementById('smapAILoading');
  var outputEl  = document.getElementById('smapAIOutput');
  var runBtn    = document.getElementById('smapAIRunBtn');
  if (loadingEl) loadingEl.style.display = 'block';
  if (outputEl)  outputEl.style.display  = 'none';
  if (runBtn)    runBtn.disabled = true;

  SarvamAI.chat([
    { role: 'system', content: systemPrompt },
    { role: 'user',   content: userPrompt }
  ], { temperature: 0.4, max_tokens: 3000 }).then(function(res) {
    var reply = (res.choices && res.choices[0] && res.choices[0].message && res.choices[0].message.content) || '';

    // Strip <think>...</think> blocks (some models use these)
    reply = reply.replace(/<think>[\s\S]*?<\/think>/gi, '').trim();

    // Strip ALL text before the first ## heading — removes model reasoning/preamble
    var firstHeading = reply.indexOf('## ');
    if (firstHeading > 0) reply = reply.substring(firstHeading);

    // Deduplicate: if the plan section repeats itself, keep only the first occurrence
    var day1Match = reply.match(/(## 📅 Day 1)/g);
    if (day1Match && day1Match.length > 1) {
      // Find second occurrence of Day 1 and cut everything from there
      var idx = reply.indexOf('## 📅 Day 1');
      var idx2 = reply.indexOf('## 📅 Day 1', idx + 10);
      if (idx2 > idx) reply = reply.substring(0, idx2).trim();
    }

    // Also deduplicate ## Coverage Summary repeated sections
    var summaryMatch = reply.match(/(## 🗺️)/g);
    if (summaryMatch && summaryMatch.length > 1) {
      var s1 = reply.indexOf('## 🗺️');
      var s2 = reply.indexOf('## 🗺️', s1 + 5);
      if (s2 > s1) reply = reply.substring(0, s2).trim();
    }

    if (!reply) reply = 'No plan generated. Please try again.';

    _smapAILastPlan = { text: reply, schools: schoolList };
    _smapRenderAIPlan(reply);

    if (loadingEl) loadingEl.style.display = 'none';
    if (outputEl)  outputEl.style.display  = 'block';
    if (runBtn)    { runBtn.disabled = false; }
    if (typeof showToast === 'function') showToast('AI Visit Plan generated! 🗺️', 'success');

  }).catch(function(err) {
    if (loadingEl) loadingEl.style.display = 'none';
    if (runBtn)    runBtn.disabled = false;
    if (typeof showToast === 'function') showToast('AI error: ' + err.message, 'error');
  });
}

function _smapRenderAIPlan(mdText) {
  var el = document.getElementById('smapAIOutputContent');
  if (!el) return;

  // Convert basic markdown to HTML
  var html = mdText
    .replace(/^## (.+)$/gm, '<h3 style="color:#34d399;margin:18px 0 8px;font-size:14px;border-bottom:1px solid rgba(52,211,153,0.2);padding-bottom:6px;">$1</h3>')
    .replace(/^\*\*(.+?)\*\*/gm, '<strong style="color:#e2e8f0;">$1</strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong style="color:#e2e8f0;">$1</strong>')
    .replace(/^- 🏫 (.+)$/gm, '<div style="margin:10px 0 4px;padding:10px 12px;background:rgba(99,102,241,0.08);border-left:3px solid #6366f1;border-radius:0 8px 8px 0;font-size:13px;color:#cbd5e1;">🏫 $1</div>')
    .replace(/^  - ⏰ (.+)$/gm, '<div style="margin-left:20px;font-size:12px;color:#94a3b8;">⏰ $1</div>')
    .replace(/^  - 💡 (.+)$/gm, '<div style="margin-left:20px;font-size:12px;color:#a5b4fc;">💡 $1</div>')
    .replace(/^- (.+)$/gm, '<div style="font-size:12px;color:#94a3b8;margin:3px 0;padding-left:8px;">• $1</div>')
    .replace(/\n{2,}/g, '<br>')
    .replace(/\n/g, '<br>');

  el.innerHTML = '<div style="font-size:13px;line-height:1.7;color:#cbd5e1;">' + html + '</div>';
}

function smapAICopyPlan() {
  if (!_smapAILastPlan) {
    if (typeof showToast === 'function') showToast('Generate a plan first', 'warning');
    return;
  }
  navigator.clipboard.writeText(_smapAILastPlan.text).then(function() {
    if (typeof showToast === 'function') showToast('Plan copied to clipboard! 📋', 'success');
  }).catch(function() {
    if (typeof showToast === 'function') showToast('Could not copy — try manually selecting the text', 'error');
  });
}

function smapAIHighlightOnMap() {
  if (!_smapAILastPlan || !SchoolMap._map) {
    if (typeof showToast === 'function') showToast('Generate a plan first', 'warning');
    return;
  }

  // Pulse/highlight all planned school markers on the map
  var plannedNames = (_smapAILastPlan.schools || []).map(function(s) { return s.name.toLowerCase(); });
  var highlightCount = 0;

  SchoolMap._markers.forEach(function(marker) {
    var popup = marker.getPopup();
    if (!popup) return;
    var content = popup.getContent ? popup.getContent() : '';
    var matched = plannedNames.some(function(n) {
      return typeof content === 'string' && content.toLowerCase().includes(n);
    });
    if (matched) {
      marker.openPopup();
      highlightCount++;
    }
  });

  // Fly to fit all planned schools
  var locs = SchoolMap.getAll();
  var latLngs = plannedNames.map(function(n) {
    var key = Object.keys(locs).find(function(k) { return k.toLowerCase() === n; });
    if (key && locs[key].lat) return [locs[key].lat, locs[key].lng];
    return null;
  }).filter(Boolean);

  if (latLngs.length > 1) {
    try { SchoolMap._map.fitBounds(L.latLngBounds(latLngs), { padding: [40, 40], maxZoom: 13 }); } catch(e) {}
  }

  if (typeof showToast === 'function') showToast(highlightCount + ' planned schools highlighted on map 📍', 'success');
}

