/* AEDナビ — script.js
 * AED.xlsx を読み込み、出発地から最も近い（徒歩経路が最短の）AEDへの経路を表示する。
 *
 * 最短経路の求め方:
 *   1. 全AEDの直線距離（ハバーサイン）を計算し、近い順に上位 N 件を候補にする
 *   2. 候補ごとに徒歩ルートを OSRM で取得し、実際の経路距離が最短のものを選ぶ
 *   3. 経路が取得できなかった候補は「直線距離×1.3」で概算する（地図上は点線）
 */
(function () {
  'use strict';

  /* ===== 設定 ===== */
  const CONFIG = {
    xlsxUrl: 'AED.xlsx',
    candidateCount: 5,          // 経路を計算する候補数
    walkMetersPerMin: 80,       // 徒歩速度（m/分）
    detourFactor: 1.3,          // 経路未取得時の概算係数
    routeEndpoint: 'https://routing.openstreetmap.de/routed-foot/route/v1/foot/',
    routeTimeoutMs: 9000,
    tileUrl: 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png',
    tileAttribution: '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">国土地理院</a>',
    defaultCenter: [35.68, 139.76],
    defaultZoom: 5
  };

  /* Excel見出しの別名（正規化後の文字列で比較） */
  const HEADER_ALIASES = {
    name:    ['施設名', '名称', '名前', '設置施設', 'name'],
    address: ['住所', '所在地', 'address'],
    lat:     ['緯度', 'lat', 'latitude'],
    lng:     ['経度', 'lng', 'lon', 'long', 'longitude'],
    place:   ['設置場所', '設置位置', '設置箇所', '場所', 'place', 'location'],
    hours:   ['利用可能時間', '利用時間', '使用可能時間', '利用可能日時', 'hours'],
    tel:     ['電話番号', '電話', 'tel', 'phone'],
    note:    ['備考', 'メモ', 'note', 'notes']
  };

  /* ===== 状態 ===== */
  const state = {
    aeds: [],
    markers: new Map(),      // aed.id -> L.marker
    start: null,             // {lat, lng}
    cache: new Map(),        // aed.id -> {distance, approx, coords, straight}
    ranked: [],              // [{aed, info}] 近い順
    selected: null,
    picking: false,
    only24: false,
    token: 0                 // 検索の競合防止
  };

  let map, startMarker, accuracyCircle, routeLayers = [];

  /* ===== DOM ===== */
  const $ = (id) => document.getElementById(id);
  const ui = {
    map: $('map'),
    btnLocate: $('btnLocate'),
    btnPick: $('btnPick'),
    chk24: $('chk24'),
    status: $('status'),
    result: $('result'),
    candidates: $('candidates'),
    candList: $('candList'),
    dataInfo: $('dataInfo')
  };

  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;   // textContent でXSSを防ぐ
    return node;
  }

  function setStatus(message, kind) {
    ui.status.textContent = message || '';
    ui.status.className = 'status' + (kind ? ' ' + kind : '');
  }

  /* ===== ユーティリティ ===== */
  const normalize = (s) => String(s).normalize('NFKC').replace(/\s+/g, '').toLowerCase();

  function parseNumber(v) {
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).normalize('NFKC').replace(/[^\d.\-]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  }

  function haversine(a, b) {
    const R = 6371000;
    const rad = Math.PI / 180;
    const dLat = (b.lat - a.lat) * rad;
    const dLng = (b.lng - a.lng) * rad;
    const h = Math.sin(dLat / 2) ** 2 +
      Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLng / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  }

  function formatDistance(m) {
    if (m < 1000) return { value: String(Math.max(10, Math.round(m / 10) * 10)), unit: 'm' };
    return { value: (m / 1000).toFixed(1), unit: 'km' };
  }

  const walkMinutes = (m) => Math.max(1, Math.round(m / CONFIG.walkMetersPerMin));

  /* ===== Excel読み込み ===== */
  async function loadAEDs() {
    let res;
    try {
      res = await fetch(CONFIG.xlsxUrl, { cache: 'no-cache' });
    } catch (e) {
      throw new Error('fetch');
    }
    if (!res.ok) throw new Error('http ' + res.status);

    const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (!rows.length) throw new Error('empty');

    // 見出し名 -> 内部キー
    const colMap = {};
    Object.keys(rows[0]).forEach((header) => {
      const h = normalize(header);
      for (const key of Object.keys(HEADER_ALIASES)) {
        if (!colMap[key] && HEADER_ALIASES[key].some((alias) => normalize(alias) === h)) {
          colMap[key] = header;
        }
      }
    });
    if (!colMap.lat || !colMap.lng) throw new Error('columns');

    const aeds = [];
    let skipped = 0;

    rows.forEach((row, i) => {
      let lat = parseNumber(row[colMap.lat]);
      let lng = parseNumber(row[colMap.lng]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) { skipped++; return; }
      // 緯度経度が逆に入力されている場合の救済（日本付近など経度が90超の地域）
      if (lat > 90 && lng <= 90) { [lat, lng] = [lng, lat]; }
      if (Math.abs(lat) > 90 || Math.abs(lng) > 180) { skipped++; return; }

      const get = (k) => (colMap[k] ? String(row[colMap[k]]).trim() : '');
      const hours = get('hours');
      aeds.push({
        id: i,
        name: get('name') || 'AED（名称未登録）',
        address: get('address'),
        lat, lng,
        place: get('place'),
        hours,
        tel: get('tel'),
        note: get('note'),
        is24h: /24時間|24h|終日|常時/i.test(hours.normalize('NFKC'))
      });
    });

    if (!aeds.length) throw new Error('empty');
    return { aeds, skipped };
  }

  /* ===== 地図 ===== */
  function initMap() {
    map = L.map(ui.map, { zoomControl: false }).setView(CONFIG.defaultCenter, CONFIG.defaultZoom);
    L.control.zoom({ position: 'bottomright', zoomInTitle: '拡大', zoomOutTitle: '縮小' }).addTo(map);
    L.tileLayer(CONFIG.tileUrl, {
      maxZoom: 18,
      attribution: CONFIG.tileAttribution
    }).addTo(map);

    map.on('click', (e) => {
      if (!state.picking) return;
      setPicking(false);
      setStart({ lat: e.latlng.lat, lng: e.latlng.lng }, 0);
      search();
    });
  }

  function aedIcon(aed) {
    const isTarget = state.selected && state.selected.id === aed.id;
    const dim = state.only24 && !aed.is24h;
    const size = isTarget ? 46 : 34;
    return L.divIcon({
      className: '',
      html: '<div class="aed-pin' + (isTarget ? ' is-target' : '') + (dim ? ' is-off' : '') + '">AED</div>',
      iconSize: [size, size],
      iconAnchor: [size / 2, size / 2]
    });
  }

  function addMarkers() {
    state.aeds.forEach((aed) => {
      const marker = L.marker([aed.lat, aed.lng], {
        icon: aedIcon(aed),
        title: aed.name,
        keyboard: true,
        riseOnHover: true
      }).addTo(map);
      marker.on('click', () => showAED(aed));
      state.markers.set(aed.id, marker);
    });
    map.fitBounds(L.latLngBounds(state.aeds.map((a) => [a.lat, a.lng])).pad(0.2));
  }

  function refreshIcons() {
    state.aeds.forEach((aed) => {
      const m = state.markers.get(aed.id);
      if (!m) return;
      m.setIcon(aedIcon(aed));
      m.setZIndexOffset(state.selected && state.selected.id === aed.id ? 1000 : 0);
    });
  }

  function clearRoute() {
    routeLayers.forEach((l) => map.removeLayer(l));
    routeLayers = [];
  }

  function drawRoute(info) {
    clearRoute();
    if (!state.start) return;

    let latlngs;
    if (info.coords) {
      latlngs = info.coords.map(([lng, lat]) => [lat, lng]);
      routeLayers.push(L.polyline(latlngs, { color: '#fff', weight: 10, opacity: .95, lineCap: 'round', lineJoin: 'round' }).addTo(map));
      routeLayers.push(L.polyline(latlngs, { color: '#1a5fd0', weight: 6, opacity: 1, lineCap: 'round', lineJoin: 'round' }).addTo(map));
    } else {
      latlngs = [[state.start.lat, state.start.lng], [info.aed.lat, info.aed.lng]];
      routeLayers.push(L.polyline(latlngs, { color: '#1a5fd0', weight: 5, dashArray: '2 10', lineCap: 'round' }).addTo(map));
    }
    map.fitBounds(L.latLngBounds(latlngs).pad(0.25), { maxZoom: 18 });
  }

  /* ===== 出発地 ===== */
  function setStart(pos, accuracy) {
    state.start = pos;
    state.cache = new Map();
    state.ranked = [];

    if (startMarker) map.removeLayer(startMarker);
    if (accuracyCircle) { map.removeLayer(accuracyCircle); accuracyCircle = null; }

    startMarker = L.marker([pos.lat, pos.lng], {
      icon: L.divIcon({ className: '', html: '<div class="me-pin"></div>', iconSize: [22, 22], iconAnchor: [11, 11] }),
      title: '出発地',
      interactive: false,
      zIndexOffset: 2000
    }).addTo(map);

    if (accuracy && accuracy > 20) {
      accuracyCircle = L.circle([pos.lat, pos.lng], {
        radius: accuracy, color: '#1a5fd0', weight: 1, fillColor: '#1a5fd0', fillOpacity: .1, interactive: false
      }).addTo(map);
    }
  }

  function locate() {
    if (!navigator.geolocation) {
      setStatus('この端末では現在地を取得できません。「地図をタップして出発地を指定」を使ってください。', 'error');
      return;
    }
    setPicking(false);
    setStatus('現在地を取得中…', 'busy');
    ui.btnLocate.disabled = true;

    navigator.geolocation.getCurrentPosition(
      (p) => {
        ui.btnLocate.disabled = false;
        setStart({ lat: p.coords.latitude, lng: p.coords.longitude }, p.coords.accuracy);
        search();
      },
      (err) => {
        ui.btnLocate.disabled = false;
        const reason = err.code === 1
          ? '位置情報の利用が許可されていません。ブラウザの設定を確認してください。'
          : '現在地を取得できませんでした。';
        setStatus(reason + ' 地図をタップして出発地を指定することもできます。', 'error');
        setPicking(true);
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 10000 }
    );
  }

  function setPicking(on) {
    state.picking = on;
    ui.btnPick.setAttribute('aria-pressed', String(on));
    ui.map.classList.toggle('picking', on);
    if (on) setStatus('地図上の出発地点をタップしてください。');
  }

  /* ===== 経路計算 ===== */
  async function fetchWalkingRoute(from, to) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), CONFIG.routeTimeoutMs);
    try {
      const url = CONFIG.routeEndpoint +
        from.lng + ',' + from.lat + ';' + to.lng + ',' + to.lat +
        '?overview=full&geometries=geojson';
      const res = await fetch(url, { signal: ctrl.signal });
      if (!res.ok) throw new Error('http ' + res.status);
      const data = await res.json();
      if (data.code !== 'Ok' || !data.routes || !data.routes.length) throw new Error('no route');
      return { distance: data.routes[0].distance, coords: data.routes[0].geometry.coordinates };
    } finally {
      clearTimeout(timer);
    }
  }

  async function computeInfo(aed) {
    const straight = haversine(state.start, aed);
    try {
      const r = await fetchWalkingRoute(state.start, aed);
      return { aed, straight, distance: r.distance, coords: r.coords, approx: false };
    } catch (e) {
      return { aed, straight, distance: straight * CONFIG.detourFactor, coords: null, approx: true };
    }
  }

  async function search() {
    if (!state.start) return;
    const token = ++state.token;

    setStatus('最短経路を計算中…', 'busy');
    ui.result.hidden = true;
    ui.candidates.hidden = true;
    clearRoute();

    const pool = state.only24 ? state.aeds.filter((a) => a.is24h) : state.aeds;
    if (!pool.length) {
      setStatus('条件に合うAEDがありません。「24時間使えるAEDのみ」を外してください。', 'error');
      return;
    }

    const nearest = pool
      .map((aed) => ({ aed, straight: haversine(state.start, aed) }))
      .sort((a, b) => a.straight - b.straight)
      .slice(0, CONFIG.candidateCount);

    const infos = await Promise.all(nearest.map((c) => computeInfo(c.aed)));
    if (token !== state.token) return;   // 後から別の検索が始まっていたら破棄

    infos.forEach((info) => state.cache.set(info.aed.id, info));
    state.ranked = infos.sort((a, b) => a.distance - b.distance);

    renderCandidates();
    const best = state.ranked[0];
    const offline = state.ranked.some((r) => r.approx);
    await showAED(best.aed);

    if (token !== state.token) return;
    if (offline) {
      setStatus('一部の経路を取得できなかったため、直線距離から概算しています。', 'error');
    } else {
      setStatus('最寄りのAEDまでの最短経路を表示しました。');
    }
  }

  /* ===== 表示 ===== */
  async function showAED(aed) {
    state.selected = aed;
    refreshIcons();

    if (!state.start) {
      renderResult(aed, null);
      map.setView([aed.lat, aed.lng], 17);
      return;
    }

    let info = state.cache.get(aed.id);
    if (!info) {
      setStatus('経路を計算中…', 'busy');
      info = await computeInfo(aed);
      if (!state.start) return;
      state.cache.set(aed.id, info);
      setStatus('');
    }
    if (state.selected !== aed) return;

    renderResult(aed, info);
    drawRoute(info);
    markActiveCandidate();
  }

  function row(dl, label, value, href) {
    if (!value) return;
    dl.appendChild(el('dt', null, label));
    const dd = el('dd');
    if (href) {
      const a = el('a', null, value);
      a.href = href;
      dd.appendChild(a);
    } else {
      dd.textContent = value;
    }
    dl.appendChild(dd);
  }

  function renderResult(aed, info) {
    const box = ui.result;
    box.textContent = '';

    const rankIndex = state.ranked.findIndex((r) => r.aed.id === aed.id);
    box.appendChild(el('span', 'tag', rankIndex === 0 ? '最寄りのAED' : '選択中のAED'));
    box.appendChild(el('h2', null, aed.name));

    if (info) {
      const d = formatDistance(info.distance);
      const dist = el('p', 'dist');
      dist.appendChild(el('strong', null, d.value + ' ' + d.unit));
      dist.appendChild(el('span', null, '徒歩 約' + walkMinutes(info.distance) + '分'));
      box.appendChild(dist);
      if (info.approx) box.appendChild(el('p', 'approx', '※ 経路を取得できないため、直線距離からの概算です。'));
      if (info.distance > 2000) {
        box.appendChild(el('p', 'approx', '※ 距離があります。車や自転車で取りに行けるか、周囲の人に相談してください。'));
      }
    }

    const dl = el('dl');
    row(dl, '住所', aed.address);
    row(dl, '設置場所', aed.place);
    row(dl, '利用時間', aed.hours);
    row(dl, '電話', aed.tel, aed.tel ? 'tel:' + aed.tel.replace(/[^\d+]/g, '') : '');
    row(dl, '備考', aed.note);
    if (dl.children.length) box.appendChild(dl);

    const dest = aed.lat + ',' + aed.lng;
    const origin = state.start ? '&origin=' + state.start.lat + ',' + state.start.lng : '';
    const link = el('a', 'nav-link', 'Googleマップで道案内を開く');
    link.href = 'https://www.google.com/maps/dir/?api=1' + origin + '&destination=' + dest + '&travelmode=walking';
    link.target = '_blank';
    link.rel = 'noopener';
    box.appendChild(link);

    box.hidden = false;
  }

  function renderCandidates() {
    ui.candList.textContent = '';
    state.ranked.forEach((r, i) => {
      const li = el('li');
      const btn = el('button', 'cand-item');
      btn.type = 'button';
      btn.dataset.id = r.aed.id;

      btn.appendChild(el('span', 'rank', String(i + 1)));

      const nm = el('span', 'nm', r.aed.name);
      if (r.aed.is24h) nm.appendChild(el('span', 'tag24', '24h'));
      btn.appendChild(nm);

      const d = formatDistance(r.info ? r.info.distance : r.distance);
      btn.appendChild(el('span', 'ds', (r.approx ? '約' : '') + d.value + d.unit));

      btn.addEventListener('click', () => showAED(r.aed));
      li.appendChild(btn);
      ui.candList.appendChild(li);
    });
    ui.candidates.hidden = false;
    markActiveCandidate();
  }

  function markActiveCandidate() {
    ui.candList.querySelectorAll('.cand-item').forEach((btn) => {
      const active = state.selected && Number(btn.dataset.id) === state.selected.id;
      if (active) btn.setAttribute('aria-current', 'true');
      else btn.removeAttribute('aria-current');
    });
  }

  /* ===== 初期化 ===== */
  function bindEvents() {
    ui.btnLocate.addEventListener('click', locate);
    ui.btnPick.addEventListener('click', () => {
      const next = !state.picking;
      setPicking(next);
      if (!next) setStatus('');
    });
    ui.chk24.addEventListener('change', () => {
      state.only24 = ui.chk24.checked;
      refreshIcons();
      if (state.start) search();
    });
  }

  async function init() {
    initMap();
    bindEvents();

    try {
      const { aeds, skipped } = await loadAEDs();
      state.aeds = aeds;
      addMarkers();

      ui.btnLocate.disabled = false;
      ui.btnPick.disabled = false;
      ui.dataInfo.textContent = 'AEDデータ: ' + aeds.length + '件を登録' +
        (skipped ? '（座標不備の' + skipped + '件を除く）' : '');
      setStatus('ボタンを押すと、現在地から最も近いAEDを探します。');
    } catch (e) {
      const messages = {
        fetch: 'AED.xlsx を読み込めません。GitHub Pages上、またはローカルサーバー（例: python -m http.server）で開いてください。',
        columns: 'AED.xlsx に「緯度」「経度」の列が見つかりません。1行目の見出しを確認してください。',
        empty: 'AED.xlsx に有効なデータ（施設名・緯度・経度）がありません。'
      };
      setStatus(messages[e && e.message] || messages.fetch, 'error');
      console.error(e);
    }
  }

  init();
})();
