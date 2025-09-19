document.addEventListener('DOMContentLoaded', function () {
const map = new maplibregl.Map({
    container: 'map',
    style: 'https://raw.githubusercontent.com/NISRA-Tech-Lab/map_tiles/main/basemap_styles/style-omt.json',
    center: [-6.8, 54.65],
    zoom: 7.5,
    minZoom: 7.5,
    maxZoom: 13,
    maxBounds: [[-9.20, 53.58], [-4.53, 55.72]]
});

let activeZone = 'sdz';
let drawToolActive = false;

// Event listener for choosing SDZ or DZ
let currentZoneType = 'sdz';
const zoneSelector = document.getElementById('zone-selector');
zoneSelector.value = 'sdz';

zoneSelector.addEventListener('change', onZoneChange);
function onZoneChange(e) {
  const selected = e.target.value;                 // 'sdz' | 'dz' | 'dea'
  activeZone = selected;
  currentZoneType = selected;
  window.selectedZoneType = selected;
  updateSourceLink();

  const vis = (id, show) => map.setLayoutProperty(id, 'visibility', show ? 'visible' : 'none');
  vis('sdz-fill', selected === 'sdz');  vis('dz-fill', selected === 'dz');   vis('dea-fill', selected === 'dea');
  vis('sdz-outline-default', selected === 'sdz');  vis('dz-outline-default', selected === 'dz');  vis('dea-outline-default', selected === 'dea');
  vis('sdz-outline-hover', selected === 'sdz');    vis('dz-outline-hover', selected === 'dz');    vis('dea-outline-hover', selected === 'dea');

  AREA_INDEX[activeZone] = null;
  populateDatalist(activeZone);

  // Clear everything and reset UI
  clearSelections();

  const defaultCategories = ['Age (7 Categories)', 'Sex Label'];
  window.chosenCategories = defaultCategories;
  selectedCategories = defaultCategories;

  document.querySelectorAll('#category-form input[type="checkbox"]').forEach(cb => {
    cb.checked = defaultCategories.includes(cb.value);
  });

  updateTables([]);        // reset tables/outputs
  popup.remove();

  syncPreviewVisibility();
  updateSummaryPreview();  
  ensureSummaryHero();
}

// Initial link setup on page load
updateSourceLink();

const selectedIds = new Set();
let sdzData = {};
let dzData = {};
let deaData = {};
let niTotals = {};
const popup = new maplibregl.Popup({ closeButton: false, closeOnClick: false });
const selectedLGDs = new Set();

fetch('./data.json')

.then(response => response.json())
.then(data => {  
    Year_Data = data 
    sdzData = data["Super Data Zone"] || {};
    dzData = data["Data Zone"] || {};
    deaData = data["District Electoral Area"] || {};
    niTotals = data["NI Total"] || {};
    window.niTotals = niTotals; 
    AREA_INDEX.sdz = AREA_INDEX.dz = AREA_INDEX.dea = null;
    ensureIndexFor(activeZone);
    populateDatalist(activeZone);
decorateCategoryBadges()    
});

// Function that toggles urban/rural fill based on zone selection
let fillVisible = true;

const AREA_INDEX = { sdz: null, dz: null, dea: null }; // built on demand

function getDataSourceFor(zone) {
  return zone === 'dz' ? dzData : zone === 'dea' ? deaData : sdzData;
}
function getLabelKeyFor(zone) {
  return zone === 'dz'
    ? "Census 2021 Data Zone Label"
    : zone === 'dea'
      ? "District Electoral Area 2014 Label"
      : "Census 2021 Super Data Zone Label";
}
function getZoneIdsFor(zone) {
  if (zone === 'dz')  return { source:'dz2021',  sourceLayer:'DZ2021_clipped',  fillLayer:'dz-fill'  };
  if (zone === 'dea') return { source:'dea2014', sourceLayer:'DEA2014_clipped', fillLayer:'dea-fill' };
  return                { source:'sdz2021', sourceLayer:'SDZ2021_clipped',      fillLayer:'sdz-fill' };
}

let previewMap = null;
let previewReady = false;
let previewActiveZone = null;
const previewSelectedIds = new Set();

function ensureSummaryHero() {
  const breakdownContainer = document.getElementById('breakdown-container');
  if (!breakdownContainer || document.getElementById('summary-map')) return;

  const mapDiv = document.createElement('div');
  mapDiv.id = 'summary-map';
  mapDiv.style.width = '260px';
  mapDiv.style.height = '180px';
  mapDiv.style.borderRadius = '8px';
  mapDiv.style.overflow = 'hidden';
  mapDiv.style.border = '1px solid #ccc';
  mapDiv.style.position = 'absolute'; // for bottom-left positioning
  mapDiv.style.bottom = '400px';
  mapDiv.style.left = '20px';

  breakdownContainer.style.position = 'relative'; // required for absolute positioning
  breakdownContainer.appendChild(mapDiv);

  setTimeout(() => {
    initSummaryPreviewMap();
  }, 100);
}

// Initialize the small preview map
function initSummaryPreviewMap() {
  const container = document.getElementById('summary-map');
  if (!container || previewMap) return;

  previewMap = new maplibregl.Map({
    container: 'summary-map',
    style: {
      version: 8,
      sources: {
        osmRaster: {
          type: 'raster',
          tiles: ['https://tile.openstreetmap.org/{z}/{x}/{y}.png'],
          tileSize: 256,
          attribution: '© OpenStreetMap'
        },
        sdz2021: {
          type: 'vector',
          tiles: ['https://raw.githubusercontent.com/nisra-explore/map_tiles/main/sdz_2021/{z}/{x}/{y}.pbf'],
          promoteId: 'sdz_code'
        },
        dz2021: {
          type: 'vector',
          tiles: ['https://raw.githubusercontent.com/nisra-explore/map_tiles/main/dz_2021/{z}/{x}/{y}.pbf'],
          promoteId: 'dz_code'
        },
        dea2014: {
          type: 'vector',
          tiles: ['https://raw.githubusercontent.com/nisra-explore/map_tiles/main/dea_2014/{z}/{x}/{y}.pbf'],
          promoteId: 'dea_code'
        }
      },
      layers: [
        { id:'bg', type:'raster', source:'osmRaster' }
      ]
    },
    center: [-6.8, 54.65],
    zoom: 7.5,
    interactive: false, // no interactions
    dragPan: false, scrollZoom: false, boxZoom: false, keyboard: false, doubleClickZoom: false, touchZoomRotate: false
  });

  previewMap.on('load', () => {
    // Add the three zone layers (fill highlight + outline)
    function addZoneLayers(idPrefix, src, srcLayer) {
      previewMap.addLayer({
        id: `${idPrefix}-fill`,
        type: 'fill',
        source: src,
        'source-layer': srcLayer,
        layout: { visibility: 'none' },
        paint: {
          'fill-color': '#3878c5',
          'fill-opacity': [
            'case',
            ['boolean', ['feature-state', 'hovered'], false], 0.35, 0
          ]
        }
      });
      previewMap.addLayer({
        id: `${idPrefix}-outline`,
        type: 'line',
        source: src,
        'source-layer': srcLayer,
        layout: { visibility: 'none' },
        paint: { 'line-color': '#666', 'line-width': 1 }
      });
    }

    addZoneLayers('sdz-preview',  'sdz2021',  'SDZ2021_clipped');
    addZoneLayers('dz-preview',   'dz2021',   'DZ2021_clipped');
    addZoneLayers('dea-preview',  'dea2014',  'DEA2014_clipped');

    previewReady = true;
    previewActiveZone = activeZone;
    syncPreviewVisibility();
    updateSummaryPreview(); 
  });
}

// Toggle which zone layers are visible in the preview
function syncPreviewVisibility() {
  if (!previewReady) return;
  const showSDZ = activeZone === 'sdz';
  const showDZ  = activeZone === 'dz';
  const showDEA = activeZone === 'dea';
  previewMap.setLayoutProperty('sdz-preview-fill',     'visibility', showSDZ ? 'visible' : 'none');
  previewMap.setLayoutProperty('sdz-preview-outline',  'visibility', showSDZ ? 'visible' : 'none');
  previewMap.setLayoutProperty('dz-preview-fill',      'visibility', showDZ  ? 'visible' : 'none');
  previewMap.setLayoutProperty('dz-preview-outline',   'visibility', showDZ  ? 'visible' : 'none');
  previewMap.setLayoutProperty('dea-preview-fill',     'visibility', showDEA ? 'visible' : 'none');
  previewMap.setLayoutProperty('dea-preview-outline',  'visibility', showDEA ? 'visible' : 'none');
}

// Apply selection highlight + fit bounds in the preview
function waitForPreviewTiles(cb, tries = 12) {
  if (!previewMap) return;

  const ready = typeof previewMap.areTilesLoaded === 'function'
    ? previewMap.areTilesLoaded()
    : true;
  if (ready) {
    return requestAnimationFrame(cb); // 1 frame for placement
  }
  if (tries <= 0) return;
  previewMap.once('render', () => waitForPreviewTiles(cb, tries - 1));
}

function updateSummaryPreview() {
  if (!previewReady) return;

  const z = activeZone;
  const { source, sourceLayer } = getZoneIdsFor(z);
  const layerId = z === 'sdz' ? 'sdz-preview-fill' : z === 'dz' ? 'dz-preview-fill' : 'dea-preview-fill';

  // zone switch: clear old feature-state & show correct layers
  if (previewActiveZone !== z) {
    const { source: oldSrc, sourceLayer: oldSL } = getZoneIdsFor(previewActiveZone || z);
    Array.from(previewSelectedIds).forEach(id => {
      try { previewMap.setFeatureState({ source: oldSrc, sourceLayer: oldSL, id }, { hovered: false }); } catch {}
    });
    previewSelectedIds.clear();
    previewActiveZone = z;
    syncPreviewVisibility();
  }

  // sync feature-state for current selection 
  const sel = new Set(Array.from(selectedIds).map(String));
  Array.from(previewSelectedIds).forEach(id => {
    if (!sel.has(String(id))) {
      try { previewMap.setFeatureState({ source, sourceLayer, id }, { hovered: false }); } catch {}
      previewSelectedIds.delete(id);
    }
  });
  sel.forEach(id => {
    if (!previewSelectedIds.has(id)) {
      try { previewMap.setFeatureState({ source, sourceLayer, id }, { hovered: true }); } catch {}
      previewSelectedIds.add(id);
    }
  });

  // nothing selected -> reset
  if (sel.size === 0) {
    previewMap.easeTo({ center: [-6.8, 54.65], zoom: 7.5, duration: 0 });
    return;
  }

  // ensure we’re zoomed out enough that all tiles for NI are present
  if (previewMap.getZoom() > 7.3) {
    previewMap.jumpTo({ center: [-6.8, 54.65], zoom: 7.3 });
  }
  previewMap.resize();
  
    previewMap.jumpTo({
    center: map.getCenter(),
    zoom: map.getZoom()
    });

  waitForPreviewTiles(() => {
    // Prefer source-level query so we get features even if off-screen
    let feats = [];
    try {
      feats = previewMap.querySourceFeatures(source, { sourceLayer });
    } catch (_) { /* some builds throw; will fall back below */ }

    // Filter to our selected ids
    const picked = feats.filter(f => sel.has(String(f.id)));

    // Fallback: if source query gave nothing (older builds), use rendered layer
    const useFeats = picked.length
      ? picked
      : previewMap.queryRenderedFeatures({ layers: [layerId] }).filter(f => sel.has(String(f.id)));

    if (!useFeats.length) return;

    // Union bbox
    let minX =  Infinity, minY =  Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const f of useFeats) {
      try {
        const bb = turf.bbox({ type: 'Feature', geometry: f.geometry, properties: {} });
        minX = Math.min(minX, bb[0]); minY = Math.min(minY, bb[1]);
        maxX = Math.max(maxX, bb[2]); maxY = Math.max(maxY, bb[3]);
      } catch {}
    }
    if (isFinite(minX)) {
      previewMap.fitBounds([[minX, minY], [maxX, maxY]], { padding: 18, duration: 0 });
    }
  });
}

// Build a lightweight search index

function buildAreaIndexFor(zone) {
  const dataSource = getDataSourceFor(zone);
  const labelKey   = getLabelKeyFor(zone);

  const byKey  = new Map(); // id -> item, lgd:* -> lgd shim
  const byName = new Map(); // name -> [items]
  const items  = [];

  if (!dataSource || !Object.keys(dataSource).length) {
    AREA_INDEX[zone] = { byKey, byName, items };
    return;
  }

  for (const [idRaw, rec] of Object.entries(dataSource)) {
    const id    = isNaN(+idRaw) ? idRaw : +idRaw;
    const lobj  = rec?.[labelKey] || {};
    const name  = Object.keys(lobj)[0] || String(id);
    const lgd   = rec?.LGD || '';

    const item = { id, name, lgd, zone, bbox: null, center: null };
    items.push(item);

    // id lookup
    byKey.set(String(id).toLowerCase(), item);

    // name lookup (allow duplicates)
    const nkey = name.toLowerCase();
    if (!byName.has(nkey)) byName.set(nkey, []);
    byName.get(nkey).push(item);

    // LGD group lookup
    if (lgd) byKey.set(`lgd:${lgd.toLowerCase()}`, { type: 'lgd', lgd, zone });
  }

  AREA_INDEX[zone] = { byKey, byName, items };
}

function ensureIndexFor(zone) {
  if (!AREA_INDEX[zone]) buildAreaIndexFor(zone);
}

function populateDatalist(zone) {
  ensureIndexFor(zone);
  const dl = document.getElementById('apb-area-list');
  if (!dl) return;
  dl.innerHTML = '';

  const idx = AREA_INDEX[zone];
  if (!idx) return;

  // LGDs from the whole dataset
  const dataSource = getDataSourceFor(zone);
  const lgds = new Set(Object.values(dataSource || {}).map(r => r?.LGD).filter(Boolean));
  Array.from(lgds).sort().forEach(lgd => {
    const opt = document.createElement('option');
    opt.value = `LGD: ${lgd}`;
    opt.label = `LGD: ${lgd} (select all in ${zone.toUpperCase()})`;
    dl.appendChild(opt);
  });

  // Names (all areas). Keep value as the plain name; label shows LGD to help disambiguate.
  idx.items
    .sort((a,b)=> a.name.localeCompare(b.name) || a.lgd.localeCompare(b.lgd))
    .forEach(it => {
      const opt = document.createElement('option');
      opt.value = it.name;
      opt.label = it.lgd ? `${it.name} — ${it.lgd}` : it.name;
      dl.appendChild(opt);
    });
}

function flyToBbox(bbox) {
  if (!bbox) return;
  try {
    map.fitBounds(bbox, { padding: 40, duration: 600 });
  } catch {}
}

// Add a single id to current selection (no clearing)
function addSelectById(zone, id) {
  const { source, sourceLayer } = getZoneIdsFor(zone);
  if (!selectedIds.has(id)) {
    selectedIds.add(id);
    map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
  }
}

// Add all areas for LGD (for current active zone)
function addSelectByLGD(zone, lgd) {
  const dataSource = getDataSourceFor(zone);
  const { source, sourceLayer } = getZoneIdsFor(zone);

  Object.entries(dataSource || {}).forEach(([id, rec]) => {
    if (rec?.LGD === lgd) {
      if (!selectedIds.has(id)) {
        selectedIds.add(id);
        map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
      }
    }
  });
}

// After selection changes, update your existing UI
function refreshOutputs() {
  const arr = Array.from(selectedIds);
  window.selectedIdsExcel = selectedIds;
  updateTables(arr);
  renderZoneBreakdownTable(arr);
  updateCtaEnabled();
  updateSummaryPreview();  
  ensureSummaryHero(); // <-- Add this here

}


map.on('load', () => {

    map.addSource('sdz2021', {
        type: 'vector',
        tiles: [
        'https://raw.githubusercontent.com/nisra-explore/map_tiles/main/sdz_2021/{z}/{x}/{y}.pbf'
        ],
        promoteId: { 'SDZ2021_clipped': 'sdz_code' }
    });

    map.addSource('dz2021', {
        type: 'vector',
        tiles: [
        'https://raw.githubusercontent.com/nisra-explore/map_tiles/main/dz_2021/{z}/{x}/{y}.pbf'
        ],
        promoteId: { 'DZ2021_clipped': 'dz_code' }
    });

    map.addSource('dea2014', {
        type: 'vector',
        tiles: [
        'https://raw.githubusercontent.com/nisra-explore/map_tiles/main/dea_2014/{z}/{x}/{y}.pbf'
        ],
        promoteId: { 'DEA2014_clipped': 'dea_code' }
    });

    map.addLayer({
    id: 'sdz-fill',
    type: 'fill',
    source: 'sdz2021',
    'source-layer': 'SDZ2021_clipped',            
    layout: {
        visibility: 'visible'
    },
    paint: {
        'fill-color': '#3878c5', 
        'fill-opacity': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], 0.35, 
        0          
        ]
    }
    });

    map.addLayer({
    id: 'dz-fill',
    type: 'fill',
    source: 'dz2021',
    'source-layer': 'DZ2021_clipped',            
    layout: {
        visibility: 'none'
    },
    paint: {
        'fill-color': '#3878c5', 
        'fill-opacity': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], 0.35, 
        0          
        ]
    }
    });

    map.addLayer({
    id: 'dea-fill',
    type: 'fill',
    source: 'dea2014',
    'source-layer': 'DEA2014_clipped',            
    layout: {
        visibility: 'none'
    },
    paint: {
        'fill-color': '#3878c5', 
        'fill-opacity': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], 0.35, 
        0          
        ]
    }
    });

    map.addLayer({
    id: 'sdz-outline-default',
    type: 'line',
    layout: {
        visibility: 'visible'
        },
    source: 'sdz2021',
    'source-layer': 'SDZ2021_clipped',
    paint: {
        'line-color': '#666666',
        'line-width': 1
    }
    });

    map.addLayer({
    id: 'dz-outline-default',
    type: 'line',
    layout: {
        visibility: 'none' 
        },
    source: 'dz2021',
    'source-layer': 'DZ2021_clipped',
    paint: {
        'line-color': '#666666',
        'line-width': 1
    }
    });

    map.addLayer({
    id: 'dea-outline-default',
    type: 'line',
    layout: {
        visibility: 'none' 
        },
    source: 'dea2014',
    'source-layer': 'DEA2014_clipped',
    paint: {
        'line-color': '#666666',
        'line-width': 1
    }
    });

    map.addLayer({
    id: 'sdz-outline-hover',
    type: 'line',
    source: 'sdz2021',
    'source-layer': 'SDZ2021_clipped',
    paint: {
        'line-color': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], '#000000',
        'rgba(0,0,0,0)'
        ],
        'line-width': 1.5
    }
    });

    map.addLayer({
    id: 'dz-outline-hover',
    type: 'line',
    layout: {
        visibility: 'none'
        },
    source: 'dz2021',
    'source-layer': 'DZ2021_clipped',
    paint: {
        'line-color': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], '#000000',
        'rgba(0,0,0,0)'
        ],
        'line-width': 1.5
    }
    });

    map.addLayer({
    id: 'dea-outline-hover',
    type: 'line',
    layout: {
        visibility: 'none'
        },
    source: 'dea2014',
    'source-layer': 'DEA2014_clipped',
    paint: {
        'line-color': [
        'case',
        ['boolean', ['feature-state', 'hovered'], false], '#000000',
        'rgba(0,0,0,0)'
        ],
        'line-width': 1.5
    }
    });

    initSummaryPreviewMap();
    syncPreviewVisibility();
    updateSummaryPreview();

    map.on('mousemove', 'sdz-fill', (e) => {
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;
    const data = sdzData?.[id];

    let popupHtml = '<div><strong>Data not found</strong></div>';
    if (data?.["Census 2021 Super Data Zone Label"]) {
        const labelObj = data["Census 2021 Super Data Zone Label"];
        const zoneName = Object.keys(labelObj)[0];
        popupHtml = `<div><strong>${zoneName}</strong>: ${labelObj[zoneName]}</div>`;
    }

    map.getCanvas().style.cursor = 'pointer';
    popup.setLngLat(e.lngLat).setOffset([0,-10]).setHTML(popupHtml).addTo(map);
    });

    map.on('mousemove', 'dz-fill', (e) => {
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;
    const data = dzData?.[id];

    let popupHtml = '<div><strong>Data not found</strong></div>';
    if (data?.["Census 2021 Data Zone Label"]) {
        const labelObj = data["Census 2021 Data Zone Label"];
        const zoneName = Object.keys(labelObj)[0];
        popupHtml = `<div><strong>${zoneName}</strong>: ${labelObj[zoneName]}</div>`;
    }

    map.getCanvas().style.cursor = 'pointer';
    popup.setLngLat(e.lngLat).setOffset([0,-10]).setHTML(popupHtml).addTo(map);
    });

    map.on('mousemove', 'dea-fill', (e) => {
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;
    const data = deaData?.[id];

    let popupHtml = '<div><strong>Data not found</strong></div>';
    if (data?.["District Electoral Area 2014 Label"]) {
        const labelObj = data["District Electoral Area 2014 Label"];
        const zoneName = Object.keys(labelObj)[0];
        popupHtml = `<div><strong>${zoneName}</strong>: ${labelObj[zoneName]}</div>`;
    }

    map.getCanvas().style.cursor = 'pointer';
    popup.setLngLat(e.lngLat).setOffset([0,-10]).setHTML(popupHtml).addTo(map);
    });

    map.on('mouseleave', 'sdz-fill', () => {
    map.getCanvas().style.cursor = '';
    popup.remove();
    });

    map.on('mouseleave', 'dz-fill', () => {
    map.getCanvas().style.cursor = '';
    popup.remove();
    });

    map.on('mouseleave', 'dea-fill', () => {
    map.getCanvas().style.cursor = '';
    popup.remove();
    });

    map.on('click', 'sdz-fill', (e) => {
    if (drawToolActive) return;
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;

    const isSelected = selectedIds.has(id);

    if (isSelected) {
        selectedIds.delete(id);
        map.setFeatureState(
        { source: 'sdz2021', sourceLayer: 'SDZ2021_clipped', id },
        { hovered: false }
        );
    } else {
        selectedIds.add(id);
        map.setFeatureState(
        { source: 'sdz2021', sourceLayer: 'SDZ2021_clipped', id },
        { hovered: true }
        );
    }

    let selectedTab = document.querySelector('.view-tab.selected');
    if (!selectedTab) {
        const chartsTab = document.querySelector('.view-tab[data-view="charts"]');
        chartsTab.classList.add("selected");
    }

    document.getElementById("charts-container").style.display = "flex";
    document.getElementById("tables-container").style.display = "none";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";

    window.selectedIdsExcel = selectedIds;

    updateTables(Array.from(selectedIds));
    updateCtaEnabled();
    updateSummaryPreview();
    });

    map.on('click', 'dz-fill', (e) => {
    if (drawToolActive) return;
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;

    const isSelected = selectedIds.has(id);

    if (isSelected) {
        selectedIds.delete(id);
        map.setFeatureState(
        { source: 'dz2021', sourceLayer: 'DZ2021_clipped', id },
        { hovered: false }
        );
    } else {
        selectedIds.add(id);
        
        map.setFeatureState(
        { source: 'dz2021', sourceLayer: 'DZ2021_clipped', id },
        { hovered: true }
        );
    }

    let selectedTab = document.querySelector('.view-tab.selected');
    if (!selectedTab) {
        const chartsTab = document.querySelector('.view-tab[data-view="charts"]');
        chartsTab.classList.add("selected");
    }

    document.getElementById("charts-container").style.display = "flex";
    document.getElementById("tables-container").style.display = "none";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";

    window.selectedIdsExcel = selectedIds;

    updateTables(Array.from(selectedIds));
    updateCtaEnabled();
    updateSummaryPreview();
    });

    map.on('click', 'dea-fill', (e) => {
    if (drawToolActive) return;
    if (!e.features.length) return;

    const feature = e.features[0];
    const id = feature.id;

    const isSelected = selectedIds.has(id);

    if (isSelected) {
        selectedIds.delete(id);
        map.setFeatureState(
        { source: 'dea2014', sourceLayer: 'DEA2014_clipped', id },
        { hovered: false }
        );
    } else {
        selectedIds.add(id);
        
        map.setFeatureState(
        { source: 'dea2014', sourceLayer: 'DEA2014_clipped', id },
        { hovered: true }
        );
    }

    let selectedTab = document.querySelector('.view-tab.selected');
    if (!selectedTab) {
        const chartsTab = document.querySelector('.view-tab[data-view="charts"]');
        chartsTab.classList.add("selected");
    }

    document.getElementById("charts-container").style.display = "flex";
    document.getElementById("tables-container").style.display = "none";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";

    window.selectedIdsExcel = selectedIds;

    updateTables(Array.from(selectedIds));
    updateCtaEnabled();
    updateSummaryPreview();
    });

    map.addSource('draw-geom', {
    type: 'geojson',
    data: { type:'FeatureCollection', features: [] }
    });
    map.addLayer({
    id: 'draw-fill',
    type: 'fill',
    source: 'draw-geom',
    paint: { 'fill-color': '#ff8800', 'fill-opacity': 0.15 }
    });
    map.addLayer({
    id: 'draw-line',
    type: 'line',
    source: 'draw-geom',
    paint: { 'line-color': '#ff8800', 'line-width': 2 }
    });

    let lastDrawnFeature = null;  // keep the most recent boundary user created
    addDrawToolbar();

    // pick correct source/layer ids for current zone type
    function zoneIds() {
    if (activeZone === 'dz')   return { source:'dz2021',  sourceLayer:'DZ2021_clipped',  fillLayer:'dz-fill'  };
    if (activeZone === 'dea')  return { source:'dea2014', sourceLayer:'DEA2014_clipped', fillLayer:'dea-fill' };
    return                       { source:'sdz2021', sourceLayer:'SDZ2021_clipped',  fillLayer:'sdz-fill'  };
    }

    function selectByGeometry(geom, mode = 'add') {
        const feature = geom.type === 'Feature' ? geom : { type:'Feature', geometry: geom, properties:{} };

        const bbox = turf.bbox(feature);
        const sw = map.project([bbox[0], bbox[1]]);
        const ne = map.project([bbox[2], bbox[3]]);
        const { source, sourceLayer, fillLayer } = zoneIds();

        const candidates = map.queryRenderedFeatures([sw, ne], { layers: [fillLayer] })
        .filter(f => {
            try {
            return turf.booleanContains(feature.geometry, f.geometry);
            } catch {
            return false;
            }
        });


        if (mode === 'replace') {
            selectedIds.forEach(id => map.setFeatureState({ source, sourceLayer, id }, { hovered: false }));
            selectedIds.clear();
        }

        if (mode === 'subtract') {
            candidates.forEach(f => {
            const id = f.id;
            if (selectedIds.has(id)) {
                selectedIds.delete(id);
                map.setFeatureState({ source, sourceLayer, id }, { hovered: false });
            }
            });
        } else {
            // 'add' (default) or after 'replace'
            candidates.forEach(f => {
            const id = f.id;
            if (!selectedIds.has(id)) {
                selectedIds.add(id);
                map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
            }
            });
        }

        const arr = Array.from(selectedIds);
        window.selectedIdsExcel = selectedIds;
        updateTables(arr);
        renderZoneBreakdownTable(arr);
        updateCtaEnabled();
    }


    // simplify in metres
    function simplifyMeters(feature, toleranceM = 10) {
    const merc = turf.toMercator(feature);
    const simp = turf.simplify(merc, { tolerance: Math.max(0, toleranceM), highQuality: false });
    return turf.toWgs84(simp);
    }

    // Download helper
    function downloadGeoJSON(feat, filename = 'custom-area.geojson') {
    const blob = new Blob([JSON.stringify(feat)], { type: 'application/geo+json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
    }

window.createDrawToolbar = function (targetElementId = "draw-toolbar-container") {
  const target = document.getElementById(targetElementId);
  if (!target) {
    console.warn("Target element for draw toolbar not found:", targetElementId);
    return;
  }

  // Inject CSS once
  if (!document.getElementById("draw-toolbar-styles")) {
    const css = document.createElement("style");
    css.id = "draw-toolbar-styles";
    css.textContent = `
      #draw-toolbar {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 6px 8px;
        background: rgba(0,0,0,.78);
        border-radius: 10px;
        color: #fff;
        font: 14px/1.2 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
        margin-top: 1rem;
      }
      #draw-toolbar .icon-btn {
        width: 36px;
        height: 36px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: transparent;
        border: 1px solid rgba(255,255,255,.15);
        border-radius: 8px;
        color: #fff;
        cursor: pointer;
      }
      #draw-toolbar .icon-btn:hover {
        background: rgba(255,255,255,.12);
      }
      #draw-toolbar .icon-btn[data-badge]::after {
        content: attr(data-badge);
        position: absolute;
        right: -60px;
        bottom: -60px;
        min-width: 18px;
        height: 18px;
        padding: 0 3px;
        background: #1ea672;
        color: #fff;
        border-radius: 9px;
        font: 600 11px/18px system-ui;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,.3);
      }
    `;
    document.head.appendChild(css);
  }

  // Create toolbar container
  const bar = document.createElement("div");
  bar.id = "draw-toolbar";

  // SVG icons
  const svgs = {
    circle: `<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/></svg>`,
    zoomIn: `<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.35-4.35M11 8v6M8 11h6"/></svg>`,
    zoomOut: `<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.35-4.35M8 11h6"/></svg>`,
    radius: `<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="2"/></svg>`
  };

  // Helper to create buttons
  const makeBtn = (id, title, svg, badge = false) => {
    const btn = document.createElement("button");
    btn.className = "icon-btn";
    btn.id = id;
    btn.title = title;
    btn.innerHTML = svg;
    if (badge) btn.setAttribute("data-badge", "2k");
    return btn;
  };

  // Add buttons
  bar.appendChild(makeBtn("circleSelectBtn", "Circle select", svgs.circle));
  bar.appendChild(makeBtn("radiusBtn", "Radius (km)", svgs.radius, true));
  bar.appendChild(makeBtn("zoomInBtn", "Zoom in", svgs.zoomIn));
  bar.appendChild(makeBtn("zoomOutBtn", "Zoom out", svgs.zoomOut));

  // Append toolbar to target
  target.appendChild(bar);
  console.log("Draw toolbar inserted into #draw-toolbar-container");
};

function addDrawToolbar() {
        const mapEl = map.getContainer();
        if (!mapEl) return;
        if (getComputedStyle(mapEl).position === 'static') {
            mapEl.style.position = 'relative';
        }

        // Inject CSS once
        if (!document.getElementById('draw-toolbar-styles')) {
            const css = document.createElement('style');
            css.id = 'draw-toolbar-styles';
            css.textContent = `
#draw-toolbar {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 8px;
  background: rgba(0,0,0,.78);
  border-radius: 10px;
  color: #fff;
  font: 14px/1.2 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  margin-top: 1rem;  
margin: 1rem auto;
  width: fit-content; /* or max-width: 100%; */

}
#draw-toolbar .icon-btn {
  width: 36px;
  height: 36px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: transparent;
  border: 1px solid rgba(255,255,255,.15);
  border-radius: 8px;
  color: #fff;
  cursor: pointer;
}
#draw-toolbar .icon-btn:hover {
  background: rgba(255,255,255,.12);
}
#draw-toolbar .icon-btn.active {
  outline: 2px solid rgba(255,255,255,.35);
}
#draw-toolbar .icon-btn svg,
#draw-toolbar .icon-btn svg * {
  width: 18px;
  height: 18px;
  stroke: currentColor;
  fill: none;
  stroke-width: 2;
}
#draw-toolbar .icon-btn[data-badge]::after {
  content: attr(data-badge);
  position: absolute;
  right: 196px;
  bottom: 106px;
  min-width: 18px;
  height: 18px;
  padding: 0 3px;
  background: #1ea672;
  color: #fff;
  border-radius: 9px;
  font: 600 11px/18px system-ui;
  text-align: center;
  box-shadow: 0 2px 6px rgba(0,0,0,.3);
}
`;
            document.head.appendChild(css);
        }


const bar = document.createElement('div');
bar.id = 'draw-toolbar';

const container = document.getElementById('draw-toolbar-container');
if (container) {
  container.appendChild(bar);
  console.log("Toolbar added to #draw-toolbar-container");
} else {
  mapEl.appendChild(bar);
  console.log("Toolbar added to map container");
}

        // Icons
        const svgs = {
            circle: `<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"/></svg>`,
            lasso:  `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M4 7c0-2 3-4 8-4s8 2 8 4-3 4-8 4c-3 0-5 .7-5 2s2 2 5 2"/>
                    <path d="M9 15c0 2-1.5 5-4 5"/></svg>`,
            upload: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 3v12"/><path d="M8 7l4-4 4 4"/>
                    <rect x="4" y="15" width="16" height="6" rx="2"/></svg>`,
            download:`<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 3v12"/><path d="M16 11l-4 4-4-4"/>
                    <rect x="4" y="17" width="16" height="4" rx="2"/></svg>`,
            trash:  `<svg viewBox="0 0 24 24" aria-hidden="true">
                <g stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <polygon points="6 4 18 4 22 12 18 20 6 20 2 12"/>
                <path d="M9 9l6 6M15 9l-6 6"/>
                </g>
            </svg>`,
            radius: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="2"/></svg>`,
            simplify:`<svg viewBox="0 0 24 24" aria-hidden="true">
                    <polyline points="3,16 8,8 13,14 21,6"/></svg>`,
            buffer: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="9"/></svg>`,
            zoomIn: `
            <svg viewBox="0 0 24 24" aria-hidden="true">
            <g stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="11" cy="11" r="7"/>
                <path d="M21 21l-4.35-4.35M11 8v6M8 11h6"/>
            </g>
            </svg>`,
        zoomOut: `
            <svg viewBox="0 0 24 24" aria-hidden="true">
            <g stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="11" cy="11" r="7"/>
                <path d="M21 21l-4.35-4.35M8 11h6"/>
            </g>
            </svg>`,
        home: `
            <svg viewBox="0 0 24 24" aria-hidden="true">
            <g stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <path d="M3 10.5l9-7 9 7"/>
                <path d="M6 10v9h12v-9"/>
            </g>
            </svg>`,
            search: `<svg class="search-icon" viewBox="0 0 24 24" aria-hidden="true">
                    <circle cx="11" cy="11" r="7" stroke-width="2"/><path d="M20 20l-3.5-3.5" stroke-width="2"/></svg>`
        };

        const makeBtn = (id, title, svg, withBadge=false) => {
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'icon-btn'; b.id = id; b.title = title;
            b.innerHTML = svg; bar.appendChild(b);
            if (withBadge) b.setAttribute('data-badge','');
            return b;
        };

        // Tool buttons
        const circleBtn = makeBtn('circleSelectBtn', 'Circle select (click map to select)', svgs.circle);
        // const lassoBtn  = makeBtn('lassoSelectBtn',  'Lasso select (drag to sketch polygon)', svgs.lasso);
        // const uploadBtn = makeBtn('uploadGeoBtn',    'Upload GeoJSON', svgs.upload);
        // const exportBtn = makeBtn('exportGeoBtn',    'Export drawn boundary (applies simplify/buffer)', svgs.download);
        // const clearBtn  = makeBtn('clearGeoBtn',     'Clear drawn boundary', svgs.trash);

        // Parameter buttons (same size with value badge)
        const radiusBtn   = makeBtn('radiusBtn',   'Radius (km): click to cycle, Shift+Click to set', svgs.radius,   true);
        // const simplifyBtn = makeBtn('simplifyBtn', 'Simplify tol (m): click to cycle, Shift+Click to set', svgs.simplify, true);
        // const bufferBtn   = makeBtn('bufferBtn',   'Buffer (m): positive grows, negative shrinks. Click to cycle, Shift+Click to set', svgs.buffer,   true);

        // Zoom controls
        const zoomOutBtn = makeBtn('dtZoomOut',  'Zoom out',  svgs.zoomOut);
        const zoomInBtn  = makeBtn('dtZoomIn',   'Zoom in',   svgs.zoomIn);
        // const homeBtn    = makeBtn('dtZoomHome', 'Reset view', svgs.home);

        // Separator + Search
        // const sepEl = document.createElement('div');
        // sepEl.className = 'dt-sep';
        // bar.appendChild(sepEl);

        // Separator + Search

// Removed search bar from draw toolbar
// const searchWrap = document.createElement('div');
// searchWrap.className = 'dt-search';
// searchWrap.innerHTML = `${svgs.search}
//   <input id="dt-search-input" list="apb-area-list" placeholder="Search area or LGD…" autocomplete="off" />
// `;
// bar.appendChild(searchWrap);


        // datalist -create if not there yet
        // let dl = document.getElementById('apb-area-list');
        // if (!dl) {
        //     dl = document.createElement('datalist');
        //     dl.id = 'apb-area-list';
        //     document.body.appendChild(dl);
        // }

        // Drawing states & params
        let circleMode = false, lassoMode = false, drawing = false;
        let lasso = [];
        let radiusKm = 2, simplifyTolMeters = 10, bufferMeters = 0;

        const setBadge = (btn, txt) => btn.setAttribute('data-badge', txt);
        const refreshBadges = () => {
            setBadge(radiusBtn,   `${radiusKm}k`);
            // setBadge(simplifyBtn, `${simplifyTolMeters}m`);
            // setBadge(bufferBtn,   `${bufferMeters}m`);
        };
        refreshBadges();

        // Hidden file input for upload
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.geojson,.json';
        fileInput.style.display = 'none';
        bar.appendChild(fileInput);

        // // Mode toggles
        // function setCircleMode(on){
        //     circleMode = !!on;
        //     if (circleMode) { lassoMode = false; lassoBtn.classList.remove('active'); }
        //     circleBtn.classList.toggle('active', circleMode);
        //     drawToolActive = circleMode || lassoMode;
        //     map.getCanvas().style.cursor = drawToolActive ? 'crosshair' : '';
        // }
        // function setLassoMode(on){
        //     lassoMode = !!on;
        //     if (lassoMode) { circleMode = false; circleBtn.classList.remove('active'); }
        //     lassoBtn.classList.toggle('active', lassoMode);
        //     drawToolActive = circleMode || lassoMode;
        //     map.getCanvas().style.cursor = drawToolActive ? 'crosshair' : '';
        // }
        // circleBtn.addEventListener('click', () => setCircleMode(!circleMode));
        // lassoBtn .addEventListener('click', () => setLassoMode(!lassoMode));

        // Mode toggle
        function setCircleMode(on){
            circleMode = !!on;
            circleBtn.classList.toggle('active', circleMode);
            drawToolActive = circleMode;
            map.getCanvas().style.cursor = drawToolActive ? 'crosshair' : '';
        }

        circleBtn.addEventListener('click', () => setCircleMode(!circleMode));

        // Circle select
        map.on('click', (e) => {
            if (!circleMode) return;
            const center = [e.lngLat.lng, e.lngLat.lat];
            const circle = turf.circle(center, Math.max(0.05, +radiusKm || 2), { steps: 128, units:'kilometers' });
            map.getSource('draw-geom').setData(circle);
            lastDrawnFeature = circle;

            const oe = e.originalEvent || {};
            const mode = (oe.altKey || oe.ctrlKey || oe.metaKey) ? 'subtract'
                    : (oe.shiftKey ? 'replace' : 'add');
            selectByGeometry(circle, mode);
        });

        // Lasso select
        map.on('mousedown', (e) => {
            if (!lassoMode || e.originalEvent.button !== 0) return;
            drawing = true;
            lasso = [[e.lngLat.lng, e.lngLat.lat]];
            map.dragPan.disable();
        });
        map.on('mousemove', (e) => {
            if (!drawing) return;
            lasso.push([e.lngLat.lng, e.lngLat.lat]);
            if (lasso.length > 2) {
            const ring = [...lasso, lasso[0]];
            map.getSource('draw-geom').setData(turf.polygon([ring]));
            }
        });
        function finishLasso() {
            if (!drawing) { drawToolActive = circleMode || lassoMode; return; }
            drawing = false;
            map.dragPan.enable();
            if (lasso.length > 2) {
            const ring = [...lasso, lasso[0]];
            const poly = turf.polygon([ring]);
            lastDrawnFeature = poly;
            map.getSource('draw-geom').setData(poly);
            selectByGeometry(poly, 'add'); // non-destructive add
            try { map.fitBounds(turf.bbox(poly), { padding: 30, animate: true }); } catch {}
            }
            lasso = [];
            drawToolActive = circleMode || lassoMode;
            map.getCanvas().style.cursor = drawToolActive ? 'crosshair' : '';
        }
        map.on('mouseup',   finishLasso);
        map.on('dragstart', finishLasso);
        map.getCanvas().addEventListener('mouseleave', finishLasso);

        // Escape cancels modes
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') { setCircleMode(false); setLassoMode(false); }
        });

        // Upload
        // uploadBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', () => {
            const file = fileInput.files?.[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
            try {
                const gj = JSON.parse(reader.result);
                const feats = gj.type === 'FeatureCollection' ? gj.features
                        : gj.type === 'Feature' ? [gj] : [];
                if (!feats.length) throw new Error('No features found');
                let merged = null;
                feats.forEach(f => { if (f && f.geometry) merged = merged ? turf.union(merged, f) : f; });
                if (merged) {
                lastDrawnFeature = merged;
                map.getSource('draw-geom').setData(merged);
                try { map.fitBounds(turf.bbox(merged), { padding: 30 }); } catch {}
                selectByGeometry(merged, 'add');
                }
            } catch {
                alert('Invalid GeoJSON file');
            }
            };
            reader.readAsText(file);
            fileInput.value = '';
        });

        // Parameter buttons
        radiusBtn.addEventListener('click', (e) => {
            if (e.shiftKey) {
            const v = prompt('Radius (km):', String(radiusKm));
            if (v !== null && !isNaN(+v) && +v > 0) radiusKm = +v;
            } else {
            const opts = [0.5, 1, 2, 5, 10];
            const i = opts.indexOf(radiusKm);
            radiusKm = opts[(i + 1) % opts.length];
            }
            refreshBadges();
        });
        // simplifyBtn.addEventListener('click', (e) => {
        //     if (e.shiftKey) {
        //     const v = prompt('Simplify tolerance (meters):', String(simplifyTolMeters));
        //     if (v !== null && !isNaN(+v) && +v >= 0) simplifyTolMeters = +v;
        //     } else {
        //     const opts = [0, 5, 10, 25, 50, 100];
        //     const i = opts.indexOf(simplifyTolMeters);
        //     simplifyTolMeters = opts[(i + 1) % opts.length];
        //     }
        //     refreshBadges();
        // });
        // bufferBtn.addEventListener('click', (e) => {
        //     if (e.shiftKey) {
        //     const v = prompt('Buffer (meters, negative shrinks):', String(bufferMeters));
        //     if (v !== null && !isNaN(+v)) bufferMeters = +v;
        //     } else {
        //     const opts = [-100, -50, -10, 0, 10, 50, 100];
        //     const i = opts.indexOf(bufferMeters);
        //     bufferMeters = opts[(i + 1) % opts.length];
        //     }
        //     refreshBadges();
        // });

        // Export/Clear
        // exportBtn.addEventListener('click', () => {
        //     if (!lastDrawnFeature) { alert('Draw or upload a boundary first.'); return; }
        //     let out = lastDrawnFeature;
        //     if (bufferMeters !== 0) {
        //     try { out = turf.buffer(out, bufferMeters, { units:'meters' }); } catch {}
        //     }
        //     if (simplifyTolMeters > 0) out = simplifyMeters(out, simplifyTolMeters);
        //     const blob = new Blob([JSON.stringify(out)], { type:'application/geo+json' });
        //     const a = document.createElement('a');
        //     a.href = URL.createObjectURL(blob);
        //     a.download = 'custom-area.geojson';
        //     a.click();
        //     URL.revokeObjectURL(a.href);
        // });
        // clearBtn.addEventListener('click', () => {
        //     map.getSource('draw-geom').setData({ type:'FeatureCollection', features: [] });
        //     lastDrawnFeature = null;
        // });

        zoomInBtn.addEventListener('click',  () => map.zoomIn({ duration: 250 }));
        zoomOutBtn.addEventListener('click', () => map.zoomOut({ duration: 250 }));
        // homeBtn.addEventListener('click',    () => {
        //     map.easeTo({ center: [-6.8, 54.65], zoom: 7.5, duration: 600 });
        // });

        // SEARCH WIRING
        // const input = document.getElementById('dt-search-input');

        // Guard so we don't double-wire if toolbar is re-inited
        if (!window.__apbSearchWired) {
            window.__apbSearchWired = true;

            // Build initial index & datalist for current active zone
            ensureIndexFor(activeZone);
            populateDatalist(activeZone);

            // Rebuild index when zone changes
            document.getElementById('zone-selector')?.addEventListener('change', () => {
            ensureIndexFor(activeZone);
            populateDatalist(activeZone);
            });

            // Rebuild index & datalist after map renders 
            let rebuildTimer = null;
            function scheduleRebuild() {
            clearTimeout(rebuildTimer);
            rebuildTimer = setTimeout(() => {
                AREA_INDEX[activeZone] = null;
                buildAreaIndexFor(activeZone);
                populateDatalist(activeZone);
            }, 150);
            }
            map.on('idle', scheduleRebuild);
        }

        function handleSearchCommit() {
            const raw = (input.value || '').trim();
            if (!raw) return;

            // LGD bulk select
            if (/^lgd[:\s]/i.test(raw) || raw.startsWith('LGD:')) {
                const lgdName = raw.replace(/^lgd[:\s]*/i, '').trim();
                addSelectByLGD(activeZone, lgdName);
                refreshOutputs();
                input.blur();
                return;
            }

            ensureIndexFor(activeZone);
            const idx = AREA_INDEX[activeZone];
            const key = raw.toLowerCase();

            // try id first
            let hit = idx?.byKey?.get(key);

            // name disambiguation
            if (!hit) {
                const nameMatches = idx?.byName?.get(key);
                if (nameMatches && nameMatches.length) {
                // If user typed “Name — LGD” or “Name, LGD”, use LGD to disambiguate
                const lgdHint = raw.split(/—|-|,|-/).slice(1).join('').trim();
                if (lgdHint) {
                    const byLgd = nameMatches.find(m => m.lgd.toLowerCase() === lgdHint.toLowerCase());
                    hit = byLgd || nameMatches[0];
                } else {
                    // Prefer a visible match if any, fallback to the first
                    const { fillLayer } = getZoneIdsFor(activeZone);
                    const visIds = new Set(map.queryRenderedFeatures({ layers:[fillLayer] }).map(f => f.id));
                    hit = nameMatches.find(m => visIds.has(m.id)) || nameMatches[0];
                }
                }
            }

            // loose substring fallback
            if (!hit && idx?.items?.length) {
                const lc = raw.toLowerCase();
                hit = idx.items.find(it =>
                it.name.toLowerCase().includes(lc) || String(it.id).toLowerCase() === lc
                );
            }

            if (!hit) {
                alert('No matching area found in the current geography level.');
                return;
            }

            // compute bbox on-the-fly
            if (!hit.bbox) {
                const { fillLayer } = getZoneIdsFor(activeZone);
                const feats = map.queryRenderedFeatures({ layers:[fillLayer] }).filter(f => f.id === hit.id);
                if (feats[0]) {
                try {
                    const gj = { type: 'Feature', geometry: feats[0].geometry, properties:{} };
                    hit.bbox = turf.bbox(gj);
                } catch(_) {}
                }
            }

            if (hit.bbox) { try { map.fitBounds(hit.bbox, { padding: 40, duration: 600 }); } catch {} }
            addSelectById(activeZone, hit.id);
            refreshOutputs();
            input.blur();
            }
            // input.addEventListener('change', handleSearchCommit);
            // input.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleSearchCommit(); });
    }

    const buildBtn = document.getElementById("build-profile-btn");
    const changeBtn = document.getElementById("change-selection");
    const outputContent = document.getElementById("output-content");
    const mapContent = document.getElementById("map-content");

    function openSelectModal() {
    document.getElementById('select-areas-modal').hidden = false;
    }
    function closeSelectModal() {
    document.getElementById('select-areas-modal').hidden = true;
    }
    function focusSelectorBox() {
    const box = document.querySelector('.lgd-selector') || document.getElementById('map-wrapper');
    if (!box) return;
    box.scrollIntoView({ behavior: 'smooth', block: 'center' });
    box.style.boxShadow = '0 0 0 4px rgba(4,134,62,.35)';
    setTimeout(() => (box.style.boxShadow = ''), 1200);
    }

    // expose so other code can call it safely
    window.updateCtaEnabled = function () {
    if (buildBtn) buildBtn.disabled = selectedIds.size === 0;
    };

    // wire once (no nesting)
    buildBtn?.addEventListener("click", () => {
    if (selectedIds.size === 0) { openSelectModal(); return; }
    outputContent.classList.remove("hidden-section");
    mapContent.classList.add("hidden-section");

    previewMap?.resize();
    updateSummaryPreview();
    });

    changeBtn?.addEventListener("click", () => {
    outputContent.classList.add("hidden-section");
    mapContent.classList.remove("hidden-section");
    });

    document.getElementById('apb-modal-close')?.addEventListener('click', closeSelectModal);
    document.getElementById('apb-modal-focus')?.addEventListener('click', () => {
    closeSelectModal();
    focusSelectorBox();
    });

    // set initial disabled state
    updateCtaEnabled();

});

// map reset zoom button
document.getElementById('resetZoomBtn').addEventListener('click', () => {
    map.easeTo({
    center: [-6.8, 54.65],
    zoom: 7.5,
    duration: 2000
    });
});

async function updateSourceLink() {
    const zoneType = window.selectedZoneType || 'sdz';
    const selectedLabels = window.chosenCategories || ['Age (7 Categories)', 'Sex Label'];
    const response = await fetch('category_lookup.json');
    const lookup = await response.json();

    const labelToCode = {};
    const labelToSource = {};
    lookup.forEach(item => {
        labelToCode[item.nested_list_names] = item.further_breakdown_df;
        labelToSource[item.nested_list_names] = item.Source; // "Flexible Table Builder" or "Data Portal"
    });

    // make available elsewhere
    window.labelToCode = labelToCode;
    window.labelToSource = labelToSource;

    const sourceLinkContainer = document.getElementById("sourceId");
    sourceLinkContainer.innerHTML = '';

    selectedLabels.forEach(label => {
    const source = labelToSource[label];
    const categoryCode = labelToCode[label];
    let fullUrl = '';

    window.urlZoneType = ZoneType;
    window.urlcategoryCode = categoryCode;
    window.urllabel = label;
    window.urlsource = source;

    if (source === "Flexible Table Builder" && categoryCode) {
    fullUrl = `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}21&v=${categoryCode}`;
    } else if (source === "Data Portal") {
    if (label === "Benefits Statistics" && zoneType === "dz") {
    fullUrl = "https://data.nisra.gov.uk/table/BSDZ";
    } else if (label === "Benefits Statistics" && zoneType === "sdz") {
    fullUrl = "https://data.nisra.gov.uk/table/BSSDZ";
    } else if ((label === "Age (MYE)" || label === "Sex (MYE)") && zoneType === "sdz") {
    fullUrl = "https://data.nisra.gov.uk/table/MYE01T012";
    }
    }

    if (fullUrl) {
    const wrapper = document.createElement('div');
    const zoneTypeText = zoneType === 'sdz' ? 'Super Data Zone' : 'Data Zone';
    const labelText = document.createTextNode(`${label} by ${zoneTypeText}: `);
    const link = document.createElement('a');
    link.href = fullUrl;
    link.textContent = fullUrl;
    link.target = "_blank";
    wrapper.appendChild(labelText);
    wrapper.appendChild(link);
    sourceLinkContainer.appendChild(wrapper);
    } else {
    console.warn(`No valid URL found for label: ${label}`);
    }
    });
    decorateCategoryBadges();
}

function decorateCategoryBadges() {
  const mapSrc = window.labelToSource || {};
  document.querySelectorAll('#category-form label').forEach(label => {
    const input = label.querySelector('input[type="checkbox"]');
    if (!input) return;
    const name = input.value;
    const src = mapSrc[name];

    // remove old badge (if any)
    label.querySelectorAll('.source-badge').forEach(b => b.remove());

    if (!src) return; 

    const badge = document.createElement('span');
    const isPortal = src === 'Data Portal';
    badge.className = 'source-badge ' + (isPortal ? 'badge-portal' : 'badge-ftb');
    badge.textContent = isPortal ? 'DP' : 'FTB';
    badge.title = src;
    label.appendChild(badge);
  });
}


let latestAggregatedData = {}; 
let selectedCategories = ['Age (7 Categories)', 'Sex Label'];
let currentView = 'charts';

document.getElementById("category-form").addEventListener("change", () => {
    selectedCategories = Array.from(document.querySelectorAll('#category-form input:checked'))
    .map(input => input.value);

    const selectedArray = Array.from(selectedIds);
    updateTables(selectedArray);
    
    document.getElementById("charts-container").style.display = "none";
    document.getElementById("tables-container").style.display = "none";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";

    if (selectedCategories.length === 0) return;

    const availableKeys = Object.keys(latestAggregatedData);
    const validCategories = selectedCategories.filter(cat => availableKeys.includes(cat));

    if (validCategories.length === 0) return;

    if (currentView === 'charts') {
    renderAggregatedCharts(latestAggregatedData, validCategories);
    document.getElementById("charts-container").style.display = "flex";
    } else if (currentView === 'tables') {
    renderAggregatedTables(latestAggregatedData, validCategories);
    document.getElementById("tables-container").style.display = "block";
    } else if (currentView === 'tableComparison') {
    renderUrbanRuralComparison(selectedArray);
    document.getElementById("urban-rural-comparison").style.display = "block";
    } else if (currentView === 'chartComparison') {
    renderUrbanRuralCharts(selectedArray);
    document.getElementById("urban-rural-charts").style.display = "block";
    }

    window.chosenCategories = selectedCategories;
    updateSourceLink();

    updateTables(Array.from(selectedIds));
});


    document.getElementById('zone-selector').addEventListener('change', function () {
    currentZoneType = this.value;
    updateTables(Array.from(selectedIds));
});

const selector = document.getElementById("zone-selector");

// Set default value
selector.value = "sdz";

// Listen for changes
selector.addEventListener("change", function () {
    currentZoneType = this.value;
    updateTables(Array.from(selectedIds));
});

// Function to hide categories not present in the data
function updateCategorySelector(availableKeys) {
    const checkboxes = document.querySelectorAll('#category-form input[type="checkbox"]');

    checkboxes.forEach(checkbox => {
    const label = checkbox.closest('label');
    if (!availableKeys.includes(checkbox.value)) {
        label.style.display = 'none';
        checkbox.checked = false;
    } else {
        label.style.display = 'block';
        checkbox.checked = window.chosenCategories?.includes(checkbox.value);
    }
    });
}

document.querySelectorAll('[data-nav="howto"]').forEach(btn => {
    btn.addEventListener('click', () => {
    window.location.href = 'landing.html';
});
});

function updateTables(selectedIdsArray) {
    const tablesContainer = document.getElementById("tables-container");
    const comparisonTableDiv = document.getElementById("urban-rural-comparison");
    const comparisonChartDiv = document.getElementById("urban-rural-charts");
    const chartsContainer = document.getElementById("charts-container");

    let selectedTab = document.querySelector('.view-tab.selected');
    if (!selectedTab) {
    selectedTab = document.querySelector('.view-tab[data-view="charts"]');
    selectedTab.classList.add('selected');
    }
    const selectedView = selectedTab.getAttribute('data-view');

    // Clear and hide all containers
    tablesContainer.innerHTML = "";
    chartsContainer.innerHTML = "";
    tablesContainer.style.display = "none";
    chartsContainer.style.display = "none";
    comparisonTableDiv.style.display = "none";
    comparisonChartDiv.style.display = "none";

    const excludedKeys = [
    "Urban_mixed_rural_status",
    "Census 2021 Super Data Zone Label",
    "Census 2021 Data Zone Label"
    ];

    const aggregatedData = {};
    let totalPopulation = 0;
    const dataSource =
        currentZoneType === 'dz'  ? dzData  :
        currentZoneType === 'dea' ? deaData :
                                    sdzData;

    selectedIdsArray.forEach(id => {
    const mapData = dataSource[id];
    if (!mapData) return;

    totalPopulation += mapData.population || 0;

    for (const [category, values] of Object.entries(mapData)) {
        if (excludedKeys.includes(category)) continue;
        if (typeof values !== 'object') continue;

        if (!aggregatedData[category]) {
        aggregatedData[category] = {};
        }

        for (const [label, count] of Object.entries(values)) {
        aggregatedData[category][label] = (aggregatedData[category][label] || 0) + count;
        }
    }
    });

    latestAggregatedData = aggregatedData;

    window.latestAggregatedData = aggregatedData;
    window.chosenCategories = selectedCategories;

    const totalPopElem = document.getElementById("totalPopulation");
    if (totalPopElem) {
    totalPopElem.textContent = totalPopulation;
    }
    renderZoneBreakdownTable(selectedIdsArray);

    const availableKeys = Object.keys(aggregatedData);
    const validCategories = selectedCategories.filter(cat => availableKeys.includes(cat));

    updateCategorySelector(availableKeys);
    updateSourceLink();

    if (selectedView === 'charts') {
    chartsContainer.style.display = "flex";
    renderAggregatedCharts(aggregatedData, validCategories);
    } else if (selectedView === 'tables') {
    tablesContainer.style.display = "block";
    renderAggregatedTables(aggregatedData, validCategories);
    } else if (selectedView === 'tableComparison') {
    comparisonTableDiv.style.display = "block";
    renderUrbanRuralComparison(selectedIdsArray);
    } else if (selectedView === 'chartComparison') {
    comparisonChartDiv.style.display = "block";
    renderUrbanRuralCharts(selectedIdsArray);
    }

}

function renderZoneBreakdownTable(selectedIdsArray) {
  const container = document.getElementById("breakdown-container");
  const titleEl = document.getElementById("areaProfileTitle");
  const summaryList = document.getElementById("summaryList");
  const populationEl = document.getElementById("totalPopulation");

  if (!selectedIdsArray.length) {
    container.style.display = "none";
    summaryList.innerHTML = "";
    populationEl.textContent = "0";
    window.areaProfileTitle = undefined;
    window.lastSelectionHash = undefined;
    return;
  }

  container.style.display = "block";
  summaryList.innerHTML = "";

  const dataSource =
    currentZoneType === 'dz' ? dzData :
    currentZoneType === 'dea' ? deaData :
    sdzData;

  const labelKey =
    currentZoneType === 'dz' ? "Census 2021 Data Zone Label" :
    currentZoneType === 'dea' ? "District Electoral Area 2014 Label" :
    "Census 2021 Super Data Zone Label";

  const lgdStats = {};
  const lgdTotals = {};
  let totalPopulation = 0;
  const currentSelectionHash = selectedIdsArray.slice().sort().join(",");
  const isSameSelection = currentSelectionHash === window.lastSelectionHash;
  window.lastSelectionHash = currentSelectionHash;

  if (!isSameSelection) {
    window.areaProfileTitle = undefined;
  }

  for (const id in dataSource) {
    const mapData = dataSource[id];
    const lgd = mapData?.["LGD"];
    if (lgd) {
      lgdTotals[lgd] = (lgdTotals[lgd] || 0) + 1;
    }
  }

  selectedIdsArray.forEach(id => {
    const mapData = dataSource[id];
    if (!mapData) return;
    const lgd = mapData["LGD"];
    const status = mapData["Urban_mixed_rural_status"];
    const labelObj = mapData[labelKey];
    const zoneName = labelObj ? Object.keys(labelObj)[0] : null;
    const population = zoneName ? labelObj[zoneName] : 0;

    if (!window.zoneNames) window.zoneNames = [];
    if (zoneName) window.zoneNames.push(zoneName);

    if (!lgdStats[lgd]) {
      lgdStats[lgd] = { total: 0, Urban: 0, Rural: 0, Mixed: 0 };
    }

    lgdStats[lgd].total++;
    if (status && lgdStats[lgd][status] !== undefined) {
      lgdStats[lgd][status]++;
    }

    totalPopulation += typeof population === "number" ? population : 0;

    if (!window.selectedZoneDetails) window.selectedZoneDetails = {};
    window.selectedZoneDetails[id] = mapData;
  });

  const savedTitle = window.areaProfileTitle || "Click to edit title for area profile";
  titleEl.textContent = savedTitle;
  titleEl.addEventListener("blur", () => {
    window.areaProfileTitle = titleEl.textContent.trim();
  });

  const sortedLGDs = Object.keys(lgdStats).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  let totalZonesSelected = 0;
  sortedLGDs.forEach(lgd => {
    const stats = lgdStats[lgd];
    const totalInLGD = lgdTotals[lgd] || stats.total;
    const parts = [];
    if (stats.Urban) parts.push(`${stats.Urban} urban`);
    if (stats.Rural) parts.push(`${stats.Rural} rural`);
    if (stats.Mixed) parts.push(`${stats.Mixed} mixed`);
    totalZonesSelected += stats.total;

    const li = document.createElement("li");
    li.innerHTML = `${lgd}: <strong>${stats.total} of ${totalInLGD}</strong> zones selected (${parts.join(", ")})`;
    summaryList.appendChild(li);
  });

  window.totalZonesSelected = totalZonesSelected;
  populationEl.textContent = totalPopulation.toLocaleString();

  updateSummaryPreview();
  ensureSummaryHero();
}

// document.getElementById("urban-rural-btn").addEventListener("click", () => {
//   const tablesContainer = document.getElementById("tables-container");
//   const comparisonDiv = document.getElementById("urban-rural-comparison");
//   const toggleBtn = document.getElementById("urban-rural-btn");

//   const isComparisonVisible = comparisonDiv.style.display === "block";

//   if (isComparisonVisible) {
//     // Show main tables again
//     comparisonDiv.style.display = "none";
//     tablesContainer.style.display = "block";
//     renderAggregatedTables(latestAggregatedData, selectedCategories);

//     // Update button label
//     toggleBtn.textContent = "Urban Rural Comparison";
//   } else {
//     // Show comparison view
//     tablesContainer.style.display = "none";
//     comparisonDiv.style.display = "block";
//     renderUrbanRuralComparison(Array.from(selectedIds));

//     // Update button label
//     toggleBtn.textContent = "View All Data";
//   }
// });

const tabButtons = document.querySelectorAll('.view-tab');
const views = {
    charts: document.getElementById("charts-container"),
    tables: document.getElementById("tables-container"),
    tableComparison: document.getElementById("urban-rural-comparison"),
    chartComparison: document.getElementById("urban-rural-charts")
};

tabButtons.forEach(tab => {
    tab.addEventListener("click", () => {
    
    tabButtons.forEach(btn => btn.classList.remove("selected"));
    tab.classList.add("selected");

    Object.values(views).forEach(div => div.style.display = "none");
    const view = tab.getAttribute("data-view");
    currentView = view;

    // Prevent rendering if no categories selected
    if (!selectedCategories || selectedCategories.length === 0) {
        return; // nothing to render
    }

    // Render valid view
    if (view === "charts") {
        renderAggregatedCharts(latestAggregatedData, selectedCategories);
        views[view].style.display = "flex";
    } else if (view === "tables") {
        renderAggregatedTables(latestAggregatedData, selectedCategories);
        views[view].style.display = "block";
    } else if (view === "tableComparison") {
        renderUrbanRuralComparison(Array.from(selectedIds));
        views[view].style.display = "block";
    } else if (view === "chartComparison") {
        renderUrbanRuralCharts(Array.from(selectedIds));
        views[view].style.display = "block";
    }
    });
});

function renderAggregatedTables(aggregatedData, selectedCategories = []) {
    if (!aggregatedData || Object.keys(aggregatedData).length === 0) return;

    if (selectedCategories.length > 0) {
    const valid = selectedCategories.filter(cat => Object.keys(aggregatedData).includes(cat));
    if (valid.length === 0) return;
    }

    const container = document.getElementById("tables-container");
    container.innerHTML = "";

    const grid = document.createElement("div");
    grid.className = "table-grid"; // use consistent class
    grid.style.display = "flex";
    grid.style.flexWrap = "wrap";
    grid.style.gap = "2rem";

    const entries = Object.entries(aggregatedData).filter(([key]) =>
    selectedCategories.length === 0 || selectedCategories.includes(key)
    );

  entries.forEach(([category, values]) => {
    const wrapper = document.createElement("div");
    wrapper.className = "table-wrapper";
    wrapper.style.flex = "0 0 calc(50% - 1rem)";
    wrapper.style.background = "#fff";
    wrapper.style.boxShadow = "0 2px 4px rgba(0,0,0,0.1)";
    wrapper.style.padding = "16px";
    wrapper.style.borderRadius = "8px";
    wrapper.style.boxSizing = "border-box";

    const title = document.createElement("h3");
    title.textContent = category;

    const year = Year_Data?.Year?.[category];

    if (year) {
        const yearEl = document.createElement("div");
        yearEl.textContent = `Year: ${year}`;
        yearEl.style.marginTop = "4px";
        yearEl.style.fontSize = "0.9rem";
        yearEl.style.color = "#555";
        title.appendChild(document.createElement("br")); // optional line break
        title.appendChild(yearEl);
    }

    wrapper.appendChild(title);

    const src = (window.labelToSource || {})[category];
    if (src) {
    const badge = document.createElement('span');
    const isPortal = src === 'Data Portal';
    badge.className = 'source-badge ' + (isPortal ? 'badge-portal' : 'badge-ftb');
    badge.textContent = isPortal ? 'Data Portal' : 'Flexible Table Builder';
    badge.style.marginLeft = '8px';
    badge.title = src;
    title.appendChild(badge);
    }
    wrapper.appendChild(title);

    const table = document.createElement("table");
    table.className = "data-table";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    ["Category", "Count", "Percentage", "NI %"].forEach(text => {
        const th = document.createElement("th");
        th.textContent = text;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const totalCount = Object.values(values).reduce((acc, val) => acc + val, 0);

    for (const [label, count] of Object.entries(values)) {
        const row = document.createElement("tr");

        const tdLabel = document.createElement("td");
        tdLabel.textContent = label;
        row.appendChild(tdLabel);

        const tdCount = document.createElement("td");
        tdCount.textContent = count;
        row.appendChild(tdCount);

        const tdPercentage = document.createElement("td");
        tdPercentage.textContent = totalCount > 0 ? ((count / totalCount) * 100).toFixed(2) + "%" : "0%";
        row.appendChild(tdPercentage);

        const tdNI = document.createElement("td");
        const niVal = niTotals[category]?.[label];
        tdNI.textContent = typeof niVal === "number" ? niVal.toFixed(1) + "%" : "–";
        row.appendChild(tdNI);

        tbody.appendChild(row);
    }

    table.appendChild(tbody);
    wrapper.appendChild(table);
    grid.appendChild(wrapper);
    });

    container.appendChild(grid);

    window.latestAggregatedData = aggregatedData;
}

//Chart.register(ChartDataLabels);

let labelToCode = {};
let labelToSource = {};

async function loadLookupData() {
    const response = await fetch('category_lookup.json');
    const lookup = await response.json();

    lookup.forEach(item => {
        labelToCode[item.nested_list_names] = item.further_breakdown_df;
        labelToSource[item.nested_list_names] = item.Source;
    });

    window.labelToCode = labelToCode;
    window.labelToSource = labelToSource;
}

loadLookupData().then(() => {
    renderAggregatedCharts(data, selectedCategories);
});

function getCategoryURL(label, zoneType = 'sdz') {

  const source = window.labelToSource?.[label];
  const categoryCode = window.labelToCode?.[label];
 
  if (source === "Flexible Table Builder" && categoryCode) {
    return `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}21&v=${categoryCode}`;
  } else if (source === "Data Portal") {
    if (label === "Benefits Statistics" && zoneType === "dz") {
      return "https://data.nisra.gov.uk/table/BSDZ";
    } else if (label === "Benefits Statistics" && zoneType === "sdz") {
      return "https://data.nisra.gov.uk/table/BSSDZ";
    } else if ((label === "Age (MYE)" || label === "Sex (MYE)") && zoneType === "sdz") {
      return "https://data.nisra.gov.uk/table/MYE01T012";
    }
  }
  return null;
}

function renderAggregatedCharts(data, selectedCategories = []) {
    // wait until element has a real width 
    function whenVisible(el, cb) {
        if (el.offsetParent !== null && el.clientWidth > 0) return cb();
        const ro = new ResizeObserver(() => {
            if (el.clientWidth > 0) {
            ro.disconnect();
            cb();
            }
        });
        ro.observe(el);
    }

    if (!data || Object.keys(data).length === 0) return;

    const EXCLUDED_KEYS = [
    "Census 2021 Data Zone Label",
    "Census 2021 Super Data Zone Label"
    ];

    const allCats = selectedCategories.length
    ? selectedCategories
    : Object.keys(data).filter(k => !EXCLUDED_KEYS.includes(k)).sort();

    const categories = allCats.filter(k => !EXCLUDED_KEYS.includes(k));
    if (categories.length === 0) return;

    const container = document.getElementById("charts-container");
    container.innerHTML = "";

    // destroy existing charts
    window.chartInstances?.forEach(c => c.destroy());
    window.chartInstances = [];

    // ---- constants  ----
    const FONT               = "14px sans-serif";
    const LINE_HEIGHT        = 16;
    const BAR_HEIGHT         = 32;
    const BAR_SPACING        = 4;
    const LABEL_BLOCK_HEIGHT = LINE_HEIGHT * 2 + 4;
    const CHART_TOP_PADDING  = 15;
    const LABEL_TO_BAR_GAP = 1;

    // 2-column grid
    const grid = document.createElement("div");
    Object.assign(grid.style, {
    display: "grid",
    gridTemplateColumns: "repeat(2, 1fr)",
    gap: "2rem",
    width: "100%"
    });
    container.appendChild(grid);

categories.forEach(category => {
    const values = data[category];
    if (!values) return;

const wrapper = document.createElement("div");
Object.assign(wrapper.style, {
    position: "relative",
    display: "flex",
    flexDirection: "column",
    background: "#fff",
    boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
    padding: "16px",
    borderRadius: "8px",
    boxSizing: "border-box",
    width: "100%"
});

// Create a container for title and info button
const titleWrapper = document.createElement("div");
titleWrapper.style.display = "flex";
titleWrapper.style.alignItems = "center";
titleWrapper.style.justifyContent = "space-between";

// Create the title element
const title = document.createElement("h3");
title.textContent = category;
title.style.margin = "0";

// Create a wrapper for the info button and tooltip
const infoWrapper = document.createElement("div");
infoWrapper.style.position = "relative";
infoWrapper.style.display = "inline-block";

// Create the info button (SVG)
const infoButton = document.createElement("img");
infoButton.src = "img/i-button.svg";
infoButton.alt = "Information";
infoButton.title = "More information";
infoButton.style.width = "20px";
infoButton.style.height = "20px";
infoButton.style.cursor = "pointer";

// Create the tooltip
const tooltip = document.createElement("div");

const zoneType = window.selectedZoneType || 'sdz';
const url = getCategoryURL(category, zoneType);

if (url) {
  tooltip.innerHTML = `<strong>Source:</strong> <a href="${url}" target="_blank" style="
    color: #fff;
    text-decoration: underline;
    font-size: 0.8rem;
    white-space: nowrap;
  ">${url}</a>`;
} else {
  tooltip.textContent = `More info about ${category}`;
}

tooltip.style.position = "absolute";
tooltip.style.bottom = "125%";
tooltip.style.left = "50%";
tooltip.style.transform = "translateX(-50%)";
tooltip.style.backgroundColor = "#333";
tooltip.style.color = "#fff";
tooltip.style.padding = "6px 8px";
tooltip.style.borderRadius = "4px";
tooltip.style.fontSize = "0.8rem";
tooltip.style.whiteSpace = "nowrap";
tooltip.style.visibility = "hidden";
tooltip.style.opacity = "0";
tooltip.style.transition = "opacity 0.3s";
tooltip.style.pointerEvents = "auto"; // ensure it can receive hover
tooltip.style.zIndex = "1000"; // make sure it's above other elements

let tooltipTimeout;

infoWrapper.addEventListener("mouseenter", () => {
  clearTimeout(tooltipTimeout);
  tooltip.style.visibility = "visible";
  tooltip.style.opacity = "1";
});

infoWrapper.addEventListener("mouseleave", () => {
  tooltipTimeout = setTimeout(() => {
    tooltip.style.visibility = "hidden";
    tooltip.style.opacity = "0";
  }, 200); // slight delay to allow hover into tooltip
});

tooltip.addEventListener("mouseenter", () => {
  clearTimeout(tooltipTimeout);
  tooltip.style.visibility = "visible";
  tooltip.style.opacity = "1";
});

tooltip.addEventListener("mouseleave", () => {
  tooltipTimeout = setTimeout(() => {
    tooltip.style.visibility = "hidden";
    tooltip.style.opacity = "0";
  }, 200);
});

// Assemble the info button and tooltip
infoWrapper.appendChild(infoButton);
infoWrapper.appendChild(tooltip);

// Assemble the title and info button
const titleContent = document.createElement("div");
titleContent.style.display = "flex";
titleContent.style.alignItems = "center";
titleContent.style.gap = "8px";
titleContent.appendChild(title);
titleContent.appendChild(infoWrapper);

// Add to wrapper
titleWrapper.appendChild(titleContent);
wrapper.appendChild(titleWrapper);

// Add year info if available
const year = Year_Data?.Year?.[category];
if (year) {
    const yearEl = document.createElement("div");
    yearEl.textContent = `Year: ${year}`;
    yearEl.style.marginTop = "4px";
    yearEl.style.fontSize = "0.9rem";
    yearEl.style.color = "#555";
    title.appendChild(document.createElement("br"));
    title.appendChild(yearEl);
}

// Add source badge if available
const src = (window.labelToSource || {})[category];
if (src) {
    const badge = document.createElement('span');
    const isPortal = src === 'Data Portal';
    badge.className = 'source-badge ' + (isPortal ? 'badge-portal' : 'badge-ftb');
    badge.textContent = isPortal ? 'Data Portal' : 'Flexible Table Builder';
    badge.style.marginLeft = '8px';
    badge.title = src;
    title.appendChild(badge);
}

// Append title wrapper to main wrapper
wrapper.appendChild(titleWrapper);

    const labels = Object.keys(values);
    const total  = Object.values(values).reduce((a, v) => a + v, 0);
    const barPercents = labels.map(l => total > 0 ? +(((values[l] || 0) / total * 100).toFixed(2)) : 0);

    const datasets = [
        { label: "Percentage", data: barPercents, backgroundColor: "#3878c5", barThickness: BAR_HEIGHT },
        { label: "NI", data: [], type: "line", borderColor: "#222", borderWidth: 2, fill: false, pointRadius: 0 }
    ];

    // legend
    const legendEl = document.createElement("div");
    Object.assign(legendEl.style, {
        display: "flex",
        justifyContent: "center",
        gap: "1rem",
        alignItems: "center",
        marginTop: "12px",
        marginBottom: "8px"
    });
    datasets.forEach(ds => {
        const item = document.createElement("div");
        Object.assign(item.style, { display: "flex", alignItems: "center" });
        const swatch = document.createElement("span");
        Object.assign(swatch.style, {
        display: "inline-block",
        width: ds.type === "line" ? "4px" : "12px",
        height: "15px",
        marginRight: "6px",
        backgroundColor: ds.type === "line" ? ds.borderColor : ds.backgroundColor,
        borderRadius: "0"
        });
        const text = document.createElement("span");
        text.textContent = ds.label;
        item.appendChild(swatch);
        item.appendChild(text);
        legendEl.appendChild(item);
    });
    wrapper.appendChild(legendEl);

    // height
    const barsPerGroup = datasets.length - 1; 
    const GAP_BELOW_GROUP_2 = 48; 
    const GAP_BELOW_GROUP_3 = 36;  
    const GAP_BELOW_GROUP_N = 26; 

    const tailGap =
        (labels.length <= 2) ? GAP_BELOW_GROUP_2 :
        (labels.length === 3) ? GAP_BELOW_GROUP_3 :
        GAP_BELOW_GROUP_N;

    const GROUP_SPACING =
        LABEL_BLOCK_HEIGHT +
        barsPerGroup * (BAR_HEIGHT + BAR_SPACING) +
        tailGap;
    const canvasHeight = labels.length * GROUP_SPACING + CHART_TOP_PADDING;

    // canvas (set CSS width 100% so it fills card; height from pixels)
    const canvas = document.createElement("canvas");
    Object.assign(canvas.style, { display: "block", width: "100%", maxHeight: "none" });
    canvas.height = canvasHeight;
    wrapper.appendChild(canvas);

    const spacer = document.createElement("div");
    spacer.style.flex = "1";
    wrapper.appendChild(spacer);

    grid.appendChild(wrapper);

    // render once the wrapper has a real width; set bitmap & CSS width once (no loops)
    whenVisible(wrapper, () => {
        const drawWidth = Math.max(0, wrapper.clientWidth - 32); // 16px padding each side
        canvas.width = drawWidth;                                 // bitmap width
        canvas.style.width = `${drawWidth}px`;                    // CSS width to match

        // NI values
        const niMap =
        (typeof window !== "undefined" && window.niTotals && window.niTotals[category]) ||
        (typeof niTotals !== "undefined" && niTotals[category]) || {};
        const niValues = labels.map(l => (typeof niMap[l] === "number" ? niMap[l] : null));

        const rawMax = Math.max(
        ...datasets[0].data,
        ...niValues.filter(v => typeof v === "number")
        ) * 1.05;
        const cappedMax = Number.isFinite(rawMax) ? Math.min(rawMax, 100) : 100;

        const chart = new Chart(canvas, {
        type: "bar",
        data: { labels, datasets },
        options: {
            indexAxis: "y",
            responsive: false,
            maintainAspectRatio: false,
            animation: false, // keep overlays stable; prevents repeated reflows
            layout: { padding: { top: CHART_TOP_PADDING, left: 10, right: 10, bottom: 0 } },
            plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.raw}%` } }
            },
            scales: {
            x: {
                beginAtZero: true,
                suggestedMax: cappedMax,
                title: { display: true, text: "Percentage" },
                ticks: { callback: (v) => `${v}%` }
            },
            y: { ticks: { display: false }, grid: { display: false }, offset: true }
            }
        },
        plugins: [
            {
            id: "aboveGroupLabels",
            afterDatasetsDraw(chartInst) {
                const ctx  = chartInst.ctx;
                const card = chartInst.canvas.parentNode;
                card.querySelectorAll(".label-overlay").forEach(el => el.remove());

                ctx.save();
                ctx.font = FONT;
                ctx.fillStyle = "#000";
                ctx.textAlign = "left";
                ctx.textBaseline = "top";

                const xStart = chartInst.chartArea.left + 4;
                const meta   = chartInst.getDatasetMeta(0);

                meta.data.forEach((bar, i) => {
                const topY = bar.y - bar.height / 2;  
                const blockBottom = topY - LABEL_TO_BAR_GAP;  
                const labelY = blockBottom - LABEL_BLOCK_HEIGHT;

                const label      = chartInst.data.labels[i];
                const textWidth  = ctx.measureText(label).width;
                const maxAllowed = chartInst.chartArea.right - xStart - 30;

                if (textWidth < maxAllowed) {
                    ctx.fillText(label, xStart, labelY);
                } else {
                    const shortLabel = label.slice(0, 25) + "...";
                    ctx.fillText(shortLabel, xStart, labelY);

                    const infoBtn = document.createElement("div");
                    infoBtn.className = "label-overlay";
                    infoBtn.textContent = "ⓘ";
                    infoBtn.style.position = "absolute";

                    const chartRect = chartInst.canvas.getBoundingClientRect();
                    const contRect  = card.getBoundingClientRect();
                    const labelW    = ctx.measureText(shortLabel).width;
                    const topOff    = labelY + chartRect.top - contRect.top + LINE_HEIGHT / 2 - 15;
                    const leftOff   = chartInst.canvas.offsetLeft + xStart + labelW + 6;

                    infoBtn.style.top    = `${topOff}px`;
                    infoBtn.style.left   = `${leftOff}px`;
                    infoBtn.style.cursor = "pointer";
                    infoBtn.style.color  = "#0074D9";
                    infoBtn.title        = label;
                    infoBtn.addEventListener("click", () => alert(`Full label:\n${label}`));
                    card.appendChild(infoBtn);
                }

                const value = datasets[0].data[i];
                const niVal = niValues[i];
                const niText = typeof niVal === "number" ? ` | NI: ${niVal.toFixed(1)}%` : "";
                ctx.fillText(`Percentage: ${value}%${niText}`, xStart, labelY + LINE_HEIGHT);
                });

                ctx.restore();
            }
            },
            {
            id: "drawNILines",
            afterDatasetsDraw(chartInst) {
                const ctx = chartInst.ctx;
                const xScale = chartInst.scales.x;

                ctx.save();
                ctx.strokeStyle = "#222";
                ctx.lineWidth = 4;
                ctx.setLineDash([]);

                const bars = chartInst.getDatasetMeta(0).data;
                bars.forEach((bar, i) => {
                const val = niValues[i];
                if (typeof val !== "number") return;

                const x = xScale.getPixelForValue(val);
                const yTop    = bar.y - bar.height / 2;
                const yBottom = bar.y + bar.height / 2;

                ctx.beginPath();
                ctx.moveTo(x, yTop);
                ctx.lineTo(x, yBottom);
                ctx.stroke();
                });

                ctx.restore();
            }
            }
        ]
        });

        window.chartInstances.push(chart);
    });
    });
}

function clearSelections() {
    selectedIds.forEach(id => {
    if (sdzData[id]) {
        map.setFeatureState(
        { source: 'sdz2021', sourceLayer: 'SDZ2021_clipped', id },
        { hovered: false }
        );
    }
    if (dzData[id]) {
        map.setFeatureState(
        { source: 'dz2021', sourceLayer: 'DZ2021_clipped', id },
        { hovered: false }
        );
    }
    if (deaData[id]) {
      map.setFeatureState(
        { source: 'dea2014', sourceLayer: 'DEA2014_clipped', id },
        { hovered: false }
      );
    }
    });

    selectedIds.clear();
    selectedLGDs.clear();
    popup.remove();

    // Reset UI
    document.getElementById("tables-container").innerHTML = "";
    document.getElementById("breakdown-container").innerHTML = "";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";
    document.getElementById("tables-container").style.display = "none";

    document.querySelectorAll('#lgd-buttons input[type="checkbox"]').forEach(cb => {
    cb.checked = false;
    });

    document.querySelectorAll('.lgd-btn').forEach(label => {
    label.classList.remove('selected');
    });

    const totalPop = document.getElementById("totalPopulation");
    if (totalPop) totalPop.textContent = "0";

    const center = map.getCenter();
    const zoom = map.getZoom();
    map.jumpTo({ center, zoom: zoom + 0.00001 });
    
    // Destroy all Chart.js instances
    if (window.chartInstances && window.chartInstances.length > 0) {
    window.chartInstances.forEach(chart => {
        if (chart && typeof chart.destroy === 'function') {
        chart.destroy();
        }
    });
    window.chartInstances = []  ;
    }
    
    // Optionally clear chart container
    const container = document.getElementById("charts-container");
    if (container) {
    container.innerHTML = "";
    }
    const out  = document.getElementById("output-content");
    const mapc = document.getElementById("map-content");
    if (out && mapc) {
    out.classList.add("hidden-section");
    mapc.classList.remove("hidden-section");
    }
    if (typeof updateCtaEnabled === 'function') updateCtaEnabled();
}

document.getElementById("clear-selection-btn").addEventListener("click", function (e) {
    e.preventDefault(); 
    this.blur(); 
    clearSelections(); 

    map.getSource('draw-geom').setData({ type:'FeatureCollection', features: [] });
    lastDrawnFeature = null;
});

document.querySelectorAll('#lgd-buttons input[type="checkbox"]').forEach(checkbox => {
  checkbox.addEventListener('change', () => {
    const lgdCode = checkbox.value;
    const label = document.querySelector(`label[for="${checkbox.id}"]`);

    const zone = currentZoneType; 

    const dataSource =
      zone === 'dz'  ? dzData  :
      zone === 'dea' ? deaData :
                       sdzData;

    const source =
      zone === 'dz'  ? 'dz2021'  :
      zone === 'dea' ? 'dea2014' :
                       'sdz2021';

    const sourceLayer =
      zone === 'dz'  ? 'DZ2021_clipped'  :
      zone === 'dea' ? 'DEA2014_clipped' :
                       'SDZ2021_clipped';

    if (checkbox.checked) {
      selectedLGDs.add(lgdCode);
      label?.classList.add('selected');

      const newIds = Object.entries(dataSource)
        .filter(([, rec]) => rec?.LGD === lgdCode)
        .map(([id]) => id);

      newIds.forEach(id => {
        if (!selectedIds.has(id)) {
          selectedIds.add(id);
          map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
        }
      });

    } else {
      selectedLGDs.delete(lgdCode);
      label?.classList.remove('selected');

      const toRemove = Array.from(selectedIds).filter(id => dataSource[id]?.LGD === lgdCode);
      toRemove.forEach(id => {
        selectedIds.delete(id);
        map.setFeatureState({ source, sourceLayer, id }, { hovered: false });
      });
    }

    const selectedArray = Array.from(selectedIds);
    window.selectedIdsExcel = selectedIds;
    updateTables(selectedArray);
    renderZoneBreakdownTable(selectedArray);
    updateCtaEnabled();
    
  });
});

function aggregateUrbanRuralData(selectedIdsArray, selectedCategories) {
    const excludedKeys = ["Urban_mixed_rural_status", "Census 2021 Super Data Zone Label", "LGD"];
    const groups = { Urban: {}, Rural: {}, Mixed: {} };

    selectedIdsArray.forEach(id => {
    const dataSource = currentZoneType === 'dz' ? dzData : sdzData;
    const data = dataSource[id];
    if (!data || !["Urban", "Rural", "Mixed"].includes(data.Urban_mixed_rural_status)) return;

    const group = data.Urban_mixed_rural_status;

    for (const [category, values] of Object.entries(data)) {
        if (excludedKeys.includes(category) || typeof values !== 'object') continue;

        if (selectedCategories.length > 0 && !selectedCategories.includes(category)) continue;

        if (!groups[group][category]) {
        groups[group][category] = {};
        }

        for (const [label, count] of Object.entries(values)) {
        groups[group][category][label] = (groups[group][category][label] || 0) + count;
        }
    }
    });

    return groups;
}

window.renderUrbanRuralComparison = function (selectedIdsArray) {
    const container = document.getElementById("urban-rural-comparison");
    container.innerHTML = "";

    const groups = aggregateUrbanRuralData(selectedIdsArray, selectedCategories);
    window.urbanRuralComparisonData = groups;

    const categoriesToDisplay = selectedCategories.length
    ? selectedCategories
    : Array.from(new Set([
        ...Object.keys(groups.Urban || {}),
        ...Object.keys(groups.Rural || {}),
        ...Object.keys(groups.Mixed || {})
        ]));

    const rowWrapper = document.createElement("div");
    rowWrapper.style.display = "flex";
    rowWrapper.style.flexWrap = "wrap";
    rowWrapper.style.gap = "20px";

    categoriesToDisplay.forEach((category) => {
    const hasUrban = !!groups.Urban?.[category];
    const hasRural = !!groups.Rural?.[category];
    const hasMixed = !!groups.Mixed?.[category];
    if (!hasUrban && !hasRural && !hasMixed) return;

    const allLabels = new Set([
        ...Object.keys(groups.Urban?.[category] || {}),
        ...Object.keys(groups.Rural?.[category] || {}),
        ...Object.keys(groups.Mixed?.[category] || {})
    ]);
    const labels = Array.from(allLabels).sort();

    const columns = [
        { key: "label", title: "Category", width: "30%" }
    ];
    if (hasUrban) {
        columns.push({ key: "urbanCount", title: "Urban\nCount" });
        columns.push({ key: "urbanPct", title: "Urban %" });
    }
    if (hasRural) {
        columns.push({ key: "ruralCount", title: "Rural\nCount" });
        columns.push({ key: "ruralPct", title: "Rural %" });
    }
    if (hasMixed) {
        columns.push({ key: "mixedCount", title: "Mixed\nCount" });
        columns.push({ key: "mixedPct", title: "Mixed %" });
    }
    columns.push({ key: "niPct", title: "NI %" });

    const wrapper = document.createElement("div");
    wrapper.style.flex = "0 0 100%";
    wrapper.style.background = "#fff";
    wrapper.style.padding = "16px";
    wrapper.style.borderRadius = "8px";
    wrapper.style.boxShadow = "0 1px 3px rgba(0,0,0,0.1)";
    wrapper.style.boxSizing = "border-box";
    wrapper.style.alignSelf = "flex-start"; // Ensures independent height

    const title = document.createElement("h3");
    title.textContent = `${category} – Urban/Rural Comparison`;
    title.style.marginBottom = "12px";
    wrapper.appendChild(title);

    const table = document.createElement("table");
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.style.tableLayout = "fixed";

    // Table header
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    columns.forEach((col, i) => {
        const th = document.createElement("th");
        th.textContent = col.title;
        th.style.padding = "8px";
        th.style.border = "1px solid #ccc";
        th.style.backgroundColor = "#3878c5";
        th.style.color = "#fff";
        th.style.fontWeight = "bold";
        th.style.fontSize = "14px";
        th.style.textAlign = "left";
        th.style.verticalAlign = "top";
        th.style.whiteSpace = "normal";
        th.style.wordBreak = "break-word";
        th.style.overflowWrap = "break-word";

        if (i === 0 && col.width) th.style.width = col.width;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Table body
    const tbody = document.createElement("tbody");

    labels.forEach((label) => {
        const row = document.createElement("tr");

        const urbanCount = groups.Urban?.[category]?.[label] || 0;
        const ruralCount = groups.Rural?.[category]?.[label] || 0;
        const mixedCount = groups.Mixed?.[category]?.[label] || 0;

        const urbanTotal = groups.Urban?.[category]
        ? Object.values(groups.Urban[category]).reduce((a, b) => a + b, 0)
        : 0;
        const ruralTotal = groups.Rural?.[category]
        ? Object.values(groups.Rural[category]).reduce((a, b) => a + b, 0)
        : 0;
        const mixedTotal = groups.Mixed?.[category]
        ? Object.values(groups.Mixed[category]).reduce((a, b) => a + b, 0)
        : 0;

        const data = {
        label,
        urbanCount,
        ruralCount,
        mixedCount,
        urbanPct: urbanTotal ? ((urbanCount / urbanTotal) * 100).toFixed(2) + "%" : "–",
        ruralPct: ruralTotal ? ((ruralCount / ruralTotal) * 100).toFixed(2) + "%" : "–",
        mixedPct: mixedTotal ? ((mixedCount / mixedTotal) * 100).toFixed(2) + "%" : "–",
        niPct: typeof window.niTotals?.[category]?.[label] === "number"
            ? window.niTotals[category][label].toFixed(1) + "%"
            : "–"
        };

        columns.forEach((col, i) => {
        const td = document.createElement("td");
        td.textContent = data[col.key] || "–";
        td.style.padding = "8px";
        td.style.border = "1px solid #ccc";
        td.style.fontSize = "14px";
        td.style.textAlign = "left";
        td.style.verticalAlign = "top";
        td.style.whiteSpace = "normal";
        td.style.wordBreak = "break-word";
        td.style.overflowWrap = "break-word";

        if (i === 0 && col.width) td.style.width = col.width;
        row.appendChild(td);
        });

        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);
    rowWrapper.appendChild(wrapper);
    });

    container.appendChild(rowWrapper);
};

// Render grouped (Urban / Rural / Mixed) horizontal bar charts for selected categories
function renderUrbanRuralCharts(selectedIdsArray) {
    // Utility: run drawing code only when the element is visible and has a real width
    function whenVisible(el, cb) {
    if (el.offsetParent !== null && el.clientWidth > 0) return cb();
    const ro = new ResizeObserver(() => {
        if (el.clientWidth > 0) {
        ro.disconnect();
        cb();
        }
    });
    ro.observe(el);
    }

    // Build grouped data structures for Urban/Rural/Mixed using current selection + categories
    const groups = aggregateUrbanRuralData(selectedIdsArray, selectedCategories);

    // Reset container and any existing charts
    const container = document.getElementById("urban-rural-charts");
    if (!container) return;
    container.innerHTML = "";

    // Destroy existing charts
    if (Array.isArray(window.chartInstances)) {
    window.chartInstances.forEach(c => { try { c.destroy(); } catch(_){} });
    }
    window.chartInstances = [];

    // Decide which categories to plot (selected, or all categories with any data)
    const categories = selectedCategories && selectedCategories.length
    ? selectedCategories
    : Array.from(new Set([
        ...Object.keys(groups.Urban || {}),
        ...Object.keys(groups.Rural || {}),
        ...Object.keys(groups.Mixed || {})
        ]));

    // Layout constants
    const FONT = "bold 12px sans-serif";
    const LINE_HEIGHT = 16;
    const BAR_HEIGHT = 16;
    const BAR_SPACING = 2;
    const LABEL_BLOCK_HEIGHT = LINE_HEIGHT * 2 + 4;
    const CHART_TOP_PADDING = 15;

    // Two-column grid
    const grid = document.createElement("div");
    Object.assign(grid.style, {
    display: "grid",
    gridTemplateColumns: "repeat(2, 1fr)",
    gap: "2rem",
    width: "100%"
    });
    container.appendChild(grid);

    // For each category: build a card, compute datasets, and render a Chart.js chart
    categories.forEach(category => {
    const hasData = groups.Urban?.[category] || groups.Rural?.[category] || groups.Mixed?.[category];
    if (!hasData) return;

    // Card wrapper + title
    const wrapper = document.createElement("div");
    Object.assign(wrapper.style, {
        position: "relative",
        display: "flex",
        flexDirection: "column",
        background: "#fff",
        boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
        padding: "16px",
        borderRadius: "8px",
        boxSizing: "border-box",
        width: "100%"
    });

    // Title
    const title = document.createElement("h3");
    title.textContent = `${category} – Urban/Rural/Mixed`;
    wrapper.appendChild(title);

    // Assemble union of labels across U/R/M for this category
    const labelSet = new Set([
        ...Object.keys(groups.Urban?.[category] || {}),
        ...Object.keys(groups.Rural?.[category] || {}),
        ...Object.keys(groups.Mixed?.[category] || {})
    ]);
    const labels = Array.from(labelSet);

    // Build bar datasets
    const colorMap = { Urban: "#2c7fb8", Rural: "#7fcdbb", Mixed: "#edf8b1" };
    const barDatasets = ["Urban","Rural","Mixed"]
        .filter(g => groups[g]?.[category])
        .map(g => {
        const vals  = groups[g][category];
        const total = Object.values(vals).reduce((a,b)=>a+b,0);
        return {
            label: g,
            data:  labels.map(l => total>0 ? +(((vals[l]||0)/total*100).toFixed(2)) : 0),
            backgroundColor: colorMap[g],
            borderColor: "#000",
            borderWidth: 1.5,
            barThickness: BAR_HEIGHT
        };
        });

    // Add NI “dataset” placeholder so legend shows an NI swatch
    barDatasets.push({
        label: "NI",
        data: [],
        type: "line",
        borderColor: "#222",
        borderWidth: 2,
        fill: false,
        pointRadius: 0
    });

    // Legend
    const legendEl = document.createElement("div");
    Object.assign(legendEl.style, {
        display: "flex",
        justifyContent: "center",
        gap: "1rem",
        alignItems: "center",
        marginTop: "12px",
        marginBottom: "8px"
    });
    barDatasets.forEach(ds => {
        const item = document.createElement("div");
        Object.assign(item.style, { display: "flex", alignItems: "center" });
        const swatch = document.createElement("span");
        Object.assign(swatch.style, {
        display: "inline-block",
        width: ds.type === "line" ? "4px" : "12px",
        height: "15px",
        marginRight: "4px",
        backgroundColor: ds.type==="line" ? ds.borderColor : ds.backgroundColor,
        borderRadius: "0"
        });
        const text = document.createElement("span");
        text.textContent = ds.label;
        item.appendChild(swatch);
        item.appendChild(text);
        legendEl.appendChild(item);
    });
    wrapper.appendChild(legendEl);

    // Compute canvas height dynamically based on number of label groups
    const barsPerGroup = barDatasets.length - 1;
    const GAP_BELOW_GROUP_2 = 48;
    const GAP_BELOW_GROUP_3 = 36;
    const GAP_BELOW_GROUP_N = 26;

    const tailGap =
        (labels.length <= 2) ? GAP_BELOW_GROUP_2 :
        (labels.length === 3) ? GAP_BELOW_GROUP_3 :
        GAP_BELOW_GROUP_N;

    const GROUP_SPACING =
        LABEL_BLOCK_HEIGHT +
        barsPerGroup * (BAR_HEIGHT + BAR_SPACING) +
        tailGap;

    const canvasHeight = labels.length * GROUP_SPACING + CHART_TOP_PADDING;

    // Canvas host for Chart.js
    const canvas = document.createElement("canvas");
    Object.assign(canvas.style, { display: "block", width: "100%", maxHeight: "none" });
    canvas.height = canvasHeight;
    wrapper.appendChild(canvas);

    // Spacer to allow equal row heights without changing canvas size
    const spacer = document.createElement("div");
    spacer.style.flex = "1";
    wrapper.appendChild(spacer);

    grid.appendChild(wrapper);

    // Render when visible (one-time sizing like aggregated)
    whenVisible(wrapper, () => {
        const drawWidth = Math.max(0, wrapper.clientWidth - 32); // 16px padding each side
        canvas.width = drawWidth;                                 // bitmap width
        canvas.style.width = `${drawWidth}px`;                    // CSS width to match

        // NI values
        const niValues = labels.map(l =>
        typeof window.niTotals?.[category]?.[l] === "number"
            ? window.niTotals[category][l]
            : null
        );

        // x-axis max
        const rawMax = Math.max(
        ...barDatasets.flatMap(ds=>ds.data||[]),
        ...niValues.filter(v=>typeof v==="number")
        ) * 1.05;
        const cappedMax = Math.min(isFinite(rawMax) ? rawMax : 100, 100);

        // Chart
        const dpr = Math.max(1, Math.min(3, window.devicePixelRatio || 1));
        const chart = new Chart(canvas, {
        type: "bar",
        data: { labels, datasets: barDatasets },
        options: {
            devicePixelRatio: dpr,
            indexAxis: "y",
            responsive: false,
            maintainAspectRatio: false,
            animation: false,
            layout: { padding: { top: CHART_TOP_PADDING, left:10, right:10, bottom:0 } },
            plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: ctx=>`${ctx.dataset.label}: ${ctx.raw}%` } }
            },
            scales: {
            x: {
                beginAtZero: true,
                suggestedMax: cappedMax,
                title: { display: true, text: "Percentage" },
                ticks: { callback: v=>`${v}%` }
            },
            y: { ticks:{display:false}, grid:{display:false}, offset:true }
            }
        },
        plugins: [
            {
            id: "aboveGroupLabels",
            afterDatasetsDraw(chartInst) {
                const ctx  = chartInst.ctx;
                const card = chartInst.canvas.parentNode;
                card.querySelectorAll(".label-overlay").forEach(el => el.remove());

                ctx.save();
                ctx.font = FONT;
                ctx.fillStyle = "#000";
                ctx.textAlign = "left";
                ctx.textBaseline = "top";

                const xStart = chartInst.chartArea.left + 4;

                chartInst.data.labels.forEach((label, i) => {
                const bars = chartInst.data.datasets
                    .map((ds, idx) => ({ ds, idx }))
                    .filter(o => o.ds.type !== "line")
                    .map(o => chartInst.getDatasetMeta(o.idx).data[i])
                    .sort((a,b) => a.y - b.y);

                if(!bars.length) return;

                const topBar = bars[0];
                const topY = topBar.y - topBar.height/2;
                const labelY = topY - LABEL_BLOCK_HEIGHT - 2;

                const textWidth  = ctx.measureText(label).width;
                const maxAllowed = chartInst.chartArea.right - xStart - 30;

                if (textWidth < maxAllowed) {
                    ctx.fillText(label, xStart, labelY);
                } else {
                    const shortLabel = label.slice(0, 25) + "...";
                    ctx.fillText(shortLabel, xStart, labelY);

                    const infoBtn = document.createElement("div");
                    infoBtn.className = "label-overlay";
                    infoBtn.textContent = "ⓘ";
                    infoBtn.style.position = "absolute";

                    const chartRect = chartInst.canvas.getBoundingClientRect();
                    const contRect  = card.getBoundingClientRect();
                    const labelW    = ctx.measureText(shortLabel).width;
                    const topOff    = labelY + chartRect.top - contRect.top + LINE_HEIGHT/2 - 15;
                    const leftOff   = chartInst.canvas.offsetLeft + xStart + labelW + 6;

                    Object.assign(infoBtn.style, {
                    top: `${topOff}px`,
                    left: `${leftOff}px`,
                    cursor: "pointer",
                    color: "#0074D9"
                    });
                    infoBtn.title = label;
                    infoBtn.addEventListener("click", () => alert(`Full label:\n${label}`));
                    card.appendChild(infoBtn);
                }

                const breakdown = chartInst.data.datasets
                    .filter(ds => ds.type !== "line")
                    .map(ds => `${ds.label}: ${ds.data[i]}%`)
                    .join(" | ");
                const niVal  = niValues[i];
                const niText = typeof niVal === "number" ? ` | NI: ${niVal.toFixed(1)}%` : "";
                ctx.fillText(`${breakdown}${niText}`, xStart, labelY + LINE_HEIGHT);
                });

                ctx.restore();
            }
            },
            {
            id: "drawNILines",
            afterDatasetsDraw(chartInst) {
                const ctx     = chartInst.ctx;
                const xScale  = chartInst.scales.x;

                ctx.save();
                ctx.strokeStyle = "#222";
                ctx.lineWidth = 4;
                ctx.setLineDash([]);

                chartInst.data.labels.forEach((_, i) => {
                const val = niValues[i];
                if (typeof val !== "number") return;

                const x = xScale.getPixelForValue(val);

                const bars = chartInst.data.datasets
                    .map((ds, idx) => ({ ds, idx }))
                    .filter(o => o.ds.type !== "line")
                    .map(o => chartInst.getDatasetMeta(o.idx).data[i])
                    .sort((a,b) => a.y - b.y);

                if(!bars.length) return;

                const yTop    = bars[0].y - bars[0].height/2;
                const yBottom = bars[bars.length-1].y + bars[bars.length-1].height/2;

                ctx.beginPath();
                ctx.moveTo(x, yTop);
                ctx.lineTo(x, yBottom);
                ctx.stroke();
                });

                ctx.restore();
            }
            }
        ]
        });

        window.chartInstances.push(chart);
    });
    });
}

document.querySelectorAll(".group-toggle").forEach((button) => {
    const content = button.nextElementSibling;
    const label = button.textContent.trim();

    if (label.startsWith("People and communities")) {
    content.style.display = "block";
    button.innerHTML = label.replace("▼", "▲");
    } else {
    content.style.display = "none";
    }

    // Toggle behavior
    button.addEventListener("click", () => {
    const isVisible = content.style.display !== "none";
    content.style.display = isVisible ? "none" : "block";
    button.innerHTML = label.replace(isVisible ? "▲" : "▼", isVisible ? "▼" : "▲");
    });
});

});

// scroll to and from geog selector and area profile builder
function smoothScrollTo(targetId, offset) {
const target = document.getElementById(targetId);
if (!target) return;

const elementPosition = target.getBoundingClientRect().top;
const offsetPosition = elementPosition + window.pageYOffset - offset;

window.scrollTo({
    top: offsetPosition,
    behavior: 'smooth'
});
}

function waitForImagesToLoad(container) {
const images = container.querySelectorAll('img');
const promises = Array.from(images).map(img => {
    return new Promise(resolve => {
    if (img.complete) resolve();
    else img.onload = img.onerror = resolve;
    });
});
return Promise.all(promises);
}

function downloadSummaryImage() {
    const selectedTab = document.querySelector('.view-tab.selected');
    const view = selectedTab ? selectedTab.getAttribute('data-view') : 'charts';

    const breakdownContainer = document.getElementById('breakdown-container');
    const contentSource = {
        charts: document.getElementById('charts-container'),
        tables: document.getElementById('tables-container'),
        tableComparison: document.getElementById('urban-rural-comparison'),
        chartComparison: document.getElementById('urban-rural-charts')
    }[view];

    if (!breakdownContainer || !contentSource) return;
    function replaceCloneCanvasesWithImages(originalRoot, cloneRoot) {
        const origCanvases  = Array.from(originalRoot.querySelectorAll('canvas'));
        const cloneCanvases = Array.from(cloneRoot.querySelectorAll('canvas'));
        const count = Math.min(origCanvases.length, cloneCanvases.length);

        for (let i = 0; i < count; i++) {
        const srcCanvas   = origCanvases[i];
        const destCanvas  = cloneCanvases[i];
        try {
            const dataUrl = srcCanvas.toDataURL('image/png');
            const img = new Image();
            img.src = dataUrl;
            img.style.maxWidth = '100%';
            img.style.display  = 'block';
            img.style.marginTop = '10px';

            // Swap the canvas for an <img> in the CLONE
            destCanvas.parentNode.replaceChild(img, destCanvas);
        } catch (e) {
        }
        }
    }
    const cloneWrapper = document.createElement('div');
    cloneWrapper.style.background  = '#fff';
    cloneWrapper.style.padding     = '20px';
    cloneWrapper.style.fontFamily  = 'sans-serif';
    cloneWrapper.style.maxWidth    = '1200px';
    cloneWrapper.style.margin      = '0 auto';

    const breakdownClone = breakdownContainer.cloneNode(true);
    const contentClone   = contentSource.cloneNode(true);

    // For CHARTS + CHART COMPARISON replace canvases in the clone.
    if (view === 'charts' || view === 'chartComparison') {
        replaceCloneCanvasesWithImages(contentSource, contentClone);
        contentClone.style.display = 'block';
        contentClone.style.visibility = 'visible';
    }

    // TABLES
    if (view === 'tables') {
        const originalTables = contentSource.querySelectorAll('.table-wrapper');
        contentClone.innerHTML = '';

        const rowContainer = document.createElement('div');
        Object.assign(rowContainer.style, {
        display: 'flex',
        flexWrap: 'wrap',
        gap: '20px',
        justifyContent: 'flex-start'
        });

        originalTables.forEach(originalWrapper => {
        const originalTable = originalWrapper.querySelector('table');
        const originalTitle = originalWrapper.querySelector('h3');
        if (!originalTable) return;

        const wrapper = document.createElement('div');
        Object.assign(wrapper.style, {
            flex: '1 1 45%',
            minWidth: '320px',
            background: '#fff',
            padding: '16px',
            borderRadius: '8px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
            boxSizing: 'border-box'
        });

        if (originalTitle) {
            const title = document.createElement('h3');
            title.textContent = originalTitle.textContent;
            title.style.margin = '0 0 12px 0';
            title.style.wordBreak = 'break-word';
            wrapper.appendChild(title);
        }

        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.style.marginTop = '8px';
        table.style.tableLayout = 'fixed';

        const rows = originalTable.querySelectorAll('tr');
        let columnCount = 0;
        if (rows.length > 0) {
            columnCount = rows[0].querySelectorAll('th, td').length;
        }

        rows.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');

            row.querySelectorAll('th, td').forEach((cell, columnIndex) => {
            const td = document.createElement('td');
            td.textContent = cell.textContent;

            if (columnIndex === 0) {
                td.style.width = '40%';
            } else {
                const remainingCols = columnCount - 1 || 1;
                td.style.width = `${60 / remainingCols}%`;
            }

            td.style.padding = '8px';
            td.style.border = '1px solid #ccc';
            td.style.fontSize = '14px';
            td.style.textAlign = 'left';
            td.style.verticalAlign = 'top';
            td.style.wordBreak = 'break-word';
            td.style.overflowWrap = 'break-word';
            td.style.whiteSpace = 'normal';

            if (rowIndex === 0) {
                td.style.fontWeight = 'bold';
                td.style.backgroundColor = '#04863E';
                td.style.color = '#fff';
            }

            tr.appendChild(td);
            });

            table.appendChild(tr);
        });

        wrapper.appendChild(table);
        rowContainer.appendChild(wrapper);
        });

        contentClone.appendChild(rowContainer);
    }

    // TABLE COMPARISON
    if (view === 'tableComparison') {
        const originalCards = contentSource.querySelectorAll(':scope > div > div');
        contentClone.innerHTML = '';

        const rowContainer = document.createElement('div');
        Object.assign(rowContainer.style, {
        display: 'flex',
        flexWrap: 'wrap',
        gap: '20px',
        justifyContent: 'flex-start'
        });

        originalCards.forEach(card => {
        const originalTable = card.querySelector('table');
        const originalTitle = card.querySelector('h3');
        if (!originalTable) return;

        const wrapper = document.createElement('div');
        Object.assign(wrapper.style, {
            flex: '0 0 100%',
            minWidth: '320px',
            background: '#fff',
            padding: '16px',
            borderRadius: '8px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
            boxSizing: 'border-box'
        });

        if (originalTitle) {
            const title = document.createElement('h3');
            title.textContent = originalTitle.textContent;
            title.style.margin = '0 0 12px 0';
            title.style.wordBreak = 'break-word';
            wrapper.appendChild(title);
        }

        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.style.marginTop = '8px';
        table.style.tableLayout = 'fixed';

        const rows = originalTable.querySelectorAll('tr');
        let columnCount = 0;
        if (rows.length > 0) {
            columnCount = rows[0].querySelectorAll('th, td').length;
        }

        rows.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');

            row.querySelectorAll('th, td').forEach((cell, columnIndex) => {
            const td = document.createElement('td');
            td.textContent = cell.textContent;

            if (columnIndex === 0) {
                td.style.width = '30%';
            } else {
                const remainingCols = columnCount - 1 || 1;
                td.style.width = `${70 / remainingCols}%`;
            }

            td.style.padding = '8px';
            td.style.border = '1px solid #ccc';
            td.style.fontSize = '14px';
            td.style.textAlign = 'left';
            td.style.verticalAlign = 'top';
            td.style.wordBreak = 'break-word';
            td.style.overflowWrap = 'break-word';
            td.style.whiteSpace = 'normal';

            if (rowIndex === 0) {
                td.style.fontWeight = 'bold';
                td.style.backgroundColor = '#04863E';
                td.style.color = '#fff';
            }

            tr.appendChild(td);
            });

            table.appendChild(tr);
        });

        wrapper.appendChild(table);
        rowContainer.appendChild(wrapper);
        });

        contentClone.appendChild(rowContainer);
    }

    cloneWrapper.appendChild(breakdownClone);
    cloneWrapper.appendChild(document.createElement('hr'));
    cloneWrapper.appendChild(contentClone);
    document.body.appendChild(cloneWrapper);

    cloneWrapper.offsetHeight;

    // Wait until images are ready
    (waitForImagesToLoad ? waitForImagesToLoad(cloneWrapper) : Promise.resolve()).then(() => {
        return new Promise(resolve => {
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
            html2canvas(cloneWrapper, {
                useCORS: true,
                scale: 2,
                backgroundColor: '#ffffff'
            }).then(resolve);
            });
        });
        });
    }).then(canvas => {
        const logo = new Image();
        logo.src = 'img/nisra-logo.svg';

        logo.onload = () => {
        const padding = 20;
        const maxLogoWidth = canvas.width * 0.25;
        const scaleFactor = Math.min(1, maxLogoWidth / logo.width);
        const scaledLogoWidth  = logo.width * scaleFactor;
        const scaledLogoHeight = logo.height * scaleFactor;

        const finalCanvas = document.createElement('canvas');
        finalCanvas.width  = canvas.width;
        finalCanvas.height = canvas.height + scaledLogoHeight + padding;

        const ctx = finalCanvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, finalCanvas.width, finalCanvas.height);

        ctx.drawImage(canvas, 0, 0);

        // logo at bottom right
        const x = finalCanvas.width - scaledLogoWidth - padding;
        const y = canvas.height + (padding / 2);
        ctx.drawImage(logo, x, y, scaledLogoWidth, scaledLogoHeight);

        // Trigger download
        const link = document.createElement('a');
        link.download = 'area-summary.png';
        link.href = finalCanvas.toDataURL('image/png');
        link.click();

        document.body.removeChild(cloneWrapper);
        };
    });
}


function downloadExcel() {
    // Grab selected IDs (zones) and trigger a refresh of comparison data
    const selectedArray = Array.from(selectedIdsExcel);
    renderUrbanRuralComparison(selectedArray);

    // Pull in cached data from the window/global scope
    const aggregated = window.latestAggregatedData || {};
    const comparison = window.urbanRuralComparisonData || {};
    const selectedCategories = window.chosenCategories || [];
    const niTotals = window.niTotals || {};

    // Setup workbook + metadata
    const workbook = XLSX.utils.book_new();
    const groups = ["Urban", "Rural", "Mixed"];
    const zoneType = window.selectedZoneType || 'sdz';
    const zoneTypeText = zoneType === 'sdz' ? 'Super Data Zone' : 'Data Zone';
    const totalZonesSelected = window.totalZonesSelected || 0;

    // Helper function: format a cell with value, type, format, and optional style
    // function formatCell(value, type, link, format, style = {}) {
    //   return { v: value, t: type, z: format, s: style };
    // }
    function formatCell(value, type, format, link) {
        let cell = { v: value, t: type,  z: format};
        if (link) {
        cell.l = { Target: link}; // Excel hyperlink object
        }
        return cell;
    }
    // ZONE BREAKDOWN SHEET
    // Groups selected zones by LGD
    const lgdGroups = {};
    selectedArray.forEach(id => {
        const mapData = window.selectedZoneDetails?.[id];
        if (!mapData) return;
        const lgd = mapData["LGD"];
        const status = mapData["Urban_mixed_rural_status"];
        const labelObj = mapData["Census 2021 Super Data Zone Label"] || mapData["Census 2021 Data Zone Label"];
        const zoneName = labelObj ? Object.keys(labelObj)[0] : null;
        if (!zoneName || !lgd) return;
        if (!lgdGroups[lgd]) lgdGroups[lgd] = [];
        lgdGroups[lgd].push({ zoneName, status });
    });
    const breakdownData = [];

    // Title line
    breakdownData.push([
        formatCell("DAERA Custom Area Profile Builder Extract", "s")
    ]);
    breakdownData.push([]);

    // Summary metadata lines
    const today = new Date().toLocaleDateString("en-GB", {
    year: "numeric",
    month: "long",
    day: "numeric"
    });
    // Add summary line
    breakdownData.push([
    formatCell(
        `The information presented in tables are combined from ${totalZonesSelected} ${zoneTypeText}s listed below:`,
        "s"
    )
    ]);
    // Add date line
    breakdownData.push([
    formatCell(
        `Date Extracted: ${today}`,
        "s",
    )
    ]);
    // Add blank row
    breakdownData.push([]);

    // LGD-by-LGD breakdown of selected zones
    Object.keys(lgdGroups).sort().forEach(lgd => {
        breakdownData.push([formatCell(lgd + " LGD", "s")]);
        breakdownData.push([
        formatCell("Area Name", "s"),
        formatCell("Area Type", "s")
        ]);

        lgdGroups[lgd].forEach(({ zoneName, status }) => {
        breakdownData.push([
            formatCell(zoneName, "s"),
            formatCell(status || "-", "s")
        ]);
        });

        breakdownData.push([]);
    });

    // Create sheet + column widths
    const breakdownSheet = XLSX.utils.aoa_to_sheet(breakdownData);
    breakdownSheet['!cols'] = [
        { wch: 13.33 },
        { wch: 9.8 }
    ];
    XLSX.utils.book_append_sheet(workbook, breakdownSheet, "Zone Breakdown");

    // CATEGORY SHEETS
    // One sheet per selected category
    (selectedCategories.length ? selectedCategories : Object.keys(aggregated)).forEach(category => {
        const sheetData = [];
        sheetData.push([formatCell(`${category}`, "s")]);

        // Category header
        sheetData.push([
        formatCell("Label", "s"),
        formatCell("Count", "s"),
        formatCell("Percentage", "s"),
        formatCell("NI %", "s")
        ]);

        // National-level data for this category
        const values = aggregated[category] || {};
        const totalCount = Object.values(values).reduce((acc, val) => acc + val, 0);
        for (const [label, count] of Object.entries(values)) {
        const percentage = totalCount > 0 ? count / totalCount : 0;
        const niVal = niTotals[category]?.[label];
        const niPercentage = typeof niVal === "number" ? niVal / 100 : null;
        sheetData.push([
            formatCell(label, "s"),
            formatCell(count, "n", "0"),
            formatCell(percentage, "n", "0.00%"),
            niPercentage !== null ? formatCell(niPercentage, "n", "0.0%") : formatCell("-", "s")
        ]);
        }
        sheetData.push([]);

        // Urban/Rural/Mixed breakdown tables
        groups.forEach(group => {
        const groupCategoryData = comparison[group]?.[category];
        if (!groupCategoryData || Object.keys(groupCategoryData).length === 0) return;
        sheetData.push([formatCell(`${category} – ${group}`, "s")]);
        sheetData.push([
            formatCell("Label", "s"),
            formatCell("Count", "s"),
            formatCell("Percentage", "s"),
            formatCell("NI %", "s")
        ]);
        const total = Object.values(groupCategoryData).reduce((acc, val) => acc + val, 0);
        for (const [label, count] of Object.entries(groupCategoryData)) {
            const pct = total > 0 ? count / total : 0;
            const niVal = niTotals[category]?.[label];
            const niPct = typeof niVal === "number" ? niVal / 100 : null;
            sheetData.push([
            formatCell(label, "s"),
            formatCell(count, "n", "0"),
            formatCell(pct, "n", "0.00%"),
            niPct !== null ? formatCell(niPct, "n", "0.0%") : formatCell("-", "s")
            ]);
        }
        sheetData.push([]);
        });

        // NOTES
        sheetData.push([])
        sheetData.push([
        formatCell("Notes on custom area aggregations", "s")
        ]);
        sheetData.push([])

        sheetData.push([
        formatCell(
            `Any aggregations created from Census Flexible Builder data may differ slightly from published Census figures.`,
            "s")
        ]);
        sheetData.push([
        formatCell(
            `For Census 2021, NISRA applied two Statistical Disclosure Control strategies: Targeted Record Swapping (TRS) and Cell Key Perturbation (CKP).`,
            "s")
        ]);
        sheetData.push([
        formatCell(
            `CKP may add small amounts of variation to some cells. Where two or more different aggregations are created, the totals of all cells may in turn be different.`,
            "s")
        ]);
        sheetData.push([
        formatCell(
            `Overall, the differences will be small and should not change the conclusions of any analysis or research.`,
            "s")
        ]);
        sheetData.push([])

        sheetData.push([
        formatCell(
            `Linked to the Statistical Disclosure Control Methods applied above, when viewing breakdowns at small geographical levels in this application,`,
            "s")
        ]);
        sheetData.push([
        formatCell(
            `cell counts of under 5 may be seen. The use of TRS and CKP, mean this number could be anything from 0-4, or could have been swapped with another census record entry elsewhere.`,
            "s")
        ]);
        sheetData.push([])

        sheetData.push([
        formatCell(
            `For more information, please refer to the NISRA statistical disclosure control methodology:`,
            "s")
        ]);
        sheetData.push([
        formatCell(
            `NISRA Statistical Disclosure Control Methodology`,
            "s", 
            undefined,
            `https://www.nisra.gov.uk/files/nisra/publications/statistical-disclosure-control-methodology-for-2021-census.pdf`)
        ]);

        // Finalise worksheet and append
        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, category.substring(0, 31));
    });

        // SAVE FILE
    const now = new Date();
    // Format date as DD-MM-YYYY
    const datePart = now.toLocaleDateString("en-GB").replace(/\//g, "-");
    // Format time as HH-MM (colon replaced with dash for filename safety)
    const timePart = now.toLocaleTimeString("en-GB", {
        hour: "2-digit",
        minute: "2-digit",
        hour12: false
    }).replace(/:/g, "-");
    // Combine into filename
    const filename = `DAERA Custom Area Profile Extract-${datePart} ${timePart}.xlsx`;
    XLSX.writeFile(workbook, filename);
}
