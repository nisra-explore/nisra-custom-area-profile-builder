/*
================================================================================
POSTCODE SEARCH FUNCTIONALITY
================================================================================
This section handles postcode lookup and automatic map selection.
Data source: CPD_LIGHT_JULY_2024.csv (hosted on GitHub)
================================================================================
*/

// Global variable to cache the postcode data (populated on first search)
let postcodeDataCache = null;
    let radiusKm = 2;   // ✅ shared
    let simplifyTolMeters = 10;
    let bufferMeters = 0;
    
/**
 * Normalize postcode to uppercase and remove spaces
 * Input: "BT12 3AB" or "bt123ab"
 * Output: "BT123AB"
 */
function normalisePostcode(postcode) {
  return postcode.replace(/\s/g, "").toUpperCase();
}

/**
 * Fetch and parse the CPD CSV file
 * Returns a Promise that resolves to an array of postcode records
 * Each record has columns indexed by position
 */
function fetchPostcodeData() {
  // Return cached data if available (avoid refetching)
  if (postcodeDataCache !== null) {
    return Promise.resolve(postcodeDataCache);
  }

  const csvUrl = 'create-js\\inputs\\CPD_LIGHT.csv';
  
  return fetch(csvUrl)
    .then(response => response.text())
    .then(data => {
      // Parse CSV into rows
      const rows = data.split('\n');
      postcodeDataCache = rows;
      return rows;
    })
    .catch(error => {
      console.error('Error fetching postcode data:', error);
      throw new Error('Failed to load postcode database');
    });
}



function parseCSVLine(line) {
  const result = [];
  let current = '';
  let insideQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if (char === ',' && !insideQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }

  result.push(current);
  return result;
}


/**
 * Look up a postcode in the CSV data
 * Parameters:
 *   - postcode: string (e.g., "BT12 3AB")
 *   - zoneType: string ('sdz', 'dz', or 'dea')
 * Returns:
 *   - Object with matched geography code and name, or null if not found
 */

function lookupPostcode(postcode, zoneType) {
  const normalised = normalisePostcode(postcode);

  const columnMap = {
    lgd: { code: 4, name: 5 },          // Local Government District
    dea: { code: 8, name: 9 },          // District Electoral Area
    sdz: { code: 20, name: 21 },      // Super Data Zone
    dz: { code: 18, name: 19 }        // Data Zone
  };

  if (!columnMap[zoneType]) {
    throw new Error(`Invalid zone type: ${zoneType}`);
  }

  const { code, name } = columnMap[zoneType];

  if (!postcodeDataCache) return null;

  for (let row of postcodeDataCache) {
    const cols = parseCSVLine(row);  // ✅ FIXED
    const csvPostcode = normalisePostcode(cols[0]);

    if (csvPostcode === normalised) {
      return {
        code: cols[code],
        name: cols[name]
      };
    }
  }

  return null;
}

/**
 * Main postcode search handler
 * This function:
 *   1. Validates the input
 *   2. Fetches postcode data
 *   3. Looks up the postcode
 *   4. Selects the matching geography on the map
 */

let latestSearchId = 0;

function handlePostcodeSearch() {
  const searchId = ++latestSearchId;
  const input = document.getElementById('postcode-input');
  const statusDiv = document.getElementById('postcode-status');
  const postcode = input.value.trim();

  statusDiv.textContent = '';
  statusDiv.className = 'postcode-status';

  if (!postcode) {
    statusDiv.textContent = 'Please enter a postcode';
    statusDiv.className = 'postcode-status error';
    return;
  }

  statusDiv.textContent = 'Searching...';
  statusDiv.className = 'postcode-status info';

  const zoneSelector = document.getElementById('zone-selector');
  const zoneType = zoneSelector ? zoneSelector.value : 'sdz';

  fetchPostcodeData()
    .then(() => {
      const result = lookupPostcode(postcode, zoneType);

      // Handle no match found
      if (!result || !result.code || result.code === 'NA' || result.code === '000000000') {
      console.warn('❌ Invalid postcode entered');
      statusDiv.textContent = 'Invalid postcode';
      statusDiv.className = 'postcode-status error';
      return;
      }

  // Select on map
  if (window.selectGeographyByCode) {
    window.selectGeographyByCode(result.code, zoneType);

    // Delay zooming slightly to ensure map has processed the selection and updated feature states
    setTimeout(() => {
      zoomToResult(result, zoneType, searchId);
    }, 100);
  }

  // Update status with found location and make it clickable to navigate to profile
  if (result && result.code && result.code !== 'NA') {

    // Update status with found location and make it clickable to navigate to profile
    statusDiv.textContent = `Found: ${zoneSelector.options[zoneSelector.selectedIndex].text} - ${result.name}`;
    statusDiv.className = 'postcode-status success';

    // Wait for selection to complete before navigating
    // Make the found result clickable to navigate to the Area Profile
    // Build an inline link so the user can choose to go to the profile page
    // (do not auto-navigate)
    const resultLinkId = 'postcode-result-link';
    // Set the status text with a clickable link
    statusDiv.innerHTML = `Found: <a href="#" id="${resultLinkId}">${zoneSelector.options[zoneSelector.selectedIndex].text} - ${result.name}</a>`;

    const resultLink = document.getElementById(resultLinkId);

    if (resultLink) {
      // On click, trigger the same action as the "Build Profile" button, but only if we have a valid selection
      resultLink.addEventListener('click', (ev) => {
        ev.preventDefault();
        
        const btn = document.getElementById('build-profile-btn');
        
        if (btn && selectedIds.size > 0) {
          btn.click();
        } else {
          // In case selection hasn't been applied yet, try selecting again then navigate
          if (window.selectGeographyByCode) {
            window.selectGeographyByCode(result.code, zoneType);
          }
          setTimeout(() => {
            const btn2 = document.getElementById('build-profile-btn');
            if (btn2 && selectedIds.size > 0) btn2.click();
          }, 400);
        }
      });
    }
  }
    })
    .catch(err => {
      statusDiv.textContent = 'Error searching postcode';
      statusDiv.className = 'postcode-status error';
      console.error(err);
    });
}

// After finding the matching geography code, this function will select it on the map and zoom to it
function zoomToResult(result, zoneType, searchId, attempts = 0) {

  console.log('--- zoomToResult START ---');
  console.log('Result:', result);
  console.log('Zone type:', zoneType);
  console.log('Search ID:', searchId);

  if (searchId !== latestSearchId) {
    console.log(' Ignoring stale zoom call');
    return;
  }


  // wait until map finishes updating
  // map.once('idle', () => {
  //   console.log('Map idle triggered');

    // Query features for the current zone type and filter to our selected IDs
    const zoneIds = getZoneIdsFor(zoneType);
    console.log('Zone IDs:', zoneIds);

    const startingFeatures = map.querySourceFeatures(zoneIds.source, {
      sourceLayer: zoneIds.sourceLayer
    });

    const startingSelectedFeatures = startingFeatures.filter(f =>
      String(f.id) === String(result.code) ||
      String(f.properties?.id) === String(result.code) ||
      String(f.properties?.ID) === String(result.code)
    );

    if (startingSelectedFeatures.length > 0) {
      const startingBbox = turf.bbox({
        type: 'FeatureCollection',
        features: startingSelectedFeatures
      });

      const currentBounds = map.getBounds();
      const alreadyVisible = currentBounds.contains([startingBbox[0], startingBbox[1]]) &&
        currentBounds.contains([startingBbox[2], startingBbox[3]]);

      if (alreadyVisible) {
        console.log('➡️ Target already in view, skipping full zoom-out');
        map.stop();
        map.fitBounds(startingBbox, {
          padding: 40,
          maxZoom: 11,
          duration: 800,
          curve: 1.5
        });
        return;
      }
    }

    console.log('➡️ Step 1: move map to approximate location');
    
    if (attempts === 0) {
      console.log('➡️ Step 1: move map to approximate location');

      // ✅ STEP 1: force map to load tiles for wider area
      map.flyTo({
        center: [-6.5, 54.7], // NI-wide center (important!)
        zoom: 7,
        duration: 600,
        essential: true
      });
    }

    
      // ✅ STEP 2: wait for tiles to load
      map.once('moveend', () => {
        if (searchId !== latestSearchId) return;

        console.log('➡️ Step 2: tiles should now be available');

        const features = map.querySourceFeatures(zoneIds.source, {
          sourceLayer: zoneIds.sourceLayer
        });



    // Use querySourceFeatures for better performance and to get features even if off-screen
    // const features = map.querySourceFeatures(zoneIds.source, {
    //   sourceLayer: zoneIds.sourceLayer
    // });
    
    // const features = map.queryRenderedFeatures({
    //   layers: [zoneIds.fillLayer]
    // });



    console.log('Total features returned:', features.length);

    // Filter features to those matching our selected IDs (checking multiple possible ID properties)
    // const selectedFeatures = features.filter(f =>
    //   selectedIds.has(String(f.id)) ||
    //   selectedIds.has(String(f.properties?.id)) ||
    //   selectedIds.has(String(f.properties?.ID))
    // );

    // Since selectedIds is a Set of strings, we need to check if any of the possible ID properties match any of the selected IDs. We can convert the feature's ID and properties to strings and check against the selectedIds set.
    const selectedFeatures = features.filter(f =>
      String(f.id) === String(result.code) ||
      String(f.properties?.id) === String(result.code) ||
      String(f.properties?.ID) === String(result.code)
    );
    
    console.log('Selected features count:', selectedFeatures.length);
    console.log('Selected feature IDs:', selectedFeatures.map(f => f.id));

    console.log('Attempt:', attempts, 'Rendered matches:', selectedFeatures.length);


    // Retry until it's actually visible/rendered
      if (selectedFeatures.length === 0) {
        if (attempts < 6) {
          requestAnimationFrame(() => {
            zoomToResult(result, zoneType, searchId, attempts + 1);
          });
        } else {
          console.warn('❌ Feature never rendered');
        }
        return;
      }


    // If we found matching features, fit bounds to them
    if (selectedFeatures.length > 0) {
      try {
        const bbox = turf.bbox({
          type: 'FeatureCollection',
          features: selectedFeatures
        });

        console.log('Sample LGD feature:', features[0]);
        console.log('Computed bbox:', bbox);
        
        // ✅ Stop any ongoing animations
        map.stop();
        console.log('Map stopped');

        // ✅ Step 1: force a consistent zoom level
        map.easeTo({
          zoom: 7.5,
          duration: 300,
          essential: true
        });

        // ✅ Step 2: then zoom to the new area
        setTimeout(() => {
          console.log('Calling fitBounds now...');
          map.fitBounds(bbox, {
            padding: 40,
            maxZoom: 11,
            duration: 800,
            curve: 1.5
          });
        }, 300);


      } catch (err) {
        console.warn('Zoom failed:', err);
      }
    } else {
      console.warn('No selected features found for zoom');
    }
  });
}


// Global variables for map and data (needed by postcode functions)
let map = null;
let selectedIds = new Set();
let lgdData = {};
let deaData = {};
let sdzData = {};
let dzData = {};
let lgdNameToId = new Map();
let lgdCodeToId = new Map();


// Helper functions moved to global scope so postcode functions can access them (MG)
function getDataSourceFor(zone) {
  return zone === 'dz' ? dzData : zone === 'dea' ? deaData : zone === 'lgd' ? lgdData : zone === 'sdz' ? sdzData : null;
}

function getZoneIdsFor(zone) {
  if (zone === 'dz') return { source: 'dz2021', sourceLayer: 'DZ2021_clipped', fillLayer: 'dz-fill' };
  if (zone === 'dea') return { source: 'dea2014', sourceLayer: 'DEA2014_clipped', fillLayer: 'dea-fill' };
  if (zone === 'lgd') return { source: 'lgd2014', sourceLayer: 'LGD2014_clipped', fillLayer: 'lgd-fill' };
  return { source: 'sdz2021', sourceLayer: 'SDZ2021_clipped', fillLayer: 'sdz-fill' };
}

function getLabelKeyFor(zone) {
  return zone === 'dz'
    ? "Census 2021 Data Zone Label"
    : zone === 'dea'
      ? "District Electoral Area 2014 Label"
      : zone === 'lgd'
        ? "Local Government District 2021 Label"
        : zone === 'sdz'
        ? "Census 2021 Super Data Zone Label"
        : null;
}

// Global wrapper for selectGeographyByCode - defined locally inside map.on('load') after it's available
// This ensures postcode search can call it without errors
window.selectGeographyByCode = function(code, zoneType) {
  // This will be replaced by the real implementation inside map.on('load')
  console.warn('selectGeographyByCode called before map loaded');
};

document.addEventListener('DOMContentLoaded', function () {
  
  markEmptyCategoryGroups();

  map = new maplibregl.Map({
    container: 'map',
    style: 'https://raw.githubusercontent.com/NISRA-Tech-Lab/map_tiles/main/basemap_styles/style-omt.json',
    center: [-6.8, 54.65],
    zoom: 7.5,
    minZoom: 7.5,
    maxZoom: 13,
    maxBounds: [[-9.20, 53.58], [-4.53, 55.72]]
  });

  const vis = (id, show) => map.setLayoutProperty(id, 'visibility', show ? 'visible' : 'none');

  function getCircleRadius() {
    return Math.max(0.05, +radiusKm || 2);
  }



  // Show coordinates on mouse move
  map.on('mousemove', (e) => {
    const lng = e.lngLat.lng.toFixed(5);
    const lat = e.lngLat.lat.toFixed(5);


  // ✅ add this block
    if (drawToolActive) {
      const center = [e.lngLat.lng, e.lngLat.lat];

      const radius = getCircleRadius();

      const circle = turf.circle(center, radius, {
        steps: 128,
        units: 'kilometers'
      });

      map.getSource('circle-preview').setData(circle);
    }


    const x = Math.round(e.point.x);
    const y = Math.round(e.point.y);

    const coordsDiv = document.getElementById('coords');
    if (coordsDiv) {
      coordsDiv.innerHTML = `
        <span class="coords-line-1">Lng: ${lng}, Lat: ${lat}</span>
        <span class="coords-line-2">X: ${x}, Y: ${y}</span>
        `;
    }
  });

  map.on('mouseleave', () => {
    const coordsDiv = document.getElementById('coords');
    if (coordsDiv) {
      coordsDiv.innerText = 'Move mouse over map';
    }
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

    vis('sdz-fill', selected === 'sdz'); vis('dz-fill', selected === 'dz'); vis('dea-fill', selected === 'dea'); vis('lgd-fill', selected === 'lgd');
    vis('sdz-outline-default', selected === 'sdz'); vis('dz-outline-default', selected === 'dz'); vis('dea-outline-default', selected === 'dea'); vis('lgd-outline-default', selected === 'lgd');
    vis('sdz-outline-hover', selected === 'sdz'); vis('dz-outline-hover', selected === 'dz'); vis('dea-outline-hover', selected === 'dea'); vis('lgd-outline-hover', selected === 'lgd');

    AREA_INDEX[activeZone] = null;
    populateDatalist(activeZone);
    populateLGDButtons();

    // Clear everything and reset UI

    const defaultCategories = ['Age (4 Categories)', 'Sex Label'];
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
    clearSelections();

  }

  // Initial link setup on page load
  updateSourceLink();

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

    // BUILD LGD DATA FROM SDZ
    lgdData = {};

    Object.entries(sdzData).forEach(([sdzCode, record]) => {
      const lgdName = record["LGD"];
      if (!lgdName) return;

      if (!lgdData[lgdName]) {
        lgdData[lgdName] = {
          "Local Government District 2021 Label": {},
          population: 0,
          LGD: lgdName
        };
        lgdData[lgdName]["Local Government District 2021 Label"][lgdName] = 0;
      }

      // get SDZ population
      const popObj = record["Census 2021 Super Data Zone Label"];
      const pop = popObj ? Object.values(popObj)[0] : 0;

      lgdData[lgdName]["Local Government District 2021 Label"][lgdName] += pop;
      lgdData[lgdName].population += pop;

      // Aggregate all category breakdowns for the LGD
      for (const [category, values] of Object.entries(record)) {
        if (category === "LGD" || category === "Census 2021 Super Data Zone Label" || category === "Urban_mixed_rural_status") {
          continue;
        }
        if (!values || typeof values !== 'object') {
          continue;
        }

        if (!lgdData[lgdName][category]) {
          lgdData[lgdName][category] = {};
        }

        for (const [label, count] of Object.entries(values)) {
          const numericValue = typeof count === 'number' ? count : Number(count);
          if (Number.isNaN(numericValue)) continue;
          lgdData[lgdName][category][label] = (lgdData[lgdName][category][label] || 0) + numericValue;
        }
      }
    });

    console.log("LGD Data:", lgdData); // debug

      AREA_INDEX.sdz = AREA_INDEX.dz = AREA_INDEX.dea = null;
      ensureIndexFor(activeZone);
      populateDatalist(activeZone);
      populateLGDButtons();
      decorateCategoryBadges()
    });

  // Function that toggles urban/rural fill based on zone selection
  let fillVisible = true;

  const AREA_INDEX = { sdz: null, dz: null, dea: null }; // built on demand

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
    mapDiv.style.borderRadius = '4px';
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
          },
            lgd2014: {
            type: 'vector',
            tiles: ['https://raw.githubusercontent.com/nisra-explore/map_tiles/main/lgd2014/{z}/{x}/{y}.pbf'],
            promoteId: 'lgd_code'
          }
        },
        layers: [
          { id: 'bg', type: 'raster', source: 'osmRaster' }
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

      addZoneLayers('sdz-preview', 'sdz2021', 'SDZ2021_clipped');
      addZoneLayers('dz-preview', 'dz2021', 'DZ2021_clipped');
      addZoneLayers('dea-preview', 'dea2014', 'DEA2014_clipped');
      addZoneLayers('lgd-preview', 'lgd2014', 'LGD2014_clipped');

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
    const showDZ = activeZone === 'dz';
    const showDEA = activeZone === 'dea';
    const showLGD = activeZone === 'lgd';
    previewMap.setLayoutProperty('sdz-preview-fill', 'visibility', showSDZ ? 'visible' : 'none');
    previewMap.setLayoutProperty('sdz-preview-outline', 'visibility', showSDZ ? 'visible' : 'none');
    previewMap.setLayoutProperty('dz-preview-fill', 'visibility', showDZ ? 'visible' : 'none');
    previewMap.setLayoutProperty('dz-preview-outline', 'visibility', showDZ ? 'visible' : 'none');
    previewMap.setLayoutProperty('dea-preview-fill', 'visibility', showDEA ? 'visible' : 'none');
    previewMap.setLayoutProperty('dea-preview-outline', 'visibility', showDEA ? 'visible' : 'none');
    previewMap.setLayoutProperty('lgd-preview-fill', 'visibility', showLGD ? 'visible' : 'none');
    previewMap.setLayoutProperty('lgd-preview-outline', 'visibility', showLGD ? 'visible' : 'none');
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
    const layerId = z === 'sdz' ? 'sdz-preview-fill' : z === 'dz' ? 'dz-preview-fill' : z === 'dea' ? 'dea-preview-fill' : 'lgd-preview-fill';

    // zone switch: clear old feature-state & show correct layers
    if (previewActiveZone !== z) {
      const { source: oldSrc, sourceLayer: oldSL } = getZoneIdsFor(previewActiveZone || z);
      Array.from(previewSelectedIds).forEach(id => {
        try { previewMap.setFeatureState({ source: oldSrc, sourceLayer: oldSL, id }, { hovered: false }); } catch { }
      });
      previewSelectedIds.clear();
      previewActiveZone = z;
      syncPreviewVisibility();
    }

    // sync feature-state for current selection 
    const sel = new Set(Array.from(selectedIds).map(String));
    Array.from(previewSelectedIds).forEach(id => {
      if (!sel.has(String(id))) {
        try { previewMap.setFeatureState({ source, sourceLayer, id }, { hovered: false }); } catch { }
        previewSelectedIds.delete(id);
      }
    });
    sel.forEach(id => {
      if (!previewSelectedIds.has(id)) {
        try { previewMap.setFeatureState({ source, sourceLayer, id }, { hovered: true }); } catch { }
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
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      for (const f of useFeats) {
        try {
          const bb = turf.bbox({ type: 'Feature', geometry: f.geometry, properties: {} });
          minX = Math.min(minX, bb[0]); minY = Math.min(minY, bb[1]);
          maxX = Math.max(maxX, bb[2]); maxY = Math.max(maxY, bb[3]);
        } catch { }
      }
      if (isFinite(minX)) {
        previewMap.fitBounds([[minX, minY], [maxX, maxY]], { padding: 18, duration: 0 });
      }
    });
  }

  // Build a lightweight search index

  function buildAreaIndexFor(zone) {
    const dataSource = getDataSourceFor(zone);
    const labelKey = getLabelKeyFor(zone);

    const byKey = new Map(); // id -> item, lgd:* -> lgd shim
    const byName = new Map(); // name -> [items]
    const items = [];

    if (!dataSource || !Object.keys(dataSource).length) {
      AREA_INDEX[zone] = { byKey, byName, items };
      return;
    }

    for (const [idRaw, rec] of Object.entries(dataSource)) {
      const id = isNaN(+idRaw) ? idRaw : +idRaw;
      const lobj = rec?.[labelKey] || {};
      const name = Object.keys(lobj)[0] || String(id);
      const lgd = rec?.LGD || '';

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
      .sort((a, b) => a.name.localeCompare(b.name) || a.lgd.localeCompare(b.lgd))
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
    } catch { }
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

    map.addSource('lgd2014', {
      type: 'vector',
      tiles: [
        'https://raw.githubusercontent.com/nisra-explore/map_tiles/main/lgd2014/{z}/{x}/{y}.pbf'
      ],
      promoteId: { 'LGD2014_clipped': 'lgd_code' }
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
      id: 'lgd-fill',
      type: 'fill',
      source: 'lgd2014',
      'source-layer': 'LGD2014_clipped',
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
      id: 'lgd-outline-default',
      type: 'line',
      layout: {
        visibility: 'none'
      },
      source: 'lgd2014',
      'source-layer': 'LGD2014_clipped',
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

    map.addLayer({
      id: 'lgd-outline-hover',
      type: 'line',
      layout: {
        visibility: 'none'
      },
      source: 'lgd2014',
      'source-layer': 'LGD2014_clipped',
      paint: {
        'line-color': [
          'case',
          ['boolean', ['feature-state', 'hovered'], false], '#000000',
          'rgba(0,0,0,0)'
        ],
        'line-width': 1.5
      }
    });

    map.addSource('circle-preview', {
      type: 'geojson',
      data: {
        type: 'FeatureCollection',
        features: []
      }
    });

    map.addLayer({
      id: 'circle-preview-fill',
      type: 'fill',
      source: 'circle-preview',
      paint: {
        'fill-color': '#f59e0b',   // ✅ orange (match final circle)
        'fill-opacity': 0.3        // match final circle
      }
    });

    map.addLayer({
      id: 'circle-preview-outline',
      type: 'line',
      source: 'circle-preview',
      paint: {
     'line-color': '#f59e0b',   // ✅ same orange
     'line-width': 2
      }
    });


    // Expose selectGeographyByCode globally so postcode search can call it
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
        const value = labelObj[zoneName];
        const formattedValue = value.toLocaleString();  // ✅ adds commas

        popupHtml = `<div><strong>${zoneName}</strong> pop. ${formattedValue}</div>`;
      }

      map.getCanvas().style.cursor = 'pointer';
      popup.setLngLat(e.lngLat).setOffset([0, -10]).setHTML(popupHtml).addTo(map);
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
        const value = labelObj[zoneName];
        const formattedValue = value.toLocaleString();  // ✅ adds commas

        popupHtml = `<div><strong>${zoneName}</strong> pop. ${formattedValue}</div>`;
      }

      map.getCanvas().style.cursor = 'pointer';
      popup.setLngLat(e.lngLat).setOffset([0, -10]).setHTML(popupHtml).addTo(map);
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

        const value = labelObj[zoneName];
        const formattedValue = value.toLocaleString();  // ✅ adds commas

        popupHtml = `<div><strong>${zoneName}</strong> pop. ${formattedValue}</div>`;
      }


      map.getCanvas().style.cursor = 'pointer';
      popup.setLngLat(e.lngLat).setOffset([0, -10]).setHTML(popupHtml).addTo(map);
    });



    map.on('mousemove', 'lgd-fill', (e) => {
      if (!e.features.length) return;

      const feature = e.features[0];

      // const id = feature.id;
      // const data = lgdData?.[id];

      const lgdName = feature.properties.LGDNAME || feature.properties.lgd_name || feature.properties.LGD;
      const data = lgdData?.[lgdName];

      let popupHtml = '<div><strong>Data not found</strong></div>';

      if (data?.["Local Government District 2021 Label"]) {
        const labelObj = data["Local Government District 2021 Label"];
        const zoneName = Object.keys(labelObj)[0];

        const value = labelObj[zoneName];
        const formattedValue = value.toLocaleString();  // ✅ adds commas

        popupHtml = `<div><strong>${zoneName}</strong>: pop. ${formattedValue}</div>`;
      }

      map.getCanvas().style.cursor = 'pointer';
      popup.setLngLat(e.lngLat).setOffset([0, -10]).setHTML(popupHtml).addTo(map);
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

    map.on('mouseleave', 'lgd-fill', () => {
      map.getCanvas().style.cursor = '';
      popup.remove();
    });

    // Click to select/deselect areas - SDZ layer
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

      // // --- Zoom to selected area(s) ---
      // // Get all selected features' geometries and union their bbox
      //  const selectedArray = Array.from(selectedIds);
      //  if (selectedArray.length > 0) {
      //    const features = map.querySourceFeatures('sdz2021', { sourceLayer: 'SDZ2021_clipped' })
      //      .filter(f => selectedArray.includes(f.id));
      //    if (features.length > 0) {
      //      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      //      features.forEach(f => {
      //        try {
      //          const bb = turf.bbox({ type: 'Feature', geometry: f.geometry, properties: {} });
      //          minX = Math.min(minX, bb[0]);
      //          minY = Math.min(minY, bb[1]);
      //          maxX = Math.max(maxX, bb[2]);
      //          maxY = Math.max(maxY, bb[3]);
      //        } catch {}
      //      });
      //      if (isFinite(minX) && isFinite(minY) && isFinite(maxX) && isFinite(maxY)) {
      //        map.fitBounds([[minX, minY], [maxX, maxY]], { padding: 40, duration: 600 });
      //      }
      //    }
      //  }
    });

    // Click to select/deselect areas - DZ layer
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

      // // --- Zoom to selected area(s) ---
      //  const selectedArray = Array.from(selectedIds);
      //  if (selectedArray.length > 0) {
      //    const features = map.querySourceFeatures('dz2021', { sourceLayer: 'DZ2021_clipped' })
      //      .filter(f => selectedArray.includes(f.id));
      //    if (features.length > 0) {
      //      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      //      features.forEach(f => {
      //        try {
      //          const bb = turf.bbox({ type: 'Feature', geometry: f.geometry, properties: {} });
      //          minX = Math.min(minX, bb[0]);
      //          minY = Math.min(minY, bb[1]);
      //          maxX = Math.max(maxX, bb[2]);
      //          maxY = Math.max(maxY, bb[3]);
      //        } catch {}
      //      });
      //      if (isFinite(minX) && isFinite(minY) && isFinite(maxX) && isFinite(maxY)) {
      //        map.fitBounds([[minX, minY], [maxX, maxY]], { padding: 40, duration: 600 });
      //      }
      //    }
      //  }
    });

    // Click to select/deselect areas - DEA layer
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

      // --- Zoom to selected area(s) ---
      // const selectedArray = Array.from(selectedIds);
      // if (selectedArray.length > 0) {
      //   const features = map.querySourceFeatures('dea2014', { sourceLayer: 'DEA2014_clipped' })
      //     .filter(f => selectedArray.includes(f.id));
      //   if (features.length > 0) {
      //     let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      //     features.forEach(f => {
      //       try {
      //         const bb = turf.bbox({ type: 'Feature', geometry: f.geometry, properties: {} });
      //         minX = Math.min(minX, bb[0]);
      //         minY = Math.min(minY, bb[1]);
      //         maxX = Math.max(maxX, bb[2]);
      //         maxY = Math.max(maxY, bb[3]);
      //       } catch {}
      //     });
      //     if (isFinite(minX) && isFinite(minY) && isFinite(maxX) && isFinite(maxY)) {
      //       map.fitBounds([[minX, minY], [maxX, maxY]], { padding: 40, duration: 600 });
      //     }
      //   }
      // }
    });


// Click to select/deselect areas - LGD layer
    map.on('click', 'lgd-fill', (e) => {
      if (drawToolActive) return;
      if (!e.features.length) return;

      const feature = e.features[0];
      const id = feature.id;
      const lgdName = feature.properties.LGDNAME || feature.properties.lgd_name || feature.properties.LGD;
      const selectionKey = lgdName || id;

      const isSelected = selectedIds.has(selectionKey);

      if (isSelected) {
        selectedIds.delete(selectionKey);
        lgdNameToId.delete(lgdName);
        map.setFeatureState(
          { source: 'lgd2014', sourceLayer: 'LGD2014_clipped', id },
          { hovered: false }
        );
      } else {
        selectedIds.add(selectionKey);
        lgdNameToId.set(lgdName, id);
        map.setFeatureState(
          { source: 'lgd2014', sourceLayer: 'LGD2014_clipped', id },
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
      data: { type: 'FeatureCollection', features: [] }
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
      if (activeZone === 'dz') return { source: 'dz2021', sourceLayer: 'DZ2021_clipped', fillLayer: 'dz-fill' };
      if (activeZone === 'dea') return { source: 'dea2014', sourceLayer: 'DEA2014_clipped', fillLayer: 'dea-fill' };
      if (activeZone === 'lgd') return { source: 'lgd2014', sourceLayer: 'LGD2014_clipped', fillLayer: 'lgd-fill' };
      return { source: 'sdz2021', sourceLayer: 'SDZ2021_clipped', fillLayer: 'sdz-fill' };
    }

    function selectByGeometry(geom, mode = 'add') {
      const feature = geom.type === 'Feature' ? geom : { type: 'Feature', geometry: geom, properties: {} };

      const bbox = turf.bbox(feature);
      const sw = map.project([bbox[0], bbox[1]]);
      const ne = map.project([bbox[2], bbox[3]]);
      const { source, sourceLayer, fillLayer } = zoneIds();

      const candidates = map.queryRenderedFeatures([sw, ne], { layers: [fillLayer] })
        .filter(f => {
          try {
            return turf.booleanIntersects(feature.geometry, f.geometry);
          } catch {
            return false;
          }
        });


      if (mode === 'replace') {
        selectedIds.forEach(key => {
          const featureId = activeZone === 'lgd' ? (lgdNameToId.get(key) || key) : key;
          map.setFeatureState({ source, sourceLayer, id: featureId }, { hovered: false });
        });
        selectedIds.clear();
        if (activeZone === 'lgd') lgdNameToId.clear();
      }

      if (mode === 'subtract') {
        candidates.forEach(f => {
          const id = f.id;
          const key = activeZone === 'lgd' ? (f.properties.LGDNAME || f.properties.lgd_name || f.properties.LGD || id) : id;
          if (selectedIds.has(key)) {
            selectedIds.delete(key);
            if (activeZone === 'lgd') lgdNameToId.delete(key);
            map.setFeatureState({ source, sourceLayer, id }, { hovered: false });
          }
        });
      } else {
        // 'add' (default) or after 'replace'
        candidates.forEach(f => {
          const id = f.id;
          const key = activeZone === 'lgd' ? (f.properties.LGDNAME || f.properties.lgd_name || f.properties.LGD || id) : id;
          if (!selectedIds.has(key)) {
            selectedIds.add(key);
            if (activeZone === 'lgd') lgdNameToId.set(key, id);
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
        // zoomIn: `<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.35-4.35M11 8v6M8 11h6"/></svg>`,
        // zoomOut: `<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.35-4.35M8 11h6"/></svg>`,
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
      // bar.appendChild(makeBtn("zoomInBtn", "Zoom in", svgs.zoomIn));
      // bar.appendChild(makeBtn("zoomOutBtn", "Zoom out", svgs.zoomOut));

      // Append toolbar to target
      target.appendChild(bar);
      console.log("Draw toolbar inserted into #draw-toolbar-container");
    };

    
    function clearDrawnCircle() {
      const source = map.getSource('draw-geom');
      if (source) {
        source.setData({
          type: 'FeatureCollection',
          features: []
        });
      }

      lastDrawnFeature = null;
    }



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
  background: #00205b;
  border-radius: 10px;
  color: #fff;
  font: 14px/1.2 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  margin-top: 1rem;  
  margin: 1rem auto;
  width: fit-content; /* or max-width: 100%; */

}
#draw-toolbar .icon-btn {
  position: relative;
  inline-size: 44px;
  block-size: 44px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: transparent;
  border: 1px solid rgba(255,255,255,.3);
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
  width: 24px;
  height: 24px;
  stroke: currentColor;
  fill: none;
  stroke-width: 2;
}
#draw-toolbar .icon-btn[data-badge]::after {
  content: attr(data-badge);
  position: absolute;
  inset-block-start: 0;            /* logical 'top'  */
  inset-inline-end: 0;             /* logical 'right'*/
  transform: translate(40%,-40%);  /* nudge outside corner */
  min-inline-size: 1.25em;         /* scales with font */
  block-size: 1.25em;
  padding-inline: .25em;
  background: #1ea672;
  color: #fff;
  border-radius: 999px;
  font: 600 12px/1 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  display: grid;
  place-items: center;
  box-shadow: 0 2px 6px rgba(0,0,0,.3);
  pointer-events: none;
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
        lasso: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M4 7c0-2 3-4 8-4s8 2 8 4-3 4-8 4c-3 0-5 .7-5 2s2 2 5 2"/>
                    <path d="M9 15c0 2-1.5 5-4 5"/></svg>`,
        upload: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 3v12"/><path d="M8 7l4-4 4 4"/>
                    <rect x="4" y="15" width="16" height="6" rx="2"/></svg>`,
        download: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 3v12"/><path d="M16 11l-4 4-4-4"/>
                    <rect x="4" y="17" width="16" height="4" rx="2"/></svg>`,
        trash: `<svg viewBox="0 0 24 24" aria-hidden="true">
                <g stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <polygon points="6 4 18 4 22 12 18 20 6 20 2 12"/>
                <path d="M9 9l6 6M15 9l-6 6"/>
                </g>
            </svg>`,
        radius: `<svg viewBox="0 0 24 24" aria-hidden="true">
                    <circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="2"/></svg>`,
        simplify: `<svg viewBox="0 0 24 24" aria-hidden="true">
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

      const makeBtn = (id, title, svg, withBadge = false) => {
        const b = document.createElement('button');
        b.type = 'button'; b.className = 'icon-btn'; b.id = id; b.title = title;
        b.innerHTML = svg; bar.appendChild(b);
        if (withBadge) b.setAttribute('data-badge', '');
        return b;
      };

      // Tool buttons
            // const pointerBtn = makeBtn('pointerBtn', 'Select individual areas', svgs.search);

      const circleBtn = makeBtn('circleSelectBtn', 'Circle select (click map to select)', svgs.circle);
      // const lassoBtn  = makeBtn('lassoSelectBtn',  'Lasso select (drag to sketch polygon)', svgs.lasso);
      // const uploadBtn = makeBtn('uploadGeoBtn',    'Upload GeoJSON', svgs.upload);
      // const exportBtn = makeBtn('exportGeoBtn',    'Export drawn boundary (applies simplify/buffer)', svgs.download);
      // const clearBtn  = makeBtn('clearGeoBtn',     'Clear drawn boundary', svgs.trash);

      // Parameter buttons (same size with value badge)
      const radiusBtn = makeBtn('radiusBtn', 'Radius (km): click to cycle, Shift+Click to set', svgs.radius, true);
      // const simplifyBtn = makeBtn('simplifyBtn', 'Simplify tol (m): click to cycle, Shift+Click to set', svgs.simplify, true);
      // const bufferBtn   = makeBtn('bufferBtn',   'Buffer (m): positive grows, negative shrinks. Click to cycle, Shift+Click to set', svgs.buffer,   true);

      // Zoom controls
       const zoomOutBtn = makeBtn('dtZoomOut',  'Zoom out',  svgs.zoomOut);
       const zoomInBtn  = makeBtn('dtZoomIn',   'Zoom in',   svgs.zoomIn);
      const homeBtn    = makeBtn('dtZoomHome', 'Reset view', svgs.home);

      zoomOutBtn.style.display = 'none';
      zoomInBtn.style.display = 'none';
      homeBtn.style.display = 'none';
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
      //let radiusKm = 2, simplifyTolMeters = 10, bufferMeters = 0;

      const setBadge = (btn, txt) => btn.setAttribute('data-badge', txt);
      const refreshBadges = () => {
        setBadge(radiusBtn, `${radiusKm}k`);
        // setBadge(simplifyBtn, `${simplifyTolMeters}m`);
        // setBadge(bufferBtn,   `${bufferMeters}m`);
      };
      refreshBadges();

      const radiusInput = document.getElementById('radius-input');

if (radiusInput) {
  radiusInput.value = radiusKm; // sync initial value

  radiusInput.addEventListener('input', (e) => {
    const v = parseFloat(e.target.value);
    if (!isNaN(v) && v > 0) {
      radiusKm = v;

      // ✅ update button badge so UI stays in sync
      refreshBadges();
    }
  });
}
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
      function setCircleMode(on) {
        circleMode = !!on;
        circleBtn.classList.toggle('active', circleMode);
        drawToolActive = circleMode;

        const canvas = map.getCanvas();

          if (drawToolActive) {
            canvas.style.cursor = 'none';   // ✅ hide hand icon
          } else {
            canvas.style.cursor = '';       // ✅ restore default
            removeCirclePreview();          // ✅ clear circle
          }
        }


      function setActiveButton(activeBtn) {
        document.querySelectorAll('#draw-toolbar .icon-btn')
          .forEach(btn => btn.classList.remove('active'));
        activeBtn.classList.add('active');
      }

      circleBtn.addEventListener('click', () => {
        setCircleMode(true);
        setActiveButton(circleBtn);
      });
      
      // Parameter buttons (same size with value badge)
      // REMOVED: radiusBtn - radius selector button
      // Replace with pointer/cursor icon to deselect circle mode
      const pointerBtn = makeBtn('pointerBtn', 'Select individual areas', svgs.search);

    // Move buttons into their new positions (MG)
    document.getElementById('pointer-slot')?.appendChild(pointerBtn);
    document.getElementById('circle-slot')?.appendChild(circleSelectBtn);


    pointerBtn.addEventListener('click', () => {
      setCircleMode(false);

      // ✅ clear active from all buttons
      document.querySelectorAll('#draw-toolbar .icon-btn')
        .forEach(btn => btn.classList.remove('active'));

      // ✅ set this one as active
      pointerBtn.classList.add('active');

      clearDrawnCircle();

    });


      // Set default radius to 5km
      radiusBtn.style.display = 'none';
      pointerBtn.innerHTML = `
        <svg viewBox="0 0 24 24">
          <path d="M17.1668 11.1733C17.1668 12.5307 17.1668 9.81592 17.1668 11.1733ZM17.1668 11.1733C17.1668 12.5307 17.1668 13.8881 17.1668 13.8881M17.1668 11.1733C17.1668 9.81592 20.0001 9.81592 20.0001 11.1733C20.0001 12.5307 20.0001 12.8701 20.0001 18.2997C20.0001 23.7294 7.85591 23.7756 5.68023 18.2997C4.87315 16.2684 5.01308 16.7027 4.2713 14.941C3.52953 13.1794 5.97114 12.1286 6.9472 13.8881C7.92326 15.6477 8.66677 18.4383 8.66677 17.2817C8.66677 16.125 8.66677 12.5307 8.66677 11.1733C8.66677 9.81592 8.66677 5.37546 8.66677 4.01805C8.66677 2.66065 11.5001 2.66065 11.5001 4.01805M17.1668 11.1733C17.1668 9.81592 14.3239 9.81592 14.3334 11.1733M14.3334 11.1733C14.3334 9.81592 11.5001 9.81592 11.5001 11.1733C11.5001 11.4976 11.5001 3.66565 11.5001 4.01805M14.3334 11.1733C14.3334 11.4976 14.3334 10.8209 14.3334 11.1733ZM14.3334 11.1733C14.3477 13.2094 14.3334 13.8881 14.3334 13.8881M11.5001 13.8881C11.5001 13.8881 11.5001 7.41019 11.5001 4.01805"
            stroke="currentColor"
            stroke-width="1.75"
            stroke-linecap="round"/>
        </svg>
      `;

      // Circle select
      map.on('click', (e) => {
        if (!circleMode) return;
        const center = [e.lngLat.lng, e.lngLat.lat];
        const circle = turf.circle(center, Math.max(0.05, +radiusKm || 2), { steps: 128, units: 'kilometers' });
        
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
          try { map.fitBounds(turf.bbox(poly), { padding: 30, animate: true }); } catch { }
        }
        lasso = [];
        drawToolActive = circleMode || lassoMode;
        map.getCanvas().style.cursor = drawToolActive ? 'crosshair' : '';
      }
      map.on('mouseup', finishLasso);
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
              try { map.fitBounds(turf.bbox(merged), { padding: 30 }); } catch { }
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
       homeBtn.addEventListener('click',    () => {
           map.easeTo({ center: [-6.8, 54.65], zoom: 7.5, duration: 600 });
       });

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
              const visIds = new Set(map.queryRenderedFeatures({ layers: [fillLayer] }).map(f => f.id));
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
          const feats = map.queryRenderedFeatures({ layers: [fillLayer] }).filter(f => f.id === hit.id);
          if (feats[0]) {
            try {
              const gj = { type: 'Feature', geometry: feats[0].geometry, properties: {} };
              hit.bbox = turf.bbox(gj);
            } catch (_) { }
          }
        }

        if (hit.bbox) { try { map.fitBounds(hit.bbox, { padding: 40, duration: 600 }); } catch { } }
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

    // expose so postcode search can call it
    window.selectGeographyByCode = selectGeographyByCode;

    // ===== POSTCODE SEARCH EVENT LISTENER =====
    // Attach click handler to postcode search button (must be inside map.on('load') after updateTables is defined)
    const postcodeSearchBtn = document.getElementById('postcode-search-btn');
    if (postcodeSearchBtn) {
      postcodeSearchBtn.addEventListener('click', handlePostcodeSearch);
    }

    const postcodeInput = document.getElementById('postcode-input');
    if (postcodeInput) {
      postcodeInput.addEventListener('keypress', function(event) {
        if (event.key === 'Enter') {
          handlePostcodeSearch();
        }
      });
    }
    // ===== END POSTCODE SEARCH EVENT LISTENER =====

   });

  document.getElementById('zone-selector').addEventListener('change', () => {
    map.easeTo({
      center: [-6.8, 54.65],
      zoom: 7.5,
      duration: 1000
    });
  });

  // map reset zoom button
  document.getElementById('resetZoomBtn').addEventListener('click', () => {
    map.easeTo({
      center: [-6.8, 54.65],
      zoom: 7.5,
      duration: 2000
    });
  });

  // LGD buttons section toggle (collapse/expand)
  document.getElementById('lgd-toggle-btn')?.addEventListener('click', () => {
    const toggleBtn = document.getElementById('lgd-toggle-btn');
    const container = document.getElementById('lgd-buttons-container');
    const icon = toggleBtn.querySelector('.lgd-toggle-icon');
    
    const isExpanded = container.style.display === 'flex';
    container.style.display = isExpanded ? 'none' : 'flex';
    container.style.flexWrap = 'wrap';
    container.style.gap = '8px';
    container.style.paddingTop = '12px';
    const newExpanded = !isExpanded;
    toggleBtn.setAttribute('aria-expanded', newExpanded);
    icon.textContent = newExpanded ? '▼' : '▶';
  });

  let DEFAULT_ZOOM = null;

  map.on('load', () => {
    DEFAULT_ZOOM = map.getZoom();
    updateZoomDisplay();
    
    const nav = new maplibregl.NavigationControl({
    showZoom: true,
    showCompass: false
  });

  map.addControl(nav, 'bottom-right'); // temporary positio

  });

  function updateZoomDisplay() {
    if (!DEFAULT_ZOOM) return;

    const zoom = map.getZoom();

    // Use actual starting zoom as baseline
    const percent = Math.round(100 + (zoom - DEFAULT_ZOOM) * 100);

    document.getElementById("zoom-level").innerText = `Zoom: ${percent}%`;
  }

  map.on('zoom', updateZoomDisplay);
  map.on('load', updateZoomDisplay);

  class PercentZoomControl {
    onAdd(map) {
      
      this._map = map;
      const container = document.createElement('div');
      container.className = 'maplibregl-ctrl maplibregl-ctrl-group custom-zoom';

      const zoomInBtn = document.createElement('button');
      zoomInBtn.innerHTML = '+';
      zoomInBtn.title = 'Zoom in (10%)';

      const zoomOutBtn = document.createElement('button');
      zoomOutBtn.innerHTML = '−';
      zoomOutBtn.title = 'Zoom out (10%)';

      // 50 percentage points step (e.g., 100% -> 150% -> 200%)
      const PERCENT_STEP = 100;

      function percentToZoom(percent) {
        if (typeof DEFAULT_ZOOM !== 'number') return map.getZoom();
        return DEFAULT_ZOOM + Math.log2(percent / 100);
      }

      function zoomToPercent(percent) {
        const z = percentToZoom(percent);
        map.easeTo({ zoom: z });
      }

      zoomInBtn.onclick = () => {
        // compute current percent relative to DEFAULT_ZOOM, then add step
        const currentZoom = map.getZoom();
        const currentPercent = Math.pow(2, currentZoom - DEFAULT_ZOOM) * 100;
        const newPercent = Math.round(currentPercent + PERCENT_STEP);
        zoomToPercent(newPercent);
      };

      zoomOutBtn.onclick = () => {
        const currentZoom = map.getZoom();
        const currentPercent = Math.pow(2, currentZoom - DEFAULT_ZOOM) * 100;
        const newPercent = Math.round(currentPercent - PERCENT_STEP);
        zoomToPercent(newPercent);
      };

      container.appendChild(zoomInBtn);
      container.appendChild(zoomOutBtn);

      // If there's a clear selections button in the page, place the zoom buttons just to its left
      // so they appear next to the clear selections control rather than in the map controls.
      try {
        const clearBtn = document.getElementById('clear-selection-btn');
        if (clearBtn && clearBtn.parentNode) {
          // ensure container has similar inline styles
          container.style.display = 'inline-flex';
          container.style.gap = '6px';
          container.style.marginRight = '8px';
          // ensure it appears above other controls (e.g., about button)
          container.style.position = 'relative';
          container.style.zIndex = '1000';
          clearBtn.parentNode.insertBefore(container, clearBtn);
          return container;
        }
      } catch (e) { /* ignore and fall back to default insertion */ }

      return container;
      
    }

    onRemove() {
      this._map = undefined;
    }
    
  }

  async function updateSourceLink() {
    const zoneType = window.selectedZoneType || 'sdz';
    const selectedLabels = window.chosenCategories || ['Age (4 Categories)', 'Sex Label'];
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

      window.urlZoneType = zoneType;
      window.urlcategoryCode = categoryCode;
      window.urllabel = label;
      window.urlsource = source;

      if (source === "Flexible Table Builder" && categoryCode) {
        if (zoneType === "dz" || zoneType === "sdz") {
                fullUrl = `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}21&v=${categoryCode}`;
              } else if (zoneType === "dea" || zoneType === "lgd") {
                fullUrl = `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}14&v=${categoryCode}`;
              } 
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

      if (!src) return;
    });
  }

  let latestAggregatedData = {};
  let selectedCategories = ['Age (4 Categories)', 'Sex Label'];
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

    // Clear selections when switching zones
    selectedIds.clear();
    lgdNameToId.clear();

    // Clear all feature states to prevent persistence across zone switches
    map.removeFeatureState({ source: 'sdz2021', sourceLayer: 'SDZ2021_clipped' });
    map.removeFeatureState({ source: 'dz2021', sourceLayer: 'DZ2021_clipped' });
    map.removeFeatureState({ source: 'dea2014', sourceLayer: 'DEA2014_clipped' });
    map.removeFeatureState({ source: 'lgd2014', sourceLayer: 'LGD2014_clipped' });

    // Update map layers
    vis('sdz-fill', currentZoneType === 'sdz'); 
    vis('dz-fill', currentZoneType === 'dz'); 
    vis('dea-fill', currentZoneType === 'dea'); 
    vis('lgd-fill', currentZoneType === 'lgd');

    vis('sdz-outline-default', currentZoneType === 'sdz'); 
    vis('dz-outline-default', currentZoneType === 'dz'); 
    vis('dea-outline-default', currentZoneType === 'dea'); 
    vis('lgd-outline-default', currentZoneType === 'lgd');

    vis('sdz-outline-hover', currentZoneType === 'sdz'); 
    vis('dz-outline-hover', currentZoneType === 'dz'); 
    vis('dea-outline-hover', currentZoneType === 'dea'); 
    vis('lgd-outline-hover', currentZoneType === 'lgd');

    populateLGDButtons();
    syncPreviewVisibility();
    updateSummaryPreview();
    ensureSummaryHero();

    updateTables([]);
  });

  const selector = document.getElementById("zone-selector");

  // Set default value
  selector.value = "sdz";

  // Listen for changes
  selector.addEventListener("change", function () {
    currentZoneType = this.value;

    // Clear selections when switching zones
    selectedIds.clear();
    lgdNameToId.clear();

    // Clear all feature states to prevent persistence across zone switches
    map.removeFeatureState({ source: 'sdz2021', sourceLayer: 'SDZ2021_clipped' });
    map.removeFeatureState({ source: 'dz2021', sourceLayer: 'DZ2021_clipped' });
    map.removeFeatureState({ source: 'dea2014', sourceLayer: 'DEA2014_clipped' });
    map.removeFeatureState({ source: 'lgd2014', sourceLayer: 'LGD2014_clipped' });

    // Update map layers
    vis('sdz-fill', currentZoneType === 'sdz'); 
    vis('dz-fill', currentZoneType === 'dz'); 
    vis('dea-fill', currentZoneType === 'dea'); 
    vis('lgd-fill', currentZoneType === 'lgd');

    vis('sdz-outline-default', currentZoneType === 'sdz'); 
    vis('dz-outline-default', currentZoneType === 'dz'); 
    vis('dea-outline-default', currentZoneType === 'dea'); 
    vis('lgd-outline-default', currentZoneType === 'lgd');

    vis('sdz-outline-hover', currentZoneType === 'sdz'); 
    vis('dz-outline-hover', currentZoneType === 'dz'); 
    vis('dea-outline-hover', currentZoneType === 'dea'); 
    vis('lgd-outline-hover', currentZoneType === 'lgd');

    populateLGDButtons();
    syncPreviewVisibility();
    updateSummaryPreview();
    ensureSummaryHero();

    updateTables([]);
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

  /**
   * Select a geography on the map by its code
   * This function integrates with the existing map selection logic
   * Parameters:
   *   - code: the geographic code (e.g., "N20000001" for a data zone)
   *   - zoneType: 'sdz', 'dz', or 'dea'
   * 
   * NOTE: This function is defined inside map.on('load') to access updateTables and other local functions
   */
  function selectGeographyByCode(code, zoneType) {
    // Ensure map is ready
    if (!map) return;

    // Ensure the correct zone type is selected
    const zoneSelector = document.getElementById('zone-selector');
    const zoneNeedsChange = zoneSelector && zoneSelector.value !== zoneType;
    
    if (zoneNeedsChange) {
      zoneSelector.value = zoneType;
      // Trigger the zone change event
      zoneSelector.dispatchEvent(new Event('change'));
    }

    // Wait for zone change to complete, then select the geography
    // Use a longer timeout to allow clearSelections animation and data re-indexing to finish
    const delayMs = zoneNeedsChange ? 500 : 250;
    
    setTimeout(() => {
      const { source, sourceLayer } = getZoneIdsFor(zoneType);
      const geoData = getDataSourceFor(zoneType);
      
      // Verify data is available
      if (!geoData || typeof geoData !== 'object' || !Object.keys(geoData).length) {
        console.warn(`No data available for zone type: ${zoneType}`);
        return;
      }
      
      // Try to find the code - handle both string and numeric ID types
      let lookupCode = code;

      if (!geoData[code]) {
        // Try converting to number if it wasn't found  
        const numCode = !isNaN(code) ? +code : code;
        if (geoData[numCode]) {
          lookupCode = numCode;
        } else {
          // Also try string version if we had a number
          lookupCode = String(code);
          if (zoneType !== 'lgd') {
            if (!geoData[lookupCode]) {
              console.error(`❌ Geography code "${code}" not found in ${zoneType} data.`);
              console.log(`Available data keys sample:`, Object.keys(geoData).slice(0, 10));
              console.log(`Searched for: "${code}", as number: ${numCode}, as string: "${lookupCode}"`);
              return;
            }
         }
        }
      }

      console.log(`✓ Code lookup successful. Found: "${lookupCode}" in ${zoneType} data`);

      // Add to selection and highlight on map
      if (zoneType === 'lgd') {
        
        // For LGD, lookupCode is the name, need to find the feature id
        const { source, sourceLayer } = getZoneIdsFor(zoneType);

        const features = map.querySourceFeatures(source, { sourceLayer });

          const feature = features.find(f =>
          String(f.properties?.LGD_CODE) === String(lookupCode) ||
          String(f.properties?.LGD2014) === String(lookupCode) ||
          String(f.id) === String(lookupCode)
        );
        
        if (feature) {
          // const featureId = feature.id;
          const featureId = feature.properties.LGD_CODE || feature.properties.lgd_code;

          // ✅ Extract the LGD NAME from the feature
          const lgdName =
            feature.properties?.LGDNAME ||
            feature.properties?.LGD2014NAME ||
            feature.properties?.lgd_name;

          // ✅ Store NAME for data (this fixes your population issue)
          if (lgdName) {
            selectedIds.add(lgdName);
          } else {
            console.warn('⚠️ LGD name not found on feature');
            // selectedIds.add(lookupCode); // fallback (optional)
            selectedIds.add(feature.properties.LGDNAME);
          }

          // ✅ Keep mapping (optional but useful)
          lgdNameToId.set(lgdName, featureId);

          // ✅ Highlight using feature ID (this fixes your map issue)
          map.setFeatureState(
            { source, sourceLayer, id: featureId },
            { hovered: true }
          );
        } else {
          console.error(`❌ Feature not found for LGD "${lookupCode}"`);
          return;
        }

      } else {
        selectedIds.add(lookupCode);
        map.setFeatureState(
          { source, sourceLayer, id: lookupCode },
          { hovered: true }
        );
      }

      // Display charts tab by default (before updateTables, same as manual map clicks)
      let selectedTab = document.querySelector('.view-tab.selected');
      if (!selectedTab) {
        const chartsTab = document.querySelector('.view-tab[data-view="charts"]');
        if (chartsTab) chartsTab.classList.add("selected");
      }

      document.getElementById("charts-container").style.display = "flex";
      document.getElementById("tables-container").style.display = "none";
      document.getElementById("urban-rural-comparison").style.display = "none";
      document.getElementById("urban-rural-charts").style.display = "none";

      // Update all UI elements
      window.selectedIdsExcel = selectedIds;
      updateTables(Array.from(selectedIds));
      renderZoneBreakdownTable(Array.from(selectedIds));
      updateCtaEnabled();
      updateSummaryPreview();
      ensureSummaryHero();
    }, delayMs);
  }

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
      currentZoneType === 'dz' ? dzData :
        currentZoneType === 'dea' ? deaData :
          currentZoneType === 'lgd' ? lgdData :
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

    document.querySelectorAll(".total-population").forEach(elem => {
      elem.textContent = totalPopulation.toLocaleString();
    });
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
    titleEl.setAttribute("lang", "en-GB");
    titleEl.setAttribute("spellcheck", "true");

    const summaryList = document.getElementById("summaryList");
    const populationEls = document.querySelectorAll(".total-population");

    if (!selectedIdsArray.length) {
      container.style.display = "none";
      summaryList.innerHTML = "";
      populationEls.forEach(el => {
        el.textContent = "0";
      });
      window.areaProfileTitle = undefined;
      window.lastSelectionHash = undefined;
      return;
    }
    container.style.display = "block";
    summaryList.innerHTML = "";

    const dataSource =
      currentZoneType === 'dz' ? dzData :
        currentZoneType === 'dea' ? deaData :
          currentZoneType === 'lgd' ? lgdData :
            sdzData;

    const labelKey =
      currentZoneType === 'dz' ? "Census 2021 Data Zone Label" :
        currentZoneType === 'dea' ? "District Electoral Area 2014 Label" :
          currentZoneType === 'lgd' ? "Local Government District 2021 Label" :
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

    const placeholderText = "Click to give your area a name";
    const savedTitle = window.areaProfileTitle?.trim();
    const applyTitleBtn = document.getElementById('apply-area-name');

    titleEl.textContent = savedTitle || placeholderText;
 


    function showApplyButton() {
      if (!applyTitleBtn) return;
      applyTitleBtn.style.display = 'inline-flex';
      requestAnimationFrame(() => applyTitleBtn.classList.add('visible'));
    }

    function hideApplyButton() {
      if (!applyTitleBtn) return;
      applyTitleBtn.classList.remove('visible');
      const onTransitionEnd = (event) => {
        if (event.propertyName === 'opacity') {
          applyTitleBtn.style.display = 'none';
          applyTitleBtn.removeEventListener('transitionend', onTransitionEnd);
        }
      };
      applyTitleBtn.addEventListener('transitionend', onTransitionEnd);
    }

    function applyAreaName() {
      const trimmedTitle = titleEl.textContent.trim();
      window.areaProfileTitle = trimmedTitle || undefined;
      titleEl.textContent = trimmedTitle || placeholderText;
      hideApplyButton();
      if (window.latestAggregatedData) {
        renderAggregatedCharts(window.latestAggregatedData, selectedCategories);
        renderAggregatedTables(window.latestAggregatedData, selectedCategories);
      }
      if (typeof updateSummaryPreview === 'function') {
        updateSummaryPreview();
      }
    }

    if (!titleEl.dataset.areaProfileTitleListenersAttached) {
      titleEl.addEventListener('focus', () => {
        if (!window.areaProfileTitle && titleEl.textContent.trim() === placeholderText) {
          titleEl.textContent = '';
        }
      });

      titleEl.addEventListener('input', () => {
        const isEditingPlaceholder = !window.areaProfileTitle;
        const hasText = titleEl.textContent.trim() !== '';
        if (isEditingPlaceholder && hasText) {
          showApplyButton();
        } else {
          hideApplyButton();
        }
      });

      titleEl.addEventListener('blur', () => {
        window.areaProfileTitle = titleEl.textContent.trim();
        if (!window.areaProfileTitle) {
          titleEl.textContent = placeholderText;
          hideApplyButton();
        }
      });

      titleEl.dataset.areaProfileTitleListenersAttached = 'true';
    }

    if (applyTitleBtn && !applyTitleBtn.dataset.applyListenerAttached) {
      applyTitleBtn.addEventListener('click', (event) => {
        event.preventDefault();
        applyAreaName();
      });
      applyTitleBtn.dataset.applyListenerAttached = 'true';
    }

    const sortedLGDs = Object.keys(lgdStats).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' })
    );

    let totalZonesSelected = 0;
    sortedLGDs.forEach(lgd => {
      const stats = lgdStats[lgd];
      const totalInLGD = lgdTotals[lgd] || stats.total;
      totalZonesSelected += stats.total;

      const li = document.createElement("li");
      li.innerHTML = `${lgd}: <strong>${stats.total} of ${totalInLGD}</strong> zones selected`;
      summaryList.appendChild(li);
    });

    window.totalZonesSelected = totalZonesSelected;

    populationEls.forEach(el => {
      el.textContent = totalPopulation.toLocaleString();
    });

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

  function getAreaLabel() {
    const saved = window.areaProfileTitle?.trim();
    return saved ? saved : "Your Area";
  }

  function renderAggregatedTables(aggregatedData, selectedCategories = []) {
    if (!aggregatedData || Object.keys(aggregatedData).length === 0) return;

    if (selectedCategories.length > 0) {
      const valid = selectedCategories.filter(cat => Object.keys(aggregatedData).includes(cat));
      if (valid.length === 0) return;
    }

    const container = document.getElementById("tables-container");
    container.innerHTML = "";

    const grid = document.createElement("div");
    grid.className = "table-grid";

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrapper";

    const entries = Object.entries(aggregatedData).filter(([key]) =>
      selectedCategories.length === 0 || selectedCategories.includes(key)
    );

    entries.forEach(([category, values]) => {
      const wrapper = document.createElement("div");
      wrapper.className = "table-wrapper";
      wrapper.style.background = "#fff";
      wrapper.style.boxShadow = "0 2px 4px rgba(0,0,0,0.1)";
      wrapper.style.padding = "16px";
      wrapper.style.borderRadius = "4px";
      wrapper.style.boxSizing = "border-box";

      const title = document.createElement("h3");
      title.textContent = category.replace(/ Label$/, "");;

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

      wrapper.appendChild(title);

      const table = document.createElement("table");
      table.className = "data-table";

      const areaLabelHeader = getAreaLabel();
      const thead = document.createElement("thead");
      const headerRow = document.createElement("tr");
      ["Category", "Count", areaLabelHeader, "NI"].forEach((text, index) => {
        const th = document.createElement("th");
        th.textContent = text;
        th.style.textAlign = index === 0 ? "left" : "right";
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
        tdCount.textContent = Number.isFinite(count) ? count.toLocaleString() : count;
        tdCount.style.textAlign = "right";
        row.appendChild(tdCount);

        const tdPercentage = document.createElement("td");
        tdPercentage.textContent = totalCount > 0 ? ((count / totalCount) * 100).toFixed(1) + "%" : "0%";
        tdPercentage.style.textAlign = "right";
        row.appendChild(tdPercentage);

        const tdNI = document.createElement("td");
        const niVal = niTotals[category]?.[label];
        tdNI.textContent = typeof niVal === "number" ? niVal.toFixed(1) + "%" : "–";
        tdNI.style.textAlign = "right";
        row.appendChild(tdNI);

        tbody.appendChild(row);
      }

      const totalRow = document.createElement("tr");
      totalRow.style.fontWeight = "bold";

      const totalLabel = document.createElement("td");
      totalLabel.textContent = "Total";
      totalRow.appendChild(totalLabel);

      const totalValue = document.createElement("td");
      totalValue.textContent = totalCount.toLocaleString();
      totalValue.style.textAlign = "right";
      totalRow.appendChild(totalValue);

      const totalPct = document.createElement("td");
      totalPct.textContent = totalCount > 0 ? "100%" : "0%";
      totalPct.style.textAlign = "right";
      totalRow.appendChild(totalPct);

      const totalNI = document.createElement("td");
      totalNI.textContent = "–";
      totalNI.style.textAlign = "right";
      totalRow.appendChild(totalNI);

      tbody.appendChild(totalRow);

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
  let labelToYear = {};


  async function loadLookupData() {
    const response = await fetch('category_lookup.json');
    const lookup = await response.json();

    lookup.forEach(item => {
      labelToCode[item.nested_list_names] = item.further_breakdown_df;
      labelToSource[item.nested_list_names] = item.Source;
      labelToYear[item.nested_list_names] = item.Year;
    });

    window.labelToCode = labelToCode;
    window.labelToSource = labelToSource;
    window.labelToYear = labelToYear;
  }

  loadLookupData().then(() => {
    renderAggregatedCharts(latestAggregatedData, selectedCategories);
  });

  function getCategoryURL(label, zoneType = 'sdz') {

    const source = window.labelToSource?.[label];
    const categoryCode = window.labelToCode?.[label];

    if (source === "Flexible Table Builder" && categoryCode) {
      if (zoneType === "dz" || zoneType === "sdz") {
          return `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}21&v=${categoryCode}`;
        } else if (zoneType === "dea" || zoneType === "lgd") {
          return `https://build.nisra.gov.uk/en/custom/data?d=PEOPLE&v=${zoneType}14&v=${categoryCode}`;
        } 
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

    window.chartInstances?.forEach(c => c.destroy());
    window.chartInstances = [];

    const FONT = "14px sans-serif";
    const LINE_HEIGHT = 10;
    const BAR_HEIGHT = 32;
    const BAR_SPACING = 4;
    const LABEL_BLOCK_HEIGHT = LINE_HEIGHT * 2 + 4;
    const CHART_TOP_PADDING = 15;
    const LABEL_TO_BAR_GAP = 1;
   
    const grid = document.createElement("div");
    grid.className = "charts-grid";
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
        borderRadius: "4px",
        boxSizing: "border-box",
        width: "100%"
      });

      // --- Title and Expand Button ---
      const titleWrapper = document.createElement("div");
      titleWrapper.style.display = "flex";
      titleWrapper.style.alignItems = "center";
      titleWrapper.style.justifyContent = "space-between";

      const title = document.createElement("h3");

      let titleText = category.replace(/ Label$/, "");
      titleText = titleText.replace(/\s*\(MYE\)\s*/i, titleText.includes("Age") ? " (4 Categories)" : "");
      title.textContent = titleText.trim();
      title.style.margin = "0";

      const expandBtn = document.createElement("button");
      expandBtn.className = "expand-toggle";
      expandBtn.setAttribute("aria-expanded", "false");
      expandBtn.style.border = "none";
      expandBtn.style.background = "none";
      expandBtn.style.padding = "0";
      expandBtn.style.marginLeft = "8px";
      expandBtn.style.cursor = "pointer";
      expandBtn.style.boxShadow = "none";

      const infoImg = document.createElement("img");
      infoImg.src = "img/i-button.svg";
      infoImg.alt = "Show more information";
      infoImg.style.width = "20px";
      infoImg.style.height = "20px";
      infoImg.style.display = "block";
      expandBtn.appendChild(infoImg);

      const expandableContent = document.createElement("div");
      expandableContent.className = "expandable-content";
      expandableContent.style.display = "none";

      const about = document.createElement("strong");
      about.textContent = "Access data at:";
      expandableContent.appendChild(about);
      expandableContent.appendChild(document.createElement("br"));

      const url = getCategoryURL(category, window.selectedZoneType || 'sdz');
      if (url) {
        const link = document.createElement("a");
        link.href = url;
        link.target = "_blank";
        link.textContent = "Source";
        expandableContent.appendChild(link);
      } else {
        expandableContent.appendChild(document.createTextNode("No additional information available."));
      }

      expandBtn.addEventListener("click", () => {
        const expanded = expandBtn.getAttribute("aria-expanded") === "true";
        expandBtn.setAttribute("aria-expanded", String(!expanded));
        expandableContent.style.display = expanded ? "none" : "block";
      });

      const titleContent = document.createElement("div");
      titleContent.style.display = "flex";
      titleContent.style.alignItems = "center";
      titleContent.style.gap = "8px";
      titleContent.appendChild(title);
      titleContent.appendChild(expandBtn);

      titleWrapper.appendChild(titleContent);
      wrapper.appendChild(titleWrapper);
      wrapper.appendChild(expandableContent);

      // Add subheading
      const subheading = document.createElement("h4");
      subheading.className = "subheading";
      subheading.style.margin = "8px 0 0 0";
      subheading.style.fontWeight = "normal";


      const src = window.labelToSource?.[category];
      const year = window.labelToYear?.[category];
      if (src === "Flexible Table Builder") {
        subheading.textContent = `Census ${year}`;
      } else if (category === "Benefits Statistics") {
        subheading.textContent = `placeholder`;
      } else {
        subheading.textContent = `Mid year population estimates ${year}`;
      }

      wrapper.appendChild(subheading);

      const areaLabel = getAreaLabel();

      // --- Legend ---
      const legendEl = document.createElement("div");
      Object.assign(legendEl.style, {
        display: "flex",
        justifyContent: "center",
        gap: "1rem",
        alignItems: "center",
        marginTop: "12px",
        marginBottom: "8px"
      });

      const datasets = [
        { label: areaLabel, backgroundColor: "#3878c5", type: "bar" },
        { label: "NI", borderColor: "#222", type: "line" }
      ];

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

      // --- Chart canvas ---
      const labels = Object.keys(values);
      const total = Object.values(values).reduce((a, v) => a + v, 0);
      const barPercents = labels.map(l => total > 0 ? +((values[l] || 0) / total * 100).toFixed(1) : 0);
      const chartDatasets = [
        { label: areaLabel, data: barPercents, backgroundColor: "#3878c5", barThickness: BAR_HEIGHT },
        { label: "NI", data: [], type: "line", borderColor: "#222", borderWidth: 2, fill: false, pointRadius: 0 }
      ];

      const canvasHeight = labels.length * (LABEL_BLOCK_HEIGHT + BAR_HEIGHT + BAR_SPACING + 26) + CHART_TOP_PADDING;
      const canvas = document.createElement("canvas");
      Object.assign(canvas.style, { display: "block", width: "100%", maxHeight: "none" });
      canvas.height = canvasHeight;
      wrapper.appendChild(canvas);

      const spacer = document.createElement("div");
      spacer.style.flex = "1";
      wrapper.appendChild(spacer);
      grid.appendChild(wrapper);

      whenVisible(wrapper, () => {
        const drawWidth = Math.max(0, wrapper.clientWidth - 32);
        canvas.width = drawWidth;
        canvas.style.width = `${drawWidth}px`;

        const niMap =
          (typeof window !== "undefined" && window.niTotals && window.niTotals[category])
          || (typeof niTotals !== "undefined" && niTotals[category])
          || {};
        const niValues = labels.map(l => (typeof niMap[l] === "number" ? niMap[l] : null));
        const rawMax = Math.max(
          ...chartDatasets[0].data,
          ...niValues.filter(v => typeof v === "number")
        ) * 1.05;
        const cappedMax = Number.isFinite(rawMax) ? Math.min(rawMax, 100) : 100;

        const chart = new Chart(canvas, {
          type: "bar",
          data: { labels, datasets: chartDatasets },
          options: {
            indexAxis: "y",
            responsive: false,
            maintainAspectRatio: false,
            animation: false,
            layout: { padding: { top: CHART_TOP_PADDING, left: 10, right: 10, bottom: 0 } },
            plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.raw}%` } }
            },
            scales: {
              x: {
                beginAtZero: true,
                suggestedMax: cappedMax,
                title: { display: true, text: "Percentage" },
                ticks: { callback: v => `${v}%` }
              },
              y: { ticks: { display: false }, grid: { display: false }, offset: true }
            }
          },
          plugins: [
            {
              id: "aboveGroupLabels",
              afterDatasetsDraw(chartInst) {
                const ctx = chartInst.ctx;
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
                    .sort((a, b) => a.y - b.y);

                  if (!bars.length) return;

                  const topBar = bars[0];
                  const topY = topBar.y - topBar.height / 2;
                  const labelY = topY - LABEL_BLOCK_HEIGHT - 2;

                  // Build combined single-line text
                  const breakdown = chartInst.data.datasets
                    .filter(ds => ds.type !== "line")
                    .map(ds => `${ds.data[i]}%`)
                    .join(" | ");
                  const niVal = niValues[i];
                  const niText = typeof niVal === "number" ? ` (NI ${niVal.toFixed(1)}%)` : "";
                  const fullText = `${label} ${breakdown}${niText}`;

                  // Available width for the label text
                  const maxAllowed = chartInst.chartArea.right - xStart - 30;

                  // Decide whether to truncate
                  let textToDraw = fullText;
                  if (ctx.measureText(fullText).width > maxAllowed) {
                    // simple truncation to fit, preserving word chars until width ok
                    let trunc = fullText;
                    while (trunc.length > 0 && ctx.measureText(trunc + "...").width > maxAllowed) {
                      trunc = trunc.slice(0, -1);
                    }
                    textToDraw = trunc + "...";

                    // info overlay to show full text on click
                    const infoBtn = document.createElement("div");
                    infoBtn.className = "label-overlay";
                    infoBtn.textContent = "ⓘ";
                    infoBtn.style.position = "absolute";

                    const chartRect = chartInst.canvas.getBoundingClientRect();
                    const contRect = card.getBoundingClientRect();
                    const labelW = ctx.measureText(textToDraw).width;
                    const topOff = labelY + chartRect.top - contRect.top + LINE_HEIGHT / 2 - 10;
                    const leftOff = chartInst.canvas.offsetLeft + xStart + labelW + 6;

                    Object.assign(infoBtn.style, {
                      top: `${topOff}px`,
                      left: `${leftOff}px`,
                      cursor: "pointer",
                      color: "#0074D9"
                    });
                    infoBtn.title = fullText;
                    infoBtn.addEventListener("click", () => alert(fullText));
                    card.appendChild(infoBtn);
                  }

                  // Draw the single-line label (vertically positioned to avoid overlapping bars)
                  ctx.fillText(textToDraw, xStart, labelY + (LABEL_BLOCK_HEIGHT - LINE_HEIGHT) / 2);
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

                chartInst.data.labels.forEach((_, i) => {
                  const val = niValues[i];
                  if (typeof val !== "number") return;

                  const x = xScale.getPixelForValue(val);

                  const bars = chartInst.data.datasets
                    .map((ds, idx) => ({ ds, idx }))
                    .filter(o => o.ds.type !== "line")
                    .map(o => chartInst.getDatasetMeta(o.idx).data[i])
                    .sort((a, b) => a.y - b.y);

                  if (!bars.length) return;

                  const yTop = bars[0].y - bars[0].height / 2;
                  const yBottom = bars[bars.length - 1].y + bars[bars.length - 1].height / 2;

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
    if (currentZoneType === 'sdz') {
      map.removeFeatureState({ source: 'sdz2021', sourceLayer: 'SDZ2021_clipped' });
    } else if (currentZoneType === 'dz') {
      map.removeFeatureState({ source: 'dz2021', sourceLayer: 'DZ2021_clipped' });
    } else if (currentZoneType === 'dea') {
      map.removeFeatureState({ source: 'dea2014', sourceLayer: 'DEA2014_clipped' });
    } else if (currentZoneType === 'lgd') {
      map.removeFeatureState({ source: 'lgd2014', sourceLayer: 'LGD2014_clipped' });
    }

    selectedIds.clear();
    lgdNameToId.clear();
    selectedLGDs.clear();
    // Reset LGD checkboxes
    document.querySelectorAll('#lgd-buttons input[type="checkbox"]').forEach(cb => {
      cb.checked = false;
    });
    document.querySelectorAll('.lgd-btn').forEach(label => {
      label.classList.remove('selected');
    });
    popup.remove();

    // Reset UI
    document.getElementById("tables-container").innerHTML = "";
    document.getElementById("breakdown-container").style.display = "none";
    document.getElementById("urban-rural-comparison").style.display = "none";
    document.getElementById("urban-rural-charts").style.display = "none";
    document.getElementById("tables-container").style.display = "none";

    document.querySelectorAll('#lgd-buttons input[type="checkbox"]').forEach(cb => {
      cb.checked = false;
    });

    document.querySelectorAll('.lgd-btn').forEach(label => {
      label.classList.remove('selected');
    });

    document.querySelectorAll(".total-population").forEach(elem => {
      elem.textContent = "0";
    });

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
      window.chartInstances = [];
    }

    // Optionally clear chart container
    const container = document.getElementById("charts-container");
    if (container) {
      container.innerHTML = "";
    }
    const out = document.getElementById("output-content");
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

    // ✅ Clear selected areas
    clearSelections();

    // ✅ Clear drawn circle
    map.getSource('draw-geom').setData({ type: 'FeatureCollection', features: [] });
    lastDrawnFeature = null;

    // ✅ Reset map view
    map.easeTo({
      center: [-6.8, 54.65],
      zoom: 7.5,
      duration: 1000
    });

    // ✅ NEW: Clear postcode input
      const input = document.getElementById('postcode-input');
      if (input) input.value = '';

      // ✅ NEW: Clear postcode result/status
      const status = document.getElementById('postcode-status');
      if (status) {
        status.textContent = '';
        status.className = 'postcode-status';
      }

      // ✅ NEW (optional): Clear preview circle if visible
      if (typeof removeCirclePreview === 'function') {
        removeCirclePreview();
      }

      // ✅ NEW: Reset radius to default (2)
      const radiusInput = document.getElementById('radius-input');
      if (radiusInput) {
        radiusInput.value = 2;      // update UI
        radiusKm = 2;               // update your JS variable
      }

  });

  // FUNCTION: Build and populate LGD checkboxes from current zone data
  function populateLGDButtons() {
    const container = document.getElementById('lgd-buttons');
    if (!container) return;

    // Clear existing buttons
    container.innerHTML = '';

    // Get current zone data
    const dataSource =
      activeZone === 'dz' ? dzData :
        activeZone === 'dea' ? deaData :
          activeZone === 'lgd' ? lgdData :
            sdzData;

    // Extract all unique LGDs from the data
    const lgds = new Set();
    Object.values(dataSource || {}).forEach(record => {
      if (record?.LGD) {
        lgds.add(record.LGD);
      }
    });

    // Sort LGDs alphabetically
    const sortedLGDs = Array.from(lgds).sort();

    // Create checkbox and label for each LGD
    sortedLGDs.forEach(lgdCode => {
      const checkboxId = `lgd-${lgdCode.replace(/\s+/g, '-').toLowerCase()}`;

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.id = checkboxId;
      checkbox.value = lgdCode;
      checkbox.className = 'lgd-checkbox';

      const label = document.createElement('label');
      label.htmlFor = checkboxId;
      label.textContent = lgdCode;
      label.className = 'lgd-btn';

      // Add checkbox to label
      label.insertBefore(checkbox, label.firstChild);

      // Add to container
      container.appendChild(label);

      // ATTACH EVENT LISTENER TO CHECKBOX (moved here so it's attached when checkboxes are created)
      checkbox.addEventListener('change', () => {
        const lgdCode = checkbox.value;
        const label = document.querySelector(`label[for="${checkbox.id}"]`);

        const zone = currentZoneType;

        const dataSource =
          zone === 'dz' ? dzData :
            zone === 'dea' ? deaData :
              zone === 'lgd' ? lgdData :
                sdzData;

        const source =
          zone === 'dz' ? 'dz2021' :
            zone === 'dea' ? 'dea2014' :
              zone === 'lgd' ? 'lgd2014' :
              'sdz2021';

        const sourceLayer =
          zone === 'dz' ? 'DZ2021_clipped' :
            zone === 'dea' ? 'DEA2014_clipped' :
              zone === 'lgd' ? 'LGD2014_clipped' :
                'SDZ2021_clipped';

        if (checkbox.checked) {
          selectedLGDs.add(lgdCode);
          label?.classList.add('selected');

          const newIds = Object.entries(dataSource)
            .filter(([, rec]) => rec?.LGD === lgdCode)
            .map(([id]) => id);

          // newIds.forEach(id => {
          //   if (!selectedIds.has(id)) {
          //     selectedIds.add(id);
          //     map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
          //   }
          // });

            if (zone === 'lgd') {

              // ✅ Find the map feature by name
              const features = map.querySourceFeatures(source, { sourceLayer });

              const feature = features.find(f =>
                f.properties.LGDNAME === lgdCode ||
                f.properties.LGD2014NAME === lgdCode ||
                f.properties.lgd_name === lgdCode
              );

              if (feature) {
                const featureId = feature.id;

                selectedIds.add(lgdCode); // ✅ keep name for data

                map.setFeatureState(
                  { source, sourceLayer, id: featureId },
                  { hovered: true }
                );

              } else {
                console.warn('❌ LGD feature not found for:', lgdCode);
              }


            } else {

              // ✅ existing behaviour for SDZ/DZ/DEA
              newIds.forEach(id => {
                if (!selectedIds.has(id)) {
                  selectedIds.add(id);
                  map.setFeatureState({ source, sourceLayer, id }, { hovered: true });
                }
              });

            }

        } else {
          selectedLGDs.delete(lgdCode);
          label?.classList.remove('selected');

          const toRemove = Array.from(selectedIds).filter(id => dataSource[id]?.LGD === lgdCode);
          toRemove.forEach(id => {
            selectedIds.delete(id);
            map.setFeatureState({ source, sourceLayer, id }, { hovered: false });
          });
        }

        // Update outputs
        let selectedTab = document.querySelector('.view-tab.selected');
        if (!selectedTab) {
          const chartsTab = document.querySelector('.view-tab[data-view="charts"]');
          if (chartsTab) chartsTab.classList.add("selected");
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
    });
  }

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
      wrapper.style.borderRadius = "4px";
      wrapper.style.boxShadow = "0 1px 3px rgba(0,0,0,0.1)";
      wrapper.style.boxSizing = "border-box";
      wrapper.style.alignSelf = "flex-start"; // Ensures independent height

      const title = document.createElement("h3");
      title.textContent = `${category.replace(/ Label$/, "")} – Urban/Rural Comparison`;
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
          urbanPct: urbanTotal ? ((urbanCount / urbanTotal) * 100).toFixed(1) + "%" : "–",
          ruralPct: ruralTotal ? ((ruralCount / ruralTotal) * 100).toFixed(1) + "%" : "–",
          mixedPct: mixedTotal ? ((mixedCount / mixedTotal) * 100).toFixed(1) + "%" : "–",
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
      window.chartInstances.forEach(c => { try { c.destroy(); } catch (_) { } });
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
       gridTemplateColumns: window.innerWidth < 768 ? "1fr" : "repeat(2, 1fr)",
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
        borderRadius: "4px",
        boxSizing: "border-box",
        width: "100%"
      });

      // Title
      const title = document.createElement("h3");
      title.textContent = `${category.replace(/ Label$/, "")} – Urban/Rural/Mixed`;
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
      const barDatasets = ["Urban", "Rural", "Mixed"]
        .filter(g => groups[g]?.[category])
        .map(g => {
          const vals = groups[g][category];
          const total = Object.values(vals).reduce((a, b) => a + b, 0);
          return {
            label: g,
            data: labels.map(l => total > 0 ? +(((vals[l] || 0) / total * 100).toFixed(1)) : 0),
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
          ...barDatasets.flatMap(ds => ds.data || []),
          ...niValues.filter(v => typeof v === "number")
        ) * 1.05;
        const cappedMax = Math.min(isFinite(rawMax) ? rawMax : 100, 100);

        // Chart
        const dpr = Math.max(1, Math.min(3, window.devicePixelRatio || 1));

        const chart = new Chart(canvas, {
          type: "bar",
          data: { labels, datasets: chartDatasets },
          options: {
            indexAxis: "y",
            responsive: false,
            maintainAspectRatio: false,
            animation: false,
            layout: { padding: { top: CHART_TOP_PADDING, left: 10, right: 10, bottom: 0 } },
            plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.raw}%` } }
            },
            scales: {
              x: {
                beginAtZero: true,
                suggestedMax: cappedMax,
                title: { display: true, text: "Percentage" },
                ticks: { callback: v => `${v}%` }
              },
              y: { ticks: { display: false }, grid: { display: false }, offset: true }
            }
          },
          plugins: [
            {
              id: "aboveGroupLabels",
              afterDatasetsDraw(chartInst) {
                // ... (your original code for drawing labels)
              }
            },
            {
              id: "drawNILines",
              afterDatasetsDraw(chartInst) {
                // ... (your original code for drawing NI lines)
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
      button.setAttribute("aria-expanded", "true");  // tab through group-content
    } else {
      content.style.display = "none";
      button.setAttribute("aria-expanded", "false");  // tab through group-content
    }

    // Toggle behavior
    button.addEventListener("click", () => {
      const isVisible = content.style.display !== "none";
      content.style.display = isVisible ? "none" : "block";
      button.innerHTML = label.replace(isVisible ? "▲" : "▼", isVisible ? "▼" : "▲");
      button.setAttribute("aria-expanded", !isVisible);  // tab through group-content

      // If expanding, focus the first checkbox in group-content
      if (!isVisible) {
        const firstCheckbox = content.querySelector("input[type='checkbox']");
        if (firstCheckbox) {
          setTimeout(() => firstCheckbox.focus(), 0);
        }
      }
    });

    // Allow Enter/Space to toggle when focused on button
    button.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        button.click();
      }                                                   // tab through group-content
    });
  });


});

// Add year to data portal dropdown categorys
document.addEventListener("DOMContentLoaded", async () => {
  const response = await fetch("category_lookup.json");
  const lookup = await response.json();

  // Build lookup map
  const yearMap = lookup.reduce((acc, item) => {
    if (!item.nested_list_names || !item.Year) return acc;

    if (
      item.Source === "Data Portal" ||
      item.nested_list_names === "Benefits Statistics"
    ) {
      acc[item.nested_list_names] = item.Year;
    }

    return acc;
  }, {});

  document.querySelectorAll(".group-content label").forEach(label => {
    const input = label.querySelector("input[type='checkbox']");
    if (!input) return;

    const key = input.value;   // ✅ use value, NOT label text
    if (!yearMap[key]) return;

    const year = yearMap[key];

    // Visible label text (remove existing year if re-running)
    let displayText = label.textContent
      .replace(/\s*\(\d{4}\)\s*$/, "")
      .trim();

    // Remove all parenthetical content first, then rebuild
    displayText = displayText.replace(/\s*\([^)]*\)/g, "").trim();

    // Determine if we need "(4 categories)" based on the original key
    let categoryNote = "";
    if (key.includes("Age") && key.includes("MYE")) {
      categoryNote = "(4 Categories)";
    }

    // Preserve checkbox + listeners
    while (label.firstChild) label.removeChild(label.firstChild);
    label.appendChild(input);
    const finalText = categoryNote ? ` ${displayText} ${year} ${categoryNote}` : ` ${displayText} ${year}`;
    label.appendChild(
      document.createTextNode(finalText)
    );
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
    const origCanvases = Array.from(originalRoot.querySelectorAll('canvas'));
    const cloneCanvases = Array.from(cloneRoot.querySelectorAll('canvas'));
    const count = Math.min(origCanvases.length, cloneCanvases.length);

    for (let i = 0; i < count; i++) {
      const srcCanvas = origCanvases[i];
      const destCanvas = cloneCanvases[i];
      try {
        const dataUrl = srcCanvas.toDataURL('image/png');
        const img = new Image();
        img.src = dataUrl;
        img.style.maxWidth = '100%';
        img.style.display = 'block';
        img.style.marginTop = '10px';

        // Swap the canvas for an <img> in the CLONE
        destCanvas.parentNode.replaceChild(img, destCanvas);
      } catch (e) {
      }
    }
  }
  const cloneWrapper = document.createElement('div');
  cloneWrapper.style.background = '#fff';
  cloneWrapper.style.padding = '20px';
  cloneWrapper.style.fontFamily = 'sans-serif';
  cloneWrapper.style.maxWidth = '1200px';
  cloneWrapper.style.margin = '0 auto';

  const headerRow = document.querySelector('.header-row')
    || document.querySelector('.profile-header-bar');

  if (headerRow) {
    const headerClone = headerRow.cloneNode(true);
    const clonedTitle = headerClone.querySelector('#areaProfileTitle');
    if (clonedTitle) {
      clonedTitle.textContent = window.areaProfileTitle?.trim() || 'Selected area';
    }
    const clonedButtons = headerClone.querySelector('.profile-nav-buttons');
    if (clonedButtons) {
      clonedButtons.remove();
    }
    cloneWrapper.appendChild(headerClone);
  }

  const breakdownClone = breakdownContainer.cloneNode(true);
  const contentClone = contentSource.cloneNode(true);

  // Put the breakdown and content into the wrapper
  cloneWrapper.appendChild(breakdownClone);
  cloneWrapper.appendChild(contentClone);

  // Hide any download hover menu in the clone so the exported image is clean.
  cloneWrapper.querySelectorAll('.dropdown-menu').forEach(menu => {
    menu.style.display = 'none';
  });

  // Fix canvases in both cloned sections
  replaceCloneCanvasesWithImages(breakdownContainer, breakdownClone);
  replaceCloneCanvasesWithImages(contentSource, contentClone);

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
        borderRadius: '4px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        boxSizing: 'border-box'
      });

      if (originalTitle) {
        const title = document.createElement('h3');
        title.textContent = originalTitle.textContent.replace(/ Label$/, "");
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
          td.style.textAlign = columnIndex === 0 ? 'left' : 'right';
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
        borderRadius: '4px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        boxSizing: 'border-box'
      });

      if (originalTitle) {
        const title = document.createElement('h3');
        title.textContent = originalTitle.textContent.replace(/ Label$/, "");
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
          td.style.textAlign = columnIndex === 0 ? 'left' : 'right';
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

    logo.onload = async () => {
      const padding = 20;
      const maxLogoWidth = canvas.width * 0.25;
      const scaleFactor = Math.min(1, maxLogoWidth / logo.width);
      const scaledLogoWidth = logo.width * scaleFactor;
      const scaledLogoHeight = logo.height * scaleFactor;

      const finalCanvas = document.createElement('canvas');
      finalCanvas.width = canvas.width;
      finalCanvas.height = canvas.height + scaledLogoHeight + padding;

      const ctx = finalCanvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, finalCanvas.width, finalCanvas.height);

      ctx.drawImage(canvas, 0, 0);

      // logo at bottom right
      const x = finalCanvas.width - scaledLogoWidth - padding;
      const y = canvas.height + (padding / 2);
      ctx.drawImage(logo, x, y, scaledLogoWidth, scaledLogoHeight);

      // Trigger download using native save picker where available
      const dataUrl = finalCanvas.toDataURL('image/png');
      const blob = await fetch(dataUrl).then(r => r.blob());
      await saveBlobWithPicker(blob, 'area-summary.png');

      document.body.removeChild(cloneWrapper);
    };
  });
}


async function downloadExcel() {
  // Grab selected IDs (zones) and trigger a refresh of comparison data
  const selectedArray = Array.from(selectedIdsExcel);
  renderUrbanRuralComparison(selectedArray);

  // Pull in cached data from the window/global scope
  const aggregated = window.latestAggregatedData || {};
  const comparison = window.urbanRuralComparisonData || {};
  const selectedCategories = window.chosenCategories || [];
  const niTotals = window.niTotals || {};

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'NISRA Custom Area Profile Builder';
  workbook.created = new Date();

  const groups = ["Urban", "Rural", "Mixed"];
  const zoneType = window.selectedZoneType || 'sdz';
  const zoneTypeText = zoneType === 'sdz'
    ? 'Super Data Zone'
    : zoneType === 'dz'
      ? 'Data Zone'
      : zoneType === 'dea'
        ? 'District Electoral Area'
        : zoneType === 'lgd'
          ? 'Local Government District'
          : 'Data Zone';
  const totalZonesSelected = window.totalZonesSelected || 0;

  const labelLookup = {
    sdz: "Census 2021 Super Data Zone Label",
    dz: "Census 2021 Data Zone Label",
    dea: "District Electoral Area 2014 Label",
    lgd: "Local Government District 2021 Label"
  };

  const areaTitle = window.areaProfileTitle?.trim();
  const excelAreaName = areaTitle || 'Selected area';
  const percentageHeader = `${excelAreaName} %`;

  // ZONE BREAKDOWN SHEET
  const breakSheet = workbook.addWorksheet('Zone Breakdown');
  const showAreaTypeColumn = zoneType === 'sdz' || zoneType === 'dz';
  breakSheet.columns = showAreaTypeColumn
    ? [{ width: 40 }, { width: 18 }]
    : [{ width: 60 }];

  breakSheet.addRow(['NISRA Custom Area Profile Builder Extract']);
  breakSheet.addRow([]);
  breakSheet.addRow([
    `The information presented in tables are combined from ${totalZonesSelected} ${zoneTypeText}s listed below:`
  ]);
  breakSheet.addRow([`Date Extracted: ${new Date().toLocaleDateString('en-GB', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  })}`]);
  if (areaTitle) {
    breakSheet.addRow([`Area Name: ${areaTitle}`]);
  }
  breakSheet.addRow([]);

  const lgdGroups = {};
  selectedArray.forEach(id => {
    const mapData = window.selectedZoneDetails?.[id];
    if (!mapData) return;
    const lgd = mapData['LGD'];
    const status = mapData['Urban_mixed_rural_status'];
    const labelObj = mapData[labelLookup[zoneType]] || mapData['Census 2021 Super Data Zone Label'] || mapData['Census 2021 Data Zone Label'];
    const zoneName = labelObj ? Object.keys(labelObj)[0] : null;
    if (!zoneName || !lgd) return;
    if (!lgdGroups[lgd]) lgdGroups[lgd] = [];
    lgdGroups[lgd].push({ zoneName, status });
  });

  Object.keys(lgdGroups).sort().forEach(lgd => {
    const titleRow = breakSheet.addRow([`${lgd} LGD`]);
    titleRow.font = { bold: true };
    if (showAreaTypeColumn) {
      const headerRow = breakSheet.addRow(['Area Name', 'Area Type']);
      headerRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
    } else {
      const headerRow = breakSheet.addRow(['Area Name']);
      headerRow.getCell(1).font = { bold: true };
      headerRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    }

    lgdGroups[lgd].forEach(({ zoneName, status }) => {
      const row = showAreaTypeColumn
        ? breakSheet.addRow([zoneName, status || '-'])
        : breakSheet.addRow([zoneName]);
      row.eachCell(cell => {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
    });

    breakSheet.addRow([]);
  });

  // CATEGORY SHEETS
  const categories = selectedCategories.length ? selectedCategories : Object.keys(aggregated);
  categories.forEach(category => {
    const sheetName = category.replace(/ Label$/, '').substring(0, 31);
    const sheet = workbook.addWorksheet(sheetName);
    // Column widths: 1=Label, 2=Count, 3=Percentage header (dynamic), 4=NI %
    const col3Width = Math.max(15, (percentageHeader || '').length);
    sheet.columns = [
      { width: 15 },
      { width: 9 },
      { width: col3Width },
      { width: 9 }
    ];

    sheet.addRow([category.replace(/ Label$/, '')]);

    const headerRow = sheet.addRow(['Label', 'Count', percentageHeader, 'NI %']);
    headerRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true };
      cell.alignment = {
        horizontal: colNumber === 1 ? 'left' : 'right',
        vertical: 'middle'
      };
    });

    const values = aggregated[category] || {};
    const totalCount = Object.values(values).reduce((acc, val) => acc + val, 0);

    Object.entries(values).forEach(([label, count]) => {
      const percentage = totalCount > 0 ? count / totalCount : 0;
      const niVal = niTotals[category]?.[label];
      const niPercentage = typeof niVal === 'number' ? niVal / 100 : null;
      const row = sheet.addRow([
        label,
        count,
        percentage,
        niPercentage !== null ? niPercentage : '-'
      ]);
      row.getCell(2).numFmt = '#,##0';
      row.getCell(3).numFmt = '0.0%';
      if (niPercentage !== null) row.getCell(4).numFmt = '0.0%';
      row.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
      row.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
    });

    const totalRow = sheet.addRow([
      'Total',
      totalCount,
      totalCount > 0 ? 1 : 0,
      '-'
    ]);
    totalRow.getCell(2).numFmt = '#,##0';
    totalRow.getCell(3).numFmt = '0.0%';
    totalRow.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
    totalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    totalRow.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
    sheet.addRow([]);

    groups.forEach(group => {
      const groupCategoryData = comparison[group]?.[category];
      if (!groupCategoryData || Object.keys(groupCategoryData).length === 0) return;

      sheet.addRow([`${category.replace(/ Label$/, '')} – ${group}`]);
      const groupHeaderRow = sheet.addRow(['Label', 'Count', percentageHeader, 'NI %']);
      groupHeaderRow.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        cell.alignment = {
          horizontal: colNumber === 1 ? 'left' : 'right',
          vertical: 'middle'
        };
      });

      const total = Object.values(groupCategoryData).reduce((acc, val) => acc + val, 0);
      Object.entries(groupCategoryData).forEach(([label, count]) => {
        const pct = total > 0 ? count / total : 0;
        const niVal = niTotals[category]?.[label];
        const niPct = typeof niVal === 'number' ? niVal / 100 : null;
        const row = sheet.addRow([
          label,
          count,
          pct,
          niPct !== null ? niPct : '-'
        ]);
        row.getCell(2).numFmt = '#,##0';
        row.getCell(3).numFmt = '0.0%';
        if (niPct !== null) row.getCell(4).numFmt = '0.0%';
        row.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
        row.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
        row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
      });

      const groupTotalRow = sheet.addRow([
        'Total',
        total,
        total > 0 ? 1 : 0,
        '-'
      ]);
      groupTotalRow.getCell(2).numFmt = '#,##0';
      groupTotalRow.getCell(3).numFmt = '0.0%';
      groupTotalRow.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
      groupTotalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      groupTotalRow.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
      sheet.addRow([]);
    });

    sheet.addRow([]);
    sheet.addRow(['Notes on custom area aggregations']);
    sheet.addRow([]);
    sheet.addRow([
      'Any aggregations created from Census Flexible Builder data may differ slightly from published Census figures.'
    ]);
    sheet.addRow([
      'For Census 2021, NISRA applied two Statistical Disclosure Control strategies: Targeted Record Swapping (TRS) and Cell Key Perturbation (CKP).'
    ]);
    sheet.addRow([
      'CKP may add small amounts of variation to some cells. Where two or more different aggregations are created, the totals of all cells may in turn be different.'
    ]);
    sheet.addRow([
      'Overall, the differences will be small and should not change the conclusions of any analysis or research.'
    ]);
    sheet.addRow([]);
    sheet.addRow([
      'Linked to the Statistical Disclosure Control Methods applied above, when viewing breakdowns at small geographical levels in this application,'
    ]);
    sheet.addRow([
      'cell counts of under 5 may be seen. The use of TRS and CKP, mean this number could be anything from 0-4, or could have been swapped with another census record entry elsewhere.'
    ]);
    sheet.addRow([]);
    sheet.addRow([
      'For more information, please refer to the NISRA statistical disclosure control methodology:'
    ]);
    const methodRow = sheet.addRow([
      'NISRA Statistical Disclosure Control Methodology'
    ]);
    const methodCell = methodRow.getCell(1);
    methodCell.value = {
      text: 'NISRA Statistical Disclosure Control Methodology',
      hyperlink: 'https://www.nisra.gov.uk/files/nisra/publications/statistical-disclosure-control-methodology-for-2021-census.pdf'
    };
    methodCell.font = { color: { argb: 'FF0000FF' }, underline: true };
  });

  const now = new Date();
  const datePart = now.toLocaleDateString('en-GB').replace(/\//g, '-');
  const timePart = now.toLocaleTimeString('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  }).replace(/:/g, '-');
  const filename = `NISRA Custom Area Profile Extract-${datePart} ${timePart}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  await saveBlobWithPicker(blob, filename);
}

async function saveBlobWithPicker(blob, suggestedName) {
  if (window.showSaveFilePicker) {
    try {
      const options = {
        suggestedName,
        types: [
          {
            description: 'File',
            accept: {
              [blob.type]: ['.' + suggestedName.split('.').pop()]
            }
          }
        ]
      };
      const handle = await window.showSaveFilePicker(options);
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return;
    } catch (err) {
      if (err && (err.name === 'AbortError' || err.name === 'NotAllowedError')) {
        console.info('Save file picker canceled by user. No file was downloaded.');
        return;
      }
      console.warn('Save file picker failed, falling back to default download.', err);
    }
  }

  if (navigator.msSaveOrOpenBlob) {
    navigator.msSaveOrOpenBlob(blob, suggestedName);
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = suggestedName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function markEmptyCategoryGroups() {
  document.querySelectorAll('.category-group').forEach(group => {
    const content = group.querySelector('.group-content');
    if (!content) return;

    // Check for indicator checkboxes
    const hasIndicators = content.querySelectorAll(
      'input[type="checkbox"]'
    ).length > 0;

    if (!hasIndicators) {
      // Prevent duplicate messages
      if (!content.querySelector('.no-indicators')) {
        const msg = document.createElement('div');
        msg.className = 'no-indicators';
        msg.textContent = 'No Indicators available for this theme';
        content.appendChild(msg);
      }
    }
  });
}

// Reset Selections
function resetCategorySelections() {
  const checkboxes = document.querySelectorAll('#category-form input[type="checkbox"]');

  // Uncheck ALL boxes
  checkboxes.forEach(cb => {
    cb.checked = false;
  });

  // Clear your tracked state
  selectedCategories = [];

  // Trigger existing update logic
  document.getElementById("category-form").dispatchEvent(new Event("change"));
}

document.getElementById("reset-categories-btn").addEventListener("click", resetCategorySelections);

document.querySelectorAll('[data-nav="howto"]').forEach(btn => {
  btn.addEventListener('click', () => {
    window.open('landing.html', '_blank');
  });
});
function removeCirclePreview() {
  const source = map.getSource('circle-preview');
  if (source) {
    source.setData({
      type: 'FeatureCollection',
      features: []
    });
  }
}

// Create Your Own Profile Button

const sidebar = document.getElementById("profile-sidebar");
const backdrop = document.getElementById("profile-backdrop");

// Open
document.getElementById("open-profile").addEventListener("click", () => {
  sidebar.classList.add("open");
  backdrop.classList.add("active");
});

// Close function
function closeProfile() {
  sidebar.classList.remove("open");
  backdrop.classList.remove("active");
}

// Close via X button
document.querySelectorAll(".close-on-mobile").forEach(btn => {
  btn.addEventListener("click", closeProfile);
});

// Close via backdrop
backdrop.addEventListener("click", closeProfile);

// Close via ESC
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && sidebar.classList.contains("open")) {
    closeProfile();
  }
});

document.addEventListener("click", (e) => {
  const isSidebarOpen = sidebar.classList.contains("open");

  if (!isSidebarOpen) return;

  const clickedInsideSidebar = sidebar.contains(e.target);
  const clickedOpenButton = e.target.closest("#open-profile");

  // If click is outside sidebar AND not the button that opens it
  if (!clickedInsideSidebar && !clickedOpenButton) {
    closeProfile();
  }
});
