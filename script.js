// ══════════════════════════════════════════
// SECTION NAVIGATION
// ══════════════════════════════════════════
function showSection(name) {
  document.getElementById('landingSection').style.display    = 'none';
  document.getElementById('attendanceSection').style.display = 'none';
  document.getElementById('meetingSection').style.display    = 'none';
  if (name === 'home') {
    document.getElementById('landingSection').style.display = 'block';
  } else if (name === 'attendance') {
    document.getElementById('attendanceSection').style.display = 'block';
  } else if (name === 'meeting') {
    document.getElementById('meetingSection').style.display = 'block';
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
// Start on landing page
showSection('home');

// ══════════════════════════════════════════
// ATTENDANCE MODULE VARIABLES
// ══════════════════════════════════════════
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const statusPanel = document.getElementById('statusPanel');
const statusText = document.getElementById('statusText');
const downloadBtnContainer = document.getElementById('downloadBtnContainer');
const downloadBtn = document.getElementById('downloadBtn');

let processedWorkbook = null;
let notMarkedWorkbook = null;
let originalFileName = "";
let globalFilteredData = [];

// Handle drag and drop
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  uploadArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
  e.preventDefault();
  e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
  uploadArea.addEventListener(eventName, () => uploadArea.classList.add('dragover'), false);
});

['dragleave', 'drop'].forEach(eventName => {
  uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('dragover'), false);
});

uploadArea.addEventListener('drop', (e) => {
  const dt = e.dataTransfer;
  const files = dt.files;
  if (files.length > 0) {
    handleFile(files[0]);
  }
});

fileInput.addEventListener('change', function(e) {
  if (e.target.files.length > 0) {
    handleFile(e.target.files[0]);
  }
});

function handleFile(file) {
  originalFileName = file.name;
  
  // Update UI
  uploadArea.style.display = 'none';
  statusPanel.style.display = 'block';
  statusText.textContent = "Reading file...";
  statusText.style.color = "var(--text-main)";
  document.querySelector('.spinner').style.display = 'block';
  downloadBtnContainer.style.display = 'none';

  const reader = new FileReader();

  reader.onload = function(e) {
    statusText.textContent = "Processing data...";
    
    // We use setTimeout to allow parsing large files without freezing UI completely
    setTimeout(() => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        
        processWorkbook(workbook);
      } catch (err) {
        console.error(err);
        statusText.textContent = "Error processing file: " + err.message;
        statusText.style.color = "#ef4444";
        document.querySelector('.spinner').style.display = 'none';
      }
    }, 100);
  };

  reader.onerror = function() {
    statusText.textContent = "Error reading file.";
    statusText.style.color = "#ef4444";
    document.querySelector('.spinner').style.display = 'none';
  };

  reader.readAsArrayBuffer(file);
}

function processWorkbook(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Find the header row by parsing as 2D array first
  const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  
  let headerRowIndex = -1;
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const isHeader = row.some(cell => {
      const s = String(cell).toLowerCase().replace(/[\s-]/g, '');
      return s.includes('punchin') || s.includes('punchtime') || s.includes('intime') || s.includes('firstpunch') || String(cell).toLowerCase().includes('punch in');
    });

    if (isHeader) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error("Could not find a 'Punch-In' or 'Time' column in the document.");
  }

  // Convert sheet to JSON starting from the found header row
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, defval: "" });

  if (jsonData.length === 0) {
    throw new Error("The Excel file has no data below the headers.");
  }

  // Find column names
  const rowObj = jsonData[0];
  const columnNames = Object.keys(rowObj);
  
  // Look for columns containing our target names, case-insensitive
  const punchInCol = columnNames.find(c => {
      const s = c.toLowerCase().replace(/[\s-]/g, '');
      return s.includes('punchin') || s.includes('punchtime') || s.includes('intime') || s.includes('firstpunch') || c.toLowerCase().includes('punch in');
  });
  const regionCol = columnNames.find(c => c.toLowerCase().trim() === 'region' || c.toLowerCase().trim() === 'branch' || c.toLowerCase().trim() === 'location');

  if (!punchInCol) {
    throw new Error("Could not find the associated 'Punch-In' column inside the parsed table.");
  }
  if (!regionCol) {
    throw new Error("Could not find a 'REGION' column.");
  }

  const allowedRegions = [
    "TELANGANA AND ANDRA PRADESH",
    "TELANGANA ANDANDRA PRADESH",
    "TELANGANA",
    "ANDRA PRADESH",
    "ANDHRA PRADESH",
    "DHARWAD", 
    "KALABURAGI", 
    "TUMKUR",
    "CHITRADURGA"
  ];
  
  // Collect ALL unique region values from the entire file for debug purposes
  const allUniqueRegions = new Set();
  jsonData.forEach(row => {
    const rv = (row[regionCol] || "").toString().trim();
    if (rv) allUniqueRegions.add(rv);
  });

  // Process data
  const filteredData = [];
  
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    const regionValue = (row[regionCol] || "").toString().trim().toUpperCase();
    
    const isAllowedRegion = allowedRegions.some(allowed => 
      regionValue === allowed.toUpperCase() || 
      regionValue.includes(allowed.toUpperCase())
    );

    if (isAllowedRegion) {
      const punchInVal = (row[punchInCol] || "").toString().trim();
      let extractedTime = "";
      let remarks = "";

      if (!punchInVal || punchInVal.toUpperCase() === "A") {
        extractedTime = punchInVal || "A";
        remarks = "not marked till 9 45 am";
      } else {
        // Parse time from strings like "2026-04-03 06:51:25 +0530"
        const timeMatch = punchInVal.match(/(\d{2}:\d{2}:\d{2})/);
        
        if (timeMatch) {
          extractedTime = timeMatch[1];
          // Extracted time is something like "06:51:25"
          const timeParts = extractedTime.split(':').map(Number);
          const hours = timeParts[0];
          const minutes = timeParts[1];
          
          const timeInMinutes = hours * 60 + minutes;
          const time730 = 7 * 60 + 30;  // 450 minutes
          const time830 = 8 * 60 + 30;  // 510 minutes
          const time945 = 9 * 60 + 45;  // 585 minutes

          if (timeInMinutes < time730) {
            remarks = "Before 7:30";
          } else if (timeInMinutes < time830) {
            remarks = "7:30 to 8:30";
          } else if (timeInMinutes < time945) {
            remarks = "8:30 and above";
          } else {
            // Punched in at 9:45 AM or later — treated as not marked
            remarks = "not marked till 9 45 am";
          }
        } else {
          // If no time is matched, leave values empty or fallback
          extractedTime = punchInVal;
          remarks = "Could not parse time";
        }
      }

      // Add to row
      row["Extracted Time"] = extractedTime;
      row["Remarks"] = remarks;

      filteredData.push(row);
    }
  }

  if (filteredData.length === 0) {
    statusText.textContent = "No rows found for the specified regions.";
    statusText.style.color = "#f59e0b"; // Warning color
    document.querySelector('.spinner').style.display = 'none';
    return;
  }

  // Provide exactly the same keys plus Extracted Time and Remarks at the end
  const newColumnOrder = [...columnNames, "Extracted Time", "Remarks"];
  
  globalFilteredData = filteredData;
  generatePivotTable(filteredData);

  const newWorksheet = XLSX.utils.json_to_sheet(filteredData, { header: newColumnOrder });
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Processed Data");
  
  processedWorkbook = newWorkbook;

  const notMarkedData = filteredData.filter(row => row["Remarks"] === "not marked till 9 45 am");
  if (notMarkedData.length > 0) {
      // Find specific columns to include in the Not Marked Excel
      const nameCol       = columnNames.find(c => c.toLowerCase().includes('name'));
      const identifierCol = columnNames.find(c => c.toLowerCase().includes('identifier'));
      const workLocCol    = columnNames.find(c => {
        const s = c.toLowerCase().replace(/\s/g, '');
        return s.includes('workloc') || s.includes('worklocation');
      }) || columnNames.find(c => c.toLowerCase().includes('location') || c.toLowerCase().includes('branch'));
      const designationCol= columnNames.find(c => c.toLowerCase().includes('designation'));
      const areaCol       = columnNames.find(c => c.toLowerCase().trim() === 'area');
      const regionColNM   = columnNames.find(c => c.toLowerCase().trim() === 'region');
      const idCol         = columnNames.find(c => c.toLowerCase().trim() === 'id' || c.toLowerCase().trim() === 'emp id' || c.toLowerCase().trim() === 'employee id');

      // Build ordered column list (only those that exist)
      const nmColumns = [nameCol, identifierCol, workLocCol, designationCol, areaCol, regionColNM, idCol, punchInCol]
        .filter(Boolean); // remove undefined entries

      // Build filtered rows with only the selected columns
      const nmFilteredRows = notMarkedData.map(row => {
        const obj = {};
        nmColumns.forEach(col => { obj[col] = row[col] ?? ""; });
        return obj;
      });

      const nmWorkbook = XLSX.utils.book_new();
      const nmWorksheet = XLSX.utils.json_to_sheet(nmFilteredRows, { header: nmColumns });
      XLSX.utils.book_append_sheet(nmWorkbook, nmWorksheet, "Not Marked 9-45 AM");
      notMarkedWorkbook = nmWorkbook;
  } else {
      notMarkedWorkbook = null;
  }

  statusText.textContent = "Processing complete! Found " + filteredData.length + " matching rows.";
  document.querySelector('.spinner').style.display = 'none';
  statusText.style.color = "var(--success)";
  downloadBtnContainer.style.display = 'flex';

  // ---- Chitradurga debug: check if 0 rows matched Chitradurga ----
  const chitradurgaCount = filteredData.filter(row => {
    const rv = (row[regionCol] || "").toString().trim().toUpperCase();
    return rv.includes("CHITRADURGA");
  }).length;

  // Find region values in the FULL file that didn't match any known region
  const unmatchedRegions = [];
  allUniqueRegions.forEach(rv => {
    const up = rv.toUpperCase();
    const matched = allowedRegions.some(a => up === a || up.includes(a));
    if (!matched) unmatchedRegions.push(rv);
  });

  // Remove any existing debug panel
  const existingDebug = document.getElementById('regionDebugPanel');
  if (existingDebug) existingDebug.remove();

  if (chitradurgaCount === 0 && unmatchedRegions.length > 0) {
    const debugPanel = document.createElement('div');
    debugPanel.id = 'regionDebugPanel';
    debugPanel.style.cssText = `
      margin-top: 1rem; padding: 1rem 1.2rem; background: #fffbeb;
      border: 2px solid #f59e0b; border-radius: 8px; color: #92400e;
      font-size: 0.88rem; line-height: 1.7;
    `;
    debugPanel.innerHTML = `
      <strong>⚠️ Chitradurga shows 0 — Region names found in your Excel that didn't match:</strong><br>
      <span style="font-family: monospace; background: #fef3c7; padding: 2px 6px; border-radius: 4px;">
        ${unmatchedRegions.map(r => `"${r}"`).join('</span>, <span style="font-family: monospace; background: #fef3c7; padding: 2px 6px; border-radius: 4px;">')}
      </span><br><br>
      👉 <strong>Please share the exact spelling used in your Excel for Chitradurga employees</strong> so it can be added to the matching list.
    `;
    document.getElementById('statusPanel').after(debugPanel);
  } else if (chitradurgaCount > 0) {
    // All good — remove debug panel if it was there
    const dp = document.getElementById('regionDebugPanel');
    if (dp) dp.remove();
  }
}

downloadBtn.addEventListener('click', () => {
  if (processedWorkbook) {
    const outputName = originalFileName.replace(/\.[^/.]+$/, "") + "_processed.xlsx";
    XLSX.writeFile(processedWorkbook, outputName);
    
    // Reset UI so user can upload another
    setTimeout(() => {
      uploadArea.style.display = 'block';
      statusPanel.style.display = 'none';
      downloadBtnContainer.style.display = 'none';
      document.getElementById('pivotContainer').style.display = 'none';
      fileInput.value = '';
    }, 2000);
  }
});

document.getElementById('downloadNotMarkedBtn')?.addEventListener('click', () => {
  if (notMarkedWorkbook) {
    const outputName = originalFileName.replace(/\.[^/.]+$/, "") + "_Not_Marked_9_45_AM.xlsx";
    XLSX.writeFile(notMarkedWorkbook, outputName);
  } else {
    alert("No employees found who haven't marked till 9:45 AM.");
  }
});

function generatePivotTable(data) {
  const regions = ["APTS", "Dharwad", "Kalburagi", "Tumkur", "Chitradurga"];
  const stats = {
    "APTS":        { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Dharwad":     { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Kalburagi":   { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Tumkur":      { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Chitradurga": { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 }
  };

  const keys = data.length > 0 ? Object.keys(data[0]) : [];
  const regionCol = keys.find(k => k.toLowerCase() === 'region');
  const punchCol = keys.find(k => k.toLowerCase().includes('punch-in time') || k.toLowerCase().includes('punch in time'));
  
  let reportDate = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
  if (data.length > 0 && punchCol) {
    for (let i = 0; i < data.length; i++) {
      const val = (data[i][punchCol] || "").toString();
      const dateMatch = val.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (dateMatch) {
        reportDate = `${dateMatch[3]}-${dateMatch[2]}-${dateMatch[1]}`;
        break;
      }
    }
  }

  data.forEach(row => {
    let rawRegion = regionCol ? (row[regionCol] || "").toString().trim().toUpperCase() : "";
    let displayRegion = null;
    
    if (rawRegion.includes("TELANGANA") || rawRegion.includes("ANDRA") || rawRegion === "APTS" || rawRegion === "TS") displayRegion = "APTS";
    else if (rawRegion.includes("DHARWAD")) displayRegion = "Dharwad";
    else if (rawRegion.includes("KALABURAGI")) displayRegion = "Kalburagi";
    else if (rawRegion.includes("TUMKUR")) displayRegion = "Tumkur";
    else if (rawRegion.includes("CHITRADURGA")) displayRegion = "Chitradurga";

    if (displayRegion) {
       stats[displayRegion].total++;
       let remark = row["Remarks"];
       if (remark === "Before 7:30") stats[displayRegion].b730++;
       else if (remark === "7:30 to 8:30") stats[displayRegion].a730_b830++;
       else if (remark === "8:30 and above") stats[displayRegion].a830++;
       else if (remark === "not marked till 9 45 am") stats[displayRegion].notMarked++;
    }
  });

  let t_total = 0, t_b730 = 0, t_a730_b830 = 0, t_a830 = 0, t_notMarked = 0;

  let tbody = "";
  regions.forEach((r, idx) => {
    t_total += stats[r].total;
    t_b730 += stats[r].b730;
    t_a730_b830 += stats[r].a730_b830;
    t_a830 += stats[r].a830;
    t_notMarked += stats[r].notMarked;
    tbody += `
      <tr>
        <td style="text-align: center;">${idx+1}</td>
        <td style="color: #000; font-weight: 800;">${r}</td>
        <td>${stats[r].total}</td>
        <td>${stats[r].b730}</td>
        <td>${stats[r].a730_b830}</td>
        <td>${stats[r].a830}</td>
        <td>${stats[r].notMarked}</td>
      </tr>
    `;
  });

  let p_b730 = t_total ? Math.round((t_b730 / t_total) * 100) : 0;
  let p_a730_b830 = t_total ? Math.round((t_a730_b830 / t_total) * 100) : 0;
  let p_a830 = t_total ? Math.round((t_a830 / t_total) * 100) : 0;
  let p_notMarked = t_total ? Math.round((t_notMarked / t_total) * 100) : 0;

  const html = `
    <table class="pivot-table">
      <thead>
        <tr class="header-main">
          <th colspan="6">Daily Attendance report</th>
          <th>${reportDate}</th>
        </tr>
        <tr class="header-sub">
          <th>Sl. No.</th>
          <th>Region</th>
          <th>Total Staffs</th>
          <th>Marked before 7:30 am</th>
          <th>Marked 7:30 to 8:30am</th>
          <th>Marked Above 8:30am</th>
          <th>Not been marked till 9:45 AM.</th>
        </tr>
      </thead>
      <tbody>
        ${tbody}
        <tr>
          <td colspan="2" style="text-align: center;">Total</td>
          <td>${t_total}</td>
          <td>${t_b730}</td>
          <td>${t_a730_b830}</td>
          <td>${t_a830}</td>
          <td>${t_notMarked}</td>
        </tr>
        <tr>
          <td colspan="2" style="text-align: center;">Percentage</td>
          <td></td>
          <td>${p_b730}%</td>
          <td>${p_a730_b830}%</td>
          <td>${p_a830}%</td>
          <td>${p_notMarked}%</td>
        </tr>
      </tbody>
    </table>
  `;

  let regionButtons = `<div style="margin-top: 3rem; background: rgba(15, 23, 42, 0.02); padding: 1.5rem; border-radius: 8px; border: 1px dashed #cbd5e1;">`;
  regionButtons += `<h3 style="color: #3b82f6; margin-bottom: 1rem; font-size: 1.1rem; text-align: center;">Drill-down: View Interactive Dashboards by Region</h3>`;
  regionButtons += `<div class="slider-container" style="justify-content: center; flex-wrap: wrap;">`;
  regions.forEach(r => {
    regionButtons += `<button class="slide-tab" onclick="showRegionDetails('${r}')" style="font-size: 1rem; padding: 0.8rem 1.5rem; border: 2px solid #3b82f6; color: #3b82f6; box-shadow: 0 4px 6px rgba(59, 130, 246, 0.1);">${r} Dashboard &rarr;</button>`;
  });
  regionButtons += `</div></div>`;

  document.getElementById('pivotReport').innerHTML = html;
  document.getElementById('regionButtonsContainer').innerHTML = regionButtons;
  document.getElementById('pivotContainer').style.display = 'block';
}

document.getElementById('downloadImgBtn').addEventListener('click', () => {
  const pivotEl = document.getElementById('pivotReport');

  // Apply maximum-contrast styles before capture
  pivotEl.classList.add('capture-mode');

  // Force every cell to pitch-black with max weight + shadow for extra darkness
  const cells = pivotEl.querySelectorAll('td, th');
  cells.forEach(cell => {
    cell.dataset.oldColor  = cell.style.color;
    cell.dataset.oldWeight = cell.style.fontWeight;
    cell.dataset.oldShadow = cell.style.textShadow;
    cell.dataset.oldSize   = cell.style.fontSize;
    cell.dataset.oldOpacity= cell.style.opacity;

    cell.style.color      = '#000000';
    cell.style.fontWeight = '900';
    cell.style.fontSize   = '13px';
    cell.style.opacity    = '1';
    // Subtle repeat-shadow thickens thin strokes — makes numbers visibly bolder
    cell.style.textShadow = '0 0 0.3px #000, 0 0 0.3px #000, 0 0 0.5px #000';
  });

  // Also force ALL spans / divs inside to black
  const allText = pivotEl.querySelectorAll('*');
  allText.forEach(el => {
    el.dataset.oldElColor = el.style.color;
    el.style.color = '#000000';
  });

  html2canvas(pivotEl, {
    scale: 3,                    // 3× resolution — very crisp & sharp
    backgroundColor: '#ffffff',
    useCORS: true,
    logging: false,
    allowTaint: true
  }).then(canvas => {
    // Restore all original styles
    pivotEl.classList.remove('capture-mode');
    cells.forEach(cell => {
      cell.style.color      = cell.dataset.oldColor  || '';
      cell.style.fontWeight = cell.dataset.oldWeight || '';
      cell.style.textShadow = cell.dataset.oldShadow || '';
      cell.style.fontSize   = cell.dataset.oldSize   || '';
      cell.style.opacity    = cell.dataset.oldOpacity|| '';
      delete cell.dataset.oldColor; delete cell.dataset.oldWeight;
      delete cell.dataset.oldShadow; delete cell.dataset.oldSize;
      delete cell.dataset.oldOpacity;
    });
    allText.forEach(el => {
      el.style.color = el.dataset.oldElColor || '';
      delete el.dataset.oldElColor;
    });

    const link = document.createElement('a');
    link.download = 'attendance_pivot.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  }).catch(err => {
    pivotEl.classList.remove('capture-mode');
    cells.forEach(cell => {
      cell.style.color = cell.dataset.oldColor || '';
      delete cell.dataset.oldColor;
    });
    console.error('Image capture failed:', err);
  });
});

function showRegionDetails(region) {
  document.getElementById('pivotContainer').style.display = 'none';
  document.getElementById('downloadBtnContainer').style.display = 'none';
  
  const keys = globalFilteredData.length > 0 ? Object.keys(globalFilteredData[0]) : [];
  const regionCol = keys.find(k => k.toLowerCase() === 'region');

  const detailData = globalFilteredData.filter(row => {
    let rawRegion = regionCol ? (row[regionCol] || "").toString().trim().toUpperCase() : "";
    let displayRegion = null;
    
    if (rawRegion.includes("TELANGANA") || rawRegion.includes("ANDRA") || rawRegion === "APTS" || rawRegion === "TS") displayRegion = "APTS";
    else if (rawRegion.includes("DHARWAD")) displayRegion = "Dharwad";
    else if (rawRegion.includes("KALABURAGI")) displayRegion = "Kalburagi";
    else if (rawRegion.includes("TUMKUR")) displayRegion = "Tumkur";
    else if (rawRegion.includes("CHITRADURGA")) displayRegion = "Chitradurga";

    return displayRegion === region;
  });

  document.getElementById('detailTitle').textContent = region + " - Work Location Dashboard";

  if (detailData.length === 0) {
     document.getElementById('detailReport').innerHTML = "<p style='color:#000;'>No data found for this region.</p>";
  } else {
     const workLocCol = keys.find(k => k.toLowerCase().replace(/ /g, '') === 'worklocation' || k.toLowerCase() === 'location' || k.toLowerCase() === 'branch') || keys.find(k => k.toLowerCase().includes('location'));
     
     if (!workLocCol) {
        document.getElementById('detailReport').innerHTML = "<p style='color:#000;'>Work Location column not found in data.</p>";
     } else {
        const locSet = new Set();
        detailData.forEach(r => locSet.add((r[workLocCol] || "Unknown").toString().trim()));
        const locations = Array.from(locSet).sort();
        window.currentDetailData = detailData;
        window.currentRegion = region;
        window.currentLocCol = workLocCol;
        window.currentKeys = keys;
        
        let tabsHtml = `<button class="slide-tab active" onclick="renderSlide('All')">All Locations Summary</button>`;
        locations.forEach(l => {
            tabsHtml += `<button class="slide-tab" onclick="renderSlide('${l.replace(/'/g, "\\'")}')">${l}</button>`;
        });
        document.getElementById('locationTabs').innerHTML = tabsHtml;
        
        renderSlide('All');
     }
  }
  
  document.getElementById('detailContainer').style.display = 'block';
}

window.renderSlide = function(locName) {
   const tabs = document.querySelectorAll('.slide-tab');
   tabs.forEach(t => {
       if (t.textContent === locName || (locName === 'All' && t.textContent === 'All Locations Summary')) t.classList.add('active');
       else t.classList.remove('active');
   });

   let slideData = window.currentDetailData;
   if (locName !== 'All') {
       slideData = window.currentDetailData.filter(r => (r[window.currentLocCol] || "Unknown").toString().trim() === locName);
   }

   const locStats = {};
   slideData.forEach(row => {
      let loc = (row[window.currentLocCol] || "Unknown").toString().trim();
      if (!locStats[loc]) {
          locStats[loc] = { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 };
      }
      locStats[loc].total++;
      let remark = row["Remarks"];
      if (remark === "Before 7:30") locStats[loc].b730++;
      else if (remark === "7:30 to 8:30") locStats[loc].a730_b830++;
      else if (remark === "8:30 and above") locStats[loc].a830++;
      else if (remark === "not marked till 9 45 am") locStats[loc].notMarked++;
   });

   const locations = Object.keys(locStats).sort();
   let tbody = "";
   let t_total = 0, t_b730 = 0, t_a730_b830 = 0, t_a830 = 0, t_notMarked = 0;

   locations.forEach((loc, idx) => {
       t_total += locStats[loc].total;
       t_b730 += locStats[loc].b730;
       t_a730_b830 += locStats[loc].a730_b830;
       t_a830 += locStats[loc].a830;
       t_notMarked += locStats[loc].notMarked;

       tbody += `
         <tr>
           <td style="text-align: center;">${idx+1}</td>
           <td style="text-align: left;">${loc}</td>
           <td>${locStats[loc].total}</td>
           <td>${locStats[loc].b730}</td>
           <td>${locStats[loc].a730_b830}</td>
           <td>${locStats[loc].a830}</td>
           <td>${locStats[loc].notMarked}</td>
         </tr>
       `;
   });

   let p_b730 = t_total ? Math.round((t_b730 / t_total) * 100) : 0;
   let p_a730_b830 = t_total ? Math.round((t_a730_b830 / t_total) * 100) : 0;
   let p_a830 = t_total ? Math.round((t_a830 / t_total) * 100) : 0;
   let p_notMarked = t_total ? Math.round((t_notMarked / t_total) * 100) : 0;

   const punchCol = window.currentKeys.find(k => {
     let s = k.toLowerCase().replace(/[\s-]/g, '');
     return s.includes('punchin') || s.includes('punchtime') || s.includes('intime') || k.toLowerCase().includes('punch in');
   });
   let reportDate = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
   if (punchCol) {
       for (let i = 0; i < slideData.length; i++) {
          const val = (slideData[i][punchCol] || "").toString();
          const dateMatch = val.match(/(\d{4})-(\d{2})-(\d{2})/);
          if (dateMatch) {
              reportDate = `${dateMatch[3]}-${dateMatch[2]}-${dateMatch[1]}`;
              break;
          }
       }
   }

   let dynamicHeader = locName === 'All' ? `${window.currentRegion} Region - Locational Summary` : `${locName} Branch - Live Dashboard`;

   const html = `
     <table class="pivot-table">
       <thead>
         <tr class="header-main">
           <th colspan="6">${dynamicHeader}</th>
           <th>${reportDate}</th>
         </tr>
         <tr class="header-sub">
           <th>Sl. No.</th>
           <th>Work Location</th>
           <th>Total Staffs</th>
           <th>Marked before 7:30 am</th>
           <th>Marked 7:30 to 8:30am</th>
           <th>Marked Above 8:30am</th>
           <th>Not been marked till 9:45 AM.</th>
         </tr>
       </thead>
       <tbody>
         ${tbody}
         <tr>
           <td colspan="2" style="text-align: center;">Total</td>
           <td>${t_total}</td>
           <td>${t_b730}</td>
           <td>${t_a730_b830}</td>
           <td>${t_a830}</td>
           <td>${t_notMarked}</td>
         </tr>
         <tr>
           <td colspan="2" style="text-align: center;">Percentage</td>
           <td></td>
           <td>${p_b730}%</td>
           <td>${p_a730_b830}%</td>
           <td>${p_a830}%</td>
           <td>${p_notMarked}%</td>
         </tr>
       </tbody>
     </table>
   `;

   const nameCol = window.currentKeys.find(k => k.toLowerCase().includes('name') || k.toLowerCase().includes('emp')) || "";
   const timeRawCol = window.currentKeys.find(k => {
     let s = k.toLowerCase().replace(/[\s-]/g, '');
     return s.includes('punchin') || s.includes('punchtime') || s.includes('intime') || k.toLowerCase().includes('punch in');
   });
   
   let rawBody = "";
   let filteredAbsences = 0;
   slideData.forEach((row, idx) => {
     let empName = nameCol ? (row[nameCol] || "-") : "-";
     let loc = window.currentLocCol ? (row[window.currentLocCol] || "-") : "-";
     let punchTime = timeRawCol ? (row[timeRawCol] || "-") : "-";
     let remark = row["Remarks"] || "-";
     if (remark.includes("not marked") || remark.includes("Not marked")) filteredAbsences++;

     let trStyle = remark.includes("not marked") ? "background-color: rgba(239, 68, 68, 0.05);" : "";
     
     rawBody += `
       <tr style="font-size: 0.85rem; cursor: default; box-shadow: none; ${trStyle}">
         <td style="text-align: center;">${idx + 1}</td>
         <td style="text-align: left;">${empName}</td>
         <td style="text-align: left;">${loc}</td>
         <td style="text-align: center;">${punchTime}</td>
         <td style="text-align: center;">${remark}</td>
       </tr>
     `;
   });

   let leafDetailsHeader = locName === 'All' ? `All Punch Logs & Leave Updates (${filteredAbsences} total absences)` : `${locName} Punch Logs & Leave Updates (${filteredAbsences} absences here)`;

   const rawHtml = `
     <div style="margin-top: 3rem;">
         <h3 style="color: #0bd18a; margin-bottom: 1rem; border-bottom: 2px solid #334155; padding-bottom: 0.5rem; font-size: 1.2rem;">${leafDetailsHeader}</h3>
         <table class="pivot-table">
           <thead>
             <tr class="header-sub">
                 <th>Sl. No.</th>
                 <th>Employee / Staff Name</th>
                 <th>Work Location</th>
                 <th>Actual Punch-In Time</th>
                 <th>Remarks Bucket</th>
             </tr>
           </thead>
           <tbody>
             ${rawBody}
           </tbody>
         </table>
     </div>
   `;

   document.getElementById('detailReport').innerHTML = html + rawHtml;
}

document.getElementById('backBtn').addEventListener('click', () => {
  document.getElementById('detailContainer').style.display = 'none';
  document.getElementById('pivotContainer').style.display = 'block';
  document.getElementById('downloadBtnContainer').style.display = 'flex';
});

// ══════════════════════════════════════════
// MEETING MODULE
// ══════════════════════════════════════════
let meetingWorkbook   = null;
let meetingFileName   = '';
let meetingAllColumns = [];
let meetingAllRows    = [];
let mCols = {}; // { region, area, workLoc, b730, a730_830, a830, total }

const meetingUploadArea  = document.getElementById('meetingUploadArea');
const meetingFileInput   = document.getElementById('meetingFileInput');
const meetingStatusPanel = document.getElementById('meetingStatusPanel');
const meetingStatusText  = document.getElementById('meetingStatusText');

// ── Drag & drop ─────────────────────────
['dragenter','dragover','dragleave','drop'].forEach(ev =>
  meetingUploadArea.addEventListener(ev, preventDefaults, false));
['dragenter','dragover'].forEach(ev =>
  meetingUploadArea.addEventListener(ev, () => meetingUploadArea.classList.add('dragover'), false));
['dragleave','drop'].forEach(ev =>
  meetingUploadArea.addEventListener(ev, () => meetingUploadArea.classList.remove('dragover'), false));

meetingUploadArea.addEventListener('drop', e => {
  if (e.dataTransfer.files.length > 0) handleMeetingFile(e.dataTransfer.files[0]);
});
meetingFileInput.addEventListener('change', e => {
  if (e.target.files.length > 0) handleMeetingFile(e.target.files[0]);
});

// ── File handler ─────────────────────────
async function handleMeetingFile(file) {
  meetingFileName = file.name;
  meetingUploadArea.style.display  = 'none';
  meetingStatusPanel.style.display = 'block';
  meetingStatusText.textContent    = 'Reading file...';
  meetingStatusText.style.color    = '';
  document.querySelector('.meeting-spinner').style.display = 'block';
  document.getElementById('meetingDownloadBtnContainer').style.display = 'none';
  document.getElementById('meetingTableContainer').style.display       = 'none';

  try {
    const arrayBuffer = await file.arrayBuffer();
    
    // 1. Load via SheetJS (for fast JSON data)
    const wbX = XLSX.read(new Uint8Array(arrayBuffer), { type:'array' });
    
    // 2. Load via ExcelJS (for images)
    const workbookE = new ExcelJS.Workbook();
    await workbookE.xlsx.load(arrayBuffer);
    
    // Process
    processMeetingWorkbook(wbX, workbookE);
  } catch(err) {
    console.error(err);
    meetingStatusText.textContent = 'Error: ' + err.message;
    meetingStatusText.style.color = '#ef4444';
    document.querySelector('.meeting-spinner').style.display = 'none';
  }
}

// ── Smart workbook processor ─────────────
function processMeetingWorkbook(wb, workbookE) {
  meetingWorkbook = wb;
  const ws    = wb.Sheets[wb.SheetNames[0]];
  const raw2d = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // ── STEP 1: Find the real header row (Scoring-based) ─────────────────
  let hIdx = 0, maxScore = -1;
  const keywords = ['REGION','AREA','DIVISION','WORK','LOCATION','BRANCH','TIME','COMPLETED','IDENTIFIER','STAFF','ID','CODE'];

  for (let i = 0; i < Math.min(raw2d.length, 30); i++) {
    const row = (raw2d[i] || []).map(c => String(c).toUpperCase().trim());
    let score = 0;
    keywords.forEach(kw => { if (row.some(v => v.includes(kw))) score++; });
    if (score > maxScore) { maxScore = score; hIdx = i; }
    if (score >= 4) { hIdx = i; break; }
  }

  // ── STEP 2: Parse from real header row ───────────────────────────────
  const jsonData = XLSX.utils.sheet_to_json(ws, { range: hIdx, defval: '' });
  if (!jsonData.length) {
    meetingStatusText.textContent = 'No data found in the file.';
    meetingStatusText.style.color = '#f59e0b';
    document.querySelector('.meeting-spinner').style.display = 'none';
    return;
  }

  meetingAllColumns = Object.keys(jsonData[0] || {});

  // ── STEP 3: Identify key columns ─────────────────────────────────────
  const findCol = (tests) =>
    meetingAllColumns.find(c => tests.some(t =>
      c.toUpperCase().replace(/[\s_\-. ]/g,'').includes(t.toUpperCase())
    )) || '';

  let timeCol     = findCol(['COMPLETEDTIME','COMPLETED TIME','COMPLETED']);
  if (!timeCol) timeCol = findCol(['TIME','PUNCHIN','PUNCHIN TIME','PUNCH']);

  let idCol       = findCol(['IDENTIFIER','STAFFCODE','SCODE','EMPLOYEEID','ID','CODE','SL.NO']);
  let divisionCol = findCol(['DIVISION','DIV.','DIV']);
  let areaCol     = findCol(['AREA']);
  let regionCol   = findCol(['REGION','REG.','REG']);
  let workLocCol  = findCol(['WORKLOCATION','WORKLOC','LOCATION','BRANCH','LOC.','LOC','OFFICE']);

  let b730Col     = findCol(['BEFORE7','BEFORE730']);
  let a730Col     = findCol(['7:30TO8','730TO830','7:30-8:30']);
  let a830Col     = findCol(['8:30AND','8:30ABV','ABOVE8','8:30-']);
  let totalCol    = findCol(['GRANDTOTAL','TOTAL','G.TOTAL']);
  let photoCol    = findCol(['PHOTO','IMAGE','PIC','METTINGPHOTO','MEETINGPHOTO']);
  let nameCol     = findCol(['NAME','STAFFNAME','EMPLOYEENAME', 'FULLNAME', 'STAFF', 'EMPLOYEE']);

  // ── STEP 4: Position-based fallback ──────────────────────────────────
  if (!regionCol || !workLocCol) {
    const colPositions = {};
    for (let rIdx = 0; rIdx < Math.min(jsonData.length, 10); rIdx++) {
      const row = jsonData[rIdx];
      meetingAllColumns.forEach((col, ci) => {
        const v = String(row[col] || '').trim();
        if (!v || !isNaN(Number(v))) return;
        const vu = v.toUpperCase();
        if (vu === 'REGION' || vu === 'AREA' || vu.includes('LOCATION') || vu.includes('WORK')) return;
        if (!colPositions[ci]) colPositions[ci] = 0;
        colPositions[ci]++;
      });
    }
    const textCols = Object.entries(colPositions)
      .filter(([,cnt]) => cnt > 0)
      .sort((a,b) => Number(a[0]) - Number(b[0]))
      .map(([i]) => meetingAllColumns[Number(i)]);

    if (!regionCol  && textCols[0]) regionCol  = textCols[0];
    if (!areaCol    && textCols[1]) areaCol    = textCols[1];
    if (!workLocCol && textCols[2]) workLocCol = textCols[2];
  }

  // Helper: parse time
  const parseTimeValue = (val) => {
    if (typeof val === 'number') return val * 24 * 60;
    if (!val) return null;
    const s = String(val).trim().toUpperCase();
    const match = s.match(/(\d+)[:.](\d+)(?::\d+)?\s*(AM|PM)?/);
    if (!match) return null;
    let h = parseInt(match[1]), m = parseInt(match[2]);
    if (match[3] === 'PM' && h < 12) h += 12;
    if (match[3] === 'AM' && h === 12) h = 0;
    return h * 60 + m;
  };

  // Image/Link extraction using ExcelJS
  const imgMap = {};
  const linkMap = {};
  const photoColIdx = meetingAllColumns.indexOf(photoCol);

  try {
    const worksheetE = workbookE.getWorksheet(1);
    const media = workbookE.model.media || [];
    
    // 1. Map embedded images
    worksheetE.getImages().forEach(img => {
      const v = media[img.imageId];
      if (v && v.buffer) {
        // Mime-type mapping
        const mimeType = v.type === 'jpg' ? 'image/jpeg' : `image/${v.type}`;
        
        let binary = '';
        const bytes = new Uint8Array(v.buffer);
        for (let j = 0; j < bytes.byteLength; j++) {
          binary += String.fromCharCode(bytes[j]);
        }
        const b64 = `data:${mimeType};base64,${btoa(binary)}`;
        
        const r = Math.floor(img.range.tl.row);
        const c = Math.floor(img.range.tl.col);
        if (!imgMap[r]) imgMap[r] = {};
        imgMap[r][c] = b64;
      }
    });

    // 2. Map hyperlinks in the photo column
    if (photoColIdx !== -1) {
      for (let i = 0; i < jsonData.length; i++) {
        const excelRowIdx = hIdx + 2 + i; // 1-indexed row in Excel
        const cell = worksheetE.getRow(excelRowIdx).getCell(photoColIdx + 1);
        if (cell.hyperlink) {
          let target = typeof cell.hyperlink === 'string' ? cell.hyperlink : cell.hyperlink.target;
          if (target) {
            // Attempt to transform SharePoint/Teams links to "Download" direct links
            if (target.includes('sharepoint.com') && !target.includes('download=1')) {
              target += (target.includes('?') ? '&' : '?') + 'download=1';
            }
            linkMap[i] = target;
          }
        }
      }
    }
  } catch(e) { console.warn("Resource extraction failed:", e); }

  // Fill-down and categorize
  let curRegion = '', curArea = '';
  jsonData.forEach((row, i) => {
    // Priority 1: Hyperlink in the photo cell
    if (linkMap[i]) {
      row.mPhoto = linkMap[i];
    } 
    // Priority 2: Embedded image anchored to this row (+/- 1 row tolerance)
    else {
      const excelRow0 = hIdx + 1 + i; // 0-indexed row for mapping
      const pIdx = meetingAllColumns.indexOf(photoCol);
      // Try exact row, row above, or row below
      [excelRow0, excelRow0 - 1, excelRow0 + 1].some(r => {
        if (imgMap[r]) {
          // If we know the photo column, try that first
          if (pIdx !== -1 && imgMap[r][pIdx]) {
            row.mPhoto = imgMap[r][pIdx];
            return true;
          }
          // Otherwise take the first image found in this row
          const firstImg = Object.values(imgMap[r])[0];
          if (firstImg) {
            row.mPhoto = firstImg;
            return true;
          }
        }
        return false;
      });
    }

    const r = String(row[regionCol] || '').trim();
    const a = String(row[areaCol]   || '').trim();
    const rUp = r.toUpperCase(), aUp = a.toUpperCase();
    if (r && rUp !== 'REGION' && rUp !== 'GRAND TOTAL') curRegion = r; else if (regionCol) row[regionCol] = curRegion;
    if (a && aUp !== 'AREA' && aUp !== 'DIVISION' && aUp !== 'GRAND TOTAL') curArea = a; else if (areaCol) row[areaCol] = curArea;

    if (timeCol) {
      const mins = parseTimeValue(row[timeCol]);
      if (mins === null) row.mBucket = '';
      else if (mins <= 7 * 60 + 30) row.mBucket = 'Before 7:30';
      else if (mins <= 8 * 60 + 30) row.mBucket = '7:30 to 8:30';
      else row.mBucket = '8:30 and above';
    }
  });

  mCols = {
    region   : regionCol,
    area     : areaCol,
    division : divisionCol,
    workLoc  : workLocCol,
    time     : timeCol,
    id       : idCol,
    b730     : b730Col,
    a730_830 : a730Col,
    a830     : a830Col,
    total    : totalCol,
    photo    : photoCol,
    name     : nameCol
  };

  const dataRows = jsonData.filter(row => {
    const rg = String(row[regionCol]  || '').trim().toUpperCase();
    const wl = String(row[workLocCol] || '').trim().toUpperCase();
    const exclude = ['HEAD OFFICE', 'CORPORATE OFFICE'];
    if (exclude.some(ex => wl.includes(ex) || rg.includes(ex))) return false;
    if (!rg && !wl) return false;
    if (rg === 'REGION' || wl === 'WORK LOCATION') return false;
    return true;
  });

  meetingAllRows = dataRows;

  // ── STEP 7: Render ────────────────────────────────────────────────────
  renderMeetingTabs('region');

  let statusMsg = `Done! ${dataRows.length} records loaded.`;
  if (timeCol)  statusMsg += ` | Using Time Formula on [${timeCol}]`;
  if (idCol)    statusMsg += ` | Distinct ID: [${idCol}]`;

  meetingStatusText.textContent = statusMsg;
  meetingStatusText.style.color = 'var(--success)';
  document.querySelector('.meeting-spinner').style.display = 'none';
  document.getElementById('meetingDownloadBtnContainer').style.display = 'flex';
}

function mIdDistinctCount(rows, tc) {
  if (!tc) return 0;
  let matches = [];

  const bucketLabels = ['Before 7:30', '7:30 to 8:30', '8:30 and above'];

  // If using raw Time Formula
  if (mCols.time) {
    if (bucketLabels.includes(tc.label)) {
      matches = rows.filter(r => r.mBucket === tc.label);
    } else if (tc.label === 'Grand Total' || tc.key === '__count__') {
      // For Grand Total, count all rows that were categorized
      matches = rows.filter(r => r.mBucket && r.mBucket !== '');
    }
  } 
  // If using pre-pivoted numeric columns
  else {
    if (tc.key === '__count__' || tc.label === 'Grand Total') {
      matches = rows;
    } else {
      matches = rows.filter(r => (parseFloat(r[tc.key]) || 0) > 0);
    }
  }

  if (!matches.length) return 0;
  if (!mCols.id) return matches.length;

  // Unique IDs
  const set = new Set();
  matches.forEach(r => {
    const val = String(r[mCols.id] || '').trim();
    if (val) set.add(val);
  });
  return set.size;
}

function buildMeetingPivot(groupCols, labels, rows) {
  const map = {};
  const gCols = Array.isArray(groupCols) ? groupCols : [groupCols];
  const gLabs = Array.isArray(labels) ? labels : [labels];

  rows.forEach(row => {
    // Key is a joined string of all group columns
    const key = gCols.map(col => String(row[col] || '(Blank)').trim() || '(Blank)').join(' | ');
    if (!map[key]) map[key] = [];
    map[key].push(row);
  });

  const keys = Object.keys(map).filter(k => !k.includes('(Blank)')).sort();
  // Handle blanks at the end
  const blankKeys = Object.keys(map).filter(k => k.includes('(Blank)')).sort();
  keys.push(...blankKeys);

  const timeCols = [];
  if (mCols.time) {
    timeCols.push({ key: 'B730', label: 'Before 7:30' });
    timeCols.push({ key: 'A730', label: '7:30 to 8:30' });
    timeCols.push({ key: 'A830', label: '8:30 and above' });
    timeCols.push({ key: 'ROW_TOTAL', label: 'Grand Total' });
  } else {
    if (mCols.b730)     timeCols.push({ key: mCols.b730,     label: 'Before 7:30 AM'  });
    if (mCols.a730_830) timeCols.push({ key: mCols.a730_830, label: '7:30 to 8:30 AM' });
    if (mCols.a830)     timeCols.push({ key: mCols.a830,     label: '8:30 AM & Above'  });
    if (mCols.total)    timeCols.push({ key: mCols.total,    label: 'Grand Total'       });
  }
  if (!timeCols.length) timeCols.push({ key: '__count__', label: 'Total Records' });

  const colSpan   = gLabs.length + timeCols.length;

  const today = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
  const totals = {};
  timeCols.forEach(tc => totals[tc.key] = 0);

  let tbody = '';
  keys.forEach(k => {
    const grpRows = map[k];
    const kParts = k.split(' | ');
    let labelCells = '';
    kParts.forEach(val => {
      labelCells += `<td style="font-weight:800;color:#000;text-align:left;min-width:140px;">${val}</td>`;
    });

    let dataCells = '';
    timeCols.forEach(tc => {
      const v = mIdDistinctCount(grpRows, tc);
      totals[tc.key] += v;
      const isGrand = tc.label === 'Grand Total';
      const style = isGrand ? 'font-weight:800;background:#f0fdf4;' : '';
      dataCells += `<td style="${style}">${v || 0}</td>`;
    });
    
    tbody += `<tr>${labelCells}${dataCells}</tr>`;
  });

  const totalsCells = timeCols.map(tc =>
    `<td style="font-weight:700;">${Math.round(totals[tc.key])}</td>`
  ).join('');

  return `
    <table class="pivot-table" id="meetingPivotTable">
      <thead>
        <tr class="header-main" style="background:#00ff00;color:#000;">
          <th colspan="${colSpan}">Meeting Report — Hierarchical Summary &nbsp;&nbsp; ${today}</th>
        </tr>
        <tr class="header-sub">
          ${gLabs.map(l => `<th style="text-align:left;">${l}</th>`).join('')}
          ${timeCols.map(tc => `<th>${tc.label}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${tbody}
        <tr style="background:#f0fdf4;">
          <td colspan="${gLabs.length}" style="font-weight:700;text-align:center;">Column Grand Total</td>
          ${totalsCells}
        </tr>
      </tbody>
    </table>`;
}

// ── Three-tab renderer ────────────────────
window.renderMeetingTabs = function(activeTab) {
  const report    = document.getElementById('meetingTableReport');
  const container = document.getElementById('meetingTableContainer');
  const rows      = meetingAllRows;

  // Define tabs
  const allTabs = [];
  
  // Master hierarchical view in specific user order: DIVISION., AREA, REGION, Work Location
  if (mCols.division || mCols.area || mCols.region || mCols.workLoc) {
    const masterCols = [];
    const masterLabs = [];
    if (mCols.division) { masterCols.push(mCols.division); masterLabs.push('DIVISION.'); }
    if (mCols.area)     { masterCols.push(mCols.area);     masterLabs.push('AREA'); }
    if (mCols.region)   { masterCols.push(mCols.region);   masterLabs.push('REGION'); }
    if (mCols.workLoc)  { masterCols.push(mCols.workLoc);  masterLabs.push('Work Location'); }
    
    allTabs.push({ 
      id: 'master', 
      label: '🗺️ Master View', 
      cols: masterCols, 
      labs: masterLabs 
    });
  }

  // Individual tabs
  if (mCols.region)   allTabs.push({ id: 'region',   label: '📍 Region Wise',          cols: [mCols.region],   labs: ['REGION'] });
  if (mCols.area)     allTabs.push({ id: 'area',     label: '🏢 Area Wise',            cols: [mCols.area],     labs: ['AREA'] });
  if (mCols.division) allTabs.push({ id: 'division', label: '📊 Division Wise',        cols: [mCols.division], labs: ['DIVISION.'] });
  if (mCols.workLoc)  allTabs.push({ id: 'workloc',  label: '📌 Work Location Wise',    cols: [mCols.workLoc],  labs: ['Work Location'] });

  if (!allTabs.length) {
    report.innerHTML = '<p style="color:#ef4444;padding:1rem;">Could not detect Region/Area/Work Location columns in this file.</p>';
    container.style.display = 'block';
    return;
  }

  // Set default to master if not specified or invalid
  const activeTabId = (activeTab && allTabs.some(t=>t.id===activeTab)) ? activeTab : allTabs[0].id;

  // Tab buttons UI
  let tabHtml = `<div style="display:flex;gap:0.6rem;flex-wrap:wrap;margin-bottom:1.5rem;">`;
  allTabs.forEach(t => {
    const active = t.id === activeTabId;
    tabHtml += `
      <button onclick="renderMeetingTabs('${t.id}')"
        style="padding:0.65rem 1.4rem;border-radius:20px;font-weight:700;font-size:0.92rem;
               cursor:pointer;border:2px solid #8b5cf6;transition:all 0.2s;font-family:inherit;
               ${active ? 'background:#8b5cf6;color:#fff;box-shadow:0 4px 12px rgba(139,92,246,.35);'
                        : 'background:#fff;color:#8b5cf6;'}">
        ${t.label}
      </button>`;
  });
  tabHtml += `</div>`;

  const activeObj = allTabs.find(t => t.id === activeTabId);
  const pivotHtml = buildMeetingPivot(activeObj.cols, activeObj.labs, rows);

  // Update Report (Tabs + Table)
  report.innerHTML = tabHtml + `<div id="meetingPivotCaptureArea">${pivotHtml}</div>`;
  
  // Show/Reposition Buttons
  const btnContainer = document.getElementById('meetingDownloadBtnContainer');
  btnContainer.style.display = 'flex';

  // Update Gallery separately
  const gallery = document.getElementById('meetingPhotoGallery');
  gallery.innerHTML = renderMeetingPhotoGallery(activeObj, rows);

  container.style.display = 'block';
};

function renderMeetingPhotoGallery(tabObj, rows) {
  // De-duplicate photos based on the source link/data
  const seen = new Set();
  const photoRows = rows.filter(r => {
    if (!r.mPhoto) return false;
    if (seen.has(r.mPhoto)) return false;
    seen.add(r.mPhoto);
    return true;
  });

  if (!photoRows.length) return '';

  // Deeply nested grouping: Region > Division > Area
  const hierarchy = {};

  photoRows.forEach(row => {
    const reg = String(row[mCols.region] || '(No Region)').trim();
    const div = String(row[mCols.division] || '(No Division)').trim();
    const area = String(row[mCols.area] || '(No Area)').trim();

    if (!hierarchy[reg]) hierarchy[reg] = {};
    if (!hierarchy[reg][div]) hierarchy[reg][div] = {};
    if (!hierarchy[reg][div][area]) hierarchy[reg][div][area] = [];
    
    hierarchy[reg][div][area].push(row);
  });

  const regions = Object.keys(hierarchy).sort();
  let sectionsHtml = '';

  regions.forEach(reg => {
    let divHtml = '';
    const divisions = Object.keys(hierarchy[reg]).sort();

    divisions.forEach(div => {
      let areaHtml = '';
      const areas = Object.keys(hierarchy[reg][div]).sort();

      areas.forEach(area => {
        let galleryItems = '';
        hierarchy[reg][div][area].forEach(row => {
          const isBase64 = row.mPhoto.startsWith('data:');
          const loc  = row[mCols.workLoc] || 'Location';
          const time = row[mCols.time] || '';
          
          let sourceLabel = 'Web Link';
          if (!isBase64) {
            try {
              const url = new URL(row.mPhoto);
              sourceLabel = url.hostname.replace('www.', '');
            } catch(e) {}
          }

          galleryItems += `
            <div class="photo-card" style="
              background: white; border: 1px solid #f1f5f9; border-radius: 12px; 
              padding: 10px; box-shadow: 0 4px 15px -10px rgba(0,0,0,0.15);
              display: flex; flex-direction: column; align-items: center; width: 160px; gap: 8px; transition: transform 0.2s;">
              <div style="width: 100%; height: 150px; overflow: hidden; border-radius: 8px; background: #fff; display: flex; align-items: center; justify-content: center; position: relative;">
                <img src="${row.mPhoto}" 
                  referrerpolicy="no-referrer"
                  style="max-width: 100%; max-height: 100%; object-fit: cover; cursor: pointer; transition: all 0.3s;"
                  onmouseover="this.style.transform='scale(1.1)'"
                  onmouseout="this.style.transform='scale(1.0)'"
                  onclick="window.open('${row.mPhoto}', '_blank')"
                  onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';">
                
                <div style="display:none; width: 100%; height: 100%; flex-direction: column; align-items: center; justify-content: center; background: #f8fafc; padding: 10px; text-align: center;">
                  <div style="margin-bottom: 8px; color: #a78bfa; background: white; padding: 8px; border-radius: 50%; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>
                  </div>
                  <button onclick="window.open('${row.mPhoto}', '_blank')" 
                    style="background: #8b5cf6; color: white; padding: 5px 12px; border-radius: 15px; border: none; font-weight: 800; font-size: 0.65rem; cursor: pointer; box-shadow: 0 4px 10px rgba(139, 92, 246, 0.3);">
                    CLICK NOW
                  </button>
                  <div style="font-size: 0.55rem; margin-top: 6px; color: #94a3b8; font-weight: 700;">${sourceLabel}</div>
                </div>
              </div>
              <div style="width: 100%; text-align: center;">
                <div style="font-weight: 900; font-size: 0.8rem; color: #1e293b; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${loc}">${loc}</div>
                <div style="font-size: 0.65rem; font-weight: 700; color: #a78bfa; margin-top: 2px;">${time}</div>
              </div>
            </div>`;
        });

        areaHtml += `
          <div style="margin-top: 1.5rem; padding-left: 1.5rem; border-left: 3px solid #e2e8f0;">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 1rem;">
               <span style="background: #e0f2fe; color: #0369a1; font-weight: 900; font-size: 0.7rem; padding: 3px 10px; border-radius: 20px; border: 1px solid #bae6fd;">AREA: ${area} (${hierarchy[reg][div][area].length} Photos)</span>
            </div>
            <div style="display: flex; flex-wrap: wrap; gap: 1rem;">
              ${galleryItems}
            </div>
          </div>`;
      });
      
      // Calculate total photos for division
      const divCount = areas.reduce((sum, a) => sum + hierarchy[reg][div][a].length, 0);

      divHtml += `
        <div style="margin-top: 2rem; border-radius: 15px; background: white; border: 1px solid #f1f5f9; padding: 1.5rem; box-shadow: 0 2px 10px rgba(0,0,0,0.02);">
          <h4 style="color: #475569; font-size: 1rem; font-weight: 900; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 12px;">
            <div style="width: 28px; height: 3px; background: #cbd5e1; border-radius: 10px;"></div>
            DIVISION: ${div} <small style="font-size: 0.7rem; color: #94a3b8; font-weight: 700; margin-left: 4px;">(${divCount} Photos)</small>
          </h4>
          ${areaHtml}
        </div>`;
    });

    // Calculate total photos for region
    const regCount = Object.keys(hierarchy[reg]).reduce((sum, d) => {
      const arKeys = Object.keys(hierarchy[reg][d]);
      return sum + arKeys.reduce((s, a) => s + hierarchy[reg][d][a].length, 0);
    }, 0);

    sectionsHtml += `
      <div style="margin-top: 4rem; position: relative;">
        <h3 style="background: linear-gradient(135deg, #6d28d9, #4c1d95); color: white; padding: 1rem 1.8rem; border-radius: 20px 20px 0 0; font-size: 1.4rem; font-weight: 900; display: inline-flex; align-items: center; gap: 12px; box-shadow: 0 10px 30px -10px rgba(109, 40, 217, 0.5);">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
          REGION: ${reg} <span style="font-size: 0.95rem; opacity: 0.8; font-weight: 700; margin-left: 8px;">(${regCount} Photos)</span>
        </h3>
        <div style="background: #f8fafc; border: 2px solid #6d28d9; border-radius: 0 25px 25px 25px; padding: 2rem; padding-top: 1.5rem; position: relative; top: -1px;">
           ${divHtml}
        </div>
      </div>
    `;
  });

  return `
    <div style="margin-top: 5rem; border-top: 8px solid #8b5cf6; padding-top: 3rem;">
      <div style="text-align: center; margin-bottom: 3rem;">
        <h2 style="color: #1e1b4b; font-size: 2.2rem; font-weight: 900; margin-bottom: 0.5rem; letter-spacing: -0.02em; text-transform: uppercase;">
          📸 MEETING PHOTOS
        </h2>
      </div>
      ${sectionsHtml}
    </div>
  `;
}

// ── Download Excel ────────────────────────
document.getElementById('meetingDownloadBtn').addEventListener('click', () => {
  if (!meetingAllRows.length) return;
  const nmWB = XLSX.utils.book_new();
  const nmWS = XLSX.utils.json_to_sheet(meetingAllRows, { header: meetingAllColumns });
  XLSX.utils.book_append_sheet(nmWB, nmWS, 'Meeting Data');
  XLSX.writeFile(nmWB, meetingFileName.replace(/\.[^/.]+$/, '') + '_meeting.xlsx');
});

// ── Download picture ──────────────────────
document.getElementById('meetingDownloadImgBtn').addEventListener('click', () => {
  const el    = document.getElementById('meetingPivotCaptureArea');
  if (!el) return;
  
  const cells = el.querySelectorAll('td, th');
  cells.forEach(c => {
    c.dataset.oc   = c.style.color;
    c.style.color      = '#000000';
    c.style.fontWeight = c.tagName === 'TH' ? '700' : '600';
    c.style.textShadow = '0 0 0.3px #000';
  });
  html2canvas(el, { scale: 3, backgroundColor: '#ffffff', useCORS: true, logging: false })
    .then(canvas => {
      cells.forEach(c => { c.style.color = c.dataset.oc || ''; c.style.textShadow = ''; delete c.dataset.oc; });
      const lnk = document.createElement('a');
      lnk.download = meetingFileName.replace(/\.[^/.]+$/, '') + '_meeting.png';
      lnk.href = canvas.toDataURL('image/png');
      lnk.click();
    });
});

// ── Reset ─────────────────────────────────
document.getElementById('meetingResetBtn').addEventListener('click', () => {
  meetingWorkbook = null; meetingAllRows = []; meetingAllColumns = []; meetingFileName = ''; mCols = {};
  meetingFileInput.value = '';
  document.getElementById('meetingTableContainer').style.display       = 'none';
  document.getElementById('meetingDownloadBtnContainer').style.display = 'none';
  meetingStatusPanel.style.display = 'none';
  meetingUploadArea.style.display  = 'block';
  meetingStatusText.style.color    = '';
});
