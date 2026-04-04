const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const statusPanel = document.getElementById('statusPanel');
const statusText = document.getElementById('statusText');
const downloadBtnContainer = document.getElementById('downloadBtnContainer');
const downloadBtn = document.getElementById('downloadBtn');

let processedWorkbook = null;
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
    "TUMKUR"
  ];
  
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

      if (!punchInVal) {
        extractedTime = "";
        remarks = "";
      } else if (punchInVal.toUpperCase() === "A") {
        extractedTime = "A";
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
          const time730 = 7 * 60 + 30; // 450 minutes
          const time830 = 8 * 60 + 30; // 510 minutes

          if (timeInMinutes < time730) {
            remarks = "Before 7:30";
          } else if (timeInMinutes < time830) {
            remarks = "7:30 to 8:30";
          } else {
            remarks = "8:30 and above";
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

  statusText.textContent = "Processing complete! Found " + filteredData.length + " matching rows.";
  document.querySelector('.spinner').style.display = 'none';
  statusText.style.color = "var(--success)";
  downloadBtnContainer.style.display = 'flex';
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

function generatePivotTable(data) {
  const regions = ["APTS", "Dharwad", "Kalburagi", "Tumkur"];
  const stats = {
    "APTS": { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Dharwad": { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Kalburagi": { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 },
    "Tumkur": { total: 0, b730: 0, a730_b830: 0, a830: 0, notMarked: 0 }
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
  html2canvas(pivotEl).then(canvas => {
    const link = document.createElement('a');
    link.download = 'attendance_pivot.png';
    link.href = canvas.toDataURL();
    link.click();
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
