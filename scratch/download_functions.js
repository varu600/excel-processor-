// Global reference for the pivot map to enable download
let lastCollMap = new Map();

function downloadCollectionPivot() {
    if (!lastCollMap || lastCollMap.size === 0) {
        alert("No collection pivot data available to download.");
        return;
    }

    const data = Array.from(lastCollMap.entries()).map(([id, amt]) => ({
        "AccountID": id,
        "Collection Total": amt
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Collection Pivot");
    XLSX.writeFile(wb, `Collection_Pivot_${Date.now()}.xlsx`);
}

function downloadDpdProcessedData() {
    if (!window.processedDpdData) return;
    const ws = XLSX.utils.json_to_sheet(window.processedDpdData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Combined Report");
    XLSX.writeFile(wb, "DPD_Combined_Report_" + new Date().getTime() + ".xlsx");
}
