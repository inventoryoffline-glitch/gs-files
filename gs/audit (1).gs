// ============================================================
// audit.gs — Enhanced Server-Side Audit with Additional Checks
// ============================================================

function runServerAudit() {
  const data = getAllData();
  const results = [];

  // ── EXISTING CHECKS ────────────────────────────────────────
  
  // Check 1: No negative stock
  const negStock = data.stock.filter(s => parseInt(s.qty) < 0);
  results.push({
    check: "No Negative Stock",
    status: negStock.length === 0 ? "pass" : "fail",
    detail: negStock.length === 0 ? "All items OK" : negStock.map(s => s.name + " (" + s.qty + ")").join(", ")
  });

  // Check 2: Allocation dispatchedQty never exceeds plannedQty
  const overAlloc = data.distAllocations.filter(a => parseInt(a.dispatchedQty) > parseInt(a.plannedQty));
  results.push({
    check: "Allocations Within Planned Limits",
    status: overAlloc.length === 0 ? "pass" : "warn",
    detail: overAlloc.length === 0 ? "All allocations OK" : overAlloc.length + " over-dispatched allocations"
  });

  // Check 3: No duplicate active vouchers across different institutes
  const activeRecs = data.records.filter(r => (r.status||"active") === "active");
  const vMap = {};
  activeRecs.forEach(r => {
    if (!vMap[r.voucherNo]) vMap[r.voucherNo] = new Set();
    vMap[r.voucherNo].add(r.instId);
  });
  const dups = Object.entries(vMap).filter(([v, s]) => s.size > 1).map(([v]) => v);
  results.push({
    check: "No Duplicate Voucher Numbers",
    status: dups.length === 0 ? "pass" : "fail",
    detail: dups.length === 0 ? "All vouchers unique" : "Duplicates: " + dups.join(", ")
  });

  // ── NEW ENHANCED CHECKS ────────────────────────────────────

  // Check 4: Stock Consistency — Dispatched count matches records
  const stockMismatch = [];
  data.stock.forEach(s => {
    const recDispatched = data.records
      .filter(r => r.itemId === s.id && (r.status || "active") === "active")
      .reduce((sum, r) => sum + (parseInt(r.qty) || 0), 0);
    const stockDispatched = parseInt(s.dispatched) || 0;
    if (recDispatched !== stockDispatched) {
      stockMismatch.push(s.name + ": stock=" + stockDispatched + ", records=" + recDispatched);
    }
  });
  results.push({
    check: "Stock Dispatched Count Matches Records",
    status: stockMismatch.length === 0 ? "pass" : "warn",
    detail: stockMismatch.length === 0 ? "All counts match" : stockMismatch.join(" | ")
  });

  // Check 5: Institute Activity — Institutes with no recent activity
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  const inactiveInsts = [];
  data.institutes.forEach(inst => {
    const lastDispatch = data.records
      .filter(r => r.instId === inst.id && (r.status || "active") === "active")
      .sort((a, b) => new Date(b.date) - new Date(a.date))[0];
    if (!lastDispatch || new Date(lastDispatch.date) < thirtyDaysAgo) {
      inactiveInsts.push(inst.name);
    }
  });
  results.push({
    check: "Institute Activity (Last 30 Days)",
    status: inactiveInsts.length === 0 ? "pass" : "info",
    detail: inactiveInsts.length === 0 ? "All institutes active" : inactiveInsts.length + " institutes inactive: " + inactiveInsts.slice(0, 5).join(", ") + (inactiveInsts.length > 5 ? "..." : "")
  });

  // Check 6: Expiry Trends — Items expiring soon
  const today = new Date();
  const expiringInDays = 30;
  const expiringDate = new Date();
  expiringDate.setDate(expiringDate.getDate() + expiringInDays);
  const expiringItems = data.stock.filter(s => {
    if (!s.expiry) return false;
    const expDate = safeParseDate(s.expiry);
    return expDate >= today && expDate <= expiringDate && parseInt(s.qty) > 0;
  });
  results.push({
    check: "Items Expiring in Next 30 Days",
    status: expiringItems.length === 0 ? "pass" : "warn",
    detail: expiringItems.length === 0 ? "No items expiring soon" : expiringItems.length + " item(s) expiring: " + expiringItems.map(s => s.name).slice(0, 5).join(", ") + (expiringItems.length > 5 ? "..." : "")
  });

  // Check 7: Record Integrity — All records reference valid institutes and items
  const orphanRecs = data.records.filter(r => {
    const instOk = data.institutes.find(i => i.id === r.instId);
    const itemOk = data.stock.find(s => s.id === r.itemId);
    return !instOk || !itemOk;
  });
  results.push({
    check: "Record Integrity (Valid References)",
    status: orphanRecs.length === 0 ? "pass" : "warn",
    detail: orphanRecs.length === 0 ? "All records valid" : orphanRecs.length + " records have invalid references"
  });

  // Check 8: Zero or Negative Allocations
  const zeroAllocs = data.distAllocations.filter(a => parseInt(a.plannedQty) <= 0);
  results.push({
    check: "No Zero-Quantity Allocations",
    status: zeroAllocs.length === 0 ? "pass" : "warn",
    detail: zeroAllocs.length === 0 ? "All allocations positive" : zeroAllocs.length + " allocations have zero or negative quantities"
  });

  // Check 9: Cancelled Records Cleanup
  const cancelledCount = data.records.filter(r => (r.status || "active") === "cancelled").length;
  results.push({
    check: "Cancelled Records Tracking",
    status: "info",
    detail: cancelledCount + " cancelled record(s) archived for reference"
  });

  // Check 10: Low Stock Items
  const lowStockItems = data.stock.filter(s => 
    s.showLow !== false && s.showLow !== 'false' && 
    parseInt(s.qty) > 0 && 
    parseInt(s.qty) <= parseInt(s.low || 10) &&
    (!s.expiry || new Date(safeParseDate(s.expiry)) >= new Date())
  );
  results.push({
    check: "Low Stock Alert Items",
    status: lowStockItems.length === 0 ? "pass" : "warn",
    detail: lowStockItems.length === 0 ? "No items below threshold" : lowStockItems.length + " item(s) low: " + lowStockItems.map(s => s.name + " (" + s.qty + ")").slice(0, 5).join(", ") + (lowStockItems.length > 5 ? "..." : "")
  });

  return { success: true, results };
}
