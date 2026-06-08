// ============================================================
// audit.gs — Server-Side Audit
// ============================================================

function runServerAudit() {
  const data    = getAllData();
  const results = [];

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

  return { success: true, results };
}
