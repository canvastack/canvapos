# Performance Budget — CanvaPOS

## Execution Time Targets

| Function | Max Time (ms) | Instrumentation |
|----------|-------------|-----------------|
| `simpanTransaksi()` — Total | 15,000 | `timeStart/End("simpanTransaksi")` |
|  ├─ Phase 1: Read & Validate | 3,000 | `timeEnd("simpanTransaksi:P1_Read")` |
|  ├─ Phase 2: Compute BOM | 5,000 | `timeEnd("simpanTransaksi:P2_Compute")` |
|  └─ Phase 3: Write (incl. LockService) | 10,000 | `timeEnd("simpanTransaksi:P3_Write")` |
| `refreshLaporan()` | 20,000 | `timeStart/End("refreshLaporan")` |
| `stockEngineBOM()` | 10,000 | `timeStart/End("stockEngineBOM")` |
| `getHPPLookup()` (cache miss) | 5,000 | Logger (built-in) |
| `protectAll()` / `unprotectAll()` | 30,000 | Logger |

## Scaling Limits

| Dimension | Limit | Impact |
|-----------|-------|--------|
| POS order rows / batch | 50 | Phase 1 read < 3s |
| BOM ingredients / product | 20 | Phase 2 compute < 5s |
| Transaksi rows (historical) | 10,000 | refreshLaporan < 20s |
| Stock items | 200 | BOM deduction < 5s |
| Topping types per order | 10 | BOM < 5s |

## Monitoring

All functions log execution time via `Logger.log()`:
- `timeStart(label)` — start timer
- `timeEnd(label)` — end timer, log `⏱ label: Xms`
- `logPerf(label, elapsedMs)` — log custom timing

Enable View → Logs in Apps Script editor to see timing output.

> ⚠ If any function exceeds the Max Time target, a `⚠ PERINGATAN` is logged and the function should be optimized or split into batch operations.
