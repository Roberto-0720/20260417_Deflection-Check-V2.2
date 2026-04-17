# Fix Virtual Node Displacement Retrieval

## Problem

The Deflection Check Tool shows incorrect results (Deflection=L/0, Max=70468mm for a 6200mm beam) because:

1. **Wrong COM interface**: Code uses `FrameElm.GetPoints()` but SAP2000 COM binding exposes `LineElm.GetPoints()` instead
2. **Wrong coordinate API**: `PointObj.GetCoordCartesian()` returns `(0,0,0)` for virtual mesh nodes (e.g., `~205`) — must use `PointElm.GetCoordCartesian()`
3. **Wrong displacement API**: `Results.JointDispl(name, 0)` returns 0 results for virtual nodes — must use `Results.JointDispl(name, 1)` (`ObjectElm=1` for element-level joints)
4. **DatabaseTables API inaccessible**: All `DatabaseTables` calls fail (`"DatabaseTables"` error), so the current primary displacement method never works — must rely entirely on `Results.JointDispl`

## Diagnostic Evidence

| API Call | Result |
|----------|--------|
| `model.FrameElm.GetPoints("2439-1")` | ❌ `FrameElm` error |
| `model.LineElm.GetPoints("2439-1")` | ✅ `['1909', '~205', 0]` |
| `model.PointObj.GetCoordCartesian("~205")` | ❌ `(0, 0, 0)` |
| `model.PointElm.GetCoordCartesian("~205")` | ✅ `(333.3, 56800.0, 41703.2)` |
| `model.Results.JointDispl("~205", 0)` | ❌ `count=0` |
| `model.Results.JointDispl("~205", 1)` | ✅ `U1=47.58, U2=8.12, U3=-1.03` |
| `model.DatabaseTables.GetTableForDisplayArray(...)` | ❌ `DatabaseTables` error |

## Proposed Changes

### SAP Connector

#### [MODIFY] [sap_connector.py](file:///d:/09 PT Code/14 Deflection Check/utils/sap_connector.py)

**Fix 1 — `get_frame_obj_mesh_points` (line 113-140):**
- Change `self.sap_model.FrameElm.GetPoints()` → `self.sap_model.LineElm.GetPoints()`

**Fix 2 — `get_node_coordinates` (line 142-164):**
- Try `PointElm.GetCoordCartesian` **first** for nodes starting with `~`, then fallback to `PointObj`
- Keep existing fallback chain for non-virtual nodes

**Fix 3 — `get_joint_displacements` (line 190-215):**
- Since `DatabaseTables` is inaccessible, go directly to `JointDispl` method
- For virtual nodes (starting with `~`), use `ObjectElm=1` (element-level)
- For regular nodes, use `ObjectElm=0` (object-level)

**Fix 4 — Remove/skip DatabaseTables as primary method:**
- `_get_disp_via_db_tables` always fails, make `JointDispl` the primary method instead

---

### Main Window

#### [MODIFY] [main_window.py](file:///d:/09 PT Code/14 Deflection Check/ui/main_window.py)

No additional changes needed (COM threading fix already applied).

## Verification Plan

### Manual Verification (User)
1. Close the running app, re-run `python main.py`  
2. Connect to SAP2000, select group `2439`, select load combinations (e.g., LC5001-LC5009)
3. Click **RUN CHECK**
4. Expected: Log shows `Group '2439': X elements, 8 nodes, L=6200mm` (8 nodes including virtual)
5. Expected: Results show reasonable deflection values (not L/0 or 70468mm)
6. Verify by cross-checking with SAP2000's own Joint Displacements table for virtual nodes
