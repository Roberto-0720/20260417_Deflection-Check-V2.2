"""SAP2000 COM API connector for Deflection Check."""
import time
from typing import List, Dict, Optional, Tuple
from data.models import NodeInfo, FrameElement, DisplacementResult


class SapConnector:
    UNIT_KN_MM = 9

    def __init__(self):
        self.sap_object = None
        self.sap_model = None
        self._connected = False
        self._model_path = ""

    def connect(self) -> bool:
        """
        Connect to a running SAP2000 instance via COM.

        Tries three methods in order (matching PileCapacityCheck V3.4):
          1. SAP2000v1.Helper  → works for V21+ (including V26)
          2. GetActiveObject("CSI.SAP2000.API.SapObject") → V17‑V20
          3. GetActiveObject("SAP2000.cOAPI") → legacy fallback
        """
        import comtypes.client
        connection_errors = []

        # Method 1: SAP2000v1.Helper (v21+, including V26)
        try:
            helper = comtypes.client.CreateObject("SAP2000v1.Helper")
            sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
            self.sap_object = sap_object
            self.sap_model = sap_object.SapModel
            self._connected = True
            self._model_path = self.sap_model.GetModelFilename()
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            self._init_results_setup()
            print("[SAP] Connected via Method 1 (SAP2000v1.Helper)")
            return True
        except Exception as e:
            connection_errors.append(f"Method 1 (SAP2000v1.Helper): {e}")

        # Method 2: GetActiveObject — CSI.SAP2000.API.SapObject (V17‑V20)
        try:
            sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
            self.sap_object = sap_object
            self.sap_model = sap_object.SapModel
            self._connected = True
            self._model_path = self.sap_model.GetModelFilename()
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            self._init_results_setup()
            print("[SAP] Connected via Method 2 (CSI.SAP2000.API.SapObject)")
            return True
        except Exception as e:
            connection_errors.append(f"Method 2 (GetActiveObject CSI): {e}")

        # Method 3: GetActiveObject — SAP2000.cOAPI (legacy fallback)
        try:
            sap_object = comtypes.client.GetActiveObject("SAP2000.cOAPI")
            self.sap_object = sap_object
            self.sap_model = sap_object.SapModel
            self._connected = True
            self._model_path = self.sap_model.GetModelFilename()
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            self._init_results_setup()
            print("[SAP] Connected via Method 3 (SAP2000.cOAPI)")
            return True
        except Exception as e:
            connection_errors.append(f"Method 3 (SAP2000.cOAPI): {e}")

        # All methods failed
        error_detail = "\n  ".join(connection_errors)
        print(f"[SAP] Connection failed — all 3 methods exhausted:\n  {error_detail}")
        self._connected = False
        return False

    def _init_results_setup(self):
        """
        Wake up the SAP2000 results engine with dummy queries.
        SAP2000 COM quirk: the engine is not fully ready until at least one
        results query has been executed.
        """
        try:
            self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()

            # Dummy query with the first available load case
            try:
                ret = self.sap_model.LoadCases.GetNameList()
                if ret[0] > 0:
                    first_case = str(ret[1][0])
                    self.sap_model.Results.Setup.SetCaseSelectedForOutput(first_case)
                    try:
                        ret2 = self.sap_model.PointObj.GetNameList()
                        if ret2[0] > 0:
                            self.sap_model.Results.JointDispl(str(ret2[1][0]), 0)
                    except Exception:
                        pass
                    self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
            except Exception:
                pass

            # Also try with first combo
            try:
                ret = self.sap_model.RespCombo.GetNameList()
                if ret[0] > 0:
                    first_combo = str(ret[1][0])
                    self.sap_model.Results.Setup.SetComboSelectedForOutput(first_combo)
                    try:
                        ret2 = self.sap_model.PointObj.GetNameList()
                        if ret2[0] > 0:
                            self.sap_model.Results.JointDispl(str(ret2[1][0]), 0)
                    except Exception:
                        pass
                    self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
            except Exception:
                pass

            print("[SAP] Results engine initialized")
        except Exception as e:
            print(f"[SAP] Warning during results init: {e}")

    def _warmup_in_current_thread(self):
        """
        Lightweight warm-up for the results engine in the CURRENT thread.

        Must be called from the background worker thread before any real
        JointDispl queries. _init_results_setup() runs on the main thread
        during connect(); this ensures the background thread COM context
        is also ready.
        """
        try:
            self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
            ret_cases = self.sap_model.LoadCases.GetNameList()
            if ret_cases[0] > 0:
                first_case = str(ret_cases[1][0])
                self.sap_model.Results.Setup.SetCaseSelectedForOutput(first_case)
                ret_nodes = self.sap_model.PointObj.GetNameList()
                if ret_nodes[0] > 0:
                    self.sap_model.Results.JointDispl(str(ret_nodes[1][0]), 0)
            self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        except Exception:
            pass

    @property
    def is_connected(self) -> bool:
        return self._connected

    @property
    def model_path(self) -> str:
        return self._model_path

    @property
    def model_directory(self) -> str:
        import os
        return os.path.dirname(self._model_path) if self._model_path else ""

    # =========================================================================
    # Group / Frame queries
    # =========================================================================
    def get_group_names(self) -> List[str]:
        if not self._connected:
            return []
        try:
            ret = self.sap_model.GroupDef.GetNameList()
            return [str(n) for n in ret[1]] if ret[0] > 0 else []
        except Exception as e:
            print(f"[SAP] Error getting groups: {e}")
            return []

    def get_group_frames(self, group_name: str) -> List[str]:
        """Get all frame element names assigned to a group."""
        if not self._connected:
            return []
        try:
            ret = self.sap_model.GroupDef.GetAssignments(group_name)
            if ret[0] == 0:
                return []
            # ObjectType 2 = Frame
            return [str(ret[2][i]) for i in range(ret[0]) if ret[1][i] == 2]
        except Exception as e:
            print(f"[SAP] Error getting group frames: {e}")
            return []

    def get_group_joints(self, group_name: str) -> List[str]:
        """Get all joint names assigned to a group."""
        if not self._connected:
            return []
        try:
            ret = self.sap_model.GroupDef.GetAssignments(group_name)
            if ret[0] == 0:
                return []
            return [str(ret[2][i]) for i in range(ret[0]) if ret[1][i] == 1]
        except Exception as e:
            print(f"[SAP] Error: {e}")
            return []

    def get_frame_info(self, frame_name: str) -> Optional[FrameElement]:
        """Get frame element connectivity and section info."""
        if not self._connected:
            return None
        try:
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            ret_pts = self.sap_model.FrameObj.GetPoints(frame_name)
            joint_i = str(ret_pts[0])
            joint_j = str(ret_pts[1])
            ret_sec = self.sap_model.FrameObj.GetSection(frame_name)
            section = str(ret_sec[0])
            ci = self.sap_model.PointObj.GetCoordCartesian(joint_i)
            cj = self.sap_model.PointObj.GetCoordCartesian(joint_j)
            dx = float(cj[0]) - float(ci[0])
            dy = float(cj[1]) - float(ci[1])
            dz = float(cj[2]) - float(ci[2])
            length = (dx**2 + dy**2 + dz**2)**0.5
            return FrameElement(
                name=frame_name, joint_i=joint_i, joint_j=joint_j,
                section=section, length=length
            )
        except Exception as e:
            print(f"[SAP] Error getting frame info for {frame_name}: {e}")
            return None

    def is_frame_beam(self, frame_name: str, max_angle_deg: float = 15.0,
                      max_plan_skew_deg: float = 15.0) -> bool:
        """
        Determine if a frame is a beam based on geometry.
        Two conditions must both be satisfied:
          1. Near-horizontal: angle from XY plane <= max_angle_deg
             (eliminates columns ~90° and inclined braces 30-60°)
          2. Plan projection is parallel to X or Y axis within max_plan_skew_deg
             (eliminates horizontal diagonal bracing)
        """
        if not self._connected:
            return False
        try:
            import math
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            ret_pts = self.sap_model.FrameObj.GetPoints(frame_name)
            ji, jj = str(ret_pts[0]), str(ret_pts[1])
            ci = self.sap_model.PointObj.GetCoordCartesian(ji)
            cj = self.sap_model.PointObj.GetCoordCartesian(jj)
            dx = float(cj[0]) - float(ci[0])
            dy = float(cj[1]) - float(ci[1])
            dz = float(cj[2]) - float(ci[2])
            length = math.sqrt(dx**2 + dy**2 + dz**2)
            if length < 1e-6:
                return False

            # Condition 1: near-horizontal (angle with XY plane <= max_angle_deg)
            elev_angle = math.degrees(math.asin(min(abs(dz) / length, 1.0)))
            if elev_angle > max_angle_deg:
                return False

            # Condition 2: plan projection parallel to X or Y axis
            # Compute horizontal projection length
            horiz_len = math.sqrt(dx**2 + dy**2)
            if horiz_len < 1e-6:
                # Perfectly vertical → already caught above, but guard anyway
                return False
            # Angle of plan projection from the nearest axis (X or Y):
            # min(|dx|, |dy|) / horiz_len = sin(skew angle from nearest axis)
            plan_skew = math.degrees(math.asin(min(abs(dx), abs(dy)) / horiz_len))
            return plan_skew <= max_plan_skew_deg

        except Exception:
            return False

    def get_frame_obj_mesh_points(self, frame_name: str):
        """
        Get all element-level joint names along a frame object,
        including virtual/mesh nodes from Auto Frame Mesh.
        Returns (ordered_joint_names, elm_joint_map)
        where elm_joint_map = {elm_name: (joint_i, joint_j)}
        """
        if not self._connected:
            return [], {}
        try:
            ret = self.sap_model.FrameObj.GetElm(frame_name)
            n_elm = ret[0]
            elm_names = ret[1] if n_elm > 0 else []

            elm_joint_map = {}
            all_joints = []
            for elm_name in elm_names:
                ret_pts = self.sap_model.LineElm.GetPoints(str(elm_name))
                ji = str(ret_pts[0])
                jj = str(ret_pts[1])
                elm_joint_map[str(elm_name)] = (ji, jj)
                if not all_joints:
                    all_joints.append(ji)
                all_joints.append(jj)
            return all_joints, elm_joint_map
        except Exception as e:
            print(f"[SAP] Error getting mesh points for {frame_name}: {e}")
            return [], {}

    def get_node_coordinates(self, node_names: List[str]) -> Dict[str, NodeInfo]:
        nodes = {}
        if not self._connected:
            return nodes
        self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
        for name in node_names:
            try:
                # Virtual mesh nodes (prefixed with ~) must use PointElm
                # PointObj.GetCoordCartesian silently returns (0,0,0) for ~ nodes
                if name.startswith('~'):
                    ret = self.sap_model.PointElm.GetCoordCartesian(name)
                else:
                    ret = self.sap_model.PointObj.GetCoordCartesian(name)
                nodes[name] = NodeInfo(name=name,
                                       x=float(ret[0]),
                                       y=float(ret[1]),
                                       z=float(ret[2]))
            except Exception:
                # Fallback: try the other API
                try:
                    if name.startswith('~'):
                        ret = self.sap_model.PointObj.GetCoordCartesian(name)
                    else:
                        ret = self.sap_model.PointElm.GetCoordCartesian(name)
                    nodes[name] = NodeInfo(name=name,
                                           x=float(ret[0]),
                                           y=float(ret[1]),
                                           z=float(ret[2]))
                except Exception as e:
                    print(f"[SAP] Cannot get coords for {name}: {e}")
        return nodes

    # =========================================================================
    # Load Case / Combo queries
    # =========================================================================
    def get_load_case_names(self) -> List[str]:
        if not self._connected:
            return []
        try:
            ret = self.sap_model.LoadCases.GetNameList()
            return [str(n) for n in ret[1]] if ret[0] > 0 else []
        except Exception:
            return []

    def get_combo_names(self) -> List[str]:
        if not self._connected:
            return []
        try:
            ret = self.sap_model.RespCombo.GetNameList()
            return [str(n) for n in ret[1]] if ret[0] > 0 else []
        except Exception:
            return []

    # =========================================================================
    # Displacement Results - Results.JointDispl (primary) + DatabaseTables fallback
    # =========================================================================
    def get_joint_displacements(
        self,
        node_names: List[str],
        load_cases: Optional[List[str]] = None,
        load_combos: Optional[List[str]] = None,
        elm_joint_map: Optional[Dict] = None,
    ) -> List[DisplacementResult]:
        """
        Get joint displacement results for the given nodes and load cases.

        Uses Results.JointDispl as the primary method — this works for BOTH
        PointObj (model joints) and PointElm (auto-mesh intermediate joints).
        DatabaseTables only covers PointObj, so it would miss intermediate
        mesh nodes that are critical for computing beam deflection.
        """
        if not self._connected:
            return []

        self.sap_model.SetPresentUnits(self.UNIT_KN_MM)

        # CRITICAL: Wake up the results engine in this thread before querying.
        # _init_results_setup() runs on the main thread at connect() time;
        # this warm-up is needed for the background worker thread's COM context.
        self._warmup_in_current_thread()

        # STEP 1: Clear all output selections
        self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()

        # STEP 2: Select load cases for output
        lc_count = 0
        if load_cases:
            for lc in load_cases:
                ret = self.sap_model.Results.Setup.SetCaseSelectedForOutput(lc)
                if ret == 0:
                    lc_count += 1
                else:
                    print(f"[SAP] Warning: Could not select LC '{lc}' (ret={ret})")

        # STEP 3: Select load combinations for output
        lcb_count = 0
        if load_combos:
            for lc in load_combos:
                ret = self.sap_model.Results.Setup.SetComboSelectedForOutput(lc)
                if ret == 0:
                    lcb_count += 1
                else:
                    print(f"[SAP] Warning: Could not select LCB '{lc}' (ret={ret})")

        print(f"[SAP] Output selection: {lc_count} LCs, {lcb_count} LCBs")

        if lc_count == 0 and lcb_count == 0:
            print("[SAP] ERROR: No load cases/combos selected for output!")
            return []

        # STEP 4: Brief pause for SAP to process selection
        time.sleep(0.2)

        # STEP 5: Query displacements with one automatic retry on 0 records
        results = []
        for attempt in range(2):
            attempt_results = []
            for node_name in node_names:
                try:
                    # ObjectElm: 0=object-level (PointObj), 1=element-level (PointElm/~nodes)
                    # Virtual mesh nodes (~xxx) only return results with ObjectElm=1
                    obj_elm = 1 if node_name.startswith('~') else 0
                    ret = self.sap_model.Results.JointDispl(node_name, obj_elm)
                    n = ret[0]
                    if n == 0:
                        continue
                    for i in range(n):
                        case_name = str(ret[3][i])
                        case_type = ""
                        try:
                            case_type = str(ret[4][i]) if ret[4] else ""
                        except Exception:
                            pass
                        attempt_results.append(DisplacementResult(
                            node_name=node_name,
                            load_case=case_name,
                            case_type=case_type,
                            ux=float(ret[6][i]),
                            uy=float(ret[7][i]),
                            uz=float(ret[8][i])
                        ))
                except Exception as e:
                    print(f"[SAP] Error getting displacements for {node_name}: {e}")

            if attempt_results:
                results = attempt_results
                break

            if attempt == 0:
                # Re-apply selections and retry after 1 second
                print("[SAP] 0 records on first attempt, retrying...")
                self.sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
                if load_cases:
                    for lc in load_cases:
                        self.sap_model.Results.Setup.SetCaseSelectedForOutput(lc)
                if load_combos:
                    for lc in load_combos:
                        self.sap_model.Results.Setup.SetComboSelectedForOutput(lc)
                time.sleep(1.0)

        if not results:
            print("[SAP] WARNING: 0 displacement records retrieved!")
            # Last resort: try DatabaseTables
            desired_lcs = set(load_cases or []) | set(load_combos or [])
            results = self._get_disp_via_db_tables(set(node_names), desired_lcs)
            if results:
                print(f"[SAP] DatabaseTables fallback: {len(results)} records")

        # STEP 6: Get displacements for PointElm (auto-mesh) nodes via FrameJointDispl.
        # Results.JointDispl() only works for PointObj joints; auto-mesh intermediate
        # nodes (~xxx) require Results.FrameJointDispl() queried by element.
        if elm_joint_map:
            nodes_with_data = {r.node_name for r in results}
            missing = {j for pair in elm_joint_map.values()
                       for j in pair if j not in nodes_with_data}
            if missing:
                elm_results = self._get_disp_via_frame_joint_displ(elm_joint_map, missing)
                print(f"[SAP] FrameJointDispl: {len(elm_results)} records for {len(missing)} PointElm nodes")
                results.extend(elm_results)

        return results

    def _get_disp_via_frame_joint_displ(
        self,
        elm_joint_map: Dict,
        missing_joints: set,
    ) -> List[DisplacementResult]:
        """
        Get displacements for PointElm joints (auto-mesh intermediate nodes)
        via Results.FrameJointDispl(), queried per element.

        SAP2000 returns two records per element per load case: I-end first,
        J-end second. We validate ordering using PointObj joints where
        possible (boundary elements whose I or J IS a PointObj with known data).
        For interior elements (both joints are PointElm), we trust the order.
        """
        results = []
        from collections import defaultdict

        for elm_name, (joint_i, joint_j) in elm_joint_map.items():
            need_i = joint_i in missing_joints
            need_j = joint_j in missing_joints
            if not need_i and not need_j:
                continue
            try:
                # ItemTypeElm=1 → Name is a LineElm (analysis element)
                ret = self.sap_model.Results.FrameJointDispl(elm_name, 1)
                n = ret[0]
                if n == 0:
                    continue

                # FrameJointDispl return layout (comtypes):
                # [0]=n, [1]=Obj, [2]=ObjSta, [3]=Elm, [4]=ElmSta,
                # [5]=ACase, [6]=StepType, [7]=StepNum,
                # [8]=U1, [9]=U2, [10]=U3, [11]=R1, [12]=R2, [13]=R3
                lc_records: Dict[str, list] = defaultdict(list)
                for i in range(n):
                    lc = str(ret[5][i])        # ACase  (was ret[3] → wrong: Elm[])
                    lc_records[lc].append((
                        float(ret[8][i]),      # U1     (was ret[6] → wrong: StepType)
                        float(ret[9][i]),      # U2     (was ret[7] → wrong: StepNum)
                        float(ret[10][i]),     # U3     (was ret[8] → wrong: U1)
                    ))

                for lc, recs in lc_records.items():
                    if len(recs) < 2:
                        continue
                    disp_i, disp_j = recs[0], recs[1]
                    if need_i:
                        results.append(DisplacementResult(
                            node_name=joint_i, load_case=lc, case_type="",
                            ux=disp_i[0], uy=disp_i[1], uz=disp_i[2]
                        ))
                    if need_j:
                        results.append(DisplacementResult(
                            node_name=joint_j, load_case=lc, case_type="",
                            ux=disp_j[0], uy=disp_j[1], uz=disp_j[2]
                        ))
            except Exception as e:
                print(f"[SAP] FrameJointDispl error for elm '{elm_name}': {e}")
        return results

    def _get_disp_via_db_tables(self, node_set, desired_lcs) -> List[DisplacementResult]:
        """Fallback: get displacements via DatabaseTables API (PointObj only)."""
        results = []
        try:
            ret = self.sap_model.DatabaseTables.GetTableForDisplayArray(
                "Joint Displacements", "", "", 0, [], 0, []
            )
            field_keys = ret[2]
            num_records = ret[3]
            table_data = ret[4]
            if num_records == 0:
                return []
            num_fields = len(field_keys)
            fm = {str(field_keys[i]).strip(): i for i in range(num_fields)}

            jc = self._find_col(fm, ["Joint", "joint"])
            cc = self._find_col(fm, ["OutputCase", "Output Case"])
            tc = self._find_col(fm, ["CaseType", "Case Type"])
            u1c = self._find_col(fm, ["U1", "u1"])
            u2c = self._find_col(fm, ["U2", "u2"])
            u3c = self._find_col(fm, ["U3", "u3"])
            if jc is None or cc is None or u1c is None:
                return []

            for r in range(num_records):
                b = r * num_fields
                try:
                    jn = str(table_data[b + jc]).strip()
                    cn = str(table_data[b + cc]).strip()
                    if jn not in node_set:
                        continue
                    if desired_lcs and cn not in desired_lcs:
                        continue
                    ct = str(table_data[b + tc]).strip() if tc is not None else ""
                    results.append(DisplacementResult(
                        node_name=jn, load_case=cn, case_type=ct,
                        ux=self._sf(table_data[b + u1c]),
                        uy=self._sf(table_data[b + u2c]) if u2c else 0,
                        uz=self._sf(table_data[b + u3c]) if u3c else 0
                    ))
                except Exception:
                    continue
        except Exception as e:
            print(f"[SAP] DatabaseTables error: {e}")
        return results

    @staticmethod
    def _find_col(fm, candidates):
        for c in candidates:
            if c in fm:
                return fm[c]
        return None

    @staticmethod
    def _sf(val) -> float:
        try:
            s = str(val).strip()
            return float(s) if s else 0.0
        except Exception:
            return 0.0

    # =========================================================================
    # Manual Group Creation
    # =========================================================================

    def build_node_degree_map(self) -> Dict[str, int]:
        """
        Return {node_name: degree} counting how many FRAME OBJECTS
        connect to each node. Uses only GetNameList + GetPoints (no coords).
        Fast — suitable to call once per Run Check.
        """
        if not self._connected:
            return {}
        try:
            ret = self.sap_model.FrameObj.GetNameList()
            frame_names = list(ret[1]) if ret[0] > 0 else []
            degree: Dict[str, int] = {}
            for fname in frame_names:
                try:
                    pts = self.sap_model.FrameObj.GetPoints(str(fname))
                    ni, nj = str(pts[0]), str(pts[1])
                    degree[ni] = degree.get(ni, 0) + 1
                    degree[nj] = degree.get(nj, 0) + 1
                except Exception:
                    pass
            return degree
        except Exception as e:
            print(f"[SAP] build_node_degree_map error: {e}")
            return {}

    def get_selected_frames(self) -> List[str]:
        """Return names of frame elements currently selected in SAP2000."""
        if not self._connected:
            return []
        try:
            ret = self.sap_model.SelectObj.GetSelected()
            n = ret[0]
            if n == 0:
                return []
            return [str(ret[2][i]) for i in range(n) if ret[1][i] == 2]
        except Exception as e:
            print(f"[SAP] get_selected_frames error: {e}")
            return []

    def get_selected_frames_and_points(self) -> tuple:
        """Return (frame_names, point_names) of currently selected objects."""
        if not self._connected:
            return [], []
        try:
            ret = self.sap_model.SelectObj.GetSelected()
            n = ret[0]
            if n == 0:
                return [], []
            frames = [str(ret[2][i]) for i in range(n) if ret[1][i] == 2]
            points = [str(ret[2][i]) for i in range(n) if ret[1][i] == 1]
            return frames, points
        except Exception as e:
            print(f"[SAP] get_selected_frames_and_points error: {e}")
            return [], []

    def get_frame_to_group_map(self) -> Dict[str, str]:
        """Return {frame_name: G-... group_name} for all G-... groups."""
        mapping: Dict[str, str] = {}
        for gname in self.get_group_names():
            if not gname.startswith('G-'):
                continue
            for fname in self.get_group_frames(gname):
                mapping[fname] = gname
        return mapping

    @staticmethod
    def _is_chain(frame_names: List[str],
                  frame_nodes: Dict[str, tuple]) -> bool:
        """
        Return True if frame_names form a single connected chain (no branches,
        no gaps). A single frame is always valid.
        """
        if len(frame_names) <= 1:
            return True
        from collections import defaultdict
        node_to_frames: Dict[str, List[str]] = defaultdict(list)
        for name in frame_names:
            ni, nj = frame_nodes[name]
            node_to_frames[ni].append(name)
            node_to_frames[nj].append(name)

        # Endpoint nodes connect to exactly 1 selected frame
        endpoints = [n for n, fl in node_to_frames.items() if len(fl) == 1]
        if len(endpoints) != 2:
            return False  # branch or disconnected

        # Walk from one endpoint and verify all frames are visited
        visited: set = set()
        current = endpoints[0]
        while True:
            candidates = [f for f in node_to_frames[current] if f not in visited]
            if not candidates:
                break
            if len(candidates) > 1:
                return False  # branch
            frame = candidates[0]
            visited.add(frame)
            ni, nj = frame_nodes[frame]
            current = nj if ni == current else ni

        return len(visited) == len(frame_names)

    def manual_create_group(self, selected_frames: List[str]) -> tuple:
        """
        Merge selected frames into one G-... group, updating existing groups.

        Rules
        -----
        - selected_frames must form a connected chain (validated before call)
        - New group name = G-{numerically smallest frame name}
        - Each old G-... group that contained any selected frame is updated:
            remaining = (old group frames) - (selected frames)
            → delete old group
            → if remaining non-empty: recreate old group with remaining frames
        - Unrelated groups and non-G-... groups are untouched.

        Returns
        -------
        (success: bool, message: str)
        """
        if not self._connected:
            return False, "Not connected to SAP2000."
        if not selected_frames:
            return False, "No frames provided."

        errors: List[str] = []

        # Determine new group name
        def _num_key(n: str):
            try:
                return (0, int(n))
            except ValueError:
                return (1, n)

        new_gname = 'G-' + min(selected_frames, key=_num_key)
        selected_set = set(selected_frames)

        # Find old G-... groups that contain any selected frame
        frame_to_group = self.get_frame_to_group_map()
        affected_groups: Dict[str, set] = {}
        for fname in selected_frames:
            gname = frame_to_group.get(fname)
            if gname and gname not in affected_groups:
                affected_groups[gname] = set(self.get_group_frames(gname))

        # Update each affected old group
        for old_gname, old_frames in affected_groups.items():
            remaining = old_frames - selected_set
            try:
                self.sap_model.GroupDef.Delete(old_gname)
            except Exception as e:
                errors.append(f"Delete '{old_gname}': {e}")
            if remaining:
                try:
                    self.sap_model.GroupDef.SetGroup(old_gname)
                    for f in remaining:
                        self.sap_model.FrameObj.SetGroupAssign(str(f), old_gname)
                except Exception as e:
                    errors.append(f"Recreate '{old_gname}': {e}")

        # Create new group with all selected frames
        try:
            self.sap_model.GroupDef.SetGroup(new_gname)
            for fname in selected_frames:
                self.sap_model.FrameObj.SetGroupAssign(str(fname), new_gname)
        except Exception as e:
            errors.append(f"Create '{new_gname}': {e}")
            return False, f"Failed: {'; '.join(errors)}"

        msg = (f"'{new_gname}' created with {len(selected_frames)} frame(s)."
               + (f"\nWarnings: {'; '.join(errors)}" if errors else ""))
        return True, msg

    @staticmethod
    def _get_chain_endpoints(frame_names: List[str],
                             frame_nodes: Dict[str, tuple]) -> tuple:
        """
        Return (start_node, end_node) of a validated chain.
        Assumes _is_chain() already passed.
        For a single frame: (node_i, node_j).
        """
        if len(frame_names) == 1:
            ni, nj = frame_nodes[frame_names[0]]
            return ni, nj
        from collections import defaultdict
        node_to_frames: Dict[str, List[str]] = defaultdict(list)
        for name in frame_names:
            ni, nj = frame_nodes[name]
            node_to_frames[ni].append(name)
            node_to_frames[nj].append(name)
        endpoints = [n for n, fl in node_to_frames.items() if len(fl) == 1]
        return endpoints[0], endpoints[1]

    def cantilever_create_group(self, selected_frames: List[str],
                                free_end_node: str) -> tuple:
        """
        Create a GC-..._S or GC-..._E group for a cantilever beam.

        The suffix _S means free end = start node of chain,
        _E means free end = end node of chain.

        Rules are identical to manual_create_group except:
        - Group prefix is 'GC-' instead of 'G-'
        - Suffix '_S' or '_E' is appended based on free_end_node position
        - Existing G-... and GC-... groups containing selected frames are updated
        """
        if not self._connected:
            return False, "Not connected to SAP2000."
        if not selected_frames:
            return False, "No frames provided."

        errors: List[str] = []

        def _num_key(n: str):
            try:
                return (0, int(n))
            except ValueError:
                return (1, n)

        # Get chain endpoints to determine _S or _E
        frame_nodes: Dict[str, tuple] = {}
        for fname in selected_frames:
            try:
                pts = self.sap_model.FrameObj.GetPoints(str(fname))
                frame_nodes[fname] = (str(pts[0]), str(pts[1]))
            except Exception as e:
                return False, f"Cannot read nodes of frame '{fname}': {e}"

        start_node, end_node = self._get_chain_endpoints(selected_frames, frame_nodes)
        if free_end_node == start_node:
            suffix = '_S'
        elif free_end_node == end_node:
            suffix = '_E'
        else:
            return False, (f"Node '{free_end_node}' is not an endpoint of the chain.\n"
                           f"Chain endpoints: '{start_node}' and '{end_node}'.")

        base_name = min(selected_frames, key=_num_key)
        new_gname = 'GC-' + base_name + suffix
        selected_set = set(selected_frames)

        # Find old G-... / GC-... groups that contain any selected frame
        frame_to_group: Dict[str, str] = {}
        for gname in self.get_group_names():
            if not (gname.startswith('G-') or gname.startswith('GC-')):
                continue
            for fname in self.get_group_frames(gname):
                if fname in selected_set:
                    frame_to_group[fname] = gname

        affected_groups: Dict[str, set] = {}
        for fname, gname in frame_to_group.items():
            if gname not in affected_groups:
                affected_groups[gname] = set(self.get_group_frames(gname))

        for old_gname, old_frames in affected_groups.items():
            remaining = old_frames - selected_set
            try:
                self.sap_model.GroupDef.Delete(old_gname)
            except Exception as e:
                errors.append(f"Delete '{old_gname}': {e}")
            if remaining:
                try:
                    self.sap_model.GroupDef.SetGroup(old_gname)
                    for f in remaining:
                        self.sap_model.FrameObj.SetGroupAssign(str(f), old_gname)
                except Exception as e:
                    errors.append(f"Recreate '{old_gname}': {e}")

        try:
            self.sap_model.GroupDef.SetGroup(new_gname)
            for fname in selected_frames:
                self.sap_model.FrameObj.SetGroupAssign(str(fname), new_gname)
        except Exception as e:
            errors.append(f"Create '{new_gname}': {e}")
            return False, f"Failed: {'; '.join(errors)}"

        msg = (f"'{new_gname}' created — free end: '{free_end_node}' ({suffix[1:]})."
               + (f"\nWarnings: {'; '.join(errors)}" if errors else ""))
        return True, msg

    # =========================================================================
    # PMember Group Creation  (new — does not affect existing methods above)
    # =========================================================================

    def get_all_frames_raw(self) -> List[Dict]:
        """
        Fetch every frame object with its end-node coordinates.

        Returns list of dicts:
            {'name', 'node_i', 'node_j', 'xi', 'yi', 'zi', 'xj', 'yj', 'zj'}
        Node coordinates are in the current model units (kN/mm after connect).
        """
        if not self._connected:
            return []
        try:
            self.sap_model.SetPresentUnits(self.UNIT_KN_MM)
            ret = self.sap_model.FrameObj.GetNameList()
            frame_names = list(ret[1]) if ret[0] > 0 else []

            node_cache: Dict[str, tuple] = {}

            def coord(node: str) -> tuple:
                if node not in node_cache:
                    c = self.sap_model.PointObj.GetCoordCartesian(node)
                    node_cache[node] = (float(c[0]), float(c[1]), float(c[2]))
                return node_cache[node]

            frames = []
            for fname in frame_names:
                try:
                    fname = str(fname)
                    pts = self.sap_model.FrameObj.GetPoints(fname)
                    ni, nj = str(pts[0]), str(pts[1])
                    xi, yi, zi = coord(ni)
                    xj, yj, zj = coord(nj)
                    sec = self.sap_model.FrameObj.GetSection(fname)
                    section = str(sec[0]) if sec else ''
                    frames.append({
                        'name': fname,
                        'node_i': ni, 'node_j': nj,
                        'xi': xi, 'yi': yi, 'zi': zi,
                        'xj': xj, 'yj': yj, 'zj': zj,
                        'section': section,
                    })
                except Exception as e:
                    print(f"[SAP] Warning: skipping frame '{fname}': {e}")
            return frames
        except Exception as e:
            print(f"[SAP] get_all_frames_raw error: {e}")
            return []

    def get_all_releases(self) -> Dict[str, tuple]:
        """
        Get release flags for every frame element.

        Returns {frame_name: (start_release: bool, end_release: bool)}
        True means the I-end or J-end has at least one DOF released.
        """
        if not self._connected:
            return {}
        try:
            ret = self.sap_model.FrameObj.GetNameList()
            frame_names = list(ret[1]) if ret[0] > 0 else []
            releases: Dict[str, tuple] = {}
            for fname in frame_names:
                fname = str(fname)
                try:
                    r = self.sap_model.FrameObj.GetReleases(fname)
                    ii, jj = r[0], r[1]   # boolean arrays: [P, V2, V3, T, M2, M3]
                    releases[fname] = (any(ii), any(jj))
                except Exception:
                    releases[fname] = (False, False)
            return releases
        except Exception as e:
            print(f"[SAP] get_all_releases error: {e}")
            return {}

    def get_post_frame_names(self) -> set:
        """Return the set of frame names inside the SAP2000 group 'G_Post'."""
        try:
            return set(self.get_group_frames("G_Post"))
        except Exception:
            return set()

    def create_pmember_groups(self, pmembers: List[Dict], progress_cb=None) -> tuple:
        """
        Delete all existing 'G-...' groups, then create new ones.

        Parameters
        ----------
        pmembers  list of {'group_name': str, 'frames': [str, ...]}

        Returns
        -------
        (n_created: int, n_deleted: int, errors: list[str])
        """
        if not self._connected:
            return 0, 0, ["Not connected to SAP2000"]

        errors: List[str] = []
        n_deleted = 0
        n_created = 0

        # Step 1 — delete existing G-... groups
        try:
            for gname in self.get_group_names():
                if gname.startswith('G-'):
                    try:
                        ret = self.sap_model.GroupDef.Delete(gname)
                        if ret == 0:
                            n_deleted += 1
                        else:
                            errors.append(f"Delete '{gname}' ret={ret}")
                    except Exception as e:
                        errors.append(f"Delete '{gname}': {e}")
        except Exception as e:
            errors.append(f"Cannot list groups: {e}")

        # Step 2 — create new groups and assign frames
        for i, pm in enumerate(pmembers):
            gname      = pm['group_name']
            frame_list = pm['frames']
            try:
                ret = self.sap_model.GroupDef.SetGroup(gname)
                if ret != 0:
                    errors.append(f"SetGroup '{gname}' ret={ret}")
                    continue
                for fname in frame_list:
                    try:
                        ret2 = self.sap_model.FrameObj.SetGroupAssign(str(fname), gname)
                        if ret2 != 0:
                            errors.append(f"Assign '{fname}'→'{gname}' ret={ret2}")
                    except Exception as e:
                        errors.append(f"Assign '{fname}'→'{gname}': {e}")
                n_created += 1
            except Exception as e:
                errors.append(f"Create '{gname}': {e}")

            if progress_cb and (i % 10 == 0 or i == len(pmembers) - 1):
                try:
                    progress_cb(i + 1)
                except Exception:
                    pass

        return n_created, n_deleted, errors
