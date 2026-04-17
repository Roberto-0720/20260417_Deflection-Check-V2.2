"""Deflection calculation engine.

For each physical beam (group of elements):
1. Collect all nodes along the beam (including virtual mesh nodes), ordered I→J
2. For each load case, compute deflection:
   - Draw chord line between displaced start and end nodes
   - At each intermediate node, compute perpendicular distance from chord
   - Max perpendicular distance = deflection for that LC
3. Report controlling LC with max deflection
"""
import math
from typing import List, Dict, Tuple, Optional
from data.models import (
    NodeInfo, FrameElement, BeamGroup, DisplacementResult,
    DeflectionResult, BeamCheckSummary
)


class DeflectionCalculator:

    def __init__(self, allowable_ratio: float = 360.0, abs_limit_mm: float = 25.0):
        self.allowable_ratio = allowable_ratio
        self.abs_limit_mm    = abs_limit_mm

    def build_beam_group(
        self,
        group_name: str,
        frame_elements: List[FrameElement],
        all_nodes_ordered: List[str],
        nodes_coords: Dict[str, NodeInfo]
    ) -> BeamGroup:
        """Build a BeamGroup from frame elements belonging to one group."""
        bg = BeamGroup(group_name=group_name)
        bg.element_names = [fe.name for fe in frame_elements]

        if frame_elements:
            bg.section = frame_elements[0].section

        if all_nodes_ordered and len(all_nodes_ordered) >= 2:
            bg.node_names_ordered = all_nodes_ordered
            bg.start_node = all_nodes_ordered[0]
            bg.end_node = all_nodes_ordered[-1]

            # Calculate total span
            n_start = nodes_coords.get(bg.start_node)
            n_end = nodes_coords.get(bg.end_node)
            if n_start and n_end:
                dx = n_end.x - n_start.x
                dy = n_end.y - n_start.y
                dz = n_end.z - n_start.z
                bg.total_length = math.sqrt(dx**2 + dy**2 + dz**2)

        return bg

    def compute_deflection_for_lc(
        self,
        beam: BeamGroup,
        nodes_coords: Dict[str, NodeInfo],
        disp_map: Dict[str, Tuple[float, float, float]],
    ) -> Tuple[float, str]:
        """
        Compute max perpendicular deflection for one beam under one load case.

        The deflection is computed as the maximum perpendicular distance
        from the displaced intermediate nodes to the displaced chord line
        (line connecting displaced start and end nodes).

        Args:
            beam: The beam group
            nodes_coords: Original node coordinates
            disp_map: node_name -> (ux, uy, uz) for this specific load case

        Returns:
            (max_deflection_mm, critical_node_name)
        """
        if len(beam.node_names_ordered) < 3:
            # Only 2 nodes (start and end) → no intermediate nodes to check
            # Still compute: the deflection at the two ends relative to chord = 0
            return 0.0, ""

        start = beam.start_node
        end = beam.end_node

        # Original positions
        s0 = nodes_coords.get(start)
        e0 = nodes_coords.get(end)
        if not s0 or not e0:
            return 0.0, ""

        # ── Original (undeflected) chord ──────────────────────────────────────
        # Using original geometry as the reference eliminates geometric offsets
        # (camber, pitch, inverted-V rise) so only the true load deflection is
        # measured. This is the "chord-rotation" method standard in structural
        # engineering and works for horizontal, inclined, and pitched beams.
        # ──────────────────────────────────────────────────────────────────────
        ocx = e0.x - s0.x
        ocy = e0.y - s0.y
        ocz = e0.z - s0.z
        orig_chord_len = math.sqrt(ocx**2 + ocy**2 + ocz**2)
        if orig_chord_len < 1e-6:
            return 0.0, ""

        # Unit vector along original chord
        ux = ocx / orig_chord_len
        uy = ocy / orig_chord_len
        uz = ocz / orig_chord_len

        ds = disp_map.get(start, (0.0, 0.0, 0.0))
        de = disp_map.get(end,   (0.0, 0.0, 0.0))

        max_defl = 0.0
        crit_node = ""

        for node_name in beam.node_names_ordered[1:-1]:  # skip start and end
            n0 = nodes_coords.get(node_name)
            if not n0:
                continue
            dn = disp_map.get(node_name, (0.0, 0.0, 0.0))

            # Position parameter t along the original chord (0=start, 1=end)
            vox = n0.x - s0.x
            voy = n0.y - s0.y
            voz = n0.z - s0.z
            t = (vox * ux + voy * uy + voz * uz) / orig_chord_len
            t = max(0.0, min(1.0, t))

            # Interpolated (rigid-body) displacement at this position
            ix = (1.0 - t) * ds[0] + t * de[0]
            iy = (1.0 - t) * ds[1] + t * de[1]
            iz = (1.0 - t) * ds[2] + t * de[2]

            # Relative deformation = actual displacement minus rigid-body part
            rx = dn[0] - ix
            ry = dn[1] - iy
            rz = dn[2] - iz

            # Use only vertical (U3) component of relative deformation.
            # This matches the reference algorithm: δ = |U3_mid - interp(U3_start, U3_end)|
            # 3D perpendicular would incorrectly inflate results for beams with
            # lateral drift (wind/seismic), which is NOT part of vertical deflection.
            perp_dist = abs(rz)

            if perp_dist > max_defl:
                max_defl = perp_dist
                crit_node = node_name

        return max_defl, crit_node

    def compute_cantilever_deflection_for_lc(
        self,
        beam: BeamGroup,
        nodes_coords: Dict[str, NodeInfo],
        disp_map: Dict[str, Tuple[float, float, float]],
    ) -> Tuple[float, str]:
        """
        Compute deflection for a cantilever beam under one load case.

        The deflection = perpendicular displacement of the free end relative
        to the fixed end, measured perpendicular to the ORIGINAL beam axis.

        This correctly handles:
        - Cantilevers with only 2 nodes (chord method gives 0 — wrong)
        - The free end displacing both vertically and horizontally
        """
        free_node  = beam.free_end_node
        fixed_node = beam.end_node if free_node == beam.start_node else beam.start_node

        f0 = nodes_coords.get(fixed_node)
        e0 = nodes_coords.get(free_node)
        if not f0 or not e0:
            return 0.0, free_node

        # Original beam axis unit vector (fixed → free)
        odx = e0.x - f0.x
        ody = e0.y - f0.y
        odz = e0.z - f0.z
        orig_len = math.sqrt(odx**2 + ody**2 + odz**2)
        if orig_len < 1e-6:
            return 0.0, free_node
        ux, uy, uz = odx / orig_len, ody / orig_len, odz / orig_len

        # Relative displacement: free_end - fixed_end
        df = disp_map.get(fixed_node, (0.0, 0.0, 0.0))
        de = disp_map.get(free_node,  (0.0, 0.0, 0.0))
        dux = de[0] - df[0]
        duy = de[1] - df[1]
        duz = de[2] - df[2]

        # Use only vertical (U3) component — consistent with relative check algorithm.
        # duz = U3_free - U3_fixed = net vertical displacement at free end
        # relative to fixed end support.
        return abs(duz), free_node

    def compute_abs_deflection_for_lc(
        self,
        beam: BeamGroup,
        nodes_coords: Dict[str, NodeInfo],
        disp_map: Dict[str, Tuple[float, float, float]],
    ) -> Tuple[float, str]:
        """
        Absolute deflection: max vertical displacement U3 (index 2) of
        ANY node along the beam — raw SAP2000 U3 value, no chord
        subtraction. Positive or negative (taken as absolute value).
        """
        max_abs   = 0.0
        crit_node = ""

        for node_name in beam.node_names_ordered:
            dn = disp_map.get(node_name)
            if not dn:
                continue
            abs_d = abs(dn[2])   # U3 = vertical (index 2)
            if abs_d > max_abs:
                max_abs   = abs_d
                crit_node = node_name

        return max_abs, crit_node

    def check_beam(
        self,
        beam: BeamGroup,
        nodes_coords: Dict[str, NodeInfo],
        displacements: List[DisplacementResult],
    ) -> BeamCheckSummary:
        """
        Check deflection for one beam across all load cases.
        Returns the summary with controlling load case.
        """
        # Cantilever allowable = 2L/ratio → equivalent to L/(ratio/2)
        effective_ratio = (self.allowable_ratio / 2.0
                           if beam.is_cantilever else self.allowable_ratio)

        summary = BeamCheckSummary(
            group_name=beam.group_name,
            section=beam.section,
            span_mm=beam.total_length,
            allowable_ratio=effective_ratio,
            element_list=beam.element_names,
        )

        # Build lookup: {load_case: {node_name: (ux, uy, uz)}}
        lc_disp: Dict[str, Dict[str, Tuple[float, float, float]]] = {}
        relevant_nodes = set(beam.node_names_ordered)
        for d in displacements:
            if d.node_name in relevant_nodes:
                if d.load_case not in lc_disp:
                    lc_disp[d.load_case] = {}
                lc_disp[d.load_case][d.node_name] = (d.ux, d.uy, d.uz)

        max_defl_overall = 0.0
        ctrl_lc   = ""
        ctrl_node = ""
        max_abs_overall = 0.0
        abs_ctrl_lc     = ""

        for lc, disp_map in lc_disp.items():
            if beam.is_cantilever:
                defl, crit = self.compute_cantilever_deflection_for_lc(
                    beam, nodes_coords, disp_map)
            else:
                defl, crit = self.compute_deflection_for_lc(
                    beam, nodes_coords, disp_map)

            result = DeflectionResult(
                group_name=beam.group_name,
                section=beam.section,
                load_case=lc,
                span_mm=beam.total_length,
                max_deflection_mm=defl,
                critical_node=crit,
                allowable_ratio=effective_ratio,
                element_list=beam.element_names
            )
            summary.all_results.append(result)

            if defl > max_defl_overall:
                max_defl_overall = defl
                ctrl_lc   = lc
                ctrl_node = crit

            # Absolute deflection uses the SAME chord-interpolation value as
            # the relative check — only the comparison limit differs (fixed mm
            # instead of L/ratio). Raw max|U3| would wrongly include support
            # settlement / rigid-body vertical movement of the whole structure.
            if defl > max_abs_overall:
                max_abs_overall = defl
                abs_ctrl_lc     = lc

        summary.max_deflection_mm     = max_defl_overall
        summary.controlling_lc        = ctrl_lc
        summary.critical_node         = ctrl_node
        summary.max_abs_deflection_mm = max_abs_overall
        summary.abs_controlling_lc    = abs_ctrl_lc
        summary.abs_limit_mm          = self.abs_limit_mm

        if max_defl_overall > 0:
            summary.actual_ratio = beam.total_length / max_defl_overall
        else:
            summary.actual_ratio = float('inf')

        return summary

    def run_full_check(
        self,
        beams: List[BeamGroup],
        nodes_coords: Dict[str, NodeInfo],
        displacements: List[DisplacementResult],
    ) -> List[BeamCheckSummary]:
        """Run deflection check for all beams."""
        results = []
        for beam in beams:
            if len(beam.node_names_ordered) < 2:
                continue
            summary = self.check_beam(beam, nodes_coords, displacements)
            results.append(summary)
        return results
