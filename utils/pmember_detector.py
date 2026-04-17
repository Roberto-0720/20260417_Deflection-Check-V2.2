"""
PMember Detector — groups SAP2000 frame elements into Physical Members.

Algorithm adapted from Tool 10 (10 SAP Parameter / processing.py).
Works entirely in-memory: no CSV files, no disk I/O.

Each Physical Member (PMember) represents one continuous structural
member in the real structure — e.g., a multi-span continuous beam
modelled as several shorter frame elements.
"""
import math
from typing import Dict, List, Set, Tuple


class FrameData:
    """Lightweight frame descriptor built from SAP2000 raw geometry."""

    __slots__ = (
        'name', 'node_i', 'node_j',
        'dx', 'dy', 'dz', 'alpha',
        'start_release', 'end_release',
        'element_type', 'pmember', 'section',
    )

    def __init__(
        self,
        name: str,
        node_i: str, node_j: str,
        xi: float, yi: float, zi: float,
        xj: float, yj: float, zj: float,
        section: str = '',
    ):
        self.name = name
        self.node_i = node_i
        self.node_j = node_j

        self.dx = xj - xi
        self.dy = yj - yi
        self.dz = zj - zi

        horiz = math.sqrt(self.dx ** 2 + self.dy ** 2)
        if horiz < 1e-9 and abs(self.dz) < 1e-9:
            self.alpha = 0.0
        elif horiz < 1e-9:
            self.alpha = 90.0
        else:
            self.alpha = math.degrees(math.atan2(abs(self.dz), horiz))

        self.start_release = False   # any DOF released at I-end
        self.end_release   = False   # any DOF released at J-end
        self.element_type  = ''
        self.pmember       = ''
        self.section       = section


class PmemberDetector:
    """
    Groups frame elements into Physical Members.

    Usage
    -----
    detector = PmemberDetector()
    pmembers = detector.detect(frames, releases, post_frame_names)

    Returns
    -------
    List of dicts: [{'group_name': 'G-205', 'frames': ['205', '206']}, ...]
    Group name = 'G-' + smallest (numeric) frame name in the PMember.
    """

    BEAM_ALPHA_MAX = 11.5   # degrees — near-horizontal threshold
    COL_ALPHA_MIN  = 78.5   # degrees — near-vertical  threshold

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def detect(
        self,
        frames: List[FrameData],
        releases: Dict[str, Tuple[bool, bool]],
        post_frame_names: Set[str],
    ) -> List[Dict]:
        """
        Parameters
        ----------
        frames            list of FrameData objects (geometry already set)
        releases          {frame_name: (start_release, end_release)}
                          True = I-end / J-end has at least one DOF released
        post_frame_names  set of frame names in SAP2000 group "G_Post"
        """
        # Apply releases onto FrameData objects
        for f in frames:
            if f.name in releases:
                f.start_release, f.end_release = releases[f.name]

        # Node degree: how many frames touch each node
        node_degree: Dict[str, int] = {}
        for f in frames:
            node_degree[f.node_i] = node_degree.get(f.node_i, 0) + 1
            node_degree[f.node_j] = node_degree.get(f.node_j, 0) + 1

        # Classify each frame
        for f in frames:
            f.element_type = self._classify(f, node_degree, post_frame_names)

        # Node → list[frame_index] (undirected, both ends)
        node_to_idx: Dict[str, List[int]] = {}
        for idx, f in enumerate(frames):
            node_to_idx.setdefault(f.node_i, []).append(idx)
            node_to_idx.setdefault(f.node_j, []).append(idx)

        # Nodes that are adjacent to at least one column element
        col_nodes: Set[str] = {
            node
            for f in frames if f.element_type == 'COLUMN'
            for node in (f.node_i, f.node_j)
        }

        # --- Assign PMember IDs ---
        counter = [0]

        def new_pm() -> str:
            counter[0] += 1
            return f'PM{counter[0]}'

        # Pass 1: chain columns (both directions from each seed, same section only)
        for idx, f in enumerate(frames):
            if f.pmember or f.element_type != 'COLUMN':
                continue
            pm = new_pm()
            f.pmember = pm
            self._chain_col(idx, f.node_i, frames, node_to_idx, pm, f.section)
            self._chain_col(idx, f.node_j, frames, node_to_idx, pm, f.section)

        # Pass 2: standalone types — one element = one PMember
        for f in frames:
            if not f.pmember and f.element_type in ('VBR', 'HBR', 'CONSOLE'):
                f.pmember = new_pm()

        # Pass 3a: POST — chaining logic mirrors Beam (release-based),
        # but without column-node boundaries (Posts live between beams, not columns).
        for idx, f in enumerate(frames):
            if f.pmember or f.element_type != 'POST':
                continue
            sr, er = f.start_release, f.end_release

            if sr and er:
                # Single-element Post (hinge at both ends) → isolated PMember
                f.pmember = new_pm()

            elif sr:
                # I-end hinged → seed here, chain forward from J until J-end hinge
                pm = new_pm()
                f.pmember = pm
                self._chain_beam_fwd(
                    idx, f.node_j, 'POST',
                    frames, node_to_idx, pm,
                    stop_at_er=True, origin_section=f.section,
                )
            # No start_release → part of a chain started elsewhere; pass 4 fallback

        # Pass 3b: Beam X / Beam Y with release-based chaining rules
        for idx, f in enumerate(frames):
            if f.pmember or f.element_type not in ('BEAM X', 'BEAM Y'):
                continue
            sr, er = f.start_release, f.end_release

            if sr and er:
                # Both ends pinned → isolated span
                f.pmember = new_pm()

            elif sr:
                # I-end pinned → chain forward from J-end;
                # stop when we reach a frame whose J-end is also pinned
                pm = new_pm()
                f.pmember = pm
                self._chain_beam_fwd(
                    idx, f.node_j, f.element_type,
                    frames, node_to_idx, pm,
                    stop_at_er=True, origin_section=f.section,
                )

            elif f.node_i in col_nodes:
                # Frame starts at a column → chain forward until next column node
                pm = new_pm()
                f.pmember = pm
                self._chain_beam_fwd(
                    idx, f.node_j, f.element_type,
                    frames, node_to_idx, pm,
                    stop_at_col=True, col_nodes=col_nodes,
                    origin_section=f.section,
                )
            # else → handled in pass 4

        # Pass 4: fallback — for beams, still stop at column nodes and section change
        for idx, f in enumerate(frames):
            if f.pmember:
                continue
            pm = new_pm()
            f.pmember = pm
            is_beam = f.element_type in ('BEAM X', 'BEAM Y')
            self._chain_beam_fwd(
                idx, f.node_j, f.element_type,
                frames, node_to_idx, pm,
                stop_at_col=is_beam, col_nodes=col_nodes if is_beam else None,
                origin_section=f.section,
            )

        # --- Collect and name results ---
        pm_to_frames: Dict[str, List[str]] = {}
        for f in frames:
            pm_to_frames.setdefault(f.pmember, []).append(f.name)

        return [
            {'group_name': 'G-' + self._min_name(fnames), 'frames': fnames}
            for fnames in pm_to_frames.values()
        ]

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _classify(
        self,
        f: FrameData,
        node_degree: Dict[str, int],
        post_frames: Set[str],
    ) -> str:
        if f.name in post_frames:
            return 'POST'
        if f.alpha >= self.COL_ALPHA_MIN:
            # Vertical element: Post if either end is hinged, Column otherwise.
            # Post = vertical member supported BY beams (hinges at both physical ends).
            # Column = vertical member that SUPPORTS beams (rigid connections).
            if f.start_release or f.end_release:
                return 'POST'
            return 'COLUMN'
        ni_deg = node_degree.get(f.node_i, 0)
        nj_deg = node_degree.get(f.node_j, 0)
        if ni_deg == 1 or nj_deg == 1:
            return 'CONSOLE'
        if f.alpha > self.BEAM_ALPHA_MAX:
            return 'VBR'
        # Near-horizontal
        dx_zero = abs(f.dx) < 1e-4
        dy_zero = abs(f.dy) < 1e-4
        if dy_zero and not dx_zero:
            return 'BEAM X'
        if dx_zero and not dy_zero:
            return 'BEAM Y'
        return 'HBR'

    def _chain_col(
        self,
        origin_idx: int,
        start_node: str,
        frames: List[FrameData],
        node_to_idx: Dict[str, List[int]],
        pm: str,
        origin_section: str = '',
    ) -> None:
        """
        Chain column frames in one direction (undirected walk).
        Stops when no unassigned column with the same section is connected.
        """
        current  = start_node
        prev_idx = origin_idx
        while True:
            found = False
            for cidx in node_to_idx.get(current, []):
                cf = frames[cidx]
                if cidx == prev_idx or cf.pmember or cf.element_type != 'COLUMN':
                    continue
                if origin_section and cf.section != origin_section:
                    continue                    # section change → new PMember
                cf.pmember = pm
                current  = cf.node_j if cf.node_i == current else cf.node_i
                prev_idx = cidx
                found    = True
                break
            if not found:
                break

    def _chain_beam_fwd(
        self,
        origin_idx: int,
        start_node: str,
        etype: str,
        frames: List[FrameData],
        node_to_idx: Dict[str, List[int]],
        pm: str,
        stop_at_er: bool = False,
        stop_at_col: bool = False,
        col_nodes: Set[str] = None,
        origin_section: str = '',
    ) -> None:
        """
        Chain beams/posts in the forward direction (I→J orientation).

        Col-node check runs at the TOP of each iteration — before looking for
        the next frame. This ensures a beam ending AT a column is included in
        the current PMember, but the next span starting FROM that same column
        node is NOT chained in (it belongs to the next PMember).
        """
        current = start_node
        while True:
            # Stop BEFORE crossing into a column node
            if stop_at_col and col_nodes and current in col_nodes:
                return
            found = False
            for cidx in node_to_idx.get(current, []):
                cf = frames[cidx]
                if cf.pmember or cf.element_type != etype:
                    continue
                if cf.node_i != current:           # only follow I→J direction
                    continue
                if origin_section and cf.section != origin_section:
                    continue                        # section change → new PMember
                cf.pmember = pm
                current = cf.node_j
                found   = True
                if stop_at_er and cf.end_release:
                    return
                break
            if not found:
                break

    @staticmethod
    def _min_name(names: List[str]) -> str:
        """Return numerically-smallest name (or lexically smallest)."""
        def key(n: str):
            try:
                return (0, int(n))
            except ValueError:
                return (1, n)
        return min(names, key=key)
