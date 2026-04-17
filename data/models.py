"""Data models for Deflection Check tool."""
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
import math


@dataclass
class NodeInfo:
    name: str
    x: float  # mm
    y: float  # mm
    z: float  # mm


@dataclass
class FrameElement:
    name: str
    joint_i: str
    joint_j: str
    section: str
    length: float  # mm


@dataclass
class BeamGroup:
    """A physical beam = group of one or more frame elements."""
    group_name: str
    element_names: List[str] = field(default_factory=list)
    node_names_ordered: List[str] = field(default_factory=list)
    section: str = ""
    total_length: float = 0.0
    start_node: str = ""
    end_node: str = ""
    is_cantilever: bool = False   # True if one end is a free end (degree=1)
    free_end_node: str = ""       # the free-end node name (degree=1 in model)


@dataclass
class DisplacementResult:
    node_name: str
    load_case: str
    case_type: str
    ux: float  # mm
    uy: float  # mm
    uz: float  # mm


@dataclass
class DeflectionResult:
    """Deflection check result for one beam under one load case."""
    group_name: str
    section: str
    load_case: str
    span_mm: float
    max_deflection_mm: float  # absolute max perpendicular deflection
    critical_node: str        # node where max deflection occurs
    allowable_ratio: float    # e.g. 360 for L/360
    element_list: List[str] = field(default_factory=list)

    @property
    def actual_ratio(self) -> float:
        """L/delta — higher is better."""
        if self.max_deflection_mm == 0:
            return float('inf')
        return self.span_mm / self.max_deflection_mm

    @property
    def is_ok(self) -> bool:
        return self.actual_ratio >= self.allowable_ratio

    @property
    def ratio_str(self) -> str:
        if self.max_deflection_mm == 0:
            return "L/∞"
        return f"L/{min(int(self.actual_ratio), 99999)}"


@dataclass
class BeamCheckSummary:
    """Summary for one physical beam across all load cases."""
    group_name: str
    section: str
    span_mm: float
    controlling_lc: str = ""
    max_deflection_mm: float = 0.0
    actual_ratio: float = float('inf')
    critical_node: str = ""
    allowable_ratio: float = 360.0
    element_list: List[str] = field(default_factory=list)
    all_results: List[DeflectionResult] = field(default_factory=list)
    # Absolute deflection check
    max_abs_deflection_mm: float = 0.0   # max perp displacement from original pos
    abs_limit_mm: float = 25.0           # allowable absolute (0 = disabled)
    abs_controlling_lc: str = ""

    @property
    def rel_is_ok(self) -> bool:
        return self.actual_ratio >= self.allowable_ratio

    @property
    def abs_is_ok(self) -> bool:
        return (self.abs_limit_mm <= 0 or
                self.max_abs_deflection_mm <= self.abs_limit_mm)

    @property
    def is_ok(self) -> bool:
        return self.rel_is_ok and self.abs_is_ok

    @property
    def ratio_str(self) -> str:
        if self.max_deflection_mm == 0:
            return "L/∞"
        return f"L/{min(int(self.actual_ratio), 99999)}"
