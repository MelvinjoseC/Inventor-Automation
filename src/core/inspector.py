import logging
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("InventorAutomation.Inspector")


class ModelInspector:
    """Extracts physical properties and executes automated Design Rule Checks (DRC) on Inventor parts."""

    def __init__(self, inventor_app: Any, inspection_rules: Optional[Dict[str, Any]] = None):
        self.app = inventor_app
        self.rules = inspection_rules or {
            "min_wall_thickness_mm": 1.0,
            "min_hole_edge_dist_mm": 2.0,
            "max_aspect_ratio": 50.0,
        }

    def get_physical_properties(self, part_doc: Any) -> Dict[str, Any]:
        """Extracts mass, volume, surface area, bounding box, and center of gravity."""
        props: Dict[str, Any] = {
            "mass_kg": 0.0,
            "volume_cm3": 0.0,
            "surface_area_cm2": 0.0,
            "center_of_gravity": {"x_cm": 0.0, "y_cm": 0.0, "z_cm": 0.0},
            "bounding_box_cm": {"dx": 0.0, "dy": 0.0, "dz": 0.0},
            "density_g_cm3": 0.0,
        }

        try:
            cd = part_doc.ComponentDefinition
            mass_props = getattr(cd, "MassProperties", None)
            if mass_props:
                props["mass_kg"] = round(float(getattr(mass_props, "Mass", 0.0)), 4)
                props["volume_cm3"] = round(float(getattr(mass_props, "Volume", 0.0)), 4)
                props["surface_area_cm2"] = round(float(getattr(mass_props, "Area", 0.0)), 4)
                
                # Center of gravity
                cg = getattr(mass_props, "CenterOfMass", None)
                if cg:
                    props["center_of_gravity"] = {
                        "x_cm": round(float(getattr(cg, "X", 0.0)), 3),
                        "y_cm": round(float(getattr(cg, "Y", 0.0)), 3),
                        "z_cm": round(float(getattr(cg, "Z", 0.0)), 3),
                    }

                # Material density
                mat = getattr(cd, "Material", None)
                if mat and hasattr(mat, "Density"):
                    props["density_g_cm3"] = round(float(mat.Density), 4)

            # Bounding Range Box
            range_box = getattr(cd, "RangeBox", None)
            if range_box:
                min_pt = range_box.MinPoint
                max_pt = range_box.MaxPoint
                dx = round(abs(float(max_pt.X) - float(min_pt.X)), 3)
                dy = round(abs(float(max_pt.Y) - float(min_pt.Y)), 3)
                dz = round(abs(float(max_pt.Z) - float(min_pt.Z)), 3)
                props["bounding_box_cm"] = {"dx": dx, "dy": dy, "dz": dz}

            logger.info(
                f"Physical properties extracted: Mass={props['mass_kg']}kg, "
                f"Volume={props['volume_cm3']}cm³, BoundingBox={props['bounding_box_cm']}"
            )
        except Exception as e:
            logger.warning(f"Could not extract complete physical properties: {e}")

        return props

    def run_drc(self, part_doc: Any, bom_item: Optional[Any] = None) -> Dict[str, Any]:
        """Runs automated Design Rule Checks (DRC) on the part geometry and dimensions."""
        checks: List[Dict[str, Any]] = []
        overall_pass = True

        props = self.get_physical_properties(part_doc)

        # 1. Non-zero Volume & Mass Check (detect empty models or hollow voids)
        vol_check = props["volume_cm3"] > 0.0001
        checks.append({
            "rule": "SolidGeometryPresence",
            "passed": vol_check,
            "details": f"Volume = {props['volume_cm3']} cm³",
        })
        if not vol_check:
            overall_pass = False

        # 2. Aspect Ratio Check (bounding box length vs thickness)
        bb = props["bounding_box_cm"]
        dims = sorted([bb["dx"], bb["dy"], bb["dz"]])
        max_dim = dims[-1]
        min_dim = dims[0] if dims[0] > 0 else 0.001
        aspect_ratio = max_dim / min_dim
        max_allowed_aspect = float(self.rules.get("max_aspect_ratio", 50.0))

        aspect_check = aspect_ratio <= max_allowed_aspect
        checks.append({
            "rule": "AspectRatioLimit",
            "passed": aspect_check,
            "details": f"Aspect ratio {aspect_ratio:.1f}:1 (limit {max_allowed_aspect}:1)",
        })
        if not aspect_check:
            overall_pass = False

        # 3. Minimum Wall Thickness Check (from BOM if available)
        if bom_item:
            min_wall_mm = float(self.rules.get("min_wall_thickness_mm", 1.0))
            part_thickness_mm = getattr(bom_item, "thickness_cm", 0.0) * 10.0

            if part_thickness_mm > 0:
                wall_check = part_thickness_mm >= min_wall_mm
                checks.append({
                    "rule": "MinimumWallThickness",
                    "passed": wall_check,
                    "details": f"Thickness = {part_thickness_mm:.2f} mm (minimum {min_wall_mm} mm)",
                })
                if not wall_check:
                    overall_pass = False

            # 4. Hole-to-Edge Clearance Check
            min_hole_edge_mm = float(self.rules.get("min_hole_edge_dist_mm", 2.0))
            margin_x = getattr(bom_item, "hole_margin_x_cm", None)
            margin_y = getattr(bom_item, "hole_margin_y_cm", None)
            hole_dia = getattr(bom_item, "hole_diameter_cm", None)
            if margin_x is not None or margin_y is not None:
                eff_margin_x = (margin_x * 10.0) if margin_x is not None else 100.0
                eff_margin_y = (margin_y * 10.0) if margin_y is not None else 100.0
                edge_clearance = min(eff_margin_x, eff_margin_y)
                if hole_dia:
                    edge_clearance -= (hole_dia * 5.0)  # subtract radius
                edge_clearance = max(0.0, edge_clearance)
                hole_edge_check = edge_clearance >= min_hole_edge_mm
                checks.append({
                    "rule": "MinHoleEdgeDistance",
                    "passed": hole_edge_check,
                    "details": f"Clearance = {edge_clearance:.2f} mm (min allowed {min_hole_edge_mm} mm)",
                })
                if not hole_edge_check:
                    overall_pass = False

        # 5. Machine Build Envelope Check
        max_envelope_cm = self.rules.get("max_envelope_cm", [100.0, 100.0, 100.0])
        env_x, env_y, env_z = max_envelope_cm[0], max_envelope_cm[1], max_envelope_cm[2]
        envelope_check = (bb["dx"] <= env_x and bb["dy"] <= env_y and bb["dz"] <= env_z)
        checks.append({
            "rule": "BuildEnvelopeLimit",
            "passed": envelope_check,
            "details": f"Bounds {bb['dx']}x{bb['dy']}x{bb['dz']} cm (max {env_x}x{env_y}x{env_z} cm)",
        })
        if not envelope_check:
            overall_pass = False

        # 6. Mass Budget Check (if configured)
        max_mass_kg = self.rules.get("max_mass_kg")
        if max_mass_kg is not None:
            mass_budget_check = props["mass_kg"] <= float(max_mass_kg)
            checks.append({
                "rule": "MassBudgetLimit",
                "passed": mass_budget_check,
                "details": f"Mass = {props['mass_kg']:.3f} kg (budget {max_mass_kg} kg)",
            })
            if not mass_budget_check:
                overall_pass = False

        # 7. Surface Area-to-Volume Ratio
        if props["volume_cm3"] > 0:
            sa_v_ratio = round(props["surface_area_cm2"] / props["volume_cm3"], 2)
            checks.append({
                "rule": "SurfaceToVolumeRatio",
                "passed": True,
                "details": f"S/V ratio = {sa_v_ratio} cm⁻¹",
            })

        drc_summary = {
            "overall_passed": overall_pass,
            "checks": checks,
            "properties": props,
            "rule_count": len(checks),
            "passed_count": sum(1 for c in checks if c["passed"]),
        }

        status_str = "PASS" if overall_pass else "FAIL"
        logger.info(f"DRC Inspection complete: {status_str} ({drc_summary['passed_count']}/{drc_summary['rule_count']} passed)")
        return drc_summary
