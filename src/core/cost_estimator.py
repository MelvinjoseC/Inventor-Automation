import logging
from typing import Any, Dict, Optional

logger = logging.getLogger("InventorAutomation.CostEstimator")


class CostEstimator:
    """Estimates raw material stock dimensions, chip removal volume, machining cycle time,
    and manufacturing cost for automated fabrication planning.
    """

    # Industry standard defaults: Density (g/cm3), Cost per kg (USD), Machine MRR (cm3/min)
    MATERIAL_PROFILES: Dict[str, Dict[str, float]] = {
        "steel": {"density": 7.85, "cost_per_kg": 3.50, "mrr_cm3_min": 12.0},
        "steel, mild": {"density": 7.85, "cost_per_kg": 3.50, "mrr_cm3_min": 12.0},
        "stainless steel": {"density": 8.00, "cost_per_kg": 9.80, "mrr_cm3_min": 8.0},
        "aluminum": {"density": 2.70, "cost_per_kg": 8.50, "mrr_cm3_min": 35.0},
        "aluminum 6061": {"density": 2.70, "cost_per_kg": 8.50, "mrr_cm3_min": 35.0},
        "brass": {"density": 8.50, "cost_per_kg": 14.00, "mrr_cm3_min": 25.0},
        "titanium": {"density": 4.43, "cost_per_kg": 45.00, "mrr_cm3_min": 4.0},
        "plastic": {"density": 1.40, "cost_per_kg": 11.00, "mrr_cm3_min": 50.0},
    }

    def __init__(
        self,
        machining_hourly_rate: float = 75.0,
        stock_allowance_cm: float = 0.5,
        setup_minutes: float = 15.0,
    ):
        self.hourly_rate = machining_hourly_rate
        self.stock_allowance = stock_allowance_cm
        self.setup_minutes = setup_minutes

    def get_material_profile(self, material_name: Optional[str]) -> Dict[str, float]:
        """Resolves material profile from name or returns default mild steel values."""
        if not material_name:
            return self.MATERIAL_PROFILES["steel"]

        mat_lower = material_name.strip().lower()
        for key, profile in self.MATERIAL_PROFILES.items():
            if key in mat_lower:
                return profile
        return self.MATERIAL_PROFILES["steel"]

    def estimate_part_cost(
        self,
        bounding_box_cm: Dict[str, float],
        finished_volume_cm3: float,
        material_name: Optional[str] = "Steel, Mild",
        custom_density: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Estimates raw stock requirements, material scrap, machining time, and fabrication cost.

        Args:
            bounding_box_cm: Dictionary with 'dx', 'dy', 'dz' dimensions in centimeters.
            finished_volume_cm3: Net finished part volume in cubic centimeters.
            material_name: Material description string.
            custom_density: Optional density override in g/cm3.
        """
        profile = self.get_material_profile(material_name)
        density_g_cm3 = custom_density if custom_density and custom_density > 0 else profile["density"]
        cost_per_kg = profile["cost_per_kg"]
        mrr = profile["mrr_cm3_min"]

        # Calculate bounding stock with envelope margin allowance
        dx = bounding_box_cm.get("dx", 1.0)
        dy = bounding_box_cm.get("dy", 1.0)
        dz = bounding_box_cm.get("dz", 1.0)

        stock_dx = round(dx + self.stock_allowance, 2)
        stock_dy = round(dy + self.stock_allowance, 2)
        stock_dz = round(dz + self.stock_allowance, 2)

        stock_volume_cm3 = round(stock_dx * stock_dy * stock_dz, 2)
        stock_mass_kg = round((stock_volume_cm3 * density_g_cm3) / 1000.0, 3)

        # Finished mass & waste calculation
        finished_mass_kg = round((finished_volume_cm3 * density_g_cm3) / 1000.0, 3)
        scrap_volume_cm3 = max(0.0, round(stock_volume_cm3 - finished_volume_cm3, 2))
        material_removal_pct = round((scrap_volume_cm3 / stock_volume_cm3 * 100.0), 1) if stock_volume_cm3 > 0 else 0.0

        # Cost components
        material_cost = round(stock_mass_kg * cost_per_kg, 2)

        # Machining time: roughing removal + finish pass + setup
        machining_minutes = (scrap_volume_cm3 / mrr) if mrr > 0 else 10.0
        total_time_minutes = round(self.setup_minutes + machining_minutes, 1)
        machining_cost = round((total_time_minutes / 60.0) * self.hourly_rate, 2)
        total_cost = round(material_cost + machining_cost, 2)

        result = {
            "material": material_name,
            "density_g_cm3": density_g_cm3,
            "stock_dimensions_cm": {"dx": stock_dx, "dy": stock_dy, "dz": stock_dz},
            "stock_volume_cm3": stock_volume_cm3,
            "stock_mass_kg": stock_mass_kg,
            "finished_mass_kg": finished_mass_kg,
            "scrap_volume_cm3": scrap_volume_cm3,
            "material_removal_percentage": material_removal_pct,
            "machining_time_minutes": total_time_minutes,
            "material_cost_usd": material_cost,
            "machining_cost_usd": machining_cost,
            "total_estimated_cost_usd": total_cost,
        }

        logger.info(
            f"Cost estimate calculated for {material_name}: "
            f"Stock={stock_mass_kg}kg (${material_cost}), "
            f"Machining={total_time_minutes}min (${machining_cost}), "
            f"Total=${total_cost}"
        )
        return result
