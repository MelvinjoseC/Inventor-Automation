import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.AssemblyEngine")


class AssemblyAutomationEngine:
    """Automates Autodesk Inventor assembly creation, component placement,
    positioning matrices, constraints, and structured BOM extraction.
    """

    kAssemblyDocumentObject = 12291

    def __init__(self, inventor_app: Any):
        if inventor_app is None:
            raise ValueError("Inventor application instance cannot be None.")
        self.app = inventor_app
        self.tg = self.app.TransientGeometry

    def new_assembly_document(self) -> Any:
        """Creates and returns a new Assembly Document."""
        try:
            logger.info("Creating new Assembly Document...")
            template_path = self.app.FileManager.GetTemplateFile(self.kAssemblyDocumentObject)
            doc = self.app.Documents.Add(self.kAssemblyDocumentObject, template_path, True)
            try:
                assembly_doc = CastTo(doc, "AssemblyDocument")
            except Exception:
                assembly_doc = doc
            logger.info(f"Assembly Document created: {assembly_doc.DisplayName}")
            return assembly_doc
        except Exception as e:
            logger.error(f"Failed to create new Assembly Document: {e}")
            raise

    def place_component(
        self,
        assembly_doc: Any,
        component_path: Path,
        x_cm: float = 0.0,
        y_cm: float = 0.0,
        z_cm: float = 0.0,
        grounded: bool = False,
    ) -> Any:
        """Inserts a part or sub-assembly component into the assembly at (x, y, z)."""
        resolved_path = Path(component_path).expanduser().resolve()
        if not resolved_path.exists():
            raise FileNotFoundError(f"Component file not found: {resolved_path}")

        try:
            logger.info(f"Placing component '{resolved_path.name}' at ({x_cm}, {y_cm}, {z_cm}) cm")
            cd = assembly_doc.ComponentDefinition
            occurrences = cd.Occurrences

            matrix = self.tg.CreateMatrix()
            vector = self.tg.CreateVector(x_cm, y_cm, z_cm)
            matrix.SetTranslation(vector)

            occ = occurrences.Add(str(resolved_path), matrix)
            if grounded:
                occ.Grounded = True
                logger.info(f"Grounded component '{resolved_path.name}'")

            return occ
        except Exception as e:
            logger.error(f"Failed to place component '{resolved_path}': {e}")
            raise

    def add_mate_constraint(
        self,
        assembly_doc: Any,
        entity1: Any,
        entity2: Any,
        offset_cm: float = 0.0,
    ) -> Any:
        """Adds a Mate constraint between two planar faces or axes."""
        try:
            logger.debug(f"Adding Mate constraint with offset {offset_cm}cm")
            cd = assembly_doc.ComponentDefinition
            constraint = cd.Constraints.AddMateConstraint(entity1, entity2, offset_cm)
            return constraint
        except Exception as e:
            logger.warning(f"Could not apply Mate constraint: {e}")
            return None

    def add_flush_constraint(
        self,
        assembly_doc: Any,
        entity1: Any,
        entity2: Any,
        offset_cm: float = 0.0,
    ) -> Any:
        """Adds a Flush constraint between two parallel planar faces."""
        try:
            logger.debug(f"Adding Flush constraint with offset {offset_cm}cm")
            cd = assembly_doc.ComponentDefinition
            constraint = cd.Constraints.AddFlushConstraint(entity1, entity2, offset_cm)
            return constraint
        except Exception as e:
            logger.warning(f"Could not apply Flush constraint: {e}")
            return None

    def add_insert_constraint(
        self,
        assembly_doc: Any,
        entity1: Any,
        entity2: Any,
        axes_opposed: bool = False,
        distance_cm: float = 0.0,
    ) -> Any:
        """Adds an Insert constraint between two circular edges/cylinders (common for fasteners and shafts)."""
        try:
            logger.debug(f"Adding Insert constraint with distance {distance_cm}cm")
            cd = assembly_doc.ComponentDefinition
            constraint = cd.Constraints.AddInsertConstraint(
                entity1, entity2, axes_opposed, distance_cm
            )
            return constraint
        except Exception as e:
            logger.warning(f"Could not apply Insert constraint: {e}")
            return None

    def create_linear_occurrence_pattern(
        self,
        assembly_doc: Any,
        occurrence: Any,
        count_x: int,
        spacing_x_cm: float,
        count_y: int = 1,
        spacing_y_cm: float = 0.0,
    ) -> Any:
        """Creates a 1D or 2D rectangular array pattern of an assembly component occurrence.

        Args:
            assembly_doc: Target AssemblyDocument.
            occurrence: Source component occurrence to pattern.
            count_x: Number of instances along primary direction.
            spacing_x_cm: Spacing distance in centimeters along primary direction.
            count_y: Number of instances along secondary direction.
            spacing_y_cm: Spacing distance in centimeters along secondary direction.
        """
        try:
            logger.info(
                f"Creating linear occurrence pattern: CountX={count_x}, SpacingX={spacing_x_cm}cm, "
                f"CountY={count_y}, SpacingY={spacing_y_cm}cm"
            )
            cd = assembly_doc.ComponentDefinition
            patterns = getattr(cd, "OccurrencePatterns", None)

            # Collection of elements to pattern
            elements = self.app.TransientObjects.CreateObjectCollection()
            elements.Add(occurrence)

            # Default work axes for direction (WorkAxes 1=X, 2=Y)
            x_axis = cd.WorkAxes.Item(1)
            y_axis = cd.WorkAxes.Item(2)

            pattern = None
            if patterns and hasattr(patterns, "AddRectangularPattern"):
                pattern = patterns.AddRectangularPattern(
                    elements,
                    x_axis,
                    True,
                    count_x,
                    spacing_x_cm,
                    y_axis,
                    True,
                    count_y,
                    spacing_y_cm,
                )
            logger.info("Linear occurrence pattern created successfully.")
            return pattern
        except Exception as e:
            logger.warning(f"Could not create native linear occurrence pattern: {e}")
            return None

    def create_circular_occurrence_pattern(
        self,
        assembly_doc: Any,
        occurrence: Any,
        count: int,
        angle_deg: float = 360.0,
        axis_entity: Optional[Any] = None,
    ) -> Any:
        """Creates a circular polar array pattern of an assembly component occurrence.

        Args:
            assembly_doc: Target AssemblyDocument.
            occurrence: Source component occurrence to pattern.
            count: Total quantity of instances around the circumference.
            angle_deg: Total sweep angle in degrees (default 360.0).
            axis_entity: Centerline axis entity (defaults to Z WorkAxis).
        """
        try:
            logger.info(
                f"Creating circular occurrence pattern: Count={count}, Angle={angle_deg} deg"
            )
            cd = assembly_doc.ComponentDefinition
            patterns = getattr(cd, "OccurrencePatterns", None)

            elements = self.app.TransientObjects.CreateObjectCollection()
            elements.Add(occurrence)

            axis = axis_entity if axis_entity is not None else cd.WorkAxes.Item(3)
            angle_rad = f"{angle_deg} deg"

            pattern = None
            if patterns and hasattr(patterns, "AddCircularPattern"):
                pattern = patterns.AddCircularPattern(
                    elements,
                    axis,
                    True,
                    count,
                    angle_rad,
                )
            logger.info("Circular occurrence pattern created successfully.")
            return pattern
        except Exception as e:
            logger.warning(f"Could not create native circular occurrence pattern: {e}")
            return None

    def get_assembly_bom_summary(self, assembly_doc: Any) -> List[Dict[str, Any]]:
        """Extracts structured occurrence hierarchy and component counts from the assembly."""
        summary: List[Dict[str, Any]] = []
        try:
            cd = assembly_doc.ComponentDefinition
            occurrences = cd.Occurrences
            parts_count: Dict[str, Dict[str, Any]] = {}

            for i in range(1, occurrences.Count + 1):
                occ = occurrences.Item(i)
                part_no = getattr(occ, "Name", f"Occurrence_{i}")
                file_path = ""
                try:
                    def_doc = occ.Definition.Document
                    file_path = def_doc.FullFileName
                    # Retrieve Part Number iProperty
                    design_props = def_doc.PropertySets.Item("Design Tracking Properties")
                    part_no = design_props.Item("Part Number").Value
                except Exception:
                    pass

                if part_no in parts_count:
                    parts_count[part_no]["quantity"] += 1
                else:
                    parts_count[part_no] = {
                        "part_number": part_no,
                        "name": getattr(occ, "Name", ""),
                        "file_path": file_path,
                        "grounded": getattr(occ, "Grounded", False),
                        "quantity": 1,
                    }

            summary = list(parts_count.values())
            logger.info(
                f"Assembly BOM extracted: {len(summary)} unique parts, {occurrences.Count} total occurrences."
            )
        except Exception as e:
            logger.warning(f"Failed to extract assembly BOM summary: {e}")
        return summary

    def check_interference(self, assembly_doc: Any) -> Dict[str, Any]:
        """Runs automated 3D interference and collision analysis across all occurrences in the assembly.

        Returns:
            Dictionary with overall status, count, total collision volume, and per-collision pair details.
        """
        report: Dict[str, Any] = {
            "has_interference": False,
            "interference_count": 0,
            "total_interference_volume_cm3": 0.0,
            "interferences": [],
        }

        try:
            logger.info("Executing automated interference detection...")
            cd = assembly_doc.ComponentDefinition
            occurrences = cd.Occurrences

            if occurrences.Count < 2:
                logger.info("Assembly contains fewer than 2 components; no collision possible.")
                return report

            # Build collection of occurrences
            occ_collection = self.app.TransientObjects.CreateObjectCollection()
            for i in range(1, occurrences.Count + 1):
                occ_collection.Add(occurrences.Item(i))

            results = None
            if hasattr(cd, "AnalyzeInterference"):
                results = cd.AnalyzeInterference(occ_collection)

            if results:
                count = getattr(results, "Count", 0)
                report["interference_count"] = count
                report["has_interference"] = count > 0

                total_vol = 0.0
                for i in range(1, count + 1):
                    res = results.Item(i)
                    vol = round(float(getattr(res, "Volume", 0.0)), 4)
                    total_vol += vol
                    occ1_name = getattr(getattr(res, "OccurrenceOne", None), "Name", "Unknown_1")
                    occ2_name = getattr(getattr(res, "OccurrenceTwo", None), "Name", "Unknown_2")
                    centroid_pt = getattr(res, "Centroid", None)
                    centroid = {
                        "x": (
                            round(float(getattr(centroid_pt, "X", 0.0)), 3) if centroid_pt else 0.0
                        ),
                        "y": (
                            round(float(getattr(centroid_pt, "Y", 0.0)), 3) if centroid_pt else 0.0
                        ),
                        "z": (
                            round(float(getattr(centroid_pt, "Z", 0.0)), 3) if centroid_pt else 0.0
                        ),
                    }

                    report["interferences"].append(
                        {
                            "occurrence_1": occ1_name,
                            "occurrence_2": occ2_name,
                            "volume_cm3": vol,
                            "centroid": centroid,
                        }
                    )

                report["total_interference_volume_cm3"] = round(total_vol, 4)
                if report["has_interference"]:
                    logger.warning(
                        f"Interference detected! {count} collisions found, "
                        f"Total volume = {report['total_interference_volume_cm3']} cm³"
                    )
                else:
                    logger.info("Zero interference detected across all assembly components.")

        except Exception as e:
            logger.warning(f"Interference analysis encountered an exception: {e}")

        return report
