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
            constraint = cd.Constraints.AddInsertConstraint(entity1, entity2, axes_opposed, distance_cm)
            return constraint
        except Exception as e:
            logger.warning(f"Could not apply Insert constraint: {e}")
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
            logger.info(f"Assembly BOM extracted: {len(summary)} unique parts, {occurrences.Count} total occurrences.")
        except Exception as e:
            logger.warning(f"Failed to extract assembly BOM summary: {e}")
        return summary
