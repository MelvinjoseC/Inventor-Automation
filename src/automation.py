import logging
from pathlib import Path
from typing import Any, Optional
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.Engine")

class InventorAutomationEngine:
    """Core automation engine for Autodesk Inventor tasks like modeling and document handling."""

    # DocumentTypeEnum constants
    kPartDocumentObject = 12290
    kAssemblyDocumentObject = 12291
    kDrawingDocumentObject = 12292

    def __init__(self, inventor_app: Any):
        if inventor_app is None:
            raise ValueError("Inventor application instance cannot be None.")
        self.app = inventor_app
        self.tg = self.app.TransientGeometry

    def new_part_document(self) -> Any:
        """Creates and returns a new Part Document."""
        try:
            logger.info("Creating a new Part Document...")
            template_path = self.app.FileManager.GetTemplateFile(self.kPartDocumentObject)
            doc = self.app.Documents.Add(self.kPartDocumentObject, template_path, True)
            part_doc = CastTo(doc, "PartDocument")
            logger.info(f"New Part Document created successfully. Handle: {part_doc.DisplayName}")
            return part_doc
        except Exception as e:
            logger.error(f"Failed to create new Part Document: {e}")
            raise

    def new_assembly_document(self) -> Any:
        """Creates and returns a new Assembly Document."""
        try:
            logger.info("Creating a new Assembly Document...")
            template_path = self.app.FileManager.GetTemplateFile(self.kAssemblyDocumentObject)
            doc = self.app.Documents.Add(self.kAssemblyDocumentObject, template_path, True)
            assembly_doc = CastTo(doc, "AssemblyDocument")
            logger.info(f"New Assembly Document created successfully. Handle: {assembly_doc.DisplayName}")
            return assembly_doc
        except Exception as e:
            logger.error(f"Failed to create new Assembly Document: {e}")
            raise

    def create_block(
        self, 
        part_doc: Any, 
        width_cm: float, 
        height_cm: float, 
        thickness_cm: float,
        material: Optional[str] = None,
        color: Optional[str] = None
    ) -> Any:
        """Generates a 3D block (rectangular cuboid) geometry on the XY work plane."""
        try:
            logger.info(
                f"Generating block: Width={width_cm}cm, Height={height_cm}cm, Thickness={thickness_cm}cm"
            )
            cd = part_doc.ComponentDefinition
            
            # Select the XY Plane (typically Index 3 in WorkPlanes collection)
            xy_plane = cd.WorkPlanes.Item(3)
            
            # Add a 2D sketch
            sketch = cd.Sketches.Add(xy_plane)
            
            # Draw a 2D Rectangle from origin (0,0)
            p1 = self.tg.CreatePoint2d(0, 0)
            p2 = self.tg.CreatePoint2d(width_cm, height_cm)
            sketch.SketchLines.AddAsTwoPointRectangle(p1, p2)
            
            # Create profiles
            profile = sketch.Profiles.AddForSolid()
            
            # Perform Extrusion
            extrudes = cd.Features.ExtrudeFeatures
            
            # Try using constant, fallback to integer values if constants are not loaded
            try:
                join_op = constants.kJoinOperation
                pos_dir = constants.kPositiveExtentDirection
            except AttributeError:
                join_op = 20481  # kJoinOperation
                pos_dir = 20993  # kPositiveExtentDirection

            ext_def = extrudes.CreateExtrudeDefinition(profile, join_op)
            ext_def.SetDistanceExtent(thickness_cm, pos_dir)
            
            extrusion = extrudes.Add(ext_def)
            logger.info("Extrusion feature created successfully.")

            # Apply material and color if specified
            if material:
                self.apply_material(part_doc, material)
            if color:
                self.apply_color(part_doc, color)

            return extrusion
        except Exception as e:
            logger.error(f"Failed to generate block geometry: {e}")
            raise

    def apply_material(self, part_doc: Any, material_name: str) -> None:
        """Applies a material to the part document."""
        try:
            cd = part_doc.ComponentDefinition
            # Attempt to retrieve the material from the document's materials database
            try:
                material = part_doc.Materials.Item(material_name)
                cd.Material = material
                logger.info(f"Applied material '{material_name}' to part.")
            except Exception:
                # If local materials collection doesn't work directly, try default material assignment
                logger.warning(
                    f"Material '{material_name}' not found in active document collection. Attempting fallback assignment."
                )
                cd.Material = cd.Material  # Keep existing material or default
        except Exception as e:
            logger.warning(f"Unable to apply material '{material_name}': {e}")

    def apply_color(self, part_doc: Any, color_name: str) -> None:
        """Applies a color style/appearance name to the part document."""
        try:
            # Set the active appearance of the part document
            # Note: Appearance structures differ between older/newer Inventor versions.
            try:
                appearance = part_doc.Appearances.Item(color_name)
                part_doc.ActiveAppearance = appearance
                logger.info(f"Applied color/appearance '{color_name}' to part.")
            except Exception:
                # Fallback to older render style property
                try:
                    render_style = part_doc.RenderStyles.Item(color_name)
                    part_doc.ActiveRenderStyle = render_style
                    logger.info(f"Applied render style '{color_name}' (fallback).")
                except Exception:
                    logger.warning(f"Appearance/Color style '{color_name}' could not be resolved in Inventor.")
        except Exception as e:
            logger.warning(f"Unable to apply color '{color_name}': {e}")

    def save_document(self, doc: Any, save_path: Path) -> None:
        """Saves any Inventor document to the given file path."""
        try:
            target = save_path.expanduser().resolve()
            target.parent.mkdir(parents=True, exist_ok=True)
            
            # If target file exists, delete it before overwriting to prevent save dialogue conflicts
            if target.exists():
                logger.info(f"File exists at {target}. Unlinking before overwrite...")
                target.unlink()
                
            logger.info(f"Saving document to: {target}")
            doc.SaveAs(str(target), True)
            logger.info("Document saved successfully.")
        except Exception as e:
            logger.error(f"Failed to save document to {save_path}: {e}")
            raise

    def close_document(self, doc: Any, save_changes: bool = False) -> None:
        """Closes the Inventor document."""
        try:
            doc.Close(SkipSave=(not save_changes))
            logger.info("Document closed.")
        except Exception as e:
            logger.warning(f"Failed to close document cleanly: {e}")
