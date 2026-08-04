import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.DrawingGenerator")

class DrawingGenerator:
    """Automates technical drawing document generation, sheet settings, and placing part views."""

    kDrawingDocumentObject = 12292

    # View Orientation Constants fallbacks
    kFrontViewOrientation = 10764
    kTopViewOrientation = 10765
    kIsoTopRightViewOrientation = 10774

    # View Style Constants fallbacks
    kHiddenLineDrawingViewStyle = 32258
    kHiddenLineRemovedDrawingViewStyle = 32257
    kShadedDrawingViewStyle = 32259

    def __init__(self, inventor_app: Any, drawing_settings: Dict[str, Any]):
        self.app = inventor_app
        self.settings = drawing_settings
        self.tg = self.app.TransientGeometry

    def create_drawing(self, part_doc: Any, part_number: str, part_name: str) -> Any:
        """Creates a drawing document for a part and places Front, Top, and Isometric views."""
        try:
            logger.info(f"Generating drawing for part {part_number} ({part_name})...")
            
            # Update part properties so they propagate to drawing title block
            self._update_part_properties(part_doc, part_number, part_name)

            # Create drawing from template
            template_path = self.app.FileManager.GetTemplateFile(self.kDrawingDocumentObject)
            doc = self.app.Documents.Add(self.kDrawingDocumentObject, template_path, True)
            drawing_doc = CastTo(doc, "DrawingDocument")
            
            sheet = drawing_doc.Sheets.Item(1)
            
            # Configure Sheet Size
            self._configure_sheet_size(sheet)
            
            # Add Views (Base + Projected)
            self._place_views(sheet, part_doc)

            logger.info(f"Successfully generated drawing for {part_number}")
            return drawing_doc
        except Exception as e:
            logger.error(f"Failed to generate drawing: {e}")
            raise

    def _update_part_properties(self, part_doc: Any, part_number: str, part_name: str) -> None:
        """Sets internal standard properties (iProperties) of the part."""
        try:
            # Set Title and Part Number in "Design Tracking Properties"
            design_props = part_doc.PropertySets.Item("Design Tracking Properties")
            design_props.Item("Part Number").Value = part_number
            design_props.Item("Description").Value = part_name
            
            # Set Author in "Summary Information"
            summary_props = part_doc.PropertySets.Item("Inventor Summary Information")
            summary_props.Item("Author").Value = "DevOps Automation"
            logger.info(f"Updated iProperties for {part_number}.")
        except Exception as e:
            logger.warning(f"Could not update part iProperties: {e}")

    def _configure_sheet_size(self, sheet: Any) -> None:
        """Sets drawing sheet size (e.g. A3, A4) from configuration settings."""
        size_str = self.settings.get("sheet_size", "A3").upper()
        
        try:
            # SheetSizeEnum mapping
            size_map = {
                "A0": 9990, "A1": 9991, "A2": 9992, "A3": 9993, "A4": 9994,
                "A": 9985, "B": 9986, "C": 9987, "D": 9988, "E": 9989
            }
            
            size_enum = size_map.get(size_str, 9993)  # default A3
            try:
                # If constants are loaded, override with exact enum constant
                enum_name = f"k{size_str}SheetSize"
                if hasattr(constants, enum_name):
                    size_enum = getattr(constants, enum_name)
            except Exception:
                pass

            sheet.Size = size_enum
            logger.info(f"Set sheet size to {size_str}")
        except Exception as e:
            logger.warning(f"Could not set sheet size: {e}. Using default.")

    def _place_views(self, sheet: Any, part_doc: Any) -> None:
        """Places base view (Front) and projected views (Top, Isometric)."""
        views = sheet.DrawingViews
        
        # Calculate positions in cm (Drawing sheet coordinates)
        # Assuming sheet width is ~42cm (A3) and height is ~29.7cm
        w = sheet.Width
        h = sheet.Height
        
        # Base View (Front View) placed at center-left
        front_pos = self.tg.CreatePoint2d(w * 0.3, h * 0.4)
        
        # Projected Top View placed directly above base view
        top_pos = self.tg.CreatePoint2d(w * 0.3, h * 0.75)
        
        # Projected Isometric View placed at top-right
        iso_pos = self.tg.CreatePoint2d(w * 0.75, h * 0.7)

        # Get orientation and styles
        try:
            front_orient = constants.kFrontViewOrientation
            top_style = constants.kHiddenLineDrawingViewStyle
            iso_orient = constants.kIsoTopRightViewOrientation
            iso_style = constants.kShadedDrawingViewStyle
        except AttributeError:
            front_orient = self.kFrontViewOrientation
            top_style = self.kHiddenLineDrawingViewStyle
            iso_orient = self.kIsoTopRightViewOrientation
            iso_style = self.kShadedDrawingViewStyle

        view_scale = float(self.settings.get("view_scale", 1.0))

        try:
            # 1. Create Base View (Front)
            logger.info("Placing Front View...")
            base_view = views.AddBaseView(
                part_doc, 
                front_pos, 
                view_scale, 
                front_orient, 
                top_style
            )
            
            # 2. Create Projected Top View
            logger.info("Placing Projected Top View...")
            views.AddProjectedView(base_view, top_pos, top_style)
            
            # 3. Create Projected Isometric View
            logger.info("Placing Projected Isometric View...")
            # We can create a base view for Isometric directly, or project it.
            # Direct base view placement of Isometric is often cleaner for layouts.
            views.AddBaseView(
                part_doc, 
                iso_pos, 
                view_scale, 
                iso_orient, 
                iso_style
            )
            
            logger.info("Drawing views placed successfully.")
        except Exception as e:
            logger.error(f"Error placing drawing views on sheet: {e}")
            raise
