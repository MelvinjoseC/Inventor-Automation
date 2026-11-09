import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.DrawingGenerator")


class DrawingGenerator:
    """Automates technical drawing document generation, sheet settings, intelligent view layouts,
    dimensions, center marks, and title block synchronization.
    """

    kDrawingDocumentObject = 12292

    # View Orientation Constants fallbacks
    kFrontViewOrientation = 10764
    kTopViewOrientation = 10765
    kRightViewOrientation = 10766
    kIsoTopRightViewOrientation = 10774

    # View Style Constants fallbacks
    kHiddenLineDrawingViewStyle = 32258
    kHiddenLineRemovedDrawingViewStyle = 32257
    kShadedDrawingViewStyle = 32259

    def __init__(self, inventor_app: Any, drawing_settings: Optional[Dict[str, Any]] = None):
        self.app = inventor_app
        self.settings = drawing_settings or {}
        self.tg = self.app.TransientGeometry

    def create_drawing(
        self,
        part_doc: Any,
        part_number: str,
        part_name: str,
        revision: str = "A",
        material: Optional[str] = None,
    ) -> Any:
        """Creates a drawing document for a part, configures sheet, and places orthographic and isometric views."""
        try:
            logger.info(f"Generating drawing for part {part_number} ({part_name})...")

            # Update part properties so they propagate to drawing title block
            self._update_part_properties(part_doc, part_number, part_name, revision)

            # Create drawing from template
            template_path = self.app.FileManager.GetTemplateFile(self.kDrawingDocumentObject)
            doc = self.app.Documents.Add(self.kDrawingDocumentObject, template_path, True)
            try:
                drawing_doc = CastTo(doc, "DrawingDocument")
            except Exception:
                drawing_doc = doc

            sheet = drawing_doc.Sheets.Item(1)

            # Configure Sheet Size
            self._configure_sheet_size(sheet)

            # Add Views (Base + Projected + Isometric)
            base_view = self._place_views(sheet, part_doc)

            # Optional automated center marks & dimensions
            if self.settings.get("auto_centermarks", True) and base_view:
                self._add_centermarks(sheet, base_view)

            logger.info(f"Successfully generated drawing for {part_number}")
            return drawing_doc
        except Exception as e:
            logger.error(f"Failed to generate drawing: {e}")
            raise

    def create_assembly_drawing(
        self,
        assembly_doc: Any,
        assembly_number: str,
        assembly_name: str,
        revision: str = "A",
        include_parts_list: bool = True,
        include_balloons: bool = True,
        include_revision_table: bool = True,
    ) -> Any:
        """Creates an assembly drawing sheet with orthographic views, isometric view, automatic Parts List,
        balloon annotations, and revision tracking table.
        """
        try:
            logger.info(f"Generating assembly drawing for {assembly_number} ({assembly_name})...")
            self._update_part_properties(assembly_doc, assembly_number, assembly_name, revision)

            template_path = self.app.FileManager.GetTemplateFile(self.kDrawingDocumentObject)
            doc = self.app.Documents.Add(self.kDrawingDocumentObject, template_path, True)
            try:
                drawing_doc = CastTo(doc, "DrawingDocument")
            except Exception:
                drawing_doc = doc

            sheet = drawing_doc.Sheets.Item(1)
            self._configure_sheet_size(sheet)

            # Place assembly views
            base_view = self._place_views(sheet, assembly_doc)

            # Insert Parts List (Drawing BOM Table)
            parts_list = None
            if include_parts_list and base_view:
                parts_list = self._add_parts_list(sheet, base_view)

            # Automated Item Balloons
            if include_balloons and base_view:
                self._add_balloons(sheet, base_view)

            # Revision History Table
            if include_revision_table:
                self._add_revision_table(sheet, revision=revision)

            logger.info(f"Successfully generated assembly drawing for {assembly_number}")
            return drawing_doc
        except Exception as e:
            logger.error(f"Failed to generate assembly drawing: {e}")
            raise

    def _add_balloons(self, sheet: Any, base_view: Any) -> int:
        """Places item callout balloons on the drawing sheet referencing the assembly parts list."""
        balloon_count = 0
        try:
            balloons = getattr(sheet, "Balloons", None)
            if not balloons:
                return 0

            # Attempt native AutoBalloon if supported
            if hasattr(balloons, "AutoBalloon"):
                curves = getattr(base_view, "DrawingCurves", None)
                if curves and curves.Count > 0:
                    leader_points = self.tg.CreatePoint2d(base_view.Center.X + 6.0, base_view.Center.Y + 4.0)
                    balloons.AutoBalloon(base_view, leader_points)
                    logger.info("Auto-ballooning applied across assembly views.")
                    return 1

            # Fallback: manual placement for candidate curves
            curves = getattr(base_view, "DrawingCurves", None)
            if curves and hasattr(balloons, "Add"):
                limit = min(curves.Count + 1, 6)
                for i in range(1, limit):
                    curve = curves.Item(i)
                    try:
                        intent = sheet.CreateGeometryIntent(curve)
                        pos = self.tg.CreatePoint2d(base_view.Center.X + (i * 2.0), base_view.Center.Y + 3.0)
                        balloons.Add(intent, pos)
                        balloon_count += 1
                    except Exception:
                        continue
            logger.info(f"Added {balloon_count} item balloons to sheet.")
        except Exception as e:
            logger.debug(f"Could not automatically place balloons: {e}")
        return balloon_count

    def _add_revision_table(
        self,
        sheet: Any,
        revision: str = "A",
        description: str = "Automated Production Release",
    ) -> Optional[Any]:
        """Inserts an engineering revision tracking table onto the drawing sheet."""
        try:
            rev_tables = getattr(sheet, "RevisionTables", None)
            if not rev_tables or not hasattr(rev_tables, "Add"):
                return None

            pos = self.tg.CreatePoint2d(2.0, float(sheet.Height) - 2.0)
            rev_table = rev_tables.Add(pos)

            # Populate revision entry
            try:
                rows = rev_table.RevisionTableRows
                if rows.Count > 0:
                    last_row = rows.Item(rows.Count)
                    if hasattr(last_row, "Item"):
                        # Set Revision Tag, Description, Approved
                        last_row.Item("REV").Text = revision
                        last_row.Item("DESCRIPTION").Text = description
            except Exception:
                pass

            logger.info(f"Revision table (Rev {revision}) inserted on drawing sheet.")
            return rev_table
        except Exception as e:
            logger.debug(f"Could not automatically place revision table: {e}")
            return None

    def _add_parts_list(self, sheet: Any, base_view: Any) -> Optional[Any]:
        """Inserts a structured Parts List table onto the top-right corner of the drawing sheet."""
        try:
            w = float(sheet.Width)
            h = float(sheet.Height)
            # Position at top-right corner with 2cm margin
            pos = self.tg.CreatePoint2d(w - 14.0, h - 2.0)
            parts_lists = sheet.PartsLists
            parts_list = parts_lists.Add(base_view, pos)
            logger.info("Structured Parts List inserted on drawing sheet.")
            return parts_list
        except Exception as e:
            logger.warning(f"Could not automatically place Parts List: {e}")
            return None

    def _update_part_properties(
        self,
        part_doc: Any,
        part_number: str,
        part_name: str,
        revision: str = "A",
    ) -> None:
        """Sets internal standard properties (iProperties) of the part."""
        try:
            design_props = part_doc.PropertySets.Item("Design Tracking Properties")
            design_props.Item("Part Number").Value = part_number
            design_props.Item("Description").Value = part_name

            summary_props = part_doc.PropertySets.Item("Inventor Summary Information")
            summary_props.Item("Author").Value = "DevOps Automation"
            summary_props.Item("Revision Number").Value = revision
            logger.info(f"Updated iProperties for {part_number} (Rev {revision}).")
        except Exception as e:
            logger.warning(f"Could not update part iProperties: {e}")

    def _configure_sheet_size(self, sheet: Any) -> None:
        """Sets drawing sheet size (e.g. A3, A4, B, D) from configuration settings."""
        size_str = self.settings.get("sheet_size", "A3").upper()

        try:
            size_map = {
                "A0": 9990, "A1": 9991, "A2": 9992, "A3": 9993, "A4": 9994,
                "A": 9985, "B": 9986, "C": 9987, "D": 9988, "E": 9989,
            }
            size_enum = size_map.get(size_str, 9993)  # default A3
            try:
                enum_name = f"k{size_str}SheetSize"
                if hasattr(constants, enum_name):
                    size_enum = getattr(constants, enum_name)
            except Exception:
                pass

            sheet.Size = size_enum
            logger.info(f"Set sheet size to {size_str}")
        except Exception as e:
            logger.warning(f"Could not set sheet size: {e}. Using default.")

    def _place_views(self, sheet: Any, part_doc: Any) -> Any:
        """Places base view (Front), projected views (Top, Right), and Isometric View."""
        views = sheet.DrawingViews
        w = float(sheet.Width)
        h = float(sheet.Height)

        # Dynamic layout coordinates based on sheet proportions
        front_pos = self.tg.CreatePoint2d(w * 0.3, h * 0.4)
        top_pos = self.tg.CreatePoint2d(w * 0.3, h * 0.75)
        iso_pos = self.tg.CreatePoint2d(w * 0.75, h * 0.7)

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
            # 1. Base Front View
            logger.info("Placing Front View...")
            base_view = views.AddBaseView(part_doc, front_pos, view_scale, front_orient, top_style)

            # 2. Projected Top View
            logger.info("Placing Projected Top View...")
            views.AddProjectedView(base_view, top_pos, top_style)

            # 3. Projected / Base Isometric View
            logger.info("Placing Isometric View...")
            views.AddBaseView(part_doc, iso_pos, view_scale, iso_orient, iso_style)

            logger.info("Drawing views placed successfully.")
            return base_view
        except Exception as e:
            logger.error(f"Error placing drawing views on sheet: {e}")
            raise

    def _add_centermarks(self, sheet: Any, view: Any) -> int:
        """Iterates over circular curves in view and places center marks."""
        count = 0
        try:
            centermarks = getattr(sheet, "Centermarks", None)
            if not centermarks or not hasattr(centermarks, "Add"):
                return 0

            curves = getattr(view, "DrawingCurves", None)
            if curves:
                for i in range(1, min(curves.Count + 1, 15)):
                    curve = curves.Item(i)
                    # Check if curve type is circle / circular arc (kCircleCurve = 5124)
                    curve_type = getattr(curve, "CurveType", None)
                    if curve_type in (5124, 5125):
                        try:
                            intent = sheet.CreateGeometryIntent(curve)
                            centermarks.Add(intent)
                            count += 1
                        except Exception:
                            continue
            if count > 0:
                logger.info(f"Added {count} center marks to drawing sheet.")
        except Exception as e:
            logger.debug(f"Could not automatically place center marks: {e}")
        return count
