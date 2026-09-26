import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.Exporter")


class InventorExporter:
    """Manages document format translation (exporting to DWG, DXF, PDF, STEP, IGES, STL, and raster images)."""

    # Official Autodesk Inventor Translator CLSIDs
    DWG_DXF_CLSID = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"
    PDF_CLSID = "{0AC6DF95-25A1-11D5-8E73-0010B541CD80}"
    STEP_CLSID = "{90AF203A-3FC1-4883-85FA-630C2918072E}"
    IGES_CLSID = "{90AF203B-3FC1-4883-85FA-630C2918072E}"
    STL_CLSID = "{53324084-226B-4895-A446-24E6E342E701}"
    SAT_CLSID = "{244F7499-D58B-4DC0-82C3-E0A96B9A8B74}"

    def __init__(self, inventor_app: Any):
        self.app = inventor_app

    def export_pdf(self, doc: Any, output_path: Path) -> Path:
        """Exports a document (Part, Assembly, or Drawing) as PDF."""
        logger.info(f"Exporting to PDF: {output_path}")
        return self._translate_file(doc, self.PDF_CLSID, output_path, format_type="pdf")

    def export_dwg(self, doc: Any, output_path: Path) -> Path:
        """Exports a Drawing document as AutoCAD DWG."""
        logger.info(f"Exporting to DWG: {output_path}")
        return self._translate_file(doc, self.DWG_DXF_CLSID, output_path, format_type="dwg")

    def export_dxf(self, doc: Any, output_path: Path) -> Path:
        """Exports a sheet drawing or flat pattern sheet as DXF."""
        logger.info(f"Exporting to DXF: {output_path}")
        return self._translate_file(doc, self.DWG_DXF_CLSID, output_path, format_type="dxf")

    def export_step(self, doc: Any, output_path: Path, protocol: int = 214) -> Path:
        """Exports a 3D Part or Assembly as STEP (AP214 or AP203)."""
        logger.info(f"Exporting to STEP: {output_path} (Protocol AP{protocol})")
        return self._translate_file(doc, self.STEP_CLSID, output_path, format_type="step", step_protocol=protocol)

    def export_iges(self, doc: Any, output_path: Path) -> Path:
        """Exports a 3D Part or Assembly as IGES."""
        logger.info(f"Exporting to IGES: {output_path}")
        return self._translate_file(doc, self.IGES_CLSID, output_path, format_type="iges")

    def export_stl(self, doc: Any, output_path: Path, resolution: str = "High") -> Path:
        """Exports a 3D Part or Assembly as STL for 3D printing / additive manufacturing."""
        logger.info(f"Exporting to STL: {output_path} (Resolution: {resolution})")
        return self._translate_file(doc, self.STL_CLSID, output_path, format_type="stl", stl_resolution=resolution)

    def export_image(
        self,
        doc: Any,
        output_path: Path,
        width: int = 1920,
        height: int = 1080,
    ) -> Path:
        """Exports a high-resolution 2D bitmap snapshot (PNG/JPEG) of the active view."""
        target = Path(output_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            target.unlink()

        try:
            logger.info(f"Capturing high-res view snapshot ({width}x{height}) to {target}")
            # Try saving via active view
            active_view = getattr(self.app, "ActiveView", None)
            if active_view and hasattr(active_view, "SaveAsBitmap"):
                active_view.SaveAsBitmap(str(target), width, height)
            else:
                # Fallback SaveCopyAs directly on document
                doc.SaveAs(str(target), True)
            logger.info(f"Snapshot successfully saved: {target}")
            return target
        except Exception as e:
            logger.error(f"Image snapshot export failed: {e}")
            raise RuntimeError(f"Snapshot export failed: {e}") from e

    def export_batch(
        self,
        doc: Any,
        base_name: str,
        output_dir: Path,
        formats: List[str],
        drawing_doc: Optional[Any] = None,
    ) -> Dict[str, Path]:
        """Runs batch export across multiple target formats, routing drawing/3D docs appropriately."""
        results: Dict[str, Path] = {}
        out_dir = Path(output_dir).expanduser().resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        for fmt in formats:
            fmt_clean = fmt.strip().lower()
            target_path = out_dir / f"{base_name}.{fmt_clean}"
            try:
                if fmt_clean == "pdf":
                    target_doc = drawing_doc if drawing_doc else doc
                    results["pdf"] = self.export_pdf(target_doc, target_path)
                elif fmt_clean == "dwg":
                    target_doc = drawing_doc if drawing_doc else doc
                    results["dwg"] = self.export_dwg(target_doc, target_path)
                elif fmt_clean == "dxf":
                    target_doc = drawing_doc if drawing_doc else doc
                    results["dxf"] = self.export_dxf(target_doc, target_path)
                elif fmt_clean in ("step", "stp"):
                    results["step"] = self.export_step(doc, target_path)
                elif fmt_clean in ("iges", "igs"):
                    results["iges"] = self.export_iges(doc, target_path)
                elif fmt_clean == "stl":
                    results["stl"] = self.export_stl(doc, target_path)
                elif fmt_clean in ("png", "jpg", "jpeg"):
                    results[fmt_clean] = self.export_image(doc, target_path)
                else:
                    logger.warning(f"Unrecognized export format requested: '{fmt_clean}'")
            except Exception as e:
                logger.error(f"Failed exporting '{fmt_clean}' for {base_name}: {e}")

        logger.info(f"Batch export completed: {len(results)}/{len(formats)} formats generated.")
        return results

    def _translate_file(
        self,
        doc: Any,
        clsid: str,
        output_path: Path,
        format_type: str = "pdf",
        step_protocol: int = 214,
        stl_resolution: str = "High",
    ) -> Path:
        """Generic translator invoke logic."""
        target = output_path.expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)

        if target.exists():
            logger.info(f"File exists at {target}. Removing before export...")
            target.unlink()

        try:
            translator = self.app.ApplicationAddIns.ItemById(clsid)
            if not translator.Activated:
                translator.Activate()
            try:
                translator = CastTo(translator, "TranslatorAddIn")
            except Exception:
                pass

            context = self.app.TransientObjects.CreateTranslationContext()
            try:
                context.Type = constants.kFileBrowseIOMechanism
            except AttributeError:
                context.Type = 1  # kFileBrowseIOMechanism raw value

            options = self.app.TransientObjects.CreateNameValueMap()
            data_medium = self.app.TransientObjects.CreateDataMedium()
            data_medium.FileName = str(target)

            # Define specific options based on output format
            if translator.HasSaveCopyAsOptions(doc, context, options):
                if format_type == "pdf":
                    options.Value["All_Sheets"] = 1
                    options.Value["Vector_Resolution"] = 4800
                elif format_type == "dxf":
                    try:
                        options.Value["DXFVersion"] = 31  # AutoCAD 2018 DXF
                    except Exception:
                        pass
                elif format_type == "dwg":
                    try:
                        options.Value["Export_Units"] = 1
                    except Exception:
                        pass
                elif format_type == "step":
                    try:
                        options.Value["ApplicationProtocolType"] = 3 if step_protocol == 214 else 2
                        options.Value["IncludeSketches"] = 0
                    except Exception:
                        pass
                elif format_type == "stl":
                    try:
                        options.Value["ExportUnits"] = 4  # Millimeters
                        options.Value["Resolution"] = 1 if stl_resolution.lower() == "high" else 2
                    except Exception:
                        pass

            logger.info(f"Calling SaveCopyAs for translator {clsid}")
            translator.SaveCopyAs(doc, context, options, data_medium)
            logger.info(f"Export complete: {target}")
            return target

        except Exception as e:
            logger.error(f"Translation/Export failed for {target}: {e}")
            raise RuntimeError(f"Export failed: {e}") from e
