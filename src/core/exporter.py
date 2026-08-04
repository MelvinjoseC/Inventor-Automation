import logging
from pathlib import Path
from typing import Any
from win32com.client import CastTo, constants

logger = logging.getLogger("InventorAutomation.Exporter")

class InventorExporter:
    """Manages document format translation (exporting to DWG, DXF, PDF)."""

    # Inventor Translator CLSIDs
    DWG_DXF_CLSID = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"
    PDF_CLSID = "{0AC6DF95-25A1-11D5-8E73-0010B541CD80}"

    def __init__(self, inventor_app: Any):
        self.app = inventor_app

    def export_pdf(self, doc: Any, output_path: Path) -> None:
        """Exports a document (Part, Assembly, or Drawing) as PDF."""
        logger.info(f"Exporting to PDF: {output_path}")
        self._translate_file(doc, self.PDF_CLSID, output_path, is_pdf=True)

    def export_dwg(self, doc: Any, output_path: Path) -> None:
        """Exports a Drawing document as AutoCAD DWG."""
        logger.info(f"Exporting to DWG: {output_path}")
        self._translate_file(doc, self.DWG_DXF_CLSID, output_path, is_pdf=False)

    def export_dxf(self, doc: Any, output_path: Path) -> None:
        """Exports a sheet drawing or flat pattern sheet as DXF."""
        logger.info(f"Exporting to DXF: {output_path}")
        self._translate_file(doc, self.DWG_DXF_CLSID, output_path, is_pdf=False, is_dxf=True)

    def _translate_file(
        self, 
        doc: Any, 
        clsid: str, 
        output_path: Path, 
        is_pdf: bool = False, 
        is_dxf: bool = False
    ) -> None:
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
            translator = CastTo(translator, "TranslatorAddIn")
            
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
                if is_pdf:
                    # PDF Options
                    options.Value["All_Sheets"] = 1
                    options.Value["Vector_Resolution"] = 4800
                elif is_dxf:
                    # DXF specific config if applicable
                    try:
                        options.Value["DXFVersion"] = 31  # AutoCAD 2018 DXF
                    except Exception:
                        pass
                else:
                    # DWG Options
                    try:
                        options.Value["Export_Units"] = 1  # mm or inches
                    except Exception:
                        pass

            logger.info(f"Calling SaveCopyAs for translator {clsid}")
            translator.SaveCopyAs(doc, context, options, data_medium)
            logger.info(f"Export complete: {target}")
            
        except Exception as e:
            logger.error(f"Translation/Export failed for {target}: {e}")
            raise RuntimeError(f"Export failed: {e}") from e
