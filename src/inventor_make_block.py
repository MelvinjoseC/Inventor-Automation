import sys
from pathlib import Path

# Add root directory to path to enable importing core modules
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.core.connection import InventorConnectionManager
from src.core.drawing_generator import DrawingGenerator
from src.core.exporter import InventorExporter
from src.automation import InventorAutomationEngine

def main():
    print("Starting legacy block generator (refactored integration)...")
    
    # Establish connection
    conn = InventorConnectionManager(visible=True, silent_mode=False)
    
    try:
        with conn as app:
            # Instantiate modular engines
            engine = InventorAutomationEngine(app)
            drawing_gen = DrawingGenerator(app, {"sheet_size": "A3", "view_scale": 1.0})
            exporter = InventorExporter(app)
            
            # 1. Create a new part document
            part_doc = engine.new_part_document()
            
            # 2. Model the block geometry: 20x10x55 mm (or 2.0 x 1.0 x 5.5 cm)
            engine.create_block(
                part_doc=part_doc, 
                width_cm=2.0, 
                height_cm=1.0, 
                thickness_cm=5.5,
                material="Steel, Mild",
                color="Polished"
            )
            
            # 3. Save part file (.ipt) to Desktop
            desktop_path = Path.home() / "Desktop"
            part_out = desktop_path / "Block.ipt"
            engine.save_document(part_doc, part_out)
            
            # 4. Generate drawing sheet for the part
            drawing_doc = drawing_gen.create_drawing(part_doc, "BM-BLOCK", "Auto-Generated Block Part")
            
            # 5. Export drawing to DWG
            dwg_out = desktop_path / "OutputbyMcj.dwg"
            engine.save_document(drawing_doc, dwg_out)
            
            # 6. Export drawing to PDF (new capability added by refactor!)
            pdf_out = desktop_path / "OutputbyMcj.pdf"
            exporter.export_pdf(drawing_doc, pdf_out)
            
            # Clean up documents
            engine.close_document(drawing_doc)
            engine.close_document(part_doc)
            
            print(f"✅ Success! Created part: {part_out}")
            print(f"✅ Success! Exported DWG: {dwg_out}")
            print(f"✅ Success! Exported PDF: {pdf_out}")
            
    except Exception as e:
        print(f"❌ Error occurred during automation script execution: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
