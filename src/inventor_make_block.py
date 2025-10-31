# src/inventor_make_block.py
from pathlib import Path
from win32com.client import gencache, CastTo, constants

PROGID = "Inventor.Application"

def get_inventor():
    try:
        inv = gencache.EnsureDispatch(PROGID)  # loads type lib (gives constants + typed interfaces)
    except Exception:
        inv = gencache.EnsureDispatch(PROGID)
    inv.Visible = True
    return inv

def new_part(inv):
    kPartDoc = 12290  # DocumentTypeEnum.kPartDocumentObject
    template = inv.FileManager.GetTemplateFile(kPartDoc)
    doc = inv.Documents.Add(kPartDoc, template, True)
    # Cast to PartDocument so .ComponentDefinition exists
    return CastTo(doc, "PartDocument")

def make_block(part_doc, width_cm=2.0, height_cm=1.0, thickness_cm=0.5):
    app = part_doc.Parent
    cd = part_doc.ComponentDefinition
    tg = app.TransientGeometry

    # Sketch on XY plane
    xy = cd.WorkPlanes.Item(3)
    sk = cd.Sketches.Add(xy)

    p1 = tg.CreatePoint2d(0, 0)
    p2 = tg.CreatePoint2d(width_cm, height_cm)  # geometry units = cm
    sk.SketchLines.AddAsTwoPointRectangle(p1, p2)

    # Profile
    prof = sk.Profiles.AddForSolid()

    # Extrude
    extrudes = cd.Features.ExtrudeFeatures
    ext_def = extrudes.CreateExtrudeDefinition(prof, constants.kJoinOperation)
    ext_def.SetDistanceExtent(thickness_cm, constants.kPositiveExtentDirection)
    extrudes.Add(ext_def)

    return sk

def save_part(part_doc, path: Path):
    target = path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    part_doc.SaveAs(str(target), True)

def export_part_as_dwg(inv, part_doc, path: Path):
    target = path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()

    cls_id = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"  # Translator: DWG
    translator = inv.ApplicationAddIns.ItemById(cls_id)
    if not translator.Activated:
        translator.Activate()
    translator = CastTo(translator, "TranslatorAddIn")

    context = inv.TransientObjects.CreateTranslationContext()
    try:
        context.Type = constants.kFileBrowseIOMechanism
    except AttributeError:
        # Fallback for Inventor versions without the constant
        context.Type = 1

    options = inv.TransientObjects.CreateNameValueMap()
    data_medium = inv.TransientObjects.CreateDataMedium()
    data_medium.FileName = str(target)

    try:
        if translator.HasSaveCopyAsOptions(part_doc, context, options):
            pass
    except Exception:
        # Some Inventor versions raise if options are not supported; ignore.
        options = inv.TransientObjects.CreateNameValueMap()
    translator.SaveCopyAs(part_doc, context, options, data_medium)

def main():
    inv = get_inventor()

    part = new_part(inv)
    make_block(part, 2.0, 1.0, 5.5)  # 20x10x55 mm (thickness passed in cm)

    part_out = Path.home() / "Desktop" / "Block.ipt"
    save_part(part, part_out)

    dwg_out = Path.home() / "Desktop" / "OutputbyMcj.dwg"
    export_part_as_dwg(inv, part, dwg_out)

    print(f"Created part: {part_out}")
    print(f"Exported DWG: {dwg_out}")

if __name__ == "__main__":
    main()
