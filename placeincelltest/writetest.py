import os
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED
from xlcellimage import ImageLoader

#############################################
# Test to see if I can write to an excel file through xml.etree
IMAGE_NAME = "image.png"
XL_PATH = "test.xlsx"
RICH_PATH = "xl/richData/_rels/richValueRel.xml.rels"

# This write the image to xl/media/*.png, but makes the sheet bugged. Upon auto-fixing, the image is removed.
# Probably need to add connections before opening again?
def uploadImg(imgb: bytes):
    with ZipFile(XL_PATH, "a") as zip:
        zip.writestr(f"xl/media/{IMAGE_NAME}", imgb)

def createRelationXML(xlPath: str, imgName: str):
    richValueRelExists = False
    with ZipFile(xlPath, "r") as zip:
        for item in zip.infolist():
            if item.filename == RICH_PATH:
                richValueRelExists = True
                break

    if richValueRelExists:

        # Append the relation
        tmp_path = xlPath + ".tmp"
        with ZipFile(xlPath, "r") as src, ZipFile(tmp_path, 'w', compression=ZIP_DEFLATED) as tmp:
            for item in src.infolist():
                if item.filename != RICH_PATH:
                    tmp.writestr(item, src.read(item.filename))
                else:

                    old_xml = src.read(item.filename)
                    root = ET.fromstring(old_xml)
                    rels_ns = "https://openxmlformats.org"
                    ET.register_namespace('', rels_ns)

                    # Generate new, unused rId:
                    idNumMax = 1
                    for rel in root.findall(f".//{{{rels_ns}}}Relationship"):
                        id = rel.get("Id")
                        if id:
                            match = re.search(r'\d+', id)
                            if match:
                                idNumMax = max(idNumMax, int(match.group()))
                                
                    _ = ET.SubElement(root, f"{{{rels_ns}}}Relationship", attrib={
                        "Id": f"rId{idNumMax+1}",
                        "Type": "http://openxmlformats.org",
                        "Target": f"../media/{imgName}"
                    })

                    new_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)                
                    tmp.writestr(RICH_PATH, new_xml)

        os.replace(tmp_path, xlPath)
        return

    
    # Function here to generate a new node
    # new_root = generateRelsXML(NameSpace, ......)
    rels_ns = "https://openxmlformats.org"
    ET.register_namespace('', rels_ns)

    new_root = ET.Element(f"{{{rels_ns}}}Relationships")
    _ = ET.SubElement(new_root, f"{{{rels_ns}}}Relationship", attrib={
        "Id": "rId1",
        "Type": "http://openxmlformats.org",
        "Target": f"../media/{imgName}"
    })

    new_xml = ET.tostring(new_root, encoding='utf-8', xml_declaration=True)  # pyright: ignore[reportAny]
    print("RichvalueRel doesn't exist yet, here is what is being put there now:")
    print(new_xml)

    # new_xml is the string data to be put in the file
    # Put it in file
    tmp_path = xlPath + ".tmp"
    with ZipFile(xlPath, "r") as src, ZipFile(tmp_path, 'w', compression=ZIP_DEFLATED) as tmp:
        for item in src.infolist():
            tmp.writestr(item, src.read(item.filename))

        tmp.writestr(RICH_PATH, new_xml)
        print("Rels file was NOT found. Here is the new file contents:")
        print(new_xml)

    os.replace(tmp_path, xlPath)




if __name__ == '__main__':
    # Get image bytes
    xli = ImageLoader("src.xlsx")
    imgb = xli.getImage("sheet1", "A1")

    uploadImg(imgb)
    createRelationXML(XL_PATH, IMAGE_NAME)
