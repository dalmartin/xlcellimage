import xml.etree.ElementTree as ET
from zipfile import ZipFile

#######################################################################
# WorkbookParser is a simple class that opens a workbook and parses it 
# to lazy load mappings from cells -> images and eager load images as 
# bytes in self.images.
#
# Use WorkbookParser.getImage(("sheet1", "A1")) to get the image stored
# in the cell A1.
#######################################################################

##############Files to Parse ##########################################
METADATAFILES = [
                    "xl/metadata.xml",
                    "xl/richData/richValueRel.xml",
                    "xl/richData/_rels/richValueRel.xml.rels",
                ]
################### Namespace Constant for XML parsing #################
NS = {
         "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
         "rvr": "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
         "relations": "http://schemas.openxmlformats.org/package/2006/relationships",
     }
########################################################################

class WorkbookParser:

    def __init__(self, path: str):
        # Cached workbook values to reference in O(1)
        self.zip: ZipFile = ZipFile(path)
        self.parsed: dict[str, ET.Element] = {}
        self.images: dict[str, bytes] = {}

        # Cached data paths to images
        #Intermediate dictionaries:
        self.cellToVM: dict[tuple[str, str], str] = {}
        self.cellToV : dict[tuple[str, str], str] = {}
        self.cellToRID : dict[tuple[str, str], str] = {}
        self.cellToPath: dict[tuple[str, str], str] = {}

        # Read workbook
        self.sheets: list[str] = [n for n in self.zip.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
        self.parsefiles: list[str] = METADATAFILES + self.sheets
        
        for parsefile in self.parsefiles:
            self._readXml(parsefile)

        # Populate data paths
        self._getVMs()
        self._getV()
        self._getRID()
        self._getImgPath()

    ######## Eager load information (except for image bytes) from the workbook

    def _readXml(self, path: str):
        bytes = self.zip.read(path)
        root = ET.fromstring(bytes)
        self.parsed[path] = root

    def _getVMs(self):
        # populate self.sheetcellIMG with mappings of (sheet, cell) -> img path
        for sheet in self.sheets:
            # get sheet name
            sheetName: str = ""
            i_start: int = sheet.find("xl/worksheets/") + len("xl/worksheets/")
            i_end: int = sheet.find(".xml")
            if i_start != -1 and i_end != -1:
                sheetName = sheet[i_start:i_end]
            else:
                raise NameError(f"{sheet} is not named properly; Failed to store sheet name")

            # Get sheet data
            sheetData = self.getData(sheet)
            for cell in sheetData.findall(".//main:c", NS):
                vm = cell.get("vm")
                cell = cell.get('r')
                if vm is not None and cell is not None:
                    self.cellToVM[(sheetName, cell)] = vm

    def _getV(self):
        # get the metadata parsed
        data = self.getData("xl/metadata.xml")
        idxV: dict[int, str] = {}

        # populate a metadata value index to a metadata value
        for i, rc in enumerate(data.findall(".//main:rc", NS)):
            v = rc.get("v")
            if v:
                idxV[i+1] = v

        # populate cell to metadata value
        for cell, vm in self.cellToVM.items():
            self.cellToV[cell] = idxV[int(vm)]

    def _getRID(self):
        # get rich value relation data
        data = self.getData("xl/richData/richValueRel.xml")
        idxRID: dict[int, str] = {}

        for i, rel in enumerate(data.findall(".//rvr:rel", NS)):
            rid = rel.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if rid:
                idxRID[i] = rid

        for cell, v in self.cellToV.items():
            self.cellToRID[cell] = idxRID[int(v)]

    def _getImgPath(self):
        # _get rich value relationships (actual path to images)
        data = self.getData("xl/richData/_rels/richValueRel.xml.rels")
        ridImgPath: dict[str, str] = {}

        for relation in data.findall(".//relations:Relationship", NS):
            rid = relation.get("Id")
            path = relation.get("Target")
            if rid and path and relation.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                pathPrefix = "../media/"
                pathStart = path.find(pathPrefix) + len(pathPrefix)
                ridImgPath[rid] = path[pathStart:]

        for cell, rId in self.cellToRID.items():
            self.cellToPath[cell] = ridImgPath[rId]


    ######## API for the ImageLoader to retrieve cached information (getters)

    def getData(self, path: str) -> ET.Element:
        if path not in self.parsed:
            raise KeyError(f"Cannot retrieve xml data for the following file: {path}. specified path has not been parsed")

        return self.parsed[path]

    def hasImage(self, sheetCell: tuple[str, str]) -> bool:
        return sheetCell in self.cellToPath

    def getImage(self, sheetCell: tuple[str, str]) -> bytes:
        if not self.hasImage(sheetCell):
            raise KeyError(f"Attempted to access an image from a cell that does not contain an image: {sheetCell}.")

        file = self.cellToPath[(sheetCell)]

        if file not in self.images:
            img: bytes = self.zip.read(f"xl/media/{file}")
            if img:
                self.images[file] = img
            else:
                raise FileNotFoundError(f"Could not access image: {file}. Please check that this exists before attempting to access.")
        
        return self.images[file]
