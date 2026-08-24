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
                    "xl/richData/rdrichvalue.xml",
                    "xl/richData/rdrichvaluestructure.xml",
                    "xl/richData/richValueRel.xml",
                    "xl/richData/_rels/richValueRel.xml.rels",
                ]
################### Namespace Constant for XML parsing #################
NS = {
         "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
         "rvr": "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
         "relations": "http://schemas.openxmlformats.org/package/2006/relationships",
         # xlrd: prefix inside metadata.xml, and the root namespace of
         # rdrichvalue.xml / rdrichvaluestructure.xml
         "richdata": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
         # namespace of the r:id attribute on richValueRel.xml's <rel> elements
         "odoc": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
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
        try:
            bytes = self.zip.read(path)
        except KeyError:
            # Some rich-data parts (e.g. rdrichvalue.xml / rdrichvaluestructure.xml)
            # only exist when the workbook actually contains in-cell images.
            # Treat an absent part as "no data" rather than raising, so a
            # workbook without rich-value images still parses cleanly and
            # simply resolves no images (see hasImage/getImage).
            return
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
        # Resolve each cell's vm (hop 1, already in cellToVM, 1-based) to the
        # futureMetadata block index (hops 2-4):
        #   2. valueMetadata/bk[vm - 1]   (vm is 1-based, so decrement)
        #   3. within that <bk>, a cell can have more than one <rc> (e.g. one
        #      for XLDAPR dynamic-array metadata and one for XLRICHVALUE), so
        #      find the specific <rc> whose t (1-based) indexes
        #      metadataTypes/metadataType[t - 1] and whose name is
        #      "XLRICHVALUE" -- never just take the first <rc>.
        #   4. that <rc>'s v attribute is the futureMetadata block index,
        #      0-based (do NOT decrement).
        if "xl/metadata.xml" not in self.parsed:
            return

        data = self.getData("xl/metadata.xml")

        metadataTypes = data.findall("main:metadataTypes/main:metadataType", NS)
        valueMetadataBks = data.findall("main:valueMetadata/main:bk", NS)

        for cell, vm in self.cellToVM.items():
            bkIndex = int(vm) - 1  # hop 2: vm is 1-based
            if not (0 <= bkIndex < len(valueMetadataBks)):
                continue

            for rc in valueMetadataBks[bkIndex].findall("main:rc", NS):
                t = rc.get("t")
                v = rc.get("v")
                if t is None or v is None:
                    continue
                typeIndex = int(t) - 1  # hop 3: t is 1-based
                if 0 <= typeIndex < len(metadataTypes) and metadataTypes[typeIndex].get("name") == "XLRICHVALUE":
                    self.cellToV[cell] = v  # hop 4: 0-based, do NOT decrement
                    break

    def _getRID(self):
        # Resolve each cell's futureMetadata block index (cellToV, hop 4) to
        # the final rId string referenced by richValueRel.xml (hops 5-9):
        #   5. futureMetadata[@name='XLRICHVALUE']/bk[v] -> xlrd:rvb[@i]
        #      (v is 0-based, do NOT decrement; i is the 0-based rv index,
        #      do NOT decrement)
        #   6. rdrichvalue.xml: the <rv> at document-order position i; its s
        #      attribute is the 0-based structure index
        #   7. rdrichvaluestructure.xml: the <s> at document-order position
        #      s; find the 0-based position of its
        #      <k n="_rvRel:LocalImageIdentifier"> child
        #   8. back on the <rv> from hop 6: the <v> child at that same
        #      position (v children correspond positionally to the
        #      structure's k children); its text is the
        #      LocalImageIdentifier, 0-based, do NOT decrement
        #   9. richValueRel.xml: the <rel> at document-order LIST POSITION
        #      LocalImageIdentifier (never by parsing/sorting the rId
        #      numeric suffix); read its r:id attribute string
        required = (
            "xl/metadata.xml",
            "xl/richData/rdrichvalue.xml",
            "xl/richData/rdrichvaluestructure.xml",
            "xl/richData/richValueRel.xml",
        )
        if any(path not in self.parsed for path in required):
            return

        metadata = self.getData("xl/metadata.xml")
        futureMetadataBks = metadata.findall(
            "main:futureMetadata[@name='XLRICHVALUE']/main:bk", NS
        )

        rvDataRoot = self.getData("xl/richData/rdrichvalue.xml")
        rvs = rvDataRoot.findall("richdata:rv", NS)

        rvStructuresRoot = self.getData("xl/richData/rdrichvaluestructure.xml")
        structures = rvStructuresRoot.findall("richdata:s", NS)

        richValueRelRoot = self.getData("xl/richData/richValueRel.xml")
        rels = richValueRelRoot.findall("rvr:rel", NS)

        for cell, v in self.cellToV.items():
            vIndex = int(v)  # hop 4 result: 0-based, do NOT decrement
            if not (0 <= vIndex < len(futureMetadataBks)):
                continue

            rvb = futureMetadataBks[vIndex].find(".//richdata:rvb", NS)
            if rvb is None or rvb.get("i") is None:
                continue
            rvIndex = int(rvb.get("i"))  # hop 5: 0-based, do NOT decrement  # pyright: ignore[reportArgumentType]
            if not (0 <= rvIndex < len(rvs)):
                continue
            rv = rvs[rvIndex]  # hop 6

            structIndexAttr = rv.get("s")
            if structIndexAttr is None:
                continue
            structIndex = int(structIndexAttr)
            if not (0 <= structIndex < len(structures)):
                continue
            structure = structures[structIndex]  # hop 7

            keys = structure.findall("richdata:k", NS)
            keyPosition = None
            for i, k in enumerate(keys):
                if k.get("n") == "_rvRel:LocalImageIdentifier":
                    keyPosition = i
                    break
            if keyPosition is None:
                continue

            values = rv.findall("richdata:v", NS)  # hop 8
            if keyPosition >= len(values) or values[keyPosition].text is None:
                continue
            localImageIdentifier = int(values[keyPosition].text)  # 0-based, do NOT decrement  # pyright: ignore[reportArgumentType]

            if not (0 <= localImageIdentifier < len(rels)):
                continue
            rid = rels[localImageIdentifier].get(f"{{{NS['odoc']}}}id")  # hop 9
            if rid:
                self.cellToRID[cell] = rid

    def _getImgPath(self):
        # _get rich value relationships (actual path to images)
        if "xl/richData/_rels/richValueRel.xml.rels" not in self.parsed:
            return

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
