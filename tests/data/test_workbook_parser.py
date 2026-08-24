import pytest

from xlcellimage.workbook_parser import WorkbookParser


class TestWorkbookParser:

    def test_zip_init(self, wbp: WorkbookParser):
        assert wbp.zip is not None

    def test_sheets_init(self, wbp: WorkbookParser):
        assert wbp.sheets == ["xl/worksheets/sheet1.xml"]

    def test_parsing_init(self, wbp: WorkbookParser):
        assert wbp.parsed["xl/worksheets/sheet1.xml"] is not None
        assert wbp.parsed["xl/metadata.xml"] is not None
        assert wbp.parsed["xl/richData/richValueRel.xml"] is not None
        assert wbp.parsed["xl/richData/_rels/richValueRel.xml.rels"] is not None

    def test_image_init(self, wbp: WorkbookParser):
        assert wbp.images == {}

    def test_cellToVM(self, wbp: WorkbookParser):
        assert wbp.cellToVM[("sheet1", "C4")] == "2"
        assert wbp.cellToVM[("sheet1", "I1")] == "1"
        assert wbp.cellToVM[("sheet1", "I23")] == "3"

    def test_cellToV(self, wbp: WorkbookParser):
        assert wbp.cellToV[("sheet1", "C4")] == "1"

    def test_cellToRID(self, wbp: WorkbookParser):
        assert wbp.cellToRID[("sheet1", "C4")] == 'rId2'

    def test_cellToPath(self, wbp: WorkbookParser):
        assert wbp.cellToPath[("sheet1", "C4")] == 'image2.png'

    def test_getData1(self, wbp: WorkbookParser):
        assert wbp.getData("xl/metadata.xml") is not None

    def test_getData2(self, wbp: WorkbookParser):
        with pytest.raises(KeyError):
            _ = wbp.getData("lol")

    def test_hasImage(self, wbp: WorkbookParser):
        assert wbp.hasImage(("sheet1", "C4"))

    def test_hasImageFalse(self, wbp: WorkbookParser):
        assert not wbp.hasImage(("sheet1", "A1"))

    def test_getImage(self, wbp: WorkbookParser):
        assert wbp.getImage(("sheet1", "C4")) is not None

    def test_getImageFail(self, wbp: WorkbookParser):
        with pytest.raises(KeyError):
            _ = wbp.getImage(("sheet1", "A1"))
        
        
