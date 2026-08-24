import shutil
from pathlib import Path

import pytest

from xlcellimage.workbook_parser import WorkbookParser

DATA = Path(__file__).parent

#######################################################################
# Regression fixture for the vm -> bk -> rc -> futureMetadata -> rv ->
# structure -> LocalImageIdentifier -> rel -> Target resolution chain.
#
# test_workbook_multi_rc.xlsx is a hand-patched copy of test_workbook.xlsx
# (see scratchpad build_fixture.py used to generate it) where xl/metadata.xml
# was modified so the numeric coincidences present in the original fixture
# no longer hold:
#
#   * metadataTypes now has TWO entries: a decoy "XLDAPR" at index 0 and the
#     real "XLRICHVALUE" at index 1. valueMetadata/bk[1] (used by cell C4,
#     vm="2") now carries a decoy <rc t="1" v="0"/> for the fake XLDAPR type
#     *before* the real <rc t="2" v="1"/> for XLRICHVALUE. A parser that
#     just takes the first <rc> in a <bk> (instead of checking which one's
#     indexed metadataType is actually named XLRICHVALUE) will resolve C4
#     to the decoy's v="0" instead of the real v="1".
#
#   * futureMetadata/bk[0] (used by cell I1, vm="1") now points its
#     <xlrd:rvb i="2"/> at rv index 2 instead of rv index 0, breaking the
#     block-index == rv-index == LocalImageIdentifier identity that made
#     the original fixture numerically coincidental. A parser that skips
#     the rdrichvalue.xml / rdrichvaluestructure.xml hops and treats the
#     futureMetadata block index as if it were already the final
#     LocalImageIdentifier will resolve I1 to the wrong image.
#
# Expected (correct, hand-traced against the raw XML independent of the
# implementation) resolution under this fixture:
#   I1  -> image3.png   (was image1.png under the old identity mapping)
#   C4  -> image2.png   (unchanged; the decoy must be skipped)
#   I23 -> image3.png   (unchanged)
#######################################################################


@pytest.fixture
def multi_rc_workbook_path(tmp_path) -> str:
    dest = tmp_path / "test_workbook_multi_rc.xlsx"
    shutil.copy(DATA / "test_workbook_multi_rc.xlsx", dest)
    return dest


@pytest.fixture
def wbp_multi_rc(multi_rc_workbook_path):
    return WorkbookParser(multi_rc_workbook_path)


class TestWorkbookParserMultiRC:

    def test_cellToV_skips_decoy_rc(self, wbp_multi_rc: WorkbookParser):
        # C4's bk has a decoy XLDAPR-typed <rc v="0"/> before the real
        # XLRICHVALUE-typed <rc v="1"/>. Must resolve to the real one.
        assert wbp_multi_rc.cellToV[("sheet1", "C4")] == "1"

    def test_cellToV_all_cells(self, wbp_multi_rc: WorkbookParser):
        assert wbp_multi_rc.cellToV[("sheet1", "I1")] == "0"
        assert wbp_multi_rc.cellToV[("sheet1", "C4")] == "1"
        assert wbp_multi_rc.cellToV[("sheet1", "I23")] == "2"

    def test_cellToPath_i1_goes_through_rv_hops(self, wbp_multi_rc: WorkbookParser):
        # I1's futureMetadata block (index 0) points (via xlrd:rvb i="2")
        # at rv index 2, not rv index 0. A parser that skips hops 5-8 and
        # uses the block index directly as the final rel-list position
        # would wrongly resolve I1 to image1.png.
        assert wbp_multi_rc.cellToPath[("sheet1", "I1")] == "image3.png"

    def test_cellToPath_c4_skips_decoy(self, wbp_multi_rc: WorkbookParser):
        assert wbp_multi_rc.cellToPath[("sheet1", "C4")] == "image2.png"

    def test_cellToPath_i23_unaffected(self, wbp_multi_rc: WorkbookParser):
        assert wbp_multi_rc.cellToPath[("sheet1", "I23")] == "image3.png"

    def test_getImage_i1_returns_bytes(self, wbp_multi_rc: WorkbookParser):
        assert wbp_multi_rc.hasImage(("sheet1", "I1"))
        assert wbp_multi_rc.getImage(("sheet1", "I1")) is not None
