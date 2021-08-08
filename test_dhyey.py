import pytest
from dhyeymain import sheets_access

def test_sheet_access():
    assert sheets_access()[1] == ['Applied_SDLC','Adv_Python','MBSE']

