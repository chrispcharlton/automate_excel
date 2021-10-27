import pytest
import os
import shutil
from src.workbook import Workbook

WB_NUMBER = 0

@pytest.fixture(scope='class')
def testdir(tmp_path_factory):
    testdir = tmp_path_factory.mktemp('tests')
    yield testdir
    shutil.rmtree(testdir)

@pytest.fixture(scope='class')
def open_workbook(testdir):
    with Workbook(testdir.joinpath('test')) as wb:
        yield wb

@pytest.fixture(scope='function')
def unique_workbook(testdir):
    global WB_NUMBER
    with Workbook(testdir.joinpath(f"test{WB_NUMBER}")) as wb:
        yield wb
    WB_NUMBER += 1