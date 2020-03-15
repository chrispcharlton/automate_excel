import os


class TestSheet:
    def test_name(self, open_workbook):
        """Tests getting and setting the .name attribute."""
        sheet = open_workbook.active_sheet
        assert sheet.name == open_workbook.workbook.ActiveSheet.Name
        sheet.name = 'new_name'
        assert sheet.name == 'new_name'
        assert open_workbook.workbook.ActiveSheet.Name == 'new_name'

    def test_to_csv(self, open_workbook):
        """Tests the .to_csv method creates a .csv.

        The .csv file should be created in the same directory as the workbook if a path is not provided.
        """
        open_workbook.active_sheet.to_csv()
        assert 'test.csv' in os.listdir(open_workbook.dir)

    def test_open_in_new_workbook(self, open_workbook):
        """Tests the .open_in_new_workbook method.

        Tests that the .open_in_new_workbook method creates a new workbook with a single sheet of the same name
        as the Sheet object.
        """
        original_workbook = open_workbook.name
        sheet_name = open_workbook.active_sheet.name
        open_workbook.active_sheet.open_in_new_workbook()
        assert open_workbook.name != original_workbook
        assert open_workbook.sheet_names == [sheet_name]
