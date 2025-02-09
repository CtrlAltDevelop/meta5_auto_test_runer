import logging
import math
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from configparser import ConfigParser
from tkinter import filedialog
from typing import Iterable, Tuple, Optional, Dict, Any

import pandas as pd
from tqdm import tqdm
import win32com.client as win32

from source.common.main_class import MainClass


class CaseSensitiveConfigParser(ConfigParser):
    def optionxform(self, option):
        return option


class Meta5AutoTestRunner(MainClass):
    def __init__(self, base_path: Path, debug: bool = False):
        super().__init__(base_path)
        self.debug: bool = debug
        self.config_path = base_path / "configs"
        self.data_path = base_path / "test_data"
        self.result_path = base_path / "results"

        self.config_path.mkdir(parents=True, exist_ok=True)
        self.result_path.mkdir(parents=True, exist_ok=True)

        self._config = CaseSensitiveConfigParser()
        self._config.read(base_path / "settings.ini", encoding="utf-8")

    def _remove_cache(self):
        folder_path = Path(self._config['Meta']['DataFolderPath']) / 'Tester' / 'cache'
        if not folder_path.exists():
            logging.warning(f"The path {folder_path} does not exist.")
            return

        if not folder_path.is_dir():
            logging.warning(f"The path {folder_path} is not a directory.")
            return

        logging.info(f"Clearing folder: {folder_path}")

        for item in folder_path.iterdir():
            try:
                if item.is_file():
                    logging.info(f"Removing file: {item}")
                    item.unlink()
                elif item.is_dir():
                    logging.info(f"Removing directory: {item}")
                    shutil.rmtree(item)
            except Exception as e:
                logging.error(f"Failed to remove {item}: {e}")

        logging.info("Folder cleared successfully.")

    def __run__(self):
        if self.debug:
            _path = self.data_path / 'Cleaned_TF15-401-ReportOptimizer-USDCHF-93008552.csv'
        else:
            print("Select Report Optimizer file")
            _path = self._get_file_via_dialog(
                title=f"Report Optimizer file",
                filetypes=[("Report Optimizer", "*.csv")]
            )

        data_path = Path(self._config['Meta']['DataFolderPath']) / 'reports'
        data_path.mkdir(parents=True, exist_ok=True)
        for _pass, values in tqdm(self._read_optimize_result_file(_path).items(), desc='Run Strategy Tester'):
            self._remove_cache()
            filename = f'Res{_pass}_{int(datetime.now().timestamp())}'
            _config = self._update_config(self._config, filename, values)
            _config_path = self.config_path / f'{filename}.ini'
            with _config_path.open(mode="w", encoding="utf-8") as f:
                _config.write(f)
            subprocess.run([self._config['Meta']['TerminalPath'], f"/config:{_config_path}"])
            self._html_to_excel(data_path / f'{filename}.htm', self.result_path / f'{filename}.xlsx')

    @staticmethod
    def _update_config(_config: ConfigParser, filename: str, inputs: Dict[str, Any]) -> ConfigParser:
        config = CaseSensitiveConfigParser()
        config.add_section("Common")
        config.add_section("Tester")
        config.add_section("TesterInputs")

        for key, value in _config["Account"].items():
            config.set("Common", key, value)

        for key, value in _config["Tester"].items():
            config.set("Tester", key, value)

        for key, value in {
            # Optimization mode:
            # 0 = No optimization (single test)
            # 1 = Slow, complete optimization
            # 2 = Fast genetic-based optimization
            # 3 = All symbols selected in Market Watch
            # 4 = All symbols in the tester's symbol list
            'Optimization': '0',

            # The backtest model (how ticks are simulated):
            #  0 = Every tick
            #  1 = 1 minute OHLC
            #  2 = Open prices only
            'Model': '2',

            # Whether to enable/disable the use of custom dates:
            #  0 = Use the full available data
            #  1 = Use the FromDate/ToDate
            'Dates': '1',

            # Forward testing mode (split the test period):
            #  0 = No forward testing
            #  1 = Forward testing on 1/2 of the period
            #  2 = Forward testing on 1/3 of the period
            #  3 = Forward testing on 1/4 of the period
            #  4 = Custom
            'ForwardMode': '0',

            # Deposit currency (USD, EUR, etc.)
            'Currency': 'USD',

            # If 1, profits are shown in pips instead of currency. 0 means disabled.
            'ProfitInPips': '0',

            # Account leverage for the test (1:100, etc.)
            'Leverage': '100',

            # Execution mode:
            #  0 = Execution without delay
            #  1 = Execution with random delay
            'ExecutionMode': '0',

            # Optimization criterion:
            #  0 = Maximize balance
            #  1 = Maximize profit factor
            #  2 = Maximize expected payoff
            #  3 = Minimize drawdown
            'OptimizationCriterion': '0',

            # Whether to run a visual backtest (0 = no, 1 = yes)
            'Visual': '0',

            # Replace htm file if exist (0 = no, 1 = yes)
            # 0 = create new file
            # 1 = replace file if exist
            'ReplaceReport': '1',

            # close MetaTrader after test done (0 = no, 1 = yes)
            'ShutdownTerminal': '1',

            # path and filename for htm report file
            'Report': f'reports\\{filename}',
        }.items():
            config.set("Tester", key, value)

        for key, value in _config["TesterInputs"].items():
            config.set("TesterInputs", key, value)

        for key, value in inputs.items():
            config.set("TesterInputs", key, str(value))

        return config

    def _get_file_via_dialog(self, title: str, filetypes: Iterable[Tuple[str, str]], optional: bool = False) \
            -> Optional[Path]:
        """Open a file dialog to select a file.

        :param title: Title of the dialog window.
        :param filetypes: List of file types for filtering.
        :return: Path to the selected file.
        :raises FileNotFoundError: If no file is selected.
        """
        logging.debug(f"Opening file dialog: {title}")
        with self._tkinter_root():
            file = filedialog.askopenfile(title=title, filetypes=filetypes, initialdir=self.base_path)
            if not file and not optional:
                logging.error(f"{title} not selected.")
                raise FileNotFoundError(f"{title} not selected.")
            if file:
                selected_path = Path(file.name)
                logging.debug(f"File selected: {selected_path}")
                return selected_path
            return None

    @staticmethod
    def _read_optimize_result_file(path: Path) -> Dict[int, Any]:
        df = pd.read_csv(path)
        df[df.select_dtypes(include=[bool]).columns] = \
            df.select_dtypes(include=[bool]).apply(lambda x: x.astype(str).str.lower())
        result = df.filter(items=['Pass'] + [col for col in df.columns if col.startswith('_')])
        result['Pass'] = result['Pass'].astype(int)
        return result.set_index('Pass').to_dict(orient='index')

    @staticmethod
    def _html_to_excel(html_path: Path, output_path: Path):
        """
        Opens an HTML file in Excel (using COM) and saves it as XLSX.

        The conversion performs the following steps:
          - Renames the first sheet to "Deals and Orders".
          - In "Deals and Orders":
                * Removes row 3 if it is empty.
                * Inserts an empty column at column 12.
          - Adds a new sheet "Backtest", then:
                * Moves the "Deals and Orders" sheet to the first position.
                * Finds a block of rows in "Deals and Orders" (between the rows containing
                  "Results" and "Orders") and copies it to "Backtest".
                  aligns columns (odd left–aligned, even right–aligned), and auto–fits columns.
          - Copies any shapes (images) from all sheets (except "Backtest") into "Backtest"
            below the report data with an extra three blank rows between images.
          - If the output file already exists, it is replaced.
          - If the HTML file does not exist, an error message is printed and the function exits.
        """
        # Check if the HTML file exists
        if not html_path.exists():
            print("Meta Test Done with ERROR (HTM file does not exist). Please check your input file and try again.")
            return

        logging.info(f"Converting HTML to Excel: {html_path} -> {output_path}")

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(str(html_path))

        # Rename the first sheet to "Deals and Orders"
        deals_sheet = wb.Sheets(1)
        deals_sheet.Name = "Deals and Orders"

        # --- Modification: Adjust Deals and Orders Sheet ---
        # Remove row 3 if it is empty.
        if excel.WorksheetFunction.CountA(deals_sheet.Rows(3)) == 0:
            deals_sheet.Rows(3).Delete()

        # Insert an empty column at column 13 (existing columns shift to the right).
        deals_sheet.Columns(13).Insert()

        # Add a new sheet for the report/backtest data.
        report_sheet = wb.Sheets.Add()
        report_sheet.Name = "Backtest"

        # --- Modification: Move "Deals and Orders" to be the first sheet ---
        deals_sheet.Move(Before=wb.Sheets(1))

        used_range = deals_sheet.UsedRange
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        # Find the start and end rows of the report block based on "Results" and "Orders" markers.
        start_row = None
        end_row = None
        for i in range(1, row_count + 1):
            if start_row is None:
                for j in range(1, col_count + 1):
                    if deals_sheet.Cells(i, j).Value == "Results":
                        start_row = i + 1
                        break
            elif end_row is None:
                for j in range(1, col_count + 1):
                    if deals_sheet.Cells(i, j).Value == "Orders":
                        end_row = i - 1
                        break
            if start_row is not None and end_row is not None:
                break

        if start_row is not None and end_row is not None:
            source_range = deals_sheet.Range(deals_sheet.Cells(start_row, 1),
                                             deals_sheet.Cells(end_row, col_count))
            source_range.Copy(report_sheet.Range("A1"))
        else:
            logging.warning("Could not find both 'Results' and 'Orders' markers in the Deals and Orders sheet.")

        # --- Modification: Data Cleanup in the Backtest Sheet ---

        # Clear any cells whose value is NaN.
        report_used = report_sheet.UsedRange
        for i in range(1, report_used.Rows.Count + 1):
            for j in range(1, report_used.Columns.Count + 1):
                cell = report_sheet.Cells(i, j)
                cell_val = cell.Value
                if isinstance(cell_val, float) and math.isnan(cell_val):
                    cell.ClearContents()

        # Remove any empty columns.
        report_used = report_sheet.UsedRange
        ncols = report_used.Columns.Count
        for j in range(ncols, 0, -1):
            if excel.WorksheetFunction.CountA(report_sheet.Columns(j)) == 0:
                report_sheet.Columns(j).Delete()

        # Remove any empty rows in the Backtest sheet.
        report_used = report_sheet.UsedRange
        nrows = report_used.Rows.Count
        for i in range(nrows, 0, -1):
            if excel.WorksheetFunction.CountA(report_sheet.Rows(i)) == 0:
                report_sheet.Rows(i).Delete()

        # Recalculate the used range after row and column deletion.
        report_used = report_sheet.UsedRange
        report_row_count = report_used.Rows.Count
        report_col_count = report_used.Columns.Count

        # Apply alternating row colors.
        grey_color = 14474460  # roughly 0xDCDCDC (light grey)
        white_color = 16777215  # white 0xFFFFFF

        for i in range(1, report_row_count + 1):
            row_range = report_sheet.Range(report_sheet.Cells(i, 1), report_sheet.Cells(i, report_col_count))
            if i % 2 == 1:
                row_range.Interior.Color = grey_color
            else:
                row_range.Interior.Color = white_color

        # Format columns: odd columns left-aligned, even columns right-aligned, and auto-fit.
        for j in range(1, report_col_count + 1):
            col_range = report_sheet.Range(report_sheet.Cells(1, j), report_sheet.Cells(report_row_count, j))
            if excel.WorksheetFunction.CountA(col_range) > 0:
                if j % 2 == 1:
                    col_range.HorizontalAlignment = -4131  # xlHAlignLeft
                else:
                    col_range.HorizontalAlignment = -4152  # xlHAlignRight
                col_range.EntireColumn.AutoFit()

        report_used = report_sheet.UsedRange
        last_used_row = report_used.Row + report_used.Rows.Count - 1
        dest_row = last_used_row + 2  # Leave one blank row

        # Copy shapes (images) from all sheets (except "Backtest") into "Backtest".
        for sheet in wb.Sheets:
            if sheet.Name == "Backtest":
                continue
            shapes = [shape for shape in sheet.Shapes]
            for shape in shapes:
                shape.CopyPicture(Appearance=1, Format=-4147)  # xlScreen=1, xlPicture=-4147
                report_sheet.Paste()
                pasted_shape = report_sheet.Shapes(report_sheet.Shapes.Count)
                pasted_shape.Top = report_sheet.Cells(dest_row, 1).Top
                pasted_shape.Left = report_sheet.Cells(dest_row, 1).Left + 10
                row_height = report_sheet.Rows(dest_row).RowHeight
                rows_occupied = pasted_shape.Height / row_height
                dest_row += math.ceil(rows_occupied) + 3
                shape.Delete()

        # If the output file exists, remove it.
        if output_path.exists():
            os.remove(output_path)

        wb.SaveAs(str(output_path), FileFormat=51)  # FileFormat 51 corresponds to .xlsx
        wb.Close(False)
        excel.Quit()
        print(f"Excel file saved: {output_path}")
