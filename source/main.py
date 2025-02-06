import logging
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
        result.columns = result.columns.str.lstrip('_')
        result['Pass'] = result['Pass'].astype(int)
        return result.set_index('Pass').to_dict(orient='index')

    @staticmethod
    def _html_to_excel(html_path: Path, output_path: Path):
        """
        Opens an HTML file in Excel (using COM) and saves it as XLSX,
        then moves all images to a separate sheet named 'Images'.
        """
        logging.info(f"Converting HTML to Excel: {html_path} -> {output_path}")

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(str(html_path))

        # Create a new sheet for images
        image_sheet = wb.Sheets.Add()
        image_sheet.Name = "Images"

        for sheet in wb.Sheets:
            if sheet.Name == "Images":
                continue  # Skip the newly created image sheet

            image_row = 1  # Track row position for placing images in the new sheet
            for shape in sheet.Shapes:
                # Move the shape to the 'Images' sheet
                shape.Copy()
                image_sheet.Paste()
                pasted_shape = image_sheet.Shapes(image_sheet.Shapes.Count)

                # Adjust position in the new sheet
                pasted_shape.Top = image_row * 50  # Adjust spacing between images
                pasted_shape.Left = 10  # Keep images aligned on the left
                image_row += 5  # Move down for next image
                shape.Delete()

        wb.SaveAs(str(output_path), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.Close(False)
        excel.Quit()
        print(f"Excel file saved: {output_path}")

