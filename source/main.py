import logging
from copy import copy
from datetime import date
from pathlib import Path
from configparser import ConfigParser
from tkinter import filedialog
from typing import Iterable, Tuple, Optional, Dict, Any

import pandas as pd
from tqdm import tqdm

from source.common.main_class import MainClass


class CaseSensitiveConfigParser(ConfigParser):
    def optionxform(self, option):
        return option


class Meta5AutoTestRunner(MainClass):
    def __init__(self, base_path: Path, debug: bool = True):
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

        for _pass, values in tqdm(self._read_optimize_result_file(_path).items(), desc='Run Strategy Tester'):
            filename = f'Res{_pass}_{date.today().strftime("%Y-%m-%d")}'
            _config = self._update_config(self._config, filename)
            _config_path = self.config_path / f'{filename}.ini'
            with _config_path.open(mode="w", encoding="utf-8") as f:
                _config.write(f)

    @staticmethod
    def _update_config(_config: ConfigParser, filename: str) -> ConfigParser:
        config = CaseSensitiveConfigParser()
        config.add_section("Common")
        config.add_section("Tester")

        for key, value in _config["Account"].items():
            config.set("Common", key, value)

        for key, value in _config["Tester"].items():
            config.set("Tester", key, value)

        for key, value in {
            "ReplaceReport": "1",
            "ShutdownTerminal": "1",
            "Report": f"reports/{filename}"
        }.items():
            config.set("Tester", key, value)

        config.remove_section("Meta")
        config.remove_section("Account")
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
        df[df.select_dtypes(include=[bool]).columns] = df.select_dtypes(include=[bool]).astype(int)
        result = df.filter(items=['Pass'] + [col for col in df.columns if col.startswith('_')])
        result.columns = result.columns.str.lstrip('_')
        result['Pass'] = result['Pass'].astype(int)
        return result.set_index('Pass').to_dict(orient='index')