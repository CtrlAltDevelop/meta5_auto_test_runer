from pathlib import Path

from source.common.main_class import MainClass


class Meta5AutoTestRunner(MainClass):
    def __init__(self, base_path: Path, debug: bool = False):
        super().__init__(base_path)
        self.debug: bool = debug
        self.result_path = base_path / "results"
        self.result_path.mkdir(parents=True, exist_ok=True)

    def __run__(self):
        pass
