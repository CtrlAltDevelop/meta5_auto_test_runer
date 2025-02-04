import warnings
from pathlib import Path

from source.main import Meta5AutoTestRunner


if __name__ == '__main__':
    """
    The main entry point of the application.
    """
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    app = Meta5AutoTestRunner(Path.cwd())
    app.safe_run()
