"""Set bundled Tcl/Tk resource paths before the frozen GUI imports tkinter."""

import os
import sys
from pathlib import Path


if getattr(sys, "frozen", False):
    _bundle_root = Path(sys._MEIPASS)
    os.environ["TCL_LIBRARY"] = str(_bundle_root / "tcl" / "tcl8.6")
    os.environ["TK_LIBRARY"] = str(_bundle_root / "tcl" / "tk8.6")
