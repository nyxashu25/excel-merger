"""Merge many Excel and CSV files into one."""
from .core import MergeError, MergeResult, SheetInfo, collect_files, merge_files

__version__ = "1.0.0"
__all__ = ["merge_files", "collect_files", "MergeError", "MergeResult", "SheetInfo", "__version__"]
