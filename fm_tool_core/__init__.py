"""
fm_tool_core package

Exposes `run_flow` so callers can simply:

    from fm_tool_core import run_flow
"""

# Import the helper that lets us load modules by name
from importlib import import_module

# Load the processing module only when this package is imported so startup is
# quick
process_mod = import_module(".process_fm_tool", package=__name__)

# Expose the run_flow function at the package level
run_flow = process_mod.run_flow

# Declare the public symbols for `from fm_tool_core import *`
__all__ = ["run_flow"]
