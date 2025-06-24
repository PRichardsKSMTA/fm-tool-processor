"""
fm_tool_core package

Exposes `run_flow` so callers can simply:

    from fm_tool_core import run_flow
"""

from importlib import import_module

# Lazy-import to keep package initialization fast
process_mod = import_module(".process_fm_tool", package=__name__)
run_flow = process_mod.run_flow

__all__ = ["run_flow"]
