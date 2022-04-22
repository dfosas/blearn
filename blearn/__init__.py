# coding=utf-8
"""blearn"""

major = 0  # Incompatible API changes
minor = 1  # Add functionality in a backwards-compatible manner
micro = 0  # Backwards-compatible bug fixes
local = None  # Local changes
__version__ = f"{major}.{minor}.{micro}"
if local is not None:
    __version__ += f"+{local}"
