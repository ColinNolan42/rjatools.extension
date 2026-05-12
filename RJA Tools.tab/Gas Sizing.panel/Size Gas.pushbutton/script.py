# -*- coding: utf-8 -*-
"""
Size Gas.pushbutton/script.py
Phase 2 — IFGC Longest Run Gas Pipe Sizing

LOCKED: This button is not active until the user states
"Phase 1 is complete. Move to Phase 2."
"""
from pyrevit import forms

forms.alert(
    'Size Gas is not yet active.\n\n'
    'This tool will be enabled in Phase 2.\n'
    'Complete and validate all Phase 1 sub-tasks first.',
    title='Phase 2 — Locked'
)
