#!/usr/bin/env python3
"""Régénère le PDF imprimable du Module 2 (délègue au générateur commun)."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from generer_pdfs import main

if __name__ == "__main__":
    raise SystemExit(main(["2"]))
