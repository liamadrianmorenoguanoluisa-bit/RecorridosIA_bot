#!/usr/bin/env python3
"""
RecorridosIA — Punto de entrada para Render
Render ejecuta: python main.py
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from bot.main import main

if __name__ == "__main__":
    main()
