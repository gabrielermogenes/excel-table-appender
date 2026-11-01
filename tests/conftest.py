"""Carrega config.example.py como se fosse o config.py, para os testes rodarem em qualquer máquina."""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

spec = importlib.util.spec_from_file_location("config", ROOT / "config.example.py")
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)
sys.modules["config"] = config
