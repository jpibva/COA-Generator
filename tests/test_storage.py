import json
import tempfile
import unittest
from pathlib import Path

from coa_storage import (
    BACKUP_DIRNAME,
    load_config,
    load_micro_history,
    load_session_data,
    save_config,
    save_micro_history_record,
    save_session_data,
)

try:
    import openpyxl  # noqa: F401
    OPENPYXL_AVAILABLE = True
except ModuleNotFoundError:
    OPENPYXL_AVAILABLE = False


class SessionStorageTests(unittest.TestCase):
    def test_session_roundtrip_and_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            session_one = {"pdf_path": "uno.pdf", "micro_results": {"P": {"L1": {"TPC": "10"}}}}
            session_two = {"pdf_path": "dos.pdf", "micro_results": {"P": {"L2": {"TPC": "20"}}}}

            save_session_data(session_one, tmp)
            self.assertEqual(load_session_data(tmp), session_one)

            save_session_data(session_two, tmp)
            self.assertEqual(load_session_data(tmp), session_two)

            backup_dir = Path(tmp) / BACKUP_DIRNAME
            backups = list(backup_dir.glob("session_*.json.bak"))
            self.assertGreaterEqual(len(backups), 1)


class ConfigStorageTests(unittest.TestCase):
    def test_save_config_creates_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_path = Path(tmp) / "config.json"
            config_one = load_config(config_file=str(config_path))
            save_config(config_one, config_file=str(config_path))

            config_two = dict(config_one)
            config_two["pdf_folder"] = "D:/nuevo"
            save_config(config_two, config_file=str(config_path))

            self.assertEqual(json.loads(config_path.read_text(encoding="utf-8"))["pdf_folder"], "D:/nuevo")
            backups = list((Path(tmp) / BACKUP_DIRNAME).glob("config_*.json.bak"))
            self.assertGreaterEqual(len(backups), 1)


@unittest.skipUnless(OPENPYXL_AVAILABLE, "openpyxl no está instalado")
class MicroHistoryStorageTests(unittest.TestCase):
    def test_history_update_keeps_single_row_and_creates_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            history_path = Path(tmp) / "Historial_Microbiologia.xlsx"
            config = load_config()

            save_micro_history_record(
                "Cliente Demo",
                "Blueberry 400g",
                "LOT-01",
                {"TPC": "100", "Yeast": "5"},
                config,
                formato_nombre="Estándar",
                history_file=str(history_path),
            )
            save_micro_history_record(
                "Cliente Demo",
                "Blueberry 400g",
                "LOT-01",
                {"TPC": "200", "Yeast": "7"},
                config,
                formato_nombre="Estándar",
                history_file=str(history_path),
            )

            history = load_micro_history(str(history_path))
            self.assertEqual(history[("cliente demo", "blueberry", "LOT-01")]["TPC"], "200")
            self.assertEqual(history[("cliente demo", "blueberry", "LOT-01")]["Yeast"], "7")

            import openpyxl

            wb = openpyxl.load_workbook(history_path)
            ws = wb["Estándar"]
            self.assertEqual(ws.max_row, 2)

            backups = list((Path(tmp) / BACKUP_DIRNAME).glob("Historial_Microbiologia_*.xlsx.bak"))
            self.assertGreaterEqual(len(backups), 1)


if __name__ == "__main__":
    unittest.main()
