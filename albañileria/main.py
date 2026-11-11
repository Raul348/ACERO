"""ALBAÑILERÍA E.070 — arranca la interfaz."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


def main() -> int:
    from PySide6.QtWidgets import QApplication

    from albanileria.ui.main_window import VentanaPrincipal
    from albanileria.ui.theme import aplicar_tema

    app = QApplication(sys.argv)
    app.setApplicationName("ALBAÑILERÍA E.070")
    aplicar_tema(app)
    w = VentanaPrincipal()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
