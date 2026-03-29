from __future__ import annotations

from .db import VocabRepository, default_db_path
from .ui import MainWindow


def main() -> None:
    repository = VocabRepository(default_db_path())
    repository.initialize()

    app = MainWindow(repository)
    app.mainloop()


if __name__ == "__main__":
    main()
