"""JSON storage helpers for planner tasks."""

from json import dump, load
from pathlib import Path

from .models import Task


DEFAULT_STORAGE_PATH = Path("data/tasks.json")


def load_tasks(path: Path = DEFAULT_STORAGE_PATH) -> list[Task]:
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8") as file:
        raw_tasks = load(file)

    return [Task.from_dict(item) for item in raw_tasks]


def save_tasks(tasks: list[Task], path: Path = DEFAULT_STORAGE_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as file:
        dump([task.to_dict() for task in tasks], file, ensure_ascii=False, indent=2)

