"""Application logic for managing personal Agile tasks."""

from pathlib import Path

from .models import Task
from .storage import DEFAULT_STORAGE_PATH, load_tasks, save_tasks


class Planner:
    """A small facade around task storage and task operations."""

    def __init__(self, storage_path: Path = DEFAULT_STORAGE_PATH) -> None:
        self.storage_path = storage_path

    def list_tasks(self) -> list[Task]:
        return load_tasks(self.storage_path)

    def add_task(
        self,
        title: str,
        description: str = "",
        priority: str = "medium",
        due_date: str = "",
    ) -> Task:
        tasks = self.list_tasks()
        task = Task(
            title=title,
            description=description,
            priority=priority,
            due_date=due_date,
        )
        tasks.append(task)
        save_tasks(tasks, self.storage_path)
        return task

