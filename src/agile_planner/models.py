"""Domain models for the personal Agile planner."""

from dataclasses import asdict, dataclass
from enum import Enum


class TaskStatus(str, Enum):
    """Supported task workflow states."""

    BACKLOG = "backlog"
    TODO = "todo"
    IN_PROGRESS = "in_progress"
    DONE = "done"


@dataclass
class Task:
    """A planner task stored in the local JSON file."""

    title: str
    description: str = ""
    priority: str = "medium"
    status: TaskStatus = TaskStatus.BACKLOG
    due_date: str = ""

    def to_dict(self) -> dict[str, str]:
        data = asdict(self)
        data["status"] = self.status.value
        return data

    @classmethod
    def from_dict(cls, data: dict[str, str]) -> "Task":
        return cls(
            title=data["title"],
            description=data.get("description", ""),
            priority=data.get("priority", "medium"),
            status=TaskStatus(data.get("status", TaskStatus.BACKLOG.value)),
            due_date=data.get("due_date", ""),
        )

