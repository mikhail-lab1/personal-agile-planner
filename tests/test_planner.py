import unittest
from pathlib import Path

from src.agile_planner.planner import Planner


class PlannerTest(unittest.TestCase):
    def test_add_task(self) -> None:
        storage_path = Path("data/test_tasks.json")
        storage_path.unlink(missing_ok=True)

        try:
            planner = Planner(storage_path)

            task = planner.add_task("Prepare report", priority="high")
            tasks = planner.list_tasks()

            self.assertEqual(task.title, "Prepare report")
            self.assertEqual(task.priority, "high")
            self.assertEqual(len(tasks), 1)
            self.assertEqual(tasks[0].title, "Prepare report")
        finally:
            storage_path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
