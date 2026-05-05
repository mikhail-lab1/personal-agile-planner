"""Command line interface for Personal Agile Planner."""

from argparse import ArgumentParser

from .planner import Planner


def build_parser() -> ArgumentParser:
    parser = ArgumentParser(description="Personal Agile Planner")
    subparsers = parser.add_subparsers(dest="command")

    add_parser = subparsers.add_parser("add", help="Add a new task")
    add_parser.add_argument("title", help="Task title")
    add_parser.add_argument("--description", default="", help="Task description")
    add_parser.add_argument("--priority", default="medium", help="Task priority")
    add_parser.add_argument("--due-date", default="", help="Task due date")

    subparsers.add_parser("list", help="Show all tasks")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    planner = Planner()

    if args.command == "add":
        task = planner.add_task(
            title=args.title,
            description=args.description,
            priority=args.priority,
            due_date=args.due_date,
        )
        print(f"Task added: {task.title}")
    elif args.command == "list":
        tasks = planner.list_tasks()
        if not tasks:
            print("No tasks found.")
            return

        for index, task in enumerate(tasks, start=1):
            print(f"{index}. [{task.status.value}] {task.title} ({task.priority})")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

