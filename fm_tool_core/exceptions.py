class FlowError(Exception):
    """Domain-specific exception returned to Power Automate Cloud Flow."""

    def __init__(self, msg: str, *, work_completed: bool = False) -> None:
        super().__init__(msg)
        self.work_completed = work_completed

    def __str__(self) -> str:  # pragma: no cover - trivial
        return self.args[0]


__all__ = ["FlowError"]
