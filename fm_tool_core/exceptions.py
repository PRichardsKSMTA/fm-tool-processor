# Custom exception type used to signal errors back to the Flow
class FlowError(Exception):
    """Domain-specific exception returned to Power Automate Cloud Flow."""

    def __init__(self, msg: str, *, work_completed: bool = False) -> None:
        # Pass the message to the base Exception class
        super().__init__(msg)
        # Indicates if partial work was completed before the error
        self.work_completed = work_completed

    def __str__(self) -> str:  # pragma: no cover - trivial
        # Return just the message text
        return self.args[0]


# Exported symbols when importing * from this module
__all__ = ["FlowError"]
