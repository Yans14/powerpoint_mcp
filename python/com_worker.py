from __future__ import annotations

import queue
import threading
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from errors import BridgeError

try:
    import pythoncom
except Exception:  # pragma: no cover
    pythoncom = None  # type: ignore[assignment]


@dataclass
class _Task:
    func: Callable[..., Any]
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    done_event: threading.Event
    result: Any = None
    error: Exception | None = None


class COMWorker:
    """Single STA thread owning all COM calls and COM objects."""

    def __init__(self) -> None:
        if pythoncom is None:
            raise BridgeError(
                code="dependency_missing",
                message="pythoncom is not available. Install pywin32 on Windows.",
            )

        self._queue: queue.Queue[_Task | None] = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True, name="pptx-com-sta")
        self._thread.start()

    def _run(self) -> None:
        assert pythoncom is not None
        pythoncom.CoInitialize()
        try:
            while True:
                task = self._queue.get()
                if task is None:
                    return
                try:
                    task.result = task.func(*task.args, **task.kwargs)
                except Exception as exc:  # pragma: no cover
                    task.error = exc
                finally:
                    task.done_event.set()
        finally:
            pythoncom.CoUninitialize()

    def call(self, func: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        done = threading.Event()
        task = _Task(func=func, args=args, kwargs=kwargs, done_event=done)
        self._queue.put(task)
        done.wait()
        if task.error:
            raise task.error
        return task.result

    def shutdown(self) -> None:
        self._queue.put(None)
        self._thread.join(timeout=2.0)
