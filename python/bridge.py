from __future__ import annotations

import json
import sys
import traceback
from typing import Any

from errors import BridgeError
from models.ooxml import BridgeRequest
from service import PowerPointService


def _write_response(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=True, default=str) + "\n")
    sys.stdout.flush()


def _write_log(message: str) -> None:
    sys.stderr.write(message + "\n")
    sys.stderr.flush()


def main() -> None:
    service = PowerPointService()
    _write_log(f"PowerPoint bridge started with engine={service.engine.name}")

    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue

        request_id = ""
        try:
            payload = json.loads(line)
            request = BridgeRequest.model_validate(payload)
            request_id = request.id

            if request.method == "__shutdown__":
                service.shutdown()
                _write_response({"id": request.id, "result": {"ok": True}})
                return

            result = service.dispatch(request.method, request.params)
            _write_response({"id": request.id, "result": result})
        except BridgeError as exc:
            _write_response(
                {
                    "id": request_id or "unknown",
                    "error": exc.to_payload(),
                }
            )
        except Exception as exc:  # pragma: no cover
            trace = traceback.format_exc()
            _write_log(trace)
            _write_response(
                {
                    "id": request_id or "unknown",
                    "error": {
                        "code": "internal_error",
                        "message": str(exc),
                        "details": {"traceback": trace},
                    },
                }
            )


if __name__ == "__main__":
    main()
