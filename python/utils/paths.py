from __future__ import annotations

import os
from pathlib import Path

from errors import BridgeError


def _assert_absolute(path: str) -> Path:
    candidate = Path(path)
    if not candidate.is_absolute():
        raise BridgeError(
            code="validation_error",
            message="Path must be absolute.",
            details={"path": path},
        )
    return candidate


def validate_existing_file(path: str, expected_suffixes: tuple[str, ...] | None = None) -> Path:
    candidate = _assert_absolute(path)

    if not candidate.exists() or not candidate.is_file():
        raise BridgeError(
            code="not_found",
            message="Input file does not exist.",
            details={"path": path},
        )

    if expected_suffixes and candidate.suffix.lower() not in expected_suffixes:
        raise BridgeError(
            code="validation_error",
            message="Unexpected file extension.",
            details={"path": path, "expected_suffixes": list(expected_suffixes)},
        )

    if not os.access(candidate, os.R_OK):
        raise BridgeError(
            code="validation_error",
            message="Input file is not readable.",
            details={"path": path},
        )

    return candidate


def validate_output_file(path: str, expected_suffixes: tuple[str, ...] | None = None) -> Path:
    candidate = _assert_absolute(path)

    if expected_suffixes and candidate.suffix.lower() not in expected_suffixes:
        raise BridgeError(
            code="validation_error",
            message="Unexpected output extension.",
            details={"path": path, "expected_suffixes": list(expected_suffixes)},
        )

    parent = candidate.parent
    if not parent.exists() or not parent.is_dir():
        raise BridgeError(
            code="validation_error",
            message="Output parent directory does not exist.",
            details={"path": path, "parent": str(parent)},
        )

    if not os.access(parent, os.W_OK):
        raise BridgeError(
            code="validation_error",
            message="Output parent directory is not writable.",
            details={"path": path, "parent": str(parent)},
        )

    return candidate
