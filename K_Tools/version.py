"""Единый источник версии K Tools для GUI и сборки."""

from __future__ import annotations


APP_VERSION = "7.2.13"
RELEASE_CHANNEL = ""

DISPLAY_VERSION = f"{APP_VERSION} {RELEASE_CHANNEL}".strip()
EXECUTABLE_BASENAME = f"K_Tools {DISPLAY_VERSION}"


if __name__ == "__main__":
    print(EXECUTABLE_BASENAME)
