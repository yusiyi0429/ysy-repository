#!/usr/bin/env python3
import re
from pathlib import Path

path = Path(__file__).resolve().parent.parent / "app_server.py"
text = path.read_text(encoding="utf-8")

# Remove # region agent log ... # endregion blocks (with optional leading whitespace)
text = re.sub(
    r"[ \t]*# region agent log\r?\n[ \t]*_agent_debug_log\([\s\S]*?[ \t]*# endregion\r?\n",
    "",
    text,
)

# Remove _debug_log(...) calls
while True:
    idx = text.find("_debug_log(")
    if idx == -1:
        break
    # find start of line
    line_start = text.rfind("\n", 0, idx) + 1
    # find matching paren
    depth = 0
    i = idx + len("_debug_log")
    while i < len(text):
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth -= 1
            if depth == 0:
                end = i + 1
                if end < len(text) and text[end : end + 1] == "\n":
                    end += 1
                text = text[:line_start] + text[end:]
                break
        i += 1
    else:
        break

path.write_text(text, encoding="utf-8")
print("remaining", len(re.findall(r"_debug_log|_agent_debug_log", text)))
