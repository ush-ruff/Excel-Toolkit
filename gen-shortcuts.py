"""Parse excel-toolkit.ahk and generate a markdown shortcut table."""

import re
import sys

ahk_file = sys.argv[1] if len(sys.argv) > 1 else "excel-toolkit.ahk"
out_file = "shortcuts.md"

with open(ahk_file, encoding="utf-8") as f:
    lines = f.readlines()

rows = []
i = 0
while i < len(lines):
    line = lines[i]
    m = re.match(r'^;\s*Key:\s*(.+)$', line)
    if m:
        key_combo = m.group(1).strip()
        # Collect description from lines above
        desc_lines = []
        j = i - 1
        while j >= 0:
            prev = lines[j].strip()
            if prev.startswith("; ===") or prev.startswith("; Key:"):
                break
            if prev == "" or prev.startswith("; ---"):
                j -= 1
                continue
            # Skip commented-out hotkeys (e.g. ; ^!4::...)
            if prev.startswith("; ") and "::" in prev[2:]:
                j -= 1
                continue
            if prev.startswith("; "):
                desc_lines.insert(0, prev[2:].lstrip())
            elif prev.startswith(";"):
                desc_lines.insert(0, prev[1:].lstrip())
            else:
                break
            j -= 1
        description = " ".join(desc_lines).strip()
        if not description:
            # Fall back to function name from the hotkey line below
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if '::' in nxt:
                    fn = nxt.split('::', 1)[1].strip()
                    fn = re.sub(r'\(.*', '', fn)
                    # Make it readable
                    fn = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', fn)
                    description = fn
        rows.append((key_combo, description))
    i += 1

# Write output
with open(out_file, "w", encoding="utf-8") as f:
    f.write("# Excel Toolkit Shortcuts\n\n")
    f.write("![Shortcuts](images/shortcuts.png)\n\n")
    f.write("|Key|Action|\n")
    f.write("|---|---|\n")
    for key, action in rows:
        f.write(f"|{key}|{action}|\n")

print(f"Written {len(rows)} shortcuts to {out_file}")
