from pathlib import Path

# -----------------------------
# CONFIG
# -----------------------------
ROOT_DIR = Path(".")          # โฟลเดอร์โปรเจกต์
OUTPUT_FILE = "project_tree.txt"

EXCLUDE_DIRS = {
    ".venv", "venv", "env",
    "__pycache__",
    ".git", ".idea", ".vscode",
    "node_modules",
    "dist", "build",
    "output", "logs",
}

EXCLUDE_SUFFIXES = {
    ".pyc", ".pyo", ".log", ".exe", ".zip", ".tmp"
}

MAX_DEPTH = 6   # กันไม่ให้ลึกเกินจำเป็น


# -----------------------------
# TREE BUILDER
# -----------------------------
def build_tree(path: Path, prefix="", depth=0, lines=None):
    if lines is None:
        lines = []

    if depth > MAX_DEPTH:
        return lines

    entries = sorted(
        path.iterdir(),
        key=lambda p: (p.is_file(), p.name.lower())
    )

    filtered = []
    for p in entries:
        if p.is_dir() and p.name in EXCLUDE_DIRS:
            continue
        if p.is_file() and p.suffix in EXCLUDE_SUFFIXES:
            continue
        filtered.append(p)

    for idx, p in enumerate(filtered):
        connector = "└─ " if idx == len(filtered) - 1 else "├─ "
        lines.append(f"{prefix}{connector}{p.name}")

        if p.is_dir():
            extension = "   " if idx == len(filtered) - 1 else "│  "
            build_tree(
                p,
                prefix + extension,
                depth + 1,
                lines
            )

    return lines


# -----------------------------
# MAIN
# -----------------------------
def main():
    root_name = ROOT_DIR.resolve().name
    lines = [f"{root_name}/"]
    build_tree(ROOT_DIR, lines=lines)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"Project structure written to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
