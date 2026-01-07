from __future__ import annotations

import os
import re
import sys
import unittest
from pathlib import Path
from subprocess import run
from typing import Generator, List

import importtime_convert

forbidden_imports: List[re.Pattern[str]] = [
    # These imports are forbidden at startup because they are slow.
    re.compile(r"^pandas(\..+)?$"),
    re.compile(r"^numpy(\..+)?$"),
]

repo_dir = Path(__file__).resolve().parent.parent
module_name = "aa_py_openpyxl_util"


class TestImportSpeed(unittest.TestCase):
    """
    Prevent imports from slowing down the startup time.
    """

    def test_module_import(self) -> None:
        completed = run(
            args=[
                sys.executable,
                # Note: Using "-X importtime" here confuses the Python debugger in IntelliJ IDEA.
                "-X",
                "importtime",
                # Import our module and show where it was imported from.
                "-c",
                f"import {module_name}; print({module_name}.__path__[0]);",
            ],
            capture_output=True,
            check=False,
            text=True,
            env={
                **os.environ,
                "PYTHONPATH": repo_dir.as_posix(),
            },
        )

        stdout = completed.stdout or ""
        stderr = completed.stderr or ""

        # This is useful in case of test failures.
        print(
            "\n".join(
                line
                for line in stderr.splitlines()
                if not line.startswith("import time:")
            )
        )

        top_level_imports = importtime_convert.parse(stderr)
        for top_level_import in top_level_imports:
            assign_parents(imp=top_level_import)
            for imp in traverse_import_tree(imp=top_level_import):
                if any(pattern.match(imp.package) for pattern in forbidden_imports):
                    self.fail(
                        f"The import of '{imp.package}' is forbidden at startup because it is slow.\n"
                        "It should be converted to a lazy import or type-checking import.\n"
                        f"{format_import_chain(imp)}"
                    )

        # Check that our module was indeed imported.
        self.assertTrue(
            any(
                imp.package == module_name
                for top_level_import in top_level_imports
                for imp in traverse_import_tree(imp=top_level_import)
            )
        )

        # Check that our module was imported from the correct location.
        self.assertEqual(str(repo_dir / module_name), stdout.strip())


def format_import_chain(imp: importtime_convert.Import) -> str:
    lines = []
    current: importtime_convert.Import | None = imp
    while current is not None:
        lines.append(
            f"  - {current.package} (Cumulative: {current.cumulative_us:.2f} µs, Self: {current.self_us:.2f} µs)"
        )
        current = getattr(current, "parent", None)
    return "Import chain:\n" + "\n".join(reversed(lines))


def traverse_import_tree(
    *,
    imp: importtime_convert.Import,
) -> Generator[importtime_convert.Import, None, None]:
    yield imp
    for child in imp.subimports:
        yield from traverse_import_tree(imp=child)


def assign_parents(*, imp: importtime_convert.Import) -> None:
    for child in imp.subimports:
        setattr(child, "parent", imp)
        assign_parents(imp=child)
