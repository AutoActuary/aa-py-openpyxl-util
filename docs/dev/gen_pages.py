from pathlib import Path
import mkdocs_gen_files


def main() -> None:
    script_dir = Path(".").resolve()
    repo_dir = script_dir.parent.parent
    package_name = "aa_py_openpyxl_util"
    src_dir = repo_dir.joinpath(package_name)

    nav = mkdocs_gen_files.Nav()

    # Inject the repo's README.md file into the documentation as the starting page.
    nav["README.md"] = "index.md"
    with mkdocs_gen_files.open("index.md", "w") as dst:
        dst.write(
            """---
title: README.md
---
"""
        )
        dst.write(repo_dir.joinpath("README.md").read_text(encoding="utf8"))

    # Generate docs from source code and inject navigation for generated files.
    # https://github.com/mkdocstrings/mkdocstrings/blob/5802b1ef5ad9bf6077974f777bd55f32ce2bc219/docs/gen_doc_stubs.py
    for path in sorted(src_dir.rglob("*.py")):
        if path.name in ["__init__.py"]:
            module_path = path.parent.relative_to(repo_dir).with_suffix("")
        else:
            module_path = path.relative_to(repo_dir).with_suffix("")
        nav_path_parts = path.relative_to(repo_dir).with_suffix("").parts[1:]
        doc_path = Path(package_name, path.relative_to(src_dir).with_suffix(".md"))
        nav[nav_path_parts] = doc_path.as_posix()

        with mkdocs_gen_files.open(doc_path, "w") as f:
            f.write(
                f"""\
---
title: '{path.relative_to(src_dir).name}'
---

# `{path.relative_to(src_dir)}`

::: {".".join(module_path.parts)}
"""
            )

    with mkdocs_gen_files.open(f"nav.md", "w") as f:
        f.writelines(nav.build_literate_nav())


main()
