#!/usr/bin/env python3
"""Checks that every member function declared in the classes of the
wxAutoExcel headers is defined, either inline or in the sources.

A declared but undefined function compiles, but code calling it fails to
link, which no build of the library itself detects.

Exits with 1 and lists the undefined functions if there are any.
"""

import glob
import os
import re
import sys

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
INCLUDE_DIR = os.path.join(ROOT, "include", "wx")
SOURCE_DIR = os.path.join(ROOT, "src")


def strip_comments(text):
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    return re.sub(r"//[^\n]*", "", text)


def class_bodies(text):
    """Yields (class name, body) of the exported classes."""
    for match in re.finditer(r"class\s+WXDLLIMPEXP_WXAUTOEXCEL\s+(\w+)[^{;]*\{", text):
        end = match.end()
        depth = 1
        while depth and end < len(text):
            depth += {"{": 1, "}": -1}.get(text[end], 0)
            end += 1
        yield match.group(1), text[match.end():end - 1]


def declared_functions(body):
    """Returns the names of the functions declared but not defined in the
    class body, skipping pure virtual functions, operators and macros."""
    names = set()
    statement = ""
    depth = 0
    for ch in body:
        if ch == "{":
            depth += 1
            if depth == 1:
                statement = ""  # an inline definition
            continue
        if ch == "}":
            depth -= 1
            continue
        if depth:
            continue
        if ch != ";":
            statement += ch
            continue
        declaration = re.sub(r".*\b(?:public|protected|private)\s*:", "", statement, flags=re.S).strip()
        statement = ""
        if re.search(r"\b(operator|friend|typedef)\b|=\s*0\s*$", declaration):
            continue
        match = re.match(r"(?:(?:virtual|static|explicit)\s+)*(?:[\w:<>*&,\s]+?\s+)?[*&]?(~?\w+)\s*\(", declaration)
        # skip macros such as wxDECLARE_NO_COPY_CLASS()
        if match and not re.fullmatch(r"(?:wx)?[A-Z0-9_]+", match.group(1)):
            names.add(match.group(1))
    return names


def main():
    sources = "".join(strip_comments(open(path, encoding="latin-1").read())
                      for path in glob.glob(os.path.join(SOURCE_DIR, "*.cpp")))
    defined = set(re.findall(r"\b(\w+)::(~?\w+)\s*\(", sources))

    undefined = []
    for path in sorted(glob.glob(os.path.join(INCLUDE_DIR, "wxAutoExcel*.h"))):
        text = strip_comments(open(path, encoding="latin-1").read())
        for class_name, body in class_bodies(text):
            for name in sorted(declared_functions(body)):
                if (class_name, name) not in defined:
                    undefined.append(f"{os.path.basename(path)}: {class_name}::{name}()")

    if undefined:
        print("Declared but not defined:")
        print("\n".join("  " + line for line in undefined))
        return 1
    print("All declared functions are defined.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
