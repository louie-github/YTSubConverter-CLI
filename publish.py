#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import builtins
import logging
import os
import shlex
import tempfile

from pathlib import Path
from typing import Iterable, Optional, Union

# Build target / Runtime Identifier (RID)
# Reference: https://docs.microsoft.com/en-us/dotnet/core/rid-catalog
WORKING_DIRECTORY = Path(__file__).parent
os.chdir(WORKING_DIRECTORY)
OUTPUT_DIRECTORY = Path(__file__).parent / "build"
OUTPUT_DIRECTORY.mkdir(exist_ok=True)

DOTNET_PUBLISH_COMMAND = ["dotnet", "publish"]
FLAGS_PORTABLE = ["--self-contained", "true"]
FLAGS_NON_PORTABLE = ["--self-contained", "false"]
# Use IncludeNativeLibrariesForSelfExtract instead of
# IncludeAllContentForSelfExtract (which mimics .NET Core 3.x behavior)
# as per Single-file Publish .NET Design Proposal:
# https://github.com/dotnet/designs/blob/main/accepted/2020/single-file/design.md
FLAGS_SINGLE_FILE = [
    "/p:PublishSingleFile=true",
    # The proposal says we don't need this on Linux, but when I remove
    # it, an extra library is generated in the publish folder. Weird.
    "/p:IncludeNativeLibrariesForSelfExtract=true",
]
FILES_WITH_WINDOWS_SYMBOLS = ["YTSubConverter.CLI/Program.cs"]

GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS") == "true"


class GithubActionsMessageError(Exception):
    pass


class CustomFormatter(logging.Formatter):
    default_fmt = "<{name}> [{levelname}] {message}"
    style = "{"

    def __init__(
        self,
        datefmt: Optional[str] = None,
        validate: bool = True,
        use_github_actions_format: bool = GITHUB_ACTIONS,
    ):
        self.use_github_actions_format = use_github_actions_format
        super().__init__(
            fmt=self.default_fmt, datefmt=datefmt, style=self.style, validate=validate
        )

    def build_github_actions_message(
        self,
        level: str,
        message: str,
        filename: Optional[str] = None,
        lineno: Optional[Union[str, int]] = None,
        colno: Optional[Union[str, int]] = None,
    ):
        if level not in {"debug", "warning", "error"}:
            raise GithubActionsMessageError(
                f"Expected a level of debug, warning, or error, got {level}"
            )
        message_attributes = []
        if filename:
            message_attributes.append(f"file={filename}")
        if lineno:
            message_attributes.append(f"line={lineno}")
        if colno:
            message_attributes.append(f"col={colno}")
        return f"::{level} {','.join(message_attributes)}::{message}"

    def format(self, record: logging.LogRecord) -> str:
        # This also sets the record.message attribute
        default_fmt_output = super().format(record)
        if self.use_github_actions_format:
            try:
                github_actions_message = self.build_github_actions_message(
                    level=record.levelname.casefold(),
                    message=record.message,
                    filename=record.filename,
                    lineno=record.lineno,
                )
            except GithubActionsMessageError:
                return default_fmt_output
            else:
                return "\n".join((github_actions_message, default_fmt_output))
        else:
            return default_fmt_output


def init_logging(formatter: Optional[logging.Formatter] = None):
    formatter = formatter if formatter is not None else CustomFormatter()
    global logger, console_handler
    logger = logging.getLogger("publish.py")
    logger.setLevel(logging.DEBUG)
    # Avoid double output when init_logging is called several times
    for handler in logger.handlers:
        logger.removeHandler(handler)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)


init_logging()


class BuildScriptError(Exception):
    ...


def remove_windows_symbols(files: Iterable[str] = FILES_WITH_WINDOWS_SYMBOLS):
    for fname in files:
        if not fname.casefold().endswith(".cs"):
            raise BuildScriptError(
                "Expected only .cs files to be stripped of WINDOWS symbols"
            )
        if not Path(fname).exists():
            logging.warning(
                f"Specified file to be stripped [{shlex.quote(fname)}] does not exist."
            )
            continue
        # utf-8-sig: UTF-8 with BOM
        encoding = "utf-8-sig"
        with open(fname, encoding=encoding) as in_f, tempfile.TemporaryFile(
            mode="w+t"
        ) as out_f:
            # Append #undef WINDOWS after all preprocessor directives
            line_iter = iter(in_f)
            for line in line_iter:
                if line.strip().startswith("#"):
                    out_f.write(line)
                else:
                    out_f.write("#undef WINDOWS\n")
                    break
            # Add remaining elements (if any)
            out_f.writelines(line_iter)


def parse_args(args=None):
    parser = argparse.ArgumentParser(
        description=(
            "Build script for YTSubConverter-CLI with support for "
            "specifying target platforms via Resource Identifiers or "
            ".NET RIDs."
        )
    )
    parser.add_argument(
        "-r",
        "--rid",
        "--target",
        help='Resource Identifier (RID) to target. Defaults to "any".',
        nargs="?",
        metavar="RID",
        default="any",
        dest="rid",
    )
    parser.add_argument(
        "-L",
        "--not-windows",
        "--remove-windows-symbols",
        help=(
            'Remove "WINDOWS" symbols used for conditional '
            "compilation. By default, if the specified RID starts with "
            '"win", this is set to False, otherwise True.'
        ),
        action="store_true",
        default=None,
        dest="remove_windows_symbols",
    )
    parser.add_argument(
        "-p",
        "--portable",
        help=(
            "Enable portable build (overrides -n / --non-portable)."
            "Packages the .NET Runtime along with the resulting "
            "binary. The resulting build is also referred to as "
            '"self-contained."'
        ),
        action="store_true",
        dest="portable",
    )
    parser.add_argument(
        "-n",
        "--non-portable",
        help=(
            "Disable portable build. The resulting binaries will "
            "require that the .NET Runtime be installed separately on "
            "the system to function. The resulting build is also "
            'referred to as "framework-dependent"'
        ),
        action="store_true",
        dest="non_portable",
    )
    parser.add_argument(
        "-s",
        "--single-file",
        help=(
            "Enable single-file build (either portable or "
            "non-portable). Note that this option is independent of "
            "[-p / --portable] or [-n / --non-portable]"
        ),
        action="store_true",
        dest="single_file",
    )
    if args is None:
        parsed_args = parser.parse_args()
    else:
        parsed_args = parser.parse_args(args)

    parsed_args.rid = parsed_args.rid.strip().casefold()
    if parsed_args.remove_windows_symbols is None:
        parsed_args.remove_windows_symbols = parsed_args.rid.startswith("win")
    if parsed_args.remove_windows_symbols:
        if not parsed_args.rid.startswith("win"):
            logger.warning(
                (
                    'Target RID "{parsed_args.rid}" does not start '
                    'with "win", but "WINDOWS" symbols are not going '
                    "to be stripped. If you did not intend to do this, "
                    "add [-L / --not-windows / --remove-windows-symbols] "
                    "as a command line argument."
                )
            )
    else:
        if parsed_args.rid.startswith("win"):
            logger.warning(
                (
                    'Target RID "{parsed_args.rid}" starts with "win", '
                    'but "WINDOWS" symbols are going to be stripped. '
                    "If you did not intend to do this, remove [-L / "
                    "--not-windows / --remove-windows-symbols] from "
                    "the command line arguments."
                )
            )


if __name__ == "__main__":
    init_logging(CustomFormatter(use_github_actions_format=True))
    logger.warning("Test message")
