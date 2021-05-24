#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import logging
import os
import shlex
import shutil
import subprocess
import tempfile

from pathlib import Path
from typing import Iterable, Optional, Union

# Build target / Runtime Identifier (RID)
# Reference: https://docs.microsoft.com/en-us/dotnet/core/rid-catalog
WORKING_DIRECTORY = Path(__file__).parent
os.chdir(WORKING_DIRECTORY)
OUTPUT_DIRECTORY = "build"

DOTNET_PUBLISH_COMMAND = ["dotnet", "publish"]
PROJECT_FILE = (
    Path(__file__).parent / "YTSubConverter.CLI" / "YTSubConverter.CLI.csproj"
)
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
    default_fmt = "[{levelname}] {message}"
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


def init_logging(
    *,
    level: Union[int, str] = logging.INFO,
    formatter: Optional[logging.Formatter] = None,
):
    formatter = formatter if formatter is not None else CustomFormatter()
    global logger, console_handler
    logger = logging.getLogger("publish.py")
    logger.setLevel(level)
    # Avoid double output when init_logging is called several times
    for handler in logger.handlers:
        logger.removeHandler(handler)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)


init_logging()


class BuildScriptError(Exception):
    ...


def remove_windows_symbols(
    files: Iterable[str] = FILES_WITH_WINDOWS_SYMBOLS, dry_run: bool = False
):
    for fname in files:
        if not fname.casefold().endswith(".cs"):
            raise BuildScriptError(
                "Expected only .cs files to be stripped of WINDOWS symbols"
            )
        if not Path(fname).exists():
            logger.warning(
                f"Specified file to be stripped [{shlex.quote(fname)}] does not exist."
            )
            continue
        if dry_run:
            logger.info(
                f"Going to strip WINDOWS symbols from file [{shlex.quote(fname)}]."
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
        logger.info(f"Stripped WINDOWS symbols from file {shlex.quote(fname)}")


def parse_args(args=None):
    parser = argparse.ArgumentParser(
        description=(
            "Publish script for YTSubConverter-CLI with support for "
            "specifying target platforms via Resource Identifiers or "
            '.NET RIDs. Generates and runs a "dotnet publish" command.'
        )
    )
    parser.add_argument(
        "project",
        help="The project file to publish.",
        nargs="?",
        default=str(PROJECT_FILE),
    )
    parser.add_argument(
        "--dry-run",
        help="Do not run anything; print the generated publish command and exit.",
        action="store_true",
    )
    # Some argument descriptions are taken and/or adapted from the
    # command `dotnet publish --help`
    parser.add_argument(
        "-v",
        "--verbose",
        help="Enable verbose output (set logger level to DEBUG).",
        action="store_true",
    )
    parser.add_argument(
        "-r",
        "--runtime",
        "--target",
        help=(
            'Resource Identifier (RID) to target. Defaults to "any". '
            "This is ignored when single-file build is disabled. "
            "[-s / --single-file is not specified]."
        ),
        nargs="?",
        metavar="RUNTIME_IDENTIFIER",
        default="any",
    )
    parser.add_argument(
        "-c",
        "--configuration",
        help='The configuration to publish for. The default is "Release".',
        nargs="?",
        metavar="CONFIGURATION",
        default="Release",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "The output directory to place the published artifacts in. "
            'Defaults to "./build".'
        ),
        nargs="?",
        default=str(OUTPUT_DIRECTORY),
    )
    parser.add_argument(
        "-k",
        "--keep",
        help="Do not delete the contents of an already-existing build directory.",
        action="store_true",
    )
    parser.add_argument(
        "-f",
        "--force-restore",
        help=(
            "Force all dependencies to be resolved even if the last "
            "restore was successful."
        ),
        action="store_true",
    )
    parser.add_argument(
        "-p",
        "--portable",
        help=(
            "Enable portable build (overrides -n / --non-portable). "
            "Packages the .NET Runtime along with the resulting "
            "binary. The resulting build is also referred to as "
            '"self-contained".'
        ),
        action="store_true",
    )
    parser.add_argument(
        "-n",
        "--non-portable",
        help=(
            "Disable portable build. The resulting binaries will "
            "require that the .NET Runtime be installed separately on "
            "the system to function. The resulting build is also "
            'referred to as "framework-dependent".'
        ),
        action="store_true",
    )
    parser.add_argument(
        "-s",
        "--single-file",
        help=(
            "Enable single-file build (either portable or "
            "non-portable). Note that this option is independent of "
            "[-p / --portable] or [-n / --non-portable]."
        ),
        action="store_true",
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
    if args is None:
        return parser.parse_args()
    else:
        return parser.parse_args(args)


def main(argv: bool = None, publish_command: Iterable[str] = DOTNET_PUBLISH_COMMAND):
    publish_command = [s for s in publish_command]

    args = parse_args(argv)

    if args.verbose:
        init_logging(level=logging.DEBUG)
    logger.debug(f"Raw command-line arguments: {args}")

    if args.remove_windows_symbols is None:
        if not args.runtime.startswith("win"):
            args.remove_windows_symbols = True
            logger.info(
                "Detected a non-Windows runtime identifer "
                f"[{args.runtime}], stripping WINDOWS symbols from "
                ".cs files."
            )
    elif args.remove_windows_symbols:
        if args.runtime.startswith("win"):
            logger.warning(
                'Target RID "{args.runtime}" starts with "win" but '
                '"WINDOWS" symbols are going to be stripped. If you '
                "did not intend to do this, remove (-L / "
                "--not-windows / --remove-windows-symbols) from the "
                "command-line arguments."
            )
    else:
        if not args.runtime.startswith("win"):
            logger.warning(
                'Target RID "{args.runtime}" does not start with '
                '"win", but "WINDOWS" symbols are not going to be '
                "stripped. If you did not intend to do this, add "
                "[-L / --not-windows / --remove-windows-symbols] as a "
                "command-line argument."
            )
    logger.debug(f"Parsed arguments: {args}")

    # Things to do before compilation
    if args.remove_windows_symbols:
        remove_windows_symbols(dry_run=args.dry_run)

    output_path = Path(args.output)
    if output_path.exists() and output_path.is_file():
        print(output_path)
        raise NotADirectoryError(
            f"Output directory [{shlex.quote(str(output_path))}] "
            "exists and is a file."
        )
    if not args.keep:
        logger.info(f"Cleaning up build directory [{shlex.quote(str(output_path))}].")
        shutil.rmtree(output_path, ignore_errors=True)
    output_path.mkdir(exist_ok=True)

    # Generation of final publish command
    if args.force_restore:
        publish_command.append("--force")

    publish_command.append("--configuration")
    publish_command.append(args.configuration)

    publish_command.append("--runtime")
    publish_command.append(args.runtime)

    publish_command.append("--output")
    publish_command.append(args.output)

    if args.single_file:
        publish_command.extend(FLAGS_SINGLE_FILE)
    if args.portable:
        publish_command.extend(FLAGS_PORTABLE)
    elif args.non_portable:
        publish_command.extend(FLAGS_NON_PORTABLE)
    else:
        logger.warning("The portable build was neither enabled nor disabled.")
        logger.warning(
            'Falling back to "dotnet publish" default behavior '
            "(framework-dependent / non-portable)."
        )

    publish_command.append(args.project)

    if args.dry_run:
        logger.info(f"Publish command: {publish_command}")
    else:
        logger.info(f"Publish command: {publish_command}")
        subprocess.run(publish_command)


if __name__ == "__main__":
    main()
