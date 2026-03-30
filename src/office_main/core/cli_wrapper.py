"""
CLI wrapper for cli-anything-libreoffice.

This module provides a Python interface to the cli-anything-libreoffice
command-line tool, handling command construction, execution, and error parsing.
"""

import json
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from cli_anything.libreoffice.core.document import create_document

    HAS_CLI_ANYTHING = True
except ImportError:
    HAS_CLI_ANYTHING = False
    create_document = None


class LibreOfficeCLI:
    """Main interface to cli-anything-libreoffice."""

    def __init__(self, project_path: Optional[str] = None, json_output: bool = True):
        """
        Initialize the CLI wrapper.

        Args:
            project_path: Path to .lo-cli.json project file
            json_output: Whether to request JSON output from commands
        """
        self.project_path = project_path
        self.json_output = json_output
        self._session_file: Optional[str] = None

    def _run_command(self, args: List[str]) -> Dict[str, Any]:
        """
        Run a cli-anything-libreoffice command and return parsed output.

        Args:
            args: Command arguments (e.g., ["writer", "list"])

        Returns:
            Parsed JSON output or raises exception on error
        """
        cmd = ["cli-anything-libreoffice"]

        if self.json_output:
            cmd.append("--json")

        if self.project_path:
            cmd.extend(["--project", self.project_path])

        cmd.extend(args)

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=False,  # Don't decode yet, handle encoding manually
                check=True,
            )

            if self.json_output:
                # Try to decode with UTF-8, fallback to latin-1 for any binary bytes
                try:
                    stdout_text = result.stdout.decode("utf-8")
                except UnicodeDecodeError:
                    stdout_text = result.stdout.decode("latin-1")
                try:
                    return json.loads(stdout_text)
                except json.JSONDecodeError as e:
                    error_msg = f"Failed to parse JSON output from command: {' '.join(cmd)}\n"
                    error_msg += f"Output (first 500 chars): {stdout_text[:500]}\n"
                    error_msg += f"JSON decode error: {e}"
                    raise RuntimeError(error_msg) from e
            else:
                stdout_text = result.stdout.decode("utf-8", errors="replace")
                stderr_text = result.stderr.decode("utf-8", errors="replace")
                return {"output": stdout_text, "stderr": stderr_text}

        except subprocess.CalledProcessError as e:
            error_msg = f"Command failed: {' '.join(cmd)}\n"
            if e.stdout:
                try:
                    error_msg += f"STDOUT: {e.stdout.decode('utf-8')}\n"
                except UnicodeDecodeError:
                    error_msg += "STDOUT: (contains non-UTF-8 output)\n"
            if e.stderr:
                try:
                    error_msg += f"STDERR: {e.stderr.decode('utf-8')}\n"
                except UnicodeDecodeError:
                    error_msg += "STDERR: (contains non-UTF-8 output)\n"
            raise RuntimeError(error_msg) from e

    def _get_document_type(self, file_path: str) -> str:
        """Determine document type from file extension."""
        path = Path(file_path)
        ext = path.suffix.lower()
        if ext in [".docx", ".doc", ".odt", ".rtf"]:
            return "writer"
        elif ext in [".xlsx", ".xls", ".ods", ".csv"]:
            return "calc"
        elif ext in [".pptx", ".ppt", ".odp"]:
            return "impress"
        else:
            # Default to writer for unknown types
            return "writer"

    def start_session(self, document_path: str) -> str:
        """
        Start a new operation log session for tracking document operations.

        Note: This is NOT the same as the cli-anything-libreoffice 'session' command
        (which handles undo/redo). This creates a temporary JSON file to track
        operations performed through this wrapper.

        Args:
            document_path: Path to the document to open

        Returns:
            Path to the temporary session file (can be used as --project argument)
        """
        document_path = os.path.abspath(document_path)

        if HAS_CLI_ANYTHING and create_document:
            # Create a valid .lo-cli.json project file
            doc_type = self._get_document_type(document_path)
            name = Path(document_path).stem
            project = create_document(doc_type=doc_type, name=name)

            # Add original document path to metadata for reference
            project["original_document"] = document_path

            with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
                json.dump(project, f)
                self._session_file = f.name
        else:
            # Fallback to old behavior (creates invalid project file)
            with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
                session_data = {"document": document_path, "operations": []}
                json.dump(session_data, f)
                self._session_file = f.name

        assert self._session_file is not None
        return self._session_file

    def end_session(self):
        """End the current operation log session and clean up temporary files.

        Note: This cleans up the temporary file created by start_session().
        It does NOT affect the cli-anything-libreoffice 'session' command state.
        """
        if self._session_file and os.path.exists(self._session_file):
            os.unlink(self._session_file)
            self._session_file = None

    def writer(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a writer (Word) command.

        Args:
            subcommand: Writer subcommand (e.g., 'add-paragraph', 'list')
            positional: Optional list of positional arguments
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["writer", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                option = f"--{key.replace('_', '-')}"
                if isinstance(value, bool):
                    # Boolean flags: only include if True
                    if value:
                        args.append(option)
                else:
                    args.extend([option, str(value)])
        return self._run_command(args)

    def calc(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a calc (Excel) command.

        Args:
            subcommand: Calc subcommand (e.g., 'set-cell', 'get-cell')
            positional: Optional list of positional arguments
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["calc", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                option = f"--{key.replace('_', '-')}"
                if isinstance(value, bool):
                    # Boolean flags: only include if True
                    if value:
                        args.append(option)
                else:
                    args.extend([option, str(value)])
        return self._run_command(args)

    def impress(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute an impress (PowerPoint) command.

        Args:
            subcommand: Impress subcommand (e.g., 'add-slide', 'list-slides')
            positional: Optional list of positional arguments
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["impress", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                option = f"--{key.replace('_', '-')}"
                if isinstance(value, bool):
                    # Boolean flags: only include if True
                    if value:
                        args.append(option)
                else:
                    args.extend([option, str(value)])
        return self._run_command(args)

    def export(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute an export command.

        Args:
            subcommand: Export subcommand (e.g., 'render', 'presets')
            positional: Optional list of positional arguments (e.g., output path for render)
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["export", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                option = f"--{key.replace('_', '-')}"
                if isinstance(value, bool):
                    # Boolean flags: only include if True
                    if value:
                        args.append(option)
                else:
                    args.extend([option, str(value)])
        return self._run_command(args)

    def document(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a document management command.

        Args:
            subcommand: Document subcommand (e.g., 'new', 'open', 'save', 'info')
            positional: Optional list of positional arguments (e.g., path for open)
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["document", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def session(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a session management command.

        Args:
            subcommand: Session subcommand (e.g., 'history', 'status', 'undo', 'redo')
            positional: Optional list of positional arguments
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["session", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def style(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a style management command.

        Args:
            subcommand: Style subcommand (e.g., 'create', 'modify', 'apply', 'list')
            positional: Optional list of positional arguments (e.g., style name)
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["style", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def batch(
        self, subcommand: str, positional: Optional[List[str]] = None, **kwargs
    ) -> Dict[str, Any]:
        """Execute a batch processing command.

        Args:
            subcommand: Batch subcommand (e.g., '', 'run' - often empty for default)
            positional: Optional list of positional arguments (e.g., input file)
            **kwargs: Keyword arguments converted to command options

        Returns:
            Parsed JSON output or raw output dictionary
        """
        if positional is None:
            positional = []
        args = ["batch", subcommand]
        args.extend(positional)
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def repl(self):
        """Start interactive REPL session.

        Note: Interactive REPL requires terminal input/output handling.
        Consider using batch() method for non-interactive command sequences.

        Raises:
            NotImplementedError: REPL mode not supported in this wrapper
        """
        # This would require interactive input
        raise NotImplementedError(
            "REPL mode requires interactive terminal. "
            "Use batch() method for non-interactive command sequences."
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_session()
