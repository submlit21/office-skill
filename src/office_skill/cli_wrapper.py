"""
CLI wrapper for cli-anything-libreoffice.

This module provides a Python interface to the cli-anything-libreoffice
command-line tool, handling command construction, execution, and error parsing.
"""

import json
import subprocess
import tempfile
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Union


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
        self._session_file = None

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
                text=True,
                check=True
            )

            if self.json_output:
                return json.loads(result.stdout)
            else:
                return {"output": result.stdout, "stderr": result.stderr}

        except subprocess.CalledProcessError as e:
            error_msg = f"Command failed: {' '.join(cmd)}\n"
            error_msg += f"STDOUT: {e.stdout}\n" if e.stdout else ""
            error_msg += f"STDERR: {e.stderr}\n" if e.stderr else ""
            raise RuntimeError(error_msg) from e

    def start_session(self, document_path: str) -> str:
        """
        Start a new session with a document.

        Args:
            document_path: Path to the document to open

        Returns:
            Session ID for subsequent operations
        """
        # Create a temporary session file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            session_data = {
                "document": os.path.abspath(document_path),
                "operations": []
            }
            json.dump(session_data, f)
            self._session_file = f.name

        return self._session_file

    def end_session(self):
        """End the current session and clean up temporary files."""
        if self._session_file and os.path.exists(self._session_file):
            os.unlink(self._session_file)
            self._session_file = None

    def writer(self, subcommand: str, **kwargs) -> Dict[str, Any]:
        """Execute a writer (Word) command."""
        args = ["writer", subcommand]
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def calc(self, subcommand: str, **kwargs) -> Dict[str, Any]:
        """Execute a calc (Excel) command."""
        args = ["calc", subcommand]
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def impress(self, subcommand: str, **kwargs) -> Dict[str, Any]:
        """Execute an impress (PowerPoint) command."""
        args = ["impress", subcommand]
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def export(self, subcommand: str, **kwargs) -> Dict[str, Any]:
        """Execute an export command."""
        args = ["export", subcommand]
        for key, value in kwargs.items():
            if value is not None:
                args.extend([f"--{key.replace('_', '-')}", str(value)])
        return self._run_command(args)

    def batch(self, commands_file: str) -> Dict[str, Any]:
        """
        Execute multiple commands from a file.

        Args:
            commands_file: Path to file containing commands (one per line)
        """
        return self._run_command(["batch", commands_file])

    def repl(self):
        """Start interactive REPL session."""
        # This would require interactive input
        raise NotImplementedError("REPL mode requires interactive terminal")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_session()