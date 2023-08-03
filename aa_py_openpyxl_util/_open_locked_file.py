import os
from typing import Callable, Union, Tuple, IO, Any, Optional
import os
from pathlib import Path


open_locked = open

if os.name == "nt":
    import msvcrt
    import ctypes
    from ctypes.wintypes import HANDLE

    class open_locked:  # type: ignore
        def __init__(self, filename: Union[str, Path], mode: str) -> None:
            """
            Constructor for the `open_locked` class.

            Args:
                filename (Union[str, Path]): The path of the file to open.
                mode (str): The mode in which the file is to be opened.
            """
            self.filename = filename
            self.mode = mode
            self.file: Optional[IO[Any]] = None
            self.hfile = None

        def __enter__(self) -> IO[Any]:
            """
            Enter method for the `open_locked` class.

            Returns:
                IO[Any]: The opened file.
            """
            self.file, self.hfile = _open_locked(self.filename, self.mode)
            return self.file

        def __exit__(self, exc_type: str, exc_val: str, exc_tb: str) -> None:
            """
            Exit method for the `open_locked` class.

            Args:
                exc_type (str): The type of exception.
                exc_val (str): The value of exception.
                exc_tb (str): The traceback of exception.
            """
            if self.file is not None:
                self.file.close()
            ctypes.windll.kernel32.CloseHandle(self.hfile)

    def _open_locked(
        filename: Union[str, Path], mode: str = "r"
    ) -> Tuple[IO[Any], Any]:
        """
        Helper function to open a file with a lock.

        Args:
            filename (Union[str, Path]): The path of the file to open.
            mode (str, optional): The mode in which the file is to be opened. Defaults to "r".

        Raises:
            ValueError: If an unsupported file access mode is provided.
            FileNotFoundError: If the file is not found.
            PermissionError: If permission is denied.
            ValueError: If an invalid parameter is provided.
            FileExistsError: If the file already exists.
            IsADirectoryError: If the given filename is a directory.
            OSError: If an unknown error occurs.

        Returns:
            Tuple[IO[Any], Any]: The opened file and the file handle.
        """
        GENERIC_READ = 0x80000000
        GENERIC_WRITE = 0x40000000
        FILE_SHARE_READ = 1
        FILE_SHARE_WRITE = 2
        FILE_SHARE_DELETE = 4

        OPEN_EXISTING = 3
        CREATE_ALWAYS = 2
        OPEN_ALWAYS = 4
        FILE_END = 2

        access_flags = {
            "r": GENERIC_READ,
            "w": GENERIC_WRITE,
            "rw": GENERIC_READ | GENERIC_WRITE,
            "rb": GENERIC_READ,
            "wb": GENERIC_WRITE,
            "a": GENERIC_WRITE,  # for 'append' mode
            "ab": GENERIC_WRITE,  # for 'append' binary mode
        }

        dispositions = {
            "r": OPEN_EXISTING,
            "w": CREATE_ALWAYS,
            "rw": OPEN_ALWAYS,
            "rb": OPEN_EXISTING,
            "wb": CREATE_ALWAYS,
            "a": OPEN_ALWAYS,  # for 'append' mode
            "ab": OPEN_ALWAYS,  # for 'append' binary mode
        }

        CreateFileW = ctypes.windll.kernel32.CreateFileW

        access_mode = access_flags.get(mode)
        if access_mode is None:
            raise ValueError(
                f"Invalid mode '{mode}', only 'r', 'w', 'rw', 'rb', 'wb', 'a', and 'ab' are supported"
            )

        disposition = dispositions.get(mode)
        if disposition is None:
            raise ValueError(
                f"Invalid mode '{mode}', only 'r', 'w', 'rw', 'rb', 'wb', 'a', and 'ab' are supported"
            )

        hfile = CreateFileW(
            str(Path(filename)),
            access_mode,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            disposition,
            0,
            None,
        )

        UNKNOWN_ERROR = 0
        INVALID_HANDLE_VALUE = -1
        ERROR_FILE_NOT_FOUND = 2
        ERROR_ACCESS_DENIED = 5
        ERROR_INVALID_PARAMETER = 87
        ERROR_FILE_EXISTS = 80
        ERROR_ALREADY_EXISTS = 183
        ERROR_DIRECTORY = 267

        if hfile == INVALID_HANDLE_VALUE:
            # Get the last error code
            error_code = ctypes.GetLastError()

            if error_code == ERROR_FILE_NOT_FOUND:
                raise FileNotFoundError(f"No such file or directory: '{filename}'")
            elif error_code == ERROR_ACCESS_DENIED:
                raise PermissionError("Permission denied: '{}'".format(filename))
            elif error_code == ERROR_INVALID_PARAMETER:
                raise ValueError("Invalid parameter")
            elif error_code == ERROR_FILE_EXISTS or error_code == ERROR_ALREADY_EXISTS:
                raise FileExistsError("File already exists: '{}'".format(filename))
            elif error_code == ERROR_DIRECTORY:
                raise IsADirectoryError("Is a directory: '{}'".format(filename))
            elif error_code == UNKNOWN_ERROR:
                raise OSError("Unknown error")
            else:
                raise ctypes.WinError()

        # for 'append' mode, you'd also need to move the file pointer to the end
        if mode in {"a", "ab"}:
            ctypes.windll.kernel32.SetFilePointer(hfile, 0, None, FILE_END)

        # Convert the Windows handle into a C runtime file descriptor
        fd = msvcrt.open_osfhandle(hfile, os.O_BINARY if "b" in mode else os.O_TEXT)

        # Create a Python file object from the file descriptor
        file = os.fdopen(fd, mode)

        return file, hfile
