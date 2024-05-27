import time
import subprocess
import sys
from textwrap import dedent
import unittest
import atexit
import warnings
import os
from aa_py_openpyxl_util import remove_atexit_permission_error

if os.name == "nt":

    def create_subprocess_lock(filename):
        lock_code = dedent(
            f"""
            import msvcrt
            import time
            f = open(r'{filename}', 'wb')
            msvcrt.locking(f.fileno(), msvcrt.LK_LOCK, 0)
            time.sleep(10)
            """
        )

        subprocess.Popen([sys.executable, "-c", lock_code])
        time.sleep(0.1)

    class TestRemoveAtexitPermissionError(unittest.TestCase):
        def test_permission_error(self) -> None:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", ResourceWarning)

                import openpyxl.worksheet._writer

                tmp = openpyxl.worksheet._writer.create_temporary_file()
                create_subprocess_lock(tmp)

                with self.assertRaises(PermissionError):
                    atexit._run_exitfuncs()
                    openpyxl.worksheet._writer._openpyxl_shutdown()

        def test_no_more_permission_error(self) -> None:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", ResourceWarning)

                import openpyxl.worksheet._writer

                tmp = openpyxl.worksheet._writer.create_temporary_file()
                create_subprocess_lock(tmp)

                remove_atexit_permission_error()
                atexit._run_exitfuncs()
                openpyxl.worksheet._writer._openpyxl_shutdown()

        def test_only_register_once(self) -> None:
            remove_atexit_permission_error()
            n = atexit._ncallbacks()
            self.assertEqual(n, atexit._ncallbacks())

            remove_atexit_permission_error()
            self.assertEqual(n, atexit._ncallbacks())

            remove_atexit_permission_error()
            self.assertEqual(n, atexit._ncallbacks())
