import sys
import os
import pynvim
import string
import subprocess
import time
from pathlib import Path
from typing import Union

import psutil
import win32process
import win32gui
import win32con


def topWindowByPid(pidChk: int):
    def topWindowByPid_wrap(hwnd: int, top_windows: list[int]):
        pid = win32process.GetWindowThreadProcessId(hwnd)[-1]
        if pid == pidChk:
            if win32gui.GetParent(hwnd) == 0:
                top_windows.append(hwnd)

    return topWindowByPid_wrap


def raiseWindow(pid: Union[int, str]):
    top_windows = []
    nvimQtPid = psutil.Process(int(pid)).ppid()
    win32gui.EnumWindows(topWindowByPid(nvimQtPid), top_windows)
    if not top_windows:
        return
    hwnd = top_windows[0]

    # from: https://stackoverflow.com/a/19733298
    # Make sure the window isn't minimized
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # flags for "Don't move or resize the window"
    flags = win32con.SWP_NOSIZE | win32con.SWP_NOMOVE

    # Make the window temporarily "always on top", then put it back
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
    win32gui.SetWindowPos(
        hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags | win32con.SWP_SHOWWINDOW
    )


def findParentGitRepo(path: str):
    for par in Path(path).parents:
        if (par / ".git").is_dir():
            return str(par)
    return None


def launchAndAttach(cmd: list, pipe: str):
    openCmd = cmd + ["--listen", pipe]
    # start the subprocess
    subprocess.Popen(openCmd, env=os.environ)

    # try attaching until we get it
    for _ in range(100):
        try:
            nvim = pynvim.attach("socket", path=pipe)
        except (OSError, RuntimeError):
            time.sleep(0.01)
        else:
            return nvim

    raise RuntimeError("Too ManyLoops")


def waitUntilVimEnter(nvim):
    # make an event in nvim that notifies us when its done loading
    event = "readytoload"
    nvim.subscribe(event)
    nvim.command(
        f'autocmd VimEnter * ++once call rpcnotify({nvim.channel_id}, "{event}")'
    )

    # wait until that event is triggered
    nvim.next_message()


def openWithNeovim(paths: list[str]):
    pipeName = "nvim-listen"
    root = findParentGitRepo(paths[0])
    if root is not None:
        pipeName = "".join([i for i in root if i in string.ascii_letters])

    os.environ["QT_SCALE_FACTOR"] = "1"
    serverpipe = r"\\.\pipe\{0}".format(pipeName)

    try:
        nvim = pynvim.attach("socket", path=serverpipe)
    except (OSError, RuntimeError):
        cmd = [r"C:\Program Files\Neovim\bin\nvim-qt.exe", "--"]
        nvim = launchAndAttach(cmd, serverpipe)
        waitUntilVimEnter(nvim)
        if root is not None:
            nvim.command(f"cd {root}")

    for path in paths:
        nvim.command(f"e {path}")

    raiseWindow(nvim.command_output("echo getpid()"))

    nvim.close()


if __name__ == "__main__":
    openWithNeovim(sys.argv[1:])
