"""
A really hacky Python script to set monitor settings using DDC/CI
and MCCS on Windows.
Huge thanks go out to @ThiefMaster on StackOverflow
https://stackoverflow.com/a/18065609

Tested with a BenQ GL2450H and a Dell U2417H. Your mileage
may vary.

I *highly* recommend that if you switch monitor inputs with this that
you switch in the order you expect them to be. E.g. when monitor #1
disconnects then  monitor #2 at some point will become monitor #1.
I don't know the timings for this, so work backwards (as I did in the
example scripts attached).

You can do anything with this. I just use 0x60 at the moment to choose
monitor input.

The full command list is available in mccs.pdf (copied into this repo, came
from the following URL: https://milek7.pl/ddcbacklight/mccs.pdf)
Usually I wouldn't mirror a file like this on GitHub but the previous URL I
looked for has been taken down.

General usage:
mccs.py: shows monitors and their indices
mccs.py script.file: runs all commands in a script file. # lines denote comments.
mccs.py x 0x0A 0x0B: Runs command 0x0A with parameter 0x0B on monitor index x.

Good luck!
"""
import sys

from ctypes import windll, byref, Structure, WinError, POINTER, WINFUNCTYPE
from ctypes.wintypes import BOOL, HMONITOR, HDC, RECT, LPARAM, DWORD, BYTE, WCHAR, HANDLE


_MONITORENUMPROC = WINFUNCTYPE(BOOL, HMONITOR, HDC, POINTER(RECT), LPARAM)


class _PHYSICAL_MONITOR(Structure):
    _fields_ = [('handle', HANDLE),
                ('description', WCHAR * 128)]


MONITORS = []

def _enumerate_monitors():
    def callback(hmonitor, hdc, lprect, lparam):
        MONITORS.append(HMONITOR(hmonitor))
        return True

    if not windll.user32.EnumDisplayMonitors(None, None, _MONITORENUMPROC(callback), None):
        raise WinError('EnumDisplayMonitors failed')

def _iter_physical_monitors(close_handles=True):
    """Iterates physical monitors.

    The handles are closed automatically whenever the iterator is advanced.
    This means that the iterator should always be fully exhausted!

    If you want to keep handles e.g. because you need to store all of them and
    use them later, set `close_handles` to False and close them manually."""

    counter = 0
    for monitor in MONITORS:
        # Get physical monitor count
        count = DWORD()
        if not windll.dxva2.GetNumberOfPhysicalMonitorsFromHMONITOR(monitor, byref(count)):
            raise WinError()
        # Get physical monitor handles
        physical_array = (_PHYSICAL_MONITOR * count.value)()
        if not windll.dxva2.GetPhysicalMonitorsFromHMONITOR(monitor, count.value, physical_array):
            raise WinError()
        for physical in physical_array:
            print("Monitor {}: {}".format(counter, physical.description))
            yield physical.handle
            if close_handles:
                if not windll.dxva2.DestroyPhysicalMonitor(physical.handle):
                    raise WinError()
        counter += 1

def _get_monitor_by_index(monitor_index):
    counter = 0
    for monitor in MONITORS:
        # Get physical monitor count
        count = DWORD()
        if not windll.dxva2.GetNumberOfPhysicalMonitorsFromHMONITOR(monitor, byref(count)):
            raise WinError()
        # Get physical monitor handles
        physical_array = (_PHYSICAL_MONITOR * count.value)()
        if not windll.dxva2.GetPhysicalMonitorsFromHMONITOR(monitor, count.value, physical_array):
            raise WinError()

        for physical in physical_array:
            if counter == monitor_index:
                return physical.handle

        counter += 1


def set_vcp_feature(monitor, code, value):
    """Sends a DDC command to the specified monitor.
    """
    if not windll.dxva2.SetVCPFeature(HANDLE(monitor), BYTE(code), DWORD(value)):
        raise WinError()


# Switch to SOFT-OFF, wait for the user to press return and then back to ON
def toggle_off_on():
    for handle in _iter_physical_monitors():
        set_vcp_feature(handle, 0xd6, 0x04)
        input()
        set_vcp_feature(handle, 0xd6, 0x01)

def process_command(monitor_id, command, parameter):
    print("Executing Command on monitor {}".format(monitor_id))
    print("Command: {}, Parameter: {}".format(command, parameter))

    try:
        monitor_id_int = int(monitor_id)
        command_int = int(command, 16)
        parameter_int = int(parameter, 16)
    except ValueError:
        print("ERROR: Monitor ID must be an integer, commands and parameter must be hex codes.")

    set_vcp_feature(
        _get_monitor_by_index(monitor_id_int),
        command_int,
        parameter_int
    )

    print("Done")

def process_script(script_path):
    with open(script_path, mode='r') as script_file:
        commands = script_file.readlines()

    commands_to_process = []

    failed = False

    for command in commands:
        if command.startswith("#"):
            # It's a comment
            continue
        parts = command.split()
        if len(parts) == 3:
            command_dict = {
                "monitor_id": parts[0],
                "command": parts[1],
                "parameter": parts[2]
            }
            commands_to_process.append(command_dict)
        else:
            print("The following line has invalid syntax:")
            print(command)
            failed = True

    if not failed:
        for command in commands_to_process:
            process_command(
                command["monitor_id"],
                command["command"],
                command["parameter"]
            )

if __name__ == '__main__':
    _enumerate_monitors()
    print("Attached Monitors")
    for handle in _iter_physical_monitors(close_handles=False):
        pass
    if len(sys.argv) == 1:
        print(
            "Please provide a script path or monitor index, command and parameter"
        )
    elif len(sys.argv) == 2:
        process_script(sys.argv[1])
    elif len(sys.argv) == 4:
        process_command(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        print("Wrong parameters")
