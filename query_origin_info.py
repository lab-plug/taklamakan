import os
import winreg
from ctypes import *

def get_origin_version():
    """Return the file version of Origin according to `origin8.tlb`.
    
    >>> ver = get_origin_version()
    >>> len(ver.split('.')) == 3
    True
    """
    origin_version = None
    reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    with winreg.OpenKey(reg, r"SOFTWARE\Classes\TypeLib\{DBC515E6-9735-4D78-A75C-3DE67DF252D0}\8.0\0\win64") as tlb_key:
        tlb = winreg.EnumValue(tlb_key, 0)[1]
        tlb_dir = tlb + '\\..'
        exe = None
        for path in os.listdir(tlb_dir):
            if os.path.isfile(os.path.join(tlb_dir, path)):
                if path.lower().startswith("origin") and path.lower().endswith(".exe"):
                    exe = path
                    break
        if not exe:
            return

        exe_path = tlb + "\\..\\" + exe
        size = windll.version.GetFileVersionInfoSizeW(exe_path, None)
        buffer = create_string_buffer(size)
        windll.version.GetFileVersionInfoW(exe_path, None, size, buffer)
        value = c_void_p(0)
        value_size = c_uint(0)
        windll.version.VerQueryValueW(buffer, wstring_at(r'\StringFileInfo\040904b0\FileVersion'), byref(value),
                                      byref(value_size))
        origin_version = wstring_at(value.value, value_size.value - 1)
    return origin_version

if __name__ == '__main__':
    import doctest
    doctest.testmod()
