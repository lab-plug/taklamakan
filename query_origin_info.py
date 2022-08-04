import winreg
from ctypes import *

def get_origin_version():
    """Return the file version of Origin.Application
    
    >>> ver = get_origin_version()
    >>> len(ver.split('.')) == 3
    True
    """
    origin_version = None
    reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    with winreg.OpenKey(reg, r"SOFTWARE\Classes\Origin.Application\CLSID") as clsid_key:
        clsid = winreg.EnumValue(clsid_key, 0)[1]
        with winreg.OpenKey(reg, f"SOFTWARE\\Classes\\WOW6432Node\\CLSID\\{clsid}\\LocalServer32") as server32_key:
            path = winreg.EnumValue(server32_key, 0)[1]
            size = windll.version.GetFileVersionInfoSizeW(path, None)
            buffer = create_string_buffer(size)
            windll.version.GetFileVersionInfoW(path, None, size, buffer)
            value = c_void_p(0)
            value_size = c_uint(0)
            windll.version.VerQueryValueW(buffer, wstring_at(r'\StringFileInfo\040904b0\FileVersion'), byref(value), byref(value_size))
            origin_version = wstring_at(value.value, value_size.value - 1)
    return origin_version

if __name__ == '__main__':
    import doctest
    doctest.testmod()
