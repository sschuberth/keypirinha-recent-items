from ctypes import *
from ctypes.wintypes import BOOL, DWORD, FILETIME, HWND, LPCWSTR, LPWSTR, UINT
from comtypes import CoCreateInstance, GUID, IUnknown, STDMETHOD
import sys

REFGUID = POINTER(GUID)
REFIID = REFGUID


class IObjectArray(IUnknown):
    _iid_ = GUID("{92ca9dcd-5622-4bba-a805-5e9f541bd8c9}")
    _methods_ = [
        STDMETHOD(HRESULT, "GetCount", [POINTER(UINT)]),
        STDMETHOD(HRESULT, "GetAt", [UINT, REFIID, POINTER(POINTER(IUnknown))])
    ]

    def GetCount(self):
        count = UINT()
        self.__com_GetCount(byref(count))
        return count.value

    def GetAt(self, index, interface = IUnknown):
        ptr = POINTER(interface)()
        self.__com_GetAt(index, interface._iid_, byref(ptr))
        return ptr


class IObjectCollection(IObjectArray):
    _iid_ = GUID("{5632b1a4-e38a-400a-928a-d4cd63230295}")
    _methods_ = [
        STDMETHOD(HRESULT, "AddObject", [POINTER(IUnknown)]),
        STDMETHOD(HRESULT, "AddFromArray", [POINTER(IObjectArray)]),
        STDMETHOD(HRESULT, "RemoveObjectAt", [UINT]),
        STDMETHOD(HRESULT, "Clear")
    ]


class IShellItem(IUnknown):
    _iid_ = GUID("{43826d1e-e718-42ee-bc55-a1e261c37bfe}")
    _methods_ = [
        STDMETHOD(HRESULT, "BindToHandler", [c_void_p, REFGUID, REFIID, POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "GetParent", [POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "GetDisplayName", [c_int, POINTER(LPWSTR)]),
        STDMETHOD(HRESULT, "GetAttributes", [c_int, POINTER(c_void_p)]),
        STDMETHOD(HRESULT, "Compare", [POINTER(IUnknown), c_int, POINTER(c_int)])
    ]

    SIGDN_FILESYSPATH = 0x80058000

    def GetDisplayName(self, form):
        name = LPWSTR()
        self.__com_GetDisplayName(form, byref(name))
        return name.value


class IAutomaticDestinationList8(IUnknown):
    _iid_ = GUID("{bc10dce3-62f2-4bc6-af37-db46ed7873c4}")
    _methods_ = [
        STDMETHOD(HRESULT, "Initialize", [LPCWSTR, LPCWSTR, LPCWSTR]),
        STDMETHOD(HRESULT, "HasList", [POINTER(BOOL)]),
        STDMETHOD(HRESULT, "GetList", [c_int, c_uint, REFIID, POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "AddUsagePoint", [POINTER(IUnknown)]),
        STDMETHOD(HRESULT, "PinItem", [POINTER(IUnknown), BOOL]),
        STDMETHOD(HRESULT, "IsPinned", [POINTER(IUnknown), POINTER(BOOL)]),
        STDMETHOD(HRESULT, "RemoveDestination", [POINTER(IUnknown)]),
        STDMETHOD(HRESULT, "SetUsageData", [POINTER(IUnknown), POINTER(c_float), POINTER(FILETIME)]),
        STDMETHOD(HRESULT, "GetUsageData", [POINTER(IUnknown), POINTER(c_float), POINTER(FILETIME)]),
        STDMETHOD(HRESULT, "ResolveDestination", [HWND, DWORD, c_void_p, REFIID, POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "ClearList", [c_int])
    ]

    def GetList(self, type, max_count):
        collection = POINTER(IObjectCollection)()
        self.__com_GetList(type, max_count, IObjectCollection._iid_, byref(collection))
        return collection


class IAutomaticDestinationList10(IUnknown):
    _iid_ = GUID("{e9c5ef8d-fd41-4f72-ba87-eb03bad5817c}")
    _methods_ = [
        STDMETHOD(HRESULT, "Initialize", [LPCWSTR, LPCWSTR, LPCWSTR]),
        STDMETHOD(HRESULT, "HasList", [POINTER(BOOL)]),
        STDMETHOD(HRESULT, "GetList", [c_int, c_uint, c_int, REFIID, POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "AddUsagePoint", [POINTER(IUnknown)]),
        STDMETHOD(HRESULT, "PinItem", [POINTER(IUnknown), BOOL]),
        STDMETHOD(HRESULT, "IsPinned", [POINTER(IUnknown), POINTER(BOOL)]),
        STDMETHOD(HRESULT, "RemoveDestination", [POINTER(IUnknown)]),
        STDMETHOD(HRESULT, "SetUsageData", [POINTER(IUnknown), POINTER(c_float), POINTER(FILETIME)]),
        STDMETHOD(HRESULT, "GetUsageData", [POINTER(IUnknown), POINTER(c_float), POINTER(FILETIME)]),
        STDMETHOD(HRESULT, "ResolveDestination", [HWND, DWORD, c_void_p, REFIID, POINTER(POINTER(IUnknown))]),
        STDMETHOD(HRESULT, "ClearList", [c_int])
    ]

    FLAG_EXCLUDE_UNNAMED_DESTINATIONS = 1

    def GetList(self, type, max_count, flags):
        collection = POINTER(IObjectCollection)()
        self.__com_GetList(type, max_count, flags, IObjectCollection._iid_, byref(collection))
        return collection


class AutomaticDestinationList:
    CLSID_AutomaticDestinationList = GUID("{f0ae1542-f497-484b-a175-a20db09144ba}")

    def __init__(self, appid):
        if sys.getwindowsversion().major < 10:
            self.list = CoCreateInstance(self.CLSID_AutomaticDestinationList, IAutomaticDestinationList8)
        else:
            self.list = CoCreateInstance(self.CLSID_AutomaticDestinationList, IAutomaticDestinationList10)

        self.list.Initialize(appid, None, None)

    TYPE_PINNED = 0
    TYPE_RECENT = 1
    TYPE_FREQUENT = 2

    def GetList(self, list_type, max_count):
        if sys.getwindowsversion().major < 10:
            collection = self.list.GetList(list_type, max_count)
        else:
            collection = self.list.GetList(list_type, max_count, IAutomaticDestinationList10.FLAG_EXCLUDE_UNNAMED_DESTINATIONS)

        return collection
