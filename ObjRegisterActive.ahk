/*
    ObjRegisterActive(Object, CLSID, Flags:=0)
    
        Registers an object as the active object for a given class ID.
        Requires AutoHotkey v1.1.17+; may crash earlier versions.
    
    Object:
            Any AutoHotkey object.
    CLSID:
            A GUID or ProgID of your own making.
            Pass an empty string to revoke (unregister) the object.
    Flags:
            One of the following values:
              0 (ACTIVEOBJECT_STRONG)
              1 (ACTIVEOBJECT_WEAK)
            Defaults to 0.
    
    Related:
        http://goo.gl/KJS4Dp - RegisterActiveObject
        http://goo.gl/no6XAS - ProgID
        http://goo.gl/obfmDc - CreateGUID()
*/
RegisterIDs(CLSID, APPID)
{
	RegRead, OutputVar, HKCU, Software\Classes\%APPID%
    if (ErrorLevel)
    {
        RegWrite, REG_SZ, HKCU, Software\Classes\%APPID%,, %APPID%
        RegWrite, REG_SZ, HKCU, Software\Classes\%APPID%\CLSID,, %CLSID%
        RegWrite, REG_SZ, HKCU, Software\Classes\CLSID\%CLSID%,, %APPID%

        RegWrite, REG_SZ, HKLM, Software\Classes\%APPID%,, %APPID%
        RegWrite, REG_SZ, HKLM, Software\Classes\%APPID%\CLSID,, %CLSID%
        RegWrite, REG_SZ, HKLM, Software\Classes\CLSID\%CLSID%,, %APPID%
    }
}

RevokeIDs(CLSID, APPID)
{
	RegDelete, HKCU, Software\Classes\%APPID%
	RegDelete, HKCU, Software\Classes\CLSID\%CLSID%

	RegDelete, HKLM, Software\Classes\%APPID%
	RegDelete, HKLM, Software\Classes\CLSID\%CLSID%
}

Str2GUID(ByRef var, str)
{
	VarSetCapacity(var, 16)
	DllCall("ole32\CLSIDFromString", "wstr", str, "ptr", &var)
	return &var
}

ObjRegisterActive(Object, CLSID, Flags:=0, APPID:=0) {
    static cookieJar := {}
    if (!CLSID) {
        if (cookie := cookieJar.Remove(Object)) != ""
            DllCall("oleaut32\RevokeActiveObject", "uint", cookie, "ptr", 0)
        return
    }
    if cookieJar[Object]
        throw Exception("Object is already registered", -1)
    VarSetCapacity(_clsid, 16, 0)
    if ((hr := DllCall("ole32\CLSIDFromString", "wstr", CLSID, "ptr", &_clsid)) < 0)
        throw Exception("Invalid CLSID", -1, CLSID)
    hr := DllCall("oleaut32\RegisterActiveObject"
        , "ptr", &Object, "ptr", &_clsid, "uint", Flags, "uint*", cookie
        , "uint")
    if hr < 0
        throw Exception(format("Error 0x{:x}", hr), -1)
    cookieJar[Object] := cookie
    if (APPID){
        RegisterIDs(CLSID, APPID)
    }
}