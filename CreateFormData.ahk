; CreateFormData() by tmplinshi, AHK Topic: https://autohotkey.com/boards/viewtopic.php?t=7647
; Thanks to Coco: https://autohotkey.com/boards/viewtopic.php?p=41731#p41731
; Modified version by SKAN, 09/May/2016
; Rewritten by iseahound in September 2022

CreateFormData(ByRef retData, ByRef retHeader, objParam, noInit := 0) {
    New CreateFormData(retData, retHeader, objParam, noInit)
}

Class CreateFormData {

    __New(ByRef retData, ByRef retHeader, objParam, noInit := 0) {

        Local CRLF := "`r`n", i, k, v, str, pvData
        ; Create a random Boundary
        Local Boundary := this.RandomBoundary()
        Local BoundaryLine := "------------------------------" . Boundary

        ; Create an IStream backed with movable memory.
        hData := DllCall("GlobalAlloc", "uint", 0x2, "uptr", 0, "ptr")
        DllCall("ole32\CreateStreamOnHGlobal", "ptr", hData, "int", False, "ptr*", pStream:=0, "uint")
        this.pStream := pStream

        ; Loop input paramters
        For k, v in objParam
        {
            If IsObject(v) {
                For i, FileName in v
                {
                    str := BoundaryLine . CRLF
                        . "Content-Disposition: form-data; name=""" . k . """; filename=""" . FileName . """" . CRLF
                        . "Content-Type: " . this.MimeType(FileName) . CRLF . CRLF

                    if (!noInit)
                        this.StrPutUTF8( str )
                    this.LoadFromFile( Filename )
                    if (!noInit)
                        this.StrPutUTF8( CRLF )
                }
            } Else {
                str := BoundaryLine . CRLF
                    . "Content-Disposition: form-data; name=""" . k """" . CRLF . CRLF
                    . v . CRLF
                this.StrPutUTF8( str )
            }
        }

        this.StrPutUTF8( BoundaryLine . "--" . CRLF )

        this.pStream := ObjRelease(pStream) ; Should be 0.
        pData := DllCall("GlobalLock", "ptr", hData, "ptr")
        size := DllCall("GlobalSize", "ptr", pData, "uptr")

        ; Create a bytearray and copy data in to it.
        retData := ComObjArray( 0x11, size ) ; Create SAFEARRAY = VT_ARRAY|VT_UI1
        pvData  := NumGet( ComObjValue( retData ), 8 + A_PtrSize , "ptr" )
        DllCall( "RtlMoveMemory", "Ptr", pvData, "Ptr", pData, "Ptr", size )

        DllCall("GlobalUnlock", "ptr", hData)
        DllCall("GlobalFree", "Ptr", hData, "Ptr")                   ; free global memory

        retHeader := "multipart/form-data; boundary=----------------------------" . Boundary
    }

    StrPutUTF8( str ) {
        length := StrPut(str, "UTF-8") - 1 ; remove null terminator
        VarSetCapacity(utf8, length)
        StrPut(str, &utf8, length, "UTF-8")
        DllCall("shlwapi\IStream_Write", "ptr", this.pStream, "ptr", &utf8, "uint", length, "uint")
    }

    LoadFromFile( filepath ) {
        DllCall("shlwapi\SHCreateStreamOnFileEx"
                    ,   "wstr", filepath
                    ,   "uint", 0x0             ; STGM_READ
                    ,   "uint", 0x80            ; FILE_ATTRIBUTE_NORMAL
                    ,    "int", False            ; fCreate is ignored when STGM_CREATE is set.
                    ,    "ptr", 0               ; pstmTemplate (reserved)
                    ,   "ptr*", pFileStream:=0
                    ,   "uint")
        DllCall("shlwapi\IStream_Size", "ptr", pFileStream, "uint64*", size:=0, "uint")
        DllCall("shlwapi\IStream_Copy", "ptr", pFileStream , "ptr", this.pStream, "uint", size, "uint")
        ObjRelease(pFileStream)
    }

    RandomBoundary() {
        str := "0|1|2|3|4|5|6|7|8|9|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z"
        Sort, str, D| Random
        str := StrReplace(str, "|")
        Return SubStr(str, 1, 12)
    }

    MimeType(FileName) {
        n := FileOpen(FileName, "r").ReadUInt()
        Return (n        = 0x474E5089) ? "image/png"
            :  (n        = 0x38464947) ? "image/gif"
            :  (n&0xFFFF = 0x4D42    ) ? "image/bmp"
            :  (n&0xFFFF = 0xD8FF    ) ? "image/jpeg"
            :  (n&0xFFFF = 0x4949    ) ? "image/tiff"
            :  (n&0xFFFF = 0x4D4D    ) ? "image/tiff"
            :  "application/octet-stream"
    }

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Working example
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


; whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
; whr.Open("POST", "https://file.io", true)

; objParam := { maxDownloads  : "1"
;             , autoDelete    : "true"
;             , file          : ["a.jpg"] } ; Change filepath with an existing image.

; CreateFormData(PostData, hdr_ContentType, objParam)

; whr.SetRequestHeader("Content-Type", hdr_ContentType)
; whr.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko")
; whr.Option(6) := False ; No auto redirect
; whr.Send(PostData)
; whr.WaitForResponse()
; MsgBox % whr.ResponseText
; Run % RegExReplace(whr.ResponseText, "^.*?(https://file.io/.+?)" Chr(34) ".*$", "$1")
