#Include OAuth2.ahk ; Include an OAuth 2.0 library for AutoHotkey
#Include Jxon.ahk
#Include Socket.ahk
#SingleInstance, Force
; You can find or write an OAuth 2.0 library that fits your needs.

if (!FileExist("token.json"))
{
    ; Load your credentials from the downloaded JSON file
    FileRead, json, % A_ScriptDir . "\credentials.json"
    credentials := Jxon_Load(json)
    
    ; Extract Client ID and Client Secret
    client_id := credentials.installed.client_id
    client_secret := credentials.installed.client_secret
    redirect_uri := credentials.installed.redirect_uris[1]
    ; MsgBox, Client ID: %client_id%`nClient Secret: %client_secret%`nRedirect URI: %redirect_uri%
    
    port := 8089
    ; Initialize the OAuth 2.0 flow
    oauth := new OAuth2(client_id, client_secret, redirect_uri . ":" . port, "https://accounts.google.com/o/oauth2/auth", "https://accounts.google.com/o/oauth2/token")
    
    ; Generate the authorization URL
    auth_url := oauth.GenerateAuthURL("https://www.googleapis.com/auth/calendar.readonly")
    Run, %auth_url%
    
    socket := CreateServerSocket(port)
    if (socket = -1)
    {
        MsgBox, Failed to create server socket on port %port%.
        ExitApp
    }
    
    Loop
    {
        clientSocket := DllCall("Ws2_32\accept", "UInt", socket, "Ptr", 0, "Ptr", 0)
        if (clientSocket = -1)
            continue
    
        VarSetCapacity(buffer, 4096, 0)
        bytesReceived := DllCall("Ws2_32\recv", "UInt", clientSocket, "Ptr", &buffer, "Int", 4096, "Int", 0)
        if (bytesReceived > 0)
        {
            text:=StrGet(&Buffer, bytesReceived, "UTF-8")
            code := ExtractCode(text)
            if (code)
            {
                response := "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`n`r`n<h1>Authorization successful! You can close this window.</h1>"
                DllCall("Ws2_32\send", "UInt", clientSocket, "Str", response, "Int", StrLen(response), "Int", 0)
                Global OAuthCode := code
                SetTimer, CloseSocket, -1000
                break
            }
        }
    }
    ; Sleep, 10000
    ; Exchange the authorization code for an access token
    token := oauth.ExchangeAuthCodeForToken(OAuthCode)
    if token{
        ; Save the token for future use
        if (FileExist("token.json"))
            FileDelete, token.json
        FileAppend, % token, token.json
        MsgBox, Authorization successful!
    } else {
        MsgBox, Could not authorize!
    }
}

; Read the token from the file
FileRead, token, token.json
token_obj:=Jxon_Load(token)

; Request Google Calendar API list of events for calendarId nhurst@ndscognitivelabs.com
; https://www.googleapis.com/calendar/v3/calendars/calendarId/events
; The access token must be included in the Authorization header
; The token is valid for 1 hour
; The token can be refreshed using the refresh token
; The refresh token is valid until the user revokes access to the application

; set timeMin to Today in yyyy-MM-ddTHH:mm:ssZ format
now:=A_Now
now+=6, Hours
timeMin := now
; timeMin += -1, Days
FormatTime, timeMin, %timeMin%, yyyy-MM-ddTHH:mm:ssZ 
; set timeMax to Tomorrow at the same time
timeMax := now
; timeMax += +1, Days
timeMax += +13, Hours
FormatTime, timeMax, %timeMax%, yyyy-MM-ddTHH:mm:ssZ
params:="?timeMin=" . timeMin . "&timeMax=" . timeMax . "&singleEvents=true" . "&timeZone=America/Mexico_City"
url := "https://www.googleapis.com/calendar/v3/calendars/nhurst%40ndscognitivelabs.com/events" . params
headers := {Authorization: "Bearer " . token_obj.access_token}
ComObjError(false)
; ComObjError(true)
http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
http.open("GET", url, true)
http.SetRequestHeader("Authorization", "Bearer " . token_obj.access_token)
; params := "{""timeMin"": """ . timeMin . """, ""timeMax"": """ . timeMax . """}"
; params := {"timeMin": timeMin, "timeMax": timeMax}
http.Send()
http.WaitForResponse()
response:=http.responseText
ComObjError(true)
response_obj:=Jxon_Load(response)
; filter out events by duration longer than 1 day
events:=response_obj.items
DateTimeFromIsoDate(isoDateTime){
    StringLeft, year, isoDateTime, 4
    StringMid, month, isoDateTime, 6, 2
    StringMid, day, isoDateTime, 9, 2
    StringMid, hour, isoDateTime, 12, 2
    StringMid, minute, isoDateTime, 15, 2
    StringMid, second, isoDateTime, 18, 2
    dateTimeObject := DateTimeFromParts(year, month, day, hour, minute, second)
    return dateTimeObject
}
; Function to create a DateTime object from parts
DateTimeFromParts(year, month, day, hour, minute, second)
{
    dateTime := year . month . day . hour . minute . second
    return dateTime
}
newEvents:=[]
Loop % events.MaxIndex(){
    if (events[A_Index].end.dateTime){
        start:=events[A_Index].start.dateTime
        end:=events[A_Index].end.dateTime
        start:=DateTimeFromIsoDate(start)
        end:=DateTimeFromIsoDate(end)
        duration:=var1
        EnvSub, duration, %start%, days
        if (duration<1)
            newEvents.Push(events[A_Index])
    }
}
response_obj.items:=newEvents
response:=Jxon_Dump(response_obj, 2)
if FileExist("response.json")
    FileDelete, response.json
FileAppend, % response, response.json
; Close the server socket
CloseSocket:
DllCall("Ws2_32\closesocket", "UInt", clientSocket)
DllCall("Ws2_32\WSACleanup")
return
