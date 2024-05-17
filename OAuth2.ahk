class OAuth2 {
    __New(client_id, client_secret, redirect_uri, auth_endpoint, token_endpoint) {
        this.client_id := client_id
        this.client_secret := client_secret
        this.redirect_uri := redirect_uri
        this.auth_endpoint := auth_endpoint
        this.token_endpoint := token_endpoint
    }

    GenerateAuthURL(scope) {
        params := "client_id=" . this.client_id
        params .= "&redirect_uri=" . this.redirect_uri
        params .= "&response_type=code"
        params .= "&scope=" . scope
        return this.auth_endpoint . "?" . params
    }

    ExchangeAuthCodeForToken(auth_code) {
        params := "code=" . auth_code
        params .= "&client_id=" . this.client_id
        params .= "&client_secret=" . this.client_secret
        params .= "&redirect_uri=" . this.redirect_uri
        params .= "&grant_type=authorization_code"
        
        http := ComObjCreate("MSXML2.XMLHTTP.6.0")
        http.Open("POST", this.token_endpoint, false)
        http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        http.Send(params)
        status:=http.Status
        if(http.Status != 200){
            MsgBox, Failed to retrieve access token. Status: %status%
            return
        }
        
        response := http.responseText
        return response
        ; token_data := {}
        ; Loop, Parse, response, `n, `r 
        ; {
        ;     StringSplit, pair, A_LoopField, `:, %A_Space%
        ;     pair1:=Trim(pair1, """")
        ;     pair2:=Trim(pair2, """")
        ;     token_data[pair1] := pair2
        ; }
        ; return token_data
    }
}

ExtractCode(buffer)
{
    if RegExMatch(buffer, "GET /.*[&?]code=([^& ]*)", m)
        return m1
    return ""
}
