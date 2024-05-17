class OAuth2 {
    __New(client_id, client_secret, redirect_uri, auth_endpoint, token_endpoint, token_info_endpoint) {
        this.client_id := client_id
        this.client_secret := client_secret
        this.redirect_uri := redirect_uri
        this.auth_endpoint := auth_endpoint
        this.token_endpoint := token_endpoint
        this.token_info_endpoint := token_info_endpoint
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

    RefreshOAuth2Token(refresh_token) {
        ; global oauth2TokenEndpoint, client_id, client_secret, refresh_token
        this.refresh_token := refresh_token
        ; Prepare the POST request
        url := this.token_endpoint
        headers := "Content-Type: application/x-www-form-urlencoded"
        body := "grant_type=refresh_token&client_id=" . this.client_id . "&client_secret=" . this.client_secret . "&refresh_token=" . this.refresh_token
    
        ; Send the HTTP POST request
        http := ComObjCreate("MSXML2.ServerXMLHTTP.6.0")
        http.Open("POST", url, false)
        http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        http.Send(body)
    
        ; Check the response status
        status:=http.status
        statusText:=http.statusText
        if (status != 200) {
            MsgBox, Error refreshing token: %status% - %statusText%
            return
        }
    
        ; Parse the JSON response
        response := http.responseText
        return response
    }

    ValidateOAuth2Token(token) {
        ; global oauth2TokenInfoEndpoint
        ; Prepare the GET request
        url := this.token_info_endpoint . "?access_token=" . token
    
        ; Send the HTTP GET request
        http := ComObjCreate("MSXML2.ServerXMLHTTP.6.0")
        http.Open("GET", url, false)
        http.Send()
    
        ; Check the response status
        status:=http.status
        statusText:=http.statusText
        if (status != 200) {
            MsgBox, Error validating token: %status% - %statusText%
            return
        }
    
        response := http.responseText
        return response
    }
}

ExtractCode(buffer)
{
    if RegExMatch(buffer, "GET /.*[&?]code=([^& ]*)", m)
        return m1
    return ""
}
