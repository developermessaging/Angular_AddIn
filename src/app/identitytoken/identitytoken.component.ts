var _xhr:any;

var msexchuid:string;
var amurl:string;
var uniqueID:string;
var aud:string;
var iss:string;
var x5t:string;
var nbf:string;
var exp:string;
var rsp:string;
var tok:string;
var result:string;
var status:string;

this.rsp = _xhr.responseText;

function getUserIdentityTokenCallback(asyncResult) {
    this.status= "getUserIdentityTokenCallback in progress";
    var token = asyncResult.value;
    
    if (asyncResult.status === "succeeded") {
        _xhr = new XMLHttpRequest();
        _xhr.open("POST", "https://dsmsgeu-identitytokenservice.azurewebsites.net/api/IdentityToken/");
        _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        _xhr.onreadystatechange = readyStateChange;

        var request:any = new Object();
        request.token = token;
        this.tok=JSON.stringify(request);
        this.result = asyncResult.status;
        _xhr.send(JSON.stringify(request));
    }
    else { this.result = "Failed: " + asyncResult.error.errorMessage; }
    this.status ="getUserIdentityTokenCallback in complete";
}

function readyStateChange() {
    if (_xhr.readyState == 4 && _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.errorMessage) {
            this.msexchuid = response.token.msexchuid;
            this.amurl = response.token.amurl;
            this.uniqueID = response.token.uniqueID;
            this.aud = response.token.aud;
            this.iss = response.token.iss;
            this.x5t = response.token.x5t;
            this.nbf = response.token.nbf;
            this.exp = response.token.exp;

            this.rsp = _xhr.responseText;
            this.error = "none";
        }
        else {
            this.error = response.error;

            //app.showNotification("Error!", response.errorMessage);
        }
    }
}