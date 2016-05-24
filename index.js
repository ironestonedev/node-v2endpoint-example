var express = require('express');
var session = require('express-session')
var request = require('request');

var app = express();
app.use(session({
    secret: 'keyboard cat',
    cookie: {}
}))

// Define the various URIs we need
var auth_uri = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
var token_uri = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
var redirect_uri = 'http://localhost:3000/returned';

// Define the client id, secret and scope
var client_id = 'fc9a6fcf-bcd0-451c-9641-12ca455b9a34';
var client_secret = 'tXVebOX8jdhgv5i6HN9Ep9w';

// Define scopes
// NOTE: You must request offline_access in order to recieve a refresh_token. Without it the
// the autorization will only live for a limited time (typically 1 hour).
var client_scopes = 'https://graph.microsoft.com/User.Read offline_access openid';


// Forms
// There are two forms we post to the v2 endpoint, one to request the initial token and another 
// to request a new token using a previous refresh_token
var token_request = {
    form: {
        grant_type: 'authorization_code',
        code: '', // We set this at runtime
        client_id: client_id,
        client_secret: client_secret,
        scope: client_scopes,
        redirect_uri: redirect_uri
    }
}

var refresh_token_request = {
    form: {
        grant_type: 'refresh_token',
        refresh_token: '', // We set this at runtime
        client_id: client_id,
        client_secret: client_secret,
        scope: client_scopes,
        redirect_uri: redirect_uri
    }
}

// This is the web root and provides a simple "Click to Start" link that kicks off the OAUTH process 
app.get('/', function (req, res) {
    var oauth_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=code&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    res.send('<a href="' + oauth_endpoint + '" target="_blank">Click to Start</a>');
});

// This is the return point after we have executed the first OAUTH step.
// Here we convert the auth code returned by the endpoint into a bearer token 
// we can use to call the API. We also return a link to refresh this token 
// when it expires.
app.get('/returned', function (req, res) {

    // Grab the auth code from the query params
    var auth_code = req.query.code;

    // Add this auto_code to our token_request
    token_request.form.code = auth_code;

    // Post this form to the v2 Endpoint and display the result in the browser    
    request.post(token_uri, token_request, function (err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + auth_code + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)
    })
});

// This is where we convert the refresh token into a usable 
// brearer token we can use for API calls. 
app.get('/refresh', function (req, res) {

    // Grab the refresh_token from the query params
    var refresh_token = req.query.code;

    // Add this auto_code to our refresh_token_request
    refresh_token_request.form.refresh_token = refresh_token;

    // Post this form to the v2 Endpoint and display the result in the browser
    request.post(token_uri, refresh_token_request, function (err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + refresh_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)
    })
});

// Start listening on port 3000
app.listen(3000, function () {
    console.log('Example app listening on port 3000!');
});