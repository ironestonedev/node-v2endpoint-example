// Add your app's registration data here
const client_id = '30004c22-c36d-492a-8835-24c3e584d692';
const client_secret = 'ZuhrNN3xYHM3arBeh1ookTk';

const graph = require("@microsoft/microsoft-graph-client");
const express = require('express');
const session = require('express-session')
const request = require('request');
const https = require('https');
const http = require('http');

var app = express();
app.use('/static', express.static('public'));
app.use(session({
    secret: 'keyboard cat',
    cookie: {}
}))

// Define the various URIs we need
var auth_uri = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
var token_uri = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
var admin_uri = 'https://login.microsoftonline.com/common/adminconsent';
var redirect_uri = 'http://localhost:3000/returned';
var admin_redirect_uri = 'http://localhost:3000/consentReturn';

// Define scopes
// NOTE: You must request offline_access in order to recieve a refresh_token. Without it the
// the autorization will only live for a limited time (typically 1 hour).
var client_scopes = 'User.Read Mail.Read User.Read.All offline_access';

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

// This is the web root and provides a link that kicks off the OAUTH process 
app.get('/', function(req, res) {
    var codegrant_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=code&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    var implicit_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=token&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    var admin_endpoint = admin_uri + '?client_id=' + client_id + '&admin_redirect_uri=' + redirect_uri;
    res.send('<div><a href="' + codegrant_endpoint + '" target="_blank">Code Grant Workflow</a></div><div><a href="' + implicit_endpoint + '" target="_blank">Implicit Grant Workflow</a></div><div><a href="' + admin_endpoint + '" target="_blank">Admin Consent Workflow</a></div>');
});


app.get('/implicit', function(req, res) {
    // Grab the token from the query params
    var access_token = req.query.access_token;
    var token_type = req.query.token_type;
    var expires_in = req.query.expires_in;
    var scope = req.query.scope;



    // Add this auto_code to our token_request
    token_request.form.code = auth_code;

    // Post this form to the v2 Endpoint and display the result in the browser    
    request.post(token_uri, token_request, function(err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + auth_code + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)
    })


});

// This is the return point after we have executed the first OAUTH step.
// Here we convert the auth code returned by the endpoint into a bearer token 
// we can use to call the API. We also return a link to refresh this token 
// when it expires.
app.get('/returned', function(req, res) {
    if(req.query.error != null)
    {        
        var content = "<p>";
        content += req.query.error;
        content += "</p>";
        content += "<p>";
        content += req.query.error_description;
        content += "</p>";
        res.end(content)
    }
    else if (req.query.code != null) {
        // This is an OAUTH Code Grant workflow

        // Grab the auth code from the query params
        var auth_code = req.query.code;

        // Add this auto_code to our token_request
        token_request.form.code = auth_code;

        // Post this form to the v2 Endpoint and display the result in the browser    
        request.post(token_uri, token_request, function(err, httpResponse, body) {
            var result = JSON.parse(body);

            var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
            content += "</p>";

            getProfile(result.access_token, function(profile) {
                content += "<pre>" + JSON.stringify(profile, null, 2) + "</pre>";
                content += "</p>";

                content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>';
                res.end(content)
            });
        })
    } else {
        // This is an OAUTH Implicit Grant workflow

        var token = {
            access_token: req.query.access_token,
            token_type: req.query.token_type,
            expires_in: req.query.expires_in,
            scope: req.query.scope
        }

        var content = "<pre>" + JSON.stringify(token, null, 2) + "</pre>"
        content += "<pre>" + token.access_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)

    }
});

// This is where we convert the refresh token into a usable 
// brearer token we can use for API calls. 
app.get('/refresh', function(req, res) {

    // Grab the refresh_token from the query params
    var refresh_token = req.query.code;

    // Add this auto_code to our refresh_token_request
    refresh_token_request.form.refresh_token = refresh_token;

    // Post this form to the v2 Endpoint and display the result in the browser
    request.post(token_uri, refresh_token_request, function(err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + refresh_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)
    })
});

// This is where we convert the refresh token into a usable 
// brearer token we can use for API calls. 
app.get('/consentReturn', function(req, res) {

    if(req.query.error != null)
    {        
        var content = "<p>";
        content += req.query.error;
        content += "</p>";
        content += "<p>";
        content += req.query.error_description;
        content += "</p>";
        res.end(content)
    }

    var codegrant_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=code&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    var implicit_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=token&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    var admin_endpoint = admin_uri + '?client_id=' + client_id + '&redirect_uri=' + redirect_uri + '&scope=' + client_scopes;
    res.send('<div>Consent Completed</div><div><a href="' + codegrant_endpoint + '" target="_blank">Code Grant Workflow</a></div><div><a href="' + implicit_endpoint + '" target="_blank">Implicit Grant Workflow</a></div><div><a href="' + admin_endpoint + '" target="_blank">Admin Consent Workflow</a></div>');
});

function getProfile(access_token, callback) {
    var client = graph.Client.init({
        authProvider: (done) => {
            done(null, access_token); //first parameter takes an error if you can't get an access token
        }
    });
    client
        .api('/me')
        .get((err, res) => {
            callback(res);
        });
}

http.createServer(app).listen(3000);

console.log('Example app listening on port 3000!');