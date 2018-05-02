//Install express server
const express = require('express');
var basicAuth = require('express-basic-auth');
const path = require('path');

const app = express();

app.use(basicAuth({
    users: { 'compro': 'c0mpr0' },
    challenge: true,
    unauthorizedResponse: getUnauthorizedResponse,
    realm: 'Leo Credential'
}))

function getUnauthorizedResponse(req) {
    return req.auth ?
        ('Credentials ' + req.auth.user + ':' + req.auth.password + ' rejected') :
        'No credentials provided'
}

// Serve only the static files form the dist directory
app.use(express.static(__dirname + '/dist'));

app.get('/*', function(req,res) {

res.sendFile(path.join(__dirname+'/dist/heroku.html'));
});

// Start the app by listening on the default Heroku port
app.listen(process.env.PORT || 8080);