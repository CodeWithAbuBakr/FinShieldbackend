const express = require('express');
const app = express();
const PORT = 4000;
const cors = require('cors');
const { AuthorizationCode } = require('simple-oauth2');
const fetch = require('node-fetch');
const dotenv = require('dotenv');
dotenv.config();
// Example for Node.js backend
app.use(cors({
  origin: 'https://localhost:3000', // or '*' for development
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

// Middleware to parse JSON requests
app.use(express.json());

// Basic route
app.get('/', (req, res) => {
  res.send('Welcome to the Node.js Backend!');
});


const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const redirectUri = process.env.REDIRECT_URI;

const oauth2Client = new AuthorizationCode({
  client: { id: clientId, secret: clientSecret },
  auth: {
    tokenHost: 'https://login.microsoftonline.com/common',
    authorizePath: '/oauth2/v2.0/authorize',
    tokenPath: '/oauth2/v2.0/token',
  },
});

// Simple in-memory token storage (replace with DB for production)
let tokenStore = null;

// Step 1: Redirect user to Microsoft login
app.get('/auth/login', (req, res) => {
  const authorizationUri = oauth2Client.authorizeURL({
    redirect_uri: redirectUri,
    scope: 'https://graph.microsoft.com/Mail.ReadWrite offline_access',
    response_mode: 'query',
  });
  res.redirect(authorizationUri);
});

// Step 2: OAuth2 callback - exchange code for tokens
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  const options = { code, redirect_uri: redirectUri, scope: 'https://graph.microsoft.com/Mail.ReadWrite offline_access' };

  try {
    const accessToken = await oauth2Client.getToken(options);
    tokenStore = oauth2Client.createToken(accessToken.token); // store token object
    console.log('Access token acquired');
    res.send('Auth successful! You can close this tab.');
  } catch (error) {
    console.error('Access Token Error', error.message);
    res.status(500).send('Authentication failed');
  }
});

// Helper: Refresh token if expired
async function getValidAccessToken() {
  if (!tokenStore) throw new Error('No token stored, please authenticate first');
  if (tokenStore.expired()) {
    console.log('Access token expired, refreshing...');
    tokenStore = await tokenStore.refresh();
    console.log('Token refreshed');
  }
  return tokenStore.token.access_token;
}

// Step 3: Get mails with Red category
app.get('/mails/red', async (req, res) => {
  try {
    const accessToken = await getValidAccessToken();
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=categories/any(c:c eq 'Red')&$top=50`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const data = await response.json();
    res.json(data.value || []);
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// Step 4: Delete a mail by id
app.delete('/mails/:id', async (req, res) => {
  try {
    const messageId = req.params.id;
    const accessToken = await getValidAccessToken();
    const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}`;
    const response = await fetch(url, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (response.status === 204) {
      res.send('Mail deleted successfully');
    } else {
      const errorText = await response.text();
      res.status(response.status).send(errorText);
    }
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// Step 5: Delete all Red category mails endpoint
app.delete('/mails/red/delete-all', async (req, res) => {
  try {
    const accessToken = await getValidAccessToken();
    const getUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=categories/any(c:c eq 'Red')&$top=50`;

    const response = await fetch(getUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const data = await response.json();
    const mails = data.value || [];

    for (const mail of mails) {
      const delUrl = `https://graph.microsoft.com/v1.0/me/messages/${mail.id}`;
      await fetch(delUrl, {
        method: 'DELETE',
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      console.log(`Deleted mail: ${mail.id}`);
    }
    res.send(`Deleted ${mails.length} Red category mails`);
  } catch (err) {
    res.status(500).send(err.message);
  }
});



// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
