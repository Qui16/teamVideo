const express = require('express');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require("@azure/identity");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");

const app = express();

// Replace with your Microsoft Graph API credentials
const clientId = '61af1d10-38df-4b9b-b4b4-3963cb53c547';
const clientSecret = 'j758Q~T_AceWIbNTGZvv7RqHoaK8F.6Xo9~4hcZ8';
const tenantId = '5c0d2688-aa31-4ac3-bd1c-f31c96bca508';
const userId = '9fd3f7d6-7b3c-4a07-96cc-9006192a6bf8';

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
// Function to create Teams meeting and get join link
async function createTeamsMeeting() {
    const client = new Client({
        authProvider: {
            getAccessToken: async () => {
                const token = await credential.getToken('https://graph.microsoft.com/.default');
                return token.token;
            },
        },
    });
    const credential = new ClientSecretCredential(
        tenantId,
        clientId,
        clientSecret,
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        // The client credentials flow requires that you request the
        // /.default scope, and pre-configure your permissions on the
        // app registration in Azure. An administrator must grant consent
        // to those permissions beforehand.
        scopes: ['https://graph.microsoft.com/.default'],
    });

    const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

    const meeting = {
        subject: 'Meeting from React App', // Replace with desired subject
        startDateTime: new Date(Date.now() + 10000).toISOString(), // Start after 10 seconds
        joinMeetingIdSettings: {
            isPasscodeRequired: false
        }
    };

    try {
        const createdMeeting = await graphClient.api(`/users/${userId}/onlineMeetings`).post(meeting);
        return createdMeeting.joinUrl;
    } catch (error) {
        console.error('Error creating Teams meeting:', error);
        return null;
    }
}

// Route to handle call button click
app.get('/getTeamsLink', async (req, res) => {
    const teamsLink = await createTeamsMeeting();
    if (teamsLink) {
        res.json({ teamsLink });
    } else {
        res.status(500).send('Error creating Teams meeting');
    }
});

// const fetchToken = async () => {
//     const authUrl = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?' +
//         'client_id={client_id}' +
//         '&response_type=code' +
//         '&redirect_uri={redirect_uri}' +
//         '&response_mode=query' +
//         '&scope=https%3A%2F%2Fgraph.microsoft.com%2FCommunications.OnlineMeetings.ReadWrite' +
//         '&state=12345';
//     // Redirect the user to the authorize endpoint
//     window.location.href = authUrl;

//     const tokenUrl = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token';

//     const data = {
//         client_id: '{client_id}',
//         scope: 'https://graph.microsoft.com/Communications.OnlineMeetings.ReadWrite',
//         code: '{authorization_code}',  // Replace with the authorization code from the query string
//         redirect_uri: '{redirect_uri}',
//         grant_type: 'authorization_code',
//         client_secret: '{client_secret}'  // Replace with your client secret
//     };

//     try {
//         const response = await fetch(tokenUrl, {
//             method: 'POST',
//             headers: {
//                 'Content-Type': 'application/x-www-form-urlencoded',
//             },
//             body: new URLSearchParams(data),
//         });
//         const token = await response.json();
//         return token.access_token;

//     } catch (error) {
//         console.error('Error getting token:', error);
//     }
// };

// app.get('/getToken', (req, res) => {
//     const token = fetchToken();
//     if (token) {
//         res.json({ token });
//     } else {
//         res.status(500).send('Error getting token');
//     }

// });

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server listening on port ${port}`));