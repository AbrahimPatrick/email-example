import * as dotenv from 'dotenv';
dotenv.config();

import * as imaps from 'imap-simple';
import axios from 'axios';

const tenantId = process.env.TENANT_ID as string;
const clientId = process.env.CLIENT_ID as string;
const clientSecret = process.env.CLIENT_SECRET as string;
const email = process.env.EMAIL as string;
const scope = 'https://outlook.office365.com/.default';
const grantType = 'client_credentials';

async function getAccessToken() {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append('grant_type', grantType);
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', scope);

    try {
        const response = await axios.post(tokenUrl, params);
        console.log({ response })
        return response.data.access_token;
    } catch (error: any) {
        console.error('Error getting access token:', error.response.data);
        throw error;
    }
}

async function connectToIMAP() {
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
    const accessToken = await getAccessToken();

    const config = {
        imap: {
            user: email,
            password: accessToken,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            authTimeout: 3000,
        }
    };

    imaps.connect(config).then(connection => {
        return connection.openBox('INBOX').then(() => {
            return connection.search(['UNSEEN'], { bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)'], markSeen: false });
        });
    }).then(messages => {
        messages.forEach(message => {
            console.log(`Subject: ${message.parts[0].body.subject}`);
            console.log(`From: ${message.parts[0].body.from}`);
        });
    }).catch(err => {
        console.error('IMAP Error:', err);
    });
}

connectToIMAP();