import * as dotenv from 'dotenv';
dotenv.config();

import Imap, { parseHeader } from 'imap';
import axios from 'axios';

const tenantId = process.env.TENANT_ID as string;
const clientId = process.env.CLIENT_ID as string;
const clientSecret = process.env.CLIENT_SECRET as string;
const emailAddress = process.env.EMAIL as string;
const oauthScope = 'https://outlook.office365.com/.default';
const grantType = 'client_credentials';

const buildXOAuth2Token = (user = '', accessToken = '') => Buffer
    .from([`user=${user}`, `auth=Bearer ${accessToken}`, '', '']
        .join('\x01'), 'utf-8')
    .toString('base64');


async function retrieveAccessToken() {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append('grant_type', grantType);
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', oauthScope);

    try {
        const response = await axios.post(tokenUrl, params);
        console.log({ response });
        return response.data.access_token;
    } catch (error: any) {
        console.error('Error getting access token:', error.response.data);
        throw error;
    }
}

async function connectToImapServer() {
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
    const accessToken = await retrieveAccessToken();

    const xoauth2Token: string = buildXOAuth2Token(emailAddress, accessToken);

    const imapConfig = new Imap({
        xoauth2: xoauth2Token,
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        tlsOptions: {
            rejectUnauthorized: false,
            servername: 'outlook.office365.com'
        }
    } as any);

    imapConfig.once('ready', () => {
        imapConfig.openBox('INBOX', false, (err: any, box: any) => {
            if (err) throw err;
            imapConfig.search(['UNSEEN'], (err: any, results: any) => {
                if (err) throw err;

                const fetch = imapConfig.fetch(results, { bodies: '' });
                fetch.on('message', (msg: any, seqno: any) => {
                    console.log('Message #%d', seqno);
                    const prefix = '(#' + seqno + ') ';
                    msg.on('body', (stream: any, info: any) => {
                        let buffer = '';
                        stream.on('data', (chunk: any) => {
                            buffer += chunk.toString('utf8');
                        });
                        stream.once('end', () => {
                            console.log(prefix + 'Parsed header: %s', parseHeader(buffer));
                        });
                    });
                    msg.once('attributes', (attrs: any) => {
                        console.log(prefix + 'Attributes: %j', attrs);
                    });
                    msg.once('end', () => {
                        console.log(prefix + 'Finished');
                    });
                });

                fetch.once('end', () => {
                    console.log('Done fetching all messages!');
                    imapConfig.end();
                });
            });
        });
    });

    imapConfig.once('error', (err: any) => {
        console.error('IMAP Error:', err);
    });

    imapConfig.connect();
}

connectToImapServer();
