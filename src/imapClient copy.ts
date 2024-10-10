import * as dotenv from 'dotenv';
dotenv.config();

import * as imaps from 'imap-simple';
import { parseHeader } from 'imap'; // Importa a função parseHeader
import Imap from 'imap';
import axios from 'axios';

const tenantId = process.env.TENANT_ID as string; // Altere para o seu Tenant ID
const clientId = process.env.CLIENT_ID as string; // Altere para o seu Client ID
const clientSecret = process.env.CLIENT_SECRET as string; // Altere para o seu Client Secret 
const scope = 'https://outlook.office365.com/.default'; // Escopo para acesso a e-mails
const grantType = 'client_credentials'; // Tipo de concessão

async function getAccessToken() {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append('grant_type', grantType);
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', scope);

    try {
        const response = await axios.post(tokenUrl, params);
        console.log({ response }); // Isso ajudará a verificar a resposta do servidor
        return response.data.access_token;
    } catch (error: any) {
        console.error('Error getting access token:', error.response.data);
        throw error;
    }
}

async function connectToIMAP() {
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0'; // Ignora erros de certificado
    const accessToken = await getAccessToken();

    const xoauth2Token = `user=${encodeURIComponent('testpai@plusoft.com')}\x01auth=Bearer ${accessToken}\x01\x01`;

    const imap = new Imap({
        user: 'testpai@plusoft.com', // Endereço de e-mail
        xoauth2: btoa(xoauth2Token), // Token em Base64
        host: 'outlook.office365.com', // Servidor IMAP
        port: 993, // Porta para SSL
        tls: true, // Habilita SSL/TLS
        authTimeout: 3000, // Tempo limite de autenticação
        password: ""
    });

    imap.once('ready', () => {
        imap.openBox('INBOX', false, (err: any, box: any) => {
            if (err) throw err;
            imap.search(['UNSEEN'], (err: any, results: any) => {
                if (err) throw err;

                const f = imap.fetch(results, { bodies: '' });
                f.on('message', (msg: any, seqno: any) => {
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

                f.once('end', () => {
                    console.log('Done fetching all messages!');
                    imap.end();
                });
            });
        });
    });

    imap.once('error', (err: any) => {
        console.error('IMAP Error:', err);
    });

    imap.connect();
}

connectToIMAP();