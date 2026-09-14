import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { InjectModel } from '@nestjs/mongoose';
import { UnipileAccount, UnipileAccountDocument } from './schema/unipile-account.schema';
import { Model } from 'mongoose';
import { UnipileInbox, UnipileInboxDocument } from './schema/unipile-inbox.schema';
import { UnipileOutbox, UnipileOutboxDocument } from './schema/unipile-outbox.schema';

@Injectable()
export class AuthService {
    constructor(

        @InjectModel(UnipileAccount.name)
        private readonly unipileAccountModel: Model<UnipileAccountDocument>,

        @InjectModel(UnipileInbox.name)
        private readonly unipileInboxModel: Model<UnipileInboxDocument>,

        @InjectModel(UnipileOutbox.name)
        private readonly unipileOutboxModel: Model<UnipileOutboxDocument>,

        private readonly configService: ConfigService,
    ) { }

    //Generate unipile google auth link
    async createUnipileGoogleAuthLink() {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');
        const redirectUri = this.configService.get<string>('unipile.redirectUri');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        if (!redirectUri) {
            throw new Error('UNIPILE_REDIRECT_URI is not configured');
        }

        const expiresOn = new Date(Date.now() + 10 * 60 * 1000).toISOString();

        const response = await fetch(`https://${dsn}/api/v1/hosted/accounts/link`, {
            method: 'POST',

            headers: {
                'X-API-KEY': apiKey,
                Accept: 'application/json',
                'Content-Type': 'application/json',
            },

            body: JSON.stringify({
                type: 'create',
                providers: ['GOOGLE'],
                api_url: `https://${dsn}`,
                expiresOn,
                success_redirect_url: redirectUri,
                failure_redirect_url: redirectUri,
            }),
        });

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile Hosted Auth request failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        return data.url;
    }

    // Generate Unipile Microsoft auth link
    async createUnipileMicrosoftAuthLink() {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');
        const redirectUri = this.configService.get<string>('unipile.redirectUri');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        if (!redirectUri) {
            throw new Error('UNIPILE_REDIRECT_URI is not configured');
        }

        const expiresOn = new Date(
            Date.now() + 10 * 60 * 1000,
        ).toISOString();

        const response = await fetch(
            `https://${dsn}/api/v1/hosted/accounts/link`,
            {
                method: 'POST',

                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                    'Content-Type': 'application/json',
                },

                body: JSON.stringify({
                    type: 'create',
                    providers: ['OUTLOOK'],
                    api_url: `https://${dsn}`,
                    expiresOn,
                    success_redirect_url: redirectUri,
                    failure_redirect_url: redirectUri,
                }),
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile Microsoft Hosted Auth request failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        return data.url;
    }



    //Save Both Microsoft and Google Unipile Account
    // Save Unipile account ID to database
    async saveUnipileAccount(
        accountId: string,
        email: string,
        connectionType: 'GOOGLE' | 'MICROSOFT',
        name?: string,
        provider?: string,
    ) {
        return this.unipileAccountModel.findOneAndUpdate(
            { accountId },
            {
                accountId,
                email,
                connectionType,
                name,
                provider,
            },
            {
                new: true,
                upsert: true,
                setDefaultsOnInsert: true,
            },
        );
    }

    //Fetch Unipile Inbox Messages
    async fetchUnipileInbox(accountId: string) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const account = await this.unipileAccountModel.findOne({
            accountId,
        });

        if (!account) {
            throw new Error('Unipile account not found');
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails?limit=50&account_id=${encodeURIComponent(accountId)}`,
            {
                method: 'GET',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                },
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile email request failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        const messages = (data.items || []).map((email: any) => ({
            messageId: email.id || '',
            providerMessageId: email.provider_id || '',

            threadId: email.thread_id || '',
            subject: email.subject || '',

            senderEmail: email.from_attendee?.identifier || '',
            senderName: email.from_attendee?.display_name || '',

            message: email.body_plain || email.body || '',

            receivedAt: email.date
                ? new Date(email.date)
                : new Date(),

            attachments: (email.attachments || []).map(
                (attachment: any) => ({
                    filename: attachment.name || '',
                    mimeType:
                        attachment.mime ||
                        attachment.content_type ||
                        '',
                    attachmentId: attachment.id || '',
                }),
            ),
        }));
        const inbox =
            await this.unipileInboxModel.findOneAndUpdate(
                { accountId },
                {
                    accountId,
                    email: account.email,
                    connectionType: account.connectionType,
                    messages,
                },
                {
                    new: true,
                    upsert: true,
                    setDefaultsOnInsert: true,
                },
            );

        console.log(
            `Fetched and saved ${messages.length} Unipile messages`,
        );

        return inbox;
    }

    //Fetch sent Messages

    async fetchUnipileOutbox(accountId: string) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const account = await this.unipileAccountModel.findOne({
            accountId,
        });

        if (!account) {
            throw new Error('Unipile account not found');
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails?limit=50&account_id=${encodeURIComponent(accountId)}&folder=SENT`,
            {
                method: 'GET',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                },
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile sent email request failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        const messages = (data.items || []).map((email: any) => ({
            messageId: email.id || '',
            providerMessageId: email.provider_id || '',

            threadId: email.thread_id || '',
            subject: email.subject || '',

            receiverEmail:
                email.to_attendees?.[0]?.identifier || '',

            receiverName:
                email.to_attendees?.[0]?.display_name || '',

            message: email.body_plain || email.body || '',

            sentAt: email.date
                ? new Date(email.date)
                : new Date(),

            attachments: (email.attachments || []).map(
                (attachment: any) => ({
                    filename: attachment.name || '',
                    mimeType:
                        attachment.mime ||
                        attachment.content_type ||
                        '',
                    attachmentId: attachment.id || '',
                }),
            ),
        }));

        const outbox =
            await this.unipileOutboxModel.findOneAndUpdate(
                { accountId },
                {
                    accountId,
                    email: account.email,
                    connectionType: account.connectionType,
                    messages,
                },
                {
                    new: true,
                    upsert: true,
                    setDefaultsOnInsert: true,
                },
            );

        console.log(
            `Fetched and saved ${messages.length} Unipile sent messages`,
        );

        return outbox;
    }

    async syncUnipileInbox(accountId: string) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const account = await this.unipileAccountModel.findOne({
            accountId,
        });

        if (!account) {
            throw new Error('Unipile account not found');
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails?limit=50&account_id=${encodeURIComponent(accountId)}`,
            {
                method: 'GET',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                },
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile inbox sync failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        const messages = (data.items || []).map((email: any) => ({
            messageId: email.id || '',
            providerMessageId: email.provider_id || '',

            threadId: email.thread_id || '',
            subject: email.subject || '',

            senderEmail: email.from_attendee?.identifier || '',
            senderName: email.from_attendee?.display_name || '',

            message: email.body_plain || email.body || '',

            receivedAt: email.date
                ? new Date(email.date)
                : new Date(),

            attachments: (email.attachments || []).map(
                (attachment: any) => ({
                    filename: attachment.name || '',
                    mimeType:
                        attachment.mime ||
                        attachment.content_type ||
                        '',
                    attachmentId: attachment.id || '',
                }),
            ),
        }));
        const inbox =
            await this.unipileInboxModel.findOneAndUpdate(
                { accountId },
                {
                    accountId,
                    email: account.email,
                    connectionType: account.connectionType,
                    messages,
                },
                {
                    new: true,
                    upsert: true,
                    setDefaultsOnInsert: true,
                },
            );

        console.log(
            `Synced ${messages.length} inbox messages for ${accountId}`,
        );

        return inbox;
    }

    async syncUnipileOutbox(accountId: string) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const account = await this.unipileAccountModel.findOne({
            accountId,
        });

        if (!account) {
            throw new Error('Unipile account not found');
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails?limit=50&account_id=${encodeURIComponent(accountId)}&folder=SENT`,
            {
                method: 'GET',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                },
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile outbox sync failed: ${response.status} ${error}`,
            );
        }

        const data = await response.json();

        const messages = (data.items || []).map((email: any) => ({
            messageId: email.id || '',
            providerMessageId: email.provider_id || '',

            threadId: email.thread_id || '',
            subject: email.subject || '',


            receiverEmail:
                email.to_attendees?.[0]?.identifier || '',

            receiverName:
                email.to_attendees?.[0]?.display_name || '',

            message: email.body_plain || email.body || '',

            sentAt: email.date
                ? new Date(email.date)
                : new Date(),

            attachments: (email.attachments || []).map(
                (attachment: any) => ({
                    filename: attachment.name || '',
                    mimeType:
                        attachment.mime ||
                        attachment.content_type ||
                        '',
                    attachmentId: attachment.id || '',
                }),
            ),
        }));

        const outbox =
            await this.unipileOutboxModel.findOneAndUpdate(
                { accountId },
                {
                    accountId,
                    email: account.email,
                    connectionType: account.connectionType,
                    messages,
                },
                {
                    new: true,
                    upsert: true,
                    setDefaultsOnInsert: true,
                },
            );

        console.log(
            `Synced ${messages.length} outbox messages for ${accountId}`,
        );

        return outbox;
    }

    //delete or thrash a message from  from the inbox and outbox 

    async deleteUnipileMessage(accountId: string, messageId: string) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails/${encodeURIComponent(messageId)}?account_id=${encodeURIComponent(accountId)}`,
            {
                method: 'DELETE',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                },
            },
        );

        if (!response.ok) {
            const error = await response.text();

            throw new Error(
                `Unipile delete failed: ${response.status} ${error}`,
            );
        }

        return true;
    }

    async downloadUnipileAttachment(
        accountId: string,
        messageId: string,
        attachmentId: string,
    ) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        const url =
            `https://${dsn}/api/v1/emails/` +
            `${encodeURIComponent(messageId)}/attachments/` +
            `${encodeURIComponent(attachmentId)}` +
            `?account_id=${encodeURIComponent(accountId)}`;

        console.log('========== UNIPILE ATTACHMENT TEST ==========');
        console.log('URL:', url);
        console.log('accountId:', accountId);
        console.log('messageId:', messageId);
        console.log('attachmentId:', attachmentId);

        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-API-KEY': apiKey,
                Accept: '*/*',
            },
        });

        console.log('Unipile status:', response.status);
        console.log(
            'Unipile content-type:',
            response.headers.get('content-type'),
        );

        if (!response.ok) {
            const error = await response.text();

            console.log('Unipile error:', error);

            throw new Error(
                `Unipile attachment download failed: ${response.status} ${error}`,
            );
        }

        const buffer = Buffer.from(
            await response.arrayBuffer(),
        );

        return {
            buffer,
            contentType:
                response.headers.get('content-type') ||
                'application/octet-stream',
            contentDisposition:
                response.headers.get('content-disposition') ||
                undefined,
        };
    }

    async sendUnipileEmail(
        accountId: string,
        to: string,
        subject: string,
        message: string,
    ) {
        const apiKey = this.configService.get<string>('unipile.apiKey');
        const dsn = this.configService.get<string>('unipile.dsn');

        if (!apiKey) {
            throw new Error('UNIPILE_API_KEY is not configured');
        }

        if (!dsn) {
            throw new Error('UNIPILE_DSN is not configured');
        }

        // Make sure the Unipile account exists in your database
        const account = await this.unipileAccountModel.findOne({
            accountId,
        });

        if (!account) {
            throw new Error(
                `Unipile account not found: ${accountId}`,
            );
        }

        const response = await fetch(
            `https://${dsn}/api/v1/emails`,
            {
                method: 'POST',
                headers: {
                    'X-API-KEY': apiKey,
                    Accept: 'application/json',
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    account_id: accountId,

                    to: [
                        {
                            identifier: to,
                        },
                    ],

                    subject,

                    body: message,
                }),
            },
        );

        const responseText = await response.text();

        if (!response.ok) {
            console.error(
                'Unipile send email failed:',
                response.status,
                responseText,
            );

            throw new Error(
                `Unipile send email failed: ${response.status} ${responseText}`,
            );
        }

        let data: any;

        try {
            data = JSON.parse(responseText);
        } catch {
            data = responseText;
        }

        console.log('Unipile email sent successfully:', data);

        return data;
    }
}