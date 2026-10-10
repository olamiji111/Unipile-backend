import {
    Controller,
    Get,
    Req,
    UseGuards,
    Res,
    Param,
    Post,
    HttpStatus,
    HttpCode,
    Body,
    Delete,
    UseInterceptors,
    UploadedFile,
    BadRequestException,
} from '@nestjs/common';
import type { Response } from 'express';
import { AuthService } from './auth.service';
import { SendEmailDto } from './dto/send-email.dto';
import { SendUnipileHtmlEmailDto } from './dto/send-email-html.dto';
import { FileInterceptor } from '@nestjs/platform-express';

@Controller('auth')
export class AuthController {
    constructor(private readonly authService: AuthService) { }

    @Get()
    async googleAuth(@Res() res: Response) {
        const authUrl = await this.authService.createUnipileGoogleAuthLink();
        return res.redirect(authUrl);
    }

    @Get('/microsoft')
    async microsoftAuth(@Res() res: Response) {
        const authUrl =
            await this.authService.createUnipileMicrosoftAuthLink();

        return res.redirect(authUrl);
    }

    @Get('/smtp')
    async smtpAuth(@Res() res: Response) {
        const authUrl =
            await this.authService.createUnipileSmtpAuthLink();

        return res.redirect(authUrl);
    }

    @Get('/unipile/callback')
    async unipileCallback(
        @Req() req: any,
        @Res() res: Response,
    ) {
        console.log('Unipile callback:', req.query);

        const accountId = req.query.account_id;

        if (!accountId) {
            return res.status(400).send('Unipile account ID not found');
        }

        const provider = req.query.provider || '';

        const connectionType =
            provider.toUpperCase() === 'OUTLOOK'
                ? 'MICROSOFT'
                : 'GOOGLE';

        // Save the connected Unipile account
        await this.authService.saveUnipileAccount(
            accountId,
            req.query.email || '',
            connectionType,
            req.query.name,
            provider,
        );

        console.log(
            `Unipile ${connectionType} account saved: ${accountId}`,
        );

        // Fetch 50 inbox messages
        await this.authService.fetchUnipileInbox(accountId);

        console.log(
            `Unipile inbox saved for account: ${accountId}`,
        );

        // Fetch 50 sent messages
        await this.authService.fetchUnipileOutbox(accountId);

        console.log(
            `Unipile outbox saved for account: ${accountId}`,
        );

        // Redirect after everything is saved
        if (connectionType === 'MICROSOFT') {

            return res.redirect('https://www.microsoft.com');

        }

        return res.redirect('https://www.google.com');
    }


    @Get('/unipile/sync/inbox/:accountId')
    async syncUnipileInbox(
        @Param('accountId') accountId: string,
    ) {
        return this.authService.syncUnipileInbox(accountId);
    }

    @Get('/unipile/sync/outbox/:accountId')
    async syncUnipileOutbox(
        @Param('accountId') accountId: string,
    ) {
        return this.authService.syncUnipileOutbox(accountId);
    }

    @Get('/unipile/delete/:accountId/:messageId')
    async deleteUnipileMessage(
        @Param('accountId') accountId: string,
        @Param('messageId') messageId: string,
    ) {
        return this.authService.deleteUnipileMessage(
            accountId,
            messageId,
        );
    }



    @Get('/unipile/attachment/:accountId/:messageId/:attachmentId')
    async downloadUnipileAttachment(
        @Param('accountId') accountId: string,
        @Param('messageId') messageId: string,
        @Param('attachmentId') attachmentId: string | string[],
        @Res() res: Response,
    ) {
        try {
            const id = Array.isArray(attachmentId)
                ? attachmentId.join('/')
                : attachmentId;

            console.log('Downloading Unipile attachment:', {
                accountId,
                messageId,
                attachmentId: id,
            });

            const file =
                await this.authService.downloadUnipileAttachment(
                    accountId,
                    messageId,
                    id,
                );

            res.setHeader(
                'Content-Type',
                file.contentType,
            );

            if (file.contentDisposition) {
                res.setHeader(
                    'Content-Disposition',
                    file.contentDisposition,
                );
            }

            return res.send(file.buffer);
        } catch (error) {
            console.error(
                'Attachment download failed:',
                error,
            );

            return res.status(500).json({
                success: false,
                error:
                    error instanceof Error
                        ? error.message
                        : String(error),
            });
        }
    }
    @Post('/unipile/send-html')
    @UseInterceptors(
        FileInterceptor('file'),
    )
    async sendHtmlEmail(
        @UploadedFile() file: Express.Multer.File,
        @Body() dto: SendUnipileHtmlEmailDto,
    ) {
        if (!file) {
            throw new BadRequestException(
                'Attachment is required',
            );
        }

        return this.authService.sendUnipileHtmlEmail(
            dto.accountId,
            dto.to,
            dto.subject,
            dto.html,
            file.buffer,
            file.originalname,
        );
    }

    @Get('/unipile/send')
    getSendEmailPage(@Res() res: Response) {
        res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Send Email - Unipile</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    max-width: 500px;
                    margin: 50px auto;
                    padding: 20px;
                }

                input, textarea {
                    width: 100%;
                    padding: 10px;
                    margin: 8px 0 15px;
                    box-sizing: border-box;
                }

                textarea {
                    height: 150px;
                }

                button {
                    padding: 10px 20px;
                    cursor: pointer;
                }

                #result {
                    margin-top: 20px;
                }
            </style>
        </head>

        <body>
            <h2>Send Email</h2>

            <label>Unipile Account ID</label>
            <input
                id="accountId"
                type="text"
                placeholder="Enter Unipile account ID"
            />

            <label>To</label>
            <input
                id="to"
                type="email"
                placeholder="recipient@example.com"
            />

            <label>Subject</label>
            <input
                id="subject"
                type="text"
                placeholder="Subject"
            />

            <label>Message</label>
            <textarea
                id="message"
                placeholder="Write your message..."
            ></textarea>

            <button onclick="sendEmail()">Send Email</button>

            <div id="result"></div>

            <script>
                async function sendEmail() {
                    const accountId =
                        document.getElementById('accountId').value;

                    const to =
                        document.getElementById('to').value;

                    const subject =
                        document.getElementById('subject').value;

                    const message =
                        document.getElementById('message').value;

                    if (!accountId || !to || !subject || !message) {
                        document.getElementById('result').innerText =
                            'Please fill all fields.';
                        return;
                    }

                    try {
                        const response = await fetch(
                            '/auth/unipile/send/' +
                            encodeURIComponent(accountId),
                            {
                                method: 'POST',
                                headers: {
                                    'Content-Type': 'application/json'
                                },
                                body: JSON.stringify({
                                    to,
                                    subject,
                                    message
                                })
                            }
                        );

                        const data = await response.json();

                        if (!response.ok) {
                            throw new Error(
                                data.message || data.error || 'Failed to send email'
                            );
                        }

                        document.getElementById('result').innerText =
                            'Email sent successfully!';

                        console.log(data);

                    } catch (error) {
                        document.getElementById('result').innerText =
                            'Error: ' + error.message;
                    }
                }
            </script>
        </body>
        </html>
    `);
    }

    @Post('/unipile/send/:accountId')
    sendUnipileEmail(
        @Param('accountId') accountId: string,
        @Body()
        body: SendEmailDto
    ) {
        return this.authService.sendUnipileEmail(
            accountId,
            body.to,
            body.subject,
            body.message,
        );
    }

    @Get('/unipile/attachment')
    getDownloadAttachmentPage(@Res() res: Response) {
        res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Download Attachment - Unipile</title>

            <style>
                body {
                    font-family: Arial, sans-serif;
                    max-width: 500px;
                    margin: 50px auto;
                    padding: 20px;
                }

                input {
                    width: 100%;
                    padding: 10px;
                    margin: 8px 0 15px;
                    box-sizing: border-box;
                }

                button {
                    padding: 10px 20px;
                    cursor: pointer;
                }

                #result {
                    margin-top: 20px;
                }
            </style>
        </head>

        <body>

            <h2>Download Attachment</h2>

            <label>Unipile Account ID</label>
            <input
                id="accountId"
                type="text"
                placeholder="Enter Unipile account ID"
            />

            <label>Message ID</label>
            <input
                id="messageId"
                type="text"
                placeholder="Enter message ID"
            />

            <label>Attachment ID</label>
            <input
                id="attachmentId"
                type="text"
                placeholder="Enter attachment ID"
            />

            <button onclick="downloadAttachment()">
                Download Attachment
            </button>

            <div id="result"></div>

            <script>
                async function downloadAttachment() {

                    const accountId =
                        document.getElementById('accountId').value.trim();

                    const messageId =
                        document.getElementById('messageId').value.trim();

                    const attachmentId =
                        document.getElementById('attachmentId').value.trim();

                    if (!accountId || !messageId || !attachmentId) {
                        document.getElementById('result').innerText =
                            'Please fill all fields.';
                        return;
                    }

                    try {

                        document.getElementById('result').innerText =
                            'Downloading...';

                        const url =
                            '/auth/unipile/attachment/' +
                            encodeURIComponent(accountId) +
                            '/' +
                            encodeURIComponent(messageId) +
                            '/' +
                            encodeURIComponent(attachmentId);

                        const response = await fetch(url);

                        if (!response.ok) {
                            const errorText = await response.text();

                            throw new Error(
                                errorText || 'Failed to download attachment'
                            );
                        }

                        const blob = await response.blob();

                        const downloadUrl =
                            window.URL.createObjectURL(blob);

                        const link =
                            document.createElement('a');

                        link.href = downloadUrl;

                        link.download = 'attachment';

                        document.body.appendChild(link);

                        link.click();

                        link.remove();

                        window.URL.revokeObjectURL(downloadUrl);

                        document.getElementById('result').innerText =
                            'Attachment downloaded successfully!';

                    } catch (error) {

                        document.getElementById('result').innerText =
                            'Error: ' + error.message;

                        console.error(
                            'Attachment download failed:',
                            error
                        );
                    }
                }
            </script>

        </body>
        </html>
    `);
    }

    @Get('/unipile/send-html')
    getSendHtmlEmailPage(@Res() res: Response) {
        res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Send HTML Email - Unipile</title>

            <style>
                body {
                    font-family: Arial, sans-serif;
                    max-width: 600px;
                    margin: 50px auto;
                    padding: 20px;
                }

                input,
                textarea {
                    width: 100%;
                    padding: 10px;
                    margin: 8px 0 15px;
                    box-sizing: border-box;
                }

                textarea {
                    height: 200px;
                    font-family: monospace;
                }

                button {
                    padding: 10px 20px;
                    cursor: pointer;
                }

                #result {
                    margin-top: 20px;
                }
            </style>
        </head>

        <body>

            <h2>Send HTML Email</h2>

            <label>Unipile Account ID</label>

            <input
                id="accountId"
                type="text"
                placeholder="Enter Unipile account ID"
            />

            <label>To</label>

            <input
                id="to"
                type="email"
                placeholder="recipient@example.com"
            />

            <label>Subject</label>

            <input
                id="subject"
                type="text"
                placeholder="Invoice #INV-2026-10398"
            />

            <label>HTML Message</label>

            <textarea
                id="html"
                placeholder="<p>Hi Peter,</p><p>Please find attached your invoice.</p>"
            ></textarea>

            <label>Attachment</label>

            <input
                id="file"
                type="file"
            />

            <button onclick="sendEmail()">
                Send Email
            </button>

            <div id="result"></div>

            <script>
                async function sendEmail() {

                    const accountId =
                        document.getElementById('accountId').value;

                    const to =
                        document.getElementById('to').value;

                    const subject =
                        document.getElementById('subject').value;

                    const html =
                        document.getElementById('html').value;

                    const file =
                        document.getElementById('file').files[0];

                    if (
                        !accountId ||
                        !to ||
                        !subject ||
                        !html ||
                        !file
                    ) {
                        document.getElementById('result').innerText =
                            'Please fill all fields and select a file.';

                        return;
                    }

                    const formData = new FormData();

                    formData.append(
                        'accountId',
                        accountId
                    );

                    formData.append(
                        'to',
                        to
                    );

                    formData.append(
                        'subject',
                        subject
                    );

                    formData.append(
                        'html',
                        html
                    );

                    formData.append(
                        'file',
                        file
                    );

                    try {

                        document.getElementById('result').innerText =
                            'Sending...';

                        const response = await fetch(
                            '/auth/unipile/send-html',
                            {
                                method: 'POST',
                                body: formData
                            }
                        );

                        const data =
                            await response.json();

                        if (!response.ok) {
                            throw new Error(
                                data.message ||
                                data.error ||
                                'Failed to send email'
                            );
                        }

                        document.getElementById('result').innerText =
                            'Email sent successfully!';

                        console.log(data);

                    } catch (error) {

                        document.getElementById('result').innerText =
                            'Error: ' + error.message;

                        console.error(error);
                    }
                }
            </script>

        </body>
        </html>
    `);
    }


    @Get('/unipile/mail-inbox')
    showUnipileInbox(@Res() res: Response) {
        return res.send(`
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mail Inbox</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #f4f6fb;
      color: #202938;
      font-family: Arial, Helvetica, sans-serif;
    }
    button, input { font: inherit; }
    .app {
      max-width: 1500px;
      min-height: 100vh;
      margin: auto;
      padding: 24px;
    }
    .topbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 16px;
      margin-bottom: 24px;
    }
    .brand {
      font-size: 25px;
      font-weight: 750;
      letter-spacing: -.7px;
    }
    .brand span { color: #3978f6; }
    .account-form {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }
    .account-input {
      width: 280px;
      max-width: 100%;
      padding: 12px 14px;
      border: 1px solid #dce2ec;
      border-radius: 10px;
      outline: none;
      background: #fff;
    }
    .account-input:focus {
      border-color: #3978f6;
      box-shadow: 0 0 0 3px #3978f61c;
    }
    .btn {
      padding: 11px 16px;
      border: 0;
      border-radius: 9px;
      cursor: pointer;
      transition: background .2s, transform .2s;
    }
    .btn:disabled {
      opacity: .55;
      cursor: not-allowed;
    }
    .btn-primary {
      color: #fff;
      background: #3978f6;
    }
    .btn-primary:hover:not(:disabled) {
      background: #245fda;
      transform: translateY(-1px);
    }
    .btn-secondary {
      color: #334155;
      background: #e9eef6;
    }
    .btn-secondary:hover:not(:disabled) {
      background: #dce5f2;
    }
    .status {
      margin: 0 0 16px;
      color: #64748b;
      font-size: 13px;
    }
    .mail-layout {
      display: grid;
      grid-template-columns: minmax(280px, 380px) minmax(0, 1fr);
      min-height: 650px;
      overflow: hidden;
      background: #fff;
      border: 1px solid #e4e9f1;
      border-radius: 16px;
      box-shadow: 0 8px 35px #17255408;
    }
    .inbox-panel {
      min-width: 0;
      border-right: 1px solid #e8edf4;
      background: #fff;
    }
    .panel-heading {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 20px;
      border-bottom: 1px solid #edf0f5;
    }
    .panel-heading h2 {
      margin: 0;
      font-size: 17px;
    }
    .count {
      padding: 5px 9px;
      border-radius: 20px;
      background: #edf4ff;
      color: #3978f6;
      font-size: 12px;
      font-weight: 700;
    }
    .message-list {
      max-height: 760px;
      overflow-y: auto;
    }
    .message-item {
      display: block;
      width: 100%;
      padding: 18px;
      text-align: left;
      color: inherit;
      background: #fff;
      border: 0;
      border-bottom: 1px solid #edf0f5;
      border-left: 3px solid transparent;
      cursor: pointer;
      transition: background .18s, border-color .18s, box-shadow .18s;
    }
    .message-item:hover {
      background: #f2f7ff;
      border-left-color: #7aa7ff;
      box-shadow: inset 0 0 0 1px #e1ecff;
    }
    .message-item.active {
      background: #edf4ff;
      border-left-color: #3978f6;
    }
    .message-sender {
      display: block;
      overflow: hidden;
      margin-bottom: 8px;
      font-size: 14px;
      font-weight: 700;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .message-subject {
      display: block;
      overflow: hidden;
      margin-bottom: 7px;
      font-size: 13px;
      font-weight: 600;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .message-preview {
      display: block;
      overflow: hidden;
      color: #7b8798;
      font-size: 12px;
      line-height: 1.6;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .message-date {
      display: block;
      margin-top: 9px;
      color: #94a0b1;
      font-size: 11px;
    }
    .reader {
      display: flex;
      min-width: 0;
      flex-direction: column;
      background: #fff;
    }
    .reader-header {
      padding: 26px;
      border-bottom: 1px solid #edf0f5;
    }
    .reader-subject {
      margin: 0 0 22px;
      overflow-wrap: anywhere;
      font-size: 23px;
      line-height: 1.4;
    }
    .sender-row {
      display: flex;
      align-items: center;
      gap: 13px;
    }
    .avatar {
      display: flex;
      width: 43px;
      height: 43px;
      flex-shrink: 0;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      color: #245fda;
      background: #e5efff;
      font-weight: 700;
    }
    .sender-info { min-width: 0; }
    .sender-name {
      margin-bottom: 5px;
      overflow-wrap: anywhere;
      font-size: 14px;
      font-weight: 700;
    }
    .sender-email {
      overflow-wrap: anywhere;
      color: #7b8798;
      font-size: 12px;
    }
    .reader-date {
      margin-top: 15px;
      color: #7b8798;
      font-size: 12px;
    }
    .reader-body {
      min-height: 420px;
      flex: 1;
      padding: 22px;
      background: #fff;
    }
    .email-frame {
      display: block;
      width: 100%;
      min-height: 480px;
      border: 0;
      background: #fff;
    }
    .empty-state {
      display: flex;
      min-height: 480px;
      padding: 30px;
      align-items: center;
      justify-content: center;
      flex-direction: column;
      text-align: center;
      color: #8793a5;
    }
    .empty-icon {
      display: flex;
      width: 64px;
      height: 64px;
      margin-bottom: 18px;
      align-items: center;
      justify-content: center;
      border-radius: 20px;
      background: #edf4ff;
      color: #3978f6;
      font-size: 29px;
    }
    .empty-title {
      margin-bottom: 9px;
      color: #334155;
      font-size: 18px;
      font-weight: 700;
    }
    .empty-description {
      max-width: 300px;
      font-size: 13px;
      line-height: 1.7;
    }
    .mobile-back { display: none; }
    @media (max-width: 850px) {
      .app { padding: 14px; }
      .mail-layout {
        display: block;
        min-height: 75vh;
      }
      .inbox-panel { border-right: 0; }
      .reader { display: none; }
      .mail-layout.show-reader .inbox-panel { display: none; }
      .mail-layout.show-reader .reader { display: flex; }
      .mobile-back {
        display: inline-block;
        margin-bottom: 15px;
      }
      .reader-header { padding: 20px; }
      .reader-subject { font-size: 20px; }
      .reader-body { padding: 12px; }
      .email-frame { min-height: 65vh; }
      .account-form { width: 100%; }
      .account-input {
        flex: 1;
        width: auto;
        min-width: 0;
      }
    }
    @media (max-width: 450px) {
      .app { padding: 10px; }
      .brand { font-size: 22px; }
      .topbar {
        align-items: flex-start;
        flex-direction: column;
      }
      .account-input { width: 100%; }
      .account-form > button { flex: 1; }
      .reader-subject { font-size: 18px; }
      .message-item { padding: 15px; }
    }
  </style>
</head>
<body>
  <main class="app">
    <header class="topbar">
      <div class="brand">Mail<span>Box</span></div>
      <form id="accountForm" class="account-form">
        <input
          id="accountId"
          class="account-input"
          type="text"
          placeholder="Enter your account ID"
          autocomplete="off"
          required
        >
        <button id="loadButton" class="btn btn-primary" type="submit">
          Load Inbox
        </button>
        <button
          id="refreshButton"
          class="btn btn-secondary"
          type="button"
          hidden
        >
          ↻ Refresh
        </button>
      </form>
    </header>
    <div id="status" class="status" role="status" aria-live="polite">
      Enter your account ID to load your messages.
    </div>
    <section id="mailLayout" class="mail-layout">
      <aside class="inbox-panel">
        <div class="panel-heading">
          <h2>Inbox</h2>
          <span id="messageCount" class="count">0 messages</span>
        </div>
        <div id="messageList" class="message-list">
          <div class="empty-state">
            <div class="empty-icon">✉</div>
            <div class="empty-title">Your inbox</div>
            <div class="empty-description">
              Load your account to see your messages here.
            </div>
          </div>
        </div>
      </aside>
      <section class="reader" id="messageDetail">
        <div class="empty-state">
          <div class="empty-icon">✉</div>
          <div class="empty-title">Select a message</div>
          <div class="empty-description">
            Choose an email from your inbox to read its full contents.
          </div>
        </div>
      </section>
    </section>
  </main>
  <script>
    const form = document.getElementById('accountForm');
    const accountInput = document.getElementById('accountId');
    const loadButton = document.getElementById('loadButton');
    const refreshButton = document.getElementById('refreshButton');
    const status = document.getElementById('status');
    const messageList = document.getElementById('messageList');
    const messageDetail = document.getElementById('messageDetail');
    const messageCount = document.getElementById('messageCount');
    const mailLayout = document.getElementById('mailLayout');
    let messages = [];
    let activeIndex = -1;
    function getMessageBody(message) {
      return message.bodyHtml ||
        message.body_html ||
        message.message ||
        message.body ||
        message.body_plain ||
        '';
    }
    function getPlainText(value) {
      const element = document.createElement('div');
      element.innerHTML = String(value || '');
      return (element.textContent || element.innerText || '')
        .replace(/\\s+/g, ' ')
        .trim();
    }
    function formatDate(value) {
      if (!value) return '';
      const date = new Date(value);
      return Number.isNaN(date.getTime())
        ? ''
        : date.toLocaleString();
    }
    function escapeHtml(value) {
      return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }
    async function loadInbox(event) {
      if (event) event.preventDefault();
      const accountId = accountInput.value.trim();
      if (!accountId) {
        status.textContent = 'Enter your account ID first.';
        return;
      }
      localStorage.setItem('unipileInboxAccountId', accountId);
      loadButton.disabled = true;
      refreshButton.disabled = true;
      refreshButton.hidden = false;
      status.textContent = 'Loading your latest messages...';
      try {
        const response = await fetch(
          '/auth/unipile/sync/inbox/' + encodeURIComponent(accountId),
          {
            headers: { Accept: 'application/json' },
            cache: 'no-store'
          }
        );
        if (!response.ok) {
          throw new Error('Inbox request failed: ' + response.status);
        }
        const data = await response.json();
        messages = Array.isArray(data)
          ? data
          : Array.isArray(data.messages)
            ? data.messages
            : [];
        renderInbox();
        status.textContent =
          'Inbox updated successfully. ' +
          messages.length + ' messages loaded.';
      } catch (error) {
        status.textContent =
          error.message || 'Unable to load your inbox.';
        if (!messages.length) {
          showEmptyState(
            'Unable to load messages',
            'Check your account ID and try again.'
          );
        }
      } finally {
        loadButton.disabled = false;
        refreshButton.disabled = false;
      }
    }
    function showEmptyState(title, description) {
      messageDetail.replaceChildren();
      const state = document.createElement('div');
      state.className = 'empty-state';
      const icon = document.createElement('div');
      icon.className = 'empty-icon';
      icon.textContent = '✉';
      const heading = document.createElement('div');
      heading.className = 'empty-title';
      heading.textContent = title;
      const detail = document.createElement('div');
      detail.className = 'empty-description';
      detail.textContent = description;
      state.append(icon, heading, detail);
      messageDetail.appendChild(state);
    }
    function renderInbox() {
      messageList.replaceChildren();
      messageDetail.replaceChildren();
      activeIndex = -1;
      messageCount.textContent =
        messages.length + (messages.length === 1 ? ' message' : ' messages');
      if (!messages.length) {
        const state = document.createElement('div');
        state.className = 'empty-state';
        const icon = document.createElement('div');
        icon.className = 'empty-icon';
        icon.textContent = '✉';
        const heading = document.createElement('div');
        heading.className = 'empty-title';
        heading.textContent = 'No messages found';
        const detail = document.createElement('div');
        detail.className = 'empty-description';
        detail.textContent = 'Your inbox is empty or no messages were returned.';
        state.append(icon, heading, detail);
        messageList.appendChild(state);
        showEmptyState(
          'Your inbox is empty',
          'When messages are available, select one to read it here.'
        );
        mailLayout.classList.remove('show-reader');
        return;
      }
      messages.forEach(function(message, index) {
        const item = document.createElement('button');
        item.type = 'button';
        item.className = 'message-item';
        const sender = document.createElement('span');
        sender.className = 'message-sender';
        sender.textContent =
          message.senderName ||
          message.senderEmail ||
          'Unknown sender';
        const subject = document.createElement('span');
        subject.className = 'message-subject';
        subject.textContent = message.subject || '(No subject)';
        const preview = document.createElement('span');
        preview.className = 'message-preview';
        preview.textContent = getPlainText(getMessageBody(message)) ||
          'No message preview available';
        const date = document.createElement('span');
        date.className = 'message-date';
        date.textContent = formatDate(
          message.receivedAt || message.sentAt || message.date
        );
        item.append(sender, subject, preview, date);
        item.addEventListener('click', function() {
          showMessage(index);
        });
        messageList.appendChild(item);
      });
      showMessage(0);
    }
    function showMessage(index) {
      const message = messages[index];
      if (!message) return;
      activeIndex = index;
      Array.from(messageList.querySelectorAll('.message-item'))
        .forEach(function(item, itemIndex) {
          item.classList.toggle('active', itemIndex === index);
        });
      messageDetail.replaceChildren();
      const header = document.createElement('div');
      header.className = 'reader-header';
      const backButton = document.createElement('button');
      backButton.type = 'button';
      backButton.className = 'btn btn-secondary mobile-back';
      backButton.textContent = '← Back to inbox';
      backButton.addEventListener('click', function() {
        mailLayout.classList.remove('show-reader');
      });
      const subject = document.createElement('h1');
      subject.className = 'reader-subject';
      subject.textContent = message.subject || '(No subject)';
      const senderRow = document.createElement('div');
      senderRow.className = 'sender-row';
      const avatar = document.createElement('div');
      avatar.className = 'avatar';
      const senderName =
        message.senderName || message.senderEmail || 'Unknown sender';
      avatar.textContent = senderName.trim().charAt(0).toUpperCase() || '?';
      const senderInfo = document.createElement('div');
      senderInfo.className = 'sender-info';
      const sender = document.createElement('div');
      sender.className = 'sender-name';
      sender.textContent = senderName;
      const email = document.createElement('div');
      email.className = 'sender-email';
      email.textContent = message.senderEmail || '';
      senderInfo.append(sender, email);
      senderRow.append(avatar, senderInfo);
      const date = document.createElement('div');
      date.className = 'reader-date';
      date.textContent = formatDate(
        message.receivedAt || message.sentAt || message.date
      );
      header.append(backButton, subject, senderRow, date);
      const bodyContainer = document.createElement('div');
      bodyContainer.className = 'reader-body';
      const frame = document.createElement('iframe');
      frame.className = 'email-frame';
      frame.title = 'Email message content';
      // Prevent email HTML from accessing the parent page.
      frame.setAttribute('sandbox', 'allow-popups');
      const rawBody = getMessageBody(message);
      const hasHtml = /<[a-z][\\s\\S]*>/i.test(String(rawBody));
      const emailDocument = hasHtml
        ? String(rawBody)
        : '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
          '<meta name="viewport" content="width=device-width,initial-scale=1">' +
          '</head><body style="font-family:Arial,sans-serif;font-size:14px;' +
          'line-height:1.7;white-space:pre-wrap;overflow-wrap:anywhere;color:#202938;">' +
          escapeHtml(rawBody) +
          '</body></html>';
      frame.srcdoc = emailDocument;
      bodyContainer.appendChild(frame);
      messageDetail.append(header, bodyContainer);
      mailLayout.classList.add('show-reader');
    }
    form.addEventListener('submit', loadInbox);
    refreshButton.addEventListener('click', loadInbox);
    const savedAccountId = '';

accountInput.value = savedAccountId;
refreshButton.hidden = true;
  </script>
</body>
</html>
  `);
    }

}