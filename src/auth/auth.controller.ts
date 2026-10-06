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

}