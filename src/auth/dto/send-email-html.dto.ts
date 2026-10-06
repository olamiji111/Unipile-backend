// dto/send-unipile-html-email.dto.ts

import { IsEmail, IsNotEmpty, IsString } from 'class-validator';

export class SendUnipileHtmlEmailDto {
    @IsNotEmpty()
    @IsString()
    accountId: string;

    @IsNotEmpty()
    @IsEmail()
    to: string;

    @IsNotEmpty()
    @IsString()
    subject: string;

    @IsNotEmpty()
    @IsString()
    html: string;
}