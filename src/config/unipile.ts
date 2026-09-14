import { registerAs } from '@nestjs/config';

export const unipileConfig = registerAs('unipile', () => ({
    apiKey: process.env.UNIPILE_API_KEY || '',
    dsn: process.env.UNIPILE_DSN || '',
    redirectUri: process.env.UNIPILE_REDIRECT_URI || '',
}));
