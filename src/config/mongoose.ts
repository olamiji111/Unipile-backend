import { registerAs } from '@nestjs/config';

export const mongooseConfig = registerAs('mongoose', () => ({
    uri: process.env.MONGO_URI || '',
    options: {},
}));
