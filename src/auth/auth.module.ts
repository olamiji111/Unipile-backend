import { Module } from '@nestjs/common';
import { AuthController } from './auth.controller';
import { AuthService } from './auth.service';
import { MongooseModule } from '@nestjs/mongoose';
import { UnipileAccount, UnipileAccountSchema } from './schema/unipile-account.schema';
import { UnipileInbox, UnipileInboxSchema } from './schema/unipile-inbox.schema';
import { UnipileOutbox, UnipileOutboxSchema } from './schema/unipile-outbox.schema';

@Module({
    imports: [
        MongooseModule.forFeature([


            {
                name: UnipileAccount.name,
                schema: UnipileAccountSchema,
            },
            {
                name: UnipileInbox.name,
                schema: UnipileInboxSchema
            },
            {
                name: UnipileOutbox.name,
                schema: UnipileOutboxSchema
            }
        ]),
    ],

    controllers: [AuthController],

    providers: [AuthService],

    exports: [AuthService],
})
export class AuthModule { }
