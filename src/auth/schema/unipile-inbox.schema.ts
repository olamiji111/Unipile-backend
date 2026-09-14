import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument } from 'mongoose';

export type UnipileInboxDocument =
    HydratedDocument<UnipileInbox>;

@Schema({
    timestamps: true,
})
export class UnipileInbox {
    @Prop({
        required: true,
        index: true,
    })
    accountId: string;

    @Prop({
        required: true,
    })
    email: string;

    @Prop({
        required: true,
        enum: ['GOOGLE', 'MICROSOFT'],
    })
    connectionType: 'GOOGLE' | 'MICROSOFT';

    @Prop({
        type: [
            {
                messageId: String,

                providerMessageId: String,

                threadId: String,
                subject: String,

                senderEmail: String,
                senderName: String,

                message: String,

                receivedAt: Date,

                attachments: [
                    {
                        filename: String,
                        mimeType: String,
                        attachmentId: String,
                    },
                ],
            },
        ],
        default: [],
    })
    messages: {
        messageId: string;

        providerMessageId: string;

        threadId: string;

        subject: string;

        senderEmail: string;
        senderName: string;

        message: string;

        receivedAt: Date;

        attachments: {
            filename: string;
            mimeType: string;
            attachmentId: string;
        }[];
    }[];
}

export const UnipileInboxSchema =
    SchemaFactory.createForClass(UnipileInbox);