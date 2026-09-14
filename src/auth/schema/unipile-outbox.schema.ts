import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument } from 'mongoose';

export type UnipileOutboxDocument =
    HydratedDocument<UnipileOutbox>;

@Schema({
    timestamps: true,
})
export class UnipileOutbox {
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

                receiverEmail: String,
                receiverName: String,

                message: String,

                sentAt: Date,

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

        receiverEmail: string;
        receiverName: string;

        message: string;

        sentAt: Date;

        attachments: {
            filename: string;
            mimeType: string;
            attachmentId: string;
        }[];
    }[];
}

export const UnipileOutboxSchema =
    SchemaFactory.createForClass(UnipileOutbox);