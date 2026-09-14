import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument } from 'mongoose';

export type UnipileAccountDocument =
    HydratedDocument<UnipileAccount>;

@Schema({
    timestamps: true,
})
export class UnipileAccount {
    @Prop({
        required: true,
        unique: true,
        index: true,
    })
    accountId: string;

    @Prop({
        required: true,
        default: '',
    })
    email: string;

    @Prop()
    name?: string;

    @Prop({
        required: true,
        enum: ['GOOGLE', 'MICROSOFT'],
    })
    connectionType: 'GOOGLE' | 'MICROSOFT';

    @Prop()
    provider?: string;
}

export const UnipileAccountSchema =
    SchemaFactory.createForClass(UnipileAccount);