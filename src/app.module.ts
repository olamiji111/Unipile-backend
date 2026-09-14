import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { unipileConfig } from './config/unipile';
import { mongooseConfig } from './config/mongoose';
import { MongooseModule } from '@nestjs/mongoose';
import { Connection } from 'mongoose';
import { AuthModule } from './auth/auth.module';

@Module({
  imports: [
    AuthModule,
    ConfigModule.forRoot({
      isGlobal: true,
      load: [mongooseConfig, unipileConfig],
    }),

    // MongoDB Database Connection
    MongooseModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        uri: configService.get<string>('mongoose.uri'),
        onConnectionCreate: (connection: Connection) => {
          connection.on('connected', () => {
            console.log('Database connected successfully');
          });
          connection.on('error', (err) => {
            console.error('Database connection error:', err);
          });
          connection.on('disconnected', () => {
            console.log('Database disconnected');
          });
          return connection;
        },
      }),
    }),
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule { }
