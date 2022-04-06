import { AadHttpClient } from "@microsoft/sp-http";
import { IMessage, IMessageDetails } from "./IMessage";
import { IMessagesService } from "./IMessagesService";
import { MessagesServiceError } from "./MessagesServiceError";

export class MessagesService implements IMessagesService {

    constructor(private aadClient: AadHttpClient, private baseUrl: string) { 
        if (aadClient === undefined){
            throw new Error('error null agrument aadClient');
        }
        if (baseUrl === undefined){
            throw new Error('error null argument serviceBaseUrl');
        }
        this.aadClient = aadClient;
        this.baseUrl = baseUrl;
    }

    public async getSentMessages(): Promise<IMessage[]> {
        try {
            const httpResponse = await this.aadClient.get(`${this.baseUrl}/api/sentnotifications`, 
            AadHttpClient.configurations.v1);

            if (httpResponse.status === 403){
                throw new MessagesServiceError('error forbidden');
            } else if (httpResponse.status !== 200){
                throw new MessagesServiceError('error retrieve messages');
            }

            const result: IMessage[] = await httpResponse.json();
            return result;
        }
        catch (error) {

        }
    }
    public async getMessage(id: string): Promise<IMessage> {
        try {
            const httpResponse = await this.aadClient.get(`${this.baseUrl}/api/sentnotifications/${id}`, 
            AadHttpClient.configurations.v1);
            const result: IMessage = await httpResponse.json();
            return result;
        }
        catch (error) {
            throw new Error("Method not implemented.");
        }        
    }

    public async getMessageDetails(id: string): Promise<IMessageDetails> {
        try {
            const httpResponse = await this.aadClient.get(`${this.baseUrl}/api/sentnotifications/${id}`, 
            AadHttpClient.configurations.v1);
            const result: IMessageDetails = await httpResponse.json();
            return result;
        }
        catch (error) {
            throw new Error("Method not implemented.");
        } 
    }
}





