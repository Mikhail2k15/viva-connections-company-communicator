import { AadHttpClient } from "@microsoft/sp-http";
import { IMessage, IMessageDetails } from "./IMessage";
import { IMessagesService } from "./IMessagesService";
import { MessagesServiceException } from "./MessagesServiceException";

export class MessagesService implements IMessagesService {

    constructor(private aadClient: AadHttpClient, private baseUrl: string) {
        this.aadClient = aadClient;
        this.baseUrl = baseUrl;
    }

    public async getSentMessages(): Promise<IMessage[]> {
        try {
            const httpResponse = await this.aadClient.get(`${this.baseUrl}/api/sentnotifications`, 
            AadHttpClient.configurations.v1);

            const result: IMessage[] = await httpResponse.json();
            return result;
        }
        catch (error) {
            throw new MessagesServiceException(`Can't getSentMessages, error: ${error}`);
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
            console.log(error);
            throw new MessagesServiceException(`Can't getMessage with ${id}, error: ${error}`);
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
            console.log(error);
            throw new MessagesServiceException(`Can't getMessageDetails with ${id}, error: ${error}`);
        } 
    }
}





