import { IMessage } from "./IMessage";

export interface IMessagesService {
    /**
     * Retrieves the last sent 25 messages.     
     * @returns {Promise<IMessage[]>} the sent messages     
     */
    getSentMessages(): Promise<IMessage[]>;

    /**
     * Retrieves a message by id
     * @param id the id of the message to return
     * @returns {Promise<IMessage>} the message with the given id
     */
    getMessage(id: string): Promise<IMessage>;
}