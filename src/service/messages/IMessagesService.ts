import { IMessage } from "./IMessage";

/**
 * Defines the abstract interface for the Company Communicator API
 */
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