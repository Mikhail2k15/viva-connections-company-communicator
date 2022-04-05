
export interface IMessage {
    id: string;
    allUsers?: boolean;
    title: string;
    summary?: string;
    imageLink?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;

    sentDate?: Date;
    sendingStartedDate?: Date;
    status?: string;
    succeeded?: number;
    failed?: number;
}

export interface IMessageDetails extends IMessage {
    sendingCompleted?: boolean;
    viewCount?: number;
    sentFormattedDate?: string;
    formattedStatus?: string;
}