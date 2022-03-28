
export interface IMessage {
    id: string;
    allUsers?: boolean;
    title: string;
    summary?: string;
    imageLink?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
}