import { ISoftwareListItem } from "./ISoftwareListItem";

export interface ICrudWithReactState {
    status: string;
    SoftwareListItems: ISoftwareListItem[];
    SoftwareListItem: ISoftwareListItem;
}