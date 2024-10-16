import { ISoftwareListItem } from "./ISoftwareListItem";

export interface IReactCrudState {
  status: string;
  SoftwareListItems: ISoftwareListItem[];
  SoftwareListItem: ISoftwareListItem;
}
