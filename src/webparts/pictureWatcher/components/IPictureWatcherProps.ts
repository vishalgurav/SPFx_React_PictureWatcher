import { ListSubscriptionFactory } from "@microsoft/sp-list-subscription";
export interface IPictureWatcherProps {
  pictureLibraryId: string;
  listSubscriptionFactory: ListSubscriptionFactory;
  onConfigure: () => void;
}
