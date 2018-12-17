export interface IPictureWatcherState {
  pictures: IPicture[];
  loading: boolean;
}

export interface IPicture {
  title: string;
  serverRelativeUrl: string;
}
