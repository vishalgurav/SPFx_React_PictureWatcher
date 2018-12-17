import * as React from "react";
import styles from "./PictureWatcher.module.scss";
import { IPictureWatcherProps } from "./IPictureWatcherProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IPictureWatcherState, IPicture } from "./IPictureWatcherState";
import { IPictureItem } from "./IPictureItem";
import { sp, Web, Site } from "@pnp/sp";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

import {
  ListSubscriptionFactory,
  IListSubscription
} from "@microsoft/sp-list-subscription";
import { Guid } from "@microsoft/sp-core-library";

export default class PictureWatcher extends React.Component<
  IPictureWatcherProps,
  IPictureWatcherState
> {
  private _subscription: IListSubscription;

  constructor(props: IPictureWatcherProps) {
    super(props);

    this.state = {
      pictures: [],
      loading: true
    };
  }
  public componentDidMount(): void {
    if (!this.props.pictureLibraryId) {
      return;
    }
    this._configureSubscription();
    this._loadPictures();
  }
  public componentDidUpdate(
    prevProps: Readonly<IPictureWatcherProps>,
    prevState: Readonly<IPictureWatcherState>,
    snapshot?: any
  ): void {
    if (this.props.pictureLibraryId === prevProps.pictureLibraryId) {
      return;
    }
    this._configureSubscription();
    this._loadPictures();
  }
  private _configureSubscription(): void {
    if (!this.props.pictureLibraryId && this._subscription) {
      this.props.listSubscriptionFactory.deleteSubscription(this._subscription);
      return;
    }
    this.props.listSubscriptionFactory.createSubscription({
      listId: Guid.parse(this.props.pictureLibraryId),
      callbacks: {
        notification: this._loadPictures
      }
    });
  }
  private _loadPictures = () => {
    this.setState({
      pictures: [],
      loading: true
    });
    sp.web.lists
      .getById(this.props.pictureLibraryId)
      .items.select("Title", "FileRef")
      .orderBy("Modified", false)
      .getAll()
      .then((pictures: IPictureItem[]) => {
        const newpicture: IPicture[] = pictures.map<IPicture>(p => ({
          title: p.title,
          serverRelativeUrl: p.FileRef
        }));
        this.setState({
          pictures: newpicture,
          loading: false
        });
      })
      .catch(err => {
        this.setState({ loading: false });
        console.log(err);
      });
  };
  public render(): React.ReactElement<IPictureWatcherProps> {
    const needsConfiguration: boolean = !this.props.pictureLibraryId;

    return (
      <div className={styles.pictureWatcher}>
        {needsConfiguration && (
          <Placeholder
            iconName="Edit"
            iconText="Configure your webpart"
            description="Please configure the webpart"
            buttonLabel="Configure"
            onConfigure={this.props.onConfigure}
          />
        )}
        {!needsConfiguration && this.state.loading && (
          <div style={{ textAlign: "center" }}>
            <Spinner size={SpinnerSize.large} label="Loading pictures..." />
          </div>
        )}
        {!needsConfiguration && !this.state.loading && (
          <div className={styles.container}>
            <div>
              {this.state.pictures != undefined
                ? this.state.pictures.map(pictures => (
                    <img
                      src={pictures.serverRelativeUrl}
                      alt={pictures.title}
                    />
                  ))
                : undefined}
            </div>
          </div>
        )}
      </div>
    );
  }
}
