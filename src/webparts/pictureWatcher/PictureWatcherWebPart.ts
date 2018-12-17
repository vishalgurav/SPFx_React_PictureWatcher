import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { ListSubscriptionFactory } from "@microsoft/sp-list-subscription";
import * as strings from "PictureWatcherWebPartStrings";
import PictureWatcher from "./components/PictureWatcher";
import { IPictureWatcherProps } from "./components/IPictureWatcherProps";

export interface IPictureWatcherWebPartProps {
  pictureLibraryId: string;
}

export default class PictureWatcherWebPart extends BaseClientSideWebPart<
  IPictureWatcherWebPartProps
> {
  private _onConfigure = () => {
    this.context.propertyPane.open();
  };
  public render(): void {
    const element: React.ReactElement<
      IPictureWatcherProps
    > = React.createElement(PictureWatcher, {
      pictureLibraryId: this.properties.pictureLibraryId,
      listSubscriptionFactory: new ListSubscriptionFactory(this),
      onConfigure: this._onConfigure
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker("pictureLibraryId", {
                  label: "Select a libary...",
                  selectedList: this.properties.pictureLibraryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  baseTemplate: 109,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "pictureLibraryIdPicker"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
