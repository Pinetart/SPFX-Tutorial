import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ReadReceiptWebpartWebPartStrings";
import ReadReceiptWebpart from "./components/ReadReceiptWebpart";
import { IReadReceiptWebpartProps } from "./components/IReadReceiptWebpartProps";

import { sp } from "@pnp/sp/presets/all";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls";

export interface IReadReceiptWebpartWebPartProps {
  description: string;
}

export default class ReadReceiptWebpartWebPart extends BaseClientSideWebPart<IReadReceiptWebpartWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IReadReceiptWebpartProps> =
      React.createElement(ReadReceiptWebpart, {
        description: this.properties.description,
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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
