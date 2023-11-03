import * as React from 'react';
import { IPnPjsExampleProps } from './IPnPjsExampleProps';

import { IItem, IResponseItem } from "./interfaces";

import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";

import { DefaultButton } from 'office-ui-fabric-react';

interface IPnPjsExampleState {
  items: IItem[];
  buttonClickCount: number;
}

export default class PnPjsExample extends React.Component<IPnPjsExampleProps, IPnPjsExampleState> {
  private _sp: SPFI;

  constructor(props: IPnPjsExampleProps) {
    super(props);

    this.state = {
      items: [],
      buttonClickCount: 0
    };
    this._sp = getSP();
  }

  public render(): React.ReactElement<IPnPjsExampleProps> {
    return (
      <div>
        <DefaultButton
          text="CREATE"
          onClick={this._handleCREATEButtonClick}
          style={{ margin: '10px' }}
        />
        <DefaultButton
          text="READ"
          onClick={this._handleREADButtonClick}
          style={{ margin: '10px' }}
        />
        {this.state.items.length > 0 && this._renderREADButton()}
      </div>
    );
  }

  private _handleCREATEButtonClick = async (): Promise<void> => {
    try {
      spfi(this._sp).web.lists.getByTitle("Test").items.add({
        Title: "added"
      });
    } catch (err) {
      Logger.write(`PnPjsExample (_createItem) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

  private _handleREADButtonClick = (): void => {
    const newCount = this.state.buttonClickCount + 1;
    this.setState({ buttonClickCount: newCount });

    if (newCount % 2 === 0) {
      this._clearItems();
    } else {
      this._getListItems();
    }
  }

  private _renderREADButton = (): JSX.Element => {
    return (
      <table>
        <tbody>
          {this.state.items.map((item, idx) => {
            return (
              <tr key={idx}>
                <td>{item.Id}</td>
                <td>{item.Title}</td>
                <td>{item.Created}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    );
  }

  private _clearItems = (): void => {
    this.setState({ items: [] });
  }

  private _getListItems = async (): Promise<void> => {
    try {
      const spCache = spfi(this._sp)

      const response: IResponseItem[] = await spCache.web.lists
        .getByTitle("Test")
        .items
        .select("Id", "Title", "Created")();

      const items: IItem[] = response.map((item: IResponseItem) => {
        return {
          Id: item.Id,
          Title: item.Title || "Unknown",
          Created: item.Created,
        };
      });

      this.setState({ items });
    } catch (err) {
      Logger.write(`PnPjsExample (_getList) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }
}
