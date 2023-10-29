import * as React from 'react';
import styles from './PnPjsExample.module.scss';
import { IPnPjsExampleProps } from './IPnPjsExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IItem, IResponseItem } from "./interfaces";

import { Caching } from "@pnp/queryable";
import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";

export default class PnPjsExample extends React.Component<IPnPjsExampleProps, {}> {
  private _sp: SPFI;

  constructor(props: IPnPjsExampleProps) {
    super(props);
    // set initial state
    this.state = {
    };
    this._sp = getSP();
  }

  public render(): React.ReactElement<IPnPjsExampleProps> {

    return (
    );
  }

  private _getListItems = async (): Promise<void> => {
    try {
      const spCache = spfi(this._sp).using(Caching({ store: "session" }));

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
