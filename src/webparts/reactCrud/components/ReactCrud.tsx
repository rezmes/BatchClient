import * as React from "react";
import styles from "./ReactCrud.module.scss";
import { ISoftwareListItem } from "./ISoftwareListItem";
import { IReactCrudProps } from "./IReactCrudProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IReactCrudState } from "./IReactCrudState";

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse,
  SPHttpClientBatch,
} from "@microsoft/sp-http";

import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  IDropdown,
  DetailsRowCheck,
  Selection,
} from "office-ui-fabric-react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

let _softwareListColumns = [
  {
    key: "ID",
    name: "ID",
    fieldName: "ID",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
  {
    key: "Title",
    name: "Title",
    fieldName: "Title",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
  {
    key: "SoftwareName",
    name: "SoftwareName",
    fieldName: "SoftwareName",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
  {
    key: "SoftwareVendor",
    name: "SoftwareVendor",
    fieldName: "SoftwareVendor",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
  {
    key: "SoftwareDescription",
    name: "SoftwareDescription",
    fieldName: "SoftwareDescription",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
  {
    key: "SoftwareVersion",
    name: "SoftwareVersion",
    fieldName: "SoftwareVersion",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
  },
];

const textFieldStyles = { width: 300 };
const narrowDropdownStyles = { width: 100 };

export default class ReactCrud extends React.Component<
  IReactCrudProps,
  IReactCrudState
> {
  private dropdownRef: Dropdown | null = null; // Define the dropdown ref;
  private _selection: Selection;
  private _onItemsSelectionChanged = () => {
    this.setState({
      SoftwareListItem: this._selection.getSelection()[0] as ISoftwareListItem,
    });
  };

  constructor(props: IReactCrudProps, states: IReactCrudState) {
    super(props);

    this.state = {
      status: "Ready",
      SoftwareListItems: [],
      SoftwareListItem: {
        Id: 0,
        Title: "",
        SoftwareName: "",
        SoftwareVendor: "Select an Option",
        SoftwareDescription: "",
        SoftwareVersion: "",
      },
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    });
  }

  private _getListItems(): Promise<ISoftwareListItem[]> {
    const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items`;
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<ISoftwareListItem[]>;
  }

  public bindDetailsList(message: string): void {
    this._getListItems().then((listItems) => {
      this.setState({ SoftwareListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All Records have been Loaded successfully");
  }
  // @autobind

  // ADD Button
  private _onAddClick = (): void => {
    const newItem: ISoftwareListItem = {
      Id: this.state.SoftwareListItem.Id,
      Title: this.state.SoftwareListItem.Title,
      SoftwareName: this.state.SoftwareListItem.SoftwareName,
      SoftwareVendor: this.state.SoftwareListItem.SoftwareVendor,
      SoftwareDescription: this.state.SoftwareListItem.SoftwareDescription,
      SoftwareVersion: this.state.SoftwareListItem.SoftwareVersion,
    };
    const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items`;

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(this.state.SoftwareListItem),
    };

    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          this.bindDetailsList(
            "Record added and All Records were loaded Successfully"
          );
        } else {
          let errormessage = `An error has occurred i.e. ${response.status} - ${response.statusText}`;
          this.setState({ status: errormessage });
        }
      });
  };
  // Update Button
  public _onUpdateClick = (): void => {
    let id: number = this.state.SoftwareListItem.Id;
    const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items(${id})`;
    const header: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: header,
      body: JSON.stringify(this.state.SoftwareListItem),
    };
    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          this.bindDetailsList("Record updated Successfully");
        } else {
          let errormessage = `An error has occurred i.e. ${response.status} - ${response.statusText}`;
          this.setState({ status: errormessage });
        }
      });
  };
  // Delete button
  public _onDeleteClick = (): void => {
    let id: number = this.state.SoftwareListItem.Id;
    const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items(${id})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    };
    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          this.bindDetailsList("Record deleted Successfully");
        } else {
          let errormessage = `An error has occurred i.e. ${response.status} - ${response.statusText}`;
          this.setState({ status: errormessage });
        }
      });
  };
  //////////////////////////////////////////////////////////////////////////////////////
  public createAndLoadInOneGo(): Promise<ISoftwareListItem[]> {
    let promise: Promise<ISoftwareListItem[]> = new Promise<
      ISoftwareListItem[]
    >((resolve, reject) => {
      const clientBatch: SPHttpClientBatch =
        this.props.context.spHttpClient.beginBatch();
      const batchOperations: Promise<
        ISoftwareListItem | ISoftwareListItem[]
      >[] = [
        this._addListItemsAsBatch(clientBatch),
        this._getListItemsAsBatch(clientBatch),
      ];
      clientBatch
        .execute()
        .then(() => {
          return Promise.all(batchOperations);
        })
        .then((values: any) => {
          resolve(values[values.length - 1]);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
    return promise;
  }

  private _getListItemsAsBatch(
    clientBatch: SPHttpClientBatch
  ): Promise<ISoftwareListItem[]> {
    const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items`;
    return clientBatch
      .get(url, SPHttpClientBatch.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<ISoftwareListItem[]>;
  }

  public _addListItemsAsBatch(
    clientBatch: SPHttpClientBatch
  ): Promise<ISoftwareListItem> {
    let promise: Promise<ISoftwareListItem> = new Promise<ISoftwareListItem>(
      (resolve, reject) => {
        const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items`;

        const spHttpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify(this.state.SoftwareListItem),
        };
        clientBatch
          .post(url, SPHttpClientBatch.configurations.v1, spHttpClientOptions)
          .then(
            (response: SPHttpClientResponse): Promise<ISoftwareListItem> => {
              return response.json();
            }
          )
          .then((softwareListItem: ISoftwareListItem): void => {
            resolve(softwareListItem);
          })
          .catch((error: any) => {
            reject(error);
          });
      }
    );

    return promise;
  }

  public btnAddAndLoadInOneGo_Click(): void {
    const newItem: ISoftwareListItem = {
      Id: this.state.SoftwareListItem.Id,
      Title: this.state.SoftwareListItem.Title,
      SoftwareName: this.state.SoftwareListItem.SoftwareName,
      SoftwareVendor: this.state.SoftwareListItem.SoftwareVendor,
      SoftwareDescription: this.state.SoftwareListItem.SoftwareDescription,
      SoftwareVersion: this.state.SoftwareListItem.SoftwareVersion,
    };
    this.createAndLoadInOneGo().then(
      (softwareListItem: ISoftwareListItem[]) => {
        this.setState({
          SoftwareListItems: softwareListItem,
          status: "Added and Loaded Successfully as one Go",
        });
      }
    );
  }
  ////////////////////////////////////////////////////////////////////////////////
  public render(): React.ReactElement<IReactCrudProps> {
    return (
      <div className={styles.reactCrud}>
        <TextField
          label="ID"
          required={false}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.Id.toString()}
          onChange={(
            _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              Id: parseInt(newValue || "0"),
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />
        <TextField
          label="Title"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.Title}
          onChange={(
            _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              Title: newValue,
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />
        <TextField
          label="Software Name"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareName}
          onChange={(
            _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              softwareName: newValue,
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />
        <TextField
          label="Software Version"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareVersion}
          onChange={(
            _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              SoftwareVersion: newValue,
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />
        <TextField
          label="Description"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareDescription}
          onChange={(
            _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              SoftwareDescription: newValue,
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />

        <Dropdown
          componentRef={(ref) => {
            this.dropdownRef = ref as any as Dropdown;
          }}
          placeHolder="Select an option"
          label="Software Vendor"
          options={[
            { key: "Sun", text: "Sun" },
            { key: "Microsoft", text: "Microsoft" },
            { key: "Google", text: "Google" },
          ]}
          defaultSelectedKey={this.state.SoftwareListItem.SoftwareVendor}
          required
          style={narrowDropdownStyles}
          onChanged={(option) => {
            const newItem = {
              ...this.state.SoftwareListItem,
              SoftwareVendor: option.text,
            };
            this.setState({ SoftwareListItem: newItem });
          }}
        />

        <p className={styles.title}>
          <PrimaryButton text="Add" title="Add" onClick={this._onAddClick} />
          <PrimaryButton
            text="Update"
            title="Update"
            onClick={this._onUpdateClick}
          />
          <PrimaryButton
            text="Delete"
            title="Delete"
            onClick={this._onDeleteClick}
          />
          <PrimaryButton
            text="Add &Load In One Go"
            title="Add"
            onClick={this.btnAddAndLoadInOneGo_Click}
          />
        </p>
        <div id="divStatus">{this.state.status}</div>
        <div>
          <DetailsList
            items={this.state.SoftwareListItems}
            columns={_softwareListColumns}
            setKey="Id"
            checkboxVisibility={CheckboxVisibility.onHover}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.single}
            compact={true}
            selection={this._selection}
          />
        </div>
      </div>
    );
  }
}
