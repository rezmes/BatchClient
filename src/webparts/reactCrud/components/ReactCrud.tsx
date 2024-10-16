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
} from "@microsoft/sp-http";

import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  // Dropdown,
  IDropdown,
  // IDropdownOption,
  // ITextFieldStyles,
  // IDropdownStyles,
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

// // const textFieldStyles: Partial<ITextFieldStyles> = {
// //   fieldGroup: {
// //     width: 300,
// //   },
// // };
// // const narrowDropdownStyles: Partial<IDropdownStyles> = {
// //   fieldGroup: {
// //     width: 100,
// //   },
// // };

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
  //   //  @autobind;
  //   // private _onAddClick = () => {
  //   //   const newItem: ISoftwareListItem = {
  //   //     Id: this.state.SoftwareListItem.Id,
  //   //     Title: this.state.SoftwareListItem.Title,
  //   //     SoftwareName: this.state.SoftwareListItem.SoftwareName,
  //   //     SoftwareVendor: this.state.SoftwareListItem.SoftwareVendor,
  //   //     SoftwareDescription: this.state.SoftwareListItem.SoftwareDescription,
  //   //     SoftwareVersion: this.state.SoftwareListItem.SoftwareVersion,
  //   //   };

  //   //   const url: string = `${this.props.siteUrl}/_api/web/lists/GetByTitle('SoftwareCatalog')/items`;
  //   //   const spHttpClientOptions: ISPHttpClientOptions = {
  //   //     body: JSON.stringify(this.state.SoftwareListItem),
  //   //   };
  //   //   this.props.context.spHttpClient
  //   //     .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
  //   //     .then((response: SPHttpClientResponse) => {
  //   //       if (response.status === 201) {
  //   //         this.bindDetailsList(
  //   //           "Record added and All Records were loaded Successfully"
  //   //         );
  //   //       } else {
  //   //         let errormessage: string =
  //   //           "An error has occured i.e. " +
  //   //           response.status +
  //   //           " - " +
  //   //           response.statusText;
  //   //         this.setState({ status: errormessage });
  //   //       }
  //   //     });
  //   // };

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
        {/*                <TextField
          label="ID"
          required={false}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.Id.toString()}
          onChanged={(e) => {
            this.state.SoftwareListItem.Id = e;
          }}
        />
         <TextField
          label="Title"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.Title}
          onChanged={(e) => {
            this.state.SoftwareListItem.Title = e;
          }}
        />
        <TextField
          label="SoftwareName"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareName}
          onChanged={(e) => {
            this.state.SoftwareListItem.SoftwareName = e;
          }}
        />
        <TextField
          label="SoftwareVersion"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareVersion}
          onChanged={(e) => {
            this.state.SoftwareListItem.SoftwareVersion = e;
          }}
        />
        <TextField
          label="description"
          required={true}
          style={textFieldStyles}
          value={this.state.SoftwareListItem.SoftwareDescription}
          onChanged={(e) => {
            this.state.SoftwareListItem.SoftwareDescription = e;
          }}
        /> */}
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
