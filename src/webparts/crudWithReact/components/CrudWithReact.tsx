import * as React from 'react';
import styles from './CrudWithReact.module.scss';
import { ICrudWithReactProps } from './ICrudWithReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICrudWithReactState } from './ICrudWithReactState';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { 
  TextField,
  PrimaryButton,
  Selection,
  IDropdown,
  IDropdownStyles,
  Dropdown,
  DetailsList,
  CheckboxVisibility,
  SelectionMode,
  DetailsListLayoutMode,
  ITextFieldStyles
 } from 'office-ui-fabric-react';
import { ISoftwareListItem } from './ISoftwareListItem';

 let _softwareListColumns = [
   {
     key: 'ID',
     name: 'ID',
     fieldName: 'ID',
     minWidth: 50,
     maxWidth: 100,
     isResizable: true
   },
   {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'SoftwareName',
    name: 'SoftwareName',
    fieldName: 'SoftwareName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'SoftwareVendor',
    name: 'SoftwareVendor',
    fieldName: 'SoftwareVendor',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'SoftwareVersion',
    name: 'SoftwareVersion',
    fieldName: 'SoftwareVersion',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'SoftwareDescription',
    name: 'SoftwareDescription',
    fieldName: 'SoftwareDescription',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
 ];

 const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
 const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 100 } };
 const narrowDropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };


export default class CrudWithReact extends React.Component<ICrudWithReactProps, ICrudWithReactState> {
  
  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    this.setState({
      SoftwareListItem: (this._selection.getSelection()[0] as ISoftwareListItem)
    });
  }

  constructor(props: ICrudWithReactProps, state: ICrudWithReactState) {
    super(props);

    this.state = {
      status: 'Ready',
      SoftwareListItems: [],
      SoftwareListItem: {
        Id: 0,
        Title: "",
        SoftwareName: "",
        SoftwareDescription: "",
        SoftwareVendor: "Select an option",
        SoftwareVersion: ""
      }
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged
    });
  }

  btnAdd_click = () => {
    
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicroSoftware')/items"



    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(this.state.SoftwareListItem)
    };
    
    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status === 201) {
        
        this.bindDetailsList("Record added and All Records were loaded Successfully");
      
      } else {

        let errormessage: string = "An error has occured i.e. " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});
      }
    })
  }

  btnUpdate_click = () => {
    let id: number = this.state.SoftwareListItem.Id;

    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicroSoftware')/items("+ id +")";

    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(this.state.SoftwareListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status === 204) {
        this.bindDetailsList("Record Updated and All records were loaded Successfully");
      } else {
        let errormessage: string = "An error has occured i.e. " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});
      }
    })
  }

  btnDelete_click = () => {
    let id: number = this.state.SoftwareListItem.Id;

    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicroSoftware')/items("+ id +")";

    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*"
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status === 204) {
        this.bindDetailsList("Record Deleted and All records were loaded Successfully");
      } else {
        let errormessage: string = "An error has occured i.e. " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});
      }
    })
  }

  private _getListItems():Promise<ISoftwareListItem[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicroSoftware')/items";
    return this.props.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {
      return json.value;
    }) as Promise<ISoftwareListItem[]>;
  }

  public bindDetailsList(message: string) : void {
    this._getListItems().then(listItems => {
      this.setState({ SoftwareListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All Records have been loaded Successfully");
  }
  
  

  public render(): React.ReactElement<ICrudWithReactProps> {
    
    const dropdownRef = React.createRef<IDropdown>();

    return (
      <div className={styles.crudWithReact}>
        <TextField 
          label="ID" 
          required={false} 
          value={(this.state.SoftwareListItem.Id).toString()} 
          styles={textFieldStyles} 
          onChange={(ev, value) => {
            this.setState(prevState => {
              return {...prevState, SoftwareListItem: {...prevState.SoftwareListItem, Id: +value}}
          });
          }}
        />

        <TextField 
          label="Software Title" 
          required={true} 
          value={(this.state.SoftwareListItem.Title)} 
          styles={textFieldStyles} 
          onChange={(ev, value) => {
            this.setState(prevState => {
              return {...prevState, SoftwareListItem: {...prevState.SoftwareListItem, Title: value}}
          });
          }}
        />

        <TextField 
          label="Software Name" 
          required={true} 
          value={(this.state.SoftwareListItem.SoftwareName)} 
          styles={textFieldStyles} 
          onChange={(ev, value) => {
            this.setState(prevState => {
              return {...prevState, SoftwareListItem: {...prevState.SoftwareListItem, SoftwareName: value}}
          });
          }}
        />

        <TextField 
          label="Software Description" 
          required={true} 
          value={(this.state.SoftwareListItem.SoftwareDescription)} 
          styles={textFieldStyles} 
          onChange={(ev, value) => {
            this.setState(prevState => {
              return {...prevState, SoftwareListItem: {...prevState.SoftwareListItem, SoftwareDescription: value}}
          });
          }}
        />

        <TextField 
          label="Software Version" 
          required={true} 
          value={(this.state.SoftwareListItem.SoftwareVersion)} 
          styles={textFieldStyles} 
          onChange={(ev, value) => {
            this.setState(prevState => {
              return {...prevState, SoftwareListItem: {...prevState.SoftwareListItem, SoftwareVersion: value}}
          });
          }}
        />

        <Dropdown 
          label="Software Vendor"
          componentRef={dropdownRef}
          placeholder = "Select an option"
          options={[
            { key: 'Microsoft', text: 'Microsoft' },
            { key: 'Sun', text: 'Sun' },
            { key: 'Oracle', text: 'Oracle' },
            { key: 'Google', text: 'Google' },
          ]}
          defaultSelectedKey={this.state.SoftwareListItem.SoftwareVendor}
          required
          styles={narrowDropdownStyles}
          onChange={(ev, value) => {
            this.state.SoftwareListItem.SoftwareVendor=value.text
          }}
        />

        <p className={styles.title}>
          <PrimaryButton text='Add' title='Add' onClick={this.btnAdd_click}/>
          <PrimaryButton text='Update' title='Update' onClick={this.btnUpdate_click}/>
          <PrimaryButton text='Delete' title='Delete' onClick={this.btnDelete_click}/>
        </p>

        <div id="divStatus">
          {this.state.status}
        </div>

        <DetailsList 
          items={this.state.SoftwareListItems}
          columns={_softwareListColumns}
          setKey='id'
          checkboxVisibility={CheckboxVisibility.always}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={ true }
          selection={this._selection}
        />
      </div>
    );
  }
}
