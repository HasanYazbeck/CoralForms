import * as React from "react";
// import { useState } from "react";

// Components
import { DefaultPalette } from "@fluentui/react";
import type { IPpeFormWebPartProps, IPpeFormWebPartState, } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { ComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DetailsList, IColumn, Selection, SelectionMode } from '@fluentui/react';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
// import CircularProgress from "@mui/material/CircularProgress";

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";

const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeTertiary,
    display: "inline",
  },
};

const datePickerStyles = mergeStyleSets({
  root: { selectors: { '> *': { marginBottom: 15 } } },
  control: { maxWidth: 300, marginBottom: 15 },
});

export default class PpeForm extends React.Component<IPpeFormWebPartProps, IPpeFormWebPartState> {

  // private spCrudOperations: SPCrudOperations;
  // private spHelpers: SPHelpers = new SPHelpers();
  private _selection: Selection;

  constructor(props: IPpeFormWebPartProps) {
    super(props);
    this.state = {
      JobTitle: "",
      Department: "",
      Division: "",
      Company: "",
      Employee: [],
      EmployeeId: undefined,
      Submitter: [],
      Requester: [],
      isReplacementChecked: false,
      PPEItems: this.props.PPEItems,
      PPEItemsDetails: this.props.PPEItemDetails,
      CoralFormsList: { Id: "" },
    };
    this._selection = new Selection();
  }

  // Dynamic table rows for items requested
  private createEmptyRow = () => ({ Item: '', Brands: '', Required: false, Details: '', Qty: '', Size: '', Selected: false });

  private addRow = () => {
    // If PPEItemsRows already exists use it; otherwise start from the currently visible default row
    const existing = this.state.PPEItemsRows && this.state.PPEItemsRows.length > 0 ? [...this.state.PPEItemsRows] : [this.createEmptyRow()];
    existing.push(this.createEmptyRow());
    this.setState({ PPEItemsRows: existing });
  }

  private deleteSelectedRows = () => {
    const rows = this.state.PPEItemsRows && this.state.PPEItemsRows.length > 0 ? [...this.state.PPEItemsRows] : [];
    const selectedIndices = this._selection ? this._selection.getSelectedIndices() : [];
    if (selectedIndices && selectedIndices.length > 0) {
      // remove items by indices
      const filtered = rows.filter((r, idx) => selectedIndices.indexOf(idx) === -1);
      this._selection.setAllSelected(false);
      this.setState({ PPEItemsRows: filtered });
    } else {
      const filtered = rows.filter(r => !r.Selected);
      this.setState({ PPEItemsRows: filtered });
    }
  }

  private onRowChange = (index: number, field: string, value: any) => {
    // Ensure we have a rows array to update (handles the fallback visible row)
    const rows = this.state.PPEItemsRows && this.state.PPEItemsRows.length > 0 ? [...this.state.PPEItemsRows] : [this.createEmptyRow()];
    // Grow array if needed
    while (rows.length <= index) {
      rows.push(this.createEmptyRow());
    }
    // @ts-ignore - dynamic field write
    rows[index][field] = value;
    this.setState({ PPEItemsRows: rows });
  }

  // State handlers

  // Handle Employee selection
  private handleEmployeeChange = (items: IPersonaProps[]) => {
    if (items && items.length > 0) {
      const selected = items[0];
      // Find the user in the Users prop
      const user = this.props.Users.find(u => u.id === selected.id);
      this.setState({
        Employee: [selected],
        EmployeeId: selected.id,
        JobTitle: user?.jobTitle || "",
        Department: user?.department || "",
        Company: user?.company || "" // Make sure your user object has a company property
      });
    } else {
      this.setState({
        Employee: [],
        EmployeeId: undefined,
        JobTitle: "",
        Department: "",
        Company: ""
      });
    }
  };

  // Handle Requester selection
  private handleNewRequestChange = (ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    if (checked) {
      this.setState({ isReplacementChecked: false });
    }
  };

  private handleReplacementChange = (ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    this.setState({ isReplacementChecked: !!checked });
  };

  async componentDidMount(): Promise<void> {

    // Get current user from context or Users prop
    const currentUserEmail = this.props.context.pageContext.user.email;
    const currentUser = this.props.Users.find(u => u.email === currentUserEmail);

    if (currentUser) {
      this.setState({
        Submitter: [{
          text: currentUser.displayName,
          secondaryText: currentUser.email,
          id: currentUser.id
        }]
      });
    }

    // await this.getCoralFormsList();
    // await this.getPPEItemsDetails();
  }

  public render(): React.ReactElement<IPpeFormWebPartProps> {

      if (this.props.IsLoading) {
        return (
          <div className={styles.loadingContainer}>
            <Spinner label={"Preparing PPE form — fresh items coming right up!"} size={SpinnerSize.large} />
          </div>
        );
      }

    const delayResults = false;
  
    const logoUrl = `${this.props.context.pageContext.web.absoluteUrl}/SiteAssets/coral-logo.png`;
    const peopleList: IPersonaProps[] = this.props.Users.map(user => ({
      text: user.displayName || "",
      secondaryText: user.email || "",
      id: user.id,
    }));
    // const themeColor = this.props.ThemeColor || DefaultPalette.themePrimary;
    const filterPromise = (personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
      if (delayResults) {
        return convertResultsToPromise(personasToReturn);
      } else {
        return personasToReturn;
      }
    };
    const onFilterChanged = (
      filterText: string,
      currentPersonas: IPersonaProps[],
      limitResults?: number,
    ): IPersonaProps[] | Promise<IPersonaProps[]> => {
      if (filterText) {
        let filteredPersonas: IPersonaProps[] = filterPersonasByText(filterText);

        filteredPersonas = removeDuplicates(filteredPersonas, currentPersonas);
        filteredPersonas = limitResults ? filteredPersonas.slice(0, limitResults) : filteredPersonas;
        return filterPromise(filteredPersonas);
      } else {
        return [];
      }
    };
    const filterPersonasByText = (filterText: string): IPersonaProps[] => {
      return peopleList.filter(item => doesTextStartWith(item.text as string, filterText));
    };
    function doesTextStartWith(text: string, filterText: string): boolean {
      return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }

    function removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
      return personas.filter(persona => !listContainsPersona(persona, possibleDupes));
    }

    function listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
      if (!personas || !personas.length || personas.length === 0) {
        return false;
      }
      return personas.filter(item => item.text === persona.text).length > 0;
    }

    function convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
      return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }

    function onInputChange(input: string): string {
      const outlookRegEx = /<.*>/g;
      const emailAddress = outlookRegEx.exec(input);

      if (emailAddress && emailAddress[0]) {
        return emailAddress[0].substring(1, emailAddress[0].length - 1);
      }

      return input;
    }


    return (
      <div className={styles.ppeFormBackground} >
        <form>
          <div className={styles.formHeader}>
            <img src={logoUrl} alt="Logo" className={styles.formLogo} />
            <span className={styles.formTitle}>
              PERSONAL PROTECTIVE EQUIPMENT (PPE) REQUISITION FORM
            </span>
          </div>

          <Stack horizontal styles={stackStyles} >

            <div className="row">
              <div className="form-group col-md-6">
                <NormalPeoplePicker
                  label={"Employee Name"}
                  itemLimit={1}
                  onResolveSuggestions={onFilterChanged}
                  className={'ms-PeoplePicker'}
                  key={'normal'}
                  removeButtonAriaLabel={'Remove'}
                  inputProps={{
                    onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                    onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                    'aria-label': 'People Picker',
                  }}
                  onInputChange={onInputChange}
                  resolveDelay={300}
                  disabled={false}
                  onChange={this.handleEmployeeChange}
                // required={true}
                // onGetErrorMessage={this.onGetErrorMessage}
                />
              </div>

              <div className="form-group col-md-6">
                <DatePicker
                  disabled
                  value={new Date(Date.now())}
                  label="Date Requested"
                  className={datePickerStyles.control}
                  // DatePicker uses English strings by default. For localized apps, you must override this prop.
                  strings={defaultDatePickerStrings}
                />
              </div>
            </div>
            <div className="row">
              <div className="form-group col-md-6">
                <TextField label="Job Title"
                  value={this.state.JobTitle} />
              </div>

              <div className="form-group col-md-6">
                <TextField label="Department" value={this.state.Department} />
              </div>
            </div>

            <div className="row">
              <div className="form-group col-md-6">
                <TextField label="Division" />
              </div>

              <div className="form-group col-md-6">
                <TextField label="Company" />
              </div>
            </div>

            <div className="row">
              <div className="form-group col-md-6">
                <NormalPeoplePicker
                  label={"Requester Name"}
                  itemLimit={1}
                  onResolveSuggestions={onFilterChanged}
                  className={'ms-PeoplePicker'}
                  key={'normal'}
                  removeButtonAriaLabel={'Remove'}
                  inputProps={{
                    onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                    onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                    'aria-label': 'People Picker',
                  }}
                  onInputChange={onInputChange}
                  resolveDelay={300}
                  disabled={false}
                  onChange={this.handleEmployeeChange}
                // required={true}
                // onGetErrorMessage={this.onGetErrorMessage}
                />
              </div>

              <div className="form-group col-md-6">
                <NormalPeoplePicker
                  label={"Submitter Name"}
                  itemLimit={1}
                  onResolveSuggestions={onFilterChanged}
                  className={'ms-PeoplePicker'}
                  key={'normal'}
                  removeButtonAriaLabel={'Remove'}
                  inputProps={{
                    onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                    onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                    'aria-label': 'People Picker',
                  }}
                  onInputChange={onInputChange}
                  resolveDelay={300}
                  disabled={true}
                  selectedItems={this.state.Submitter}
                />
              </div>
            </div>

            <div className={`row  ${styles.mt10}`}>
              <div className="form-group col-md-12 d-flex justify-content-between" >
                <Label htmlFor={""}>Reason for Request</Label>

                <Checkbox label="New Request"
                  className="align-items-center"
                  checked={!this.state.isReplacementChecked}
                  onChange={this.handleNewRequestChange} />

                <Checkbox label="Replacement"
                  className="align-items-center"
                  checked={this.state.isReplacementChecked}
                  onChange={this.handleReplacementChange} />

                <TextField placeholder="Reason" disabled={!this.state.isReplacementChecked} />
              </div>
            </div>

          </Stack>

          <Separator />

          <div className="mb-2 text-center">
            <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '1.05rem' }}>
              Please complete the table below in the blank spaces; grey spaces are for administrative use only.
            </small>
          </div>

          <Stack horizontal styles={stackStyles} >

            <div className="row">
              <div className="form-group col-md-12">
                {(() => {
                  const commandBarItems: ICommandBarItemProps[] = [
                    {
                      key: 'addItem',
                      text: 'Add Item',
                      iconProps: { iconName: 'Add' },
                      onClick: this.addRow,
                    },
                    {
                      key: 'deleteSelected',
                      text: 'Delete',
                      iconProps: { iconName: 'Delete' },
                      onClick: this.deleteSelectedRows,
                    }
                  ];

                  return <CommandBar items={commandBarItems} styles={{ root: { marginBottom: 8 } }} />;
                })()}

                {(() => {
                  const defaultRows = [this.createEmptyRow()];
                  const rows = this.state.PPEItemsRows && this.state.PPEItemsRows.length > 0 ? this.state.PPEItemsRows : defaultRows;

                  // Decorate items with __index so Selection can map back
                  const items = rows.map((r, idx) => ({ ...r, __index: idx } as any));

                  // Build distinct options for the Item combobox from PPEItems (IPPEItemDetails.Title or Title)
                  const itemTitles: string[] = (this.props.PPEItemDetails || []).map((p: any) => {
                    return (p && p.PPEItem && p.PPEItem.Title) ? p.PPEItem.Title : (p && p.Title ? p.Title : undefined);
                  }).filter(Boolean) as string[];
                  const distinctTitles = Array.from(new Set(itemTitles));
                  const itemOptions: IComboBoxOption[] = distinctTitles.map(t => ({ key: t, text: t }));

                  const columns: IColumn[] = [
                    {
                      key: 'columnItem',
                      name: 'Item',
                      fieldName: 'Item',
                      minWidth: 150,
                      isResizable: true,
                      onRender: (item: any) => (
                        <ComboBox
                          allowFreeform
                          autoComplete="on"
                          selectedKey={item.Item || undefined}
                          options={itemOptions}
                          onChange={(ev, option, index, value) => {
                            const newVal = option ? option.key : value;
                            this.onRowChange(item.__index, 'Item', newVal || '');
                          }}
                        />
                      )
                    },
                    {
                      key: 'columnBrand',
                      name: 'Brand',
                      fieldName: 'Brands',
                      minWidth: 120,
                      isResizable: true,
                      onRender: (item: any) => <TextField value={item.Brands || ''} onChange={(ev, val) => this.onRowChange(item.__index, 'Brands', val || '')} underlined={true} />
                    },
                    {
                      key: 'columnRequired',
                      name: 'Required',
                      className: `text-center align-middle ${styles.justifyItemsCenter}`,
                      fieldName: 'Required',
                      minWidth: 90,
                      maxWidth: 120,
                      isResizable: false,
                      onRender: (item: any) => <div className={`table-secondary ${styles.justifyItemsCenter}`}><Checkbox checked={!!item.Required} onChange={(ev, checked) => this.onRowChange(item.__index, 'Required', !!checked)} /></div>
                    },
                    {
                      key: 'columnDetails',
                      name: 'Specific Details',
                      fieldName: 'Details',
                      minWidth: 180,
                      isResizable: true,
                      onRender: (item: any) => <div className="table-secondary"><TextField value={item.Details} onChange={(ev, val) => this.onRowChange(item.__index, 'Details', val || '')} underlined={true} /></div>
                    },
                    {
                      key: 'columnQty',
                      name: 'Qty',
                      fieldName: 'Qty',
                      minWidth: 70,
                      maxWidth: 90,
                      isResizable: false,
                      onRender: (item: any) => <div className="table-secondary text-center align-middle"><TextField value={item.Qty} onChange={(ev, val) => this.onRowChange(item.__index, 'Qty', val || '')} underlined={true} /></div>
                    },
                    {
                      key: 'columnSize',
                      name: 'Size',
                      fieldName: 'Size',
                      minWidth: 100,
                      maxWidth: 140,
                      isResizable: true,
                      onRender: (item: any) => <TextField value={item.Size} onChange={(ev, val) => this.onRowChange(item.__index, 'Size', val || '')} underlined={true} />
                    }
                  ];

                  return (
                    <DetailsList
                      items={items}
                      columns={columns}
                      selection={this._selection}
                      selectionMode={SelectionMode.multiple}
                      setKey="ppeItemsList"
                      layoutMode={0}
                      isHeaderVisible={true}
                      className={styles.detailsListHeaderCenter}
                    />
                  );
                })()}
              </div>
            </div>

          </Stack>
        </form>
      </div>
    );
  }

}
