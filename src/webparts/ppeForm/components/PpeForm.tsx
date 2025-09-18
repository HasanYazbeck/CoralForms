import * as React from "react";
// import { useState } from "react";

// Components
import { DefaultPalette } from "@fluentui/react";
import type { IPpeFormWebPartProps, IPpeFormWebPartState, } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
// import CircularProgress from "@mui/material/CircularProgress";

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";

// Classes
import { SPCrudOperations } from '../../../Classes/SPCrudOperations';
import { IUser } from "../../../Interfaces/IUser";
import { SPHelpers } from '../../../Classes/SPHelpers';
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";

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

  private spCrudOperations: SPCrudOperations;
  private spHelpers: SPHelpers = new SPHelpers();

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
      PPEItems: [],
      CoralFormsList: { Id: "" },
    };
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
    const filtered = rows.filter(r => !r.Selected);
    this.setState({ PPEItemsRows: filtered });
  }

  // private removeRow = (index: number) => {
  //   const rows = this.state.PPEItemsRows ? [...this.state.PPEItemsRows] : [];
  //   rows.splice(index, 1);
  //   this.setState({ PPEItemsRows: rows });
  // }

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

    this.getCoralFormsList();
    this.getPPEItemsDetails();
  }

  public getCoralFormsList = async (): Promise<void> => {
    let result: ICoralFormsList = { Id: "" };
    try {
      const searchFormName = "PERSONAL PROTECTIVE EQUIPMENT";
      const searchEscaped = searchFormName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created` +
        `&$filter=substringof('${searchEscaped}', Title)`;
      this.spCrudOperations = new SPCrudOperations(this.props.context.spHttpClient,
        this.props.context.pageContext.web.absoluteUrl, 'CoralFormsList', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this.props.Users !== undefined && this.props.Users.length > 0 ? this.props.Users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }

              result = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                hasInstructionForUse: obj.hasInstructionForUse !== undefined ? obj.hasInstructionForUse : undefined,
                hasWorkflow: obj.hasWorkflow !== undefined ? obj.hasWorkflow : undefined,
                Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
              }
            }
          });
          this.setState({ CoralFormsList: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

  public getPPEItems = async (): Promise<void> => {
    const result: IPPEItem[] = [];
    try {
      const query: string = `?$select=Id,Title,Required,hasInstructionForUse,hasWorkflow,Created`;
      // `PPEDetails/Id,PPEDetails/Title&$expand=PPEDetails`;
      this.spCrudOperations = new SPCrudOperations(this.props.context.spHttpClient,
        this.props.context.pageContext.web.absoluteUrl, 'PPEItems', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this.props.Users !== undefined && this.props.Users.length > 0 ? this.props.Users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {// Convert string to Date first
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }
              const temp: IPPEItem = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                // hasInstructionForUse: obj.hasInstructionForUse !== undefined ? obj.hasInstructionForUse : undefined,
                // hasWorkflow: obj.hasWorkflow !== undefined ? obj.hasWorkflow : undefined,
                Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
                Required: obj.Required !== undefined ? obj.Required : undefined,
                // PPEDetails: obj.PPEDetails !== undefined ? obj.PPEDetails : undefined,
              }


              console.log(temp);
              // Get PPEDetails for each one (Types, Sizez )
              result.push(temp);
            }
          });
          this.setState({ PPEItems: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

  public getPPEItemsDetails = async (): Promise<void> => {
    const result: IPPEItemDetails[] = [];
    try {
      const query: string = `?$select=Id,Title,PPEItem,Types,Sizes,Created,` +
        `PPEItem/Id,PPEItem/Title,PPEItem/Required,Types/Id,Types/Title&$expand=PPEItem,Types`;
      this.spCrudOperations = new SPCrudOperations(this.props.context.spHttpClient,
        this.props.context.pageContext.web.absoluteUrl, 'PPEItemsDetails', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this.props.Users !== undefined && this.props.Users.length > 0 ? this.props.Users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {// Convert string to Date first
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }
              const temp: IPPEItemDetails = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                PPEItem: obj.PPEItem !== undefined ? {
                  Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
                  Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,

                  // Required: obj.PPEItem.Required !== undefined ? obj.PPEItem.Required : undefined,
                  // Brands: obj.PPEItem.Brands !== undefined && obj.PPEItem.Brands !== null ? obj.PPEItem.Brands.split(",") : undefined,

                } : undefined,
                Types: obj.Types !== undefined && obj.Types !== null ? obj.Types : undefined,
                Sizes: obj.Sizes !== undefined && obj.Sizes !== null ? obj.Sizes.split(",") : undefined,
              };
              console.log(temp);
              // Get PPEDetails for each one (Types, Sizez )
              result.push(temp);
            }
          });
          this.setState({ PPEItems: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

  public render(): React.ReactElement<IPpeFormWebPartProps> {
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

    if (this.props.IsLoading) {
      return (
        <div className={styles.loadingContainer}>
          {/* <CircularProgress variant="indeterminate" /> */}
        </div>
      );
    }
    return (
      <div>
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

                <div className="table-responsive">
                  <table id="itemsTable" className="table table-bordered">
                    <thead className="thead-light">
                      <tr>
                        <th className="align-items-center text-center" style={{ width: 80 }}>Select</th>
                        <th className="align-items-center">Item</th>
                        <th className="text-center">Brand</th>
                        <th className="text-center align-middle justify-content-center">Required</th>
                        <th className="">Specific Details</th>
                        <th className="text-center align-middle" style={{ width: 80 }}>Qty</th>
                        <th className="text-center align-items-center" style={{ width: 120 }}>Size</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const defaultRows = [this.createEmptyRow()];
                        const rows = this.state.PPEItemsRows && this.state.PPEItemsRows.length > 0 ? this.state.PPEItemsRows : defaultRows;
                        return rows.map((row, i) => (
                          <tr key={i}>
                              <td className="text-center align-middle">
                                <input type="checkbox" checked={!!row.Selected} onChange={(ev) => this.onRowChange(i, 'Selected', ev.currentTarget.checked)} />
                              </td>
                              <td>
                                <TextField value={row.Item} onChange={(ev, val) => this.onRowChange(i, 'Item', val || '')} underlined={true} />
                              </td>
                              <td>
                                <TextField value={row.Brands || ''} onChange={(ev, val) => this.onRowChange(i, 'Brands', val || '')} underlined={true} />
                              </td>
                              <td className={`table-secondary text-center align-middle ${styles["justify-items-center"]}`}>
                                <Checkbox checked={!!row.Required} onChange={(ev, checked) => this.onRowChange(i, 'Required', !!checked)} />
                              </td>
                              <td className="table-secondary">
                                <TextField value={row.Details} onChange={(ev, val) => this.onRowChange(i, 'Details', val || '')} underlined={true} />
                              </td>
                              <td className="table-secondary text-center align-middle">
                                <TextField value={row.Qty} onChange={(ev, val) => this.onRowChange(i, 'Qty', val || '')} underlined={true} />
                              </td>
                              <td>
                                <TextField value={row.Size} onChange={(ev, val) => this.onRowChange(i, 'Size', val || '')} underlined={true} />
                              </td>
                            </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

          </Stack>
        </form>
      </div>
    );
  }

}
