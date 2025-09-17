import * as React from "react";
// import { useState } from "react";
import type { IPpeFormWebPartProps, IPpeFormWebPartState, } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";
// import CircularProgress from "@mui/material/CircularProgress";
import { DefaultPalette } from "@fluentui/react";

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
      isReplacementChecked: false
    };
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

  public render(): React.ReactElement<IPpeFormWebPartProps> {
    const delayResults = false;
    const peopleList: IPersonaProps[] = this.props.Users.map(user => ({
      text: user.displayName || "",
      secondaryText: user.email || "",
      id: user.id,
    }));

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

    if (this.props.IsLoading) {
      return (
        <div className={styles.loadingContainer}>
          {/* <CircularProgress variant="indeterminate" /> */}
        </div>
      );
    }

    // const logoUrl = `${this.props.Context.pageContext.web.absoluteUrl}/SiteAssets/coral-logo.png`;
    const logoUrl = `https://softflowcloud.sharepoint.com/sites/DEMO/DanielTest/SiteAssets/coral-logo.png`;

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


          <Stack horizontal styles={stackStyles} >

            <div className="row">
              <div className="form-group col-md-12">
                <Label htmlFor={"itemsTable"}>Items Requested</Label>

                <div className="table-responsive">
                  <table id="itemsTable" className="table table-bordered">
                    <thead className="thead-light">
                      <tr>
                        <th className="align-items-center">Item</th>
                        <th className="align-items-center">Required</th>
                        <th className="align-items-center">Specific Details</th>
                        <th className="align-items-center" style={{width: 80}}>Qty</th>
                        <th className="align-items-center" style={{width: 120}}>Size</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[1,2,3,4,5].map((i) => (
                        <tr key={i}>
                          <td>
                            <TextField placeholder={`Item ${i}`} underlined={true} />
                          </td>
                          <td>
                            <Checkbox className="align-items-center"/>
                          </td>
                          <td>
                            <TextField placeholder={`Details for item ${i}`} />
                          </td>
                          <td>
                            <TextField placeholder="Qty" underlined={true} />
                          </td>
                          <td>
                            <TextField placeholder="Size" underlined={true} />
                          </td>
                        </tr>
                      ))}
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

  componentDidMount() {
    // Get current user from context or Users prop
    const currentUserEmail = this.props.Context.pageContext.user.email;
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
  }
}
