import * as React from "react";
// import { useState } from "react";
import type {
  IPpeFormWebPartProps,
  IPpeFormWebPartState,
} from "./IPpeFormProps";
import { IPPEForm } from "../../../Interfaces/IEmployeeProps";

import { TextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from "@microsoft/sp-http";
import { IPersonaProps } from "@fluentui/react";
// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";
import CircularProgress from "@mui/material/CircularProgress";
import { DefaultPalette } from "@fluentui/react";

const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.white,
  },
};

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
    };
  }

  // private handleEmployeeChange = (selectedOption: { label: string;id: string;}) => {
  //   if (selectedOption) {
  //     const selectedUser = this.props.Users.find(
  //       (u) => u.id === selectedOption.id
  //     );
  //     if (selectedUser) {
  //       this.setState({
  //         selectedEmployeeId: selectedOption.id,
  //         jobTitle: selectedUser.jobTitle || "",
  //         department: selectedUser.department || "",
  //         // division: selectedUser.division || "",
  //         // company: selectedUser.company || ""
  //       });
  //     }
  //   }
  // };
  private async fetchUserId(loginName: string): Promise<number | null> {
    const userLookupUrl = `
      https://softflowcloud.sharepoint.com/sites/Presales/HRDept/_api/web/siteusers?$filter=Email eq '${encodeURIComponent(
        loginName
      )}'&$select=Id`;

    try {
      const response = await this.props.Context.spHttpClient.get(userLookupUrl,SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const userData = await response.json();
        if (userData.value && userData.value.length > 0) {
          return userData.value[0].Id;
        } else {
          console.error("User not found:", loginName);
          return null;
        }
      } else {
        console.error("Error fetching user ID:", response.statusText);
        return null;
      }
    } catch (error) {
      console.error("Error fetching user ID:", error);
      return null;
    }
  }
  private handleEmployeeChange = async (items: IPersonaProps[]): Promise<void> => {
    if (items && items.length > 0) {
      const selected = items[0];
      const employeeEmail = selected.secondaryText;
      const employeeName = selected.text;

      if (employeeEmail) {
        const EmployeeId = await this.fetchUserId(employeeEmail);
        if (EmployeeId !== null) {
          this.setState({
            Employee: [{ text: employeeName, secondaryText: employeeEmail }],
            EmployeeId: EmployeeId.toString(),
          });
        }
      }
    } else {
      this.setState({
        Employee: [],
        EmployeeId: "",
      });
    }
  }
  public render(): React.ReactElement<IPPEForm> {

     const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.Context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.context.msGraphClientFactory,
      spHttpClient: this.context.spHttpClient,
    };


    if (this.props.IsLoading) {
      return (
        <div className={styles.loadingContainer}>
          <CircularProgress variant="indeterminate" />
        </div>
      );
    }

    return (
      <div>
        <Stack horizontal styles={stackStyles} tokens={{ childrenGap: 5 }}>
          <form>
            <div className="row">
              <div className="form-group col-md-6">
                  <PeoplePicker
                    context={peoplePickerContext}
                    personSelectionLimit={1}
                    defaultSelectedUsers={this.state.Employee.map((emp) => emp.text || "")}
                    showtooltip={true}
                    required={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    onChange={this.handleEmployeeChange}/>

                {/* <AutoComplete label="Employee Name"
                  keyId="1"
                  className={styles.formField}
                  options={this.props.Users.map((user) => ({
                    label: user.displayName || "",
                    id: user.id,
                  }))}
                // OnChange={this.handleEmployeeChange}
                /> */}
              </div>
            </div>
            <div className="row">
              <div className="form-group col-md-6">
                <TextField label="Department" />
              </div>

              <div className="form-group col-md-6">
                <TextField label="Company" />
              </div>
            </div>

            <div className="row">
              <div className="form-group col-md-6">
                <TextField label="Division" />
              </div>

              <div className="form-group col-md-6">
                <TextField label="Job Title" />
              </div>
            </div>

          </form>
        </Stack>
      </div>
    );
  }
}
