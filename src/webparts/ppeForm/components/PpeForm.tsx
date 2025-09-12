import * as React from "react";
// import { useState } from "react";
import type {
  IPpeFormWebPartProps,
  IPpeFormWebPartState,
} from "./IPpeFormProps";
import { IPPEForm } from "../../../Interfaces/IEmployeeProps";
import Accordion from "@mui/material/Accordion";
// import AccordionActions from '@mui/material/AccordionActions';
import AccordionSummary from "@mui/material/AccordionSummary";
import AccordionDetails from "@mui/material/AccordionDetails";
import Typography from "@mui/material/Typography";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import { AutoComplete } from "./AutoComplete/AutoComplete";

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";
import CircularProgress from "@mui/material/CircularProgress";

export default class PpeForm extends React.Component<
  IPpeFormWebPartProps,
  IPpeFormWebPartState
> {
  constructor(props: IPpeFormWebPartProps) {
    super(props);
    this.state = {
      selectedEmployeeId: undefined,
      jobTitle: "",
      department: "",
      division: "",
      company: "",
    };
  }

  private handleEmployeeChange = (selectedOption: { label: string; id: string }) => {
    if(selectedOption) {
const selectedUser = this.props.Users.find((u) => u.id === selectedOption.id);
      if (selectedUser) {
        this.setState({
          selectedEmployeeId:  selectedOption.id,
          jobTitle: selectedUser.jobTitle || "",
          department: selectedUser.department || "",
          // division: selectedUser.division || "",
          // company: selectedUser.company || ""
        });
      }
    }
      
    };

  public render(): React.ReactElement<IPPEForm> {

    if (this.props.IsLoading) {
    return (
      <div className={styles.loadingContainer}>
        <CircularProgress variant="indeterminate" />
      </div>
    );
  }

    return (
      <div>
        <Accordion defaultExpanded>
          <AccordionSummary
            expandIcon={<ExpandMoreIcon />}
            aria-controls="panel1-content"
            id="panel1-header"
          >
            <Typography component="span">Employee Info</Typography>
          </AccordionSummary>
          <AccordionDetails>
            <form>
              <div className="row">
                <div className="form-group col-md-6">
                  <AutoComplete
                    label="Employee Name"
                    keyId="1"
                    className={styles.formField}
                    options={this.props.Users.map((user) => ({
                      label: user.displayName || "",
                      id: user.id,
                    }))}
                    OnChange={this.handleEmployeeChange}
                  />
                </div>

                <div className="form-group col-md-6">
                  <AutoComplete
                    label="Job Title"
                    keyId="2"
                    value={this.state.jobTitle}
                    options={this.props.JobTitles.map((jobTitles) => ({
                      label: jobTitles,
                      id: jobTitles,
                    }))}
                  />
                </div>
              </div>

              <div className="row">
                <div className="form-group col-md-6">
                  <AutoComplete
                    label="Company"
                    keyId="3"
                    value={this.state.company}
                    options={this.props.Users.map((user) => ({
                      label: user.displayName,
                      id: user.id,
                    }))}
                  />
                </div>

                <div className="form-group col-md-6">
                  <AutoComplete
                    label="Division"
                    keyId="4"
                    value={this.state.division}
                    options={this.props.Users.map((user) => ({
                      label: user.displayName,
                      id: user.id,
                    }))}
                  />
                </div>
              </div>

              <div className="row">
                <div className="form-group col-md-6">
                  <AutoComplete
                    label="Department"
                    keyId="5"
                    value={this.state.department}
                    options={this.props.Departments.map((dep) => ({
                      label: dep,
                      id: dep,
                    }))}
                  />
                </div>

                <div className="form-group col-md-6"></div>
              </div>
            </form>
          </AccordionDetails>
        </Accordion>
      </div>
    );
  }
}
