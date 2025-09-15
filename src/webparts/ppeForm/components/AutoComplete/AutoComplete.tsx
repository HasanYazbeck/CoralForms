import * as React from "react";
import TextField from "@mui/material/TextField";
import Autocomplete from "@mui/material/Autocomplete";
import { ICommon } from "../../../../Interfaces/ICommon";

export interface IAutocompleteOption {
  label?: string;
  id: string;
}

interface IAutocompleteProps<T> {
  onSelectItem?: (item: any | undefined) => void;
  // OnChange?: (selected: IAutocompleteOption | null) => void;
  searchResults?: T[];
  options: ICommon[];
  keyId: string;
  className?: string | undefined;
  placeholder?: string;
  value?: any | undefined;
  label: string;
}

interface IAutocompleteState<T> {
  searchResults: T[];
  selectedItem: T | undefined;
  // options: ICommon[];
}

export class AutoComplete extends React.Component<
  IAutocompleteProps<any>,
  IAutocompleteState<any>
> {
  state: IAutocompleteState<any> = {
    searchResults: [],
    selectedItem: undefined,
    // options: [],
  };

  constructor(props: IAutocompleteProps<any>) {
    super(props);
  }

  public render(): React.ReactElement<{}> {
    return (
      <Autocomplete 
        //  onChange={(event) => this.props.OnChange}
      
      value={this.props.options.find((o) => o.label === this.props.value) || null}
        disablePortal
        options={this.props.options}
        sx={{ width: 300 }}
        renderInput={(params) => (
          <TextField {...params} label={this.props.label} />
        )}
      />
    );
  }
}
