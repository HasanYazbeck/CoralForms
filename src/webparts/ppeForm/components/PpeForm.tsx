import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGraphResponse, IGraphUserResponse } from "../../../Interfaces/ICommon";


// Components
import { DefaultPalette, DetailsListLayoutMode } from "@fluentui/react";
import type { IPpeFormWebPartProps } from "./IPpeFormProps";
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
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IUser } from "../../../Interfaces/IUser";
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";

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

export default function PpeForm(props: IPpeFormWebPartProps) {
  // Helpers and refs
  const spHelpers = useMemo(() => new SPHelpers(), []);
  const spCrudRef = useRef<SPCrudOperations | undefined>(undefined);
  const selectionRef = useRef(new Selection());

  // Local state (converted from class state)
  const [jobTitle, setJobTitle] = useState("");
  const [department, setDepartment] = useState("");
  const [division] = useState("");
  const [company, setCompany] = useState("");
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_employeeId, setEmployeeId] = useState<string | undefined>(undefined);
  const [submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [isReplacementChecked, setIsReplacementChecked] = useState(false);

  // New hook state requested by you
  const [users, setUsers] = useState<IUser[]>([]);
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);

  // Rows for the items table
  const [ppeItemsRows, setPpeItemsRows] = useState<Array<any>>([]);

  // Helper: ensure we return a string[] or undefined from strings or arrays
  const normalizeToStringArray = useCallback((val: any): string[] | undefined => {
    if (val === undefined || val === null) return undefined;
    if (Array.isArray(val)) return val.map((v) => (v !== undefined && v !== null ? String(v) : '')).filter(Boolean);
    if (typeof val === 'string') return val.split(',').map(s => s.trim()).filter(Boolean);
    // If it's an object with 'results' (SharePoint REST sometimes), handle that
    if (val && typeof val === 'object' && Array.isArray((val as any).results)) return (val as any).results.map((v: any) => String(v)).filter(Boolean);
    return undefined;
  }, []);

  // ---------------------------
  // Data-loading functions (ported)
  // ---------------------------
  const _getUsers = useCallback(async (): Promise<IUser[]> => {
    let fetched: IUser[] = [];
    let endpoint: string | null = "/users?$select=id,displayName,mail,department,jobTitle,mobilePhone,officeLocation&$expand=manager($select=id,displayName)";
    try {
      do {
        const client: MSGraphClientV3 = await (props.context as any).msGraphClientFactory.getClient("3");
        const response: IGraphResponse = await client.api(endpoint).get();
        if (response?.value && Array.isArray(response.value)) {
          const seenIds = new Set<string>();
          const mappedUsers = response.value
            .filter((u: IGraphUserResponse) => u.mail)
            .filter((user) => user.mail && !user.mail?.toLowerCase().includes("healthmailbox") && !user.mail?.toLowerCase().includes("softflow-intl.com") && !user.mail?.toLowerCase().includes("sync"))
            .filter(user => {
              if (seenIds.has(user.id)) return false;
              seenIds.add(user.id);
              return true;
            })
            .map((user: IGraphUserResponse) => ({
              id: user.id,
              displayName: user.displayName,
              email: user.mail,
              jobTitle: user.jobTitle,
              department: user.department,
              officeLocation: user.officeLocation,
              mobilePhone: user.mobilePhone,
              profileImageUrl: undefined,
              isSelected: false,
              manager: user.manager ? { id: user.manager.id, displayName: user.manager.displayName } : undefined,
            } as IUser));

          fetched.push(...mappedUsers);
          endpoint = response["@odata.nextLink"] || null;
        } else {
          break;
        }
      } while (endpoint);

      if (fetched.length > 0) {
        setUsers(fetched);
      }
      return fetched;
    } catch (error) {
      console.error("Error fetching users:", error);
      setUsers([]);
      return [];
    }
  }, [props.context]);

  const _getCoralFormsList = useCallback(async (usersArg?: IUser[]) => {
    try {
      const searchFormName = "PERSONAL PROTECTIVE EQUIPMENT";
      const searchEscaped = searchFormName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'CoralFormsList', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      let result: ICoralFormsList = { Id: "" };
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          result = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            hasInstructionForUse: obj.hasInstructionForUse !== undefined ? obj.hasInstructionForUse : undefined,
            hasWorkflow: obj.hasWorkflow !== undefined ? obj.hasWorkflow : undefined,
          };
        }
      });
      console.log("Coral Forms List:", result);
      setCoralFormsList(result);
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }, [props.context, spHelpers]);

  const _getPPEItems = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,Brands,Created`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPEItems', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItem[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItem = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            // IsRequired: obj.Required !== undefined ? obj.Required : undefined,
            Brands: normalizeToStringArray(obj.Brands),
            PPEItemsDetails: []
          };
          result.push(temp);
        }
      });
      console.log("PPE Item:", result);
      setPpeItems(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItems([]);
    }
  }, [props.context, spHelpers]);

  const _getPPEItemsDetails = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,PPEItem,Types,Sizes,Created,PPEItem/Id,PPEItem/Title,Types/Id,Types/Title&$expand=PPEItem,Types`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPEItemsDetails', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItemDetails[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItemDetails = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Types: obj.Types !== undefined && obj.Types !== null ? obj.Types : undefined,
            Sizes: normalizeToStringArray(obj.Sizes),
            PPEItem: obj.PPEItem !== undefined ? {
              Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
              Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,
              IsRequired: obj.PPEItem.IsRequired !== undefined ? obj.PPEItem.IsRequired : undefined,
              Brands: normalizeToStringArray(obj.PPEItem.Brands),
            } : undefined,
          };
          result.push(temp);
        }
      });
      console.log("PPE Item Details:", result);
      setPpeItemDetails(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItemDetails([]);
    }
  }, [props.context, spHelpers]);

  // ---------------------------
  // useEffect: load data on mount
  // ---------------------------
  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      setLoading(true);
      const fetchedUsers = await _getUsers();
      await _getCoralFormsList(fetchedUsers);
      await _getPPEItems(fetchedUsers);
      await _getPPEItemsDetails(fetchedUsers);
      
      if (!cancelled) {
        try {
          const currentUserEmail = props.context.pageContext.user.email;
          const current = fetchedUsers.find(u => u.email === currentUserEmail);
          if (current) setSubmitter([{ text: current.displayName || '', secondaryText: current.email || '', id: current.id }]);
        } catch (e) {
          // ignore if context not available
        }
        setLoading(false);
      }
    };

    load();

    return () => { cancelled = true; };
  }, [_getUsers, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, props.context]);

  // ---------------------------
  // Row helpers
  // ---------------------------
  const createEmptyRow = useCallback(() => ({ Item: '', Brands: '', Required: false, Details: [] as string[], Qty: '', Size: '', Selected: false }), []);

  const addRow = useCallback(() => {
    setPpeItemsRows(prev => {
      const base = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      base.push(createEmptyRow());
      return base;
    });
  }, [createEmptyRow]);

  const deleteSelectedRows = useCallback(() => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [];
      const selectedIndices = selectionRef.current ? selectionRef.current.getSelectedIndices() : [];
      if (selectedIndices && selectedIndices.length > 0) {
        const filtered = rows.filter((r, idx) => selectedIndices.indexOf(idx) === -1);
        selectionRef.current.setAllSelected(false);
        return filtered;
      } else {
        return rows.filter(r => !r.Selected);
      }
    });
  }, []);

  const onRowChange = useCallback((index: number, field: string, value: any) => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= index) rows.push(createEmptyRow());
      // @ts-ignore
      rows[index][field] = value;
      return rows;
    });
  }, [createEmptyRow]);

  // Toggle a PPEItemDetail title in the row's Details array
  const toggleDetail = useCallback((index: number, detailTitle: string) => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= index) rows.push(createEmptyRow());
      const currentDetails = Array.isArray(rows[index].Details) ? [...rows[index].Details] : [];
      const normalized = String(detailTitle || '').trim();
      const idx = currentDetails.findIndex((d: any) => String(d).trim() === normalized);
      if (idx === -1) {
        currentDetails.push(normalized);
      } else {
        currentDetails.splice(idx, 1);
      }
      // @ts-ignore
      rows[index].Details = currentDetails;
      return rows;
    });
  }, [createEmptyRow]);

  // ---------------------------
  // Handlers
  // ---------------------------
  // Map of Item Title -> Brands[] (deduped)
  const brandsMap = useMemo(() => {
    const map: Record<string, string[]> = {};
    // First, populate from the main PPE items list (ppeItems) which contains Brands
    (ppeItems || []).forEach((pi: any) => {
      const title = pi && pi.Title ? String(pi.Title).trim() : undefined;
      const brandsArr = normalizeToStringArray(pi && pi.Brands ? pi.Brands : undefined) || [];
      if (title) {
        if (!map[title]) map[title] = [];
        map[title] = Array.from(new Set(map[title].concat(brandsArr)));
      }
    });

    // Then merge in any brands from PPEItemDetails (to capture detail-level brands)
    (ppeItemDetails || []).forEach((p: any) => {
      const title = p && p.PPEItem && p.PPEItem.Title ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined);
      const brandsArr = normalizeToStringArray(p && p.PPEItem && p.PPEItem.Brands ? p.PPEItem.Brands : p && p.Brands ? p.Brands : undefined) || [];
      if (title) {
        if (!map[title]) map[title] = [];
        map[title] = Array.from(new Set(map[title].concat(brandsArr)));
      }
    });

    return map;
  }, [ppeItemDetails, ppeItems, normalizeToStringArray]);

  // Map of Item Title -> Sizes[] (deduped) from PPEItemDetails
  const sizesMap = useMemo(() => {
    const map: Record<string, string[]> = {};
    (ppeItemDetails || []).forEach((p: any) => {
      const title = p && p.PPEItem && p.PPEItem.Title ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined);
      const sizesArr = normalizeToStringArray(p && p.Sizes ? p.Sizes : undefined) || [];
      if (title) {
        if (!map[title]) map[title] = [];
        map[title] = Array.from(new Set(map[title].concat(sizesArr)));
      }
    });
    return map;
  }, [ppeItemDetails, normalizeToStringArray]);

  // When the Item for a row is changed, also pre-fill the Brands field with the first matching brand (if any)
  const handleItemChange = useCallback((index: number, newItem: any) => {
  const newItemStr = newItem !== undefined && newItem !== null ? String(newItem).trim() : '';
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= index) rows.push(createEmptyRow());
      rows[index].Item = newItemStr;
      // Clear Brands first so the Brand ComboBox re-renders with the new options
      rows[index].Brands = '';
      const options = brandsMap[newItemStr] || [];
      const defaultBrand = options && options.length > 0 ? options[0] : '';
      const sizeOptions = sizesMap[newItemStr] || [];
      const defaultSize = sizeOptions && sizeOptions.length > 0 ? sizeOptions[0] : '';
      // set defaultBrand on next tick to allow ComboBox to pick up new options
      setTimeout(() => {
        setPpeItemsRows(current => {
          const copy = current && current.length > 0 ? [...current] : [createEmptyRow()];
          while (copy.length <= index) copy.push(createEmptyRow());
          copy[index].Brands = defaultBrand;
          copy[index].Size = defaultSize;
          return copy;
        });
      }, 0);
      return rows;
    });
  }, [createEmptyRow, brandsMap]);

  const handleEmployeeChange = useCallback((items: IPersonaProps[]) => {
    if (items && items.length > 0) {
      const selected = items[0];
      const user = users.find(u => u.id === selected.id);
      setEmployee([selected]);
      setEmployeeId(selected.id as string);
      setJobTitle(user?.jobTitle || '');
      setDepartment(user?.department || '');
      setCompany(user?.company || '');
    } else {
      setEmployee([]);
      setEmployeeId(undefined);
      setJobTitle('');
      setDepartment('');
      setCompany('');
    }
  }, [users]);

  const handleNewRequestChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    if (checked) setIsReplacementChecked(false);
  }, []);

  const handleReplacementChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    setIsReplacementChecked(!!checked);
  }, []);

  // ---------------------------
  // Render
  // ---------------------------
  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label={"Preparing PPE form — fresh items coming right up!"} size={SpinnerSize.large} />
      </div>
    );
  }

  const delayResults = false;
  const logoUrl = `${props.context.pageContext.web.absoluteUrl}/SiteAssets/coral-logo.png`;
  const peopleList: IPersonaProps[] = users.map(user => ({ text: user.displayName || '', secondaryText: user.email || '', id: user.id }));

  const filterPromise = (personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (delayResults) return convertResultsToPromise(personasToReturn);
    return personasToReturn;
  };

  const onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = filterPersonasByText(filterText);
      filteredPersonas = removeDuplicates(filteredPersonas, currentPersonas);
      filteredPersonas = limitResults ? filteredPersonas.slice(0, limitResults) : filteredPersonas;
      return filterPromise(filteredPersonas);
    }
    return [];
  };

  const filterPersonasByText = (filterText: string): IPersonaProps[] => peopleList.filter(item => doesTextStartWith(item.text as string, filterText));
  function doesTextStartWith(text: string, filterText: string): boolean { return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0; }
  function removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) { return personas.filter(persona => !listContainsPersona(persona, possibleDupes)); }
  function listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) { if (!personas || !personas.length) return false; return personas.filter(item => item.text === persona.text).length > 0; }
  function convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> { return new Promise<IPersonaProps[]>((resolve) => setTimeout(() => resolve(results), 2000)); }
  function onInputChange(input: string): string { const outlookRegEx = /<.*>/g; const emailAddress = outlookRegEx.exec(input); if (emailAddress && emailAddress[0]) return emailAddress[0].substring(1, emailAddress[0].length - 1); return input; }

  return (
    <div className={styles.ppeFormBackground}>
      <form>
        <div className={styles.formHeader}>
          <img src={logoUrl} alt="Logo" className={styles.formLogo} />
          <span className={styles.formTitle}>PERSONAL PROTECTIVE EQUIPMENT (PPE) REQUISITION FORM</span>
        </div>

        <Stack horizontal styles={stackStyles}>
          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker
                label={"Employee Name"}
                itemLimit={1}
                onResolveSuggestions={onFilterChanged}
                className={'ms-PeoplePicker'}
                key={'normal'}
                removeButtonAriaLabel={'Remove'}
                inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }}
                onInputChange={onInputChange}
                resolveDelay={300}
                disabled={false}
                onChange={handleEmployeeChange}
              />
            </div>

            <div className="form-group col-md-6">
              <DatePicker disabled value={new Date(Date.now())} label="Date Requested" className={datePickerStyles.control} strings={defaultDatePickerStrings} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <TextField label="Job Title" value={jobTitle} />
            </div>
            <div className="form-group col-md-6">
              <TextField label="Department" value={department} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6"><TextField label="Division" value={division} /></div>
            <div className="form-group col-md-6"><TextField label="Company" value={company} /></div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker label={"Requester Name"} itemLimit={1} onResolveSuggestions={onFilterChanged} className={'ms-PeoplePicker'} key={'normal'} removeButtonAriaLabel={'Remove'} inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }} onInputChange={onInputChange} resolveDelay={300} disabled={false} onChange={handleEmployeeChange} />
            </div>

            <div className="form-group col-md-6">
              <NormalPeoplePicker label={"Submitter Name"} itemLimit={1} onResolveSuggestions={onFilterChanged} className={'ms-PeoplePicker'} key={'normal'} removeButtonAriaLabel={'Remove'} inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }} onInputChange={onInputChange} resolveDelay={300} disabled={true} selectedItems={submitter} />
            </div>
          </div>

          <div className={`row  ${styles.mt10}`}>
            <div className="form-group col-md-12 d-flex justify-content-between" >
              <Label htmlFor={""}>Reason for Request</Label>

              <Checkbox label="New Request" className="align-items-center" checked={!isReplacementChecked} onChange={handleNewRequestChange} />

              <Checkbox label="Replacement" className="align-items-center" checked={isReplacementChecked} onChange={handleReplacementChange} />

              <TextField placeholder="Reason" disabled={!isReplacementChecked} />
            </div>
          </div>
        </Stack>

        <Separator />

        <div className="mb-2 text-center">
          <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '1.0rem' }}>Please complete the table below in the blank spaces; grey spaces are for administrative use only.</small>
        </div>

        <Stack horizontal styles={stackStyles}>
          <div className="row">
            <div className="form-group col-md-12">
              {(() => {
                const commandBarItems: ICommandBarItemProps[] = [
                  { key: 'addItem', text: 'Add Item', iconProps: { iconName: 'Add' }, onClick: addRow },
                  { key: 'deleteSelected', text: 'Delete', iconProps: { iconName: 'Delete' }, onClick: deleteSelectedRows }
                ];
                return <CommandBar items={commandBarItems} styles={{ root: { marginBottom: 8 } }} />;
              })()}

              {(() => {
                const defaultRows = [createEmptyRow()];
                const rows = ppeItemsRows && ppeItemsRows.length > 0 ? ppeItemsRows : defaultRows;
                const items = rows.map((r, idx) => ({ ...r, __index: idx } as any));

                const itemDetailsFromDetails: string[] = (ppeItemDetails || []).map((p: any) => (p && p.PPEItem && p.PPEItem.Title) ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined)).filter(Boolean) as string[];
                const itemDetailsFromItems: string[] = (ppeItems || []).map((pi: any) => pi && pi.Title ? String(pi.Title).trim() : undefined).filter(Boolean) as string[];
                const allTitles = itemDetailsFromDetails.concat(itemDetailsFromItems).filter(Boolean) as string[];
                const distinctTitles = Array.from(new Set(allTitles)).sort((a, b) => a.localeCompare(b));
                const itemOptions: IComboBoxOption[] = distinctTitles.map(t => ({ key: t, text: t }));

                const columns: IColumn[] = [
                  {
                    key: 'columnItem', name: 'Item', fieldName: 'Item', minWidth: 150, isResizable: true, onRender: (item: any) => (
                      <ComboBox allowFreeform autoComplete="on" selectedKey={item.Item || undefined} options={itemOptions} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; handleItemChange(item.__index, newVal || ''); }} />
                    )
                  },
                  { key: 'columnBrand', name: 'Brand', fieldName: 'Brands', minWidth: 120, isResizable: true, onRender: (item: any) => {
                      const options = (brandsMap && brandsMap[item.Item]) ? brandsMap[item.Item].map((b: string) => ({ key: b, text: b })) : [];
                      return (
                        <div className={styles.comboCell}>
                          <div style={{ width: '100%' }}>
                            <ComboBox allowFreeform autoComplete="on" selectedKey={item.Brands || undefined} options={options} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; onRowChange(item.__index, 'Brands', newVal || ''); }} />
                          </div>
                        </div>
                      );
                    } },
                  { key: 'columnRequired', name: 'Required', className: `text-center align-middle ${styles.justifyItemsCenter}`, fieldName: 'Required', minWidth: 90, maxWidth: 120, isResizable: false, onRender: (item: any) => <div className={`table-secondary ${styles.justifyItemsCenter}`}><Checkbox checked={!!item.Required} onChange={(ev, checked) => onRowChange(item.__index, 'Required', !!checked)} /></div> },
                  { key: 'columnDetails', name: 'Specific Details', fieldName: 'Details', minWidth: 180, isResizable: true, onRender: (item: any) => {
                      // find PPEItemDetails entries that match the selected Item title
                      const itemTitle = item && item.Item ? String(item.Item).trim() : '';
                      const detailRows = (ppeItemDetails || []).filter((d: any) => {
                        const title = d && d.PPEItem && d.PPEItem.Title ? String(d.PPEItem.Title).trim() : (d && d.Title ? String(d.Title).trim() : undefined);
                        return title === itemTitle;
                      });
                      // collect unique detail titles for checkboxes
                      const detailTitles = Array.from(new Set(detailRows.map((d: any) => d && d.Title ? String(d.Title).trim() : undefined).filter(Boolean)));
                      const selectedDetails = Array.isArray(item.Details) ? item.Details.map((d: any) => String(d).trim()) : [];
                      return (
                        <div className="table-secondary">
                          {detailTitles.length === 0 ? <small className="text-muted">No details</small> : detailTitles.map((title: string) => (
                            <div key={title} className="form-check">
                              <input className="form-check-input" type="checkbox" id={`${item.__index}_detail_${title}`} checked={selectedDetails.indexOf(title) !== -1} onChange={() => toggleDetail(item.__index, title)} />
                              <label className="form-check-label" htmlFor={`${item.__index}_detail_${title}`}>{title}</label>
                            </div>
                          ))}
                        </div>
                      );
                    } },
                  { key: 'columnQty', name: 'Qty', fieldName: 'Qty', minWidth: 70, maxWidth: 90, isResizable: false, onRender: (item: any) => <div className="table-secondary text-center align-middle"><TextField value={item.Qty} onChange={(ev, val) => onRowChange(item.__index, 'Qty', val || '')} underlined={true} /></div> },
                  { key: 'columnSize', name: 'Size', fieldName: 'Size', minWidth: 100, maxWidth: 140, isResizable: true, onRender: (item: any) => <TextField value={item.Size} onChange={(ev, val) => onRowChange(item.__index, 'Size', val || '')} underlined={true} /> }
                ];

                return (
                  <DetailsList items={items} columns={columns} selection={selectionRef.current} selectionMode={SelectionMode.single} setKey="ppeItemsList" layoutMode={DetailsListLayoutMode.justified} isHeaderVisible={true} className={styles.detailsListHeaderCenter} />
                );
              })()}
            </div>
          </div>
        </Stack>
      </form>
    </div>
  );
}
