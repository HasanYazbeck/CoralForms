import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from "../../../Interfaces/ICommon";


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
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';

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
  const formName = "PERSONAL PROTECTIVE EQUIPMENT";
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

  // New hook state
  const [users, setUsers] = useState<IUser[]>([]);
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [itemInstructionsForUse, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);

  // Rows for the items table
  const [ppeItemsRows, setPpeItemsRows] = useState<Array<any>>([]);
  // Approvals sign-off rows (Department, HR, HSE, Warehouse)
  const [approvalsRows, setApprovalsRows] = useState<Array<any>>([
    { SignOff: 'Department Approval', Name: '', Signature: '', Date: undefined, __index: 0 },
    { SignOff: 'HR Approval', Name: '', Signature: '', Date: undefined, __index: 1 },
    { SignOff: 'HSE Approval', Name: '', Signature: '', Date: undefined, __index: 2 },
    { SignOff: 'Warehouse Approval', Name: '', Signature: '', Date: undefined, __index: 3 }
  ]);

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

  const _getCoralFormsList = useCallback(async (usersArg?: IUser[]): Promise<ICoralFormsList | undefined> => {
    try {
      const searchEscaped = formName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'CoralFormsList', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      const ppeform = data.find((obj: any) => obj !== null);
      let result: ICoralFormsList = { Id: "" };
      
      if (ppeform) {
        const createdBy = usersToUse?.find(u => u.id.toString() === ppeform.AuthorId?.toString());
        const created = ppeform.Created ? new Date(spHelpers.adjustDateForGMTOffset(ppeform.Created)) : undefined;
        result = {
          Id: ppeform.Id ?? undefined,
          Title: ppeform.Title ?? undefined,
          CreatedBy: createdBy,
          Created: created,
          hasInstructionForUse: ppeform.hasInstructionForUse ?? undefined,
          hasWorkflow: ppeform.hasWorkflow ?? undefined,
        };
      }
      setCoralFormsList(result);
      return result;
    } catch (error) {
      console.error('An error has occurred!', error);
      setCoralFormsList({ Id: '' });
      return undefined;
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
      // console.log("PPE Item:", result);
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
      // console.log("PPE Item Details:", result);
      setPpeItemDetails(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItemDetails([]);
    }
  }, [props.context, spHelpers]);

  const _getLKPItemInstructionsForUse = useCallback(async (usersArg?: IUser[], formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName,Order,Description,Created&$filter=substringof('${formName}', FormName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKPItemInstructionsForUse', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ILKPItemInstructionsForUse[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: ILKPItemInstructionsForUse = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            Description: obj.Description !== undefined && obj.Description !== null ? obj.Description : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };

          result.push(temp);
        }
      });
  console.log("Item Instrunctions For Use:", result);
  // sort by Order (ascending). If Order is missing, place those items at the end.
  result.sort((a, b) => {
    const aOrder = (a && a.Order !== undefined && a.Order !== null) ? Number(a.Order) : Number.POSITIVE_INFINITY;
    const bOrder = (b && b.Order !== undefined && b.Order !== null) ? Number(b.Order) : Number.POSITIVE_INFINITY;
    return aOrder - bOrder;
  });
  setItemInstructionsForUse(result);
    } catch (error) {
  console.error('An error has occurred while retrieving items!', error);
  setItemInstructionsForUse([]);
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
      const coralListResult = await _getCoralFormsList(fetchedUsers);
      console.log(coralListResult);
      await _getPPEItems(fetchedUsers);
      await _getPPEItemsDetails(fetchedUsers);

      // Use the returned result from _getCoralFormsList instead of the (possibly stale) coralFormsList state
      if (coralListResult && coralListResult.hasInstructionForUse) {
        await _getLKPItemInstructionsForUse(fetchedUsers, formName);
      }

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
  }, [_getUsers, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, _getLKPItemInstructionsForUse, props.context]);

  // ---------------------------
  // Row helpers
  // ---------------------------
  const createEmptyRow = useCallback(() => ({ Item: '', Brands: '', Required: false, Details: [] as string[], Qty: '', Size: '', SizesByType: {} as Record<string, string>, Selected: false }), []);

  const addRow = useCallback(() => {
    setPpeItemsRows(prev => {
      const base = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      base.push(createEmptyRow());
      return base;
    });
  }, [createEmptyRow]);

  const handleSizeByTypeChange = useCallback((rowIndex: number, typeKey: string, newVal: any) => {
    const newValStr = newVal !== undefined && newVal !== null ? String(newVal) : '';
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= rowIndex) rows.push(createEmptyRow());
      // @ts-ignore
      const current = rows[rowIndex].SizesByType || {};
      // @ts-ignore
      rows[rowIndex].SizesByType = { ...(current), [typeKey]: newValStr };
      return rows;
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

  // Approvals handlers
  const onApprovalChange = useCallback((index: number, field: string, value: any) => {
    setApprovalsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [];
      while (rows.length <= index) rows.push({ SignOff: '', Name: '', Signature: '', Date: undefined, __index: rows.length });
      // @ts-ignore
      rows[index][field] = value;
      return rows;
    });
  }, []);

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

  // Map of Item Title -> (TypeTitle -> [sizes])
  const sizesByTypeMap = useMemo(() => {
    const map: Record<string, Record<string, string[]>> = {};
    (ppeItemDetails || []).forEach((p: any) => {
      const itemTitle = p && p.PPEItem && p.PPEItem.Title ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined);
      if (!itemTitle) return;
      const types = p && p.Types ? p.Types : undefined;
      const typeTitles: string[] = [];
      if (Array.isArray(types)) {
        types.forEach((t: any) => { if (t && t.Title) typeTitles.push(String(t.Title).trim()); });
      } else if (types && types.Title) {
        typeTitles.push(String(types.Title).trim());
      }
      const sizesArr = normalizeToStringArray(p && p.Sizes ? p.Sizes : undefined) || [];
      if (!map[itemTitle]) map[itemTitle] = {};
      typeTitles.forEach(tt => {
        if (!map[itemTitle][tt]) map[itemTitle][tt] = [];
        map[itemTitle][tt] = Array.from(new Set(map[itemTitle][tt].concat(sizesArr)));
      });
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
      const defaultSizesByType: Record<string, string> = {};
      const byType = sizesByTypeMap[newItemStr] || {};
      Object.keys(byType).forEach(tt => { const arr = byType[tt]; if (arr && arr.length) defaultSizesByType[tt] = arr[0]; else defaultSizesByType[tt] = ''; });
      // set defaultBrand on next tick to allow ComboBox to pick up new options
      setTimeout(() => {
        setPpeItemsRows(current => {
          const copy = current && current.length > 0 ? [...current] : [createEmptyRow()];
          while (copy.length <= index) copy.push(createEmptyRow());
          copy[index].Brands = defaultBrand;
          copy[index].Size = defaultSize;
          // set per-type defaults
          // @ts-ignore
          copy[index].SizesByType = defaultSizesByType;
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

  const filterPersonasByText = (filterText: string): IPersonaProps[] => peopleList.filter(item => doesTextContain(item.text as string, filterText));
  function doesTextContain(text: string, filterText: string): boolean { if (!text || !filterText) return false; return text.toLowerCase().includes(filterText.toLowerCase()); }
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
                      <div className={styles.comboCell}>
                        <ComboBox allowFreeform autoComplete="on" selectedKey={item.Item || undefined} options={itemOptions} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; handleItemChange(item.__index, newVal || ''); }} />
                      </div>
                    )
                  },
                  {
                    key: 'columnBrand', name: 'Brand', fieldName: 'Brands', minWidth: 120, isResizable: true, onRender: (item: any) => {
                      const options = (brandsMap && brandsMap[item.Item]) ? brandsMap[item.Item].map((b: string) => ({ key: b, text: b })) : [];
                      return (
                        <div className={styles.comboCell}>
                          <div style={{ width: '100%' }}>
                            <ComboBox allowFreeform autoComplete="on" selectedKey={item.Brands || undefined} options={options} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; onRowChange(item.__index, 'Brands', newVal || ''); }} />
                          </div>
                        </div>
                      );
                    }
                  },
                  { key: 'columnRequired', name: 'Required', className: `text-center align-middle ${styles.justifyItemsCenter}`, fieldName: 'Required', minWidth: 90, maxWidth: 120, isResizable: false, onRender: (item: any) => <div className={`${styles.tableSecondaryBg} ${styles.justifyItemsCenter}`}><Checkbox checked={!!item.Required} onChange={(ev, checked) => onRowChange(item.__index, 'Required', !!checked)} /></div> },
                  {
                    key: 'columnDetails', name: 'Specific Details', fieldName: 'Details', minWidth: 320, isResizable: true, onRender: (item: any) => {
                      // Special-case: if Item === 'Others' render a Purpose TextField
                      const itemTitle = item && item.Item ? String(item.Item).trim() : '';
                      if (itemTitle === 'Others') {
                        const purposeVal = item && item.Purpose ? item.Purpose : '';
                        return (
                          <div className={`${styles.tableSecondaryBg} ${styles.detailsCell}`}>
                            <TextField value={purposeVal} onChange={(ev, val) => onRowChange(item.__index, 'Purpose', val || '')} />
                          </div>
                        );
                      }

                      // find PPEItemDetails entries that match the selected Item title
                      const detailRows = (ppeItemDetails || []).filter((d: any) => {
                        const title = d && d.PPEItem && d.PPEItem.Title ? String(d.PPEItem.Title).trim() : (d && d.Title ? String(d.Title).trim() : undefined);
                        return title === itemTitle;
                      });
                      // collect unique detail titles for checkboxes
                      const detailTitles = Array.from(new Set(detailRows.map((d: any) => d && d.Title ? String(d.Title).trim() : undefined).filter(Boolean)));
                      const selectedDetails = Array.isArray(item.Details) ? item.Details.map((d: any) => String(d).trim()) : [];
                      return (
                        <div className={`${styles.tableSecondaryBg} ${styles.detailsCell}`}>
                          {detailTitles.length === 0 ? <small className="text-muted">No details</small> : detailTitles.map((title: string) => (
                            <div key={title} className={styles.detailItem}>
                              <Checkbox checked={selectedDetails.indexOf(title) !== -1} onChange={() => toggleDetail(item.__index, title)} />
                              <span>{title}</span>
                            </div>
                          ))}
                        </div>
                      );
                    }
                  },
                  { key: 'columnQty', name: 'Qty', fieldName: 'Qty', minWidth: 70, maxWidth: 90, isResizable: false, onRender: (item: any) => <div className={`${styles.tableSecondaryBg} text-center align-middle`}><TextField value={item.Qty} onChange={(ev, val) => onRowChange(item.__index, 'Qty', val || '')} underlined={true} /></div> },
                  {
                    key: 'columnSize', name: 'Size', fieldName: 'Size', minWidth: 140, maxWidth: 260, isResizable: true, onRender: (item: any) => {
                      const itemTitle = item && item.Item ? String(item.Item).trim() : '';
                      // Special-case: if Item === 'Others' render a freeform Size text field
                      if (itemTitle === 'Others') {
                        const sizeVal = item && item.Size ? item.Size : '';
                        return <div className={styles.tableSecondaryBg}><TextField value={sizeVal} onChange={(ev, val) => onRowChange(item.__index, 'Size', val || '')} /></div>;
                      }
                      const byType = sizesByTypeMap[itemTitle] || {};
                      const typeKeys = Object.keys(byType || {});
                      if (!typeKeys.length) {
                        // no types => either single size options or N/A
                        const options = (sizesMap && sizesMap[itemTitle]) ? sizesMap[itemTitle].map((s: string) => ({ key: s, text: s })) : [];
                        if (!options.length) return <div className={styles.tableSecondaryBg}><small className="text-muted">N/A</small></div>;
                        return <ComboBox allowFreeform autoComplete="on" selectedKey={item.Size || undefined} options={options} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; onRowChange(item.__index, 'Size', newVal || ''); }} />;
                      }

                      // if all type lists are empty, show a single N/A
                      const allEmpty = typeKeys.every(tk => !(byType[tk] && byType[tk].length));
                      if (allEmpty) return <div className={styles.tableSecondaryBg}><small className="text-muted">N/A</small></div>;

                      // render per-type ComboBoxes side-by-side
                      const sizesByTypeSelected: Record<string, string> = (item && item.SizesByType) ? item.SizesByType : {};
                      return (
                        <div className={styles.sizeGroup}>
                          {typeKeys.map((tk: string) => {
                            const options = (byType[tk] || []).map((s: string) => ({ key: s, text: s }));
                            const selected = sizesByTypeSelected[tk] || '';
                            return (
                              <div key={tk} className={styles.sizeBox}>
                                <div style={{ fontSize: 11, color: '#666', marginBottom: 4 }}>{tk}</div>
                                {options.length ? (
                                  <ComboBox allowFreeform autoComplete="on" selectedKey={selected || undefined} options={options} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; handleSizeByTypeChange(item.__index, tk, newVal || ''); }} />
                                ) : (
                                  <small className="text-muted">N/A</small>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      );
                    }
                  }
                ];

                return (
                  <>
                    <DetailsList items={items} columns={columns} selection={selectionRef.current} selectionMode={SelectionMode.single} setKey="ppeItemsList" layoutMode={DetailsListLayoutMode.fixedColumns} isHeaderVisible={true} className={styles.detailsListHeaderCenter} />

                    {itemInstructionsForUse && itemInstructionsForUse.length > 0 && (
                      <div style={{ marginTop: 12 }}>
                        <Label>Instructions for Use:</Label>
                        {itemInstructionsForUse.map((instr: ILKPItemInstructionsForUse , idx: number) => (
                          <MessageBar key={instr.Id ?? instr.Order}  isMultiline styles={{ root: { marginBottom: 6 } }}>
                            <strong>{`${idx + 1}. `}</strong>
                            {instr.Description}
                          </MessageBar>
                        ))}
                      </div>
                    )}
                    {/* Approvals sign-off table */}
                    <div style={{ marginTop: 18 }}>
                      <Separator />
                      <Label>Approvals / Sign-off</Label>
                      <DetailsList
                        items={approvalsRows}
                        columns={[
                          { key: 'colSignOff', name: 'Sign off', fieldName: 'SignOff', minWidth: 160, isResizable: true },
                          { key: 'colName', name: 'Name', fieldName: 'Name', minWidth: 260, isResizable: true, onRender: (item: any) => (
                            <div style={{ minWidth: 220 }}>
                              <NormalPeoplePicker
                                itemLimit={1}
                                onResolveSuggestions={onFilterChanged}
                                onChange={(items: IPersonaProps[] | undefined) => {
                                  const sel = items && items.length ? items[0] : undefined;
                                  // store display name and id (if available) in the Name field
                                  onApprovalChange(item.__index, 'Name', sel ? (sel.text || '') : '');
                                  // also store an id field for potential persistence
                                  onApprovalChange(item.__index, 'NameId', sel ? sel.id : undefined);
                                }}
                                selectedItems={item.Name ? [{ text: item.Name, id: item.NameId }] : []}
                                resolveDelay={300}
                                inputProps={{ 'aria-label': 'Approvee' }}
                              />
                            </div>
                          ) },
                          { key: 'colSignature', name: 'Signature', fieldName: 'Signature', minWidth: 220, isResizable: true, onRender: (item: any) => (<TextField value={item.Signature || ''} onChange={(ev, val) => onApprovalChange(item.__index, 'Signature', val || '')} />) },
                          { key: 'colDate', name: 'Date', fieldName: 'Date', minWidth: 140, isResizable: true, onRender: (item: any) => (<DatePicker value={item.Date ? new Date(item.Date) : undefined} onSelectDate={(date) => onApprovalChange(item.__index, 'Date', date)} strings={defaultDatePickerStrings} />) }
                        ]}
                        selectionMode={SelectionMode.none}
                        setKey="approvalsList"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                      />
                    </div>
                  </>
                );
              })()}
            </div>
          </div>
        </Stack>
      </form>
    </div>
  );
}
