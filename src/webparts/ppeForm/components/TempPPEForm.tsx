import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from "../../../Interfaces/ICommon";

// Components
import { DefaultPalette } from "@fluentui/react";
import type { IPpeFormWebPartProps } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
// import { MessageBar } from '@fluentui/react/lib/MessageBar';
// import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IUser } from "../../../Interfaces/IUser";
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";
import { IEmployeeProps, IEmployeesPPEItemsCriteria } from "../../../Interfaces/IEmployeeProps";

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

  // Local state (converted from class state)
  const [jobTitle, setJobTitle] = useState("");

  const [department, setDepartment] = useState("");
  const [division, setDivision] = useState("");
  const [company, setCompany] = useState("");
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_employeeId, setEmployeeId] = useState<number | undefined>(undefined);
  const [submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [requester, setRequester] = useState<IPersonaProps[]>([]);
  const [isReplacementChecked, setIsReplacementChecked] = useState(false);

  // New hook state
  const [users, setUsers] = useState<IUser[]>([]);
  // Employees list items fetched via _getEmployees (used for picker search)
  const [employees, setEmployees] = useState<IEmployeeProps[]>([]);
  const [, setPpeItems] = useState<IPPEItem[]>([]);
  const [, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);



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

  const _getEmployees = useCallback(async (usersArg?: IUser[], employeeFullName?: string): Promise<IEmployeeProps[]> => {
    try {
      const query: string = `?$select=Id,EmployeeID,FullName,Division/Id,Division/Title,Company/Id,Company/Title,EmploymentStatus,JobTitle/Id,JobTitle/Title,Department/Id,Department/Title,Manager/Id,Manager/FullName,Created&$expand=Division,Company,JobTitle,Department,Manager&$filter=substringof('${employeeFullName}', FullName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Employee', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeeProps[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IEmployeeProps = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            employeeID: obj.EmployeeID !== undefined && obj.EmployeeID !== null ? obj.EmployeeID : 0,
            fullName: obj.FullName !== undefined && obj.FullName !== null ? obj.FullName : undefined,
            jobTitle: obj.JobTitle !== undefined && obj.JobTitle !== null ? { id: obj.JobTitle.Id, title: obj.JobTitle.Title } : undefined,
            company: obj.Company !== undefined && obj.Company !== null ? { id: obj.Company.Id, title: obj.Company.Title } : undefined,
            department: obj.Department !== undefined && obj.Department !== null ? { id: obj.Department.Id, title: obj.Department.Title } : undefined,
            manager: obj.Manager !== undefined && obj.Manager !== null ? { Id: obj.Manager.Id, fullName: obj.Manager.FullName } as IEmployeeProps : undefined,
            employmentStatus: obj.EmploymentStatus !== undefined && obj.EmploymentStatus !== null ? obj.EmploymentStatus : undefined,
            division: obj.Division !== undefined && obj.Division !== null ? { id: obj.Division.Id, title: obj.Division.Title } : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };

          result.push(temp);
        }
      });
      setEmployees(result);
      return result;
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setEmployees([]);
      return [];
    }
  }, [props.context, spHelpers]);

    const _getEmployeesPPEItemsCriteria = useCallback(async (usersArg?: IUser[], employeeID?: string): Promise<IEmployeeProps[]> => {
    try {
      const query: string = `?$select=Id,Employee/Id,Employee/FullName,Created,SafetyHelmet,ReflectiveVest,SafetyShoes,` +
      `Employee/Id,Employee/FullName,RainSuit/Id,RainSuit/DisplayText,UniformCoveralls/Id,UniformCoveralls/DisplayText,UniformTop/Id,UniformTop/DisplayText,`+
      `UniformPants/Id,UniformPants/DisplayText,WinterJacket/Id,WinterJacket/DisplayText`+
      `&$expand=Employee,RainSuit,UniformCoveralls,UniformTop,UniformPants,WinterJacket`+
      `&$filter=substringof('${employeeID}', Employee/Id)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'EmployeePPEItemsCriteria', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeesPPEItemsCriteria[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IEmployeesPPEItemsCriteria = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            employeeID: obj.EmployeeID !== undefined && obj.EmployeeID !== null ? obj.EmployeeID : 0,
            fullName: obj.FullName !== undefined && obj.FullName !== null ? obj.FullName : undefined,
            jobTitle: obj.JobTitle !== undefined && obj.JobTitle !== null ? { id: obj.JobTitle.Id, title: obj.JobTitle.Title } : undefined,
            company: obj.Company !== undefined && obj.Company !== null ? { id: obj.Company.Id, title: obj.Company.Title } : undefined,
            department: obj.Department !== undefined && obj.Department !== null ? { id: obj.Department.Id, title: obj.Department.Title } : undefined,
            manager: obj.Manager !== undefined && obj.Manager !== null ? { Id: obj.Manager.Id, fullName: obj.Manager.FullName } as IEmployeeProps : undefined,
            employmentStatus: obj.EmploymentStatus !== undefined && obj.EmploymentStatus !== null ? obj.EmploymentStatus : undefined,
            division: obj.Division !== undefined && obj.Division !== null ? { id: obj.Division.Id, title: obj.Division.Title } : undefined,
            reflectiveVest: obj.ReflectiveVest !== undefined && obj.ReflectiveVest !== null ? obj.ReflectiveVest : undefined,
            safetyHelmet: obj.SafetyHelmet !== undefined && obj.SafetyHelmet !== null ? obj.SafetyHelmet : undefined,
            safetyShoes: obj.SafetyShoes !== undefined && obj.SafetyShoes !== null ? obj.SafetyShoes : undefined,
            rainSuit: obj.RainSuit !== undefined && obj.RainSuit !== null ? { id: obj.RainSuit.Id, label: obj.RainSuit.DisplayText } : undefined,
            uniformCoveralls: obj.UniformCoveralls !== undefined && obj.UniformCoveralls !== null ? { id: obj.UniformCoveralls.Id, label: obj.UniformCoveralls.DisplayText } : undefined,
            uniformTop: obj.UniformTop !== undefined && obj.UniformTop !== null ? { id: obj.UniformTop.Id, label: obj.UniformTop.DisplayText } : undefined,
            uniformPants: obj.UniformPants !== undefined && obj.UniformPants !== null ? { id: obj.UniformPants.Id, label: obj.UniformPants.DisplayText } : undefined,
            winterJacket: obj.WinterJacket !== undefined && obj.WinterJacket !== null ? { id: obj.WinterJacket.Id, label: obj.WinterJacket.DisplayText } : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };

          result.push(temp);
        }
      });
      setEmployees(result);
      return result;
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setEmployees([]);
      return [];
    }
  }, [props.context, spHelpers]);

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
      // console.error('An error has occurred!', error);
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
            Brands: normalizeToStringArray(obj.Brands),
            PPEItemsDetails: []
          };
          result.push(temp);
        }
      });
      setPpeItems(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItems([]);
    }
  }, [props.context, spHelpers]);

  const _getPPEItemsDetails = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,PPEItem,Sizes,Created,PPEItem/Id,PPEItem/Title&$expand=PPEItem`;
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
            Sizes: normalizeToStringArray(obj.Sizes),
            PPEItem: obj.PPEItem !== undefined ? {
              Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
              Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,
              Brands: normalizeToStringArray(obj.PPEItem.Brands),
            } : undefined,
          };
          result.push(temp);
        }
      });
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
  }, [_getEmployees,_getEmployeesPPEItemsCriteria, _getUsers, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, _getLKPItemInstructionsForUse, props.context]);





  const handleEmployeeChange = useCallback((items: IPersonaProps[]) => {
    if (items && items.length > 0) {
      const selected = items[0];
      // First try to find in employees list by FullName (fullName -> persona.text)
      const emp = employees.find(e => (e.fullName || '').toLowerCase() === (selected.text || '').toLowerCase());
      // Fallback to users (Graph) if not found
      const user = users.find(u => u.displayName?.toLowerCase() === (selected.text || '').toLowerCase() || u.id === selected.id);
      setEmployee([selected]);
      setEmployeeId(emp?.employeeID);
      setJobTitle(emp?.jobTitle?.title || user?.jobTitle || '');
      setDepartment(emp?.department?.title || user?.department || '');
      // Ensure division stored as a simple string
      setDivision(emp?.division?.title || '');
      setCompany(emp?.company?.title || user?.company || '');
      // Auto-set requester ONLY if Employee list record has a manager; otherwise leave empty
      if (emp?.manager?.fullName) {
        setRequester([{ text: emp.manager.fullName, id: emp.manager.Id ? String(emp.manager.Id) : emp.manager.fullName }]);
      } else {
        setRequester([]);
      }
    } else {
      setEmployee([]);
      setEmployeeId(undefined);
      setJobTitle('');
      setDepartment('');
      setDivision('');
      setCompany('');
      setRequester([]);
    }
  }, [employees, users]);

  // Employee picker dynamic resolver using Employee list instead of raw users
  const employeeOnFilterChanged = useCallback((filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): Promise<IPersonaProps[]> | IPersonaProps[] => {
    if (!filterText || filterText.trim().length === 0) return [];
    // Always return a promise so the picker waits for async results
    return _getEmployees(undefined, filterText).then(fetched => {
      const list = fetched.length ? fetched : employees; // fallback to existing state
      const matches = list
        .filter(e => (e.fullName || '').toLowerCase().includes(filterText.toLowerCase()))
        .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName) }) as IPersonaProps);
      const deduped: IPersonaProps[] = [];
      const seen = new Set<string>();
      matches.forEach(m => { const key = (m.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(m); } });
      return limitResults ? deduped.slice(0, limitResults) : deduped;
    });
  }, [_getEmployees, employees]);

  // Requester resolver (merge employees and Graph users for broader search)
  const requesterOnFilterChanged = useCallback((filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (!filterText || filterText.trim().length === 0) return [];
    const lower = filterText.toLowerCase();
    const employeeMatches = employees
      .filter(e => (e.fullName || '').toLowerCase().includes(lower))
      .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName) }) as IPersonaProps);
    const userMatches = users
      .filter(u => (u.displayName || '').toLowerCase().includes(lower))
      .map(u => ({ text: u.displayName || '', secondaryText: u.jobTitle, id: u.id }) as IPersonaProps);
    const combined = employeeMatches.concat(userMatches);
    const deduped: IPersonaProps[] = [];
    const seen = new Set<string>();
    combined.forEach(p => { const key = (p.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(p); } });
    return limitResults ? deduped.slice(0, limitResults) : deduped;
  }, [employees, users]);

  const handleRequesterChange = useCallback((items: IPersonaProps[] | undefined) => {
    if (items && items.length) setRequester([items[0]]); else setRequester([]);
  }, []);

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
            <div className="form-group col-md-6"><TextField label="Employee ID" value={_employeeId?.toString()} disabled={true} /></div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker
                label={"Employee Name"}
                itemLimit={1}
                // Use employee list based resolver
                onResolveSuggestions={employeeOnFilterChanged}
                className={'ms-PeoplePicker'}
                key={'employee'}
                removeButtonAriaLabel={'Remove'}
                inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'Employee Picker' }}
                onInputChange={onInputChange}
                resolveDelay={50}
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
              <TextField label="Job Title" value={jobTitle} disabled={true} />
            </div>
            <div className="form-group col-md-6">
              <TextField label="Department" value={department} disabled={true} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6"><TextField label="Division" value={division} disabled={true} /></div>
            <div className="form-group col-md-6"><TextField label="Company" value={company} disabled={true} /></div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker
                label={"Requester Name"}
                itemLimit={1}
                onResolveSuggestions={requesterOnFilterChanged}
                className={'ms-PeoplePicker'}
                key={'requester'}
                removeButtonAriaLabel={'Remove'}
                inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'Requester Picker' }}
                onInputChange={onInputChange}
                resolveDelay={150}
                disabled={false}
                onChange={handleRequesterChange}
                selectedItems={requester}
              />
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


      </form>
    </div>
  );
}
