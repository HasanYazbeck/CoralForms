import * as React from "react";
import { ILookupItem } from "../../../Interfaces/PtwForm/IPTWForm";
import { Checkbox, TextField } from "@fluentui/react";
import styles from "./PtwForm.module.scss";

interface ICheckBoxDistributerComponentProps {
  id: string;
  optionList: ILookupItem[];
  className?: string;
  colSpacing?: 'col-1' | 'col-2' | 'col-3' | 'col-4' | 'col-6';
  selectedIds?: number[]; // if provided, component acts controlled
  onChange?: (selectedIds: number[]) => void;
  countOthersAsSelection?: boolean; // default true
  onOthersChange?: (checked: boolean, text: string) => void;
  othersTextValue?: string | undefined;
}

export function CheckBoxDistributerComponent(props: ICheckBoxDistributerComponentProps): JSX.Element {
  const [othersChecked, setOthersChecked] = React.useState(false);
  const [othersText, setOthersText] = React.useState('');
  const [internalSelectedIds, setInternalSelectedIds] = React.useState<number[]>([]);
  const effectiveSelectedIds = React.useMemo(
    () => (props.selectedIds !== undefined ? props.selectedIds : internalSelectedIds),
    [props.selectedIds, internalSelectedIds]
  );
  const countOthers = props.countOthersAsSelection ?? true;

  const { regularCategories, othersCategory } = React.useMemo(() => {
    const items = props.optionList?.slice()?.sort((a, b) => a.orderRecord - b.orderRecord) ?? [];
    const others = items.find(c => c.title === 'Others' || c.title === 'Other' || c.title === 'Other(s)');
    const regular = items.filter(c => c.title !== 'Others' && c.title !== 'Other' && c.title !== 'Other(s)');
    return { regularCategories: regular, othersCategory: others };
  }, [props.optionList]);

   // NEW: derive active state and sync checkbox when text/value exists or Others selected
  const isOthersActive = React.useMemo(() => {
    const hasText = !!props.othersTextValue && props.othersTextValue.trim() !== '';
    const othersSelected = !!othersCategory && effectiveSelectedIds.includes(othersCategory.id);
    return othersChecked || hasText || (countOthers && othersSelected);
  }, [othersChecked, props.othersTextValue, effectiveSelectedIds, othersCategory, countOthers]);

  React.useEffect(() => {
    const hasText = !!props.othersTextValue && props.othersTextValue.trim() !== '';
    if (hasText && !othersChecked) setOthersChecked(true);
    // keep local mirror so user can edit when controlled
    if (hasText && props.othersTextValue !== othersText) setOthersText(props.othersTextValue!);
  }, [props.othersTextValue, othersChecked, othersText]);

  const setSelectedIds = (next: number[]) => {
    if (props.selectedIds === undefined) setInternalSelectedIds(next);
    props.onChange?.(next);
  };

  const toggle = (id: number, checked: boolean) => {
    const has = effectiveSelectedIds.includes(id);
    let next: number[];
    if (checked && !has) next = [...effectiveSelectedIds, id];
    else if (!checked && has) next = effectiveSelectedIds.filter(x => x !== id);
    else next = effectiveSelectedIds;

    props.onChange?.(next);
    setSelectedIds(next);
  };

  return (
    <div className="form-group col-md-12" id={props.id}>
      <div className="row">
        {regularCategories.map(category => (
          <div key={category.id} className={props.colSpacing ? props.colSpacing : 'col-3'}>
            <div className="my-2">
              <Checkbox
                label={category.title}
                checked={effectiveSelectedIds.includes(category.id)}
                onChange={(_, checked) => toggle(category.id, !!checked)}
              />
            </div>
          </div>
        ))}
      </div>

      {othersCategory && (
        <div className="row mt-1">
          <div className={styles.checkboxItem}>
            <Checkbox
              label={othersCategory.title}
              checked={othersChecked || (countOthers && effectiveSelectedIds.includes(othersCategory.id))}
              onChange={(_, checked) => {
                const isChecked = !!checked;
                setOthersChecked(isChecked);
                if (!isChecked) setOthersText('');
                if (countOthers) toggle(othersCategory.id, isChecked);
                props.onOthersChange?.(isChecked, isChecked ? othersText : '');
              }}
            />
          </div>

          <div className={styles.othersTextWrapper}>
            <TextField
              type="text"
              className={styles.othersText}
              placeholder="Please specify"
              value={props.othersTextValue !== undefined ? props.othersTextValue : othersText}
              onChange={(_, v) => {
                const val = v || '';
                setOthersText(val);
                props.onOthersChange?.(othersChecked, val);
              }}
              disabled={!isOthersActive}
            />
          </div>
        </div>
      )}
    </div>
  );
}