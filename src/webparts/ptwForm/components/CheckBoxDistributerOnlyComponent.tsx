import * as React from "react";
import { ILookupItem } from "../../../Interfaces/PtwForm/IPTWForm";
import { Checkbox, ITextFieldStyles, TextField } from "@fluentui/react";
import styles from "./PtwForm.module.scss";

interface ICheckBoxDistributerOnlyComponentProps {
    id: string;
    optionList: ILookupItem[];
    className?: string;
    colSpacing?: 'col-1' | 'col-2' | 'col-3' | 'col-4' | 'col-6';
    onChange?: (checked: boolean | undefined, selectedItem: ILookupItem | undefined) => void;
    selectedIds?: number[]; // controlled selection support
}

export function CheckBoxDistributerOnlyComponent(props: ICheckBoxDistributerOnlyComponentProps): JSX.Element {
    const [othersChecked, setOthersChecked] = React.useState(false);
    const [othersText, setOthersText] = React.useState('');
    const textFieldBlackStyles: Partial<ITextFieldStyles> = {
        // Applies to both input and textarea
        field: {
            color: '#000', // <-- main text
            selectors: {
                '&::placeholder': { color: '#666', fontWeight: 500, },        // optional: darker placeholder
                '&:disabled': { color: '#000', fontWeight: 500, }             // ensure disabled still renders black
            },
            subComponentStyles: {
                label: { root: { color: '#000', fontWeight: 500, } }
            }
        }
    };
    const { regularCategories, othersCategory } = React.useMemo(() => {
        const items = props.optionList?.slice()?.sort((a, b) => a.orderRecord - b.orderRecord) ?? [];
        const others = items.find(c => c.title === 'Others' || c.title === 'Other');
        const regular = items.filter(c => c.title !== 'Others' && c.title !== 'Other');
        return { regularCategories: regular, othersCategory: others };
    }, [props.optionList]);

    return (
        <div className="form-group col-md-12" id={props.id}>
            <div className="row p-2">
                {regularCategories.map(category => (
                    <div key={category.id} className={props.colSpacing ? props.colSpacing : 'col-3'}>
                        <div className="my-2">
                            <Checkbox
                                label={category.title}
                                checked={props.selectedIds ? props.selectedIds.includes(category.id) : undefined}
                                onChange={(_, checked) => props.onChange?.(checked, category)}
                            />
                        </div>
                    </div>
                ))}
            </div>

            {othersCategory && (
                <div className="row p-2">
                    <div className={styles.checkboxItem}>
                        <Checkbox label={othersCategory.title}
                            checked={props.selectedIds ? props.selectedIds.includes(othersCategory.id) : othersChecked}
                            onChange={(_, checked) => {
                                const isChecked = !!checked;
                                setOthersChecked(isChecked);
                                if (!isChecked) setOthersText('');
                                props.onChange?.(isChecked, othersCategory);
                            }}
                        />
                    </div>

                    <div className={styles.othersTextWrapper}>
                        <TextField type="text"
                            className={styles.othersText}
                            placeholder="Please specify"
                            value={othersText}
                            // onChange={(e) => setOthersText(e.target.value)}
                            disabled={!othersChecked}
                            styles={!othersChecked ? textFieldBlackStyles : undefined}
                        />
                    </div>
                </div>
            )}
        </div>

    );
}