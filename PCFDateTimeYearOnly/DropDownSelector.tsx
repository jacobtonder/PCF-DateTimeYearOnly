import * as React from 'react';
import { dropdownStyles } from "./styles/dropdownStyle";
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

export interface DropDownSelectorProps
{
    selectedValue: string | number;
    availableOptions: IDropdownOption[];
    isDisabled: boolean;
    onChange: (selectedOption?: IDropdownOption) => void;
}

export const DropDownSelector: React.FunctionComponent<DropDownSelectorProps> = props => {
    return (
        <Dropdown
            selectedKey = { props.selectedValue }
            options = { props.availableOptions }
            disabled = { props.isDisabled }
            onChange = { (e, option?) => { props.onChange(option); } }
            styles={dropdownStyles}
        />
    );
}
