import React = require("react");
import ReactDom = require("react-dom");
import { DropDownSelector } from "./DropDownSelector";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";

export class PCFDateTimeYearOnly implements ComponentFramework.StandardControl<IInputs, IOutputs>
{

    private _availableOptions: IDropdownOption[];
    private _context : ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _currentValue?: number;
    private _currentYear: number;
    private _notifyOutputChanged: () => void;

    /**
     * Empty constructor.
     */
    constructor()
    {
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._availableOptions = new Array();
        this._container = container;
        this._context = context;
        this._currentYear = new Date().getFullYear();
        this._notifyOutputChanged = notifyOutputChanged;

        // Push years before
        var yearsBeforeAmount = this._context.parameters.yearsBeforeAmount?.raw ?? 1;
        for (let i = yearsBeforeAmount; i > 0; i--)
        {
            this._availableOptions.push({ key: this._currentYear - i, text: (this._currentYear - i).toString() });
        }

        // Push current year
        this._availableOptions.push({ key: this._currentYear, text: this._currentYear.toString() });

        // Push years after
        var yearsAfterAmount = context.parameters.yearsAfterAmount?.raw ?? 1;
        for (let i = 1; i <= yearsAfterAmount; i++)
        {
            this._availableOptions.push({ key: this._currentYear + i, text:(this._currentYear + i).toString() });
        }

        this.renderView(context);
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this.renderView(context);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        if(this._currentValue === -1 || this._currentValue === undefined)
        {
            return { value: undefined };
        }

        let year = this._currentValue ?? null;
        let date = this.correctTimeZone( new Date(year, 0, 2));
        return { value: date };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        ReactDom.unmountComponentAtNode(this._container);
    }

    private renderView(context: ComponentFramework.Context<IInputs>)
    {
        let currentValue: number;
        currentValue = context.parameters.value != null && context.parameters.value.raw != null ? (<Date>context.parameters.value.raw).getFullYear() : -1;

        const dropDownSelector = React.createElement(DropDownSelector,
        {
            selectedValue: currentValue,
            availableOptions:  [{ key: -1, text: '----'}, ... this._availableOptions],
            isDisabled: context.mode.isControlDisabled,
            onChange: (selectedOption?: IDropdownOption) =>
            {
                this._currentValue = typeof selectedOption === 'undefined' || selectedOption.key === -1 ? undefined : <number>selectedOption.key;
                this._notifyOutputChanged();
            }
        });

        ReactDom.render(dropDownSelector, this._container);
    }

    private correctTimeZone(date: Date): Date
    {  
        const TIMEZONE_INDEPENDENT_BEHAVIOR = 3;
        const fieldBehavior = this._context.parameters.value.attributes!.Behavior;
        const timezoneOffsetInMinutes = fieldBehavior === TIMEZONE_INDEPENDENT_BEHAVIOR
          ? 0
          : this._context.userSettings.getTimeZoneOffsetMinutes(date);

        const newDate = new Date(date).setMinutes(date.getMinutes() + date.getTimezoneOffset() + timezoneOffsetInMinutes);

        return new Date(newDate);
      }
}
