import { lookup } from "dns";
import { IInputs, IOutputs } from "./generated/ManifestTypes";


// Define the LookupValue interface inside the class
interface LookupValue {
	id: string;
	name: string;
	entityType: string;
}

export class AutoPopulateTextColumn implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    // reference to Power Apps component framework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// reference to the component container HTMLDivElement
	private _container: HTMLDivElement;
	// power Apps component framework delegate which will be assigned to this object which would be called whenever any update happens.
	private _notifyOutputChanged: () => void;
	// input element that is used to auto populate
	private inputElement: HTMLInputElement;
	// value of the column is stored and used inside the control
	private _value: string;
	//private _regardingValue: any;
	private _configValue: string;
	private _string1: string;
	private _string2: string;
	private _string3: string;
	private _lookup1: ComponentFramework.LookupValue | null;
	private _lookup2: ComponentFramework.LookupValue | null;


    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */

	public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        // Add control initialization code
        
        this._context = context;
		this._container = document.createElement("div");
		this._notifyOutputChanged = notifyOutputChanged;
		// creating HTML elements for the auto populate text column
		this.inputElement = document.createElement("input");
		this.inputElement.setAttribute("class", "boundText");
		this.inputElement.setAttribute("placeholder", "---");

        // Log the context parameters to debug
        console.log("Context Parameters: ", context.parameters);

		// retrieving the latest value from the component.
		this._value = context.parameters.boundTextProperty.raw
		? context.parameters.boundTextProperty.raw
		: "";
		this.inputElement.value = this._value;
		this.inputElement.addEventListener("blur", this.onBlur.bind(this));
		this._configValue = context.parameters.configValue.raw
		? context.parameters.configValue.raw
		: "";
		this._string1 = context.parameters.String1.raw ? context.parameters.String1.raw : "" ;
		this._string2 = context.parameters.String2.raw ? context.parameters.String2.raw : "" ;
		this._string3 = context.parameters.String3.raw ? context.parameters.String3.raw : "" ;
		this._lookup1 = context.parameters.Lookup1.raw ? context.parameters.Lookup1.raw[0] : null;
		this._lookup2 = context.parameters.Lookup2.raw ? context.parameters.Lookup2.raw[0] : null ;
		
		const parameter=[];
		for(const param of Object.keys(context.parameters)){
			parameter.push(param);
		}
		this.getFormattedText(this._configValue, parameter);

		// appending the HTML elements to the component's HTML container element.
		this._container.appendChild(this.inputElement);
		container.appendChild(this._container);

    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view 0.14
		this._context = context;

        const newConfigValue = context.parameters.configValue.raw
		? context.parameters.configValue.raw
		: "";
		const String1 = context.parameters.String1.raw ? context.parameters.String1.raw : "" ;
		const String2 = context.parameters.String2.raw ? context.parameters.String2.raw : "" ;
		const String3 = context.parameters.String3.raw ? context.parameters.String3.raw : "" ;
		const Lookup1 = context.parameters.Lookup1.raw ? context.parameters.Lookup1.raw : null ;
		const Lookup2 = context.parameters.Lookup2.raw ? context.parameters.Lookup2.raw : null ;
		

		if(this._configValue != newConfigValue){
			this._configValue = newConfigValue;
			//this.inputElement.value = this._value;
		}
		const parameter=[];
		for(const param of Object.keys(context.parameters)){
			parameter.push(param);
		}
		this.getFormattedText(newConfigValue, parameter);


    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
			boundTextProperty: this._value
		};
    }
    private onBlur(event: Event): void {
        // Handle the blur event
        this._value = (event.target as HTMLInputElement).value;
        this._notifyOutputChanged();
    }

	private notify(newText:string):void {
		this.inputElement.value = newText;
		this._value = this.inputElement.value;
		this._notifyOutputChanged();
	}

	/**
	 * Called this function to retrive the regarding record and format the output.
	 * @param fconfigValue string that holds the new config value.
	 * @param tablename The tablename. Has value only when the datasource is table or lookup.
	 */
	async getFormattedText(fconfigValue: string, fparameter:string[]): Promise<void> {
		let output = "";
		const multiTypeConfig: string[] = fconfigValue.split("|");
		if(multiTypeConfig.length === 0 || multiTypeConfig === null) {
			this.notify(output);
			return;
		}
		for(const typeConfig of multiTypeConfig) {
			const config: string[] = typeConfig.split(",");
			
			//Append a direct string of text
			if(config.length === 1){
				output = output + " " + config[0];
			}
			//TestString - | columnValue,AddressLine1 | lookup,cityValue,name,City:name | lookup,stateValue,ais_code,State:ais_code
			//Evaluate and append data from same table
			if(config.length > 1 ){
				
				//check if the value requested is a direct column value				
				if(config[0] === "columnValue"){
					const columnName = config[1];
					if( this._context.parameters[columnName as keyof IInputs] ){
						output = output + " " + this._context.parameters[columnName as keyof IInputs]?.raw || "";
					}else{
						output = output + " [columnValue for " + columnName + "not available]";
					}
				}
				
				//check if the value requested is a lookup value
				if(config[0] === "lookup"){
					const parameterName = config[1];
					const columnName = config[2];					
					if(this._context.parameters[parameterName as keyof IInputs]){
						const lookupdataArray = this._context.parameters[parameterName as keyof IInputs]?.raw;
						if (lookupdataArray && lookupdataArray.length > 0) {
							const lookupdata = lookupdataArray[0];				
							if (this.isLookupValue(lookupdata) && lookupdata !== null && columnName in lookupdata) {
								output = output + " " + lookupdata[columnName as keyof ComponentFramework.LookupValue];
							} else if(this.isLookupValue(lookupdata)){
								try{
									const luValue = await this.getLookupValue(lookupdata.entityType, columnName, lookupdata.id);
									output = output + " " + luValue;
								}
								catch(error){
									output = output + " [Error retrieving " + columnName + " for " + parameterName + "]";
								}								
							}
							else {
								output = output + " [" + columnName + " for " + parameterName + " not available]";
							}
						}else {
							output = output + " [lookupValue array " + parameterName + "is empty]";
						}
					}
				}
			}
			
		}

		this.notify(output.trim())
	}

	isLookupValue(value: unknown): value is ComponentFramework.LookupValue {
		if(value && typeof value === 'object' && 'entityType' in value && 'id' in value){
			return true;
		}
		else{return false;}
	}

	/**
	 * Called to get the column value from the lookup reference
	 */
	async getLookupValue(tableName: string, columnName:string, id: string): Promise<string> {
		//let lookupValue = "nodata";
		return new Promise((resolve, reject) => {
			this._context.webAPI.retrieveRecord(tableName, id, "?$select=" + columnName).then(
				(result) => {
					const temp_result=result;
					console.log("Lookup value found: " + result[columnName]);
					const temp_output=result[columnName];
					return resolve(result[columnName]);
				},
				(error) => {
					console.log("Lookup value not found: " + error.message);
					return resolve("null");
				}
			).catch((error) => {
				console.log("Lookup value not found: " + error.message);
				return reject("error");
			});
			
		});
	}

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
	
}
