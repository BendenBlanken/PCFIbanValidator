import { IInputs, IOutputs } from "./generated/ManifestTypes";
var IBAN = require('iban');

export class IbanValidator implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // parameter declarations
    private _ibanElement: HTMLInputElement;
    private _ibanImageElement: HTMLImageElement;

    private _iban: string;

    private _ibanChanged: EventListenerOrEventListenerObject;

    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;	
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
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
        // assigning environment variables. 
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        // Add control initialization code
        this._ibanChanged = this.ibanChanged.bind(this);

        // textbox control
        this._ibanElement = document.createElement("input");
        this._ibanElement.setAttribute("type", "text");
        this._ibanElement.setAttribute("class", "ibanelement");
        this._ibanElement.addEventListener("change", this._ibanChanged);

        this._ibanImageElement = <HTMLImageElement>document.createElement("img");
        this._ibanImageElement.setAttribute("height", "20");
        this._ibanImageElement.setAttribute("width", "16");
        this._ibanImageElement.setAttribute("class", "ibanhidden");
		
		this.findImage(this._ibanImageElement, "valid.png");

        this._container.appendChild(this._ibanElement);
        this._container.appendChild(this._ibanImageElement);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
        // Add code to update control view

        // @ts-ignore
        var crmIbanAttribute = this._context.parameters.Iban.attributes.LogicalName;

        // @ts-ignore 
        Xrm.Page.getAttribute(crmIbanAttribute).setValue(this._context.parameters.Iban.formatted);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
        return {
            Iban: this._iban
        };
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
        this._ibanElement.removeEventListener("change", this._ibanChanged);
    }

    private ibanChanged(evt: Event): void {
        var iban = this._ibanElement.value;
        if (iban != null && iban.length > 0) {
            if (IBAN.isValid(iban)) {
                this._ibanImageElement.setAttribute("title", "This is a valid IBAN.");
                this._ibanImageElement.removeAttribute("class");
				
				this.findImage(this._ibanImageElement, "valid.png");
				
				// @ts-ignore 
				Xrm.Page.ui.clearFormNotification("IbanValidator");
            }
            else {
                this._ibanImageElement.setAttribute("title", "this is a invalid IBAN.");
                this._ibanImageElement.removeAttribute("class");
				
				this.findImage(this._ibanImageElement, "invalid.png");
				
				// @ts-ignore 
				Xrm.Page.ui.setFormNotification("Iban is invalid.", "2", "IbanValidator");

            }
            
			this._iban = iban;

        }
        else {
            this._ibanImageElement.setAttribute("class", "ibanhidden");
            this._ibanImageElement.removeAttribute("title");
            this._iban = "";
			
			this.findImage(this._ibanImageElement, "valid.png");
			// @ts-ignore 
			Xrm.Page.ui.clearFormNotification("IbanValidator");
        }
        this._notifyOutputChanged();
    }
	
	private findImage(img: HTMLImageElement, imageName: string) {
		//find the image
		this._context.resources.getResource(imageName,
		content => {
			this.setImage(img, "png", content);
		},
		() => {
			this.showError();
		});
	}

	private setImage(element: HTMLImageElement, fileType: string, fileContent: string): void {
		//set the image to the img element
		fileType = fileType.toLowerCase();
		let imageUrl: string = `data:image/${fileType};base64, ${fileContent}`;
		element.src = imageUrl;
	}

	private showError(): void {
		console.log('error while downloading .png');
	}
}