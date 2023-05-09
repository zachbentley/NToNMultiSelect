import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { Select2 } from "select2";
import * as $ from 'jquery';
import "./scripts/select2.min.js";

declare var Xrm: any;

class DataAction{
	guid : string;
	associate: boolean;
}
class NToNData{
	len : string;
    ida : string;
    na  : string;
    re : string;
    rn : string;
    actions : DataAction[];
}

export class NToNMultiSelect implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	//private _value: string;
	private contextObj: ComponentFramework.Context<IInputs>;
    // Div element created as part of this control's main container
	private mainContainer: HTMLSelectElement;
	private errorElement: HTMLDivElement;
	private selectedItems: string[] = [];
	private overlayDiv: HTMLDivElement;
	private container: HTMLDivElement;
	private _isValidState : boolean = true;

	private _relData : NToNData;
	
	private _linkedEntityName: string;
	private _relationshipEntity: string;
	private _relationshipName: string; 
	private _idAttribute: string;
	private _nameAttribute: string;
	private _linkedEntityFetchXmlResource: string;
	private _minRequiredSelections: number;
	private _maxRequiredSelections: number;

	private _numRelatedSelections: number;
	private _isError: boolean;
	private _errorLabelText: string;
	private labelElement: HTMLLabelElement;

	private _linkedEntityCollectionName: string;
	private _mainEntityCollectionName: string;
	
	private _entityMetadataSuccessCallback: any;
	private _linkedEntityMetadataSuccessCallback: any;
	private _relationshipSuccessCallback: any;
	private _successCallback : any;

	private _ctrlId : string;

	private _notifyOutputChanged: () => void;


	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}


	public S4() {
		return (((1+Math.random())*0x10000)|0).toString(16).substring(1); 
	}

	public newGuid(){
		var result : string = (this.S4() + this.S4() + "-" + this.S4() + "-4" + this.S4().substr(0,3) + "-" + this.S4() + "-" + this.S4() + this.S4() + this.S4()).toLowerCase();
	    return result;
	}
	 
	// then to call it, plus stitch in '4' in the third group
	
	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		//debugger;
		this.container = container;
		this.contextObj = context;
		this.labelElement = document.createElement("label");
		if(typeof Xrm == 'undefined')
		{
			this.errorElement = document.createElement("div");
			this.errorElement.innerHTML = "<H2>This control only works on model-driven forms!</H2>";
			container.appendChild(this.errorElement);
		}
		else{
			
			this._ctrlId = this.newGuid();
			this._relData = new NToNData();
			this._relData.actions = [];
			
			if(context.parameters.linkedEntityName.raw != null){
			  this._linkedEntityName = context.parameters.linkedEntityName.raw;
			  this._relData.len = this._linkedEntityName;
			}
			if(context.parameters.idAttribute.raw != null){
			  this._idAttribute = context.parameters.idAttribute.raw;
			  this._relData.ida = this._idAttribute;
			}
			if(context.parameters.nameAttribute.raw != null){
			  this._nameAttribute = context.parameters.nameAttribute.raw;
			  this._relData.na = this._nameAttribute;
			}
			if(context.parameters.relationshipEntity.raw != null){
			  this._relationshipEntity = context.parameters.relationshipEntity.raw;
			  this._relData.re = this._relationshipEntity;
			}
			if(context.parameters.relationshipName.raw != null){
			  this._relationshipName = context.parameters.relationshipName.raw;
			  this._relData.rn = this._relationshipName;
			}
			if(context.parameters.linkedEntityFetchXmlResource.raw != null){
			  this._linkedEntityFetchXmlResource = context.parameters.linkedEntityFetchXmlResource.raw;
			}
			if(context.parameters.minRequiredSelections.raw != null){
			  this._minRequiredSelections = context.parameters.minRequiredSelections.raw;
			}
			if(context.parameters.maxRequiredSelections.raw != null){
			  this._maxRequiredSelections = context.parameters.maxRequiredSelections.raw;
			}
			
			context.mode.trackContainerResize(true);
			container.classList.add("pcf_container_element");
			
			this.overlayDiv = document.createElement("div");
			this.overlayDiv.classList.add("pcf_overlay_element");
			container.appendChild(this.overlayDiv);
			
			this.mainContainer = document.createElement("select");
			this.mainContainer.id = this._ctrlId;
			this.mainContainer.classList.add("js-example-basic-multiple");
			this.mainContainer.classList.add("pcf_main_element");
			this.mainContainer.multiple = true;
			this.mainContainer.name = "states[]";
			container.appendChild(this.mainContainer);

			this._entityMetadataSuccessCallback = this.entityMetadataSuccessCallback.bind(this);
			this._linkedEntityMetadataSuccessCallback = this.linkedEntityMetadataSuccessCallback.bind(this);
			this._relationshipSuccessCallback = this.relationshipSuccessCallback.bind(this);
			this._successCallback = this.successCallback.bind(this);
			
			this._notifyOutputChanged = notifyOutputChanged;
			
			(<any>Xrm).Utility.getEntityMetadata((<any>this.contextObj).page.entityTypeName,[]).then(this._entityMetadataSuccessCallback, this.errorCallback);
			(<any>Xrm).Utility.getEntityMetadata(this._linkedEntityName,[]).then(this._linkedEntityMetadataSuccessCallback, this.errorCallback);
			
			if((<any>this.contextObj).page.entityId != null 
			   && (<any>this.contextObj).page.entityId != "00000000-0000-0000-0000-000000000000")
			{
				this.contextObj.webAPI.retrieveMultipleRecords(this._relationshipEntity, "?$filter="+ (<any>this.contextObj).page.entityTypeName+"id eq " + (<any>this.contextObj).page.entityId, 5000).then(this._relationshipSuccessCallback, this.errorCallback);
			}
			else{
				this.relationshipSuccessCallback(null);
			}
			this._numRelatedSelections=this.selectedItems.length;
			this.setErrorState();
			var thisVar : any;
			thisVar = this;
			$(document).ready(function() {
				thisVar.setReadonly();
				$('#'+ thisVar._ctrlId).select2({closeOnSelect: false}).on('select2:select', function (e) {
					var data = e.params.data;
					thisVar.selectAction("select", data.id);
				  }).on('select2:unselect', function (e) {
					var data = e.params.data;
					thisVar.selectAction("unselect", data.id);
				});
			});
		}
		
	}

	public entityMetadataSuccessCallback(value: any) : void | PromiseLike<void>
	{
		this._mainEntityCollectionName = value.EntitySetName;
	}
	
	public linkedEntityMetadataSuccessCallback(value: any) : void | PromiseLike<void>
	{
		this._linkedEntityCollectionName = value.EntitySetName;
	}
	
	public addOptions(value: any)
	{
		for(var i in value.entities)
		{
			var current : any = value.entities[i];

			var checked = this.selectedItems.indexOf(<string>current[this._idAttribute]) > -1;
			var newOption = new Option(current[this._nameAttribute], current[this._idAttribute], checked, checked);
	        $('#'+ this._ctrlId).append(newOption);
		}

	}

	public successCallback(value: any) : void | PromiseLike<void>
	{
		this.addOptions(value);
		//this.initTree();
	}

	public relationshipSuccessCallback(value: any) : void | PromiseLike<void>
	{
		if(value != null)
		{
			for(var i in value.entities)
			{
				this.selectedItems.push(value.entities[i][this._idAttribute]);
			}
		}
		if(this._linkedEntityFetchXmlResource != null)
		{
			this.contextObj.webAPI.retrieveMultipleRecords(this._linkedEntityName, "?fetchXml=" + this._linkedEntityFetchXmlResource, 5000).then(this._successCallback, this.errorCallback);
		}
		else
		{
  		    this.contextObj.webAPI.retrieveMultipleRecords(this._linkedEntityName, "?$orderby=" + this._nameAttribute + " asc", 5000).then(this._successCallback, this.errorCallback);
		}
	}

	public errorCallback(value: any)
	{
		alert(value);
	}

	public setErrorState()
	{
		if(this._maxRequiredSelections != null && this._numRelatedSelections>this._maxRequiredSelections) {
			this._errorLabelText = "This field accepts no more than " + this._maxRequiredSelections + " selections.";
			this.labelElement.innerHTML = this._errorLabelText;
		  	this._isError = true;
			//this.errorCallback(this.labelElement.innerHTML);
		} else if(this._minRequiredSelections != null && this._numRelatedSelections<this._minRequiredSelections) {
			this._errorLabelText = "This field requires at least " + this._minRequiredSelections + " selections.";
			this.labelElement.innerHTML = this._errorLabelText;
		  	this._isError = true;
			//this.errorCallback(this.labelElement.innerHTML);
		} else {
			this.labelElement.innerHTML = "";
		  	this._isError = false;
		}
	  	
	}

	public setReadonly(): void
	{
		(<HTMLElement>this.container.firstElementChild).style.display = this.contextObj.mode.isControlDisabled == false ? "none" : "block";
	}
	
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		if(this._isValidState == false) return;
		// Add code to update control view
		this.contextObj = context;
		this.setReadonly();
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		if (this._isValidState == false)
		{
		  return {
			value: ""
		  };
		}
		else if (this._isError == true) {
			this.container.appendChild(this.labelElement);
			return {
				value: ""
			};
		}
		else {
		  return {
			value: "NTONDATA:"+JSON.stringify(this._relData)
		  };
		}
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	public selectAction(action : string, id : string)
	{
		//Control is present on a New record form
		if((<any>this.contextObj).page.entityId == null
		   || (<any>this.contextObj).page.entityId == "00000000-0000-0000-0000-000000000000")
		{
			if(action == "select")
			{
				//debugger;
				var act = new DataAction();
				act.associate = true;
				act.guid = id;
				this._relData.actions.push(act);
				this._numRelatedSelections++;
			}
			else{
				for(var i=0; i < this._relData.actions.length; i++)
				{
					var act = this._relData.actions[i];
					if(act.guid == id){
						this._relData.actions.splice(i,1);
						break;
					}
				}
				this._numRelatedSelections--;
			}
		}

		//Control is present on an Edit or View record form
		else {
			var url: string = (<any>Xrm).Utility.getGlobalContext().getClientUrl();
			var recordUrl: string = url + "/api/data/v9.1/"+ this._mainEntityCollectionName + "(" + (<any>this.contextObj).page.entityId + ")";
			
			if(action == "select")
			{
				//See himbap samples here: http://himbap.com/blog/?p=2063
				var associate = {
					"@odata.id": recordUrl
				};
				
				var req = new XMLHttpRequest();
				req.open("POST", url + "/api/data/v9.1/"+ this._linkedEntityCollectionName +"(" + id + ")/" + this._relationshipName + "/$ref", true);
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.onreadystatechange = function() {
					if (this.readyState == 4 /* complete */ ) {
						req.onreadystatechange = null;
						if (this.status == 204) {
							//alert('Record Associated');
						} else {
							var error = JSON.parse(this.response).error;
							alert(error.message);
						}
					}
				};
				req.send(JSON.stringify(associate));
				this._numRelatedSelections++;		
			}
			else if(action == "unselect")
			{
				var req = new XMLHttpRequest();
				req.open("DELETE",url + "/api/data/v9.1/"+ this._linkedEntityCollectionName +"(" + id + ")/" + this._relationshipName + "/$ref"+"?$id="+recordUrl, true);
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.onreadystatechange = function() {
					if (this.readyState == 4 /* complete */ ) {
						req.onreadystatechange = null;
						if (this.status == 204) {
							//alert('Record Disassociated');
						} else {
							var error = JSON.parse(this.response).error;
							alert(error.message);
						}
					}
				};
				req.send();
				this._numRelatedSelections--;
			}
		}
		this.setErrorState();
		this._notifyOutputChanged();
	}

}