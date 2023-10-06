import { IInputs, IOutputs } from "./generated/ManifestTypes";
// eslint-disable-next-line no-undef
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

import * as ReactDOM from 'react-dom'; 
import { Render } from "./Render";
export class MultiSelectLookupControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    notifyOutputChanged: () => void;
    container: HTMLDivElement;
    context: ComponentFramework.Context<IInputs>;
    recordOfArray: { id: string; url: string; entityType: string; name: string; refId: string; }[];

    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code

        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        this.context = context;
        console.log('MultiSelectLookup', this.context)
         

    }
    RemoveItemFromGrid = (id: any) => {
        if(this.context.parameters.ConnectionEntityLogicalName.raw)
        this.context.webAPI.deleteRecord(this.context.parameters.ConnectionEntityLogicalName.raw,id)
        .then(
            (success: any) => {
                console.log("Success", success);
                this.context.parameters.records.refresh();
            },
            (error: any) => {
                console.log("Error", error);
            } 
        );
    }
    OpenLookups = () => {
         let entityToOpen =this.context.parameters.RelatedTableLogicalName.raw? this.context.parameters.RelatedTableLogicalName.raw:"contact"
         let primaryEntityId = ((this.context) as any)?.page?.entityId;
        var lookupOptions: any =
        {
            defaultEntityType: entityToOpen,
            entityTypes: [entityToOpen],
            allowMultiSelect: true,
            disableMru: true
        };        
       var createRecordReqRef= this.context.webAPI.createRecord;
         this.context.utils.lookupObjects(lookupOptions).then((res) => {
            console.log(res);
            if (res && res.length > 0) {
                let createRecordsPromise:any[]=[]; 
                 res.forEach(element => {
                    var dataToCreate={} as any;
                    if(this.context.parameters.RelatedTableTypeColumnLogicalName.raw&&this.context.parameters.RelatedTableTypeColumnDefaultValue.raw)
                    dataToCreate[this.context.parameters.RelatedTableTypeColumnLogicalName.raw]= + this.context.parameters.RelatedTableTypeColumnDefaultValue.raw;
                    dataToCreate[this.context.parameters.RelatedTableLookupSchemaName.raw+"@odata.bind"]=  "/"+this.context.parameters.RelatedTablePluralLogicalName.raw+"("+element.id.replace('{', '').replace('}', '')+")";
                    dataToCreate[this.context.parameters.ParentTableSchemaName.raw+"@odata.bind"]="/"+this.context.parameters.ParentTablePluralName.raw+"("+primaryEntityId+")";
                    let isRecordIsPresent=this.recordOfArray.findIndex(it=>it.id.toLowerCase() == element.id.replace('{', '').replace('}','').toLowerCase()) >-1;
                     if(this.context.parameters.ConnectionEntityLogicalName.raw && !isRecordIsPresent)
                    createRecordsPromise.push(
                        createRecordReqRef(this.context.parameters.ConnectionEntityLogicalName.raw,dataToCreate))
                });               
                
                if(createRecordsPromise.length){
                    Promise.all(createRecordsPromise).then(it=>{
                        console.log('Records Created',it);
                        this.context.parameters.records.refresh();
                    }).catch(err=>{
                        console.log('Records Not Created',err);
                    })
                }               
            }
        }, err => {
            console.log(err)
        })


    }
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        let gridParams = context.parameters.records;
        let recordOfArray = [];
        let records = gridParams.records;//context.parameters.records.records || gridParams;
        for (let key in records) {
            const contactRef:{etn:string,id:{guid:string},name:string}=records[key].getValue((this.context.parameters.RelatedTableLookupSchemaName.raw as any).toLowerCase()) as any;
            let type = contactRef.etn; 
            recordOfArray.push({
                "id": contactRef.id.guid,
                url: '?pagetype=entityrecord&etn=' + type + '&id=' + contactRef.id.guid,
                entityType: type,
                name: contactRef.name,
                refId:key


            });
        }
        let customControlProperties = (this.context?.utils as any)?._customControlProperties;
        let labelEnabled = customControlProperties?.descriptor?.ShowLabel;
        let labelText = customControlProperties?.descriptor?.Label;
        let image = context.parameters.controlImageURl.raw ? context.parameters.controlImageURl.raw : "";
        this.recordOfArray=recordOfArray;
        Render(recordOfArray,labelEnabled,labelText,image,this.OpenLookups,this.RemoveItemFromGrid,this.container);       
    }


    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {

        } as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }

}
