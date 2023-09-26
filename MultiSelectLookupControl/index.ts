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
        let relationShipName = ((this.context.utils) as any)?._customControlProperties?.descriptor?.Parameters?.["RelationshipName"];
        let entityToOpen = ((this.context.utils) as any)?._customControlProperties?.descriptor?.Parameters?.["TargetEntityType"];
        if (!entityToOpen) {
            entityToOpen = this.context.parameters.records.getTargetEntityType();

        }
        let primaryEntity = (this.context as any)?.page?.entityTypeName;
        let primaryEntityId = ((this.context) as any)?.page?.entityId;
        var disassociateRequest = {
            getMetadata: () => ({
                boundParameter: null,
                parameterTypes: {},
                operationType: 2,
                operationName: "Disassociate"
            }),


            relationship: relationShipName,


            target: {
                entityType: primaryEntity,
                id: primaryEntityId
            },

            relatedEntityId: id
        }

        let excuteRef = ((this.context.webAPI) as any).execute;
        excuteRef(disassociateRequest).then(
            (success: any) => {
                console.log("Success", success);
                this.context.parameters.records.refresh();
            },
            (error: any) => {
                console.log("Error", error);
            }
        )
    }
    OpenLookups = () => {
        let relationShipName = ((this.context.utils) as any)?._customControlProperties?.descriptor?.Parameters?.["RelationshipName"];
        let entityToOpen = ((this.context.utils) as any)?._customControlProperties?.descriptor?.Parameters?.["TargetEntityType"];
        if (!entityToOpen) {
            entityToOpen = this.context.parameters.records.getTargetEntityType();

        }
        let primaryEntity = ((this.context) as any)?.page?.entityTypeName;
        let primaryEntityId = ((this.context) as any)?.page?.entityId;
        var lookupOptions: any =
        {
            defaultEntityType: entityToOpen,
            entityTypes: [entityToOpen],
            allowMultiSelect: true,
            disableMru: true
        };
        var manyToManyAssociateRequest = {
            getMetadata: () => ({
                boundParameter: null,
                parameterTypes: {},
                operationType: 2,
                operationName: "Associate"
            }),
            relationship: relationShipName,

            target: {
                entityType: primaryEntity,
                id: primaryEntityId
            },
        }
        let excuteRef = ((this.context.webAPI) as any).execute;
        this.context.utils.lookupObjects(lookupOptions).then((res) => {
            console.log(res);
            if (res && res.length > 0) {
                (manyToManyAssociateRequest as any).relatedEntities =
                    res.map(it => { return { id: it.id.replace('{', '').replace('}', ''), entityType: it.entityType } });
                if (excuteRef) {
                    excuteRef(manyToManyAssociateRequest).then(
                        (success: any) => {
                            console.log("Success", success);
                            this.context.parameters.records.refresh();
                        },
                        (error: any) => {
                            console.log("Error", error);
                        }
                    );
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
            let type = "contact";
            if ((records as any)[key].getNamedReference)
                type = (records as any)[key]?.getNamedReference()?.logicalName
            recordOfArray.push({
                "id": key,
                url: '?pagetype=entityrecord&etn=' + type + '&id=' + key,
                entityType: type,
                name: (records as any)[key]._record.fields.fullname.value //(records as any)[key].getValue("fullname")


            })
        }

        let customControlProperties = (this.context?.utils as any)?._customControlProperties;
        let labelEnabled = customControlProperties?.descriptor?.ShowLabel;
        let labelText = customControlProperties?.descriptor?.Label;
        let image = context.parameters.controlImageURl.raw ? context.parameters.controlImageURl.raw : "";
        Render(recordOfArray,labelEnabled,labelText,image,this.OpenLookups,this.RemoveItemFromGrid,this.container);
        // ReactDOM.render(<MultiSelectLookup lookups={recordOfArray}  labelEnabled={labelEnabled}
        //     labelText={labelText} image={image} OpenLookupObject={this.OpenLookups} 
        //     RemoveItemFromGrid={this.RemoveItemFromGrid} />, this.container);

    //    let rootContainer= createRoot(this.container);     //    React.
    //    rootContainer.render(
    //     <MultiSelectLookup />
    //     );
        // ReactDOM.render(
        //     React.createElement(MultiSelectLookup, {
        //         lookups: recordOfArray,
        //         labelEnabled: labelEnabled,
        //         labelText: labelText,
        //         image: image,
        //         OpenLookupObject: this.OpenLookups,
        //         RemoveItemFromGrid: this.RemoveItemFromGrid
        //     }),
        //     this.container
        // );
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
