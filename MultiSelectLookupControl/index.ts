import { IInputs, IOutputs } from "./generated/ManifestTypes";
// eslint-disable-next-line no-undef
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import * as React from "react";
import * as ReactDOM from "react-dom";
import MultiSelectLookup from './MultiSelectLookup'
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
    public RemoveItemFromGrid(id: any) {
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
                this.notifyOutputChanged();

            },
            (error: any) => {
                console.log("Error", error);
            }
        )
    }
    public OpenLookups() {

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
       let changeDeactor= this.notifyOutputChanged;
        this.context.utils.lookupObjects(lookupOptions).then((res) => {
            console.log(res);
            if (res && res.length>0) {
                (manyToManyAssociateRequest as any).relatedEntities =
                    res.map(it => { return { id: it.id.replace('{', '').replace('}', ''), entityType: it.entityType } });
                if (excuteRef) {
                    excuteRef(manyToManyAssociateRequest).then(
                        (success: any) => {
                            console.log("Success", success);
                            changeDeactor();
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

    public getEntityMetaData(name: string) {
        this.context.utils.getEntityMetadata(name).then((res) => {
            res.EntitySetName
        }, (err) => {

        })
    }
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view 
        // let gridParams = {
        //     "loading": false,
        //     "columns": [
        //         {
        //             "name": "fullname",
        //             "displayName": "Full Name",
        //             "dataType": "SingleLine.Text",
        //             "alias": "fullname",
        //             "order": 0,
        //             "visualSizeFactor": 300,
        //             "isHidden": false,
        //             "imageProviderWebresource": "",
        //             "imageProviderFunctionName": "",
        //             "isPrimary": true,
        //             "cellType": "",
        //             "disableSorting": false
        //         },
        //         {
        //             "name": "emailaddress1",
        //             "displayName": "Email",
        //             "dataType": "SingleLine.Email",
        //             "alias": "emailaddress1",
        //             "order": 1,
        //             "visualSizeFactor": 150,
        //             "isHidden": false,
        //             "imageProviderWebresource": "",
        //             "imageProviderFunctionName": "",
        //             "isPrimary": false,
        //             "cellType": "",
        //             "disableSorting": false
        //         },
        //         {
        //             "name": "parentcustomerid",
        //             "displayName": "Company Name",
        //             "dataType": "Lookup.Customer",
        //             "alias": "parentcustomerid",
        //             "order": 2,
        //             "visualSizeFactor": 150,
        //             "isHidden": false,
        //             "imageProviderWebresource": "",
        //             "imageProviderFunctionName": "",
        //             "isPrimary": false,
        //             "cellType": "",
        //             "disableSorting": false
        //         },
        //         {
        //             "name": "telephone1",
        //             "displayName": "Business Phone",
        //             "dataType": "SingleLine.Phone",
        //             "alias": "telephone1",
        //             "order": 3,
        //             "visualSizeFactor": 125,
        //             "isHidden": false,
        //             "imageProviderWebresource": "",
        //             "imageProviderFunctionName": "",
        //             "isPrimary": false,
        //             "cellType": "",
        //             "disableSorting": false
        //         },
        //         {
        //             "name": "statecode",
        //             "displayName": "Status",
        //             "dataType": "OptionSet",
        //             "alias": "statecode",
        //             "order": 4,
        //             "visualSizeFactor": 125,
        //             "isHidden": false,
        //             "imageProviderWebresource": "",
        //             "imageProviderFunctionName": "",
        //             "isPrimary": false,
        //             "cellType": "",
        //             "disableSorting": false
        //         }
        //     ],
        //     "error": false,
        //     "errorMessage": null,
        //     "innerError": null,
        //     "sortedRecordIds": [
        //         "a1cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //         "a3cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //         "a5cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //         "a7cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //     ],
        //     "records": {
        //         "a1cb8e99-bb2a-ed11-9db1-000d3a53a6df": {
        //             "_record": {
        //                 "initialized": 2,
        //                 "identifier": {
        //                     "etn": "contact",
        //                     "id": {
        //                         "guid": "a1cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                     }
        //                 },
        //                 "fields": {
        //                     "fullname": {
        //                         "value": "Yvonne McKay",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "emailaddress1": {
        //                         "value": "someone_a@example.com",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "parentcustomerid": {
        //                         "reference": {
        //                             "etn": "account",
        //                             "id": {
        //                                 "guid": "8dcb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                             },
        //                             "name": "Fourth Coffee"
        //                         },
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "statecode": {
        //                         "label": "Active",
        //                         "valueString": "0",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     }
        //                 }
        //             },
        //             "_columnAliasNameMap": {},
        //             "_primaryFieldName": "fullname",
        //             "_isDirty": false,
        //             "_entityReference": {
        //                 "_etn": "contact",
        //                 "_id": "a1cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //                 "_name": "Yvonne McKay"
        //             }
        //         },
        //         "a3cb8e99-bb2a-ed11-9db1-000d3a53a6df": {
        //             "_record": {
        //                 "initialized": 2,
        //                 "identifier": {
        //                     "etn": "contact",
        //                     "id": {
        //                         "guid": "a3cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                     }
        //                 },
        //                 "fields": {
        //                     "fullname": {
        //                         "value": "Susanna Stubberod",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "emailaddress1": {
        //                         "value": "someone_b@example.com",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "parentcustomerid": {
        //                         "reference": {
        //                             "etn": "account",
        //                             "id": {
        //                                 "guid": "8fcb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                             },
        //                             "name": "Litware, Inc."
        //                         },
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "statecode": {
        //                         "label": "Active",
        //                         "valueString": "0",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     }
        //                 }
        //             },
        //             "_columnAliasNameMap": {},
        //             "_primaryFieldName": "fullname",
        //             "_isDirty": false,
        //             "_entityReference": {
        //                 "_etn": "contact",
        //                 "_id": "a3cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //                 "_name": "Susanna Stubberod"
        //             }
        //         },
        //         "a5cb8e99-bb2a-ed11-9db1-000d3a53a6df": {
        //             "_record": {
        //                 "initialized": 2,
        //                 "identifier": {
        //                     "etn": "contact",
        //                     "id": {
        //                         "guid": "a5cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                     }
        //                 },
        //                 "fields": {
        //                     "fullname": {
        //                         "value": "Nancy Anderson",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "emailaddress1": {
        //                         "value": "someone_c@example.com",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "parentcustomerid": {
        //                         "reference": {
        //                             "etn": "account",
        //                             "id": {
        //                                 "guid": "91cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                             },
        //                             "name": "Adventure Works"
        //                         },
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "statecode": {
        //                         "label": "Active",
        //                         "valueString": "0",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     }
        //                 }
        //             },
        //             "_columnAliasNameMap": {},
        //             "_primaryFieldName": "fullname",
        //             "_isDirty": false,
        //             "_entityReference": {
        //                 "_etn": "contact",
        //                 "_id": "a5cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //                 "_name": "Nancy Anderson"
        //             }
        //         },
        //         "a7cb8e99-bb2a-ed11-9db1-000d3a53a6df": {
        //             "_record": {
        //                 "initialized": 2,
        //                 "identifier": {
        //                     "etn": "contact",
        //                     "id": {
        //                         "guid": "a7cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                     }
        //                 },
        //                 "fields": {
        //                     "fullname": {
        //                         "value": "Maria Campbell",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "emailaddress1": {
        //                         "value": "someone_d@example.com",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "parentcustomerid": {
        //                         "reference": {
        //                             "etn": "account",
        //                             "id": {
        //                                 "guid": "93cb8e99-bb2a-ed11-9db1-000d3a53a6df"
        //                             },
        //                             "name": "Fabrikam, Inc."
        //                         },
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     },
        //                     "statecode": {
        //                         "label": "Active",
        //                         "valueString": "0",
        //                         "timestamp": "2023-09-25T16:49:22.449Z",
        //                         "validationResult": {
        //                             "errorId": null,
        //                             "errorMessage": null,
        //                             "isValueValid": true,
        //                             "userInput": null,
        //                             "isOfflineSyncError": false
        //                         }
        //                     }
        //                 }
        //             },
        //             "_columnAliasNameMap": {},
        //             "_primaryFieldName": "fullname",
        //             "_isDirty": false,
        //             "_entityReference": {
        //                 "_etn": "contact",
        //                 "_id": "a7cb8e99-bb2a-ed11-9db1-000d3a53a6df",
        //                 "_name": "Maria Campbell"
        //             }
        //         }
        //     },
        //     "sorting": [],
        //     "filtering": {
        //         "aliasMap": {
        //             "fullname": "fullname",
        //             "emailaddress1": "emailaddress1",
        //             "parentcustomerid": "parentcustomerid",
        //             "telephone1": "telephone1",
        //             "statecode": "statecode"
        //         }
        //     },
        //     "paging": {
        //         "pageNumber": 1,
        //         "totalResultCount": 14,
        //         "firstPageNumber": 1,
        //         "lastPageNumber": 1,
        //         "pageSize": 4,
        //         "hasNextPage": true,
        //         "hasPreviousPage": false
        //     },
        //     "linking": {},
        //     "entityDisplayCollectionName": "Contacts",
        //     "_capabilities": {
        //         "hasRecordNavigation": true
        //     }
        // };
        let gridParams= context.parameters.records;
        let recordOfArray = [];
        let records = gridParams.records;//context.parameters.records.records || gridParams;
        for (let key in records) {
            let type = "contact";
            if ((records as any)[key].getNamedReference)
                type = (records as any)[key]?.getNamedReference()?.logicalName
            recordOfArray.push({
                "id": key,
                entityType: type,
                name: (records as any)[key]._record.fields.fullname.value //(records as any)[key].getValue("fullname")


            })
        }
        ReactDOM.render(
            React.createElement(MultiSelectLookup, { 
                context: context, 
                lookups: recordOfArray,
                OpenLookupObject: this.OpenLookups,
                RemoveItemFromGrid: this.RemoveItemFromGrid
            }),
            this.container
        ); 
    }

   
    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

}
