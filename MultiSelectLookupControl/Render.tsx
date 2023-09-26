import React from "react";
import MultiSelectLookup from "./MultiSelectLookup";
import ReactDOM from "react-dom";
import { createRoot } from 'react-dom/client';

export const Render = (recordOfArray: any, labelEnabled: any, labelText: any, image: any, OpenLookups: any, RemoveItemFromGrid: any, container: any) => {

    createRoot(container).render(
        <MultiSelectLookup lookups={recordOfArray} labelEnabled={labelEnabled}
            labelText={labelText} image={image} OpenLookupObject={OpenLookups}
            RemoveItemFromGrid={RemoveItemFromGrid} />
    )

}