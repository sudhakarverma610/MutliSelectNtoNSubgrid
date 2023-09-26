import * as React from 'react';
import { IInputs } from './generated/ManifestTypes';

const MultiSelectLookup = (props: any) => {
  let lookupObj: any[] = props.lookups;
  let defaultformMapptedlookup = lookupObj.map(it => {
    return { ...it, url: '?pagetype=entityrecord&etn=' + it.entityType + '&id=' + it.id }
  });
  const [formMapptedlookup, setformMapptedlookup] = React.useState(defaultformMapptedlookup);
  React.useEffect(() => {
    console.log('useEffect calling')
    setformMapptedlookup(lookupObj.map(it => {
      return { ...it, url: '?pagetype=entityrecord&etn=' + it.entityType + '&id=' + it.id }
    }));
  }, lookupObj);
  let context = props.context as any;//ComponentFramework.Context<IInputs>;
  let notifyOutputChanged: any = props.notifyOutputChanged;
  let labelEnabled=context?.utils?._customControlProperties?.descriptor?.ShowLabel;
  let labelText=context?.utils?._customControlProperties?.descriptor?.Label;
  console.log(props.lookups);
  let image = context.parameters.controlImageURl.raw ? context.parameters.controlImageURl.raw : "";
  const OpenLookup = () => {
    props.OpenLookupObject();
  }
  const removeItem = (selectedIt: any) => {
    setformMapptedlookup(formMapptedlookup.filter(it => !(it.id == selectedIt.id && it.entityType == selectedIt.entityType)));
    props.RemoveItemFromGrid(selectedIt.id);
  }
  return <>
    <div className='multiSelectLookup'>
      {
        labelEnabled &&

        <label>{labelText}</label>
      }
      <div className='lookupObjectListing'>
        {formMapptedlookup.map(it => {
          return <div className='lookupObjectListing-item' key={it.id}>
            {image && <img src={image} alt="" height={20} width={20} />}
            <a href={it.url}>{it.name}</a>
            <span className='deleticon' onClick={() => removeItem(it)}>
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 384 512"><path d="M376.6 84.5c11.3-13.6 9.5-33.8-4.1-45.1s-33.8-9.5-45.1 4.1L192 206 56.6 43.5C45.3 29.9 25.1 28.1 11.5 39.4S-3.9 70.9 7.4 84.5L150.3 256 7.4 427.5c-11.3 13.6-9.5 33.8 4.1 45.1s33.8 9.5 45.1-4.1L192 306 327.4 468.5c11.3 13.6 31.5 15.4 45.1 4.1s15.4-31.5 4.1-45.1L233.7 256 376.6 84.5z" /></svg>
            </span>

          </div>
        }

        )}
        <span className='divSearch ' onClick={OpenLookup}>
        </span>
      </div>
    </div>
  </>
}
export default MultiSelectLookup;