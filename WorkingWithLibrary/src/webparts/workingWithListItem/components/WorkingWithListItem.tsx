import * as React from 'react';
import styles from './WorkingWithListItem.module.scss';
import { IWorkingWithListItemProps } from './IWorkingWithListItemProps';
import {IWorkingWithListItemState} from './IWorkingWithListItemState';
import {Items, sp} from '@pnp/sp/presets/all';
import { escape } from '@microsoft/sp-lodash-subset';
import {PrimaryButton,TextField} from 'office-ui-fabric-react';

export default class WorkingWithListItem extends React.Component<IWorkingWithListItemProps, IWorkingWithListItemState> {
 
  constructor(props:IWorkingWithListItemProps){
    super(props);
    this.state={
      items:[],
      textValue:'',
    };
  }
 
public async componentDidMount() : Promise<void> {
  sp.setup(this.props.context);
  await this.loaddata();
}

private async loaddata(): Promise<void>{
  const listitems: any[]= await sp.web.lists.getByTitle('Crew').items.select('Title').get();
  this.setState({
    items:listitems.map(i=> i.Title)
  })
}
 
private async AddItem(){
 const additem = await sp.web.lists.getByTitle('Crew').items.add({
   Title:'abc'
 });
 await this.loaddata();
}

private async UpdateItem() : Promise<void> {
const updateitem= await sp.web.lists.getByTitle('Crew').items.getById(9).update({
  Title :'Item Updated by PnPJS'
});
await this.loaddata();
}


private async DeleteItem() : Promise<void> {
const deleteitem=await sp.web.lists.getByTitle('Crew').items.getById(9).delete();
await this.loaddata();
}

private onTextValueChnaged = ( event, newvalue:string) :void =>
{
this.setState({
  textValue:newvalue,
})
}

private async filesave(): Promise<void> {
  const str=['internal','external'];
  let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
  await sp.web.lists.getByTitle('demo').get().then(async (library) => {
    let libraryUrl = library.DocumentTemplateUrl.split('Forms');
    sp.web.getFolderByServerRelativeUrl(libraryUrl[0])
      .files.add(myfile.name, myfile, true).then((f) => {
        f.file.listItemAllFields.get().then((item) => {
            sp.web.lists.getByTitle("Demo").items.getById(item.Id).update({
              ImmersionType: { results: str },
            });
          }).then((response :any):void => { console.log('success')})
        });
      });

}

  public render(): React.ReactElement<IWorkingWithListItemProps> {
    return (
      <div className={ styles.workingWithListItem }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <div>List Items</div>
    <div>{this.state.items.map(i=> <li> {i}</li>)}</div>
            </div>
            <p><TextField onChange={this.onTextValueChnaged} value={'Please Enter Value'} /></p>
            <input type="file" name="myFile"   id="newfile"   ></input>
            <PrimaryButton onClick={this.filesave} text="Save" />
            <PrimaryButton text="Add Item" onClick={this.AddItem}/> <br/><br/>
            <PrimaryButton text="Update Item" onClick={this.UpdateItem}/> <br/><br/>
            <PrimaryButton text="Delete Item" onClick={this.DeleteItem}/> <br/><br/>
          </div>
        </div>
      </div>
    );
  }
}
