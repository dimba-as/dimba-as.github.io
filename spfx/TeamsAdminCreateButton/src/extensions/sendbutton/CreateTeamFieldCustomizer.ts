import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Guid, Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as strings from 'CreateTeamFieldCustomizerStrings';
import Sendbutton, { ISendbuttonProps } from './components/Sendbutton';
import { sp } from '@pnp/sp';

export interface ICreateTeamFieldCustomizerProperties {
  // This is an example; replace with your own property 
}

const LOG_SOURCE: string = 'CreateTeamFieldCustomizer';

export default class CreateTeamFieldCustomizer
  extends BaseFieldCustomizer<ICreateTeamFieldCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    console.info(Environment);
    console.info("Initializing CreateTeamFieldCustomizer...");
    Log.info(LOG_SOURCE, 'Activated CreateTeamFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "CreateTeamFieldCustomizer" and "${strings.Title}"`);
  }

  @override
  public async onRenderCell(event: IFieldCustomizerCellEventParameters): Promise<void> {
    // Use this method to perform your custom cell rendering. 
    const id: number = event.listItem.getValueByName('ID').toString();
    const status: string = event.listItem.getValueByName('avStatus').toString();
    const listId: string = this.context.pageContext.list.id.toString();
       
    let functionUrl = 'https://dimbateamsadmin.azurewebsites.net/api/TeamsAdminStarter?code=E6WQhnw00v2iJwViLiUe5d90/kwryEHq9vTZv4QAUKDucpfXEjiSrw==&listItemId={id}';
    // let configItems= await sp.web.lists.getByTitle("Konfigurasjon").items.filter("Title eq 'FunctionUrl'").usingCaching().get();
    // if(configItems && configItems.length>0){
    //   functionUrl = configItems[0].avVerdi;
    // }
    //Get from Config
    console.log(functionUrl);
    const sendBtn: React.ReactElement<{}> = React.createElement(Sendbutton, {
      context: this.context,
      id: id,
      listId:listId,
      status: status,
      functionUrl: functionUrl
    } as ISendbuttonProps);
    if(functionUrl===''){
      ReactDOM.render(React.createElement("Dette funker ikke"), event.domElement);
    }else{
      ReactDOM.render(sendBtn, event.domElement);
    }
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }

}
