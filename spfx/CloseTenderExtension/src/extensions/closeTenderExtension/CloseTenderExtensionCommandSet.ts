import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'CloseTenderExtensionCommandSetStrings';
import ApproveDialog from './Approve';
import DenyDialog from './Deny';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICloseTenderExtensionCommandSetProperties {
  approveText: string;
  denyText: string;
}

const LOG_SOURCE: string = 'CloseTenderExtensionCommandSet';

export default class CloseTenderExtensionCommandSet extends BaseListViewCommandSet<ICloseTenderExtensionCommandSetProperties> {
  private _selectedRow = null;
  private _listId = null;
  private _props = {};
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CloseTenderExtensionCommandSet');
    
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const approveCommand: Command = this.tryGetCommand('APPROVE_COMMAND');
    const denyCommand: Command = this.tryGetCommand('DENY_COMMAND');
    const libraryUrl = this.context.pageContext.list.title;
    this._listId = this.context.pageContext.list.id.toString();
    if (approveCommand && denyCommand) {
      // This command should be hidden unless exactly one row is selected.
      approveCommand.visible = false;
      denyCommand.visible =  event.selectedRows.length === 1 && libraryUrl == "Anbud";
      if (event.selectedRows.length === 1) {
        this._selectedRow = event.selectedRows[0];
        this._props = {
          listItemId: this._selectedRow.getValueByName("ID"),
          listId: this.context.pageContext.list.id.toString(),
          siteUrl: this.context.pageContext.web.absoluteUrl
        };
      }

    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'APPROVE_COMMAND':
        const dialog: ApproveDialog = new ApproveDialog();
        dialog.message = 'Vil du lukke dette anbudet?';
        dialog.props = this._props;
        dialog.show().then(() => {
          location.href = location.href;  
        });
        break;
      case 'DENY_COMMAND':
        const denyDialog: DenyDialog = new DenyDialog();
        denyDialog.message = 'Vil du lukke dette anbudet?';
        denyDialog.props = this._props;

        denyDialog.show().then(() => {
          location.href = location.href;  
        });
        break;
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
