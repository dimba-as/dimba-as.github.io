import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {  
PrimaryButton,
DefaultButton,
DialogFooter,
DialogContent

} from 'office-ui-fabric-react';
import { sp } from "@pnp/sp/presets/all";
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

interface IDenyDialogContentProps {
    message: string;
    close: () => void;
    submit: (value: boolean) => void;
  
  }
  interface IDenyDialogContentState {
    sending: boolean;
   
  
  }
class DenyDialogContent extends React.Component<IDenyDialogContentProps, IDenyDialogContentState> {
  
    constructor(props) {
        super(props);
        this.state = {sending:false};
    }
  
    public render(): JSX.Element {
        return <DialogContent
        title='AVSLÅ FORESPØRSEL'
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
        >
        <DialogFooter>
            <DefaultButton text='Avbryt' title='Avbryt' onClick={this.props.close} disabled={this.state.sending} />
            <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(true); this.setState({sending:true});}} disabled={this.state.sending} />
        </DialogFooter>
        </DialogContent>;
    }
  
 
  }
  export default class DenyDialog extends BaseDialog {
    public message: string;
    public listId: string;
    public listItemId:number;  
    public context:ListViewCommandSetContext;
    public render(): void {
        ReactDOM.render(<DenyDialogContent
        close={ this.close }
        message={ this.message }
        submit={ this._submit }
        />, this.domElement);
    }
  
    public getConfig(): IDialogConfiguration {
        return {
        isBlocking: false
        };
    }
  
    protected onAfterClose(): void {
        super.onAfterClose();
        
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
  
    private _submit = (value: boolean) => {
        sp.setup({
            spfxContext: this.context
          });
        sp.web.lists.getById(this.listId).items.getById(this.listItemId).update({avAnbudsForesporsel:"Avslått"}).then(()=>{
            location.href = location.href;  
          });
    }
  }