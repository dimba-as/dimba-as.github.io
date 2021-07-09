import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {  
PrimaryButton,
DefaultButton,
DialogFooter,
DialogContent

} from 'office-ui-fabric-react';
import axios from 'axios';
interface IApproveDialogContentProps {
    message: string;
    props:any;
    close: () => void;
    submit: (value: boolean) => void;
  
  }
  interface IApproveDialogContentState {
    sending: boolean;
   
  
  }
class ApproveDialogContent extends React.Component<IApproveDialogContentProps, IApproveDialogContentState> {
  
    constructor(props) {
        super(props);
        this.state = {sending:false};
    }
  
    public render(): JSX.Element {
        return <DialogContent
        title='LUKK ANBUDET'
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
  export default class ApproveDialog extends BaseDialog {
    public message: string;
    public props: any;
  
    public render(): void {
        ReactDOM.render(<ApproveDialogContent
        props={this.props}
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
        axios.post("https://dimbateamsadmin.azurewebsites.net/api/CloseTender?code=UP7VTz2A72uj3pYbKjsZsx3nY9VDBrgzTe9GHlp8tVzenDpzf9GAVw==",this.props).then((res)=>{
            this.close();
          });
    }
  }