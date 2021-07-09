import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  autobind,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent,
  Spinner
} from 'office-ui-fabric-react';

interface IInfoDialogContentProps {
    //title: string;
    message: string;
    close: () => void;
  }

  class InfoDialogContent extends React.Component<IInfoDialogContentProps, {}> {
    
    constructor(props) {
      super(props);
    }
  
    public render(): JSX.Element {
      return <DialogContent            
        title={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={false}
      >           
        <DialogFooter>
          <PrimaryButton text='OK' title='OK' onClick={this.props.close} />         
        </DialogFooter>
      </DialogContent>;
    }
  }

  export default class InfoDialog extends BaseDialog {
    public title: string;
    public message: string;
  
    public render(): void {
      ReactDOM.render(<InfoDialogContent
        close={ this.close }
        //title={ this.title }
        message={ this.message }        
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
  }