import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';
import { ExtensionContext } from "@microsoft/sp-extension-base";
import styles from './CreditCheckButton.module.scss';
import { PrimaryButton, Spinner, SpinnerSize, Stack } from "office-ui-fabric-react";
import { sp } from "@pnp/sp/presets/all";
import fetch from "node-fetch";
export interface ICreditCheckButtonProps {
  id: number;
  listId: string;
  status: string;
  context: ExtensionContext;
  functionUrl: string;
}
export interface ICreditCheckButtonState {
  text: string;
  sending: boolean;
}
const LOG_SOURCE: string = 'CreditCheckButton';

export default class CreditCheckButton extends React.Component<ICreditCheckButtonProps, ICreditCheckButtonState> {
  constructor(props: ICreditCheckButtonProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      text: "Kredittsjekk",
      sending: false
    };
  }

  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: CreditCheckButton mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: CreditCheckButton unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
         {this.state.sending ?
          <Stack {...{ horizontal: true }} tokens={{ childrenGap: 10 }}>
            <div>Sjekker..</div>
            <Spinner size={SpinnerSize.xSmall} />
          </Stack>
          :
          <PrimaryButton
            text={this.state.text}
            onClick={this.clicked.bind(this)}
            aria-roledescription={"Gjør en kredittsjekk"}
            iconProps={{ iconName: 'Spinner' }}
          />
        }
      </div>
    );
  
  }
  private clicked = async () => {

    //let item: any = await sp.web.lists.getById(this.props.listId).items.getById(this.props.id).get();

    this.setState({
      sending: true,
      text: 'Sjekker',
    });
    const functionUrlWithParam = this.props.functionUrl;

    const result = await fetch(functionUrlWithParam);
    // const dataJSON = await result.json();
    const dataText = await result.text();
    await sp.web.lists.getById(this.props.listId).items.getById(this.props.id).update({
      avCreditCheckResult:dataText
    });
    this.setState({
      // text: (response == "OK") ? 'Sendt' : 'Send',
      text: 'Kredittsjekk',
      sending: false
    });

  }
}
