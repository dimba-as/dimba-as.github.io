import { Log } from "@microsoft/sp-core-library";
import { override } from "@microsoft/decorators";
import { ExtensionContext } from "@microsoft/sp-extension-base";
import * as React from "react";
import { sp } from '@pnp/sp';
import { PrimaryButton, Spinner, SpinnerSize, Stack } from "office-ui-fabric-react";
import styles from "./Sendbutton.module.scss";
import Utils from "../../../common/Utils";

export interface ISendbuttonProps {
  id: number;
  listId: string;
  status: string;
  context: ExtensionContext;
  functionUrl: string;
}

export interface ISendbuttonState {
  text: string;
  sending: boolean;
}

const LOG_SOURCE: string = "CreateTeamButton";

export default class Sendbutton extends React.Component<ISendbuttonProps, ISendbuttonState> {

  public Utility: Utils;

  constructor(props: ISendbuttonProps) {
    super(props);
    this.Utility = new Utils(props.context);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      text: "Opprett Team",
      sending: false
    };
  }

  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, "React Element: Create Team button mounted");
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, "React Element: Create Team button unmounted");
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        {this.state.sending ?
          <Stack {...{ horizontal: true }} tokens={{ childrenGap: 10 }}>
            <div>Oppretter..</div>
            <Spinner size={SpinnerSize.xSmall} />
          </Stack>
          :
          <PrimaryButton
            text={this.state.text}
            onClick={this.clicked.bind(this)}
            aria-roledescription={"Opprett Team"}
            iconProps={{ iconName: 'Spinner' }}
            disabled={this.props.status!=="Bestilt"}
          />
        }
      </div>
    );
  }

  private clicked = async () => {

    let item: any = await sp.web.lists.getById(this.props.listId).items.getById(this.props.id).get();

    this.setState({
      sending: true
    });
    //https://afgteamsadmin.azurewebsites.net/api/Teams/TeamsAdminStarter/{id}?code={code}
    const functionUrlWithParam = this.props.functionUrl.replace("{id}", this.props.id.toString());
    let response = await this.Utility.createTeam(functionUrlWithParam);


    this.setState({
      // text: (response == "OK") ? 'Sendt' : 'Send',
      text: 'Opprett Team',
      sending: false
    });

  }

}