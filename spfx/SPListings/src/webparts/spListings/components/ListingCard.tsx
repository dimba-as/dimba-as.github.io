import * as React from 'react';
import { Listing } from '../../../models/models';
import "./materialize-custom.scss";
import { format } from 'date-fns';
import Panel from '../components/Panel/Panel';
import { PanelPosition } from '../components/Panel/Panel';
import { PrimaryButton } from '@fluentui/react';
import { flagInterest, getSPListings } from '../../../apiHelper';
import { fromPairs } from 'lodash';
import { WebPartContext } from '@microsoft/sp-webpart-base';


export interface IListingCardProps {
  Listing: Listing;
  panelPosition?: PanelPosition;
  context: WebPartContext;
}

export interface IListingCardState {
  isOpen?: boolean;
  sendMessage?: string;
  sending: boolean;
}

export default class ListingCard extends React.Component<IListingCardProps, IListingCardState> {

  constructor(props: IListingCardProps) {
    super(props);
    this.state = { sendMessage: "", sending: false };
  }
  private openApplication() {
    this.setState({
      isOpen: !this.state.isOpen
    });
  }
  private async sendApplication() {
    const user = this.props.context.pageContext.user;
    this.setState({ sending: true });
    let from = user.loginName;


    const props = {
      webUrl: this.props.Listing.SPSiteUrl,
      listItemId: this.props.Listing.ListItemID,
      from: from,
      displayName: this.props.context.pageContext.user.displayName
    };
    const res = await flagInterest(props);

    this.setState({
      sendMessage: "Takk for din interesse. Du vil få tilbakemelding av entreprenøren."
    });

  }
  private onPanelClosed() {
    this.setState({
      isOpen: false, sendMessage: "", sending: false
    });
  }
  public render(): React.ReactElement<IListingCardProps> {
    const panelPosition = !this.props.panelPosition && this.props.panelPosition !== 0
      ? PanelPosition.Right : this.props.panelPosition;
    return (
      <div className="col s4 m4">
        <div className="card large light-blue darken-2">
          <div className="card-title white-text indigo darken-4">{this.props.Listing.Title}</div>
          <div className="card-content white-text indigo darken-1">
            <div className="col s12 m12 l12 description"><p dangerouslySetInnerHTML={{ __html: this.props.Listing.avCustomAnbudsbeskrivelse }} /></div>
          </div>
          <div className="card-content light-blue darken-2 white-text ">

            <div className="row">
              <div className="col s6 m6 l6">Adresse:</div><div className="col s6 m6 l6">{this.props.Listing.avAdresseOWSTEXT}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Anbudsfrist:</div><div className="col s6 m6 l6">{format(new Date(this.props.Listing.avAnbudsfristOWSDATE), 'dd.MM.yyyy')}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Oppstart år:</div><div className="col s6 m6 l6">{Math.round(this.props.Listing.avOppstartAarOWSNMBR)}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Oppstart periode:</div><div className="col s6 m6 l6">{this.props.Listing.avOppstartsperiodeOWSCHCS}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Ferdig år:</div><div className="col s6 m6 l6">{Math.round(this.props.Listing.avFerdigstillelseAarOWSNMBR)}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Ferdig periode:</div><div className="col s6 m6 l6">{this.props.Listing.avFerdigstillelseperiodeOWSCHCS}</div>
            </div>
            <div className="row">
              <div className="col s6 m6 l6">Fag:</div><div className="col s6 m6 l6">{this.props.Listing.avCustomFag}</div>
            </div>
          </div>
          <div className="card-action">
            <a href="#" onClick={this.openApplication.bind(this)}>Meld din interesse</a>
          </div>
        </div>
        <Panel isOpen={this.state.isOpen} position={panelPosition} onDismiss={this.onPanelClosed.bind(this)}>
          <div className="table">
            <h2>{this.props.Listing.Title}</h2>
            <div className="row">
              <div className="col s12 m12 l12">Ja! Jeg vil melde min interesse for dette prosjektet.</div>
            </div>
            <div className="row">
              <div className="col s12 m12 l12">&nbsp;</div>
            </div>
            <div className="row">
              <div className="col s12 m12 l12">&nbsp;</div>
            </div>
            <div className="row buttonRow">
              <div className="col s12 m12 l12"> <PrimaryButton disabled={this.state.sending} text="Send inn forespørsel" onClick={this.sendApplication.bind(this)} /></div>
            </div>
            <div className="row buttonRow">
              <div className="col s12 m12 l12" style={{ padding: '20px' }}><h3>{this.state.sendMessage}</h3></div>
            </div>
            {/* <div className="row buttonRow">
              <div className="col s12 m12 l12">{this.context}</div>
            </div> */}
          </div>
        </Panel>
      </div>
    );
  }
}
