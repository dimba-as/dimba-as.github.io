import * as React from 'react';
import { Listing } from '../../../models/models';
import "./materialize-custom.scss";
import {format } from 'date-fns';
export interface IListingCardProps {
  Listing: Listing;
}

export interface IListingCardState {

}

export default class ListingCard extends React.Component<IListingCardProps, IListingCardState> {

  constructor(props: IListingCardProps) {
    super(props);
    this.state = {};
  }

  public render(): React.ReactElement<IListingCardProps> {
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
        </div>
      </div>
    );
  }
}
