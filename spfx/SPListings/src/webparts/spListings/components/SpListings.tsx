import * as React from 'react';
import styles from './SpListings.module.scss';
import { ISpListingsProps, ISpListingsState } from './ISpListingsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { getSPListings } from '../../../apiHelper';
import ListingCard from './ListingCard';
import { setup as pnpSetup } from "@pnp/common";
export default class SpListings extends React.Component<ISpListingsProps, ISpListingsState> {
  constructor(props) {
    super(props);
   
    this.state = { listings: [] };
    
    this.getSPListings();

  }
  private async getSPListings() {
    const listingsResponse = await getSPListings();
    this.setState({ listings: listingsResponse });

  }
  private listings = () => {
    if (this.state.listings.length === 0) return <span></span>;
    return (
      <div className={styles.row}>

        {
          this.state.listings.map((listing, i) => {
            return (<ListingCard Listing={listing} key={i} context={this.props.context}></ListingCard>);
          })
        }
      </div>
    );
  }

  public render(): React.ReactElement<ISpListingsProps> {
    return (

      <div className="spListingsApp">
        <div className="row">
          <div className="col s12 m12 l12">
            {this.listings()}
          </div>
        </div>
      </div>
    );
  }
}
