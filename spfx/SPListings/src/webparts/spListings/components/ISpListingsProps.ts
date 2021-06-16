import { Listing } from '../../../models/models';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpListingsProps {
  context:WebPartContext;
}
export interface ISpListingsState {
  listings: Listing[];
}
