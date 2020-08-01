import {SPHttpClient} from '@microsoft/sp-http';
export interface ICrudeOprationReactProps {
  description: string;
  siteUrl: string;
  spHttpClient:SPHttpClient;
}
