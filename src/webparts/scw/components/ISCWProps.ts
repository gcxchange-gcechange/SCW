import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISCWProps {
  listTemplate: string;
  context: WebPartContext;
  prefLang: string;
}
