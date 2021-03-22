import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ISCWWebPartProps {
    listTemplate: string;
}
export default class SCWWebPart extends BaseClientSideWebPart<ISCWWebPartProps> {
    render(): void;
    onInit(): Promise<void>;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SCWWebPart.d.ts.map