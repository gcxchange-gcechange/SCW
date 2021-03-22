import * as React from 'react';
import { ISCWProps } from './ISCWProps';
import { ISCWState } from './ISCWState';
import { Selection } from 'office-ui-fabric-react/lib/DetailsList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseWizard } from "../../../common/components/Wizard";
export declare enum MyWizardSteps {
    None = 0,
    FirstStep = 1,
    SecondStep = 2,
    ThirdStep = 4,
    FourthStep = 8,
    LastStep = 16
}
export declare class MyWizard extends BaseWizard<MyWizardSteps> {
}
export default class SCW extends React.Component<ISCWProps, ISCWState> {
    private _teamsContext;
    protected onInit(): Promise<any>;
    private _selection;
    selection1: Selection;
    constructor(props: ISCWProps, state: ISCWState);
    private imagesTemplate;
    private _closeWizard;
    private _onValidateStep;
    private _renderMyWizard;
    private _openWizard;
    private ResetScreen;
    render(): React.ReactElement<ISCWProps>;
    protected functionTemplateImg: string;
    componentDidMount(): Promise<void>;
    private loadTemplate;
    private _getSelectionDetails;
    private _getOwners;
    private onchangedTitle;
    private onchangedFrName;
    private _searchSite;
    protected functionUrl: string;
    private callAzureFunction;
    protected emailQueueUrl: string;
    private SendEmail;
}
//# sourceMappingURL=SCW.d.ts.map