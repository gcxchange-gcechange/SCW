import { MessageBarType, IDropdownOption } from 'office-ui-fabric-react';
import { ISiteItem } from './ISiteItem';
import { ITemplate } from './ITemplate';
import { ISelected } from './ISelected';
export interface ISCWState {
    title: string;
    showMessageBar: boolean;
    messageType?: MessageBarType;
    message?: string;
    frName?: string;
    items: IDropdownOption[];
    enDes?: string;
    frDes?: string;
    reason?: string;
    sites: Array<ISiteItem>;
    isAvailiability: string;
    error: string;
    isSiteEnNameRight: boolean;
    isSiteFrNameRight: boolean;
    ownersNumber: Number;
    currentPage: number;
    templateItems: ITemplate[];
    selectionDetails: string;
    selectedTempalteTitle: string;
    isCurrentPage: boolean;
    isWizardOpened: boolean;
    statusMessage: string;
    statusType: "OK" | "KO" | null;
    firstStepInput: string;
    thirdStepInput: string;
    wizardValidatingMessage: string;
    tellusEn: string;
    tellusFr: string;
    BusinessReason: string;
    selected: ISelected[];
    checkSite: boolean;
    loading: boolean;
}
//# sourceMappingURL=ISCWState.d.ts.map