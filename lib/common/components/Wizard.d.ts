import * as React from "react";
import { IPivotItemProps } from "office-ui-fabric-react/lib/Pivot";
export interface IWizardStepProps<TStep extends number> extends IPivotItemProps {
    step: TStep;
    caption: string;
}
export declare class WizardStep<TStep extends number> extends React.Component<IWizardStepProps<TStep>, {}> {
}
export interface IWizardStepValidationResult {
    isValidStep: boolean;
    errorMessage?: string;
}
export interface IWizardProps<TStep extends number> {
    defaultCurrentStep?: TStep;
    onValidateStep?: (currentStep: TStep) => IWizardStepValidationResult | Promise<IWizardStepValidationResult>;
    onCompleted?: () => void;
    onCancel?: () => void;
    nextButtonLabel?: string;
    previousButtonLabel?: string;
    cancelButtonLabel?: string;
    finishButtonLabel?: string;
    validatingMessage?: string;
    mainCaption?: string;
    disableStep1?: boolean;
    disableStep2?: boolean;
    disableStep4?: boolean;
    disableStep8?: boolean;
}
export interface IWizardState<TStep extends number> {
    currentStep: TStep;
    completedSteps: TStep;
    errorMessage: string;
    isValidatingStep: boolean;
}
export declare abstract class BaseWizard<TStep extends number> extends React.Component<IWizardProps<TStep>, IWizardState<TStep>> {
    constructor(props: IWizardProps<TStep>);
    private renderStepProgress;
    private readonly firstStep;
    private readonly lastStep;
    private _validateWithCallback;
    private _goToStep;
    private _validateStep;
    private readonly hasNextStep;
    private readonly hasPreviousStep;
    private _goToNextStep;
    private _goToPreviousStep;
    private _cancel;
    private _finish;
    private readonly cancelButton;
    private readonly previousButton;
    private readonly nextButton;
    private readonly finishButton;
    render(): React.ReactElement<IWizardProps<TStep>>;
}
//# sourceMappingURL=Wizard.d.ts.map