import * as React from "react";
import { ActionButton } from "office-ui-fabric-react/lib/Button";
import { Pivot, PivotItem, IPivotItemProps, IPivotStyles } from "office-ui-fabric-react/lib/Pivot";
import styles from "./Wizard.module.scss";
import * as strings from 'SCWWebPartStrings';

export interface IWizardStepProps<TStep extends number> extends IPivotItemProps {
    step: TStep;
    caption: string;
}

export class WizardStep<TStep extends number> extends React.Component<IWizardStepProps<TStep>, {}> {

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

const DEFAULT_NEXT_BUTTON_LABEL = strings.NextBtn;
const DEFAULT_PREVIOUS_BUTTON_LABEL = strings.BackBtn;
const DEFAULT_FINISH_BUTTON_LABEL = "Finish";
const DEFAULT_CANCEL_BUTTON_LABEL = strings.CancelBtn;
const DEFAULT_VALIDATING_MESSAGE = "Validating step...";

export abstract class BaseWizard<TStep extends number> extends React.Component<IWizardProps<TStep>, IWizardState<TStep>> {

    constructor(props: IWizardProps<TStep>) {
        super(props);

        this.state = {
            currentStep: props.defaultCurrentStep || this.firstStep,
            completedSteps: null,
            errorMessage: null,
            isValidatingStep: false
        };
    }

    private renderStepProgress(type) {
        const stepChildren = React.Children.toArray(this.props.children)
        .filter((reactChild: React.ReactElement) => reactChild.type == WizardStep && reactChild.props.step);

        if (stepChildren.length == 0) {
            throw new Error("The specified wizard steps are not valid");
        }
        if(type == "content"){
            return stepChildren
            .map((reactChild: React.ReactElement) => {
                return <PivotItem key={`WizardStep__${reactChild.props.step}`}
                    itemKey={reactChild.props.step.toString()} >
                    {reactChild.props.children}
                </PivotItem>;
            });
        }else{
            return stepChildren
            .map((reactChild: React.ReactElement) => {
               return<li className={(this.state.currentStep > reactChild.props.step ? styles.active : "")}>
                   {reactChild.props.caption}
                </li>;
            });
        }
    }

    private get firstStep(): TStep {
        const stepValues = React.Children.toArray(this.props.children)
            .filter((c: React.ReactElement) => c.props.step as number > 0)
            .map((c: React.ReactElement) => c.props.step as number);
        if (stepValues.length < 1) {
            throw new Error("The specified step values are invalid. First step value must be higher than 0");
        }
        return Math.min(...stepValues) as TStep;
    }

    private get lastStep(): TStep {
        const stepValues = React.Children.toArray(this.props.children)
            .filter((c: React.ReactElement) => c.props.step as number > 0)
            .map((c: React.ReactElement) => c.props.step as number);
        if (stepValues.length < 1) {
            throw new Error("The specified step values are invalid. First step value must be higher than 0");
        }
        return Math.max(...stepValues) as TStep;
    }

    private _validateWithCallback = (validationCallback: (validationResult: IWizardStepValidationResult) => void) => {

        if (!validationCallback) {
            return;
        }

        const validationResult = this._validateStep(this.state.currentStep);
        if (typeof (validationResult as Promise<IWizardStepValidationResult>).then === "function") {
            this.setState({
                isValidatingStep: true,
                errorMessage: null
            });
            const promiseResult = validationResult as Promise<IWizardStepValidationResult>;
            promiseResult.then(result => {
                validationCallback(result);
            }).catch(error => {
                if (error as string) {
                    validationCallback({
                        isValidStep: false,
                        errorMessage: error
                    });
                }
            });
        }
        else {
            const directResult = validationResult as IWizardStepValidationResult;
            if (!directResult) {
                throw new Error("The validation result has unexpected format.");
            }
            validationCallback(directResult);
        }
    }

    private _goToStep = (step: TStep, completedSteps?: TStep, skipValidation: boolean = false) => {

        if (!skipValidation) {
            this._validateWithCallback(result => {
                if (result.isValidStep) {

                    this.setState({
                        currentStep: step,
                        completedSteps,
                        errorMessage: null,
                        isValidatingStep: false
                    });

                    if(this.state.currentStep == 8){
                        //trigger focus on peoplepicker
                        let element: HTMLElement = document.getElementsByClassName("ms-BasePicker-input")[1] as HTMLElement;
                        element.focus();
                    }
                    console.log("Current step: ", this.state.currentStep,  "completeSteps: ", this.state.completedSteps);
                } else {
                    this.setState({
                        errorMessage: result.errorMessage,
                        isValidatingStep: false
                    });
                }
            });
        } else {
            this.setState({ currentStep: step, completedSteps });
        }
    }

    private _validateStep = (step: TStep) => {
        if (this.props.onValidateStep) {
            return this.props.onValidateStep(step);
        }

        return {
            isValidStep: true,
            errorMessage: null
        };
    }

    private get hasNextStep(): boolean {
        return this.state.currentStep < this.lastStep;
    }

    private get hasPreviousStep(): boolean {
        return this.state.currentStep > this.firstStep;
    }

    private _goToNextStep = () => {
        let completedWizardSteps = (this.state.completedSteps | this.state.currentStep) as TStep;
        const nextStep = ((this.state.currentStep as number) << 1) as TStep;
        console.log("Current step: ", this.state.currentStep, " next step: ", nextStep, "completeSteps: ", this.state.completedSteps);
        this._goToStep(nextStep, completedWizardSteps);
    }

    private _goToPreviousStep = () => {
        const previousStep = ((this.state.currentStep as number) >> 1) as TStep;
        console.log("Current step: ", this.state.currentStep, " previous step: ", previousStep);
        this._goToStep(previousStep, null, true);
    }


    private _cancel = () => {
        if (this.props.onCancel) {
            this.props.onCancel();
        }
    }

    private _finish = () => {
        this._validateWithCallback((result) => {
            if (result.isValidStep) {
                if (this.props.onCompleted) {
                    this.props.onCompleted();
                }
            } else {
                this.setState({
                    errorMessage: result.errorMessage,
                    isValidatingStep: false
                });
            }
        });
    }

    private get cancelButton(): JSX.Element {
        return <ActionButton title={strings.tooltipBtnCancel} iconProps={{ iconName: "Cancel" }} text={this.props.cancelButtonLabel || DEFAULT_CANCEL_BUTTON_LABEL} onClick={this._cancel} />;
    }

    private get previousButton(): JSX.Element {
        if (this.hasPreviousStep) {
            return <ActionButton title={strings.tooltipBtnBack} className={styles.nextBtn} iconProps={{ iconName: "ChevronLeft" }} styles={{icon: {color: 'white', fontSize: 16}, iconHovered: {color:"white"}}} text={this.props.previousButtonLabel || DEFAULT_PREVIOUS_BUTTON_LABEL} onClick={this._goToPreviousStep} />;
        }
        return null;
    }

    private get nextButton(): JSX.Element {
        var isDisable = "disableStep"+this.state.currentStep;
        if (this.hasNextStep) {
            return <ActionButton title={strings.tooltipBtnNext} disabled={this.props[isDisable]} className={styles.nextBtn} styles={{flexContainer: { flexDirection: 'row-reverse' }, icon: {color: 'white', fontSize: 16}, iconHovered: {color:"white"}}} iconProps={{ iconName: "ChevronRight" }} text={this.props.nextButtonLabel || DEFAULT_NEXT_BUTTON_LABEL} onClick={this._goToNextStep} />;
        }
        return null;
    }

    private get finishButton(): JSX.Element {
        if (!this.hasNextStep) {
            return <ActionButton title={strings.tootltipBtnEnd} autoFocus className={styles.nextBtn} text={this.props.finishButtonLabel || DEFAULT_FINISH_BUTTON_LABEL} onClick={this._finish} />;
        }

        return null;
    }

    public render(): React.ReactElement<IWizardProps<TStep>> {
        const pivotStyles: IPivotStyles = {
            root: [
              {
                display: 'flex',
                justifyContent: 'center',
                paddingTop:"5%"
              }
            ],
            link:[{
                display: 'none',
            }],
            linkIsSelected: [{
                display: 'none',
                selectors: {
                    ':before': {
                        borderBottom: 'none',
                    }
                  }
            }],
            icon: [],
            count:[],
            linkContent: [{
                'display': 'none',
            }],
            text:[],
          };
        return <div className={styles.wizardComponent}>
            <div className={`${styles.canceled}`}>
                    {this.cancelButton}
                </div>
            <div className={styles.container}>
               <ul className={styles.progressbar}>  
               {this.renderStepProgress("bar")}
               </ul>
               </div>
            <Pivot styles={pivotStyles} selectedKey={this.state.currentStep.toString()}>
                {this.renderStepProgress("content")}
            </Pivot>
            {this.state.isValidatingStep && <div>{this.props.validatingMessage || DEFAULT_VALIDATING_MESSAGE}</div>}
            {this.state.errorMessage && <div className={styles.error}>{this.state.errorMessage}</div>}

            <div className={styles.row}>
                <div className={`${styles.righted}`}>
                    {this.nextButton}
                    {this.finishButton}
                </div>
                <div className={`${styles.lefted}`}>
                    {this.previousButton}
                </div>
            </div>
        </div>;
    }
}