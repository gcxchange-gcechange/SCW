var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from "react";
import { ActionButton } from "office-ui-fabric-react/lib/Button";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import styles from "./Wizard.module.scss";
import { initializeIcons } from '@uifabric/icons';
import * as strings from 'SCWWebPartStrings';
initializeIcons();
var WizardStep = /** @class */ (function (_super) {
    __extends(WizardStep, _super);
    function WizardStep() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return WizardStep;
}(React.Component));
export { WizardStep };
var DEFAULT_NEXT_BUTTON_LABEL = strings.NextBtn;
var DEFAULT_PREVIOUS_BUTTON_LABEL = strings.BackBtn;
var DEFAULT_FINISH_BUTTON_LABEL = "Finish";
var DEFAULT_CANCEL_BUTTON_LABEL = strings.CancelBtn;
var DEFAULT_VALIDATING_MESSAGE = "Validating step...";
var BaseWizard = /** @class */ (function (_super) {
    __extends(BaseWizard, _super);
    function BaseWizard(props) {
        var _this = _super.call(this, props) || this;
        _this._validateWithCallback = function (validationCallback) {
            if (!validationCallback) {
                return;
            }
            var validationResult = _this._validateStep(_this.state.currentStep);
            if (typeof validationResult.then === "function") {
                _this.setState({
                    isValidatingStep: true,
                    errorMessage: null
                });
                var promiseResult = validationResult;
                promiseResult.then(function (result) {
                    validationCallback(result);
                }).catch(function (error) {
                    if (error) {
                        validationCallback({
                            isValidStep: false,
                            errorMessage: error
                        });
                    }
                });
            }
            else {
                var directResult = validationResult;
                if (!directResult) {
                    throw new Error("The validation result has unexpected format.");
                }
                validationCallback(directResult);
            }
        };
        _this._goToStep = function (step, completedSteps, skipValidation) {
            if (skipValidation === void 0) { skipValidation = false; }
            if (!skipValidation) {
                _this._validateWithCallback(function (result) {
                    if (result.isValidStep) {
                        _this.setState({
                            currentStep: step,
                            completedSteps: completedSteps,
                            errorMessage: null,
                            isValidatingStep: false
                        });
                        if (_this.state.currentStep == 8) {
                            //trigger focus on peoplepicker
                            var element = document.getElementsByClassName("ms-BasePicker-input")[1];
                            element.focus();
                        }
                        console.log("Current step: ", _this.state.currentStep, "completeSteps: ", _this.state.completedSteps);
                    }
                    else {
                        _this.setState({
                            errorMessage: result.errorMessage,
                            isValidatingStep: false
                        });
                    }
                });
            }
            else {
                _this.setState({ currentStep: step, completedSteps: completedSteps });
            }
        };
        _this._validateStep = function (step) {
            if (_this.props.onValidateStep) {
                return _this.props.onValidateStep(step);
            }
            return {
                isValidStep: true,
                errorMessage: null
            };
        };
        _this._goToNextStep = function () {
            var completedWizardSteps = (_this.state.completedSteps | _this.state.currentStep);
            var nextStep = (_this.state.currentStep << 1);
            console.log("Current step: ", _this.state.currentStep, " next step: ", nextStep, "completeSteps: ", _this.state.completedSteps);
            _this._goToStep(nextStep, completedWizardSteps);
        };
        _this._goToPreviousStep = function () {
            var previousStep = (_this.state.currentStep >> 1);
            console.log("Current step: ", _this.state.currentStep, " previous step: ", previousStep);
            _this._goToStep(previousStep, null, true);
        };
        _this._cancel = function () {
            if (_this.props.onCancel) {
                _this.props.onCancel();
            }
        };
        _this._finish = function () {
            _this._validateWithCallback(function (result) {
                if (result.isValidStep) {
                    if (_this.props.onCompleted) {
                        _this.props.onCompleted();
                    }
                }
                else {
                    _this.setState({
                        errorMessage: result.errorMessage,
                        isValidatingStep: false
                    });
                }
            });
        };
        _this.state = {
            currentStep: props.defaultCurrentStep || _this.firstStep,
            completedSteps: null,
            errorMessage: null,
            isValidatingStep: false
        };
        return _this;
    }
    BaseWizard.prototype.renderStepProgress = function (type) {
        var _this = this;
        var stepChildren = React.Children.toArray(this.props.children)
            .filter(function (reactChild) { return reactChild.type == WizardStep && reactChild.props.step; });
        if (stepChildren.length == 0) {
            throw new Error("The specified wizard steps are not valid");
        }
        if (type == "content") {
            return stepChildren
                .map(function (reactChild) {
                return React.createElement(PivotItem, { key: "WizardStep__" + reactChild.props.step, itemKey: reactChild.props.step.toString() }, reactChild.props.children);
            });
        }
        else {
            return stepChildren
                .map(function (reactChild) {
                return React.createElement("li", { className: (_this.state.currentStep > reactChild.props.step ? styles.active : "") }, reactChild.props.caption);
            });
        }
    };
    Object.defineProperty(BaseWizard.prototype, "firstStep", {
        get: function () {
            var stepValues = React.Children.toArray(this.props.children)
                .filter(function (c) { return c.props.step > 0; })
                .map(function (c) { return c.props.step; });
            if (stepValues.length < 1) {
                throw new Error("The specified step values are invalid. First step value must be higher than 0");
            }
            return Math.min.apply(Math, stepValues);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "lastStep", {
        get: function () {
            var stepValues = React.Children.toArray(this.props.children)
                .filter(function (c) { return c.props.step > 0; })
                .map(function (c) { return c.props.step; });
            if (stepValues.length < 1) {
                throw new Error("The specified step values are invalid. First step value must be higher than 0");
            }
            return Math.max.apply(Math, stepValues);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "hasNextStep", {
        get: function () {
            return this.state.currentStep < this.lastStep;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "hasPreviousStep", {
        get: function () {
            return this.state.currentStep > this.firstStep;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "cancelButton", {
        get: function () {
            return React.createElement(ActionButton, { title: strings.tooltipBtnCancel, iconProps: { iconName: "Cancel" }, text: this.props.cancelButtonLabel || DEFAULT_CANCEL_BUTTON_LABEL, onClick: this._cancel });
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "previousButton", {
        get: function () {
            if (this.hasPreviousStep) {
                return React.createElement(ActionButton, { title: strings.tooltipBtnBack, className: styles.nextBtn, iconProps: { iconName: "ChevronLeft" }, styles: { icon: { color: 'white', fontSize: 16 }, iconHovered: { color: "white" } }, text: this.props.previousButtonLabel || DEFAULT_PREVIOUS_BUTTON_LABEL, onClick: this._goToPreviousStep });
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "nextButton", {
        get: function () {
            var isDisable = "disableStep" + this.state.currentStep;
            if (this.hasNextStep) {
                return React.createElement(ActionButton, { title: strings.tooltipBtnNext, disabled: this.props[isDisable], className: styles.nextBtn, styles: { flexContainer: { flexDirection: 'row-reverse' }, icon: { color: 'white', fontSize: 16 }, iconHovered: { color: "white" } }, iconProps: { iconName: "ChevronRight" }, text: this.props.nextButtonLabel || DEFAULT_NEXT_BUTTON_LABEL, onClick: this._goToNextStep });
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseWizard.prototype, "finishButton", {
        get: function () {
            if (!this.hasNextStep) {
                return React.createElement(ActionButton, { title: strings.tootltipBtnEnd, autoFocus: true, className: styles.nextBtn, text: this.props.finishButtonLabel || DEFAULT_FINISH_BUTTON_LABEL, onClick: this._finish });
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    BaseWizard.prototype.render = function () {
        var pivotStyles = {
            root: [
                {
                    display: 'flex',
                    justifyContent: 'center',
                    paddingTop: "5%"
                }
            ],
            link: [],
            linkIsSelected: [{
                    selectors: {
                        ':before': {
                            borderBottom: 'none',
                        }
                    }
                }],
            icon: [],
            count: [],
            linkContent: [],
            text: [],
        };
        return React.createElement("div", { className: styles.wizardComponent, style: { backgroundColor: "#e6e6e6" } },
            React.createElement("div", { className: "" + styles.canceled }, this.cancelButton),
            React.createElement("div", { className: styles.container },
                React.createElement("ul", { className: styles.progressbar }, this.renderStepProgress("bar"))),
            React.createElement(Pivot, { styles: pivotStyles, selectedKey: this.state.currentStep.toString() }, this.renderStepProgress("content")),
            this.state.isValidatingStep && React.createElement("div", null, this.props.validatingMessage || DEFAULT_VALIDATING_MESSAGE),
            this.state.errorMessage && React.createElement("div", { className: styles.error }, this.state.errorMessage),
            React.createElement("div", { className: styles.row, style: { backgroundColor: "#e6e6e6" } },
                React.createElement("div", { className: "" + styles.righted },
                    this.nextButton,
                    this.finishButton),
                React.createElement("div", { className: "" + styles.lefted }, this.previousButton)));
    };
    return BaseWizard;
}(React.Component));
export { BaseWizard };
//# sourceMappingURL=Wizard.js.map