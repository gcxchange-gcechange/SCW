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
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './SCW.module.scss';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { MessageBarType, Label, Spinner, Image, DefaultButton, ImageFit } from 'office-ui-fabric-react';
import { Selection } from 'office-ui-fabric-react/lib/DetailsList';
import { autobind } from 'office-ui-fabric-react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { AadHttpClient } from "@microsoft/sp-http";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { BaseWizard, WizardStep } from "../../../common/components/Wizard";
import * as strings from 'SCWWebPartStrings';
import { spaceDescFr } from 'SCWWebPartStrings';
var owners = [];
var totalPages = 1;
var allTemplateItems = [];
var selTemplate = [];
var currentSelectedKey = -1;
export var MyWizardSteps;
(function (MyWizardSteps) {
    MyWizardSteps[MyWizardSteps["None"] = 0] = "None";
    MyWizardSteps[MyWizardSteps["FirstStep"] = 1] = "FirstStep";
    MyWizardSteps[MyWizardSteps["SecondStep"] = 2] = "SecondStep";
    MyWizardSteps[MyWizardSteps["ThirdStep"] = 4] = "ThirdStep";
    MyWizardSteps[MyWizardSteps["FourthStep"] = 8] = "FourthStep";
    MyWizardSteps[MyWizardSteps["LastStep"] = 16] = "LastStep";
})(MyWizardSteps || (MyWizardSteps = {}));
var MyWizard = /** @class */ (function (_super) {
    __extends(MyWizard, _super);
    function MyWizard() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return MyWizard;
}(BaseWizard));
export { MyWizard };
var SCW = /** @class */ (function (_super) {
    __extends(SCW, _super);
    function SCW(props, state) {
        var _this = _super.call(this, props) || this;
        _this._selection = new Selection({
            onSelectionChanged: function () {
                if (_this._selection.count != 0) {
                    currentSelectedKey = _this._selection.getSelection()[0].key;
                }
                _this.setState({ selectionDetails: _this._getSelectionDetails() });
            }
        });
        _this.selection1 = new Selection;
        _this.functionTemplateImg = "https://gettemplate.azurewebsites.net/api/HttpTriggerCSharp1";
        _this._searchSite = function () {
            // Log the current operation
            _this.props.context.msGraphClientFactory
                .getClient()
                .then(function (client) {
                // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
                client
                    .api("groups")
                    .filter("displayName eq '" + _this.state.title + "-" + _this.state.frName + "'")
                    .select("displayName,id,name")
                    .get(function (err, res) {
                    if (err) {
                        console.error(err);
                        return;
                    }
                    // Prepare the output array
                    var sites = new Array();
                    // Map the JSON response to the output array
                    res.value.map(function (item) {
                        sites.push({
                            displayName: item.displayName,
                            id: item.id,
                        });
                    });
                    if (sites.length != 0) {
                        _this.setState({
                            sites: sites,
                            isAvailiability: strings.siteTaken,
                            checkSite: true
                        });
                    }
                    else {
                        _this.setState({
                            sites: sites,
                            isAvailiability: strings.greatChoice,
                            checkSite: false
                        });
                    }
                });
            });
        };
        _this.functionUrl = "https://scwsitecreation.azurewebsites.net/api/HttpTrigger1?";
        _this.emailQueueUrl = "https://createsitefunchttp20210105171931.azurewebsites.net/api/SendStatusToQueue?";
        _this.state = {
            title: '',
            showMessageBar: false,
            frName: '',
            items: [],
            enDes: '',
            sites: [],
            isAvailiability: '',
            error: '',
            isSiteEnNameRight: true,
            isSiteFrNameRight: true,
            ownersNumber: 1,
            currentPage: 1,
            templateItems: [],
            selectionDetails: _this._getSelectionDetails(),
            selectedTempalteTitle: '',
            isCurrentPage: true,
            isWizardOpened: false,
            statusMessage: null,
            statusType: null,
            firstStepInput: null,
            thirdStepInput: null,
            tellusEn: "",
            tellusFr: "",
            BusinessReason: "",
            wizardValidatingMessage: 'Validating...',
            selected: [],
            checkSite: true,
            loading: false
        };
        return _this;
    }
    SCW.prototype.onInit = function () {
        var _this = this;
        var retVal = Promise.resolve();
        if (this.context.microsoftTeams) {
            retVal = new Promise(function (resolve, reject) {
                _this.context.microsoftTeams.getContext(function (context) {
                    _this._teamsContext = context;
                    resolve();
                });
            });
        }
        return retVal;
    };
    SCW.prototype.imagesTemplate = function (key, title) {
        var templateSel = {
            key: key,
            title: title,
        };
        selTemplate = [];
        selTemplate.push(templateSel);
        this.setState({
            selected: selTemplate,
        });
    };
    SCW.prototype._closeWizard = function (completed) {
        var _this = this;
        if (completed === void 0) { completed = false; }
        this.setState({
            isWizardOpened: false,
            statusMessage: completed ? "The wizard has been completed" : "The wizard has been canceled",
            statusType: completed ? "OK" : "KO"
        });
        setTimeout(function () {
            _this.setState({
                statusMessage: null,
                statusType: null
            });
        }, 3000);
    };
    SCW.prototype._onValidateStep = function (step) {
        var _this = this;
        var isValid = true;
        var isValid1 = true;
        var isValid2 = true;
        var ValidResult = true;
        switch (step) {
            case MyWizardSteps.FirstStep:
                isValid = this.state.selected[0] !== undefined;
                return {
                    isValidStep: isValid,
                    errorMessage: !isValid ? "Select a template" : null
                };
            case MyWizardSteps.ThirdStep:
                return new Promise(function (resolve) {
                    isValid = _this.state.tellusEn.length >= 5 && _this.state.tellusEn.length <= 500;
                    isValid1 = _this.state.tellusFr.length >= 5 && _this.state.tellusFr.length <= 500;
                    isValid2 = _this.state.BusinessReason.length >= 5 && _this.state.BusinessReason.length <= 500;
                    if (isValid == true && isValid2 == true && isValid1 == true) {
                        ValidResult = true;
                    }
                    else {
                        ValidResult = false;
                    }
                    setTimeout(function () {
                        resolve({
                            isValidStep: ValidResult,
                            errorMessage: !ValidResult ? "Your input to third step is invalid" : null
                        });
                    });
                });
            default:
                return { isValidStep: true };
        }
    };
    SCW.prototype._renderMyWizard = function () {
        var _this = this;
        var listOwners = "";
        for (var step = 0; step < owners.length; step++) {
            if (listOwners == "") {
                listOwners = owners[step];
            }
            else {
                listOwners = listOwners + ', ' + owners[step];
            }
        }
        return React.createElement(MyWizard, { mainCaption: "", onCancel: function () { return _this._closeWizard(false); }, onCompleted: function () { return _this.callAzureFunction(); }, onValidateStep: function (step) { return _this._onValidateStep(step); }, validatingMessage: this.state.wizardValidatingMessage, disableStep1: (this.state.selected[0] !== undefined ? false : true), disableStep2: this.state.checkSite, disableStep4: (this.state.tellusEn.length >= 5 && this.state.tellusEn.length < 500 && this.state.tellusFr.length >= 5 && this.state.tellusFr.length < 500 && this.state.BusinessReason.length >= 5 && this.state.BusinessReason.length < 500 ? false : true), disableStep8: (this.state.ownersNumber >= 2 ? false : true), finishButtonLabel: strings.btnSubmit },
            React.createElement(WizardStep, { caption: strings.menuTemplate, step: MyWizardSteps.FirstStep },
                React.createElement("div", { className: styles.wizardStep },
                    React.createElement("h1", { className: styles.titleStep }, strings.titleTemplate),
                    React.createElement("p", null, strings.paragrapheTemplate),
                    React.createElement("div", { className: "ms-Grid", dir: "ltr" }, this.state.templateItems.map(function (item) { return (React.createElement("button", { autoFocus: (item.key == 0 ? true : false), className: styles.imagetest + " " + (_this.state.selected[0] !== undefined ? _this.state.selected[0]["key"] == item.key ? styles.selected : "" : "") + " ms-Grid-col ms-sm12 ms-md6 ms-lg6 ", onClick: function () { return _this.imagesTemplate(item.key, item.title); }, "aria-label": "" + strings.templateButtonLabel + (strings.userLang == "EN" ? item.title : item.titleFR) },
                        React.createElement(Image, { title: "" + strings.altTemplate + (strings.userLang == "EN" ? item.title : item.titleFR), src: item.url, alt: "" + strings.altTemplate + (strings.userLang == "EN" ? item.title : item.titleFR), width: 150, height: 250, className: "ms-Grid-col ms-sm12 ms-md6 ms-lg6" }),
                        React.createElement("div", { className: "ms-Grid-col ms-sm5 ms-md5 ms-lg5" },
                            React.createElement("h4", null, (strings.userLang == "EN" ? item.title : item.titleFR)),
                            React.createElement("p", { title: (strings.userLang == "EN" ? item.description : item.descriptionFR), className: styles.templateDesc }, (strings.userLang == "EN" ? item.description : item.descriptionFR))))); })))),
            React.createElement(WizardStep, { caption: strings.menuSpace, step: MyWizardSteps.SecondStep },
                React.createElement("div", { className: styles.wizardStep },
                    React.createElement("h1", { className: styles.titleStep }, strings.titleSpace),
                    React.createElement("p", null, strings.paragrapheSpace),
                    React.createElement("em", null, strings.validationTxtSpace),
                    React.createElement("section", { className: styles.SectiontextField },
                        React.createElement("div", { className: "form-group" },
                            React.createElement(Label, { htmlFor: "englishLabelTitle", className: styles.labelBulingue, required: true }, strings.english),
                            React.createElement(TextField, { title: strings.tooltipspaceNameEn, autoFocus: true, id: "englishLabelTitle", onChanged: this.onchangedTitle }),
                            React.createElement("span", { style: { color: "#C70000" } }, this.state.error),
                            React.createElement("br", null)),
                        React.createElement("div", { className: "form-group" },
                            React.createElement(Label, { htmlFor: "frenchLabelTitle", required: true, className: styles.labelBulingue }, strings.french),
                            React.createElement(TextField, { title: strings.tooltipspaceNameFr, id: "frenchLabelTitle", onChanged: this.onchangedFrName }),
                            React.createElement("span", { style: { color: "#C70000" } }, this.state.error),
                            React.createElement("div", { className: styles.yes + " form-group" },
                                React.createElement("p", null,
                                    React.createElement("label", { className: (this.state.checkSite == false ? styles.greencheckSite : styles.redcheckSite) },
                                        " ",
                                        this.state.isAvailiability)))),
                        React.createElement("br", null),
                        React.createElement(DefaultButton, { title: strings.tooltipchecksite, className: styles.checkSiteBtn, disabled: (this.state.title.length >= 5 && this.state.title.length <= 125 && this.state.frName.length >= 5 && this.state.frName.length <= 125 ? false : true), onClick: this._searchSite }, strings.btnChecksite)))),
            React.createElement(WizardStep, { caption: strings.menuTell, step: MyWizardSteps.ThirdStep },
                React.createElement("div", { className: styles.wizardStep },
                    React.createElement("h1", null, strings.titleTellUs),
                    React.createElement("p", null, strings.paragrapheTellUs),
                    React.createElement("em", null, strings.validationTxtTellUs),
                    React.createElement("section", { className: styles.SectiontextField },
                        React.createElement(Label, { htmlFor: "englishLabelDesc", className: styles.labelBulingue, required: true }, strings.english),
                        React.createElement(TextField, { title: strings.tooltipdescEn, autoFocus: true, multiline: true, rows: 4, value: this.state.tellusEn, placeholder: strings.phLetus, id: "englishLabelDesc", onChanged: function (v) { return _this.setState({ tellusEn: v }); } }),
                        React.createElement(Label, { htmlFor: "frenchLabelDesc", className: styles.labelBulingue, required: true }, strings.french),
                        React.createElement(TextField, { title: strings.tooltipdescFr, multiline: true, rows: 4, value: this.state.tellusFr, id: "frenchLabelDesc", placeholder: strings.phLetus, onChanged: function (v) { return _this.setState({ tellusFr: v }); } }),
                        React.createElement(Label, { htmlFor: "businessLabel", className: styles.labelBulingue, required: true }, strings.businessReason),
                        React.createElement(TextField, { title: strings.tooltipBusReason, multiline: true, rows: 4, id: "businessLabel", value: this.state.BusinessReason, placeholder: strings.phBusinessReason, onChanged: function (v) { return _this.setState({ BusinessReason: v }); } })))),
            React.createElement(WizardStep, { caption: strings.menuOwners, step: MyWizardSteps.FourthStep },
                React.createElement("div", { className: styles.wizardStep },
                    React.createElement("h1", { className: styles.titleStep }, strings.titleOwners),
                    React.createElement("p", null, strings.paragrapheOwners),
                    React.createElement("p", null, strings.validationTxtOwners),
                    React.createElement("div", { className: "form-group" },
                        React.createElement(Label, { htmlFor: "peopleLabel", className: styles.labelBulingue, required: true }, strings.owners),
                        React.createElement(PeoplePicker, { context: this.props.context, personSelectionLimit: 1, groupName: "", showHiddenInUI: false, defaultSelectedUsers: [this.props.context.pageContext.user.email], required: true, ensureUser: false, disabled: true }),
                        React.createElement(PeoplePicker, { showtooltip: true, tooltipMessage: strings.tooltipOwners, context: this.props.context, personSelectionLimit: 2, groupName: "", showHiddenInUI: false, required: true, onChange: this._getOwners, ensureUser: false })),
                    React.createElement("p", null, strings.ownerInfo1),
                    React.createElement("p", null, strings.ownerInfo2),
                    React.createElement("p", null, strings.ownerInfo3),
                    React.createElement("p", null, strings.ownerInfo4))),
            React.createElement(WizardStep, { caption: strings.menuFinal, step: MyWizardSteps.LastStep },
                React.createElement("div", { className: styles.wizardStep },
                    React.createElement("h1", { className: styles.titleStep }, strings.titleReview),
                    (this.state.loading ?
                        React.createElement("div", null,
                            React.createElement(Label, null, strings.textLoading),
                            React.createElement(Spinner, { label: strings.iconLoading, ariaLive: "assertive", labelPosition: "left" }))
                        :
                            React.createElement("div", null,
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md6 ms-lg6" },
                                    React.createElement(Label, { htmlFor: "templateLabel", className: styles.labelBulingue, required: true }, strings.templateTitle),
                                    React.createElement(TextField, { title: strings.templateTitle, id: "templateLabel", readOnly: true, value: (this.state.selected[0] !== undefined ? this.state.selected[0]["title"] : ""), placeholder: "template" }),
                                    React.createElement(Label, { htmlFor: "spaceEnLabel", className: styles.labelBulingue, required: true }, strings.spaceDescEn),
                                    React.createElement(TextField, { title: strings.spaceDescEn, id: "spaceEnLabel", readOnly: true, defaultValue: this.state.tellusEn, placeholder: "Descripton en" }),
                                    React.createElement(Label, { htmlFor: "ownersLabel", className: styles.labelBulingue, required: true }, strings.owners),
                                    React.createElement(TextField, { title: strings.owners, id: "ownersLabel", multiline: true, autoAdjustHeight: true, readOnly: true, value: listOwners, placeholder: "Owners" })),
                                React.createElement("div", { className: "ms-Grid-col ms-sm12 ms-md6 ms-lg6" },
                                    React.createElement(Label, { htmlFor: "spaceNameLabel", className: styles.labelBulingue, required: true }, strings.spaceName),
                                    React.createElement(TextField, { title: strings.spaceName, id: "spaceNameLabel", readOnly: true, value: "'" + this.state.title + "-" + this.state.frName + "'", placeholder: "Space Name" }),
                                    React.createElement(Label, { htmlFor: "spaceFrLabel", className: styles.labelBulingue, required: true }, strings.spaceDescFr),
                                    React.createElement(TextField, { title: spaceDescFr, id: "spaceFrLabel", readOnly: true, defaultValue: this.state.tellusFr, placeholder: "Description fr" }),
                                    React.createElement(Label, { htmlFor: "teamPurposeLabel", className: styles.labelBulingue, required: true }, strings.teamPurpose),
                                    React.createElement(TextField, { title: strings.teamPurpose, id: "teamPurposeLabel", multiline: true, autoAdjustHeight: true, readOnly: true, value: this.state.BusinessReason, placeholder: "Business Reason" })))))),
            React.createElement("div", null, "Invalid element here, will be ignored"));
    };
    SCW.prototype._openWizard = function () {
        this.setState({
            isWizardOpened: true
        });
    };
    SCW.prototype.ResetScreen = function () {
        this.setState({ title: '',
            showMessageBar: false,
            frName: '',
            items: [],
            enDes: '',
            sites: [],
            isAvailiability: '',
            error: '',
            isSiteEnNameRight: true,
            isSiteFrNameRight: true,
            ownersNumber: 1,
            currentPage: 1,
            //templateItems: [],
            selectionDetails: this._getSelectionDetails(),
            selectedTempalteTitle: '',
            isCurrentPage: true,
            isWizardOpened: false,
            statusMessage: null,
            statusType: null,
            firstStepInput: null,
            thirdStepInput: null,
            tellusEn: "",
            tellusFr: "",
            BusinessReason: "",
            wizardValidatingMessage: 'Validating...',
            selected: [],
            checkSite: true,
            loading: false
        });
    };
    SCW.prototype.render = function () {
        var _this = this;
        var imageWelcome = {
            src: require("../../../../assets/sharepoint_teams_graphic.png"),
            imageFit: ImageFit.contain,
            width: 300,
            height: 150,
        };
        var imageCongrat = {
            src: require("../../../../assets/gcxchange_support_pencil.png"),
            imageFit: ImageFit.contain,
            width: 300,
            height: 150,
        };
        return (React.createElement("div", { className: styles.container },
            React.createElement("div", { className: styles.row },
                React.createElement("div", null, this.state.isWizardOpened
                    ? this._renderMyWizard()
                    : this.state.showMessageBar
                        ?
                            React.createElement("div", { className: styles.congratScreen },
                                React.createElement(Image, __assign({}, imageCongrat, { alt: strings.altCongrat, className: styles.imageFit })),
                                React.createElement("h1", null, strings.congrats),
                                React.createElement("p", null, strings.congratPara1),
                                React.createElement("p", null, strings.congratPara2),
                                React.createElement("button", { autoFocus: true, onClick: function () { return _this.ResetScreen(); } }, strings.congratHome),
                                React.createElement("p", null,
                                    strings.congratPara3,
                                    " ",
                                    React.createElement("a", { href: "https://tbssctdev.sharepoint.com/teams/scw/Lists/sr/AllItems.aspx" },
                                        " ",
                                        strings.congratLink)))
                        :
                            React.createElement("div", { className: styles.welcomeContainer },
                                React.createElement(Image, __assign({}, imageWelcome, { alt: strings.altWelcome, className: styles.imageFit, title: strings.tooltipWelImg })),
                                React.createElement("h1", null, strings.createSpace),
                                React.createElement("p", null, strings.paragrapheHome),
                                this.state.templateItems.length != 0 ?
                                    React.createElement(DefaultButton, { title: strings.startButton, className: styles.GoButton, text: strings.startButton, onClick: function () { return _this._openWizard(); } })
                                    :
                                        React.createElement("div", null,
                                            React.createElement(Spinner, { label: strings.iconLoading, ariaLive: "assertive" })),
                                React.createElement("h6", null,
                                    strings.powered,
                                    React.createElement("br", null),
                                    strings.gcx))))));
    };
    SCW.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadTemplate()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    SCW.prototype.loadTemplate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log("before");
                        console.log(allTemplateItems);
                        return [4 /*yield*/, this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then(function (client) {
                                client.get(_this.functionTemplateImg, AadHttpClient.configurations.v1).then(function (response) {
                                    console.log("Status code: " + response.status);
                                    response.json().then(function (responseJSON) {
                                        var i = 0;
                                        //allTemplateItems = [];
                                        for (var k in responseJSON) {
                                            var template = {
                                                key: i,
                                                title: responseJSON[k].TitleEn,
                                                titleFR: responseJSON[k].TitleFr,
                                                description: responseJSON[k].DescriptionEn,
                                                descriptionFR: responseJSON[k].DescriptionFr,
                                                url: responseJSON[k].TemplateImgUrl
                                            };
                                            allTemplateItems.push(template);
                                            i++;
                                        }
                                        totalPages = Math.ceil(allTemplateItems.length / 4);
                                        if (response.ok) {
                                            console.log("response OK");
                                            _this.setState({
                                                templateItems: allTemplateItems,
                                            });
                                        }
                                        else {
                                            console.log("Response error");
                                        }
                                    })
                                        .catch(function (response) {
                                        var errMsg = "WARNING - error when calling URL " + _this.functionUrl + ". Error = " + response.message + response.status + JSON.stringify(response);
                                        console.log("err is ", errMsg);
                                    });
                                });
                            })];
                    case 1:
                        _a.sent();
                        console.log("from call");
                        console.log(allTemplateItems);
                        return [2 /*return*/];
                }
            });
        });
    };
    SCW.prototype._getSelectionDetails = function () {
        var selectionCount = this._selection.count;
        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return (this._selection.getSelection()[0].title);
            default:
                return selectionCount + " items selected";
        }
    };
    SCW.prototype._getOwners = function (ownersFromPeoplePicker) {
        owners = [this.props.context.pageContext.user.email];
        for (var item in ownersFromPeoplePicker) {
            owners.push(ownersFromPeoplePicker[item].secondaryText);
            console.log("owner is", owners);
        }
        this.setState({ ownersNumber: owners.length });
    };
    SCW.prototype.onchangedTitle = function (title) {
        // check length, only include letter、number and -   title.length < 5 || title.length > 10 ||
        if (title.match("^([a-zA-Z0-9 ]*)+$") == null || title.length < 5 || title.length > 125) {
            this.setState({ isSiteEnNameRight: true });
        }
        else {
            this.setState({ error: "" });
            this.setState({ isSiteEnNameRight: false });
        }
        this.setState({
            title: title,
            isAvailiability: "",
            error: strings.ErrMustLetter
        });
    };
    SCW.prototype.onchangedFrName = function (frName) {
        if (frName.match("^([A-Za-z0-9àâäèéêëîïôœùûüÿçÀÂÄÈÉÊËÎÏÔŒÙÛÜŸÇ ]*)+$") == null || frName.length < 5 || frName.length > 125) {
            this.setState({ error: strings.ErrMustLetter });
            this.setState({ isSiteFrNameRight: true });
        }
        else {
            this.setState({ error: "" });
            this.setState({ isSiteFrNameRight: false });
        }
        this.setState({
            frName: frName,
            isAvailiability: "",
            error: strings.ErrMustLetter
        });
    };
    // code=C6K3k07tDzwSaI/PZhAr/rJrMFY1pSaHTpDIc7c0sVn3q75cFVtCJg==";    
    SCW.prototype.callAzureFunction = function () {
        var _this = this;
        var requestHeaders = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
        var siteUrl = this.props.context.pageContext.web.absoluteUrl;
        var owner1, owner2, owner3;
        if (owners.length == 2) {
            owner1 = owners[0];
            owner2 = owners[1];
            owner3 = "";
        }
        else {
            owner1 = owners[0];
            owner2 = owners[1];
            owner3 = owners[2];
        }
        console.log("owner3 is ", owner3);
        var postOptions = {
            headers: requestHeaders,
            body: "\n        {\n          \"name\": \n          {\n            \"title\": \"" + this.state.title + "-" + this.state.frName + "\",\n            \"spacenamefr\": \"" + this.state.frName + "\",\n            \"owner1\": \"" + owner1 + "\",\n            \"owner2\": \"" + owner2 + "\",\n            \"owner3\": \"" + owner3 + "\",\n            \"description\": \"" + this.state.tellusEn + "\",\n            \"descriptionFr\": \"" + this.state.tellusFr + "\",\n            \"business\":\"" + this.state.BusinessReason + "\",\n            \"template\": \"" + this.state.selected[0]["title"] + "\",\n            \"requester_name\": \"" + this.props.context.pageContext.user.displayName + "\",\n            \"requester_email\": \"" + this.props.context.pageContext.user.email + "\",\n          }\n        }"
        };
        var responseText = "";
        // use aad authentication
        this.setState({ loading: true }, function () {
            _this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then(function (client) {
                client.post(_this.functionUrl, AadHttpClient.configurations.v1, postOptions).then(function (response) {
                    console.log("Status code: " + response.status);
                    _this.setState({
                        showMessageBar: true,
                        messageType: MessageBarType.success,
                        isWizardOpened: false,
                        loading: false
                    });
                    _this.SendEmail();
                    response.json().then(function (responseJSON) {
                        responseText = JSON.stringify(responseJSON);
                        console.log("respond is ", responseText);
                        if (response.ok) {
                            console.log("response OK");
                        }
                        else {
                            console.log("Response error");
                        }
                    })
                        .catch(function (response) {
                        var errMsg = "WARNING - error when calling URL " + _this.functionUrl + ". Error = " + response.message;
                        console.log("err is ", errMsg);
                    });
                });
            });
        });
    };
    SCW.prototype.SendEmail = function () {
        var _this = this;
        var requestHeaders = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
        var postQueue = {
            headers: requestHeaders,
            body: "\n      {\n          \"name\": \"" + this.state.title + "-" + this.state.frName + "\",\n          \"status\": \"Submitted\",\n          \"requesterName\": \"" + this.props.context.pageContext.user.displayName + "\",\n          \"requesterEmail\": \"" + this.props.context.pageContext.user.email + "\"\n      }"
        };
        this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then(function (client) {
            client.post(_this.emailQueueUrl, AadHttpClient.configurations.v1, postQueue).then(function (response) {
                console.log("Status code:", response.status);
                console.log('respond is ', response.ok);
                console.log('send reject message to queue successful.');
                console.log("requester Email", _this.props.context.pageContext.user.email);
            });
        });
    };
    __decorate([
        autobind
    ], SCW.prototype, "_getOwners", null);
    __decorate([
        autobind
    ], SCW.prototype, "onchangedTitle", null);
    __decorate([
        autobind
    ], SCW.prototype, "onchangedFrName", null);
    __decorate([
        autobind
    ], SCW.prototype, "callAzureFunction", null);
    return SCW;
}(React.Component));
export default SCW;
//# sourceMappingURL=SCW.js.map