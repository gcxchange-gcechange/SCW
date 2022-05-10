import * as React from 'react';
import styles from './SCW.module.scss';
import { ISCWProps } from './ISCWProps';
import { ISCWState } from './ISCWState';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { MessageBarType, Stack, Label, Spinner, Image, DefaultButton, ImageFit, IImageProps, ActionButton } from 'office-ui-fabric-react';
import { Selection} from 'office-ui-fabric-react/lib/DetailsList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { MSGraphClient, HttpClientResponse, IHttpClientOptions, AadHttpClient } from "@microsoft/sp-http";
import { ISiteItem } from './ISiteItem';
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Pagination from 'office-ui-fabric-react-pagination';
import { ITemplate } from './ITemplate';
import { BaseWizard, WizardStep, IWizardStepValidationResult } from "../../../common/components/Wizard";
//import * as strings from 'SCWWebPartStrings';
import { ISelected } from './ISelected';
import * as microsoftTeams from '@microsoft/teams-js';
import { spaceDescFr } from 'SCWWebPartStrings';

import { SelectLanguage } from './SelectLanguage';

var owners = [];
let totalPages: number = 1;
var allTemplateItems: ITemplate[] = [];
var selTemplate: ISelected[] = [];
var currentSelectedKey: number = -1;

export enum MyWizardSteps {
  None = 0,
  FirstStep = 1,
  SecondStep = 2,
  ThirdStep = 4,
  FourthStep = 8,
  LastStep = 16
}

export class MyWizard extends BaseWizard<MyWizardSteps> {
}

export default class SCW extends React.Component<ISCWProps, ISCWState> {
  private _teamsContext: microsoftTeams.Context;

  public strings = SelectLanguage(this.props.prefLang);

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  private _selection = new Selection({
    onSelectionChanged: () => {
      if (this._selection.count != 0) {
        currentSelectedKey = (this._selection.getSelection()[0] as ITemplate).key;
      }
      this.setState({ selectionDetails: this._getSelectionDetails() });
    }
  });

  public selection1 = new Selection;

  constructor(props: ISCWProps, state: ISCWState) {
    super(props);
    this.state = {
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
      selectionDetails: this._getSelectionDetails(),
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
    this.onchangedTitle = this.onchangedTitle.bind(this);
    this.onchangedFrName = this.onchangedFrName.bind(this);
    this._getOwners = this._getOwners.bind(this);
  }

  private imagesTemplate(key, title){
    var templateSel: ISelected = {
      key: key,
      title: title,
    };
    selTemplate = [];

    selTemplate.push(templateSel);
    this.setState({
      selected: selTemplate,
    });
    console.log(this.state.selected);
  }

  private _closeWizard(completed: boolean = false) {
    this.setState({
      isWizardOpened: false,
      statusMessage: completed ? "The wizard has been completed" : "The wizard has been canceled",
      statusType: completed ? "OK" : "KO"
    });

    setTimeout(() => {
      this.setState({
        statusMessage: null,
        statusType: null
      });
    }, 3000);
  }

  private _onValidateStep(step: MyWizardSteps): IWizardStepValidationResult | Promise<IWizardStepValidationResult> {

    let isValid = true;
    let isValid1 = true;
    let isValid2 = true;
    let ValidResult = true;

    switch (step) {
      case MyWizardSteps.FirstStep:
        isValid = this.state.selected[0] !== undefined;
        return {
          isValidStep: isValid,
          errorMessage: !isValid ? "Select a template" : null
        };

      case MyWizardSteps.SecondStep:
        isValid = (!this.state.checkSite) ? true : false;
        return {
          isValidStep: isValid
        }
      case MyWizardSteps.ThirdStep:

        return new Promise((resolve) => {
          isValid = this.state.tellusEn.length >= 5 && this.state.tellusEn.length <= 500 ;
          isValid1 = this.state.tellusFr.length >= 5 && this.state.tellusFr.length <= 500 ;
          isValid2 = this.state.BusinessReason.length >= 5 && this.state.BusinessReason.length <= 500 ;

          if(isValid == true && isValid2== true && isValid1==true ){
            ValidResult = true;
          }else{
            ValidResult = false;
          }
          setTimeout(() => {
            resolve({
              isValidStep: ValidResult,
              errorMessage: !ValidResult? "Your input to third step is invalid" : null
            });
          });
        });
      default:
        return { isValidStep: true };
    }
  }

  private _renderMyWizard() {
    var listOwners = "";
    for (let step = 0; step < owners.length; step++) {
      if (listOwners == ""){
        listOwners = owners[step];
      }else{
        listOwners= listOwners+', '+owners[step];
      }
    }
    return <MyWizard
      mainCaption=""
      onCancel={() => this._closeWizard(false)}
      onCompleted={() => this.callAzureFunction()}
      onValidateStep={(step) => this._onValidateStep(step)}
      onTitleCheck={() => this._searchSite()}
      checkSite={this.state.checkSite}
      validatingMessage={this.state.wizardValidatingMessage}
      disableStep1={(this.state.selected[0] !== undefined ? false : true)}
      disableStep2={(this.state.title.length >= 5 && this.state.title.length <= 125 && this.state.frName.length >= 5 && this.state.frName.length <= 125 ? false : true)}
      disableStep4={(this.state.tellusEn.length >= 5 && this.state.tellusEn.length < 500 && this.state.tellusFr.length >= 5 && this.state.tellusFr.length < 500 && this.state.BusinessReason.length >= 5 && this.state.BusinessReason.length < 500 ? false : true)}
      disableStep8={(this.state.ownersNumber >= 2 ? false : true)}
      finishButtonLabel= {this.strings.btnSubmit}
      nextButtonLabel={this.strings.NextBtn}
      previousButtonLabel={this.strings.BackBtn}
    >
    <WizardStep caption={this.strings.menuTemplate} step={MyWizardSteps.FirstStep}>
      <div className={styles.wizardStep}>
        <h1 className={styles.titleStep}>{this.strings.titleTemplate}</h1>
        <p>{this.strings.paragrapheTemplate}</p>
          <div className="ms-Grid" dir="ltr">
            {this.state.templateItems.map(item => (
              <span key={item.key} className={styles.templateHolder}>
                <input
                 autoFocus={(item.key == 0 ? true : false)}
                  type="radio"
                  name="template"
                  id={`template-${item.key}`}
                  onClick={() => this.imagesTemplate(item.key, item.title)}
                  aria-label={`${this.strings.templateButtonLabel}${(this.strings.userLang == "EN" ? item.title : item.titleFR)}`}
                />
                <label
                  htmlFor={`template-${item.key}`}
                  className={`${styles.imagetest} ${(this.state.selected[0] !== undefined ? this.state.selected[0]["key"] == item.key ? styles.selected : "" : "")} ms-Grid-col`}
                >
                  <Image
                      title={`${this.strings.altTemplate}${(this.strings.userLang == "EN" ? item.title : item.titleFR)}`}
                      src= {item.url}
                      alt={`${this.strings.altTemplate}${(this.strings.userLang == "EN" ? item.title : item.titleFR)}`}
                      width={150}
                      height={250}
                      className="ms-Grid-col ms-sm12 ms-md6 ms-lg6"
                    />
                  <div className={"ms-Grid-col ms-sm5 ms-md5 ms-lg5"}>
                    <h4>{(this.strings.userLang == "EN" ? item.title : item.titleFR)}</h4>
                    <p title={(this.strings.userLang == "EN" ? item.description : item.descriptionFR)} className={styles.templateDesc}>{(this.strings.userLang == "EN" ? item.description : item.descriptionFR)}</p>
                  </div>
                </label>
              </span>
            ))}
          </div>
        {/* <Stack horizontal verticalAlign="center" horizontalAlign="center">
          <br /><br />
          <div className={styles.pagination}>
          <Pagination
              style={{margin : "100px", border: "solid blue 3px"}}
              currentPage={this.state.currentPage}
              totalPages={totalPages}
              siblingCount={0}
              className={styles.pagination}
              onChange={(page: number) => {
                if (this._selection.count != 0 ){
                  currentSelectedKey = (this._selection.getSelection()[0] as ITemplate).key;
                  console.log('current key is ', currentSelectedKey);
                }
                if (currentSelectedKey >= (page - 1) * 4 && currentSelectedKey <= page*4 ){
                  const newSelection1 = this._selection;
                  newSelection1.setItems(allTemplateItems);
                  newSelection1.setKeySelected(`${currentSelectedKey}`, true, false);
                }
                this.setState({
                  currentPage: page,
                  templateItems: allTemplateItems.slice((page - 1) * 4, page * 4),
                  selectionDetails: this._getSelectionDetails(),
                });
              }}
            />
          </div>
        </Stack> */}
      </div>
    </WizardStep>

      <WizardStep caption={this.strings.menuSpace} step={MyWizardSteps.SecondStep}>
        <div className={styles.wizardStep}>
          <h1 className={styles.titleStep}>{this.strings.titleSpace}</h1>
          <p>{this.strings.paragrapheSpace}</p>
          <em>{this.strings.validationTxtSpace}</em>
          <section className={styles.SectiontextField}>
            <div className="form-group">
              <Label htmlFor="englishLabelTitle" className={styles.labelBulingue} required>{this.strings.english}</Label>
              <span id="englishLabelDesc" style={{color: "#777777", textAlign: 'left', display: 'block'}}>{this.strings.ErrMustLetter}</span>
              <TextField title={this.strings.tooltipspaceNameEn} autoFocus  id="englishLabelTitle" onChanged={this.onchangedTitle} aria-labelledby="englishLabelDesc" errorMessage={(!this.state.isSiteEnNameRight) && this.state.error} />
              <br></br>
            </div>
            <div className="form-group">
              <Label htmlFor="frenchLabelTitle" required className={styles.labelBulingue}>{this.strings.french}</Label>
              <span id="frenchLabelDesc" style={{color: "#777777", textAlign: 'left', display: 'block'}}>{this.strings.ErrMustLetter}</span>
              <TextField title={this.strings.tooltipspaceNameFr} id="frenchLabelTitle"  onChanged={this.onchangedFrName} aria-labelledby="frenchLabelDesc" errorMessage={(!this.state.isSiteFrNameRight) && this.state.error} />
              <div className={`${styles.yes} form-group`}>
                <p>
                  <label aria-live="polite" className={(this.state.checkSite == false ? styles.greencheckSite : styles.redcheckSite)}> {this.state.isAvailiability}</label>
                </p>
              </div>
            </div>
          </section>
        </div>
      </WizardStep>

      <WizardStep caption={this.strings.menuTell} step={MyWizardSteps.ThirdStep}>
        <div className={styles.wizardStep}>
          <h1>{this.strings.titleTellUs}</h1>
          <p>{this.strings.paragrapheTellUs}</p>
          <em>{this.strings.validationTxtTellUs}</em>
          <section className={styles.SectiontextField}>
            <Label htmlFor="englishLabelDesc" className={styles.labelBulingue} required>{this.strings.spaceDescEn}</Label>
            <span id="englishLabelDescriber" style={{color: "#777777", textAlign: 'left', display: 'block'}}>{this.strings.DescriptionHelperText}</span>
            <TextField
              title={this.strings.tooltipdescEn}
              autoFocus
              multiline rows={4}
              value={this.state.tellusEn}
              placeholder={this.strings.phLetusEn}
              id="englishLabelDesc"
              aria-labelledby="englishLabelDescriber"
              errorMessage={(this.state.tellusEn.length >= 5 && this.state.tellusEn.length <= 500 ) ? '': (this.state.tellusEn.length > 1) && this.strings.DescriptionHelperText}
              onChanged={(v) => this.setState({ tellusEn: v })}></TextField>
            <Label htmlFor="frenchLabelDesc" className={styles.labelBulingue} required>{this.strings.spaceDescFr}</Label>
            <span id="frenchLabelDescriber" style={{color: "#777777", textAlign: 'left', display: 'block'}}>{this.strings.DescriptionHelperText}</span>
            <TextField
              title={this.strings.tooltipdescFr}
              multiline rows={4}
              value={this.state.tellusFr}
              id="frenchLabelDesc"
              aria-labelledby="frenchLabelDescriber"
              placeholder={this.strings.phLetusFr}
              errorMessage={(this.state.tellusFr.length >= 5 && this.state.tellusFr.length <= 500 ) ? '': (this.state.tellusFr.length > 1) && this.strings.DescriptionHelperText}
              onChanged={(v) => this.setState({ tellusFr: v })}></TextField>
            <Label htmlFor="businessLabel" className={styles.labelBulingue} required>{this.strings.businessReason}</Label>
            <span id="businessLabelDescriber" style={{color: "#777777", textAlign: 'left', display: 'block'}}>{this.strings.DescriptionHelperText}</span>
            <TextField
              title={this.strings.tooltipBusReason}
              multiline rows={4}
              id="businessLabel"
              aria-labelledby="businessLabelDescriber"
              value={this.state.BusinessReason}
              placeholder={this.strings.phBusinessReason}
              errorMessage={(this.state.BusinessReason.length >= 5 && this.state.BusinessReason.length <= 500 ) ? '': (this.state.BusinessReason.length > 1) && this.strings.DescriptionHelperText}
              onChanged={(v) => this.setState({ BusinessReason: v })}></TextField>
          </section>
        </div>
      </WizardStep>

      <WizardStep caption={this.strings.menuOwners} step={MyWizardSteps.FourthStep}>
        <div className={styles.wizardStep}>
          <h1 className={styles.titleStep}>{this.strings.titleOwners}</h1>
          <p>{this.strings.paragrapheOwners}</p>
          <p>{this.strings.validationTxtOwners}</p>

          <div className="form-group">
            
            <PeoplePicker
              context={this.props.context}
              peoplePickerCntrlclassName={styles.peoplePickerHidden}       
              disabled={true}/>
            <Label htmlFor="peopleLabel" className={styles.labelBulingue} required>{this.strings.owners}</Label>  
            <PeoplePicker
            placeholder={this.strings.owners}
            showtooltip={true}
              tooltipMessage={this.strings.tooltipOwners}
              context={this.props.context}
              personSelectionLimit={3}
              peoplePickerCntrlclassName={styles.peoplePickerTest}
              groupName={""}
              showHiddenInUI={false}
              defaultSelectedUsers = {[this.props.context.pageContext.user.email]}
              required = {true} 
              onChange = {this._getOwners}                      
              ensureUser={false}/>
              
          </div>
          <p>{this.strings.ownerInfo1}</p>
          <p>{this.strings.ownerInfo2}</p>
          <p>{this.strings.ownerInfo3}</p>
          <p>{this.strings.ownerInfo4}</p>
        </div>
      </WizardStep>

      <WizardStep caption={this.strings.menuFinal} step={MyWizardSteps.LastStep}>
        <div className={styles.wizardStep}>
        <h1 className={styles.titleStep}>{this.strings.titleReview}</h1>
        <p>{this.strings.lastStepMakeChanges}</p>
        
          {(this.state.loading ?
            <div>
              <Label>{this.strings.textLoading}</Label>
              <Spinner label={this.strings.iconLoading} ariaLabel={this.strings.iconLoading} ariaLive="assertive" labelPosition="left" />
            </div>
            :
            <div>
              <div className="ms-Grid-row ms-sm12 ms-md12 ms-lg12">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Label htmlFor="templateLabel" className={styles.labelBulingue} required>{this.strings.templateTitle}</Label>
                <TextField
                  autoFocus
                  title={this.strings.templateTitle}
                  id="templateLabel"
                  readOnly
                  disabled
                  value={(this.state.selected[0] !== undefined ? this.state.selected[0]["title"]: "")}
                  placeholder="template"></TextField>
                  </div>
                  <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Label htmlFor="spaceNameLabel" className={styles.labelBulingue} required>{this.strings.spaceName}</Label>
                <TextField
                  title={this.strings.spaceName}
                  id="spaceNameLabel"
                  readOnly
                  disabled
                  value={`'${this.state.title} - ${this.state.frName}'`}
                  placeholder="Space Name"></TextField>
                </div>

              </div>
              <div className="ms-Grid-row ms-sm12 ms-md12 ms-lg12">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
              <Label htmlFor="spaceEnLabel" className={styles.labelBulingue} required>{this.strings.spaceDescEn}</Label>
                <TextField
                  title={this.strings.spaceDescEn}
                  id="spaceEnLabel"
                  readOnly
                  disabled
                  defaultValue={this.state.tellusEn}
                  placeholder="Descripton en"></TextField>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Label htmlFor="spaceFrLabel" className={styles.labelBulingue} required>{this.strings.spaceDescFr}</Label>
                <TextField
                  title={this.strings.spaceDescFr}
                  id="spaceFrLabel"
                  readOnly
                  disabled
                  defaultValue={this.state.tellusFr}
                  placeholder="Description fr"></TextField>
                  </div>
              </div>

              <div className="ms-Grid-row ms-sm12 ms-md12 ms-lg12">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
              <Label htmlFor="ownersLabel" className={styles.labelBulingue} required>{this.strings.owners}</Label>
                <TextField
                  title={this.strings.owners}
                  id="ownersLabel"
                  multiline
                  autoAdjustHeight
                  readOnly
                  disabled
                  value={listOwners}
                  placeholder="Owners"></TextField>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Label htmlFor="teamPurposeLabel" className={styles.labelBulingue} required>{this.strings.teamPurpose}</Label>
                <TextField
                  title={this.strings.teamPurpose}
                  id="teamPurposeLabel"
                  multiline
                  autoAdjustHeight
                  readOnly
                  disabled
                  value={this.state.BusinessReason}
                  placeholder="Business Reason"></TextField>
                </div>
              </div>
            </div>
          )}
        </div>
      </WizardStep>
      <div>
        Invalid element here, will be ignored
      </div>
    </MyWizard >;
  }

  private _openWizard() {
    this.setState({
      isWizardOpened: true
    });
  }

  private ResetScreen() {
    this.setState({ title: '',
    showMessageBar: false, //need to be false
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
  }
  public render(): React.ReactElement<ISCWProps> {
    const imageWelcome: IImageProps = {
      src: require("../../../../assets/sharepoint_teams_graphic.png"),
      imageFit: ImageFit.contain,
      width: 300,
      height: 150,
    };

    const imageCongrat: IImageProps = {
      src: require("../../../../assets/gcxchange_support_pencil.png"),
      imageFit: ImageFit.contain,
      width: 300,
      height: 150,
    };

    return (
      <div className={styles.container}>
        <div className={styles.row}>
          <div>
            {this.state.isWizardOpened
              ? this._renderMyWizard()
              : this.state.showMessageBar
                ?
                  <div className={styles.congratScreen}>
                     <Image
                       {...imageCongrat}
                      alt={this.strings.altCongrat}
                      className={styles.imageFit}
                    />
                    <h1>{this.strings.congrats}</h1>
                    <p>{this.strings.congratPara1}</p>
                    <p aria-live="polite">{this.strings.congratPara2}</p>
                    <DefaultButton className={styles.homeButton} autoFocus href="https://gcxgce.sharepoint.com/">{this.strings.congratHome}</DefaultButton>
                    <p>{this.strings.congratPara3} <a href="https://gcxgce.sharepoint.com/sites/Support"> {this.strings.congratLink}</a></p>
                  </div>
                :
                  <div className={styles.welcomeContainer}>
                     <Image
                       {...imageWelcome}
                      alt={this.strings.altWelcome}
                      className={styles.imageFit}
                      title={this.strings.tooltipWelImg}
                    />
                    <h1 className={ styles.titleStep }>{this.strings.createSpace}</h1>
                    <p>{this.strings.paragrapheHome}</p>
                    {this.state.templateItems.length !=0 ?
                      <ActionButton title={this.strings.startButton} className={styles.GoButton} text={this.strings.startButton} onClick={() => this._openWizard()} />
                     :
                      <div>
                        <Spinner label={this.strings.iconLoading} ariaLabel={this.strings.iconLoading} ariaLive="assertive" />
                      </div>
                    }
                    <div className={ styles.poweredByText }>{this.strings.powered}
                    <br></br>{this.strings.gcx}</div>
                  </div>
            }
          </div>
        </div>
      </div>
    );
  }
  protected functionTemplateImg: string = "";
  public async componentDidMount(): Promise<void> {
    await this.loadTemplate();
  }

  private async loadTemplate(){
    var allTemplateItems = [];
    await this.props.context.aadHttpClientFactory.getClient("").then((client: AadHttpClient) => {
      client.get(this.functionTemplateImg, AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        console.log(`Status code: ${response.status}`);
        response.json().then((responseJSON: JSON) => {
        var i = 0;
        for (var k in responseJSON) {

          var template: ITemplate = {
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
            this.setState({
              templateItems: allTemplateItems,
            });
          } else {
          console.log("Response error");
          }
        })
        .catch((response: any) => {
          let errMsg: string = `WARNING - error when calling URL ${this.functionUrl}. Error = ${response.message}${response.status}${JSON.stringify(response)}`;
          console.log("err is ", errMsg);
        });
      });

    });
  }


  private _getSelectionDetails(): string {
    const selectionCount = this._selection.count;
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return (
          (this._selection.getSelection()[0] as ITemplate).title
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _getOwners(ownersFromPeoplePicker: any[]) {
    owners = [this.props.context.pageContext.user.email];
    for (let item in ownersFromPeoplePicker) {
      if (this.props.context.pageContext.user.email === ownersFromPeoplePicker[item].secondaryText) {
        console.log(ownersFromPeoplePicker);
      } else {
        owners.push(ownersFromPeoplePicker[item].secondaryText);
        console.log("owner is", owners); 
      }
    }
    this.setState({ ownersNumber: owners.length });
  }

  private onchangedTitle(title: string): void {
    // check length, only include letter、number and -   title.length < 5 || title.length > 10 ||
    if (title.match("^([a-zA-Z0-9'()&,/#’ ]*)+$") == null || title.length < 5 || title.length > 125) {
      this.setState({
        isSiteEnNameRight: false,
        error: this.strings.ErrMustLetter,
      });
    } else {
      this.setState({ error: "" });
      this.setState({ isSiteEnNameRight: true });
    }
    this.setState({
      title: title,
      isAvailiability: "",
    });
  }

  private onchangedFrName(frName: any): void {
    if (frName.match("^([A-Za-z0-9àâäèéêëîïôœùûüÿçÀÂÄÈÉÊËÎÏÔŒÙÛÜŸÇ'()&,/#’ ]*)+$") == null || frName.length < 5 || frName.length > 125) {
      this.setState({ error: this.strings.ErrMustLetter });
      this.setState({ isSiteFrNameRight: false });
    } else {
      this.setState({ error: "" });
      this.setState({ isSiteFrNameRight: true });
    }
    this.setState({
      frName: frName,
      isAvailiability: "",
    });

  }

  private _searchSite = (): void => {

    // Log the current operation
    // Grab the first few characters of the title to perform a search or similar titles
    var searchItem = this.state.title.slice(0,5);
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("groups")
          .filter(`startswith(displayName, '${searchItem}')`)
          .select("displayName,id,name")
          .get((err, res) => {

            if (err) {
              console.error(err);
              return;
            }
            // Prepare the output array
            var sites: Array<ISiteItem> = new Array<ISiteItem>();
            // Map the JSON response to the output array
            if (res) {
              res.value.map((item: any) => {
                var enValid = true;
                var frValid = true;
                var split  = item.displayName.split("-");
                if (split.length === 1) {
                  // Checking titles with no split and stop
                  if (split[0].trim().toLowerCase() === this.state.title.toLowerCase()){
                    enValid = false;
                    frSplit = false;
                  }
                } else {
                  var enSplit = split[0].trim().toLowerCase();
                  var frSplit = split[1].trim().toLowerCase();

                  if (enSplit === this.state.title.toLowerCase()) {
                    enValid = false;
                  }
                  if (frSplit === this.state.frName.toLowerCase()) {
                    frValid = false
                  }
                }

                sites.push({
                  displayName: item.displayName,
                  id: item.id,
                  enIsValid: enValid,
                  frIsValid: frValid,
                });
              });
            }
            var testing = [];
            if (sites.length != 0) {
              sites.map((s: any) => {
                if (!s.enIsValid) {
                  this.setState({
                    error: this.strings.siteTaken,
                    isSiteEnNameRight: false,
                  })
                  testing.push('False');
                }
                if(!s.frIsValid) {
                  this.setState({
                    error: this.strings.siteTaken,
                    isSiteFrNameRight: false,
                  })
                  testing.push('False');
                }
              })
             if (testing.length != 0) {
              console.log('stay');
             } else {
               console.log('go');
               this.setState(
                {
                  sites: sites,
                  isAvailiability: this.strings.greatChoice,
                  checkSite: false
                }
              );
              if (document.getElementById('nextBtn')) {
                document.getElementById('nextBtn').click();
              }
             }
              this.setState(
                {
                  sites: sites,
                  isAvailiability: this.strings.siteTaken,
                  checkSite: true
                }
              );

            } else {
              this.setState(
                {
                  sites: sites,
                  isAvailiability: this.strings.greatChoice,
                  checkSite: false
                }
              );
              if (document.getElementById('nextBtn')) {
                document.getElementById('nextBtn').click();
              }
            }
          });
      });
  }

  protected functionUrl: string = "";

  private callAzureFunction(): void {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");

    var siteUrl: string = this.props.context.pageContext.web.absoluteUrl;
    var owner1: any, owner2: any, owner3: string;
    if (owners.length == 2) {
      owner1 = owners[0];
      owner2 = owners[1];
      owner3 = "";
    } else {
      owner1 = owners[0];
      owner2 = owners[1];
      owner3 = owners[2];
    }

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `
        {
          "name":
          {
            "title": "${this.state.title} - ${this.state.frName}",
            "spacenamefr": "${this.state.frName}",
            "owner1": "${owner1}",
            "owner2": "${owner2}",
            "owner3": "${owner3}",
            "description": "${this.state.tellusEn}",
            "descriptionFr": "${this.state.tellusFr}",
            "business":"${this.state.BusinessReason}",
            "template": "${this.state.selected[0]["title"]}",
            "requester_name": "${this.props.context.pageContext.user.displayName}",
            "requester_email": "${this.props.context.pageContext.user.email}",
          }
        }`
    };

    let responseText: string = "";
    // use aad authentication
    this.setState({loading:true}, () => {
      this.props.context.aadHttpClientFactory.getClient("").then((client: AadHttpClient) => {
        client.post(this.functionUrl, AadHttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
          console.log(`Status code: ${response.status}`);
          this.setState({
            showMessageBar: true,
            messageType: MessageBarType.success,
            isWizardOpened: false,
            loading: false
          });
          this.SendEmail();
          response.json().then((responseJSON: JSON) => {
            responseText = JSON.stringify(responseJSON);
            console.log("respond is ", responseText);
            if (response.ok) {
              console.log("response OK");
            } else {
            console.log("Response error");
            }
          })
          .catch((response: any) => {
            let errMsg: string = `WARNING - error when calling URL ${this.functionUrl}. Error = ${response.message}`;
            console.log("err is ", errMsg);
          });
        });
      });
    });
  }

  protected emailQueueUrl: string = "";

  private SendEmail(): void {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postQueue: IHttpClientOptions = {
      headers: requestHeaders,
      body: `
      {
          "name": "${this.state.title}-${this.state.frName}",
          "status": "Submitted",
          "requesterName": "${this.props.context.pageContext.user.displayName}",
          "requesterEmail": "${this.props.context.pageContext.user.email}"
      }`
  };
    this.props.context.aadHttpClientFactory.getClient("").then((client: AadHttpClient) => {
      client.post(this.emailQueueUrl, AadHttpClient.configurations.v1, postQueue).then((response: HttpClientResponse) => {
          console.log(`Status code:`, response.status);
          console.log('respond is ', response.ok);
          console.log('send reject message to queue successful.');
          console.log(`requester Email`, this.props.context.pageContext.user.email);
      });
  });

            const postHD: IHttpClientOptions = {
            headers: requestHeaders,
            body: `
      {
          "name": "${this.state.title}-${this.state.frName}",
          "status": "Notif_HD",
      }`
        };

    this.props.context.aadHttpClientFactory.getClient("").then((client: AadHttpClient) => {
            client.post(this.emailQueueUrl, AadHttpClient.configurations.v1, postHD).then((response: HttpClientResponse) => {
                console.log(`Status code:`, response.status);
                console.log('respond is ', response.ok);
                console.log('send notif_HD message to queue successful.');
                console.log(`requester Email`, this.props.context.pageContext.user.email);
            });
        });
  }
}
