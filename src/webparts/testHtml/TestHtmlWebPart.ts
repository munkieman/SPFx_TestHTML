import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TestHtmlWebPart.module.scss';
import * as strings from 'TestHtmlWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

require('bootstrap');

export interface ITestHtmlWebPartProps {
  description: string;
}

export default class TestHtmlWebPart extends BaseClientSideWebPart<ITestHtmlWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private checkFolder(subfolder1Flag:boolean,subfolder2Flag:boolean,subfolder3Flag:boolean, libraryName:string) : void {
    let html: string = ``; // <div class="accordion mt-1" id="accordionPF"><div class="accordion-item">;
    let htmlEnd : string = ``;

    alert(subfolder1Flag);
    alert(subfolder2Flag);
    alert(subfolder3Flag);

    if(subfolder1Flag){       
      html+=`<div class="accordion" id="accordionPF">
                <div class="accordion-item">
                  <h2 class="accordion-header" id="headingPF">
                    <button class="btn btn-primary accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF1" aria-expanded="true" aria-controls="collapseSF1">
                      <i class="bi bi-folder2"></i>
                      <a href="#" class="text-white ms-1">Folder</a>
                      <span class="badge bg-secondary">0</span>                    
                    </button>
                  </h2>`;

    }else{
      html+=`<div class="accordion" id="accordionPF">
                <div class="accordion-item">
                  <h2 class="accordion-header" id="headingOne">
                    <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseOne">
                      <i class="bi bi-folder2"></i>
                      <a href="#" class="text-white ms-1">Folder</a>
                      <span class="badge bg-secondary">0</span>                    
                    </button>
                  </h2>`;
    }

    if(subfolder1Flag){

      if(subfolder2Flag){
        html+=`<div id="collapseSF1" class="accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF">
                  <div class="accordion-body">
                    <div class="accordion" id="accordionSF1">                             
                      <div class="accordion-item">
                        <h2 class="accordion-header" id="headingSF1">
                          <button class="btn btn-primary accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2" aria-expanded="false" aria-controls="collapseSF3-1">
                            <i class="bi bi-folder2"></i>
                            <a href="#" class="text-white ms-1">SubFolder 1</a>
                            <span class="badge bg-secondary">0</span>                        
                          </button>
                        </h2>`;
        htmlEnd+=`</div></div></div></div>`;

      }else{
        html+=`<div id="collapseSF1" class="ms-1 accordion-collapse collapse" aria-labelledby="headingPF" data-bs-parent="#accordionPF">
                <div class="accordion-body">
                  <div class="accordion" id="accordionSF1"
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="headingSF1">
                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1-1">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1">SubFolder 1</a>
                          <span class="badge bg-secondary">0</span>                    
                        </button>
                      </h2>
                    </div>
                  </div>
                </div>`;
      }               
    }

    if(subfolder2Flag){

      if(subfolder3Flag){
        html+=`<div id="collapseSF2" class="accordion-collapse collapse" aria-labelledby="headingSF3-1" data-bs-parent="accordionSF1">
                <div class="accordion-body">
                  <div class="accordion" id="accordionSF3">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="headingSF3">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2" aria-expanded="false" aria-controls="collapseSF2">
                          SubFolder 2  
                        </button>
                      </h2>`;
        htmlEnd+=`</div></div></div></div>`;

      }else{
        html+=`<div id="collapseSF2" class="accordion-collapse collapse" aria-labelledby="headingSF3-1" data-bs-parent="accordionSF1">
                <div class="accordion-body">
                  <div class="accordion" id="accordionSF2">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="headingSF2">
                        <button class="btn btn-primary" type="button" data-bs-target="#collapseSF2" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF2">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1">SubFolder 2</a>
                          <span class="badge bg-secondary">0</span>                    
                        </button>
                      </h2>
                    </div>
                  </div>
                </div>
              </div>`;
      }               
    }          

    if(subfolder3Flag){
      html+=`<div id="collapseSF3" class="accordion-collapse collapse" aria-labelledby="headingSF3-2" data-bs-parent="accordionSF2">
              <div class="accordion-body">
                <h2 class="accordion-header" id="headerSF3">
                  <button 
                    class="btn btn-primary" 
                    type="button" 
                    data-bs-toggle="collapse" 
                    data-bs-target="#collapseSF3" 
                    aria-expanded="true" 
                    aria-controls="collapseSF3">
                      <i class="bi bi-folder2"></i>
                      <a href="#" id="subfolder3Name"> 
                        Sub Folder 3
                      </a>
                  </button>
                </h2>
              </div>
            </div>`;
    }    

    html+=htmlEnd;
    html+=`</div></div>`;
    console.log(html);
   
    const listContainer: Element = this.domElement.querySelector('#'+libraryName)!;
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.testHtml} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div class="container-fluid">
        <div class="row">
          <div class="col-2 v-scrollbar">
            <div class="d-flex mt-1 align-items-start libraryContainer">
              <div class="nav flex-column nav-pills me-3 libraryList" id="v-pills-tab" role="tablist" aria-orientation="vertical">
                <button class="nav-link" id="policiesTab" data-bs-toggle="pill" data-bs-target="#policies" type="button" role="tab" aria-controls="policies" aria-selected="true">Policies</button>
                <button class="nav-link" id="proceduresTab" data-bs-toggle="pill" data-bs-target="#procedures" type="button" role="tab" aria-controls="procedures" aria-selected="false">Procedures</button>
                <button class="nav-link" id="guidesTab" data-bs-toggle="pill" data-bs-target="#guides" type="button" role="tab" aria-controls="guides" aria-selected="false">Guides</button>
                <button class="nav-link" id="formsTab" data-bs-toggle="pill" data-bs-target="#forms" type="button" role="tab" aria-controls="forms" aria-selected="false">Forms</button>
                <button class="nav-link" id="generalTab" data-bs-toggle="pill" data-bs-target="#general" type="button" role="tab" aria-controls="general" aria-selected="false">General</button>
              </div>
            </div>
          </div>
          <div class="col-10">
            <div class="tab-content" id="v-pills-tabContent">
              <div class="tab-pane fade" id="policies" role="tabpanel" aria-labelledby="policiesTab"> 
                <div class="row">
                  <div class="col-6 v-scrollbar" id="policyFolders"></div>
                  <div class="col-6 v-scrollbar" id="policyFiles"></div>
                </div>               
              </div>
              <div class="tab-pane fade" id="procedures" role="tabpanel" aria-labelledby="proceduresTab">
                <div class="row">
                  <div class="col-6 v-scrollbar" id="procsFolders"></div>
                  <div class="col-6 v-scrollbar" id="procsFiles"></div>
                </div> 
              </div>
              <div class="tab-pane fade" id="guides" role="tabpanel" aria-labelledby="guidesTab">
                <div class="row">
                  <div class="col-6 v-scrollbar" id="guidesFolders"></div>
                  <div class="col-6 v-scrollbar" id="guidesFiles"></div>
                </div> 
              </div>
              <div class="tab-pane fade" id="forms" role="tabpanel" aria-labelledby="formsTab">
                <div class="row">
                  <div class="col-6 v-scrollbar" id="formsFolders"></div>
                  <div class="col-6 v-scrollbar" id="formsFiles"></div>
                </div> 
              </div>
              <div class="tab-pane fade" id="general" role="tabpanel" aria-labelledby="generalTab">
                <div class="row">
                  <div class="col-6 v-scrollbar" id="generalFolders"></div>
                  <div class="col-6 v-scrollbar" id="generalFiles"></div>
                </div> 
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>`;


    //this.checkFolder(true,true,false,"procsFolders");
    document.getElementById('policiesTab')!.addEventListener("click",(e:Event) => this.checkFolder(true,true,true,"policyFolders"));
    document.getElementById('proceduresTab')!.addEventListener("click",(e:Event) => this.checkFolder(true,true,false,"procsFolders"));
    document.getElementById('guidesTab')!.addEventListener("click",(e:Event) => this.checkFolder(true,false,false,"guidesFolders"));
    document.getElementById('formsTab')!.addEventListener("click",(e:Event) => this.checkFolder(true,true,true,"formsFolders"));
    document.getElementById('generalTab')!.addEventListener("click",(e:Event) => this.checkFolder(true,true,false,"generalFolders"));

  }

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}