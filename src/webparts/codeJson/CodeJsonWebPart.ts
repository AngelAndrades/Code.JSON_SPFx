import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CodeJsonWebPartStrings';

//import * as $ from 'jquery';
import '@progress/kendo-ui';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { sp } from '@pnp/sp/presets/all';
import { Store } from './state/store';
import { SPA } from './apps/spa';

export interface ICodeJsonWebPartProps {
  organization: string;
  contactName: string;
  contactEmail: string;
  vcs: string;
  homeLink: string;
  vasiExtractList: string;
  codeJsonList: string;
  instructionsLink: string;
  licensing: string;
  disclaimer: string;
}

export default class CodeJsonWebPart extends BaseClientSideWebPart<ICodeJsonWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
          spfxContext: this.context,
          sp: {
            headers: { Accept: 'application/json;odata=nometadata' }
          }
      });
    });
  }

  private store = new Store();

  private validateDisclaimer(value: string): string {
    if (value === null || value.trim().length === 0) return 'Disclaimer text is a required field.';
    else return '';
  }

  private validateOrganization(value: string): string {
    if (value === null || value.trim().length === 0) return 'Organization is a required field.';
    else return '';
  }

  private validateContactName(value: string): string {
    if (value === null || value.trim().length === 0) return 'Contact Name is a required field.';
    else return '';
  }

  private validateContactEmail(value: string): string {
    if (value === null || value.trim().length === 0) return 'Contact Email is a required field.';
    else return '';
  }

  private validateVersionControl(value: string): string {
    if (value === null || value.trim().length === 0) return 'Version Control System is a required field.';
    else return '';
  }

  private validateRepoUrl(value: string): string {
    if (value === null || value.trim().length === 0) return 'Repo Homepage URL is a required field.';
    else return '';
  }

  private validateVASIExtractList(value: string): string {
    if (value === 'NO_LIST_SELECTED') return 'VASI Extract List is a required field.';
    else return '';
  }

  private validateVASIAppendList(value: string): string {
    if (value === 'NO_LIST_SELECTED') return 'code.JSON List is a required field.';
    else return '';
  }

  public render(): void {
    if (this.properties.organization != null) this.store.set('organization', this.properties.organization);
    if (this.properties.contactName != null) this.store.set('contactName', this.properties.contactName);
    if (this.properties.contactEmail != null) this.store.set('contactEmail', this.properties.contactEmail);
    if (this.properties.vcs != null) this.store.set('vcs', this.properties.vcs);
    if (this.properties.homeLink != null) this.store.set('homeLink', this.properties.homeLink);
    if (this.properties.licensing != null) this.store.set('licensing', true);
    if (this.properties.disclaimer != null) this.store.set('disclaimer', this.properties.disclaimer);

    // only render if SharePoint Lists are set
    if(this.properties.vasiExtractList != null && this.properties.codeJsonList != null && this.properties.instructionsLink != null) {
      this.store.set('importList', this.properties.vasiExtractList);
      this.store.set('appendList', this.properties.codeJsonList);
      this.store.set('spLink', this.properties.instructionsLink);

      // load additional kendo dependencies
      SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.2.617/styles/kendo.common-material.min.css');
      SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.2.617/styles/kendo.material.min.css');
  
      // SPComponentLoader.loadScript('https://kendo.cdn.telerik.com/2020.2.617/js/kendo.all.min.js');
      SPComponentLoader.loadScript('https://kendo.cdn.telerik.com/2020.2.617/js/jszip.min.js');

      // web part DOM
      this.domElement.innerHTML = `
        <style>
          .k-edit-form-container {
              width: 600px;
          }
  
          .k-edit-label {
              width: 30%;
              margin: 0;
          }
  
          .k-input, .k-combobox, .k-dropdown, .k-numerictextbox, .k-textbox {
              width: 85% !important
          }
  
          .k-grid .k-grid-header .k-header .k-link { height: auto; }
  
          .k-grid .k-grid-header .k-header { white-space: normal; }
  
          .k-grid .k-header .k-grid-search { max-width: 20% !important }
  
          .k-toolbar>* {
              min-width: 10vw;
          }
        </style>
        <div id="dialog"></div>
        <div id="tabStrip">
          <ul>
            <li class="k-state-active">Appended Data</li>
            <li>Pre-filtered VASI Data Extract</li>
            <li>User Guide</li>
          </ul>
          <div>
            <h2>Predefined filter applied to the VASI Data Extract</h2>
            <div id="filter"></div>
            <div style="padding: 10px;"></div>
            <div id="appendGrid"></div>
          </div>
          <div>
            <div id="importGrid"></div>
          </div>
          <div>Redirecting you to the user guide...</div>
        </div>
      `;

      const app = SPA.getInstance(this.store);
    } else {
      this.domElement.innerHTML = `<div>
      <strong>Select the Property Panel for this web part and provide valid inputs for the following fields:</strong>
      <ul>
      <li>Organization</li>
      <li>Contact Name</li>
      <li>Contact Email</li>
      <li>Version Control System</li>
      <li>Repo Homepage URL</li>
      <li>Instruction Link: paste the link to the wiki page containing the instruction guide.</li>
      <li>VASI Extract List: choose the SharePoint custom list containing the imported VASI Data Extract.</li>
      <li>code.JSON List: choose the SharePoint custom list containing the appended information.</li>
      <li>Apply Creative Commons Zero Licensing: sets this license information as the default license, can be overridden at the project entry. If unchecked, there will be no default license information associate with each project record unless specifically entered for each project entry.</li>
      <li>Disclaimer Text: Sets the default disclaimer text, can be overridden at the project entry</li>
      </ul>
      </div>`;
    }
  }

  // prevent the property pane from rendering on changes (causing data leaks)
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Default property values',
              groupFields: [
                PropertyPaneTextField('organization', {
                  label: 'Organization',
                  onGetErrorMessage: this.validateOrganization.bind(this)
                }),
                PropertyPaneTextField('contactName', {
                  label: 'Contact Name',
                  onGetErrorMessage: this.validateContactName.bind(this)
                }),
                PropertyPaneTextField('contactEmail', {
                  label: 'Contact Email',
                  onGetErrorMessage: this.validateContactEmail.bind(this)
                }),
                PropertyPaneTextField('vcs', {
                  label: 'Version Control System',
                  onGetErrorMessage: this.validateVersionControl.bind(this)
                }),
                PropertyPaneTextField('homeLink', {
                  label: 'Repo Homepage URL',
                  onGetErrorMessage: this.validateRepoUrl.bind(this)
                }),
                PropertyPaneTextField('instructionsLink', {
                  label: 'Instruction Link'
                }),
                PropertyFieldListPicker('vasiExtractList', {
                  label: 'Select VASI Extract List',
                  selectedList: this.properties.vasiExtractList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: this.validateVASIExtractList.bind(this),
                  deferredValidationTime: 0,
                  key: 'vasiExtractListId'
                }),
                PropertyFieldListPicker('codeJsonList', {
                  label: 'Select code.JSON List',
                  selectedList: this.properties.codeJsonList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: this.validateVASIAppendList.bind(this),
                  deferredValidationTime: 0,
                  key: 'codeJsonListId'
                }),
                PropertyPaneCheckbox('licensing', {
                  text: 'Apply Creative Commons Zero Licensing'
                }),
                PropertyPaneTextField('disclaimer', {
                  label: 'Disclaimer Text',
                  multiline: true,
                  onGetErrorMessage: this.validateDisclaimer.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}