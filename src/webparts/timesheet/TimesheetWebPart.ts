import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'TimesheetWebPartStrings';
import TimesheetForm, { ITimesheetFormProps } from './components/TimesheetForm';
import { getCurrentEmployee, GraphMailer } from '../../services/graph';
import { SpSubmissionStore, SpTemplateSource } from '../../services/sharepoint';
import { submitTimesheet } from '../../services/submit';

export interface ITimesheetWebPartProps {
  emailTo: string;
  listTitle: string;
  templateLibraryUrl: string;
}

export default class TimesheetWebPart extends BaseClientSideWebPart<ITimesheetWebPartProps> {
  private _root: Root | undefined;
  private _graph: Promise<MSGraphClientV3> | undefined;

  private get graph(): Promise<MSGraphClientV3> {
    this._graph =
      this._graph ||
      this.context.msGraphClientFactory.getClient('3').catch((err: Error) => {
        this._graph = undefined; // retry on the next call
        throw new Error(`Connecting to Microsoft Graph failed: ${err.message}`);
      });
    return this._graph;
  }

  public render(): void {
    const web = this.context.pageContext.web;
    const { emailTo, listTitle } = this.properties;
    const templateLibraryUrl = this.properties.templateLibraryUrl || `${web.serverRelativeUrl.replace(/\/$/, '')}/Timesheet Templates`;

    const props: ITimesheetFormProps = {
      emailTo,
      getName: async () => getCurrentEmployee(await this.graph),
      submit: async (data) =>
        submitTimesheet(data, {
          store: new SpSubmissionStore(this.context.spHttpClient, web.absoluteUrl, listTitle),
          templates: new SpTemplateSource(this.context.spHttpClient, web.absoluteUrl, templateLibraryUrl),
          mailer: new GraphMailer(await this.graph),
          emailTo,
          employeeId: this.context.pageContext.legacyPageContext.userId
        })
    };

    this._root = this._root || createRoot(this.domElement);
    this._root.render(React.createElement(TimesheetForm, props));
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('emailTo', { label: strings.EmailToFieldLabel }),
                PropertyPaneTextField('listTitle', { label: strings.ListTitleFieldLabel }),
                PropertyPaneTextField('templateLibraryUrl', {
                  label: strings.TemplateLibraryUrlFieldLabel,
                  description: strings.TemplateLibraryUrlFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
