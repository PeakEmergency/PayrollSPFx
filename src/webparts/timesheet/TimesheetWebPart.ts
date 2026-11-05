import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { type IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { ThemeProvider, createTheme, initializeIcons, type ITheme } from '@fluentui/react';
import ExpenseReport from '../../expenseReport/components/ExpenseReport';
import { SharePointReportService } from '../../expenseReport/services/SharePointReportService';

import * as strings from 'TimesheetWebPartStrings';
import TimesheetForm, { ITimesheetFormProps } from './components/TimesheetForm';
import { EMPLOYEE_LIST, Employee, PAY_PERIOD_LIST, PAY_RATE_LIST, SpEmployeeSource, SpPayPeriodSource, SpPayRateSource, SpSubmissionStore, SpTemplateSource } from '../../services/sharepoint';
import { EXPENSE_TEMPLATE_FILE } from '../../excel/fillExpenseReport';
import { submitTimesheet } from '../../services/submit';

export interface ITimesheetWebPartProps {
  emailTo: string;
  listTitle: string;
  templateLibraryUrl: string;
  // Expense report settings (blank uses the defaults below).
  expenseSiteUrl: string;
  reportsListTitle: string;
  defaultRate: string;
}

export default class TimesheetWebPart extends BaseClientSideWebPart<ITimesheetWebPartProps> {
  private _root: Root | undefined;
  private _employees: Promise<Employee[]> | undefined;
  private _expenseService: SharePointReportService | undefined;
  private _expenseKey = '';
  private _theme: ITheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    initializeIcons();
  }

  private renderExpenseReport(templateLibraryUrl: string): React.ReactElement {
    const siteUrl = (this.properties.expenseSiteUrl || this.context.pageContext.web.absoluteUrl).replace(/\/$/, '');
    const reportsListTitle = this.properties.reportsListTitle || 'ExpenseReports';
    const key = JSON.stringify([siteUrl, reportsListTitle]);
    if (!this._expenseService || this._expenseKey !== key) {
      this._expenseService = new SharePointReportService(
        {
          request: async (url: string, init: RequestInit): Promise<Response> => {
            const response = await this.context.spHttpClient.fetch(url, SPHttpClient.configurations.v1, {
              method: init.method,
              headers: new Headers(init.headers),
              body: init.body as string | Blob | ArrayBuffer | undefined
            });
            return response as unknown as Response;
          }
        },
        { siteUrl, reportsListTitle }
      );
      this._expenseKey = key;
    }
    return React.createElement(
      ThemeProvider,
      { theme: this._theme },
      React.createElement(ExpenseReport, {
        key,
        service: this._expenseService,
        defaultRate: this.properties.defaultRate || '0.48',
        getEmployees: this.getEmployees,
        userEmail: this.context.pageContext.user.email,
        getPayPeriods: () => new SpPayPeriodSource(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, PAY_PERIOD_LIST).getPayPeriods(),
        submission: {
          emailTo: this.properties.emailTo,
          senderEmail: this.context.pageContext.user.email,
          loadTemplate: () => new SpTemplateSource(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, templateLibraryUrl).loadFile(EXPENSE_TEMPLATE_FILE)
        },
        instanceId: this.instanceId,
        isDarkTheme: this._theme?.isInverted || false
      })
    );
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    if (!theme) return;
    this._theme = createTheme({ isInverted: theme.isInverted, palette: theme.palette, semanticColors: theme.semanticColors });
    if (this._root) this.render();
  }

  /** One load shared by both forms; a failure retries on the next call. */
  private getEmployees = (): Promise<Employee[]> => {
    this._employees =
      this._employees ||
      new SpEmployeeSource(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, EMPLOYEE_LIST).getEmployees().catch((err: Error) => {
        this._employees = undefined;
        throw err;
      });
    return this._employees;
  };

  public render(): void {
    const web = this.context.pageContext.web;
    const { emailTo, listTitle } = this.properties;
    const templateLibraryUrl = this.properties.templateLibraryUrl || `${web.serverRelativeUrl.replace(/\/$/, '')}/Timesheet Templates`;

    const props: ITimesheetFormProps = {
      emailTo,
      getEmployees: this.getEmployees,
      getPayPeriods: () => new SpPayPeriodSource(this.context.spHttpClient, web.absoluteUrl, PAY_PERIOD_LIST).getPayPeriods(),
      getPayRates: () => new SpPayRateSource(this.context.spHttpClient, web.absoluteUrl, PAY_RATE_LIST).getAllRates(),
      submit: async (data) =>
        submitTimesheet(data, {
          store: new SpSubmissionStore(this.context.spHttpClient, web.absoluteUrl, listTitle),
          templates: new SpTemplateSource(this.context.spHttpClient, web.absoluteUrl, templateLibraryUrl),
          emailTo,
          employeeId: this.context.pageContext.legacyPageContext.userId,
          senderEmail: this.context.pageContext.user.email
        }),
      expenseReport: this.renderExpenseReport(templateLibraryUrl)
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
            },
            {
              groupName: 'Expense report',
              groupFields: [
                PropertyPaneTextField('expenseSiteUrl', { label: 'SharePoint site URL', placeholder: this.context.pageContext.web.absoluteUrl }),
                PropertyPaneTextField('reportsListTitle', { label: 'Reports list', placeholder: 'ExpenseReports' }),
                PropertyPaneTextField('defaultRate', {
                  label: 'Default mileage rate per km',
                  description: 'New reports only. Saved reports keep their rates.',
                  onGetErrorMessage: (value: string): string =>
                    value && !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(value.trim()) ? 'Enter a decimal rate, for example 0.48.' : ''
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
