import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';

import PpeForm from "./components/PpeForm";
import { IPpeFormWebPartProps } from "./components/IPpeFormProps";
// import TEMPPpeForm from "./components/TempPPEForm";


export default class PpeFormWebPart extends BaseClientSideWebPart<IPpeFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    return new Promise<void>(async (resolve, reject) => {
      this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();
      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
      resolve();
    });
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  public render(): void {

    const element: React.ReactElement<IPpeFormWebPartProps> =
      React.createElement(PpeForm, {
        context: this.context,
        ThemeColor: this._themeVariant?.palette?.themePrimary,
        IsDarkTheme: this._isDarkTheme,
        HasTeamsContext: !!this.context.sdks.microsoftTeams,
      });

    ReactDom.render(element, this.domElement);
  }
}
