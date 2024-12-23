import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
	PropertyPaneHorizontalRule,
	IPropertyPaneConfiguration,
  	PropertyPaneTextField,
	IPropertyPaneGroup,
	PropertyPaneSlider,
	PropertyPaneCheckbox,
 	PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";

import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SovaPopupWebPartStrings';
import SovaPopup from './components/SovaPopup';
import { ISovaPopupProps } from './components/ISovaPopupProps';

export interface ISovaPopupWebPartProps {
	popupContentType: number, 				//	1 = HTML text, 2 = URL to HTML file
	HTMLContent: string,
	HTMLFilePickerResult: IFilePickerResult,
	closeButtonCaption: string,

	autoShowOnDelay: boolean;
	delaySeconds: number,
	openButtonCaption: string,					//	open popup button caption

	vPos: number,
	hPos: number,
	widthType: number,						// pixels or percentage
	widthPercentage: number,
	widthPixels: number,
	heightType: number,						// pixels or percentage
	heightPercentage: number,
	heightPixels: number,

	redirectOnClose: boolean,
	redirectURL: string,

	rememberUserCloseAction: boolean,
	rememberUserCloseActionMaxN: number
}

export default class SovaPopupWebPart extends BaseClientSideWebPart<ISovaPopupWebPartProps> {

	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = '';
	private _isInEditMode: boolean = false;

	protected get isRenderAsync(): boolean {
		return true;
	}

	//public render(): void {
	public async render(): Promise<void> {
		if(Environment.type == EnvironmentType.SharePoint){
			if(this.displayMode == DisplayMode.Edit) this._isInEditMode = true;
		}

		// IF POPUP CONTENT SHOULD BE READ FROM A FILE, READ THE CONTENT FROM THE FILE
		let sContentToDisplay = "";
		if (this.properties.popupContentType == 1) {
			sContentToDisplay = this.properties.HTMLContent;
		}
		else if (this.properties.popupContentType == 2){	// load data from a standalone document
			if (this.properties.HTMLFilePickerResult && this.properties.HTMLFilePickerResult.fileAbsoluteUrl && this.properties.HTMLFilePickerResult.fileAbsoluteUrl != ""){
				let fileUrl = this.properties.HTMLFilePickerResult.fileAbsoluteUrl;
				try{
					let res:SPHttpClientResponse = await this.context.spHttpClient.get(fileUrl, SPHttpClient.configurations.v1);
					if (res.ok){
						sContentToDisplay = await res.text();
					}else {
						console.log("POPUP ERROR: Error loading content to display from the file URL - " + fileUrl);
						sContentToDisplay = "";
					}
				}catch{
					console.log("POPUP ERROR: General error notification while access file URL - " + fileUrl);
				}
			}
		}

		const wpProperties:ISovaPopupProps = {
			popupContentType: this.properties.popupContentType,
			popupContentToDisplay: sContentToDisplay,
			closeButtonCaption: this.properties.closeButtonCaption,

			autoShowOnDelay: this.properties.autoShowOnDelay,
			delaySeconds: this.properties.delaySeconds,
			openButtonCaption: this.properties.openButtonCaption,

			vPos: this.properties.vPos,
			hPos: this.properties.hPos,
			widthType: this.properties.widthType,
			widthPercentage: this.properties.widthPercentage,
			widthPixels: this.properties.widthPixels,
			heightType: this.properties.heightType,
			heightPercentage: this.properties.heightPercentage,
			heightPixels: this.properties.heightPixels,

			redirectOnClose: this.properties.redirectOnClose,
			redirectURL: this.properties.redirectURL,

			rememberUserCloseAction: this.properties.rememberUserCloseAction,
			rememberUserCloseActionMaxN: this.properties.rememberUserCloseActionMaxN,

			webPartId: this.context.instanceId,
			context:this.context,
			isEditMode: this._isInEditMode,
			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName
		};

		const element: React.ReactElement<any, any> = React.createElement(SovaPopup, wpProperties, null);
		ReactDom.render(element, this.domElement);

		this.renderCompleted();
	}

	protected onInit(): Promise<void> {
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

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		// *****************************************************
		// SET IMPORTANT INITIAL VALUES IF NOT ALREADY SET
		// *****************************************************

		/*if (!this.properties.popupContentType) this.properties.popupContentType = 1;
		if (!this.properties.HTMLContent) this.properties.HTMLContent = "";
		if (!this.properties.openButtonCaption) this.properties.openButtonCaption = "Show popup";

		if (!this.properties.autoShowOnDelay) this.properties.autoShowOnDelay = false;
		if (!this.properties.delaySeconds) this.properties.delaySeconds = 3;
		if (!this.properties.closeButtonCaption) this.properties.closeButtonCaption = "";

		if (!this.properties.vPos) this.properties.vPos = 2;
		if (!this.properties.hPos) this.properties.hPos = 2;

		if (!this.properties.widthType) this.properties.widthType = 1;
		if (!this.properties.widthPercentage) this.properties.widthPercentage = 50;
		if (!this.properties.widthPixels) this.properties.widthPixels = 400;

		if (!this.properties.heightType) this.properties.heightType = 1;
		if (!this.properties.heightPercentage) this.properties.heightPercentage = 50;
		if (!this.properties.heightPixels) this.properties.heightPixels = 300;

		if (!this.properties.redirectOnClose) this.properties.redirectOnClose = false;
		if (!this.properties.redirectURL) this.properties.redirectURL = "";

		if (!this.properties.rememberUserCloseAction) this.properties.rememberUserCloseAction = false;
		if (!this.properties.rememberUserCloseActionMaxN) this.properties.rememberUserCloseActionMaxN = 0;*/

		// *****************************************************
		// BASIC GROUP FIELDS
		// *****************************************************

		const basicGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('popupContentType', {
				label: "Popup content type",
				options: [
					{key: 1, text: "HTML text"},
					{key: 2, text: "URL to HTML file"}
				],
				selectedKey: this.properties.popupContentType
			})
		];
		if (this.properties.popupContentType == 1){
			basicGroupFields.push(
				PropertyPaneTextField('HTMLContent', {
					label: "HTML content for popup",
					multiline: true,
					rows:8,
					value: this.properties.HTMLContent
				}));
		}
		if (this.properties.popupContentType == 2){
			basicGroupFields.push(
				PropertyFieldFilePicker('filePicker', {
					context: this.context,
					accepts:[".html", ".htm"],
					hideStockImages: true,
					hideLocalUploadTab: true,
					hideOrganisationalAssetTab: true,
					filePickerResult: this.properties.HTMLFilePickerResult,
					onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
					properties: this.properties,
					onSave: (e: IFilePickerResult) => {
						//console.log(e);
						this.properties.HTMLFilePickerResult = e;  },
					onChanged: (e: IFilePickerResult) => {
						//console.log(e);
						this.properties.HTMLFilePickerResult = e; },
					key: "filePickerId",
					buttonLabel: "Choose HTML file",
					label: "Select the file with HTML content",
				})
			);
		}
		basicGroupFields.push(
			PropertyPaneHorizontalRule(),
			PropertyPaneTextField('closeButtonCaption', {
				label: "Caption for close popup button",
				description: "If this field is empty, the close popup button won't be displayed. An 'X' icon will be displayed to close the popup.",
				value: this.properties.closeButtonCaption
			})
		);
		basicGroupFields.push(
			PropertyPaneHorizontalRule(),
			PropertyPaneCheckbox("autoShowOnDelay", {
				text: "Auto-show popup",
				checked: this.properties.autoShowOnDelay
			}),
			PropertyFieldNumber('delaySeconds',{
				key: "delaySeconds",
				label: "Delay (seconds)",
				value: this.properties.delaySeconds,
				minValue: 0,
				maxValue: 3000,
				description: "After how many seconds does the popup automatically appears",
				disabled: !this.properties.autoShowOnDelay,

			}),
			PropertyPaneTextField('openButtonCaption', {
			label: "Open button caption",
			description: "Text to display on the button that opens the popup",
			value: this.properties.openButtonCaption,
			disabled: this.properties.autoShowOnDelay,
			})
		);

		// *****************************************************
		// POSITION AND DIMENSIONS
		// *****************************************************

		const positionGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('hPos', {
				label: "Horizontal position",
				options: [
					{key: 1, text: "Left"},
					{key: 2, text:"Center"},
					{key: 3, text:"Right"}
				],
				selectedKey: this.properties.hPos
			}),
			PropertyPaneDropdown('vPos', {
				label: "Vertical position",
				options: [
					{key: 1, text: "Top"},
					{key: 2, text:"Center"},
					{key: 3, text:"Bottom"}
				],
				selectedKey: this.properties.vPos
			})
		];

		const dimensionsGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('widthType', {
				label: "Width type",
				options: [
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.widthType
			}),
			PropertyPaneSlider('widthPercentage',{
				label: "Width (percentage)",
				min:10,
				max:100,
				value:this.properties.widthPercentage,
				showValue:true,
				step:5,
				disabled: this.properties.widthType == 2
			}),
			PropertyFieldNumber('widthPixels',{
				key: "widthPixels",
				label: "Width (pixels)",
				value: this.properties.widthPixels,
				minValue: 10,
				maxValue: 1000,
				disabled: this.properties.widthType == 1
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneDropdown('heightType', {
				label: "Height type",
				options: [
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.heightType
			}),
			PropertyPaneSlider('heightPercentage',{
				label: "Height (percentage)",
				min:10,
				max:90,
				value:this.properties.heightPercentage,
				showValue:true,
				step:5,
				disabled: this.properties.heightType == 2
			}),
			PropertyFieldNumber('heightPixels',{
				key: "heightPixels",
				label: "Height (pixels)",
				value: this.properties.heightPixels,
				minValue: 10,
				maxValue: 1000,
				disabled: this.properties.heightType == 1
			})
		];

		// *****************************************************
		// POPUP REDIRECT
		// *****************************************************

		const popupRedirectGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneCheckbox("redirectOnClose", {
				text: "Redirect on popup close?",
				checked: this.properties.redirectOnClose
			}),
			PropertyPaneTextField('redirectURL', {
				disabled: !this.properties.redirectOnClose,
				label: "Redirect URL",
				value: this.properties.redirectURL
			})
		];

		// *****************************************************
		// REMEMBER USER CLOSE ACTION
		// *****************************************************

		const rememberUserCloseActionGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneCheckbox("rememberUserCloseAction", {
				text: "Remember user close action?",
				checked: this.properties.rememberUserCloseAction
			}),
			PropertyPaneSlider('rememberUserCloseActionMaxN',{
				label: "Hide popup afer N close actions",
				min:1,
				max:10,
				value:this.properties.rememberUserCloseActionMaxN,
				showValue:true,
				step:1,
				disabled: !this.properties.rememberUserCloseAction
			})
		];

		// real popup allows definition of position
		return {
		pages: [
			{
			header: {
				description: "Please configure the web part"
			},
			displayGroupsAsAccordion: true,
			groups: [
				{
					groupName: "Basic configuration",
					isCollapsed: false,
					groupFields: basicGroupFields
				},
				{
					groupName: "Position",
					isCollapsed: true,
					groupFields: positionGroupFields
				},
				{
					groupName: "Dimensions",
					isCollapsed: true,
					groupFields: dimensionsGroupFields
				},
				{
					groupName: "Redirect on popup close",
					isCollapsed: true,
					groupFields: popupRedirectGroupFields
				},
				{
					groupName: "Remember user close action",
					isCollapsed: true,
					groupFields: rememberUserCloseActionGroupFields
				},
			]
			}
		]
		};
	}
}
