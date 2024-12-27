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

import { PropertyFieldColorPicker} from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';

import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'SovaContentWebPartStrings';
import SovaContent from './components/SovaContent';
import { ISovaContentProps } from './components/ISovaContentProps';

export interface ISovaContentWebPartProps {
	contentType: number, 				//	0 = plain text (cleaned up), 1 = HTML text, 2 = URL to HTML file
	plainTextContent: string,
	HTMLContent: string,
	HTMLFilePickerResult: IFilePickerResult,

	widthType: number,						// pixels or percentage
	widthPercentage: number,
	widthPixels: number,
	heightPixels: number,

	contentPadding:number,
	marginTop:number,
	marginLeft:number,
	marginBottom:number,
	zIndex:number,
	shrinkWebPart:boolean,

	showIcon: boolean,
	icon: any,
	iconColor: string,
	iconSize: string,

	fontSize: string,
	color: string,
	backgroundColor: string,

	borderRadius: number,
	borderWidth: string,
	borderColor: string,
	dropShadow: boolean,

	xShow: boolean,
	xColor: string,

	rememberUserCloseAction: boolean,
	rememberUserCloseActionMaxN: number
}

export default class SovaContentWebPart extends BaseClientSideWebPart<ISovaContentWebPartProps> {

	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = '';
	private _isInEditMode: boolean = false;

	protected get isRenderAsync(): boolean {
		return true;
	}

	public async render(): Promise<void> {
		if(Environment.type == EnvironmentType.SharePoint){
			if(this.displayMode == DisplayMode.Edit) this._isInEditMode = true;
		}

		// IF CONTENT SHOULD BE READ FROM A FILE, READ THE CONTENT FROM THE FILE
		// To execute async calls, the render method is converted to async (note also the isRenderAsync method above)
		let sContentToDisplay = "";
		if (this.properties.contentType == 0) {
			const langCode = document.documentElement.lang || navigator.language;
			const formatter = new Intl.DateTimeFormat(langCode);
			const tempsDate = formatter.format(new Date());
			sContentToDisplay = escape(this.properties.plainTextContent.replace("%%USERNAME%%", this.context.pageContext.user.displayName).replace("%%CURRENTDATE%%", tempsDate));
		}
		if (this.properties.contentType == 1) {
			sContentToDisplay = this.properties.HTMLContent;
		}
		else if (this.properties.contentType == 2){	// load data from a standalone document
			if (this.properties.HTMLFilePickerResult && this.properties.HTMLFilePickerResult.fileAbsoluteUrl && this.properties.HTMLFilePickerResult.fileAbsoluteUrl != ""){
				let fileUrl = this.properties.HTMLFilePickerResult.fileAbsoluteUrl;
				try{
					let res:SPHttpClientResponse = await this.context.spHttpClient.get(fileUrl, SPHttpClient.configurations.v1);
					if (res.ok){
						sContentToDisplay = await res.text();
					}else {
						console.log("WEBPART ERROR: Error loading content to display from the file URL - " + fileUrl);
						sContentToDisplay = "";
					}
				}catch{
					console.log("WEBPART ERROR: General error notification while access file URL - " + fileUrl);
				}
			}
		}

		const wpProperties:ISovaContentProps = {
			contentType: this.properties.contentType,
			contentToDisplay: sContentToDisplay,

			widthType: this.properties.widthType,
			widthPercentage: this.properties.widthPercentage,
			widthPixels: this.properties.widthPixels,
			heightPixels: this.properties.heightPixels,

			contentPadding:this.properties.contentPadding,
			marginTop: this.properties.marginTop,
			marginLeft:this.properties.marginLeft,
			marginBottom: this.properties.marginBottom,
			zIndex:this.properties.zIndex,
			shrinkWebPart:this.properties.shrinkWebPart,

			borderRadius: this.properties.borderRadius,
			borderWidth: this.properties.borderWidth,
			borderColor: this.properties.borderColor,
			backgroundColor: this.properties.backgroundColor,

			showIcon: this.properties.showIcon,
			iconName: this.properties.icon,
			iconColor: this.properties.iconColor,
			iconSize: this.properties.iconSize,

			fontSize: this.properties.fontSize,
			color:this.properties.color,
			dropShadow: this.properties.dropShadow,

			xShow: this.properties.xShow,
			xColor: this.properties.xColor,

			rememberUserCloseAction: this.properties.rememberUserCloseAction,
			rememberUserCloseActionMaxN: this.properties.rememberUserCloseActionMaxN,

			webPartId: this.context.instanceId,
			context:this.context,
			isEditMode: this._isInEditMode,
			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName,
			domElement: this.domElement
		};

		const element: React.ReactElement<any, any> = React.createElement(SovaContent, wpProperties, null);
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
		// CONTENT
		// *****************************************************

		const contentGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('contentType', {
				label: "Type of content to display",
				options: [
					{key: 0, text: "Plain text"},
					{key: 1, text: "HTML text"},
					{key: 2, text: "URL to HTML file"}
				],
				selectedKey: this.properties.contentType
			})
		];
		if (this.properties.contentType == 0){
			contentGroupFields.push(
				PropertyPaneTextField('plainTextContent', {
					label: "Text for banner",
					multiline: true,
					rows:8,
					value:this.properties.plainTextContent,
					description: "You can use %%USERNAME%% int text to display current user's display name. User %%CURENTDATE%% to display date in user's locale."
				}));

			contentGroupFields.push(PropertyPaneCheckbox("showIcon", {
				text: "Show icon",
				checked: this.properties.showIcon
			}));

			if (this.properties.showIcon){
				contentGroupFields.push(
					PropertyFieldIconPicker('icon', {
						currentIcon: this.properties.icon,
						key: "iconPickerId",
						onSave: (icon: string) => { this.properties.icon = icon; },
						onChanged:(icon: string) => {  },
						buttonLabel: "Select icon",
						renderOption: "panel",
						properties: this.properties,
						onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
						label: "Icon Picker",
						disabled: !this.properties.showIcon
					  }),
					  PropertyFieldColorPicker("iconColor", {
						label: "Icon color",
						selectedColor: this.properties.iconColor,
						onPropertyChange: this.onPropertyPaneFieldChanged,
						properties: this.properties,
						showPreview: true,
						key: "basicUsage"
					}),
					PropertyPaneDropdown('iconSize', {
						label: "Icon size",
						options: [
							{key: "0.8em", text: "Smaller"},
							{key: "1em", text: "Normal"},
							{key: "1.4em", text: "Larger"}
						],
						selectedKey: this.properties.iconSize
					})
				);
			}
		}
		if (this.properties.contentType == 1){
			contentGroupFields.push(
				PropertyPaneTextField('HTMLContent', {
					label: "HTML content (IFRAME)",
					multiline: true,
					rows:8,
					value: this.properties.HTMLContent
				}));
		}
		if (this.properties.contentType == 2){
			contentGroupFields.push(
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
					label: "File with HTML content (IFRAME)"
				})
			);
		}

		// *****************************************************
		// COLORS AND STYLES FOR CLOSABLE BANNER
		// *****************************************************

		const styleGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('fontSize', {
				label: "Font size",
				options: [
					{key: "0.8em", text: "Smaller"},
					{key: "1em", text: "Normal"},
					{key: "1.4em", text: "Larger"}
				],
				selectedKey: this.properties.fontSize
			}),
			PropertyFieldColorPicker("color", {
				label: "Foreground color",
				selectedColor: this.properties.color,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				key: "basicUsage"
			}),
			PropertyFieldColorPicker("backgroundColor", {
				label: "Background color",
				selectedColor: this.properties.backgroundColor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				key: "basicUsage"
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneDropdown('borderWidth', {
				label: "Border width",
				options: [
					{key: "", text: "None"},
					{key: "1px", text: "1px"},
					{key: "2px", text: "2px"},
					{key: "3px", text: "3px"},
					{key: "4px", text: "4px"},
					{key: "5px", text: "5px"}
				],
				selectedKey: this.properties.borderWidth
			}),
			PropertyPaneSlider('borderRadius',{
				label: "Border radius",
				min:0,
				max:200,
				value:this.properties.borderRadius,
				showValue:true,
				step:1
			}),
			PropertyFieldColorPicker("borderColor", {
				label: "Border color",
				selectedColor: this.properties.borderColor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				key: "basicUsage"
			}),
			PropertyPaneCheckbox("dropShadow", {
				text: "Drop shadow",
				checked: this.properties.dropShadow
			})
		];

		// *****************************************************
		// HEIGHT AND PADDING
		// *****************************************************
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
			PropertyFieldNumber('heightPixels',{
				key: "heightPixels",
				label: "Height (pixels)",
				value: this.properties.heightPixels,
				minValue: 0,
				maxValue: 1000
			}),
			PropertyPaneHorizontalRule(),
			PropertyFieldNumber('contentPadding',{
				key: "contentPadding",
				label: "Content padding",
				value: this.properties.contentPadding,
				minValue: 0,
				maxValue: 1000
			})
		];

		// *****************************************************
		// MARGINS
		// *****************************************************
		const marginsGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyFieldNumber('marginTop',{
				key: "marginTop",
				label: "Margin top",
				value: this.properties.marginTop,
				minValue: -1000,
				maxValue: 1000
			}),
			PropertyFieldNumber('marginLeft',{
				key: "marginLeft",
				label: "Margin left",
				value: this.properties.marginLeft,
				minValue: -1000,
				maxValue: 1000
			}),
			PropertyFieldNumber('marginBottom',{
				key: "marginBottom",
				label: "Margin bottom",
				value: this.properties.marginBottom,
				minValue: -1000,
				maxValue: 1000
			}),
			PropertyFieldNumber('zIndex',{
				key: "zIndex",
				label: "Z-Index",
				value: this.properties.zIndex,
				minValue: -1000,
				maxValue: 1000,
				description: "Keep 0 in most cases"
			}),
			PropertyPaneCheckbox("shrinkWebPart", {
				text: "Shrink web part?",
				checked: this.properties.shrinkWebPart
			})
		];

		// *****************************************************
		// REMEMBER USER CLOSE ACTION
		// *****************************************************

		const rememberUserCloseActionGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneCheckbox("xShow", {
				text: "Show close button 'X'",
				checked: this.properties.xShow
			}),
			PropertyFieldColorPicker("xColor", {
				label: "Color for closing X",
				selectedColor: this.properties.xColor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				disabled: !this.properties.xShow,
				key: "basicUsage"
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneCheckbox("rememberUserCloseAction", {
				text: "Remember user close action?",
				checked: this.properties.rememberUserCloseAction
			}),
			PropertyPaneSlider('rememberUserCloseActionMaxN',{
				label: "Hide afer N close actions",
				min:1,
				max:5,
				value:this.properties.rememberUserCloseActionMaxN,
				showValue:true,
				step:1,
				disabled: !this.properties.rememberUserCloseAction
			})
		];

		return {
			pages: [
				{
				header: {
					description: "Please configure the web part"
				},
				displayGroupsAsAccordion: true,
					groups: [
					{
						groupName: "Content",
						isCollapsed: true,
						groupFields: contentGroupFields
					},
					{
						groupName: "Banner styles",
						isCollapsed: true,
						groupFields: styleGroupFields
					},
					{
						groupName: "Dimensions",
						isCollapsed: true,
						groupFields: dimensionsGroupFields
					},
					{
						groupName: "Margins",
						isCollapsed: true,
						groupFields: marginsGroupFields
					},
					{
						groupName: "Close action",
						isCollapsed: true,
						groupFields: rememberUserCloseActionGroupFields
					}
				]
				}
			]
		};
	}
}
