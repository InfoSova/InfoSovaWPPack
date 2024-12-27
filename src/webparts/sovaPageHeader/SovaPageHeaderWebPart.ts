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
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { PropertyFieldColorPicker} from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

import { Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'SovaPageHeaderWebPartStrings';
import SovaPageHeader from './components/SovaPageHeader';
import { ISovaPageHeaderProps } from './components/ISovaPageHeaderProps';

export interface ISovaPageHeaderWebPartProps {
	ImageFilePickerResult: IFilePickerResult,

	backgroundType:number,
	withVerticalSection: boolean,

	overlayPosition: number,

	imageWidthType: number,
	imageWidthPercentage: number,
	imageWidthPixels: number,
	imageHeightType: number,
	imageHeightPercentage: number,
	imageHeightPixels: number,

	imageMinWidthType: number,
	imageMinWidthPercentage: number,
	imageMinWidthPixels: number,
	imageMinHeightType: number,
	imageMinHeightPercentage: number,
	imageMinHeightPixels: number,

	paddingTop: number,
	paddingLeft: number,
	paddingRight: number,

	overlayWidthType: number,
	overlayWidthPixels: number,
	overlayWidthPercentage: number,
	overlayHeightType: number,
	overlayHeightPixels: number,
	overlayHeightPercentage: number,

	overlayMinWidthType: number,
	overlayMinWidthPixels: number,
	overlayMinWidthPercentage: number,
	overlayMinHeightType: number,
	overlayMinHeightPixels: number,
	overlayMinHeightPercentage: number,

	backgroundColor: string,

	overlayBackgroundColor: string,
	borderRadius: number,

	HTMLContentToDisplay: string,

	webPartId: any,
	context: any,
	isEditMode: boolean,

	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
	domElement: any
}

export default class SovaPageHeaderWebPart extends BaseClientSideWebPart<ISovaPageHeaderWebPartProps> {

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

		// get page title from page list item properties
		let sHTMLContentToDisplay = this.properties.HTMLContentToDisplay;
		let currentWebUrl = this.context.pageContext.web.absoluteUrl;
		if (this.context.pageContext.list && this.context.pageContext.list.title && this.context.pageContext.list.title != "" && this.context.pageContext.listItem){
			let requestUrl = currentWebUrl.concat("/_api/web/Lists/GetByTitle('" + this.context.pageContext.list?.title + "')/items(" + this.context.pageContext.listItem?.id + ")");
			try{
				let res:SPHttpClientResponse = await this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
				if (res.ok){
					let tempResponse = await res.json();
					sHTMLContentToDisplay = sHTMLContentToDisplay.replace("%%PAGETITLE%%", tempResponse.Title)
				}else {
					console.log("WEBPART ERROR: Error loading current page information");
				}
			}catch{
				console.log("WEBPART ERROR: General error notification while loading current page information");
			}
		}else console.log("WEBPART ERROR: Page elements are undefined");

		const wpProperties:ISovaPageHeaderProps = {
			imageUrl: (this.properties.ImageFilePickerResult && this.properties.ImageFilePickerResult.fileAbsoluteUrl)?this.properties.ImageFilePickerResult.fileAbsoluteUrl:"",

			backgroundType: this.properties.backgroundType,
			withVerticalSection: this.properties.withVerticalSection,

			overlayPosition: this.properties.overlayPosition,

			imageWidthType: this.properties.imageWidthType,
			imageWidthPixels: this.properties.imageWidthPixels,
			imageWidthPercentage: this.properties.imageWidthPercentage,
			imageHeightType: this.properties.imageHeightType,
			imageHeightPixels: this.properties.imageHeightPixels,
			imageHeightPercentage: this.properties.imageHeightPercentage,

			imageMinWidthType: this.properties.imageMinWidthType,
			imageMinWidthPixels: this.properties.imageMinWidthPixels,
			imageMinWidthPercentage: this.properties.imageMinWidthPercentage,
			imageMinHeightType: this.properties.imageMinHeightType,
			imageMinHeightPixels: this.properties.imageMinHeightPixels,
			imageMinHeightPercentage: this.properties.imageMinHeightPercentage,

			paddingTop: this.properties.paddingTop,
			paddingLeft: this.properties.paddingLeft,
			paddingRight: this.properties.paddingRight,

			overlayWidthType: this.properties.overlayWidthType,
			overlayWidthPixels: this.properties.overlayWidthPixels,
			overlayWidthPercentage: this.properties.overlayWidthPercentage,
			overlayHeightType: this.properties.overlayHeightType,
			overlayHeightPixels: this.properties.overlayHeightPixels,
			overlayHeightPercentage: this.properties.overlayHeightPercentage,

			overlayMinWidthType: this.properties.overlayMinWidthType,
			overlayMinWidthPixels: this.properties.overlayMinWidthPixels,
			overlayMinWidthPercentage: this.properties.overlayMinWidthPercentage,
			overlayMinHeightType: this.properties.overlayMinHeightType,
			overlayMinHeightPixels: this.properties.overlayMinHeightPixels,
			overlayMinHeightPercentage: this.properties.overlayMinHeightPercentage,

			backgroundColor: this.properties.backgroundColor,

			overlayBackgroundColor: this.properties.overlayBackgroundColor,
			borderRadius: this.properties.borderRadius,

			HTMLContentToDisplay: sHTMLContentToDisplay,

			webPartId: this.context.instanceId,
			context: this.context,
			isEditMode: this._isInEditMode,

			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName,
			domElement: this.domElement
		};

		const element: React.ReactElement<ISovaPageHeaderProps> = React.createElement(SovaPageHeader, wpProperties, null);
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
		const basicGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneTextField('HTMLContentToDisplay', {
				label: "HTML content (IFRAME)",
				multiline: true,
				rows:16,
				value: this.properties.HTMLContentToDisplay
			}),
			PropertyPaneDropdown('overlayPosition', {
				label: "Overlay position",
				options: [
					{key: 0, text: "Left"},
					{key: 1, text: "Center"},
					{key: 2, text: "Right"}
				],
				selectedKey: this.properties.overlayPosition
			}),
			PropertyPaneCheckbox("withVerticalSection", {
				text: "With vertical section?",
				checked: this.properties.withVerticalSection
			}),
			PropertyPaneDropdown('backgroundType', {
				label: "Background type",
				options: [
					{key: 0, text: "Image"},
					{key: 1, text: "Color"}
				],
				selectedKey: this.properties.backgroundType
			})
		];

		// image background
		if (this.properties.backgroundType == 0){
			basicGroupFields.push(
				PropertyFieldFilePicker('filePicker', {
					context: this.context,
					accepts:[".jpg", ".png", ".jpeg", ".gif"],
					hideStockImages: false,
					hideLocalUploadTab: false,
					hideOrganisationalAssetTab: false,
					filePickerResult: this.properties.ImageFilePickerResult,
					onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
					properties: this.properties,
					onSave: (e: IFilePickerResult) => {
						//console.log(e);
						this.properties.ImageFilePickerResult = e;  },
					onChanged: (e: IFilePickerResult) => {
						//console.log(e);
						this.properties.ImageFilePickerResult = e; },
					key: "filePickerId",
					buttonLabel: "Choose image",
					label: "Image to display"
				})
			);

		}
		// image bg color
		if (this.properties.backgroundType == 1){
			basicGroupFields.push(
				PropertyFieldColorPicker("backgroundColor", {
					label: "Header background color",
					selectedColor: this.properties.backgroundColor,
					onPropertyChange: this.onPropertyPaneFieldChanged,
					properties: this.properties,
					showPreview: true,
					key: "basicUsage"
				})
			);
		}

		/****************************************************************/
		/* IMAGE DIMENSIONS  											*/
		/****************************************************************/

		const imageDimGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('imageWidthType', {
				label: "Image width type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.imageWidthType
			})];
		if (this.properties.imageWidthType == 1) imageDimGroupFields.push(
			PropertyPaneSlider('imageWidthPercentage',{
				label: "Width (percentage)",
				min:10,
				max:100,
				value:this.properties.imageWidthPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.imageWidthType == 2) imageDimGroupFields.push(
			PropertyFieldNumber('imageWidthPixels',{
				key: "imageWidthPixels",
				label: "Width (pixels)",
				value: this.properties.imageWidthPixels,
				minValue: 10,
				maxValue: 1000
			}));

		imageDimGroupFields.push(
			PropertyPaneDropdown('imageHeightType', {
				label: "Image height type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.imageHeightType
			}));
		if (this.properties.imageHeightType == 1) imageDimGroupFields.push(
			PropertyPaneSlider('imageHeightPercentage',{
				label: "Height (percentage)",
				min:10,
				max:100,
				value:this.properties.imageHeightPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.imageHeightType == 2) imageDimGroupFields.push(
			PropertyFieldNumber('imageHeightPixels',{
				key: "imageHeightPixels",
				label: "Height (pixels)",
				value: this.properties.imageHeightPixels,
				minValue: 10,
				maxValue: 1000
			}));


		imageDimGroupFields.push(
			PropertyPaneDropdown('imageMinWidthType', {
				label: "Image min-width type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.imageMinWidthType
			}));
		if (this.properties.imageMinWidthType == 1) imageDimGroupFields.push(
			PropertyPaneSlider('imageMinWidthPercentage',{
				label: "Min Width (percentage)",
				min:10,
				max:100,
				value:this.properties.imageMinWidthPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.imageMinWidthType == 2) imageDimGroupFields.push(
			PropertyFieldNumber('imageMinWidthPixels',{
				key: "imageMinWidthPixels",
				label: "Min Width (pixels)",
				value: this.properties.imageMinWidthPixels,
				minValue: 10,
				maxValue: 1000
			}));

		imageDimGroupFields.push(
			PropertyPaneDropdown('imageMinHeightType', {
				label: "Image min-height type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.imageMinHeightType
			}));
		if (this.properties.imageMinHeightType == 1) imageDimGroupFields.push(
			PropertyPaneSlider('imageMinHeightPercentage',{
				label: "Min Height (percentage)",
				min:10,
				max:100,
				value:this.properties.imageMinHeightPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.imageMinHeightType == 2) imageDimGroupFields.push(
			PropertyFieldNumber('imageMinHeightPixels',{
				key: "imageMinHeightPixels",
				label: "Min Height (pixels)",
				value: this.properties.imageMinHeightPixels,
				minValue: 10,
				maxValue: 1000
			}));

		/****************************************************************/
		/* PADDING 			  											*/
		/****************************************************************/
		const paddingGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneSlider('paddingTop',{
				label: "Padding top",
				min:0,
				max:50,
				value:this.properties.paddingTop,
				showValue:true,
				step:1
			}),
			PropertyPaneSlider('paddingLeft',{
				label: "Padding left (default = 32)",
				min:0,
				max:50,
				value:this.properties.paddingLeft,
				showValue:true,
				step:1
			}),
			PropertyPaneSlider('paddingRight',{
				label: "Padding right (default = 32)",
				min:0,
				max:50,
				value:this.properties.paddingRight,
				showValue:true,
				step:1
			})
		];

		/****************************************************************/
		/* OVERLAY DIMENSIONS  											*/
		/****************************************************************/

		const overlayDimGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('overlayWidthType', {
				label: "Overlay width type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayWidthType
			})];
		if (this.properties.overlayWidthType == 1) overlayDimGroupFields.push(
			PropertyPaneSlider('overlayWidthPercentage',{
				label: "Width (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayWidthPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayWidthType == 2) overlayDimGroupFields.push(
			PropertyFieldNumber('overlayWidthPixels',{
				key: "overlayWidthPixels",
				label: "Width (pixels)",
				value: this.properties.overlayWidthPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayDimGroupFields.push(
			PropertyPaneDropdown('overlayHeightType', {
				label: "Overlay height type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayHeightType
			}));
		if (this.properties.overlayHeightType == 1) overlayDimGroupFields.push(
			PropertyPaneSlider('overlayHeightPercentage',{
				label: "Height (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayHeightPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayHeightType == 2) overlayDimGroupFields.push(
			PropertyFieldNumber('overlayHeightPixels',{
				key: "overlayHeightPixels",
				label: "Height (pixels)",
				value: this.properties.overlayHeightPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayDimGroupFields.push(
			PropertyPaneDropdown('overlayMinWidthType', {
				label: "Overlay min-width type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayMinWidthType
			}));
		if (this.properties.overlayMinWidthType == 1) overlayDimGroupFields.push(
			PropertyPaneSlider('overlayMinWidthPercentage',{
				label: "min-width (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayMinWidthPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayMinWidthType == 2) overlayDimGroupFields.push(
			PropertyFieldNumber('overlayMinWidthPixels',{
				key: "overlayMinWidthPixels",
				label: "min-width (pixels)",
				value: this.properties.overlayMinWidthPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayDimGroupFields.push(
			PropertyPaneDropdown('overlayMinHeightType', {
				label: "Overlay min-height type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayMinHeightType
			}));
		if (this.properties.overlayMinHeightType == 1) overlayDimGroupFields.push(
			PropertyPaneSlider('overlayMinHeightPercentage',{
				label: "min-height (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayMinHeightPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayMinHeightType == 2) overlayDimGroupFields.push(
			PropertyFieldNumber('overlayMinHeightPixels',{
				key: "overlayMinHeightPixels",
				label: "min-height (pixels)",
				value: this.properties.overlayMinHeightPixels,
				minValue: 10,
				maxValue: 1000
			}));

		/****************************************************************/
		/* OVERLAY STYLE  												*/
		/****************************************************************/
		const overlayStyleGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyFieldColorPicker("overlayBackgroundColor", {
				label: "Overlay background color",
				selectedColor: this.properties.overlayBackgroundColor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				key: "basicUsage"
			}),
			PropertyPaneSlider('borderRadius',{
				label: "Border radius",
				min:0,
				max:200,
				value:this.properties.borderRadius,
				showValue:true,
				step:1
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
						groupName: "Basic settings",
						isCollapsed: true,
						groupFields: basicGroupFields
					},
					{
						groupName: "Image dimensions",
						isCollapsed: true,
						groupFields: imageDimGroupFields
					},
					{
						groupName: "Padding",
						isCollapsed: true,
						groupFields: paddingGroupFields
					},
					{
						groupName: "Overlay dimensions",
						isCollapsed: true,
						groupFields: overlayDimGroupFields
					},
					{
						groupName: "Overlay style",
						isCollapsed: true,
						groupFields: overlayStyleGroupFields
					}
				]
			}
			]
		};
	}
}
