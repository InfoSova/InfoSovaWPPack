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

import * as strings from 'SovaImageWebPartStrings';
import SovaImage from './components/SovaImage';
import { ISovaImageProps } from './components/ISovaImageProps';

export interface ISovaImageWebPartProps {
	ImageFilePickerResult: IFilePickerResult,

	positionType: number,

	withVerticalSection: boolean,

	imageWidthType: number,
	imageWidthPixels: number,
	imageWidthPercentage: number,
	imageHeightType: number,
	imageHeightPixels: number,
	imageHeightPercentage: number,

	imageMinWidthType: number,
	imageMinWidthPixels: number,
	imageMinWidthPercentage: number,
	imageMinHeightType: number,
	imageMinHeightPixels: number,
	imageMinHeightPercentage: number,


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


	overlayTopType: number,			// 0 = none, 1 = pixels, 2 = percentage
	overlayTopPixels: number,
	overlayTopPercentage: number,
	overlayLeftType: number,
	overlayLeftPixels: number,
	overlayLeftPercentage: number,
	overlayRightType: number,
	overlayRightPixels: number,
	overlayRightPercentage: number,
	overlayBottomType: number,
	overlayBottomPixels: number,
	overlayBottomPercentage: number,

	backgroundColor: string,
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

export default class SovaImageWebPart extends BaseClientSideWebPart<ISovaImageWebPartProps> {

	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = '';
	private _isInEditMode: boolean = false;

	public render(): void {
		if(Environment.type == EnvironmentType.SharePoint){
			if(this.displayMode == DisplayMode.Edit) this._isInEditMode = true;
		}

		const wpProperties:ISovaImageProps = {
			imageUrl: (this.properties.ImageFilePickerResult && this.properties.ImageFilePickerResult.fileAbsoluteUrl)?this.properties.ImageFilePickerResult.fileAbsoluteUrl:"",

			positionType: this.properties.positionType,

			withVerticalSection: this.properties.withVerticalSection,

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


			overlayTopType: this.properties.overlayTopType,			// 0 = none, 1 = pixels, 2 = percentage
			overlayTopPixels: this.properties.overlayTopPixels,
			overlayTopPercentage: this.properties.overlayTopPercentage,
			overlayLeftType: this.properties.overlayLeftType,
			overlayLeftPixels: this.properties.overlayLeftPixels,
			overlayLeftPercentage: this.properties.overlayLeftPercentage,
			overlayRightType: this.properties.overlayRightType,
			overlayRightPixels: this.properties.overlayRightPixels,
			overlayRightPercentage: this.properties.overlayRightPercentage,
			overlayBottomType: this.properties.overlayBottomType,
			overlayBottomPixels: this.properties.overlayBottomPixels,
			overlayBottomPercentage: this.properties.overlayBottomPercentage,

			backgroundColor: this.properties.backgroundColor,
			borderRadius: this.properties.borderRadius,

			HTMLContentToDisplay: this.properties.HTMLContentToDisplay,

			webPartId: this.context.instanceId,
			context: this.context,
			isEditMode: this._isInEditMode,

			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName,
			domElement: this.domElement
		};

		const element: React.ReactElement<ISovaImageProps> = React.createElement(SovaImage, wpProperties, null);
		ReactDom.render(element, this.domElement);
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
			}),
			PropertyPaneTextField('HTMLContentToDisplay', {
				label: "HTML content (IFRAME)",
				multiline: true,
				rows:16,
				value: this.properties.HTMLContentToDisplay
			}),
			PropertyPaneDropdown('positionType', {
				label: "Position type",
				options: [
					{key: 0, text: "Top header"},
					{key: 1, text: "Inside content"}
				],
				selectedKey: this.properties.positionType
			})
		];
		if (this.properties.positionType == 0){
			basicGroupFields.push(PropertyPaneCheckbox("withVerticalSection", {
				text: "With vertical section?",
				checked: this.properties.withVerticalSection
			}));
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
			PropertyFieldNumber('imageWidthPixels',{
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
		/* OVERLAY POSITION  											*/
		/****************************************************************/
		const overlayPosGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneDropdown('overlayTopType', {
				label: "Overlay top type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayTopType
			})];
		if (this.properties.overlayTopType == 1) overlayPosGroupFields.push(
			PropertyPaneSlider('overlayTopPercentage',{
				label: "Top (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayTopPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayTopType == 2) overlayPosGroupFields.push(
			PropertyFieldNumber('overlayTopPixels',{
				key: "overlayTopPixels",
				label: "Top  (pixels)",
				value: this.properties.overlayTopPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayPosGroupFields.push(
			PropertyPaneDropdown('overlayLeftType', {
				label: "Overlay left type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayLeftType
			}));
		if (this.properties.overlayLeftType == 1) overlayPosGroupFields.push(
			PropertyPaneSlider('overlayLeftPercentage',{
				label: "Left (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayLeftPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayLeftType == 2) overlayPosGroupFields.push(
			PropertyFieldNumber('overlayLeftPixels',{
				key: "overlayLeftPixels",
				label: "Left (pixels)",
				value: this.properties.overlayLeftPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayPosGroupFields.push(
			PropertyPaneDropdown('overlayRightType', {
				label: "Overlay right type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayRightType
			}));
		if (this.properties.overlayRightType == 1) overlayPosGroupFields.push(
			PropertyPaneSlider('overlayRightPercentage',{
				label: "Right (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayRightPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayRightType == 2) overlayPosGroupFields.push(
			PropertyFieldNumber('overlayRightPixels',{
				key: "overlayRightPixels",
				label: "Right (pixels)",
				value: this.properties.overlayRightPixels,
				minValue: 10,
				maxValue: 1000
			}));

		overlayPosGroupFields.push(
			PropertyPaneDropdown('overlayBottomType', {
				label: "Overlay bottom type",
				options: [
					{key: 0, text: "None"},
					{key: 1, text: "Percentage"},
					{key: 2, text: "Pixels"}
				],
				selectedKey: this.properties.overlayBottomType
			}));
		if (this.properties.overlayBottomType == 1) overlayPosGroupFields.push(
			PropertyPaneSlider('overlayBottomPercentage',{
				label: "Bottom (percentage)",
				min:10,
				max:100,
				value:this.properties.overlayBottomPercentage,
				showValue:true,
				step:5
			}));
		if (this.properties.overlayBottomType == 2) overlayPosGroupFields.push(
			PropertyFieldNumber('overlayBottomPixels',{
				key: "overlayBottomPixels",
				label: "Bottom (pixels)",
				value: this.properties.overlayBottomPixels,
				minValue: 10,
				maxValue: 1000
			}));

		/****************************************************************/
		/* OVERLAY STYLE  												*/
		/****************************************************************/
		const overlayStyleGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyFieldColorPicker("backgroundColor", {
				label: "Overlay background color",
				selectedColor: this.properties.backgroundColor,
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


		if (this.properties.positionType == 0){
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
		}else
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
							groupName: "Overlay dimensions",
							isCollapsed: true,
							groupFields: overlayDimGroupFields
						},
						{
							groupName: "Overlay position",
							isCollapsed: true,
							groupFields: overlayPosGroupFields
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
