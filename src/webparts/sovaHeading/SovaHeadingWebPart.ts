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
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';

import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'SovaHeadingWebPartStrings';
import SovaHeading from './components/SovaHeading';
import { ISovaHeadingProps } from './components/ISovaHeadingProps';

export interface ISovaHeadingWebPartProps {
	textToDisplay: string,

	heightPixels: number,

	contentPaddingVertical:number,
	contentPaddingHorizontal:number,
	marginBottom:number,

	showIcon: boolean,
	icon: any,
	iconColor: string,
	iconSize: string,

	fontSize: string,
	color: string,
	backgroundColor: string,

	bold: boolean,

	borderRadius: number,
	borderWidth: string,
	borderColor: string,

	dropShadow: boolean,

	separatorStyle: string,
	separatorWidth: string,
	separatorColor: string,

	redirectOnClick: boolean,
	redirectURL: string,
	linkHoverUnderline:boolean,
	linkHoverBold:boolean
}

export default class SovaHeadingWebPart extends BaseClientSideWebPart<ISovaHeadingWebPartProps> {

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

		const wpProperties:ISovaHeadingProps = {
			textToDisplay: escape(this.properties.textToDisplay),

			showIcon: this.properties.showIcon,
			icon: this.properties.icon,
			iconColor: this.properties.iconColor,
			iconSize: this.properties.iconSize,

			heightPixels: this.properties.heightPixels,
			contentPaddingVertical:this.properties.contentPaddingVertical,
			contentPaddingHorizontal:this.properties.contentPaddingHorizontal,
			marginBottom: this.properties.marginBottom,

			fontSize: this.properties.fontSize,
			color: this.properties.color,
			backgroundColor: this.properties.backgroundColor,

			bold: this.properties.bold,

			borderRadius: this.properties.borderRadius,
			borderWidth: this.properties.borderWidth,
			borderColor: this.properties.borderColor,

			dropShadow: this.properties.dropShadow,

			separatorStyle: this.properties.separatorStyle,
			separatorWidth: this.properties.separatorWidth,
			separatorColor: this.properties.separatorColor,

			redirectOnClick: this.properties.redirectOnClick,
			redirectURL: this.properties.redirectURL,
			linkHoverUnderline:this.properties.linkHoverUnderline,
			linkHoverBold:this.properties.linkHoverBold,

			webPartId: this.context.instanceId,
			context:this.context,
			isEditMode: this._isInEditMode,
			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName
		};

		const element: React.ReactElement<any, any> = React.createElement(SovaHeading, wpProperties, null);
		ReactDom.render(element, this.domElement);

		this.renderCompleted();
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

		if (!this.properties.textToDisplay) this.properties.textToDisplay = "Hello world";

		if (!this.properties.heightPixels) this.properties.heightPixels = 0;

		if (!this.properties.contentPaddingVertical) this.properties.contentPaddingVertical = 2;
		if (!this.properties.contentPaddingHorizontal) this.properties.contentPaddingVertical = 4;
		if (!this.properties.marginBottom) this.properties.marginBottom = 4;

		if (!this.properties.showIcon) this.properties.showIcon = false;
		if (!this.properties.icon) this.properties.icon = "";
		if (!this.properties.iconColor) this.properties.iconColor = "#000000";
		if (!this.properties.iconSize) this.properties.iconSize = "1em";

		if (!this.properties.fontSize) this.properties.fontSize = "1em";
		if (!this.properties.color) this.properties.color = "#000000";
		if (!this.properties.backgroundColor) this.properties.backgroundColor = "#ffffff";

		if (!this.properties.bold) this.properties.bold = false;

		if (!this.properties.borderRadius) this.properties.borderRadius = 2;
		if (!this.properties.borderWidth) this.properties.borderWidth = "";
		if (!this.properties.borderColor) this.properties.borderColor = "#666666";

		if (!this.properties.dropShadow) this.properties.dropShadow = false;

		if (!this.properties.separatorStyle) this.properties.separatorStyle = "solid";
		if (!this.properties.separatorWidth) this.properties.separatorWidth = "1px";
		if (!this.properties.separatorColor) this.properties.separatorWidth = "#000000";

		if (!this.properties.redirectOnClick) this.properties.redirectOnClick = false;
		if (!this.properties.redirectURL) this.properties.redirectURL = "";
		if (!this.properties.linkHoverUnderline) this.properties.linkHoverUnderline = false;
		if (!this.properties.linkHoverBold) this.properties.linkHoverBold = false;

		// *****************************************************
		// CONTENT
		// *****************************************************

		const contentGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneTextField('textToDisplay', {
				label: "Text for heading",
				value: this.properties.textToDisplay
			})
		];

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
					selectedKey: this.properties.fontSize
				})
			);
		}


		// *****************************************************
		// HEIGHT AND PADDING
		// *****************************************************
		const dimensionsGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyFieldNumber('heightPixels',{
				key: "heightPixels",
				label: "Height (pixels)",
				value: this.properties.heightPixels,
				minValue: 0,
				maxValue: 1000
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneSlider('contentPaddingVertical',{
				label: "Content padding vertical",
				min:0,
				max:20,
				value:this.properties.contentPaddingVertical,
				showValue:true,
				step:1
			}),
			PropertyPaneSlider('contentPaddingHorizontal',{
				label: "Content padding horizontal",
				min:0,
				max:20,
				value:this.properties.contentPaddingHorizontal,
				showValue:true,
				step:1
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneSlider('marginBottom',{
				label: "Bottom margin",
				min:-40,
				max:40,
				value:this.properties.marginBottom,
				showValue:true,
				step:1
			})
		];

		// *****************************************************
		// COLORS AND STYLES FOR HEADING
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
			PropertyPaneCheckbox("bold", {
				text: "Bold caption",
				checked: this.properties.bold
			}),
			PropertyPaneHorizontalRule(),
			PropertyPaneCheckbox("dropShadow", {
				text: "Drop shadow",
				checked: this.properties.dropShadow
			})
		];

		// *****************************************************
		// BORDER AND SEPARATOR
		// *****************************************************
		const borderAndSeparatorGroupFields: IPropertyPaneGroup["groupFields"] = [
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
				max:20,
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
			PropertyPaneDropdown('separatorStyle', {
				label: "Separator style",
				options: [
					{key: "none", text: "None"},
					{key: "solid", text: "Solid"},
					{key: "dashed", text: "Dashed"},
					{key: "dotted", text: "Dotted"}
				],
				selectedKey: this.properties.separatorStyle
			}),
			PropertyFieldColorPicker("separatorColor", {
				label: "Separator color",
				selectedColor: this.properties.separatorColor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				showPreview: true,
				key: "basicUsage"
			}),
			PropertyPaneDropdown('separatorWidth', {
				label: "Separator width",
				options: [
					{key: "", text: "None"},
					{key: "1px", text: "1px"},
					{key: "2px", text: "2px"},
					{key: "3px", text: "3px"},
					{key: "4px", text: "4px"},
					{key: "5px", text: "5px"}
				],
				selectedKey: this.properties.separatorWidth
			})
		];

		// *****************************************************
		// REDIRECT ON CLICK
		// *****************************************************

		const redirectGroupFields: IPropertyPaneGroup["groupFields"] = [
			PropertyPaneCheckbox("redirectOnClick", {
				text: "Redirect on heading click?",
				checked: this.properties.redirectOnClick
			}),
			PropertyPaneTextField('redirectURL', {
				disabled: !this.properties.redirectOnClick,
				label: "Redirect URL",
				value: this.properties.redirectURL
			}),
			PropertyPaneCheckbox("linkHoverUnderline", {
				text: "Link hover underline",
				checked: this.properties.linkHoverUnderline
			}),
			PropertyPaneCheckbox("linkHoverBold", {
				text: "Link hover bold",
				checked: this.properties.linkHoverBold
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
						groupName: "Heigh and padding",
						isCollapsed: true,
						groupFields: dimensionsGroupFields
					},
					{
						groupName: "Heading styles",
						isCollapsed: true,
						groupFields: styleGroupFields
					},
					{
						groupName: "Border and separator",
						isCollapsed: true,
						groupFields: borderAndSeparatorGroupFields
					},
					{
						groupName: "Redirect on click",
						isCollapsed: true,
						groupFields: redirectGroupFields
					}
				]
				}
			]
		};
	}
}
