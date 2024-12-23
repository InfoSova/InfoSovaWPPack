import * as React from 'react';

// just for "X" to close banner
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { FontIcon } from '@fluentui/react/lib/Icon';

import styles from './SovaHeading.module.scss';
import type { ISovaHeadingProps } from './ISovaHeadingProps';

export default class SovaHeading extends React.Component<ISovaHeadingProps, any> {
	public constructor(props: ISovaHeadingProps, state: any) {
		super(props);

		this.state = state != null? {}: state;

		// required for FluentUI, to get the icon for the 'X' button
		initializeIcons();
	}

	public render(): React.ReactElement<ISovaHeadingProps> {

		// ***************************
		// FIX STYLE FOR BANNER
		// ***************************

		let sHeight = this.props.heightPixels + "px";

		let borderStyle:string = "";
		if (this.props.borderWidth == "") borderStyle = "none";
			else borderStyle = "solid " + this.props.borderWidth + " " + this.props.borderColor;

		let sBorderRadius = "" + this.props.borderRadius + "px";

		let headingDivStyle:any = {
			"boxSizing": "border-box",
			"borderRadius": sBorderRadius,
			"border": borderStyle,
			"background-color": this.props.backgroundColor,
			"color": this.props.color,
			"fontSize": this.props.fontSize,
			"width": "100%",
			"marginBottom": this.props.marginBottom
		};
		if (this.props.heightPixels > 0) headingDivStyle["height"] = sHeight;
		if (this.props.separatorStyle != "none") headingDivStyle["border-bottom"] = this.props.separatorStyle + " " + this.props.separatorWidth + " " + this.props.separatorColor;
		if (this.props.bold) headingDivStyle["font-weight"] = "bold";

		let headingDivContentStyle:any = {
			"paddingTop": (this.props.contentPaddingVertical + "px"),
			"paddingBottom": (this.props.contentPaddingVertical + "px"),
			"paddingLeft": (this.props.contentPaddingHorizontal + "px"),
			"paddingRight": (this.props.contentPaddingHorizontal + "px"),
			"width": "100%",
			"height": "100%",
			"position": "relative"
		};

		return (
			<div className={this.props.dropShadow?styles.hooDropShadow:""} style={headingDivStyle}>
				<div style={headingDivContentStyle}>
					{!this.props.redirectOnClick && !this.props.showIcon?<div>{this.props.textToDisplay}</div>:""}
					{!this.props.redirectOnClick && this.props.showIcon?
						<div style={{display:"flex", alignItems:"center"}}>
							<div style={{display:"flex"}}><FontIcon iconName={this.props.icon} style={{color: this.props.iconColor, fontSize: this.props.iconSize}} /></div>
							<div style={{paddingLeft:"8px",display:"inline-block"}}>{this.props.textToDisplay}</div>
						</div>
						:""
					}
					{this.props.redirectOnClick && !this.props.showIcon?<div><a href={this.props.redirectURL} style={{color: this.props.color, textDecoration: "none"}}
						className={(this.props.linkHoverUnderline?styles.hooLinkHoverUnderline:styles.hooLinkHoverUnderlineNO) + " " + (this.props.linkHoverBold?styles.hooLinkHoverBold:styles.hooLinkHoverBoldNO)}>{this.props.textToDisplay}</a></div>:""}
					{this.props.redirectOnClick && this.props.showIcon?
						<div style={{display:"flex", alignItems:"center"}}>
							<div style={{display:"flex"}}><FontIcon iconName={this.props.icon} style={{color: this.props.iconColor, fontSize: this.props.iconSize}} /></div>
							<div style={{paddingLeft:"8px",display:"inline-block"}}><a href={this.props.redirectURL} style={{color: this.props.color, textDecoration: "none"}}
								className={(this.props.linkHoverUnderline?styles.hooLinkHoverUnderline:styles.hooLinkHoverUnderlineNO) + " " + (this.props.linkHoverBold?styles.hooLinkHoverBold:styles.hooLinkHoverBoldNO)}>{this.props.textToDisplay}</a></div>
						</div>
						:""
					}
				</div>
			</div>
		);
	}
}
