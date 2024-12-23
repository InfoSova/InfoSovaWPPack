import * as React from 'react';

// for icons and closing 'X'
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { FontIcon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';

// for storing user close activities
import {PnPClientStorage, dateAdd} from "@pnp/common";

import styles from './SovaContent.module.scss';
import type { ISovaContentProps } from './ISovaContentProps';

export default class SovaContent extends React.Component<ISovaContentProps, any> {
	private PnPStorage:PnPClientStorage;
	private PnPStorageIdentifier = this.props.webPartId + "-NumberOfCloseActions";

	public constructor(props: ISovaContentProps, state: any) {
		super(props);

		this.showBanner = this.showBanner.bind(this);
		this.hideBanner = this.hideBanner.bind(this);

		this.state = state != null? {isElementVisible: false}: state;

		// to store information about the user close actions
		this.PnPStorage = new PnPClientStorage();

		// required for FluentUI, to get the icon for the 'X' button
		initializeIcons();
	}

	public async componentDidMount() {
		this.showBanner();
	}

	public hideBanner(){
		this.setState({isElementVisible: false});
		if (this.props.rememberUserCloseAction) this.incrementNumberForUserInStorage();
	}

	public showBanner(){
		this.setState({isElementVisible: true});
	}

	public getNumberForUserFromStorage():number{
		let nForUser:string = this.PnPStorage.local.get(this.PnPStorageIdentifier);
        if (!isNaN(Number(nForUser))) return Number(nForUser);
		return 0;
	}

	public incrementNumberForUserInStorage(){
		let nForUser:string = this.PnPStorage.local.get(this.PnPStorageIdentifier);
        if (nForUser && !isNaN(Number(nForUser))) {
			let numValue:number = Number(nForUser);
			numValue++;
			this.PnPStorage.local.put(this.PnPStorageIdentifier, String(numValue), dateAdd(new Date(), 'year', 1));
		}else{
			// set the initial value for user close actions
			this.PnPStorage.local.put(this.PnPStorageIdentifier, String(1), dateAdd(new Date(), 'year', 1));
		}
	}

	public render(): React.ReactElement<ISovaContentProps> {

		if (this.props.rememberUserCloseAction){
			// if users closed the element enough times, don't display
			if (this.getNumberForUserFromStorage() >=  this.props.rememberUserCloseActionMaxN) return (<></>);
		}

		let showIFrame:boolean = this.props.contentType != 0;

		// ***************************
		// FIX STYLE FOR BANNER
		// ***************************

		let sHeight = this.props.heightPixels + "px";

		let borderStyle:string = "";
		if (this.props.borderWidth == "") borderStyle = "none";
			else borderStyle = "solid " + this.props.borderWidth + " " + this.props.borderColor;

		let sBorderRadius = "" + this.props.borderRadius + "px";

		let sWidth = this.props.widthPercentage + "%";
		if (this.props.widthType == 2) sWidth = this.props.widthPixels + "px";

		let bannerDivStyle:any = {
			"position": "relative",
			"boxSizing": "border-box",
			"borderRadius": sBorderRadius,
			"border": borderStyle,
			"background-color": this.props.backgroundColor,
			"color": this.props.color,
			"padding": (this.props.contentPadding + "px"),
			"fontSize": this.props.fontSize,
			"width": sWidth,
			"marginTop": this.props.marginTop,
			"marginLeft": this.props.marginLeft,
			"marginBottom": this.props.marginBottom
		};
		if (this.props.heightPixels > 0) bannerDivStyle["height"] = sHeight;
		if (this.props.zIndex > 0) bannerDivStyle["zIndex"] = this.props.zIndex;

		if (this.props.shrinkWebPart &&
			this.props.domElement.parentElement &&
			this.props.domElement.parentElement.parentElement &&
			this.props.domElement.parentElement.parentElement &&
			this.props.domElement.parentElement.parentElement.parentElement)
			{
				// set spacing on the web part container div
				this.props.domElement.parentElement.parentElement.parentElement.style.marginTop = "0px";
				this.props.domElement.parentElement.parentElement.parentElement.style.marginBottom = "0px";
				this.props.domElement.parentElement.parentElement.parentElement.style.padding = "1px";		// set padding to 1 to prevent margin collapse

				// remove top margin on next webparts's DIV container to take care of spacing problem
				if (this.props.domElement.parentElement.parentElement.parentElement.nextSibling){
					this.props.domElement.parentElement.parentElement.parentElement.nextSibling.style.marginTop = "0px";
				}
			}

		return (
			<>
				{this.props.isEditMode?<div>[SovaContentWebPart]</div>:""}
				{this.state.isElementVisible?
					<div className={this.props.dropShadow?styles.hooDropShadow:""} style={bannerDivStyle}>
						<div style={{width:"100%", height:"100%", position:"relative"}}>

							{showIFrame?
								<iframe sandbox='allow-scripts allow-top-navigation-by-user-activation' srcDoc={"<style>body {color:" + this.props.color + ";}</style>" + this.props.contentToDisplay} className={styles.hooIFrame} />:""
							}
							{!showIFrame && !this.props.showIcon?
								<div>{this.props.contentToDisplay}</div>:""
							}
							{!showIFrame && this.props.showIcon?
								<div style={{display:"flex", alignItems:"center"}}>
									<div style={{display:"flex"}}><FontIcon iconName={this.props.iconName} style={{color: this.props.iconColor, fontSize: this.props.iconSize}} /></div>
									<div style={{paddingLeft:"8px",display:"inline-block"}}>{this.props.contentToDisplay}</div>
								</div>
								:""
							}

							{this.props.xShow?
							<IconButton iconProps={{iconName: "ChromeClose"}} title="Close" ariaLabel="Close" styles={{
								icon: {color: this.props.xColor, fontWeight:"bold", fontSize: 12},
								root: {
									position: 'absolute',
									width: 12,
									height: 12,
									right: '12px',
									top: '50%',
									color: this.props.xColor,
									transform: 'translate(0%, -50%)'
								},
								rootHovered: {backgroundColor: this.props.backgroundColor},
								rootPressed: {backgroundColor: this.props.backgroundColor}
								}} onClick={this.hideBanner} />: ""
						}
						</div>
					</div>:""}
			</>
		);

	}
}
