import * as React from 'react';

// just for "X" to close banner
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { IconButton } from '@fluentui/react/lib/Button';

import { DefaultButton, FocusTrapZone, Layer, Overlay, Popup } from '@fluentui/react';

// for storing user close popup or banner activities
import {PnPClientStorage, dateAdd} from "@pnp/common";

import styles from './SovaPopup.module.scss';
import type { ISovaPopupProps } from './ISovaPopupProps';

export default class SovaPopup extends React.Component<ISovaPopupProps, any> {
	private PnPStorage:PnPClientStorage;
	private PnPStorageIdentifier = this.props.webPartId + "-NumberOfClosePopupActions";

	public constructor(props: ISovaPopupProps, state: any) {
		super(props);

		this.showPopup = this.showPopup.bind(this);
		this.hidePopup = this.hidePopup.bind(this);

		this.state = state != null? {isPopupVisible: false}: state;

		// to store information about the user closing the popup
		this.PnPStorage = new PnPClientStorage();

		// required for FluentUI, to get the icon for the 'X' button
		initializeIcons();
	}

	public async componentDidMount() {
		// SET UP THE TIMED POPUP
		if (this.props.autoShowOnDelay){
			if ((this.props.delaySeconds != null) && (this.props.delaySeconds !== 0)){
				if(this.props.delaySeconds > 0) await new Promise(resolve => setTimeout(resolve, this.props.delaySeconds*1000));
				this.showPopup();
			}
			else this.showPopup();
		}
	}

	public hidePopup(){
		this.setState({isPopupVisible: false});
		if (this.props.rememberUserCloseAction) this.incrementNumberForUserInStorage();
		if (this.props.redirectOnClose && this.props.redirectURL.trim() != "") window.location.href = this.props.redirectURL;
	}

	public showPopup(){
		this.setState({isPopupVisible: true});
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
			this.PnPStorage.local.put(this.PnPStorageIdentifier, String(numValue), dateAdd(new Date(), 'year', 100));
		}else{
			// set the initial value for user close popup actions
			this.PnPStorage.local.put(this.PnPStorageIdentifier, String(1), dateAdd(new Date(), 'year', 100));
		}
	}

	public render(): React.ReactElement<ISovaPopupProps> {

		if (this.props.rememberUserCloseAction){
			// if users closed popup enough times, don't display popup or banner
			if (this.getNumberForUserFromStorage() >=  this.props.rememberUserCloseActionMaxN) return (<></>);
		}

		let showCloseButton = this.props.closeButtonCaption && this.props.closeButtonCaption != "";

		let sWidth = this.props.widthPercentage + "%";
		if (this.props.widthType == 2) sWidth = this.props.widthPixels + "px";
		let sHeight = this.props.heightPercentage + "%";
		if (this.props.heightType == 2) sHeight = this.props.heightPixels + "px";

		// ***************************
		// FIX DISPLAY FOR POPUP
		// ***************************
		let popupAddOnStyles = {};
		if (this.props.vPos == 1 && this.props.hPos == 1) popupAddOnStyles = {
			left: '5%', top: '5%', transform: 'translate(-5%, -5%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 1 && this.props.hPos == 2) popupAddOnStyles = {
			left: '50%', top: '5%', transform: 'translate(-50%, -5%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 1 && this.props.hPos == 3) popupAddOnStyles = {
			left: '95%', top: '5%', transform: 'translate(-95%, -5%)', width: sWidth, height: sHeight
			};

		if (this.props.vPos == 2 && this.props.hPos == 1) popupAddOnStyles = {
			left: '5%', top: '50%', transform: 'translate(-5%, -50%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 2 && this.props.hPos == 2) popupAddOnStyles = {
			left: '50%', top: '50%', transform: 'translate(-50%, -50%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 2 && this.props.hPos == 3) popupAddOnStyles = {
			left: '95%', top: '50%', transform: 'translate(-95%, -50%)', width: sWidth, height: sHeight
			};

		if (this.props.vPos == 3 && this.props.hPos == 1) popupAddOnStyles = {
			left: '5%', top: '95%', transform: 'translate(-5%, -95%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 3 && this.props.hPos == 2) popupAddOnStyles = {
			left: '50%', top: '95%', transform: 'translate(-50%, -95%)', width: sWidth, height: sHeight
			};
		if (this.props.vPos == 3 && this.props.hPos == 3) popupAddOnStyles = {
			left: '95%', top: '95%', transform: 'translate(-95%, -95%)', width: sWidth, height: sHeight
			};

		return (
			<>
				{!this.props.autoShowOnDelay?<DefaultButton onClick={this.showPopup} text={this.props.openButtonCaption} />:""}
				{this.state.isPopupVisible?<Layer>
						<Popup className={styles.hooPopup} role="dialog" aria-modal="true" onDismiss={this.hidePopup}>
							<Overlay isDarkThemed={true} onClick={this.hidePopup} />
							<FocusTrapZone>
								<div role="document" style={popupAddOnStyles} className={styles.hooPopupBase}>
									<div>
										<div className={styles.hooPopupContentContainer} style={{bottom: (showCloseButton?"45px":"10px")}}>
											<iframe sandbox='allow-scripts allow-top-navigation-by-user-activation' srcDoc={this.props.popupContentToDisplay} className={styles.hooPopupIFrame} />
										</div>
										{
											showCloseButton?
											<div className={styles.hooPopupButtonContainer} >
												<DefaultButton onClick={this.hidePopup} className={styles.hooCloseButtonStyle}>{this.props.closeButtonCaption}</DefaultButton>
											</div>: ""
										}
									</div>
									{!showCloseButton?
									<IconButton iconProps={{iconName: "ChromeClose"}} title="Close popup" ariaLabel="Close popup" styles={{
											icon: {color: "#666", fontSize: 16},
											root: {
												width: 20,
												height: 20,
												position: 'absolute',
												top: "-24px",
												right: "0px",
												backgroundColor: "#fff",
												color: "#ccc"
												//left: '98%',
												//top: '9%',
												//transform: 'translate(-98%, -9%)'*/
											},
											rootHovered: {backgroundColor: "#eee"},
											rootPressed: {backgroundColor: "#ddd"}
											}} onClick={this.hidePopup} />: ""
									}
								</div>
							</FocusTrapZone>
						</Popup>
					</Layer>:""}
			</>
		);
	}
}
