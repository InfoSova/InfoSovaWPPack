import * as React from 'react';
import styles from './SovaPageHeader.module.scss';
import type { ISovaPageHeaderProps } from './ISovaPageHeaderProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SovaPageHeader extends React.Component<ISovaPageHeaderProps, any> {
	public render(): React.ReactElement<ISovaPageHeaderProps> {

		let imageStyle:any = {};
		if (this.props.imageWidthType == 1) imageStyle["width"] = this.props.imageWidthPercentage + "%";
		if (this.props.imageWidthType == 2) imageStyle["width"] = this.props.imageWidthPixels + "px";
		if (this.props.imageHeightType == 1) imageStyle["height"] = this.props.imageHeightPercentage + "%";
		if (this.props.imageHeightType == 2) imageStyle["height"] = this.props.imageHeightPixels + "px";
		if (this.props.imageMinWidthType == 1) imageStyle["min-width"] = this.props.imageMinWidthPercentage + "%";
		if (this.props.imageMinWidthType == 2) imageStyle["min-width"] = this.props.imageMinWidthPixels + "px";
		if (this.props.imageMinHeightType == 1) imageStyle["min-height"] = this.props.imageMinHeightPercentage + "%";
		if (this.props.imageMinHeightType == 2) imageStyle["min-height"] = this.props.imageMinHeightPixels + "px";

		imageStyle["overflow"] = "hidden";
		imageStyle["position"] = "relative";

		// if image background
		if (this.props.backgroundType == 0) imageStyle["backgroundImage"] = "url('" + this.props.imageUrl + "') ";
		if (this.props.backgroundType == 0) imageStyle["backgroundRepeat"] = "no-repeat";
		if (this.props.backgroundType == 0) imageStyle["backgroundPosition"] = "center";
		if (this.props.backgroundType == 0) imageStyle["backgroundSize"] = "cover";

		if (this.props.backgroundType == 1) imageStyle["backgroundColor"] = this.props.backgroundColor;


		let containerStyle:any = {};
		// 32px is in MS web part
		containerStyle["paddingTop"] = this.props.paddingTop + "px";
		containerStyle["paddingLeft"] = this.props.paddingLeft + "px";
		containerStyle["paddingRight"] = this.props.paddingRight + "px";


		let overlayStyle:any = {};
		overlayStyle["backgroundColor"] = this.props.overlayBackgroundColor;
		overlayStyle["borderRadius"] = this.props.borderRadius;
		overlayStyle["position"] = "absolute";

		if (this.props.overlayPosition == 0) overlayStyle["left"] = "0";
		if (this.props.overlayPosition == 1) {
			overlayStyle["left"] = "50%";
			overlayStyle["transform"] = "translate(-50%, 0%)";
		}
		if (this.props.overlayPosition == 2) overlayStyle["right"] = "0";

		if (this.props.overlayWidthType == 1) overlayStyle["width"] = this.props.overlayWidthPercentage + "%";
		if (this.props.overlayWidthType == 2) overlayStyle["width"] = this.props.overlayWidthPixels + "px";
		if (this.props.overlayHeightType == 1) overlayStyle["height"] = this.props.overlayHeightPercentage + "%";
		if (this.props.overlayHeightType == 2) overlayStyle["height"] = this.props.overlayHeightPixels + "px";

		return (
			<div style={imageStyle}>
				<div style={containerStyle}>
					<div style={{width:"100%",position:"relative"}} className={(this.props.withVerticalSection?styles.overLayPositionWithVerticalSection:styles.overLayPosition)}>
						<div style={overlayStyle}>
							<iframe sandbox='allow-scripts allow-top-navigation-by-user-activation' srcDoc={this.props.HTMLContentToDisplay} className={styles.hooIFrame} />
						</div>
					</div>
				</div>
			</div>
		);
	}
}
