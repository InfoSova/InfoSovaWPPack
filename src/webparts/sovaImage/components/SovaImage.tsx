import * as React from 'react';
import styles from './SovaImage.module.scss';
import type { ISovaImageProps } from './ISovaImageProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SovaImage extends React.Component<ISovaImageProps, any> {
	public render(): React.ReactElement<ISovaImageProps> {

		let imageStyle:any = {};
		if (this.props.imageWidthType == 1) imageStyle["width"] = this.props.imageWidthPercentage + "%";
		if (this.props.imageWidthType == 2) imageStyle["width"] = this.props.imageWidthPixels + "px";
		if (this.props.imageHeightType == 1) imageStyle["height"] = this.props.imageHeightPercentage + "%";
		if (this.props.imageHeightType == 2) imageStyle["height"] = this.props.imageHeightPixels + "px";
		if (this.props.imageMinWidthType == 1) imageStyle["min-width"] = this.props.imageMinWidthPercentage + "%";
		if (this.props.imageMinWidthType == 2) imageStyle["min-width"] = this.props.imageMinWidthPixels + "px";
		if (this.props.imageMinHeightType == 1) imageStyle["min-height"] = this.props.imageMinHeightPercentage + "%";
		if (this.props.imageMinHeightType == 2) imageStyle["min-height"] = this.props.imageMinHeightPixels + "px";

		// only if positioning is top header - imitate OOB functionality for banner, but with HTML content
		if (this.props.positionType == 0){
			imageStyle["overflow"] = "hidden";
			imageStyle["backgroundImage"] = "url('" + this.props.imageUrl + "') ";
			imageStyle["backgroundRepeat"] = "no-repeat";
			imageStyle["backgroundPosition"] = "center";
			imageStyle["backgroundSize"] = "cover";
			imageStyle["position"] = "relative";
			//imageStyle["paddingLeft"] = "30px";
			//imageStyle["paddingRight"] = "30px";
			//imageStyle["paddingTop"] = "30px";
		}

		let overlayStyle:any = {};
		if (this.props.overlayWidthType == 1) overlayStyle["width"] = this.props.overlayWidthPercentage + "%";
		if (this.props.overlayWidthType == 2) overlayStyle["width"] = this.props.overlayWidthPixels + "px";
		if (this.props.overlayHeightType == 1) overlayStyle["height"] = this.props.overlayHeightPercentage + "%";
		if (this.props.overlayHeightType == 2) overlayStyle["height"] = this.props.overlayHeightPixels + "px";

		// only if positioning is within content, absolute positioning of overlay
		if (this.props.positionType == 1){
			overlayStyle["position"] = "absolute";
			if (this.props.overlayTopType == 1) overlayStyle["top"] = this.props.overlayTopPercentage + "%";
			if (this.props.overlayTopType == 2) overlayStyle["top"] = this.props.overlayTopPixels + "px";
			if (this.props.overlayLeftType == 1) overlayStyle["left"] = this.props.overlayLeftPercentage + "%";
			if (this.props.overlayLeftType == 2) overlayStyle["left"] = this.props.overlayLeftPixels + "px";
			if (this.props.overlayRightType == 1) overlayStyle["right"] = this.props.overlayRightPercentage + "%";
			if (this.props.overlayRightType == 2) overlayStyle["right"] = this.props.overlayRightPixels + "px";
			if (this.props.overlayBottomType == 1) overlayStyle["bottom"] = this.props.overlayBottomPercentage + "%";
			if (this.props.overlayBottomType == 2) overlayStyle["bottom"] = this.props.overlayBottomPixels + "px";
		}

		overlayStyle["backgroundColor"] = this.props.backgroundColor;
		overlayStyle["borderRadius"] = this.props.borderRadius;

		if (this.props.positionType == 0){
			return (
				<div style={imageStyle}>
					<div style={{paddingLeft:"32px", paddingRight:"32px", paddingTop:"32px"}}>
						<div style={{width:"100%"}} className={(this.props.withVerticalSection?styles.overLayPositionWithVerticalSection:styles.overLayPosition)}>
							<div style={overlayStyle}>
								<iframe sandbox='allow-scripts allow-top-navigation-by-user-activation' srcDoc={this.props.HTMLContentToDisplay} className={styles.hooIFrame} />
							</div>
						</div>
					</div>
				</div>
			);
		}
		else
		{
			return (
			<div style={{position:"relative",overflow:"hidden"}}>
				<img src={this.props.imageUrl} style={imageStyle} />
				<div style={overlayStyle} >
					<iframe sandbox='allow-scripts allow-top-navigation-by-user-activation' srcDoc={this.props.HTMLContentToDisplay} className={styles.hooIFrame} />
				</div>
			</div>
			);
		}
	}
}
