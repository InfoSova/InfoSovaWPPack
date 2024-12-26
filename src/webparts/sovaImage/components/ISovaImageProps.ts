export interface ISovaImageProps {
	imageUrl: string,

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
