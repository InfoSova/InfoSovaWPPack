export interface ISovaPageHeaderProps {
	imageUrl: string,

	backgroundType:number,
	withVerticalSection: boolean,

	overlayPosition: number,

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

	paddingTop: number,
	paddingLeft:number,
	paddingRight:number,

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
