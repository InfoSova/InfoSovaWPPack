export interface ISovaPopupProps {
	popupContentType: number, 	// HTML text or URL to file to display
	popupContentToDisplay: string,
	closeButtonCaption: string,

	autoShowOnDelay: boolean,
	delaySeconds: number,
	openButtonCaption: string,

	vPos: number,
	hPos: number,
	widthType: number,		// pixels or percentage
	widthPercentage: number,
	widthPixels: number,
	heightType: number,		// pixels or percentage
	heightPercentage: number,
	heightPixels: number,

	redirectOnClose: boolean,
	redirectURL: string,

	rememberUserCloseAction: boolean,
	rememberUserCloseActionMaxN: number,

	webPartId: any,
	context: any,
	isEditMode: boolean,
	isDarkTheme: boolean,
	environmentMessage: string,
	hasTeamsContext: boolean,
	userDisplayName: string
}
