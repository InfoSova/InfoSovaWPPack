export interface ISovaHeadingProps {
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
	linkHoverBold:boolean,

	webPartId: any,
	context: any,
	isEditMode: boolean,

	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
}
