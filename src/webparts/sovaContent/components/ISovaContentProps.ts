export interface ISovaContentProps {
	contentType: number, 		//	0 = plain text (cleaned up), 1 = HTML text, 2 = URL to HTML file
	contentToDisplay: string,

	widthType: number,						// pixels or percentage
	widthPercentage: number,
	widthPixels: number,
	heightPixels: number,

	contentPadding:number,
	marginTop:number,
	marginLeft:number,
	marginBottom:number,
	zIndex:number,
	shrinkWebPart:boolean,

	showIcon:boolean,
	iconName: string,
	iconColor: string,
	iconSize: string,

	fontSize: string,
	color: string,
	backgroundColor: string,
	borderWidth: string,
	borderRadius: number,
	borderColor: string,
	dropShadow: boolean,

	xShow: boolean,
	xColor: string,

	rememberUserCloseAction: boolean,
	rememberUserCloseActionMaxN: number,

	webPartId: any,
	context: any,
	isEditMode: boolean,

	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
	domElement: any
}
