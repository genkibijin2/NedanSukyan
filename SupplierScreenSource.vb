'//--PUBLIC VARS=========================
Dim currentSupplierName As String
Dim currentSupplierAddress As String
Dim currentInstructions As String 'Stores Supplier guide text

Private Sub Form_Activate(ByVal bActivate As Boolean)
	SelectSupplierInfo
	supplierName.Text = currentSupplierName
	supplyInfo.Text = currentSupplierAddress
	Set loadingBox.Image = IconLoading
End Sub

Private Sub OutputSlowly(textOutput As String)
	Dim instructions As String 'text to be output
	Dim stackOfChars As String 'The output which is constructed by sticking chars together.
	Dim howManyCharsToLoop As Integer 'Length of the input which determines how many loops happen
	Dim looper As Integer
	instructions = textOutput
	howManyCharsToLoop = Len(instructions)
	For looper = 1 To howManyCharsToLoop
		stackOfChars = Left(instructions, looper)
		Sleep(10)
		supplyInfo.Text = stackOfChars
	Next looper
End Sub

Private Sub SelectSupplierInfo() 'Uses the companyID to check which supplier it should begin writing address info for
	currentInstructions = "No supplier guide to show, yet..." 'Fallback if there is no currentInstructions
	If selectedCompanyID = 1 Then
		currentSupplierName = "Avantek"
		currentSupplierAddress = "Avantek Machinery Ltd, AVA001 \nFlamstead House, Denby Hall Business Park, Denby, DE5 8JX"
		Set CompanyIconHeaderTop.Image = IconAvantek
	End If
	If selectedCompanyID = 2 Then
		currentSupplierName = "Bevelpane"
		currentSupplierAddress = "Bevelpane Panels BEV001 \nT/A Future Products, Units 4B & C Isabelle Court, Millennium Business Park, Mansfield, NG19 7JX"
		Set CompanyIconHeaderTop.Image = IconBevelpane
	End If
	If selectedCompanyID = 3 Then
		currentSupplierName = "Carl F"
		currentSupplierAddress = "Carl F Petersen, CAR002 \nCulley Court, Orton Southgate, Peterborough, Corby, PE2 6WA"
		Set CompanyIconHeaderTop.Image = IconCarlF
	End If
	If selectedCompanyID = 4 Then
		currentSupplierName = "Cotswold"
		currentSupplierAddress = "Cotswold Architectural Products, COT001 \nManor Park Industrial Estate, Manor Road, Cheltenham, Gloucestershire, GL51 9SQ"
		Set CompanyIconHeaderTop.Image = IconCotswold
	End If
	If selectedCompanyID = 5 Then
		currentSupplierName = "DGS"
		currentSupplierAddress = "DGS, DGS101 \nUnit 7 Techno Trading Estate, Station Road, Morley, Leeds, LS27 8JR"
		Set CompanyIconHeaderTop.Image = IconDGS
		currentInstructions = "1. Open 'Purchase Orders' in Eurostock\n2. Copy the 'Order Ref' code into the 'PO' box, and disable the use of dates\n3. Open up supplier in a seperate Eurostock instasnce and select DGS (DGS101)\n4. Cross reference the items in the purchase order with the prices in the supplier section\n\n-Example ref E-39465-\n\n\n\n"
	End If
	If selectedCompanyID = 6 Then
		currentSupplierName = "Distinction Doors"
		currentSupplierAddress = "Distinction Doors, DIS001 \n36 Wentworth Way, Wentworth Industrial Estate, Tankersley, South Yorkshire, S75 3DH"
		Set CompanyIconHeaderTop.Image = IconDistinctionDoors
	End If
	If selectedCompanyID = 7 Then
		currentSupplierName = "ERA (Fab &Fix)"
		currentSupplierAddress = "ERA Security Hardware, ERA001 \nValient Way, Wolverhampton, West Midlands, Willenhall, WV9 5GB"
		Set CompanyIconHeaderTop.Image = IconERA
	End If
	If selectedCompanyID = 8 Then
		currentSupplierName = "Exitex"
		currentSupplierAddress = "Exitex LTD, EXI001 \nDundalk, Ireland"
		Set CompanyIconHeaderTop.Image = IconExitex
	End If
	If selectedCompanyID = 9 Then
		currentSupplierName = "Liniar"
		currentSupplierAddress = "H.L. Plastics, LIN001 \nFlamstead House, Denby Hall Business Park, Denby, Derbyshire, DE5 8JX"
		Set CompanyIconHeaderTop.Image = IconLiniar
		currentInstructions = "1. Open up dashboard\n2. Navigate to 'Stock Control Tools'\n3. Locate your item code in the invoice - It should be under the 'Part code' header\n4. Place this code into the stock tools 'Supplier code' section to show price\n\n-Code example: LPD187WZ-\n\n\n\n\n"
	End If
	If selectedCompanyID = 10 Then
		currentSupplierName = "Maco"
		currentSupplierAddress = "Maco Door & Window Hardware (UK) Ltd, MAC001 \nEurolink Industrial Center, Castle Road, Sittingbourne, Kent ME10 3LY"
		Set CompanyIconHeaderTop.Image = IconMaco
	End If
	If selectedCompanyID = 11 Then
		currentSupplierName = "NW Metal"
		currentSupplierAddress = "NW Metal Sections, NW00001 \nUnit B, Moss Park Warehousing, Dairy Farm Road, Rainford, WA11 7RJ"
		Set CompanyIconHeaderTop.Image = IconNWMetal
	End If
	If selectedCompanyID = 12 Then
		currentSupplierName = "ODL"
		currentSupplierAddress = "ODL Europe, ODL001 \n1 Brook Road, Bootle, Merseyside, L20 4XP"
		Set CompanyIconHeaderTop.Image = IconODL
	End If
	If selectedCompanyID = 13 Then
		currentSupplierName = "Premier Arches"
		currentSupplierAddress = "Premier Arches Ltd PRE001 \nUnit 1a, Cochrane Street, Bolton, Greater Manchester, BL3 6BN"
		Set CompanyIconHeaderTop.Image = IconPremier
		currentInstructions = "1. Open Window Designer\n2. Find the 'Reference' number at the bottom of the premier invoice\n3. In Window Designer, click 'Job' At the top left, then 'Locate Job'\n4. Paste your reference into the 'Job Number' Field and click 'Load'\n5. Select the correct job listing and click 'Ok'\n6. Navigate to 'reports' at the top, and access either 'Glass reports', 'Customer reports - Job quotes with discount' etc, depending on the type of invoice\n\n\n\n\n"
	End If
	If selectedCompanyID = 14 Then
		currentSupplierName = "Rapidrop"
		currentSupplierAddress = "Rapidrop Global Lts T/A IFI, IFI001 \nUnits 1, 2 & 3, Rutland Business Park, Newark Road, Peterborough, Cambridgeshire, PE1 5WA"
		Set CompanyIconHeaderTop.Image = IconRapidrop
	End If
	If selectedCompanyID = 15 Then
		currentSupplierName = "Rehau"
		currentSupplierAddress = "Rehau, R0001 \nBrinel Drive, Irlam, Manchester, Lancashire, M44 5BL"
		Set CompanyIconHeaderTop.Image = IconRehau
		currentInstructions = "1. Navigate to the online 'REHAU CUSTOMER PORTAL' in your browser\n2. Log in with credentials\n3. Click 'Order System', then 'Price Lists'\n4. Copy the 'article number' from the invoice into the respective box on the search\n5. If the price is off, you can change it in Eurostock - the rehau portal is the correct price, always\n6. If the item doesn't exist on Eurostock, first check if it isn't a one-off purchase order using the Eurostock 'purchase orders'"
	End If
	If selectedCompanyID = 16 Then
		currentSupplierName = "Window Ware"
		currentSupplierAddress = "Window Ware, WIN003 \nTelford Way, Cross Park, Bedford, Bedforshire, MK42 0PQ"
		Set CompanyIconHeaderTop.Image = IconWindowWare
		currentInstructions = "1. Open up Eurostock and navigate to 'Maintenance -> Supplier'\n2. Find 'Window Widgets' under supplier code WIN003\n3. Navigate to your item code provided in the invoice\n4. Check respective pack prices; Update if necessary\n\n-Example item code: GTV205101-\n\n\n\n\n"
	End If
	If selectedCompanyID = 17 Then
		currentSupplierName = "Window Widgets"
		currentSupplierAddress = "Window Widgets LLP, WIN002 \nUnit 1 Venture Business Park, Madleaze Road, Gloucestershire, GL1 5SJ"
		Set CompanyIconHeaderTop.Image = IconWindowWidgets

	End If
	If selectedCompanyID = 18 Then
		currentSupplierName = "Press Glass"
		currentSupplierAddress = "Press Glass UK, PRE002"
		Set CompanyIconHeaderTop.Image = IconPressglass
		currentInstructions = "1. Open up the Press Glass pricesheet excel document; It should start with the phrase 'Liorder'\n2. Search for each element of an item listed in the description of the press glass invoice\n3. Total up the prices to find the £/sqm value\n4. Remember that some items are single price, regardless of area, and should be added on at the end\n\n\n\n\n"
	End If
End Sub

Private Sub LoadButton_Click()
	loadingBox.Visible = true
	Set loadingbox.Image = IconLoading
	'Slowly outputs the supplier guide to the textbox; Supplier guide text is changed based on the companyID in the selection sub
	'//--Special additions due to horizontal size of IDE; essentially adding on new lines to certain suppliers with long instrucitons--\\
		'//REHAU
		If selectedCompanyID = 15 Then
			currentInstructions = currentInstructions & " section\n7. Otherwise, you can create a new entry in Eurostock\n\n\n\n\n"
		End If
	'\\--=================================================END OF SPECIAL ADDITIONS===================================================--//

	'//--Output--\\
	OutputSlowly(currentInstructions)
	Set loadingbox.Image = IconDoneLoad
End Sub

'//--Company ID Reference List--\\
'Avantek = 1										|
'Bevelpane = 2									|
'Carl F = 3									  	|
'Cotswold = 4										|
'DGS = 5												|
'Distinction Doors = 6					|
'ERA (Fab & Fix) = 7						|
'Exitex = 8											|
'Liniar = 9											|
'Maco = 10									    |
'NW Metal = 11									|
'ODL = 12												|
'Premier Arches = 13						|
'Rapidrop = 14									|
'Rehau = 15									  	|
'Window Ware = 16								|
'Window Widgets = 17						|
'Press Glass = 18								|
'\\--Company ID Reference List--//



Private Sub Button1_Click() 'Return menu button
  dim f as new launchScreen
  f.Show hbFormModeless+hbFormGoto
End Sub
