#NoTrayIcon
Opt('MustDeclareVars', 1)

;;;;;;;;;;;;;;; OPTIONS!
Global $requiredProgramPath="c:\Program Files (x86)\Vodafone\Vodafone Communication Centre\PCCommunicator.exe"
Global $txtAddressName = "Vodafone Communication Centre Line"
Global $Number = ""
;;;;;;;;;;;;;; END OF OPTIONS


;Based on from https://www.autoitscript.com/forum/topic/106221-tapi-call/
;Kudos to Tec from Autoit Forum!

Const $LINEADDRESSTYPE_PHONENUMBER = "&H1"
Const $LINEMEDIAMODE_INTERACTIVEVOICE = "&H4"
Const $TAPI3_ALL_TAPI_EVENTS = "&H1FFFF"
Local $lAddressType = $LINEADDRESSTYPE_PHONENUMBER

;Check if a required program exists
If($requiredProgramPath<>"") Then
	If(Not FileExists($requiredProgramPath)) Then
		MsgBox(16,"TapiCall - Error","Required program has not been found.")
		Exit
	EndIf
EndIf

;Associate this program with the tel: protocol
RegWrite("HKCU64\SOFTWARE\RegisteredApplications","TapiCall","REG_SZ","SOFTWARE\\TapiCall\\TapiCall\\Capabilities")
RegWrite("HKCU64\SOFTWARE\TapiCall\TapiCall\Capabilities","ApplicationDescription","REG_SZ","Interface between tel: protocol and TAPI.")
RegWrite("HKCU64\SOFTWARE\TapiCall\TapiCall\Capabilities","ApplicationName","REG_SZ","TapiCall")
RegWrite("HKCU64\SOFTWARE\TapiCall\TapiCall\Capabilities\UrlAssociations","tel","REG_SZ","TapiCall.TapiCall")
RegWrite("HKCU64\SOFTWARE\Classes\TapiCall.TapiCall","","REG_SZ","TapiCall")
RegWrite("HKCU64\SOFTWARE\Classes\TapiCall.TapiCall","FriendlyTypeName","REG_SZ","TapiCall")
RegWrite("HKCU64\SOFTWARE\Classes\TapiCall.TapiCall\DefaultIcon","","REG_SZ",'"'&$requiredProgramPath&',0"')
RegWrite("HKCU64\SOFTWARE\Classes\TapiCall.TapiCall\shell\open\command","","REG_SZ",@ScriptFullPath&" %1")

;Receive tel: string from command line
If($cmdLine[0]=1) Then
	$Number=StringReplace($cmdLine[1],"tel:","")
	if(StringLeft($Number,1)<>"+") Then
		$Number="+"&$Number
	EndIf
Else
	MsgBox(16,"TapiCall - Error","Wrong command line parameters."&@CRLF&@CRLF&"Usage: "&@ScriptName&" tel:+PHONENUMBER")
	Exit
EndIf

;Find out if the required TAPI "address" does exist
Local $objTapi = ObjCreate("TAPI.TAPI.1")
Local $mapper = ObjCreate("DispatchMapper.DispatchMapper.1")
$objTapi.Initialize
$objTapi.EventFilter = $TAPI3_ALL_TAPI_EVENTS

Local $bolFoundLine = False
Local $objCollAddresses = $objTapi.Addresses
Local $objCrtAddress
Local $gobjAddress
For $lLoop = 1 To $objCollAddresses.Count
	$objCrtAddress = $objCollAddresses.Item($lLoop)
	If $objCrtAddress.AddressName = $txtAddressName Then
		$gobjAddress = $objCrtAddress
		$bolFoundLine = True
		ExitLoop
	EndIf
Next

;TAPI address does exist, we make the call
If $bolFoundLine = True Then
	Local $TestCall = $gobjAddress.CreateCall($Number, $lAddressType, $LINEMEDIAMODE_INTERACTIVEVOICE)
	$TestCall.connect(False)
Else ;TAPI address was not found
	MsgBox(16,"TapiCall - Error",'Could not find TAPI address "'&$txtAddressName&'"')
EndIf
