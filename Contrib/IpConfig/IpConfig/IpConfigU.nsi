!include MUI2.nsh
!include LogicLib.nsh

Name `Unicode Test IpConfig plugin`
OutFile `IpConfigU.exe`
ShowInstDetails show

!define MUI_FINISHPAGE_NOAUTOCLOSE

!insertmacro MUI_PAGE_INSTFILES

!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE English

Section
	IpConfig::GetNetworkAdapterIDFromIPAddress "192.168.1.33"
	Pop $0
	Pop $1
	StrCpy $0 "ok" 0 error
	IpConfig::GetNetworkAdapterMACAddress $1
	Pop $0
	Pop $2
	DetailPrint "MACAddress from adapter #$1: $2"
	Goto end
error:
	DetailPrint "$1"
end:
SectionEnd

