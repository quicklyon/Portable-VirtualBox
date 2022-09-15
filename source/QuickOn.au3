; Language       : multilanguage
; Author         : ZhouYueQiu
; e-Mail         : zhouyueqiu@easycorp.ltd
; License        : http://creativecommons.org/licenses/by-nc-sa/3.0/
; Version        : 1.0.0
; Download       : https://www.qucheng.com/page/download.html
; Support        : https://www.qucheng.com/book/Installation-manual/quick-install-6.html

#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=QuickOn.ico
#AutoIt3Wrapper_Compile_Both=Y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <ColorConstants.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <DirConstants.au3>
#include <FileConstants.au3>
#include <IE.au3>
#include <ProcessConstants.au3>
#include <String.au3>
#include <WinAPIError.au3>

Opt("GUIOnEventMode", 1)
Opt("TrayAutoPause", 0)
Opt("TrayMenuMode", 11)
Opt("TrayOnEventMode", 1)

TraySetClick(16)
TraySetState()
TraySetToolTip("QuickOn")

Global $version = "1.0.0"
Global $cfg = @ScriptDir & "\data\settings\settings.ini"
Global $langDir = @ScriptDir & "\data\language\"
Global $lng = IniRead($cfg, "language", "key", "NotFound")
Global $updateUrl = IniRead(@ScriptDir & "\data\settings\vboxinstall.ini", "download", "update", "NotFound")

Global $new1 = 0, $new2 = 0

Global $startvbox = 1

; 设置VirtualBox.xml配置文件
If (FileExists(@ScriptDir & "\app64\virtualbox.exe")) And ($startvbox = 1 Or IniRead(@ScriptDir & "\data\settings\vboxinstall.ini", "startvbox", "key", "NotFound") = 1) Then
	Global $arch = "app64"

	If FileExists(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml-prev") Then
		FileDelete(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml-prev")
	EndIf

	If FileExists(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml-tmp") Then
		FileDelete(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml-tmp")
	EndIf

	If FileExists(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml") Or (FileExists(@ScriptDir & "\data\.VirtualBox\Machines\") And FileExists(@ScriptDir & "\data\.VirtualBox\HardDisks\")) Then
		Local $values0, $values1, $values2, $values3, $values4, $values5, $values6, $values7, $values8, $values9, $values10, $values11, $values12, $values13
		Local $line, $content, $i, $j, $k, $l, $m, $n
		Local $file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128)
		If $file <> -1 Then
			$line = FileRead($file)
			$values0 = _StringBetween($line, '<MachineRegistry>', '</MachineRegistry>')
			If $values0 = 0 Then
				$values1 = 0
			Else
				$values1 = _StringBetween($values0[0], 'src="', '"')
			EndIf
			$values2 = _StringBetween($line, '<HardDisks>', '</HardDisks>')
			If $values2 = 0 Then
				$values3 = 0
			Else
				$values3 = _StringBetween($values2[0], 'location="', '"')
			EndIf
			$values4 = _StringBetween($line, '<DVDImages>', '</DVDImages>')
			If $values4 = 0 Then
				$values5 = 0
			Else
				$values5 = _StringBetween($values4[0], '<Image', '/>')
			EndIf
			$values10 = _StringBetween($line, '<Global>', '</Global>')
			If $values10 = 0 Then
				$values11 = 0
			Else
				$values11 = _StringBetween($values10[0], '<SystemProperties', '/>')
			EndIf

			For $i = 0 To UBound($values1) - 1
				$values6 = _StringBetween($values1[$i], 'Machines', '.vbox')
				If $values6 <> 0 Then
					$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
					$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
					FileWrite($file, StringReplace($content, $values1[$i], "Machines" & $values6[0] & ".vbox"))
					FileClose($file)
				EndIf
			Next

			For $j = 0 To UBound($values3) - 1
				$values7 = _StringBetween($values3[$j], 'HardDisks', '.vdi')
				If $values7 <> 0 Then
					$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
					$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
					FileWrite($file, StringReplace($content, $values3[$j], "HardDisks" & $values7[0] & ".vdi"))
					FileClose($file)
				EndIf
			Next

			For $k = 0 To UBound($values3) - 1
				$values8 = _StringBetween($values3[$k], 'Machines', '.vdi')
				If $values8 <> 0 Then
					$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
					$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
					FileWrite($file, StringReplace($content, $values3[$k], "Machines" & $values8[0] & ".vdi"))
					FileClose($file)
				EndIf
			Next

			For $l = 0 To UBound($values5) - 1
				$values9 = _StringBetween($values5[$l], 'location="', '"')
				If $values9 <> 0 Then
					If Not FileExists($values9[0]) Then
						$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
						$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
						FileWrite($file, StringReplace($content, "<Image" & $values5[$l] & "/>", ""))
						FileClose($file)
					EndIf
				EndIf
			Next

			For $m = 0 To UBound($values11) - 1
				$values12 = _StringBetween($values11[$m], 'defaultMachineFolder="', '"')
				If $values12 <> 0 Then
					If Not FileExists($values10[0]) Then
						$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
						$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
						FileWrite($file, StringReplace($content, $values12[0], @ScriptDir & "\data\.VirtualBox\Machines"))
						FileClose($file)
					EndIf
				EndIf
			Next

			For $n = 0 To UBound($values1) - 1
				$values13 = _StringBetween($values1[$n], 'Machines', '.xml')
				If $values13 <> 0 Then
					$content = FileRead(FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 128))
					$file = FileOpen(@ScriptDir & "\data\.VirtualBox\VirtualBox.xml", 2)
					FileWrite($file, StringReplace($content, $values1[$n], "Machines" & $values13[0] & ".xml"))
					FileClose($file)
				EndIf
			Next

			FileClose($file)
		EndIf
	Else
		MsgBox(0, IniRead($langDir & $lng & ".ini", "download", "15", "NotFound"), IniRead($langDir & $lng & ".ini", "download", "16", "NotFound"))
	EndIf


	If FileExists(@ScriptDir & "\" & $arch & "\VirtualBox.exe") And FileExists(@ScriptDir & "\" & $arch & "\VBoxSVC.exe") And FileExists(@ScriptDir & "\" & $arch & "\VBoxC.dll") Then
		If Not ProcessExists("VirtualBox.exe") Or Not ProcessExists("VBoxManage.exe") Then
			If FileExists(@ScriptDir & "\data\settings\SplashScreen.jpg") Then
				SplashImageOn("Portable-VirtualBox", @ScriptDir & "\data\settings\SplashScreen.jpg", 480, 360, -1, -1, 1)
			Else
				SplashTextOn("Portable-VirtualBox", IniRead($langDir & $lng & ".ini", "messages", "06", "NotFound"), 220, 40, -1, -1, 1, "arial", 12)
			EndIf

			If Not FileExists(@SystemDir & "\msvcp100.dll") Or Not FileExists(@SystemDir & "\msvcr100.dll") Then
				FileCopy(@ScriptDir & "\app64\msvcp100.dll", @SystemDir, 9)
				FileCopy(@ScriptDir & "\app64\msvcr100.dll", @SystemDir, 9)
				Local $msv = 3
			Else
				Local $msv = 0
			EndIf

			RunWait($arch & "\VBoxSVC.exe /reregserver", @ScriptDir, @SW_HIDE)
			RunWait(@SystemDir & "\regsvr32.exe /S " & $arch & "\VBoxC.dll", @ScriptDir, @SW_HIDE)
			DllCall($arch & "\VBoxRT.dll", "hwnd", "RTR3Init")

			SplashOff()

			If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxDRV", "DisplayName") <> "VirtualBox Service" Then
				RunWait("cmd /c sc create VBoxDRV binpath= ""%CD%\" & $arch & "\drivers\VBoxDrv\VBoxDrv.sys"" type= kernel start= auto error= normal displayname= PortableVBoxDRV", @ScriptDir, @SW_HIDE)
				Local $DRV = 1
			Else
				Local $DRV = 0
			EndIf

			If IniRead($cfg, "usb", "key", "NotFound") = 1 Then
				If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSB", "DisplayName") <> "VirtualBox USB" Then
					If @OSArch = "x86" Then
						RunWait(@ScriptDir & "\data\tools\devcon_x86.exe install .\" & $arch & "\drivers\USB\device\VBoxUSB.inf ""USB\VID_80EE&PID_CAFE""", @ScriptDir, @SW_HIDE)
					EndIf
					If @OSArch = "x64" Then
						RunWait(@ScriptDir & "\data\tools\devcon_x64.exe install .\" & $arch & "\drivers\USB\device\VBoxUSB.inf ""USB\VID_80EE&PID_CAFE""", @ScriptDir, @SW_HIDE)
					EndIf
					FileCopy(@ScriptDir & "\" & $arch & "\drivers\USB\device\VBoxUSB.sys", @SystemDir & "\drivers", 9)
					Local $USB = 1
				Else
					Local $USB = 0
				EndIf
			Else
				Local $USB = 0
			EndIf

			If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSBMon", "DisplayName") <> "VirtualBox USB Monitor Driver" Then
				RunWait("cmd /c sc create VBoxUSBMon binpath= ""%CD%\" & $arch & "\drivers\USB\filter\VBoxUSBMon.sys"" type= kernel start= auto error= normal displayname= PortableVBoxUSBMon", @ScriptDir, @SW_HIDE)
				Local $MON = 1
			Else
				Local $MON = 0
			EndIf

			If IniRead($cfg, "net", "key", "NotFound") = 1 Then
				If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetAdp", "DisplayName") <> "VirtualBox Host-Only Network Adapter" Then
					If @OSArch = "x86" Then
						RunWait(@ScriptDir & "\data\tools\devcon_x86.exe install .\" & $arch & "\drivers\network\netadp\VBoxNetAdp.inf ""sun_VBoxNetAdp""", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\devcon_x86.exe install .\" & $arch & "\drivers\network\netadp6\VBoxNetAdp6.inf ""sun_VBoxNetAdp""", @ScriptDir, @SW_HIDE)
					EndIf
					If @OSArch = "x64" Then
						RunWait(@ScriptDir & "\data\tools\devcon_x64.exe install .\" & $arch & "\drivers\network\netadp\VBoxNetAdp.inf ""sun_VBoxNetAdp""", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\devcon_x64.exe install .\" & $arch & "\drivers\network\netadp6\VBoxNetAdp6.inf ""sun_VBoxNetAdp""", @ScriptDir, @SW_HIDE)
					EndIf
					FileCopy(@ScriptDir & "\" & $arch & "\drivers\network\netadp\VBoxNetAdp.sys", @SystemDir & "\drivers", 9)

					FileCopy(@ScriptDir & "\" & $arch & "\drivers\network\netadp6\VBoxNetAdp6.sys", @SystemDir & "\drivers", 9)
					Local $ADP = 1
				Else
					Local $ADP = 0
				EndIf
			Else
				Local $ADP = 0
			EndIf

			If IniRead($cfg, "net", "key", "NotFound") = 1 Then
				If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetFlt", "DisplayName") <> "VBoxNetFlt Service" Then
					If @OSArch = "x86" Then
						RunWait(@ScriptDir & "\data\tools\snetcfg_x86.exe -v -u sun_VBoxNetFlt", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\snetcfg_x86.exe -v -l .\" & $arch & "\drivers\network\netflt\VBoxNetFlt.inf -m .\" & $arch & "\drivers\network\netflt\VBoxNetFltM.inf -c s -i sun_VBoxNetFlt", @ScriptDir, @SW_HIDE)

						RunWait(@ScriptDir & "\data\tools\snetcfg_x86.exe -v -u oracle_VBoxNetLwf", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\snetcfg_x86.exe -v -l .\" & $arch & "\drivers\network\netlwf\VBoxNetLwf.inf -m .\" & $arch & "\drivers\network\netlwf\VBoxNetLwf.inf -c s -i oracle_VBoxNetLwf", @ScriptDir, @SW_HIDE)
					EndIf
					If @OSArch = "x64" Then
						RunWait(@ScriptDir & "\data\tools\snetcfg_x64.exe -v -u sun_VBoxNetFlt", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\snetcfg_x64.exe -v -l .\" & $arch & "\drivers\network\netflt\VBoxNetFlt.inf -m .\" & $arch & "\drivers\network\netflt\VBoxNetFltM.inf -c s -i sun_VBoxNetFlt", @ScriptDir, @SW_HIDE)
						
						RunWait(@ScriptDir & "\data\tools\snetcfg_x64.exe -v -u oracle_VBoxNetLwf", @ScriptDir, @SW_HIDE)
						RunWait(@ScriptDir & "\data\tools\snetcfg_x64.exe -v -l .\" & $arch & "\drivers\network\netlwf\VBoxNetLwf.inf -m .\" & $arch & "\drivers\network\netlwf\VBoxNetLwf.inf -c s -i oracle_VBoxNetLwf", @ScriptDir, @SW_HIDE)
					EndIf
					FileCopy(@ScriptDir & "\" & $arch & "\drivers\network\netflt\VBoxNetFltNobj.dll", @SystemDir, 9)
					FileCopy(@ScriptDir & "\" & $arch & "\drivers\network\netflt\VBoxNetFlt.sys", @SystemDir & "\drivers", 9)
					RunWait(@SystemDir & "\regsvr32.exe /S " & @SystemDir & "\VBoxNetFltNobj.dll", @ScriptDir, @SW_HIDE)
					
					FileCopy(@ScriptDir & "\" & $arch & "\drivers\network\netlwf\VBoxNetLwf.sys", @SystemDir & "\drivers", 9)
					Local $NET = 1
				Else
					Local $NET = 0
				EndIf
			Else
				Local $NET = 0
			EndIf

			If $DRV = 1 Then
				RunWait("sc start VBoxDRV", @ScriptDir, @SW_HIDE)
			EndIf

			If $USB = 1 Then
				RunWait("sc start VBoxUSB", @ScriptDir, @SW_HIDE)
			EndIf

			If $MON = 1 Then
				RunWait("sc start VBoxUSBMon", @ScriptDir, @SW_HIDE)
			EndIf

			If $ADP = 1 Then
				RunWait("sc start VBoxNetAdp", @ScriptDir, @SW_HIDE)
			EndIf

			If $NET = 1 Then
				RunWait("sc start VBoxNetFlt", @ScriptDir, @SW_HIDE)
			EndIf

		Else
			WinSetState("Oracle VM VirtualBox Manager", "", BitAND(@SW_SHOW, @SW_RESTORE))
			WinSetState("] - Oracle VM VirtualBox", "", BitAND(@SW_SHOW, @SW_RESTORE))
		EndIf
	Else
		SplashOff()
		MsgBox(0, IniRead($langDir & $lng & ".ini", "messages", "01", "NotFound"), IniRead($langDir & $lng & ".ini", "start", "01", "NotFound"))
	EndIf
EndIf

Break(1)
Exit