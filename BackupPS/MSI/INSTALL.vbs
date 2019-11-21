
    'Title:   Multiple MSI/Patch/Setup Install Script.
    'Version: 1.2.0
    'Author:  Guntars Eitvids
    'Contact: Guntars.Eitvids@atea.com 
    'Date:    12/22/2017
    '
    'Description: VB Script was designed to perform automatic silent install of multiple Application components (MSI, Patches, Setups)
    '			  The implemented "RunCheckExitCode()" Macro performs installation using input "Command" and "Arguments" parameters.
    '			  is required to finish installation. Script also returns error code if any error occured during installation.
    '
    'Usage:   Run "INSTALL.vbs" from administrative console or deploy through SCCM
    '
    'Notes:   DO NOT MOVE SCRIPT ANYWHERE OUTSIDE FOLDER PROVIDED WITH ORIGINAL PACKAGE, THAT CONTAINED THIS INSTALLATION SCRIPT, OTHERWISE
    '	  IT WON'T WORK. DO NOT EDIT, MODIFY OR USE THIS SCRIPT FOR OTHER PURPOSES EXCEPT FOR INSTALLING ORIGINAL PACKAGE THAT CONTAINED
    '	  THIS SCRIPT. AUTHOR DOESN'T TAKE RESPONSIBILITY FOR ANY DAMAGE MADE BY THIS SCRIPT AS A RESULT OF IT'S INPROPER USE.

    Option Explicit

    DIM oShell,oEnv
    DIM SCRIPT_PATH, SYSTEM32_PATH
    DIM Command,Arguments
    DIM ReturnCode,ScriptExitCode,QM
    DIM MSI_INSTALL_MODE,DEBUG_MODE

    Set oShell = WScript.CreateObject ("WSCript.shell")

    'Temporarily Disables ZONE Checking. Allows executing applications from Network
    set oEnv = oShell.Environment("PROCESS")
    oEnv("SEE_MASK_NOZONECHECKS") = 1

    'Defines variable to add quotation Marks to handle Paths with spaces. Just to make code more readable.
    QM = Chr(34)
    
    'Enable/Disable Debug Mode to see Command, Arguments formatting and Execution return code
    DEBUG_MODE = False

    'Defines path to this script
    SCRIPT_PATH = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - (Len(WScript.ScriptName)))

    'Defines path to "System32" folder
    SYSTEM32_PATH = oShell.ExpandEnvironmentStrings("%WINDIR%") & "\System32\"

    'Sets default Return code for this script to "0" (0 - means successful)
    'Will be overwriten by RunCheckExitCode() routine if any errors occur.
    ScriptExitCode = 0

    Dim ARG,ARGS,MSI_ARGUMENTS
    Dim ATEA_MSI_CACHE
    MSI_ARGUMENTS = " "
    Set ARGS = WScript.Arguments
    If (ARGS.Count>0) Then
      For each ARG in ARGS
        If InStr(ARG,"=") Then
          If InStr(ARG,"ATEAMSICACHE=") Then
            ATEA_MSI_CACHE = " " & Left(ARG,InStr(ARG,"=")) & Chr(34) & Right(ARG,Len(ARG)-InStr(ARG,"=")) & Chr(34) & " "
          End If
          MSI_ARGUMENTS = MSI_ARGUMENTS & " " & Left(ARG,InStr(ARG,"=")) & Chr(34) & Right(ARG,Len(ARG)-InStr(ARG,"=")) & Chr(34)
        End If
      Next
    End If
    
    'START OF Application (MSI/Patch/Setup) Installation Sequence:

    
    If  IsProductInstalled("{046806D1-0A38-3FCA-AF84-F71C50A0C363}") Then
	Command = QM & SCRIPT_PATH & "PreviousPackageUninstall\vs_premium.exe" & QM
	Arguments = " /Uninstall /quiet /norestart /noweb"
	RunCheckExitCode Command,Arguments
    End if
  
If  IsProductInstalled("{030A6785-C3A9-37DA-8530-444C320629FA}") Then
    Command = QM & SCRIPT_PATH & "PreviousPackageUninstall\vs_enterprise.exe" & QM
    Arguments = " /Uninstall /quiet /norestart /noweb" & MSI_ARGUMENTS
    RunCheckExitCode Command, Arguments
end if  

    If  IsProductInstalled("{5DD91CC6-7665-4567-949D-6892999CA12A}") Then
    Command = QM & SYSTEM32_PATH & "msiexec.exe" & QM
    Arguments = " /x {5DD91CC6-7665-4567-949D-6892999CA12A} /qn" & MSI_ARGUMENTS
    RunCheckExitCode Command, Arguments
    End if

    Command = QM & SCRIPT_PATH & "vs_Enterprise.exe" & QM
    Arguments = " --ProductKey VN6CV-86GKQ-DHCM7-4T3QG-VCF3G --quiet --norestart --noweb --wait" & MSI_ARGUMENTS
    RunCheckExitCode Command, Arguments

DeleteStartMenuShortcut "Visual Studio Installer"

Sub DeleteStartMenuShortcut (ShortcutName)
dim filesys, demofolder, oShell, ProgramFiles,AppFolder, PUBLICE
set filesys = CreateObject ("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
PUBLICE = oShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")

If filesys.FileExists(PUBLICE+"\Start Menu\Programs\"+ShortcutName+".lnk") Then
filesys.DeleteFile(PUBLICE+"\Start Menu\Programs\"+ShortcutName+".lnk"), True
End if

If filesys.FileExists(PUBLICE+"\Microsoft\Windows\Start Menu\Programs\"+ShortcutName+".lnk") Then
filesys.DeleteFile(PUBLICE+"\Microsoft\Windows\Start Menu\Programs\"+ShortcutName+".lnk"), True
End if
End Sub
  
AddNoModifyRegistryValue "{6F320B93-EE3C-4826-85E0-ADF79F8D4C61}","Contact your local administrator.","","","",0,0,0
AddNoModifyRegistryValue "e8782f6b","Contact your local administrator.","","","",0,0,0

Sub AddNoModifyRegistryValue(UninstallSubkey,Contact,Comments,HelpTelephone,HelpLink,NoRepair,NoRemove,SYSTEMCOMPONENT)
	Dim oShell
	Dim RegistryPath
	Set oShell = CreateObject("WScript.Shell")
	if KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"+UninstallSubkey+"\") then
		RegistryPath="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"+UninstallSubkey
	end if
	if KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"+UninstallSubkey+"\") then
		RegistryPath="HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"+UninstallSubkey
	end if
if KeyExists(RegistryPath+"\") then
	oShell.RegWrite RegistryPath+"\NoModify", "1", "REG_DWORD"
	if KeyExists(RegistryPath+"\URLUpdateInfo") then 
		oShell.RegWrite RegistryPath+"\URLUpdateInfo", "", "REG_SZ" 
	end if	
	if KeyExists(RegistryPath+"\URLInfoAbout") then 
		oShell.RegWrite RegistryPath+"\URLInfoAbout", "", "REG_SZ" 
	end if		
	if KeyExists(RegistryPath+"\URLUpdateInfo") then 
		oShell.RegWrite RegistryPath+"\URLUpdateInfo", "", "REG_SZ" 
	end if
	if KeyExists(RegistryPath+"\Readme") then 
		oShell.RegWrite RegistryPath+"\Readme", "", "REG_SZ" 
	end if
	if Comments<>"" then
		oShell.RegWrite RegistryPath+"\Comments", Comments, "REG_SZ" 
	else
		if KeyExists(RegistryPath+"\Comments") then 
			oShell.RegWrite RegistryPath+"\Comments", "", "REG_SZ" 
		end if
	end if
	if Contact<>"" then
		oShell.RegWrite RegistryPath+"\Contact", Contact, "REG_SZ" 
	else
		if KeyExists(RegistryPath+"\Contact") then 
			oShell.RegWrite RegistryPath+"\Contact", "", "REG_SZ" 
		end if
	end if
	if HelpTelephone<>"" then
		oShell.RegWrite RegistryPath+"\HelpTelephone", HelpTelephone, "REG_SZ" 
	else
		if KeyExists(RegistryPath+"\HelpTelephone") then 
			oShell.RegWrite RegistryPath+"\HelpTelephone", "", "REG_SZ" 
		end if
	end if
	if HelpLink<>"" then
		oShell.RegWrite RegistryPath+"\HelpLink", HelpLink, "REG_SZ" 
	else
		if KeyExists(RegistryPath+"\HelpLink") then 
			oShell.RegWrite RegistryPath+"\HelpLink", "", "REG_SZ" 
		end if
	end if
	if NoRepair<>0 then
		oShell.RegWrite RegistryPath+"\NoRepair", NoRepair, "REG_DWORD" 		
	end if
	if NoRemove<>0 then
		oShell.RegWrite RegistryPath+"\NoRemove", NoRemove, "REG_DWORD" 		
	end if
	if SYSTEMCOMPONENT=1 then
		oShell.RegWrite RegistryPath+"\SYSTEMCOMPONENT", SYSTEMCOMPONENT, "REG_DWORD" 		
	end if
end if	
End Sub
Function KeyExists(key)
    Dim oShell
    On Error Resume Next
    Set oShell = CreateObject("WScript.Shell")
        oShell.RegRead (key)
    Set oShell = Nothing
    If Err = 0 Then KeyExists = True
End Function

    'END OF Application (MSI/Patch/Setup) Installation Sequence

    'Remove old GUIDs for upgrade
    RemoveWrapperGuidFromRegistry "{5E1BD337-2951-4A3D-A597-EEBD85B688BA}"

    'Adds registry for Sccm application detection

    AddWrapperGuidToRegistry

    'Disables ZONE Checking.
    oEnv.Remove("SEE_MASK_NOZONECHECKS")

    Set oShell = Nothing
    WScript.Quit(ScriptExitCode)

    Sub RunCheckExitCode(Command,Arguments)

    DIM ReturnCode

    ReturnCode = oShell.Run(Command & " " & Arguments, 0, True)

    ' Below code is used for debugging. It shows "Command", "Arguments" values and ReturnCode value
    ' as a result of above "oShell.Run()" execution.
    If (DEBUG_MODE = True) Then
    WScript.Echo "Running Command:", Command & Arguments, "Exit Code: " & ReturnCode
    End If
    'Checks whether execution wasn't successful
    If (ReturnCode <> 0) AND (ReturnCode <> 3010) AND (ReturnCode <> 1605) Then
    'If any error occured, then script execution is terminated and error code is returned.
    ScriptExitCode = ReturnCode

    'Disables ZONE Checking.
    oEnv.Remove("SEE_MASK_NOZONECHECKS")

    Set oShell = Nothing
    WScript.Quit(ScriptExitCode)
    ElseIf (ReturnCode = 3010) Then
    'If ReturnCode is "3010", set this value for VB Script return code and continue installation.
    ScriptExitCode = 3010
    End If
    End Sub

    Sub RemoveWrapperGuidFromRegistry(ByVal WrapperGuid)
      Dim StubReturnCode
      Const strKeyPath = "HKLM\SOFTWARE\Atea\Applications\"
      Command  = "cmd /c REG DELETE " &  QM &  strKeyPath &  WrapperGuid &  QM &  " /f"
      Arguments = ""
      StubReturnCode = oShell.Run(Command &  " " &  Arguments, 0, True)
    End Sub

    
    Sub AddWrapperGuidToRegistry()
      Dim StubReturnCode
      Const strKeyPath = "HKLM\SOFTWARE\Applications\"
      Const WrapperGuid = "{9E9FE62C-97DF-4DDB-879E-8F2F54CF8C5C}"
      Command   = "cmd /c REG ADD " & QM & strKeyPath & WrapperGuid & QM & " /f"
      Arguments = ""
      StubReturnCode = oShell.Run(Command & " " & Arguments, 0, True)
    End Sub
  
    
   Function IsProductInstalled(ByVal sProductCode)
    Dim installer,productCode
    Set installer = CreateObject("WindowsInstaller.Installer")
    For Each productCode In installer.Products
      If LCase(productCode) = LCasE(sProductCode) Then
        IsProductInstalled = true
        Exit function
      End If
    Next
    IsProductInstalled = false
   End Function
  

  