
    'Title:   Multiple MSI/Patch/Setup Uninstall Script.
    'Version: 1.2.0
    'Author:  Guntars Eitvids
    'Contact: Guntars.Eitvids@atea.com
    'Date:    12/22/2017
    '
    'Description: VB Script was designed to perform automatic silent uninstall of multiple Application components (MSI, Patches, Setups)
    '			  The implemented "RunCheckExitCode()" Macro performs uninstallation using input "Command" and "Arguments" parameters.
    '			  Script returns 0 or 3010 code if uninstallation was successful. In case 3010 code uninstallation is successful, but reboot
    '			  is required to finish uninstallation. Script also returns error code if any error occured during uninstallation.
    '
    'Usage:   Run "INSTALL.vbs" from administrative console or deploy through SCCM
    '
    'Notes:   DO NOT MOVE SCRIPT ANYWHERE OUTSIDE FOLDER PROVIDED WITH ORIGINAL PACKAGE, THAT CONTAINED THIS UNINSTALLATION SCRIPT, OTHERWISE
    '	  IT WON'T WORK. DO NOT EDIT, MODIFY OR USE THIS SCRIPT FOR OTHER PURPOSES EXCEPT FOR UNINSTALLING ORIGINAL PACKAGE THAT CONTAINED
    '	  THIS SCRIPT. AUTHOR DOESN'T TAKE RESPONSIBILITY FOR ANY DAMAGE MADE BY THIS SCRIPT AS A RESULT OF IT'S INPROPER USE.

    Option Explicit

    DIM oShell,oEnv
    DIM SCRIPT_PATH,SYSTEM32_PATH
    DIM Command,Arguments
    DIM ReturnCode,ScriptExitCode,QM
    DIM MSI_INSTALL_MODE,DEBUG_MODE

    Set oShell = WScript.CreateObject ("WSCript.shell")

    'Temporarily Disables ZONE Checking. Allows executing applications from Network
    set oEnv = oShell.Environment("PROCESS")
    oEnv("SEE_MASK_NOZONECHECKS") = 1

    'Defines variable to add Quotation Marks to handle Paths with spaces. Just to make code more readable.
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
    
    'START OF Application (MSI/Patch/Setup) Uninstallation Sequence:

    dim ProgramFilesFolder

ProgramFilesFolder = oShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%")+"\"
if ProgramFilesFolder = "%PROGRAMFILES(x86)%"+"\" then ProgramFilesFolder = oShell.ExpandEnvironmentStrings("%PROGRAMFILES%")+"\" 

    Command = QM & SCRIPT_PATH & "vs_Enterprise.exe" & QM
    Arguments = " Uninstall --quiet --norestart --noweb --wait --installPath """+ProgramFilesFolder+"Microsoft Visual Studio\2017\Enterprise""" & MSI_ARGUMENTS
    RunCheckExitCode Command, Arguments

Dim filesys
set filesys = CreateObject ("Scripting.FileSystemObject")
if filesys.FileExists(ProgramFilesFolder+"Microsoft Visual Studio\Installer\vs_installer.exe") then
Command = QM & ProgramFilesFolder+"Microsoft Visual Studio\Installer\vs_installer.exe" & QM
    Arguments = " /Uninstall --quiet --norestart" & MSI_ARGUMENTS
    RunCheckExitCode Command, Arguments
end if
  


    'END OF Application (MSI/Patch/Setup) Uninstallation Sequence

    RemoveWrapperGuidFromRegistry

    'Disables ZONE Checking.
    oEnv.Remove("SEE_MASK_NOZONECHECKS")
    Set oEnv = Nothing

    Set oShell = Nothing
    WScript.Quit(ScriptExitCode)


    Sub RunCheckExitCode(Command,Arguments)

    DIM ReturnCode
    
    ReturnCode = oShell.Run(Command & " " & Arguments, 0, True)

    ' Below code is used for debugging. It shows "Command", "Arguments" values and ReturnCode value
    ' as a result of above "oShell.Run()" execution.
    If (DEBUG_MODE = True) Then
    WScript.Echo "Running Command:", _
    Command & Arguments, _
    "Exit Code: " & ReturnCode
    End If
    'Checks whether execution wasn't successful
    If (ReturnCode <> 0) AND (ReturnCode <> 3010) AND (ReturnCode <> 1605) Then
    'If any error occured, then script execution is terminated and error code is returned.
    ScriptExitCode = ReturnCode

    'Disables ZONE Checking.
    oEnv.Remove("SEE_MASK_NOZONECHECKS")
    Set oEnv = Nothing

    Set oShell = Nothing
    WScript.Quit(ScriptExitCode)
    ElseIf (ReturnCode = 3010) Then
    'If ReturnCode is "3010", set this value for VB Script return code and continue uninstallation.
    ScriptExitCode = 3010
    End If
    End Sub



  
    Sub RemoveWrapperGuidFromRegistry()
      Dim StubReturnCode
      Const strKeyPath = "HKLM\SOFTWARE\Applications\"
      Const WrapperGuid = "{9E9FE62C-97DF-4DDB-879E-8F2F54CF8C5C}"
      Command  = "cmd /c REG DELETE " & QM & strKeyPath & WrapperGuid & QM & " /f"
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
  

