
' #============================================================#
' #  SECDELW.VBS                                               #
' #============================================================#
' #  Securely deletes a file or a directory or securely wipes  #
' #  an unused disk space using Sysinternals SDelete.exe.      #
' #                                                            #
' #  Usage:                                                    #
' #      secdelw.vbs "target object"                           #
' #  or                                                        #
' #      wscript.exe secdelw.vbs "target object"               #
' #                                                            #
' #           Copyright(C) ZulNs, Yogyakarta, July 2'nd, 2013  #
' #============================================================#

Option Explicit

Const TITLE = "ZulNs: Securely Delete"
Const ADMIN = "~adm"
Const sDelFile = "sdelete.exe"

Dim quote
quote = Chr(34)

If Wscript.Arguments.Count = 0 Then
	MsgBox "Securely deletes a file or directory or " & _
			"wipes unused disk spaces." & vbCrLf & vbCrLf & _
			"Usage:" & vbCrLf & vbCrLf & _
			"wscript.exe " & Wscript.ScriptName & " " & _
			quote & "target object" & quote, _
			vbInformation, TITLE
	Wscript.Quit
End If

Dim shell, target, triQuote
target = Wscript.Arguments.Item(0)

If target <> ADMIN Then
	triQuote = quote & quote & quote
	Set shell = CreateObject("Shell.Application")
	target = quote & WScript.ScriptFullName & quote & _
			" " & ADMIN & " " & triQuote & target & triQuote
	shell.ShellExecute "wscript.exe", target, "", "runas", 1
	Set shell = Nothing
	Wscript.Quit
End If

target = Wscript.Arguments.Item(1)
Dim fso, retCode, program
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(fso.GetSpecialFolder(1) & "\" & sDelFile) Then
	MsgBox "Can't find " & quote & fso.GetSpecialFolder(1) & "\" & _
			sDelFile & quote & " file!!!", vbCritical, TITLE
	set fso = Nothing
	Wscript.Quit
End If

Set shell = CreateObject("Wscript.Shell")
program = "cmd.exe /C " & sDelFile & " -p 3 "

If fso.FileExists(target) Then
	' This is for file object
	If MsgBox("Are you sure to securely delete this below file?" & _
			vbCrLf & vbCrLf & quote & target & quote, _
			vbYesNo + vbQuestion, TITLE) = vbNo Then
		Set fso = Nothing
		Set shell = Nothing
		Wscript.Quit
	End If
	
	retCode = shell.Run(program & "-a " & _
			quote & target & quote, 1, True)
		
	MsgBox "DONE!!!", vbInformation, TITLE
	Set fso = Nothing
	Set shell = Nothing
	Wscript.Quit
End If

If fso.FolderExists(target) Then
	If fso.GetBaseName(target) = "" Then
		' This is for drive object
		If MsgBox("Are you sure to securely wipe unused space of " & _
				"drive " & fso.GetDriveName(target) & " ?", _
				vbYesNo + vbQuestion, TITLE) = vbNo Then
			Set fso = Nothing
			Set shell = Nothing
			Wscript.Quit
		End If
		
		retCode = shell.Run(program & "-c " & _
				fso.GetDriveName(target), 1, True)
	Else
		' This is for directory object
		If MsgBox("Are you sure to securely delete this below " & _
				"folder and all of its contents?" & vbCrLf & vbCrLf & _
				quote & target & quote, _
				vbYesNo + vbQuestion, TITLE) = vbNo Then
			Set fso = Nothing
			Set shell = Nothing
			Wscript.Quit
		End If
		
		retCode = shell.Run(program & "-a -s " & _
				quote & target & quote, 1, True)
	End If
	MsgBox "DONE!!!", vbInformation, TITLE
	Set fso = Nothing
	Set shell = Nothing
	Wscript.Quit
End If

MsgBox "Can't find below target:" & vbCrLf & vbCrLf & _
		quote & target & quote, vbCritical, TITLE

set fso = Nothing
Set shell = Nothing
