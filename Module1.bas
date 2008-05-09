Attribute VB_Name = "mMisc"
'    Ultimate Webshots Converter 1
'    Copyright (C) 2007  Hervé "Setaou" BRY <uwc at apinc dot org>
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Option Explicit

Public Const BETA As String = ""

Public Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetTempPathA Lib "kernel32" _
   (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const WBC_File_Marker As Long = &H95FA16AB
Public Const WBZ_File_Marker As Long = &H6791AB43
Public Const WB1_File_Marker_NoCode As Long = &HE0FFD8FF 'Header JPG non codé
Public Const WB1_File_Marker_Code As Long = &H42425757 '"WWBB" en hexa

Public Const CTRL_ESPACEMENT As Integer = 120

Public Const FAE_SKIP = 0
Public Const FAE_OVERWRITE = 1
Public Const FAE_RENAME = 2

Public WebshotsExtensions As String

Public EvalExpr As New cEvalExpr

'---------- Toolbar

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
            ByVal lpsz2 As String) As Long

' Toolbar constants
Public Const WM_USER = &H400
Public Const TBSTYLE_FLAT As Long = &H800
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TB_SETIMAGELIST = WM_USER + 48
Public Const TB_SETHOTIMAGELIST = WM_USER + 52
Public Const TB_SETDISABLEDIMAGELIST = WM_USER + 54


'--------- Win version
Private Declare Function GetVersionEx _
    Lib "kernel32" Alias "GetVersionExA" _
    (VersionInfo As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'-------- Listbox
Public Declare Function SendMessageByNum Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const LB_SETHORIZONTALEXTENT = &H194

'-------- Liste des fichiers
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_NO_MORE_FILES = 18
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_DEVICE = &H40

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'Gestion clavier
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'Shellexecute
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Function FormatPath(ByVal path As String, Optional ByVal Directory As Boolean = False) As String
    Dim ff As String
    ff = path
    
    ff = Replace(ff, "/", "_")
    If Not Directory Then ff = Replace(ff, "\", "_")
    ff = Replace(ff, "*", "_")
    ff = Replace(ff, "?", "_")
    If Not Directory Then ff = Replace(ff, ":", "_")
    ff = Replace(ff, """", "_")
    ff = Replace(ff, "<", "_")
    ff = Replace(ff, ">", "_")
    ff = Replace(ff, "|", "_")
    
    ff = Replace(ff, vbCr, "")
    ff = Replace(ff, vbLf, "")
    
    If Directory Then
        If Right(path, 1) <> "\" Then path = path + "\"
    End If
    
    FormatPath = ff
End Function
Sub lstAddHScroll(lst As ListBox)
    ' depends on the scalewidth
    ' if scalemode is Twips then Divide M by 15 to get Pixels
    Dim a As Integer, m As Integer

    For a = 0 To lst.ListCount - 1
        If MainForm.TextWidth(lst.List(a)) > m Then m = MainForm.TextWidth(lst.List(a))
    Next

    SendMessageByNum lst.hWnd, LB_SETHORIZONTALEXTENT, m / Screen.TwipsPerPixelX + 10, 0
End Sub

Public Function WindowsVersion() As String
    Dim VerInfo As OSVERSIONINFO
    Dim WindowsName As String

    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    GetVersionEx VerInfo

    Select Case VerInfo.dwPlatformId
        Case VER_PLATFORM_WIN32_NT
            If VerInfo.dwMajorVersion = 3 Then
                WindowsName = "Windows NT 3"
            ElseIf VerInfo.dwMajorVersion = 4 Then
                WindowsName = "Windows NT 4"
            ElseIf VerInfo.dwMajorVersion = 5 And VerInfo.dwMinorVersion = 0 Then
                WindowsName = "Windows 2000"
            ElseIf VerInfo.dwMajorVersion = 5 And VerInfo.dwMinorVersion = 1 Then
                WindowsName = "Windows XP"
            ElseIf VerInfo.dwMajorVersion = 6 And VerInfo.dwMinorVersion = 0 Then
                WindowsName = "Windows Vista"
            End If
        Case VER_PLATFORM_WIN32_WINDOWS
            If VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 0 Then
                WindowsName = "Windows 95"
            ElseIf VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 10 Then
                WindowsName = "Windows 98"
            ElseIf VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 90 Then
                WindowsName = "Windows Me"
            End If
    End Select

    WindowsVersion = WindowsName
End Function


Public Function GetTempPath() As String
    Dim s As String
    Dim i As Integer
    i = GetTempPathA(0, "")
    s = Space(i)
    Call GetTempPathA(i, s)
    s = Left$(s, i - 1)
    
    If Len(s) > 0 Then
    
        If Right$(s, 1) <> "\" Then
            GetTempPath = s + "\"
        Else
            GetTempPath = s
        End If
    
    Else
    
        GetTempPath = App.path + "\"
      
    End If
End Function

Public Function GetDirectory(ByVal path As String) As String
    GetDirectory = Mid(path, 1, InStrRev(path, "\"))
End Function

Public Function GetFileName(ByVal path As String) As String
    GetFileName = Mid(path, InStrRev(path, "\") + 1, Len(path) - InStrRev(path, "\"))
End Function

Public Function DirectoryExists(ByVal Doss As String) As Boolean
    Dim attrib As Long
    If Right(Doss, 1) <> "\" Then Doss = Doss + "\"
    
    attrib = GetFileAttributes(Doss)
    DirectoryExists = (Not (attrib = INVALID_HANDLE_VALUE)) And (attrib And FILE_ATTRIBUTE_DIRECTORY)
End Function

Public Function FileExists(ByVal File As String) As Boolean
    Dim attrib As Long
    
    attrib = GetFileAttributes(File)
    FileExists = (Not (attrib = INVALID_HANDLE_VALUE)) And (Not (attrib And FILE_ATTRIBUTE_DIRECTORY))
End Function

Public Sub OpenFile(Fichier As String)
    Dim file_marker As Long, f As Integer
    
    'Si le fichier existe
    If FileExists(Fichier) Then
        f = FreeFile
        Open Fichier For Binary As f
            Get f, , file_marker
        Close f
        
        Select Case file_marker
            Case WBC_File_Marker 'Fichier WBC
                Dim NewWBC As New WBC
                Load NewWBC
                NewWBC.Show
                NewWBC.Chargement Fichier
                
            Case WB1_File_Marker_Code, WB1_File_Marker_NoCode 'Fichier WB1, WBD, JPG
                Dim NewWB1 As New WB1
                Load NewWB1
                NewWB1.Show
                NewWB1.Chargement Fichier
                
            Case WBZ_File_Marker 'Fichier WBZ
                Dim NewWBZ As New WBZ
                Load NewWBZ
                NewWBZ.Show
                NewWBZ.Chargement Fichier
                
            Case Else 'Type inconnu
                MsgBox sprintf(GetTranslation(8), Fichier), vbCritical, GetTranslation(9)
                
        End Select
        
    'Si le fichier n'existe pas
    Else
        MsgBox sprintf(GetTranslation(13, "Le fichier que vous tentez d'ouvrir n'existe pas ! \r\n\n Fichier : %s"), Fichier), vbCritical + vbOKOnly, GetTranslation(14, "Erreur : Fichier inexistant !")
    
    End If
End Sub

Public Function KeyPressed(Key As Long) As Boolean
    KeyPressed = ((GetKeyState(Key) And &H80) = &H80)
End Function

Public Function Version() As String
    Version = CStr(App.Major) + "." + CStr(App.Minor)
End Function
Public Function VersionLong() As String
    VersionLong = CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Function

Public Function CreateFolder(Folder As String) As Integer
    Dim i As Integer, start As Integer

    If Right(Folder, 1) <> "\" Then Folder = Folder + "\"
    
    If Left(Folder, 2) = "\\" Then ' If the folder is on an network share
        start = InStr(3, Folder, "\") + 1 ' The test starts after the hostname (so after the 3rd "\" : ex \\host\share\folder)
    Else
        start = 1
    End If
    
    For i = start To Len(Folder)
        i = InStr(i, Folder, "\")
        If i = 0 Then Exit For
        
        If Not DirectoryExists(Left(Folder, i)) Then MkDir Left(Folder, i)
    Next
End Function

Public Function CmdLineParam(Cmdline As String, Param As Integer) As String
    Dim Element() As String
    Dim NumElement As Integer, i As Integer, j As Integer
    Dim LongElement As Boolean
    
    Element = Split(Cmdline, " ")
    
    LongElement = False
    
    For i = LBound(Element) To UBound(Element)
        j = InStr(1, Element(i), """")
        
        'Si pas d'element long en cours on incrémente le compteur
        If LongElement = False Then NumElement = NumElement + 1
        
        'Si l'element en cours fait partie du param recherché, on l'ajoute
        If NumElement = Param Then CmdLineParam = CmdLineParam + Element(i) + " "
        
        'Si l'element en cours dépasse l'élement recherché, sort de la boucle
        If NumElement > Param Then Exit For
        
        'Si " au début => début d'élément long
        If j = 1 Then LongElement = True
        'Si " a la fin => fin d'élément long
        If j = Len(Element(i)) Then LongElement = False
    Next i
    
    CmdLineParam = Trim(CmdLineParam)
        
    'Supprime les guillemets
    If Left(CmdLineParam, 1) = """" And Right(CmdLineParam, 1) = """" Then CmdLineParam = Mid(CmdLineParam, 2, Len(CmdLineParam) - 2)
End Function
