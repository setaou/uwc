VERSION 5.00
Begin VB.Form ChLang 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UWC - Language choice"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblDate 
      Caption         =   "Date :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblVer 
      Caption         =   "Version :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblTrans 
      Caption         =   "Translator :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Language :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ChLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Type tLanguageInfo
    Language As String
    Version As String
    Translator As String
    Date As String
    File As String
End Type

Private Languages() As tLanguageInfo
Private Sub cmdOk_Click()
    'Sauvegarde
    SaveSetting "UWC", "Language", "Current", Languages(Combo1.ListIndex).File
    
    'Charge
    LoadTranslations Languages(Combo1.ListIndex).File
    
    'Si les versions diffèrent
    If Version <> Languages(Combo1.ListIndex).Version Then MsgBox sprintf(GetTranslation(306, "La version du fichier de langue (%s) est différente de la version de UWC (%s). Certains textes ne seront peut-être pas affichés correctement."), Languages(Combo1.ListIndex).Version, Version), vbInformation
    
    Unload Me
End Sub

Private Sub Combo1_Click()
    'Si une traduction a été chargée
    If TransLoaded Then
        With Languages(Combo1.ListIndex)
            lblVer.Caption = sprintf(GetTranslation(302, "Version : %s"), .Version)
            lblTrans.Caption = sprintf(GetTranslation(303, "Translator : %s"), .Translator)
            lblDate.Caption = sprintf(GetTranslation(304, "Date : %s"), .Date)
        End With
    'Si aucune traduction n'a été chargée
    Else
        With Languages(Combo1.ListIndex)
            lblVer.Caption = sprintf("Version : %s", .Version)
            lblTrans.Caption = sprintf("Translator : %s", .Translator)
            lblDate.Caption = sprintf("Date : %s", .Date)
        End With
    End If
End Sub


Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    Dim fileinfo As WIN32_FIND_DATA, handle As Long, File As String
    Dim LastTransName As String
    
    ReDim Languages(0 To 0)
    
    'Si une langue a été chargée
    If TransLoaded Then
        'Traduit
        TranslateForm
        
        'Sauvegarde le nom de la trans actuelle
        LastTransName = GetTransMetadata("file")
    End If
    
    'Lance la recherche de fichiers de langues
    handle = FindFirstFile(App.path + "\" + "*.lng", fileinfo)
    
    'Si aucun fichier n'est trouvé
    If handle = INVALID_HANDLE_VALUE Then
        MsgBox "No language files were found ! Please reinstall UWC.", vbCritical, "Critical error !"
        End
    End If
    
    Do
        File = Mid(fileinfo.cFileName, 1, InStr(1, fileinfo.cFileName, Chr(0)) - 1)
        
        LoadTranslations File
        
        ReDim Preserve Languages(0 To UBound(Languages) + 1)
        With Languages(UBound(Languages) - 1)
            .Language = GetTransMetadata("language")
            .Version = GetTransMetadata("version")
            .Translator = GetTransMetadata("translator")
            .Date = GetTransMetadata("date")
            .File = File
        End With
        
        Combo1.AddItem GetTransMetadata("language")
        
        'Si le fichier en cours est celui qui était chargé pour la traduction
        'l'activer dans la liste
        If Len(LastTransName) > 0 And (LastTransName = File) Then Combo1.ListIndex = UBound(Languages) - 1
        
    Loop Until FindNextFile(handle, fileinfo) = 0

    FindClose handle
    
    'Reset la traduction
    ClearTranslation
    
    'Si une langue était déjà chargée, la remet
    If Len(LastTransName) > 0 Then
        LoadTranslations LastTransName
    'Sinon on en profite pour sélectionner le 1er élément de la liste
    Else
        Combo1.ListIndex = 0
    End If
End Sub
Private Sub TranslateForm()
    Me.Caption = "UWC - " + GetTranslation(300, "Language choice")
    Label1.Caption = sprintf(GetTranslation(301, "Language : %s"), "")
    lblVer.Caption = sprintf(GetTranslation(302, "Version : %s"), "")
    lblTrans.Caption = sprintf(GetTranslation(303, "Translator : %s"), "")
    lblDate.Caption = sprintf(GetTranslation(304, "Date : %s"), "")
    cmdOk.Caption = GetTranslation(305, "&Ok")
End Sub


