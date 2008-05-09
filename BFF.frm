VERSION 5.00
Object = "{9DD53B4F-5DB5-4C19-AFDC-291A7557D5C8}#16.0#0"; "DirExplo.ocx"
Begin VB.Form BFF 
   Caption         =   "Choisissez le dossier à ajouter"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   Icon            =   "BFF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin DirExplorerOCX.DirExplorer DirExplorer 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8916
   End
   Begin VB.CommandButton Annuler 
      Caption         =   "A&nnuler"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CheckBox Rec 
      Caption         =   "Parcourir les sous-dossiers"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   4095
   End
   Begin VB.Label Choix 
      AutoSize        =   -1  'True
      Caption         =   "Choix"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   390
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BFF"
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

Dim Owner As Form

Private Sub Annuler_Click()
    Me.Tag = False
    Me.Hide
End Sub

Private Sub Ok_Click()
    Me.Tag = True
    Me.Hide
End Sub


Public Function BrowseForFolder(ByRef Folder As String, ByRef Recurse As Boolean, ByRef Owner As Form) As Boolean
    DirExplorer.Chemin = Folder
    UpdateLabel
    
    If Recurse Then Rec.Value = vbChecked Else Rec.Value = vbUnchecked
    
    Me.Show vbModal, Owner
    
    'Le résultat est valide seulement si le bouton Ok a été pressé et si le dossier choisi n'est pas nul
    BrowseForFolder = (CBool(Me.Tag) And (Len(DirExplorer.Chemin) > 0))
    
    If BrowseForFolder Then
        Folder = DirExplorer.Chemin
        If Rec.Value = vbChecked Then Recurse = True Else Recurse = False
    End If
    
    Unload Me
End Function

Private Sub DirExplorer_Click()
    UpdateLabel
    
    Form_Resize
End Sub


Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    'Traduction
    TranslateForm

    Me.Tag = False
End Sub

Private Sub Form_Resize()
    Dim l_minimum, h_minimum, h_minimum_explorateur

    h_minimum_explorateur = 3000
    
    l_minimum = Annuler.Width + Ok.Width + 3 * CTRL_ESPACEMENT
    h_minimum = Annuler.Height + Rec.Height + Choix.Height + h_minimum_explorateur + 3 * CTRL_ESPACEMENT

    If BFF.ScaleWidth < l_minimum Then BFF.Width = l_minimum + (BFF.Width - BFF.ScaleWidth)
    If BFF.ScaleHeight < h_minimum Then BFF.Height = h_minimum + (BFF.Height - BFF.ScaleHeight)
    
    Ok.Top = BFF.ScaleHeight - Ok.Height - CTRL_ESPACEMENT
    Annuler.Top = BFF.ScaleHeight - Annuler.Height - CTRL_ESPACEMENT
    Ok.Left = BFF.ScaleWidth - Ok.Width - CTRL_ESPACEMENT
    Annuler.Left = Ok.Left - Annuler.Width - CTRL_ESPACEMENT
    
    Rec.Left = CTRL_ESPACEMENT
    Rec.Top = Ok.Top - Rec.Height - CTRL_ESPACEMENT
    
    Choix.Left = CTRL_ESPACEMENT
    Choix.Width = BFF.ScaleWidth - 2 * CTRL_ESPACEMENT
    Choix.Top = Rec.Top - Choix.Height - CTRL_ESPACEMENT
    
    DirExplorer.Width = BFF.ScaleWidth
    DirExplorer.Height = Choix.Top - CTRL_ESPACEMENT
End Sub

Public Sub UpdateLabel()
    If Len(DirExplorer.Chemin) > 0 Then
        Choix.Caption = DirExplorer.Chemin
    Else
        Choix.Caption = GetTranslation(51)
    End If
End Sub

Private Sub TranslateForm()
    Me.Caption = GetTranslation(50)
    Rec.Caption = GetTranslation(52)
    Annuler.Caption = GetTranslation(53)
    Ok.Caption = GetTranslation(54)
End Sub
