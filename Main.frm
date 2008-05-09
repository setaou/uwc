VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultimate Webshots Converter"
   ClientHeight    =   2370
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "Main.frx":1A40A
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Paypal 
      Height          =   540
      Left            =   360
      MouseIcon       =   "Main.frx":3D6CE
      MousePointer    =   99  'Custom
      ToolTipText     =   "Not mandatory, but the author would appreciate it !"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu mnuFichierOuvrir 
         Caption         =   "Ouvrir"
      End
      Begin VB.Menu mnuFichierSepar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuConvert 
      Caption         =   "Conversion"
      Begin VB.Menu mnuMassConvert 
         Caption         =   "Conversion en masse -> JPG"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsLangue 
         Caption         =   "Choisir la langue"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "?"
      Begin VB.Menu mnuAboutApropos 
         Caption         =   "A Propos"
      End
      Begin VB.Menu mnuAproposSepar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAproposWeb 
         Caption         =   "Site WEB"
      End
      Begin VB.Menu mnuAproposContact 
         Caption         =   "Contact / Report de bug"
      End
   End
End
Attribute VB_Name = "MainForm"
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

Private Sub Form_Initialize()
    'Initialisation des Common Controls (pour le style XP)
    InitCommonControls
End Sub

Private Sub Form_Load()
    TranslateForm
End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    If data.GetFormat(15) Then
        For i = 1 To data.Files.Count
            OpenFile data.Files(i)
        Next i
    End If
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If data.GetFormat(15) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub


Private Sub Form_Paint()
    If Len(BETA) Then
        Dim Texte As String
    
        Texte = VersionLong + " " + BETA
    
        MainForm.CurrentX = MainForm.ScaleWidth - MainForm.TextWidth(Texte) - 5
        MainForm.CurrentY = MainForm.ScaleHeight - MainForm.TextHeight(Texte) - 5
    
        MainForm.Print Texte
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub mnuAboutApropos_Click()
    MsgBox sprintf("Ultimate Webshots Converter\n" _
         + GetTranslation(20, "Version %s") + "\n\n" _
         + "(c) 2003-2007 Setaou ( uwc@apinc.org )\n\n" _
         + "http://uwc.apinc.org\n\n" _
         + GetTranslation(12, "Remerciements") + "\n" _
         + "  - Renfield & Yomm (DirExplorer)\n" _
         + "  - VBaccelerator (cDibSection & cProgressbar)\n" _
         + "  - VBnet (cCommonDialogs)\n" _
         + "  - Benoit Frigon (sprintf)\n\n" _
         + "  - " + sprintf(GetTranslation(17, "Traduction en %s : %s"), GetTransMetadata("language"), GetTransMetadata("translator")) _
         , VersionLong + IIf(Len(BETA) > 0, " " + BETA, "")) _
         , vbInformation + vbOKOnly, sprintf(GetTranslation(11), "UWC")
End Sub

Private Sub mnuAproposContact_Click()
    ShellExecute MainForm.hWnd, "open", "mailto:uwc@apinc.org", 0&, 0&, 0&
End Sub

Private Sub mnuAproposWeb_Click()
    'Ouvre le site WEB
    ShellExecute MainForm.hWnd, "open", "http://uwc.apinc.org", 0&, 0&, 0&
End Sub


Private Sub mnuFichierOuvrir_Click()
    Dim cd As New cCommonDialog
    Dim ret As String
    Dim f As Integer
    Dim file_marker As Long
    
    cd.VBGetOpenFileName ret, , , , , True, WebshotsExtensions, , , GetTranslation(10), , MainForm.hWnd
    If Len(ret) > 0 Then OpenFile ret
End Sub

Private Sub mnuFichierQuitter_Click()
    End
End Sub

Private Sub mnuMassConvert_Click()
    MassConvert.Show
End Sub



Private Sub TranslateForm()
    mnuFichier.Caption = GetTranslation(1)
    mnuFichierOuvrir.Caption = GetTranslation(2)
    mnuFichierQuitter.Caption = GetTranslation(3)
    mnuConvert.Caption = GetTranslation(4)
    mnuMassConvert.Caption = GetTranslation(5)
    mnuAboutApropos.Caption = GetTranslation(7, "A Propos")
    mnuOptions.Caption = GetTranslation(15, "Options")
    mnuOptionsLangue.Caption = GetTranslation(16, "Choisir la langue")
    mnuAproposWeb.Caption = GetTranslation(18, "Site Web")
    mnuAproposContact.Caption = GetTranslation(19, "Contact / Report de bug")
End Sub

Private Sub mnuOptionsLangue_Click()
    'Affiche la fenêtre de choix de langue
    ChLang.Show vbModal, MainForm
    
    'Retraduit la fenêtre
    TranslateForm
End Sub



Private Sub Paypal_Click()
    ShellExecute MainForm.hWnd, "open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=uwc%40apinc%2eorg&item_name=Ultimate%20Webshots%20Converter&no_shipping=0&no_note=1&tax=0&currency_code=EUR&bn=PP%2dDonationsBF&charset=UTF%2d8", 0&, 0&, 0&
End Sub
