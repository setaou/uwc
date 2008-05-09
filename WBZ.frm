VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form WBZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UWC - Archive WebShots"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "WBZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   510
      Left            =   3210
      TabIndex        =   0
      Top             =   6240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Sauvegarder l'image"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Ajuster à la fenêtre"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Taille réelle"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.ListBox Infos 
      Height          =   855
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   7695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox Preview 
         BorderStyle     =   0  'None
         Height          =   5655
         Left            =   120
         ScaleHeight     =   377
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   497
         TabIndex        =   2
         Top             =   240
         Width           =   7455
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":2A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":32D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":3B28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":4282
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":4994
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":50A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WBZ.frx":57B8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "WBZ"
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

Private w As New cWBZ_File

Private Type tPreview_Infos
    Preview As New cDIBSection
    Original As New cDIBSection
    previewX As Integer
    previewY As Integer
    previewX2 As Integer
    previewY2 As Integer
    Zoom As Boolean
End Type

Private PreviewInfos As tPreview_Infos
Private Sub TranslateForm()
    Me.Caption = "UWC - " + GetTranslation(252, "Image Webshots")
    Frame1.Caption = GetTranslation(200, "Aperçu")
    Toolbar1.Buttons(1).ToolTipText = GetTranslation(209, "Sauvegarder l'image")
    Toolbar1.Buttons(3).ToolTipText = GetTranslation(210, "Ajuster à la fenêtre")
    Toolbar1.Buttons(4).ToolTipText = GetTranslation(211, "Taille réelle")
End Sub

Public Sub Chargement(Fichier As String)
        Dim i As Integer
        Dim Height As Integer, Width As Integer
        Dim Fact1 As Single, Fact2 As Single
        
        If w.OpenFile(Fichier) = 1 Then
        
            Me.Caption = "UWC - " + GetTranslation(251, "Archive Webshots") + " - (" + Fichier + ")"
        
            w.Save_JPG_Picture (GetTempPath() + "temp.jpg")
            
            PreviewInfos.Original.CreateFromPicture LoadPicture(GetTempPath() + "temp.jpg")
            
            Fact1 = Preview.ScaleHeight / PreviewInfos.Original.Height
            Fact2 = Preview.ScaleWidth / PreviewInfos.Original.Width
            
            If Fact1 < Fact2 Then
                Height = PreviewInfos.Original.Height * Fact1
                Width = PreviewInfos.Original.Width * Fact1
            Else
                Height = PreviewInfos.Original.Height * Fact2
                Width = PreviewInfos.Original.Width * Fact2
            End If
            
            Set PreviewInfos.Preview = PreviewInfos.Original.Resample(Height, Width)
            PreviewInfos.Zoom = False
            
            Kill GetTempPath() + "temp.jpg"
            
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(4).Enabled = True
            
            'Infos supplémentaires
            Infos.AddItem GetTranslation(202, "> Informations supplémentaires :")
            If Len(w.PictureTitle) Then Infos.AddItem sprintf(GetTranslation(203, " - Titre : %s"), w.PictureTitle)
            If Len(w.PictureCredit) Then Infos.AddItem sprintf(GetTranslation(205, " - Credits : %s"), w.PictureCredit)
            If Len(w.PictureRootCategory) Then
                Infos.AddItem GetTranslation(208, " - Catégorie : ")
                Infos.AddItem "    o " + w.PictureRootCategory
            End If
            If Len(w.PictureChildCategory) Then
                Infos.AddItem "       '-> " + w.PictureChildCategory
            End If
            If Len(w.PictureID) > 0 Then
                Infos.AddItem sprintf(GetTranslation(206, " - ID : %s"), w.PictureID)
            End If
            
            Infos.AddItem sprintf(GetTranslation(207, " - Dimensions : %dx%d"), PreviewInfos.Original.Width, PreviewInfos.Original.Height)
            Infos.AddItem sprintf(GetTranslation(223, " - Daily Date : %s"), w.PictureDailyDate)
            
            Frame1.Caption = sprintf(GetTranslation(201, "Aperçu : %s") + " (%dx%d)", w.PictureTitle, PreviewInfos.Original.Width, PreviewInfos.Original.Height)

        Else 'Aucune image dans le fichier
        
            MsgBox GetTranslation(213, "Ce fichier n'est pas valide ou ne contient aucune image!"), vbExclamation + vbOKOnly, GetTranslation(214, "Erreur : Fichier incorrect")
            Unload Me
        
        End If
End Sub
Private Function InitTB(Optional intSize As Integer = 24) As Long

    ' Set up the toolbar
    Dim lngStyle As Long
    Dim lRes As Long
    
    ' Get the toolbar handle (we cannot just use tbrMain.hwnd as this is a container
    ' window for the actual toolbar control)
    Dim hTBar As Long
    hTBar = FindWindowEx(Me.Toolbar1.hWnd, 0&, "ToolbarWindow32", vbNullString)
        
    ' The style "TBSTYLE_FLAT" MUST be added.  Although this option is available
    ' in the property pages for the toolbar, it needs to be set here.
    
    ' Get the current style
    lngStyle = SendMessage(hTBar, TB_GETSTYLE, 0&, ByVal 0&)
    
    ' Add the TBSTYLE_FLAT style (could also apply other styles here)
    lngStyle = lngStyle Or TBSTYLE_FLAT
        
    ' Set the new style
    Call SendMessage(hTBar, TB_SETSTYLE, 0&, ByVal lngStyle)
    
    lRes = SendMessage(hTBar, TB_SETIMAGELIST, 0, ByVal ImageList1.hImageList)
    lRes = SendMessage(hTBar, TB_SETHOTIMAGELIST, 0, ByVal ImageList1.hImageList)
    lRes = SendMessage(hTBar, TB_SETDISABLEDIMAGELIST, 0, ByVal ImageList2.hImageList)

    Me.Toolbar1.Refresh
End Function
Private Sub Form_Initialize()
    InitCommonControls
End Sub


Private Sub Form_Load()
    'Toolbar
    InitTB
    
    'Traduction
    TranslateForm
End Sub




Private Sub Preview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    PreviewInfos.previewX2 = x
    PreviewInfos.previewY2 = y
End Sub

Private Sub Preview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button > 0 And PreviewInfos.Zoom = True Then
        PreviewInfos.previewX = PreviewInfos.previewX + (x - PreviewInfos.previewX2)
        PreviewInfos.previewY = PreviewInfos.previewY + (y - PreviewInfos.previewY2)
        
        PreviewInfos.previewX2 = x
        PreviewInfos.previewY2 = y
        
        If PreviewInfos.previewX > 0 Then PreviewInfos.previewX = 0
        If PreviewInfos.previewY > 0 Then PreviewInfos.previewY = 0
        
        If PreviewInfos.previewX < -(PreviewInfos.Original.Width - Preview.ScaleWidth) Then PreviewInfos.previewX = -(PreviewInfos.Original.Width - Preview.ScaleWidth)
        If PreviewInfos.previewY < -(PreviewInfos.Original.Height - Preview.ScaleHeight) Then PreviewInfos.previewY = -(PreviewInfos.Original.Height - Preview.ScaleHeight)
        
        If PreviewInfos.Original.Width < Preview.ScaleWidth Then PreviewInfos.previewX = 0
        If PreviewInfos.Original.Height < Preview.ScaleHeight Then PreviewInfos.previewY = 0
        
        PreviewInfos.Original.PaintPicture Preview.hDC, PreviewInfos.previewX, PreviewInfos.previewY
    End If
End Sub

Private Sub Preview_Paint()
    If PreviewInfos.Zoom = False Then
        PreviewInfos.Preview.PaintPicture Preview.hDC
    Else
        PreviewInfos.Original.PaintPicture Preview.hDC, PreviewInfos.previewX, PreviewInfos.previewY
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.index
        Case 1 'Sauvegarder
        
            Dim cd As New cCommonDialog
            Dim ret As String
            
            If Len(w.PictureTitle) > 0 Then ret = FormatPath(w.PictureTitle)
            
            cd.VBGetSaveFileName ret, , , GetTranslation(253, "Images JPEG") + " (*.jpg)|*.jpg", , , , "jpg"
            If Len(ret) > 0 Then w.Save_JPG_Picture ret
            
        Case 3 'Taille preview
            
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(4).Enabled = True
            PreviewInfos.Zoom = False
            Preview.Refresh
            Preview.MousePointer = 0
            
        Case 4 'taille reelle
            
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(4).Enabled = False
            PreviewInfos.Zoom = True
            Preview.Refresh
            Preview.MousePointer = 15
        
        End Select
End Sub


