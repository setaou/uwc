VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form WBC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UWC - Collection WebShots"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   510
      Left            =   5070
      TabIndex        =   2
      Top             =   5460
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Sauvegarder toutes les images"
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
            Object.ToolTipText     =   "Sauvegarder l'image"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Ajuster à la fenêtre"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Taille réelle"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Liste des images"
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   6345
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   5295
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox preview 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   329
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   449
         TabIndex        =   4
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.ListBox Infos 
      Height          =   855
      IntegralHeight  =   0   'False
      Left            =   2640
      TabIndex        =   1
      Top             =   6000
      Width           =   6975
   End
   Begin VB.PictureBox progressbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   9495
      TabIndex        =   0
      Top             =   6960
      Width           =   9495
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   8880
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3196
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":38A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3FBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4714
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5538
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5D8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":65DC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "WBC"
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

Private progress As New cProgressBar
Private w As New cWBC_File

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
    Me.Caption = "UWC - " + GetTranslation(250, "Collection Webshots")
    Frame1.Caption = GetTranslation(200, "Aperçu")
    Frame2.Caption = GetTranslation(215, "Liste des images")
    Toolbar1.Buttons(1).ToolTipText = GetTranslation(212, "Sauvegarder toutes les images")
    Toolbar1.Buttons(3).ToolTipText = GetTranslation(209, "Sauvegarder l'image")
    Toolbar1.Buttons(5).ToolTipText = GetTranslation(210, "Ajuster à la fenêtre")
    Toolbar1.Buttons(6).ToolTipText = GetTranslation(211, "Taille réelle")
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
    'Style XP Progressbar
    With progress
        If WindowsVersion = "Windows XP" Or WindowsVersion = "Windows Vista" Then
            .XpStyle = True
        Else
            .BarColor = RGB(180, 200, 250)
            .Segments = True
            .BorderStyle = EVPRGInset
            .ForeColor = vbBlack
            .BarForeColor = vbBlack
        End If
        .DrawObject = progressbar
        .Min = 0
        .Max = 100
        .Value = 0
        .ShowText = True
    End With
    
    'Toolbar
    InitTB
    
    'Traduction
    TranslateForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set w = Nothing
End Sub

Private Sub List1_Click()
    Dim Height As Integer, Width As Integer
    Dim Fact1 As Single, Fact2 As Single
    
    'Si on a une image sélectionnée
    If List1.ListIndex >= 0 Then
        
        'Extraction
        progress.Min = 0
        progress.Max = 3
        progress.Text = GetTranslation(217, "Extraction ...")
        
        w.Picture(List1.ListIndex + 1).Save_JPG_Picture GetTempPath() + "temp.jpg"
        
        'Chargement
        progress.Value = 1
        progress.Text = GetTranslation(218, "Chargement ...")
        
        PreviewInfos.Original.CreateFromPicture LoadPicture(GetTempPath() + "temp.jpg")
        
        'Redimentionnement de la preview
        progress.Value = 2
        progress.Text = GetTranslation(219, "Redimentionnement ...")
        
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
        
        'Affichage
        progress.Value = 3
        progress.Text = GetTranslation(220, "Affichage ...")
        
        Preview.Cls
        PreviewInfos.Preview.PaintPicture Preview.hDC
        Kill GetTempPath() + "temp.jpg"
        
        progress.Value = 0
        progress.Text = ""
        
        'Infos sur l'image
        Infos.Clear
        Infos.AddItem GetTranslation(202, "> Informations supplémentaires :")
        If Len(w.Picture(List1.ListIndex + 1).Image_Title) Then Infos.AddItem sprintf(GetTranslation(203, " - Titre : %s"), w.Picture(List1.ListIndex + 1).Image_Title)
        If Len(w.Picture(List1.ListIndex + 1).Image_Description) Then Infos.AddItem sprintf(GetTranslation(204, " - Description : %s"), w.Picture(List1.ListIndex + 1).Image_Description)
        If Len(w.Picture(List1.ListIndex + 1).Image_Credits) Then Infos.AddItem sprintf(GetTranslation(205, " - Credits : %s"), w.Picture(List1.ListIndex + 1).Image_Credits)
        If Len(w.Picture(List1.ListIndex + 1).Root_Category) Then
            Infos.AddItem GetTranslation(208, " - Catégorie : ")
            Infos.AddItem "    o " + w.Picture(List1.ListIndex + 1).Root_Category
        End If
        If Len(w.Picture(List1.ListIndex + 1).Child_Category) Then
            Infos.AddItem "       '-> " + w.Picture(List1.ListIndex + 1).Child_Category
        End If
        If Len(w.Picture(List1.ListIndex + 1).Picture_ID) Then
            Infos.AddItem sprintf(GetTranslation(206, " - ID : %s"), w.Picture(List1.ListIndex + 1).Picture_ID)
        End If
        
        Infos.AddItem sprintf(GetTranslation(207, " - Dimensions : %dx%d"), PreviewInfos.Original.Width, PreviewInfos.Original.Height)
        Infos.AddItem sprintf(GetTranslation(223, " - Daily Date : %s"), w.Picture(List1.ListIndex + 1).Daily_Date)
        
        Frame1.Caption = sprintf(GetTranslation(201, "Aperçu : %s") + " (%dx%d)", w.Picture(List1.ListIndex + 1).Image_Title, PreviewInfos.Original.Width, PreviewInfos.Original.Height)
        
        PreviewInfos.Zoom = False
        PreviewInfos.previewX = 0
        PreviewInfos.previewY = 0
        Preview.MousePointer = 0
        
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = True
    End If
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
    If PreviewInfos.Zoom = True Then
        PreviewInfos.Original.PaintPicture Preview.hDC, PreviewInfos.previewX, PreviewInfos.previewY
    Else
        PreviewInfos.Preview.PaintPicture Preview.hDC
    End If
End Sub


Private Sub progressbar_Paint()
    progress.Draw
End Sub

Public Sub Chargement(Fichier As String)
        Dim i As Integer
    
        w.OpenFile Fichier
        Me.Caption = "UWC - " + GetTranslation(250, "Collection Webshots") + " - " + w.File_Title + " (" + Fichier + ")"
        
        If w.PictureCount > 0 Then 'Il y a des images
            
            progress.Max = w.PictureCount
            progress.Min = 0
            For i = 1 To w.PictureCount
                progress.Text = sprintf(GetTranslation(216, "Image n°%d"), i)
                progress.Value = i
                List1.AddItem w.Picture(i).Image_Title
            Next i
            
            progress.Value = 0
            progress.Text = ""
            
            Toolbar1.Buttons(1).Enabled = True
            
        Else 'Aucune image dans le fichier
        
            MsgBox GetTranslation(213, "Ce fichier n'est pas valide ou ne contient aucune image!"), vbExclamation + vbOKOnly, GetTranslation(214, "Erreur : Fichier incorrect")
            Unload Me
        
        End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.index
        Case 1 'Sauvegarder tout
            Dim i As Integer, b As New cBrowseForFolder, fold As String
            
            Call b.Display(Me.hWnd, GetTranslation(221, "Choisissez l'emplacement où sauvegarder l'intégralité des images de la collection."))
            
            If b.successful Then
                progress.Max = w.PictureCount
                progress.Min = 0
                
                fold = b.folderName
                If Right(fold, 1) <> "\" Then fold = fold + "\"
                
                For i = 1 To w.PictureCount
                    progress.Text = sprintf(GetTranslation(222, "Sauvegarde : %s ..."), w.Picture(i).Image_Title)
                    progress.Value = i
                    w.Picture(i).Save_JPG_Picture fold + w.Picture(i).Image_Title + ".jpg"
                Next i
                
                progress.Value = 0
                progress.Text = ""
            End If
        Case 3 'Sauvegarder
            Dim cd As New cCommonDialog
            Dim ret As String
            
            ret = FormatPath(w.Picture(List1.ListIndex + 1).Image_Title) & ".jpg"
            
            cd.VBGetSaveFileName ret, , , GetTranslation(253, "Images JPEG") + " (*.jpg)|*.jpg", , , , "jpg"
            If ret <> "" Then w.Picture(List1.ListIndex + 1).Save_JPG_Picture ret
        Case 5 'Stretch
            Preview.Cls
            PreviewInfos.Preview.PaintPicture Preview.hDC
            PreviewInfos.Zoom = False
            Preview.MousePointer = 0
            Toolbar1.Buttons(5).Enabled = False
            Toolbar1.Buttons(6).Enabled = True
        Case 6 'Taille reelle
            Preview.Cls
            PreviewInfos.Original.PaintPicture Preview.hDC, PreviewInfos.previewX, PreviewInfos.previewY
            PreviewInfos.Zoom = True
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(5).Enabled = True
            Preview.MousePointer = 15
    End Select
End Sub

