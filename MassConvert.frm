VERSION 5.00
Begin VB.Form MassConvert 
   Caption         =   "UWC - Conversion en masse"
   ClientHeight    =   6615
   ClientLeft      =   2580
   ClientTop       =   2475
   ClientWidth     =   9495
   Icon            =   "MassConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_dest 
      Caption         =   "Destination"
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   4695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4455
         TabIndex        =   19
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton Command7 
            Caption         =   "+"
            Enabled         =   0   'False
            Height          =   285
            Left            =   4140
            TabIndex        =   25
            Top             =   720
            Width           =   315
         End
         Begin VB.ComboBox ParamDest 
            Height          =   315
            ItemData        =   "MassConvert.frx":2372
            Left            =   0
            List            =   "MassConvert.frx":2374
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   720
            Width           =   4095
         End
         Begin VB.CheckBox MemeDossier 
            Caption         =   "Même dossier que l'original"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   360
            Width           =   4455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   285
            Left            =   4140
            TabIndex        =   21
            Top             =   0
            Width           =   315
         End
         Begin VB.TextBox Destination 
            Height          =   285
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox progressbar 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   9255
      TabIndex        =   16
      Top             =   6240
      Width           =   9255
   End
   Begin VB.Frame frm_log 
      Caption         =   "Log"
      Height          =   3255
      Left            =   4920
      TabIndex        =   14
      Top             =   2880
      Width           =   4455
      Begin VB.TextBox Log 
         Height          =   2895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frm_para 
      Caption         =   "Paramètres"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   4695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4530
         TabIndex        =   15
         Top             =   240
         Width           =   4530
         Begin VB.CommandButton Command8 
            Caption         =   "+"
            Enabled         =   0   'False
            Height          =   285
            Left            =   4140
            TabIndex        =   26
            Top             =   0
            Width           =   315
         End
         Begin VB.ComboBox ParamFichier 
            Height          =   315
            ItemData        =   "MassConvert.frx":2376
            Left            =   0
            List            =   "MassConvert.frx":2378
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Width           =   4095
         End
         Begin VB.OptionButton FAE 
            Caption         =   "Renommer"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton FAE 
            Caption         =   "Ecraser"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   7
            Top             =   600
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton FAE 
            Caption         =   "Passer"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Convertir !"
            Height          =   375
            Left            =   2760
            TabIndex        =   9
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Si un fichier existe déjà :"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   360
            Width           =   4455
         End
      End
   End
   Begin VB.Frame frm_liste 
      Caption         =   "Fichiers"
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9255
      Begin VB.PictureBox frm_liste_pic 
         BorderStyle     =   0  'None
         Height          =   2370
         Left            =   120
         ScaleHeight     =   2370
         ScaleWidth      =   9075
         TabIndex        =   12
         Top             =   240
         Width           =   9075
         Begin VB.CheckBox EviterDoublons 
            Caption         =   "Eviter les doublons"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   2040
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Effacer"
            Height          =   375
            Left            =   8160
            TabIndex        =   4
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "+ &Fichier"
            Height          =   375
            Left            =   8160
            TabIndex        =   2
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "- F&ichier(s)"
            Height          =   375
            Left            =   8160
            TabIndex        =   3
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "+ &Dossier"
            Height          =   375
            Left            =   8160
            TabIndex        =   1
            Top             =   0
            Width           =   855
         End
         Begin VB.ListBox List1 
            Height          =   2025
            IntegralHeight  =   0   'False
            Left            =   0
            MultiSelect     =   2  'Extended
            OLEDropMode     =   1  'Manual
            TabIndex        =   0
            Top             =   0
            Width           =   8055
         End
      End
   End
   Begin VB.Menu menu_popup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_ttselect 
         Caption         =   "Tout sélectionner"
      End
      Begin VB.Menu mnu_invselect 
         Caption         =   "Inverser la sélection"
      End
      Begin VB.Menu mnu_separ 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_filtres 
         Caption         =   "Filtres"
         Begin VB.Menu mnu_filtre_vignettes 
            Caption         =   "Supprimer les vignettes"
         End
         Begin VB.Menu mnu_filtre_doublons 
            Caption         =   "Supprimer les doublons (bientôt)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_separ3 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_filtre_perso 
            Caption         =   "Personnalisé (bientôt)"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnu_separ2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ouvrir 
         Caption         =   "Ouvrir"
      End
      Begin VB.Menu mnu_supprimer 
         Caption         =   "Supprimer"
      End
   End
End
Attribute VB_Name = "MassConvert"
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

Private cd As New cCommonDialog
Private progress As New cProgressBar

Private LastBFFPath As String
Private ConvEnCours As Boolean

Private ExprCheminPerso As String
Private ExprFichierPerso As String

Private Type ConvStats
    Fichiers As Integer
    Images As Integer
    Convertis As Integer
    Renommes As Integer
    Passes As Integer
End Type
'Ajout du contenu d'un dossier à la liste
Private Sub Command1_Click()

    Dim Doss As String, Fichier As String, Recurse As Boolean
    
    'Reprise du dernier dossier parcouru
    Doss = LastBFFPath
    
    'Choix d'un dossier -> la suite n'est executée que si le dossier choisi est valide
    'et que l'utilisateur n'a pas cliqué sur Annuler
    If BFF.BrowseForFolder(Doss, Recurse, Me) Then
    
        LastBFFPath = Doss
        
        'Sauvegarde du dernier dossier parcouru dans la base de registre
        SaveSetting "UWC", "Batch Conversion", "Folder", LastBFFPath
        
        If Recurse Then 'Parcourir les sous-dossiers
            'Affichage du texte de progressbar
            progress.ShowText = True
            'Lancement de l'ajout récursif
            AjoutDossierRecurse Doss, (EviterDoublons.Value = vbUnchecked)
            'Affichage du texte de progressbar désactivé
            progress.Text = ""
            progress.ShowText = False
        Else 'Ajouter seulement un dossier
            AjoutDossier Doss, (EviterDoublons.Value = vbUnchecked)
        End If
        
        'Mise à jour de la liste
        lstAddHScroll List1
        List1.SetFocus
    End If
End Sub

'Suppression des fichiers sélectionnés dans la liste
Private Sub Command2_Click()
    Dim k As Integer
    
    k = 0
    
    Do Until k + 1 > List1.ListCount
        If List1.Selected(k) = True Then
            List1.RemoveItem k
        Else
            k = k + 1
        End If
    Loop
    
    'Mise a jour de la liste (scrollbar)
    lstAddHScroll List1
    List1.SetFocus
End Sub


'Choix du dossier de destination
Private Sub Command3_Click()
    Dim b As New cBrowseForFolder
    
    Call b.Display(Me.hWnd, GetTranslation(130))
    
    If b.successful Then
        Destination.Text = b.folderName
        If Right(Destination.Text, 1) <> "\" Then Destination.Text = Destination.Text + "\"
    End If
End Sub

'Conversion des fichiers de la liste
Private Sub Command4_Click()
    Dim i As Integer, j As Integer
    Dim Fichier As String, Dossier As String
    Dim NouveauFichier As String, NouveauDossier As String
    
    Dim FileAlreadyExists As Integer
    
    Dim PhotosTXT As New cPhotos_Txt_File
    Dim WB1 As cWB1_File
    Dim WBC As cWBC_File
    Dim WBZ As cWBZ_File
    
    Dim Stats As ConvStats
    
    Dim ExprChemin As String
    Dim ExprFichier As String
    
    ''''
    ' Si une conversion est en cours => Gestion de l'arrêt utilisateur
    ''
    
    If ConvEnCours Then
        ConvEnCours = False
        
        Command4.Caption = GetTranslation(123)
        Exit Sub
    End If
    
    ''''
    ' Récupération des paramètres et tests de validité
    ''
    
    'Modèles de nommage
    If MemeDossier.Value = vbUnchecked Then
        Select Case ParamDest.ListIndex
            Case 0
                ExprChemin = "(%dest.path)"
            Case 1
                ExprChemin = "(%dest.path((%col.title|%col.filename)\)|%dest.path)"
            Case 2
                ExprChemin = "(%dest.path((%pic.cat.root\%pic.cat.child|%pic.cat.root|%pic.cat.child)\)|%dest.path)"
            Case 3
                ExprChemin = ExprCheminPerso
        End Select
    Else
        ExprChemin = "(%file.path)"
    End If
    
    Select Case ParamFichier.ListIndex
        Case 0
            ExprFichier = "(%file.name).jpg"
        Case 1
            ExprFichier = "(%pic.title|%file.name).jpg"
        Case 2
            ExprFichier = "(%pic.id|%file.name).jpg"
        Case 3
            ExprFichier = ExprFichierPerso
    End Select
    
    'Récupère le dossier de destination
    Dossier = Destination.Text
    
    'Récupère la réaction en cas de fichier deja existant
    If FAE(FAE_SKIP).Value Then FileAlreadyExists = FAE_SKIP
    If FAE(FAE_OVERWRITE).Value Then FileAlreadyExists = FAE_OVERWRITE
    If FAE(FAE_RENAME).Value Then FileAlreadyExists = FAE_RENAME
    
    'Si aucun fichier dans la liste
    If List1.ListCount = 0 Then
        MsgBox GetTranslation(125), vbExclamation + vbOKOnly, GetTranslation(129)
        Exit Sub
    End If
    
    If MemeDossier.Value = vbUnchecked Then
        'Si pas de sossier de destination
        If Len(Destination) = 0 Then
            MsgBox GetTranslation(126), vbExclamation + vbOKOnly, GetTranslation(129)
            Exit Sub
        End If
        
        'Ajout d'un "\" de fin au dossier de destination s'il n'est pas présent
        If Right(Destination, 1) <> "\" Then Destination = Destination + "\"
        
        'Si le dossier de destination n'existe pas
        If Not DirectoryExists(Destination) Then
            MsgBox GetTranslation(127), vbExclamation + vbOKOnly, GetTranslation(129)
            Exit Sub
        End If
    End If
    
    ''''
    ' A partir d'ici la conversion commence vraiment
    ''
    
    'Notification que la conversion est en cours
    ConvEnCours = True
    
    'Changement de texte du bouton de conversion
    Command4.Caption = GetTranslation(124)
    
    'Désactivation de tous les contrôles
    EnableAll False
    
    'Initialisation de la barre de progression
    progress.Min = 0
    progress.Max = List1.ListCount

    'Initialisation du log
    Log.Text = ""
    AddToLog GetTranslation(150) + vbCrLf
    
    'Pour chaque element de la liste
    For i = 0 To List1.ListCount - 1
    
        'Gestion de l'arrêt utilisateur en cours de conversion
        If ConvEnCours = False Then
            AddToLog GetTranslation(151)
            Exit For
        End If
    
        'MAJ de la barre de progression
        progress.Value = i + 1
        
        'Traitement des évènements de la fenêtre
        DoEvents
        
        'Récupère le nom du fichier en cours
        Fichier = List1.List(i)
        
        'Log
        AddToLog sprintf(GetTranslation(152), GetFileName(Fichier))
        
        'Selon l'extension
        Select Case LCase(Right(Fichier, 3))
            
            Case "wbd", "wb1", "wb0" ' **** WB1, WBD, WB0 ****
                Set WB1 = New cWB1_File
                
                If WB1.OpenFile(Fichier) = 1 Then
                    'Stats
                    Stats.Images = Stats.Images + 1
                    
                    'Cherche les fichiers Photos.txt/Album.txt et les ouvre s'ils ne sont pas déjà ouverts
                    PhotosTXT.AutoOpen GetDirectory(Fichier)

                    'Cherche la photo dans le fichier photos.txt/album.txt
                    PhotosTXT.SeekPicture GetFileName(Fichier)
                    
                    'Mise à jour des variables
                    With EvalExpr
                        .ClearVars
                        .SetVar "dest.path", Dossier
                        .SetVar "file.name", GetFileName(Fichier)
                        .SetVar "file.path", GetDirectory(Fichier)
                        .SetVar "pic.title", FormatPath(PhotosTXT.PictureTitle)
                        .SetVar "pic.id", FormatPath(PhotosTXT.PictureID)
                        .SetVar "pic.descr", FormatPath(PhotosTXT.PictureCaption)
                        .SetVar "pic.cat.child", FormatPath(PhotosTXT.AlbumTitle)
                        .SetVar "pic.cat.root", FormatPath(PhotosTXT.AlbumTopicName)
                    End With
                    
                    'Evaluation des nouveaux noms de dossier/fichier
                    NouveauDossier = EvalExpr.Eval(ExprChemin, True)
                    NouveauFichier = EvalExpr.Eval(ExprFichier, True)
                    
                    'Création du dossier s'il n'existe pas
                    If NouveauDossier <> Dossier Then CreateFolder (NouveauDossier)
                    
                    'Si le fichier doit être écrit (cf. GestionFAE)
                    If GestionFAE(FileAlreadyExists, NouveauDossier, NouveauFichier, Stats) Then
                        'Enregistrement
                        WB1.Save_JPG_Picture NouveauDossier + NouveauFichier
                        
                        'Log
                        AddToLog sprintf(GetTranslation(153), NouveauDossier + NouveauFichier) + vbCrLf
                    'Si le fichier doit être passé
                    Else
                        'Log
                        AddToLog sprintf(GetTranslation(154), NouveauDossier + NouveauFichier) + vbCrLf
                    End If
                   
                Else
                    
                    'Echec
                    AddToLog GetTranslation(155) + vbCrLf
                    
                End If
                
                Set WB1 = Nothing
            
            Case "wbz" '**** WBZ ****
                Set WBZ = New cWBZ_File
                
                If WBZ.OpenFile(Fichier) = 1 Then
                    'Stats
                    Stats.Images = Stats.Images + 1
                    
                    'Mise à jour des variables
                    With EvalExpr
                        .ClearVars
                        .SetVar "dest.path", Dossier
                        .SetVar "file.name", GetFileName(Fichier)
                        .SetVar "file.path", GetDirectory(Fichier)
                        .SetVar "pic.title", FormatPath(WBZ.PictureTitle)
                        .SetVar "pic.id", FormatPath(WBZ.PictureID)
                        .SetVar "pic.descr", FormatPath(WBZ.Description)
                        .SetVar "pic.cat.child", FormatPath(WBZ.PictureChildCategory)
                        .SetVar "pic.cat.root", FormatPath(WBZ.PictureRootCategory)
                        .SetVar "pic.credit", FormatPath(WBZ.PictureCredit)
                        .SetVar "pic.dailydate", FormatPath(WBZ.PictureDailyDate)
                    End With
                    
                    'Evaluation des nouveaux noms de dossier/fichier
                    NouveauDossier = EvalExpr.Eval(ExprChemin, True)
                    NouveauFichier = EvalExpr.Eval(ExprFichier, True)
                    
                    'Création du dossier s'il n'existe pas
                    If NouveauDossier <> Dossier Then CreateFolder (NouveauDossier)
                    
                    'Si le fichier doit être écrit (cf. GestionFAE)
                    If GestionFAE(FileAlreadyExists, NouveauDossier, NouveauFichier, Stats) Then
                        'Enregistrement
                        WBZ.Save_JPG_Picture NouveauDossier + NouveauFichier
                        
                        'Log
                        AddToLog sprintf(GetTranslation(153), NouveauDossier + NouveauFichier) + vbCrLf
                    'Si le fichier doit être passé
                    Else
                        'Log
                        AddToLog sprintf(GetTranslation(154), NouveauDossier + NouveauFichier) + vbCrLf
                    End If
                    
                Else
                    
                    'Echec
                    AddToLog GetTranslation(155) + vbCrLf
                
                End If
                
                Set WBZ = Nothing
            
            Case "wbc", "wbp" '**** WBC, WBP ****
                Set WBC = New cWBC_File
                
                If WBC.OpenFile(Fichier) = 1 Then
                    
                    'Pour chaque image dans la collection
                    For j = 1 To WBC.PictureCount
                        'Stats
                        Stats.Images = Stats.Images + 1
                        
                        'Mise à jour des variables
                        With EvalExpr
                            .ClearVars
                            .SetVar "dest.path", Dossier
                            .SetVar "file.name", WBC.Picture(j).Original_Filename
                            .SetVar "file.path", GetDirectory(Fichier)
                            .SetVar "pic.title", FormatPath(WBC.Picture(j).Image_Title)
                            .SetVar "pic.id", FormatPath(WBC.Picture(j).Picture_ID)
                            .SetVar "pic.descr", FormatPath(WBC.Picture(j).Image_Description)
                            .SetVar "pic.cat.child", FormatPath(WBC.Picture(j).Child_Category)
                            .SetVar "pic.cat.root", FormatPath(WBC.Picture(j).Root_Category)
                            .SetVar "pic.credit", FormatPath(WBC.Picture(j).Image_Credits)
                            .SetVar "pic.dailydate", FormatPath(WBC.Picture(j).Daily_Date)
                            .SetVar "col.title", FormatPath(WBC.File_Title)
                            .SetVar "col.filename", GetFileName(Fichier)
                        End With
                        
                        'Evaluation des nouveaux noms de dossier/fichier
                        NouveauDossier = EvalExpr.Eval(ExprChemin, True)
                        NouveauFichier = EvalExpr.Eval(ExprFichier, True)
                        
                        'Création du dossier s'il n'existe pas
                        If NouveauDossier <> Dossier Then CreateFolder (NouveauDossier)
                        
                        'Si le fichier doit être écrit (cf. GestionFAE)
                        If GestionFAE(FileAlreadyExists, NouveauDossier, NouveauFichier, Stats) Then
                            'Enregistrement
                            WBC.Picture(j).Save_JPG_Picture NouveauDossier + NouveauFichier
                            
                            'Log
                            AddToLog sprintf(GetTranslation(160), WBC.Picture(j).Image_Title, NouveauDossier + NouveauFichier)
                        'Si le fichier doit être passé
                        Else
                            'Log
                            AddToLog sprintf(GetTranslation(161), WBC.Picture(j).Image_Title, NouveauDossier + NouveauFichier)
                        End If
                    Next j
                    
                    'Terminé
                    AddToLog GetTranslation(159) + vbCrLf
                Else
                
                    'Echec
                    AddToLog GetTranslation(155) + vbCrLf
                    
                End If
                
                Set WBC = Nothing
        
        End Select
    Next i
    
    'Log + Statistiques
    AddToLog GetTranslation(156) + vbCrLf
    
    AddToLog sprintf(GetTranslation(157), List1.ListCount, Stats.Images)
    AddToLog sprintf(GetTranslation(158), Stats.Convertis, Stats.Renommes, Stats.Passes)
    
    'Mise a jour de la barre de progression
    progress.Value = 0

    'Activation de tous les contrôles
    EnableAll True
    
    'Notification de la fin de la conversion
    ConvEnCours = False

    'Mise a jour du texte du bouton start/stop
    Command4.Caption = GetTranslation(123)
End Sub

'Effacer toute la liste
Private Sub Command5_Click()
    List1.Clear
    lstAddHScroll List1
    List1.SetFocus
End Sub

'Ajout d'un fichier a la liste
Private Sub Command6_Click()
    Dim ret As String

    cd.VBGetOpenFileName ret, , , , , True, WebshotsExtensions, , , GetTranslation(131), , MassConvert.hWnd
    
    'Si un fichier a été choisi, ajoute celui-ci a la liste
    If Len(ret) > 0 Then AjoutFichier ret, (EviterDoublons.Value = vbUnchecked)
    
    'Mise à jour de la liste
    lstAddHScroll List1
    List1.SetFocus
End Sub


Private Sub Command7_Click()
    If ExprEdit.EditExpr(ExprCheminPerso, MassConvert) Then SaveSetting "UWC", "Batch Conversion", "DestCustom", ExprCheminPerso
End Sub

Private Sub Command8_Click()
    If ExprEdit.EditExpr(ExprFichierPerso, MassConvert) Then SaveSetting "UWC", "Batch Conversion", "FileCustom", ExprFichierPerso
End Sub


Private Sub Destination_Change()
    'Sauvegarde de l'état dans la base de registre
    SaveSetting "UWC", "Batch Conversion", "DestPath", Destination.Text
End Sub

Private Sub EviterDoublons_Click()
    'Sauvegarde de l'état dans la base de registre
    SaveSetting "UWC", "Batch Conversion", "Doubles", IIf(EviterDoublons.Value = vbChecked, 1, 0)
End Sub

Private Sub Form_Initialize()
    'Initialisation des common controls
    InitCommonControls
End Sub

Private Sub Form_Load()
    'Initialisation de la barre de progression, application des styles XP si le système le permet.
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
    
    'Traduction
    TranslateForm
    
    'Chargement de l'état de la checkbox Eviter les doublons
    EviterDoublons.Value = IIf(GetSetting("UWC", "Batch Conversion", "Doubles", "1") = "1", vbChecked, vbUnchecked)
    LastBFFPath = GetSetting("UWC", "Batch Conversion", "Folder", "")
    
    ParamDest.ListIndex = GetSetting("UWC", "Batch Conversion", "Dest", 0)
    ParamFichier.ListIndex = GetSetting("UWC", "Batch Conversion", "File", 0)

    Destination.Text = GetSetting("UWC", "Batch Conversion", "DestPath", "")

    ExprFichierPerso = GetSetting("UWC", "Batch Conversion", "FileCustom", "(%pic.title|%file.name).jpg")
    ExprCheminPerso = GetSetting("UWC", "Batch Conversion", "DestCustom", "(%dest.path)")
End Sub
Private Sub Form_Resize()
    'Redimentionnement des controles
    Dim l_minimum As Integer, h_minimum As Integer

    If MassConvert.WindowState = vbMinimized Then Exit Sub

    l_minimum = 10000
    h_minimum = 7000

    If MassConvert.ScaleWidth < l_minimum Then MassConvert.Width = l_minimum + (MassConvert.Width - MassConvert.ScaleWidth)
    If MassConvert.ScaleHeight < h_minimum Then MassConvert.Height = h_minimum + (MassConvert.Height - MassConvert.ScaleHeight)

    progressbar.Left = CTRL_ESPACEMENT
    progressbar.Top = MassConvert.ScaleHeight - progressbar.Height - CTRL_ESPACEMENT
    progressbar.Width = MassConvert.ScaleWidth - 2 * CTRL_ESPACEMENT
    progressbar.Refresh
    
    frm_log.Top = progressbar.Top - frm_log.Height - CTRL_ESPACEMENT
    frm_log.Width = MassConvert.ScaleWidth - frm_para.Width - 3 * CTRL_ESPACEMENT
        Log.Top = 2 * CTRL_ESPACEMENT
        Log.Left = CTRL_ESPACEMENT
        Log.Width = frm_log.Width - 2 * CTRL_ESPACEMENT
        Log.Height = frm_log.Height - 3 * CTRL_ESPACEMENT
    
    frm_dest.Top = frm_log.Top
    frm_para.Top = progressbar.Top - frm_para.Height - CTRL_ESPACEMENT
    
    frm_liste.Width = MassConvert.ScaleWidth - 2 * CTRL_ESPACEMENT
    frm_liste.Height = frm_log.Top - 2 * CTRL_ESPACEMENT
        frm_liste_pic.Top = CTRL_ESPACEMENT * 2
        frm_liste_pic.Left = CTRL_ESPACEMENT
        frm_liste_pic.Height = frm_liste.Height - 3 * CTRL_ESPACEMENT
        frm_liste_pic.Width = frm_liste.Width - 2 * CTRL_ESPACEMENT
            EviterDoublons.Top = frm_liste_pic.ScaleHeight - EviterDoublons.Height
            List1.Top = 0
            List1.Left = 0
            List1.Height = EviterDoublons.Top - CTRL_ESPACEMENT
            List1.Width = frm_liste_pic.ScaleWidth - Command1.Width - CTRL_ESPACEMENT
            Command1.Top = 0
            Command1.Left = frm_liste_pic.ScaleWidth - Command1.Width
            Command6.Top = Command1.Top + Command1.Height + CTRL_ESPACEMENT
            Command6.Left = Command1.Left
            Command5.Top = frm_liste_pic.ScaleHeight - Command5.Height
            Command5.Left = Command1.Left
            Command2.Top = Command5.Top - Command2.Height - CTRL_ESPACEMENT
            Command2.Left = Command1.Left
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Gestion du CTRL+A
    If KeyPressed(vbKeyControl) And KeyPressed(vbKeyA) Then
        ListeTtSelect
    End If
    
    'Gestion de la touche Suppr
    If KeyCode = vbKeyDelete Then
        ListeSupprSelect
    End If
End Sub




Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Menu contextuel sur clic droit ds la liste
    If Button = 2 And Len(List1.Text) Then PopupMenu menu_popup
End Sub

'Ajout de fichiers a la liste par drag'n'drop
Private Sub List1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If data.GetFormat(15) Then
        For i = 1 To data.Files.Count
            'Si c'est un dossier, l'ajoute avec tous ses sous dossiers
            If DirectoryExists(data.Files(i)) Then
                'Affichage du texte de progressbar
                progress.ShowText = True
                'Lancement de l'ajout récursif
                AjoutDossierRecurse data.Files(i), (EviterDoublons.Value = vbUnchecked)
                'Affichage du texte de progressbar désactivé
                progress.ShowText = False
            'Si c'est un fichier
            Else
                'Avec la bonne extension
                Select Case LCase(Right(data.Files(i), 4))
                    Case ".wbc", ".wbz", ".wbd", ".wb1", ".wbp", ".wb0"
                        AjoutFichier data.Files(i), (EviterDoublons.Value = vbUnchecked)
                End Select
            End If
        Next i
        
        'Mise à jour de la liste
        lstAddHScroll List1
    End If
End Sub

Private Sub List1_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    'N'autorise que les fichiers en drag'n'drop
    If data.GetFormat(15) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub MemeDossier_Click()
    Destination.Enabled = IIf(MemeDossier.Value = vbChecked, False, True)
    ParamDest.Enabled = IIf(MemeDossier.Value = vbChecked, False, True)
    Command3.Enabled = IIf(MemeDossier.Value = vbChecked, False, True)
    Command7.Enabled = IIf(MemeDossier.Value = vbChecked, False, IIf(ParamDest.ListIndex = 3, True, False))
End Sub

Private Sub mnu_filtre_vignettes_Click()
    ListeSupprVignettes
End Sub

Private Sub mnu_invselect_Click()
    ListeInvSelect
End Sub

'Element Ouvrir du menu contextuel de la liste
Private Sub mnu_ouvrir_Click()
    ListeOuvreSelect
End Sub

'Element Supprimer du menu contextuel de la liste
Private Sub mnu_supprimer_Click()
    ListeSupprSelect
End Sub

Private Sub mnu_ttselect_Click()
    ListeTtSelect
End Sub






Private Sub ParamDest_Click()
    Command7.Enabled = IIf(ParamDest.ListIndex = 3, True, False)

    'Sauvegarde de l'état dans la base de registre
    SaveSetting "UWC", "Batch Conversion", "Dest", ParamDest.ListIndex
End Sub




Private Sub ParamFichier_Click()
    Command8.Enabled = IIf(ParamFichier.ListIndex = 3, True, False)

    'Sauvegarde de l'état dans la base de registre
    SaveSetting "UWC", "Batch Conversion", "File", ParamFichier.ListIndex
End Sub


'Raffraichissement de la barre de progression
Private Sub progressbar_Paint()
    progress.Draw
End Sub

'Ajout d'un texte au log
Private Sub AddToLog(Texte As String, Optional LogFileOnly As Boolean = False)
    If Not LogFileOnly Then
        Log.Text = Right(Log.Text, 30000) + Texte + vbCrLf
        Log.SelStart = Len(Log.Text)
    End If
    
    If LogFileNumber <> 0 Then Print #LogFileNumber, CStr(Time) + " > " + Texte
End Sub

'Ajout d'un fichier a la liste
'
'* BypassTest permet de se passer de la vérification de doublons
Private Sub AjoutFichier(ByVal Fichier As String, Optional ByVal BypassTest As Boolean = False)
    Dim i As Integer
    
    If Not BypassTest Then
        For i = 0 To List1.ListCount - 1
            If List1.List(i) = Fichier Then Exit Sub
        Next i
    End If
    
    List1.AddItem Fichier
    AddToLog "[Conversion de masse][Ajout de fichier] Ajout de " + Fichier, True
End Sub

'Ajout d'un dossier a la liste
'
'* BypassTest permet de se passer de la vérification de doublons
Private Sub AjoutDossier(ByVal Doss As String, Optional ByVal BypassTest As Boolean = False)
    Dim fileinfo As WIN32_FIND_DATA, handle As Long
    Dim Fichier As String

    If Right(Doss, 1) <> "\" Then Doss = Doss + "\"
    
    AddToLog "[Conversion de masse][Ajout de dossier] Parcours de " + Doss, True
    
    handle = FindFirstFile(Doss + "*.wb*", fileinfo)
    If handle = INVALID_HANDLE_VALUE Then Exit Sub
    
    Do
        Fichier = Mid(fileinfo.cFileName, 1, InStr(1, fileinfo.cFileName, Chr(0)) - 1)
        
        Select Case LCase(Right(Fichier, 4))
            Case ".wbc", ".wbz", ".wb1", ".wbd", ".wbp", ".wb0"
                AjoutFichier Doss + Fichier, BypassTest
        End Select
        
        'Gère les évènements
        DoEvents
        
        If KeyPressed(vbKeyEscape) Then Exit Sub

    Loop Until FindNextFile(handle, fileinfo) = 0

    FindClose handle
End Sub

'Ajout d'un dossier et de des sous-dossiers à la liste
'
'* BypassTest permet de se passer de la vérification de doublons
Private Sub AjoutDossierRecurse(ByVal Doss As String, Optional ByVal BypassTest As Boolean = False)
    Dim fileinfo As WIN32_FIND_DATA, handle As Long, d As String
    
    If Right(Doss, 1) <> "\" Then Doss = Doss + "\"
    
    handle = FindFirstFile(Doss + "*", fileinfo)
    If handle = INVALID_HANDLE_VALUE Then Exit Sub
    
    progress.Text = Doss
    AjoutDossier Doss, BypassTest
    
    'Boucle pour tous les dossiers contenus dans ce dossier
    Do
        d = Mid(fileinfo.cFileName, 1, InStr(1, fileinfo.cFileName, Chr(0)) - 1)
        
        If (fileinfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) And d <> "." And d <> ".." Then
            AjoutDossierRecurse Doss + d, BypassTest
        End If
        
        'Gère les évènements
        DoEvents
        
        If KeyPressed(vbKeyEscape) Then Exit Sub
    
    Loop Until FindNextFile(handle, fileinfo) = 0
    
    FindClose handle
End Sub



'Sélectionne toute la liste
Private Sub ListeTtSelect()
    Dim i As Integer

    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
    Next i
End Sub
'Inverse la sélection de la liste
Private Sub ListeInvSelect()
    Dim i As Integer

    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = Not List1.Selected(i)
    Next i
End Sub

'Supprime tous les fichiers sélectionnés dans la liste
Private Sub ListeSupprSelect()
    Dim k As Integer
    
    k = 0
    
    Do Until k + 1 > List1.ListCount
        If List1.Selected(k) = True Then
            List1.RemoveItem k
        Else
            k = k + 1
        End If
    Loop
    
    'Mise à jour de la liste
    lstAddHScroll List1
    List1.SetFocus
End Sub

'Supprime toutes les vignettes de la liste
Private Sub ListeSupprVignettes()
    Dim k As Integer
    
    k = 0
    
    Do Until k + 1 > List1.ListCount
        If LCase(Mid(GetFileName(List1.List(k)), 1, 2)) = "th" Then
            List1.RemoveItem k
        Else
            k = k + 1
        End If
    Loop
    
    'Mise à jour de la liste
    lstAddHScroll List1
    List1.SetFocus
End Sub
'Ouvre tous les fichiers sélectionnés dans la liste
Private Sub ListeOuvreSelect()
    Dim k As Integer
    
    k = 0
    
    Do Until k + 1 > List1.ListCount
        If List1.Selected(k) = True Then OpenFile List1.List(k)
        k = k + 1
    Loop
End Sub

'Gestion du comportement pour un fichier
' - Si il n'existe pas, renvoie True
' - Si il existe déjà
'    - Si Comportement = FAE_RENAME, renomme Fichier et renvoie True
'    - Si Comportement = FAE_OVERWRITE, supprime le fichier Fichier et renvoie True
'    - Si Comportement = FAE_SKIP, renvoie False
'
'L'argument Stats gère les statistiques (cF. type ConvStats)
Private Function GestionFAE(ByVal Comportement As Integer, ByVal Dossier As String, ByRef Fichier As String, ByRef Stats As ConvStats) As Boolean
    Dim Void As ConvStats

    'Si le fichier existe
    If FileExists(Dossier + Fichier) Then
        Select Case Comportement
            Case FAE_RENAME
                'Renomme le nouveau fichier
                Fichier = "_" + Fichier
                'Teste si le fichier renommé existe déjà et le renomme encore si nécessaire etc...
                GestionFAE FAE_RENAME, Dossier, Fichier, Void
                GestionFAE = True
                
                Stats.Renommes = Stats.Renommes + 1
                Stats.Convertis = Stats.Convertis + 1
            Case FAE_OVERWRITE
                'Supprime le fichier
                Kill Dossier + Fichier
                GestionFAE = True
                Stats.Convertis = Stats.Convertis + 1
            Case FAE_SKIP
                Stats.Passes = Stats.Passes + 1
                GestionFAE = False
        End Select
    'Si le fichier n'existe pas
    Else
        Stats.Convertis = Stats.Convertis + 1
        GestionFAE = True
    End If
End Function

Public Sub EnableAll(ByVal State As Boolean)
    Command1.Enabled = State
    Command2.Enabled = State
    Command3.Enabled = State
    Command5.Enabled = State
    Command6.Enabled = State
    
    List1.Enabled = State
    
    Destination.Enabled = State
    
    EviterDoublons.Enabled = State
    FAE(0).Enabled = State
    FAE(1).Enabled = State
    FAE(2).Enabled = State
End Sub

Private Sub TranslateForm()
    MassConvert.Caption = "UWC - " + GetTranslation(5)

    frm_liste.Caption = GetTranslation(100)
    Command1.Caption = GetTranslation(101)
    Command6.Caption = GetTranslation(102)
    Command2.Caption = GetTranslation(103)
    Command5.Caption = GetTranslation(104)
    EviterDoublons.Caption = GetTranslation(105)
    
    List1.ToolTipText = GetTranslation(132)
    
    mnu_ttselect.Caption = GetTranslation(106)
    mnu_invselect.Caption = GetTranslation(107)
    mnu_filtres.Caption = GetTranslation(108)
    mnu_filtre_vignettes.Caption = GetTranslation(109)
    mnu_filtre_doublons.Caption = GetTranslation(110)
    mnu_filtre_perso.Caption = GetTranslation(111)
    mnu_ouvrir.Caption = GetTranslation(112)
    mnu_supprimer.Caption = GetTranslation(113)
    
    frm_dest.Caption = GetTranslation(115)
    MemeDossier.Caption = GetTranslation(133)
    
    frm_para.Caption = GetTranslation(114)
    Label2.Caption = GetTranslation(119)
    FAE(0).Caption = GetTranslation(120)
    FAE(1).Caption = GetTranslation(121)
    FAE(2).Caption = GetTranslation(122)
    
    frm_log.Caption = GetTranslation(128)
    
    Command4.Caption = GetTranslation(123)
    
    ParamDest.Clear
    ParamDest.AddItem GetTranslation(134)
    ParamDest.AddItem GetTranslation(135)
    ParamDest.AddItem GetTranslation(136)
    ParamDest.AddItem GetTranslation(140)
    
    ParamFichier.Clear
    ParamFichier.AddItem GetTranslation(137)
    ParamFichier.AddItem GetTranslation(138)
    ParamFichier.AddItem GetTranslation(139)
    ParamFichier.AddItem GetTranslation(140)
End Sub
