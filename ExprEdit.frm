VERSION 5.00
Begin VB.Form ExprEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UWC - Editeur d'expression"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "ExprEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Tag             =   "-1"
   Begin VB.CommandButton Aide 
      Caption         =   "&Aide"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   6480
      TabIndex        =   9
      Top             =   720
      Width           =   735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   240
         Width           =   495
         Begin VB.CommandButton Command1 
            Caption         =   "( )"
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "|"
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   675
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "( | )"
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   1320
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Annuler 
      Caption         =   "A&nnuler"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame frmVariables 
      Caption         =   "Variables"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   2280
         ScaleHeight     =   1815
         ScaleWidth      =   3855
         TabIndex        =   6
         Top             =   120
         Width           =   3855
         Begin VB.Frame frmDescr 
            Caption         =   "Description"
            Height          =   1695
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   3855
            Begin VB.Label lblDescr 
               Caption         =   "-"
               Height          =   1335
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   3495
               WordWrap        =   -1  'True
            End
         End
      End
      Begin VB.ListBox lstVariables 
         Height          =   1695
         IntegralHeight  =   0   'False
         ItemData        =   "ExprEdit.frx":2372
         Left            =   120
         List            =   "ExprEdit.frx":239A
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtExpr 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label lblExpr 
      Caption         =   "Expression :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "ExprEdit"
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





Private Sub Aide_Click()
    MsgBox sprintf(GetTranslation(365)), vbOKOnly + vbInformation, "UWC - " + GetTranslation(350) + " - " + GetTranslation(366)
End Sub

Private Sub Annuler_Click()
    Me.Tag = False
    Me.Hide
End Sub

Private Sub Command1_Click()
    txtExpr.SelText = "(" + txtExpr.SelText + ")"
End Sub

Private Sub Command2_Click()
txtExpr.SelText = "|"
End Sub


Private Sub Command3_Click()
txtExpr.SelText = "( 1sr possibility | 2nd possibility )"
End Sub


Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    TranslateForm
End Sub


Private Sub lstVariables_Click()
    Select Case lstVariables.Text
        Case "%pic.title"
            lblDescr = sprintf(GetTranslation(354))
        Case "%pic.id"
            lblDescr = sprintf(GetTranslation(355))
        Case "%pic.desc"
            lblDescr = sprintf(GetTranslation(356))
        Case "%pic.credit"
            lblDescr = sprintf(GetTranslation(357))
        Case "%pic.dailydate"
            lblDescr = sprintf(GetTranslation(367, " > %pic.dailydate\n\nDaily date"))
        Case "%pic.cat.root"
            lblDescr = sprintf(GetTranslation(358))
        Case "%pic.cat.child"
            lblDescr = sprintf(GetTranslation(359))
        Case "%col.title"
            lblDescr = sprintf(GetTranslation(360))
        Case "%col.filename"
            lblDescr = sprintf(GetTranslation(361))
        Case "%file.name"
            lblDescr = sprintf(GetTranslation(362))
        Case "%file.path"
            lblDescr = sprintf(GetTranslation(363))
        Case "%dest.path"
            lblDescr = sprintf(GetTranslation(364))
    End Select
End Sub



Public Function EditExpr(ByRef Expr As String, ByRef Owner As Form) As Boolean
    txtExpr.Text = Expr
    Me.Tag = False
    
    Me.Show vbModal, Owner
    
    EditExpr = CBool(ExprEdit.Tag)
    
    If EditExpr Then
        Expr = txtExpr.Text
    End If
    
    Unload Me
End Function

Private Sub lstVariables_DblClick()
    txtExpr.SelText = lstVariables.Text
End Sub

Private Sub Ok_Click()
    Me.Tag = True
    Me.Hide
End Sub



Public Sub TranslateForm()
    Me.Caption = "UWC - " + GetTranslation(350)
    lblExpr.Caption = GetTranslation(351)
    frmVariables.Caption = GetTranslation(352)
    frmDescr.Caption = GetTranslation(353)
    
    Annuler.Caption = GetTranslation(53)
    Ok.Caption = GetTranslation(54)
    
    lblDescr.Caption = ""
    Aide.Caption = GetTranslation(366)
End Sub
