VERSION 5.00
Object = "{9DD53B4F-5DB5-4C19-AFDC-291A7557D5C8}#16.0#0"; "DirExplorer.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin DirExplorerOCX.DirExplorer DirExplorer1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7223
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'    DirExplorer1.Chemin = Text1.Text
    DirExplorer1.Root = PosteDeTravail
    
End Sub

Private Sub DirExplorer1_Change()
Dim toto As Integer

Debug.Print "change"
End Sub

Private Sub DirExplorer1_Click()
Me.Caption = DirExplorer1.Chemin
End Sub

