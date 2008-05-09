VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl DirExplorer 
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   ScaleHeight     =   2985
   ScaleWidth      =   2040
   ToolboxBitmap   =   "DirExplorer.ctx":0000
   Begin ComctlLib.TreeView TView 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   5212
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "DirExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private ObjPath As String
Private ObjRoot As Long

Public Event Click()
Public Event Change()

Public Enum TypeRoot
    Bureau = CSIDL_DESKTOP
    PosteDeTravail = CSIDL_DRIVES
    Windows = CSIDL_WINDOWS
    System32 = CSIDL_SYSTEM32
    ProgramFiles = CSIDL_PROGAMFILES
    MesDocuments = CSIDL_PERSONAL
    Reseau = CSIDL_NETWORK
    'InternetExplorer = CSIDL_INTERNET
    'MenuProgram = CSIDL_PROGRAMS
    'PanneauDeConfig = CSIDL_CONTROLS
    'Imprimante = CSIDL_PRINTERS
    'Favoris = CSIDL_FAVORITES
    'RepDemarrer = CSIDL_STARTUP
    'DocRecent = CSIDL_RECENT
    'SendTo = CSIDL_SENDTO
    'Corbeille = CSIDL_BITBUCKET
    'MenuDemarrer = CSIDL_STARTMENU
    'RepBureau = CSIDL_DESKTOPDIRECTORY
    'VoisinageReseau = CSIDL_NETHOOD
    'Police = CSIDL_FONTS
    'Modeles = CSIDL_TEMPLATES
    'MenuDemarrerAll = CSIDL_COMMON_STARTMENU
    'MenuProgramAll = CSIDL_COMMON_PROGRAMS
    'RepDemarrerAll = CSIDL_COMMON_STARTUP
    'RepBureauAll = CSIDL_COMMON_DESKTOPDIRECTORY
    'AppliDate = CSIDL_APPDATA
    'VoisinageImpr =  CSIDL_PRINTHOOD
    ' ??? = CSIDL_ALTSTARTUP
    ' ??? = CSIDL_COMMON_ALTSTARTUP
    'FavorisAll = CSIDL_COMMON_FAVORITES
    'TempInternetFile = CSIDL_INTERNET_CACHE
    'Cookies = CSIDL_COOKIES
    'Historique = CSIDL_HISTORY
End Enum

Public Property Get Chemin() As String
    
    Chemin = ObjPath
    
End Property

Public Property Let Chemin(ByVal Path As String)

    ObjPath = Path
    ExpandNode LCase(Path)
    PropertyChanged "Chemin"
    RaiseEvent Change
    
End Property

Public Property Get Root() As TypeRoot
    
    Root = ObjRoot
    
End Property

Public Property Let Root(ByVal Root As TypeRoot)
    
    ObjRoot = Root
    Call InsertRootFolder(GetPIDLFromFolderID(hWnd, ObjRoot))
    ExpandNode LCase(ObjPath)
    PropertyChanged "Root"
    
End Property

Private Sub TView_Collapse(ByVal Node As ComctlLib.Node)
#If 0 Then
    If Node.FullPath = Left(TView.SelectedItem.FullPath, Len(Node.FullPath)) Then
        ObjPath = mTVItems(Node.Index).Path
        If mTVItems(Node.Index).Path <> "" Then
            RaiseEvent Click
        End If
    End If
#End If
End Sub

Private Sub TView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Branche As Node
    Dim R As RECT
    
    'récupération de l'objet placé sous le curseur
    Set Branche = TView.HitTest(x, y)
    
    'si l'objet est une node
    If (Branche Is Nothing) = False Then
        'obtention du rectangle encadrant le teste de la node
        R.Left = mTVItems(Branche.Tag).hNode
        SendMessage TView.hWnd, TVM_GETITEMRECT, True, R
        
        'test du point droit par rapport à la largeur du treeview (taille treeview = taille usercontrol)
        If R.Right > UserControl.ScaleWidth \ 15 Then
            TView.ToolTipText = Branche.Text
        Else
           TView.ToolTipText = ""
        End If
    Else
        TView.ToolTipText = ""
    End If
    
    Set Branche = Nothing
    
End Sub

Private Sub TView_NodeClick(ByVal Node As ComctlLib.Node)

    ObjPath = mTVItems(Node.Index).Path
    If mTVItems(Node.Index).Path <> "" Then
        RaiseEvent Click
    End If

End Sub


Private Sub UserControl_Initialize()

    Set mTv = TView
    Call TV_SetImageList(GetSystemImageList(SHGFI_SMALLICON), TVSIL_NORMAL)
    Call Subclass(mTv.hWnd, AddressOf TVWndProc)
    
    Call InsertRootFolder(GetPIDLFromFolderID(hWnd, CSIDL_DESKTOP))
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ObjRoot = PropBag.ReadProperty("Root", 0)
    ObjPath = PropBag.ReadProperty("Chemin", "")
    InitTree
    
End Sub

Private Sub UserControl_Resize()

    TView.Width = UserControl.Width
    TView.Height = UserControl.Height

End Sub

Private Sub UserControl_Terminate()

    Call RemoveRootFolder
    Call UnSubclass(mTv.hWnd)

    Call TV_SetImageList(0, TVSIL_NORMAL)
  
End Sub

Sub ExpandNode(ByVal Path As String)

    Dim Node As Node
    Dim TVI As cTVItem
    Dim sPath() As String
    Dim i As Long
    
    On Error Resume Next
    
    If Trim(Path) = "" Then
        For Each TVI In mTVItems
            If GetNodeFromlParam(TVI.lParam).Parent Is Nothing Then
                GetNodeFromlParam(TVI.lParam).Expanded = True
                GetNodeFromlParam(TVI.lParam).Selected = True
                GetNodeFromlParam(TVI.lParam).EnsureVisible
            Else
                GetNodeFromlParam(TVI.lParam).Expanded = False
            End If
        Next TVI
    Else
        sPath = Split(Path, "\")
        Path = sPath(0)
        For i = 0 To UBound(sPath)
            For Each TVI In mTVItems
                If TVI.Path = vbNullString Then
                    If GetNodeFromlParam(TVI.lParam).Text = "Poste de travail" Then
                        GetNodeFromlParam(TVI.lParam).Expanded = True
                    End If
                ElseIf GetNodeFromlParam(TVI.lParam).Text = "Bureau" Then
                    GetNodeFromlParam(TVI.lParam).Expanded = True
                ElseIf LCase(TVI.Path) = Path Or LCase(TVI.Path) = Path & "\" Then
                    Set Node = GetNodeFromlParam(TVI.lParam)
                    Node.Expanded = True
                    Path = Path & "\" & sPath(i + 1)
                    Exit For
                End If
            Next TVI
        Next i
        With Node
            .EnsureVisible
            .Selected = True
        End With
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Chemin", ObjPath, "")
    Call PropBag.WriteProperty("Root", ObjRoot, 0)
    
End Sub

Sub InitTree()

    Call InsertRootFolder(GetPIDLFromFolderID(hWnd, ObjRoot))
    ExpandNode ObjPath
    
End Sub
