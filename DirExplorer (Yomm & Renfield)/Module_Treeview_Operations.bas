Attribute VB_Name = "Module_Treeview_Operations"
Option Explicit
Private drapeau As Long
'# Procedure dite de rappel du TreeView
Public Function TVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case TVM_SETIMAGELIST:
            Exit Function
           
        Case OCM_NOTIFY
            Dim dwRtn As Long
  
            dwRtn = TV_Notify(hWnd, lParam)
            If dwRtn Then
                TVWndProc = dwRtn
                Exit Function
            End If
                        
        Case WM_DESTROY
            Dim oldProc  As Long
            oldProc = GetProp(hWnd, "OldProc")
            Call SetWindowLong(hWnd, GWL_WNDPROC, oldProc)
    
    End Select
    TVWndProc = CallWindowProc(GetProp(hWnd, "OldProc"), hWnd, uMsg, wParam, lParam)
End Function

'# Cette fonction traite les messages de notification
Public Function TV_Notify(hwndTV As Long, ByVal lParam As Long) As Long
    Dim nmtv As NMTREEVIEW
    Dim tvid As cTVItem
  
    '# lParam pointe sur un NMHDR, ou sur une structure plus complexe, dont le premier membre est un NMHDR
    MoveMemory nmtv, ByVal lParam, Len(nmtv)
  
    Select Case nmtv.hdr.code
        Case TVN_ITEMEXPANDING
            If (TV_GetChild(nmtv.itemNew.hItem) = 0) Then
                Call InsertSubfolders(mTVItems(CStr(nmtv.itemNew.lParam)).pidlFQ, nmtv.itemNew.hItem, GetNodeFromlParam(nmtv.itemNew.lParam))
            End If
      
        Case TVN_SELCHANGED
        '# Ajout ultérieur d'un ListView
    
        Case NM_RCLICK
        '# Menu Contextuel
    
      Case TVN_DELETEITEM
        Set tvid = mTVItems(CStr(nmtv.itemOld.lParam))
        If (tvid Is Nothing) = False Then
              isMalloc.Free ByVal tvid.pidlRel
              isMalloc.Free ByVal tvid.pidlFQ
            Call mTVItems.Remove(CStr(nmtv.itemOld.lParam))
        End If
    End Select
End Function

'# Insertion de la racine
Public Function InsertRootFolder(pidlFQ As Long) As Long
    
    Dim pidlRel As Long
    Dim hItem As Long
   
    '# On vide l'arbre
    'Call RemoveRootFolder
    mTv.Nodes.Clear
    
    '# Cree un pid, relatif a la racine passee en parametre
    pidlRel = GetItemID(pidlFQ, GIID_LAST)
    If pidlRel Then
        '# On ajoute ce noeud
        hItem = InsertFolder(Nothing, GetIShellFolderParent(pidlFQ), pidlFQ, pidlRel, 0, 0)
        If hItem Then
            '# On ouvre le premier noeud
            mTv.Nodes(1).Selected = True
            mTv.Nodes(1).Expanded = True
        
            InsertRootFolder = hItem
        End If
        '# On libere l'espace alloué
        Call FreePIDL(pidlRel)
    End If
    
End Function

'# Permet de vider l'arbre
Public Sub RemoveRootFolder()
    If mTv.Nodes.Count Then
        mTv.Nodes(1).Root.Expanded = False
        Call mTv.Nodes.Remove(mTv.Nodes(1).Root.Index)
  End If
End Sub

Public Function InsertFolder(nodParent As Node, isfParent As IShellFolder, pidlfqChild As Long, pidlrelChild As Long, hitemParent As Long, hitemPrevChild As Long) As Long
    Dim ulAttrs As ESFGAO
    Dim TVI As TVITEM
    Dim tvid As New cTVItem
  
    '# On précise les informations que l'on souhaite certaines informations
    ulAttrs = SFGAO_HASSUBFOLDER Or SFGAO_SHARE
    '# On demandes ces informations sur l'item
    Call isfParent.GetAttributesOf(1, pidlrelChild, ulAttrs)
    
    '# l'Item contient les elements suivant :
    TVI.mask = TVIF_CHILDREN Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    
    '# l'Item possede-t'il des sous-item
    TVI.cChildren = Abs(CBool(ulAttrs And SFGAO_HASSUBFOLDER))
    
    '# On recupere l'indice des icones
    TVI.iImage = GetFileIconIndexPIDL(pidlfqChild, SHGFI_SMALLICON)
    TVI.iSelectedImage = GetFileIconIndexPIDL(pidlfqChild, SHGFI_SMALLICON Or SHGFI_OPENICON)
    
    '# Si l'item est en partage, on ajoute l'overlay correspondant (la petite main !)
    If (ulAttrs And SFGAO_SHARE) Then
        TVI.mask = TVI.mask Or TVIF_STATE
        TVI.state = TVIS_OVERLAYMASK

        TVI.stateMask = TVIO_SHARE
    End If
  
    '# ON AJOUTE LES CHECKBOXES
    TVI.mask = TVI.mask Or TVIF_HANDLE  'Or TVIF_STATE
    TVI.stateMask = TVI.stateMask Or TVIS_STATEIMAGEMASK
    TVI.state = TVI.state Or INDEXTOSTATEIMAGEMASK(1)
  
    Dim Node As Node
  drapeau = drapeau + 1
    '# On ajoute le noeud
    
    If (nodParent Is Nothing) Then
        Set Node = mTv.Nodes.Add(Text:=GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER))
    Else
        Set Node = mTv.Nodes.Add(nodParent, tvwChild, Text:=GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER))
    End If
    
    '# on recuperes le pid de l'item que l'on vien d'ajouter
    If (hitemParent = 0) Then
        TVI.hItem = TV_GetRoot()
    ElseIf (hitemPrevChild = 0) Then
        TVI.hItem = TV_GetChild(hitemParent)
    Else
        TVI.hItem = TV_GetNextSibling(hitemPrevChild)
    End If
  
    '# et on place toutes les informations precedemment etablies dans le noeud
    Call TV_SetItem(TVI)
    
    Node.Tag = CStr(GetTVItemlParam(TVI.hItem))
    '# on alloue de nouveaux emplacements memoire, pour copier les pid ( absolus : Fully Qualified et relatifs )
    '# puis on place le tout dans la collection
    tvid.pidlFQ = CopyPIDL(pidlfqChild)
    tvid.pidlRel = CopyPIDL(pidlrelChild)
    tvid.Path = GetPathFromPIDL(tvid.pidlFQ)
    tvid.lParam = CStr(GetTVItemlParam(TVI.hItem))
    tvid.hNode = TVI.hItem
    On Error Resume Next
    Call mTVItems.Add(tvid, CStr(GetTVItemlParam(TVI.hItem)))
    
    '# on renvoi le handle de l'item cree
    InsertFolder = TVI.hItem
End Function

'# permet de retrouver les sous elements d'un item donné , et de les ajouter au treeview
Public Sub InsertSubfolders(pidlfqParent As Long, hitemParent As Long, nodParent As Node)
    Dim hwndOwner As Long
    Dim isfParent As IShellFolder
    Dim ieidl As IEnumIDList
    Dim pidlrelChild As Long
    Dim pidlfqChild As Long
    Dim hitemChild As Long
    Dim tvscb As TVSORTCB
    Dim TVI As TVITEM
  
    hwndOwner = GetTopLevelParent(mTv.hWnd)
    
    '# On recupere le ShellFolder parent, afin de pouvoir effectuer sur celui-ci une enumeration de son contenu.
    Set isfParent = GetIShellFolder(isfDesktop, pidlfqParent)
    'yommfile
    'If isfParent.EnumObjects(hwndOwner, SHCONTF_INCLUDEHIDDEN Or SHCONTF_NONFOLDERS, ieidl) >= 0 Then
    If isfParent.EnumObjects(hwndOwner, SHCONTF_FOLDERS Or SHCONTF_INCLUDEHIDDEN, ieidl) >= 0 Then
        '# On fait une enumeration du contenu du repertoire parent
        While (ieidl.Next(1, pidlrelChild, 0) = 0)
            '# On reference l'IDL enfant par rapport a la racine
            pidlfqChild = CombinePIDLs(pidlfqParent, pidlrelChild)
            If pidlfqChild Then
                '# On ajoute le repertoire
                hitemChild = InsertFolder(nodParent, isfParent, pidlfqChild, pidlrelChild, hitemParent, hitemChild)
                
                '# On libere la memoire allouée
                isMalloc.Free ByVal pidlfqChild
            End If
            
            '# On libere la memoire de l' IDL qui nous a servit pour l'enumeration
            isMalloc.Free ByVal pidlrelChild
        Wend
    End If
    
    '# Permet de trier les elements
    If hitemChild Then
        '# On utilise un fonction de rappel pour trier
        '# les elements lui seront envoyés, et selon la valeur quelle renverra, les elements seront inversés automatiquement..
        tvscb.hParent = hitemParent
        MoveMemory tvscb.lpfnCompare, AddressOf TreeViewCompareProc, 4
        tvscb.lParam = ObjPtr(isfParent)
        Call TV_SortChildrenCB(tvscb, 0)
    Else
        '# on ne peut acceder au noeuds enfants  : on supprime le [+]
        TVI.hItem = hitemParent
        TVI.mask = TVIF_CHILDREN
        TVI.cChildren = 0
        Call TV_SetItem(TVI)
    End If
End Sub

'# Permet de rafraichir l'arborescence
Public Sub RefreshTreeview(nodSibling As Node)
    Dim nodChild  As Node
  
    While Not (nodSibling Is Nothing)
        If nodSibling.Expanded Then
            '# le noeud est ouvert, on procede recursivement....
            Call RefreshTreeview(nodSibling.Child)
        Else
            '# le noeud est ferme, on procede avec le noeud de meme niveau suivant
            While Not (nodChild Is Nothing)
                mTv.Nodes.Remove nodChild.Index
                Set nodChild = nodSibling.Child
            Wend
        End If
        Set nodSibling = nodSibling.Next
    Wend
End Sub

'# Fonction de tri des items
Public Function TreeViewCompareProc(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParamSort As Long) As Long
    Dim isfParent As IShellFolder
    Dim result As Long
  
    '# On initialise le tri
    MoveMemory isfParent, lParamSort, 4

    '# On demande a l'interface de trier les elements
    result = isfParent.CompareIDs(0, mTVItems(CStr(lParam1)).pidlRel, mTVItems(CStr(lParam2)).pidlRel)

    '# Si le resultat obtenu est cohérent, ne garde que le Lo-Word
    If (result >= 0) Then TreeViewCompareProc = LOWORD(result)
  
    '# On remet le tri a 0 dans l'interface
    FillMemory isfParent, 4, 0
End Function

Public Function GetTVItemlParam(hItem As Long) As Long
    Dim TVI As TVITEM
    
    TVI.hItem = hItem
    TVI.mask = TVIF_PARAM
  
    If TV_GetItem(TVI) Then GetTVItemlParam = TVI.lParam
End Function

Public Function GetTVItemData(hItem As Long) As cTVItem
    Set GetTVItemData = mTVItems(CStr(GetTVItemlParam(hItem)))
End Function

Public Function GetNodeFromhItem(hwndTV As Long, hItem As Long) As Node
    Set GetNodeFromhItem = GetNodeFromlParam(GetTVItemlParam(hItem))
End Function

Public Function GetNodeFromlParam(lParam As Long) As Node
    Dim pNode As Long
    Dim nod As Node
    
    If lParam Then
        MoveMemory pNode, ByVal lParam + 8, 4
        If pNode Then
            MoveMemory nod, pNode, 4
            Set GetNodeFromlParam = nod
            FillMemory nod, 4, 0
        End If
    End If
End Function

Public Function TV_SetImageList(himl As Long, iImage As Long) As Long
    TV_SetImageList = SendMessage(mTv.hWnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)
End Function

Public Function TV_GetNextItem(hItem As Long, flag As Long) As Long
    TV_GetNextItem = SendMessage(mTv.hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function

Public Function TV_GetChild(hItem As Long) As Long
    TV_GetChild = TV_GetNextItem(hItem, TVGN_CHILD)
End Function

Public Function TV_GetNextSibling(hItem As Long) As Long
    TV_GetNextSibling = TV_GetNextItem(hItem, TVGN_NEXT)
End Function

Public Function TV_GetSelection() As Long
    TV_GetSelection = TV_GetNextItem(0, TVGN_CARET)
End Function

Public Function TV_GetRoot() As Long
    TV_GetRoot = TV_GetNextItem(0, TVGN_ROOT)
End Function

Public Function TV_GetItem(pItem As TVITEM) As Boolean
    TV_GetItem = SendMessage(mTv.hWnd, TVM_GETITEM, 0, pItem)
End Function

Public Function TV_SetItem(pItem As TVITEM) As Boolean
    TV_SetItem = SendMessage(mTv.hWnd, TVM_SETITEM, 0, pItem)
End Function

Public Function TV_HitTest(lpht As TVHITTESTINFO) As Long
    TV_HitTest = SendMessage(mTv.hWnd, TVM_HITTEST, 0, lpht)
End Function

Public Function TV_SortChildrenCB(psort As TVSORTCB, fRecurse As Boolean) As Boolean
    TV_SortChildrenCB = SendMessage(mTv.hWnd, TVM_SORTCHILDRENCB, ByVal fRecurse, psort)
End Function
