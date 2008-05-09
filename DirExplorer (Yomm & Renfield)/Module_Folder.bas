Attribute VB_Name = "Module_Folder"
Option Explicit

'#Renvoie une interface au IShellFolder racin
Public Function isfDesktop() As IShellFolder
    Static isf As IShellFolder
    '# Si l'element n'existe pas, on l'initialise
    If (isf Is Nothing) Then Call SHGetDesktopFolder(isf)
    Set isfDesktop = isf
End Function

'# Renvoie le nom d'un Item
Public Function GetFolderDisplayName(isfParent As IShellFolder, pidlRel As Long, uFlags As SHGFI_Enum) As String
    Dim lpStr As STRRET
    If 0 <= isfParent.GetDisplayNameOf(pidlRel, uFlags, lpStr) Then
        GetFolderDisplayName = GetStrRet(lpStr, pidlRel)
    End If
End Function

'# Renvoie un IShellFolder, d'apres un Pid relatif (relatif au IShellFolder Parent specifié !)
Public Function GetIShellFolder(isfParent As IShellFolder, pidlRel As Long) As IShellFolder
    Dim isf As IShellFolder
    On Error GoTo Fin
       
    Call isfParent.BindToObject(pidlRel, 0, GetGUID(IID_IShellFolder), isf)
    
Fin:
    If Err Or (isf Is Nothing) Then
        Set GetIShellFolder = isfDesktop
    Else
        Set GetIShellFolder = isf
    End If
End Function

'# Renvoie une reference vers l' IShellFolder parent du dernier Item du PID spécifié
Public Function GetIShellFolderParent(ByVal pidlFQ As Long, Optional fRtnDesktop As Boolean = True) As IShellFolder
    Dim pidlParent As Long

    pidlParent = GetPIDLParent(pidlFQ, fRtnDesktop)
    If pidlParent Then
        Set GetIShellFolderParent = GetIShellFolder(isfDesktop, pidlParent)
        isMalloc.Free ByVal pidlParent
    End If
End Function
