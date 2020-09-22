VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//Name        : clsMenus
'//Author      : Abdalla Mahmoud
'//E-Mail      : la3toot@hotmail.com
'//Description : get the MouseMove and MouseOut
                'event of the menus
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private m_SubClass As New cSuperClass
Implements ISuperClass

Private m_VBMenus      As New Collection
Private m_CollMenus    As New Collection

Private m_hMenu  As Long
Private m_hWnd   As Long

Private m_MenusCount As Long
Private Const WM_MENUSELECT = &H11F

Public Event MouseMove(ByVal MenuName As String, ByVal IsSeparator As Boolean)
Public Event MouseOut()

Private Sub Class_Terminate()
    Call m_SubClass.UnSubclass
    Set m_SubClass = Nothing
    Set m_CollMenus = Nothing
    Set m_VBMenus = Nothing
End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'On Error Resume Next
    'Dim DrawInfo As DRAWITEMSTRUCT
    Dim IsSelected As Boolean
    Dim IsDisabled As Boolean
    Dim hMenu As Long
    Dim lMenuCount As Long
    Dim lMenuId As Long
    Dim lItem As Long
    Dim bFound As Boolean
    Dim I As Long
    Dim MenuName As String
    Static X As Long
    Select Case uMsg
    Case WM_MENUSELECT
            hMenu = lParam
            lMenuId = wParam And &HFFFF&
            ' Now check if the message is a menu item higlight,
            ' or whether it is indicating exit from the menu:
            lMenuCount = GetMenuItemCount(hMenu)
            For lItem = 0 To lMenuCount - 1
                If (lMenuId = GetMenuItemID(hMenu, lItem)) Then
                    bFound = True
                    Exit For
                End If
            Next
        If bFound Then
            'frmMain.Caption = "true"
            For I = 1 To m_CollMenus.Count
                If m_CollMenus(I).ID = lMenuId Then
                    MenuName = m_CollMenus(I).Name
                    RaiseEvent MouseMove(MenuName, m_CollMenus(I).vbMenu.Caption = "-")
                    Exit For
                End If
            Next
        Else
            RaiseEvent MouseOut
        End If
    End Select
End Sub

Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub

Public Sub SubClass(XForm As Object)
    Dim CTL As Control
    Dim CIndex As Long
    Dim hMenu As Long
    m_hWnd = XForm.hWnd
    For Each CTL In XForm.Controls
        If TypeOf CTL Is Menu Then
            Call m_VBMenus.Add(CTL)
        End If
    Next
    m_hMenu = GetMenu(m_hWnd)
    Call EnumMenus(m_hMenu)
    Call m_SubClass.SubClass(m_hWnd, Me, True)
End Sub

Private Sub EnumMenus(ByVal hMenu As Long)
    On Error Resume Next
    Dim I       As Long
    Dim hSubMenu As Long
    Dim MenuCount As Long
    Dim Store_Menu As clsStoreMenu
    Dim chk     As Long
    Dim CKey As String
    MenuCount = GetMenuItemCount(hMenu)
    For I = 0 To MenuCount - 1
        Set Store_Menu = New clsStoreMenu
        With Store_Menu
            m_MenusCount = m_MenusCount + 1
            Set .vbMenu = m_VBMenus(m_MenusCount)
            .ID = GetMenuItemID(hMenu, I)
            Call Err.Clear
            CKey = .vbMenu.Name
            chk = .vbMenu.Index
            If Err.Number = 0 Then
                CKey = CKey & "(" & chk & ")"
            End If
            .Name = CKey
            Call m_CollMenus.Add(Store_Menu, CKey)
            hSubMenu = GetSubMenu(hMenu, I)
            If hSubMenu > 0 Then
                Call EnumMenus(hSubMenu)
            End If
        End With
    Next
End Sub

Private Function GetItem(ByVal Key As String) As clsStoreMenu
    On Error Resume Next
    Set GetItem = m_CollMenus.Item(Key)
    Call Err.Clear
End Function

Public Property Get HelpText(ByVal Key As String) As String
    On Error Resume Next
    Dim Index As Long
    Index = GetItem(Key)
    If Index Then
        HelpText = m_CollMenus.Item(Index).HelpText
    End If
End Property

Public Property Let HelpText(ByVal Key As String, ByVal vNewValue As String)
    On Error Resume Next
    Dim Index As Long
    Index = GetItem(Key)
    If Index Then
        m_CollMenus.Item(Index).HelpText = vNewValue
    End If
End Property
