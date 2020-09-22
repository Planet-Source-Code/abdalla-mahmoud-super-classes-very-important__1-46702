VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHookToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//Name        : clsHookToolbar
'//Author      : Abdalla Mahmoud
'//E-Mail      : la3toot@hotmail.com
'//Description : With this class you can
            'get the event of the MouseMove on
            'the toolbar buttons
Option Explicit

Private Type POINTAPI
        X As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private WithEvents m_Extender As Toolbar
Attribute m_Extender.VB_VarHelpID = -1
Private WithEvents m_Timer    As CTimer
Attribute m_Timer.VB_VarHelpID = -1

Public Event ButtonMouseMove(ByVal Index As Long, ByVal IsSeparator As Boolean)
Public Event ButtonMouseOut()

Private m_Rect() As RECT
Private m_HelpText() As String
Private m_Count  As Long
Private m_OldButton As Long
Private m_hWnd As Long
Private m_MouseOut As Boolean

Public Sub SubClass(ByVal XToolbar As Object)
    If Not TypeOf XToolbar Is Toolbar Then Exit Sub
    Dim I As Long
    Dim TP_X As Long
    Dim TP_Y As Long
    Set m_Extender = XToolbar
    With m_Extender
        m_Count = .Buttons.Count
        Erase m_Rect
        ReDim m_Rect(1 To m_Count) As RECT
        Erase m_HelpText
        ReDim m_HelpText(1 To m_Count) As String
        For I = 1 To .Buttons.Count
            With .Buttons(I)
                m_Rect(I).Left = .Left / Screen.TwipsPerPixelX
                m_Rect(I).Top = .Top / Screen.TwipsPerPixelY
                m_Rect(I).Right = (.Left + .Width) / Screen.TwipsPerPixelX
                m_Rect(I).Bottom = (.Top + .Height) / Screen.TwipsPerPixelY
            End With
        Next
    End With
End Sub

Private Function GetIndex(ByVal Key) As Long
    On Error GoTo FunErr
    If IsNumeric(Key) Then
        GetIndex = CLng(Key)
    Else
        GetIndex = m_Extender.Buttons(Key).Index
    End If
    Exit Function
FunErr:
    GetIndex = 0
End Function

Private Sub Class_Initialize()
    Set m_Timer = New CTimer
    With m_Timer
        .Enabled = False
        .Interval = 20
    End With
End Sub

Private Sub Class_Terminate()
    Set m_Timer = Nothing
    Set m_Extender = Nothing
    Erase m_HelpText
    Erase m_Rect
End Sub

Private Sub m_Extender_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim PO As POINTAPI
    Dim HT As Long
    Dim I  As Long
    Call GetCursorPos(PO)
    Call ScreenToClient(m_Extender.hWnd, PO)
    For I = 1 To m_Count
        If PtInRect(m_Rect(I), PO.X, PO.y) Then
            m_Timer.Enabled = True
            If m_hWnd = 0 Then
                Call GetCursorPos(PO)
                m_hWnd = WindowFromPoint(PO.X, PO.y)
            End If
            If I <> m_OldButton Then
                m_OldButton = I
                RaiseEvent ButtonMouseMove(m_OldButton, m_Extender.Buttons(m_OldButton).Style = tbrSeparator)
                m_MouseOut = False
            End If
            Exit Sub
        End If
    Next
    If Not m_MouseOut Then
        m_MouseOut = True
        m_Timer.Enabled = False
        m_OldButton = 0
        RaiseEvent ButtonMouseOut
    End If
End Sub

Public Property Get HelpText(ByVal Key) As Variant
    On Error Resume Next
    Dim Index As Long
    Index = GetIndex(Key)
    If Index Then
        HelpText = m_HelpText(Index)
    End If
End Property

Public Property Let HelpText(ByVal Key, ByVal vNewValue As Variant)
    On Error Resume Next
    Dim Index As Long
    Index = GetIndex(Key)
    If Index Then
        m_HelpText(Index) = vNewValue
    End If
End Property

Private Sub m_Timer_Time()
    Dim PO As POINTAPI
    Dim lhWnd As Long
    Call GetCursorPos(PO)
    lhWnd = WindowFromPoint(PO.X, PO.y)
    If lhWnd <> m_hWnd And m_hWnd <> 0 Then
        RaiseEvent ButtonMouseOut
        m_Timer.Enabled = False
        m_OldButton = 0
    End If
End Sub
