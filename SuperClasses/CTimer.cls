VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//Name        : CTimer
'//Author      : Abdalla Mahmoud
'//E-Mail      : la3toot@hotmail.com
'//Description :create a timer class without
                'timer control
Option Explicit

Private m_ID       As Long
Private m_Interval As Long
Private m_Enabled  As Long
Private m_Init     As Boolean
Private m_Tag      As String
Private m_Action   As String

Public Event Time()

Friend Property Get ID() As Long
    ID = m_ID
End Property

Friend Property Let ID(ByVal vNewValue As Long)
    m_ID = vNewValue
End Property

Friend Function RaiseTimer()
    RaiseEvent Time
End Function

Public Property Get Interval() As Long
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal vNewValue As Long)
    If m_Interval = vNewValue Then Exit Property
    m_Interval = vNewValue
    If m_Enabled = False Then Exit Property
    Call tUnRegisterTimer(Me)
    If m_Enabled And m_Interval > 0 Then
        Call tRegisterTimer(Me)
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    If m_Enabled = vNewValue Then Exit Property
    m_Enabled = vNewValue
    If m_Enabled Then
        If m_Interval Then Call tRegisterTimer(Me)
    Else
        Call tUnRegisterTimer(Me)
    End If
End Property

Private Sub Class_Initialize()
    Enabled = True
End Sub

Private Sub Class_Terminate()
    Call tUnRegisterTimer(Me)
End Sub

Public Property Get Action() As String
    Action = m_Action
End Property

Public Property Let Action(ByVal vNewValue As String)
    m_Action = vNewValue
End Property

Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal vNewValue As String)
    m_Tag = vNewValue
End Property


