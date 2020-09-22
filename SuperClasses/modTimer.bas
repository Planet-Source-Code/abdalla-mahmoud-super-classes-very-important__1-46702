Attribute VB_Name = "modTimer"
'//Name        : modTimer
'//Author      : Abdalla Mahmoud
'//E-Mail      : la3toot@hotmail.com
'//Description :this muse be with the CTimer
                'class
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private m_Timers() As CTimer
Private m_TimersCount As Long

Public Sub tRegisterTimer(ByRef Class As CTimer)
    Class.ID = SetTimer(0, 0, Class.Interval, AddressOf TimerProc)
    m_TimersCount = m_TimersCount + 1
    ReDim Preserve m_Timers(1 To m_TimersCount) As CTimer
    Set m_Timers(m_TimersCount) = Class
End Sub

Public Sub tUnRegisterTimer(ByRef Class As CTimer)
    Dim I As Long
    Dim C As Long
    Dim holdArr() As CTimer
    
    If m_TimersCount = 0 Then Exit Sub
    For I = 1 To m_TimersCount
        If Class Is m_Timers(I) Then
            Call KillTimer(0, m_Timers(I).ID)
        Else
            C = C + 1
            ReDim Preserve holdArr(1 To C) As CTimer
            Set holdArr(C) = m_Timers(I)
        End If
    Next
    Erase m_Timers
    If C Then m_Timers = holdArr
    Erase holdArr
    m_TimersCount = C 'm_TimersCount - 1
End Sub

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  Dim I As Long
    For I = 1 To m_TimersCount
        If idEvent = m_Timers(I).ID Then
            m_Timers(I).RaiseTimer
            Exit For
        End If
    Next
End Sub
