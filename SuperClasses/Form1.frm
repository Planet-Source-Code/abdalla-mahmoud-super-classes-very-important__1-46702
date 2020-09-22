VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "Comct332.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Classes"
   ClientHeight    =   5115
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4740
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
            Text            =   "Help Text Displays Here"
            TextSave        =   "Help Text Displays Here"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   2275
      _CBWidth        =   6150
      _CBHeight       =   1290
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   600
      Width1          =   2205
      NewRow1         =   0   'False
      Child2          =   "Toolbar2"
      MinHeight2      =   600
      Width2          =   3360
      NewRow2         =   -1  'True
      Child3          =   "Toolbar3"
      MinHeight3      =   600
      Width3          =   1350
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   600
         Left            =   3555
         TabIndex        =   4
         Top             =   660
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1058
         ButtonWidth     =   1323
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sample 1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sample 2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sample 3"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   600
         Left            =   165
         TabIndex        =   3
         Top             =   660
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   1058
         ButtonWidth     =   1111
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abdalla"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abdalla"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abdalla"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abdalla"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1058
         ButtonWidth     =   1191
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button3"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button4"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Change the GUI of the Coolbar then exit and come again .. the Coolbar saved into the file in the App.Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Open This Menu And Move Mouse"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&See The Status Bar"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//timer without timer class
Private WithEvents m_Timer As CTimer
Attribute m_Timer.VB_VarHelpID = -1
'//mouse move events for the buttons
Private WithEvents m_HootToolbar1 As clsHookToolbar
Attribute m_HootToolbar1.VB_VarHelpID = -1
Private WithEvents m_HootToolbar2 As clsHookToolbar
Attribute m_HootToolbar2.VB_VarHelpID = -1
Private WithEvents m_HootToolbar3 As clsHookToolbar
Attribute m_HootToolbar3.VB_VarHelpID = -1
'//mouse move events for the normal menus
Private WithEvents m_HookMenu As clsMenus
Attribute m_HookMenu.VB_VarHelpID = -1
'//Coolbar GUI Saver
Private m_CoolbarSaver As New clsCoolbarSaver
Attribute m_CoolbarSaver.VB_VarHelpID = -1

Private Sub Form_Activate()
    '//set the filename we will store thre GUI
    '//of the Coolbar
    m_CoolbarSaver.FileName = App.Path & "\GUI.fil"
    '//subclass
    Call m_CoolbarSaver.SubClass(CoolBar1)
    '//read the file .. if the file not exitsts
    '//or the file is empty then will return
    '//false and load default GUI
    Call m_CoolbarSaver.DoRead
End Sub

Private Sub Form_Load()
    '//set classes
    
    '//create class
    Set m_Timer = New CTimer
    '//set the interval
    With m_Timer
        .Interval = 1000
    End With
    
    '//create the class
    Set m_HootToolbar1 = New clsHookToolbar
    With m_HootToolbar1
        '//subclassing
        Call .SubClass(Toolbar1)
        '//set help texts
        .HelpText(1) = "This is Toolbar1 . Button1"
        .HelpText(2) = "This is Toolbar1 . Button2"
        .HelpText(4) = "This is Toolbar1 . Button3"
        .HelpText(5) = "This is Toolbar1 . Button4"
    End With
    
    '//same above
    Set m_HootToolbar2 = New clsHookToolbar
    With m_HootToolbar2
        Call .SubClass(Toolbar2)
        .HelpText(1) = "This is Toolbar2 . Button1"
        .HelpText(2) = "This is Toolbar2 . Button2"
        .HelpText(3) = "This is Toolbar2 . Button3"
        .HelpText(4) = "This is Toolbar2 . Button4"
    End With
    
    '//same above
    Set m_HootToolbar3 = New clsHookToolbar
    With m_HootToolbar3
        Call .SubClass(Toolbar3)
        .HelpText(1) = "This is Toolbar3 . Button1"
        .HelpText(2) = "This is Toolbar3 . Button2"
        .HelpText(3) = "This is Toolbar3 . Button3"
    End With
    
    '//create class
    Set m_HookMenu = New clsMenus
    '//subclass
    Call m_HookMenu.SubClass(Me)
    
    'Call MsgBox("Don't Forget To Vote .", vbInformation, "Please Vote")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//kill classes
    Set m_Timer = Nothing
    Set m_HootToolbar1 = Nothing
    Set m_HootToolbar2 = Nothing
    Set m_HootToolbar3 = Nothing
    Set m_HookMenu = Nothing
    '//save the Last GUI first to the stored file
    '//if error then return false
    Call m_CoolbarSaver.DoSave
    Set m_CoolbarSaver = Nothing
End Sub

Private Sub m_HookMenu_MouseMove(ByVal MenuName As String, ByVal IsSeparator As Boolean)
    If IsSeparator Then
        StatusBar1.Panels(2).Text = "Mouse Move Over Separator"
    Else
        StatusBar1.Panels(2).Text = "Mouse Move Over " & MenuName
    End If
End Sub

Private Sub m_HookMenu_MouseOut()
    StatusBar1.Panels(2).Text = "Mouse not on any hooked ."
End Sub

Private Sub m_HootToolbar1_ButtonMouseMove(ByVal Index As Long, ByVal IsSeparator As Boolean)
    If IsSeparator Then
        StatusBar1.Panels(2).Text = "Mouse move over separator"
    Else
        StatusBar1.Panels(2).Text = m_HootToolbar1.HelpText(Index)
    End If
End Sub

Private Sub m_HootToolbar1_ButtonMouseOut()
    StatusBar1.Panels(2).Text = "Mouse not on any hooked ."
End Sub

Private Sub m_HootToolbar2_ButtonMouseMove(ByVal Index As Long, ByVal IsSeparator As Boolean)
    If IsSeparator Then
        StatusBar1.Panels(2).Text = "Mouse move over separator"
    Else
        StatusBar1.Panels(2).Text = m_HootToolbar2.HelpText(Index)
    End If
End Sub

Private Sub m_HootToolbar2_ButtonMouseOut()
    StatusBar1.Panels(2).Text = "Mouse not on any hooked ."
End Sub

Private Sub m_HootToolbar3_ButtonMouseMove(ByVal Index As Long, ByVal IsSeparator As Boolean)
    If IsSeparator Then
        StatusBar1.Panels(2).Text = "Mouse move over separator"
    Else
        StatusBar1.Panels(2).Text = m_HootToolbar3.HelpText(Index)
    End If
End Sub

Private Sub m_HootToolbar3_ButtonMouseOut()
    StatusBar1.Panels(2).Text = "Mouse not on any hooked ."
End Sub

Private Sub m_Timer_Time()
    StatusBar1.Panels(1).Text = Time
End Sub

