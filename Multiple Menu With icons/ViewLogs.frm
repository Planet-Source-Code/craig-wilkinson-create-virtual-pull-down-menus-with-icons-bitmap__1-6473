VERSION 5.00
Begin VB.Form CreateMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Blank3 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   5040
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Blank2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -60
      X2              =   5100
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Blank1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   5100
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label MenuLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Text Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
   Begin VB.Image Selector 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   -60
      Top             =   90
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Image ImageBorder 
      Height          =   285
      Index           =   0
      Left            =   420
      Picture         =   "ViewLogs.frx":0000
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image MenuIcon 
      Height          =   240
      Index           =   0
      Left            =   420
      Top             =   540
      Width           =   240
   End
   Begin VB.Image SideBar 
      Height          =   2100
      Index           =   1
      Left            =   0
      Picture         =   "ViewLogs.frx":00EA
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Line Blank4 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   -60
      X2              =   5040
      Y1              =   2940
      Y2              =   2940
   End
End
Attribute VB_Name = "CreateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LastIcon As Long
Private MaxWidth As Long

Const MenuHeight = 285

Public ForeColorOver As Long
Public BackColorOver As Long

Public Event Click(ByVal Index As Long, Tag As String)
Public Event Closed()
Public Event MouseDown(ByVal Index As Long, Tag As String, Button As Integer, Shift As Integer)
Public Event MouseUp(ByVal Index As Long, Tag As String, Button As Integer, Shift As Integer)
Public Event MouseMove(ByVal Index As Long, Tag As String)

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    
Public Sub ShowMenu(Optional ByVal Left As Long = -1, Optional ByVal Top As Long = -1)
    Dim i As Long
    Dim CurPos As POINTAPI
    
    Me.Height = MenuHeight * MenuLabel.Count + 90
    
   SideBar(1).Height = MenuHeight * MenuLabel.Count
    
    
   If Left = -1 Then
    GetCursorPos CurPos
        Me.Left = CurPos.X * Screen.TwipsPerPixelX
        Me.Top = CurPos.Y * Screen.TwipsPerPixelY
    Else
        Me.Left = Left
        Me.Top = Top
    End If
      
    Call MenuLabel_MouseMove(0, 0, 0, 0, 0)
    Me.Show
End Sub
Public Sub SetItem(ByVal Index As Long, ByVal Caption As String, Optional Icon As IPictureDisp, Optional Key As String)
 
On Error Resume Next ' Errors Here If Already Loaded
Load MenuLabel(Index)
Load MenuIcon(Index)
Load Selector(Index)

' 1st Blank Line
If Caption$ = "BLANK" Then
Me.Blank1.Y1 = MenuHeight * Index + 150: Me.Blank1.Y2 = MenuHeight * Index + 150
Me.Blank2.Y1 = MenuHeight * Index + 150: Me.Blank2.Y2 = MenuHeight * Index + 150

If Me.SideBar(1).Visible = True Then Me.Blank1.X1 = 420: Me.Blank2.X1 = 420 Else Me.Blank1.X1 = 42: Me.Blank2.X1 = 42

Exit Sub
End If

' 2nd Blank Line
If Caption$ = "BLANK2" Then
Me.Blank3.Y1 = MenuHeight * Index + 150: Me.Blank3.Y2 = MenuHeight * Index + 150: Me.Blank3.X1 = 420
Me.Blank4.Y1 = MenuHeight * Index + 150: Me.Blank4.Y2 = MenuHeight * Index + 150: Me.Blank4.X1 = 420
If Me.SideBar(1).Visible = True Then Me.Blank3.X1 = 420: Me.Blank4.X1 = 420 Else Me.Blank3.X1 = 42: Me.Blank4.X1 = 42

Exit Sub
End If

' !!Add More Lines And Above Code If More Blanks Required On Menu

MenuLabel(Index).Caption = Space(2) & Caption

If SideBar(1).Visible Then MenuLabel(Index).Left = 800 Else MenuLabel(Index).Left = 440
MenuLabel(Index).Width = 2790
MenuLabel(Index).Visible = True
    
Set MenuIcon(Index).Picture = Icon
    

MenuLabel(Index).Tag = Key
MenuLabel(Index).Top = MenuHeight * Index + 25
MenuLabel(Index).Visible = True


MenuIcon(Index).Top = MenuHeight * Index + 15
If SideBar(1).Visible Then MenuIcon(Index).Left = 440 Else MenuIcon(Index).Left = 70
MenuIcon(Index).Visible = True

Selector(Index).Top = MenuHeight * Index + 15
Selector(Index).Visible = True

LastIcon = MenuLabel.UBound
End Sub
    
    
Private Sub Form_LostFocus()
   Unload Me
   End Sub
    
Private Sub Form_Initialize()
    LastIcon = 255
    BackColorOver = vbHighlight
    ForeColorOver = vbHighlightText
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub MenuIcon_Click(Index As Integer)
Call MenuLabel_Click(Index)
End Sub

Private Sub ImageBorder_Click(Index As Integer)
Call MenuLabel_Click(Index)
End Sub

Private Sub Selector_Click(Index As Integer)
Call MenuLabel_Click(Index)
End Sub

Private Sub Selector_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MenuLabel_MouseMove(Index, Button, Shift, X, Y)
End Sub
Private Sub MenuIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MenuLabel_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub MenuLabel_Click(Index As Integer)
'Enter Function Calls Here
If MenuLabel(Index).Tag = "ConfigureLockout" Then MsgBox "Configure Lockout"
If MenuLabel(Index).Tag = "ProtectSystem" Then MsgBox "Protect System"
If MenuLabel(Index).Tag = "Exit" Then Unload Me: End
'-----------------------------------------------------------------------------------------------
If MenuLabel(Index).Tag = "ViewDailyLogs" Then MsgBox "View Daily Logs"
If MenuLabel(Index).Tag = "ViewInternetLog" Then MsgBox "View Internet Logs"
If MenuLabel(Index).Tag = "ViewLockoutFileBlocks" Then MsgBox "View Lockout File Blocks"
If MenuLabel(Index).Tag = "InternalLogViewer" Then MsgBox "Internal Log Viewer"
'-----------------------------------------------------------------------------------------------
If MenuLabel(Index).Tag = "LockoutHomePage" Then MsgBox "Lockout Home Page (www.lockout.co.uk)"
If MenuLabel(Index).Tag = "CheckForUpdates" Then MsgBox "Check For Updates"
'-----------------------------------------------------------------------------------------------
If MenuLabel(Index).Tag = "HelpContents" Then MsgBox "Help Contents"
If MenuLabel(Index).Tag = "MoreInfo" Then MsgBox "MoreInfo"
'-----------------------------------------------------------------------------------------------
If MenuLabel(Index).Tag = "Register" Then MsgBox "Register"
If MenuLabel(Index).Tag = "VersionInfo" Then MsgBox "Version Info"
If MenuLabel(Index).Tag = "AboutLockout" Then MsgBox "About Lockout"

Unload Me
End Sub

Private Sub MenuLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If LastIcon = 255 Then GoSub Skip
If LastIcon = Index Then Exit Sub

    MenuLabel(LastIcon).BackColor = Me.BackColor
    MenuLabel(LastIcon).ForeColor = Me.ForeColor
  '  ImageBorder(LastIcon).Visible = False
    MenuLabel(LastIcon).BackStyle = 0
Skip:
    MenuLabel(Index).BackStyle = 1
    MenuLabel(Index).BackColor = BackColorOver
    MenuLabel(Index).ForeColor = ForeColorOver
    
If MenuLabel(Index).Left = 0 Then
   ImageBorder(0).Visible = False
   Else
    ImageBorder(0).Visible = True
    ImageBorder(0).Top = MenuHeight * Index
  End If
  LastIcon = Index

End Sub




