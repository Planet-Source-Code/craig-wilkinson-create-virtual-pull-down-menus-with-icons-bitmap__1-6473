VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lockout - Desktop Security"
   ClientHeight    =   2730
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4635
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "craig@dynamicdesigns.co.uk"
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   1020
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "This Version Modified By Craig Wilkinson"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   4035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Create Menus With Icons"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Based On ICO_menu From FredJust"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   4035
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   360
      Top             =   0
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   13
      Left            =   1500
      Picture         =   "MainMenu.frx":0E42
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   12
      Left            =   1500
      Picture         =   "MainMenu.frx":0F44
      Top             =   1860
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   11
      Left            =   1500
      Picture         =   "MainMenu.frx":1046
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   10
      Left            =   1140
      Picture         =   "MainMenu.frx":1148
      Top             =   1860
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   9
      Left            =   1140
      Picture         =   "MainMenu.frx":124A
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   8
      Left            =   780
      Picture         =   "MainMenu.frx":134C
      Top             =   1860
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   7
      Left            =   780
      Picture         =   "MainMenu.frx":144E
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   6
      Left            =   420
      Picture         =   "MainMenu.frx":1550
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   420
      Picture         =   "MainMenu.frx":1652
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   420
      Picture         =   "MainMenu.frx":1754
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   420
      Picture         =   "MainMenu.frx":1856
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "MainMenu.frx":1958
      Top             =   2100
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   60
      Picture         =   "MainMenu.frx":1A5A
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   60
      Picture         =   "MainMenu.frx":1B5C
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image SideBar 
      Height          =   1830
      Index           =   1
      Left            =   0
      Picture         =   "MainMenu.frx":1C5E
      Stretch         =   -1  'True
      Top             =   -420
      Width           =   345
   End
   Begin VB.Menu Menu 
      Caption         =   "&File"
   End
   Begin VB.Menu View_Logs 
      Caption         =   "&View Logs"
   End
   Begin VB.Menu Online_Menu 
      Caption         =   "&Online"
   End
   Begin VB.Menu Help_Menu 
      Caption         =   "&Help"
   End
   Begin VB.Menu About_Menu 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-===============================-
' Original Code By FredJust - IcoMenu Form
' http://fred.just.free.fr/
' fred.just@free.fr
' fredjust@hotmail.com
'-===============================-
' Updated By Craig Wilkinson 8th March 2000
' http://www.dynamicdesigns.co.uk
' http://www.lockout.co.uk
' craig@dynamicdesigns.co.uk
' lockout@lockout.co.uk
'-===============================-

' To Remove The Side Border Open CreateMenu Form
' Select SideBar And Set Visible = False


Option Explicit
Dim WithEvents VirtualFileMenu As CreateMenu
Attribute VirtualFileMenu.VB_VarHelpID = -1
Dim WithEvents VirtualViewLogsMenu As CreateMenu
Attribute VirtualViewLogsMenu.VB_VarHelpID = -1
Dim WithEvents VirtualOnlineMenu As CreateMenu
Attribute VirtualOnlineMenu.VB_VarHelpID = -1
Dim WithEvents VirtualHelpMenu As CreateMenu
Attribute VirtualHelpMenu.VB_VarHelpID = -1
Dim WithEvents VirtualAboutMenu As CreateMenu
Attribute VirtualAboutMenu.VB_VarHelpID = -1


Private Sub Form_Load()
Me.Height = 2010                ' Set Height Of Main Form
End Sub

Private Sub Menu_Click()
Set VirtualFileMenu = New CreateMenu

If VirtualFileMenu.SideBar(1).Visible = True Then
VirtualFileMenu.ImageBorder(0).Left = 420
VirtualFileMenu.Width = 2400                                                                ' Set Width Of Menu Including Side Bar
Else
VirtualFileMenu.Width = 2400 - VirtualFileMenu.SideBar(1).Width        ' Set Width Of Menu Without Side Bar
VirtualFileMenu.ImageBorder(0).Left = 60
End If

With VirtualFileMenu
.SetItem 0, "&Configure Lockout", MainMenu.Image1(0), "ConfigureLockout"
.SetItem 1, "&Protect System", MainMenu.Image1(1), "ProtectSystem"
.SetItem 2, "BLANK", , "BLANK"
.SetItem 3, "&Exit", MainMenu.Image1(2), "Exit"
End With

VirtualFileMenu.ShowMenu 0 + MainMenu.Left, MainMenu.Top + MainMenu.Height - MainMenu.ScaleHeight - 50
End Sub

Private Sub View_Logs_Click()
Set VirtualViewLogsMenu = New CreateMenu

If VirtualViewLogsMenu.SideBar(1).Visible = True Then
VirtualViewLogsMenu.ImageBorder(0).Left = 420
VirtualViewLogsMenu.Width = 3280                                                                         ' Set Width Of Menu Including Side Bar
Else
VirtualViewLogsMenu.ImageBorder(0).Left = 60
VirtualViewLogsMenu.Width = 3280 - VirtualViewLogsMenu.SideBar(1).Width         ' Set Width Of Menu Without Side Bar
End If

With VirtualViewLogsMenu
.SetItem 0, "View Lockout Daily Logs", MainMenu.Image1(3), "ViewDailyLogs"
.SetItem 1, "View Lockout Internet Log", MainMenu.Image1(4), "ViewInternetLog"
.SetItem 2, "View Lockout File Blocks", MainMenu.Image1(5), "ViewLockoutFileBlocks"
.SetItem 3, "BLANK", , "BLANK"                                                                                          'Insert A Blank Divider
.SetItem 4, "Internal Log Viewer", MainMenu.Image1(6), "InternalLogViewer"
.SetItem 5, "BLANK2", , "BLANK2"                                                                                        'Insert Second Blank Divider
End With
VirtualViewLogsMenu.ShowMenu 460 + MainMenu.Left, MainMenu.Top + MainMenu.Height - MainMenu.ScaleHeight - 50
End Sub

Private Sub Online_Menu_Click()
Set VirtualOnlineMenu = New CreateMenu


If VirtualOnlineMenu.SideBar(1).Visible = True Then
VirtualOnlineMenu.ImageBorder(0).Left = 420
VirtualOnlineMenu.Width = 2500                                                                           ' Set Width Of Menu Including Side Bar
Else
VirtualOnlineMenu.ImageBorder(0).Left = 60
VirtualOnlineMenu.Width = 2500 - VirtualOnlineMenu.SideBar(1).Width             ' Set Width Of Menu Without Side Bar
End If

VirtualOnlineMenu.SetItem 0, "&Lockout Home Page", MainMenu.Image1(7), "LockoutHomePage"
VirtualOnlineMenu.SetItem 1, "Check For &Updates", MainMenu.Image1(8), "CheckForUpdates"

VirtualOnlineMenu.ShowMenu 1380 + MainMenu.Left, MainMenu.Top + MainMenu.Height - MainMenu.ScaleHeight - 50
End Sub

Private Sub Help_Menu_Click()
Set VirtualHelpMenu = New CreateMenu

If VirtualHelpMenu.SideBar(1).Visible = True Then
VirtualHelpMenu.ImageBorder(0).Left = 420
VirtualHelpMenu.Width = 2050                                                                                 ' Set Width Of Menu Including Side Bar
Else
VirtualHelpMenu.ImageBorder(0).Left = 60
VirtualHelpMenu.Width = 2050 - VirtualHelpMenu.SideBar(1).Width                         ' Set Width Of Menu Without Side Bar
End If

With VirtualHelpMenu
.SetItem 0, "H&elp Contents", MainMenu.Image1(9), "HelpContents"
.SetItem 1, "More &Info", MainMenu.Image1(10), "MoreInfo"
End With

VirtualHelpMenu.ShowMenu 1990 + MainMenu.Left, MainMenu.Top + MainMenu.Height - MainMenu.ScaleHeight - 50
End Sub

Private Sub About_Menu_Click()
Set VirtualAboutMenu = New CreateMenu

If VirtualAboutMenu.SideBar(1).Visible = True Then
VirtualAboutMenu.ImageBorder(0).Left = 420
VirtualAboutMenu.Width = 2100                                                                               ' Set Width Of Menu Including Side Bar
Else
VirtualAboutMenu.ImageBorder(0).Left = 60
VirtualAboutMenu.Width = 2100 - VirtualAboutMenu.SideBar(1).Width                   ' Set Width Of Menu Without Side Bar
End If

With VirtualAboutMenu
.SetItem 0, "&Register", MainMenu.Image1(11), "Register"
.SetItem 1, "&Version Info", MainMenu.Image1(12), "VersionInfo"
.SetItem 2, "BLANK", , "BLANK"
.SetItem 3, "&About Lockout", MainMenu.Image1(13), "AboutLockout"
.SetItem 4, "BLANK2", , "BLANK2"
End With

VirtualAboutMenu.ShowMenu 2500 + MainMenu.Left, MainMenu.Top + MainMenu.Height - MainMenu.ScaleHeight - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next ' Skip If Form Not Loaded (Form May Not Have Been Selected)
Unload Me
Unload CreateMenu
Unload VirtualFileMenu
Unload VirtualViewLogsMenu
Unload VirtualOnlineMenu
Unload VirtualHelpMenu
Unload VirtualAboutMenu
End
End Sub










