VERSION 5.00
Object = "*\AScrollLbl.vbp"
Begin VB.Form Scrolllblfrm 
   Caption         =   "ArcSoftware Design@Hotmail.com"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin ScrollLbl.ScrollLabel ScrollLabel4 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ForeColor       =   65535
      Picture         =   "Scrolllblfrm.frx":0000
      Caption         =   ""
   End
   Begin ScrollLbl.ScrollLabel ScrollLabel3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ForeColor       =   255
      Picture         =   "Scrolllblfrm.frx":08DA
   End
   Begin ScrollLbl.ScrollLabel ScrollLabel2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Picture         =   "Scrolllblfrm.frx":11B4
      Caption         =   ""
   End
   Begin ScrollLbl.ScrollLabel ScrollLabel1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ForeColor       =   65535
      Picture         =   "Scrolllblfrm.frx":6C35
   End
End
Attribute VB_Name = "Scrolllblfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Stop" Then
Command1.Caption = "Start"
Else
Command1.Caption = "Stop"
End If
If Command1.Caption = "Start" Then
ScrollLabel1.StartIt
ScrollLabel2.StartIt
ScrollLabel3.StartIt
ScrollLabel4.StartIt
End If
If Command1.Caption = "Stop" Then
ScrollLabel1.StopIt
ScrollLabel2.StopIt
ScrollLabel3.StopIt
ScrollLabel4.StopIt
End If
End Sub

Private Sub Form_Load()
'Start the timers and set the scroll speed and movement
'Failure to set the speed and movement will result in
'a label that doesnt move.
ScrollLabel1.StartIt
ScrollLabel1.Speed (40)
ScrollLabel1.Movement (40)
ScrollLabel2.StartIt
ScrollLabel2.Speed (10)
ScrollLabel2.Movement (50)
ScrollLabel3.StartIt
ScrollLabel3.Speed (20)
ScrollLabel3.Movement (70)
ScrollLabel4.StartIt
ScrollLabel4.Speed (20)
ScrollLabel4.Movement (10)
End Sub
