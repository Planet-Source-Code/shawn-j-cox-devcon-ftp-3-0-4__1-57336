VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1890
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   1890
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   2490
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1110
      Top             =   780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "An FTP Client written in Visual Basic"
      Height          =   210
      Left            =   255
      TabIndex        =   2
      Top             =   1485
      Width           =   3120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 3.0.4 Fatty Editon"
      Height          =   450
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DevCon FTP "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
    frmMain.Show
    Unload Me
End Sub
