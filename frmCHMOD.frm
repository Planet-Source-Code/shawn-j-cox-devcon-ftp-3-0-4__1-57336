VERSION 5.00
Begin VB.Form frmCHMOD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHMOD Options:"
   ClientHeight    =   2805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "User &:"
      Height          =   765
      Left            =   0
      TabIndex        =   11
      Top             =   1590
      Width           =   4545
      Begin VB.TextBox txtCHMOD 
         Height          =   270
         Index           =   2
         Left            =   3930
         TabIndex        =   17
         Text            =   "0"
         Top             =   270
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   1335
      End
      Begin VB.CheckBox chkWrite 
         Caption         =   "Write"
         Height          =   210
         Index           =   2
         Left            =   1440
         TabIndex        =   13
         Top             =   345
         Width           =   1170
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Execute"
         Height          =   300
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   1620
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Common"
      Height          =   375
      Left            =   45
      TabIndex        =   9
      Top             =   2385
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2385
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      Top             =   2385
      Width           =   825
   End
   Begin VB.Frame fraUser 
      Caption         =   "Moderator &:"
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   4545
      Begin VB.TextBox txtCHMOD 
         Height          =   270
         Index           =   1
         Left            =   3930
         TabIndex        =   16
         Text            =   "0"
         Top             =   315
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Execute"
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   300
         Width           =   1620
      End
      Begin VB.CheckBox chkWrite 
         Caption         =   "Write"
         Height          =   210
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   345
         Width           =   1155
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.Frame fraAdmin 
      Caption         =   "Admin &:"
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   4545
      Begin VB.TextBox txtCHMOD 
         Height          =   270
         Index           =   0
         Left            =   3930
         TabIndex        =   15
         Text            =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Execute"
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   8
         Top             =   315
         Width           =   1620
      End
      Begin VB.CheckBox chkWrite 
         Caption         =   "Write"
         Height          =   210
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Top             =   345
         Width           =   1125
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "CHMOD #:"
      Height          =   255
      Left            =   1215
      TabIndex        =   18
      Top             =   2445
      Width           =   1470
   End
   Begin VB.Menu mnu_Common 
      Caption         =   "common"
      Visible         =   0   'False
      Begin VB.Menu mnu_Script 
         Caption         =   "CGI Script"
      End
      Begin VB.Menu mnu_All 
         Caption         =   "All Access"
      End
      Begin VB.Menu mnu_Admin 
         Caption         =   "Admin Only"
      End
   End
End
Attribute VB_Name = "frmCHMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer
Dim strCHMOD As String
' I got this idea from the project CHMOD
' only mine is back asswards
Private Sub chkDelete_Click(Index As Integer)
 Select Case Index
  Case 0
   ' code for 7 Other
   If chkDelete(Index).Value = 0 Then
    c3 = c3 - 1
   Else
    c3 = c3 + 1
   End If
   txtCHMOD(Index).Text = c3
  Case 1
   ' code for 4
   If chkDelete(Index).Value = 0 Then
    c2 = c2 - 1
   Else
    c2 = c2 + 1
   End If
   txtCHMOD(Index).Text = c2
  Case 2
   ' code for 1
   If chkDelete(Index).Value = 0 Then
    c1 = c1 - 1
   Else
    c1 = c1 + 1
   End If
   txtCHMOD(Index).Text = c1
  End Select
  Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
  strCHMOD = Right(Label1.Caption, 3)
End Sub

Private Sub chkRead_Click(Index As Integer)
 Select Case Index
  Case 0
   ' code for 7 Other
   If chkRead(Index).Value = 0 Then
    c3 = c3 - 4
   Else
    c3 = c3 + 4
   End If
   txtCHMOD(Index).Text = c3
  Case 1
   ' code for 4
   If chkRead(Index).Value = 0 Then
    c2 = c2 - 4
   Else
    c2 = c2 + 4
   End If
   txtCHMOD(Index).Text = c2
  Case 2
   ' code for 1
   If chkRead(Index).Value = 0 Then
    c1 = c1 - 4
   Else
    c1 = c1 + 4
   End If
   txtCHMOD(Index).Text = c1
  End Select
  Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
  strCHMOD = Right(Label1.Caption, 3)
End Sub

Private Sub chkWrite_Click(Index As Integer)
 Select Case Index
  Case 0
   ' code for 7 Other
   If chkWrite(Index).Value = 0 Then
    c3 = c3 - 2
   Else
    c3 = c3 + 2
   End If
   txtCHMOD(Index).Text = c3
  Case 1
   ' code for 4
   If chkWrite(Index).Value = 0 Then
    c2 = c2 - 2
   Else
    c2 = c2 + 2
   End If
   txtCHMOD(Index).Text = c2
  Case 2
   ' code for 1
   If chkWrite(Index).Value = 0 Then
    c1 = c1 - 2
   Else
    c1 = c1 + 2
   End If
   txtCHMOD(Index).Text = c1
  End Select
  Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
  strCHMOD = Right(Label1.Caption, 3)
End Sub

Private Sub cmdOK_Click()
 Dim strCHMOD2 As String
 Dim lngResult As Long, ftpSession As Long
 Dim bolResult As Boolean
 strCHMOD2 = "SITE CHMOD " & strCHMOD & " " & Klic & strCFile

'To CHMOD using the FTP Protocol, there is a FTP command
'named SITE. SITE allows you to send commands unique to the
'FTP server, which other servers have no need for.
'(and as such are not specified in the FTP Protocol document RCF759
'I used the FtpCommand API from wininet to
'send the SITE command.
'SITE usage is as follows:
' SITE CHMOD 0777 for read/write/delete all users
' The SITE command is not supported by all FTPd's, you might still
'need to open up the telnet terminal (if your host allows it) and manually
'CHMOD files from there. There is nothing I can do to fix that one.

bolResult = FtpCommand(session, True, FTP_TRANSFER_TYPE_ASCII, strCHMOD2, poo, lngResult)
MsgBox lngResult & "//" & bolResult & "//" & strCHMOD2
'==========FtpCommand API (wininet.dll)==================
'ByVal hConnect as long
' This is the handle returned by the InternetOpen API
'ByVal fExpectResponse As Boolean
' TRUE/FALSE if you expect a response, such as a success response.
'ByVal dwFlags As Long
' FTP_TRANSFER_TYPE_BINARY
' FTP_TRANSFER_TYPE_ASCII
' Specify type of connection for this command.
'ByVal lpszCommand As String
' Command to send to FTP in string format
'ByVal dwContext As Long
'
'ByVal phFtpCommand As Long
' long variable to recieve response.
End Sub

Private Sub Command1_Click()
 PopupMenu mnu_Common
End Sub


Private Sub mnu_Admin_Click()
'Typical settings for files are 777, 755, 666 or 644.
'Typical settings for directories are 777 or 755.
'Cgi scripts 755, data files 666, and configuration files 644.
 Dim i As Integer
 txtCHMOD(0).Text = "0"
 txtCHMOD(1).Text = "0"
 txtCHMOD(2).Text = "7"
 For i = 0 To 2
  Select Case i
   Case 0
    chkRead(0).Value = 1
    chkWrite(0).Value = 1
    chkDelete(0).Value = 1
   Case 1
    chkRead(i).Value = 0
    chkWrite(i).Value = 0
    chkDelete(i).Value = 0
   Case 2
    chkRead(i).Value = 0
    chkWrite(i).Value = 0
    chkDelete(i).Value = 0
   End Select
 Next i
 Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
 strCHMOD = Right(Label1.Caption, 3)
End Sub

Private Sub mnu_All_Click()
' All Access means every user has ability to read/write/delete files.
' Usually this is only given to user files, like on a Forum.
 txtCHMOD(0).Text = "7"
 txtCHMOD(1).Text = "7"
 txtCHMOD(2).Text = "7"
 Dim i As Integer
 For i = 0 To 2
  chkRead(i).Value = 1
  chkWrite(i).Value = 1
  chkDelete(i).Value = 1
 Next i
 Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
 strCHMOD = Right(Label1.Caption, 3)
End Sub

Private Sub mnu_Script_Click()
 Dim i As Integer
 txtCHMOD(0).Text = "5"
 txtCHMOD(1).Text = "5"
 txtCHMOD(2).Text = "7"
 For i = 0 To 2
  Select Case i
   Case 0
    chkRead(0).Value = 1
    chkWrite(0).Value = 1
    chkDelete(0).Value = 1
   Case 1
    chkRead(1).Value = 1
    chkWrite(1).Value = 0
    chkDelete(1).Value = 1
   Case 2
    chkRead(2).Value = 1
    chkWrite(2).Value = 0
    chkDelete(2).Value = 2
   End Select
 Next i
 Label1.Caption = "CHMOD #: " & txtCHMOD(0).Text & txtCHMOD(1).Text & txtCHMOD(2).Text
 strCHMOD = Right(Label1.Caption, 3)
End Sub
