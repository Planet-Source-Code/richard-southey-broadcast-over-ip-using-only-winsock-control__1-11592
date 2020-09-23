VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBroadcast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Broadcast Using UDP"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4290
   Icon            =   "frmBroadcast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReceived 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton cmdBroadcast 
      Caption         =   "Broadcast"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock udpBroadcast 
      Left            =   3600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblReceived 
      BackStyle       =   0  'Transparent
      Caption         =   "Received Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblMail 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "richard@richsoftcomputing.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmBroadcast.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "mailto:richard@richsoftcomputing.cjb.net?subject=Broadcast Over UDP"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "www.richsoftcomputing.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmBroadcast.frx":0594
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "www.richsoftcomputing.cjb.net"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblData 
      Caption         =   "Data to broadcast:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Sending data over UDP to address 255.255.255.255 will broadcast to all computiers on the network."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmBroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

Const BROADCASTPORT = 1055
'API Call which drives the Hyperlink
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Sub HyperJump(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub
Private Sub cmdBroadcast_Click()
    'Broadcast the data in txtData
    'Ignore error 126 which sometimes occurs
    On Error GoTo ErrorHandler
    udpBroadcast.SendData txtData.Text
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 126 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
    Resume Next
        
End Sub

Private Sub Form_Load()
    'Set up the Winsock control
    With udpBroadcast
        .Protocol = sckUDPProtocol
        .LocalPort = BROADCASTPORT
        .RemotePort = BROADCASTPORT
        .RemoteHost = "255.255.255.255" ' This is the broadcast IP
    End With
End Sub


Private Sub lblEmail_Click()
    'Activate the default e-mail client
    HyperJump lblEmail.Tag
End Sub

Private Sub lblWebsite_Click()
    'Activate the hyperlink
    HyperJump lblWebsite.Tag
End Sub

Private Sub udpBroadcast_DataArrival(ByVal bytesTotal As Long)
    Dim IncomingData As String
    'Data has arrived so display it
    udpBroadcast.GetData IncomingData
    txtReceived.Text = IncomingData
End Sub

Private Sub udpBroadcast_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'An error has occurred so display a description
    MsgBox Err.Description, vbCritical, Err.Source
End Sub




