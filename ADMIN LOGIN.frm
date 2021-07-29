VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "ADMIN LOGIN.frx":0000
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox textusername 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   5160
      Width           =   3615
   End
   Begin VB.TextBox textpass 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdlogin 
      Appearance      =   0  'Flat
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      MaskColor       =   &H00400000&
      Picture         =   "ADMIN LOGIN.frx":1D165
      TabIndex        =   2
      Top             =   6705
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   480
      Left            =   8880
      TabIndex        =   6
      Top             =   4320
      Width           =   2280
   End
   Begin VB.Label changepass 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password ?"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   9270
      TabIndex        =   5
      Top             =   7800
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   8280
      TabIndex        =   4
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   8310
      TabIndex        =   3
      Top             =   5760
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   9120
      Picture         =   "ADMIN LOGIN.frx":1EDA8
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public nn As New ADODB.Recordset

Private Sub changepass_Click()
Form2.Show
Unload Me
con.Close
End Sub

Private Sub cmdlogin_Click()
rs.Open "select *from login", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
Dim i As Integer
i = 0
While Not rs.EOF = True
If rs.Fields("username").Value = textusername.Text And rs.Fields("password").Value = textpass.Text Then
i = i + 1
Form3.Show
Unload Me
End If
rs.MoveNext
Wend
If i = 0 Then
MsgBox "User Does Not Exist", vbOKOnly, "Invalid"
textusername = ""
textpass = ""
textusername.SetFocus
End If
rs.Close
End Sub

Private Sub Command1_Click()
nn.CursorLocation = adUseClient
nn.Open "select *from login", con, adOpenDynamic, adLockPessimistic
a = InputBox("enter a username")
b = InputBox("enter password")
nn.AddNew
nn.Fields("username").Value = a
nn.Fields("password").Value = b
nn.Update
nn.Close
End Sub

Private Sub Form_Load()
WindowState = 2
con.Open "Provider=MSDASQL.1;Password=root;Persist Security Info=True;User ID=root;Data Source=yogesh"
rs.CursorLocation = adUseClient
nn.CursorLocation = adUseClient
End Sub
