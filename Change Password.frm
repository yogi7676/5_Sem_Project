VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Change Password.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textusername 
      Appearance      =   0  'Flat
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
      Left            =   8160
      TabIndex        =   0
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox Textnwpass 
      Appearance      =   0  'Flat
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
      Left            =   8160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox textconpass 
      Appearance      =   0  'Flat
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
      Left            =   8160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "RESET"
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
      Left            =   8160
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdverify 
      Caption         =   "VERIFY USERNAME"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "Change Password.frx":1D165
      TabIndex        =   1
      Top             =   4500
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   9360
      Picture         =   "Change Password.frx":1D544
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RESET PASSWORD"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   8640
      TabIndex        =   8
      Top             =   3000
      Width           =   2595
   End
   Begin VB.Label labelusername 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   8160
      TabIndex        =   7
      Top             =   3600
      Width           =   1290
   End
   Begin VB.Label labelnew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD"
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
      Left            =   8160
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Labelconfirm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   2400
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset

Private Sub cmdreset_Click()
On Error GoTo err:
If Textnwpass.Text = textconpass.Text Then
rs.Fields(1).Value = textconpass.Text
rs.Update
MsgBox "Password Successfully Updated", vbInformation + vbOKOnly, "Change password"
con.Close
Form1.Show
Unload Me
Else
MsgBox "Password Not Matched", vbInformation + vbOKOnly, "Change password"
Text1 = ""
Text2 = ""
Text1.SetFocus
End If
err:
err.clear
MsgBox "Do Not use Old Password again", vbOKOnly + vbInformation, "Password Repetation"
End Sub

Private Sub cmdverify_Click()
If rs.Fields(0).Value = textusername.Text Then
labelusername.Visible = False
textusername.Visible = False
cmdverify.Visible = False
labelnew.Visible = True
labelnew.Top = 3600
Textnwpass.Visible = True
Textnwpass.Top = 3840
Labelconfirm.Visible = True
Labelconfirm.Top = 4500
textconpass.Visible = True
textconpass.Top = 4740
cmdreset.Visible = True
cmdreset.Top = 5500
Textnwpass.SetFocus
Else
MsgBox "Invalid User", vbOKCancel + vbDefaultButton2 + vbCritical, "User Validation"
textusername = ""
textusername.SetFocus
End If
End Sub

Private Sub Form_Load()
WindowState = 2
con.Open "Provider=MSDASQL.1;Password=root;Persist Security Info=True;User ID=root;Data Source=yogesh"
rs.CursorLocation = adUseClient
rs.Open "select *from login", con, adOpenDynamic, adLockPessimistic
End Sub
Private Sub Textnwpass_LostFocus()
Dim byt As Byte
Dim aa As Integer
Dim strc As String
If Len(Textnwpass.Text) >= 8 Then
For aa = 1 To Len(Textnwpass.Text)
strc = Mid$(Textnwpass.Text, aa, 1)
If strc >= "!" And strc <= "/" Then
byt = byt Or &H8
End If
If strc >= "0" And strc <= "9" Then
byt = byt Or &H4
End If
If strc >= ":" And strc <= "@" Then
byt = byt Or &H8
End If
If strc >= "A" And strc <= "Z" Then
byt = byt Or &H2
End If
If strc >= "a" And strc <= "z" Then
byt = byt Or &H1
End If
Next aa
End If
If byt <> &HF Then
MsgBox "Password Should Contain atleast one Uppercase Letter,Lowercase Letter,One Number and a Special Character And Must Contain Minimum of 8 Charcters", vbOKOnly + vbInformation, "Password Information"
Textnwpass = ""
Textnwpass.SetFocus
End If
End Sub
