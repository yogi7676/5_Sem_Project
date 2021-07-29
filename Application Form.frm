VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Application Form"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   Picture         =   "Application Form.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "DONE"
      Height          =   495
      Left            =   1560
      TabIndex        =   103
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Height          =   735
      Left            =   -120
      TabIndex        =   102
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   1560
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   735
      Left            =   -120
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   3600
      TabIndex        =   99
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3600
      TabIndex        =   98
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   3600
      TabIndex        =   95
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3600
      TabIndex        =   94
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   3600
      TabIndex        =   93
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3600
      TabIndex        =   92
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   30
      Left            =   3120
      TabIndex        =   91
      Top             =   2760
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   30
      Left            =   2520
      TabIndex        =   90
      Top             =   1800
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   30
      Left            =   3720
      TabIndex        =   89
      Top             =   480
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox queaadharveri 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   17280
      MaxLength       =   12
      TabIndex        =   82
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton quecommandverify 
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   17280
      TabIndex        =   81
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox quecombo1 
      Height          =   315
      Left            =   17280
      TabIndex        =   80
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton queuedone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   17280
      TabIndex        =   79
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2640
      Picture         =   "Application Form.frx":1D165
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Cancel"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   405
      Left            =   1560
      TabIndex        =   30
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Format          =   96206849
      CurrentDate     =   43744
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   405
      Left            =   1560
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Format          =   96272385
      CurrentDate     =   43744
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   405
      Left            =   1560
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Format          =   96272385
      CurrentDate     =   43744
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   8040
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   714
      _Version        =   393216
      Format          =   96272385
      CurrentDate     =   43744
   End
   Begin VB.OptionButton Optionyes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.OptionButton Optionno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13080
      TabIndex        =   104
      Top             =   5880
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.TextBox Textfirstname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox textmidname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox textlastname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.OptionButton Optionmale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Male"
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   8040
      TabIndex        =   97
      Top             =   4080
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.OptionButton Optionfemale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Female"
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   10080
      TabIndex        =   96
      Top             =   4080
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.OptionButton Optionothers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Others"
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   12120
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox texttmpadd 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   8040
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   4680
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox textdistrict 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textpicode 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   11880
      MaxLength       =   6
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox textperadd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   6360
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox textperdist 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textperpin 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   11880
      MaxLength       =   6
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox textreligion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   12
      Text            =   "Text10"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textcaste 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   11880
      TabIndex        =   13
      Text            =   "Text11"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox textemail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   8040
      TabIndex        =   14
      Text            =   "Text12"
      Top             =   8760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textphno 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   11880
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "Text13"
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   3135
      TabIndex        =   36
      Top             =   9360
      Visible         =   0   'False
      Width           =   3135
      Begin VB.OptionButton Optionfyes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Microsoft Himalaya"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1515
      End
      Begin VB.OptionButton Optionfno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Microsoft Himalaya"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   0
         Width           =   1570
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      ScaleHeight     =   375
      ScaleWidth      =   3135
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
      Begin VB.OptionButton Optionapl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "APL"
         BeginProperty Font 
            Name            =   "Microsoft Himalaya"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Top             =   0
         Width           =   1570
      End
      Begin VB.OptionButton OptionBpl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "BPL"
         BeginProperty Font 
            Name            =   "Microsoft Himalaya"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.TextBox Textaadhar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   400
      Left            =   8040
      MaxLength       =   12
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Textn1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Texta1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Textr1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   20
      Text            =   "Text3"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Textn2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Texta2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   24
      Text            =   "Text5"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Textr2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   25
      Text            =   "Text6"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Textn3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   27
      Text            =   "Text7"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Texta3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Textr3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   1560
      TabIndex        =   29
      Text            =   "Text9"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1560
      Picture         =   "Application Form.frx":1D2EF
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Save "
      Top             =   8760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label queselect 
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   405
      Left            =   15960
      TabIndex        =   88
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label queaadhar 
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhar No"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   400
      Left            =   15840
      TabIndex        =   87
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label quename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   17280
      TabIndex        =   86
      Top             =   8400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label quecopyaadhar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   15960
      TabIndex        =   85
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label queuecardtype 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15960
      TabIndex        =   84
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label queuehead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Queue"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   15840
      TabIndex        =   83
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   78
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   77
      Top             =   2880
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mid Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   76
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   75
      Top             =   4080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary Address"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   6240
      TabIndex        =   74
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Address"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   975
      Left            =   6240
      TabIndex        =   73
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Is Permanent Address  Same as Temporary Address"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   72
      Top             =   5880
      Visible         =   0   'False
      Width           =   5670
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   71
      Top             =   7560
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   70
      Top             =   8160
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   11040
      TabIndex        =   69
      Top             =   8280
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   68
      Top             =   8760
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone no"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   10800
      TabIndex        =   67
      Top             =   8760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   66
      Top             =   5280
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   10920
      TabIndex        =   65
      Top             =   5280
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   64
      Top             =   7080
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   10920
      TabIndex        =   63
      Top             =   6960
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   9975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5400
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      Height          =   9135
      Left            =   5760
      Top             =   960
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   9975
      Left            =   14760
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5400
   End
   Begin VB.Label Labelappli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICATION FORM"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   62
      Top             =   1080
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Want To Add Family Members"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6240
      TabIndex        =   61
      Top             =   9480
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   360
      TabIndex        =   60
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   360
      TabIndex        =   59
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   58
      Top             =   240
      Width           =   4935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GOVERNMENT OF KARNATAKA"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6120
      TabIndex        =   57
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   56
      Top             =   6840
      Width           =   4695
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Food,Civil Supplies and Consumer Affairs Department"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   55
      Top             =   480
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   2955
      Left            =   15120
      Picture         =   "Application Form.frx":1D959
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   4680
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   15120
      TabIndex        =   54
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   15120
      Picture         =   "Application Form.frx":33495
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   15000
      TabIndex        =   53
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Card Type"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   6240
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Labelid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8160
      TabIndex        =   51
      Top             =   1440
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhar No"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   6240
      TabIndex        =   50
      Top             =   3480
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Labelfamid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   480
      TabIndex        =   49
      Top             =   1080
      Width           =   75
   End
   Begin VB.Label relname1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   48
      Top             =   1560
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label relaadhar1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhar"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label rel1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relation"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label reldob1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   45
      Top             =   3360
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label relname2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label relaadhar2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhar"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label rel2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relation"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   42
      Top             =   5160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label reldob2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label relname3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   40
      Top             =   6360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label relaadhar3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhar"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   39
      Top             =   6960
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label rel3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relation"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   38
      Top             =   7560
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label reldob3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   240
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuqueue 
      Caption         =   "&Queue"
      Begin VB.Menu mnuaddqueue 
         Caption         =   "Add &Queue"
      End
      Begin VB.Menu mnurunqueue 
         Caption         =   "R&un Queue"
         Begin VB.Menu mnuveritoken 
            Caption         =   "VerifyToken"
         End
         Begin VB.Menu mnuexecute 
            Caption         =   "Executetoken"
         End
      End
      Begin VB.Menu mnuviewQueue 
         Caption         =   "ViewQueue"
      End
      Begin VB.Menu mnuclearqueue 
         Caption         =   "ClearQueue"
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu popnew 
         Caption         =   "New"
      End
      Begin VB.Menu popsave 
         Caption         =   "Save"
      End
      Begin VB.Menu popaddqueue 
         Caption         =   "Add Queue"
      End
      Begin VB.Menu popveritoken 
         Caption         =   "Verify Token"
      End
      Begin VB.Menu poprunqueue 
         Caption         =   "Execute Queue"
      End
      Begin VB.Menu popexit 
         Caption         =   "Exit"
      End
      Begin VB.Menu popclear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuprint 
      Caption         =   "Print"
      Begin VB.Menu mnucard 
         Caption         =   "Ca&rdPrint"
      End
   End
   Begin VB.Menu mnuupdate 
      Caption         =   "&Update"
      Begin VB.Menu mnuuserinfo 
         Caption         =   "U&serInfo"
         Begin VB.Menu mnuBasic 
            Caption         =   "BasicInfo"
         End
         Begin VB.Menu mnuadd 
            Caption         =   "UserAddress"
         End
      End
      Begin VB.Menu mnufaminfo 
         Caption         =   "&FamillyInfo"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public coo As New ADODB.Connection
Public rs As New ADODB.Recordset
Public all As New ADODB.Recordset
Public que As New ADODB.Recordset
Public user As New ADODB.Recordset
Public useradd As New ADODB.Recordset
Public family As New ADODB.Recordset
Public view As New ADODB.Recordset
Public veri As New ADODB.Recordset
Public quex As New ADODB.Recordset

Private Sub Command1_Click()
Call save
End Sub

Private Sub Command10_Click()
veri.Close
DataGrid5.Visible = False
Command10.Visible = False
End Sub

Private Sub Command2_Click()
Dim a As String
a = MsgBox("Are you sure to Cancel", vbYesNo + vbDefaultButton2 + vbInformation, "Cancel")
If a = vbYes Then
Call famvisiblefalse
Call clear
Call uservisiblefalse
rs.Delete
Shape2.Visible = True
Labelappli.Visible = True
Label27.Visible = True
Label27.Left = 6480
Picture2.Visible = True
Label27.Top = 1800
Picture2.Top = 1800
OptionBpl.Value = False
Optionapl.Value = False
Optionfemale.Value = False
Optionmale.Value = False
Optionothers.Value = False
Optionyes.Value = False
Optionno.Value = False
End If
End Sub

Private Sub Command3_Click()
Dim msg As String
Dim aa As String
msg = MsgBox("Save User Details", vbYesNo + vbDefaultButton2 + vbInformation, "Save")
If msg = vbYes Then
user.Update
DataGrid1.Visible = False
Command3.Visible = False
Command4.Visible = False
user.Close
ElseIf msg = vbNo Then
aa = MsgBox("Want to edit details", vbYesNo + vbDefaultButton2 + vbInformation, "Edit More")
If aa = vbYes Then
DataGrid1.Visible = True
ElseIf aa = vbNo Then
DataGrid1.Visible = False
Command3.Visible = False
Command4.Visible = False
user.Close
End If
End If
End Sub

Private Sub Command4_Click()
Dim msg As String
msg = MsgBox("Sure to Cancel", vbYesNo + vbDefaultButton2 + vbInformation, "Cancel")
If msg = vbYes Then
DataGrid1.Visible = False
Command3.Visible = False
Command4.Visible = False
user.Close
ElseIf msg = vbNo Then
DataGrid1.Visible = True
Command3.Visible = True
Command4.Visible = True
End If
End Sub

Private Sub Command5_Click()
Dim msg As String
Dim aa As String
msg = MsgBox("Save User Details", vbYesNo + vbDefaultButton2 + vbInformation, "Save")
If msg = vbYes Then
useradd.Update
DataGrid2.Visible = False
Command5.Visible = False
Command6.Visible = False
useradd.Close
ElseIf msg = vbNo Then
aa = MsgBox("Want to edit details", vbYesNo + vbDefaultButton2 + vbInformation, "Edit More")
If aa = vbYes Then
DataGrid2.Visible = True
ElseIf aa = vbNo Then
DataGrid2.Visible = False
Command5.Visible = False
Command6.Visible = False
useradd.Close
End If
End If
End Sub

Private Sub Command6_Click()
Dim msg As String
msg = MsgBox("Sure to Cancel", vbYesNo + vbDefaultButton2 + vbInformation, "Cancel")
If msg = vbYes Then
DataGrid2.Visible = False
Command5.Visible = False
Command6.Visible = False
useradd.Close
ElseIf msg = vbNo Then
DataGrid2.Visible = True
Command5.Visible = True
Command6.Visible = True
End If
End Sub

Private Sub Command7_Click()
Dim msg As String
Dim aa As String
msg = MsgBox("Save User Details", vbYesNo + vbDefaultButton2 + vbInformation, "Save")
If msg = vbYes Then
family.Update
DataGrid3.Visible = False
Command7.Visible = False
Command8.Visible = False
family.Close
ElseIf msg = vbNo Then
aa = MsgBox("Want to edit details", vbYesNo + vbDefaultButton2 + vbInformation, "Edit More")
If aa = vbYes Then
DataGrid3.Visible = True
ElseIf aa = vbNo Then
DataGrid3.Visible = False
Command7.Visible = False
Command8.Visible = False
family.Close
End If
End If
End Sub

Private Sub Command8_Click()
Dim msg As String
msg = MsgBox("Sure to Cancel", vbYesNo + vbDefaultButton2 + vbInformation, "Cancel")
If msg = vbYes Then
DataGrid3.Visible = False
Command7.Visible = False
Command8.Visible = False
family.Close
ElseIf msg = vbNo Then
DataGrid3.Visible = True
Command7.Visible = True
Command8.Visible = True
End If
End Sub

Private Sub Command9_Click()
view.Close
DataGrid4.Visible = False
Command9.Visible = False
End Sub

Private Sub Form_Load()
WindowState = 2
Label26.Caption = "About The Scheme"
Label21.Caption = "About The Scheme"
Label20.Caption = "The scheme launched in Karnataka State on 10 July 2013 by Honourable Chief Minister Sri.Siddaramaiah. In this scheme, poor people will be given free rice, so that they can easily get two meals in a day"
Label23.Caption = "Investment on this scheme is approximately Rs.4200 Crores."
Label25.Caption = "As of now Annabhagya Yojana Scheme has Succesfully Completed 6 years and 2 months."
Label19.Caption = "This scheme for poor peoples or you can say Below Poverty Line (BPL) peoples .In this scheme, people with single card holder person in the family will get 10 kg rice at a rate of Rs 1/kg.Family with 2 cardholders will get 20 Kg rice and with 3 or more will get a maximum of 30 Kg rice at same price rate .Not only rice peoples will also get edible oils, sugar, iodized salt, kerosene and other items in fewer price rates.So this will help poor peoples to get at least 2 times meal in a day"
con.Open "Provider=MSDASQL.1;Password=root;Persist Security Info=True;User ID=root;Data Source=yogesh"
rs.CursorLocation = adUseClient
rs.Open "select firstname,midname,lastname,aadhar,permanentadd,perdistrict,perpincode,tempadd,tempdistrict,temppincode,dob,gender,caste,religion,email,phno,cardtype,name1,name2,name3,dob1,dob2,dob3,relation1,relation2,relation3,aadhar1,aadhar2,aadhar3 from userinfo", con, adOpenDynamic, adLockPessimistic
que.Open "select aadhar,name,todaydate,tomdate,cardtype from queue", con, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rs
DataGrid1.Visible = True
Set Textfirstname.DataSource = rs
Set textmidname.DataSource = rs
Set textlastname.DataSource = rs
Set Textaadhar.DataSource = rs
Set texttmpadd.DataSource = rs
Set textperadd.DataSource = rs
Set textperdist.DataSource = rs
Set textperpin.DataSource = rs
Set textdistrict.DataSource = rs
Set textpicode.DataSource = rs
Set textreligion.DataSource = rs
Set textcaste.DataSource = rs
Set textemail.DataSource = rs
Set textphno.DataSource = rs
Set Textn1.DataSource = rs
Set Textn2.DataSource = rs
Set Textn3.DataSource = rs
Set Textr1.DataSource = rs
Set Textr2.DataSource = rs
Set Textr3.DataSource = rs
Set Texta1.DataSource = rs
Set Texta2.DataSource = rs
Set Texta3.DataSource = rs
Textn1.DataField = "name1"
Textn2.DataField = "name2"
Textn3.DataField = "name3"
Textr1.DataField = "relation1"
Textr2.DataField = "relation2"
Textr3.DataField = "relation3"
Texta1.DataField = "aadhar1"
Texta2.DataField = "aadhar2"
Texta3.DataField = "aadhar3"
Textfirstname.DataField = "firstname"
textmidname.DataField = "midname"
textlastname.DataField = "lastname"
Textaadhar.DataField = "aadhar"
textperadd.DataField = "permanentadd"
textperdist.DataField = "perdistrict"
textperpin.DataField = "perpincode"
texttmpadd.DataField = "tempadd"
textdistrict.DataField = "tempdistrict"
textpicode.DataField = "temppincode"
textreligion.DataField = "religion"
textcaste.DataField = "caste"
textphno.DataField = "phno"
textemail.DataField = "email"
Call clear
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu mnupopup, vbPopupMenuRightButton
End If
End Sub

Private Sub mnuadd_Click()
Dim msg As String
msg = InputBox("Enter User Aadharno", "User verification")
useradd.CursorLocation = adUseClient
useradd.Open "select aadhar,tempadd,tempdistrict,temppincode,permanentadd,perdistrict,perpincode from userinfo where aadhar like" & "'" & msg & "'", con, adOpenDynamic, adLockPessimistic
If useradd.EOF = True Or useradd.BOF = True Then
MsgBox "User Not found", vbOKOnly + vbCritical, "User Error"
Else
useradd.Requery
Set DataGrid2.DataSource = useradd
DataGrid2.Visible = True
DataGrid2.Width = 20490
DataGrid2.Height = 11500
DataGrid2.Top = 120
DataGrid2.Left = 0
Command5.Visible = True
Command5.Left = 8500
Command5.Top = 9000
Command6.Visible = True
Command6.Left = 9715
Command6.Top = 9000
DataGrid2.Columns(0).Caption = "AADHAR NO"
DataGrid2.Columns(1).Caption = "TEMPORARY ADDRESS"
DataGrid2.Columns(2).Caption = "TEMPORARY DISTRICT"
DataGrid2.Columns(3).Caption = "TEMPORARY PINCODE"
DataGrid2.Columns(4).Caption = "PERMANENT ADDRESS"
DataGrid2.Columns(5).Caption = "PERMANENT DISTRICT"
DataGrid2.Columns(6).Caption = "PERMANENT PINCODE"
End If
End Sub

Private Sub mnuaddqueue_Click()
Shape2.Visible = True
queuehead.Visible = True
queuehead.Left = 6480
queuehead.Top = 1200
queuehead.Width = 7695
Call verifytrue
Call uservisiblefalse
Call famvisiblefalse
Call namevifalse
Call namecombofalse
Call addfalse
queaadharveri.SetFocus
End Sub

Private Sub mnuBasic_Click()
Dim msg As String
msg = InputBox("Enter User Aadharno", "User verification")
user.CursorLocation = adUseClient
user.Open "select aadhar,firstname,midname,lastname,dob,gender,religion,caste,email,phno from userinfo where aadhar like" & "'" & msg & "'", con, adOpenDynamic, adLockPessimistic
If user.EOF = True Or user.BOF = True Then
MsgBox "User Not found", vbOKOnly + vbCritical, "User Error"
Else
user.Requery
Set DataGrid1.DataSource = user
DataGrid1.Visible = True
DataGrid1.Width = 20490
DataGrid1.Height = 11500
DataGrid1.Top = 120
DataGrid1.Left = 0
Command3.Visible = True
Command3.Left = 8500
Command3.Top = 9000
Command4.Visible = True
Command4.Left = 9715
Command4.Top = 9000
DataGrid1.Columns(0).Caption = "AADHAR NO"
DataGrid1.Columns(1).Caption = "FIRST NAME"
DataGrid1.Columns(2).Caption = "MIDDLE NAME"
DataGrid1.Columns(3).Caption = "LAST NAME"
DataGrid1.Columns(4).Caption = "DOB"
DataGrid1.Columns(5).Caption = "GENDER"
DataGrid1.Columns(6).Caption = "RELIGION"
DataGrid1.Columns(7).Caption = "CASTE"
DataGrid1.Columns(8).Caption = "EMAIL"
DataGrid1.Columns(9).Caption = "PHONE NUMBER"
'useradd.Close
'family.Close
DataGrid2.Visible = False
DataGrid3.Visible = False
End If
End Sub

Private Sub mnucard_Click()
Dim pri As String
pri = InputBox("Enter User AadharNo", "Aadhar Print")
DataEnvironment1.rscardprint.Open "select *from userinfo where aadhar like " & "'" & pri & "'", con, adOpenDynamic, adLockPessimistic
CARDPRINT.Refresh
CARDPRINT.Show
DataEnvironment1.rscardprint.Close
End Sub

Private Sub mnuclearqueue_Click()
Dim cc As New ADODB.Recordset
Dim a As String
a = MsgBox("Confirm Clear", vbYesNo + vbDefaultButton2 + vbInformation, "CLEAR")
If a = vbYes Then
cc.Open "delete from queue", con, adOpenDynamic, adLockPessimistic
ElseIf a = vbNo Then
Exit Sub
End If
End Sub

Private Sub mnudelete_Click()
Dim del As String
del = InputBox("Enter Card Id", "Card ID")
Dim d As New ADODB.Recordset
Dim ss As String
ss = MsgBox("Confirm to Delete", vbYesNo + vbDefaultButton2 + vbInformation, "Delete")
If ss = vbYes Then
d.Open "delete from userinfo where id like " & "'" & del & "'", con, adOpenDynamic, adLockPessimistic
Else
d.Close
Exit Sub
End If
End Sub

Private Sub mnuexecute_Click()
Dim a As String
a = InputBox("Enter Token ID", "Queue Execution")
quex.Open "select * from queue where id like " & "'" & a & "'", con, adOpenDynamic, adLockPessimistic
If quex.BOF = True Or quex.EOF = True Then
MsgBox "Queue Empty", vbOKOnly + vbInformation, "EMPTY"
Else
MsgBox "TOKEN EXECUTED", vbOKOnly + vbInformation, "TOKEN RUN"
End If
End Sub

Private Sub mnuexit_Click()
Call formexit
End Sub

Private Sub mnufaminfo_Click()
Dim msg As String
msg = InputBox("Enter User Aadharno", "User verification")
family.CursorLocation = adUseClient
family.Open "select aadhar,name1,aadhar1,relation1,dob1,name2,aadhar2,relation2,dob2,name3,aadhar3,relation3,dob3 from userinfo where aadhar like" & "'" & msg & "'", con, adOpenDynamic, adLockPessimistic
If family.EOF = True Or family.BOF = True Then
MsgBox "User Not found", vbOKOnly + vbCritical, "User Error"
Else
Set DataGrid3.DataSource = family
DataGrid3.Visible = True
DataGrid3.Width = 20490
DataGrid3.Height = 11500
DataGrid3.Top = 120
DataGrid3.Left = 0
DataGrid3.Columns(0).Caption = "AADHAR NO"
DataGrid3.Columns(1).Caption = "NAME"
DataGrid3.Columns(2).Caption = "AADHAR NO"
DataGrid3.Columns(3).Caption = "RELATION"
DataGrid3.Columns(4).Caption = "DOB"
DataGrid3.Columns(5).Caption = "NAME"
DataGrid3.Columns(6).Caption = "AADHAR NO"
DataGrid3.Columns(7).Caption = "RELATION"
DataGrid3.Columns(8).Caption = "DOB"
DataGrid3.Columns(9).Caption = "NAME"
DataGrid3.Columns(10).Caption = "AADHAR NO"
DataGrid3.Columns(11).Caption = "RELATION"
DataGrid3.Columns(12).Caption = "DOB"
Command7.Visible = True
Command7.Left = 8500
Command7.Top = 9000
Command8.Visible = True
Command8.Top = 9000
Command8.Left = 9715
End If
End Sub

Private Sub mnunew_Click()
Shape2.Visible = True
Call addtrue
Call verifyfalse
Call namevifalse
Call famvisiblefalse
Call uservisiblefalse
queuehead.Visible = False
End Sub
Private Sub mnusave_Click()
Call save
End Sub

Private Sub mnuveritoken_Click()
Dim a As String
a = InputBox("Enter Tokenid", "Token verification")
veri.CursorLocation = adUseClient
veri.Open "select id,aadhar,name,todaydate,tomdate,cardtype from queue where id like " & "'" & a & "'", con, adOpenDynamic, adLockPessimistic
If veri.BOF = True Or veri.EOF = True Then
MsgBox "Queue Is Empty", vbOKOnly + vbInformation, "Empty"
Else
Set DataGrid5.DataSource = veri
DataGrid5.Visible = True
DataGrid5.Width = 20490
DataGrid5.Height = 11500
DataGrid5.Top = 120
DataGrid5.Left = 0
Command10.Visible = True
Command10.Left = 8500
Command10.Top = 9000
DataGrid5.Columns(0).Caption = "ID"
DataGrid5.Columns(1).Caption = "AADHAR"
DataGrid5.Columns(2).Caption = "NAME"
DataGrid5.Columns(3).Caption = "TOKEN PRINT DATE"
DataGrid5.Columns(4).Caption = "TOKEN EXPIRY DATE"
DataGrid5.Columns(5).Caption = "CARDTYPE"
End If
End Sub

Private Sub mnuviewQueue_Click()
view.CursorLocation = adUseClient
view.Open "select id,aadhar,name,todaydate,tomdate,cardtype from queue", con, adOpenDynamic, adLockPessimistic
Set DataGrid4.DataSource = view
DataGrid4.Visible = True
DataGrid4.Width = 20490
DataGrid4.Height = 11500
DataGrid4.Top = 120
DataGrid4.Left = 0
Command9.Visible = True
Command9.Left = 8500
Command9.Top = 9000
DataGrid4.Columns(0).Caption = "ID"
DataGrid4.Columns(1).Caption = "AADHAR"
DataGrid4.Columns(2).Caption = "NAME"
DataGrid4.Columns(3).Caption = "TOKEN PRINT DATE"
DataGrid4.Columns(4).Caption = "TOKEN EXPIRY DATE"
DataGrid4.Columns(5).Caption = "CARDTYPE"
End Sub

Private Sub Optionapl_Click()
If Optionapl.Value = True Then
Label27.Visible = False
Picture2.Visible = False
Picture1.Visible = True
Call uservisibletrue
rs.AddNew
rs("cardtype").Value = "APL"
End If
End Sub

Private Sub OptionBpl_Click()
If OptionBpl.Value = True Then
Label27.Visible = False
Picture2.Visible = False
Call uservisibletrue
rs.AddNew
rs("cardtype").Value = "BPL"
End If
End Sub

Private Sub Optionfemale_Click()
If Optionfemale.Value = True Then
rs("gender").Value = "Female"
End If
End Sub

Private Sub Optionfno_Click()
Dim str As String
If Optionfno.Value = True Then
str = MsgBox("Do you want to Save User details", vbYesNo + vbDefaultButton2 + vbInformation, "Confirmation")
If str = vbYes Then
rs.Update
ElseIf str = vbNo Then
rs.Delete
Call uservisiblefalse
Call famvisiblefalse
Labelappli.Visible = False
Shape2.Visible = False
End If
End If
End Sub

Private Sub Optionfyes_Click()
Dim str As String
If Optionfyes.Value = True Then
str = MsgBox("Sure to add Family Details", vbYesNo + vbDefaultButton2 + vbInformation, "Continue...")
If str = vbYes Then
Call famvisibletrue
Call uservisiblefalse
relname1.Left = 6480
relname2.Left = 6480
relname3.Left = 6480
reldob1.Left = 6480
reldob2.Left = 6480
reldob3.Left = 6480
relaadhar1.Left = 6480
relaadhar2.Left = 6480
relaadhar3.Left = 6480
rel1.Left = 6480
rel2.Left = 6480
rel3.Left = 6480
Textr1.Left = 8280
Textr2.Left = 8280
Textr3.Left = 8280
Texta1.Left = 8280
Texta2.Left = 8280
Texta3.Left = 8280
DTPicker2.Left = 8280
DTPicker3.Left = 8280
DTPicker4.Left = 8280
Textn1.Left = 8280
Textn2.Left = 8280
Textn3.Left = 8280
Textn1.Width = 5895
Textn2.Width = 5895
Textn3.Width = 5895
Textr1.Width = 5895
Textr2.Width = 5895
Textr3.Width = 5895
Texta1.Width = 5895
Texta2.Width = 5895
Texta3.Width = 5895
DTPicker2.Width = 5895
DTPicker3.Width = 5895
DTPicker4.Width = 5895
Command1.Left = 8280
Command1.Top = 8880
Command1.Width = 2900
Command2.Width = 2900
Command2.Top = 8880
Command2.Left = 11280
ElseIf str = vbNo Then
Dim clo As String
clo = MsgBox("Do you Want to Save", vbYesNo + vbDefaultButton2 + vbInformation, "Save")
If clo = vbYes Then
rs.Update
Call uservisiblefalse
Call famvisiblefalse
ElseIf clo = vbNo Then
rs.Delete
Call uservisiblefalse
Call famvisiblefalse
Labelappli.Visible = False
Shape2.Visible = True
End If
End If
End If
End Sub

Private Sub Optionmale_Click()
If Optionmale.Value = True Then
rs("gender").Value = "Male"
End If
End Sub

Private Sub Optionno_Click()
If Optionno.Value = True Then
textperadd.TabIndex = 7
textperdist.TabIndex = 8
textperpin.TabIndex = 9
DTPicker1.TabIndex = 10
textreligion.TabIndex = 11
textcaste.TabIndex = 12
textphno.TabIndex = 13
textemail.TabIndex = 14
End If
End Sub

Private Sub Optionothers_Click()
If Optionothers.Value = True Then
rs("gender").Value = "Transgender"
End If
End Sub

Private Sub Optionyes_Click()
If Optionyes.Value = True Then
textperadd.Locked = True
textperdist.Locked = True
textperdist.Locked = True
textperadd.Text = texttmpadd.Text
textperdist.Text = textdistrict.Text
textperpin.Text = textpicode.Text
DTPicker1.TabIndex = 7
textreligion.TabIndex = 8
textcaste.TabIndex = 9
textphno.TabIndex = 11
textemail.TabIndex = 10
End If
End Sub

Public Sub cardtype()
If Optionapl.Value = True Then
rs.Fields("cardtype").Value = "APL"
ElseIf OptionBpl.Value = True Then
rs.Fields("cardtype").Value = "BPL"
End If
End Sub

Public Sub gender()
If Optionmale.Value = True Then
rs("gender").Value = "Male"
ElseIf Optionfemale.Value = True Then
rs("gender").Value = "Female"
ElseIf Optionothers.Value = True Then
rs("gender").Value = "Transgender"
End If
End Sub

Public Sub dob()
rs("dob").Value = DTPicker1.Value
rs("dob1").Value = DTPicker2.Value
rs("dob2").Value = DTPicker3.Value
rs("dob3").Value = DTPicker4.Value
End Sub

Public Sub uservisibletrue()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label18.Visible = True
Label28.Visible = True
Textfirstname.Visible = True
textmidname.Visible = True
textlastname.Visible = True
Textaadhar.Visible = True
textperadd.Visible = True
textperpin.Visible = True
textperdist.Visible = True
texttmpadd.Visible = True
textdistrict.Visible = True
textpicode.Visible = True
textemail.Visible = True
textphno.Visible = True
textreligion.Visible = True
textcaste.Visible = True
Picture1.Visible = True
Optionmale.Visible = True
Optionfemale.Visible = True
Optionothers.Visible = True
Optionyes.Visible = True
Optionno.Visible = True
DTPicker1.Visible = True
End Sub

Public Sub uservisiblefalse()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label18.Visible = False
Label28.Visible = False
Textfirstname.Visible = False
textmidname.Visible = False
textlastname.Visible = False
Textaadhar.Visible = False
textperadd.Visible = False
textperpin.Visible = False
textperdist.Visible = False
texttmpadd.Visible = False
textdistrict.Visible = False
textpicode.Visible = False
textemail.Visible = False
textphno.Visible = False
textreligion.Visible = False
textcaste.Visible = False
Picture1.Visible = False
Optionmale.Visible = False
Optionfemale.Visible = False
Optionothers.Visible = False
Optionyes.Visible = False
Optionno.Visible = False
DTPicker1.Visible = False
End Sub

Public Sub famvisibletrue()
rel1.Visible = True
rel2.Visible = True
rel3.Visible = True
relname1.Visible = True
relname2.Visible = True
relname3.Visible = True
reldob1.Visible = True
reldob2.Visible = True
reldob3.Visible = True
relaadhar1.Visible = True
relaadhar3.Visible = True
relaadhar2.Visible = True
Textn1.Visible = True
Textn2.Visible = True
Textn3.Visible = True
Texta1.Visible = True
Texta2.Visible = True
Texta3.Visible = True
Textr1.Visible = True
Textr2.Visible = True
Textr3.Visible = True
DTPicker2.Visible = True
DTPicker3.Visible = True
DTPicker4.Visible = True
Command1.Visible = True
Command2.Visible = True
End Sub

Public Sub famvisiblefalse()
rel1.Visible = False
rel2.Visible = False
rel3.Visible = False
relname1.Visible = False
relname2.Visible = False
relname3.Visible = False
reldob1.Visible = False
reldob2.Visible = False
reldob3.Visible = False
relaadhar1.Visible = False
relaadhar3.Visible = False
relaadhar2.Visible = False
Textn1.Visible = False
Textn2.Visible = False
Textn3.Visible = False
Texta1.Visible = False
Texta2.Visible = False
Texta3.Visible = False
Textr1.Visible = False
Textr2.Visible = False
Textr3.Visible = False
DTPicker2.Visible = False
DTPicker3.Visible = False
DTPicker4.Visible = False
Command1.Visible = False
Command2.Visible = False
End Sub

Public Sub clear()
Textaadhar = ""
Textfirstname = ""
textlastname = ""
textmidname = ""
texttmpadd = ""
textdistrict = ""
textpicode = ""
textperadd = ""
textperdist = ""
textperpin = ""
textcaste = ""
textreligion = ""
textemail = ""
textphno = ""
DTPicker1 = "01-01-2000"
Textfamadhar = ""
Textfamname = ""
Textfamrelat = ""
Textn1 = ""
Textn2 = ""
Textn3 = ""
Texta1 = ""
Texta2 = ""
Texta3 = ""
Textr1 = ""
Textr3 = ""
textr4 = ""
DTPicker2 = "01-01-2000"
DTPicker3 = "01-01-2000"
DTPicker4 = "01-01-2000"
End Sub

Public Sub save()
Dim str As String
str = MsgBox("Save Details", vbYesNo + vbDefaultButton2 + vbInformation, "Save")
If str = vbYes Then
Call dob
Call gender
Call cardtype
rs.Update
Call famvisiblefalse
Call uservisiblefalse
Call verifyfalse
Call namevifalse
Shape2.Visible = False
Labelappli.Visible = False
OptionBpl.Value = False
Optionapl.Value = False
Optionmale.Value = False
Optionfemale.Value = False
Optionothers.Value = False
Optionfno.Value = False
Optionfyes.Value = False
Optionyes.Value = False
Optionno.Value = False
ElseIf str = vbNo Then
Exit Sub
Call famvisiblefalse
Call uservisiblefalse
Call verifyfalse
Call namevifalse
Shape2.Visible = False
Labelappli.Visible = False
OptionBpl.Value = False
Optionapl.Value = False
Optionmale.Value = False
Optionfemale.Value = False
Optionothers.Value = False
Optionfno.Value = False
Optionfyes.Value = False
Optionyes.Value = False
Optionno.Value = False
End If
End Sub

Public Sub addtrue()
Labelappli.Visible = True
Label27.Visible = True
Label27.Left = 6480
Picture2.Left = 8280
Picture2.Visible = True
Label27.Top = 1800
Picture2.Top = 1800
Optionapl.Value = False
OptionBpl.Value = False
End Sub

Public Sub addfalse()
Labelappli.Visible = False
Label27.Visible = False
Label27.Left = 6480
Picture2.Visible = False
Label27.Top = 1800
Picture2.Top = 1800
End Sub

Public Sub datetime()
que("name").Value = quename.Caption
Dim tommorow As String
tommorow = DateAdd("n", 1200, Now)
que("tomdate").Value = tommorow
Dim today As String
today = Format(Now)
que("todaydate").Value = today
End Sub

Public Sub namevifalse()
quename.Visible = False
quecopyaadhar.Visible = False
queuedone.Visible = False
End Sub

Public Sub namevitrue()
quename.Visible = True
quecopyaadhar.Visible = True
queuedone.Visible = True
quename.Left = 8280
quename.Top = 2400
quename.Width = 5895
quecopyaadhar.Left = 8280
quecopyaadhar.Top = 1800
quecopyaadhar.Width = 5895
queuedone.Left = 8280
queuedone.Top = 3000
queuedone.Width = 5895
End Sub

Public Sub verifyfalse()
queaadhar.Visible = False
quecommandverify.Visible = False
queaadharveri.Visible = False
End Sub

Public Sub verifytrue()
queaadhar.Visible = True
quecommandverify.Visible = True
queaadharveri.Visible = True
queaadhar.Left = 6480
queaadhar.Top = 1800
queaadhar.Width = 5895
queaadharveri.Left = 8280
queaadharveri.Top = 1800
quecommandverify.Left = 8280
quecommandverify.Top = 2400
quecommandverify.Width = 5895
queaadharveri.Width = 5895
End Sub

Private Sub popaddqueue_Click()
Shape2.Visible = True
queuehead.Visible = True
queuehead.Left = 6480
queuehead.Top = 1200
queuehead.Width = 7695
Call verifytrue
Call uservisiblefalse
Call famvisiblefalse
Call namevifalse
Call namecombofalse
Call addfalse
queaadharveri.SetFocus
End Sub

Private Sub popclear_Click()
Call clear
End Sub

Private Sub popexit_Click()
Call formexit
End Sub

Private Sub popnew_Click()
Shape2.Visible = True
Call addtrue
Call verifyfalse
Call namevifalse
Call famvisiblefalse
Call uservisiblefalse
queuehead.Visible = False
End Sub

Private Sub poprunqueue_Click()
Dim a As String
a = InputBox("Enter Tokenid", "Token Execution")
quex.Open "delete from queue where id like " & " ' " & a & " '", con, adOpenDynamic, adLockPessimistic
If quex.BOF = True Or ques.EOF = True Then
MsgBox "Queue Empty", vbOKOnly + vbInformation, "Empty"
Else
MsgBox "Token Executed Successfully", vbOKOnly + vbInformation, "Success"
End If
End Sub

Private Sub popsave_Click()
Call save
End Sub

Private Sub popveritoken_Click()
Dim a As String
a = InputBox("Enter Tokenid", "Token verification")
veri.CursorLocation = adUseClient
veri.Open "select *from queue where id like " & "'" & a & "'", con, adOpenDynamic, adLockPessimistic
If veri.BOF = True Or veri.EOF = True Then
MsgBox "Queue Is Empty", vbOKOnly + vbInformation, "Empty"
Else
Set DataGrid5.DataSource = veri
DataGrid5.Visible = True
DataGrid5.Width = 20490
DataGrid5.Height = 11500
DataGrid5.Top = 120
DataGrid5.Left = 0
Command10.Visible = True
Command10.Left = 8500
Command10.Top = 9000
End If
End Sub

Private Sub queaadharveri_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub quecombo1_Click()
que.AddNew
quename.Caption = quecombo1.Text
que("aadhar").Value = quecopyaadhar.Caption
que("cardtype").Value = queuecardtype.Caption
que("name").Value = quename.Caption
Call namevitrue
queselect.Visible = False
quecombo1.Visible = False
End Sub

Private Sub quecommandverify_Click()
On Error GoTo e:
Dim qus As New ADODB.Recordset
qus.Open "select firstname,name1,name2,name3,cardtype from userinfo where aadhar like" & "'" & queaadharveri.Text & "'", con, adOpenStatic, adLockOptimistic
If qus.EOF = True Or qus.BOF = True Then
MsgBox "User ID Not Found", vbOKCancel + vbDefaultButton1 + vbExclamation, "User Not Exists"
queaadharveri = ""
queaadharveri.SetFocus
Else
Call namecombotrue
Call verifyfalse
Call namevifalse
quecombo1.AddItem (qus("firstname").Value)
quecombo1.AddItem (qus("name1").Value)
quecombo1.AddItem (qus("name2").Value)
quecombo1.AddItem (qus("name3").Value)
quecopyaadhar.Caption = queaadharveri.Text
queuecardtype.Caption = qus("cardtype").Value
End If
qus.Close
e:
err.clear
End Sub

Private Sub queuedone_Click()
On Error GoTo err
Dim a As String
a = MsgBox("Do you want to Print", vbYesNo + vbDefaultButton2 + vbInformation, "Print")
If a = vbYes Then
Call datetime
que.Update
quecombo1.clear
queaadharveri = ""
DataEnvironment1.rstokenprint.Open "select *from queue where aadhar like " & "'" & quecopyaadhar.Caption & "'", con, adOpenDynamic, adLockPessimistic
tokenprint.Refresh
tokenprint.Show
DataEnvironment1.rstokenprint.Close
Call verifytrue
Call namevifalse
queaadharveri.SetFocus
ElseIf a = vbNo Then
que.Delete
Call verifytrue
Call namevifalse
quecombo1.clear
queaadharveri = ""
queaadharveri.SetFocus
End If
err:
MsgBox (err.Description)
Shape2.Visible = True
queuehead.Visible = True
queuehead.Left = 6480
queuehead.Top = 1200
queuehead.Width = 7695
Call verifytrue
Call uservisiblefalse
Call famvisiblefalse
Call namevifalse
Call namecombofalse
Call addfalse
queaadharveri.SetFocus
queaadharveri.Text = ""
End Sub

Public Sub namecombotrue()
queselect.Visible = True
queselect.Left = 6480
queselect.Top = 1800
quecombo1.Left = 8280
quecombo1.Top = 1800
quecombo1.Width = 5895
quecombo1.Visible = True
End Sub

Public Sub namecombofalse()
queselect.Visible = False
quecombo1.Visible = False
End Sub

Public Sub runqueueaadhartrue()
labeltokenid.Visible = True
labeltokenid.Left = 6480
labeltokenid.Top = 1800
texttokenid.Visible = True
texttokenid.Left = 8280
texttokenid.Top = 1800
texttokenid.Width = 5895
cmdverifytoken.Left = 8280
cmdverifytoken.Top = 2400
cmdverifytoken.Visible = True
End Sub

Public Sub runqueueaadharfalse()
labeltokenid.Visible = False
labeltokenid.Left = 6480
labeltokenid.Top = 1800
texttokenid.Visible = False
texttokenid.Left = 8280
texttokenid.Top = 1800
texttokenid.Width = 5895
cmdverifytoken.Left = 8280
cmdverifytoken.Top = 2400
cmdverifytoken.Visible = False
End Sub

Private Sub Texta1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub Texta2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub Texta3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub Textaadhar_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub textcaste_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub


Private Sub textdistrict_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub textemail_LostFocus()
Call isemail(textemail.Text)
End Sub

Private Sub Textfirstname_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textfirstname_LostFocus()
If Textfirstname.Text = "" Then
Textfirstname.SetFocus
End If
End Sub

Private Sub textlastname_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub textmidname_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textn1_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textn2_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textn3_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub


Private Sub textperpin_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub


Private Sub textphno_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub textphno_LostFocus()
If Len(textphno.Text) < 10 Or Len(textphno.Text) > 10 Then
MsgBox "INVALID PHNO", vbOK + vbInformation, "Wrong Number"
textphno = ""
textphno.SetFocus
End If
End Sub

Private Sub textpicode_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

Private Sub Textr1_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textr2_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub Textr3_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Private Sub textreligion_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
End If
End Sub

Public Sub formexit()
Dim ex As String
ex = MsgBox("Are you sure to Exit", vbYesNo + vbDefaultButton2 + vbInformation, "Exit")
If ex = vbYes Then
Unload Me
ElseIf ex = vbNo Then
Exit Sub
End If
End Sub

Public Function isemail(email As String) As Boolean
Dim at As Integer
Dim dt As Integer
Dim dts As Integer
isemail = True
at = InStr(1, email, "@", vbTextCompare)
dt = InStr(at + 2, email, ".", vbTextCompare)
dts = InStr(at + 2, email, "..", vbTextCompare)
If at = 0 Or dt = 0 Or Not dts = 0 Or Right(email, 1) = "." Then
isemail = False
MsgBox "Email ID you Entered Is Not Valid", vbOKOnly + vbCritical, "Invalid Email"
textemail = ""
textemail.SetFocus
Else
isemail = True
End If
End Function

