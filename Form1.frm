VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D4D4D4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unique Serial Security System (v 1.5 - Expanded)"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command24 
      BackColor       =   &H00F2F2FF&
      Caption         =   "Clear Fields"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00F2F2FF&
      Caption         =   "Digit Swap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00F2F2FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00F2F2FF&
      Caption         =   "Sample Label"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0080C0FF&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Finish"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   4920
      Width           =   4335
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   29
      TabIndex        =   28
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text22 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text21 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   26
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   24
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   23
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Position Swap"
      Height          =   255
      Left            =   255
      TabIndex        =   22
      Top             =   3960
      Width           =   4200
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   21
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   20
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   19
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   17
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Alpha Replacement Sequence"
      Height          =   255
      Left            =   255
      TabIndex        =   16
      Top             =   3120
      Width           =   4200
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selective Inversion Process"
      Height          =   255
      Left            =   255
      TabIndex        =   15
      Top             =   2160
      Width           =   4200
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Split and Morph Code"
      Height          =   255
      Left            =   255
      TabIndex        =   9
      Top             =   1440
      Width           =   4200
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   25
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Unique Request Code (SN + PID)"
      Height          =   255
      Left            =   255
      TabIndex        =   2
      Top             =   480
      Width           =   4320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HDD SN"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H008080FF&
      Caption         =   "HELP!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tools & Features"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   49
      Top             =   5880
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   5640
      Width           =   5895
   End
   Begin VB.Line Line10 
      X1              =   4560
      X2              =   4560
      Y1              =   1320
      Y2              =   4680
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   120
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4725
      TabIndex        =   31
      Top             =   195
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Independant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4725
      TabIndex        =   30
      Top             =   0
      Width           =   1170
   End
   Begin VB.Line Line9 
      X1              =   4710
      X2              =   4710
      Y1              =   0
      Y2              =   5760
   End
   Begin VB.Line Line8 
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   5760
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0093B830&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   240
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   4560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0093B830&
      X1              =   3960
      X2              =   3960
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0093B830&
      X1              =   2280
      X2              =   2280
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0093B830&
      X1              =   600
      X2              =   600
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00806622&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Top             =   720
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   645
      Left            =   120
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00806622&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D86B50&
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   4680
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "..."
      Begin VB.Menu mnuLabel 
         Caption         =   "View Sample Label"
      End
      Begin VB.Menu mnuDS 
         Caption         =   "Perform Digit Swap"
      End
      Begin VB.Menu mnuReg 
         Caption         =   "Sample Registration Screen"
      End
      Begin VB.Menu s435345624523452345 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuLhelp 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################################
'#                                                #
'#            Unique Serial Generation            #
'#            ========================            #
'#                                                #
'# Copyright: Rehan Dalal                         #
'# For educational purposes only. If you use it,  #
'# please give credit.                            #
'#                                                #
'##################################################

Option Explicit

Private Declare Function GetVolumeInformation& Lib "kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName _
    As String, ByVal pVolumeNameBuffer As String, ByVal _
    nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As _
    Long, ByVal lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long)
    Const MAX_FILENAME_LEN = 256
    
Public Function SerNum(Drive$) As Long 'Find the hard disk serial number
    Dim No&, s As String * MAX_FILENAME_LEN
    Call GetVolumeInformation(Drive + ":\", s, MAX_FILENAME_LEN, _
    No, 0&, 0&, s, MAX_FILENAME_LEN)
    SerNum = No
End Function

Private Function Invert(strng As String) ' Inversion function
Dim i As Integer
Dim tmp_txt As String

i = Len(strng)

Do Until i = 0

    tmp_txt = tmp_txt & Mid(strng, i, 1)

    i = i - 1
Loop

Invert = tmp_txt

End Function

Private Sub Command1_Click()
Dim MFact As Integer
Command3_Click

MFact = Int(Val(Val(Val(Mid(Text2.Text, 1, 1)) + Val(Mid(Text2.Text, 12, 1)) + Val(Mid(Text2.Text, 24, 1)) + Val(Mid(Text2.Text, Val(Mid(Text2.Text, 1, 1)), 1))) / 4))
Text3.Text = iSplit(Text2.Text, MFact, 0)
Text4.Text = iSplit(Text2.Text, MFact, 1)
Text5.Text = iSplit(Text2.Text, MFact, 2)
Text6.Text = iSplit(Text2.Text, MFact, 3)
Text7.Text = iSplit(Text2.Text, MFact, 4)
End Sub

Private Sub Command10_Click()
If Mid(Text3.Text, 5, 1) <> 0 Then
    Text8.Text = Invert(Text3.Text)
Else
    Text8.Text = Text3.Text
End If
Text10.Text = Invert(Text5.Text)
Text12.Text = Invert(Text7.Text)
Text9.Text = Text4.Text
Text11.Text = Text6.Text
End Sub

Private Sub Command11_Click()
Text13.Text = Replace(Text8.Text, "27", "Z3")
Text13.Text = Replace(Text13.Text, "91", "8F")
Text13.Text = Replace(Text13.Text, "72", "1K")
Text13.Text = Replace(Text13.Text, "19", "PS")
Text13.Text = Replace(Text13.Text, "56", "O1")
Text13.Text = Replace(Text13.Text, "65", "M3")
Text13.Text = Replace(Text13.Text, "83", "L0")
Text13.Text = Replace(Text13.Text, "38", "E5")
Text13.Text = Replace(Text13.Text, "01", "XD")
Text13.Text = Replace(Text13.Text, "10", "PW")

Text14.Text = Replace(Text9.Text, "30", "C4")
Text14.Text = Replace(Text14.Text, "03", "UX")
Text14.Text = Replace(Text14.Text, "55", "I8")
Text14.Text = Replace(Text14.Text, "66", "PS")
Text14.Text = Replace(Text14.Text, "23", "MZ")
Text14.Text = Replace(Text14.Text, "32", "8Q")
Text14.Text = Replace(Text14.Text, "14", "0L")
Text14.Text = Replace(Text14.Text, "41", "XS")
Text14.Text = Replace(Text14.Text, "74", "9U")
Text14.Text = Replace(Text14.Text, "47", "NT")

Text15.Text = Replace(Text10.Text, "27", "Z3")
Text15.Text = Replace(Text15.Text, "91", "8F")
Text15.Text = Replace(Text15.Text, "72", "1K")
Text15.Text = Replace(Text15.Text, "19", "PS")
Text15.Text = Replace(Text15.Text, "56", "O1")
Text15.Text = Replace(Text15.Text, "32", "8Q")
Text15.Text = Replace(Text15.Text, "14", "0L")
Text15.Text = Replace(Text15.Text, "41", "XS")
Text15.Text = Replace(Text15.Text, "74", "9U")
Text15.Text = Replace(Text15.Text, "47", "NT")

Text16.Text = Replace(Text11.Text, "27", "Z3")
Text16.Text = Replace(Text16.Text, "91", "8F")
Text16.Text = Replace(Text16.Text, "72", "1K")
Text16.Text = Replace(Text16.Text, "19", "PS")
Text16.Text = Replace(Text16.Text, "56", "O1")
Text16.Text = Replace(Text16.Text, "65", "M3")
Text16.Text = Replace(Text16.Text, "83", "L0")
Text16.Text = Replace(Text16.Text, "38", "E5")
Text16.Text = Replace(Text16.Text, "01", "XD")
Text16.Text = Replace(Text16.Text, "10", "PW")

Text17.Text = Replace(Text12.Text, "30", "C4")
Text17.Text = Replace(Text17.Text, "03", "UX")
Text17.Text = Replace(Text17.Text, "55", "I8")
Text17.Text = Replace(Text17.Text, "66", "PS")
Text17.Text = Replace(Text17.Text, "23", "MZ")
Text17.Text = Replace(Text17.Text, "32", "8Q")
Text17.Text = Replace(Text17.Text, "14", "0L")
Text17.Text = Replace(Text17.Text, "41", "XS")
Text17.Text = Replace(Text17.Text, "74", "9U")
Text17.Text = Replace(Text17.Text, "47", "NT")
End Sub

Private Sub Command12_Click()
Dim i As Integer

i = Val(Mid(Text1.Text, 1, 1))

Select Case i = Val(Mid(Text1.Text, 1, 1))
    Case i = 1
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
    Case i = 2
        Text18.Text = Text16.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text17.Text
    Case i = 3
        Text18.Text = Text15.Text
        Text19.Text = Text13.Text
        Text20.Text = Text16.Text
        Text21.Text = Text17.Text
        Text22.Text = Text14.Text
    Case i = 4
        Text18.Text = Text13.Text
        Text19.Text = Text14.Text
        Text20.Text = Text16.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
    Case i = 5
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
    Case i = 6
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
    Case i = 7
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
    Case i = 8
        Text18.Text = Text16.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text17.Text
    Case i = 9
        Text18.Text = Text17.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text16.Text
End Select
End Sub

Private Sub Command13_Click()
Text23.Text = Text18.Text & "-" & Text19.Text & "-" & Text20.Text & "-" & Text21.Text & "-" & Text22.Text
End Sub

Private Sub Command14_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label2.Caption = "This generates the request code that the user of the program would send to you to recieve the serial number. This code is generated by combining the hard-disk serial number with the 24 digit Product ID that you specify. The request code is a 25 digit number that is put through a series of processes to finally generate the serial number."
Form2.Label1.Caption = "Generate Unique Request Code"
End Sub

Private Sub Command15_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label2.Caption = "This process splits the 25 digit request code into blocks of 5 digits. Each block is multiplied by a multiplication factor equal to the average of first digit, fourteenth digit, twenty fourth digit of the request and one more digit depending on the value of the first digit of the hard disk serial number. The new numbers are 6 digit numbers which are once again trimmed down to 5 digits."
Form2.Label1.Caption = "Split And Morph Code"
End Sub

Private Sub Command16_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label2.Caption = "Selective inversion basically just inverts the first, third and fifth blocks of the forming serial number. In the case that the last digit of the first, third or fifth block is zero then, it will not be inverted. To those wondering what inversion is, it is where the digits are arranged backwards. Eg: 123 --> 321."
Form2.Label1.Caption = "Selective Inversion Process"
End Sub

Private Sub Command17_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label2.Caption = "During Alpha Replacement, certain two digit blocks are replaced either by a alphabet and number or by two alphabets. For each 5 digit block there are different replacements. Eg: 27 --> M3 or 53 --> XS. These do not follow any strict pattern they are random, but defined by the developer."
Form2.Label1.Caption = "Alpha Replacement Sequence"
End Sub

Private Sub Command18_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label1.Caption = "Position Swap"
Form2.Label2.Caption = "Like the title would suggest, this basically just swaps the position of the 5 digit blocks. Once again, where these are moved is dependant on the first digit of the hard disk serial number."
End Sub

Private Sub Command19_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = True
Form2.Label2.Caption = "This basically arranges the serial number into one long string as it would appear on your CD or text file or wherever..."
Form2.Label1.Caption = "Finish"
End Sub

Private Sub Command2_Click()
Text1.Text = SerNum("C") * -1

If Text1.Text < 0 Then
    Text1.Text = Text1.Text * -1
End If

Text1.Text = Val(Invert(Text1.Text))
End Sub

Private Sub Command20_Click()
Form2.Show
Form2.Left = Me.Left + Me.Width
Form2.Top = Me.Top
Form2.Label3.Visible = False
Form2.Label2.Caption = "Everytime you click on the normal process buttons, it performs the function mentioned in the caption and all the steps before. To read more about these click on the note button in the independant control panel on the right. You may also notice the purple buttons in the independant controls panel. These are to be used to perform the various functions with non-standard value (i.e. independant of default values.). With a little experimentation you should get the hang of it... it is fairly easy."
Form2.Label1.Caption = "Instructions For Use"
End Sub

Private Sub Command21_Click()
Form3.Label1.Caption = Text23.Text
Form3.Show
End Sub

Private Sub Command22_Click()
Form4.Show
End Sub

Private Sub Command23_Click()
Dim ret

ret = MsgBox("Each time you digit swap you get a new serial. This is not a one time effect." & vbCrLf & vbCrLf & "Do you still wish to continue", vbYesNo, "Digit Swap Warning")

If ret = vbNo Then Exit Sub

If Text13.Text = "" Then
    Command5_Click
End If

SwapDigits (Val(Mid(Text1.Text, 1, 1)))

Command12_Click
Command13_Click

End Sub

Private Sub SwapDigits(WhichDigit As Integer)
Dim SD As Integer
Dim Block(0 To 5) As String
Dim tmp_dig(0 To 5) As String

SD = Mid(WhichDigit, 1, 1)

If SD = 0 Then Exit Sub

If SD > 5 Then
SD = SD - 5
End If

Block(0) = Text13.Text
Block(1) = Text14.Text
Block(2) = Text15.Text
Block(3) = Text16.Text
Block(4) = Text17.Text

tmp_dig(0) = Mid(Block(0), SD, 1)
tmp_dig(1) = Mid(Block(1), SD, 1)
tmp_dig(2) = Mid(Block(2), SD, 1)
tmp_dig(3) = Mid(Block(3), SD, 1)
tmp_dig(4) = Mid(Block(4), SD, 1)

Block(1) = Mid(Block(1), 1, SD) & tmp_dig(0) & Mid(Block(1), SD + 1, 5 - SD)
Block(2) = Mid(Block(2), 1, SD) & tmp_dig(1) & Mid(Block(2), SD + 1, 5 - SD)
Block(3) = Mid(Block(3), 1, SD) & tmp_dig(2) & Mid(Block(3), SD + 1, 5 - SD)
Block(4) = Mid(Block(4), 1, SD) & tmp_dig(3) & Mid(Block(4), SD + 1, 5 - SD)
Block(0) = Mid(Block(0), 1, SD) & tmp_dig(4) & Mid(Block(0), SD + 1, 5 - SD)

Block(0) = Mid(Block(0), 1, SD - 1) & Mid(Block(0), SD + 1, 6 - SD)
Block(1) = Mid(Block(1), 1, SD - 1) & Mid(Block(1), SD + 1, 6 - SD)
Block(2) = Mid(Block(2), 1, SD - 1) & Mid(Block(2), SD + 1, 6 - SD)
Block(3) = Mid(Block(3), 1, SD - 1) & Mid(Block(3), SD + 1, 6 - SD)
Block(4) = Mid(Block(4), 1, SD - 1) & Mid(Block(4), SD + 1, 6 - SD)

Text13.Text = Block(0)
Text14.Text = Block(1)
Text15.Text = Block(2)
Text16.Text = Block(3)
Text17.Text = Block(4)
End Sub

Private Sub Command24_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
End Sub

Private Sub Command3_Click()
Command2_Click
Text2.Text = Mid(Text1.Text & "5432463276523486583264856", 1, 25) 'replace "5432463276523486583264856" with your own 24 digit ProductID Code
End Sub

Private Function iSplit(orig As String, mFactor As Integer, Partition As Integer) As String
Dim tmp_key As String
Dim tmp_istring(0 To 5) As String

tmp_key = orig

tmp_istring(0) = Val(Mid(tmp_key, 1, 5)) * mFactor
tmp_istring(1) = Val(Mid(tmp_key, 6, 5)) * mFactor
tmp_istring(2) = Val(Mid(tmp_key, 11, 5)) * mFactor
tmp_istring(3) = Val(Mid(tmp_key, 16, 5)) * mFactor
tmp_istring(4) = Val(Mid(tmp_key, 21, 5)) * mFactor

iSplit = tmp_istring(Partition)

End Function

Private Sub Command4_Click()
Command1_Click
If Mid(Text3.Text, 5, 1) <> 0 Then
    Text8.Text = Invert(Text3.Text)
Else
    Text8.Text = Text3.Text
End If
Text10.Text = Invert(Text5.Text)
Text12.Text = Invert(Text7.Text)
Text9.Text = Text4.Text
Text11.Text = Text6.Text
End Sub

Private Sub Command5_Click()
Command4_Click

Text13.Text = Replace(Text8.Text, "27", "Z3")
Text13.Text = Replace(Text13.Text, "91", "8F")
Text13.Text = Replace(Text13.Text, "72", "1K")
Text13.Text = Replace(Text13.Text, "19", "PS")
Text13.Text = Replace(Text13.Text, "56", "O1")
Text13.Text = Replace(Text13.Text, "65", "M3")
Text13.Text = Replace(Text13.Text, "83", "L0")
Text13.Text = Replace(Text13.Text, "38", "E5")
Text13.Text = Replace(Text13.Text, "01", "XD")
Text13.Text = Replace(Text13.Text, "10", "PW")

Text14.Text = Replace(Text9.Text, "30", "C4")
Text14.Text = Replace(Text14.Text, "03", "UX")
Text14.Text = Replace(Text14.Text, "55", "I8")
Text14.Text = Replace(Text14.Text, "66", "PS")
Text14.Text = Replace(Text14.Text, "23", "MZ")
Text14.Text = Replace(Text14.Text, "32", "8Q")
Text14.Text = Replace(Text14.Text, "14", "0L")
Text14.Text = Replace(Text14.Text, "41", "XS")
Text14.Text = Replace(Text14.Text, "74", "9U")
Text14.Text = Replace(Text14.Text, "47", "NT")

Text15.Text = Replace(Text10.Text, "27", "Z3")
Text15.Text = Replace(Text15.Text, "91", "8F")
Text15.Text = Replace(Text15.Text, "72", "1K")
Text15.Text = Replace(Text15.Text, "19", "PS")
Text15.Text = Replace(Text15.Text, "56", "O1")
Text15.Text = Replace(Text15.Text, "32", "8Q")
Text15.Text = Replace(Text15.Text, "14", "0L")
Text15.Text = Replace(Text15.Text, "41", "XS")
Text15.Text = Replace(Text15.Text, "74", "9U")
Text15.Text = Replace(Text15.Text, "47", "NT")

Text16.Text = Replace(Text11.Text, "27", "Z3")
Text16.Text = Replace(Text16.Text, "91", "8F")
Text16.Text = Replace(Text16.Text, "72", "1K")
Text16.Text = Replace(Text16.Text, "19", "PS")
Text16.Text = Replace(Text16.Text, "56", "O1")
Text16.Text = Replace(Text16.Text, "65", "M3")
Text16.Text = Replace(Text16.Text, "83", "L0")
Text16.Text = Replace(Text16.Text, "38", "E5")
Text16.Text = Replace(Text16.Text, "01", "XD")
Text16.Text = Replace(Text16.Text, "10", "PW")

Text17.Text = Replace(Text12.Text, "30", "C4")
Text17.Text = Replace(Text17.Text, "03", "UX")
Text17.Text = Replace(Text17.Text, "55", "I8")
Text17.Text = Replace(Text17.Text, "66", "PS")
Text17.Text = Replace(Text17.Text, "23", "MZ")
Text17.Text = Replace(Text17.Text, "32", "8Q")
Text17.Text = Replace(Text17.Text, "14", "0L")
Text17.Text = Replace(Text17.Text, "41", "XS")
Text17.Text = Replace(Text17.Text, "74", "9U")
Text17.Text = Replace(Text17.Text, "47", "NT")
End Sub

Private Sub Command6_Click()
Dim i As Integer

Command5_Click

i = Val(Mid(Text1.Text, 1, 1))

Select Case i = Val(Mid(Text1.Text, 1, 1))
    Case i = 1
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
        Debug.Print "1"
    Case i = 2
        Text18.Text = Text16.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text17.Text
        Debug.Print "2"
    Case i = 3
        Text18.Text = Text15.Text
        Text19.Text = Text13.Text
        Text20.Text = Text16.Text
        Text21.Text = Text17.Text
        Text22.Text = Text14.Text
        Debug.Print "3"
    Case i = 4
        Text18.Text = Text13.Text
        Text19.Text = Text14.Text
        Text20.Text = Text16.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
        Debug.Print "4"
    Case i = 5
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
        Debug.Print "5"
    Case i = 6
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
        Debug.Print "6"
    Case i = 7
        Text18.Text = Text14.Text
        Text19.Text = Text16.Text
        Text20.Text = Text13.Text
        Text21.Text = Text17.Text
        Text22.Text = Text15.Text
        Debug.Print "7"
    Case i = 8
        Text18.Text = Text16.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text17.Text
        Debug.Print "8"
    Case i = 9
        Text18.Text = Text17.Text
        Text19.Text = Text13.Text
        Text20.Text = Text15.Text
        Text21.Text = Text14.Text
        Text22.Text = Text16.Text
        Debug.Print "9"
End Select

Debug.Print "i = " & i
End Sub

Private Sub Command7_Click()
Command6_Click

Text23.Text = Text18.Text & "-" & Text19.Text & "-" & Text20.Text & "-" & Text21.Text & "-" & Text22.Text
End Sub

Private Sub Command8_Click()
Dim ret As String
Dim pID As String

If Text1.Text = "" Then
Command2_Click
End If
ret = InputBox("Please enter a 24-Digit Product ID... This should contain only numbers or you will encounter an error. If your code is more or less than 24 digits, it will automatically be rectified.", "Product ID Required")
pID = ret & "232323232323232323232323"
pID = Mid(pID, 1, 24)

Text2.Text = Mid(Text1.Text & pID, 1, 25)
End Sub

Private Sub Command9_Click()
Dim MFact As Integer

MFact = Int(Val(Val(Val(Mid(Text2.Text, 1, 1)) + Val(Mid(Text2.Text, 12, 1)) + Val(Mid(Text2.Text, 24, 1)) + Val(Mid(Text2.Text, Val(Mid(Text2.Text, 1, 1)), 1))) / 4))
Text3.Text = iSplit(Text2.Text, MFact, 0)
Text4.Text = iSplit(Text2.Text, MFact, 1)
Text5.Text = iSplit(Text2.Text, MFact, 2)
Text6.Text = iSplit(Text2.Text, MFact, 3)
Text7.Text = iSplit(Text2.Text, MFact, 4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAbout_Click()
Form4.Show
End Sub

Private Sub mnuDS_Click()
Command23_Click
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLabel_Click()
Command21_Click
End Sub

Private Sub mnuLhelp_Click()
Command20_Click
End Sub

Private Sub mnuReg_Click()
If Text2.Text = "" Or Text23.Text = "" Then
    Command7_Click
End If
Form5.Text1.Text = Text2.Text
Form5.Text1.Tag = Text23.Text
Form5.Show
End Sub
