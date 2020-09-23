VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help:"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   2640
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0091EEFF&
      Height          =   1350
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   2970
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H006ACAF9&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Height = Form1.Height
End Sub

Private Sub Timer1_Timer()
Me.Left = Form1.Left + Form1.Width
Me.Top = Form1.Top
End Sub
