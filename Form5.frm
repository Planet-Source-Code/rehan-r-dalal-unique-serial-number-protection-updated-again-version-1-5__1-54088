VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register Software"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1920
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text3 
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
      Left            =   120
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text5 
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
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text7 
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
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFBF2&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "You're Request Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number: (Paste whole serial using ctrl+v into first box)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim SNo As String

SNo = UCase(Text3.Text & "-" & Text4.Text & "-" & Text5.Text & "-" & Text6.Text & "-" & Text7.Text)

If SNo = Text1.Tag Then
MsgBox "You have successfully registered you're software. Thank You!"
Unload Me
Else
MsgBox "Incorrect serial number. Try Again.", vbCritical
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 86 Then
    Text3.MaxLength = 29
End If
Timer1.Enabled = True
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim tmp_serial
Dim Block() As String

Text3.MaxLength = 5

If KeyCode = 86 Then

tmp_serial = Text3.Text
Block() = Split(tmp_serial, "-", 5)

On Error Resume Next

Text3.Text = Block(0)
Text4.Text = Block(1)
Text5.Text = Block(2)
Text6.Text = Block(3)
Text7.Text = Block(4)

End If
End Sub
