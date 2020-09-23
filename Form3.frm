VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFBF2&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sample Label Example"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":0000
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number (Keep Carefully)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   630
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00000-11111-22222-33333-44444"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000D2FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   375
      Left            =   120
      Top             =   480
      Width           =   5415
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1095
      Left            =   120
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
