VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Ping Monitor 2.02"
   ClientHeight    =   1755
   ClientLeft      =   4365
   ClientTop       =   3165
   ClientWidth     =   4770
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblProgName 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ping Monitor Pro "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Ping Server 2.00 Beta"
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

