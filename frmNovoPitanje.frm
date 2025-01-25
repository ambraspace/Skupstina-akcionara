VERSION 5.00
Begin VB.Form frmNovoPitanje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dodaj novo pitanje"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Poništi"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtNovoPitanje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaxLength       =   255
      TabIndex        =   1
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Novo pitanje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmNovoPitanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Display()
On Error Resume Next
Me.Show vbModal
Me.txtNovoPitanje.SetFocus
End Sub

Private Sub cmdCancel_Click()
Me.txtNovoPitanje = ""
Me.Hide
End Sub

Private Sub cmdOK_Click()
If Trim(Me.txtNovoPitanje) = "" Then Exit Sub
frmGlasanje.DodajNovoPitanje Me.txtNovoPitanje
cmdCancel_Click
End Sub

