VERSION 5.00
Begin VB.Form frmAkcijeModify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Akcije"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSkip 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Akcije:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "JMB:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ime:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Prezime:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmAkcijeModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

