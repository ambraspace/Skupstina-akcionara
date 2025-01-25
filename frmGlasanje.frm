VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGlasanje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Glasanje"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCommands 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   7095
      Begin VB.CommandButton cmdNovoPitanje 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Novo pitanje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdRaport 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Izvještaj"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   5640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   130
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kraj posla"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   130
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<< Korak nazad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   160
         Width           =   1335
      End
   End
   Begin MSChart20Lib.MSChart ctlPieChart 
      Height          =   1815
      Left            =   4440
      OleObjectBlob   =   "frmGlasanje.frx":0000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3900
      TabIndex        =   19
      Top             =   4140
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3900
      TabIndex        =   18
      Top             =   3780
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3900
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   4140
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   15
      Top             =   3780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   14
      Top             =   3420
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPitanje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Èetvrta taèka dnevnog reda? (4/12)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.Shape Shape18 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6720
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape17 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6720
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6720
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6360
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6360
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6360
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6000
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6000
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6000
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Top             =   4080
      Width           =   2180
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Top             =   3720
      Width           =   2180
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Top             =   3360
      Width           =   2180
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblGlas 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   38.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblGlasacCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "7/11"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblAkcije 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "54644 (3.87991 %)"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   2430
      Width           =   2175
   End
   Begin VB.Label lblJMB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "1805977100013"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   2100
      Width           =   2175
   End
   Begin VB.Label lblIme 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " Petar"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   1770
      Width           =   2175
   End
   Begin VB.Label lblPrezime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " PETROVIÆ"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prvo skupštinsko zasjedanje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmGlasanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AkcijeSuma As Long
Private KolekcijaPitanja As New Collection
Private KolekcijaGlasaca As New Collection
Private KolekcijaGlasova As New Collection

Private Sub cmdBack_Click()
    
Me.cmdEnd.Visible = False
Me.cmdRaport.Visible = False
Me.cmdNovoPitanje.Visible = False
    
If Val(Me.lblPrezime.Tag) > 1 Then
    Me.lblPrezime.Tag = Val(Me.lblPrezime.Tag) - 1
Else
    If Val(Me.lblPitanje.Tag) > 1 Then
        Me.lblPrezime.Tag = KolekcijaGlasaca.Count
        Me.lblPitanje.Tag = Val(Me.lblPitanje.Tag) - 1
    Else
        Beep
        Exit Sub
    End If
End If

rsGlasovi.MoveLast
KolekcijaGlasova.Remove "ID" & rsGlasovi("ID")
rsGlasovi.Delete

OsvježiStranicu

End Sub

Private Sub cmdBack_KeyDown(KeyCode As Integer, Shift As Integer)
Taster KeyCode
End Sub

Private Sub cmdEnd_Click()
rsSkupstine.Close
rsAkcionari.Close
rsGlasaci.Close
rsPitanja.Close
rsGlasovi.Close
dbData.Close
End
End Sub


Private Sub cmdNovoPitanje_Click()
frmNovoPitanje.Display

End Sub


Public Sub DodajNovoPitanje(sPitanje As String)
Dim cPitanjeTMP As New cPitanje
cPitanjeTMP.PitanjeText = sPitanje
rsPitanja.AddNew
rsPitanja("Pitanje") = sPitanje
rsPitanja("Skupstina") = TrenutnaSkupstinaID
cPitanjeTMP.DBID = rsPitanja("ID")
rsPitanja.Update
KolekcijaPitanja.Add cPitanjeTMP, "ID" & cPitanjeTMP.DBID
Set cPitanjeTMP = Nothing
Me.lblPitanje.Tag = KolekcijaPitanja.Count
Me.lblPrezime.Tag = 1
Me.cmdEnd.Visible = False
Me.cmdNovoPitanje.Visible = False
Me.cmdRaport.Visible = False
OsvježiStranicu
End Sub

Private Sub cmdRaport_Click()

frmRaport.Display Me.lblTitle, KolekcijaPitanja, KolekcijaGlasaca, KolekcijaGlasova

End Sub


Private Sub ctlPieChart_KeyDown(KeyCode As Integer, Shift As Integer)
Taster KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Taster KeyCode
End Sub

Private Sub Form_Resize()
On Error Resume Next

RazvuciTLHW Me.lblTitle, 120, 120, 375, 7095
RazvuciFont Me.lblTitle, 14

RazvuciTLHW Me.lblPitanje, 240, 840, 375, 4695
RazvuciFont Me.lblPitanje, 11

RazvuciTLHW Me.lblPrezime, 240, 1440, 225, 2175
RazvuciFont Me.lblPrezime, 9

RazvuciTLHW Me.lblIme, 240, 1770, 225, 2175
RazvuciFont Me.lblIme, 9

RazvuciTLHW Me.lblJMB, 240, 2100, 225, 2175
RazvuciFont Me.lblJMB, 9

RazvuciTLHW Me.lblAkcije, 240, 2430, 225, 2175
RazvuciFont Me.lblAkcije, 9

RazvuciTLHW Me.lblGlasacCount, 840, 2760, 225, 975
RazvuciFont Me.lblGlasacCount, 9

RazvuciTLHW Me.lblGlas, 2640, 1560, 975, 1095
RazvuciFont Me.lblGlas, 38

RazvuciTLHW Me.ctlPieChart, 4440, 1200, 1935, 2415

RazvuciTLHW Me.fraCommands, 120, 4440, 495, 7095

RazvuciTLHW Me.cmdBack, 120, 130, 310, 1450
RazvuciFont Me.cmdBack, 7
RazvuciTLHW Me.cmdNovoPitanje, 1650, 130, 310, 1450
RazvuciFont Me.cmdNovoPitanje, 7
RazvuciTLHW Me.cmdEnd, 4000, 130, 310, 1450
RazvuciFont Me.cmdEnd, 7
RazvuciTLHW Me.cmdRaport, 5550, 130, 310, 1450
RazvuciFont Me.cmdRaport, 7

RazvuciTLHW Me.Shape1, 240, 3360, 375, 2895
RazvuciTLHW Me.Shape2, 240, 3720, 375, 2895
RazvuciTLHW Me.Shape3, 240, 4080, 375, 2895

RazvuciTLHW Me.Shape7, 3840, 3360, 375, 2180
RazvuciTLHW Me.Shape8, 3840, 3720, 375, 2180
RazvuciTLHW Me.Shape9, 3840, 4080, 375, 2180
RazvuciTLHW Me.Shape10, 6000, 3360, 375, 375
RazvuciTLHW Me.Shape11, 6000, 3720, 375, 375
RazvuciTLHW Me.Shape12, 6000, 4080, 375, 375
RazvuciTLHW Me.Shape13, 6360, 3360, 375, 375
RazvuciTLHW Me.Shape14, 6360, 3720, 375, 375
RazvuciTLHW Me.Shape15, 6360, 4080, 375, 375
RazvuciTLHW Me.Shape16, 6720, 3360, 375, 375
RazvuciTLHW Me.Shape17, 6720, 3720, 375, 375
RazvuciTLHW Me.Shape18, 6720, 4080, 375, 375

End Sub

Private Sub RazvuciTLHW(objekt As Control, l As Single, t As Single, h As Single, w As Single)
On Error Resume Next

Dim Hstepen As Single, Wstepen As Single
Hstepen = Me.ScaleHeight / 5055
Wstepen = Me.ScaleWidth / 7320
objekt.Width = w * Wstepen
objekt.Height = h * Hstepen
objekt.Top = t * Hstepen
objekt.Left = l * Wstepen
End Sub

Private Sub RazvuciFont(objekt As Control, fs As Single)
On Error Resume Next

Dim Hstepen As Single
Hstepen = Me.ScaleHeight / 5055

objekt.FontSize = fs * Hstepen
End Sub

Public Sub BeginVote(sSkupstinaName As String, cPitanjaList As Collection, cGlasaciList As Collection)

rsSkupstine.AddNew
rsSkupstine("Naziv") = sSkupstinaName
TrenutnaSkupstinaID = rsSkupstine("ID")
rsSkupstine.Update

Dim i As Long

For i = 1 To cPitanjaList.Count
    rsPitanja.AddNew
    rsPitanja("Skupstina") = TrenutnaSkupstinaID
    rsPitanja("Pitanje") = cPitanjaList(i).PitanjeText
    cPitanjaList(i).DBID = rsPitanja("ID")
    rsPitanja.Update
Next

For i = 1 To cGlasaciList.Count
    rsGlasaci.AddNew
    rsGlasaci("Skupstina") = TrenutnaSkupstinaID
    rsGlasaci("Prezime") = cGlasaciList(i).GlasacPrezime
    rsGlasaci("Ime") = cGlasaciList(i).GlasacIme
    rsGlasaci("JMB") = cGlasaciList(i).GlasacJMB
    rsGlasaci("Broj akcija") = cGlasaciList(i).GlasacAkcije
    cGlasaciList(i).DBID = rsGlasaci("ID")
    rsGlasaci.Update
Next

Dim ctmp As cGlasac
For Each ctmp In cGlasaciList
    AkcijeSuma = AkcijeSuma + ctmp.GlasacAkcije
Next

Dim t1 As cPitanje
For Each t1 In cPitanjaList
    KolekcijaPitanja.Add t1, "ID" & t1.DBID
Next

Dim t2 As cGlasac
For Each t2 In cGlasaciList
    KolekcijaGlasaca.Add t2, "ID" & t2.DBID
Next

Me.lblPitanje.Tag = 1
Me.lblPrezime.Tag = 1

OsvježiStranicu
Me.Show vbModal

End Sub

Public Sub NastaviVote(sSkupstinaName As String, cPitanjaList As Collection, cGlasaciList As Collection, cGlasoviList As Collection, iPitanje As Long, iGlasac As Long)

Dim ctmp As cGlasac
For Each ctmp In cGlasaciList
    AkcijeSuma = AkcijeSuma + ctmp.GlasacAkcije
Next

Set KolekcijaPitanja = cPitanjaList
Set KolekcijaGlasaca = cGlasaciList
Set KolekcijaGlasova = cGlasoviList

Me.lblPitanje.Tag = iPitanje
Me.lblPrezime.Tag = iGlasac

OsvježiStranicu
Me.Show vbModal

End Sub




Private Sub Taster(key As Integer)

If (key = vbKeyZ Or key = vbKeyP Or key = vbKeyS) And Not Me.cmdEnd.Visible Then

Select Case key
    Dim TMPGlas As New cGlas
    Case vbKeyZ
        Me.lblGlas.BackColor = RGB(0, 255, 0)
        Me.lblGlas.ForeColor = RGB(0, 0, 0)
        Me.lblGlas = "Z"
        fnPause 1
        rsGlasovi.AddNew
        rsGlasovi("Pitanje") = KolekcijaPitanja(Val(Me.lblPitanje.Tag)).DBID
        rsGlasovi("Glasac") = KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).DBID
        rsGlasovi("Glas") = "Z"
            TMPGlas.DBID = rsGlasovi("ID")
            TMPGlas.Glas = "Z"
            TMPGlas.GlasacID = rsGlasovi("Glasac")
            TMPGlas.PitanjeID = rsGlasovi("Pitanje")
            KolekcijaGlasova.Add TMPGlas, "ID" & TMPGlas.DBID
            Set TMPGlas = Nothing
        rsGlasovi.Update
    Case vbKeyP
        Me.lblGlas.BackColor = RGB(255, 0, 0)
        Me.lblGlas.ForeColor = RGB(0, 0, 0)
        Me.lblGlas = "P"
        fnPause 1
        rsGlasovi.AddNew
        rsGlasovi("Pitanje") = KolekcijaPitanja(Val(Me.lblPitanje.Tag)).DBID
        rsGlasovi("Glasac") = KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).DBID
        rsGlasovi("Glas") = "P"
            TMPGlas.DBID = rsGlasovi("ID")
            TMPGlas.Glas = "P"
            TMPGlas.GlasacID = rsGlasovi("Glasac")
            TMPGlas.PitanjeID = rsGlasovi("Pitanje")
            KolekcijaGlasova.Add TMPGlas, "ID" & TMPGlas.DBID
            Set TMPGlas = Nothing
        rsGlasovi.Update
    Case vbKeyS
        Me.lblGlas.BackColor = RGB(255, 255, 255)
        Me.lblGlas.ForeColor = RGB(0, 0, 0)
        Me.lblGlas = "S"
        fnPause 1
        rsGlasovi.AddNew
        rsGlasovi("Pitanje") = KolekcijaPitanja(Val(Me.lblPitanje.Tag)).DBID
        rsGlasovi("Glasac") = KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).DBID
        rsGlasovi("Glas") = "S"
            TMPGlas.DBID = rsGlasovi("ID")
            TMPGlas.Glas = "S"
            TMPGlas.GlasacID = rsGlasovi("Glasac")
            TMPGlas.PitanjeID = rsGlasovi("Pitanje")
            KolekcijaGlasova.Add TMPGlas, "ID" & TMPGlas.DBID
            Set TMPGlas = Nothing
        rsGlasovi.Update
End Select

OsvježiStranicu

If Val(Me.lblPrezime.Tag) < KolekcijaGlasaca.Count Then
    Me.lblPrezime.Tag = Val(Me.lblPrezime.Tag) + 1
Else
DajStatistiku
    If Val(Me.lblPitanje.Tag) < KolekcijaPitanja.Count Then
        Me.lblPitanje.Tag = Val(Me.lblPitanje.Tag) + 1
        Me.lblPrezime.Tag = 1
    Else
        Me.lblPrezime.Tag = Val(Me.lblPrezime.Tag) + 1
        Završavaj
        Exit Sub
    End If
End If

OsvježiStranicu

End If
End Sub

Private Sub Završavaj()
MsgBox "Glasanje je završeno!", vbApplicationModal + vbExclamation, "Kraj"

Me.cmdEnd.Visible = True
Me.cmdRaport.Visible = True
Me.cmdNovoPitanje.Visible = True

End Sub

Private Sub OsvježiStranicu()

rsSkupstine.MoveLast
Me.lblTitle = rsSkupstine("Naziv")

Me.lblPitanje = KolekcijaPitanja(Val(Me.lblPitanje.Tag)).PitanjeText & " (" & Val(Me.lblPitanje.Tag) & "/" & KolekcijaPitanja.Count & ")"
Me.lblPrezime = " " & KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).GlasacPrezime
Me.lblIme = " " & KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).GlasacIme
Me.lblJMB = KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).GlasacJMB
Me.lblAkcije = KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).GlasacAkcije & " (" & Round(KolekcijaGlasaca(Val(Me.lblPrezime.Tag)).GlasacAkcije / AkcijeSuma * 100, 5) & " %)"
Me.lblGlasacCount = Val(Me.lblPrezime.Tag) & "/" & KolekcijaGlasaca.Count

Me.lblGlas.BackColor = RGB(0, 0, 0)
Me.lblGlas.ForeColor = RGB(255, 255, 255)
Me.lblGlas = "?"

Dim lZa As Long, lProtiv As Long, lSuzdrzani As Long
Dim t As cGlas
For Each t In KolekcijaGlasova
    If t.PitanjeID = KolekcijaPitanja(Val(Me.lblPitanje.Tag)).DBID Then
        Select Case t.Glas
            Case "Z"
                lZa = lZa + KolekcijaGlasaca("ID" & t.GlasacID).GlasacAkcije
            Case "P"
                lProtiv = lProtiv + KolekcijaGlasaca("ID" & t.GlasacID).GlasacAkcije
            Case "S"
                lSuzdrzani = lSuzdrzani + KolekcijaGlasaca("ID" & t.GlasacID).GlasacAkcije
        End Select
    End If
Next

Me.ctlPieChart.Column = 1
Me.ctlPieChart.Data = lZa / AkcijeSuma * 100
Me.ctlPieChart.Column = 2
Me.ctlPieChart.Data = lProtiv / AkcijeSuma * 100
Me.ctlPieChart.Column = 3
Me.ctlPieChart.Data = lSuzdrzani / AkcijeSuma * 100

Dim max As Single

If lZa + lProtiv + lSuzdrzani = 0 Then
    Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 100
    Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 10
Else
    max = fnMax(lZa / AkcijeSuma * 100, lProtiv / AkcijeSuma * 100, lSuzdrzani / AkcijeSuma * 100)
    
    If (max / 10) = Int(max / 10) Then
        Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = max
    Else
        Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Int((max + 10) / 10) * 10
    End If
    Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Me.ctlPieChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum / 10
End If
End Sub

Private Function fnMax(a As Single, b As Single, c As Single)
fnMax = a
If b > fnMax Then fnMax = b
If c > fnMax Then fnMax = c
End Function

Private Sub DajStatistiku()
Dim a As Integer
Dim sZa As Single, sProtiv As Single, sSuzdrzan As Single
Me.ctlPieChart.Column = 1
sZa = Me.ctlPieChart.Data
Me.ctlPieChart.Column = 2
sProtiv = Me.ctlPieChart.Data
Me.ctlPieChart.Column = 3
sSuzdrzan = Me.ctlPieChart.Data

Dim sString As String
sString = "Rezultat:" & Chr(13) & Chr(10)
sString = sString & "Za: " & Round(sZa, 5) & " %" & Chr(13) & Chr(10)
sString = sString & "Protiv: " & Round(sProtiv, 5) & " %" & Chr(13) & Chr(10)
sString = sString & "Suzdržani: " & Round(sSuzdrzan, 5) & " %"

MsgBox sString, vbApplicationModal + vbInformation, "Rezultat"

End Sub

