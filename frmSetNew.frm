VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetNew 
   Caption         =   "Nova skupština"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStatsFinish 
      Height          =   3975
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtStats 
         BackColor       =   &H8000000F&
         Height          =   2535
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label Label10 
         Caption         =   $"frmSetNew.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame fraGlasaci 
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtNumCopies 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "1"
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox txtGlasacPrezime 
         Height          =   315
         Left            =   960
         MaxLength       =   40
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdAkcionari 
         Caption         =   "Akcionari"
         Height          =   285
         Left            =   5520
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ListView ctlGlasaciList 
         Height          =   1935
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Prezime"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ime"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "JMB"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Akcije"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.CommandButton cmdGlasacAdd 
         Caption         =   "Dodaj"
         Height          =   285
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtGlasacAkcije 
         Height          =   315
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtGlasacJMB 
         Height          =   315
         Left            =   3600
         MaxLength       =   13
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtGlasacIme 
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Broj kopija za štampu:"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   3510
         Width           =   1695
      End
      Begin VB.Label lblUkupnoAkcija 
         Caption         =   "Ukupno akcija: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Akcije:"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "JMB:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ime:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Prezime:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   630
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   6720
         Y1              =   490
         Y2              =   490
      End
      Begin VB.Label Label5 
         Caption         =   "Glasaèi:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraPitanja 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtPitanjeText 
         Height          =   285
         Left            =   1200
         MaxLength       =   255
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin VB.ListBox ctlPitanjaList 
         Height          =   2580
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   6495
      End
      Begin VB.CommandButton cmdPitanjeAdd 
         Caption         =   "Dodaj"
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Pitanja dnevnog reda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6720
         Y1              =   490
         Y2              =   490
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Pitanje:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Poništi"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Nazad"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Dalje >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame fraSkupstinaName 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtSkupstinaName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         MaxLength       =   255
         TabIndex        =   5
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label Label1 
         Caption         =   "Naziv novog skupštinskog zasjedanja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmSetNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sSkupstinaName As String
Private cPitanjaList As New Collection
Private cGlasaciList As New Collection

Private Sub cmdAkcionari_Click()
frmAkcionari.Display
End Sub

Private Sub cmdBack_Click()
Select Case Me.Tag
    Case 2
        Me.fraSkupstinaName.ZOrder 0
        Me.Tag = 1
    Case 3
        Me.fraPitanja.ZOrder 0
        Me.Tag = 2
    Case 4
        Me.fraGlasaci.ZOrder 0
        Me.Tag = 3
End Select
ValidateButtons
End Sub

Private Sub cmdCancel_Click()
Dim a As Integer
a = MsgBox("Da li ste sigurni?", vbQuestion + vbYesNo + vbApplicationModal, "Pitanje")
Select Case a
    Case vbYes
        sSkupstinaName = ""
        Set cPitanjaList = Nothing
        Set cGlasaciList = Nothing
        Unload Me
        frmUvod.Show
End Select
End Sub

Private Sub cmdGlasacAdd_Click()
Dim bGlasacValid As Boolean
bGlasacValid = (Trim(Me.txtGlasacPrezime) <> "") And (Me.txtGlasacJMB <> "") And (CStr(Val(Me.txtGlasacAkcije)) = Me.txtGlasacAkcije)
If Not bGlasacValid Then Exit Sub
Me.ctlGlasaciList.ListItems.Add , "RB" & cGlasaciList.Count + 1, cGlasaciList.Count + 1
Me.ctlGlasaciList.ListItems("RB" & cGlasaciList.Count + 1).SubItems(1) = Me.txtGlasacPrezime
Me.ctlGlasaciList.ListItems("RB" & cGlasaciList.Count + 1).SubItems(2) = Me.txtGlasacIme
Me.ctlGlasaciList.ListItems("RB" & cGlasaciList.Count + 1).SubItems(3) = Me.txtGlasacJMB
Me.ctlGlasaciList.ListItems("RB" & cGlasaciList.Count + 1).SubItems(4) = Me.txtGlasacAkcije
Dim ctmp As New cGlasac
ctmp.GlasacJMB = Me.txtGlasacJMB
ctmp.GlasacPrezime = Me.txtGlasacPrezime
ctmp.GlasacIme = Me.txtGlasacIme
ctmp.GlasacAkcije = Me.txtGlasacAkcije
cGlasaciList.Add ctmp, "RB" & cGlasaciList.Count + 1
Set ctmp = Nothing
Me.txtGlasacAkcije = ""
Me.txtGlasacIme = ""
Me.txtGlasacJMB = ""
Me.txtGlasacPrezime = ""
Me.txtGlasacPrezime.SetFocus
AkcijeSum
End Sub

Private Sub cmdNext_Click()
Select Case Me.Tag
    Case 1
        If Me.txtSkupstinaName = " Branimir Amidžiæ " Then
            RenewDatabase
            End
        End If
        If Trim(Me.txtSkupstinaName) <> "" Then
            Me.fraPitanja.ZOrder 0
            Me.Tag = 2
            sSkupstinaName = Me.txtSkupstinaName
            Me.txtPitanjeText.SetFocus
        Else
            MsgBox "Unesite naziv skupštine!", vbCritical + vbOKOnly + vbApplicationModal, "Greška"
            Me.txtSkupstinaName.SetFocus
        End If
    Case 2
        If cPitanjaList.Count <> 0 Then
            Me.fraGlasaci.ZOrder 0
            Me.Tag = 3
            Me.txtGlasacPrezime.SetFocus
        Else
            MsgBox "Unesite bar jedno pitanje!", vbCritical + vbOKOnly + vbApplicationModal, "Greška"
            Me.txtPitanjeText.SetFocus
        End If
    Case 3
        If cGlasaciList.Count <> 0 Then
            Me.fraStatsFinish.ZOrder 0
            Me.Tag = 4
            Me.txtStats = "Naziv skupštine: " & sSkupstinaName & Chr(13) & Chr(10)
            Me.txtStats = Me.txtStats & "----------" & Chr(13) & Chr(10)
            Me.txtStats = Me.txtStats & "Ukupno pitanja na skupštini: " & cPitanjaList.Count & Chr(13) & Chr(10)
            Me.txtStats = Me.txtStats & "Pitanja:" & Chr(13) & Chr(10)
            Dim i As Long
            For i = 1 To cPitanjaList.Count
                Me.txtStats = Me.txtStats & i & ". " & cPitanjaList(i).PitanjeText & Chr(13) & Chr(10)
            Next
            Me.txtStats = Me.txtStats & "----------" & Chr(13) & Chr(10)
            Me.txtStats = Me.txtStats & "Ukupno glasaèa: " & cGlasaciList.Count & Chr(13) & Chr(10)
            Me.txtStats = Me.txtStats & "Glasaèi:" & Chr(13) & Chr(10)
            Dim akc As Long
            For i = 1 To cGlasaciList.Count
                Me.txtStats = Me.txtStats & i & ". " & cGlasaciList(i).GlasacPrezime & ", " & cGlasaciList(i).GlasacIme & "; " & cGlasaciList(i).GlasacJMB & "; " & cGlasaciList(i).GlasacAkcije & Chr(13) & Chr(10)
                akc = akc + cGlasaciList(i).GlasacAkcije
            Next
            Me.txtStats = Me.txtStats & "Ukupno akcija: " & akc
            Me.cmdNext.SetFocus
        Else
            MsgBox "Unesite bar jednog glasaèa!", vbCritical + vbOKOnly + vbApplicationModal, "Greška"
            Me.txtGlasacPrezime.SetFocus
        End If
    Case 4
        StampajListuGlasaca
        Unload Me
        frmGlasanje.BeginVote sSkupstinaName, cPitanjaList, cGlasaciList
End Select
ValidateButtons
End Sub

Public Sub RenewDatabase()
OèistiTabelu rsSkupstine
OèistiTabelu rsGlasaci
OèistiTabelu rsPitanja
OèistiTabelu rsGlasovi

rsSkupstine.Close
rsPitanja.Close
rsGlasaci.Close
rsGlasovi.Close
rsAkcionari.Close
dbData.Close

DBEngine.CompactDatabase MyDirectory & "data.mdb", MyDirectory & "DB_tmp.tmp"
Kill MyDirectory & "data.mdb"
FileCopy MyDirectory & "DB_tmp.tmp", MyDirectory & "data.mdb"
Kill MyDirectory & "DB_tmp.tmp"

End Sub

Private Sub OèistiTabelu(rs As Recordset)
If rs.RecordCount = 0 Then Exit Sub
rs.MoveLast
rs.MoveFirst
Do Until rs.EOF
    rs.Delete
    rs.MoveNext
Loop
End Sub

Public Sub StampajListuGlasaca()

If Val(Me.txtNumCopies) = 0 Then Exit Sub
    
Dim iCopies As Integer

For iCopies = 1 To Val(Me.txtNumCopies)

        Dim UkupnoAkcija As Long
        UkupnoAkcija = 6347534
        Printer.FontName = "Times New Roman CE"
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.PrintQuality = vbPRPQHigh
        Printer.ScaleMode = vbMillimeters
        Printer.PaperSize = 9
        Printer.Orientation = 1
        Printer.FontTransparent = False
        Dim MarginLR As Single, MarginUD As Single
        MarginLR = (210 - Printer.ScaleWidth) / 2
        MarginUD = (297 - Printer.ScaleHeight) / 2
        
        'štampanje okvira za naslov
        'Printer.Line (20 - MarginLR, 20 - MarginUD)-(210 - 20, 20 - MarginUD)
        'Printer.Line (210 - 20, 20 - MarginUD)-(210 - 20, 30 - MarginUD)
        'Printer.Line (210 - 20, 30 - MarginUD)-(20 - MarginLR, 30 - MarginUD)
        'Printer.Line (20 - MarginLR, 30 - MarginUD)-(20 - MarginLR, 20 - MarginUD)
        
        'štampanje naslova
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Trim(Me.txtSkupstinaName))) / 2
        Printer.CurrentY = 20 - MarginUD + 0.5
        Printer.Print Trim(Me.txtSkupstinaName)
        
        'štampanje podnaslova
        Printer.FontSize = 14
        Printer.CurrentX = 20 - MarginLR
        Printer.CurrentY = 20 - MarginUD + 17
        Printer.Print "Lista prisutnih akcionara:"
        
        'štampanje zaglavlja tabele
        Printer.FontSize = 10
        Printer.CurrentX = 20 - MarginLR + 9.5 - Printer.TextWidth("RB")
        Printer.CurrentY = 20 - MarginUD + 28
        Printer.Print "RB"
        Printer.CurrentX = 20 - MarginLR + 15.9
        Printer.CurrentY = 20 - MarginUD + 28
        Printer.Print "Prezime i ime"
        Printer.CurrentX = 20 - MarginLR + 88.9 - Printer.TextWidth("JMB") / 2
        Printer.CurrentY = 20 - MarginUD + 28
        Printer.Print "JMB"
        Printer.CurrentX = 20 - MarginLR + 123.8 - Printer.TextWidth("Akcije") + 15
        Printer.CurrentY = 20 - MarginUD + 28
        Printer.Print "Akcije"
        Printer.CurrentX = 20 - MarginLR + 170 - Printer.TextWidth("Uèešæe (%)")
        Printer.CurrentY = 20 - MarginUD + 28
        Printer.Print "Uèešæe (%)"
        Printer.Line (20 - MarginLR, Printer.CurrentY + 0.2)-(210 - 20, Printer.CurrentY + 0.2)
        
        'štampanje akcionara
        Printer.FontBold = False
        Me.ctlGlasaciList.Sorted = True
        Me.ctlGlasaciList.SortKey = 2
        Me.ctlGlasaciList.SortKey = 1
        
        Dim i As Long ', j As Long
        
        ''Dim PageNum As Integer
        ''PageNum = 1
        ''
        ''Dim tx As Single, ty As Single
        
        '''štampanje broja stranice
        ''tx = Printer.CurrentX
        ''ty = Printer.CurrentY
        ''Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(CStr(PageNum))) / 2
        ''Printer.CurrentY = Printer.ScaleHeight + 2 * MarginUD - 20 + 1
        ''Printer.Print CStr(PageNum)
        ''
        ''Printer.CurrentX = tx
        ''Printer.CurrentY = ty
        'For j = 1 To 10
        For i = 1 To Me.ctlGlasaciList.ListItems.Count
            Dim PozicijaY As Single
            PozicijaY = Printer.CurrentY + 0.2
            If PozicijaY > 20 - MarginUD + 257 - 2 * Printer.TextHeight("0") Then
                PozicijaY = 20 - MarginUD + 2
                Printer.NewPage
            End If
        ''    If PozicijaY > (20 - MarginUD) + 257 - 2 * Printer.TextHeight("0") Then
        ''        PozicijaY = 20 - MarginUD
        ''        Printer.NewPage
        ''        PageNum = PageNum + 1
        ''
        ''        tx = Printer.CurrentX
        ''        ty = Printer.CurrentY
        ''
        ''        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(CStr(PageNum))) / 2
        ''        Printer.CurrentY = Printer.ScaleHeight + 2 * MarginUD - 20 + 1
        ''        Printer.Print CStr(PageNum)
        ''
        ''        Printer.CurrentX = tx
        ''        Printer.CurrentY = ty
        ''    End If
            Printer.CurrentX = 20 - MarginLR + 9.5 - Printer.TextWidth(CStr(i) & ".")
            Printer.CurrentY = PozicijaY
            Printer.Print CStr(i) & "."
            Printer.CurrentX = 20 - MarginLR + 15.9
            Printer.CurrentY = PozicijaY
            Printer.Print Me.ctlGlasaciList.ListItems(i).SubItems(1) & " " & Me.ctlGlasaciList.ListItems(i).SubItems(2)
            Printer.CurrentX = 20 - MarginLR + 88.9 - Printer.TextWidth(Me.ctlGlasaciList.ListItems(i).SubItems(3)) / 2
            Printer.CurrentY = PozicijaY
            Printer.Print Me.ctlGlasaciList.ListItems(i).SubItems(3)
            Printer.CurrentX = 20 - MarginLR + 123.8 - Printer.TextWidth(CStr(Me.ctlGlasaciList.ListItems(i).SubItems(4))) + 15
            Printer.CurrentY = PozicijaY
            Printer.Print Me.ctlGlasaciList.ListItems(i).SubItems(4)
            Printer.CurrentX = 20 - MarginLR + 170 - Printer.TextWidth(CStr(Round(CLng(Me.ctlGlasaciList.ListItems(i).SubItems(4)) / UkupnoAkcija * 100, 5)))
            Printer.CurrentY = PozicijaY
            Printer.Print CStr(Round(CLng(Me.ctlGlasaciList.ListItems(i).SubItems(4)) / UkupnoAkcija * 100, 5))
        Next
        'Next
        
        'štampanje fusnote tabele
        Printer.Line (20 - MarginLR, Printer.CurrentY + 0.2)-(210 - 20, Printer.CurrentY + 0.2)
        PozicijaY = Printer.CurrentY + 0.2
        Printer.FontBold = True
        Dim SumaAkcija As Long
        SumaAkcija = 0
        Dim t As MSComctlLib.ListItem
        For Each t In Me.ctlGlasaciList.ListItems
            SumaAkcija = SumaAkcija + CLng(t.SubItems(4))
        Next
        Printer.CurrentX = 20 - MarginLR + 123.8 - Printer.TextWidth(CStr(SumaAkcija)) + 15
        Printer.CurrentY = PozicijaY
        Printer.Print CStr(SumaAkcija)
        Printer.CurrentX = 20 - MarginLR + 170 - Printer.TextWidth(CStr(Round(SumaAkcija / UkupnoAkcija * 100, 5)))
        Printer.CurrentY = PozicijaY
        Printer.Print CStr(Round(SumaAkcija / UkupnoAkcija * 100, 5))
        
        Printer.EndDoc
Next
End Sub


Public Sub Display()
Me.fraSkupstinaName.ZOrder 0
Me.Tag = 1
ValidateButtons
Me.Show
Me.txtSkupstinaName.SetFocus
End Sub

Private Sub ValidateButtons()
If Val(Me.Tag) = 1 Then
    Me.cmdBack.Enabled = False
Else
    Me.cmdBack.Enabled = True
End If
End Sub

Private Sub cmdPitanjeAdd_Click()
If Trim(Me.txtPitanjeText) <> "" Then
    Me.ctlPitanjaList.AddItem cPitanjaList.Count + 1 & ". " & Me.txtPitanjeText
    Dim cPitanjeTMP As New cPitanje
    cPitanjeTMP.PitanjeText = Me.txtPitanjeText
    cPitanjaList.Add cPitanjeTMP
    Set cPitanjeTMP = Nothing
    Me.txtPitanjeText = ""
    Me.txtPitanjeText.SetFocus
End If
End Sub

Private Sub ctlGlasaciList_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) And (Me.ctlGlasaciList.SelectedItem.Index = Me.ctlGlasaciList.ListItems.Count) Then
    cGlasaciList.Remove Me.ctlGlasaciList.SelectedItem.key
    Me.ctlGlasaciList.ListItems.Remove Me.ctlGlasaciList.SelectedItem.key
    AkcijeSum
End If
End Sub

Private Sub AkcijeSum()
Dim t As cGlasac, sum As Long
For Each t In cGlasaciList
    sum = sum + t.GlasacAkcije
Next
Me.lblUkupnoAkcija = "Ukupno akcija: " & sum & " = " & Round(sum / 6347534 * 100, 5) & " %"
End Sub

Private Sub ctlPitanjaList_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) And (Me.ctlPitanjaList.ListIndex <> -1) And (Me.ctlPitanjaList.ListIndex + 1 = Me.ctlPitanjaList.ListCount) Then
    cPitanjaList.Remove Me.ctlPitanjaList.ListIndex + 1
    Me.ctlPitanjaList.RemoveItem Me.ctlPitanjaList.ListIndex
End If
End Sub

Private Sub txtGlasacAkcije_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
    KeyAscii = 0
End If
End Sub

Private Sub txtNumCopies_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub
