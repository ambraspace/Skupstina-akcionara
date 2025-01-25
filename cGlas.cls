VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_lDBID As Long
Private m_lGlasacID As Long
Private m_lPitanjeID As Long
Private m_sGlas As String

Public Property Get DBID() As Long
DBID = m_lDBID
End Property

Public Property Let DBID(sInput As Long)
m_lDBID = sInput
End Property

Public Property Get GlasacID() As Long
GlasacID = m_lGlasacID
End Property

Public Property Let GlasacID(sInput As Long)
m_lGlasacID = sInput
End Property

Public Property Get PitanjeID() As Long
PitanjeID = m_lPitanjeID
End Property

Public Property Let PitanjeID(sInput As Long)
m_lPitanjeID = sInput
End Property

Public Property Get Glas() As String
Glas = m_sGlas
End Property

Public Property Let Glas(sInput As String)
m_sGlas = sInput
End Property
