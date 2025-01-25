VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlasac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sGlasacJMB As String
Private m_sGlasacPrezime As String
Private m_sGlasacIme As String
Private m_lGlasacAkcije As Long
Private m_lDBID As Long

Public Property Get GlasacJMB() As String
GlasacJMB = m_sGlasacJMB
End Property

Public Property Let GlasacJMB(sInput As String)
m_sGlasacJMB = sInput
End Property

Public Property Get GlasacPrezime() As String
GlasacPrezime = m_sGlasacPrezime
End Property

Public Property Let GlasacPrezime(sInput As String)
m_sGlasacPrezime = sInput
End Property

Public Property Get GlasacIme() As String
GlasacIme = m_sGlasacIme
End Property

Public Property Let GlasacIme(sInput As String)
m_sGlasacIme = sInput
End Property

Public Property Get GlasacAkcije() As Long
GlasacAkcije = m_lGlasacAkcije
End Property

Public Property Let GlasacAkcije(sInput As Long)
m_lGlasacAkcije = sInput
End Property

Public Property Get DBID() As Long
DBID = m_lDBID
End Property

Public Property Let DBID(sInput As Long)
m_lDBID = sInput
End Property


