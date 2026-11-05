VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WiaJpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_pfnJpg As String
Private m_wiaJpg As Object ' WIA.ImageFile
Private m_wiaPro As Object ' WIA.ImageProcess
        
Friend Sub New_(aPFN_jpg As String)
    m_pfnJpg = aPFN_jpg
    Set m_wiaJpg = CreateObject("WIA.ImageFile")
    Set m_wiaPro = CreateObject("WIA.ImageProcess")
    m_wiaJpg.LoadFile m_pfnJpg
    m_wiaPro.Filters.Add m_wiaPro.FilterInfos("Convert").FilterID
    m_wiaPro.Filters(1).Properties("FormatID").Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" ' GUID für JPG
End Sub
        
Public Property Get Quality() As Single
    Quality = m_wiaPro.Filters(1).Properties("Quality").Value
End Property
Public Property Let Quality(ByVal Value As Single)
    Quality = m_wiaPro.Filters(1).Properties("Quality").Value = Value
End Property

Public Function SaveAs(Optional AsPFN) As Boolean
    If IsMmissing(AsPFN) Then AsPFN = m_pfnJpg
    m_wiaJpg.SaveFile AsPFN
End Function
