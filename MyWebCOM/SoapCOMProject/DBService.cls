VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public oConn As New ADODB.Connection
Public oRs As New ADODB.Recordset

Public Function SayHello() As String
    SayHello = "Hi there...I am a COM working through SOAP ready to service you my best!!"
End Function

Public Function GetLoaner() As String

Dim oStream As New ADODB.Stream
Dim ConnStr$
ConnStr = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\..\MyDB\MyDB.mdb"

oConn.Open ConnStr
With oRs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .ActiveConnection = oConn
    .Open "select a.LoanerKey, a.LoanerCode, a.LoanerDesc, a.LoanerDate, a.LoanerAmount, " & _
    " b.StatusCode from Loaner a, Status b where a.StatusKey = b.StatusKey"
End With

'oRs.Save oStream, adPersistADTG
oRs.Save oStream, adPersistXML
GetLoaner = oStream.ReadText

oRs.Close
oConn.Close

End Function

Public Function GetDebtor() As String

Dim oStream As New ADODB.Stream
Dim ConnStr$
ConnStr = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\..\MyDB\MyDB.mdb"

oConn.Open ConnStr
With oRs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .ActiveConnection = oConn
    .Open "select a.DebtorKey, a.DebtorCode, a.DebtorDesc, a.DebtorDate, a.DebtorAmount, " & _
    " b.StatusCode from Debtor a, Status b where a.StatusKey = b.StatusKey"
End With

oRs.Save oStream, adPersistXML
GetDebtor = oStream.ReadText()

oRs.Close
oConn.Close

End Function

Public Function GetStructDebtor() As String
    GetStructDebtor = "DebtorKey I(4)" + Chr(13) + _
    "DebtorCode C(50)" + Chr(13) + Chr(10) + _
    "DebtorDesc C(100)" + Chr(13) + Chr(10) + _
    "DebtorDate T(8)" + Chr(13) + Chr(10) + _
    "DebtorAmount N(8,2)" + Chr(13) + Chr(10) + _
    "StatusKey I(4)"
End Function

Public Function GetStructLoaner() As String
    GetStructLoaner = "LoanerKey I(4)" + Chr(13) + Chr(10) + _
    "LoanerCode C(50)" + Chr(13) + Chr(10) + _
    "LoanerDesc C(100)" + Chr(13) + Chr(10) + _
    "LoanerDate T(8)" + Chr(13) + Chr(10) + _
    "LoanerAmount N(8,2)" + Chr(13) + Chr(10) + _
    "StatusKey I(4)"
End Function

Public Function ExecuteSql(sSql As String) As Integer
    On Error GoTo ErrHandler
    Dim ConnStr$
    ConnStr = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\..\MyDB\MyDB.mdb"

    oConn.Open ConnStr
    oConn.Execute sSql
    oConn.Close
    
    ExecuteSql = 1
    
    Exit Function

ErrHandler:
    ExecuteSql = 0
    
End Function


