Attribute VB_Name = "MyModule1"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
String, ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1

Public MainForm As MainFrm
Public oConn As New ADODB.Connection
Public oRs As New ADODB.Recordset

Sub Main()
    Set MainForm = New MainFrm
    MainForm.Show
End Sub

Public Function GetLoaner(ByVal cWebAddress As String) As String

Dim oSoap As New MSSOAPLib.SoapClient
oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
GetLoaner = oSoap.GetLoaner()

End Function

Public Function GetDebtor(ByVal cWebAddress As String) As String

Dim oSoap As New MSSOAPLib.SoapClient
oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
GetDebtor = oSoap.GetDebtor()

End Function

Public Function ConvertFromvCalDateTime(sDT As Variant) As Date
' From format: YYYYMMDDTHHMMSS (may be
' terminated with Z for UCT time)
Dim vTemp As Variant
sDT = Trim(sDT)
vTemp = DateSerial(Left(sDT, 4), Mid(sDT, 5, 2), Mid(sDT, 7, 2))
'vTemp = vTemp + TimeSerial(Mid(sDT, 10, 2), Mid(sDT, 12, 2), Mid(sDT, 14, 2))
'If Right(sDT, 1) = "Z" Then vTemp = vTemp - 8 / 24
ConvertFromvCalDateTime = vTemp

End Function

Public Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
        "", "C:\", SW_SHOWNORMAL)
End Function

Public Function GetStructDebtor(ByVal cWebAddress As String) As String
    Dim oSoap As New MSSOAPLib.SoapClient
    oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
    GetStructDebtor = oSoap.GetStructDebtor
End Function

Public Function GetStructLoaner(ByVal cWebAddress As String) As String
    Dim oSoap As New MSSOAPLib.SoapClient
    oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
    GetStructLoaner = oSoap.GetStructLoaner
End Function

Public Function SayHello(ByVal cWebAddress As String) As String
On Error GoTo ErrHandler
    Dim oSoap As New MSSOAPLib.SoapClient
    oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
    SayHello = oSoap.SayHello
    Exit Function
        
ErrHandler:
    SayHello = "Error connecting to the web server" + Chr(13) + _
    "Please verify the web address"
End Function

Public Function ExecuteSql(ByVal cWebAddress As String, ByVal sSql As String) As String
    Dim oSoap As New MSSOAPLib.SoapClient
    Dim nRetVal As Integer
    
    oSoap.mssoapinit cWebAddress, "MyComWSDL", "DBServiceSoapPort"
    nRetVal = oSoap.ExecuteSql(sSql)
    
    If nRetVal = 1 Then
        ExecuteSql = "Execution Successful"
    Else
        ExecuteSql = "Error, please check your syntax"
    End If
End Function

Public Function GetSqlExample() As String
GetSqlExample = _
Chr(13) + Chr(10) + "{--------- SQL example ----------}" + Chr(13) + Chr(10) + _
"insert into Loaner (LoanerCode, LoanerDesc, " + _
"LoanerDate, LoanerAmount ,StatusKey) " + _
"values ('asd', 'asdasd', '02/02/02 02:02:02 AM', 12.12, 2)" + _
Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
"insert into Debtor (DebtorCode, DebtorDesc, " + _
"DebtorDate, DebtorAmount ,StatusKey) " + _
"values ('asd', 'asdasd', '02/02/02 02:02:02 AM', 12.12, 2)"

End Function
