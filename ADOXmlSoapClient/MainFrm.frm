VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO and XML Data Retrieval"
   ClientHeight    =   9000
   ClientLeft      =   2745
   ClientTop       =   1020
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10710
   Begin VB.CommandButton CmdTest 
      Caption         =   "Test Server Connection"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   8520
      Width           =   3375
   End
   Begin VB.TextBox TextWebAddress 
      Height          =   285
      Left            =   5520
      TabIndex        =   13
      Text            =   "http://127.0.0.1/MyWebCom/SoapCOM/MyComWSDL.WSDL"
      Top             =   8160
      Width           =   5055
   End
   Begin VB.CommandButton CmdGetDebtor 
      Caption         =   "Get Table &Debtor"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton CmdGetLoaner 
      Caption         =   "Get Table &Loaner"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Frame FrameListView 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   10215
      Begin MSComctlLib.ListView ListViewLoaner 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "DblClick Item To Update"
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Loaner Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Loaner Desc"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewDebtor 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "DblClick Item To Update"
         Top             =   3840
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Debtor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Debtor Desc"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "Query Debtor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "Query Loaner"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame FrameXML 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   10215
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6615
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   9975
         ExtentX         =   17595
         ExtentY         =   11668
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame FrameTextView 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   10215
      Begin VB.TextBox TextXML 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   120
         Width           =   9975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13785
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "List View"
            Key             =   "listview"
            Object.ToolTipText     =   "View in ListView"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text View"
            Key             =   "textview"
            Object.ToolTipText     =   "View as Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "XMLBrowse"
            Key             =   "xmlview"
            Object.ToolTipText     =   "Browse XML Document Tree"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Operation"
            Key             =   "operation"
            Object.ToolTipText     =   "Do operations on the table"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameOperation 
      BorderStyle     =   0  'None
      Caption         =   "FrameOperation"
      Height          =   6735
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   10215
      Begin VB.CommandButton CmdExample 
         Caption         =   "SQL Example"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton CmdExecute 
         Caption         =   "E&xecute SQL"
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton CmdStructDebtor 
         Caption         =   "Get Table Debtor's Structure"
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   5400
         Width           =   2415
      End
      Begin VB.CommandButton CmdStructLoaner 
         Caption         =   "Get Table Loaner's Structure"
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "MainFrm.frx":0000
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Web Server Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Menu padFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuHelpText 
         Caption         =   "&Read Me"
         Index           =   1
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExample_Click()
    Text1.Text = Text1.Text + GetSqlExample
End Sub

Private Sub CmdStructDebtor_Click()
    Text1.Text = ""
    Text1.Text = GetStructDebtor(Trim(TextWebAddress.Text))
    MsgBox Text1.Text
End Sub

Private Sub CmdStructLoaner_Click()
    Text1.Text = ""
    Text1.Text = GetStructLoaner(Trim(TextWebAddress.Text))
    MsgBox Text1.Text
End Sub
Private Sub CmdGetDebtor_Click()
Dim oXmlDom As New MSXML2.DOMDocument
Dim oRs As New ADODB.Recordset
Dim oList As ListItem

TextXML.Text = GetDebtor(Trim(TextWebAddress.Text))

oXmlDom.loadXML TextXML.Text
oXmlDom.Save App.Path & "\MyXML.xml"

WebBrowser1.Navigate App.Path & "\MyXML.xml"

oRs.Open oXmlDom

With ListViewDebtor.ListItems
    .Clear
    If oRs.RecordCount > 0 Then
        While Not oRs.EOF
            Set oList = .Add(, , oRs.Fields("DebtorCode").Value)
            oList.SubItems(1) = oRs.Fields("DebtorDesc").Value
            oList.SubItems(2) = oRs.Fields("DebtorDate").Value
            oList.SubItems(3) = FormatCurrency(oRs.Fields("DebtorAmount").Value)
            oList.SubItems(4) = oRs.Fields("StatusCode").Value
            oRs.MoveNext
        Wend
    End If
End With

End Sub

Private Sub CmdGetLoaner_Click()
Dim oXmlDom As New MSXML2.DOMDocument
Dim oStream As New ADODB.Stream

TextXML.Text = GetLoaner(Trim(TextWebAddress.Text))

oXmlDom.async = False
    oXmlDom.loadXML TextXML.Text
    oXmlDom.Save App.Path & "\MyXML.xml"
    WebBrowser1.Navigate App.Path & "\MyXML.xml"    'there should be a better way to do this
    
    'MsgBox oXmlDom.documentElement.childNodes.Item(1).childNodes.Item(1).nodeName
    'MsgBox oXmlDom.documentElement.childNodes.Item(1).childNodes.Item(1).xml
    
    ' object to browse the node
    Dim objNodeList As MSXML2.IXMLDOMNodeList
    Dim objNode As MSXML2.IXMLDOMNode, i As Integer
    Dim objAttribute As MSXML2.IXMLDOMAttribute
    Dim objNodeMap As MSXML2.IXMLDOMNamedNodeMap
    Dim objNamedItem As MSXML2.IXMLDOMNode
    Dim oList As ListItem

Set objNodeList = oXmlDom.getElementsByTagName("z:row")

With ListViewLoaner.ListItems
.Clear
For i = 0 To (objNodeList.length - 1)
    Set objNode = objNodeList.nextNode
    Set objNodeMap = objNode.Attributes
    
    Set oList = .Add(, , objNodeMap.getNamedItem("LoanerCode").nodeTypedValue)
    oList.SubItems(1) = objNodeMap.getNamedItem("LoanerDesc").nodeTypedValue
    oList.SubItems(2) = ConvertFromvCalDateTime(objNodeMap.getNamedItem("LoanerDate").nodeTypedValue)
    'MsgBox ConvertFromvCalDateTime(objNodeMap.getNamedItem("LoanerDate").nodeValue)
    
    oList.SubItems(3) = FormatCurrency(objNodeMap.getNamedItem("LoanerAmount").nodeTypedValue)
    oList.SubItems(4) = objNodeMap.getNamedItem("StatusCode").nodeTypedValue
        
    ' or you can use this one
    'Set objNamedItem = objNodeMap.getNamedItem("LoanerAmount").nodeTypedValue
    'MsgBox objNamedItem.nodeTypedValue
Next i
End With

End Sub

Private Sub CmdTest_Click()
    MsgBox SayHello(Trim(TextWebAddress.Text))
End Sub

Private Sub CmdExecute_Click()
    MsgBox ExecuteSql(Trim(TextWebAddress.Text), Text1.Text)
End Sub

Private Sub mnuExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub mnuHelpText_Click(Index As Integer)
Dim r As Long, msg As String
r = StartDoc(App.Path & "\ReadMe.txt")
If r <= 32 Then
    MsgBox "Error on opening ReadMe file", , "Error"
End If
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Key
        Case "listview":
            FrameListView.ZOrder 0
        
        Case "textview":
            FrameTextView.ZOrder 0
        
        Case "xmlview":
            FrameXML.ZOrder 0
        
        Case "operation":
            FrameOperation.ZOrder 0
        
    End Select
End Sub
