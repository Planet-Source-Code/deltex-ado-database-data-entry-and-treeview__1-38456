VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTreeDataEntry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Treeview with Data-Entry"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame SSFrame3 
      Height          =   1845
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   4650
      _Version        =   65536
      _ExtentX        =   8202
      _ExtentY        =   3254
      _StockProps     =   14
      Caption         =   "Few Notes :"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         Caption         =   "* Uses only one table for simplicity. (ADO for connection)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   26
         Top             =   330
         Width           =   4305
      End
      Begin VB.Label Label4 
         Caption         =   $"frmTreeDataEntry.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   285
         TabIndex        =   25
         Top             =   645
         Width           =   4005
      End
      Begin VB.Label Label5 
         Caption         =   "* Code is auto-increment and updated only once the record is saved."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   285
         TabIndex        =   24
         Top             =   1305
         Width           =   3765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1470
      Left            =   105
      TabIndex        =   21
      Top             =   2895
      Width           =   4665
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3090
         TabIndex        =   14
         Top             =   945
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   13
         Top             =   945
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   12
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3090
         TabIndex        =   11
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   10
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   9
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Timer tmrTree 
      Interval        =   1000
      Left            =   4815
      Top             =   4350
   End
   Begin MSComctlLib.StatusBar stbTreeView 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Top             =   6390
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "by: DelTex   (Aug. 2002)"
            TextSave        =   "by: DelTex   (Aug. 2002)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Planet-Source-Code"
            TextSave        =   "Planet-Source-Code"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   4920
      TabIndex        =   19
      Top             =   4545
      Width           =   3810
      _Version        =   65536
      _ExtentX        =   6720
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Refresh Mode :"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refres&h"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   17
         Top             =   405
         Width           =   870
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Ma&nual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1335
         TabIndex        =   16
         Top             =   435
         Width           =   870
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "A&utomatic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   405
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   4755
      Top             =   3495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeDataEntry.frx":008E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeDataEntry.frx":01A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeDataEntry.frx":02B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwItems 
      Height          =   4110
      Left            =   4860
      TabIndex        =   18
      Top             =   270
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   7250
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgTree"
      Appearance      =   1
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   4635
      _Version        =   65536
      _ExtentX        =   8176
      _ExtentY        =   4710
      _StockProps     =   14
      Caption         =   "Data-Entries"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker dpkDateEnt 
         Height          =   315
         Left            =   1215
         TabIndex        =   8
         ToolTipText     =   "Date Entered"
         Top             =   1995
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19595265
         CurrentDate     =   37497
      End
      Begin VB.TextBox txtUM 
         Height          =   315
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Unit of Measurement"
         Top             =   1515
         Width           =   1275
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1215
         MaxLength       =   35
         TabIndex        =   4
         ToolTipText     =   "Item Description"
         Top             =   1005
         Width           =   3225
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         ToolTipText     =   "Item Code"
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Da&te :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   570
         TabIndex        =   7
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "U/&M :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   645
         TabIndex        =   5
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Descri&ption :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Top             =   975
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Code :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   525
         Width           =   510
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Hit <ESC> key to end without prompt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   6000
      TabIndex        =   22
      Top             =   6165
      Width           =   2805
   End
End
Attribute VB_Name = "frmTreeDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCon As String, strSQL As String, strPath As String
Public blnAddMode As Boolean    ' indicator if operation is in addition mode (true) or editing mode (false)

Private WithEvents conTree As ADODB.Connection      ' for our database connection
Attribute conTree.VB_VarHelpID = -1
Private WithEvents rsTreeItems As ADODB.Recordset   ' for tblItem recordset
Attribute rsTreeItems.VB_VarHelpID = -1

Public strNodeLabel As String                       ' Node description, to be used
                                                    ' to restore the original label
                                                    ' after the node has been edited (you'll find that out later,
                                                    ' or when you run the program and trick the treeviews).
                                                    ' It is because Treeviews, by default, have editable labels
                                                    ' If you dont want to use APIs to make treeview uneditable,
                                                    ' You are free to use this technique

Const strDbase = "mdbTree.mdb"

Private Sub cmdAdd_Click()
    strSQL = "select * from tblITems " ' order by desc"
    clearText
    enableSaveCancel
    disableDMButtons
    enableText
    blnAddMode = True
    txtDesc.SetFocus
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Cancel your changes now?", vbYesNo + vbQuestion + vbDefaultButton2, "Cancel") = vbYes Then
        disableSaveCancel
        enableDMButtons
        disableText
        blnAddMode = False
        
        cmdAdd.SetFocus
    End If
End Sub

Private Sub cmdDel_Click()
    If txtCode.Text = "" Then
        MsgBox "Please select record to be deleted.", vbOKOnly + vbInformation, "Delete"
    Else
        If MsgBox("Are you sure you want to delete the record?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete") = vbYes Then
            deleteRecord
            clearText
            refreshTree
            MsgBox "Record has been deleted.", vbOKOnly + vbInformation, "Delete"
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If txtCode.Text = "" Then
        MsgBox "Please select record to be modified.", vbOKOnly + vbInformation, "EDIT"
    Else
        enableSaveCancel
        disableDMButtons
        enableText
        strSQL = "select * from tblITems where code = " & Val(txtCode.Text)
        blnAddMode = False
        txtDesc.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    initializeTree
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Save your changes now?", vbYesNo + vbQuestion + vbDefaultButton1, "Save") = vbYes Then
        updateEntry
        disableSaveCancel
        enableDMButtons
        disableText
        refreshTree
        blnAddMode = False
        
        cmdAdd.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' ESC key is pressed
    ' make sure form's KEYPREVIEW property is set to true
    
    If KeyAscii = 27 Then
        End
    End If
End Sub

Private Sub Form_Load()
    MsgBox "Thank you very much for downloading this code.", vbOKOnly, "Thanks"
    
    connectDB
    initializeTree
    
    ' center the form
    Left = (Screen.Width - ScaleWidth) / 2
    Top = (Screen.Height - ScaleHeight) / 2

    stbTreeView.Panels(3).Text = Date
    stbTreeView.Panels(4).Text = Time
    
    disableText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbCritical + vbDefaultButton1, "Exit") = vbYes Then
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub optAuto_Click()
    cmdRefresh.Enabled = False
End Sub

Private Sub optManual_Click()
    cmdRefresh.Enabled = True
End Sub



Private Sub tmrTree_Timer()
    ' sets time
    ' make sure that the timer control's INTERVAL property is set to 1000
    stbTreeView.Panels(4).Text = Time
End Sub

Private Sub enableSaveCancel()
    ' disables save and cancel buttons
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub disableSaveCancel()
    ' enables save and cancel buttons
    
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub enableDMButtons()
    ' disables Data Manipulation (DM) buttons
    
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdDel.Enabled = True
End Sub

Private Sub disableDMButtons()
    ' enables Data Manipulation (DM) buttons
    
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDel.Enabled = False
End Sub

Private Sub enableText()
    ' enables all textboxes
    
    txtCode.Locked = True
    txtDesc.Locked = False
    txtUM.Locked = False
    dpkDateEnt.Enabled = True
End Sub

Private Sub disableText()
    txtCode.Locked = True
    txtDesc.Locked = True
    txtUM.Locked = True
    dpkDateEnt.Enabled = False
End Sub

Private Sub showData()
    ' this function will show data in all textboxes
    ' once a valid node is clicked from treeview
    txtCode.Text = rsTreeItems.Fields("code")
    txtDesc.Text = rsTreeItems.Fields("desc")
    txtUM.Text = rsTreeItems.Fields("um")
    dpkDateEnt.Value = rsTreeItems.Fields("dateEnt")
    
End Sub

Private Sub refreshTree()
    ' Refreshes the content of tree
    If optAuto.Value = True Then
        initializeTree
    End If
End Sub

Private Sub connectDB()
    ' connect to Database
    strPath = App.Path & "\"
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & strPath & strDbase
    
    Set conTree = New ADODB.Connection
    conTree.Open strCon
End Sub

Private Sub initializeTree()
    ' initializes our treeview alphabetically
    
    Dim strLetters As String, strKey As String * 1
    Dim nodLetter As Node
    Dim i As Integer
    
    ' This technique is from SAMS 21 Days book
    tvwItems.LineStyle = tvwRootLines     ' Simply change the style of line of the treeview
    tvwItems.Nodes.Clear                  ' always clear
    
    strLetters = "ABCDEFGHIJKLMNOPQRSTUVXYZ"    ' initialize key for the letter tree
    
    For i = 1 To 25
        strKey = Mid(strLetters, i, 1)       ' chop a letter
        Set nodLetter = tvwItems.Nodes.Add(, , strKey, strKey, 1) ' Notice that we have the same key with the text
                                                                    ' because strKey is unique and at the same time,
                                                                    ' we use it as a description to group the table by letters
        Set rsTreeItems = New ADODB.Recordset
        
        ' Get all item(s) that start with current letter
        strSQL = "Select * from tblItems where desc like '" & strKey & "%'" 'order by desc"
        rsTreeItems.Open strSQL, conTree, adOpenStatic, adLockOptimistic
        
        If rsTreeItems.RecordCount > 0 Then ' If we have item(s) that start with current letter, insert it as a child to that letter
            While Not rsTreeItems.EOF
                Set nodLetter = tvwItems.Nodes.Add(strKey, tvwChild, strKey & Str(rsTreeItems.Fields("code")), "(" & rsTreeItems.Fields("code") & ") " & rsTreeItems.Fields("desc"), 2)
                rsTreeItems.MoveNext
                DoEvents
            Wend
        End If
    Next i
    Set rsTreeItems = Nothing   ' Release rsTreeItems from memory
End Sub

Private Sub tvwItems_AfterLabelEdit(Cancel As Integer, NewString As String)
    NewString = strNodeLabel        ' Bring back to its original label
End Sub

Private Sub tvwItems_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 1      ' display closed folder image is parent node is collapsed
End Sub

Private Sub tvwItems_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 3      ' display open folder image if parent node is expanded
End Sub

Private Sub updateEntry()
    ' updates entries once SAVE button is presses
    
    Set rsTreeItems = New ADODB.Recordset
    rsTreeItems.Open strSQL, conTree, adOpenStatic, adLockOptimistic
    If blnAddMode = True Then   ' if operation is addition, add new record
        rsTreeItems.AddNew
    End If
    rsTreeItems.Fields("desc") = Trim(txtDesc.Text)
    rsTreeItems.Fields("um") = Trim(txtUM.Text)
    rsTreeItems.Fields("dateEnt") = dpkDateEnt.Value
    rsTreeItems.Update
    Set rsTreeItems = Nothing
End Sub

Private Sub clearText()
    ' clears the textboxes, and sets the current date to date picker control
    txtCode.Text = ""
    txtDesc.Text = ""
    txtUM.Text = ""
    dpkDateEnt.Value = Date
End Sub

Private Sub tvwItems_NodeClick(ByVal Node As MSComctlLib.Node)
    strNodeLabel = Node
    If Node.Image = 2 Then  ' if a child node (indicated by the TEXT image of the node)
        Dim dblCode As Double
        dblCode = getCode(Node) ' call function GETCODE and pass the value of the NODE
                                ' this function will extract the code of the record
                                
        strSQL = "Select * from tblITems where code = " & dblCode   ' extract the record from the database whose code is
                                                                    ' is equal to the value returned by the function above
        
        Set rsTreeItems = New ADODB.Recordset
        rsTreeItems.Open strSQL, conTree, adOpenStatic, adLockOptimistic
        
        If rsTreeItems.RecordCount = 0 Then     ' if no record was found, it must have been edited or deleted
            MsgBox "Record has been edited or deleted." & vbCrLf & "Please refresh the treeview.", vbInformation + vbOKOnly, "Record"
            
        Else            ' otherwise, show the record details
            showData
        End If
        
        Set rsTreeItems = Nothing
    End If
End Sub

Private Sub deleteRecord()
    Set rsTreeItems = New ADODB.Recordset
    strSQL = "delete from tblItems where code = " & Val(txtCode.Text)
    rsTreeItems.Open strSQL, conTree, adOpenStatic, adLockOptimistic
    Set rsTreeItems = Nothing
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUM.SetFocus
    End If
End Sub

Private Sub txtUM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dpkDateEnt.SetFocus
    End If
End Sub
