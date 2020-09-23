VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddress 
   Caption         =   "Address Database"
   ClientHeight    =   5910
   ClientLeft      =   3900
   ClientTop       =   1830
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5850
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Refresh"
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   27
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&UpData"
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   26
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Find"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   25
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdNegotiate 
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   21
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdNegotiate 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   20
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdNegotiate 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   19
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdNegotiate 
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskPhone 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      Format          =   "(###) ###-####"
      Mask            =   "(###) ###-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtField 
      DataField       =   "Comments"
      DataSource      =   "Adodc1"
      Height          =   1365
      Index           =   8
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3240
      Width           =   5655
   End
   Begin VB.TextBox txtField 
      DataField       =   "AddressID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtField 
      DataField       =   "Zip"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtField 
      DataField       =   "State"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtField 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtField 
      DataField       =   "First_Name"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtField 
      DataField       =   "Street"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox txtField 
      DataField       =   "Last_Name"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5535
      Visible         =   0   'False
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DB\Address.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DB\Address.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Address"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comments:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AddressID:"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "State:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   12
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "City:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "First Name:"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Street Address:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim varBookMark As Variant


Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
lblRecord.Caption = "Record: " & _
CStr(Adodc1.Recordset.AbsolutePosition) & _
" of " & Str(Adodc1.Recordset.RecordCount)
End Sub


Private Sub cmdAction_Click(Index As Integer)

  On Error GoTo goActionErr
  
With Adodc1

  Select Case Index
  
    Case 0  'Add
      If cmdAction(0).Caption = "&Add" Then
        varBookMark = .Recordset.BookMark
        .Recordset.AddNew
        txtField(0).SetFocus
        cmdAction(0).Caption = "&Cancel"
        SetVisible False
      Else
        .Recordset.CancelUpdate
      If varBookMark > 0 Then
        .Recordset.BookMark = varBookMark
      Else
        .Recordset.MoveFirst
      End If
        cmdAction(Index).Caption = "&Add"
        SetVisible True
     End If
     
    Case 1  'Delete
      If .Recordset.EditMode = False Then
        .Recordset.Delete
        .Recordset.MoveNext
        If .Recordset.EOF Then .Recordset.MoveLast
      Else
         MsgBox "Must update or refresh record before deleting!!"
      End If
      
     Case 2  'Find
       frmFind.Show
       
     Case 3  'Update
       .Recordset.Update
       varBookMark = .Recordset.BookMark
       .Recordset.Requery
       If varBookMark > 0 Then
         .Recordset.BookMark = varBookMark
       Else
         .Recordset.MoveLast
       End If
       cmdAction(0).Caption = "&Add"
       SetVisible True
       
      Case 4  'Refresh
        varBookMark = .Recordset.BookMark
        .Refresh
        If varBookMark > 0 Then
          .Recordset.BookMark = varBookMark
        Else
          .Recordset.MoveLast
        End If
        
      End Select
      
     End With
     
    Exit Sub
    
goActionErr:
  
    MsgBox Err.Description
    
End Sub

Private Sub cmdNegotiate_Click(Index As Integer)
   On Error GoTo goNegotiateErr
   With Adodc1.Recordset
Select Case Index
       Case 0  'First Record
         .MoveFirst
       Case 1  'Previous Record
         .MovePrevious
         If .BOF Then .MoveFirst
       Case 2  'Next Record
         .MoveNext
         If .EOF Then .MoveLast
       Case 3  'Last Record
         .MoveLast
     End Select
   End With
  Exit Sub
  
goNegotiateErr:
  
     MsgBox Err.Description
     
End Sub

Public Sub SetVisible(blnStatus As Boolean)
  
   Dim intIndex As Integer
   
   For intIndex = 0 To 3
     cmdNegotiate(intIndex).Enabled = blnStatus
   Next intIndex
   
   cmdAction(1).Enabled = blnStatus ' Delete
   cmdAction(2).Enabled = blnStatus ' Find
   cmdAction(4).Enabled = blnStatus ' Refresh
End Sub
