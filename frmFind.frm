VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   1410
   ClientLeft      =   4590
   ClientTop       =   6480
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Enter First Or Last Name you are searching for"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub FindString()

   Dim strFind As String
   Dim intFields As Integer
   
   On Error GoTo FindError
   
   If Trim(txtSearch) <> "" Then
     strFind = Trim(txtSearch)
     With frmAddress.Adodc1.Recordset
       Do Until .EOF
         For intFields = 0 To 1
           If InStr(1, frmAddress.txtField(intFields), strFind, _
                    vbTextCompare) > 0 Then
              frmAddress.txtField(intFields).SelStart = _
                      InStr(1, frmAddress.txtField(intFields), _
                            strFind, vbTextCompare) - 1
              frmAddress.txtField(intFields).SelLength = Len(strFind)
              frmAddress.txtField(intFields).SetFocus
              Exit Sub
            End If
          Next
          .MoveNext
          DoEvents
        Loop
        MsgBox "Record not found"
        .MoveFirst
      End With
     End If
     
     Exit Sub
     
FindError:
   
   MsgBox Err.Description
   Err.Clear
                    
End Sub

Private Sub cmdCancel_Click()

 Unload Me
 Set frmFind = Nothing
End Sub

Private Sub cmdFind_Click()

   If frmAddress.Adodc1.Recordset.RecordCount > 0 Then
     frmAddress.Adodc1.Recordset.MoveFirst
     Call FindString
   Else
     MsgBox "RecordSet is Empty"
   End If
End Sub

Private Sub cmdFindNext_Click()
  
   With frmAddress.Adodc1.Recordset
     .MoveNext
     If .EOF Then
        .MoveLast
        MsgBox "End of File Reached!"
      Else
        Call FindString
      End If
      End With
End Sub
