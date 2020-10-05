VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Transaction"
      Height          =   615
      Left            =   3240
      TabIndex        =   14
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   615
      Left            =   4680
      TabIndex        =   13
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   615
      Left            =   600
      TabIndex        =   11
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Spam\Visual Basic\Bank\Bank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "MASTER"
      Top             =   9240
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Opening Balance:"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Date Of Opening:"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Account Type:"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Account Number:"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "STATE BANK OF INDIA"
      Height          =   405
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim D As String
    D = Format(Now, "DD/MM/YYYY")

    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = D
    Text4.Text = " "
    Text1.SetFocus
    Command2.Enabled = True

End Sub

Private Sub Command2_Click()
    With Data1.Recordset
    .AddNew
    .Fields("AccNo") = Val(Text1.Text)
    .Fields("CName") = Text2.Text
    .Fields("AccType") = Combo1.Text
    .Fields("ODate") = CDate(Text3.Text)
    .Fields("OBalance") = Val(Text4.Text)
    .Update
    End With
    MsgBox ("One record saved!")
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
    Form2.Show
    
End Sub

Private Sub Form_Load()
    Dim D As String
    D = Format(Now, "DD/MM/YYYY")

    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = D
    Text4.Text = " "
    Combo1.Text = "Select Account Type "
    Combo1.AddItem ("Saving")
    Combo1.AddItem ("Current")
    Combo1.AddItem ("Recurring")
    Combo1.AddItem ("Fixed")
    
End Sub
Private Sub Command2_GotFocus()
    Dim I As Long
    Dim T As String
    
    I = Val(Text4.Text)
    T = UCase(Combo1.Text)
    
    If T = "SAVING" Then
        If I <= 2000 Then
            MsgBox ("You opening balance for your SAVING ACCOUNT can't be below 2000")
            Text4.SetFocus
        Else
        End If
    Else
    End If
    
    If T = "CURRENT" Then
        If I <= 5000 Then
            MsgBox ("You opening balance for your CURRENT ACCOUNT can't be below 5000")
            Text4.SetFocus
        Else
        End If
    Else
    End If
    
    If T = "RECURRING" Then
        If I <= 1000 Then
            MsgBox ("You opening balance for your RECURRING ACCOUNT can't be below 1000")
            Text4.SetFocus
        Else
        End If
    Else
    End If
    
    If T = "FIXED" Then
        If I <= 5000 Then
            MsgBox ("You opening balance for your FIXED ACCOUNT can't be below 5000")
            Text4.SetFocus
            Else
        End If
    Else
    End If
                
    
End Sub
Private Sub text3_gotfocus()
    If Combo1.Text = "Select Account Type " Then
        MsgBox ("Please select account type")
        Combo1.SetFocus
    Else
    End If
End Sub
Private Sub Combo1_GotFocus()
    Dim A As String
    A = Text2.Text
    If Text2.Text = " " Then
        MsgBox ("Name shouldn't be blank")
        Text2.SetFocus
    Else
    End If
End Sub
Private Sub text2_gotfocus()
    Dim A As Long
    A = Val(Text1.Text)
    Data1.Refresh
    While Not Data1.Recordset.EOF
        If Data1.Recordset.Fields("AccNo") = Trim(A) Then
            MsgBox ("Account number isn't unique!")
            Text1.SetFocus
        Else
        End If
    
        Data1.Recordset.MoveNext
    Wend
End Sub
