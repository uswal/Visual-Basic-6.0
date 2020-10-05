VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Text            =   "Text3"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   5520
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "Transection Yes"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "No"
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "AccType"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "O Date"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "O Balance"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   3600
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   3840
      TabIndex        =   22
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
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "MASTER"
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\Spam\Visual Basic\Bank\Bank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TRANS"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Deposit"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Withdraw"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Acc No :"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Transaction Type :"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "State Bank Of India"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Transaction Date:"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Amount"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Label10.Caption = " "
    Label11.Caption = " "
    Label12.Caption = " "
    Label13.Caption = " "
    Text1.SetFocus
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim Acc As String
    Dim TType As String
    Dim TDate As Date
    Dim Amt, ramt As Integer
    Dim A As Long
    
    A = Val(Text1.Text)
    Data1.Refresh
    While Not Data1.Recordset.EOF
        If Data1.Recordset.Fields("AccNo") = Trim(A) Then
           GoTo x
        Else
        End If
        Data1.Recordset.MoveNext
    Wend
x:
    
    ramt = Val(Data1.Recordset.Fields("OBalance"))
    Amt = Val(Text3.Text)
    
    If Option1.Value Then
        Amt = ramt + Amt
        Data1.Recordset.Edit
        Data1.Recordset.Fields("OBalance") = Val(Amt)
        Data1.Recordset.Update
        
        Data2.Recordset.AddNew
        Data2.Recordset.Fields("AccNo") = Val(Text1.Text)
        Data2.Recordset.Fields("TType") = Option1.Caption
        Data2.Recordset.Fields("TDate") = Text2.Text
        Data2.Recordset.Fields("Amount") = Text3.Text
        Data2.Recordset.Update
        MsgBox ("Deposit successful")
    End If
    
    If Option2.Value Then
    
        If ramt < Amt Then
            MsgBox ("Not enough balance")
            GoTo y
        End If
        
        Amt = ramt - Amt
        Data1.Recordset.Edit
        Data1.Recordset.Fields("OBalance") = Val(Amt)
        Data1.Recordset.Update
        
        Data2.Recordset.AddNew
        Data2.Recordset.Fields("AccNo") = Val(Text1.Text)
        Data2.Recordset.Fields("TType") = Option2.Caption
        Data2.Recordset.Fields("TDate") = Text2.Text
        Data2.Recordset.Fields("Amount") = Text3.Text
        Data2.Recordset.Update
        MsgBox ("Withdrawl successful")
        
    End If
    MsgBox ("New Balance" + Str(Amt))
y:
    MsgBox ("Press new to start new transaction")
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Label10.Caption = " "
    Label11.Caption = " "
    Label12.Caption = " "
    Label13.Caption = " "
    Text1.SetFocus
    
End Sub

Private Sub Command5_Click()
    MsgBox ("Transaction has been cancelled")
    Text1.Text = " "
    Text3.Text = " "
    
    Label10.Caption = " "
    Label11.Caption = " "
    Label12.Caption = " "
    Label13.Caption = " "
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Dim D As String
    D = Format(Now, "DD/MM/YYYY")
    Text1.Text = " "
    Text2.Text = D
    Text3.Text = " "
    Label10.Caption = " "
    Label11.Caption = " "
    Label12.Caption = " "
    Label13.Caption = " "
End Sub

Private Sub text1_lostfocus()
    Dim A As String
    Dim D As Integer
    A = Text1.Text
    D = 0
    Data1.Refresh
    While Not Data1.Recordset.EOF
        If Data1.Recordset.Fields("AccNo") = Trim(A) Then
            Label10.Caption = Data1.Recordset.Fields("CName")
            Label11.Caption = Data1.Recordset.Fields("AccType")
            Label12.Caption = Data1.Recordset.Fields("ODate")
            Label13.Caption = Data1.Recordset.Fields("OBalance")
            D = 1
            Text3.SetFocus
        Else
        End If
    
        Data1.Recordset.MoveNext
    Wend
    
    If D = 0 Then
        MsgBox ("Account not found")
        Text1.SetFocus
    End If
    

End Sub


