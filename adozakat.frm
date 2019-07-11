VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adozakat 
   BackColor       =   &H000000FF&
   Caption         =   "Sistem Maklumat Zakat"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15285
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adozakat 
      Height          =   495
      Left            =   3240
      Top             =   7440
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
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
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\FINAL PROJECT VB\adozakat.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\FINAL PROJECT VB\adozakat.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "adozakat"
      Caption         =   "adozakat"
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
   Begin VB.CommandButton cmdLast 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   18
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   14
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cboAsnaf 
      DataField       =   "KATEGORI ASNAF"
      DataSource      =   "adozakat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "adozakat.frx":0000
      Left            =   4800
      List            =   "adozakat.frx":0016
      TabIndex        =   10
      Text            =   "FAKIR"
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox txtAlamat 
      DataField       =   "ALAMAT"
      DataSource      =   "adozakat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4800
      TabIndex        =   9
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox txtJumlah 
      DataField       =   "JUMLAH BANTUAN"
      DataSource      =   "adozakat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   5880
      Width           =   4815
   End
   Begin VB.TextBox txtNo 
      DataField       =   "IC"
      DataSource      =   "adozakat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      DataField       =   "NAMA"
      DataSource      =   "adozakat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "JUMLAH BANTUAN (RM)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "KATEGORI ASNAF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "NO KAD PENGENALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MAKLUMAT PENERIMA ZAKAT"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12975
   End
End
Attribute VB_Name = "adozakat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdFirst_Click()

    adozakat.Recordset.MoveFirst

End Sub

Private Sub cmdLast_Click()

    adozakat.Recordset.MoveLast

End Sub
Private Sub cmdNext_Click()

     With adozakat.Recordset
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
    End With

End Sub

Private Sub cmdPrevious_Click()

    With adozakat.Recordset
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
    End With

End Sub

Private Sub cmdAdd_Click()

    If cmdAdd.Caption = "&Add" Then
        adozakat.Recordset.AddNew
        txtName.SetFocus
        DisableButtons
        cmdSave.Enabled = True
        cmdAdd.Caption = "&Cancel"
    Else
        adozakat.Recordset.CancelUpdate
        EnableButtons
        cmdSave.Enabled = False
        cmdAdd.Caption = "&Add"
    End If

End Sub

Private Sub cmdDelete_Click()

     With adozakat.Recordset
        .Delete
        .MoveNext
        If .EOF Then
            .MovePrevious
            If .BOF Then
                MsgBox "The recordset is empty.", vbInformation, "No Record."
                DisableButtons
            End If
        End If
    End With

End Sub

Private Sub DisableButtons()

    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdDelete.Enabled = False
    
End Sub

Private Sub EnableButtons()

    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdDelete.Enabled = True
    
End Sub

Private Sub cmdSave_Click()
    If txtName.Text <> "" Then
        adozakat.Recordset.Update
        EnableButtons
        cmdSave.Enabled = False
        cmdAdd.Caption = "&Add"
    Else
        MsgBox "The recordset is empty.", vbInformation, "No Record."
    End If
End Sub


