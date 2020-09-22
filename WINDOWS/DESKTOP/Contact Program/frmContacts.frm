VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployees 
   Caption         =   "Employee Information"
   ClientHeight    =   7830
   ClientLeft      =   -165
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   2160
      Picture         =   "frmContacts.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   735
      Left            =   1440
      Picture         =   "frmContacts.frx":1C84
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   735
      Left            =   720
      Picture         =   "frmContacts.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   3120
      TabIndex        =   21
      Top             =   4560
      Width           =   7095
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox cboTitle 
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox cboDepartment 
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtNotes 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   6855
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Caption         =   "Position Title:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Department:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblNotes 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Contact Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtBday 
         Height          =   285
         Left            =   6000
         TabIndex        =   38
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtEmergencyPhone 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   35
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtEmergecyContact 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtWorkPhone 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtPhone 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtPager 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtMobile 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "(&&&) &&& - &&&&"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtExtension 
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label lblBirthday 
         Caption         =   "Birthday:"
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblEmergencyPhone 
         Caption         =   "Emergency Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblEmergencyContactName 
         Caption         =   "Emergency Contact:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label lblWorkPhone 
         Caption         =   "Work Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Address:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblLocation 
         Caption         =   "City/State/Zip:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPhone 
         Caption         =   "Home Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblPager 
         Caption         =   "Pager:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lbExtension 
         Caption         =   "Ext."
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblEmail 
         Caption         =   "Home Email:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   735
      Left            =   0
      Picture         =   "frmContacts.frx":2C10
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":34DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":475E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwAM 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11456
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim booNewRec           As Boolean
Private mNode           As Node
Private WithEvents rs   As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents conn As ADODB.Connection
Attribute conn.VB_VarHelpID = -1


Private Sub cmdAdd_Click()
    booNewRec = True
    rs.AddNew
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
        rs.Requery
        Call AddNodes
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    If booNewRec = True Then
        rs.Update
        rs.Requery
        Call AddNodes
        booNewRec = False
    Else
        rs.Update
    End If
Exit Sub

ErrorHandler:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Load()
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    Set conn = New ADODB.Connection
    
    sql = "SELECT * FROM employee"
    
    Call Utilities.OpenEmployeeDB(rs, conn, sql)
    
    'If there are no records add a new one!
    If rs.RecordCount < 1 Then
        rs.AddNew
    End If
    
    booNewRec = False
    
    Call AddNodes
    Call SetDataSource
    Call BindData

End Sub

Private Sub SetDataSource()
    Set txtName.DataSource = rs
    Set txtAddress.DataSource = rs
    Set txtLocation.DataSource = rs
    Set txtPhone.DataSource = rs
    Set txtPager.DataSource = rs
    Set txtMobile.DataSource = rs
    Set txtEmail.DataSource = rs
    Set txtWorkPhone.DataSource = rs
    Set txtExtension.DataSource = rs
    Set cboDepartment.DataSource = rs
    Set cboTitle.DataSource = rs
    Set cboStatus.DataSource = rs
    Set txtNotes.DataSource = rs
    Set txtEmergecyContact.DataSource = rs
    Set txtEmergencyPhone.DataSource = rs
    Set txtBday.DataSource = rs
End Sub

Private Sub BindData()
    ' Bind fields to related textboxes
    txtName.DataField = "name"
    txtAddress.DataField = "address"
    txtLocation.DataField = "location"
    txtWorkPhone.DataField = "work_phone"
    txtPhone.DataField = "home_phone"
    txtPager.DataField = "pager"
    txtMobile.DataField = "mobile_phone"
    txtEmail.DataField = "home_email"
    txtExtension.DataField = "extension"
    cboDepartment.DataField = "department"
    cboTitle.DataField = "title"
    cboStatus.DataField = "status"
    txtNotes.DataField = "notes"
    txtEmergecyContact.DataField = "emergency_contact"
    txtEmergencyPhone.DataField = "emergency_phone"
    txtBday.DataField = "birthday"
End Sub

Private Sub AddNodes()
    Dim i As Long
    With tvwAM.Nodes
        .Clear
        For i = 1 To rs.RecordCount
            Set mNode = .Add(, , , rs.Fields!Name, 2)
            rs.MoveNext
        Next i
    End With
    rs.MoveFirst
End Sub


Private Sub Form_Resize()
    tvwAM.Height = Height - 1500
End Sub


Private Sub tvwAM_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Dim x As Integer
        x = MsgBox("Are you sure that you want to delete employee: (" & rs.Fields("name") & ") from the database?", _
               vbQuestion + vbYesNo, "Delete a Contact!")
        If x = vbYes Then
            rs.Delete
            MsgBox "Record has been deleted!", vbOKOnly + vbInformation, "Delete Successfull!"
            rs.Requery
            Call AddNodes
        End If
    End If
End Sub

Private Sub tvwAM_NodeClick(ByVal Node As MSComctlLib.Node)
    If booNewRec = False Then
        Dim i As Long
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            If rs.Fields!Name = Node Then
                i = rs.RecordCount
                Call BindData
            Else
                rs.MoveNext
            End If
        Next i
    Else
        MsgBox "You must finish adding the current record!  This record must be saved before another one is selected!  To cancel exit current form...", vbOKOnly + vbExclamation, _
                ":========== OC Contact Error =========:"
    End If
End Sub
