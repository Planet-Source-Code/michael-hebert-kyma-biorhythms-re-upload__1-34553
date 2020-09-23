VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmBioData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  BioData Entry Form "
   ClientHeight    =   4800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4800
   Icon            =   "frmBioData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Click to Exit this dialog"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Click to Delete Entry"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      ToolTipText     =   "Click to Add Entry"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox lstName 
      Height          =   1035
      ItemData        =   "frmBioData.frx":030A
      Left            =   360
      List            =   "frmBioData.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Click on a Name to Select"
      Top             =   840
      Width           =   2655
   End
   Begin MSACAL.Calendar calBirthDate 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Set Month, Year and Day"
      Top             =   2280
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483626
      Year            =   2002
      Month           =   4
      Day             =   6
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Enter First Name Last Name"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   4680
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   4680
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblBirthDate 
      Caption         =   "Select a Birth Date from the Month, Year and Date Boxes"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label blbName 
      Alignment       =   2  'Center
      Caption         =   "Enter a Full Name or Select from List"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmBioData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim add As Boolean
Dim newName As String

Private Sub Form_Load()
    
    mdbFile = App.Path & "\family.mdb"
    LoadDB (mdbFile)
    txtName.Text = ""
    
End Sub

Private Sub frmBioData_Unload()

    Unload Me
    
End Sub

Private Sub lstName_Click()

Set mDB = OpenDatabase(mdbFile)
        With mDB
        Set mRS = .OpenRecordset("Names")
            With mRS
                If .RecordCount <> 0 Then
                .MoveFirst
                    If !Name = lstName.Text Then
                        txtName.Text = lstName.Text
                        calBirthDate.Month = !BirthMonth
                        calBirthDate.Day = !BirthDay
                        calBirthDate.Year = !BirthYear
                    Else
                    For dby = 1 To .RecordCount - 1
                        .MoveNext
                        If .EOF Then Exit For
                        If !Name = lstName.Text Then
                            txtName.Text = lstName.Text
                            calBirthDate.Month = !BirthMonth
                            calBirthDate.Day = !BirthDay
                            calBirthDate.Year = !BirthYear
                            Exit For
                        End If
                    Next dby
                End If
            End If
        End With
        .Close
    End With
    
End Sub

Private Sub txtName_Click()
    
    txtName.Text = ""
    newName = txtName.Text
        
End Sub

Private Sub cmdAdd_Click()

    'Open the database and update with new data
    
    Set mDB = OpenDatabase(mdbFile)
    With mDB
        Set mRS = .OpenRecordset("Names")
        With mRS
            If txtName.Text <> "" Then
                .AddNew
                !Name = txtName.Text
                !BirthMonth = BirthMonth
                !BirthDay = BirthDay
                !BirthYear = BirthYear
                .Update
            End If
        End With
        .Close
    End With
    
    'Refresh both listboxes
    
    lstName.AddItem txtName.Text
    frmBio1.lstName.AddItem txtName.Text
    
    'Cleanup textbox for next entry
    
    txtName.Text = ""
    txtName.SetFocus
    
End Sub

Private Sub cmdDelete_Click()
    
    'Open database and delete selected record
    
    Set mDB = OpenDatabase(mdbFile)
    With mDB
        Set mRS = .OpenRecordset("Names")
        With mRS
        Do Until .EOF
            If lstName.Text = "<Safety>" Then
                MsgBox ("<Safety> Record may not be deleted!")
                Exit Sub
            End If
            If lstName.Text = !Name Then
                If MsgBox("Really delete " & lstName.Text & "?", vbQuestion + vbYesNo, "Delete ") = vbYes Then
                    .Delete
                    lstName.RemoveItem (lstName.ListIndex)
                Else
                    Exit Sub
                End If
            End If
            .MoveNext
            Loop
        End With
        .Close
    End With
    
    'Clear textbox
    
    txtName.Text = ""
    
    'Refresh listbox on main form
    
    frmBio1.lstName.RemoveItem frmBio1.lstName.ListIndex
    frmBio1.lstName.Clear
    frmBio1.LoadDB
    
    txtName.SetFocus
    
End Sub

Public Sub LoadDB(mdbFile As String)

    'Open the database and fill the listbox with data
    
    Set mDB = OpenDatabase(mdbFile)
    With mDB
        Set mRS = .OpenRecordset("Names")
        With mRS
            If .RecordCount <> 0 Then
                .MoveFirst
                
                'Do not display the Safety record
                
                If !Name <> "<Safety>" Then
                    lstName.AddItem !Name
                End If
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name <> "<Safety>" Then
                        lstName.AddItem !Name
                    End If
                Next dby
            End If
        End With
        .Close
    End With
    
    'Set the list box count
    
    If lstName.ListCount > 0 Then lstName.ListIndex = 0
     
End Sub

Private Sub calBirthDate_Click()
    
    BirthMonth = calBirthDate.Month
    BirthDay = calBirthDate.Day
    BirthYear = calBirthDate.Year

End Sub

Public Sub cmdExit_Click()

    frmBioData_Unload
    
End Sub
