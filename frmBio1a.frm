VERSION 5.00
Begin VB.Form frmBio1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBio1a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7883
      TabIndex        =   19
      ToolTipText     =   "Click to Exit the Program"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtReportYear 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtReportDay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtReportMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdReportDate 
      Caption         =   "Change Report Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Click to Change Report Date"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ListBox lstName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Click on a Name to Generate BioRhythm Chart"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6347
      TabIndex        =   12
      ToolTipText     =   "Click to Open the Helps"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4811
      TabIndex        =   11
      ToolTipText     =   "Click to Print the Current Chart"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdJournal 
      Caption         =   "Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3275
      TabIndex        =   10
      ToolTipText     =   "Click to Launch the Journal"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdInterpretation 
      Caption         =   "Interpret"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1739
      TabIndex        =   9
      ToolTipText     =   "Click to Interpret the Current Chart"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompatibility 
      Caption         =   "Compatibility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   203
      TabIndex        =   8
      ToolTipText     =   "Click to Start BioCompatibility Calculator"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox picBio 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2760
      ScaleHeight     =   240.677
      ScaleMode       =   0  'User
      ScaleWidth      =   37.109
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin VB.CommandButton cmdChartTitle 
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCalc 
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   18
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   17
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   16
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   15
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   14
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   13
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   12
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   11
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   10
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   9
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   8
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   7
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   6
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   5
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   4
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   3
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   2
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   0
         X2              =   1.41
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnPip 
         BorderColor     =   &H8000000E&
         Index           =   0
         X1              =   1.41
         X2              =   2.821
         Y1              =   9.04
         Y2              =   9.04
      End
      Begin VB.Line lnVert 
         BorderColor     =   &H80000005&
         DrawMode        =   5  'Not Copy Pen
         X1              =   10.577
         X2              =   10.577
         Y1              =   0
         Y2              =   153.672
      End
      Begin VB.Line lnHor 
         BorderColor     =   &H80000005&
         DrawMode        =   5  'Not Copy Pen
         X1              =   0
         X2              =   23.27
         Y1              =   72.316
         Y2              =   72.316
      End
   End
   Begin VB.OptionButton optEsoteric 
      Caption         =   "Extended"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Selects Extended Chart"
      Top             =   200
      Width           =   975
   End
   Begin VB.OptionButton optNormal 
      Alignment       =   1  'Right Justify
      Caption         =   "Traditional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Selects Traditional Chart"
      Top             =   200
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit BioData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Click to Add, Delete or Edit BioData"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtBirthYear 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtBirthDay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtBirthMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblChartOptions 
      Alignment       =   2  'Center
      Caption         =   "Chart Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   200
      Width           =   6375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Physical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   3840
      Width           =   1580
   End
   Begin VB.Label lblMagenta 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Emotional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   4275
      TabIndex        =   24
      Top             =   3840
      Width           =   1580
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Intellectual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5805
      TabIndex        =   23
      Top             =   3840
      Width           =   1780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BioData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Report Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3080
      Width           =   2295
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   4320
      Y2              =   4950
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   0
      X2              =   9240
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000014&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   4320
      Y2              =   4940
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   9240
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblCyan 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Intuitional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   3840
      Width           =   1580
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   75
      Y2              =   4200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2640
      X2              =   9240
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   2640
      X2              =   9240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   3000
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   0
      X2              =   2520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   2520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   0
      X2              =   2520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   75
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   75
      Y2              =   2870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   2520
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000014&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   3000
      Y2              =   4190
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   4190
   End
End
Attribute VB_Name = "frmBio1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Set Scale Width and Height of picturebox
    
    picBio.ScaleWidth = 30
    picBio.ScaleHeight = 200
    
    'Draw Horizontal line in picturebox
    
    lnHor.Y1 = picBio.ScaleHeight / 2
    lnHor.Y2 = lnHor.Y1
    lnHor.X1 = 0
    lnHor.X2 = picBio.ScaleWidth
    
    'Draw Vertical line in picturebox
    
    lnVert.Y1 = 0
    lnVert.Y2 = picBio.ScaleHeight
    lnVert.X1 = picBio.ScaleWidth / 2
    lnVert.X2 = lnVert.X1
        
 'Draw Graticule in picturebox
    
    Dim count As Integer
    Dim increment As Integer
    
    increment = 10
    
    For count = 0 To 9
        lnPip(count).Y1 = (picBio.ScaleHeight / 2) + increment
        lnPip(count).Y2 = lnPip(count).Y1
        lnPip(count).X1 = (picBio.ScaleWidth / 2) - 0.25
        lnPip(count).X2 = (picBio.ScaleWidth / 2) + 0.4
        increment = increment + 10
    Next count
    
    increment = 10
    
    For count = 10 To 18
        lnPip(count).Y1 = (picBio.ScaleHeight / 2) - increment
        lnPip(count).Y2 = lnPip(count).Y1
        lnPip(count).X1 = (picBio.ScaleWidth / 2) - 0.25
        lnPip(count).X2 = (picBio.ScaleWidth / 2) + 0.4
        increment = increment + 10
    Next count
        
    'Set Report Date to current date
    'And convert month numerals to names
    
    Dim rm As Integer
        rm = DatePart("m", Date)
        With txtReportMonth
            If rm = 1 Then
                .Text = "January"
            End If
            If rm = 2 Then
                .Text = "February"
            End If
            If rm = 3 Then
                .Text = "March"
            End If
            If rm = 4 Then
                .Text = "April"
            End If
            If rm = 5 Then
                .Text = "May"
            End If
            If rm = 6 Then
                .Text = "June"
            End If
            If rm = 7 Then
                .Text = "July"
            End If
            If rm = 8 Then
                .Text = "August"
            End If
            If rm = 9 Then
                .Text = "September"
            End If
            If rm = 10 Then
                .Text = "October"
            End If
            If rm = 11 Then
                .Text = "November"
            End If
            If rm = 12 Then
                .Text = "December"
            End If
        End With
    
    txtReportDay.Text = DatePart("d", Date)
    txtReportYear.Text = DatePart("yyyy", Date)
    
    'Load the default database
    
    mdbFile = App.Path & "\family.mdb"
    LoadDB
    
    'Set the default chart option
    
    optNormal.Value = True
    optEsoteric.Value = False
    
    'Set the form title
    
    cmdChartTitle_Click
    
    'Display the chart for the first name in the list
    
    calc
    
End Sub

Private Sub cmdExit_Click()

    frmBio1_Unload
        
End Sub

Private Sub frmBio1_Unload()

    'Dump the form
    
    Unload Me
    
End Sub

Private Sub cmdChartTitle_Click()

    'Make title for the form
    
    frmBio1.Caption = " BioChart for " & lstName.Text & " on " & ReportDate

End Sub

Private Sub lstName_Click()

    'Create the database object
    
    Set mDB = OpenDatabase(mdbFile)
        With mDB
        Set mRS = .OpenRecordset("Names")
            With mRS
            
                'Get the first record
                
                If .RecordCount <> 0 Then
                .MoveFirst
                    If !Name = lstName.Text Then
                        txtBirthMonth.Text = !BirthMonth
                        convertMonth
                        txtBirthDay.Text = !BirthDay
                        txtBirthYear.Text = !BirthYear
                    Else
                    
                    'Add the rest of the records
                    
                    For dby = 1 To .RecordCount - 1
                        .MoveNext
                        If .EOF Then Exit For
                        If !Name = lstName.Text Then
                            txtBirthMonth.Text = !BirthMonth
                            convertMonth
                            txtBirthDay.Text = !BirthDay
                            txtBirthYear.Text = !BirthYear
                            Exit For
                        End If
                    Next dby
                End If
            End If
        End With
        .Close              'Close the database for safety
    End With
    
    'Update the title bar
    
    cmdChartTitle_Click
    
    'Plot the curves
    
    calc

End Sub

Private Sub cmdEdit_Click()
    
    'Launch the database editing dialog
    
    frmBioData.Show

End Sub

Private Sub cmdReportDate_Click()

    'Open the ReportDate dialog
    
    frmReportDate.Show
    
    'Update the title bar
    
    cmdChartTitle_Click
    
End Sub

Private Sub cmdCompatibility_Click()

    'Launch the compatibility calculator
    
    frmCompatibility.Show
    
End Sub

Private Sub cmdInterpretation_Click()

    frmBioHelp.Show
    frmBioHelp.rtfBioHelp.FileName = App.Path & "\interp1.rtf"
    
End Sub

Private Sub cmdJournal_Click()

    'Shell to Microsoft Wordpad inserting date and time
    'May be lazy but why re-invent the wheel?
    
    Dim journal As String
    journal = Shell("C:\Program Files\Accessories\WORDPAD.exe", 1)
    SendKeys "Journal entry for " & Date & " at " & Time & vbCrLf

End Sub

Private Sub cmdPrint_Click()
    
    'Yeah, I know it's cheesy but I haven't written
    'the report generator yet.
    
    MsgBox ("Report Generator will be added later.")
    
    'Or just use frmBio1.PrintForm
    
End Sub

Private Sub cmdHelp_Click()
    
    'Start the BioHelp module and point to the help file
    
    frmBioHelp.Show
    frmBioHelp.rtfBioHelp.FileName = App.Path & "\help1.rtf"
    
End Sub

Private Sub optNormal_Click()
    
    'Plot the curves for Traditional rhythms
    
    calc
    
End Sub

Private Sub optEsoteric_Click()
    
    'Plot the curves for Extended rhythms
    
    calc
    
End Sub
    
Private Sub cmdCalc_Click()
    
    'An invisible button used for changing the report date
    
    calc
    
End Sub

    
Private Function calc()

    'Declare a few variables used in this function
    
    Dim currentPoint As Double
    Dim bioColor As ColorConstants
    
    'Clear the screen
    
    picBio.Cls
    
    'Get the complete Birth and Report dates
    
    BirthDate = txtBirthMonth.Text & " " & txtBirthDay.Text & ", " & txtBirthYear.Text
    ReportDate = txtReportMonth.Text & " " & txtReportDay.Text & ", " & txtReportYear.Text
    
    'Do not calculate if Birth Date and Report Date not complete
    
    If txtBirthMonth.Text = "" Or txtBirthDay.Text = "" Or txtBirthYear.Text = "" Then
        Exit Function
    End If
    
    If txtReportMonth.Text = "" Or txtReportDay.Text = "" Or txtReportYear.Text = "" Then
        Exit Function
    End If
    
    'Loop through the cycles and plot them
    
    For bioNum = 1 To 4
    
        'Physical period
        If bioNum = 1 And optNormal.Value = True Then bioPeriod = 23: bioColor = vbRed
        'Aesthetic period
        If bioNum = 1 And optEsoteric.Value = True Then bioPeriod = 43: bioColor = vbRed
        'Emotions period
        If bioNum = 2 And optNormal.Value = True Then bioPeriod = 28: bioColor = vbMagenta
        'Compassion period
        If bioNum = 2 And optEsoteric.Value = True Then bioPeriod = 38: bioColor = vbMagenta
        'Intellect period
        If bioNum = 3 And optNormal.Value = True Then bioPeriod = 33: bioColor = vbGreen
        'Awareness period
        If bioNum = 3 And optEsoteric.Value = True Then bioPeriod = 48: bioColor = vbGreen
        'Intuition period
        If bioNum = 4 And optNormal.Value = True Then bioPeriod = 38: bioColor = vbCyan
        'Spiritual period
        If bioNum = 4 And optEsoteric.Value = True Then bioPeriod = 53: bioColor = vbCyan
        
        'Find first peak to left of middle
        
        currentPoint = picBio.ScaleWidth / 2 - ((DateDiff("d", BirthDate, ReportDate) Mod bioPeriod) - bioPeriod / 4)
        
        'Find first peak that is off the chart
        
        Do While currentPoint > 0
            currentPoint = currentPoint - bioPeriod
        Loop
        
        'Necessary because next loop adds bioPeriod/2 back to the variable
        
        currentPoint = currentPoint - bioPeriod / 2
        
        'Find high and low points then plot parabolas
        
        Do While currentPoint < picBio.ScaleWidth
            currentPoint = currentPoint + bioPeriod / 2
            If currentPoint + bioPeriod / 4 >= 0 Then
                parabola bioNum, currentPoint, 0, currentPoint + bioPeriod / 4, bioColor
            End If
                currentPoint = currentPoint + bioPeriod / 2
            If currentPoint + bioPeriod / 4 >= 0 Then
                parabola bioNum, currentPoint, picBio.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
            End If
        Loop
        
    Next bioNum
    
End Function

Public Function parabola(bioNum As Integer, Xa As Double, Ya As Double, lastPt As Double, RedGreenBlueCyan As ColorConstants)

    'Creates a parabola when vertex and last point are given
    'Vertex = Xa, Ya
    'Last point = lastPt, horCenter
    
    'Declare the variables it needs
    
    Dim horCenter As Integer
    Dim slope As Double
    Dim Y As Double, X As Double
    
    'Set horCenter to the horizontal center of the picturebox
    
    horCenter = picBio.ScaleHeight / 2
    
    'Find slope of parabola
    
    slope = (horCenter - Ya) / ((lastPt - Xa) ^ 2)
    
    'Plot the parabola
    
    For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
        Y = slope * ((X - Xa) ^ 2) + Ya
        
            'Find the value of the curve at midpoint
            'And convert it to a percentage of scaleheight
            'With positive number on top and negative below
            
            If X = picBio.ScaleWidth / 2 Then
                bioVal(bioNum) = Y
                If bioVal(bioNum) <= 100 Then
                    bioVal(bioNum) = 100 - bioVal(bioNum)
                End If
                If bioVal(bioNum) > 100 Then
                    bioVal(bioNum) = (bioVal(bioNum) - 100) * -1
                End If
            End If
            
        picBio.PSet (X, Y), RedGreenBlueCyan
    Next X
    
    'Update the labels
    
    If optNormal.Value = True Then
        lblRed.Caption = "Physical " & Str(bioVal(1)) & "%"
    End If
    If optEsoteric.Value = True Then
        lblRed.Caption = "Aesthetic " & Str(bioVal(1)) & "%"
    End If
    If optNormal.Value = True Then
        lblMagenta.Caption = "Emotions " & Str(bioVal(2)) & "%"
    End If
    If optEsoteric.Value = True Then
        lblMagenta.Caption = "Compassion " & Str(bioVal(2)) & "%"
    End If
    If optNormal.Value = True Then
        lblGreen.Caption = "Intellect " & Str(bioVal(3)) & "%"
    End If
    If optEsoteric.Value = True Then
        lblGreen.Caption = "Awareness " & Str(bioVal(3)) & "%"
    End If
    If optNormal.Value = True Then
        lblCyan.Caption = "Intuition " & Str(bioVal(4)) & "%"
    End If
    If optEsoteric.Value = True Then
        lblCyan.Caption = "Spiritual " & Str(bioVal(4)) & "%"
    End If
    
End Function

Public Sub LoadDB()

    'Declare the database object
    
    Set mDB = OpenDatabase(mdbFile)
    
    With mDB
        Set mRS = .OpenRecordset("Names")
        With mRS
        
            'Move to the first record
            
            If .RecordCount <> 0 Then
                .MoveFirst
                
                'Do not display the Safety record
                
                If !Name <> "<Safety>" Then
                    lstName.AddItem !Name
                End If
                
                If !Name <> "<Safety>" Then
                txtBirthMonth.Text = !BirthMonth
                convertMonth
                txtBirthDay.Text = !BirthDay
                txtBirthYear = !BirthYear
                End If
                
                'Add the rest of the records
                
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name <> "<Safety>" Then
                        lstName.AddItem !Name
                    End If
                    txtBirthMonth.Text = !BirthMonth
                    convertMonth
                    txtBirthDay.Text = !BirthDay
                    txtBirthYear.Text = !BirthYear
                Next dby
            End If
        End With
        .Close          'Close the database to keep it safe
    End With
    
    'Set the default to the first name in the list
    
    If lstName.ListCount > 0 Then lstName.ListIndex = 0
     
End Sub

Private Sub convertMonth()

    'Convert months from numeral to name

    If txtBirthMonth.Text = "1" Then
        txtBirthMonth.Text = "January"
    End If
    If txtBirthMonth.Text = "2" Then
        txtBirthMonth.Text = "February"
    End If
    If txtBirthMonth.Text = "3" Then
        txtBirthMonth.Text = "March"
    End If
    If txtBirthMonth.Text = "4" Then
        txtBirthMonth.Text = "April"
    End If
    If txtBirthMonth.Text = "5" Then
        txtBirthMonth.Text = "May"
    End If
    If txtBirthMonth.Text = "6" Then
        txtBirthMonth.Text = "June"
    End If
    If txtBirthMonth.Text = "7" Then
        txtBirthMonth.Text = "July"
    End If
    If txtBirthMonth.Text = "8" Then
        txtBirthMonth.Text = "August"
    End If
    If txtBirthMonth.Text = "9" Then
        txtBirthMonth.Text = "September"
    End If
    If txtBirthMonth.Text = "10" Then
        txtBirthMonth.Text = "October"
    End If
    If txtBirthMonth.Text = "11" Then
        txtBirthMonth.Text = "November"
    End If
    If txtBirthMonth.Text = "12" Then
        txtBirthMonth.Text = "December"
    End If
    
End Sub
