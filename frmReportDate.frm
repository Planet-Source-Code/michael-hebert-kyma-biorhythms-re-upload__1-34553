VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmReportDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Report Date Selector "
   ClientHeight    =   3690
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   Icon            =   "frmReportDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin MSACAL.Calendar calReportDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4471
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   4
      Day             =   7
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
      ValueIsNull     =   0   'False
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   4800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   4800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Select a new Report Date"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmReportDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    'Set the calendar to system date
    
    calReportDate.Value = Date
    
End Sub

Private Sub frmReportDate_Unload()

    Unload Me
    
End Sub

Private Sub calReportDate_Click()

    'Get the month name
    
    frmBio1.txtReportMonth.Text = calReportDate.Month
    
    'Convert to long month name
    
    If frmBio1.txtReportMonth = "1" Then
        frmBio1.txtReportMonth.Text = "January"
    End If
    If frmBio1.txtReportMonth = "2" Then
        frmBio1.txtReportMonth.Text = "February"
    End If
    If frmBio1.txtReportMonth = "3" Then
        frmBio1.txtReportMonth.Text = "March"
    End If
    If frmBio1.txtReportMonth = "4" Then
        frmBio1.txtReportMonth.Text = "April"
    End If
    If frmBio1.txtReportMonth = "5" Then
        frmBio1.txtReportMonth.Text = "May"
    End If
    If frmBio1.txtReportMonth = "6" Then
        frmBio1.txtReportMonth.Text = "June"
    End If
    If frmBio1.txtReportMonth = "7" Then
        frmBio1.txtReportMonth.Text = "July"
    End If
    If frmBio1.txtReportMonth = "8" Then
        frmBio1.txtReportMonth.Text = "August"
    End If
    If frmBio1.txtReportMonth = "9" Then
        frmBio1.txtReportMonth.Text = "September"
    End If
    If frmBio1.txtReportMonth = "10" Then
        frmBio1.txtReportMonth.Text = "October"
    End If
    If frmBio1.txtReportMonth = "11" Then
        frmBio1.txtReportMonth.Text = "November"
    End If
    If frmBio1.txtReportMonth = "12" Then
        frmBio1.txtReportMonth.Text = "December"
    End If
    
    'Send Report Day and Year to main form
    
    frmBio1.txtReportDay.Text = calReportDate.Day
    frmBio1.txtReportYear.Text = calReportDate.Year

End Sub

Private Sub cmdSet_Click()

    'Click the invisible buttons
    
    frmBio1.cmdCalc.Value = True
    
    frmBio1.cmdChartTitle = True
    
    'And unload this form
    
    frmReportDate_Unload
    
End Sub
Private Sub cmdExit_Click()
    
    'Reset calendar control to system date
    
    calReportDate.Value = Date
    
    'And unload this form
    
    frmReportDate_Unload
    
End Sub

