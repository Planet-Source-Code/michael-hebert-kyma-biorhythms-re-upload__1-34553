VERSION 5.00
Begin VB.Form frmCompatibility 
   Caption         =   "  Compatibility Chart"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7830
   Icon            =   "frmCompatability.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4808
      TabIndex        =   26
      ToolTipText     =   "Click for Help"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3308
      TabIndex        =   25
      ToolTipText     =   "Click to print report"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdJournal 
      Caption         =   "Journal"
      Height          =   375
      Left            =   1808
      TabIndex        =   24
      ToolTipText     =   "Click to launch journal"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdInterp 
      Caption         =   "Interpret"
      Height          =   375
      Left            =   308
      TabIndex        =   23
      ToolTipText     =   "Click for interpretation"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6308
      TabIndex        =   22
      ToolTipText     =   "Click to Exit"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox picComp4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   1095
      Left            =   5280
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
      Begin VB.Line lnHor4 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   2040
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.PictureBox picComp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   1095
      Left            =   5280
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   12
      Top             =   480
      Width           =   2295
      Begin VB.Line lnHor2 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   2160
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.PictureBox picComp3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   1095
      Left            =   2880
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
      Begin VB.Line lnHor3 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   2160
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.PictureBox picComp1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   1095
      Left            =   2880
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   480
      Width           =   2295
      Begin VB.Line lnHor1 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   2160
         Y1              =   92.754
         Y2              =   92.754
      End
   End
   Begin VB.TextBox txtBirthYear2 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtBirthDay2 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtBirthMonth2 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtBirthYear1 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtBirthDay1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtBirthMonth1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox lstName2 
      ForeColor       =   &H80000012&
      Height          =   1035
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Click to select second name"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ListBox lstName1 
      ForeColor       =   &H80000012&
      Height          =   1035
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Click to select first name"
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton optNormal 
      Alignment       =   1  'Right Justify
      Caption         =   "Traditional"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      ToolTipText     =   "Click for Traditional chart"
      Top             =   200
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optEsoteric 
      Caption         =   "Extended"
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      ToolTipText     =   "Click for Extended chart"
      Top             =   200
      Width           =   975
   End
   Begin VB.Label lblPerson2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Second Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   5280
      TabIndex        =   20
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblPerson1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "First Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   7680
      X2              =   7680
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Chart Type"
      Height          =   255
      Left            =   4500
      TabIndex        =   17
      Top             =   200
      Width           =   1335
   End
   Begin VB.Label lblPic4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "lblPic4"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblPic2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "lblPic2"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblPic3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "lblPic3"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblPic1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "lblPic1"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   7680
      X2              =   7680
      Y1              =   120
      Y2              =   3840
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   2760
      X2              =   7680
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2760
      X2              =   7680
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   120
      Y2              =   3840
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   2040
      Y2              =   3840
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   120
      X2              =   2640
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2040
      Y2              =   3840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   2640
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   2640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1920
   End
End
Attribute VB_Name = "frmCompatibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Set up the picture boxes
   
    picComp1.ScaleWidth = 40
    picComp1.ScaleHeight = 100
    lnHor1.Y1 = picComp1.ScaleHeight / 2
    lnHor1.Y2 = lnHor1.Y1
    lnHor1.X1 = 0
    lnHor1.X2 = picComp1.ScaleWidth
        
    picComp2.ScaleWidth = 40
    picComp2.ScaleHeight = 100
    lnHor2.Y1 = picComp2.ScaleHeight / 2
    lnHor2.Y2 = lnHor1.Y1
    lnHor2.X1 = 0
    lnHor2.X2 = picComp2.ScaleWidth
    
    picComp3.ScaleWidth = 40
    picComp3.ScaleHeight = 100
    lnHor3.Y1 = picComp3.ScaleHeight / 2
    lnHor3.Y2 = lnHor3.Y1
    lnHor3.X1 = 0
    lnHor3.X2 = picComp3.ScaleWidth
    
    picComp4.ScaleWidth = 40
    picComp4.ScaleHeight = 100
    lnHor4.Y1 = picComp4.ScaleHeight / 2
    lnHor4.Y2 = lnHor4.Y1
    lnHor4.X1 = 0
    lnHor4.X2 = picComp4.ScaleWidth
    
    'Set the chart labels for normal display (default option)
    
    lblPic1.Caption = "Physical"
    lblPic2.Caption = "Emotional"
    lblPic3.Caption = "Intellectual"
    lblPic4.Caption = "Intuitional"
    
    'Load the database
    
    mdbFile = App.Path & "\family.mdb"
    LoadDB
    
End Sub

Private Sub frmCompatibility_Unload()

    Unload Me
    
End Sub

Private Sub lstName1_Click()

    Set mDB = OpenDatabase(mdbFile)
        With mDB
        Set mRS = .OpenRecordset("Names")
            With mRS
                If .RecordCount <> 0 Then
                .MoveFirst
                    If !Name = lstName1.Text Then
                        txtBirthMonth1.Text = !BirthMonth
                        convertMonth
                        txtBirthDay1.Text = !BirthDay
                        txtBirthYear1.Text = !BirthYear
                    Else
                    For dby = 1 To .RecordCount - 1
                        .MoveNext
                        If .EOF Then Exit For
                        If !Name = lstName1.Text Then
                            txtBirthMonth1.Text = !BirthMonth
                            convertMonth
                            txtBirthDay1.Text = !BirthDay
                            txtBirthYear1.Text = !BirthYear
                            Exit For
                        End If
                    Next dby
                End If
            End If
        End With
        .Close
    End With
    lblPerson1.Caption = lstName1.Text
    
    chartTitle
    
    calc

End Sub

Private Sub lstName2_Click()

    Set mDB = OpenDatabase(mdbFile)
        With mDB
        Set mRS = .OpenRecordset("Names")
            With mRS
                If .RecordCount <> 0 Then
                .MoveFirst
                    If !Name = lstName2.Text Then
                        txtBirthMonth2.Text = !BirthMonth
                        convertMonth
                        txtBirthDay2.Text = !BirthDay
                        txtBirthYear2.Text = !BirthYear
                    Else
                    For dby = 1 To .RecordCount - 1
                        .MoveNext
                        If .EOF Then Exit For
                        If !Name = lstName2.Text Then
                            txtBirthMonth2.Text = !BirthMonth
                            convertMonth
                            txtBirthDay2.Text = !BirthDay
                            txtBirthYear2.Text = !BirthYear
                            Exit For
                        End If
                    Next dby
                End If
            End If
        End With
        .Close
    End With
    lblPerson2.Caption = lstName2.Text
    
    chartTitle
    
    calc

End Sub

Private Sub chartTitle()

    frmCompatibility.Caption = " BioCompatibility of " & lstName1.Text & " and " & lstName2.Text

End Sub

Private Sub optNormal_Click()

    optNormal.Value = True
    optEsoteric.Value = False
    lblPic1.Caption = "Physical"
    picComp1.ScaleWidth = 40
    lblPic2.Caption = "Emotional"
    picComp2.ScaleWidth = 40
    lblPic3.Caption = "Intellectual"
    picComp3.ScaleWidth = 40
    lblPic4.Caption = "Intuitional"
    picComp4.ScaleWidth = 40

    calc
    
End Sub

Private Sub optEsoteric_Click()

    optEsoteric.Value = True
    optNormal.Value = False
    lblPic1.Caption = "Aesthetic"
    picComp1.ScaleWidth = 60
    lblPic2.Caption = "Compassion"
    picComp2.ScaleWidth = 60
    lblPic3.Caption = "Awareness"
    picComp3.ScaleWidth = 60
    lblPic4.Caption = "Spiritual"
    picComp4.ScaleWidth = 60
    
    calc
    
End Sub
Private Sub cmdExit_Click()

    frmCompatibility_Unload
    
End Sub

Private Sub cmdHelp_Click()

    frmBioHelp.Show
    frmBioHelp.rtfBioHelp.FileName = App.Path & "\help2.rtf"
    
End Sub

Private Sub cmdInterp_Click()

    frmBioHelp.Show
    frmBioHelp.rtfBioHelp.FileName = App.Path & "\interp2.rtf"
    
End Sub
Private Sub cmdJournal_Click()

    'Shell to Microsoft Wordpad inserting date and time
    
    Dim journal As String
    journal = Shell("C:\Program Files\Accessories\WORDPAD.exe", 1)
    SendKeys "Journal entry for " & Date & " at " & Time & vbCrLf
    
End Sub

Private Sub cmdPrint_Click()

    MsgBox ("Report Generator will be added later.")
    'frmCompatibility.PrintForm
    
End Sub

Private Sub calc()
    
    Dim person As Integer
    Dim currentPoint As Double
    Dim bioPeriod As Integer
    Dim bioColor As ColorConstants
    Dim bioNum As Integer
    
    'Clear the screens
    
    picComp1.Cls
    picComp2.Cls
    picComp3.Cls
    picComp4.Cls
   
    'Get the complete Birth dates
    
    BirthDate1 = txtBirthMonth1.Text & " " & txtBirthDay1.Text & ", " & txtBirthYear1.Text
    BirthDate2 = txtBirthMonth2.Text & " " & txtBirthDay2.Text & ", " & txtBirthYear2.Text
    
    'Find out which birthdate is earliest
    
    If DateValue(BirthDate1) < DateValue(BirthDate2) Then
        tempDate1 = BirthDate1
        tempDate2 = BirthDate2
    Else
        tempDate1 = BirthDate2
        tempDate2 = BirthDate1
    End If
    
    'Set colors for each person
    
    For person = 1 To 2
        If person = 1 Then
            bioColor = vbRed
        Else
            bioColor = vbGreen
        End If
        
    'Loop through the cycles and plot them
    
    For bioNum = 1 To 4
        'Physical period
        If bioNum = 1 And optNormal.Value = True Then bioPeriod = 23
        'Aesthetic period
        If bioNum = 1 And optEsoteric.Value = True Then bioPeriod = 43
        'Emotional period
        If bioNum = 2 And optNormal.Value = True Then bioPeriod = 28
        'Compassion period
        If bioNum = 2 And optEsoteric.Value = True Then bioPeriod = 38
        'Intellectual period
        If bioNum = 3 And optNormal.Value = True Then bioPeriod = 33
        'Awareness period
        If bioNum = 3 And optEsoteric.Value = True Then bioPeriod = 48
        'Intuitional period
        If bioNum = 4 And optNormal.Value = True Then bioPeriod = 38
        'Spiritual period
        If bioNum = 4 And optEsoteric.Value = True Then bioPeriod = 53
        
        'Find first peak to left of middle
        If person = 1 And bioNum = 1 Then
            currentPoint = picComp1.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate1) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 2 And bioNum = 1 Then
            currentPoint = picComp1.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate2) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 1 And bioNum = 2 Then
            currentPoint = picComp2.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate1) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 2 And bioNum = 2 Then
            currentPoint = picComp2.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate2) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 1 And bioNum = 3 Then
            currentPoint = picComp3.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate1) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 2 And bioNum = 3 Then
            currentPoint = picComp3.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate2) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 1 And bioNum = 4 Then
            currentPoint = picComp4.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate1) Mod bioPeriod) - bioPeriod / 4)
        End If
        If person = 2 And bioNum = 4 Then
            currentPoint = picComp4.ScaleWidth / 2 - ((DateDiff("d", tempDate1, tempDate2) Mod bioPeriod) - bioPeriod / 4)
        End If

        'Find first peak that is off the chart
        Do While currentPoint > 0
            currentPoint = currentPoint - bioPeriod
        Loop
        
        'Necessary because next loop adds bioPeriod/2 back to the variable
        currentPoint = currentPoint - bioPeriod / 2
        
        'Find high and low points and plot parabolas
        Do While currentPoint < picComp1.ScaleWidth
            currentPoint = currentPoint + bioPeriod / 2
            If currentPoint + bioPeriod / 4 >= 0 Then
                parabola person, bioNum, currentPoint, 0, currentPoint + bioPeriod / 4, bioColor
            End If
                currentPoint = currentPoint + bioPeriod / 2
            If currentPoint + bioPeriod / 4 >= 0 Then
                If bioNum = 1 Then
                    parabola person, bioNum, currentPoint, picComp1.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
                End If
                If bioNum = 2 Then
                    parabola person, bioNum, currentPoint, picComp2.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
                End If
                If bioNum = 3 Then
                    parabola person, bioNum, currentPoint, picComp3.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
                End If
                If bioNum = 4 Then
                    parabola person, bioNum, currentPoint, picComp4.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
                End If
            End If
        Loop
    
    Next bioNum
    
    Next person
    
    stats
    
End Sub

Public Function parabola(person As Integer, bioNum As Integer, Xa As Double, Ya As Double, lastPt As Double, RedGreen As ColorConstants)

    'Creates a parabola when vertex and last point are given
    'Vertex = Xa, Ya
    'Last point = lastPt, horCenter
    
    Dim horCenter1 As Integer
    Dim horCenter2 As Integer
    Dim horCenter3 As Integer
    Dim horCenter4 As Integer
    
    Dim slope As Double
    Dim Y As Double, X As Double
    
    If bioNum = 1 Then
        'Set horCenter to the horizontal center of the picturebox
        horCenter1 = picComp1.ScaleHeight / 2
    
        'Find slope of parabola
        slope = (horCenter1 - Ya) / ((lastPt - Xa) ^ 2)
    
        'Graph the parabola
        For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
            Y = slope * ((X - Xa) ^ 2) + Ya
            picComp1.PSet (X, Y), RedGreen
        Next X
    End If
    
    If bioNum = 2 Then
        horCenter2 = picComp2.ScaleHeight / 2
        slope = (horCenter2 - Ya) / ((lastPt - Xa) ^ 2)
        For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
            Y = slope * ((X - Xa) ^ 2) + Ya
            picComp2.PSet (X, Y), RedGreen
        Next X
    End If
    
    If bioNum = 3 Then
        horCenter3 = picComp1.ScaleHeight / 2
        slope = (horCenter3 - Ya) / ((lastPt - Xa) ^ 2)
        For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
            Y = slope * ((X - Xa) ^ 2) + Ya
            picComp3.PSet (X, Y), RedGreen
        Next X
    End If
    
    If bioNum = 4 Then
        horCenter4 = picComp1.ScaleHeight / 2
        slope = (horCenter4 - Ya) / ((lastPt - Xa) ^ 2)
        For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
            Y = slope * ((X - Xa) ^ 2) + Ya
            picComp4.PSet (X, Y), RedGreen
        Next X
    End If

End Function

Private Sub stats()

    'Provides percentage value equal to the number of days
    'in each cycle that are common to both people
    
    Dim bioNum As Integer
    Dim bioPeriod As Integer
    Dim numCycles As Single
    Dim daysLived As Long
    Dim daysInto As Single
    Dim diff As Single
    
    For bioNum = 1 To 4
    
        'Loop through all four biocycles
        
        'Physical period
        If bioNum = 1 And optNormal.Value = True Then bioPeriod = 23
        'Aesthetic period
        If bioNum = 1 And optEsoteric.Value = True Then bioPeriod = 43
        'Emotional period
        If bioNum = 2 And optNormal.Value = True Then bioPeriod = 28
        'Compassion period
        If bioNum = 2 And optEsoteric.Value = True Then bioPeriod = 38
        'Intellectual period
        If bioNum = 3 And optNormal.Value = True Then bioPeriod = 33
        'Awareness period
        If bioNum = 3 And optEsoteric.Value = True Then bioPeriod = 48
        'Intuitional period
        If bioNum = 4 And optNormal.Value = True Then bioPeriod = 38
        'Spiritual period
        If bioNum = 4 And optEsoteric.Value = True Then bioPeriod = 53
        
        'Calculate the number of Days Into the current cycle
        
        daysLived = DateDiff("d", tempDate1, tempDate2)
        numCycles = (daysLived / bioPeriod)
        daysInto = Round((numCycles - Int(numCycles)) * bioPeriod, 2)
        
        If daysInto = 0 Or daysInto = bioPeriod Then
            diff = 100
        End If
        
        If daysInto > 0 And daysInto < bioPeriod * 0.5 Then
            diff = 100 - ((100 / (bioPeriod / 2)) * daysInto) '(dayPct * daysInto)
            diff = Round(diff)
        End If
        
        If daysInto = bioPeriod * 0.5 Then
            diff = 0
        End If
        
        If daysInto > bioPeriod * 0.5 And daysInto < bioPeriod Then
            diff = 100 - (100 - ((100 / (bioPeriod / 2)) * daysInto)) ' * 100
            diff = Abs(Round(100 - diff))
        End If
        
        If bioNum = 1 And optNormal.Value = True Then
            lblPic1.Caption = "Physical " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 1 And optEsoteric.Value = True Then
            lblPic1.Caption = "Aesthetic " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 2 And optNormal.Value = True Then
            lblPic2.Caption = "Emotional " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 2 And optEsoteric.Value = True Then
            lblPic2.Caption = "Compassion " & bioPeriod & " days " & diff & "%"
        End If
        
         If bioNum = 3 And optNormal.Value = True Then
            lblPic3.Caption = "Intellect " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 3 And optEsoteric.Value = True Then
            lblPic3.Caption = "Awareness " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 4 And optNormal.Value = True Then
            lblPic4.Caption = "Intuition " & bioPeriod & " days " & diff & "%"
        End If
        
        If bioNum = 4 And optEsoteric.Value = True Then
            lblPic4.Caption = "Spiritual " & bioPeriod & " days " & diff & "%"
        End If
       
    Next bioNum
        
End Sub

Public Sub LoadDB()

    Set mDB = OpenDatabase(mdbFile)
    With mDB
        Set mRS = .OpenRecordset("Names")
        With mRS
            If .RecordCount <> 0 Then
                .MoveFirst
                
                'Do not display the Safety record
                
                If !Name <> "<Safety>" Then
                    lstName1.AddItem !Name
                End If
                If !Name <> "<Safety>" Then
                    lstName2.AddItem !Name
                End If
                txtBirthMonth1.Text = !BirthMonth
                txtBirthMonth2.Text = !BirthMonth
                txtBirthDay1.Text = !BirthDay
                txtBirthDay2.Text = !BirthDay
                txtBirthYear1.Text = !BirthYear
                txtBirthYear2.Text = !BirthYear
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name <> "<Safety>" Then
                        lstName1.AddItem !Name
                    End If
                    If !Name <> "<Safety>" Then
                        lstName2.AddItem !Name
                    End If
                    txtBirthMonth1.Text = !BirthMonth
                    txtBirthMonth2.Text = !BirthMonth
                    txtBirthDay1.Text = !BirthDay
                    txtBirthDay2.Text = !BirthDay
                    txtBirthYear1.Text = !BirthYear
                    txtBirthYear2.Text = !BirthYear
                Next dby
            End If
        End With
        .Close
    End With
    
    'Preset lstName1 listbox to display first name in list
    
    If lstName1.ListCount > 0 Then lstName1.ListIndex = 0
    
    'Preset lstName2 listbox to display second name in list
    
    If lstName2.ListCount > 0 Then lstName2.ListIndex = 1
    
    convertMonth
    
End Sub

Private Sub convertMonth()
    
    If txtBirthMonth1.Text = "1" Then
        txtBirthMonth1.Text = "January"
    End If
    If txtBirthMonth2.Text = "1" Then
        txtBirthMonth2.Text = "January"
    End If
    If txtBirthMonth1.Text = "2" Then
        txtBirthMonth1.Text = "February"
    End If
    If txtBirthMonth2.Text = "2" Then
        txtBirthMonth2.Text = "February"
    End If
    If txtBirthMonth1.Text = "3" Then
        txtBirthMonth1.Text = "March"
    End If
    If txtBirthMonth2.Text = "3" Then
        txtBirthMonth2.Text = "March"
    End If
    If txtBirthMonth1.Text = "4" Then
        txtBirthMonth1.Text = "April"
    End If
    If txtBirthMonth2.Text = "4" Then
        txtBirthMonth2.Text = "April"
    End If
    If txtBirthMonth1.Text = "5" Then
        txtBirthMonth1.Text = "May"
    End If
    If txtBirthMonth2.Text = "5" Then
        txtBirthMonth2.Text = "May"
    End If
    If txtBirthMonth1.Text = "6" Then
        txtBirthMonth1.Text = "June"
    End If
    If txtBirthMonth2.Text = "6" Then
        txtBirthMonth2.Text = "June"
    End If
    If txtBirthMonth1.Text = "7" Then
        txtBirthMonth1.Text = "July"
    End If
    If txtBirthMonth2.Text = "7" Then
        txtBirthMonth2.Text = "July"
    End If
    If txtBirthMonth1.Text = "8" Then
        txtBirthMonth1.Text = "August"
    End If
    If txtBirthMonth2.Text = "8" Then
        txtBirthMonth2.Text = "August"
    End If
    If txtBirthMonth1.Text = "9" Then
        txtBirthMonth1.Text = "September"
    End If
    If txtBirthMonth2.Text = "9" Then
        txtBirthMonth2.Text = "September"
    End If
    If txtBirthMonth1.Text = "10" Then
        txtBirthMonth1.Text = "October"
    End If
    If txtBirthMonth2.Text = "10" Then
        txtBirthMonth2.Text = "October"
    End If
    If txtBirthMonth1.Text = "11" Then
        txtBirthMonth1.Text = "November"
    End If
    If txtBirthMonth2.Text = "11" Then
        txtBirthMonth2.Text = "November"
    End If
    If txtBirthMonth1.Text = "12" Then
        txtBirthMonth1.Text = "December"
    End If
    If txtBirthMonth2.Text = "12" Then
        txtBirthMonth2.Text = "December"
    End If
    
End Sub

