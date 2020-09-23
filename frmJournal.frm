VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJournal 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   7260
   Icon            =   "frmJournal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtfJournal 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmJournal.frx":030A
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New ..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open ..."
      End
      Begin VB.Menu mnufilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As ..."
      End
      Begin VB.Menu mnufilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnufilesep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select A&ll"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A crude, limited but serviceable RTF text editor
'for maintaining a biorhythm journal.

Option Explicit

Dim FileName As String
Dim CurrentFileName As String

Private Sub Form_Load()
    
    'Make sure the clipboard is empty
    
    Clipboard.Clear
    
    'Set a caption for the form
    
    frmJournal.Caption = " ~ BioJournal Entry on " & Date & " at " & Time & " ~"
    
End Sub

Private Sub frmJournal_Unload()

    'Clear the clipboard
    
    Clipboard.Clear
    
    'And dump the form
    
    Unload Me
    
End Sub


Private Sub mnuEditSelectAll_Click()

    rtfJournal.SelStart = 0
    rtfJournal.SelLength = Len(rtfJournal.Text)
    
End Sub

Private Sub mnuFileNew_Click()

    'Prompt to save before creating a new journal
    
    If rtfJournal.Text <> "" Then
        If MsgBox("Journal entry has not been saved. Continue?", vbYesNo, "New...") = vbYes Then
            rtfJournal.Text = ""
        End If
    End If

End Sub

Private Sub mnuFileOpen_Click()

    'Handle error generate by clicking Cancel button
    
    On Error GoTo CancelError
    
    'Set file extension filter
    
    CommonDialog1.Filter = "BioJournal Files *.rtf|*.rtf|Text Files *.txt|*.txt|All Files *.*|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DefaultExt = ".rtf"
    
    'Open the file dialog
    
    CommonDialog1.Action = 1
    
    'Select and open a file
    
    Open CommonDialog1.FileName For Input As 1
    rtfJournal.FileName = CommonDialog1.FileName 'Input$(LOF(1), 1)
    Close 1
    frmJournal.Caption = " ~ BioJournal ~ " & CommonDialog1.FileName & " ~"
    CurrentFileName = CommonDialog1.FileName

'Exit if Cancel button was clicked

CancelError:
    Exit Sub
    
End Sub

Private Sub mnuFileSave_Click()

    If CurrentFileName = "" Then
        mnuFileSaveAs_Click
        Exit Sub
    End If
    
    Open CurrentFileName For Output As 1
    Print #1, rtfJournal.Text
    Close #1
    
End Sub

Private Sub mnuFileSaveAs_Click()

    'Handle Cancel Button
    
    On Error GoTo CancelError
    
    'Set the file filter
    
    CommonDialog1.Filter = "BioJournal Files *.rtf|*.rtf|Text Files *.txt|*.txt|All Files *.*|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DefaultExt = ".rtf"
    
    'Display a Save As... dialog
    
    CommonDialog1.Action = 2
    
    'Save the file
    
    Open CommonDialog1.FileName For Output As 1
    Print #1, rtfJournal.Text
    Close #1
    frmJournal.Caption = " ~ BioJournal ~ " & CommonDialog1.FileName & " ~"
    CurrentFileName = CommonDialog1.FileName
    Exit Sub
    
CancelError:
    
    Exit Sub
    
End Sub

Private Sub mnuFilePrint_Click()

    'Handle Print Error
    
    On Error GoTo PrintError
    
    'Set printer fonts
    
    Printer.FontName = rtfJournal.SelFontName
    Printer.FontSize = rtfJournal.SelFontSize
    Printer.FontBold = rtfJournal.SelBold
    Printer.FontItalic = rtfJournal.SelItalic
    Printer.FontStrikethru = rtfJournal.SelStrikeThru
    Printer.FontUnderline = rtfJournal.SelUnderline
    Printer.ForeColor = rtfJournal.SelColor
    
    'Print the journal
    
    Printer.Print rtfJournal.Text
    Printer.EndDoc
    Exit Sub
    
PrintError:
    
    If MsgBox("Unable to print document.", 21, "Printing Error...") = 4 Then
        'mnuFilePrint_Click()
    End If
    Exit Sub
    
End Sub

Private Sub mnuFileExit_Click()

    'Prompt to save file
    
    If rtfJournal.Text <> "" Then
        If MsgBox("Save the journal entry before exiting? ", vbYesNo, "Save...") = vbYes Then
        Exit Sub
        End If
    End If
    
    frmJournal_Unload
    
End Sub

Private Sub mnuEditCopy_Click()

    'Make sure the clipboard is clear
    
    Clipboard.Clear
    
    'Place selected text on the clipboard
    
    Clipboard.SetText rtfJournal.SelText
    
End Sub

Private Sub mnuEditCut_Click()
    
    'Clear the clipboard
    
    Clipboard.Clear
    
    'Copy the selected text to the clipboard
    
    Clipboard.SetText rtfJournal.SelText
    
    'Then delete it from the document
    
    rtfJournal.SelText = ""
    
End Sub

Private Sub mnuEditPaste_Click()

    'Get the text from the clipboard and paste it in
    
    rtfJournal.SelText = Clipboard.GetText()
    
End Sub

Private Sub mnuFormatFont_Click()

    'Handle Cancel Error
    
    On Error GoTo CancelError
    
    'Set flags to display printer fonts
    
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    
    'Set textbox font property
    
    With CommonDialog1
        .FontName = rtfJournal.SelFontName
        .FontSize = rtfJournal.SelFontSize
        .FontBold = rtfJournal.SelBold
        .FontItalic = rtfJournal.SelItalic
        .FontStrikethru = rtfJournal.SelStrikeThru
        .FontUnderline = rtfJournal.SelUnderline
        .Color = rtfJournal.SelColor
    End With
    
    'Display the font dialog
    
    CommonDialog1.Action = 4
    
    'OK button clicked
    
    With rtfJournal
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
        .SelStrikeThru = CommonDialog1.FontStrikethru
        .SelUnderline = CommonDialog1.FontUnderline
        .SelColor = CommonDialog1.Color
    End With
        
'Cancel button clicked

CancelError:

    Exit Sub
    
End Sub
