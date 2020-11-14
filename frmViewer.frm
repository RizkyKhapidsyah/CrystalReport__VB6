VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewer 
   Caption         =   "Print Preview"
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   7560
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   6840
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.PictureBox CRViewer1 
      Height          =   7005
      Left            =   0
      ScaleHeight     =   6945
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   360
      Width           =   5805
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Printer"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s1"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Email"
            Object.ToolTipText     =   "Email"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Graphics"
            Object.ToolTipText     =   "Show / Hide Background"
            Style           =   1
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7530
      Top             =   3240
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
            Picture         =   "frmViewer.frx":0000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":27B4
            Key             =   "Email"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEmail 
      Caption         =   "E-mail"
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/////////////////////////////////////////////////////////////////////////////////////
'   frmViewer.
'
'   Purpose : To set up the environment for the crystal report
'
'   Written By : Eugene Wolff - August 2001
'
'   Requirements :  References  - Microsoft ActiveX Data Objects 2.5 Library
'                               - Microsoft ActiveX Data Objects Recordset 2.1 Library
'                   Components  - Crystal Report Viewer Control - CRVIEWER.DLL
'                               - Microsoft MAPI Controls 6.0 - MSMAPI.OCX
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


Dim Report As New CrystalReport1        ' Set up the report

Dim bLoggedon As Boolean                ' Check if email logged on

Dim conn3 As New ADODB.Connection       ' ADO Connection
Dim rs As New ADODB.Recordset           ' RS Connection

Private Sub Form_Load()
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass       ' Set the mouse pointer to an hour glass
       
    Report.DiscardSavedData                 ' Clear all saved changes and refresh report
    Report.PaperOrientation = crPortrait    ' Set the paper orientation of the report, this is
                                            ' good practice just incase the printer is set to
                                            ' Landscape
    
    Set conn3 = New ADODB.Connection        ' Set conn3
    conn3.Open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & NormalisePath(App.Path) & "testdb.mdb" ' Open the connection
    
    Set rs = New ADODB.Recordset            ' set rs
    rs.Open "test", conn3, adOpenKeyset, adLockReadOnly ' Open the recordset
    
    Report.Database.SetDataSource rs, , 1   ' Set Table number 1 to the new file
    
    Report.PaperSize = crPaperA4            ' Set A4 Paper Size
        
    CRViewer1.ReportSource = Report         ' Set the viewer to the crystal report
    CRViewer1.DisplayTabs = False           ' don't display the tabs on the top of the crystal report
    
    CRViewer1.ViewReport                    ' Show the report in the the crystal viewer
    CRViewer1.Zoom 1                        ' Set zoom to page width
    Printer.Orientation = 1                 ' Set the printer orientation, this is just something I like to do
    
    Report.txtHeading.SetText "This is to show how to set text in a report" ' Changing a text field on the report
        
    Screen.MousePointer = vbNormal          ' Set the mouse pointer back to normal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmViewer
    End
End Sub

Private Sub Form_Resize()
    '///////////////////////////////////////////////////////////
    ' Clean up if the form is resized
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    CRViewer1.Top = 330                     ' To take into account the toolbar
    CRViewer1.Left = 0                      ' set the left
    CRViewer1.Height = ScaleHeight - 350    ' To take into account the taskbar
    CRViewer1.Width = ScaleWidth            ' set to screen width
    
End Sub

Private Sub mnuEmail_Click()
    '///////////////////////////////////////////////////////////
    ' To email the report as a .pdf file
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    On Error Resume Next
    
    Report.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the export format as .pdf
    Report.ExportOptions.DestinationType = crEDTDiskFile        ' Save it to disk
    
    P = App.Path & "\Test.pdf"                                  ' The file name and path
    
    Report.ExportOptions.DiskFileName = P                       ' Export to file
    Report.Export False 'True will prompt user for export options
        
    If bLoggedon = False Then
        ' Set mapi session first
        With mapSess
            .DownLoadMail = False   ' don't do an email download now
            .LogonUI = True         ' Use default username and password
            .SignOn                 ' open session
            .NewSession = True      ' a new email
            bLoggedon = True
            mapMess.SessionID = .SessionID  ' important!! set message sessionid to session sessionid
        End With
    End If
    
    With mapMess
        .Compose                    ' Compose a new mail
        .MsgSubject = "Demo for crystal report"   ' Set the Subject of the email
        .MsgNoteText = "This is a demo to show how to email a crystal report"   ' Set the message
                       
        .AttachmentPathName = P     ' attachment path and name
        .Send True  ' open the email to allow user to add recipient
                    ' if false the recipient must be hardcoded
    End With
    frmViewer.CRViewer1.SetFocus       'Return to viewer
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPrint_Click()
    '///////////////////////////////////////////////////////////
    ' This is thanks to Luca Minudel for his printer class
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Const ErrCancel = 32755         ' This is if cancel was pressed in the printer dialog
    Dim P As New clsPrintDialog     ' Set p = the printer class
    On Error GoTo errorPrinter
    
    Dim printing As Boolean
    P.Min = 1                       ' The first page
    P.Max = Report.PrintingStatus.NumberOfPages ' find the number of pages
    P.ToPage = P.Max                ' In the printer dialog show the to page as the last page
    printing = P.ShowPrinter        ' Show printer
    
    If printing = False Then        ' if there are no printers just exit
        Exit Sub
    End If
    Printer.Orientation = 1         ' Set the printer orientation
    Report.PaperOrientation = crPortrait    ' Set the report orientation
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port   ' Sets the selected printer to the report printer
        
    Report.PrintOut False, , , P.FromPage, P.ToPage ' Print the report
    Set P = Nothing
    Exit Sub
errorPrinter:
    If Err = ErrCancel Then Exit Sub Else Resume ' If cancel was pressed exit the printing sub
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index        ' Check which button was pressed on the toolbar
        Case 1      ' Print
            Printer.Orientation = 1 ' Set printer orientation
            Printer.PrintQuality = vbPRPQHigh   ' Set the printer Quality
            Report.SelectPrinter DefaultDriverName, DefaultDeviceName, DefaultPort 'Set the default printer
            Report.PaperOrientation = crPortrait    ' Set the report Orientation
            Report.PrintOut False   ' Print report without prompting
        Case 3      ' Email
            mnuEmail_Click  ' Do the same as the email menu
    End Select
        
End Sub
