VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =======================================================================
'
' CLASS   : clsPrintDialog
' PURPOSE : Show a dialog to select a printer and to set printer
'           properties. Selected printer will be set to Printer object
'           WITHOUT CHANGING DEFAULT PRINTER.
'           Printer object could print to selected printer then.
'           This class work like "MS Common Dialogs" ShowPrinter method,
'           but it set Printer object without changing default printer.
' HELP    : Look "MS Common Dialogs" OCX help, it applies also to this
'           class. All Flags constants are supported.
' NOTE    : Do you have found any bug or improvement ?
'           Please let me know, that's why I'm sharing source code.
' AUTHOR  : ___________________________________________________
'            Luca Minudel                    software designer
'            Italy Conegliano(TV)
'            voice & fax                     +39 (0)438 412280
'            e-mail                      luca.minudel@nline.it
'            WWW                       (italian language used)
'            http://www.geocities.com/SiliconValley/Vista/4041
'
' =======================================================================
'
' --- API CONSTANTS
'
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
'
' --- API TYPES DEFINITION
'
Private Type PRINTDLG_TYPE
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hdc As Long
  Flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type
Private Type DEVMODE_TYPE
  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type
'
' --- API DECLARATIONS
'
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" _
                         (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                    (hpvDest As Any, _
                     hpvSource As Any, _
                     ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" _
                         (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" _
                         (ByVal wFlags As Long, _
                          ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'
' --- PUBLIC ENUM
'
Public Enum PrinterConstants
  cdlPDAllPages = &H0
  cdlPDCollate = &H10
  cdlPDDisablePrintToFile = &H80000
  cdlPDHelpButton = &H800
  cdlPDHidePrintToFile = &H100000
  cdlPDNoPageNums = &H8
  cdlPDNoSelection = &H4
  cdlPDNoWarning = &H80
  cdlPDPageNums = &H2
  cdlPDPrintSetup = &H40
  cdlPDPrintToFile = &H20
  cdlPDReturnDC = &H100
  cdlPDReturnDefault = &H400
  cdlPDReturnIC = &H200
  cdlPDSelection = &H1
  cdlPDUseDevModeCopies = &H40000
End Enum
Public Enum ErrorConstants
  cdlCancel = 32755
End Enum
'
' --- PRIVATE VARIABLES
'
Private intMinPage As Integer  ' Local copy of Min
Private intMaxPage As Integer  ' Local copy of Max
Private intFromPage As Integer ' Local copy of FromPage
Private intToPage As Integer   ' Local copy of ToPage
' N.B. 0 >= Min >= FromPage >= ToPage >= Max
'      If Max=0 then no limits.
'
' --- PUBLIC VARIABLES
'
Public Flags As PrinterConstants
Public CancelError As Boolean
'
' -- INITIALIZE
'
Private Sub Class_Initialize()
  intMinPage = 0
  intMaxPage = 0
  intFromPage = 0
  intToPage = 0
  CancelError = True
End Sub
'
' -- PUBLIC MEMBERS
'
Property Get Min() As Integer
  Min = intMinPage
End Property
Property Let Min(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMinPage = intNewValue
  If intNewValue > intFromPage Then _
    intFromPage = intNewValue
  If intNewValue > intToPage Then _
    intToPage = intNewValue
  If intNewValue > intMaxPage Then _
    intMaxPage = intNewValue
End Property
Property Get FromPage() As Integer
  FromPage = intFromPage
End Property
Property Let FromPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intFromPage = intNewValue
  If intNewValue > intToPage Then _
    intToPage = intNewValue
  If intNewValue > intMaxPage Then _
    intMaxPage = intNewValue
  If intNewValue < intMinPage Then _
    intMinPage = intNewValue
End Property
Property Get ToPage() As Integer
  ToPage = intToPage
End Property
Property Let ToPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intToPage = intNewValue
  If intNewValue > intMaxPage Then _
    intMaxPage = intNewValue
  If intNewValue < intFromPage Then _
    intFromPage = intNewValue
  If intNewValue < intMinPage Then _
    intMinPage = intNewValue
End Property
Property Get Max() As Integer
  Max = intMaxPage
End Property
Property Let Max(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMaxPage = intNewValue
  If intNewValue < intToPage Then _
    intToPage = intNewValue
  If intNewValue < intFromPage Then _
    intFromPage = intNewValue
  If intNewValue < intMinPage Then _
    intMinPage = intNewValue
End Property
Public Function ShowPrinter() As Boolean
Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE
Dim lpDevMode As Long, lpDevName As Long
Dim intReturn As Integer
Dim objPrinter As Printer
Dim strNewPrinterName As String
Dim blnCancel   As Boolean
  blnCancel = False
  ' Use PrintDialog to get the handle to a memory
  ' block with a DevMode and DevName structures
  With PrintDlg
    .lStructSize = Len(PrintDlg)
    .hwndOwner = 0
    .Flags = Flags
    .nMinPage = intMinPage
    .nFromPage = intFromPage
    .nToPage = intToPage
    .nMaxPage = intMaxPage
  End With
  'Set the current orientation and duplex setting
  DevMode.dmDeviceName = Printer.DeviceName
  DevMode.dmSize = Len(DevMode)
  DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
  DevMode.dmOrientation = Printer.Orientation
  On Error Resume Next
  DevMode.dmDuplex = Printer.Duplex
  On Error GoTo 0
  'Allocate memory for the initialization hDevMode structure
  'and copy the settings gathered above into this memory
  PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
                                  GMEM_ZEROINIT, Len(DevMode))
  lpDevMode = GlobalLock(PrintDlg.hDevMode)
  If lpDevMode > 0 Then
    CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
    intReturn = GlobalUnlock(lpDevMode)
  End If
  'Set the current driver, device, and port name strings
  With DevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
  End With
  With Printer
    DevName.extra = .DriverName & Chr(0) & _
                    .DeviceName & Chr(0) & .Port & Chr(0)
  End With
  'Allocate memory for the initial hDevName structure
  'and copy the settings gathered above into this memory
  PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
                                   GMEM_ZEROINIT, Len(DevName))
  lpDevName = GlobalLock(PrintDlg.hDevNames)
  If lpDevName > 0 Then
    CopyMemory ByVal lpDevName, DevName, Len(DevName)
    intReturn = GlobalUnlock(lpDevName)
  End If
  'Call the print dialog up and let the user make changes
  If PrintDialog(PrintDlg) Then
    'First get the DevName structure.
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    CopyMemory DevName, ByVal lpDevName, 45
    intReturn = GlobalUnlock(lpDevName)
    With PrintDlg
      Flags = .Flags
      intFromPage = .nFromPage
      intToPage = .nToPage
    End With
     GlobalFree PrintDlg.hDevNames
    'Next get the DevMode structure and set the printer
    'properties appropriately
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
    intReturn = GlobalUnlock(PrintDlg.hDevMode)
    GlobalFree PrintDlg.hDevMode
    strNewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
                        InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
    If Printer.DeviceName <> strNewPrinterName Then
      For Each objPrinter In Printers
        If UCase$(objPrinter.DeviceName) = strNewPrinterName Then _
          Set Printer = objPrinter
      Next
    End If
    On Error Resume Next
    'Set printer object properties according to selections made
    'by user
    With Printer
      .Copies = DevMode.dmCopies
      .Duplex = DevMode.dmDuplex
      .Orientation = DevMode.dmOrientation
    End With
    On Error GoTo 0
  Else
    GlobalFree PrintDlg.hDevMode
    GlobalFree PrintDlg.hDevNames
    blnCancel = True
    'If CancelError Then _
     ' ShowPrinter = True
  End If
  ShowPrinter = Not blnCancel
End Function


