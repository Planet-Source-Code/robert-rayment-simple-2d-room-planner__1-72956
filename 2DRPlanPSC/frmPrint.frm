VERSION 5.00
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Sent to printer  CLOSE"
      Height          =   855
      Left            =   1410
      TabIndex        =   1
      Top             =   660
      Width           =   1650
   End
   Begin VB.PictureBox picGet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPrint.frm

Option Explicit

Private Sub DoPrint()
   Const vbHiMetric As Integer = 8
'   Const vbTwips As Integer = 1
'   Const vbPixels As Integer = 3
   Dim FullPrnWidth As Double
   Dim FullPrnHeight As Double
   
   Dim PrnWidth As Double
   Dim PrnHeight As Double
   Dim PrnLeft As Double
   Dim PrnTop As Double
   
   Dim PrnPicWidth As Double
   Dim PrnPicHeight As Double
   Dim PrnPicLeft As Double
   Dim PrnPicTop As Double
   
   Dim Frac1 As Single
   Dim Frac2 As Single

   'GoTo SkipPrinter   ' For testing
   
   Frac1 = 0.9
   Frac2 = 0.05
   
   Printer.Orientation = 2 'Orient  ' 1-portrait, 2-landscape
   
   ' Calculate the dimensions of the printable area in HiMetric.
   FullPrnWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHiMetric)
   FullPrnHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHiMetric)
   
   ' Adjust printable width & height
   PrnWidth = FullPrnWidth * Frac1
   PrnHeight = FullPrnHeight * Frac1
   ' Scale width & height
   PrnPicWidth = Printer.ScaleX(PrnWidth, vbHiMetric, Printer.ScaleMode)
   PrnPicHeight = Printer.ScaleY(PrnHeight, vbHiMetric, Printer.ScaleMode)
   
   ' Adjust left & top
   PrnLeft = FullPrnWidth * Frac2
   PrnTop = FullPrnHeight * Frac2
   ' Scale left & top
   PrnPicLeft = Printer.ScaleX(PrnLeft, vbHiMetric, Printer.ScaleMode)
   PrnPicTop = Printer.ScaleX(PrnTop, vbHiMetric, Printer.ScaleMode)
   
   Printer.PaintPicture picGet.Image, PrnPicLeft, PrnPicTop, PrnPicWidth, PrnPicHeight

   Printer.EndDoc
   Printer.Orientation = 1 'Orient  ' 1-portrait, 2-landscape
   
SkipPrinter:
   picGet.Picture = LoadPicture("")
   picGet.Width = 4
   picGet.Height = 4
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
' Image on clipboard from Form1.mnuPrintForm_Click
' Printer LIVE
   
   ShowPrinter (Me.hWnd)
   If Len(PrtName$) = 0 Then
      picGet.Picture = LoadPicture("")
      picGet.Width = 4
      picGet.Height = 4
      cmdClose.Caption = "CANCELLED   CLOSE"
      Exit Sub
   End If
   picGet.Picture = Clipboard.GetData(vbCFBitmap)
   picGet.Refresh
   DoEvents
   DoPrint
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
