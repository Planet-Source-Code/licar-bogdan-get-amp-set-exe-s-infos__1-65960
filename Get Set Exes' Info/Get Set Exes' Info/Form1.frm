VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Exe's Info"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtDes 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox txtOFN 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtINF 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open Exe"
         Height          =   1095
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSetInfo 
         Caption         =   "Set Info"
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetInfo 
         Caption         =   "Get Info"
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtCN 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtPN 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtC 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblChDes 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Description"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblChOFN 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Original Filename"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblChINF 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Internal Name"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblChCN 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblChPN 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblChC 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Company Name"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Copyright"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CM1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'                  *** Get & Set Exes' Info ***
'
'A simple program that demonstrates how easy is to change exes'
'informations. It's just an example, not a distributable application.
'It basically replaces the informations of an exe with the new
'ones. One obviously cannot add more bytes to the files, so the
'lengths of the new infos have to be the same as the old ones.
'Therefore, if you want shorter infos, just add some spaces at
'the end.
'Sometimes you'll find invalid infos, because they are seldom
'written differently.
'
'You can get & set: - Company Name
'                   - Legal Copyright
'                   - Product Name
'                   - Internal Name
'                   - Original Filename
'                   - Description

Option Explicit

Dim str As String, txt As String, c As String, Path As String

'The positions of the file informations
Dim CN As Long, PN As Long, Cp As Long, INF As Long, OFN As Long, Des As Long

'The actual informations
Dim strCN As String, strPN As String, strCp As String
Dim strINF As String, strOFN As String, strDes As String

Private Sub cmdGetInfo_Click()
strCN = "": strCp = "": strPN = "": strOFN = "": strINF = "": strDes = "": txt = ""
txtC = "": txtCN = "": txtPN = "": txtINF = "": txtOFN = "": txtDes = ""

On Error Resume Next
Me.MousePointer = 11

'Opens the exe in binary for reading
Open Path For Binary As #1
    txt = Space$(LOF(1))
    Get #1, , txt
Close #1

'Stores the file informations
FileInfo "CompanyName", 26, CN, strCN
FileInfo "Copyright", 20, Cp, strCp
FileInfo "ProductName", 26, PN, strPN
FileInfo "InternalName", 26, INF, strINF
FileInfo "OriginalFilename", 34, OFN, strOFN
FileInfo "Description", 26, Des, strDes

'Cleans the infos
strCN = Replace(strCN, c, "")
strCp = Replace(strCp, c, "")
strPN = Replace(strPN, c, "")
strOFN = Replace(strOFN, c, "")
strINF = Replace(strINF, c, "")
strDes = Replace(strDes, c, "")

txtCN = strCN: lblChCN.Caption = "0"
txtC = strCp: lblChC.Caption = "0"
txtPN = strPN: lblChPN.Caption = "0"
txtINF = strINF: lblChINF.Caption = "0"
txtOFN = strOFN: lblChOFN.Caption = "0"
txtDes = strDes: lblChDes.Caption = "0"

cmdSetInfo.Enabled = True
txtCN.SetFocus
Me.MousePointer = 0
End Sub

Private Sub cmdOpen_Click()
strCN = "": strCp = "": strPN = "": strOFN = "": strINF = "": strDes = "": txt = ""
txtC = "": txtCN = "": txtPN = "": txtINF = "": txtOFN = "": txtDes = ""

With CM1
    .DialogTitle = "Choose An Exe To Open"
    .Filter = "Executables (*.exe)|*.exe"
    .ShowOpen
    Path = .FileName
End With

Frame2.Caption = Path
cmdGetInfo.Enabled = True
cmdGetInfo.SetFocus
End Sub

Private Sub cmdSetInfo_Click()
On Error GoTo ErrorHandler

'Replaces the old infos with the new ones
Open Path For Binary As #1
    Put #1, CN + 26, BinString(txtCN)
    Put #1, Cp + 20, BinString(txtC)
    Put #1, PN + 26, BinString(txtPN)
    Put #1, INF + 26, BinString(txtINF)
    Put #1, OFN + 34, BinString(txtOFN)
    Put #1, Des + 26, BinString(txtDes)
Close #1

ErrorHandler:
If Err Then MsgBox "An error has occured: " & Err.Description: Err.Clear
End Sub

'Modifies a string for being written in binary
Function BinString(ByVal str As String) As String
Dim i As Long
    For i = 1 To Len(str)
        BinString = BinString & Mid$(str, i, 1) & c
    Next i
End Function

'Finds the postions and the values of the infos
'Input: Info, InfoDist. Output: InfoPos, InfoVal.
Sub FileInfo(ByVal Info As String, InfoDist As Long, InfoPos As Long, InfoVal As String)
Dim i As Long
InfoPos = InStr(1, txt, BinString(Info))
For i = InfoPos + InfoDist To InfoPos + 100
    InfoVal = InfoVal & Mid$(txt, i, 1)
    If Mid$(txt, i, 3) = String$(3, 0) Then Exit For
Next i
End Sub

Private Sub Form_Load()
c = Chr$(0)
cmdSetInfo.Enabled = False
cmdGetInfo.Enabled = False
End Sub

Private Sub txtC_Change()
If Len(txtC) = Len(strCp) + 1 Then txtC = Left$(txtC, Len(strCp)): txtC.SelStart = Len(strCp)
lblChC.Caption = Len(strCp) - Len(txtC)
End Sub
Private Sub txtCN_Change()
If Len(txtCN) = Len(strCN) + 1 Then txtCN = Left$(txtCN, Len(strCN)): txtCN.SelStart = Len(strCN)
lblChCN.Caption = Len(strCN) - Len(txtCN)
End Sub
Private Sub txtINF_Change()
If Len(txtINF) = Len(strINF) + 1 Then txtINF = Left$(txtINF, Len(strINF)): txtINF.SelStart = Len(strINF)
lblChINF.Caption = Len(strINF) - Len(txtINF)
End Sub
Private Sub txtOFN_Change()
If Len(txtOFN) = Len(strOFN) + 1 Then txtOFN = Left$(txtOFN, Len(strOFN)): txtOFN.SelStart = Len(strOFN)
lblChOFN.Caption = Len(strOFN) - Len(txtOFN)
End Sub
Private Sub txtPN_Change()
If Len(txtPN) = Len(strPN) + 1 Then txtPN = Left$(txtPN, Len(strPN)): txtPN.SelStart = Len(strPN)
lblChPN.Caption = Len(strPN) - Len(txtPN)
End Sub
Private Sub txtDes_Change()
If Len(txtDes) = Len(strDes) + 1 Then txtDes = Left$(txtDes, Len(strDes)): txtDes.SelStart = Len(strDes)
lblChDes.Caption = Len(strDes) - Len(txtDes)
End Sub

Private Sub txtC_GotFocus()
txtC.SelLength = 1000
End Sub
Private Sub txtCN_GotFocus()
txtCN.SelLength = 1000
End Sub
Private Sub txtPN_GotFocus()
txtPN.SelLength = 1000
End Sub
Private Sub txtINF_GotFocus()
txtINF.SelLength = 1000
End Sub
Private Sub txtOFN_GotFocus()
txtOFN.SelLength = 1000
End Sub
Private Sub txtDes_GotFocus()
txtDes.SelLength = 1000
End Sub
