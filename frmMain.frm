VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C++ Header File to VB Module Converter"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ProgressBar 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   8
      Top             =   6480
      Width           =   9255
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   7815
   End
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   120
      Width           =   1275
      Begin VB.CommandButton cmdSeperator2 
         Enabled         =   0   'False
         Height          =   1395
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Height          =   975
         Left            =   0
         Picture         =   "frmMain.frx":150C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSeperator 
         Enabled         =   0   'False
         Height          =   975
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Height          =   975
         Left            =   0
         Picture         =   "frmMain.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Translate"
         Enabled         =   0   'False
         Height          =   975
         Left            =   0
         Picture         =   "frmMain.frx":26A0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   975
         Left            =   0
         Picture         =   "frmMain.frx":2F6A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConvert_Click()

    'Disable command buttons:
    cmdConvert.Enabled = False
    cmdExport.Enabled = False
        
    C_to_VB.Convert txtLog
    
    'Enable command buttons:
    If cmdExport.Enabled = False Then cmdExport.Enabled = True
    If cmdCopy.Enabled = False Then cmdCopy.Enabled = True
       
End Sub

Private Sub cmdCopy_Click()

    C_to_VB.Copy txtLog

End Sub

Private Sub cmdExport_Click()

    With CommonDialog
        .CancelError = False
        .DialogTitle = "Export..."
        .Filter = "Visual Basic Module (*.bas)|*.bas|All Files (*.*)|*.*"
        .Flags = cdlOFNHideReadOnly
        .ShowSave
        
        If Len(.FileName) > 0 Then
            
            'Do not allow to overwrite C++ files by accident:
            If Right$(.FileName, 2) = ".h" Then .FileName = Left$(.FileName, Len(.FileName) - 2) & ".bas"
            If Right$(.FileName, 4) = ".cpp" Then .FileName = Left$(.FileName, Len(.FileName) - 4) & ".bas"
            
            'Check if the file already exists:
            If FileExist(.FileName) = True Then
                Dim lRetVal As Long: lRetVal = MsgBox("File " & .FileName & " already exists." & vbCrLf & "Do you want to overwrite it?", vbYesNo + vbExclamation, "C++ To VB Converter")
                Select Case lRetVal
                    Case vbYes
                    Case vbNo: Exit Sub
                    Case Else: Exit Sub
                End Select
            End If
            
            'Export buffer to file:
            C_to_VB.Export .FileName, txtLog
            
        End If
        
    End With

End Sub

Private Sub cmdImport_Click()

    With CommonDialog
        .CancelError = False
        .DialogTitle = "Import..."
        .Filter = "C++ Header Files (*.h)|*.h|All Files (*.*)|*.*"
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        
        If Len(.FileName) > 0 Then
        
            'Import the file:
            C_to_VB.Import .FileName, txtLog
            
            'Enable command buttons:
            If cmdConvert.Enabled = False Then cmdConvert.Enabled = True
                        
        End If
        
    End With
   
End Sub


Private Function FileExist(PathName As String) As Boolean
On Error GoTo File_Does_Not_Exist
    
    Dim lRetVal As Long: lRetVal = FileLen(PathName)
    FileExist = True
    
    Exit Function
    
File_Does_Not_Exist:

    FileExist = False
    
End Function
