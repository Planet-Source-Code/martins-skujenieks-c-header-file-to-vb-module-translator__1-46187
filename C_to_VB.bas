Attribute VB_Name = "C_to_VB"

'/*******************************************************************
'*                                                                  *
'*  ----                                                            *
'*  C++ To VB Conversion Module                                     *
'*  ----                                                            *
'*  Converts C++ header files to Visual Basic module files.         *
'*  Convert function excludes unsupported compiler flags,           *
'*  converts remarks, type and constant definitions, hex values     *
'*  and API declarations. Average efficiency is ~92%.               *
'*                                                                  *
'*                                                                  *
'*  ----                                                            *
'*  File Information                                                *
'*  ----                                                            *
'*  Version:        2.00                                            *
'*  Created:        11th June, 2003                                 *
'*  Last Modified:  14th June, 2003                                 *
'*                                                                  *
'*                                                                  *
'*  ----                                                            *
'*  Legal Copyright                                                 *
'*  ----                                                            *
'*  Copyright © Martins Skujenieks 2003                             *
'*                                                                  *
'*                                                                  *
'*  ----                                                            *
'*  End user License Agreement (EULA)                               *
'*  ----                                                            *
'*  This product is provided "as is", with no guarantee of          *
'*  completeness or accuracy and without warranty of any kind,      *
'*  express or implied.                                             *
'*                                                                  *
'*  In no event will developer be liable for damages of any         *
'*  kind that may be incurred with your hardware, peripherals       *
'*  or software programs.                                           *
'*                                                                  *
'*  You may make as much copies of this file as you wish if they    *
'*  are for your personal, non-commercial, home use only, provided  *
'*  you keep intact all copyright and other proprietary notices.    *
'*  No parts of this file may be copied, reproduced, modified,      *
'*  republished, uploaded, posted, transmitted or distributed       *
'*  in any way, without prior written consent of developer.         *
'*                                                                  *
'*                                                                  *
'*  ----                                                            *
'*  Questions, comments or suggestions?                             *
'*  ----                                                            *
'*  Visit:          http://www.exe.times.lv                         *
'*  E-Mail:         martins_s@mail.teliamtc.lv                      *
'*                                                                  *
'*******************************************************************/


Option Explicit

                                        '/* read-only */
    Public Buffer(16384) As String      '/* RAM buffer */
    Public LineCount As Long            '/* line count */
    
    Private Stream As String            '/* internal */
    
        
    
'/*******************************************************************
'*  External functions
'*******************************************************************/


'/*******************************************************************
'*                                                                  *
'*  Import          Function loads file into RAM buffer.            *
'*                                                                  *
'*  PathName        Path and name of the file to load.              *
'*                                                                  *
'*******************************************************************/
Public Sub Import(PathName As String, Log As TextBox)
On Error Resume Next

    '// Write log property:
    Log.Text = vbNullString
    Log.Text = Log.Text & "Importing..." & vbCrLf
    DoEvents

    LineCount = 0
    
    Open PathName For Input As #1
    
        '// Read line from file untile the end-of-file is reached.
        Do Until EOF(1) = True
        
            DoEvents
            
            Buffer(LineCount) = vbNullString
            
            Do
                DoEvents
                
                
                '// Read line from file byte-by-byte.
                '// Byte-by-byte reading used, because Input function
                '// removes from text commas and inserts line-breaks
                '// instead of them.
                '// This is slower, but definately we don't want
                '// to lose data.
                Stream = vbNullString
                Do
                    DoEvents
                    Stream = Stream & Input(1, #1)
                    If Right$(Stream, 2) = vbCrLf Then
                        Stream = Left$(Stream, Len(Stream) - 2)
                        Exit Do
                    End If
                Loop
                
                'Replace tab's with spaces:
                Stream = Replace(Stream, Chr$(vbKeyTab), "    ")
                
                '// Check wether current line is or is not complete line.
                '// C++ compiler handles code in way the VB doesn't, so
                '// we need to get some code lines together in one line.
                'If _
                'Len(Stream) = 0 Or _
                'InStr(1, Stream, "/") Or _
                'InStr(1, Stream, "*") Or _
                'InStr(1, Stream, "#") Or _
                'InStr(1, Stream, "{") Or _
                'InStr(1, Stream, "}") Or _
                'InStr(1, Stream, ",") Or _
                'InStr(1, Stream, ";") Then
                    Buffer(LineCount) = Buffer(LineCount) & Stream
                    Exit Do
                'Else
                '    Buffer(LineCount) = Buffer(LineCount) & " " & Stream
                'End If
                
            Loop
            
            LineCount = LineCount + 1
            
        Loop
        
    Close #1
       
    '// Write log property:
    Log.Text = Log.Text & "Buffered " & LineCount & " lines of code (" & Int(FileLen(PathName) / 1024) & " KB)!" & vbCrLf
    Log.Text = Log.Text & "Imported " & StrConv(PathName, vbProperCase) & "!" & vbCrLf
    DoEvents
    
End Sub



'/*******************************************************************
'*                                                                  *
'*  Export          Function saves RAM buffer to file.              *
'*                                                                  *
'*  PathName        Path and name of the file to save to.           *
'*                                                                  *
'*******************************************************************/
Public Sub Export(PathName As String, Log As TextBox)
On Error Resume Next

    '// Write log:
    Log.Text = Log.Text & "Exporting..." & vbCrLf
    DoEvents

    Dim Line As Long
    
    Open PathName For Output As #1
    
        For Line = 0 To LineCount
        
            Print #1, Buffer(Line)
            
        Next Line
        
    Close #1
    
    '// Write log property:
    Log.Text = Log.Text & "Exported " & StrConv(PathName, vbProperCase) & "!" & vbCrLf
    DoEvents

End Sub



'/*******************************************************************
'*                                                                  *
'*  Display         Function displays RAM buffer into TextBox.      *
'*                                                                  *
'*  ControlName     Specifies control that has Text property,       *
'*                  e.g. TextBox, RichTextBox etc.                  *
'*                                                                  *
'*******************************************************************/
Public Sub Display(Box As TextBox)
On Error GoTo Err_Handler

    '// This buffer may require up to 16MB of memory!!!
    Dim TempBuffer As String
    
    Dim Line As Long
    For Line = 0 To LineCount
        TempBuffer = TempBuffer & Buffer(Line) & vbCrLf
    Next Line

    Box.Text = vbNullString
    Box.Text = TempBuffer
    
    '// Free used memory:
    TempBuffer = vbNullString
   
    Exit Sub
Err_Handler:
    
    '// Probably, the control can't display files larger then ~52 KB.
    If Err.Number = 7 Then Box.Text = "Out of memory!"

End Sub



'/*******************************************************************
'*                                                                  *
'*  Copy            Function copies RAM buffer to Clipboard.        *
'*                                                                  *
'*******************************************************************/
Public Sub Copy(Log As TextBox)
On Error GoTo Err_Handler

    '// Write log:
    Log.Text = Log.Text & "Copying to clipboard..." & vbCrLf
    DoEvents

    '// This buffer may require up to 16MB of memory!!!
    Dim ClipboardBuffer As String
    
    Dim Line As Long
    For Line = 0 To LineCount
        ClipboardBuffer = ClipboardBuffer & Buffer(Line) & vbCrLf
    Next Line
    
    Clipboard.Clear
    Clipboard.SetText ClipboardBuffer
    
    '// Write log property:
    Log.Text = Log.Text & "Copied to clipboard!" & vbCrLf
    DoEvents
    
    '// Free used memory:
    ClipboardBuffer = vbNullString
   
    Exit Sub
Err_Handler:
    
    '// Probably, there is not enough physicall memory to complete operation:
    MsgBox Err.Number

End Sub



'/*******************************************************************
'*                                                                  *
'*  Convert         Function converts buffered C++ code to VB code. *
'*                  Function writes conversion details in Log       *
'*                  property and processing progress in Progress    *
'*                  property.                                       *
'*                                                                  *
'*******************************************************************/
Public Sub Convert(Log As TextBox)

    Log.Text = Log.Text & "Excluding unsupported compiler flags..." & vbCrLf
        Exclude_Unsupported_Compiler_Flags
        DoEvents
        
    Log.Text = Log.Text & "Converting remarks..." & vbCrLf
        Convert_Remarks
        DoEvents
        
    Log.Text = Log.Text & "Converting boolean operators..." & vbCrLf
        Convert_Boolean_Operators
        DoEvents
        
    Log.Text = Log.Text & "Converting hexadecimal values..." & vbCrLf
        Convert_Hexadecimal_Values
        DoEvents
        
    Log.Text = Log.Text & "Converting constant definitions..." & vbCrLf
        Convert_Constant_Definitions
        DoEvents
        
    Log.Text = Log.Text & "Converting type definitions..." & vbCrLf
        Convert_Type_Definitions
        DoEvents
        
    'log.text = log.text &  "Converting API declarations..." & vbcrlf
        'Convert_API_Declarations
        'DoEvents
        
    Log.Text = Log.Text & "Removing casts..." & vbCrLf
        Remove_Casts
        DoEvents
        
    Log.Text = Log.Text & "Successfully completed!" & vbCrLf
    DoEvents

End Sub



'/*******************************************************************
'*  Internal functions
'*******************************************************************/

Private Sub Exclude_Unsupported_Compiler_Flags()

    Dim LineNum As Long
    Dim LinePos As Long
    Dim Temp As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
        
        If _
        InStr(1, Temp, "#if") Or _
        InStr(1, Temp, "#ifdef") Or _
        InStr(1, Temp, "#ifndef") Or _
        InStr(1, Temp, "#define") Or _
        InStr(1, Temp, "#include") Or _
        InStr(1, Temp, "#error") Or _
        InStr(1, Temp, "#pragma") Or _
        InStr(1, Temp, "#elseif") Or _
        InStr(1, Temp, "#else") Or _
        InStr(1, Temp, "#endif") Then
            If Left$(Temp, 1) <> "'" Then Temp = "'" & Temp
        End If
    
        If _
        InStr(1, Temp, "DECLARE_HANDLE") Then
            If Left$(Temp, 1) <> "'" Then Temp = "'" & Temp
        End If
    
        If _
        InStr(1, Temp, "extern") Then
            If Left$(Temp, 1) <> "'" Then Temp = "'" & Temp
        End If
        
        Buffer(LineNum) = Temp
            
    Next

End Sub

Private Sub Convert_Remarks()

    Dim LineNum As Long
    Dim LinePos As Long
    Dim Temp As String
    Dim Remark As Boolean
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
        
        Temp = Replace(Temp, "//", "'//")
        Temp = Replace(Temp, "/*", "'/*")
        
        If Remark = True Then
            If Left$(Temp, 1) <> "'" Then Temp = "'" & Temp
        End If
        
        If InStr(1, Temp, "/*") Then Remark = True
        If InStr(1, Temp, "*/") Then Remark = False
        
        Buffer(LineNum) = Temp

    Next
        
End Sub

Private Sub Convert_Boolean_Operators()

    Dim LineNum As Long
    Dim LinePos As Long
    Dim Temp As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
        
        If _
        InStr(1, Temp, "&") Or _
        InStr(1, Temp, "|") Then
            Temp = Replace(Temp, "&", "And")
            Temp = Replace(Temp, "|", "Or")
        End If

        Buffer(LineNum) = Temp
    
    Next

End Sub

Private Sub Convert_Hexadecimal_Values()

    Dim LineNum As Long
    Dim LinePos As Long
    Dim LPos As Long
    Dim SpacePos As Long
    Dim Temp As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
        
        If InStr(1, Temp, "0x") > 0 Then
            Temp = Temp & " "
            LinePos = InStr(1, Temp, "0x")
            LPos = InStr(LinePos, Temp, "L")
            SpacePos = InStr(LinePos, Temp, " ")
            If LinePos < LPos And LPos < SpacePos Then
                Temp = Left$(Temp, LPos - 1) & "&" & Right$(Temp, Len(Temp) - LPos)
            End If
            Temp = Replace(Temp, "0x", "&H")
            If Right$(Temp, 1) = " " Then Temp = Left$(Temp, Len(Temp) - 1)
        End If
        
        Buffer(LineNum) = Temp

    Next
        
End Sub

Private Sub Convert_Constant_Definitions()

    Dim LineLen As Long
    Dim LineNum As Long
    Dim LinePos As Long
    Dim Temp As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
        
        If InStr(1, Temp, "#define") > 0 Then
        
            LineLen = Len(Temp)
            LinePos = InStr(1, Temp, "#define")
            
            Do
                If LinePos >= LineLen Then Exit Do
                If Mid$(Temp, LinePos, 1) = " " Then Exit Do
                LinePos = LinePos + 1
            Loop
            
            Do
                If LinePos >= LineLen Then Exit Do
                If Mid$(Temp, LinePos, 1) <> " " Then Exit Do
                LinePos = LinePos + 1
            Loop
            
            Do
                If LinePos >= LineLen Then Exit Do
                If Mid$(Temp, LinePos, 1) = " " Then Exit Do
                LinePos = LinePos + 1
            Loop
            
            Do While Right$(Temp, 1) = " "
                Temp = Left$(Temp, Len(Temp) - 1)
            Loop
            
            Temp = Left$(Temp, LinePos) & "=" & Right$(Temp, Len(Temp) - LinePos)
            
            If Right$(Temp, 1) = "=" Then
                Temp = Left$(Temp, Len(Temp) - 1)
                If Left$(Temp, 1) <> "'" Then Temp = "'" & Temp
            Else
                Temp = Replace(Temp, "#define", "Public Const")
                If Left$(Temp, 1) = "'" Then Temp = Right$(Temp, Len(Temp) - 1)
            End If
        
        End If
        
        Buffer(LineNum) = Temp

    Next

End Sub

Private Sub Convert_Type_Definitions()

    Dim LineNum As Long
    Dim Temp As String
    Dim TypeDef As Boolean
    Dim TypeDefEnd As Boolean
    Dim Variable As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
               
        If InStr(1, Temp, "typedef") Then TypeDef = True
        
        If TypeDef = True Then
        
        
            'Insert space at line begining:
            If Left$(Temp, 1) <> " " Then Temp = " " & Temp
        
        
            'Convert C++ style brackets to VB compatible:
            Temp = Replace(Temp, "[", "(")
            Temp = Replace(Temp, "]", ")")
            
        
            'Translate begining of type definition:
            Temp = Replace(Temp, "typedef struct _tag", "Public Type ")
            Temp = Replace(Temp, "typedef struct tag", "Public Type ")
            Temp = Replace(Temp, "typedef struct _", "Public Type ")
            Temp = Replace(Temp, "typedef struct", "Public Type")
            Temp = Replace(Temp, "{", vbNullString)
            
            
            'Translate C++ variable type to VB:
            If InStr(1, Temp, " char ") Then Variable = " As String * 1"
            If InStr(1, Temp, " CHAR ") Then Variable = " As String * 1"
            If InStr(1, Temp, " WCHAR ") Then Variable = " As String * 1"
            If InStr(1, Temp, " byte ") Then Variable = " As Byte"
            If InStr(1, Temp, " BYTE ") Then Variable = " As Byte"
            If InStr(1, Temp, " int ") Then Variable = " As Long"
            If InStr(1, Temp, " INT ") Then Variable = " As Integer"
            If InStr(1, Temp, " UINT ") Then Variable = " As Integer"
            If InStr(1, Temp, " PUINT ") Then Variable = " As Integer"
            If InStr(1, Temp, " WORD ") Then Variable = " As Integer"
            If InStr(1, Temp, " DWORD ") Then Variable = " As Long"
            If InStr(1, Temp, " long ") Then Variable = " As Long"
            If InStr(1, Temp, " LONG ") Then Variable = " As Long"
            If InStr(1, Temp, " double ") Then Variable = " As Double"
            If InStr(1, Temp, " LPSTR ") Then Variable = " As String"
            If InStr(1, Temp, " LPCSTR ") Then Variable = " As String"
            If InStr(1, Temp, " LPWSTR ") Then Variable = " As String"
            If InStr(1, Temp, " POINT ") Then Variable = " As POINT"
            If InStr(1, Temp, " RECT ") Then Variable = " As RECT"

            If InStr(1, Temp, " HWND ") Then Variable = " As Long"
            If InStr(1, Temp, " HDC ") Then Variable = " As Long"
            If InStr(1, Temp, " HGDIOBJ ") Then Variable = " As Long"
            If InStr(1, Temp, " HRGN ") Then Variable = " As Long"
            If InStr(1, Temp, " HBITMAP ") Then Variable = " As Long"
            If InStr(1, Temp, " HFONT ") Then Variable = " As Long"
            If InStr(1, Temp, " COLORREF ") Then Variable = " As Long"
            
            
            'Remove C++ variable types:
            Temp = Replace(Temp, " char ", vbNullString)
            Temp = Replace(Temp, " CHAR ", vbNullString)
            Temp = Replace(Temp, " WCHAR ", vbNullString)
            Temp = Replace(Temp, " byte ", vbNullString)
            Temp = Replace(Temp, " BYTE ", vbNullString)
            Temp = Replace(Temp, " int ", vbNullString)
            Temp = Replace(Temp, " INT ", vbNullString)
            Temp = Replace(Temp, " UINT ", vbNullString)
            Temp = Replace(Temp, " PUINT ", vbNullString)
            Temp = Replace(Temp, " WORD ", vbNullString)
            Temp = Replace(Temp, " DWORD ", vbNullString)
            Temp = Replace(Temp, " long ", vbNullString)
            Temp = Replace(Temp, " LONG ", vbNullString)
            Temp = Replace(Temp, " double ", vbNullString)
            Temp = Replace(Temp, " LPSTR ", vbNullString)
            Temp = Replace(Temp, " LPCSTR ", vbNullString)
            Temp = Replace(Temp, " LPWSTR ", vbNullString)
            Temp = Replace(Temp, " POINT ", vbNullString)
            Temp = Replace(Temp, " RECT ", vbNullString)

            Temp = Replace(Temp, " HWND ", vbNullString)
            Temp = Replace(Temp, " HDC ", vbNullString)
            Temp = Replace(Temp, " HGDIOBJ ", vbNullString)
            Temp = Replace(Temp, " HRGN ", vbNullString)
            Temp = Replace(Temp, " HBITMAP ", vbNullString)
            Temp = Replace(Temp, " HFONT ", vbNullString)
            Temp = Replace(Temp, " COLORREF ", vbNullString)
            
            
            'Check if the variable is pointer:
            If InStr(1, Temp, "*") <> 0 Then
                'Temp = Replace(Temp, "*", vbNullString)
                Variable = Variable & " '*"
            End If
            
            
            'Insert VB coded variable:
            Temp = Replace(Temp, ";", Variable)
            Variable = vbNullString
            
            
            'Translate end of type definition:
            Temp = Replace(Temp, "}", "End Type '}")
            
            
            'Format the VB code:
            If InStr(1, Temp, "Type") = 0 Then
                If Left$(Temp, 5) <> "     " Then
                    Do
                        DoEvents
                        Temp = " " & Temp
                        If Left$(Temp, 5) = "     " Then Exit Do
                    Loop
                Else
                    Do Until Left$(Temp, 6) <> "      "
                        DoEvents
                        Temp = Right$(Temp, Len(Temp) - 1)
                    Loop
                End If
            End If
            
            
            'Remove added space at line begining:
            If Left$(Temp, 1) = " " Then Temp = Right$(Temp, Len(Temp) - 1)
            
            If InStr(1, Temp, "}") Then TypeDef = False
        
                
        End If
        
        Buffer(LineNum) = Temp

    Next

End Sub

Private Sub Convert_API_Declarations()
    
    '// Under development
    
    '// My first two tries was able to translate only about 25-30% API
    '// declarations correctly. I am currently developing better algorythm.
    
    '// As there is very many ways how to write API declaration in C++,
    '// it is difficult to write good function for all types of calling.
    
End Sub

Public Sub Remove_Casts()

    Dim LineNum As Long
    Dim Temp As String
    
    For LineNum = 0 To LineCount
    
        Temp = Buffer(LineNum)
               
        'Remove C++ variable casts:
        Temp = Replace(Temp, "(char)", vbNullString)
        Temp = Replace(Temp, "(CHAR)", vbNullString)
        Temp = Replace(Temp, "(WCHAR)", vbNullString)
        Temp = Replace(Temp, "(byte)", vbNullString)
        Temp = Replace(Temp, "(BYTE)", vbNullString)
        Temp = Replace(Temp, "(int)", vbNullString)
        Temp = Replace(Temp, "(INT)", vbNullString)
        Temp = Replace(Temp, "(UINT)", vbNullString)
        Temp = Replace(Temp, "(PUINT)", vbNullString)
        Temp = Replace(Temp, "(WORD)", vbNullString)
        Temp = Replace(Temp, "(DWORD)", vbNullString)
        Temp = Replace(Temp, "(long)", vbNullString)
        Temp = Replace(Temp, "(LONG)", vbNullString)
        Temp = Replace(Temp, "(double)", vbNullString)
        Temp = Replace(Temp, "(LPSTR)", vbNullString)
        Temp = Replace(Temp, "(LPCSTR)", vbNullString)
        Temp = Replace(Temp, "(LPWSTR)", vbNullString)
        Temp = Replace(Temp, "(void)", vbNullString)
        Temp = Replace(Temp, "(void*)", vbNullString)
        Temp = Replace(Temp, "(void *)", vbNullString)
        
        Buffer(LineNum) = Temp

    Next

End Sub


