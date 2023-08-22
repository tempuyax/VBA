Attribute VB_Name = "Corloring"
'main function
'PLACE IN A STANDARD MODULE
Sub LoopThroughRows()
Dim i As Long, lastrow As Long
Dim pctdone As Single
lastrow = ActiveDocument.Paragraphs.Count
'(Step 1) Display your Progress Bar
ufprogress.LabelProgress.Width = 0
ufprogress.Show
For i = 1 To lastrow
'(Step 2) Periodically update progress bar
    pctdone = i / lastrow
    With ufprogress
        .LabelCaption.Caption = "Processing Row " & i & " of " & lastrow
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
        '--------------------------------------
        'the rest of your macro goes below here
        TextComment (i)
        '--------------------------------------
'(Step 3) Close the progress bar when you're done
    If i = lastrow Then Unload ufprogress
Next i
End Sub
Sub TextComment(p)
Dim i As Long, rEnd As Long, ct As Long
'MASM32 text
Dim MASM32Cmnt
MASM32Cmnt = Array("Include", "Invoke", "Local", "End", "DB", "DD", "Proc", "Endp", "Dup", "Equ", ".Data?", ".Data", ".Code", ".Const", "Addr")
'API Function text
Dim APIfunc(8) As String
APIfunc(0) = "InitCommonControls"
APIfunc(1) = "CoInitializeEx"
APIfunc(2) = "DialogBoxParam"
APIfunc(3) = "CoUninitialize"
APIfunc(4) = "ExitProcess"
APIfunc(5) = "MultiByteToWideChar"
APIfunc(6) = "CLSIDFromProgID"
APIfunc(7) = "CoCreateInstance"
APIfunc(8) = "EndDialog"
'Contant text
Dim ConstanTxt(16) As String
ConstanTxt(0) = "HINSTANCE"
ConstanTxt(1) = "NULL"
ConstanTxt(2) = "GUID"
ConstanTxt(3) = "HWND"
ConstanTxt(4) = "ULONG"
ConstanTxt(5) = "WPARAM"
ConstanTxt(6) = "LPARAM"
ConstanTxt(7) = "WM_INITDIALOG"
ConstanTxt(8) = "WM_COMMAND"
ConstanTxt(9) = "BN_CLICKED"
ConstanTxt(10) = "CP_ACP"
ConstanTxt(11) = "CLSCTX_SERVER"
ConstanTxt(12) = "BN_CLICKED"
ConstanTxt(13) = "TRUE"
ConstanTxt(14) = "FALSE"
ConstanTxt(15) = "WM_CLOSE"
ConstanTxt(16) = "VARIANT_TRUE"
'High Level text
Dim HighLevel
HighLevel = Array(".If", ".Endif", ".ElseIf")
'Prosesor Instruction text
Dim PROinst
PROinst = Array("Shl", "or")
' put object of range here
Dim rngParagraph As range
        Set rngParagraph = ActiveDocument.Paragraphs(p).range
        'rngParagraph.Select
        ct = (rngParagraph.Characters.Count - 1)
            For n = 1 To ct
                If rngParagraph.Characters(n).Text = ";" Then
                   rngParagraph.MoveStart Unit:=wdCharacter, Count:=(n - 1) 'set forward star
                   'rngParagraph.Select
                   rEnd = rngParagraph.Characters.Count  'safe range end
                   rngParagraph.Font.Color = wdColorGray50
                   If n = 1 Then Exit Sub Else GoTo lbNextExit
                End If
            Next n
lbNextExit:
     rngParagraph.Expand Unit:=wdParagraph 'reset to Paragraph range
     rngParagraph.MoveEnd Unit:=wdCharacter, Count:=-rEnd 'set backward end
     'rngParagraph.Select
     'rngParagraph.Font.Color = wdColorRed
     'rngParagraph.HighlightColorIndex = wdTurquoise
     'rngParagraph.Words(1).Case = wdTitleSentence
With rngParagraph.Find
    .MatchWholeWord = True
    .MatchWildcards = True
    'format text dgn tanda kutip "text" yg menjadi warna merah
    .Text = "\" + String(1, 34) + "*" + "\" + String(1, 34)
    .Replacement.Font.ColorIndex = wdRed
    .Execute Replace:=wdReplaceAll
    'format character bracet menjadi warna pink
    .Text = "[:,{}|=&#-+]" 'do not use !()
    .Replacement.Font.ColorIndex = wdPink
    .Execute Replace:=wdReplaceAll
    'format text angka hex 0000h yg menjadi warna cyan
    .Text = "0*(H)"
    .Replacement.Font.Color = RGB(0, 176, 240) 'Cyan
    .Execute Replace:=wdReplaceAll
    'format text integer menjadi warna green
    '.MatchWildcards = False
    '.Text = "*[0-9]"
    '.Replacement.Font.Color = wdGreen 'hijau
    '.Execute Replace:=wdReplaceAll
    '==============================
    .MatchWildcards = False
    For i = 0 To UBound(MASM32Cmnt)
        .Text = MASM32Cmnt(i)
        .Replacement.Font.ColorIndex = wdBlue
        If .Execute(Replace:=wdReplaceAll) = True Then
           GoTo Next100
        End If
    Next i
Next100:
    '==============================
    For i = 0 To UBound(APIfunc)
        .Text = APIfunc(i)
        .Replacement.Font.Color = RGB(153, 51, 0)
        If .Execute(Replace:=wdReplaceAll) = True Then
           GoTo Next101
        End If
    Next i
Next101:
    '==============================
    For i = 0 To UBound(HighLevel)
        .Text = HighLevel(i)
        .Replacement.Font.ColorIndex = wdGreen
        If .Execute(Replace:=wdReplaceAll) = True Then
           GoTo Next102
        End If
    Next i
Next102:
    '==============================
    For i = 0 To UBound(ConstanTxt)
        .Text = ConstanTxt(i)
        .Replacement.Font.Color = RGB(0, 0, 153)
        If .Execute(Replace:=wdReplaceAll) = True Then
           GoTo Next103
        End If
    Next i
Next103:
End With
End Sub
Sub tes()
TextComment (20)
End Sub
