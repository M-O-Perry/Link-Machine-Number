Attribute VB_Name = "Module1"
Sub HyperlinkAllNumbers()
'
' Macro1 Macro
'
    
    Dim StrFile As String
    Dim itmnum As String
    Dim firstfolder As String
    Dim secondfolder As String
    Dim dirstring As String
    Dim itemnum As String
    
    
    Dim sheet As Worksheet
    
    Set sheet = ThisWorkbook.ActiveSheet
    
    Dim rng As Range
    Set rng = sheet.UsedRange
    
    
    Set rng = Range("B1:B" & (rng.Rows.Count + 2))
    
    
    Dim cell As Range
    
    For Each cell In rng
        itmnum = cell.Value
        
        If InStr(itmnum, "-") > 0 Then
            itmnum = Left(itmnum, 4) & Right(itmnum, 4)
        End If
        
        If Len(itmnum) = 8 And InStr(itmnum, "-") = 0 Then
            
            firstfolder = Left(itmnum, 2)
            secondfolder = Left(itmnum, 4)
            
            If firstfolder = "16" Or firstfolder = "17" Then
            
                dirstring = "E:\FINAL\text\" & secondfolder & "\" & itmnum & ".DOC"
            
            Else
                dirstring = "E:\FINAL\text\" & firstfolder & "XX\" & itmnum & ".DOC"
            
            End If
            
            StrFile = Dir(dirstring)
            
            If Len(StrFile) > 0 Then
                Application.ActiveSheet.Hyperlinks.Add Anchor:=cell.Offset(0, 0), Address:=dirstring
            
            End If
            
        
        
        End If
    
    
    Next cell
End Sub

