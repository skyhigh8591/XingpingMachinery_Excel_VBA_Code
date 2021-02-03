Attribute VB_Name = "Module7"
Sub 複製貼上自訂貼上數量()
'
'
'
    Dim mypath As String, myFile As String, f As String, myname As String, n, s, I, ii As Integer, myFolder As FileDialog, MySheet As Worksheet, V, SS, ZZ As String
    
    
    MsgBox "選擇複製資料清單"
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
            .Show
                f = GetName(.SelectedItems(1))
                Workbooks.Open (.SelectedItems(1))
    End With
    MsgBox "選擇品目欄資料夾"
    Set myFolder = Application.FileDialog(msoFileDialogFolderPicker)
        myFolder.InitialFileName = ThisWorkbook.Path
        myFolder.Show
        myFile = "*.xls"
         
         
          '修改地方
         V = "R"
         SS = 17
         ZZ = 3
         



'V = Application.InputBox(Prompt:="NO欄位英文", Type:=1 + 2)




         
          mFile = Dir(myFolder.SelectedItems.Item(1) & "\" & "*.XLS")
            Do While mFile <> ""
                Workbooks.Open Filename:=myFolder.SelectedItems.Item(1) & "\" & mFile
                o = mFile
                mFile = Dir()
                
            Windows(f).Activate
            Sheets(1).Select

    For I = 3 To [C65536].End(xlUp).Row
            If Cells(I, "A") > 0 Then
            Windows(f).Activate
            Sheets(1).Select
            Range("A" & I & ":" & "L" & I).Select
            Selection.Copy
            
            Windows(o).Activate
            Sheets(2).Select
            Range("C14").Select
            Selection.End(xlDown).Select
            RR = Selection.Row
            col = Selection.Column
            Cells(RR + 1, col).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            

            
            ElseIf Cells(I, "C") = 0 Then
            Exit For
            End If
            
         
            
            
            
          
            Windows(f).Activate
            Sheets(1).Select
            For ii = 3 To [C65536].End(xlUp).Row
            Windows(f).Activate
            Sheets(1).Select
            
            If Cells(ZZ, SS) > 0 Then
            Windows(f).Activate
            Sheets(1).Select
            Cells(ZZ, SS).Select
            Selection.Copy
            
            Windows(o).Activate
            Sheets(2).Select
            '修改地方
            Range("T15").Select
            Selection.End(xlDown).Select
            TT = Selection.Row
            col = Selection.Column
            Cells(TT + 1, col).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            ZZ = ZZ + 1
            Exit For
            
            
            ElseIf Cells(ZZ, SS) = "" Then
            Exit For
            End If
            Next
            
            
            Windows(o).Activate
            Sheets(2).Select
            Range(V & "16" & ":" & V & "17").Select
            Selection.AutoFill Destination:=Range(V & "16" & ":" & V & "65"), Type:=xlFillDefault
            Range(V & "16" & ":" & V & "65").Select
            ActiveWindow.SmallScroll Down:=-63
            Range(V & "16").Select
                                                                 
            Windows(f).Activate
            Sheets(1).Select
            Range("A2").Select
            Application.CutCopyMode = False
            
            
    Next
    ZZ = 3
    SS = SS + 1
    
    
    
    
    
    
    
        
    
            
            
            

    
  

    
    'Windows(o).Activate
    'ActiveWorkbook.Save
    'ActiveWorkbook.Close False
    
    Loop
                
                    
    
End Sub

Function GetName(rFile) As String
   
    Dim nameArr As Variant
    nameArr = Split(rFile, "\")
    GetName = nameArr(UBound(nameArr))

End Function



