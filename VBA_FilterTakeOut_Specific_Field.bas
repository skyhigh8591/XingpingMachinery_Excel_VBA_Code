Attribute VB_Name = "Module1"
Sub P裁切()
'
'
'
    Dim mypath As String, myFile As String, f As String, myname As String, n, s, I As Integer, myFolder As FileDialog, MySheet As Worksheet
    
    
    MsgBox "選擇新的裁切LIST檔案(P)"
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
         
         
          mFile = Dir(myFolder.SelectedItems.Item(1) & "\" & "*.XLS")
            Do While mFile <> ""
                Workbooks.Open Filename:=myFolder.SelectedItems.Item(1) & "\" & mFile
                o = mFile
                mFile = Dir()
   
                        Windows(o).Activate
                        Sheets(2).Select
                        Range("K15").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.AutoFilter
                        
                 
                         
                        ActiveSheet.Range("$K$15:$K$92").AutoFilter Field:=1, Criteria1:=Array( _
                        "PROFILE", "PROFILE-K"), Operator:=xlFilterValues
                            
                            
                            
                            
                            
                        ActiveWindow.SmallScroll Down:=-15
                        Range("C15").Select
                        Selection.End(xlDown).Select
                        R = Selection.Row
                        col = Selection.Column
                        Cells(R, col).Select
                If Cells(R, col) > 0 Then
                        Range("C16:J16").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                        Windows(f).Activate
                        Sheets(2).Select
                        ActiveWindow.SmallScroll Down:=-12
                        Range("A1").Select
                        Range("A1").Select
                        Selection.End(xlDown).Select
                        R = Selection.Row
                        col = Selection.Column
                        Cells(R + 1, col).Select
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        Windows(o).Activate
                        Sheets(2).Select
                        Range("A2").Select
                        Application.CutCopyMode = False
                                      Else
                                      
                                      
                                      
                End If
                ActiveWorkbook.Close False
                Loop
       
  
    
    Dim y As Integer
    Windows(f).Activate
        Sheets(2).Select
            Range("C3").Select
                y = 1
                    For y = 1 To 9999
                        If Cells(y, "C") Like "*手配*" Then
                            Cells(y, "C").Select
                            Selection.ClearContents
                        End If
                    Next
                    
    
End Sub
Function GetName(rFile) As String
   
    Dim nameArr As Variant
    nameArr = Split(rFile, "\")
    GetName = nameArr(UBound(nameArr))

End Function


