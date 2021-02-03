Attribute VB_Name = "Module6"
Sub �ק�P�˪����()
Attribute �ק�P�˪����.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

Dim n, s, y, t As Integer

n = 49
y = 451

MsgBox "��ܫ~�����Ƨ�"
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

'��J�d��W


 


           
           
           
           


                
                
                
                
 '��J�d��U
                
               ' ActiveWorkbook.Save
               ' ActiveWorkbook.Close False
                Loop

End Sub

