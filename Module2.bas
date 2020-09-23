Attribute VB_Name = "Module1"
Public PCurrtPWd As String
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub pCompactDB()
On Error GoTo 100
                            '*** This is a 1 step Compacting DB and returning to Application
                            Unload Form1
DBEngine.CompactDatabase App.Path & "\Doctors.mdb", App.Path & "\Temp.mdb", , , ";pwd=" & PCurrtPWd
FileCopy App.Path & "\Temp.mdb", App.Path & "\Doctors.mdb"
                            Form1.Show
                            If PCurrtPWd <> "" Then Form1.txtPaswd = PCurrtPWd: Form1.CmdOK_Click
                            If Dir(App.Path & "\Temp.mdb") <> "" Then
                            Kill App.Path & "\Temp.mdb"
                            MsgBox "DataBase Compacted"
                            End If


Exit Sub
100
If Dir(App.Path & "\Temp.mdb") <> "" Then
Kill App.Path & "\Temp.mdb"
End If
Form1.Show
MsgBox Err.Description
End Sub

Sub pBackupDB()
                            'This is an easy way to Backup your DataBase. Just Copy the DataBase to a File
                            'or compact the DB to a File
                            'or you can write Code to Copy all records to another File
                            'Also you may Add a Date or No. at the End of the db Backed name and give the user to choose which Backed to restore
                            Unload Form1
FileCopy App.Path & "\Doctors.mdb", "DoctorsBackup.mdb"
                            Form1.Show
                            If PCurrtPWd <> "" Then Form1.txtPaswd = PCurrtPWd: Form1.CmdOK_Click
                            MsgBox "DataBase Backed Up to DoctorsBackup.mdb"

End Sub

Sub pRestoreDB()
'On Error GoTo 100
                            'This is easy way to Restore your DataBase. Just Copy Back the Backed up File
                            'or you can write Code to delete all records and copy Back all records from the backup file or any other Code
                            If MsgBox("Restoring your DataBase will Override your existing one. Do you want to Proceed with Restore?", vbApplicationModal + vbYesNo) = vbNo Then Exit Sub
                            Unload Form1
FileCopy App.Path & "\DoctorsBackup.mdb", App.Path & "\Doctors.mdb"
                            Form1.Show
                            If PCurrtPWd <> "" Then Form1.txtPaswd = PCurrtPWd: Form1.CmdOK_Click
                            MsgBox "DataBase Restored"
                            Exit Sub

100
If Dir(App.Path & "\Temp.mdb") <> "" Then
Kill App.Path & "\Temp.mdb"
End If
Form1.Show
MsgBox Err.Description
End Sub

Sub pChangeDataBasePswd()
On Error GoTo 100

                                    '*** This is a 1 step Changing DataBase Pswd and returning to Application
                                    Unload Form1
DBEngine.CompactDatabase App.Path & "\Doctors.mdb", App.Path & "\Temp.mdb", ";pwd=" & Form15.Text2.Text, , ";pwd=" & Form15.Text1.Text
FileCopy App.Path & "\Temp.mdb", App.Path & "\Doctors.mdb"
                                    'This just will store the Pwd in a File
                                    Open App.Path & "\DB_Pwd.text" For Output As #1
                                    Print #1, Form15.Text2.Text
                                    Close #1
                                    PCurrtPWd = Form15.Text2.Text
                                    'Make a Backup after Pwd Change
                                    FileCopy App.Path & "\Doctors.mdb", "DoctorsBackup.mdb"
                                    Unload Form15
                                    Form1.Show
                                    Form1.txtPaswd = PCurrtPWd: Form1.CmdOK_Click
                                    If Dir(App.Path & "\Temp.mdb") <> "" Then
                                    Kill App.Path & "\Temp.mdb"
                                    End If
                                    MsgBox "Password Changed to '" & PCurrtPWd & "' Make sure you Write it Down"
                                    

Exit Sub
100
If Dir(App.Path & "\Temp.mdb") <> "" Then
Kill App.Path & "\Temp.mdb"
End If
Form1.Show
Unload Form15
MsgBox Err.Description
End Sub
