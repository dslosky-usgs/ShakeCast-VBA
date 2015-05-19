Attribute VB_Name = "DevTools"
Sub SaveCodeModules(dirStr As String)

'This code Exports all VBA modules
Dim moduleName As String
Dim vbaType As Integer

With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.count
        If .VBComponents(i).CodeModule.CountOfLines > 0 Then
            moduleName = .VBComponents(i).CodeModule.Name
            vbaType = .VBComponents(i).Type
            
            If vbaType = 1 Then

                .VBComponents(i).Export dirStr & moduleName & ".vba"
            ElseIf vbaType = 3 Then

                .VBComponents(i).Export dirStr & moduleName & ".frm"
            ElseIf vbaType = 100 Then

                .VBComponents(i).Export dirStr & moduleName & ".cls"
            End If
            
        End If
    Next i
End With

End Sub

Sub ImportCodeModules(dirStr As String)

Dim modList(0 To 0) As String
Dim vbaType As Integer

With ThisWorkbook.VBProject
    'For i% = 1 To .VBComponents.count
    For Each comp In .VBComponents
    
        'modulename = .VBComponents(i%).CodeModule.Name
        moduleName = comp.CodeModule.Name
        
        vbaType = .VBComponents(moduleName).Type
        
        If moduleName <> "DevTools" Then
            If vbaType = 1 Or _
                vbaType = 3 Then
                
                .VBComponents.Remove .VBComponents(moduleName)
                
            ElseIf vbaType = 100 Then
                .VBComponents(moduleName).CodeModule.DeleteLines 1, .VBComponents(moduleName).CodeModule.CountOfLines
            End If
        End If
    Next comp
End With

' make a list of files in the target directory

Dim varDir As Variant
varDir = dir(dirStr, vbNormal)


With ThisWorkbook.VBProject
    'For Each moduleName In dirContents.Files
    
    While (varDir <> "")
        If varDir <> "DevTools.vba" Then
            If Right(varDir, 4) = ".vba" Or _
                Right(varDir, 4) = ".frm" Then
                .VBComponents.Import dirStr & varDir
                
            ElseIf Right(varDir, 4) = ".cls" Then
                Dim r As Integer
                Dim fullmoduleString As String
                Open dirStr & varDir For Input As #1
                
                r = 0
                fullmoduleString = ""
                Do Until EOF(1)
                    Line Input #1, moduleString
                    If r > 8 Then
                        If Right(moduleString, 1) = "_" Then
                            fullmoduleString = fullmoduleString & moduleString & vbNewLine
                        Else
                            fullmoduleString = fullmoduleString & moduleString & vbNewLine
                        End If
                    End If
                    r = r + 1
                Loop
                .VBComponents(replace(varDir, ".cls", "")).CodeModule.InsertLines .VBComponents(replace(varDir, ".cls", "")).CodeModule.CountOfLines + 1, fullmoduleString
                        
                Close #1
                
            End If
        End If
        
        varDir = dir
        
    Wend
    ' Next moduleName
End With

End Sub

Sub SaveCodeModulesWork()
    SaveCodeModules "Macintosh HD:Users:dslosky:Documents:stuff:Jobs:Worksheet Optimization:ShakeCast-VBA:source:"
End Sub

Sub ImportCodeModulesWork()
    ImportCodeModules "Macintosh HD:Users:dslosky:Documents:stuff:Jobs:Worksheet Optimization:ShakeCast-VBA:source:"
End Sub

