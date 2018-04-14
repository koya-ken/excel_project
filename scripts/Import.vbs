If WScript.Arguments.Unnamed.Count < 1 Then
    usage()
Else 
    main()
End If

' 使い方
Function usage() 
    With WScript
        .Echo "エクスポートされたマクロを指定したExcelファイルにインポートします"
        .Echo "注意:元のマクロは全て削除されます"
        .Echo ""
        .Echo WScript.ScriptName & " <input_excel_files...>"
        .Echo ""
        .Echo "/MACRO_DIR" & Chr(9) & "マクロ格納ディレクトリを指定します。(デフォルトはカレントディレクトリ)"
        .Echo Chr(9) & Chr(9) & "例: /MACRO_DIR:.\macro"
        .Echo Chr(9) & Chr(9) & "　指定したディレクトリ配下に vba_<excel_file_name> のディレクトリが含まれている必要があります"
    End With
End Function

' コンポーネントタイプから拡張子を取得する
Function ext(cType)
    Select Case cType
        Case 1      : Ext = ".bas"
        Case 3      : Ext = ".frm"
        Case Else   : Ext = ".cls"
    End Select
End Function

' マクロ格納先ディレクトリを取得
' 引数に指定された値
Function macroRootDirectory()
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = WScript.Arguments.Named.Item("MACRO_DIR")

    If dir = "" Then
        dir = "."
    End If

    dir = fso.GetAbsolutePathName(dir)
    macroRootDirectory = dir
End Function

' ファイル別マクロ格納先ディレクトリをファイル名から作成
Function fileMacroDirectory(inputExcelName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileMacroDirectory = macroRootDirectory() & "\vba_" & fso.GetFileName(inputExcelName) 
End Function

' エクセルファイルを開く
Function openWorkBooks(excelFile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")
    fileName = fso.GetAbsolutePathName(excelFile)
    updateLinks = 0
    readOnly = False
    IgnoreReadOnlyRecommended = True
    set openWorkBooks = excelApp.Workbooks.Open(fileName,updateLinks,readOnly,,,,IgnoreReadOnlyRecommended)
End Function

' モジュールが出力対象かどうか確認する
Function isOutputComponentType(component)
    If component.Type >= 1 And component.Type <= 3 Then
        isOutputComponentType = True
    Else
        isOutputComponentType = False
    End If
End Function

Function main()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")
    Set objArgs = WScript.Arguments.Unnamed

    For Each arg In objArgs
        WScript.Echo "processing " & arg & ".."

        macroDirectory = fileMacroDirectory(arg)
        set workBooks = openWorkBooks(arg)

        If workBooks.VBProject.Protection = 0 Then
            For Each component In workBooks.VBProject.VBComponents
                If isOutputComponentType(component) Then
                    call workBooks.VBProject.VBComponents.Remove(component)
                End If
            Next    
        End If
        workBooks.Save

        If workBooks.VBProject.Protection = 0 Then
            Set directory = fso.GetFolder(macroDirectory)
            For Each file in directory.files
                scriptFile = file.parentfolder & "\" & file.name
                WScript.Echo "import " & file.name
                workBooks.VBProject.VBComponents.Import scriptFile
            Next
        End If

        workBooks.Save
        workBooks.Close False
    Next

    excelApp.Quit
End Function
