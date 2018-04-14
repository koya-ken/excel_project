If WScript.Arguments.Unnamed.Count < 1 Then
    usage()
Else 
    main()
End If

' 使い方
Function usage() 
    With WScript
        .Echo "マクロが含まれているエクセルファイルの標準モジュールとクラスモジュールをエクスポートします"
        .Echo ""
        .Echo WScript.ScriptName & " <input_excel_files...>"
        .Echo ""
        .Echo "/OUT_DIR" & Chr(9) & "出力先ディレクトリを指定します。(デフォルトはカレントディレクトリ)"
        .Echo Chr(9) & Chr(9) & "例: /OUT_DIR:.\macro"
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

' 出力先ルートディレクトリを取得
' 引数に指定された値
Function outputDirectory()
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = WScript.Arguments.Named.Item("OUT_DIR")

    If dir = "" Then
        dir = "."
    End If

    dir = fso.GetAbsolutePathName(dir)
    outputDirectory = dir
End Function

' 再帰的にディレクトリを作成
Function CreateFolderRecursive(FullPath)
  Dim arr, dir, path
  Dim oFs

  Set oFs = WScript.CreateObject("Scripting.FileSystemObject")
  arr = split(FullPath, "\")
  path = ""
  For Each dir In arr
    If path <> "" Then path = path & "\"
    path = path & dir
    If oFs.FolderExists(path) = False Then oFs.CreateFolder(path)
  Next
End Function

' マクロファイル出力先ディレクトリをファイル名から作成
Function macroOutputDirectory(inputExcelName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    macroOutputDirectory = outputDirectory() & "\vba_" & fso.GetFileName(inputExcelName) 
End Function

' エクセルファイルを開く
Function openWorkBooks(excelFile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")
    fileName = fso.GetAbsolutePathName(excelFile)
    updateLinks = 0
    readOnly = True
    IgnoreReadOnlyRecommended = True
    set openWorkBooks = excelApp.Workbooks.Open(fileName,updateLinks,readOnly,,,,IgnoreReadOnlyRecommended)
End Function

' 指定したディレクトリが存在するかチェックする
' 存在しなかったらディレクトリを作成する
Function checkOutputDirectoryAndCreate(directory)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(directory) And Not fso.FileExists(directory) Then
        CreateFolderRecursive(directory)
    End If
End Function

' モジュールが出力対象かどうか確認する
Function isOutputComponentType(component)
    If component.Type >= 1 And component.Type <= 3 Then
        isOutputComponentType = True
    Else
        isOutputComponentType = False
    End If
End Function

' マクロファイルをエクスポートする
Function exportMacroFile(component,directory)
    component.Export directory & "\" & component.Name & ext(component.Type)
End Function

Function main()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")

    Set objArgs = WScript.Arguments.Unnamed
    For Each arg In objArgs
        WScript.Echo "processing " & arg & ".."
        macroDirectory = macroOutputDirectory(arg)
        set workBooks = openWorkBooks(arg)

        If workBooks.VBProject.Protection = 0 Then
            For Each component In workBooks.VBProject.VBComponents
                If isOutputComponentType(component) Then
                    checkOutputDirectoryAndCreate(macroDirectory)
                    call exportMacroFile (component,macroDirectory)
                End If
            Next    
        End If

        workBooks.Close False
    Next
    
    excelApp.Quit
End Function