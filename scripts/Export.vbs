If WScript.Arguments.Unnamed.Count < 1 Then
    usage()
Else 
    main()
End If

' �g����
Function usage() 
    With WScript
        .Echo "�}�N�����܂܂�Ă���G�N�Z���t�@�C���̕W�����W���[���ƃN���X���W���[�����G�N�X�|�[�g���܂�"
        .Echo ""
        .Echo WScript.ScriptName & " <input_excel_files...>"
        .Echo ""
        .Echo "/OUT_DIR" & Chr(9) & "�o�͐�f�B���N�g�����w�肵�܂��B(�f�t�H���g�̓J�����g�f�B���N�g��)"
        .Echo Chr(9) & Chr(9) & "��: /OUT_DIR:.\macro"
    End With
End Function

' �R���|�[�l���g�^�C�v����g���q���擾����
Function ext(cType)
    Select Case cType
        Case 1      : Ext = ".bas"
        Case 3      : Ext = ".frm"
        Case Else   : Ext = ".cls"
    End Select
End Function

' �o�͐惋�[�g�f�B���N�g�����擾
' �����Ɏw�肳�ꂽ�l
Function outputDirectory()
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = WScript.Arguments.Named.Item("OUT_DIR")

    If dir = "" Then
        dir = "."
    End If

    dir = fso.GetAbsolutePathName(dir)
    outputDirectory = dir
End Function

' �ċA�I�Ƀf�B���N�g�����쐬
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

' �}�N���t�@�C���o�͐�f�B���N�g�����t�@�C��������쐬
Function macroOutputDirectory(inputExcelName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    macroOutputDirectory = outputDirectory() & "\vba_" & fso.GetFileName(inputExcelName) 
End Function

' �G�N�Z���t�@�C�����J��
Function openWorkBooks(excelFile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")
    fileName = fso.GetAbsolutePathName(excelFile)
    updateLinks = 0
    readOnly = True
    IgnoreReadOnlyRecommended = True
    set openWorkBooks = excelApp.Workbooks.Open(fileName,updateLinks,readOnly,,,,IgnoreReadOnlyRecommended)
End Function

' �w�肵���f�B���N�g�������݂��邩�`�F�b�N����
' ���݂��Ȃ�������f�B���N�g�����쐬����
Function checkOutputDirectoryAndCreate(directory)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(directory) And Not fso.FileExists(directory) Then
        CreateFolderRecursive(directory)
    End If
End Function

' ���W���[�����o�͑Ώۂ��ǂ����m�F����
Function isOutputComponentType(component)
    If component.Type >= 1 And component.Type <= 3 Then
        isOutputComponentType = True
    Else
        isOutputComponentType = False
    End If
End Function

' �}�N���t�@�C�����G�N�X�|�[�g����
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