If WScript.Arguments.Unnamed.Count < 1 Then
    usage()
Else 
    main()
End If

' �g����
Function usage() 
    With WScript
        .Echo "�G�N�X�|�[�g���ꂽ�}�N�����w�肵��Excel�t�@�C���ɃC���|�[�g���܂�"
        .Echo "����:���̃}�N���͑S�č폜����܂�"
        .Echo ""
        .Echo WScript.ScriptName & " <input_excel_files...>"
        .Echo ""
        .Echo "/MACRO_DIR" & Chr(9) & "�}�N���i�[�f�B���N�g�����w�肵�܂��B(�f�t�H���g�̓J�����g�f�B���N�g��)"
        .Echo Chr(9) & Chr(9) & "��: /MACRO_DIR:.\macro"
        .Echo Chr(9) & Chr(9) & "�@�w�肵���f�B���N�g���z���� vba_<excel_file_name> �̃f�B���N�g�����܂܂�Ă���K�v������܂�"
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

' �}�N���i�[��f�B���N�g�����擾
' �����Ɏw�肳�ꂽ�l
Function macroRootDirectory()
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = WScript.Arguments.Named.Item("MACRO_DIR")

    If dir = "" Then
        dir = "."
    End If

    dir = fso.GetAbsolutePathName(dir)
    macroRootDirectory = dir
End Function

' �t�@�C���ʃ}�N���i�[��f�B���N�g�����t�@�C��������쐬
Function fileMacroDirectory(inputExcelName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileMacroDirectory = macroRootDirectory() & "\vba_" & fso.GetFileName(inputExcelName) 
End Function

' �G�N�Z���t�@�C�����J��
Function openWorkBooks(excelFile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set excelApp = CreateObject("Excel.Application")
    fileName = fso.GetAbsolutePathName(excelFile)
    updateLinks = 0
    readOnly = False
    IgnoreReadOnlyRecommended = True
    set openWorkBooks = excelApp.Workbooks.Open(fileName,updateLinks,readOnly,,,,IgnoreReadOnlyRecommended)
End Function

' ���W���[�����o�͑Ώۂ��ǂ����m�F����
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
