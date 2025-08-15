Attribute VB_Name = "ChromeDriverAutoUpdateModule"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long
                           
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpString As Any, _
                          ByVal lpFileName As String) As Long
                           
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                        (ByVal pCaller As Long, _
                         ByVal szURL As String, _
                         ByVal szFileName As String, _
                         ByVal dwReserved As Long, _
                         ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" _
                        (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long
                           
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpString As Any, _
                          ByVal lpFileName As String) As Long
                           
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                        (ByVal pCaller As Long, _
                         ByVal szURL As String, _
                         ByVal szFileName As String, _
                         ByVal dwReserved As Long, _
                         ByVal lpfnCB As Long) As Long
Private Declare Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" _
                        (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
#End If

Private Type VersionType '---�{���̓N���X�I�u�W�F�N�g�ɂ��������ǂ��ꂾ���̂��߂Ƀ��W���[����肽���Ȃ�
    Major As Long
    Minor As Long
    Build As Long
    Revision As Long
    BuildVersion As String
    RevisionVersion As String
End Type
Const ZIP_FILE As String = "chromedriver.zip"
Dim workPath As String
Dim objFso As New Scripting.FileSystemObject
Public Function ChromeDriverAutoUpdate(Optional ByVal ForcedExecution As Boolean = False) As Boolean
'====================================================================================================
'chrome.exe��chromedriver.exe�̃o�[�W�������r����chromedriver�������X�V����
'�������͋������s�t���O�iForcedExecution�j��True�ł����s����
'====================================================================================================
    Dim chromePath As String '---chrome.exe���ۑ�����Ă���p�X
    Dim chromeFullpath As String '---chrome.exe�܂Ŋ܂߂��t���p�X
    Dim chromeVersion As VersionType
    Dim chromedriverPath As String
    Dim chromedriverFullPath As String
    Dim objFolder As Scripting.Folder
    Dim lngRevision As Long
    Dim targetRevision As Long
    
    ' ---chromedriver���_�E�����[�h�p�̃t�H���_���쐬����@��Python�ɍ��킹�Ă�
    workPath = Environ("USERPROFILE") & "\.cache\selenium\seleniumbasic"
    Select Case SHCreateDirectoryEx(0&, workPath, 0&)
        Case 0:
            ' ---�쐬����
        Case 183
            ' ---�쐬�ς�
        Case Else:
            ' ---�쐬�ł��Ȃ�������
            MsgBox "�_�E�����[�h�p�t�H���_���쐬�ł��܂���ł���" & vbCrLf & Error(Err), vbCritical
    End Select
    
    '---chrome�{�̂̃t�H���_��T��
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramW6432") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramFiles") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\Google\Chrome\Application")
            chromePath = Environ("LOCALAPPDATA") & "\Google\Chrome\Application"
        Case Else
            MsgBox "'chrome'�t�H���_��������܂���", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---�O�̂���chrome.exe���m�F����
    If objFso.FileExists(chromePath & "\chrome.exe") = True Then
        chromeFullpath = chromePath & "\chrome.exe"
    Else
        MsgBox "'chrome.exe'��������܂���", vbCritical
        Exit Function
    End If
    
    '---SeleniumBasic�̃t�H���_��T��
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramW6432") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramFiles") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\SeleniumBasic")
            chromedriverPath = Environ("LOCALAPPDATA") & "\SeleniumBasic"
        Case Else
            MsgBox "'SeleniumBasic'�̃t�H���_��������܂���", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---�O�̂���chromedriver.exe���m�F����
    If objFso.FileExists(chromedriverPath & "\chromedriver.exe") = True Then
        chromedriverFullPath = chromedriverPath & "\chromedriver.exe"
    Else
        MsgBox "'chromedriver.exe'��������܂���", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    Set objFso = Nothing
        
    '---chrome.exe�̃o�[�W�������擾����
    If GetChromeVersion(chromeFullpath, chromeVersion) = False Then '---chrome.exe�̃o�[�W�������擾����
        MsgBox "'chrome.exe'�̃o�[�W�������擾�ł��܂���ł���", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---chrome.exe�̃o�[�W�����ɍ��킹��chromedriver.exe���_�E�����[�h����
    If ChromedriverCheck(chromedriverPath, chromeVersion) = False Then '---chromedriver�̃o�[�W�������擾����
        MsgBox "'chromedriver.exe'�̃o�[�W�������擾�ł��܂���ł���", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---���p����chromedriver�̃o�[�W������T�������r�W������chrome�ȉ��̃o�[�W�����̒��ōŐV���g��
    On Error GoTo ErrLabel
        targetRevision = 0
        For Each objFolder In objFso.GetFolder(workPath).SubFolders
            If objFolder.Name Like chromeVersion.BuildVersion & "*" Then
                lngRevision = CLng(Split(objFolder.Name, ".")(3))
                If lngRevision <= chromeVersion.Revision Then
                    If targetRevision < lngRevision Then
                        targetRevision = lngRevision
                    End If
                End If
            Else
                Call objFolder.Delete(True)
            End If
        Next
        Call objFso.GetFile(workPath & "\" & chromeVersion.BuildVersion & "." & CStr(targetRevision) & "\chromedriver.exe").Copy(chromedriverPath & "\chromedriver.exe", True)
    On Error GoTo 0
    
    '---���ʂƂ��čX�V���Ă��Ȃ��ꍇ�����邪�A�X�V���s����Ȃ��čX�V�s�v�Ȕ��肾����True��Ԃ�
    ChromeDriverAutoUpdate = True
Exit Function
ErrLabel:     '---�\�����ʃG���[�̕���
    MsgBox "chromedriver �̓��ւɎ��s���܂���" & vbCrLf & Error(Err) & vbCrLf & "�����̉�ʂ̃L���v�`�����쐬�҂֑����Ă�������"
    ChromeDriverAutoUpdate = False
End Function
Private Function GetChromeVersion(ByVal chromeFullpath As String, ByRef chromeVersion As VersionType) As Boolean
'====================================================================================================
'PowerShell��chrome.exe�̃o�[�W���������擾����@����uPowerShell�������オ��
'====================================================================================================
    Dim command As String
    Dim objRet As Object
    
    On Error GoTo ErrLabel
        '---chrome�o�[�W�������̏����l
        chromeVersion.Major = 1
        chromeVersion.Minor = 0
        chromeVersion.Build = 0
        chromeVersion.Revision = 0
        '---chrome.exe�̃o�[�W�������擾����PowerShell�R�}���h
        command = "powershell.exe -NoProfile -ExecutionPolicy Bypass (Get-Item -Path '" & chromeFullpath & "').VersionInfo.FileVersion"
        '---PowerShell�̎��s���ʂ��Z�b�g
        Set objRet = CreateObject("WScript.Shell").Exec(command)
        '---PowerShell�̃R�}���h���b�g�̎��s���ʂ��擾
        chromeVersion.RevisionVersion = Trim(objRet.StdOut.ReadAll)
        '---���̎擾���I�������I�u�W�F�N�g���N���A����
        Set objRet = Nothing
        '---���s�R�[�h���܂܂�Ă��邩��폜����
        chromeVersion.RevisionVersion = Trim(Replace(Replace(Replace(chromeVersion.RevisionVersion, vbCrLf, vbNullString), vbCr, vbNullString), vbLf, vbNullString))
        '---�o�[�W�������𕪂��ĕԂ�
        With CreateObject("VBScript.RegExp") '---���K�\���̏���
            .Pattern = "\d+\.\d+\.\d+(\.\d+)?"
            .Global = True
            If .test(chromeVersion.RevisionVersion) Then '---�O�̂��ߐ��K�\���Ńo�[�W���������`�F�b�N����
                chromeVersion.Major = CLng(Split(chromeVersion.RevisionVersion, ".")(0))
                chromeVersion.Minor = CLng(Split(chromeVersion.RevisionVersion, ".")(1))
                chromeVersion.Build = CLng(Split(chromeVersion.RevisionVersion, ".")(2))
                If UBound(Split(chromeVersion.RevisionVersion, ".")) >= 3 Then chromeVersion.Revision = CLng(Split(chromeVersion.RevisionVersion, ".")(3)) '---���r�W�����ԍ�������΁���{����͂�
                chromeVersion.BuildVersion = Join(Array(chromeVersion.Major, chromeVersion.Minor, chromeVersion.Build), ".") '---���r�W������`�����V���[�g�o�[�W���������Z�b�g����
            Else '---���K�\���s��v�Ȃ玸�s�ŕԂ�
                GetChromeVersion = False
                Exit Function
            End If
        End With
        GetChromeVersion = True
    On Error GoTo 0
    Exit Function
ErrLabel:     '---�\�����ʃG���[�̕���
    MsgBox "chrome.exe �̃o�[�W�������擾�Ɏ��s���܂���" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "�����̉�ʂ̃L���v�`�����쐬�҂֑����Ă�������"
    GetChromeVersion = False
End Function
Private Function ChromedriverCheck(chromedriverPath, chromeVersion As VersionType) As Boolean
    Dim rc As Long
    Dim url As String
    Dim objHttp As New MSXML2.XMLHTTP60
    Dim objRet As Scripting.Dictionary
    Dim objVersion As Scripting.Dictionary
    Dim idx As Variant
    Dim chromedriver As Variant
    
    Const JSON_ENDPOINTS_URL As String = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
    Const TARGET_PLATFORM As String = "win64"

    On Error GoTo ErrLabel
        With objHttp
            .Open "GET", JSON_ENDPOINTS_URL, False
            .Send
            Set objRet = JsonConverter.ParseJson(.responseText) '---JSON endpoints ����������W����
            '---�o�[�W���������ƍ����Ȃ���擾�Ώۃ`�F�b�N����
            For Each objVersion In objRet("versions")
                '---�r���h�܂ň�v���Ă���Ύ擾���ɂ���
                If objVersion("version") Like chromeVersion.BuildVersion & "*" Then
                    '---�Â�����chromedriver���C���f�b�N�X�ɂȂ��ꍇ������̂ŔO�̂��߃C���f�b�N�X�`�F�b�N����
                    For Each idx In objVersion("downloads")
                        If idx = "chromedriver" Then
                            For Each chromedriver In objVersion("downloads")(idx)
                                '---�擾�ς݃t�H���_�ɂȂ�������
                                If objFso.FolderExists(workPath & "\" & objVersion("version")) = False Then
                                    '---�Ώۂ�platform���`�F�b�N���� ��Winsows11����64bit�����Ȃ��̂Ō��ߑł�
                                    If chromedriver("platform") = TARGET_PLATFORM Then
                                        url = chromedriver("url")
                                        If DownloadChromedriver(url, objVersion("version")) = False Then
                                            MsgBox "chromedriver�̃_�E�����[�h�Ɏ��s���܂���", vbCritical
                                            ChromedriverCheck = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            Set objRet = Nothing
        End With
    On Error GoTo 0
    ChromedriverCheck = True
    Exit Function
ErrLabel:     '---�\�����ʃG���[�̕���
    MsgBox "chromedriver.exe �̍X�V�Ɏ��s���܂���" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "�����̉�ʂ̃L���v�`�����쐬�҂֑����Ă�������"
    ChromedriverCheck = False
End Function
Private Function DownloadChromedriver(ByVal url As String, targetVersion As String) As Boolean
    Dim rc As Long
    Dim downloadPath As String
    Dim newDriverPath As String
    Dim objFolder As Scripting.Folder
    downloadPath = workPath & "\" & targetVersion
    ' ---chromedriver�̃t�H���_���쐬����
    Select Case SHCreateDirectoryEx(0&, downloadPath, 0&)
        Case 0:
            ' ---�쐬����
        Case 183
            ' ---�쐬�ς�
        Case Else:
            ' ---�쐬�ł��Ȃ�������
            MsgBox "ChromeDriver�p�t�H���_���쐬�ł��܂���ł���" & vbCrLf & Error(Err), vbCritical
            DownloadChromedriver = False
            Exit Function
    End Select
    
    '---�t�@�C�����_�E�����[�h����
    If URLDownloadToFile(0, url, workPath & "\" & ZIP_FILE, 0, 0) <> 0 Then
        MsgBox "ChromeDriver���_�E�����[�h�ł��܂���ł���" & vbCrLf & Error(Err), vbCritical
        DownloadChromedriver = False
        Exit Function
    End If
    Application.DisplayAlerts = False
    '---zip������̃t�H���_�Ɍ����ĉ𓀂���
    With CreateObject("Shell.Application") '---zip������̃t�H���_�Ɍ����ĉ𓀂���
        .Namespace((downloadPath)).CopyHere .Namespace((workPath & "\" & ZIP_FILE)).Items
    End With
    '--- �𓀂����t�H���_����ċN��������chromedriver.exe�̃t���p�X���擾����
    newDriverPath = SearchFilesRecursively(downloadPath & "\", "chromedriver.exe")
    If newDriverPath = "" Then
        MsgBox "chromedriver.exe �̍X�V�Ɏ��s���܂���"
        DownloadChromedriver = False
    End If
    '---chromedriver���o�[�W�����t�H���_�����Ɉړ�����
    Call objFso.MoveFile(newDriverPath, downloadPath & "\")
    '---chromedriver���Ȃ��Ȃ����s�v�t�H���_���폜����
    For Each objFolder In objFso.GetFolder(downloadPath).SubFolders
        objFolder.Delete True
    Next
    '---zip�t�@�C�����폜����
    Call objFso.DeleteFile(workPath & "\" & ZIP_FILE, True)
    Application.DisplayAlerts = True
    DownloadChromedriver = True
End Function
Function SearchFilesRecursively(ByVal folderPath As String, fileName) As String
'====================================================================================================
' folderPath���N�_�ɍċN�����ŃT�u�t�H���_�܂őΏۂɂ���fileName��T���ăt���p�X��Ԃ�
'====================================================================================================
    Dim objFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim objFile As Scripting.File
    
    ' �t�@�C���ꗗ�\��
    For Each objFile In objFso.GetFolder(folderPath).Files
        If objFile.Name = fileName Then
            SearchFilesRecursively = objFile.Path
            Exit Function
        End If
    Next objFile
    
    ' �T�u�t�H���_���ċA�I�ɒT��
    For Each subFolder In objFso.GetFolder(folderPath).SubFolders
        If SearchFilesRecursively(subFolder.Path, fileName) = subFolder.Path & "\" & fileName Then
            SearchFilesRecursively = subFolder.Path & "\" & fileName
            Exit Function
        End If
    Next subFolder
    SearchFilesRecursively = ""
End Function
