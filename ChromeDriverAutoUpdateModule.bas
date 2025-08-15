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

Private Type VersionType '---本当はクラスオブジェクトにしたいけどこれだけのためにモジュール作りたくない
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
'chrome.exeとchromedriver.exeのバージョンを比較してchromedriverを自動更新する
'もしくは強制実行フラグ（ForcedExecution）がTrueでも実行する
'====================================================================================================
    Dim chromePath As String '---chrome.exeが保存されているパス
    Dim chromeFullpath As String '---chrome.exeまで含めたフルパス
    Dim chromeVersion As VersionType
    Dim chromedriverPath As String
    Dim chromedriverFullPath As String
    Dim objFolder As Scripting.Folder
    Dim lngRevision As Long
    Dim targetRevision As Long
    
    ' ---chromedriverをダウンロード用のフォルダを作成する　※Pythonに合わせてる
    workPath = Environ("USERPROFILE") & "\.cache\selenium\seleniumbasic"
    Select Case SHCreateDirectoryEx(0&, workPath, 0&)
        Case 0:
            ' ---作成成功
        Case 183
            ' ---作成済み
        Case Else:
            ' ---作成できなかった時
            MsgBox "ダウンロード用フォルダを作成できませんでした" & vbCrLf & Error(Err), vbCritical
    End Select
    
    '---chrome本体のフォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramW6432") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramFiles") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\Google\Chrome\Application")
            chromePath = Environ("LOCALAPPDATA") & "\Google\Chrome\Application"
        Case Else
            MsgBox "'chrome'フォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---念のためchrome.exeを確認する
    If objFso.FileExists(chromePath & "\chrome.exe") = True Then
        chromeFullpath = chromePath & "\chrome.exe"
    Else
        MsgBox "'chrome.exe'が見つかりません", vbCritical
        Exit Function
    End If
    
    '---SeleniumBasicのフォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramW6432") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramFiles") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\SeleniumBasic")
            chromedriverPath = Environ("LOCALAPPDATA") & "\SeleniumBasic"
        Case Else
            MsgBox "'SeleniumBasic'のフォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---念のためchromedriver.exeを確認する
    If objFso.FileExists(chromedriverPath & "\chromedriver.exe") = True Then
        chromedriverFullPath = chromedriverPath & "\chromedriver.exe"
    Else
        MsgBox "'chromedriver.exe'が見つかりません", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    Set objFso = Nothing
        
    '---chrome.exeのバージョンを取得する
    If GetChromeVersion(chromeFullpath, chromeVersion) = False Then '---chrome.exeのバージョンを取得する
        MsgBox "'chrome.exe'のバージョンが取得できませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---chrome.exeのバージョンに合わせたchromedriver.exeをダウンロードする
    If ChromedriverCheck(chromedriverPath, chromeVersion) = False Then '---chromedriverのバージョンを取得する
        MsgBox "'chromedriver.exe'のバージョンが取得できませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---利用するchromedriverのバージョンを探す※リビジョンがchrome以下のバージョンの中で最新を使う
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
    
    '---結果として更新していない場合もあるが、更新失敗じゃなくて更新不要な判定だからTrueを返す
    ChromeDriverAutoUpdate = True
Exit Function
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chromedriver の入替に失敗しました" & vbCrLf & Error(Err) & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
    ChromeDriverAutoUpdate = False
End Function
Private Function GetChromeVersion(ByVal chromeFullpath As String, ByRef chromeVersion As VersionType) As Boolean
'====================================================================================================
'PowerShellでchrome.exeのバージョン情報を取得する　※一瞬PowerShellが立ち上がる
'====================================================================================================
    Dim command As String
    Dim objRet As Object
    
    On Error GoTo ErrLabel
        '---chromeバージョン情報の初期値
        chromeVersion.Major = 1
        chromeVersion.Minor = 0
        chromeVersion.Build = 0
        chromeVersion.Revision = 0
        '---chrome.exeのバージョンを取得するPowerShellコマンド
        command = "powershell.exe -NoProfile -ExecutionPolicy Bypass (Get-Item -Path '" & chromeFullpath & "').VersionInfo.FileVersion"
        '---PowerShellの実行結果をセット
        Set objRet = CreateObject("WScript.Shell").Exec(command)
        '---PowerShellのコマンドレットの実行結果を取得
        chromeVersion.RevisionVersion = Trim(objRet.StdOut.ReadAll)
        '---情報の取得が終わったらオブジェクトをクリアする
        Set objRet = Nothing
        '---改行コードが含まれているから削除する
        chromeVersion.RevisionVersion = Trim(Replace(Replace(Replace(chromeVersion.RevisionVersion, vbCrLf, vbNullString), vbCr, vbNullString), vbLf, vbNullString))
        '---バージョン情報を分けて返す
        With CreateObject("VBScript.RegExp") '---正規表現の準備
            .Pattern = "\d+\.\d+\.\d+(\.\d+)?"
            .Global = True
            If .test(chromeVersion.RevisionVersion) Then '---念のため正規表現でバージョン情報をチェックする
                chromeVersion.Major = CLng(Split(chromeVersion.RevisionVersion, ".")(0))
                chromeVersion.Minor = CLng(Split(chromeVersion.RevisionVersion, ".")(1))
                chromeVersion.Build = CLng(Split(chromeVersion.RevisionVersion, ".")(2))
                If UBound(Split(chromeVersion.RevisionVersion, ".")) >= 3 Then chromeVersion.Revision = CLng(Split(chromeVersion.RevisionVersion, ".")(3)) '---リビジョン番号があれば※基本あるはず
                chromeVersion.BuildVersion = Join(Array(chromeVersion.Major, chromeVersion.Minor, chromeVersion.Build), ".") '---リビジョンを覗いたショートバージョン情報をセットする
            Else '---正規表現不一致なら失敗で返す
                GetChromeVersion = False
                Exit Function
            End If
        End With
        GetChromeVersion = True
    On Error GoTo 0
    Exit Function
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chrome.exe のバージョン情報取得に失敗しました" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
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
            Set objRet = JsonConverter.ParseJson(.responseText) '---JSON endpoints から情報を収集する
            '---バージョン情報を照合しながら取得対象チェックする
            For Each objVersion In objRet("versions")
                '---ビルドまで一致していれば取得候補にする
                If objVersion("version") Like chromeVersion.BuildVersion & "*" Then
                    '---古い情報はchromedriverがインデックスにない場合があるので念のためインデックスチェックする
                    For Each idx In objVersion("downloads")
                        If idx = "chromedriver" Then
                            For Each chromedriver In objVersion("downloads")(idx)
                                '---取得済みフォルダになかったら
                                If objFso.FolderExists(workPath & "\" & objVersion("version")) = False Then
                                    '---対象のplatformかチェックする ※Winsows11から64bitしかないので決め打ち
                                    If chromedriver("platform") = TARGET_PLATFORM Then
                                        url = chromedriver("url")
                                        If DownloadChromedriver(url, objVersion("version")) = False Then
                                            MsgBox "chromedriverのダウンロードに失敗しました", vbCritical
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
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chromedriver.exe の更新に失敗しました" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
    ChromedriverCheck = False
End Function
Private Function DownloadChromedriver(ByVal url As String, targetVersion As String) As Boolean
    Dim rc As Long
    Dim downloadPath As String
    Dim newDriverPath As String
    Dim objFolder As Scripting.Folder
    downloadPath = workPath & "\" & targetVersion
    ' ---chromedriverのフォルダを作成する
    Select Case SHCreateDirectoryEx(0&, downloadPath, 0&)
        Case 0:
            ' ---作成成功
        Case 183
            ' ---作成済み
        Case Else:
            ' ---作成できなかった時
            MsgBox "ChromeDriver用フォルダを作成できませんでした" & vbCrLf & Error(Err), vbCritical
            DownloadChromedriver = False
            Exit Function
    End Select
    
    '---ファイルをダウンロードする
    If URLDownloadToFile(0, url, workPath & "\" & ZIP_FILE, 0, 0) <> 0 Then
        MsgBox "ChromeDriverをダウンロードできませんでした" & vbCrLf & Error(Err), vbCritical
        DownloadChromedriver = False
        Exit Function
    End If
    Application.DisplayAlerts = False
    '---zipを既定のフォルダに向けて解凍する
    With CreateObject("Shell.Application") '---zipを既定のフォルダに向けて解凍する
        .Namespace((downloadPath)).CopyHere .Namespace((workPath & "\" & ZIP_FILE)).Items
    End With
    '--- 解凍したフォルダから再起処理してchromedriver.exeのフルパスを取得する
    newDriverPath = SearchFilesRecursively(downloadPath & "\", "chromedriver.exe")
    If newDriverPath = "" Then
        MsgBox "chromedriver.exe の更新に失敗しました"
        DownloadChromedriver = False
    End If
    '---chromedriverをバージョンフォルダ直下に移動する
    Call objFso.MoveFile(newDriverPath, downloadPath & "\")
    '---chromedriverがなくなった不要フォルダを削除する
    For Each objFolder In objFso.GetFolder(downloadPath).SubFolders
        objFolder.Delete True
    Next
    '---zipファイルを削除する
    Call objFso.DeleteFile(workPath & "\" & ZIP_FILE, True)
    Application.DisplayAlerts = True
    DownloadChromedriver = True
End Function
Function SearchFilesRecursively(ByVal folderPath As String, fileName) As String
'====================================================================================================
' folderPathを起点に再起処理でサブフォルダまで対象にしてfileNameを探してフルパスを返す
'====================================================================================================
    Dim objFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim objFile As Scripting.File
    
    ' ファイル一覧表示
    For Each objFile In objFso.GetFolder(folderPath).Files
        If objFile.Name = fileName Then
            SearchFilesRecursively = objFile.Path
            Exit Function
        End If
    Next objFile
    
    ' サブフォルダを再帰的に探索
    For Each subFolder In objFso.GetFolder(folderPath).SubFolders
        If SearchFilesRecursively(subFolder.Path, fileName) = subFolder.Path & "\" & fileName Then
            SearchFilesRecursively = subFolder.Path & "\" & fileName
            Exit Function
        End If
    Next subFolder
    SearchFilesRecursively = ""
End Function
