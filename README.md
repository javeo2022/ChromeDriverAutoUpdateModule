# ChromeDriverAutoUpdateModule
SeleniumBasic用のVBAでChromeDriverを自動更新するためのプログラムです

# プログラム以外に必要なもの
- VBA-JSONを利用していますのでセットで使ってください

※VBA-JSONはこちら https://github.com/VBA-tools/VBA-JSON

- VBA-JSONとSeleniumBasicをセットで使うとエラーが発生するのでVBA-JSONのDictionaryをScripting.Dictionaryへ変更してください

※詳しくは https://javeo.jp/vba-json-error/ で紹介しています

# 必要な参照設定
参照設定は下記を有効にしてください
- Microsoft Scripting Runtime
- Microsoft XML, v6.0

# 使い方
Boolean型の返値にしているChromeDriverAutoUpdateModule.ChromeDriverAutoUpdateを実行するだけです
```VB
Option Explicit
Sub main()
'---Chromedriverの自動更新
If ChromeDriverAutoUpdateModule.ChromeDriverAutoUpdate = False Then
    Exit Sub
End If
End Sub
```
