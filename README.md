# csv
CSVファイルを読み込むためのライブラリです。
* 読み込んだレコードのフィールドを String 型の配列として返します。
* ファイルの1行目をヘッダーとして扱う機能はありません。ヘッダーも通常の明細と同じように処理します。
* ダブルクォーテーションで囲まれたフィールドの前後 (ダブルクォーテーションの内側) にある半角スペースは、常にフィールドの一部として扱い、削除しません。
* 空行は1個または直前に読み込んだレコードのフィールドと同じ数の空文字の配列として返します。
  <br>※ 設定により変更可
* コメント行は読み飛ばします。
  <br>※ 設定により変更可
## 開発環境
* Windows 10 Pro 64bit
* Visual Studio Community 2022 (64 ビット)
* .NET Framework 4.8
* Visual Basic .NET
## 使い方
1. プロジェクトの参照に「Hanaco.Csv.dll」を追加します。
2. CsvConfiguration のインスタンスを生成し、設定を変更します (必要な場合のみ)。
3. 読み込むファイル名と、2 で生成した CsvConfiguration のインスタンスを指定して CsvParser のインスタンスを生成します。
4. Read メソッドで1行ずつレコードを読み込みます。
```vb
Dim config As New CsvConfiguration

config.AllowCommaTerminate = True

Using parser = New CsvParser(TextBoxCsv.Text, config)
    Do Until parser.EndOfData
        Dim rec = parser.Read()
    Loop
End Using
```
## 設定
| プロパティ                 | 説明          | デフォルト     | RFC準拠       |
| :--- | :--- | :---: | :---: |
| AllowCommaTerminate       | 最後のフィールドがカンマで終わることを許可するかどうかを制御します。<br>● True の場合、許可します。<br>● False の場合、最後のフィールドがカンマで終わっていることを検出すると、CommaTerminateException をスローします。<br>デフォルトは False です。 | False | False |
| AllowQuoteInField         | ダブルクォーテーションで囲まれていないフィールドの値に、ダブルクォーテーションを含めることを許可するかどうかを制御します。<br>● True の場合、許可します。<br>● False の場合、フィールドの値にダブルクォーテーションが含まれていることを検出すると、NotAllowedCharException をスローします。<br>デフォルトは False です。 | False | False |
| AllowVariableFieldCount   | 異なるフィールドの数を持ったレコードを許可するかどうかを制御します。<br>● True の場合、許可します。<br>● False の場合、直前に読み込んだレコードとフィールドの数が異なることを検出すると、FieldCountException をスローします。<br>デフォルトは False です。 | False | False |
| CommentChar               | コメント行の開始文字です。<br>指定した文字が行頭に配置配置された行は、パーサーによって無視されます。<br>コメント行を無視しない場合、Char.MinValue を指定します。<br>デフォルトは # です。 | # | － |
| DelimiterChar             | フィールドの区切り文字です。<br>デフォルトは , です。 | , | , |
| Encoding                  | ファイルの読み込み時に使用する文字エンコーディングです。<br>デフォルトは Shift_JIS です。 |Shift_JIS      | － |
| HasFieldsEnclosedInQuotes | フィールドがダブルクォーテーションで囲まれているかどうかを制御します。<br>● Ture の場合、閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字を検出すると、MalformedException をスローします。<br>● Ture の場合、閉じダブルクォーテーションが検出されないと、MissingCloseQuoteException をスローします。<br>● False の場合、ダブルクォーテーションは常にフィールドの値として処理されます。<br>デフォルトは True です。 | True  | True |
| IgnoreBlankLines          | 空行を無視するかどうかを制御します。<br>● True の場合、空行は読み飛ばします。<br>● False の場合、空行を検出すると1個 (AllowVariableFieldCount が True の場合、または1行目が空行だった場合) あるいは直前に読み込んだ行と同じ数 (AllowVariableFieldCount が False の場合) の空のフィールドを持った行を返します。<br>デフォルトは False です。 | False | － |
| NewLineChar               | 改行文字 (レコードの区切り文字) です。<br>デフォルトは CRLF です。 | CRLF | CRLF |
| QuoteChar                 | フィールドを囲む際に用いられる引用符です。<br>デフォルトは " です。 | " | " |
| TrimWhiteSpace            | フィールド値から前後の半角スペースを削除するかどうかを制御します。<br>● True の場合、レコードの先頭から、または区切り文字の後に続く半角スペース、およびフィールドの末尾から先頭に向かって続く半角スペースを削除します。ただし、ダブルクォーテーションで囲まれたフィールドの前後 (ダブルクォーテーションの内側) にある半角スペースは、常にフィールドの一部として扱い、削除しません。<br>● False の場合、半角スペースは削除しません。<br>デフォルトは False です。 | False | False |
## 例外
以下の場合、例外をスローします。
* 構成定義 (CsvConfiguration) の検証に失敗した場合、ConfigurationException の例外をスローします。
* AllowCommaTerminate プロパティが False の場合、最後のフィールドがカンマで終わっていることを検出すると CommaTerminateException をスローします。
* AllowQuoteInField プロパティが False の場合、フィールドの値にダブルクォーテーションが含まれていることを検出すると NotAllowedCharException をスローします。
* AllowVariableFieldCount プロパティが False の場合、直前に読み込んだレコードとフィールドの数が異なることを検出すると FieldCountException をスローします。
* HasFieldsEnclosedInQuotes プロパティが Ture の場合、閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字を検出すると MalformedException をスローします。
* HasFieldsEnclosedInQuotes プロパティが Ture の場合、閉じダブルクォーテーションが検出できなかったとき MissingCloseQuoteException をスローします。
## ライセンス
[MIT ライセンス](./LICENSE)です。ご自由にお使いください。
ただし、本リポジトリを使用して発生したいかなる問題についても、当方が責任を負うことはありませんので、ご使用は自己責任でお願いします。

© 2024 Hanaco
