Imports System.Text

Namespace Parser

    ''' <summary>
    ''' CSVファイルの解析に使用する構成定義を操作するためのメソッドとプロパティを提供します。
    ''' </summary>
    Public Interface ICsvConfiguration

#Region "プロパティ"

        ''' <summary>
        ''' 最後のフィールドがカンマで終わることを許可するかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property AllowCommaTerminate() As Boolean

        ''' <summary>
        ''' ダブルクォーテーションで囲まれていないフィールドの値に、ダブルクォーテーションを含めることを許可するかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property AllowQuoteInField() As Boolean

        ''' <summary>
        ''' 異なるフィールドの数を持ったレコードを許可するかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property AllowVariableFieldCount() As Boolean

        ''' <summary>
        ''' コメント行の開始文字を取得します。
        ''' </summary>
        ReadOnly Property CommentChar() As Char

        ''' <summary>
        ''' フィールドの区切り文字を取得します。
        ''' </summary>
        ReadOnly Property DelimiterChar() As Char

        ''' <summary>
        ''' ファイルの読み込み時に使用する文字エンコーディングを取得します。
        ''' </summary>
        ReadOnly Property Encoding() As Encoding

        ''' <summary>
        ''' フィールドが引用符で囲まれているかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property HasFieldsEnclosedInQuotes() As Boolean

        ''' <summary>
        ''' 空行を無視するかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property IgnoreBlankLines() As Boolean

        ''' <summary>
        ''' 改行文字 (レコードの区切り文字) を取得します。
        ''' </summary>
        ReadOnly Property NewLineChar() As String

        ''' <summary>
        ''' フィールドを囲む際に用いられる引用符を取得します。
        ''' </summary>
        ReadOnly Property QuoteChar() As Char

        ''' <summary>
        ''' フィールド値から前後の半角スペースを削除するかどうかを示す値を取得します。
        ''' </summary>
        ReadOnly Property TrimWhiteSpace() As Boolean

#End Region

#Region "メソッド"

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Function Clone() As CsvConfiguration

        ''' <summary>
        ''' 指定された文字がコメント行の開始文字であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>コメント行の開始文字の場合は True を返し、それ以外の場合は False を返します。</returns>
        Function IsComment(c As Char) As Boolean

        ''' <summary>
        ''' 指定された文字が区切り文字であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>区切り文字の場合は True を返し、それ以外の場合は False を返します。</returns>
        Function IsDelimiter(c As Char) As Boolean

        ''' <summary>
        ''' 指定された文字が引用符であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>引用符の場合は True を返し、それ以外の場合は False を返します。</returns>
        Function IsDoubleQuote(c As Char) As Boolean

        ''' <summary>
        ''' 構成定義を検証します。
        ''' </summary>
        ''' <returns>検証が成功した場合は True。それ以外の場合は False。</returns>
        Function Validate() As Boolean

#End Region

    End Interface

End Namespace
