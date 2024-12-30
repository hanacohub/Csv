Imports System.Text

Namespace Parser

    ''' <summary>
    ''' CSVファイルの解析に使用する構成定義を表します。
    ''' </summary>
    Public Class CsvConfiguration
        Implements ICsvConfiguration
        Implements ICloneable

#Region "ICloneable Support"

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Private Function ICloneable_Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone
        End Function

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Public Function Clone() As CsvConfiguration Implements ICsvConfiguration.Clone
            Return DirectCast(Me.MemberwiseClone, CsvConfiguration)
        End Function

#End Region

#Region "規定値"

        ''' <summary>
        ''' AllowCommaTerminate プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultAllowCommaTerminate As Boolean = False

        ''' <summary>
        ''' AllowQuoteInField プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultAllowQuoteInField As Boolean = False

        ''' <summary>
        ''' AllowVariableFieldCount プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultAllowVariableFieldCount As Boolean = False

        ''' <summary>
        ''' CommentChar プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultCommentChar As Char = "#"c

        ''' <summary>
        ''' DelimiterChar プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultDelimiterChar As Char = ","c

        ''' <summary>
        ''' Encoding プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultEncoding As Encoding = Encoding.GetEncoding("Shift_JIS")

        ''' <summary>
        ''' HasFieldsEnclosedInQuote プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultHasFieldsEnclosedInQuotes As Boolean = True

        ''' <summary>
        ''' IgnoreBlankLines プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultIgnoreBlankLines As Boolean = False

        ''' <summary>
        ''' NewLineChar プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultNewLineChar As String = Environment.NewLine

        ''' <summary>
        ''' QuoteChar プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultQuoteChar As Char = """"c

        ''' <summary>
        ''' TrimWhiteSpace プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultTrimWhiteSpace As Boolean = False

#End Region

#Region "プロパティ"

        ''' <summary>
        ''' 最後のフィールドがカンマで終わることを許可するかどうかを示す値を取得または設定します。
        ''' True の場合、許可します。
        ''' False の場合、最後のフィールドがカンマで終わっていることを検出すると、Parser.Exceptions.CommaTerminateException をスローします。
        ''' </summary>
        Public Overridable Property AllowCommaTerminate() As Boolean = DefaultAllowCommaTerminate Implements ICsvConfiguration.AllowCommaTerminate

        ''' <summary>
        ''' ダブルクォーテーションで囲まれていないフィールドの値に、ダブルクォーテーションを含めることを許可するかどうかを示す値を取得または設定します。
        ''' True の場合、許可します。
        ''' False の場合、フィールドの値にダブルクォーテーションが含まれていることを検出すると、Parser.Exceptions.NotAllowedCharException をスローします。
        ''' </summary>
        Public Overridable Property AllowQuoteInField() As Boolean = DefaultAllowQuoteInField Implements ICsvConfiguration.AllowQuoteInField

        ''' <summary>
        ''' 異なるフィールドの数を持ったレコードを許可するかどうかを示す値を取得または設定します。
        ''' True の場合、許可します。
        ''' False の場合、直前に読み込んだレコードとフィールドの数が異なることを検出すると、Parser.Exceptions.FieldCountException をスローします。
        ''' </summary>
        Public Overridable Property AllowVariableFieldCount() As Boolean = DefaultAllowVariableFieldCount Implements ICsvConfiguration.AllowVariableFieldCount

        ''' <summary>
        ''' コメント行の開始文字を取得または設定します。
        ''' 指定した文字が行頭に配置配置された行は、パーサーによって無視されます。
        ''' コメント行を無視しない場合、Char.MinValue を指定します。
        ''' </summary>
        Public Overridable Property CommentChar() As Char = DefaultCommentChar Implements ICsvConfiguration.CommentChar

        ''' <summary>
        ''' フィールドの区切り文字を取得または設定します。
        ''' </summary>
        Public Overridable Property DelimiterChar() As Char = DefaultDelimiterChar Implements ICsvConfiguration.DelimiterChar

        ''' <summary>
        ''' ファイルの読み込み時に使用する文字エンコーディングを取得または設定します。
        ''' </summary>
        Public Overridable Property Encoding() As Encoding = DefaultEncoding Implements ICsvConfiguration.Encoding

        ''' <summary>
        ''' フィールドがダブルクォーテーションで囲まれているかどうかを示す値を取得または設定します。
        ''' Ture の場合、閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字を検出すると、Parser.Exceptions.MalformedException をスローします。
        ''' Ture の場合、閉じダブルクォーテーションが検出されないと、Parser.Exceptions.MissingCloseQuoteException をスローします。
        ''' False の場合、ダブルクォーテーションは常にフィールドの値として処理されます。
        ''' </summary>
        Public Overridable Property HasFieldsEnclosedInQuotes() As Boolean = DefaultHasFieldsEnclosedInQuotes Implements ICsvConfiguration.HasFieldsEnclosedInQuotes

        ''' <summary>
        ''' 空行を無視するかどうかを示す値を取得または設定します。
        ''' True の場合、空行は読み飛ばします。
        ''' False の場合、空行を検出すると1個 (AllowVariableFieldCount が False の場合) あるいは直前に読み込んだ行と同じ数 (AllowVariableFieldCount が True の場合) の空のフィールドを持った行を返します。
        ''' </summary>
        Public Overridable Property IgnoreBlankLines() As Boolean = DefaultIgnoreBlankLines Implements ICsvConfiguration.IgnoreBlankLines

        ''' <summary>
        ''' 改行文字 (レコードの区切り文字) を取得または設定します。
        ''' </summary>
        Public Overridable Property NewLineChar() As String = DefaultNewLineChar Implements ICsvConfiguration.NewLineChar

        ''' <summary>
        ''' フィールドを囲む際に用いられる引用符を取得または設定します。
        ''' </summary>
        Public Overridable Property QuoteChar() As Char = DefaultQuoteChar Implements ICsvConfiguration.QuoteChar

        ''' <summary>
        ''' フィールド値から前後の半角スペースを削除するかどうかを示す値を取得または設定します。
        ''' True の場合、レコードの先頭から、または区切り文字の後に続く半角スペース、およびフィールドの末尾から先頭に向かって続く半角スペースを削除します。ただし、ダブルクォーテーションで囲まれたフィールドの前後 (ダブルクォーテーションの内側) にある半角スペースは、常にフィールドの一部として扱い、削除しません。
        ''' False の場合、半角スペースは削除しません。
        ''' </summary>
        Public Overridable Property TrimWhiteSpace() As Boolean = DefaultTrimWhiteSpace Implements ICsvConfiguration.TrimWhiteSpace

#End Region

#Region "文字の判定"

        ''' <summary>
        ''' 指定された文字がコメント行の開始文字であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>コメント行の開始文字の場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Overridable Function IsComment(c As Char) As Boolean Implements ICsvConfiguration.IsComment
            Return (c = _CommentChar)
        End Function

        ''' <summary>
        ''' 指定された文字が改行文字(CR)であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>改行文字(CR)の場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Function IsCr(c As Char) As Boolean
            Return (c = ControlChars.Cr)
        End Function

        ''' <summary>
        ''' 指定された文字が区切り文字であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>区切り文字の場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Overridable Function IsDelimiter(c As Char) As Boolean Implements ICsvConfiguration.IsDelimiter
            Return (c = _DelimiterChar)
        End Function

        ''' <summary>
        ''' 指定された文字が引用符であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>引用符の場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Overridable Function IsDoubleQuote(c As Char) As Boolean Implements ICsvConfiguration.IsDoubleQuote
            Return (c = _QuoteChar)
        End Function

        ''' <summary>
        ''' 指定された文字が改行文字(LF)であるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>改行文字(LF)の場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Function IsLf(c As Char) As Boolean
            Return (c = ControlChars.Lf)
        End Function

        ''' <summary>
        ''' 指定された文字がスペースであるかどうかを示す System.Boolean 値を返します。
        ''' </summary>
        ''' <param name="c">テストする文字。</param>
        ''' <returns>スペースの場合は True を返し、それ以外の場合は False を返します。</returns>
        Protected Friend Function IsWhiteSpace(c As Char) As Boolean
            Return (c = " "c)
        End Function

#End Region

#Region "Validate"

        ''' <summary>
        ''' 構成定義を検証します。
        ''' </summary>
        ''' <returns>検証が成功した場合は True。それ以外の場合は False。</returns>
        Protected Friend Overridable Function Validate() As Boolean Implements ICsvConfiguration.Validate

            '改行文字
            If (_NewLineChar <> ControlChars.CrLf) AndAlso (_NewLineChar <> ControlChars.Lf) AndAlso (_NewLineChar <> ControlChars.Cr) Then
                Throw New Exceptions.ConfigurationException("指定した改行文字は無効です。")
            End If

            '区切り文字
            If _DelimiterChar = _NewLineChar Then Throw New Exceptions.ConfigurationException("指定した区切り文字が、改行文字と同じです。")

            'コメント行の開始文字
            If _CommentChar = _DelimiterChar Then Throw New Exceptions.ConfigurationException("指定したコメント行の開始文字が、区切り文字と同じです。")
            If _CommentChar = _NewLineChar Then Throw New Exceptions.ConfigurationException("指定したコメント行の開始文字が、改行文字と同じです。")
            If _CommentChar = _QuoteChar Then Throw New Exceptions.ConfigurationException("指定したコメント行の開始文字が、引用符と同じです。")

            '引用符
            If _QuoteChar = _DelimiterChar Then Throw New Exceptions.ConfigurationException("指定した引用符が、区切り文字と同じです。")
            If _QuoteChar = _NewLineChar Then Throw New Exceptions.ConfigurationException("指定した引用符が、改行文字と同じです。")

            Return True

        End Function

#End Region

    End Class

End Namespace
