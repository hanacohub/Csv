Namespace Parser.Exceptions

    ''' <summary>
    ''' CSVファイルの解析でスローされる例外。
    ''' </summary>
    <Serializable>
    Public MustInherit Class ParserException
        Inherits Exception

        ''' <summary>
        ''' Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        Public Sub New()
            MyBase.New()

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージを使用して、Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">エラーを説明するメッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよびこの例外の原因となった内部例外への参照を使用して、Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">例外の原因を説明するエラー メッセージ。</param>
        ''' <param name="innerException">現在の例外の原因である例外。内部例外が指定されていない場合は null 参照 (Visual Basic では、Nothing)。</param>
        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)

        End Sub

        ''' <summary>
        ''' 現在行の解析結果への参照を使用して、Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="parseResult">現在行の解析結果。</param>
        Public Sub New(parseResult As ICsvParseResult)
            MyBase.New()

            _ParseResult = parseResult

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよび現在行の解析結果への参照を使用して、Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">エラーを説明するメッセージ。</param>
        ''' <param name="parseResult">現在行の解析結果。</param>
        Public Sub New(message As String, parseResult As ICsvParseResult)
            MyBase.New(message)

            _ParseResult = parseResult

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよび現在行の解析結果とこの例外の原因となった内部例外への参照を使用して、Parser.Exceptions.ParserException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">例外の原因を説明するエラー メッセージ。</param>
        ''' <param name="parseResult">現在行の解析結果。</param>
        ''' <param name="innerException">現在の例外の原因である例外。内部例外が指定されていない場合は null 参照 (Visual Basic では、Nothing)。</param>
        Public Sub New(message As String, parseResult As ICsvParseResult, innerException As Exception)
            MyBase.New(message, innerException)

            _ParseResult = parseResult

        End Sub

        ''' <summary>
        ''' 現在行の解析結果を取得します。
        ''' </summary>
        Public ReadOnly Property ParseResult() As ICsvParseResult

    End Class

End Namespace
