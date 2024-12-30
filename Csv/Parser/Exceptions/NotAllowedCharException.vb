Namespace Parser.Exceptions

    ''' <summary>
    ''' AllowQuoteInField プロパティが False の場合、フィールドの値にダブルクォーテーションが含まれていることを検出したときにスローされる例外。
    ''' </summary>
    <Serializable>
    Public Class NotAllowedCharException
        Inherits ParserException

        ''' <summary>
        ''' Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        Public Sub New()
            MyBase.New(My.Resources.Arg_NotAllowedCharException)

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージを使用して、Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">エラーを説明するメッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよびこの例外の原因となった内部例外への参照を使用して、Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">例外の原因を説明するエラー メッセージ。</param>
        ''' <param name="innerException">現在の例外の原因である例外。内部例外が指定されていない場合は null 参照 (Visual Basic では、Nothing)。</param>
        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)

        End Sub

        ''' <summary>
        ''' 現在行の解析結果への参照を使用して、Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="parseResult">現在行の解析結果。</param>
        Public Sub New(parseResult As ICsvParseResult)
            MyBase.New(My.Resources.Arg_NotAllowedCharException, parseResult)

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよび現在行の解析結果への参照を使用して、Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">エラーを説明するメッセージ。</param>
        ''' <param name="parseResult">現在行の解析結果。</param>
        Public Sub New(message As String, parseResult As ICsvParseResult)
            MyBase.New(message, parseResult)

        End Sub

        ''' <summary>
        ''' 指定したエラー メッセージおよび現在行の解析結果とこの例外の原因となった内部例外への参照を使用して、Parser.Exceptions.NotAllowedCharException クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">例外の原因を説明するエラー メッセージ。</param>
        ''' <param name="parseResult">現在行の解析結果。</param>
        ''' <param name="innerException">現在の例外の原因である例外。内部例外が指定されていない場合は null 参照 (Visual Basic では、Nothing)。</param>
        Public Sub New(message As String, parseResult As ICsvParseResult, innerException As Exception)
            MyBase.New(message, parseResult, innerException)

        End Sub

    End Class

End Namespace
