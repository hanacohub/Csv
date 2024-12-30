Namespace Parser

    ''' <summary>
    ''' 文字を解析した後の状態について記述します。
    ''' </summary>
    Public Enum CsvParserState

        ''' <summary>
        ''' 初期状態または区切り文字を検出した状態。
        ''' </summary>
        Start

        ''' <summary>
        ''' ダブルクォーテーションで囲まれていないフィールドを検出した状態。
        ''' </summary>
        Field

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、開きダブルクォーテーションを検出した状態。
        ''' </summary>
        OpenQuote

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドを検出した状態。
        ''' </summary>
        Quoted

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、エスケープされたダブルクォーテーションを検出した状態。
        ''' </summary>
        EscapedQuote

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、閉じダブルクォーテーションを検出した状態。
        ''' </summary>
        CloseQuote

        ''' <summary>
        ''' 改行文字(CR)を検出した状態。
        ''' </summary>
        ''' <remarks>構成定義の NewLineChar に CrLf が指定されている場合に使用します。</remarks>
        CarriageReturn

        ''' <summary>
        ''' コメント行の開始文字を検出した状態
        ''' </summary>
        Comment

        ''' <summary>
        ''' 半角スペースを検出した状態。
        ''' </summary>
        ''' <remarks>構成定義の TrimWhiteSpace が True が指定されている場合に使用します。</remarks>
        WhiteSpace

        ''' <summary>
        ''' レコードの終端を検出した状態。
        ''' </summary>
        EndOfRecord

        ''' <summary>
        '''エラーを検出した状態。
        ''' </summary>
        [Error]

    End Enum

End Namespace
