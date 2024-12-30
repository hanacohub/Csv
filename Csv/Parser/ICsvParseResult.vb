Namespace Parser

    ''' <summary>
    ''' 現在行 (論理行(読み込み済みのCSVレコード)) の解析結果を格納するためのメソッドやプロパティを定義します。
    ''' </summary>
    Public Interface ICsvParseResult

#Region "プロパティ"

        ''' <summary>
        ''' CSVファイルの解析に使用する構成定義を取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property Config() As ICsvConfiguration

        ''' <summary>
        ''' フィールドの値（区切り文字、改行文字、フィールドを囲む引用符、エスケープされたダブルクォーテーション、半角スペースを削除した値）を格納したコレクションを取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property Fields() As String()

        ''' <summary>
        ''' 解析結果を格納したフィールドのコレクションを持っているかどうかを示す値を取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property HasFields() As Boolean

        ''' <summary>
        ''' 最後に解析した文字の位置を示す 0 から始まるインデックス番号を取得します。
        ''' </summary>
        ReadOnly Property LastCharIndex() As Integer

        ''' <summary>
        ''' 解析した文字が含まれる行の番号を取得します。
        ''' これは実際のファイル行番号です。
        ''' </summary>
        ReadOnly Property LineNumber() As Integer

        ''' <summary>
        ''' 現在の文字を解析した結果、決定した新しい状態を取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property NewState As CsvParserState

        ''' <summary>
        ''' 現在の文字を解析する時点の状態を取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property OldState As CsvParserState

        ''' <summary>
        ''' フィールドの生の値（区切り文字、改行文字を削除し、フィールドを囲む引用符やエスケープされたダブルクォーテーション、半角スペースを削除していない値）を格納したコレクションを取得します。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property RawFields() As String()

        ''' <summary>
        ''' 読み込んだ行（編集されていない値）を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property RawRecord() As String

        ''' <summary>
        ''' 読み込んだ行を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property Record() As String

        ''' <summary>
        ''' 現在のレコード番号を取得します。
        ''' これは論理行(読み込み済みのCSVレコード)のレコード番号です。
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property RecordNumber() As Integer

#End Region

#Region "メソッド"

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Function Clone() As CsvParseResult

#End Region

    End Interface

End Namespace
