Namespace Parser

    ''' <summary>
    ''' Parser.CsvRecordParseCompleted イベントにデータを提供します。
    ''' </summary>
    Public Class CsvRecordParseCompletedEventArgs
        Inherits EventArgs

#Region "コンストラクタ"

        ''' <summary>
        ''' Parser.CsvRecordParseCompletedEventArgs クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="rawFields">フィールドの生の値（区切り文字、改行文字を削除し、フィールドを囲む引用符やエスケープされたダブルクォーテーション、半角スペースを削除していない値）を格納したコレクション。</param>
        ''' <param name="rawRecord">読み込んだ行の値（編集されていない値）。</param>
        ''' <param name="fields">フィールドの値（区切り文字、改行文字、フィールドを囲む引用符、エスケープされたダブルクォーテーション、半角スペースを削除した値）を格納したコレクション。</param>
        ''' <param name="record">読み込んだ行の値。</param>
        Public Sub New(rawFields As String(), rawRecord As String, fields As String(), record As String)
            MyBase.New()

            _RawFields = rawFields
            _RawRecord = rawRecord
            _Fields = fields
            _Record = record

        End Sub

#End Region

#Region "プロパティ"

        ''' <summary>
        ''' フィールドの生の値（区切り文字、改行文字を削除し、フィールドを囲む引用符やエスケープされたダブルクォーテーション、半角スペースを削除していない値）を格納したコレクションを取得します。
        ''' </summary>
        Public ReadOnly Property Fields() As String()
            Get
                Return _Fields
            End Get
        End Property
        Private _Fields As String() = New String() {}

        ''' <summary>
        ''' フィールドの値（区切り文字、改行文字、フィールドを囲む引用符、エスケープされたダブルクォーテーション、半角スペースを削除した値）を格納したコレクションを取得します。
        ''' </summary>
        Public ReadOnly Property RawFields() As String()
            Get
                Return _RawFields
            End Get
        End Property
        Private _RawFields As String() = New String() {}

        ''' <summary>
        ''' 読み込んだ行（編集されていない値）を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property RawRecord() As String
            Get
                Return _RawRecord
            End Get
        End Property
        Private _RawRecord As String = String.Empty

        ''' <summary>
        ''' 読み込んだ行を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property Record() As String
            Get
                Return _Record
            End Get
        End Property
        Private _Record As String = String.Empty

#End Region

    End Class

End Namespace
