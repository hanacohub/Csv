Imports System.Text

Namespace Parser

    ''' <summary>
    ''' 現在行 (論理行(読み込み済みのCSVレコード)) の解析結果を格納するためのクラスです。
    ''' </summary>
    Public Class CsvParseResult
        Implements ICloneable
        Implements ICsvParseResult

#Region "ICloneable Support"

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Private Function ICloneable_Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone
        End Function

#End Region

#Region "ICsvParseResult Support"

        ''' <summary>
        ''' 現在のインスタンスのコピーである新しいオブジェクトを作成します。
        ''' </summary>
        ''' <returns>このインスタンスのコピーである新しいオブジェクト。</returns>
        Public Function Clone() As CsvParseResult Implements ICsvParseResult.Clone

            Dim clonedParseResult As CsvParseResult = DirectCast(Me.MemberwiseClone, CsvParseResult)
            clonedParseResult._Fields = New List(Of String)(_Fields)
            clonedParseResult._RawFields = New List(Of String)(_RawFields)
            clonedParseResult._Config = _Config.Clone

            Return clonedParseResult

        End Function

#End Region

#Region "イベント"

        ''' <summary>
        ''' レコードの解析が完了した後（レコードの終端を検出した後）に発生します。
        ''' </summary>
        Public Event CsvRecordParseCompleted As CsvRecordParseCompletedEventHandler

#End Region

#Region "イベント発生メソッド"

        ''' <summary>
        ''' Parser.CsvRecordParseCompleted イベントを発生させます。
        ''' </summary>
        ''' <param name="e">イベントに関連する情報。</param>
        Protected Overridable Sub OnCsvRecordParseCompleted(e As CsvRecordParseCompletedEventArgs)
            RaiseEvent CsvRecordParseCompleted(Me, e)
        End Sub

#End Region

#Region "規定値"

        ''' <summary>
        ''' LastCharIndex プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultLastCharIndex As Integer = -1

        ''' <summary>
        ''' LineNumber プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultLineNumber As Integer = -1

        ''' <summary>
        ''' NewState プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultNewState As CsvParserState = CsvParserState.Start

        ''' <summary>
        ''' OldState プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultOldState As CsvParserState = CsvParserState.Start

        ''' <summary>
        ''' Read プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultRead As Boolean = False

        ''' <summary>
        ''' RecordNumber プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultRecordNumber As Integer = -1

#End Region

#Region "フィールド"

        ''' <summary>
        ''' 解析した文字を格納するためのコレクション。
        ''' </summary>
        Private _FieldChars As StringBuilder = Nothing

        ''' <summary>
        ''' 解析した文字を格納するためのコレクション。
        ''' </summary>
        Private _RawFieldChars As StringBuilder = Nothing

        ''' <summary>
        ''' 解析した文字を格納するためのコレクション。
        ''' </summary>
        Private _WhiteSpaces As StringBuilder = Nothing

#End Region

#Region "プロパティ"

        ''' <summary>
        ''' CSVファイルの解析に使用する構成定義を取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property Config() As ICsvConfiguration = Nothing Implements ICsvParseResult.Config

        ''' <summary>
        ''' 解析中にスローされた例外を取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property Exceptions() As List(Of Exception)
            Get
                Return _Exceptions
            End Get
        End Property
        Private _Exceptions As List(Of Exception) = Nothing

        ''' <summary>
        ''' フィールドの値（区切り文字、改行文字、フィールドを囲む引用符、エスケープされたダブルクォーテーション、半角スペースを削除した値）を格納したコレクションを取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property Fields() As String() Implements ICsvParseResult.Fields
            Get
                Return _Fields.ToArray
            End Get
        End Property
        Private _Fields As List(Of String) = Nothing

        ''' <summary>
        ''' 解析結果を格納したフィールドのコレクションを持っているかどうかを示す値を取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property HasFields() As Boolean Implements ICsvParseResult.HasFields
            Get
                Return (_Fields.Count > 0)
            End Get
        End Property

        ''' <summary>
        ''' 最後に解析した文字の位置を示す 0 から始まるインデックス番号を取得します。
        ''' </summary>
        Public ReadOnly Property LastCharIndex() As Integer = DefaultLastCharIndex Implements ICsvParseResult.LastCharIndex

        ''' <summary>
        ''' 解析した文字が含まれる行の番号を取得します。
        ''' これは実際のファイル行番号です。
        ''' </summary>
        Public Overridable ReadOnly Property LineNumber() As Integer = DefaultLineNumber Implements ICsvParseResult.LineNumber

        ''' <summary>
        ''' 現在の文字を解析した結果、決定した新しい状態を取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property NewState As CsvParserState = DefaultNewState Implements ICsvParseResult.NewState

        ''' <summary>
        ''' 現在の文字を解析する時点の状態を取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property OldState As CsvParserState = DefaultOldState Implements ICsvParseResult.OldState

        ''' <summary>
        ''' 現在の行が CsvParser.Read メソッドで読み込まれたかどうかを示す値を取得または設定します。
        ''' </summary>
        ''' <returns></returns>
        Friend Property Read As Boolean = DefaultRead

        ''' <summary>
        ''' フィールドの生の値（区切り文字、改行文字を削除し、フィールドを囲む引用符やエスケープされたダブルクォーテーション、半角スペースを削除していない値）を格納したコレクションを取得します。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property RawFields() As String() Implements ICsvParseResult.RawFields
            Get
                Return _RawFields.ToArray
            End Get
        End Property
        Private _RawFields As List(Of String) = Nothing

        ''' <summary>
        ''' 読み込んだ行（編集されていない値）を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property RawRecord() As String Implements ICsvParseResult.RawRecord
            Get
                Return String.Join(_Config.DelimiterChar, Me.RawFields)
            End Get
        End Property

        ''' <summary>
        ''' 読み込んだ行を文字列として取得します。
        ''' これは論理行(読み込み済みのCSVレコード)です。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property Record() As String Implements ICsvParseResult.Record
            Get
                Return String.Join(_Config.DelimiterChar, Me.Fields)
            End Get
        End Property

        ''' <summary>
        ''' 現在のレコード番号を取得します。
        ''' これは論理行(読み込み済みのCSVレコード)のレコード番号です。
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Property RecordNumber() As Integer = DefaultRecordNumber Implements ICsvParseResult.RecordNumber

#End Region

#Region "メソッド"

        ''' <summary>
        ''' 解析した文字を追加します。
        ''' </summary>
        ''' <param name="c">解析した文字。</param>
        ''' <param name="lineNumber">解析した文字が含まれる行の番号。</param>
        ''' <param name="lastCharIndex">最後に解析した文字の位置を示す 0 から始まるインデックス番号。</param>
        ''' <param name="fieldState">文字を解析した結果、決定したフィールドの値に関する新しい状態。</param>
        ''' <param name="oldState">現在の文字を解析する時点の状態。</param>
        ''' <param name="newState">現在の文字を解析した結果、決定した新しい状態。</param>
        Public Sub Add(c As Char, lineNumber As Integer, lastCharIndex As Integer, fieldState As CsvParserState, oldState As CsvParserState, newState As CsvParserState)

            _LineNumber = lineNumber
            _LastCharIndex = lastCharIndex
            _OldState = oldState
            _NewState = newState

            Try
                'コメント行の場合、以降の処理は行いません。
                If fieldState = CsvParserState.Comment Then Return

                '区切り文字、改行文字(およびファイルの終端)、エスケープされたダブルクォーテーション、コメント行以外の場合、RawField の値として追加します。
                If (newState <> CsvParserState.Start) AndAlso (newState <> CsvParserState.CarriageReturn) AndAlso (newState <> CsvParserState.EndOfRecord) AndAlso (newState <> CsvParserState.EscapedQuote) Then
                    _RawFieldChars.Append(c)
                End If

                'レコードの先頭から、または区切り文字の後に続く半角スペース、およびフィールドの末尾から先頭に向かって続く半角スペースを削除します。
                If newState = CsvParserState.WhiteSpace Then
                    If (fieldState = CsvParserState.Start) OrElse (fieldState = CsvParserState.CloseQuote) Then
                        'フィールドの先頭から続く半角スペースおよび閉じダブルクォーテーションの後に続く半角スペースの場合、以降の処理は行いません。
                        Return
                    Else
                        _WhiteSpaces.Append(c)
                    End If
                End If

                'ダブルクォーテーションで囲まれていないフィールド、ダブルクォーテーションで囲まれたフィールド、エスケープされたダブルクォーテーションの場合、Field の値として追加します。
                'また、ダブルクォーテーションで囲まれたフィールドで CsvParserState.CarriageReturn を検出した場合も、Field の値として追加します。
                If (newState = CsvParserState.Field) OrElse
                    (newState = CsvParserState.Quoted) OrElse
                    (newState = CsvParserState.EscapedQuote) OrElse
                    ((newState = CsvParserState.CarriageReturn) AndAlso (fieldState = CsvParserState.Quoted)) OrElse
                    ((newState = CsvParserState.CarriageReturn) AndAlso (fieldState = CsvParserState.EscapedQuote)) Then
                    'ダブルクォーテーションで囲まれたフィールドの前後(ダブルクォーテーションの内側)にある半角スペースは削除しません。
                    For i = 0 To _WhiteSpaces.Length - 1
                        _FieldChars.Append(_WhiteSpaces(i))
                    Next
                    _FieldChars.Append(c)
                    _WhiteSpaces.Clear()
                End If

                '区切り文字、改行文字(またはファイルの終端)の場合、または最後のフィールドがカンマで終わっていた場合、フィールドの値を確定します。
                If (newState = CsvParserState.Start) OrElse (newState = CsvParserState.EndOfRecord) Then
                    If (newState = CsvParserState.Start) OrElse
                        ((newState = CsvParserState.EndOfRecord) AndAlso (_RawFieldChars.Length > 0)) OrElse
                        ((newState = CsvParserState.EndOfRecord) AndAlso (_RawFieldChars.Length = 0) AndAlso (_RawFields.Count > 0)) Then
                        _RawFields.Add(_RawFieldChars.ToString)
                    End If

                    If (newState = CsvParserState.Start) OrElse
                        ((newState = CsvParserState.EndOfRecord) AndAlso (_FieldChars.Length > 0)) OrElse
                        ((newState = CsvParserState.EndOfRecord) AndAlso (_FieldChars.Length = 0) AndAlso (_Fields.Count > 0)) Then
                        _Fields.Add(_FieldChars.ToString)
                    End If

                    _RawFieldChars.Clear()
                    _FieldChars.Clear()
                End If

                '改行文字(またはファイルの終端)の場合、レコードの解析が完了したことを通知します。
                If newState = CsvParserState.EndOfRecord Then
                    Me.OnCsvRecordParseCompleted(New CsvRecordParseCompletedEventArgs(Me.RawFields, Me.RawRecord, Me.Fields, Me.Record))
                End If

            Finally
                'レコードの先頭から、または区切り文字の後に続く半角スペース、およびフィールドの末尾から先頭に向かって続く半角スペースを削除します。
                If newState <> CsvParserState.WhiteSpace Then
                    _WhiteSpaces.Clear()
                End If
            End Try

        End Sub

        ''' <summary>
        ''' 値が空のフィールドを追加します。
        ''' </summary>
        ''' <param name="fieldCount">作成するフィールドの数。</param>
        Friend Sub AddEmptyField(fieldCount As Integer)

            If fieldCount < 1 Then
                fieldCount = 1
            End If

            _RawFields.Clear()
            _Fields.Clear()

            For i = 1 To fieldCount
                _RawFields.Add(String.Empty)
                _Fields.Add(String.Empty)
            Next

        End Sub

        ''' <summary>
        ''' 解析結果をクリアします。
        ''' </summary>
        Protected Overridable Sub Clear()

            _LastCharIndex = DefaultLastCharIndex
            _Exceptions.Clear()
            _FieldChars.Clear()
            _Fields.Clear()
            _LineNumber = DefaultLineNumber
            _RawFieldChars.Clear()
            _RawFields.Clear()
            _Read = DefaultRead
            _RecordNumber = DefaultRecordNumber
            _WhiteSpaces.Clear()

        End Sub

        ''' <summary>
        ''' 解析結果を初期化します。
        ''' </summary>
        ''' <param name="recordNumber">現在のレコード番号。これは論理行(読み込み済みのCSVレコード)のレコード番号です。</param>
        Protected Friend Overridable Sub Initialize(recordNumber As Integer)

            Me.Clear()
            _RecordNumber = recordNumber

        End Sub

#End Region

#Region "コンストラクタ"

        ''' <summary>
        ''' Parser.CsvParseResult クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="config">CSVファイルの解析に使用する構成定義。</param>
        Public Sub New(config As CsvConfiguration)

            _Config = config
            _Exceptions = New List(Of Exception)
            _FieldChars = New StringBuilder
            _Fields = New List(Of String)
            _RawFieldChars = New StringBuilder
            _RawFields = New List(Of String)
            _WhiteSpaces = New StringBuilder

        End Sub

#End Region

    End Class

End Namespace
