Imports System.IO
Imports Hanaco.Csv.Parser.Exceptions

Namespace Parser

    ''' <summary>
    ''' CSVファイルの解析に使用するメソッドとプロパティを提供します。
    ''' </summary>
    Public Class CsvParser
        Implements IDisposable

#Region "IDisposable Support"
        Private disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: マネージ状態を破棄します (マネージ オブジェクト)。
                End If

                ' TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
                ' TODO: 大きなフィールドを null に設定します。

                'Reader を解放します。
                Me.Release()
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: 上の Dispose(ByVal disposing As Boolean) にアンマネージ リソースを解放するコードがある場合にのみ、Finalize() をオーバーライドします。
        'Protected Overrides Sub Finalize()
        '    ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

#Region "規定値"

        ''' <summary>
        ''' _AheadResultIndex の既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultAheadResultIndex As Integer = -1

        ''' <summary>
        ''' _BufferRemaining の既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultBufferRemaining As Integer = 0

        ''' <summary>
        ''' バッファーに読み取り可能な最大文字数。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultBufferSize As Integer = 4096

        ''' <summary>
        ''' BufferStartPos プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultBufferStartPos As Integer = 0

        ''' <summary>
        ''' CharIndex プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultCharIndex As Integer = -1

        ''' <summary>
        ''' _CurrentResultIndex の既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultCurrentResultIndex As Integer = -1

        ''' <summary>
        ''' _DetectedNewLine の既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultDetectedNewLine As Boolean = False

        ''' <summary>
        ''' LineNumber プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultLineNumber As Integer = 0

        ''' <summary>
        ''' RecordNumber プロパティの既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultRecordNumber As Integer = 0

        ''' <summary>
        ''' _SaveFieldCount 既定値。
        ''' </summary>
        ''' <returns></returns>
        Private Shared ReadOnly Property DefaultSaveFieldCount As Integer = -1

#End Region

#Region "フィールド"

        ''' <summary>
        ''' 現在行の1行先の解析結果を返す 0 から始まるインデックス番号。
        ''' </summary>
        Private _AheadResultIndex As Integer = DefaultAheadResultIndex

        ''' <summary>
        ''' 読み込んだ文字を格納するバッファー。
        ''' </summary>
        Private _BufferChars As Char() = {}

        ''' <summary>
        ''' バッファーに格納された文字のうち、まだ読み取られていない文字の数。
        ''' </summary>
        Private _BufferRemaining As Integer = DefaultBufferRemaining

        ''' <summary>
        ''' 解析した文字の位置。
        ''' </summary>
        Private _CharIndex As Integer = DefaultCharIndex

        ''' <summary>
        ''' 現在行の解析結果を返す 0 から始まるインデックス番号。
        ''' </summary>
        Private _CurrentResultIndex As Integer = DefaultCurrentResultIndex

        ''' <summary>
        ''' 改行文字を検出したかどうかを示す値。
        ''' </summary>
        Private _DetectedNewLine As Boolean = DefaultDetectedNewLine

        ''' <summary>
        ''' リーダー。
        ''' </summary>
        Private _Reader As TextReader = Nothing

        ''' <summary>
        ''' 読み込んだレコードが持っているフィールドの数。
        ''' </summary>
        Private _SaveFieldCount As Integer = DefaultSaveFieldCount

#End Region

#Region "プロパティ"

        ''' <summary>
        ''' 現在行の1行先の解析結果を取得します。
        ''' </summary>
        Private ReadOnly Property AheadResult() As CsvParseResult
            Get
                Return _ParseResults(_AheadResultIndex)
            End Get
        End Property

        ''' <summary>
        ''' バッファーの読み取り開始位置を取得します。
        ''' </summary>
        ''' <returns></returns>
        Private ReadOnly Property BufferStartPos() As Integer
            Get
                If _BufferStartPos >= DefaultBufferSize Then
                    _BufferStartPos = DefaultBufferStartPos
                End If
                Return _BufferStartPos
            End Get
        End Property
        Private _BufferStartPos As Integer = DefaultBufferStartPos

        ''' <summary>
        ''' CSVファイルの解析に使用する構成定義を取得します。
        ''' </summary>
        Public ReadOnly Property Config() As ICsvConfiguration
            Get
                Return _Config
            End Get
        End Property
        Private _Config As CsvConfiguration = Nothing

        ''' <summary>
        ''' 現在行の解析結果を取得します。
        ''' </summary>
        Private ReadOnly Property CurrentResult() As CsvParseResult
            Get
                Return _ParseResults(_CurrentResultIndex)
            End Get
        End Property

        ''' <summary>
        ''' 現在のカーソル位置以降にデータが存在しない場合、True を返します。
        ''' </summary>
        Public ReadOnly Property EndOfData() As Boolean
            Get
                Try
                    Return Me.EndOfStream AndAlso (Not (_ParseResults.Any(Function(e) e.HasFields AndAlso (Not e.Read))))
                Catch ex As System.ObjectDisposedException
                    Return True
                End Try
            End Get
        End Property

        ''' <summary>
        ''' 現在のカーソル位置以降にデータが存在しない場合、True を返します。
        ''' </summary>
        Private ReadOnly Property EndOfStream() As Boolean
            Get
                Try
                    Return (_Reader.Peek < 0) AndAlso (_BufferRemaining <= 0)
                Catch ex As System.ObjectDisposedException
                    Return True
                End Try
            End Get
        End Property

        ''' <summary>
        ''' 現在の行番号を返します。ストリームから取り出す文字がなくなった場合は -1 を返します。
        ''' これは実際のファイル行番号です。
        ''' </summary>
        Public ReadOnly Property LineNumber() As Integer
            Get
                Return If(Me.EndOfData, -1, _LineNumber)
            End Get
        End Property
        Private _LineNumber As Integer = DefaultLineNumber

        ''' <summary>
        ''' 現在行の解析結果を取得します。
        ''' </summary>
        Public ReadOnly Property ParseResult() As ICsvParseResult
            Get
                If (_ParseResults.GetLowerBound(0) <= _CurrentResultIndex) AndAlso (_CurrentResultIndex <= _ParseResults.GetUpperBound(0)) Then
                    Return _ParseResults(_CurrentResultIndex).Clone
                End If

                Return Nothing
            End Get
        End Property
        Private _ParseResults As CsvParseResult() = Nothing

        ''' <summary>
        ''' 現在のレコード番号を取得します。
        ''' これは論理行(読み込み済みのCSVレコード)のレコード番号です。
        ''' </summary>
        Public ReadOnly Property RecordNumber() As Integer = DefaultRecordNumber

#End Region

#Region "イベント ハンドラー"

        ''' <summary>
        ''' CsvRecordParseCompleted イベントの通知を受け取ります。
        ''' </summary>
        ''' <param name="sender">イベントを発生させたオブジェクト。</param>
        ''' <param name="e">イベントに関連する情報。</param>
        Protected Overridable Sub ParseResult_CsvRecordParseCompleted(sender As Object, e As CsvRecordParseCompletedEventArgs)

            '最後のフィールドがカンマで終わっていた場合、例外をスローします。
            If (Not _Config.AllowCommaTerminate) AndAlso (Me.AheadResult.RawRecord.Trim.Length > 0) Then
                If _Config.IsDelimiter(CType(Strings.Right(e.Record, 1), Char)) Then
                    Me.AheadResult.Exceptions.Add(New Exceptions.CommaTerminateException(Me.AheadResult.Clone))
                    Return
                End If
            End If

            '直前に読み込んだレコードとフィールドの数が異なる場合、例外をスローします。
            If Not _Config.AllowVariableFieldCount Then
                If _SaveFieldCount = DefaultSaveFieldCount Then
                    _SaveFieldCount = Me.AheadResult.Fields.Count
                ElseIf _SaveFieldCount <> Me.AheadResult.Fields.Count Then
                    Me.AheadResult.Exceptions.Add(New Exceptions.FieldCountException(Me.AheadResult.Clone))
                    Return
                End If
            End If

        End Sub

#End Region

#Region "NewLineDetected"

        ''' <summary>
        ''' 改行文字を検出したときの処理。
        ''' </summary>
        Protected Overridable Sub NewLineDetected()
            _LineNumber += 1
            _CharIndex = DefaultCharIndex
            _DetectedNewLine = False
        End Sub

#End Region

#Region "Read"

        ''' <summary>
        ''' 初期読み込み。
        ''' </summary>
        Protected Overridable Sub InitialRead()

            For i = 0 To _ParseResults.Count - 1
                If Me.EndOfStream Then Return
                Me.Readahead(i)
            Next

        End Sub

        ''' <summary>
        ''' 現在行の1行先のレコードを読み込みます。
        ''' </summary>
        ''' <param name="aheadResultIndex">現在行の1行先の解析結果を格納する 0 から始まるインデックス番号。</param>
        Protected Overridable Sub Readahead(aheadResultIndex As Integer)

            If Me.EndOfStream Then Return

            _AheadResultIndex = aheadResultIndex
            _RecordNumber += 1
            Me.AheadResult.Initialize(_RecordNumber)
            Me.ParseLine()

        End Sub

        ''' <summary>
        ''' 現在行のすべてのフィールドを読み込んで文字列の配列として返し、次のデータが格納されている行にカーソルを進めます。
        ''' </summary>
        ''' <returns>現在の行のフィールド値を格納する文字列の配列。</returns>
        Public Overridable Function Read() As String()

            If _CurrentResultIndex < 0 Then
                _CurrentResultIndex = 0
                _AheadResultIndex = 1
            Else
                _CurrentResultIndex = 1 - _CurrentResultIndex
                _AheadResultIndex = 1 - _AheadResultIndex
            End If

            '読み込み時に例外が発生していた場合、その例外をスローします。
            If Me.CurrentResult.Exceptions.Count > 0 Then
                Throw Me.CurrentResult.Exceptions.First
            End If

            '読み込み済みのフィールドを取得します。
            Dim fiels = Me.CurrentResult.Fields
            Me.CurrentResult.Read = True

            If Me.CurrentResult.RecordNumber > 1 Then '1行目を読み込んだ時点で、1行目の ParseResult を上書きしないため、すぐには次の行を読み込みません。
                Me.Readahead(_AheadResultIndex)
            End If

            If fiels.Count > 0 Then
                Return fiels
            Else
                Return Nothing
            End If

        End Function

#End Region

#Region "ReadChar"

        ''' <summary>
        ''' 次の文字を読み取り、１文字分だけ文字位置を進めます。
        ''' </summary>
        ''' <returns>次の文字。それ以上読み取り可能な文字がない場合は Char.MinValue。 </returns>
        Protected Overridable Function ReadChar() As Char

            If (_BufferRemaining = 0) AndAlso (Not (_Reader.Peek < 0)) Then
                _BufferRemaining += _Reader.Read(_BufferChars, _BufferRemaining - Me.BufferStartPos, _BufferChars.Length - _BufferRemaining)
            End If

            If _BufferRemaining > 0 Then
                Dim c = _BufferChars(Me.BufferStartPos)
                _CharIndex += 1
                _BufferChars(Me.BufferStartPos) = Char.MinValue
                _BufferStartPos += 1
                _BufferRemaining -= 1
                Return c
            Else
                _CharIndex = DefaultCharIndex
                Return Char.MinValue
            End If

        End Function

#End Region

#Region "Parse"

        ''' <summary>
        ''' 初期状態または区切り文字を検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInStart(c As Char) As CsvParserState

            If _Config.IsDoubleQuote(c) Then
                If _Config.HasFieldsEnclosedInQuotes Then
                    Return CsvParserState.OpenQuote
                ElseIf _Config.AllowQuoteInField Then
                    Return CsvParserState.Field
                Else
                    'フィールドの値にダブルクォーテーションが含まれている場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.NotAllowedCharException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                End If

            ElseIf _Config.IsDelimiter(c) Then
                Return CsvParserState.Start

            ElseIf _Config.IsCr(c) Then
                If _Config.NewLineChar = ControlChars.CrLf Then
                    Return CsvParserState.CarriageReturn
                ElseIf _Config.NewLineChar = ControlChars.Lf Then
                    Return CsvParserState.Field
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsLf(c) Then
                If (_Config.NewLineChar = ControlChars.CrLf) OrElse (_Config.NewLineChar = ControlChars.Cr) Then
                    Return CsvParserState.Field
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsWhiteSpace(c) Then
                If _Config.TrimWhiteSpace Then
                    Return CsvParserState.WhiteSpace
                Else
                    Return CsvParserState.Field
                End If

            ElseIf _Config.IsComment(c) Then
                Return CsvParserState.Comment

            Else
                Return CsvParserState.Field

            End If

        End Function

        ''' <summary>
        ''' ダブルクォーテーションで囲まれていないフィールドを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInField(c As Char) As CsvParserState

            If _Config.IsDoubleQuote(c) Then
                If _Config.AllowQuoteInField Then
                    Return CsvParserState.Field
                Else
                    'フィールドの値にダブルクォーテーションが含まれている場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.NotAllowedCharException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                End If

            ElseIf _Config.IsDelimiter(c) Then
                Return CsvParserState.Start

            ElseIf _Config.IsCr(c) Then
                If _Config.NewLineChar = ControlChars.CrLf Then
                    Return CsvParserState.CarriageReturn
                ElseIf _Config.NewLineChar = ControlChars.Lf Then
                    Return CsvParserState.Field
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsLf(c) Then
                If (_Config.NewLineChar = ControlChars.CrLf) OrElse (_Config.NewLineChar = ControlChars.Cr) Then
                    Return CsvParserState.Field
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsWhiteSpace(c) Then
                If _Config.TrimWhiteSpace Then
                    Return CsvParserState.WhiteSpace
                Else
                    Return CsvParserState.Field
                End If

            Else
                Return CsvParserState.Field

            End If

        End Function

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、開きダブルクォーテーションを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInOpenQuote(c As Char) As CsvParserState

            If _Config.IsDoubleQuote(c) Then
                Return CsvParserState.CloseQuote

            ElseIf _Config.IsCr(c) Then
                If _Config.NewLineChar = ControlChars.CrLf Then
                    Return CsvParserState.CarriageReturn
                ElseIf _Config.NewLineChar = ControlChars.Lf Then
                    Return CsvParserState.Quoted
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.Quoted
                End If

            ElseIf _Config.IsLf(c) Then
                If (_Config.NewLineChar = ControlChars.CrLf) OrElse (_Config.NewLineChar = ControlChars.Cr) Then
                    Return CsvParserState.Quoted
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.Quoted
                End If

            Else
                Return CsvParserState.Quoted
            End If

        End Function

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInQuoted(c As Char) As CsvParserState
            Return Me.ParseInOpenQuote(c)
        End Function

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、エスケープされたダブルクォーテーションを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInEscapedQuote(c As Char) As CsvParserState
            Return Me.ParseInOpenQuote(c)
        End Function

        ''' <summary>
        ''' ダブルクォーテーションで囲まれたフィールドで、閉じダブルクォーテーションを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        ''' <remarks>閉じダブルクォーテーションの後はダブルクォーテーション、半角スペース、区切り文字、改行文字以外はエラーとします。</remarks>
        Protected Overridable Function ParseInCloseQuote(c As Char) As CsvParserState

            If _Config.IsDelimiter(c) Then
                Return CsvParserState.Start

            ElseIf _Config.IsDoubleQuote(c) Then
                Return CsvParserState.EscapedQuote

            ElseIf _Config.IsCr(c) Then
                If _Config.NewLineChar = ControlChars.CrLf Then
                    Return CsvParserState.CarriageReturn
                ElseIf _Config.NewLineChar = ControlChars.Lf Then
                    '閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字が見つかった場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.MalformedException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsLf(c) Then
                If (_Config.NewLineChar = ControlChars.CrLf) OrElse (_Config.NewLineChar = ControlChars.Cr) Then
                    '閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字が見つかった場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.MalformedException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            ElseIf _Config.IsWhiteSpace(c) Then
                Return CsvParserState.WhiteSpace

            Else
                '閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字が見つかった場合、例外をスローします。
                Me.AheadResult.Exceptions.Add(New Exceptions.MalformedException(Me.AheadResult.Clone))
                Return CsvParserState.Error
            End If

        End Function

        ''' <summary>
        ''' コメント行の開始文字を検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInComment(c As Char) As CsvParserState

            If _Config.IsCr(c) Then
                If _Config.NewLineChar = ControlChars.CrLf Then
                    Return CsvParserState.CarriageReturn
                ElseIf _Config.NewLineChar = ControlChars.Lf Then
                    Return CsvParserState.Comment
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.Start
                End If

            ElseIf _Config.IsLf(c) Then
                If (_Config.NewLineChar = ControlChars.CrLf) OrElse (_Config.NewLineChar = ControlChars.Cr) Then
                    Return CsvParserState.Comment
                Else
                    _DetectedNewLine = True
                    Return CsvParserState.Start
                End If

            Else
                Return CsvParserState.Comment
            End If

        End Function

        ''' <summary>
        ''' 改行文字(CR)を検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <param name="fieldState">文字を解析した結果、決定したフィールドの値に関する新しい状態。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        ''' <remarks>構成定義の改行文字に CrLf が指定されている場合に使用します。</remarks>
        Protected Overridable Function ParseInCarriageReturn(c As Char, fieldState As CsvParserState) As CsvParserState

            If _Config.IsLf(c) Then
                If (fieldState = CsvParserState.OpenQuote) OrElse (fieldState = CsvParserState.Quoted) Then
                    _DetectedNewLine = True
                    Return CsvParserState.Quoted
                ElseIf fieldState = CsvParserState.Comment Then
                    Return CsvParserState.Start
                Else 'CsvParserState.Start, CsvParserState.Field, CsvParserState.CloseQuote
                    _DetectedNewLine = True
                    Return CsvParserState.EndOfRecord
                End If

            Else 'CsvParserState.Start, CsvParserState.Field, CsvParserState.OpenQuote, CsvParserState.Quoted, CsvParserState.CloseQuote
                If (fieldState = CsvParserState.Start) OrElse (fieldState = CsvParserState.Field) Then
                    Return Me.ParseInField(c)
                ElseIf (fieldState = CsvParserState.OpenQuote) OrElse (fieldState = CsvParserState.Quoted) Then
                    Return Me.ParseInQuoted(c)
                ElseIf fieldState = CsvParserState.Comment Then
                    Return Me.ParseInComment(c)
                Else 'CsvParserState.CloseQuote
                    '閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字が見つかった場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.MalformedException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                End If

            End If

        End Function

        ''' <summary>
        ''' 半角スペースを検出した状態で読み込んだ文字を解析します。
        ''' </summary>
        ''' <param name="c">解析する文字。</param>
        ''' <param name="fieldState">文字を解析した結果、決定したフィールドの値に関する新しい状態。</param>
        ''' <returns>文字を解析した後の状態。</returns>
        Protected Overridable Function ParseInWhiteSpace(c As Char, fieldState As CsvParserState) As CsvParserState

            If fieldState = CsvParserState.Start Then
                Return Me.ParseInStart(c)

            ElseIf fieldState = CsvParserState.CloseQuote Then
                If _Config.IsDoubleQuote(c) Then
                    '閉じダブルクォーテーションの後にダブルクォーテーション、半角スペース、区切り文字、改行文字以外の文字が見つかった場合、例外をスローします。
                    Me.AheadResult.Exceptions.Add(New Exceptions.MalformedException(Me.AheadResult.Clone))
                    Return CsvParserState.Error
                Else
                    Return Me.ParseInCloseQuote(c)
                End If

            Else 'CsvParserState.Field
                Return Me.ParseInField(c)

            End If

        End Function

        ''' <summary>
        ''' 現在行のすべてのフィールドを読み込んで文字列の配列として返します。
        ''' </summary>
        ''' <returns>現在の行のフィールド値を格納する文字列の配列。</returns>
        Protected Overridable Function ParseLine() As String()

            Dim fieldState As CsvParserState = CsvParserState.Start
            Dim currentState As CsvParserState = CsvParserState.Start
            Dim newState As CsvParserState = CsvParserState.Start
            Dim c As Char = Me.ReadChar

            While Not c.Equals(Char.MinValue)

                If currentState = CsvParserState.Start Then : newState = Me.ParseInStart(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.Field Then : newState = Me.ParseInField(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.OpenQuote Then : newState = Me.ParseInOpenQuote(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.Quoted Then : newState = Me.ParseInQuoted(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.CloseQuote Then : newState = Me.ParseInCloseQuote(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.EscapedQuote Then : newState = Me.ParseInEscapedQuote(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.Comment Then : newState = Me.ParseInComment(c) : fieldState = currentState
                ElseIf currentState = CsvParserState.CarriageReturn Then : newState = Me.ParseInCarriageReturn(c, fieldState)
                ElseIf currentState = CsvParserState.WhiteSpace Then : newState = Me.ParseInWhiteSpace(c, fieldState)
                End If

                Me.AheadResult.Add(c, _LineNumber, _CharIndex, fieldState, currentState, newState)

                If (newState = CsvParserState.EndOfRecord) AndAlso (newState <> CsvParserState.Error) AndAlso (Not Me.AheadResult.HasFields) Then
                    If _Config.IgnoreBlankLines Then
                        newState = CsvParserState.Start
                    Else
                        Me.AheadResult.AddEmptyField(_SaveFieldCount)
                    End If
                End If

                If (newState = CsvParserState.EndOfRecord) OrElse (newState = CsvParserState.Error) Then
                    Return Me.AheadResult.Fields()
                End If

                If _DetectedNewLine Then Me.NewLineDetected()

                currentState = newState
                c = Me.ReadChar

            End While

            '最後のレコードの末尾に改行文字がない場合、ここで処理します。
            If (currentState = CsvParserState.Quoted) OrElse (currentState = CsvParserState.EscapedQuote) Then
                '(u) 閉じダブルクォーテーションが見つからなかった場合、例外をスローします。
                Me.AheadResult.Exceptions.Add(New MissingCloseQuoteException(Me.AheadResult.Clone))
            Else
                Me.AheadResult.Add(Char.MinValue, -1, -1, fieldState, currentState, CsvParserState.EndOfRecord)
            End If

            If (Not _Config.IgnoreBlankLines) AndAlso (Not Me.AheadResult.HasFields) AndAlso (newState <> CsvParserState.Error) Then
                Me.AheadResult.AddEmptyField(_SaveFieldCount)
            End If

            Return Me.AheadResult.Fields()

        End Function

#End Region

#Region "Release"

        ''' <summary>
        ''' Reader を解放します。
        ''' </summary>
        Protected Overridable Sub Release()

            Try
                If _Reader IsNot Nothing Then
                    _Reader.Close()
                    _Reader.Dispose()
                End If

            Catch ex As Exception

            Finally
                _Reader = Nothing
            End Try

        End Sub

#End Region

#Region "コンストラクタ"

        ''' <summary>
        ''' Parser.CsvParser クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="path">解析するファイルの完全なパス。</param>
        Public Sub New(path As String)
            Me.New(path, New CsvConfiguration)
        End Sub

        ''' <summary>
        ''' Parser.CsvParser クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="path">解析するファイルの完全なパス。</param>
        ''' <param name="config">CSVファイルの解析に使用する構成定義。</param>
        Public Sub New(path As String, config As CsvConfiguration)
            Me.New(New StreamReader(path, config.Encoding), config)
        End Sub

        ''' <summary>
        ''' Parser.CsvParser クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="reader">解析するリーダー。</param>
        Public Sub New(reader As TextReader)
            Me.New(reader, New CsvConfiguration)
        End Sub

        ''' <summary>
        ''' Parser.CsvParser クラスの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="reader">解析するリーダー。</param>
        ''' <param name="config">CSVファイルの解析に使用する構成定義。</param>
        Public Sub New(reader As TextReader, config As CsvConfiguration)

            _Reader = reader
            _Config = config
            _Config.Validate()
            _LineNumber = 1

            _BufferChars = New Char(DefaultBufferSize - 1) {}

            ReDim _ParseResults(1)

            For i = 0 To _ParseResults.Count - 1
                _ParseResults(i) = New CsvParseResult(_Config)
                AddHandler _ParseResults(i).CsvRecordParseCompleted, AddressOf ParseResult_CsvRecordParseCompleted
            Next

            Me.InitialRead()

        End Sub

#End Region

    End Class

End Namespace
