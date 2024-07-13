Imports ClosedXML

Public Class StartForm

#If CONFIG = "TextExtractor" Then
    Public Const IgnoreTextMistakes As Boolean = True
#Else
    Public Const IgnoreTextMistakes As Boolean = False
#End If

    Private Function LangDictionaryPath() As String
#If CONFIG = "TextExtractor" Then
        Return PathTextBox.Text & ".extracted.txt"
#Else
        Return PathTextBox.Text & ".dictionary.txt"
#End If
    End Function
    Private Function MapPath() As String
        Return PathTextBox.Text
    End Function
    Private Function MapTranslatedPath() As String
        If Reader.GetFileType(MapPath) = Reader.FileType.Map Then
            Return PathTextBox.Text & ".translated.sg"
        ElseIf Reader.GetFileType(MapPath) = Reader.FileType.Campaign Then
            Return PathTextBox.Text & ".translated.csg"
        Else
            Throw New Exception("Unexpected file type")
        End If
    End Function
    Private Function TGlobal1Path() As String
        Return TglobalTextBox1.Text
    End Function
    Private Function Tglobal2Path() As String
        Return TglobalTextBox2.Text
    End Function
    Private Function Autotranslate() As Boolean
        Return AutotranslateCheckBox.Checked
    End Function
    Private Function ResourcesDir()
        Return ".\Resources"
    End Function

    Sub parse() Handles ParseButton.Click
        Call WriteLastPathFile()
        If Not Writer.CheckCanWrite(LangDictionaryPath) Then Exit Sub
        Dim p As New Parser
        Dim f As List(Of String) = p.Parse(Reader.ReadFile(MapPath), Reader.GetFileType(MapPath))
        Dim L As Dictionary(Of String, String) = Nothing
        Dim DBFL As Dictionary(Of String, String) = Nothing
        If Autotranslate() Then DBFL = Translator.DBFLangDictionary(TGlobal1Path, Tglobal2Path)
        Dim XlL As Dictionary(Of String, String) = Reader.ReadExelDictionaryFolder(ResourcesDir)
        Call AppendDictionary(DBFL, XlL)
        If IO.File.Exists(LangDictionaryPath) Then
            Dim t As New Translator
            L = t.ReadLangDictionary(LangDictionaryPath)
        End If
        Call (New Writer).Write(LangDictionaryPath, f, L, DBFL)
        MsgBox("done")
    End Sub
    Sub make() Handles MakeButton.Click
        Call WriteLastPathFile()
        If Not Writer.CheckCanWrite(MapTranslatedPath) Then Exit Sub
        Dim t As New Translator
        Dim L As Dictionary(Of String, String) = t.ReadLangDictionary(LangDictionaryPath)
        Dim DBFL As Dictionary(Of String, String) = Nothing
        If Autotranslate() Then DBFL = Translator.DBFLangDictionary(TGlobal1Path, Tglobal2Path)
        Dim XlL As Dictionary(Of String, String) = Reader.ReadExelDictionaryFolder(ResourcesDir)
        Call AppendDictionary(DBFL, XlL)
        Dim f() As Byte = Reader.ReadFile(MapPath)
        Dim h As Parser.CSGHeader = Nothing
        If Reader.GetFileType(MapPath) = Reader.FileType.Campaign Then
            h = New Parser.CSGHeader(f)
            Call h.CheckFile(f)
        End If
        Dim translated() As Byte = t.Translate(L, DBFL, f, h, Reader.GetFileType(MapPath))
        If Reader.GetFileType(MapPath) = Reader.FileType.Campaign Then
            Call h.CheckFile(translated)
            Call h.RewriteHeader(translated)
            Call h.CheckFile(translated)
        End If
        IO.File.WriteAllBytes(MapTranslatedPath, translated)
        MsgBox("done")
    End Sub
    Sub test() Handles TestButton.Click
        Dim t As New Translator
        Dim f() As Byte = Reader.ReadFile(MapPath)
        Dim parsed As List(Of String) = (New Parser).Parse(Reader.ReadFile(MapPath), Reader.GetFileType(MapPath))
        Dim L As New Dictionary(Of String, String)
        For Each line As String In parsed
            L.Add(line, line)
        Next line
        Dim h As Parser.CSGHeader = Nothing
        Dim h_test As Parser.CSGHeader = Nothing
        If Reader.GetFileType(MapPath) = Reader.FileType.Campaign Then
            h = New Parser.CSGHeader(f)
            h_test = New Parser.CSGHeader(f)
            Call h.CheckFile(f)
        End If
        Dim translated() As Byte = t.Translate(L, Nothing, f, h, Reader.GetFileType(MapPath), h_test)
        If Reader.GetFileType(MapPath) = Reader.FileType.Campaign Then
            Call h.CheckFile(translated)
            Call h.RewriteHeader(translated)
            Call h.CheckFile(translated)
        End If
        For i As Integer = 0 To UBound(f) Step 1
            If Not f(i) = translated(i) Then
                Throw New Exception
            End If
        Next i
        If Not f.Length = translated.Length Then
            Throw New Exception
        End If
        MsgBox("test done")
    End Sub

    Sub meload() Handles Me.Load
        TestButton.Visible = False
        PathTextBox.Text = Reader.GetMapPath
        TglobalTextBox1.Text = Reader.GetTglobal1Path
        TglobalTextBox2.Text = Reader.GetTglobal2Path
        AutotranslateCheckBox.Checked = Reader.GetAutotranslateState
#If CONFIG = "TextExtractor" Then
        TglobalTextBox2.Visible = False
        Tglobal2Button.Visible = False
        AutotranslateCheckBox.Visible = False
        AutotranslateCheckBox.Checked = False
        MakeButton.Visible = False
        HelpB.Visible = False
#End If
    End Sub
    Private Sub WriteLastPathFile()
        Call Writer.WriteLastPathFile(PathTextBox.Text, TglobalTextBox1.Text, TglobalTextBox2.Text, AutotranslateCheckBox.Checked)
    End Sub

    Private Sub AppendDictionary(ByRef destination As Dictionary(Of String, String), _
                                 ByRef source As Dictionary(Of String, String))
        If IsNothing(destination) Then destination = New Dictionary(Of String, String)
        Dim keys As List(Of String) = source.Keys.ToList
        For Each k As String In keys
            If Not destination.Keys.Contains(k) Then destination.Add(k, source.Item(k))
        Next k
    End Sub

    Private Sub SelectFile(obj As Object, e As EventArgs) Handles MapButton.Click, Tglobal1Button.Click, Tglobal2Button.Click
        Dim buttons() As Button = {MapButton, Tglobal1Button, Tglobal2Button}
        Dim boxes() As TextBox = {PathTextBox, TglobalTextBox1, TglobalTextBox2}
        Dim filters() As String = {"Map files (*.sg)|*.sg|Campaign files (*.csg)|*.csg", "TGlobal table (*.dbf)|*.dbf", "TGlobal table (*.dbf)|*.dbf"}
        For i As Integer = 0 To UBound(buttons) Step 1
            If obj.Equals(buttons(i)) Then
                Dim d As New OpenFileDialog
                Try
                    If IO.Directory.Exists(boxes(i).Text) Then
                        d.InitialDirectory = boxes(i).Text
                    Else
                        d.InitialDirectory = System.IO.Path.GetDirectoryName(boxes(i).Text)
                    End If
                Catch
                    d.InitialDirectory = "C:\"
                End Try
                d.Filter = filters(i)
                Dim path As String
                If d.ShowDialog() = DialogResult.OK Then
                    path = d.FileName
                Else
                    Exit Sub
                End If
                boxes(i).Text = path
                Exit For
            End If
        Next i
    End Sub

    Sub help() Handles HelpB.Click

        Dim msg As String = "В текстбокс забиваем адрес файла карты." & vbNewLine &
                            "Кнопка Parse - читает файл и сохраняет блоки с текстовыми сообщениями" & vbNewLine &
                            "               в файл 'путь к карте'.dictionary.txt." & vbNewLine &
                            "               Если файл с переводом существует, то в его конец допишет" & vbNewLine &
                            "               новый текст, появившийся в карте" & vbNewLine &
                            "               (измененный текст считается новым)" & vbNewLine &
                            "Кнопка Make  - собирает новый файл карты. Читает оригинальный файл\" & vbNewLine &
                            "               карты и файл с переводом и проверяет длину текста." & vbNewLine &
                            "               D2 не знает буквы ё, поэтому она автоматически." & vbNewLine &
                            "               заменяется на е." & vbNewLine &
                            "Можно добавить собственные словари (например, с именами/названиями чего-либо). " &
                            "Для этого нужно закинуть файл xmlx в папку Resources. Будут прочтены все листы файла." &
                            "Как организован сам словарь: в первом столбце английское слово или фраза, " &
                            "а в последующих столбцах один или несколько вариантов перевода на русский."
        MsgBox(msg)
    End Sub

#If CONFIG = "TextExtractor" Then
    Private Sub TglobalTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TglobalTextBox1.TextChanged
        TglobalTextBox2.Text = TglobalTextBox1.Text
    End Sub
#End If
End Class

Class Reader
    Public Const encID As Integer = 1251

    Public Shared Function ReadFile(ByRef path As String) As Byte()
        Return IO.File.ReadAllBytes(path)
    End Function

    Public Shared Function GetMapPath() As String
        Return ReadLastPathFile(0)
    End Function
    Public Shared Function GetTglobal1Path() As String
        Return ReadLastPathFile(1)
    End Function
    Public Shared Function GetTglobal2Path() As String
        Return ReadLastPathFile(2)
    End Function
    Public Shared Function GetAutotranslateState() As String
        Return CBool(ReadLastPathFile(3))
    End Function
    Private Shared Function ReadLastPathFile(ByRef line As Integer) As String
        If IO.File.Exists(My.Resources.lastPathFile) Then
            Dim lines() As String = IO.File.ReadAllLines(My.Resources.lastPathFile)
            If line <= UBound(lines) Then
                Return lines(line)
            End If
        End If
        If line = 3 Then
            Return "True"
        Else
            Return "C:\"
        End If
    End Function

    Public Shared Function ReadExelDictionaryFolder(ByVal path As String) As Dictionary(Of String, String)
        Dim dictionary As New Dictionary(Of String, String)
        Dim files() As String = IO.Directory.GetFiles(path)
        For i As Integer = 0 To UBound(files) Step 1
            If files(i).ToLower.EndsWith(".xlsx") Then
                Call ReadExcelDictionary(files(i), dictionary)
            End If
        Next i
        Return dictionary
    End Function
    Public Shared Sub ReadExcelDictionary(ByVal path As String, ByRef destDictionary As Dictionary(Of String, String))
        Dim f As New IO.FileInfo(path)
        Dim w As New Excel.XLWorkbook(path)
        Dim engWord, rusWord As String
        For Each sheet As Excel.IXLWorksheet In w.Worksheets
            Dim rows As Excel.IXLRangeRows = sheet.RangeUsed.RowsUsed()
            For Each r As Excel.IXLRangeRow In rows
                Dim cells As Excel.IXLCells = r.CellsUsed
                engWord = "" : rusWord = ""
                For Each c As Excel.IXLCell In cells
                    If c.Address.ColumnNumber = 1 Then
                        engWord = c.Value.ToString.Trim
                    Else
                        rusWord = c.Value.ToString.Trim
                        If engWord = "" Then
                            MsgBox("Empty english word at row " & c.Address.RowNumber & "(Sheet: " & sheet.Name & ")")
                            End
                        ElseIf rusWord = "" Then
                            MsgBox("Empty russian word at cell " & c.Address.ColumnLetter & c.Address.RowNumber & "(Sheet: " & sheet.Name & ")")
                            End
                        Else
                            If Not StartForm.IgnoreTextMistakes Then
                                For Each word As String In {engWord, rusWord}
                                    If word.Contains("  ") Then
                                        MsgBox("Following text contains double spaces in the excel file " &
                                               path & " (sheet " & sheet.Name & ") :" & word)
                                        End
                                    End If
                                Next word
                            End If
                        End If
                        If Not destDictionary.Keys.Contains(engWord.ToLower) Then
                            destDictionary.Add(engWord.ToLower, rusWord)
                        End If
                        If Not destDictionary.Keys.Contains(rusWord.ToLower) Then
                            destDictionary.Add(rusWord.ToLower, engWord)
                        End If
                    End If
                Next c
            Next r
        Next sheet
    End Sub

    Public Enum FileType
        Map = 1
        Campaign = 2
    End Enum
    Public Shared Function GetFileType(ByRef path As String) As FileType
        If path.ToLower.EndsWith(".sg") Then
            Return FileType.Map
        ElseIf path.ToLower.EndsWith(".csg") Then
            Return FileType.Campaign
        Else
            MsgBox("Unexpected file type. I expect .sg or .csg")
            End
        End If
    End Function

End Class

Class Writer

    Public Const BlockDelimiterKeyword As String = "--------- --------- --------- --------- --------- ---------" & vbNewLine
    Public Const OrigTextKeyword As String = "# Original text" & vbNewLine
    Public Const TransTextKeyword As String = "# Translated text" & vbNewLine
    Public Const SuggestionTextKeyword As String = "# Suggestion:"

    Private ReadOnly wordDelimiter As New List(Of Char)
    Public Sub New()
        wordDelimiter.AddRange(" `~!@#$%^&*()-=_+':;|\/№?,." & """" & "0123456789" & vbNewLine)
    End Sub

    Public Shared Function CheckCanWrite(ByRef path As String) As Boolean
        If IO.File.Exists(path) Then
            Dim answer As MsgBoxResult = MsgBox("Do you want to replace file " & path & "?", MsgBoxStyle.YesNo)
            If answer = MsgBoxResult.No Then
                Return False
#If CONFIG = "TextExtractor" Then
            Else
                Kill(path)
#End If
            End If
        End If
        Return True
    End Function

    Public Sub Write(ByRef path As String, ByRef content As List(Of String), _
                     Optional ByRef langDictionary As Dictionary(Of String, String) = Nothing, _
                     Optional ByRef DBFLangDictionary As Dictionary(Of String, String) = Nothing)
        Dim out() As String
        Dim printed As New List(Of String)
        Dim n As Integer = -1
        If IsNothing(langDictionary) Then
            ReDim out(content.Count - 1)
        Else
            ReDim out(content.Count + langDictionary.Count - 2)
            Dim keys As List(Of String) = langDictionary.Keys.ToList
            For Each k As String In keys
                n += 1
                out(n) = PrintLine(k, langDictionary.Item(k))
                printed.Add(k)
            Next k
        End If
        For Each s As String In content
            If Not printed.Contains(s) Then
                n += 1
                If Not IsNothing(DBFLangDictionary) AndAlso DBFLangDictionary.ContainsKey(s.ToLower) Then
                    out(n) = PrintLine(s, DBFLangDictionary.Item(s.ToLower))
                Else
                    out(n) = PrintLine(s, MakeSuggestion(s, DBFLangDictionary))
                End If
                printed.Add(s)
            End If
        Next s
        If UBound(out) > n Then ReDim Preserve out(n)
        IO.File.WriteAllLines(path, out, System.Text.Encoding.GetEncoding(Reader.encID))
    End Sub
    Private Function MakeSuggestion(ByRef text As String, ByRef DBFLangDictionary As Dictionary(Of String, String)) As String
        Dim result As String = ""
        If Not IsNothing(DBFLangDictionary) Then
            Dim tLower As String = text.ToLower
            For Each k As String In DBFLangDictionary.Keys
                If IsWord(tLower, k) Then
                    result &= vbNewLine & k & " -> " & DBFLangDictionary.Item(k)
                End If
            Next k
            If Not result = "" Then result = Writer.SuggestionTextKeyword & result
        End If
        Return result
    End Function
    Private Function IsWord(ByRef text As String, ByRef word As String) As Boolean
        Dim searchFrom As Integer = 0
        Dim checkStartResult, checkEndResult As Boolean
        Dim c As Char
        Do While searchFrom > -1
            searchFrom = text.IndexOf(word, searchFrom)
            If searchFrom > -1 Then
                checkStartResult = False
                checkEndResult = False
                If searchFrom = 0 Then
                    checkStartResult = True
                Else
                    c = text(searchFrom - 1)
                    If wordDelimiter.Contains(c) Then checkStartResult = True
                End If
                If checkStartResult Then
                    If searchFrom + word.Length = text.Length Then
                        checkEndResult = True
                    Else
                        c = text(searchFrom + word.Length)
                        If wordDelimiter.Contains(c) Then checkEndResult = True
                    End If
                End If
                If checkStartResult And checkEndResult Then Return True
                searchFrom += 1
            End If
        Loop
        Return False
    End Function

    Private Shared Function PrintLine(ByRef original As String, ByRef translation As String) As String
#If CONFIG = "TextExtractor" Then
        Return BlockDelimiterKeyword & _
               original
#Else
        Return BlockDelimiterKeyword & _
               OrigTextKeyword & _
               BlockDelimiterKeyword & _
               original & vbNewLine & _
               BlockDelimiterKeyword & _
               TransTextKeyword & _
               BlockDelimiterKeyword & _
               translation
#End If
    End Function

    Public Shared Sub WriteLastPathFile(ByRef mapFile As String, ByRef tglobal1 As String, ByRef tglobal2 As String, _
                                        ByRef AutotranslateCheckBox As Boolean)
        Dim lines() As String = {mapFile, tglobal1, tglobal2, AutotranslateCheckBox.ToString}
        IO.File.WriteAllLines(My.Resources.lastPathFile, lines)
    End Sub

End Class

Class Parser

    Dim dataBlocks() As Block

    Public Class CSGHeader

        Public Const RecordsCount1Byte As Integer = 92

        Public records() As MQRCRecord

        Public Class MQRCRecord
            Public dataID As IntField
            Public size As SizeField
            Public pos As IntLinkField

            Public subHeader As Subrecord

            Public Class IntField
                Public value As Integer
                Public bias As Integer
                Public size As Integer

                Public Sub New(_value As Integer, _bias As Integer, ByVal _valueLen As Integer)
                    value = _value
                    bias = _bias
                    size = _valueLen
                End Sub
                Public Sub New(_content() As Byte, _bias As Integer, ByVal _valueLen As Integer)
                    If _valueLen = 1 Then
                        value = _content(_bias)
                    ElseIf _valueLen = 2 Then
                        value = BitConverter.ToInt16(_content, _bias)
                    ElseIf _valueLen = 4 Then
                        value = BitConverter.ToInt32(_content, _bias)
                    ElseIf _valueLen = 8 Then
                        value = BitConverter.ToInt64(_content, _bias)
                    Else
                        Throw New Exception("Unexpected value length: " & _valueLen)
                    End If
                    bias = _bias
                    size = _valueLen
                End Sub

                Public Overridable Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                    If bias >= changedAt Then bias += lenghtChange
                End Sub

                Public Sub PrintToArray(ByRef dest() As Byte)
                    Dim b() As Byte = BitConverter.GetBytes(value)
                    For i As Integer = 0 To size - 1 Step 1
                        dest(bias + i) = b(i)
                    Next i
                End Sub

                Public Function Print() As String
                    Return value & "  :  " & bias
                End Function
            End Class
            Public Class IntLinkField
                Inherits IntField
                Public Sub New(ByVal _value As Integer, _bias As Integer, ByVal _valueLen As Integer)
                    Call MyBase.New(_value, _bias, _valueLen)
                End Sub
                Public Sub New(_content() As Byte, _bias As Integer, ByVal _valueLen As Integer)
                    Call MyBase.New(_content, _bias, _valueLen)
                End Sub

                Public Overrides Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                    If bias >= changedAt Then bias += lenghtChange
                    If value >= changedAt Then value += lenghtChange
                End Sub
            End Class
            Public Class StrField
                Public value As String
                Public bias As Integer

                Public Sub New(_value As String, _bias As Integer)
                    value = _value
                    bias = _bias
                End Sub
                Public Sub New(_content() As Byte, _bias As Integer, ByVal valueLen As Integer)
                    Dim byteString(valueLen - 1) As Byte
                    For i As Integer = 0 To valueLen - 1 Step 1
                        byteString(i) = _content(_bias + i)
                    Next i
                    value = Converter.ToStr(byteString)
                    bias = _bias
                End Sub

                Public Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                    If bias >= changedAt Then bias += lenghtChange
                End Sub

                Public Function Print() As String
                    Return value & "  :  " & bias
                End Function
            End Class
            Public Class SizeField
                Public data As IntField
                Public block As IntField
                Public signature As StrField

                Public Sub New(ByVal _dataValue As Integer, ByVal _dataBias As Integer, ByVal _dataValueLen As Integer, _
                               ByVal _blockValue As Integer, ByVal _blockBias As Integer, ByVal _blockValueLen As Integer, _
                               ByRef _signature As StrField)
                    data = New IntField(_dataValue, _dataBias, _dataValueLen)
                    block = New IntField(_blockValue, _blockBias, _blockValueLen)
                    signature = _signature
                End Sub
                Public Sub New(ByVal _content() As Byte, _
                               ByVal _dataBias As Integer, ByVal _dataValueLen As Integer, _
                               ByVal _blockBias As Integer, ByVal _blockValueLen As Integer, _
                               ByRef _signature As StrField)
                    data = New IntField(_content, _dataBias, _dataValueLen)
                    block = New IntField(_content, _blockBias, _blockValueLen)
                    signature = _signature
                End Sub

                Public Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                    Call data.BytesAdded(changedAt, lenghtChange)
                    Call block.BytesAdded(changedAt, lenghtChange)

                    Dim blockStartByte As Integer = Subrecord._headerLength + signature.bias
                    Dim blockEndByte As Integer = blockStartByte + block.value
                    If changedAt >= blockStartByte And changedAt <= blockEndByte Then
                        If data.value = block.value Then
                            data.value += lenghtChange
                            block.value += lenghtChange
                        Else
                            data.value += lenghtChange
                            If data.value > block.value Then
                                Throw New Exception("Unexpected data size")
                                End
                            End If
                        End If
                    End If
                End Sub

                Public Sub PrintToArray(ByRef dest() As Byte)
                    Call data.PrintToArray(dest)
                    Call block.PrintToArray(dest)
                End Sub

                Public Function Print(ByVal prefix As String) As String
                    Return prefix & data.Print & vbNewLine & _
                           prefix & block.Print
                End Function
            End Class
            Public Class Subrecord
                Public signature As StrField
                Public size As SizeField

                Public Const _dataSizeByte As Integer = 12
                Public Const _blockSizeByte As Integer = 16
                Public Const _headerLength As Integer = 28

                Public Sub New(ByRef content() As Byte, ByRef pos As Integer)
                    signature = New StrField(content, pos, 4)
                    If Not signature.value = "MQRC" Then
                        MsgBox("Unexpected block signature: " & signature.value)
                        End
                    End If
                    size = New SizeField(content, pos + _dataSizeByte, 4, pos + _blockSizeByte, 4, signature)
                End Sub

                Public Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                    Call size.BytesAdded(changedAt, lenghtChange)
                    Call signature.BytesAdded(changedAt, lenghtChange)
                End Sub

                Public Sub PrintToArray(ByRef dest() As Byte)
                    Call size.PrintToArray(dest)
                End Sub
            End Class

            Public Const _recordIDByte As Integer = 0
            Public Const _dataSizeByte As Integer = 4
            Public Const _blockSizeByte As Integer = 8
            Public Const _posByte As Integer = 12
            Public Const _recordLength As Integer = 16

            Public Sub New(ByRef content() As Byte, ByRef start As Integer)
                dataID = New IntField(content, start + _recordIDByte, 1)
                pos = New IntLinkField(content, start + _posByte, 4)
                subHeader = New Subrecord(content, pos.value)
                size = New SizeField(content, start + _dataSizeByte, 4, start + _blockSizeByte, 4, subHeader.signature)
                If Not size.data.value = subHeader.size.data.value Then
                    MsgBox("Unexpected data size")
                    End
                ElseIf Not size.block.value = subHeader.size.block.value Then
                    MsgBox("Unexpected block size")
                    End
                End If
            End Sub

            Public Sub BytesAdded(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
                Call size.BytesAdded(changedAt, lenghtChange)
                Call subHeader.BytesAdded(changedAt, lenghtChange)
                Call dataID.BytesAdded(changedAt, lenghtChange)
                Call pos.BytesAdded(changedAt, lenghtChange)
            End Sub

            Public Sub PrintToArray(ByRef dest() As Byte)
                Call dataID.PrintToArray(dest)
                Call size.PrintToArray(dest)
                Call pos.PrintToArray(dest)
                Call subHeader.PrintToArray(dest)
            End Sub

            Public Function Print() As String
                Return dataID.Print & vbNewLine & _
                       size.Print("") & vbNewLine & _
                       pos.Print & vbNewLine & _
                       " > " & subHeader.signature.Print & vbNewLine & _
                       subHeader.size.Print(" > ")
            End Function
        End Class

        Public Sub New(ByRef content() As Byte)
            Dim rCount1, rCount2 As Integer
            rCount1 = content(RecordsCount1Byte)
            ReDim records(rCount1 - 1)
            For i As Integer = 0 To UBound(records) Step 1
                records(i) = New MQRCRecord(content, RecordsCount1Byte + 4 + i * MQRCRecord._recordLength)
                Console.WriteLine(i & "  -----------------")
                Console.WriteLine(records(i).Print)
            Next i
            Console.WriteLine("#################")
            Dim RecordsCount2Byte As Integer = RecordsCount1Byte + 4 + rCount1 * MQRCRecord._recordLength
            rCount2 = content(RecordsCount2Byte)
            ReDim Preserve records(UBound(records) + rCount2)
            For i As Integer = 0 To rCount2 - 1 Step 1
                Dim m As Integer = rCount1 + i
                records(m) = New MQRCRecord(content, RecordsCount2Byte + 4 + i * MQRCRecord._recordLength)
                Console.WriteLine(m & "  -----------------")
                Console.WriteLine(records(m).Print)
            Next i
        End Sub

        Public Sub FileLengthChanged(ByRef changedAt As Integer, ByRef lenghtChange As Integer)
            For i As Integer = 0 To UBound(records) Step 1
                Call records(i).BytesAdded(changedAt, lenghtChange)
            Next i
        End Sub

        Public Sub CheckFile(ByRef content() As Byte)
            Dim id As Integer
            Dim signature As String
            For i As Integer = 0 To UBound(records) Step 1
                id = (New MQRCRecord.IntField(content, records(i).dataID.bias, 1)).value
                If Not id = records(i).dataID.value Then Throw New Exception("Unexpected id")
                signature = (New MQRCRecord.StrField(content, records(i).subHeader.signature.bias, 4)).value
                If Not signature = records(i).subHeader.signature.value Then Throw New Exception("Unexpected signature")
            Next i
        End Sub
        Public Sub RewriteHeader(ByRef dest() As Byte)
            For i As Integer = 0 To UBound(records) Step 1
                Call records(i).PrintToArray(dest)
            Next i
        End Sub

        Public Shared Sub Compare(ByRef header1 As CSGHeader, ByRef header2 As CSGHeader)
            If IsNothing(header1) Or IsNothing(header2) Then Exit Sub
            For i As Integer = 0 To UBound(header1.records) Step 1
                Call Compare(header1.records(i), header2.records(i))
            Next i
        End Sub
        Public Shared Sub Compare(ByRef r1 As MQRCRecord, ByRef r2 As MQRCRecord)
            Try
                Call Compare(r1.dataID, r2.dataID)
                Call Compare(r1.size, r2.size)
                Call Compare(r1.pos, r2.pos)
                Call Compare(r1.subHeader.signature, r2.subHeader.signature)
                Call Compare(r1.subHeader.size, r2.subHeader.size)
                Call Compare(r1.subHeader.size, r2.subHeader.size)
            Catch ex As Exception
                Dim m1() As String = r1.Print.Replace(vbNewLine, Chr(13)).Split(Chr(13))
                Dim m2() As String = r2.Print.Replace(vbNewLine, Chr(13)).Split(Chr(13))
                For i As Integer = 0 To UBound(m1) Step 1
                    If m1(i) <> m2(i) Then
                        m1(i) &= "  <--"
                        m2(i) &= "  <--"
                    End If
                Next i
                MsgBox(ex.Message & vbNewLine & _
                       String.Join(vbNewLine, m1) & vbNewLine & _
                       "----------" & vbNewLine & _
                       String.Join(vbNewLine, m2))
            End Try
        End Sub
        Public Shared Sub Compare(ByRef f1 As MQRCRecord.SizeField, ByRef f2 As MQRCRecord.SizeField)
            Call Compare(f1.data, f2.data)
            Call Compare(f1.block, f2.block)
        End Sub
        Public Shared Sub Compare(ByRef f1 As MQRCRecord.IntField, ByRef f2 As MQRCRecord.IntField)
            If Not f1.value = f2.value Then Throw New Exception("Different value")
            If Not f1.bias = f2.bias Then Throw New Exception("Different bias")
        End Sub
        Public Shared Sub Compare(ByRef f1 As MQRCRecord.StrField, ByRef f2 As MQRCRecord.StrField)
            If Not f1.value = f2.value Then Throw New Exception("Different value")
            If Not f1.bias = f2.bias Then Throw New Exception("Different bias")
        End Sub

    End Class

    Public Class Block
        Public StartsWith As String
        Public EndsWith As String = "ENDOBJECT"
        Public TextBlocks() As TxtBlock

        Public byteStartsWith() As Byte
        Public byteEndsWith() As Byte

        Public ignoreByte() As Boolean = Nothing

        Public collectWords() As String = Nothing
        Public collectWordsByte()() As Byte = Nothing
        Public collectWordsIgnore()() As Boolean = Nothing

        Public Class TxtBlock
            Public owner As String
            Public subblocks() As String = Nothing
            Public maxTextLength As Integer

            Public expectWords As New List(Of String)

            Public byteSubblocks()() As Byte = Nothing
            Public ignoreByte()() As Boolean = Nothing

            Private Const subfieldsSplitter As String = "+"

            Public Sub New(ByRef description() As String, ByRef n As Integer, ByRef o As String, ByRef maxSize As Dictionary(Of String, Integer))
                Dim s() As String = description(n).Split(subfieldsSplitter)
                ReDim subblocks(UBound(s)), byteSubblocks(UBound(s)), ignoreByte(UBound(s))
                Dim str As String
                For i As Integer = 0 To UBound(s) Step 1
                    str = s(i)
                    If str.Contains("|") Then
                        Dim splited() As String = str.Split("|")
                        For k As Integer = 0 To UBound(splited) - 1 Step 1
                            expectWords.Add(splited(k))
                        Next k
                        str = splited(UBound(splited))
                    End If
                    subblocks(i) = str
                    byteSubblocks(i) = Converter.ToByteArray(str)
                    ignoreByte(i) = MakeIgnoreBytesArray(byteSubblocks(i))
                Next i
                owner = o
                maxTextLength = maxSize.Item(owner & " " & s(0))
            End Sub
            Protected Friend Shared Function StringConversion(ByRef s As String) As String
                Dim str As String = s
                Dim i1 As Integer = -1
                Dim i2 As Integer
                Do While Str.Contains("%")
                    i1 = Str.IndexOf("%", i1 + 1)
                    If i1 = -1 Then Exit Do
                    i2 = Str.IndexOf("%", i1 + 1)
                    If i2 = -1 Then Exit Do
                    Dim b As String = Str.Substring(i1 + 1, i2 - i1 - 1)
                    str = str.Substring(0, i1) & Converter.ToStr(b) & str.Substring(i2 + 1)
                Loop
                Return str
            End Function
            Protected Friend Shared Function MakeIgnoreBytesArray(ByRef byteString() As Byte) As Boolean()
                Dim r(UBound(byteString)) As Boolean
                Dim asterix As Byte = Converter.ToByteArray("*")(0)
                For k As Integer = 0 To UBound(byteString) Step 1
                    If byteString(k) = asterix Then r(k) = True
                Next k
                Return r
            End Function

            Public Function Check(ByRef fileText() As Byte, ByRef startByte As Integer, _
                                  ByRef collectedWords As List(Of String)) As CheckResult
                Dim r As New CheckResult With {.text = "", .sizeByte = -1, .textStartByte = -1}
                For Each w As String In expectWords
                    If Not collectedWords.Contains(w) Then Return r
                Next w
                For i As Integer = 0 To UBound(byteSubblocks) Step 1
                    If IsSearchedText(fileText, startByte, byteSubblocks(i), ignoreByte(i), True) Then
                        r.status = CheckResult.State.SubblockStart
                        r.maxTextLength = maxTextLength
                        startByte += byteSubblocks(i).Length
                        If r.sizeByte = -1 Then r.sizeByte = startByte
                        Dim L As Integer = fileText(startByte) - 2
                        If UBound(byteSubblocks) > 0 Then
                            L -= 1
                            r.isLongBlock = True
                        End If
                        startByte += 4

                        If L > 0 Then
                            If r.textStartByte = -1 Then r.textStartByte = startByte
                            Dim byteText(L) As Byte
                            For j As Integer = 0 To L Step 1
                                byteText(j) = fileText(startByte)
                                startByte += 1
                            Next j
                            If byteText.Length > r.maxTextLength Then
                                MsgBox("Unexpected text length " & byteText.Length & vbNewLine & "Max length is " & r.maxTextLength & _
                                       vbNewLine & owner & " " & subblocks(i) & vbNewLine & Converter.ToStr(byteText))
                            End If
                            r.text &= Converter.ToStr(byteText)
                            If UBound(byteSubblocks) > 0 Then
                                startByte += 2
                            Else
                                startByte += 1
                            End If
                        Else
                            startByte += 1
                        End If
                    Else
                        Exit For
                    End If
                Next i
                If Not r.text = "" Then
                    r.owner = owner
                    r.textEndByte = startByte - 1
                    r.block = subblocks
                    r.byteBlock = byteSubblocks
                End If
                Return r
            End Function
        End Class

        Public Sub New(ByRef description As String, ByRef maxSize As Dictionary(Of String, Integer))
            Dim s() As String = description.Split(" ")
            StartsWith = s(0)
            ReDim TextBlocks(UBound(s) - 1)
            Dim collect As New List(Of String)
            For i As Integer = 1 To UBound(s) Step 1
                TextBlocks(i - 1) = New TxtBlock(s, i, StartsWith, maxSize)
                For Each w In TextBlocks(i - 1).expectWords
                    If Not collect.Contains(w) Then collect.Add(w)
                Next w
            Next i
            byteStartsWith = Converter.ToByteArray(StartsWith)
            byteEndsWith = Converter.ToByteArray(EndsWith)
            ignoreByte = TxtBlock.MakeIgnoreBytesArray(byteStartsWith)
            Dim n As Integer = -1
            ReDim collectWords(collect.Count - 1), collectWordsByte(collect.Count - 1), collectWordsIgnore(collect.Count - 1)
            For Each w As String In collect
                n += 1
                collectWords(n) = w
                collectWordsByte(n) = Converter.ToByteArray(w)
                collectWordsIgnore(n) = TxtBlock.MakeIgnoreBytesArray(collectWordsByte(n))
            Next w
        End Sub

        Public Function Check(ByRef fileText() As Byte, ByRef startByte As Integer, ByRef readText As Boolean, _
                              ByRef collected As List(Of String)) As CheckResult
            If IsSearchedText(fileText, startByte, byteEndsWith, ignoreByte, False) Then
                Return New CheckResult With {.status = CheckResult.State.BlockEnd}
            ElseIf IsSearchedText(fileText, startByte, byteStartsWith, ignoreByte, False) Then
                If readText Then
                    MsgBox("Unexpected block start at byte: " & startByte)
                    End
                End If
                Return New CheckResult With {.status = CheckResult.State.BlockStart}
            ElseIf readText Then
                If collectWordsByte.Length > 0 Then
                    For i As Integer = 0 To UBound(collectWords) Step 1
                        If IsSearchedText(fileText, startByte, collectWordsByte(i), collectWordsIgnore(i), False) Then
                            If Not collected.Contains(collectWords(i)) Then collected.Add(collectWords(i))
                        End If
                    Next i
                End If
                Dim r As New CheckResult
                For i As Integer = 0 To UBound(TextBlocks) Step 1
                    r = TextBlocks(i).Check(fileText, startByte, collected)
                    If r.status = CheckResult.State.SubblockStart Then Exit For
                Next i
                Return r
            Else
                Return New CheckResult
            End If
        End Function
    End Class

    Public Sub New()
        Dim sizes() As String = SplitResourcesFile(My.Resources.TextMaxSize)
        Dim keywords() As String = SplitResourcesFile(My.Resources.StartTextKeywords)

        Dim maxSize As New Dictionary(Of String, Integer)
        For i As Integer = 0 To UBound(sizes) Step 1
            sizes(i) = Block.TxtBlock.StringConversion(sizes(i))
            Dim s() As String = sizes(i).Split(" ")
            maxSize.Add(s(0) & " " & s(1), s(2))
        Next i
        ReDim dataBlocks(UBound(keywords))
        For i As Integer = 0 To UBound(keywords) Step 1
            keywords(i) = Block.TxtBlock.StringConversion(keywords(i))
            dataBlocks(i) = New Block(keywords(i), maxSize)
        Next i
    End Sub
    Private Function SplitResourcesFile(ByRef f As String) As String()
        Return f.Replace(Chr(10), Chr(13)) _
                .Replace(vbLf, Chr(13)) _
                .Replace(Chr(13) & Chr(13), Chr(13)) _
                .Replace(Chr(13) & Chr(13), vbNewLine) _
                .Split(vbNewLine)
    End Function

    Public Structure CheckResult
        Public status As State
        Public text As String
        Public blockID As Integer

        Public owner As String
        Public block() As String
        Public byteBlock()() As Byte

        Public sizeByte As Integer
        Public textStartByte As Integer
        Public textEndByte As Integer
        Public maxTextLength As Integer

        Public isLongBlock As Boolean

        Public Enum State
            None = 0
            BlockStart = 1
            BlockEnd = 2
            SubblockStart = 3
        End Enum
    End Structure

    Public Const sgInitBlock As Integer = 42
    Public Const sgDescriptionBlock As Integer = sgInitBlock + 256
    Public Const sgAuthorBlock As Integer = sgDescriptionBlock + 22
    Public Const sgNameBlock As Integer = sgAuthorBlock + 64
    Public Const sgHostLordStart As Integer = 609
    Public Const sgHostLordEnd As Integer = sgHostLordStart + 14

    Public Const csgInitBlock As Integer = 1179
    Public Const csgNameBlock As Integer = csgInitBlock + 32
    Public Const csgAuthorStart As Integer = 1469
    Public Const csgAuthorEnd As Integer = csgAuthorStart + 31
    Public Const csgSGBias As Integer = 1836

    Private blockID As Integer = -1
    Private collectedWords As New List(Of String)

    Public Function Parse(ByRef fileText() As Byte, ByRef fileType As Reader.FileType) As List(Of String)
        Dim t As CheckResult
        Dim i As Integer = 0
        Dim r As New List(Of String)
        If fileType = Reader.FileType.Campaign Then
            t = GetCampaignName(fileText)
            If Not r.Contains(t.text) Then r.Add(t.text)
            t = GetCampaignAuthor(fileText)
            If Not r.Contains(t.text) Then r.Add(t.text)
        ElseIf Not fileType = Reader.FileType.Map Then
            Throw New Exception("Unexpected type")
        End If

        t = GetMapName(fileText, fileType)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetMapAuthor(fileText, fileType)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetMapDescription(fileText, fileType)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetMapHostLordName(fileText, fileType)
        If Not r.Contains(t.text) Then r.Add(t.text)
        i = sgHostLordEnd + 1
        If fileType = Reader.FileType.Campaign Then i += csgSGBias

        Do While i <= UBound(fileText)
            t = GetText(fileText, i)
            If Not t.text = "" AndAlso Not r.Contains(t.text) Then
                r.Add(t.text)
            End If
        Loop
        For Each s As String In r
            Call Translator.TestString(s)
        Next s
        Return r
    End Function

    Public Shared Function GetMapDescription(ByRef fileText() As Byte, ByRef fileType As Reader.FileType) As CheckResult
        Dim b As Integer = 0
        If fileType = Reader.FileType.Campaign Then b = csgSGBias
        Dim r As New CheckResult With {.textStartByte = sgInitBlock + 1 + b, .textEndByte = sgDescriptionBlock - 2 + b}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetMapAuthor(ByRef fileText() As Byte, ByRef fileType As Reader.FileType) As CheckResult
        Dim b As Integer = 0
        If fileType = Reader.FileType.Campaign Then b = csgSGBias
        Dim r As New CheckResult With {.textStartByte = sgDescriptionBlock + 1 + b, .textEndByte = sgAuthorBlock - 2 + b}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetMapName(ByRef fileText() As Byte, ByRef fileType As Reader.FileType) As CheckResult
        Dim b As Integer = 0
        If fileType = Reader.FileType.Campaign Then b = csgSGBias
        Dim r As New CheckResult With {.textStartByte = sgAuthorBlock + 1 + b, .textEndByte = sgNameBlock + b}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetMapHostLordName(ByRef fileText() As Byte, ByRef fileType As Reader.FileType) As CheckResult
        Dim b As Integer = 0
        If fileType = Reader.FileType.Campaign Then b = csgSGBias
        Dim r As New CheckResult With {.textStartByte = sgHostLordStart + b, .textEndByte = sgHostLordEnd + b}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function

    Public Shared Function GetCampaignName(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = csgInitBlock + 1, .textEndByte = csgNameBlock}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetCampaignAuthor(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = csgAuthorStart, .textEndByte = csgAuthorEnd}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function

    Private Shared Function ReadFromTo(ByRef fileText() As Byte, ByRef i1 As Integer, ByRef i2 As Integer, ByRef trim As Boolean) As Byte()
        Dim b(i2 - i1) As Byte
        For i As Integer = i1 To i2 Step 1
            b(i - i1) = fileText(i)
        Next i
        If trim Then Call TrimByteArray(b)
        Return b
    End Function
    Private Shared Sub TrimByteArray(ByRef b() As Byte)
        For i As Integer = UBound(b) To 0 Step -1
            If b(i) > 0 Then
                If i < UBound(b) Then ReDim Preserve b(i)
                Exit For
            End If
        Next i
    End Sub

    Public Function GetText(ByRef fileText() As Byte, ByRef startByte As Integer) As CheckResult
        Dim r As New CheckResult
        If blockID = -1 Then
            For i As Integer = 0 To UBound(dataBlocks) Step 1
                r = dataBlocks(i).Check(fileText, startByte, False, collectedWords)
                If r.status = CheckResult.State.BlockStart Then
                    startByte += dataBlocks(i).byteStartsWith.Length
                    blockID = i
                    Exit For
                ElseIf r.status = CheckResult.State.BlockEnd Then
                    startByte += dataBlocks(i).byteEndsWith.Length
                    Exit For
                End If
                r.text = ""
            Next i
            If Not r.status = CheckResult.State.BlockStart And Not r.status = CheckResult.State.BlockEnd Then
                startByte += 1
            End If
        Else
            r = dataBlocks(blockID).Check(fileText, startByte, True, collectedWords)
            If r.status = CheckResult.State.None Then
                startByte += 1
            End If
        End If
        If r.status = CheckResult.State.BlockEnd Then
            collectedWords.Clear()
            blockID = -1
        End If
        Return r
    End Function

    Public Shared Function IsSearchedText(ByRef fileText() As Byte, ByRef startByte As Integer, ByRef text() As Byte, _
                                          ByRef ignoreByte() As Boolean, ByRef checkZeroBytesPresence As Boolean) As Boolean
        If checkZeroBytesPresence Then
            If startByte + UBound(text) + 2 > UBound(fileText) Then Return False
        Else
            If startByte + UBound(text) > UBound(fileText) Then Return False
        End If
        For i As Integer = 0 To UBound(text) Step 1
            Dim c1 As String = Converter.ToStr(fileText(startByte + i))
            Dim c2 As String = Converter.ToStr(text(i))
            If IsNothing(ignoreByte) OrElse Not ignoreByte(i) Then
                If Not fileText(startByte + i) = text(i) Then Return False
            End If
        Next i
        If checkZeroBytesPresence Then
            For i As Integer = 1 To 2 Step 1
                If Not fileText(startByte + text.Length + i) = 0 Then Return False
            Next i
        End If
        Return True
    End Function
End Class

Class Translator

    Public BlockDelimiter() As Byte = Converter.ToByteArray(Writer.BlockDelimiterKeyword)
    Public OrigText() As Byte = Converter.ToByteArray(Writer.OrigTextKeyword)
    Public TransText() As Byte = Converter.ToByteArray(Writer.TransTextKeyword)

    Public Function ReadLangDictionary(ByRef path As String) As Dictionary(Of String, String)
        Dim r As New Dictionary(Of String, String)
        Dim content() As Byte = IO.File.ReadAllBytes(path)
        Dim startRead As Boolean = False
        Dim expectOriginal As Boolean = True
        Dim except As Boolean = False
        Dim i As Integer = 0
        Dim txtO() As Byte = Nothing
        Dim txtT() As Byte = Nothing
        Do While i <= UBound(content)
            If Parser.IsSearchedText(content, i, BlockDelimiter, Nothing, False) Then
                Call AddString(r, txtO, txtT)
                startRead = True
                i += BlockDelimiter.Length
                If expectOriginal Then
                    If Parser.IsSearchedText(content, i, OrigText, Nothing, False) Then
                        i += OrigText.Length
                    Else
                        except = True
                    End If
                Else
                    If Parser.IsSearchedText(content, i, TransText, Nothing, False) Then
                        i += TransText.Length
                    Else
                        except = True
                    End If
                End If
                If Not except Then
                    If Parser.IsSearchedText(content, i, BlockDelimiter, Nothing, False) Then
                        i += BlockDelimiter.Length
                    Else
                        except = True
                    End If
                End If
                expectOriginal = Not expectOriginal
            Else
                If startRead Then
                    If expectOriginal Then
                        Call AddChar(txtT, content(i))
                    Else
                        Call AddChar(txtO, content(i))
                    End If
                End If
                i += 1
            End If
            If i > UBound(content) Then Call AddString(r, txtO, txtT)
            If except Then
                MsgBox("Unexpected dictionary format. File " & vbNewLine & path & vbNewLine & "Byte " & i)
                End
            End If
        Loop
        Return r
    End Function
    Private Sub AddChar(ByRef dest() As Byte, ByRef c As Byte)
        If IsNothing(dest) Then
            ReDim dest(0)
        Else
            ReDim Preserve dest(dest.Length)
        End If
        dest(UBound(dest)) = c
    End Sub
    Private Sub AddString(ByRef dest As Dictionary(Of String, String), ByRef orig() As Byte, ByRef trans() As Byte)
        If Not IsNothing(orig) And Not IsNothing(trans) Then
            Dim t1 As String = PrepareString(orig, False)
            Dim t2 As String = PrepareString(trans, True)
            orig = Nothing : trans = Nothing
            If Not t2.Trim(" ", vbTab, Chr(10), Chr(13)) = "" Then dest.Add(t1, t2)
            If t2.ToLower.Contains(Writer.SuggestionTextKeyword.ToLower) Then
                MsgBox("Delete suggestions from the translation of text:" & vbNewLine & t1)
                End
            Else
                For Each t As String In {t1, t2}
                    Call TestString(t)
                Next t
            End If
        End If
    End Sub
    Public Shared Sub TestString(ByRef txt As String)
        If StartForm.IgnoreTextMistakes Then Exit Sub
        If txt.Trim.Contains("  ") Then
            Dim t As String = txt.Replace(Chr(13), Chr(10))
            Dim i0 As Integer = -1
            Do While Not i0 = t.Length
                i0 = t.Length
                t = t.Replace(Chr(10) & Chr(10), Chr(10))
            Loop
            Dim s() As String = t.Split(Chr(10))
            For Each line As String In s
                If line.Trim.Contains("  ") Then
                    MsgBox("Following text contains double spaces in the map file:" & vbNewLine & line)
                    End
                End If
            Next line
        End If
    End Sub
    Private Shared Function PrepareString(ByRef txt() As Byte, ByRef replaseYo As Boolean) As String
        Dim n As Integer = UBound(txt)
        Do While (txt(n) = 10 Or txt(n) = 13) And n > 0
            n -= 1
        Loop
        If n < UBound(txt) Then ReDim Preserve txt(n)
        Dim i As Integer = 0
        Do While i < UBound(txt) - 1
            If txt(i) = 13 And txt(i + 1) = 10 Then
                txt(i) = 10
                If UBound(txt) > i + 1 Then
                    For j As Integer = i + 2 To UBound(txt) Step 1
                        txt(j - 1) = txt(j)
                    Next j
                End If
                ReDim Preserve txt(UBound(txt) - 1)
            End If
            i += 1
        Loop
        Dim t As String = Converter.ToStr(txt)
        If replaseYo Then t = t.Replace("ё", "е")
        Return t
    End Function
    Private Shared Function PrepareString(ByRef txt As String, ByRef replaseYo As Boolean) As String
        Dim n As Integer = txt.Length - 1
        Do While (txt(n) = Chr(10) Or txt(n) = Chr(13)) And n > 0
            n -= 1
        Loop
        If n < txt.Length - 1 Then txt = txt.Substring(0, n + 1)
        Dim i As Integer = 0
        Do While i < txt.Length - 2
            If txt(i) = Chr(13) And txt(i + 1) = Chr(10) Then
                txt = txt.Substring(0, i) & Chr(10) & txt.Substring(i + 2)
            End If
            i += 1
        Loop
        If replaseYo Then txt = txt.Replace("ё", "е")
        Return txt
    End Function

    Public Shared Function DBFLangDictionary(ByRef textTable1 As String, ByRef textTable2 As String) As Dictionary(Of String, String)
        For Each f As String In {textTable1, textTable2}
            If Not IO.File.Exists(f) OrElse Not IO.Path.GetExtension(f).ToLower = ".dbf" Then
                Return Nothing
            End If
        Next
        Dim d() As Dictionary(Of String, String) = {NevendaarTools.GameDataModel.ReadTextTable(textTable1), _
                                                    NevendaarTools.GameDataModel.ReadTextTable(textTable2)}
        Dim dlower(UBound(d)) As Dictionary(Of String, String)
        Dim keys As List(Of String)
        For i As Integer = 0 To UBound(d) Step 1
            dlower(i) = New Dictionary(Of String, String)
            keys = d(i).Keys.ToList
            For Each k As String In keys
                If Not StartForm.IgnoreTextMistakes Then
                    If d(i).Item(k).Contains("  ") Then
                        MsgBox("Following text contains double spaces in the dbf file " &
                               {textTable1, textTable2}(i) & " :" & vbNewLine & d(i).Item(k))
                        End
                    End If
                End If
                dlower(i).Add(k.ToLower, d(i).Item(k))
            Next k
        Next i
        Dim result As New Dictionary(Of String, String)
        keys = dlower(0).Keys.ToList
        For Each k As String In keys
            If dlower(1).ContainsKey(k) Then
                For i As Integer = 0 To 1 Step 1
                    If Not result.ContainsKey(dlower(i).Item(k).ToLower) Then
                        result.Add(dlower(i).Item(k).ToLower, PrepareString(dlower(1 - i).Item(k), True))
                    End If
                Next i
            End If
        Next k
        Return result
    End Function

    Private Class TData
        Public langDict, DBFLangDict As Dictionary(Of String, String)
        Public fileText() As Byte
        Public output() As Byte
        Public outI As Integer
        Public csgHeader As Parser.CSGHeader = Nothing

        Public Sub AddByte(ByRef b As Byte, ByVal isTextReplacing As Boolean)
            If outI > UBound(output) Then ReDim Preserve output(outI)
            output(outI) = b
            If Not IsNothing(csgHeader) And Not isTextReplacing Then
                Call csgHeader.FileLengthChanged(outI, 1)
            End If
            outI += 1
        End Sub
        Public Sub OriginalTextRemoved(ByVal startByte As Integer, ByVal bytesAmount As Integer)
            If Not IsNothing(csgHeader) And Not bytesAmount = 0 Then
                Call csgHeader.FileLengthChanged(startByte, bytesAmount)
            End If
        End Sub

        Public Function Print() As String
            ReDim Preserve output(outI - 1)
            Return Converter.ToStr(output)
        End Function
    End Class

    Public Function Translate(ByRef langDict As Dictionary(Of String, String), ByRef DBFLangDict As Dictionary(Of String, String), _
                              ByRef fileText() As Byte, _
                              ByRef csgHeader As Parser.CSGHeader, ByRef fileType As Reader.FileType, _
                              Optional ByRef csgHeader_forTest As Parser.CSGHeader = Nothing) As Byte()
        Dim d As New TData With {.fileText = fileText, .langDict = langDict, .DBFLangDict = DBFLangDict, _
                                 .csgHeader = csgHeader, .outI = 0}
        ReDim d.output(UBound(fileText))
        Call CopyHeader(d, fileType)
        Dim p As New Parser
        Dim t As Parser.CheckResult
        Dim i As Integer = Parser.sgHostLordEnd + 1
        If fileType = Reader.FileType.Campaign Then i += Parser.csgSGBias

        d.outI = i
        Dim i0 As Integer
        Do While i <= UBound(fileText)
            i0 = i
            t = p.GetText(fileText, i)
            Call Parser.CSGHeader.Compare(d.csgHeader, csgHeader_forTest)
            If t.text = "" Then
                Call CopyRange(d, i0, i - 1, True)
            Else
                Call TranslateWithDynLength(d, t, i0)
            End If
            Call Parser.CSGHeader.Compare(d.csgHeader, csgHeader_forTest)
        Loop
        Return d.output
    End Function

    Private Sub CopyHeader(ByRef d As TData, ByRef fileType As Reader.FileType)
        If fileType = Reader.FileType.Map Then
            Call CopyRange(d, 0, Parser.sgInitBlock, True)
        ElseIf fileType = Reader.FileType.Campaign Then
            Call CopyRange(d, 0, Parser.csgInitBlock, True)
            Dim rC1 As Parser.CheckResult = Parser.GetCampaignName(d.fileText)
            Call TranslateWithFixedLength(d, rC1)
            Dim rC2 As Parser.CheckResult = Parser.GetCampaignAuthor(d.fileText)
            Call TranslateWithFixedLength(d, rC2)
            Call CopyBytesBetweenBlocks(d, rC1, rC2)
            Call CopyBytesBetweenBlocks(d, rC2.textEndByte + 1, Parser.csgSGBias + Parser.sgInitBlock)
        Else
            Throw New Exception("Unexpected type")
        End If
        Dim r1 As Parser.CheckResult = Parser.GetMapDescription(d.fileText, fileType)
        Call TranslateWithFixedLength(d, r1)
        Dim r2 As Parser.CheckResult = Parser.GetMapAuthor(d.fileText, fileType)
        Call TranslateWithFixedLength(d, r2)
        Dim r3 As Parser.CheckResult = Parser.GetMapName(d.fileText, fileType)
        Call TranslateWithFixedLength(d, r3)
        Dim r4 As Parser.CheckResult = Parser.GetMapHostLordName(d.fileText, fileType)
        Call TranslateWithFixedLength(d, r4)

        Call CopyBytesBetweenBlocks(d, r1, r2)
        Call CopyBytesBetweenBlocks(d, r2, r3)
        Call CopyBytesBetweenBlocks(d, r3, r4)
    End Sub
    Private Sub CopyBytesBetweenBlocks(ByRef d As TData, ByRef b1 As Parser.CheckResult, ByRef b2 As Parser.CheckResult)
        Dim i1 As Integer = b1.textEndByte + 1
        Dim i2 As Integer = b2.textStartByte - 1
        Call CopyBytesBetweenBlocks(d, i1, i2)
    End Sub
    Private Sub CopyBytesBetweenBlocks(ByRef d As TData, ByRef i1 As Integer, ByRef i2 As Integer)
        If i1 > i2 Then Exit Sub
        For i As Integer = i1 To i2 Step 1
            d.output(i) = d.fileText(i)
        Next i
    End Sub
    Private Sub TranslateWithFixedLength(ByRef d As TData, ByRef t As Parser.CheckResult)
        Dim trText() As Byte = GetTranslation(d, t)
        Dim maxLen As Integer = t.textEndByte - t.textStartByte + 1
        If trText.Length > maxLen Then
            MsgBox("Text has length of " & trText.Length & " whereas max. is " & maxLen & "." _
                   & vbNewLine & Converter.ToStr(trText))
            End
        End If
        For i As Integer = 0 To UBound(trText) Step 1
            d.output(t.textStartByte + i) = trText(i)
        Next i
    End Sub
    Private Sub TranslateWithDynLength(ByRef d As TData, ByRef t As Parser.CheckResult, ByRef checkStarI As Integer)
        Dim trText() As Byte
        trText = GetTranslation(d, t)
        Dim initialOutI As Integer = d.outI
        Dim originalBlockLen As Integer = t.textEndByte - checkStarI + 1
        If trText.Length > t.maxTextLength * t.byteBlock.Length Then
            MsgBox("Text has length of " & trText.Length & " whereas max. is " & t.maxTextLength * t.byteBlock.Length & "." _
                   & vbNewLine & Converter.ToStr(trText))
            End
        End If
        If Not t.isLongBlock Then
            Call CopyRange(d, checkStarI, t.sizeByte - 1, True)
            Call d.AddByte(UBound(trText) + 2, True)
            d.outI += 3
            Call AddRange(d, trText, True)
            d.outI += 1
        Else
            Dim j1 As Integer = 0
            Dim j2 As Integer = -1
            For i As Integer = 0 To UBound(t.byteBlock) Step 1
                Call AddRange(d, t.byteBlock(i), True)
                j1 = j2 + 1
                j2 = Math.Min(j1 + t.maxTextLength - 1, UBound(trText))
                If j1 < j2 Then
                    Call d.AddByte(j2 - j1 + 3, True)
                    d.outI += 3
                    Call AddRange(d, trText, j1, j2, True)
                    Call AddRange(d, Converter.ToByteArray("_"), True)
                    d.outI += 1
                Else
                    Call d.AddByte(1, True)
                    d.outI += 4
                End If
            Next i
        End If
        Dim finalOutI As Integer = d.outI
        Dim lenghtChange As Integer = (finalOutI - initialOutI) - originalBlockLen
        Call d.OriginalTextRemoved(initialOutI, lenghtChange)
    End Sub
    Private Sub CopyRange(ByRef d As TData, ByRef i1 As Integer, ByRef i2 As Integer, _
                          ByVal isTextReplacing As Boolean)
        Call AddRange(d, d.fileText, i1, i2, isTextReplacing)
    End Sub
    Private Sub AddRange(ByRef d As TData, ByRef range() As Byte, _
                         ByVal isTextReplacing As Boolean)
        Call AddRange(d, range, 0, UBound(range), isTextReplacing)
    End Sub
    Private Sub AddRange(ByRef d As TData, ByRef range() As Byte, ByRef i1 As Integer, ByRef i2 As Integer, _
                         ByVal isTextReplacing As Boolean)
        If d.outI + i2 - i1 > UBound(d.output) Then ReDim Preserve d.output(d.outI + i2 - i1)
        For i As Integer = i1 To i2 Step 1
            Call d.AddByte(range(i), isTextReplacing)
        Next i
    End Sub

    Private Function GetTranslation(ByRef d As TData, ByRef t As Parser.CheckResult) As Byte()

        Dim p As String = PrepareString(t.text, False)
        If d.langDict.ContainsKey(p) Then
            Return Converter.ToByteArray(d.langDict.Item(p))
        ElseIf Not IsNothing(d.DBFLangDict) AndAlso d.DBFLangDict.ContainsKey(p.ToLower) Then
            Return Converter.ToByteArray(d.DBFLangDict.Item(p.ToLower))
        Else
            MsgBox("Could not find translation for:" & vbNewLine & t.text)
            End
        End If
    End Function
End Class

Class Converter

    Public Shared Function ToByteArray(ByRef txt As String) As Byte()
        Return System.Text.Encoding.GetEncoding(Reader.encID).GetBytes(txt)
    End Function
    Public Shared Function ToStr(ByRef b As Byte) As String
        Return System.Text.Encoding.GetEncoding(Reader.encID).GetString({b})
    End Function
    Public Shared Function ToStr(ByRef b() As Byte) As String
        Return System.Text.Encoding.GetEncoding(Reader.encID).GetString(b)
    End Function

End Class