
Public Class StartForm

    Private Function LangDictionaryPath() As String
        Return PathTextBox.Text & ".dictionary.txt"
    End Function
    Private Function MapPath() As String
        Return PathTextBox.Text
    End Function
    Private Function MapTranslatedPath() As String
        Return PathTextBox.Text & ".translated.sg"
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

    Sub parse() Handles ParseButton.Click
        Call WriteLastPathFile()
        If Not Writer.CheckCanWrite(LangDictionaryPath) Then Exit Sub
        Dim p As New Parser
        Dim f As List(Of String) = p.Parse(Reader.ReadFile(MapPath))
        Dim L As Dictionary(Of String, String) = Nothing
        Dim DBFL As Dictionary(Of String, String) = Nothing
        If Autotranslate() Then DBFL = Translator.DBFLangDictionary(TGlobal1Path, Tglobal2Path)
        If IO.File.Exists(LangDictionaryPath) Then
            Dim t As New Translator
            L = t.ReadLangDictionary(LangDictionaryPath)
        End If
        Call Writer.Write(LangDictionaryPath, f, L, DBFL)
        MsgBox("done")
    End Sub
    Sub make() Handles MakeButton.Click
        Call WriteLastPathFile()
        If Not Writer.CheckCanWrite(MapTranslatedPath) Then Exit Sub
        Dim t As New Translator
        Dim L As Dictionary(Of String, String) = t.ReadLangDictionary(LangDictionaryPath)
        Dim DBFL As Dictionary(Of String, String) = Nothing
        If Autotranslate() Then DBFL = Translator.DBFLangDictionary(TGlobal1Path, Tglobal2Path)
        Dim f() As Byte = Reader.ReadFile(MapPath)
        Dim translated() As Byte = t.Translate(L, DBFL, f, False)
        IO.File.WriteAllBytes(MapTranslatedPath, translated)
        MsgBox("done")
    End Sub
    Sub test() Handles TestButton.Click
        Dim t As New Translator
        Dim f() As Byte = Reader.ReadFile(MapPath)
        Dim translated() As Byte = t.Translate(Nothing, Nothing, f, True)
        For i As Integer = 0 To UBound(f) Step 1
            If Not f(i) = translated(i) Then
                Throw New Exception
            End If
        Next i
        If Not f.Length = translated.Length Then
            Throw New Exception
        End If
        MsgBox("done")
    End Sub

    Sub meload() Handles Me.Load
        TestButton.Visible = False
        PathTextBox.Text = Reader.GetMapPath
        TglobalTextBox1.Text = Reader.GetTglobal1Path
        TglobalTextBox2.Text = Reader.GetTglobal2Path
        AutotranslateCheckBox.Checked = Reader.GetAutotranslateState
    End Sub
    Private Sub WriteLastPathFile()
        Call Writer.WriteLastPathFile(PathTextBox.Text, TglobalTextBox1.Text, TglobalTextBox2.Text, AutotranslateCheckBox.Checked)
    End Sub

    Private Sub SelectFile(obj As Object, e As EventArgs) Handles MapButton.Click, Tglobal1Button.Click, Tglobal2Button.Click
        Dim buttons() As Button = {MapButton, Tglobal1Button, Tglobal2Button}
        Dim boxes() As TextBox = {PathTextBox, TglobalTextBox1, TglobalTextBox2}
        Dim filters() As String = {"Map files (*.sg)|*.sg", "TGlobal table (*.dbf)|*.dbf", "TGlobal table (*.dbf)|*.dbf"}
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
                            "После изменения нескольких блоков лучше пересобирать карту и" & vbNewLine &
                            "проверять, запускается ли она в редакторе карт"
        MsgBox(msg)
    End Sub
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
End Class

Class Writer

    Public Const BlockDelimiterKeyword As String = "--------- --------- --------- --------- --------- ---------" & vbNewLine
    Public Const OrigTextKeyword As String = "# Original text" & vbNewLine
    Public Const TransTextKeyword As String = "# Translated text" & vbNewLine

    Public Shared Function CheckCanWrite(ByRef path As String) As Boolean
        If IO.File.Exists(path) Then
            Dim answer As MsgBoxResult = MsgBox("Do you want to replace file " & path & "?", MsgBoxStyle.YesNo)
            If answer = MsgBoxResult.No Then Return False
        End If
        Return True
    End Function

    Public Shared Sub Write(ByRef path As String, ByRef content As List(Of String), _
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
                    out(n) = PrintLine(s, "")
                End If
                printed.Add(s)
            End If
        Next s
        If UBound(out) > n Then ReDim Preserve out(n)
        IO.File.WriteAllLines(path, out, System.Text.Encoding.GetEncoding(Reader.encID))
    End Sub
    Private Shared Function PrintLine(ByRef original As String, ByRef translation As String) As String
        Return BlockDelimiterKeyword & _
               OrigTextKeyword & _
               BlockDelimiterKeyword & _
               original & vbNewLine & _
               BlockDelimiterKeyword & _
               TransTextKeyword & _
               BlockDelimiterKeyword & _
               translation
    End Function

    Public Shared Sub WriteLastPathFile(ByRef mapFile As String, ByRef tglobal1 As String, ByRef tglobal2 As String, _
                                        ByRef AutotranslateCheckBox As Boolean)
        Dim lines() As String = {mapFile, tglobal1, tglobal2, AutotranslateCheckBox.ToString}
        IO.File.WriteAllLines(My.Resources.lastPathFile, lines)
    End Sub

End Class

Class Parser

    Dim dataBlocks() As Block

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

    Public Const initBlock As Integer = 42
    Public Const descriptionBlock As Integer = initBlock + 256
    Public Const AuthorBlock As Integer = descriptionBlock + 22
    Public Const NameBlock As Integer = AuthorBlock + 64
    Public Const HostLordStart As Integer = 609
    Public Const HostLordEnd As Integer = HostLordStart + 14
    Private blockID As Integer = -1
    Private collectedWords As New List(Of String)

    Public Function Parse(ByRef fileText() As Byte) As List(Of String)
        Dim t As CheckResult
        Dim i As Integer = 0
        Dim r As New List(Of String)
        t = GetMapName(fileText)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetMapAuthor(fileText)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetMapDescription(fileText)
        If Not r.Contains(t.text) Then r.Add(t.text)
        t = GetHostLordName(fileText)
        If Not r.Contains(t.text) Then r.Add(t.text)
        i = HostLordEnd + 1
        Do While i <= UBound(fileText)
            t = GetText(fileText, i)
            If Not t.text = "" AndAlso Not r.Contains(t.text) Then
                r.Add(t.text)
            End If
        Loop
        Return r
    End Function

    Public Shared Function GetMapDescription(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = initBlock + 1, .textEndByte = descriptionBlock - 2}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetMapAuthor(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = descriptionBlock + 1, .textEndByte = AuthorBlock - 2}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetMapName(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = AuthorBlock + 1, .textEndByte = NameBlock}
        r.text = Converter.ToStr(ReadFromTo(fileText, r.textStartByte, r.textEndByte, True))
        Return r
    End Function
    Public Shared Function GetHostLordName(ByRef fileText() As Byte) As CheckResult
        Dim r As New CheckResult With {.textStartByte = HostLordStart, .textEndByte = HostLordEnd}
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
        Dim d() As Dictionary(Of String, String) = {NevendaarTools.GameDataModel.ReadTextTable(textTable1), _
                                                    NevendaarTools.GameDataModel.ReadTextTable(textTable2)}
        Dim dlower(UBound(d)) As Dictionary(Of String, String)
        Dim keys As List(Of String)
        For i As Integer = 0 To UBound(d) Step 1
            dlower(i) = New Dictionary(Of String, String)
            keys = d(i).Keys.ToList
            For Each k As String In keys
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

        Public Sub AddByte(ByRef b As Byte)
            If outI > UBound(output) Then ReDim Preserve output(outI)
            output(outI) = b
            outI += 1
        End Sub
        Public Function Print() As String
            ReDim Preserve output(outI - 1)
            Return Converter.ToStr(output)
        End Function
    End Class

    Public Function Translate(ByRef langDict As Dictionary(Of String, String), ByRef DBFLangDict As Dictionary(Of String, String), _
                              ByRef fileText() As Byte, ByRef test As Boolean) As Byte()
        Dim d As New TData With {.fileText = fileText, .langDict = langDict, .DBFLangDict = DBFLangDict, .outI = 0}
        ReDim d.output(UBound(fileText))
        If Not test Then
            Call CopyHeader(d)
        Else
            Call CopyRange(d, 0, Parser.HostLordEnd)
        End If
        Dim p As New Parser
        Dim t As Parser.CheckResult
        Dim i As Integer = Parser.HostLordEnd + 1
        d.outI = i
        Dim i0 As Integer
        Do While i <= UBound(fileText)
            i0 = i
            t = p.GetText(fileText, i)
            If t.text = "" Then
                Call CopyRange(d, i0, i - 1)
            Else
                Call TranslateWithDynLength(d, t, i0, test)
            End If
        Loop
        Return d.output
    End Function

    Private Sub CopyHeader(ByRef d As TData)
        Call CopyRange(d, 0, Parser.initBlock)
        Dim r1 As Parser.CheckResult = Parser.GetMapDescription(d.fileText)
        Call TranslateWithFixedLength(d, r1)
        Dim r2 As Parser.CheckResult = Parser.GetMapAuthor(d.fileText)
        Call TranslateWithFixedLength(d, r2)
        Dim r3 As Parser.CheckResult = Parser.GetMapName(d.fileText)
        Call TranslateWithFixedLength(d, r3)
        Dim r4 As Parser.CheckResult = Parser.GetHostLordName(d.fileText)
        Call TranslateWithFixedLength(d, r4)
        For i As Integer = r3.textEndByte + 1 To r4.textStartByte - 1 Step 1
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
    Private Sub TranslateWithDynLength(ByRef d As TData, ByRef t As Parser.CheckResult, ByRef checkStarI As Integer, _
                                       ByRef test As Boolean)
        Dim trText() As Byte
        If Not test Then
            trText = GetTranslation(d, t)
        Else
            trText = Converter.ToByteArray(t.text)
        End If
        If trText.Length > t.maxTextLength * t.byteBlock.Length Then
            MsgBox("Text has length of " & trText.Length & " whereas max. is " & t.maxTextLength * t.byteBlock.Length & "." _
                   & vbNewLine & Converter.ToStr(trText))
            End
        End If
        If Not t.isLongBlock Then
            Call CopyRange(d, checkStarI, t.sizeByte - 1)
            Call d.AddByte(UBound(trText) + 2)
            d.outI += 3
            Call AddRange(d, trText)
            d.outI += 1
        Else
            Dim j1 As Integer = 0
            Dim j2 As Integer = -1
            For i As Integer = 0 To UBound(t.byteBlock) Step 1
                Call AddRange(d, t.byteBlock(i))
                j1 = j2 + 1
                j2 = Math.Min(j1 + t.maxTextLength - 1, UBound(trText))
                If j1 < j2 Then
                    Call d.AddByte(j2 - j1 + 3)
                    d.outI += 3
                    Call AddRange(d, trText, j1, j2)
                    Call AddRange(d, Converter.ToByteArray("_"))
                    d.outI += 1
                Else
                    Call d.AddByte(1)
                    d.outI += 4
                End If
            Next i
        End If
    End Sub
    Private Sub CopyRange(ByRef d As TData, ByRef i1 As Integer, ByRef i2 As Integer)
        Call AddRange(d, d.fileText, i1, i2)
    End Sub
    Private Sub AddRange(ByRef d As TData, ByRef range() As Byte)
        Call AddRange(d, range, 0, UBound(range))
    End Sub
    Private Sub AddRange(ByRef d As TData, ByRef range() As Byte, ByRef i1 As Integer, ByRef i2 As Integer)
        If d.outI + i2 - i1 > UBound(d.output) Then ReDim Preserve d.output(d.outI + i2 - i1)
        For i As Integer = i1 To i2 Step 1
            Call d.AddByte(range(i))
        Next i
    End Sub

    Private Function GetTranslation(ByRef d As TData, ByRef t As Parser.CheckResult) As Byte()
        If d.langDict.ContainsKey(t.text) Then
            Return Converter.ToByteArray(d.langDict.Item(t.text))
        ElseIf Not IsNothing(d.DBFLangDict) AndAlso d.DBFLangDict.ContainsKey(t.text.ToLower) Then
            Return Converter.ToByteArray(d.DBFLangDict.Item(t.text.ToLower))
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