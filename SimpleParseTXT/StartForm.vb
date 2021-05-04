Public Class StartForm


    Dim mapInfoKeys() As String

    Public chartoint As New Dictionary(Of String, Integer)
    Public inttochar As New Dictionary(Of Integer, String)

    Dim nStartMsg As Integer = 3
    Dim nEndMsg As Integer = 1

    Dim maxMsgLen As Integer = 254

    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(1251)

    Public parsed() As Block

    Structure Block
        Dim isMsg As Boolean
        Dim subtext1, subtext2 As String
        Dim Text As String
        Dim msgLen As Integer
        Dim setAsMe() As Integer
        Dim setMeAs As Integer
        Dim isSpecial As Boolean
        Dim isInfo As Boolean
    End Structure

    Sub parse() Handles ParseButton.Click
        Call read_sg()
        Call saveParsed()
        MsgBox("done")
    End Sub
    Sub make() Handles MakeButton.Click
        Call read_sg()
        Call make_sg()
        MsgBox("done")
    End Sub
    Sub test() Handles TestButton.Click
        Dim t0 As Integer
        t0 = Environment.TickCount
        Call read_sg()
        Console.WriteLine(Environment.TickCount - t0)
        t0 = Environment.TickCount
        Call saveParsed()
        Console.WriteLine(Environment.TickCount - t0)
        t0 = Environment.TickCount

        IO.File.Delete(PathTextBox.Text & ".translated_messages.txt")
        IO.File.Move(PathTextBox.Text & ".messages.txt", PathTextBox.Text & ".translated_messages.txt")
        Call make_sg()
        Console.WriteLine(Environment.TickCount - t0)
        t0 = Environment.TickCount

        Dim bytes1() As Byte = IO.File.ReadAllBytes(PathTextBox.Text)
        Dim bytes2() As Byte = IO.File.ReadAllBytes(PathTextBox.Text & ".translated.sg")
        If bytes1.Length <> bytes2.Length Then Throw New Exception("invalid length")
        For i As Integer = 0 To UBound(bytes1) Step 1
            If bytes1(i) <> bytes2(i) Then
                Console.WriteLine(bytes1(i - 1) & " " & bytes1(i) & " " & bytes1(i + 1))
                Console.WriteLine(bytes2(i - 1) & " " & bytes2(i) & " " & bytes2(i + 1))
                Console.WriteLine(inttochar(bytes1(i - 1)) & " " & inttochar(bytes1(i)) & " " & inttochar(bytes1(i + 1)))
                Console.WriteLine("")
                Console.WriteLine(inttochar(bytes2(i - 1)) & " " & inttochar(bytes2(i)) & " " & inttochar(bytes2(i + 1)))
                Throw New Exception("invalid byte")
            End If
        Next i
        MsgBox("done")
    End Sub

    Sub meload() Handles Me.Load
        TestButton.Enabled = False
        If IO.File.Exists("./lastpath.txt") Then
            PathTextBox.Text = IO.File.ReadAllText("./lastpath.txt")
        End If
        For i As Integer = 0 To 255 Step 1
            Dim s As String = enc.GetString({CByte(i)})
            chartoint.Add(s, i)
            inttochar.Add(i, s)
        Next i
    End Sub

    Sub read_sg()

        parsed = Nothing

        IO.File.WriteAllText("./lastpath.txt", PathTextBox.Text)

        If Not IO.File.Exists(PathTextBox.Text) Then
            MsgBox("Не могу найти файл " & PathTextBox.Text)
            End
        End If

        Dim bytes() As Byte = IO.File.ReadAllBytes(PathTextBox.Text)

        Dim txt As String = ""

        Dim u As Integer = bytes.Length - 1
        Dim subtxt1, subtxt2 As String
        Dim currentBlock As String = ""
        Dim subblock As String = ""
        Dim parseTxtBlock, parseMapInfoBlock, appendBlock As Boolean

        Dim maxlen As Integer = 2000

        For i = 0 To u Step 1
            subblock &= inttochar(bytes(i))
            If subblock.Length > maxlen OrElse i = u Then
                txt &= subblock
                subblock = ""
            End If
        Next i

        Dim added As New Dictionary(Of String, Integer)

        Dim initBlock As Integer = 42
        Dim descriptionBlock As Integer = initBlock + 256
        Dim AuthorBlock As Integer = descriptionBlock + 22
        Dim NameBlock As Integer = AuthorBlock + 256

        For i As Integer = 0 To u Step 1
            If i < NameBlock Then
                parseTxtBlock = False
            ElseIf i + My.Resources.MsgTextStart.Length < u Then
                parseTxtBlock = IsSearchingString(txt, i, My.Resources.MsgTextStart)
            Else
                parseTxtBlock = False
            End If
            If Not parseTxtBlock And i + My.Resources.scenDescStart.Length < u Then
                parseMapInfoBlock = IsSearchingString(txt, i, My.Resources.scenDescStart)
            End If
            If parseTxtBlock Then
                Call SaveNonMsgBlock(currentBlock, subblock)
                i += My.Resources.MsgTextStart.Length
                Dim n As Integer = chartoint.Item(txt(i)) - 1
                i += 1
                subtxt1 = txt.Substring(i, nStartMsg)
                i += nStartMsg
                For j As Integer = i To u Step 1
                    appendBlock = IsSearchingString(txt, j + nEndMsg, My.Resources.MsgTextEnd)
                    If appendBlock Then
                        subtxt2 = txt.Substring(j, nEndMsg)
                        Call addblock(subblock, True, subtxt1, subtxt2, False, False)
                        If Not n = subblock.Length Then Throw New Exception("Invalid msg length")
                        'Console.WriteLine(n & vbTab & subblock.Length & vbTab & parsed(parsed.Length - 1).subtext1 & vbTab & parsed(parsed.Length - 1).subtext2)
                        subblock = ""
                        i += parsed(parsed.Length - 1).msgLen + nEndMsg + My.Resources.MsgTextEnd.Length - 1

                        If added.ContainsKey(parsed(parsed.Length - 1).Text) Then
                            Dim id As Integer = added.Item(parsed(parsed.Length - 1).Text)
                            If IsNothing(parsed(id).setAsMe) Then
                                ReDim parsed(id).setAsMe(0)
                            Else
                                ReDim Preserve parsed(id).setAsMe(parsed(id).setAsMe.Length)
                            End If
                            parsed(id).setAsMe(UBound(parsed(id).setAsMe)) = parsed.Length - 1
                            parsed(parsed.Length - 1).Text = ""
                            parsed(parsed.Length - 1).msgLen = -1
                            parsed(parsed.Length - 1).setMeAs = id
                        Else
                            added.Add(parsed(parsed.Length - 1).Text, parsed.Length - 1)
                        End If
                        Exit For
                    Else
                        subblock &= txt(j)
                    End If
                Next j
            ElseIf parseMapInfoBlock Then
                Call SaveNonMsgBlock(currentBlock, subblock)
                i += My.Resources.scenDescStart.Length
                subtxt1 = ""
                For j As Integer = i To u Step 1
                    appendBlock = IsSearchingString(txt, j + nEndMsg, My.Resources.scenDescEnd)
                    If appendBlock Then
                        subtxt2 = ""
                        Call ParseMapInfo(subblock)
                        i = j
                        Exit For
                    Else
                        subblock &= txt(j)
                    End If
                Next j
                currentBlock = ""
                subblock = ""
                parseMapInfoBlock = False
            Else
                subblock &= txt(i)
                If subblock.Length > maxlen Or i = u Then
                    currentBlock &= subblock
                    subblock = ""
                End If
                If i = u Or i = initBlock Or i = descriptionBlock Or i = AuthorBlock Or i = NameBlock Then
                    i = i
                    currentBlock &= subblock
                    subblock = ""
                    If i = u Or i = initBlock Then
                        Call addblock(currentBlock, False, Nothing, Nothing, False, False)
                    Else
                        Call addblock(currentBlock, False, Nothing, Nothing, True, False)
                    End If
                    currentBlock = ""
                End If
            End If
        Next i
    End Sub
    Private Sub SaveNonMsgBlock(ByRef currentBlock As String, ByRef subBlock As String)
        currentBlock &= subblock
        subblock = ""
        If IsNothing(parsed) Then
            ReDim parsed(0)
        Else
            ReDim Preserve parsed(parsed.Length)
        End If
        parsed(parsed.Length - 1).isMsg = False
        parsed(parsed.Length - 1).isSpecial = False
        parsed(parsed.Length - 1).isInfo = False
        parsed(parsed.Length - 1).msgLen = currentBlock.Length
        parsed(parsed.Length - 1).Text = currentBlock
        currentBlock = ""
    End Sub
    Private Sub ParseMapInfo(ByRef str As String)
        ReDim mapInfoKeys(8)
        Dim i As Integer

        mapInfoKeys(0) = ""
        Dim part1 As String = str.Substring(0, 102)
        i = part1.Length
        mapInfoKeys(1) = My.Resources.minfoName
        Dim Name As String = ReadField(str, i, My.Resources.minfoName, False)
        mapInfoKeys(2) = My.Resources.minfoDesc
        Dim Description As String = ReadField(str, i, My.Resources.minfoDesc, False)
        mapInfoKeys(3) = My.Resources.minfoBriefing
        Dim Goal As String = ReadField(str, i, My.Resources.minfoBriefing, False)
        mapInfoKeys(4) = My.Resources.minfoWin
        Dim winText As String = ReadField(str, i, My.Resources.minfoWin, True)
        For k As Integer = 2 To 5 Step 1
            winText &= ReadField(str, i, My.Resources.minfoWin & k, True)
        Next k
        mapInfoKeys(5) = My.Resources.minfoLose
        Dim loseText As String = ReadField(str, i, My.Resources.minfoLose, False)
        mapInfoKeys(6) = My.Resources.minfoBriefingLong
        Dim Briefing As String = ""
        For k As Integer = 1 To 5 Step 1
            Briefing &= ReadField(str, i, My.Resources.minfoBriefingLong & k, True)
        Next k
        mapInfoKeys(7) = ""
        Dim part2 As String = str.Substring(i, 101)
        i += part2.Length
        mapInfoKeys(8) = My.Resources.minfoAuthor
        Dim Author As String = ReadField(str, i, My.Resources.minfoAuthor, False)

        Dim content() As String = New String() {part1, Name, Description, Goal, winText, loseText, Briefing, part2, Author}
        For k As Integer = 0 To UBound(content) Step 1
            Call addblock(content(k), False, "", "", True, True)
        Next k
    End Sub
    Private Function ReadField(ByRef txt As String, ByRef i As Integer, ByRef field As String, ByRef longText As Boolean) As String
        Dim f As String = txt.Substring(i, field.Length)
        If Not f = field Then Throw New Exception("неожиданное поле описания карты")
        i += f.Length
        Dim lenChar As String = txt.Substring(i, 1)
        Dim n As Integer = chartoint(lenChar) - 1
        i += 4
        Dim res As String = txt.Substring(i, n)
        i += res.Length + 1
        If longText And res.Length > 0 Then
            Return res.Substring(0, res.Length - 1)
        Else
            Return res
        End If
    End Function

    Sub addblock(ByRef text As String, ByRef ismsg As Boolean, ByRef subtxt1 As String, ByRef subtxt2 As String, _
                 ByRef isSpecial As Boolean, ByRef isInfo As Boolean)
        If IsNothing(parsed) Then
            ReDim parsed(0)
        Else
            ReDim Preserve parsed(parsed.Length)
        End If
        Dim u As Integer = parsed.Length - 1
        parsed(u).isMsg = ismsg
        parsed(u).msgLen = text.Length
        parsed(u).Text = text
        parsed(u).isSpecial = isSpecial
        parsed(u).isInfo = isInfo
        If ismsg Then
            For k = 0 To nStartMsg - 1 Step 1
                If k > 0 Then parsed(u).subtext1 &= " "
                parsed(parsed.Length - 1).subtext1 &= chartoint.Item(subtxt1(k))
            Next k
            For k = 0 To nEndMsg - 1 Step 1
                If k > 0 Then parsed(u).subtext2 &= " "
                parsed(u).subtext2 &= chartoint.Item(subtxt2(k))
            Next k
        ElseIf isSpecial And Not isInfo Then
            For i As Integer = text.Length - 1 To 0 Step -1
                If Not text.Substring(i, 1) = Chr(0) Or i = 0 Then
                    parsed(u).Text = parsed(u).Text.Substring(0, i + 1)
                    Exit For
                End If
            Next i
        End If
    End Sub
    Function IsSearchingString(ByRef text As String, ByRef i As Integer, ByRef word As String) As Boolean
        If Not text(i) = word.Substring(0, 1) Then
            Return False
        Else
            Dim s As String = text.Substring(i, word.Length)
            Return (s = word)
        End If
    End Function

    Sub saveParsed()
        Dim u As Integer = -1
        For i As Integer = 0 To UBound(parsed) Step 1
            Dim b As Block = parsed(i)
            If (b.isMsg Or b.isSpecial) And b.msgLen > -1 Then u += 1
        Next i
        Dim txt(u) As String
        u = -1
        For i As Integer = 0 To UBound(parsed) Step 1
            Dim b As Block = parsed(i)
            If (b.isMsg Or b.isSpecial) And b.msgLen > -1 Then
                u += 1
                Dim setAs As String = ""
                If Not IsNothing(b.setAsMe) Then
                    For k As Integer = 0 To UBound(b.setAsMe) Step 1
                        setAs &= " " & b.setAsMe(k)
                    Next k
                End If
                Dim asterix As String
                If b.isSpecial Then
                    asterix = "*"
                Else
                    asterix = ""
                End If
                txt(u) = "#block " & i & vbNewLine &
                         b.msgLen & asterix & " " & b.subtext1 & " " & b.subtext2 & setAs & vbNewLine &
                         b.Text
            End If
        Next i
        IO.File.WriteAllLines(PathTextBox.Text & ".messages.txt", txt, enc)
    End Sub

    Sub make_sg()

        If Not IO.File.Exists(PathTextBox.Text & ".translated_messages.txt") Then
            MsgBox("Не могу найти файл " & PathTextBox.Text & ".translated_messages.txt")
            End
        End If

        Dim trans() As String = IO.File.ReadAllLines(PathTextBox.Text & ".translated_messages.txt", enc)

        Dim trB(UBound(parsed)) As String

        Dim infotick As Integer = 0

        Dim tline As Integer = 0
        Dim txt As String = ""
        Dim err As String = ""
        For i As Integer = 0 To UBound(parsed) Step 1
            Dim b As Block = parsed(i)
            If b.isMsg Or b.isSpecial Then
                If b.msgLen > -1 Then
                    If trans(tline) = "#block " & i Then
                        tline += 1
                        Dim s() As String = trans(tline).Split(" ")

                        If Not b.isSpecial = (s(0).Substring(s(0).Length - 1) = "*") Then
                            Throw New Exception("Неожиданный тип блока " & i)
                        End If
                        If b.isSpecial Then s(0) = s(0).Substring(0, s(0).Length - 1)

                        Dim len As Integer = s(0)
                        Dim subtext1 As String = ""
                        Dim subtext2 As String = ""
                        If Not b.isSpecial Then
                            For k As Integer = 1 To nStartMsg Step 1
                                subtext1 &= inttochar(s(k))
                            Next k
                            For k As Integer = nStartMsg + 1 To nStartMsg + nEndMsg Step 1
                                subtext2 &= inttochar(s(k))
                            Next k
                        End If
                        tline += 1
                        Dim msg As String = ""
                        Dim addnewline As Boolean = False
                        While Not trans(tline).Contains("#block")
                            If addnewline Then
                                msg &= Chr(10)
                            Else
                                addnewline = True
                            End If
                            If Not b.isInfo Or (b.isInfo AndAlso Not mapInfoKeys(infotick) = "") Then
                                msg &= trans(tline).Replace("_", " ").Replace("ё", "е")
                            Else
                                msg &= trans(tline)
                            End If
                            If tline = UBound(trans) Then
                                Exit While
                            Else
                                tline += 1
                            End If
                        End While
                        Dim L As Integer = msg.Length
                        If Not TextStringLenTest(L, b, infotick) Then
                            err &= "Количество символов в сообщении блока " & i & " превышает " & maxMsgLen & " (" & L & ")" & vbNewLine
                        Else
                            trB(i) = ""
                            If Not b.isSpecial Then
                                trB(i) &= My.Resources.MsgTextStart & inttochar(L + 1) & subtext1
                            ElseIf b.isInfo Then
                                If infotick = 4 Or infotick = 6 Then
                                    Dim tmpT As String = msg
                                    Dim shortmsg As String
                                    ReDim Preserve trB(trB.Length + 4)
                                    For k As Integer = 1 To 5 Step 1
                                        'разбиваем msg на куски по 253 символа. если кусок сообщения не нулевой длины, добавляем в конец "_"
                                        shortmsg = tmpT.Substring(0, Math.Min(maxMsgLen - 1, tmpT.Length))
                                        tmpT = tmpT.Substring(shortmsg.Length)
                                        trB(i) &= mapInfoKeys(infotick)
                                        If k > 1 Or infotick = 6 Then
                                            trB(i) &= k
                                        End If
                                        Dim dL As Integer = 0
                                        If shortmsg.Length > 0 Then dL = 1
                                        trB(i) &= inttochar(shortmsg.Length + 1 + dL)
                                        For q As Integer = 0 To 2 Step 1
                                            trB(i) &= inttochar(0)
                                        Next q
                                        trB(i) &= shortmsg
                                        If shortmsg.Length > 0 Then trB(i) &= "_"
                                        If k < 5 Then trB(i) &= inttochar(0)
                                    Next k
                                    msg = ""
                                Else
                                    trB(i) &= mapInfoKeys(infotick)
                                    If infotick = 0 Then trB(i) &= My.Resources.scenDescStart
                                    If Not mapInfoKeys(infotick) = "" Then
                                        trB(i) &= inttochar(L + 1)
                                        For q As Integer = 0 To 2 Step 1
                                            trB(i) &= inttochar(0)
                                        Next q
                                        trB(i) &= subtext1
                                    End If
                                End If
                                infotick += 1
                            End If
                            trB(i) &= msg
                            If Not b.isSpecial Then
                                trB(i) &= subtext2 & My.Resources.MsgTextEnd
                            ElseIf Not b.isInfo Then
                                For n As Integer = msg.Length To b.msgLen - 1 Step 1
                                    trB(i) &= inttochar(0)
                                Next n
                            ElseIf b.isInfo And infotick > 1 And Not infotick = 8 Then
                                trB(i) &= inttochar(0)
                            End If
                            txt &= trB(i)
                        End If
                    Else
                        MsgBox("Неожиданный ID блока. Текущий: " & trans(tline).Split(" ")(1) & " ; Ожидаемый: " & i)
                        End
                    End If
                Else
                    txt &= trB(b.setMeAs)
                End If
            Else
                txt &= b.Text
            End If
        Next i
        If err = "" Then
            Dim bytes(txt.Length - 1) As Byte
            For i As Integer = 0 To UBound(bytes) Step 1
                bytes(i) = chartoint(txt(i))
            Next i
            IO.File.WriteAllBytes(PathTextBox.Text & ".translated.sg", bytes)
        Else
            MsgBox(err)
        End If

        'Dim initFile() As Byte = IO.File.ReadAllBytes(PathTextBox.Text)
        'Dim newFile() As Byte = IO.File.ReadAllBytes(PathTextBox.Text & ".translated.sg")
        'Dim context, c1, c2 As String
        'For i As Integer = 0 To UBound(initFile) Step 1
        '    If Not initFile(i) = newFile(i) Then
        '        context = ""
        '        For j As Integer = i - 20 To i - 1 Step 1
        '            context &= inttochar(initFile(j))
        '        Next
        '        c1 = inttochar(initFile(i))
        '        c2 = inttochar(newFile(i))
        '        i = i
        '    End If
        'Next
    End Sub
    Private Function TextStringLenTest(ByRef L As Integer, ByRef b As Block, ByRef infotick As Integer) As Boolean
        If b.isInfo Then
            If infotick = 4 Or infotick = 6 Then
                If L > (maxMsgLen - 1) * 5 Then Return False
            Else
                If L > maxMsgLen Then Return False
            End If
        Else
            If L > maxMsgLen Then Return False
        End If
        Return True
    End Function

    Sub help() Handles HelpB.Click

        Dim msg As String = "В текстбокс забиваем адрес файла карты." & vbNewLine &
                            "Кнопка Parse - читает файл и сохраняет блоки с текстовыми сообщениями" & vbNewLine &
                            "               в файл 'путь к карте'.messages.txt" & vbNewLine &
                            "               строчки, начинающиеся с #block, не трогаем. Не трогаем и" & vbNewLine &
                            "               строчки с несколькими непонятными числами." & vbNewLine &
                            "               После чтения карты переименовываем 'путь к карте'.messages.txt" & vbNewLine &
                            "               в 'путь к карте'.translated_messages.txt" & vbNewLine &
                            "Кнопка Make  - собирает новый файл карты. Читает оригинальный файл" & vbNewLine &
                            "               карты и файл с измененным блоками текста" & vbNewLine &
                            "               максимум символов в сообщении - " & maxMsgLen & ", D2 не знает буквы ё." & vbNewLine &
                            "После изменения нескольких блоков лучше пересобирать карту и" & vbNewLine &
                            "проверять, запускается ли она в редакторе карт" & vbNewLine &
                            "Кнопка Test - и не должна быть активна"
        MsgBox(msg)
    End Sub
End Class

