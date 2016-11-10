Imports System.IO
Imports System.Drawing

Module MainModule
    Private Const NameSize As Integer = 48
    Private Const DataSize As Integer = 320

    Private Class IconResources
        Property IndexFileName As String
        Property DirectoryPath As String

        Property Name As String
        Property LowResolutionDataFile As String
        Property HighResolutionDataFile As String
        Property XLowResolutionDataFile As String
        Property XHighResolutionDataFile As String

        Property Icons As New List(Of ResourceIcon)

        Private MarkerIndex(4) As Integer

        Sub New(indexFileName As String)
            Me.IndexFileName = IO.Path.GetFileName(indexFileName)
            Me.DirectoryPath = IO.Path.GetDirectoryName(indexFileName)

            Dim buf = File.ReadAllBytes(indexFileName)

            ' Read name, filename
            Dim idx = 0
            For i = 0 To buf.Count - 1
                If buf(i) = &HA Then
                    MarkerIndex(idx) = i
                    idx += 1
                    If idx > 4 Then
                        Exit For
                    End If
                End If
            Next

            Me.Name = System.Text.Encoding.ASCII.GetString(buf.Take(MarkerIndex(0)).ToArray).TrimEnd(ChrW(0))
            Me.LowResolutionDataFile = System.Text.Encoding.ASCII.GetString(buf.Skip(MarkerIndex(0) + 1).Take(MarkerIndex(1) - MarkerIndex(0) - 1).ToArray).TrimEnd(ChrW(0))
            Me.HighResolutionDataFile = System.Text.Encoding.ASCII.GetString(buf.Skip(MarkerIndex(1) + 1).Take(MarkerIndex(2) - MarkerIndex(1) - 1).ToArray).TrimEnd(ChrW(0))
            Me.XLowResolutionDataFile = System.Text.Encoding.ASCII.GetString(buf.Skip(MarkerIndex(2) + 1).Take(MarkerIndex(3) - MarkerIndex(2) - 1).ToArray).TrimEnd(ChrW(0))
            Me.XHighResolutionDataFile = System.Text.Encoding.ASCII.GetString(buf.Skip(MarkerIndex(3) + 1).Take(MarkerIndex(4) - MarkerIndex(3) - 1).ToArray).TrimEnd(ChrW(0))

            ' Read icon block
            Dim offset = MarkerIndex(4) + 1
            Dim structSize = NameSize + DataSize

            Dim count = (buf.Count - offset) / structSize
            Dim dst(structSize - 1) As Byte

            For i = 0 To count - 1
                Buffer.BlockCopy(buf, Convert.ToInt32(offset + i * structSize), dst, 0, structSize)
                Me.Icons.Add(New ResourceIcon(dst))
            Next
        End Sub

        Private Sub OutputIndexFile(workingDirectory As String)
            Dim buf = New List(Of Byte)

            Dim nameBuf = System.Text.Encoding.ASCII.GetBytes(Me.Name)
            buf.AddRange(nameBuf)
            For i = 0 To MarkerIndex(0) - nameBuf.Count - 1
                buf.Add(0)
            Next
            buf.Add(&HA)

            Dim lowFileBuf = System.Text.Encoding.ASCII.GetBytes(Me.LowResolutionDataFile)
            buf.AddRange(lowFileBuf)
            For i = 0 To MarkerIndex(1) - MarkerIndex(0) - lowFileBuf.Count - 2
                buf.Add(0)
            Next
            buf.Add(&HA)

            Dim highFileBuf = System.Text.Encoding.ASCII.GetBytes(Me.HighResolutionDataFile)
            buf.AddRange(highFileBuf)
            For i = 0 To MarkerIndex(2) - MarkerIndex(1) - highFileBuf.Count - 2
                buf.Add(0)
            Next
            buf.Add(&HA)

            Dim xLowFileBuf = System.Text.Encoding.ASCII.GetBytes(Me.XLowResolutionDataFile)
            buf.AddRange(xLowFileBuf)
            For i = 0 To MarkerIndex(3) - MarkerIndex(2) - xLowFileBuf.Count - 2
                buf.Add(0)
            Next
            buf.Add(&HA)

            Dim xHighFileBuf = System.Text.Encoding.ASCII.GetBytes(Me.XHighResolutionDataFile)
            buf.AddRange(xHighFileBuf)
            For i = 0 To MarkerIndex(4) - MarkerIndex(3) - xHighFileBuf.Count - 2
                buf.Add(0)
            Next
            buf.Add(&HA)

            For Each icon In Me.Icons
                buf.AddRange(icon.ToByteArray)
            Next

            File.WriteAllBytes(IO.Path.Combine(workingDirectory, Me.IndexFileName), buf.ToArray)
        End Sub

        Sub Pack(workingDirectory As String)

            Dim lowBuf = New List(Of Byte)
            Dim highBuf = New List(Of Byte)

            lowBuf.AddRange(New Byte() {&H66, &H64, &H72, &H61})
            highBuf.AddRange(New Byte() {&H66, &H64, &H72, &H61})

            Dim lowOffset = lowBuf.Count
            Dim highOffset = highBuf.Count

            For Each icon In Me.Icons
                For i = 0 To 7
                    If icon.Low.Pics(i).Size = 0 Then
                        Continue For
                    End If

                    Dim iconFile = Path.Combine(workingDirectory, "Low", String.Format("{0}_s{1}.png", icon.Key, i))
                    If Not IO.File.Exists(iconFile) Then
                        Console.WriteLine("File not found: " & iconFile)
                        Exit Sub
                    End If
                    Dim buf = File.ReadAllBytes(iconFile)
                    lowBuf.AddRange(buf)

                    icon.Low.Pics(i).Offset = lowOffset
                    icon.Low.Pics(i).Size = buf.Count
                    lowOffset += buf.Count
                Next

                For i = 0 To 7
                    If icon.High.Pics(i).Size = 0 Then
                        Continue For
                    End If

                    Dim iconFile = Path.Combine(workingDirectory, "High", String.Format("{0}_s{1}.png", icon.Key, i))
                    If Not IO.File.Exists(iconFile) Then
                        Console.WriteLine("File not found: " & iconFile)
                        Exit Sub
                    End If
                    Dim buf = File.ReadAllBytes(iconFile)
                    highBuf.AddRange(buf)

                    icon.High.Pics(i).Offset = highOffset
                    icon.High.Pics(i).Size = buf.Count
                    highOffset += buf.Count
                Next
            Next

            File.WriteAllBytes(IO.Path.Combine(workingDirectory, Me.LowResolutionDataFile), lowBuf.ToArray)
            File.WriteAllBytes(IO.Path.Combine(workingDirectory, Me.HighResolutionDataFile), highBuf.ToArray)
            Me.OutputIndexFile(workingDirectory)

        End Sub

        Sub Extract(workingDirectory As String)
            Directory.CreateDirectory(workingDirectory)
            Directory.CreateDirectory(IO.Path.Combine(workingDirectory, "Low"))
            Directory.CreateDirectory(IO.Path.Combine(workingDirectory, "High"))

            Dim lowBuf = File.ReadAllBytes(IO.Path.Combine(Me.DirectoryPath, Me.LowResolutionDataFile))
            Dim highBuf = File.ReadAllBytes(IO.Path.Combine(Me.DirectoryPath, Me.HighResolutionDataFile))

            Dim maxSize = Math.Max(
                Me.Icons.Max(Function(ico) ico.Low.Pics.Max(Function(pic) pic.Size)),
                Me.Icons.Max(Function(ico) ico.High.Pics.Max(Function(pic) pic.Size)))

            Dim dst() As Byte
            ReDim dst(maxSize - 1)

            For Each icon In Me.Icons
                For i = 0 To 7
                    Dim p = icon.Low.Pics(i)
                    If p.Size = 0 Then
                        Continue For
                    End If

                    Buffer.BlockCopy(lowBuf, p.Offset, dst, 0, p.Size)
                    Using fs = New IO.FileStream(Path.Combine(workingDirectory, "Low", String.Format("{0}_s{1}.png", icon.Key, i)), FileMode.Create, FileAccess.Write)
                        Using sw = New IO.BinaryWriter(fs)
                            sw.Write(dst, 0, p.Size)
                        End Using
                    End Using
                Next

                For i = 0 To 7
                    Dim p = icon.High.Pics(i)
                    If p.Size = 0 Then
                        Continue For
                    End If

                    Buffer.BlockCopy(highBuf, p.Offset, dst, 0, p.Size)
                    Using fs = New IO.FileStream(Path.Combine(workingDirectory, "High", String.Format("{0}_s{1}.png", icon.Key, i)), FileMode.Create, FileAccess.Write)
                        Using sw = New IO.BinaryWriter(fs)
                            sw.Write(dst, 0, p.Size)
                        End Using
                    End Using
                Next
            Next
        End Sub
    End Class

    Private Class IconData
        Property Width As Integer
        Property Height As Integer
        Property X As Integer
        Property Y As Integer
        Property Pics As New List(Of PicInfo)
    End Class

    Private Class PicInfo
        Property Offset As Integer
        Property Size As Integer
    End Class

    Private Class ResourceIcon
        Property Key As String
        Property Low As IconData
        Property High As IconData

        Sub New(buffer() As Byte)

            Me.Key = System.Text.Encoding.ASCII.GetString(buffer.Take(NameSize).ToArray).TrimEnd(ChrW(0))

            Dim data = New List(Of Int32)
            For j = 0 To DataSize / 4 - 1
                data.Add(BitConverter.ToInt32(buffer, Convert.ToInt32(NameSize + j * 4)))
            Next

            Me.Low = New IconData With {
                .Width = data(0),
                .Height = data(4),
                .X = data(8),
                .Y = data(12)
            }
            For i = 0 To 8 - 1
                Me.Low.Pics.Add(New PicInfo With {
                            .Offset = data(16 + i),
                            .Size = data(48 + i)})
            Next

            Me.High = New IconData With {
                .Width = data(1),
                .Height = data(5),
                .X = data(9),
                .Y = data(13)
            }
            For i = 0 To 8 - 1
                Me.High.Pics.Add(New PicInfo With {
                            .Offset = data(24 + i),
                            .Size = data(56 + i)})
            Next

        End Sub

        Function ToByteArray() As Byte()
            Dim buf = New List(Of Byte)

            ' Key
            Dim key = System.Text.Encoding.ASCII.GetBytes(Me.Key)
            buf.AddRange(key)
            For i = 0 To NameSize - key.Count - 1
                buf.Add(0)
            Next

            ' Data
            buf.AddRange(BitConverter.GetBytes(Me.Low.Width))
            buf.AddRange(BitConverter.GetBytes(Me.High.Width))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(Me.Low.Height))
            buf.AddRange(BitConverter.GetBytes(Me.High.Height))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(Me.Low.X))
            buf.AddRange(BitConverter.GetBytes(Me.High.X))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(Me.Low.Y))
            buf.AddRange(BitConverter.GetBytes(Me.High.Y))
            buf.AddRange(BitConverter.GetBytes(0))
            buf.AddRange(BitConverter.GetBytes(0))
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(Me.Low.Pics(i).Offset))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(Me.High.Pics(i).Offset))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(0))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(0))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(Me.Low.Pics(i).Size))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(Me.High.Pics(i).Size))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(0))
            Next
            For i = 0 To 8 - 1
                buf.AddRange(BitConverter.GetBytes(0))
            Next
            Return buf.ToArray
        End Function

    End Class

    Sub Main(args As String())
        Dim workingDirectory = Path.Combine(My.Application.Info.DirectoryPath, "Work")
        Dim indexFilePath As String
        Dim packing = False

        If args.Length >= 2 AndAlso (args(0) = "-e" OrElse args(0) = "-p") Then
            indexFilePath = args(1)

            If args.Length >= 3 Then
                workingDirectory = args(2)
            End If

            packing = args(0) = "-p"

        ElseIf args.Length >= 1 AndAlso Path.GetFileName(args(0)) = "IconResources.idx" Then
            ' Drag & Drop

            indexFilePath = args(0)

            If args.Length >= 2 Then
                workingDirectory = args(1)
            End If

            packing = Directory.Exists(Path.Combine(workingDirectory, "Low")) AndAlso Directory.Exists(Path.Combine(workingDirectory, "High"))

        Else
            ShowUsage()
            Exit Sub
        End If


        Dim res = New IconResources(indexFilePath)
        Console.WriteLine(res.Name)

        If packing Then
            ' Pack
            Console.WriteLine("Packing icons...")
            res.Pack(workingDirectory)
        Else
            ' Extract
            Console.WriteLine("Extracting icons...")
            res.Extract(workingDirectory)
        End If

    End Sub

    Private Sub ShowUsage()
        Console.WriteLine("Usage:")
        Console.WriteLine("  Extract icons: pscc2017Icon -e ""C:\Program Files\Adobe\Adobe Photoshop CC 2017\Resources\IconResources.idx"" ""WorkingDirectory""")
        Console.WriteLine("  Pack icons:    pscc2017Icon -p ""C:\Program Files\Adobe\Adobe Photoshop CC 2017\Resources\IconResources.idx"" ""WorkingDirectory""")
    End Sub

End Module
