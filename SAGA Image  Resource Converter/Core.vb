'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Drawing
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module CoreModule

   'This procedure is started when this program is executed.
   Public Sub Main()
      Try
         Dim OutFile As String = Nothing
         Dim ResourcePath As String = If(GetCommandLineArgs().Count > 1, GetCommandLineArgs().Last().Trim(), ".")

         With My.Application.Info
            Console.WriteLine($"{ .Title} v{ .Version.ToString()} - by: { .CompanyName}, { .Copyright}")
            Console.WriteLine($"{NewLine}{ .Description}")
         End With

         For Each InFile As String In Directory.GetFiles(ResourcePath, "*.res")
            Try
               OutFile = $"{InFile}.png"
               Console.Write($"{InFile} -> {OutFile}")
               ExtractResource(InFile, OutFile)
               Console.WriteLine()
            Catch
               Console.WriteLine(" - NOT AN IMAGE RESOURCE?")
            End Try
         Next InFile

         Console.WriteLine($"{NewLine}Done.")
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure attempts to extract an image from the specified resource file.
   Private Sub ExtractResource(InFile As String, OutFile As String)
      Dim Backtrack As New Integer
      Dim Bit As New Integer
      Dim BitfieldBytes() As Byte = {}
      Dim Data As New Byte
      Dim Data1 As New Byte
      Dim Data2 As New Byte
      Dim DataBytes() As Byte = {}
      Dim Decoded As New List(Of Byte)
      Dim ImageSize As Size = Nothing
      Dim MarkByte As New Byte
      Dim Palette As New List(Of Color)
      Dim Pixel As New Integer
      Dim RunAddend As New Integer
      Dim RunCount As New Integer

      Using ResData As New BinaryReader(New MemoryStream(File.ReadAllBytes(InFile)))
         ImageSize = New Size(ResData.ReadInt16, ResData.ReadInt16)
         ResData.BaseStream.Seek(&H4%, SeekOrigin.Current)

         For PaletteEntry As Integer = 0 To 255
            Palette.Add(Color.FromArgb(&HFF%, ResData.ReadByte(), ResData.ReadByte(), ResData.ReadByte()))
         Next PaletteEntry

         While ResData.BaseStream.Position < ResData.BaseStream.Length
            MarkByte = ResData.ReadByte()
            Select Case MarkByte And &HC0%
               Case &HC0%
                  RunCount = MarkByte And &H3F%
                  Decoded.AddRange(ResData.ReadBytes(RunCount))
               Case &H80%
                  RunCount = (MarkByte And &H3F%) + &H3%
                  Data = ResData.ReadByte()
                  Decoded.AddRange(Enumerable.Repeat(Data, RunCount))
               Case &H40%
                  RunCount = ((MarkByte >> &H3%) And &H7%) + &H3%
                  Backtrack = ResData.ReadByte()
                  DataBytes = Decoded.GetRange(Decoded.Count - Backtrack, RunCount).ToArray()
                  Decoded.AddRange(DataBytes)
               Case &H0%
                  Select Case MarkByte And &H30%
                     Case &H30%
                        RunCount = (MarkByte And &HF%) + &H1%
                        Data1 = ResData.ReadByte()
                        Data2 = ResData.ReadByte()
                        BitfieldBytes = ResData.ReadBytes(RunCount)
                        For Each ByteO As Integer In BitfieldBytes
                           For BitIndex As Integer = &H7% To &H0% Step -&H1%
                              Decoded.Add(If(((ByteO >> BitIndex) And &H1%) = &H0%, Data1, Data2))
                           Next BitIndex
                        Next ByteO
                     Case &H20%
                        RunCount = ToInt32(MarkByte And &HF%) << &H8%
                        RunAddend = ResData.ReadByte()
                        RunCount += RunAddend
                        Decoded.AddRange(ResData.ReadBytes(RunCount))
                     Case &H10%
                        Backtrack = (ToInt32(MarkByte And &HF%) << &H8%)
                        RunAddend = ResData.ReadByte()
                        Backtrack += RunAddend
                        RunCount = ResData.ReadByte()
                        DataBytes = Decoded.GetRange(Decoded.Count - Backtrack, RunCount).ToArray()
                        Decoded.AddRange(DataBytes)
                  End Select
            End Select
         End While
      End Using

      CreateImageFile(Decoded, ImageSize, Palette, OutFile)
   End Sub

   'This procedure creates an image file using the specified information.
   Private Sub CreateImageFile(Decoded As List(Of Byte), ImageSize As Size, Palette As List(Of Color), OutFile As String)
      Try
         Dim ImageO As New Bitmap(ImageSize.Width, ImageSize.Height)

         For y As Integer = 0 To ImageSize.Height - 1
            For x As Integer = 0 To ImageSize.Width - 1
               ImageO.SetPixel(x, y, Palette(Decoded((((y >> &H2%) * ImageSize.Width + x) << &H2%) + (y And &H3%))))
            Next x
         Next y

         ImageO.Save(OutFile, Imaging.ImageFormat.Png)
      Catch ExceptionO As Exception
         Console.WriteLine($"{NewLine}{ExceptionO.Message}")
      End Try
   End Sub
End Module
