' Coded by https://github.com/DosX-dev
' Telegram: @DosX_Plus

Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Text

Module Module1
    Dim Dirs As Object() = "exec-sorted|exec-sorted\net|exec-sorted\net\VB_NET|exec-sorted\net\C#_or_IL".Split("|"c)

    Sub Main()
        ClrOut("[!] GitHub of FindEXEC: ", ConsoleColor.Black, ConsoleColor.Gray, False)
        ClrOut("https://github.com/DosX-dev/FindEXEC", ConsoleColor.Black, ConsoleColor.Blue, True)

        Console.WriteLine($"[!] Output directory: \{Dirs(0)}\")
        For Each _CurDir In Dirs
            If Not Directory.Exists(_CurDir) Then Directory.CreateDirectory(_CurDir)
        Next
        Dim Counter = 0
        Dim GlobalCounter = 0
        Dim Files = Directory.GetFiles(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName))
        For Each File1 In Files
            GlobalCounter += 1
            If Not File1 = Process.GetCurrentProcess().MainModule.FileName Then
                Dim ExeData = File.ReadAllBytes(File1)
                Try
                    If IsBinaryEXE(ExeData) Then
                        Counter += 1
                        Dim FileName = Path.GetFileName(File1)
                        Dim Prefix = $"[{GlobalCounter}/{Files.Length}]"
                        If IsNET(ExeData)(0) Then
                            File.Copy(File1, $"exec-sorted\net\{IsNET(ExeData)(1)}\{FileName}")
                            ProcessLog(Prefix, FileName, ".NET", IIf(IsNET(ExeData)(1) Is "VB_NET", "VB NET", "C# or IL"), True)

                        Else
                            File.Copy(File1, $"exec-sorted\{FileName}")
                            ProcessLog(Prefix, FileName, "NATIVE", "??", False)
                        End If
                    End If
                Catch Exc As Exception : End Try
            End If
        Next
        ClrOut("Files sorted! Press any key to exit...", ConsoleColor.DarkGreen, ConsoleColor.White, True)
        Console.ReadKey()
    End Sub

    Sub ProcessLog(Prefix As String, FileName As String, Platform As String, Language As String, Detected As Boolean)
        ClrOut($"{Prefix} ", ConsoleColor.Black, ConsoleColor.Gray, False)
        ClrOut($"{FileName}", ConsoleColor.Black, ConsoleColor.DarkGray, False)
        Console.Write(" => ")
        ClrOut($"{Language} ", ConsoleColor.Black, ConsoleColor.Yellow, False)
        ClrOut($"({Platform})", ConsoleColor.Black, IIf(Detected, ConsoleColor.Green, ConsoleColor.Red), True)
    End Sub
    Sub ClrOut(Text As String, Color1 As ConsoleColor, Color2 As ConsoleColor, NewLine As Boolean)
        Console.BackgroundColor = Color1 : Console.ForegroundColor = Color2
        If NewLine Then
            Console.WriteLine(Text)
        Else
            Console.Write(Text)
        End If
        Console.BackgroundColor = ConsoleColor.Black : Console.ForegroundColor = ConsoleColor.White
    End Sub
    Function IsBinaryEXE(ExeData)
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        Dim TextSigns = ".dll,pe,�!�l"
        For Each Sign In TextSigns.Split(","c)
            If Not InputData.Contains(Sign) Then
                Return False
            End If
        Next

        If IndexOf(ExeData, ByteStr("{NUL}")) Then
            If InputData.Length > 2 Then
                If InputData.Substring(0, 2) = "mz" Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Function IsNET(ExeData) As Object()
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        Dim TextSigns = "stathreadattribute,system.,#guid,#blob,#strings"
        For Each Sign In TextSigns.Split(","c)
            If Not InputData.Contains(Sign) Then
                Return {False, "NATIVE"}
            End If
        Next
        If IndexOf(ExeData, ByteStr("{NUL}_CorExeMain{NUL}mscoree.dll{NUL}")) Then
            If IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic{NUL}")) AndAlso
               IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic.CompilerServices{NUL}")) Then
                Return {True, "VB_NET"}
            Else
                Return {True, "C#_or_IL"}
            End If
        End If
        Return {False, "NATIVE"}
    End Function


    Function ByteStr(InputStr As String) As Byte() ' {NUL}  ==>  \d{00}
        Return ReplaceBytes(Encoding.ASCII.GetBytes(InputStr), Encoding.ASCII.GetBytes("{NUL}"), {CByte(0)})
    End Function
    Public Function ReplaceBytes(DataToChange As Byte(), ToFind As Byte(), ToReplace As Byte()) As Byte()
        Dim MatchStart As Integer = -1
        Dim MatchLength As Integer = 0
        Using MemWorker = New IO.MemoryStream
            For index = 0 To DataToChange.Length - 1
                If DataToChange(index) = ToFind(MatchLength) Then
                    If MatchLength = 0 Then MatchStart = index
                    MatchLength += 1
                    If MatchLength = ToFind.Length Then
                        MemWorker.Write(ToReplace, 0, ToReplace.Length)
                        MatchLength = 0
                    End If
                Else
                    If MatchLength > 0 Then
                        MemWorker.Write(DataToChange, MatchStart, MatchLength)
                        MatchLength = 0
                    End If
                    MemWorker.WriteByte(DataToChange(index))
                End If
            Next
            If MatchLength > 0 Then
                MemWorker.Write(DataToChange, DataToChange.Length - MatchLength, MatchLength)
            End If
            Dim RetVal(MemWorker.Length - 1) As Byte
            MemWorker.Position = 0
            MemWorker.Read(RetVal, 0, RetVal.Length)
            Return RetVal
        End Using
    End Function
    Public Function IndexOf(ByVal ArrayToSearchThrough As Byte(), ByVal PatternToFind As Byte()) As Integer
        If PatternToFind.Length > ArrayToSearchThrough.Length Then Return -1

        For Arr As Integer = 0 To ArrayToSearchThrough.Length - PatternToFind.Length - 1
            Dim Found As Boolean = True

            For Searcher As Integer = 0 To PatternToFind.Length - 1

                If ArrayToSearchThrough(Arr + Searcher) <> PatternToFind(Searcher) Then
                    Found = False
                    Exit For
                End If
            Next

            If Found Then
                Return Arr
            End If
        Next
        Return 0
    End Function
End Module
