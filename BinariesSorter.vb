' MADE WITH <3 BY DOSX
' Coded by https://github.com/DosX-dev
' Telegram: @DosX_Plus

Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading

Module Module1
    Dim Dirs As Object() = {"exec-sorted\NET\VB_NET\DLL",
                            "exec-sorted\NET\C#_or_IL\DLL",
                            "exec-sorted\NET\Delphi\DLL",
                            "exec-sorted\NET\JScript\DLL"}
    Dim ConsoleTitleDefault = Console.Title

    Sub Main()
        Dim _Main = New Thread(AddressOf LetsWork) : _Main.Start()
    End Sub

    Sub LetsWork()
        ClrOut("
                                  _ 
  _______    _                   |_|   _______   _       _   _______   ________
 |_|_|_|_|  |_|   ______      ___|_|  |_|_|_|_| |_|_   _|_| |_|_|_|_| /_|_|_|_/
 |_|____     _   |_|_|_|\   _/_|_|_|  |_|____     |_|_|_|   |_|____   |_|
 |_|_|_|    |_|  |_|   |_| |_|   |_|  |_|_|_|      _|_|_    |_|_|_|   |_|
 |_|        |_|  |_|   |_| |_|___|_|  |_|______  _|_| |_|_  |_|______ |_|_____
 |_|        |_|  |_|   |_|   \_|_|_|  |_|_|_|_| |_|     |_| |_|_|_|_| \_|_|_|_\
", ConsoleColor.Black, ConsoleColor.Cyan, True)
        ClrOut(" [?] GitHub of FindEXEC: ", ConsoleColor.Black, ConsoleColor.Gray, False)
        ClrOut("https://github.com/DosX-dev/FindEXEC", ConsoleColor.Black, ConsoleColor.Blue, True)

        Console.WriteLine($" [!] Output directory: \exec-sorted\")
        Console.WriteLine()

        For Each _CurDir In Dirs
            If Not Directory.Exists(_CurDir) Then Directory.CreateDirectory(_CurDir)
        Next
        Dim Counter = 0
        Dim GlobalCounter = 0
        Dim Files = Directory.GetFiles(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName))
        For Each CurFile In Files
            GlobalCounter += 1
            If Not CurFile = Process.GetCurrentProcess().MainModule.FileName Then
                Dim ExeData = File.ReadAllBytes(CurFile)
                Try
                    If IsBinaryEXE(ExeData) Then
                        Counter += 1
                        Dim FileName = Path.GetFileName(CurFile)
                        Dim Prefix = $"[{Int(GlobalCounter / Files.Length * 100)}%][{GlobalCounter}/{Files.Length}]"

                        Dim FileSize = {ExeData.Length \ 1024, "Kb"}

                        If FileSize(0) > 1023 Then
                            FileSize(0) \= 1024
                            FileSize(1) = "Mb"
                        End If

                        Console.Title = $"{Prefix} FindEXEC [{FileName}] [{FileSize(0)} {FileSize(1)}]"
                        Dim NET_Info = IsNET(ExeData)
                        If NET_Info(0) Then
                            If NET_Info(2) = "EXE" Then
                                Dim PathToSave = $"exec-sorted\NET\{NET_Info(1)}\{FileName}"
                                If Not File.Exists(PathToSave) Then
                                    File.Copy(CurFile, PathToSave)
                                End If
                            Else
                                Dim PathToSave = $"exec-sorted\NET\{NET_Info(1)}\DLL\{FileName}"
                                If Not File.Exists(PathToSave) Then
                                    File.Copy(CurFile, PathToSave)
                                End If
                            End If
                            ProcessLog(Prefix, FileName, ".NET", NET_Info(1).Replace("_", " "), True, NET_Info(2))
                        Else
                            Dim NativeInfo = GuessNativeRuntime(ExeData)
                            Dim PathToSave = $"exec-sorted\{IIf(NativeInfo IsNot "??", NativeInfo, "Unknown")}_{FileName}"
                            If Not File.Exists(PathToSave) Then
                                File.Copy(CurFile, PathToSave)
                            End If
                            ProcessLog(Prefix, FileName, "NATIVE", NativeInfo, False, NET_Info(2))
                        End If
                    End If
                Catch Exc As Exception
                    MsgBox(Exc.Message)
                End Try
            End If
        Next
        Console.Title = ConsoleTitleDefault
        Console.WriteLine()
        ClrOut(" - - - ", ConsoleColor.Black, ConsoleColor.Yellow, False)
        ClrOut(" Files sorted! Press any key to exit... ", ConsoleColor.DarkGreen, ConsoleColor.White, False)
        ClrOut(" - - - ", ConsoleColor.Black, ConsoleColor.Yellow, True)
        Console.ReadKey()
    End Sub

    Sub ProcessLog(Prefix As String, FileName As String, Platform As String, Language As String, Detected As Boolean, FileProjectType As String)
        ClrOut($"{Prefix}", ConsoleColor.Black, ConsoleColor.DarkGray, False)
        ClrOut($" [{FileProjectType}] ", ConsoleColor.Black, ConsoleColor.Gray, False)
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
        Console.ResetColor()
    End Sub
    Function IsBinaryEXE(ExeData)
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        Dim TextSigns = ".dll,pe" ' OLD "�!�l"
        For Each Sign In TextSigns.Split(","c)
            If Not InputData.Contains(Sign) Then
                Return False
            End If
        Next

        If IndexOf(ExeData, ByteStr("{NUL}{NUL}")) Then
            If InputData.Length > 700 Then
                If InputData.Substring(0, 2) = "mz" Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Function IsNET(ExeData) As Object()
        Dim FileProjectType = "BIN"

        If Not (Convert.ToChar(ExeData(364)) = "H"c And
                Convert.ToChar(ExeData(182)) = "@"c) Then
            Return {False, "NATIVE", FileProjectType}
        End If

        '_CorExeMain - EXE; _CorDllMain - DLL
        Dim BinToLower = ToLowerInBinary(ExeData)
        If IndexOf(BinToLower, ByteStr("{NUL}mscoree.dll")) OrElse
           IndexOf(BinToLower, ByteStr("{NUL}mscorlib.dll")) Then

            If IndexOf(ExeData, ByteStr("{NUL}_CorDllMain")) Then FileProjectType = "DLL" ' .NET dll
            If IndexOf(ExeData, ByteStr("{NUL}_CorExeMain")) Then FileProjectType = "EXE" ' .NET exe

            If Not FileProjectType = "BIN" Then
                If IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic{NUL}")) AndAlso
                   IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic.CompilerServices{NUL}")) Then
                    Return {True, "VB_NET", FileProjectType}
                ElseIf IndexOf(ExeData, ByteStr("{NUL}Borland.")) Then
                    Return {True, "Delphi", FileProjectType}
                ElseIf IndexOf(ExeData, ByteStr("{NUL}Microsoft.JScript{NUL}")) AndAlso
                       IndexOf(ExeData, ByteStr("{NUL}Microsoft.JScript.Vsa{NUL}")) Then
                    Return {True, "JScript", FileProjectType}
                Else
                    Return {True, "C#_or_IL", FileProjectType}
                End If
            End If
        End If
        Return {False, "NATIVE", FileProjectType}
    End Function
    Public Detects = {"msvcp50.dll=C++ (MS 1998)", "msvcp60.dll=С++ (MS 2000-2001)", ' Microsoft C++ Runtime
                      "msvcp70.dll=С++ (MS 2002)", "msvcp71.dll=C++ (MS 2003)",
                      "msvcp80.dll=C++ (MS 2005)", "msvcp90.dll=C++ (MS 2008)",
                      "msvcp100.dll=C++ (MS 2010)", "msvcp110.dll=C++ (MS 2012)",
                      "msvcp120.dll=C++ (MS 2013)", "msvcp130.dll=C++ (MS 2013)",
                      "msvcp140.dll=C++ (MS 2015-2017)", "msvcp150.dll=C++ (MS 2017-2018)",
                      "msvcp160.dll=C++ (MS 2019)", "msvcrt.dll=C++", "vcruntime140.dll=C++",
                      "libgcj-13.dll=C++ (GCC)", "libgcc_s_dw2-1.dll=C++ (GCC)", ' GNU GCC (C++)
                      "msys-1.0.dll=C++ (GCC)", "libgcj.dll=C++ (GCC)", "cyggcj.dll=C++ (GCC)",
                      "msvcirt.dll=C++", ' Microsoft C++ Library (<iostream.h>)
                      "crtdll.dll=C", ' Microsoft C Runtime
                      "vb40032.dll=VB4", ' Microsoft Visual Basic 4
                      "msvbvm50.dll=VB5", ' Microsoft Visual Basic 5
                      "msvbvm60.dll=VB6", ' Microsoft Visual Basic 6
                      "upx0{NUL}{NUL}=UPX-Packed", ' UPX Packer
                      "{NUL}.mpress1=MPRESS-Packed"} ' MSPRESS native packer
    Function ToLowerInBinary(ExeData) ' Change registry of all chars in Byte() to lower
        Dim ChangedData = ExeData
        For Each CurStr In "QWERTYUIOPASDFGHJKLZXCVBNM"
            ChangedData = ReplaceBytes(ChangedData, ByteStr(CurStr.ToString), ByteStr(CustomToLower(CurStr.ToString)))
        Next
        Return ChangedData
    End Function

    Function CustomToLower(InputData) ' Analog of ToLower() but faster (Only for ENG)
        Dim Result = InputData
        Dim UPP = "QWERTYUIOPASDFGHJKLZXCVBNM"
        Dim DWN = "qwertyuiopasdfghjklzxcvbnm"
        For IndexToReplace = 0 To (UPP.Length - 1)
            Result = Result.Replace(UPP(IndexToReplace), DWN(IndexToReplace))
        Next
        Return Result
    End Function

    Function GuessNativeRuntime(ExeData)
        Try
            Dim AssemblyData = ToLowerInBinary(ExeData)
            For Each SearchForSigns In Detects
                Dim SignAndRuntime = SearchForSigns.Split("=")
                Dim Sign = SignAndRuntime(0)
                Dim Runtime = SignAndRuntime(1)
                If IndexOf(AssemblyData, ByteStr($"{{NUL}}{Sign}{{NUL}}")) Then
                    Return Runtime
                End If
            Next
            Return "??"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

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

' MADE WITH <3 BY DOSX
' Coded by https://github.com/DosX-dev
' Telegram: @DosX_Plus
