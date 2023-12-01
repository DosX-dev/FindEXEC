'' MADE WITH <3 BY DOSX
'' Coded by DosX
'' GitHub: https://github.com/DosX-dev

'' Attention! This is old legacy code. It doesn't work well. And I'm too lazy to fix it. Keep in mind :(

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

Module Module1

    ' ========================
    Const STD_OUTPUT_HANDLE As Integer = -11
    Const ENABLE_VIRTUAL_TERMINAL_PROCESSING As UInteger = 4
    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function GetStdHandle(ByVal nStdHandle As Integer) As IntPtr
    End Function
    <DllImport("kernel32.dll")>
    Private Function GetConsoleMode(ByVal hConsoleHandle As IntPtr, <Out> ByRef lpMode As UInteger) As Boolean
    End Function
    <DllImport("kernel32.dll")>
    Private Function SetConsoleMode(ByVal hConsoleHandle As IntPtr, ByVal dwMode As UInteger) As Boolean
    End Function
    Const _UnderLine As String = ChrW(27) & "[4m" ' Underline text format
    Const _ResetUnderLine As String = ChrW(27) & "[0m" ' Underline reset
    Sub UpgradeConsole()
        Dim ConFormatHandle = GetStdHandle(STD_OUTPUT_HANDLE)
        Dim ConMode As UInteger
        GetConsoleMode(ConFormatHandle, ConMode)
        ConMode = ConMode Or ENABLE_VIRTUAL_TERMINAL_PROCESSING
        SetConsoleMode(ConFormatHandle, ConMode)
    End Sub
    Sub ClrOut(Text As String, Color1 As ConsoleColor, Color2 As ConsoleColor, NewLine As Boolean) ' Custom colored output
        Console.BackgroundColor = Color1 : Console.ForegroundColor = Color2
        If NewLine Then
            Console.WriteLine(Text)
        Else
            Console.Write(Text)
        End If
        Console.ResetColor()
    End Sub
    Sub EndOfColoredText() ' Console window resizing fix
        ClrOut(".", Console.BackgroundColor, Console.BackgroundColor, True)
    End Sub

    Sub RemoveLastText(_Lenght)
        Try
            Console.Write(Space(10))
            Dim Len = Console.CursorLeft - _Lenght - 10
            Console.SetCursorPosition(Len, Console.CursorTop)
            Console.Write(Space(Len)) ' Remove {StartupText}
            Console.SetCursorPosition(Len, Console.CursorTop)
        Catch ex As Exception : End Try
    End Sub
    ' ========================

    ReadOnly InfoBorder = $" +----------------------------+{vbCrLf} %TEXT%{vbCrLf} +----------------------------+"
    Dim Dirs As String() = {"exec-sorted\NET\DLL"},
        ConsoleTitleDefault As String = Console.Title,
        NETStat As Integer = 0, NATIVEStat As Integer = 0, EXECount As Integer = 0,
        SelectedDirectory As String,
        IsEnd As Boolean = False ' Indicates whether the program has completed it's work
    Sub Main()
        UpgradeConsole()
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
        ClrOut($"{_UnderLine}https://github.com/DosX-dev/FindEXEC{_ResetUnderLine}", ConsoleColor.Black, ConsoleColor.Blue, False) : EndOfColoredText()

        Dim StartupText = " [~] Select a directory... "
        ClrOut(StartupText, ConsoleColor.Black, ConsoleColor.Yellow, False)

        Dim SelectDirectory = New Windows.Forms.FolderBrowserDialog
        SelectDirectory.Description = "Select a folder for sorting binary files."
        SelectDirectory.SelectedPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) ' Default directory

        If SelectDirectory.ShowDialog() = Windows.Forms.DialogResult.OK Then
            NETStat = 0 : NATIVEStat = 0 : EXECount = 0
            Dim _Main = New Thread(AddressOf LetsWork) : _Main.Start(SelectDirectory.SelectedPath) ' Let's work!
        Else
            ClrOut("Abort", ConsoleColor.Black, ConsoleColor.Red, True)
            End
        End If

        Dim CurTaskLength = StartupText.Length
        RemoveLastText(StartupText.Length)

        Do
            Dim StatCommand = Console.ReadKey(True)
            If IsEnd Then
                End
            Else
                Console.Write(Space(70))
                RemoveLastText(70)
                If Not Console.CursorLeft > 0 Then
                    Select Case StatCommand.Key
                        Case ConsoleKey.H ' Help
                            ClrOut(InfoBorder.Replace("%TEXT%", "{H} - Help | {S} - Statistics"),
                                    ConsoleColor.Black,
                                   ConsoleColor.Gray, True)
                        Case ConsoleKey.S ' Statistics
                            ClrOut(InfoBorder.Replace("%TEXT%", $"PE files detected => {EXECount & vbCrLf} | NATIVE => {NATIVEStat & vbCrLf} | NET => {NETStat}"),
                                   ConsoleColor.Black,
                                   ConsoleColor.Gray, True)
                    End Select

                End If
            End If
        Loop
    End Sub
    Sub LetsWork(DirectoryPath)
        SelectedDirectory = DirectoryPath

        Console.Write($" [!] Output directory: ")
        ClrOut($"{SelectedDirectory}\{_UnderLine}exec-sorted{_ResetUnderLine}", ConsoleColor.Black, ConsoleColor.White, False) : EndOfColoredText()
        Console.WriteLine()

        For Each _CurDir In Dirs
            If Not Directory.Exists($"{SelectedDirectory}\{_CurDir}") Then
                Directory.CreateDirectory($"{SelectedDirectory}\{_CurDir}")
            End If
        Next

        Dim Counter = 0,
            GlobalCounter = 0,
            Files = Directory.GetFiles(SelectedDirectory)

        For Each CurFile In Files
            GlobalCounter += 1
            If Not CurFile = Process.GetCurrentProcess().MainModule.FileName Then
                Dim ExeData = File.ReadAllBytes(CurFile),
                    Prefix = $"[{Int(GlobalCounter / Files.Length * 100)}%][{GlobalCounter}/{Files.Length}]",
                    FileName = Path.GetFileName(CurFile),
                    FileSize = {ExeData.Length \ 1024, "Kb"}

                Try

                    Dim ProcText = {" Analyzing ", FileName, "...", "."}
                    ClrOut(ProcText(0), ConsoleColor.DarkGray, ConsoleColor.White, False)
                    ClrOut((_UnderLine & ProcText(1) & _ResetUnderLine), ConsoleColor.DarkGray, ConsoleColor.Gray, False)
                    ClrOut(ProcText(2), ConsoleColor.DarkGray, ConsoleColor.White, False)
                    ClrOut(ProcText(3), Console.BackgroundColor, Console.BackgroundColor, False)
                    RemoveLastText(ProcText(0).Length + ProcText(1).Length + ProcText(2).Length + ProcText(3).Length)

                    If IsBinaryEXE(ExeData) Then
                        Counter += 1

                        If FileSize(0) > 1023 Then
                            FileSize = {FileSize(0) \ 1024, "Mb"}
                        End If

                        Console.Title = $"{Prefix} FindEXEC [{FileName}] [{FileSize(0)} {FileSize(1)}]"
                        Dim NET_Info = IsNET(ExeData)
                        If NET_Info(0) Then
                            If NET_Info(2) = "EXE" Then
                                Dim PathToSave = $"{SelectedDirectory}\exec-sorted\NET\{NET_Info(1)}_{FileName}"
                                If Not File.Exists(PathToSave) Then
                                    File.Copy(CurFile, PathToSave)
                                End If
                            Else
                                Dim PathToSave = $"{SelectedDirectory}\exec-sorted\NET\DLL\{NET_Info(1)}_{FileName}"
                                If Not File.Exists(PathToSave) Then
                                    File.Copy(CurFile, PathToSave)
                                End If
                            End If
                            ProcessLog(Prefix, FileName, ".NET", NET_Info(1).Replace("_", " "), True, NET_Info(2), IsIncludesPDB(ExeData))
                            NETStat += 1
                        Else

                            Dim NativeInfo = GuessNativeRuntime(ExeData),
                                PathToSave = $"{SelectedDirectory}\exec-sorted\{IIf(NativeInfo IsNot "??", NativeInfo, "Unknown")}_{FileName}"

                            If Not File.Exists(PathToSave) Then
                                File.Copy(CurFile, PathToSave)
                            End If
                            ProcessLog(Prefix, FileName, "NATIVE", NativeInfo, False, NET_Info(2), IsIncludesPDB(ExeData))
                            NATIVEStat += 1
                        End If
                        EXECount += 1
                    Else ' If file is not binary

                    End If
                Catch Exc As Exception
                    ClrOut($"Exception occurred: {_UnderLine & Exc.Message & _ResetUnderLine}", ConsoleColor.Black, ConsoleColor.Red, False) : EndOfColoredText()
                End Try
            End If
        Next
        Console.Title = ConsoleTitleDefault
        Console.WriteLine()
        ClrOut(" - - - ", ConsoleColor.Black, ConsoleColor.Yellow, False)
        ClrOut(" Files sorted! Press any key to exit... ", ConsoleColor.DarkGreen, ConsoleColor.White, False)
        ClrOut(" - - - ", ConsoleColor.Black, ConsoleColor.Yellow, True)

        IsEnd = True
    End Sub

    Sub ProcessLog(Prefix As String, FileName As String, Platform As String, Language As String, Detected As Boolean, FileProjectType As String, Optional PDB As Boolean = False)
        ClrOut($"{Prefix}", ConsoleColor.Black, ConsoleColor.DarkGray, False)
        ClrOut($" [{FileProjectType}] ", ConsoleColor.Black, ConsoleColor.Gray, False)
        ClrOut($"{FileName}", ConsoleColor.Black, ConsoleColor.DarkGray, False)
        Console.Write(" => ")
        ClrOut($"{Language} ", ConsoleColor.Black, ConsoleColor.Yellow, False)
        ClrOut($"({Platform})", ConsoleColor.Black, IIf(Detected, ConsoleColor.Green, ConsoleColor.Red), False)
        If PDB Then
            ClrOut(" {PDB}", ConsoleColor.Black, ConsoleColor.DarkGray, False)
        End If
        Console.WriteLine()
    End Sub
    Function IsIncludesPDB(ExeData)
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        If InputData.Contains(".pdb") Then
            Return True
        End If
        Return False
    End Function
    Function IsBinaryEXE(ExeData)
        Dim InputData = Encoding.UTF8.GetString(ExeData),
            TextSigns = ".dll,pe"

        For Each Sign In TextSigns.Split(","c)
            If Not InputData.ToLower().Contains(Sign) Then
                Return False
            End If
        Next

        If IndexOf(ExeData, {0, 3, 0}) = 3 Then ' Checking for "\x{00}\x{03}\x{00}"
            If InputData.Length > 700 Then
                If InputData.Substring(0, 2) = "MZ" Then ' Detect for DOS prefix
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Function IsNET(ExeData) As Object()
        Dim FileProjectType = "BIN"

        ' <BEGIN>\x{00}\x{00}PE\x{00}\x{00}<..ENTROPY (~238)..>H
        Dim HeaderShift = IndexOf(ExeData, ByteStr("{NUL}{NUL}PE{NUL}{NUL}")) ' Offset of 'PE' section; Skip [e_lfanew]
        If Not (Convert.ToChar(ExeData(HeaderShift + 238)) = "H"c AndAlso
                Convert.ToChar(ExeData(HeaderShift + 263)) = " "c AndAlso
                               ExeData(HeaderShift + 239) = 0 AndAlso
                               ExeData(HeaderShift + 249) = 0) Then
            Return {False, "NATIVE", FileProjectType}
        End If
        ' Legacy second char - Convert.ToChar(ExeData(HeaderShift + 96)) = "@"c

        '_CorExeMain - EXE; _CorDllMain - DLL
        Dim BinToLower = ToLowerInBinary(ExeData)
        If (IndexOf(BinToLower, ByteStr("{NUL}mscoree.dll")) OrElse
            IndexOf(BinToLower, ByteStr("{NUL}mscorlib.dll"))) AndAlso (IndexOf(ExeData, ByteStr("{NUL}System."))) Then

            If IndexOf(ExeData, ByteStr("{NUL}_CorExeMain")) Then
                FileProjectType = "EXE" ' .NET exe
            ElseIf IndexOf(ExeData, ByteStr("{NUL}_CorDllMain")) Then : FileProjectType = "DLL" ' .NET dll
            End If

            If Not FileProjectType = "BIN" Then
                If IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic{NUL}")) AndAlso
                   IndexOf(ExeData, ByteStr("{NUL}Microsoft.VisualBasic.CompilerServices{NUL}")) Then
                    Return {True, "VB_NET", FileProjectType}
                ElseIf IndexOf(ExeData, ByteStr("{NUL}Microsoft.JScript{NUL}")) AndAlso
                       IndexOf(ExeData, ByteStr("{NUL}Microsoft.JScript.Vsa{NUL}")) Then
                    Return {True, "JScript", FileProjectType}
                ElseIf IndexOf(ExeData, ByteStr("{NUL}Borland.")) Then
                    Return {True, "Delphi", FileProjectType}
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
            ChangedData = ReplaceBytes(ChangedData, Encoding.UTF8.GetBytes(CurStr.ToString), Encoding.UTF8.GetBytes(CustomToLower(CurStr.ToString)))
        Next
        Return ChangedData
    End Function

    Function CustomToLower(InputData) ' Analog of ToLower() but faster (Only for ENG)
        Dim Result = InputData,
            UPP = "QWERTYUIOPASDFGHJKLZXCVBNM",
            DWN = "qwertyuiopasdfghjklzxcvbnm"

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
            MsgBox(ex.Message, 16)
        End Try
    End Function

    Function ByteStr(InputStr As String) As Byte() ' {NUL}  ==>  \x{00}
        Return ReplaceBytes(Encoding.ASCII.GetBytes(InputStr), Encoding.ASCII.GetBytes("{NUL}"), {CByte(0)})
    End Function
    Public Function ReplaceBytes(DataToChange As Byte(), ToFind As Byte(), ToReplace As Byte()) As Byte()
        Dim MatchStart As Integer = -1,
            MatchLength As Integer = 0

        Using MemWorker = New IO.MemoryStream
            For Index = 0 To DataToChange.Length - 1
                If DataToChange(Index) = ToFind(MatchLength) Then
                    If MatchLength = 0 Then MatchStart = Index
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
                    MemWorker.WriteByte(DataToChange(Index))
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
            For Searcher As Integer = 0 To (PatternToFind.Length - 1)
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
