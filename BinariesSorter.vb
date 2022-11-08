' Coded by https://github.com/DosX-dev
' Telegram: @DosX_Plus

Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Text

Module Module1
    Dim Dirs As Object() = "exec-sorted|exec-sorted\net|exec-sorted\net\VB_NET|exec-sorted\net\C#_or_IL".Split("|"c)
    Sub Main()
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
                            Console.WriteLine($"{Prefix} {FileName} => .NET ({IIf(IsNET(ExeData)(1) Is "VB_NET", "VB NET", "C# or IL")})")
                        Else
                            File.Copy(File1, $"exec-sorted\{FileName}")
                            Console.WriteLine($"{Prefix} {FileName} => ?? (NATIVE)")

                        End If


                    End If
                Catch Exc As Exception : End Try
            End If

        Next
    End Sub
    Function IsBinaryEXE(ExeData)
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        Dim TextSigns = ".dll,pe,�!�l"
        For Each Sign In TextSigns.Split(","c)
            If Not InputData.Contains(Sign) Then
                Return False
            End If
        Next

        If InputData.Length > 2 Then
            If InputData.Substring(0, 2) = "mz" Then
                Return True
            End If
        End If
        Return False
    End Function

    Function IsNET(ExeData) As Object()
        Dim InputData = Encoding.UTF8.GetString(ExeData).ToLower()
        Dim TextSigns = "stathreadattribute,_corexemain,system.,#guid,#blob,#strings,mscoree.dll"
        For Each Sign In TextSigns.Split(","c)
            If Not InputData.Contains(Sign) Then
                Return {False, "NATIVE"}
            End If
        Next
        If InputData.Contains("microsoft.visualbasic.compilerservices") Then
            Return {True, "VB_NET"}
        Else
            Return {True, "C#_or_IL"}
        End If
    End Function

End Module
