Imports System.IO

Module ModuleSend

    ' 
    'https://itmanagerslife.blogspot.com/2011/12/send-attachment-with-default-email.html
    ' RFC 821 - SIMPLE MAIL TRANSFER PROTOCOL https://www.rfc-editor.org/rfc/rfc821#page-4


    Public MaxBodyLength = 1024 ' Max length of body text in characters
    Public Function Split(ByVal expression As String, ByVal delimiter As String, ByVal qualifier As String, ByVal ignoreCase As Boolean) As String()
        ' Based on the work of LSteinle
        ' http://www.codeproject.com/KB/dotnet/TextQualifyingSplit.aspx?fid=336054&select=1797240&fr=1#xx0xx

        Dim _QualifierState As Boolean = False
        Dim _StartIndex As Integer = 0
        Dim _Values As New System.Collections.ArrayList

        For _CharIndex As Integer = 0 To expression.Length - 1
            If Not qualifier Is Nothing AndAlso String.Compare(expression.Substring(_CharIndex, qualifier.Length), qualifier, ignoreCase) = 0 Then
                _QualifierState = Not _QualifierState
            ElseIf Not _QualifierState AndAlso Not delimiter Is Nothing AndAlso String.Compare(expression.Substring(_CharIndex, delimiter.Length), delimiter, ignoreCase) = 0 Then
                _Values.Add(expression.Substring(_StartIndex, _CharIndex - _StartIndex))
                _StartIndex = _CharIndex + 1
            End If
        Next

        If _StartIndex < expression.Length Then _Values.Add(expression.Substring(_StartIndex, expression.Length - _StartIndex))

        Dim _returnValues(_Values.Count - 1) As String
        _Values.CopyTo(_returnValues)
        Return _returnValues
    End Function

    ' Based on the work of David M Brooks
    ' http://www.codeproject.com/KB/IP/SendFileToNET.aspx

    Sub Main()
        'Environment.GetCommandLineArgs()
        Dim argc As Integer
        Dim argv As System.Collections.ObjectModel.ReadOnlyCollection(Of String)


        argv = My.Application.CommandLineArgs
        argc = argv.Count

        If (argc <= 1) Then
            Dim Product As String
            Product = "SendToVB Vers." & My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString & " - Freeware 2012 by IT-manager's Life."

            Console.WriteLine(Product)
            Console.WriteLine("IT-manager's Life: https://itmanagerslife.blogspot.com")
            Console.WriteLine("Programmatically send emails with attachments using the default email client")
            Console.WriteLine("It need .NET Framework 2.0 installed")
            'Console.WriteLine("")
            Console.WriteLine("USAGE1: SendToVB.exe -files <file1> -body <text> -to <address> -subject <text>")
            Console.WriteLine(" -files <file1> <file2> ... Attach multiple files separating by a space")
            Console.WriteLine(" -body <text>               Add the message enclosed in double quotes")
            Console.WriteLine(" -bodyfile <bodyfile.txt>   For long message (max 1024 chars), use a ""bodyfile.txt""")
            Console.WriteLine(" -to <address1>;<address2>  Send to multiple recipient address separating by ;")
            Console.WriteLine(" -cc <address1>;<address2>  Send to multiple Carbon Copy address")
            Console.WriteLine(" -bcc <address1>;<address2> Send to multiple Blind Carbon Copy address")
            Console.WriteLine(" -subject <content>         Add the Title of message enclosed in double quotes ")
            Console.WriteLine(" -mailto                    Force to use mailto:(pay attention with attachment)")
            Console.WriteLine(" -verbose                   Show more instructions for debugging")
            Console.WriteLine("Example1: SendToVB.exe -to me@example.com;you@example.com -subject ""Yes I am""")
            Console.WriteLine("Example2: SendToVB.exe -body ""Lorem ipsum"" -to it@example.com -subject ""Ok""")
            Console.WriteLine("Example3: SendToVB.exe -bodyfile c:\text.txt -to it@example.com -subject ""Ok""")
            Console.WriteLine("Example4: SendToVB.exe -files ""c:\my files\ok.ppt"" c:\ok.doc -to it@example.com")
            Console.WriteLine("Use first letter to abbreviate the flags: -b=body, -s=subject (except -bcc)")
            'Console.WriteLine("")
            Console.WriteLine("USAGE2: SendToVB.exe -list <listfile.txt>")
            Console.WriteLine(" -list <listfile.txt>       listfile.txt must have one line with only flags")
            Console.WriteLine("Example5: SendToVB.exe -list ""c:\flags.txt""")
            Console.WriteLine("Use this with 'Strings Too Long' problems.")
            Exit Sub
        Else
            'Console.WindowHeight = 1
            'Console.WindowWidth = 1

            'Console.SetWindowSize(1, 1)
        End If


        Dim i As Integer = 0
        Dim n As Integer = 0
        Dim ibody As Integer = -1
        Dim ibodyfile As Integer = -1
        Dim isubject As Integer = -1
        Dim ito As Integer = -1
        Dim icc As Integer = -1
        Dim ibcc As Integer = -1
        Dim imailto As Integer = -1
        Dim iverbose As Integer = -1
        Dim param As String = ""
        Dim ToText As String = ""
        Dim CCText As String = ""
        Dim BCCText As String = ""
        Dim BodyText As String = ""
        Dim SubjectText As String = ""
        Dim ListFileAttach As New Collection
        Dim PathAttach As String = ""
        Dim icommandfile As Integer = -1
        Dim CommandText As String = ""
        Dim vText As String = ""

        Dim CurrentPath As String = Directory.GetCurrentDirectory()



        '------------------------------------------------------------------------------------------------------------------------------------------------
        i = 0
        param = argv(i)
        i = i + 1

        If (System.String.Compare(param, "-list") = 0) Or (System.String.Compare(param, "-l") = 0) Then

            PathAttach = argv(i)
            If Left$(PathAttach, 2) = ".\" Then
                PathAttach = CurrentPath & "\" & Mid$(PathAttach, 3)
            End If

            icommandfile = i
            If File.Exists(PathAttach) Then
                Dim reader As IO.StreamReader = New IO.StreamReader(PathAttach)
                Try
                    Do
                        CommandText = CommandText & reader.ReadLine & vbCrLf
                    Loop Until reader.Peek = -1
                Catch

                Finally
                    reader.Close()
                End Try

                Dim Elements As New List(Of String)
                Dim readOnlyElements As New System.Collections.ObjectModel.ReadOnlyCollection(Of String)(Elements)

                CommandText = Replace(CommandText, vbCrLf, "")

                For Each _Part As String In Split(CommandText, " ", """", True)
                    Elements.Add(Replace(_Part, """", ""))
                Next

                argv = readOnlyElements
                argc = argv.Count

                vText += "CommandText =" & CommandText & vbCrLf

            Else
                Console.WriteLine("Missing CommandFile: " & PathAttach)
                Call Verbose(iverbose, vText, ToText, SubjectText, BodyText)
                Exit Sub

            End If

        ElseIf (String.Compare(param, "-verbose") = 0) Or (String.Compare(param, "-v") = 0) Then
            iverbose = i
            i = i + 1

        End If
        '------------------------------------------------------------------------------------------------------------------------------------------------

        i = 0
        While (i < argc)

            param = argv(i)
            i = i + 1
            If (String.Compare(param, "-files") = 0) Or (String.Compare(param, "-f") = 0) Then
                While (i < argc AndAlso Left$(argv(i), 1) <> "-")
                    PathAttach = argv(i)
                    If Left$(PathAttach, 2) = ".\" Then
                        PathAttach = CurrentPath & "\" & Mid$(PathAttach, 3)
                    End If
                    If File.Exists(PathAttach) Then
                        ListFileAttach.Add(PathAttach)
                        n = n + 1
                        i = i + 1
                    Else
                        Console.WriteLine("Missing Attach file: " & PathAttach)
                        Call Verbose(iverbose, vText, ToText, SubjectText, BodyText)
                        Exit Sub
                    End If
                End While

            ElseIf (String.Compare(param, "-body") = 0) Or (String.Compare(param, "-b") = 0) Then
                ibody = i
                BodyText = argv(ibody)

                If BodyText.Length > MaxBodyLength Then
                    Console.WriteLine("Body too long: " & BodyText.Length & " characters > " & MaxBodyLength)
                    Exit Sub
                End If
                i = i + 1

            ElseIf (String.Compare(param, "-bodyfile") = 0) Then
                PathAttach = argv(i)
                If Left$(PathAttach, 2) = ".\" Then
                    PathAttach = CurrentPath & "\" & Mid$(PathAttach, 3)
                End If

                ibodyfile = i
                If File.Exists(PathAttach) Then
                    Dim reader As IO.StreamReader = New IO.StreamReader(PathAttach)
                    Try
                        Do
                            BodyText = BodyText & reader.ReadLine & vbCrLf
                        Loop Until reader.Peek = -1
                    Catch

                    Finally
                        reader.Close()
                    End Try

                    If BodyText.Length > MaxBodyLength Then
                        Console.WriteLine("Body too long: " & BodyText.Length & " characters > " & MaxBodyLength)
                        Exit Sub
                    End If

                    i = i + 1
                Else
                    Console.WriteLine("Missing BodyFile: " & PathAttach)
                    Call Verbose(iverbose, vText, ToText, SubjectText, BodyText)
                    Exit Sub
                End If

            ElseIf (String.Compare(param, "-subject") = 0) Or (String.Compare(param, "-s") = 0) Then
                isubject = i
                SubjectText = argv(isubject)
                i = i + 1

            ElseIf (String.Compare(param, "-to") = 0) Or (String.Compare(param, "-t") = 0) Then
                ito = i
                ToText = argv(ito)
                i = i + 1

            ElseIf (String.Compare(param, "-cc") = 0) Or (String.Compare(param, "-c") = 0) Then
                icc = i
                CCText = argv(icc)
                i = i + 1

            ElseIf (String.Compare(param, "-bcc") = 0) Then
                ibcc = i
                BCCText = argv(ibcc)
                i = i + 1

            ElseIf (String.Compare(param, "-mailto") = 0) Or (String.Compare(param, "-m") = 0) Then
                imailto = i
                i = i + 1

            ElseIf (String.Compare(param, "-verbose") = 0) Or (String.Compare(param, "-v") = 0) Then
                iverbose = i
                i = i + 1

            End If
        End While

        Call Verbose(iverbose, vText, ToText, SubjectText, BodyText)

        If imailto > -1 Or n = 0 Then
            Dim CarbonCopy As String = ""
            If Len(CCText) > 0 Then
                CarbonCopy = CarbonCopy & "cc=" & CCText & "&"
            End If
            If Len(BCCText) > 0 Then
                CarbonCopy = CarbonCopy & "bcc=" & BCCText & "&"
            End If
            System.Diagnostics.Process.Start("mailto:" & ToText & "?" & CarbonCopy & "subject=" & SubjectText & "&body=" & BodyText) ' & "&attachment=c:\readme.txt"
        Else
            Dim mapi As New SendFileTo.MAPI

            For Each SingleFileAttach As String In ListFileAttach
                mapi.AddAttachment(SingleFileAttach)    'mapi.AddAttachment("c:\\temp\\file1.txt")
            Next
            mapi.AddRecipientTo(ToText)  'mapi.AddRecipientTo("person2@somewhere.com")
            If Len(CCText) > 0 Then
                mapi.AddRecipientCC(CCText) ' Carbon Copy
            End If
            If Len(BCCText) > 0 Then
                mapi.AddRecipientBCC(BCCText) ' BLIND Carbon Copy
            End If
            mapi.SendMailPopup(SubjectText, BodyText)
        End If

    End Sub


    Sub Verbose(ByVal iverbose As Integer, ByVal vText As String, ByVal Totext As String, ByVal SubjectText As String, ByVal BodyText As String)
        If iverbose > -1 Then
            Console.WriteLine("to: " & Totext & vbCrLf & "subject: " & SubjectText & vbCrLf & "body: " & BodyText & vbCrLf)
            Console.WriteLine(vText)
        End If
    End Sub




End Module






