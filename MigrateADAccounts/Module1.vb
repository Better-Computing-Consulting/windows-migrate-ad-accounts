Imports System.IO
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Module Module1

    Sub Main()
        ReadExcelDocCreateCmds()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
    Sub ReadExcelDocCreateCmds()
        Dim GivenName, middleName, Surname, NewPrincipalName, NewName, NewSamAccountName, Initials, NewDisplayName, Password, City, Department, Description, Title, mail As String
        GivenName = "E"
        middleName = "F"
        Initials = "I"
        Surname = "G"
        NewDisplayName = "B"
        NewPrincipalName = "O"
        NewName = "D"
        NewSamAccountName = "S"
        Password = "P"
        City = "Y"
        Department = "AF"
        Description = "AG"
        Title = "AI"
        mail = "V"
        Dim aExcelApp As New Excel.Application
        Dim aExcelWrkbook As Excel.Workbook = aExcelApp.Workbooks.Open("C:\temp\adaccountstomigrate.xlsx")
        For Each ws As Excel.Worksheet In aExcelWrkbook.Worksheets
            If ws.Name.Contains("allusers") Then
                Dim numrows As Integer = ws.Range("A2", ws.Range("A2").End(Excel.XlDirection.xlDown)).Rows.Count
                Dim qt As String = Chr(34)
                Dim unames As New List(Of String)
                For i As Int16 = 2 To numrows + 1
                    Dim UserPrincipalName As String = ws.Range(NewPrincipalName & i).Value.ToString.Trim
                    Dim oldemail As String = ws.Range(mail & i).Value.ToString.Trim.ToLower
                    Dim SamAccountName As String = ws.Range(NewSamAccountName & i).Value.ToString.Trim
                    Dim rProperties As New List(Of ADProperty)
                    With rProperties
                        Dim aname As String = ws.Range(NewName & i).Value.ToString.Trim
                        If unames.Contains(aname) Then
                            For x As Integer = 1 To 5
                                aname = aname & " (" & x & ")"
                                If Not unames.Contains(aname) Then Exit For
                            Next
                        End If
                        unames.Add(aname)
                        .Add(New ADProperty(" -Name ", qt & aname & qt))
                        .Add(New ADProperty(" -GivenName ", New String(qt & ws.Range(GivenName & i).Value.ToString.Trim & qt)))
                        .Add(New ADProperty(" -SamAccountName ", qt & SamAccountName & qt))
                        .Add(New ADProperty(" -DisplayName ", New String(qt & ws.Range(NewName & i).Value.ToString.Trim & qt)))
                        .Add(New ADProperty(" -Path ", New String(qt & "OU=Users,OU=Import,DC=healthcare,DC=org" & qt)))
                        .Add(New ADProperty(" -UserPrincipalName ", qt & UserPrincipalName & qt))
                        .Add(New ADProperty(" -AccountPassword ", New String("(ConvertTo-SecureString " & "'" & ws.Range(Password & i).Value.ToString.Trim & "'" & " -AsPlainText -Force)")))
                        .Add(New ADProperty(" -State ", qt & "California" & qt))
                        .Add(New ADProperty(" -Country ", qt & "US" & qt))
                        .Add(New ADProperty(" -Enabled ", "$true"))
                    End With
                    Dim pscmd As String = "New-ADUser"
                    For Each p As ADProperty In rProperties
                        pscmd &= p.Name & p.Value
                    Next
                    Dim oProperites As New List(Of ADProperty)
                    With oProperites
                        If Not IsNothing(ws.Range(Surname & i).Value) Then
                            .Add(New ADProperty(" -Surname ", New String(qt & ws.Range(Surname & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -Surname ", ""))
                        End If
                        If Not IsNothing(ws.Range(middleName & i).Value) Then
                            .Add(New ADProperty(" -OtherName ", New String(qt & ws.Range(middleName & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -OtherName ", ""))
                        End If
                        If Not IsNothing(ws.Range(Initials & i).Value) Then
                            .Add(New ADProperty(" -Initials ", New String(qt & ws.Range(Initials & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -Initials ", ""))
                        End If
                        If Not IsNothing(ws.Range(City & i).Value) Then
                            .Add(New ADProperty(" -City ", New String(qt & ws.Range(City & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -City ", ""))
                        End If
                        If Not IsNothing(ws.Range(Department & i).Value) Then
                            .Add(New ADProperty(" -Department ", New String(qt & ws.Range(Department & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -Department ", ""))
                        End If
                        If Not IsNothing(ws.Range(Description & i).Value) Then
                            .Add(New ADProperty(" -Description ", New String(qt & ws.Range(Description & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -Description ", ""))
                        End If
                        If Not IsNothing(ws.Range(Title & i).Value) Then
                            .Add(New ADProperty(" -Title ", New String(qt & ws.Range(Title & i).Value.ToString.Trim & qt)))
                        Else
                            .Add(New ADProperty(" -Title ", ""))
                        End If
                    End With
                    For Each p As ADProperty In oProperites
                        If p.isSet Then pscmd &= p.Name & p.Value
                    Next
                    Console.WriteLine(pscmd)
                    Dim smtpcmd As String = "Set-ADUser " & SamAccountName & " -add @{ProxyAddresses=" & qt & "SMTP:" & UserPrincipalName & ",smtp:" & oldemail & qt & " -split " & qt & "," & qt & "}"
                    Console.WriteLine(smtpcmd)
                Next
            End If
        Next
        aExcelWrkbook.Close()
        aExcelApp.Quit()
        aExcelWrkbook = Nothing
        aExcelApp = Nothing
    End Sub
End Module
Class ADProperty
    Public Name As String
    Public Value As String
    Public Sub New(inName As String, inValue As String)
        Name = inName
        Value = inValue
    End Sub
    ReadOnly Property isSet As Boolean
        Get
            If Value.Trim.Length > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
End Class