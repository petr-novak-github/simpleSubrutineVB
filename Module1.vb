Module Module1
    'cz-Vytvoř subrutinu, ktera secte dve cisla a pouzij ji v hlavnim programu.
    'en-Create a Subrutine which will sum two numbers and use it in main program.
    Sub Main()
        Dim sumTwoNums As Integer
        Dim n1 As Integer
        Dim n2 As Integer
        Dim ret As String

        n1 = InputBox("Type first number.(Integer)")
        n2 = InputBox("Type second number.(Integer)")

        sumTwoNumsS(n1, n2, sumTwoNums)
        ret = "Sumation of these two numbers (using another subrutine) is: " + Str(sumTwoNums)
        MsgBox(ret)

    End Sub
    Sub sumTwoNumsS(a As Integer, b As Integer, ByRef vys As Integer)
        vys = a + b
    End Sub

End Module

' notice ByRef in subrutine declaration need to be there
' for comparison here is how to solve same problem using function instead another subrutine

'Module Module1
'    Sub Main()
'        Dim sumTwoNums As Byte
'        Dim n1 As Integer
'        Dim n2 As Integer
'        Dim ret As String

'        n1 = InputBox("Type first number.(Integer)")
'        n2 = InputBox("Type second number.(Integer)")

'        sumTwoNums = sumTwoNumsF(n1, n2)
'        ret = "Sumation of these two numbers is: " + Str(sumTwoNums)
'        MsgBox(ret)

'    End Sub

'    Function sumTwoNumsF(a As Integer, b As Integer) As Integer
'        sumTwoNumsF = a + b
'    End Function

'End Module



