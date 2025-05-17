' Author: Md. Zahirul Islam
' Email: zahirul.islam.spa@gmail.com
' Cell: +8801711792629
' Description: VBA code to convert numbers into Bangladeshi Taka and Poisha format.

Function NumberToWords(ByVal MyNumber As String) As String
    Dim Units As Variant
    Dim Tens As Variant
    Dim Temp As String
    Dim DecimalPlace As Integer
    Dim Count As Integer
    Dim DecimalPart As String
    Dim WholePart As String

    ' Units and Tens for conversion
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", _
                  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", _
                  "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

    ' Split MyNumber into whole number (Taka) and decimal part (Poisha)
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        WholePart = Left(MyNumber, DecimalPlace - 1)
        DecimalPart = Mid(MyNumber, DecimalPlace + 1)
    Else
        WholePart = MyNumber
        DecimalPart = ""
    End If

    Temp = ""

    ' Process the whole part of the number (Taka) with "Lac" and "Thousand"
    Count = 1
    Do While WholePart <> ""
        Select Case Count
            Case 1
                Temp = ConvertHundreds(Right(WholePart, 3)) & Temp
            Case 2
                Temp = ConvertHundreds(Right(WholePart, 2)) & " Thousand " & Temp
            Case 3
                Temp = ConvertHundreds(Right(WholePart, 2)) & " Lac " & Temp
            Case 4
                Temp = ConvertHundreds(Right(WholePart, 2)) & " Crore " & Temp
        End Select

        If Len(WholePart) > 3 And Count = 1 Then
            WholePart = Left(WholePart, Len(WholePart) - 3)
        ElseIf Len(WholePart) > 2 Then
            WholePart = Left(WholePart, Len(WholePart) - 2)
        Else
            WholePart = ""
        End If
        Count = Count + 1
    Loop

    NumberToWords = Trim(Temp) & " Taka"

    ' Process the decimal part of the number (Poisha)
    If DecimalPart <> "" Then
        DecimalPart = Left(DecimalPart, 2) ' Limit decimal to two digits
        NumberToWords = NumberToWords & " and " & ConvertHundreds(DecimalPart) & " Poisha Only"
    Else
        NumberToWords = NumberToWords & " Only"
    End If

End Function

Private Function ConvertHundreds(ByVal MyNumber As String) As String
    Dim Result As String
    Dim Units As Variant
    Dim Tens As Variant

    ' Units and Tens for conversion
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", _
                  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", _
                  "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

    Result = ""

    ' Convert hundreds place
    If Val(MyNumber) > 99 Then
        Result = Units(Int(MyNumber / 100)) & " Hundred "
        MyNumber = MyNumber Mod 100
    End If

    ' Convert tens and ones place
    If Val(MyNumber) > 19 Then
        Result = Result & Tens(Int(MyNumber / 10)) & " "
        MyNumber = MyNumber Mod 10
    End If
    Result = Result & Units(MyNumber)

    ConvertHundreds = Trim(Result)
End Function
