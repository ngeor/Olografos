Attribute VB_Name = "Olografoscode"
Option Explicit

Dim N1(1 To 3, 1 To 9) As String
Dim N2(1 To 4) As String

' Represents the gender of the variable
Public Enum GenderEnum
    Masculin = 0
    Feminin = 1
    Neutral = 2
End Enum

Public Function OlografosΔρχ(Num As Double) As String
    OlografosΔρχ = Olografos(Num, Feminin) + " δρχ"
End Function

Public Function Olografos(Num As Double, Gender As GenderEnum) As String
    InitNames
    Dim iNum As Long
    iNum = Int(Num)
    If iNum = 0 Then
        Olografos = "Μηδέν"
    Else
        Olografos = OlografosInt(iNum, Gender)
    End If
End Function

Private Function OlografosInt(iNum As Long, Gender As GenderEnum) As String
    Dim Result As String
    Dim Gen As GenderEnum
    Dim triada As Integer
    Dim temp As Long
    Dim groupValue As Integer
    Dim h As String
    Dim part As String
    
    Result = ""
    
    ' copy input to avoid mutating
    temp = iNum
    
    ' η τριάδα μετριέται από αριστερά
    ' μηδέν -> 0..999
    ' ένα   -> 001_000..999_999
    triada = 0
    While temp > 0
        ' get the last 3 digits
        groupValue = temp Mod 1000
        If groupValue > 0 Then
    
            ' determine gender for this group
            If triada = 0 Then
                Gen = Gender
            ElseIf triada = 1 Then ' χιλιάδες
                Gen = Feminin
            Else ' εκατομμύρια ++
                Gen = Neutral
            End If
            
            If triada = 1 And groupValue = 1 Then
                part = "χίλιες"
            Else
                part = OloTriada(groupValue, Gen)
                If triada > 0 Then
                    h = N2(triada)
                    If triada >= 2 Then
                        If groupValue = 1 Then
                            h = h + "ο"
                        Else
                            h = h + "α"
                        End If
                    End If
                    part = Join(part, h)
                End If
            End If
            Result = Join(part, Result)
        End If
        temp = temp \ 1000
        triada = triada + 1
    Wend

GiveResult:
    Mid(Result, 1, 1) = UCase(Mid(Result, 1, 1))
    OlografosInt = Result
End Function

Private Function OloTriada(groupValue As Integer, Gender As GenderEnum) As String
    ' TODO verify groupValue is 0..999
    OloTriada = Join(Ekatontades(groupValue, Gender), DekadesMonades(groupValue, Gender))
End Function

Private Function Ekatontades(groupValue As Integer, Gender As GenderEnum) As String
    Dim value As Integer
    Dim temp As String
    Dim genderSuffix As String
    value = groupValue \ 100
    If value >= 1 And value <= 9 Then
        temp = N1(3, value)
        If value >= 2 Then
            If Gender = Masculin Then
                genderSuffix = "οι" ' διακόσιοι
            ElseIf Gender = Feminin Then
                genderSuffix = "ες" ' διακόσιες
            Else
                genderSuffix = "α"  ' διακόσια
            End If
        Else
            genderSuffix = "" ' εκατό
        End If
        Ekatontades = temp & genderSuffix
    Else
        Ekatontades = ""
    End If
End Function

Private Function DekadesMonades(groupValue As Integer, Gender As GenderEnum) As String
    Dim dekades As Integer
    Dim monades As Integer
    Dim temp As String
    
    dekades = (groupValue Mod 100) \ 10
    monades = groupValue Mod 10
    If dekades = 1 And monades = 1 Then
        DekadesMonades = "έντεκα"
    ElseIf dekades = 1 And monades = 2 Then
        DekadesMonades = "δώδεκα"
    ElseIf dekades = 0 And monades = 1 Then
        Select Case Gender
        Case Masculin
            DekadesMonades = "ένας"
        Case Feminin
            DekadesMonades = "μία"
        Case Neutral
            DekadesMonades = "ένα"
        End Select
    Else
        temp = ""
        If monades >= 1 And monades <= 9 Then
            temp = N1(1, monades)
            If monades = 3 Then
                Select Case Gender
                Case Neutral
                    temp = temp & "α"   ' τρία
                Case Else
                    temp = temp & "εις" ' τρεις
                End Select
            ElseIf monades = 4 Then
                Select Case Gender
                Case Neutral
                    temp = temp & "α"   ' τέσσερα
                Case Else
                    temp = temp & "ις" ' τέσσερις
                End Select
            End If
        End If
        If dekades >= 1 And dekades <= 9 Then
            If dekades = 1 Then
                temp = N1(2, dekades) & temp ' δεκαπέντε, χωρίς κενό
            Else
                temp = Join(N1(2, dekades), temp) ' είκοσι πέντε, με κενό
            End If
        End If
        DekadesMonades = temp
    End If
End Function

Private Sub InitNames()
    ' TODO initialize only once
    N1(1, 1) = "ένα"
    N1(1, 2) = "δύο"
    N1(1, 3) = "τρ"
    N1(1, 4) = "τέσσερ"
    N1(1, 5) = "πέντε"
    N1(1, 6) = "έξι"
    N1(1, 7) = "επτά"
    N1(1, 8) = "οχτώ"
    N1(1, 9) = "εννιά"
    N1(2, 1) = "δέκα"
    N1(2, 2) = "είκοσι"
    N1(2, 3) = "τριάντα"
    N1(2, 4) = "σαράντα"
    N1(2, 5) = "πενήντα"
    N1(2, 6) = "εξήντα"
    N1(2, 7) = "εβδομήντα"
    N1(2, 8) = "ογδόντα"
    N1(2, 9) = "ενενήντα"
    N1(3, 1) = "εκατό"
    N1(3, 2) = "διακόσι"
    N1(3, 3) = "τριακόσι"
    N1(3, 4) = "τετρακόσι"
    N1(3, 5) = "πεντακόσι"
    N1(3, 6) = "εξακόσι"
    N1(3, 7) = "επτακόσι"
    N1(3, 8) = "οχτακόσι"
    N1(3, 9) = "εννιακόσι"

    N2(1) = "χιλιάδες"
    N2(2) = "εκατομμύρι"
    N2(3) = "δισεκατομμύρι"
    N2(4) = "τρισεκατομμύρι"
End Sub

Private Function Join(First As String, Second As String) As String
    If Len(First) > 0 Then
        If Len(Second) > 0 Then
            Join = First & " " & Second
        Else
            ' second string is empty
            Join = First
        End If
    Else
        ' first string is empty
        Join = Second
    End If
End Function
