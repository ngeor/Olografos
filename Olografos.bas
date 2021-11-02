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

Public Function Olografos���(Num As Double) As String
    Olografos��� = Olografos(Num, Feminin) + " ���"
End Function

Public Function Olografos(Num As Double, Gender As GenderEnum) As String
    InitNames
    Dim iNum As Long
    iNum = Int(Num)
    If iNum = 0 Then
        Olografos = "�����"
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
    
    ' � ������ ��������� ��� ��������
    ' ����� -> 0..999
    ' ���   -> 001_000..999_999
    triada = 0
    While temp > 0
        ' get the last 3 digits
        groupValue = temp Mod 1000
        If groupValue > 0 Then
    
            ' determine gender for this group
            If triada = 0 Then
                Gen = Gender
            ElseIf triada = 1 Then ' ��������
                Gen = Feminin
            Else ' ����������� ++
                Gen = Neutral
            End If
            
            If triada = 1 And groupValue = 1 Then
                part = "������"
            Else
                part = OloTriada(groupValue, Gen)
                If triada > 0 Then
                    h = N2(triada)
                    If triada >= 2 Then
                        If groupValue = 1 Then
                            h = h + "�"
                        Else
                            h = h + "�"
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
    OlografosInt = RTrim(Result)
End Function

Private Function OloTriada(groupValue As Integer, Gender As GenderEnum) As String
    Dim b As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim Result As String
    Dim h As String

    Result = ""
    For b = 1 To 3
        Select Case b
        Case 1
            i = groupValue \ 100
        Case 2
            i = (groupValue Mod 100) \ 10
        Case 3
            i = groupValue Mod 10
        End Select

        If i <> 0 Then
            h = N1(4 - b, i)

            Select Case b
                Case 1
                    If i <> 1 Then
                        If Gender = Masculin Then
                            h = h + "��"
                        ElseIf Gender = Feminin Then
                            h = h + "��"
                        Else
                            h = h + "�"
                        End If
                    End If
                Case 2
                    If i = 1 Then
                        i2 = groupValue Mod 10
                        If i2 = 1 Or i2 = 2 Then
                            If i2 = 1 Then
                                h = "������"
                            ElseIf i2 = 2 Then
                                h = "������"
                            End If
                            Result = Result + h
                            Exit For
                        End If
                    End If
                Case 3
                    If i = 1 Then
                        If Gender = Feminin Then
                            h = "���"
                        End If
                    ElseIf i = 3 Then
                        If Gender = Neutral Then
                            h = h + "��"
                        Else
                            h = h + "���"
                        End If
                    ElseIf i = 4 Then
                        If Gender = Neutral Then
                            h = h + "�"
                        Else
                            h = h + "��"
                        End If
                    End If
            End Select

            Result = Result + h
            If b <> 3 Then Result = Result + " "
        End If
    Next b
    OloTriada = Result
End Function

Private Sub InitNames()
    N1(1, 1) = "���"
    N1(1, 2) = "���"
    N1(1, 3) = "��"
    N1(1, 4) = "������"
    N1(1, 5) = "�����"
    N1(1, 6) = "���"
    N1(1, 7) = "����"
    N1(1, 8) = "����"
    N1(1, 9) = "�����"
    N1(2, 1) = "����"
    N1(2, 2) = "������"
    N1(2, 3) = "�������"
    N1(2, 4) = "�������"
    N1(2, 5) = "�������"
    N1(2, 6) = "������"
    N1(2, 7) = "���������"
    N1(2, 8) = "�������"
    N1(2, 9) = "��������"
    N1(3, 1) = "�����"
    N1(3, 2) = "�������"
    N1(3, 3) = "��������"
    N1(3, 4) = "���������"
    N1(3, 5) = "���������"
    N1(3, 6) = "�������"
    N1(3, 7) = "��������"
    N1(3, 8) = "��������"
    N1(3, 9) = "���������"

    N2(1) = "��������"
    N2(2) = "����������"
    N2(3) = "�������������"
    N2(4) = "��������������"
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
