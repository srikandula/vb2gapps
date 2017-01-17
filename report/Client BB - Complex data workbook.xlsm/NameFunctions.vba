Attribute VB_Name = "NameFunctions"
Public Function WorkbookName(Branch As String, mp1 As Double, mp2 As Double, sourceRte1 As String, sourceRte2 As String) As String
    If Branch = "" Then
        WorkbookName = ""
        Exit Function
    End If
    Dim suffix As String
    
    If Left(Right(Branch, FindInStr(Branch, "_") + 1), 1) = "R" Then
        Branch = Left(Branch, FindInStr(Branch, "_") - 1)
    End If
    If sourceRte1 <> "" Then
        If sourceRte2 <> "" Then
            suffix = "_R" & sourceRte1 & "_" & sourceRte2
        Else
            suffix = "_R" & sourceRte1
        End If
    Else
        If sourceRte2 <> "" Then
            suffix = "_R" & sourceRte2
        End If
    End If
    If Left(Branch, 1) = "U" Then
        WorkbookName = Branch & "_" & Format(Now, "DDMMMYYYY") & suffix & ".xlsm"
    Else
        WorkbookName = Branch & "_MP" & Format(mp1, "0.0000") & "-" & Format(mp2, "0.0000") & "_" & Format(Now, "DDMMMYYYY") & suffix & ".xlsm"
    End If
End Function
Public Function UnnamedBranch(branchType As String, srcRte As String) As String
    Dim rteStr As String
    Dim sufRte As String
    If branchType = "" Then
        If srcRte = "" Then
            UnnamedBranch = ""
            Exit Function
            
        Else
            branchType = "DREG"
        End If
    End If
    'does not handle multiple source Routes appended to the suffix
    'if the srcRte is Unnamed, handle with care
    'if the srcrte is blank, nothing
    'if the srtrte is not blank, not unnamed then add "_R"
    'if unnamed srcrte, check for its srcrte suffix and pull it off
    '   then name it as srcrte + u + new timestamp + the og suffix
    rteStr = ""
    If Left(srcRte, 2) = "U_" Then
        sufRte = Right(srcRte, Len(srcRte) - FindInStr(srcRte, "_"))
        If Left(sufRte, 1) = "R" Then
            srcRte = Left(srcRte, FindInStr(srcRte, "_") - 1)
            UnnamedBranch = srcRte & "_U" & TimeStamp & "_" & sufRte
            Exit Function
        End If
        UnnamedBranch = srcRte & "_U" & TimeStamp
    Else
        If Not srcRte = "" Then
            rteStr = "_R" & srcRte
        End If
        UnnamedBranch = "U_" & branchType & "_" & TimeStamp & rteStr
    End If
End Function
Private Function TimeStamp() As String
    TimeStamp = Format(Now(), "YYYYMMDDHHNN")
End Function
Private Function FindInStr(FindIn As String, ToFind As String) As Integer
    'searches for a character starting from the end of a string
    Dim FindCha As Integer
    For FindCha = Len(FindIn) - Len(ToFind) + 1 To 1 Step -1
        If Mid(FindIn, FindCha, Len(ToFind)) = ToFind Then
            FindInStr = FindCha
            Exit Function
        End If
    Next FindCha
End Function


