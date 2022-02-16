Option Explicit

Class ConsoleTable  'Defining the Class

    Private strSeparator'This is a special character to delimit table cells
    Private strHeads    'This specifies the head of each column
    Private strWidths   'This specifies the len of each cell
    Private strSpacing  'Select 1=Left, 2=Center, 3=Rigth
    Private intRow      'Integer dedicated to rows 
    Private intCol      'Integer dedicated to columns
    'Private ContentArray(1) 'This is the table
    Private objDic
      
    
    Private Sub Class_Initialize( )
        strSeparator = "|"
        strHeads     = "n" & strSeparator & "Items"
        strWidths    = "-" & strSeparator & "-----"
        strSpacing   = "1" & strSeparator & "1" 
        Set objDic = WScript.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate(  )
        Set objDic = Nothing
    End Sub

    Public Function GetTableRows 'Max rows value +1 (size of the virtual table)
        GetTableRows = objDic.Count
    End Function

    Public Function GetTableCols 'Max columns value +1 (size of the virtual table)  
        GetTableCols = UBound(Split(strHeads,strSeparator)) + 1
    End Function

    Public Sub SetHeaders(strInput1)
        If UBound(Split(strInput1,strSeparator)) > 0 Then
            strHeads = strInput1
            strWidths = strInput1
            UpdateWidths(strInput1) 
            SetSpacing(1)
        End If
    End Sub

    Public Sub SetSpacing(intInput2)
        Dim strTemp
        For intCol = 0 to GetTableCols - 1
            strTemp = strTemp & intInput2 & strSeparator
        Next
        strSpacing = strTemp
    End Sub


    Public Sub AddRow(strInput2)
        Dim tmpArray1
            tmpArray1 = Split(strInput2,strSeparator)
        Redim Preserve tmpArray1(GetTableCols-1)
        objDic.Add GetTableRows, Join(tmpArray1,strSeparator)
        UpdateWidths(Join(tmpArray1,strSeparator))
    End Sub

    Private Sub UpdateWidths(strInput6)
        Dim tmpArray7
            tmpArray7 = Split(strWidths,strSeparator)
        Dim tmpArray8
            tmpArray8 = Split(strInput6,strSeparator)
        For intCol = 0 to UBound(tmpArray7)
            If intCol > UBound(tmpArray8) Then Exit For
            If Len(tmpArray8(intCol)) > len(tmpArray7(intCol)) Then 
                tmpArray7(intCol) = tmpArray8(intCol)
            End If
        Next
        strWidths = GetHorizontalRules(Join(tmpArray7,strSeparator)) 
    End Sub

    Public Sub Write
        Dim objKey
        PrintRow strHeads                       'Print heads
        PrintRow GetHorizontalRules(strHeads)   'Print horizontal rules
        For Each objKey In objDic               'Explore rows one by one
            PrintRow objDic(objKey)             'Print row
        Next
    End Sub

    Private Function GetHorizontalRules(strInput5)
        Dim tmpArray6
            tmpArray6 = Split(strInput5,strSeparator)
        For intCol = 0 to UBound(tmpArray6)
            tmpArray6(intCol) = String(Len(tmpArray6(intCol)),"-")
        Next
        GetHorizontalRules = Join(tmpArray6,strSeparator)
    End Function

    Private Sub PrintRow(strInput3)
        Dim tmpArray3
            tmpArray3 = Split(strInput3,strSeparator)
        For intCol = 0 to UBound(tmpArray3)
            WScript.StdOut.Write GetSpacedCell(tmpArray3(intCol),intCol) & " "
        Next
        WScript.StdOut.Write vbCrLf
    End Sub

    Private Function GetSpacedCell(strInput4,intColumnNumber)
        Dim intLSpaces
        Dim intRSpaces
        Dim intSpaces
        Dim tmpArray5
            tmpArray5 = Split(strWidths,strSeparator)
            strInput4 = Trim(strInput4)
        intSpaces = Len(tmpArray5(intColumnNumber)) - Len(strInput4)
        intLSpaces = CInt(intSpaces/2)
        intRSpaces = intSpaces - intLSpaces 
        Select Case GetSpacingMode(intColumnNumber)
            Case 1 'Left
            GetSpacedCell = strInput4 & Space(intSpaces)
            Case 2 'Center
            GetSpacedCell = Space(intLSpaces) & strInput4 & Space(intRSpaces)
            Case 3 'Rigth
            GetSpacedCell = Space(intSpaces) & strInput4 
            Case Else
        End Select 
    End Function

    Private Function GetSpacingMode(intInput1)
        Dim tmpArray4
            tmpArray4 = Split(strSpacing,strSeparator)
        If (intInput1 => LBound(tmpArray4)) AND (intInput1 <= UBound(tmpArray4)) Then
            GetSpacingMode = CInt(tmpArray4(intInput1))
        End If
    End Function


    'Public Sub SetCellSpacing(intColumn,intSpacingMode)
        'Validate intSpacingMode
    '    If (CInt(intSpacingMode) > 0) AND (CInt(intSpacingMode) < 4) Then
            'Validate intColumn
    '        If (CInt(intColumn) => 0) AND (CInt(intColumn) < GetTableCols) Then
                'Split spacing
    '            Dim tmpArray
    '                tmpArray = Split(strSpacing,strSeparator)
                'Change value
    '                tmpArray(intColumn) = intSpacingMode
                'Join spacing
    '            strSpacing = Join(tmpArray,strSeparator)
    '        End If
    '    End If
    'End Sub

    
    

End Class

' Instantiation of the Class
Dim objTable
Set objTable = New ConsoleTable

With objTable
    .SetHeaders("Numero|Pais|Fondos")
    .AddRow("1|Cuba|1000")
    .AddRow("2|Rusia|10000")
    .AddRow("3|Estados Unidos|1000000")
    .AddRow("4|Brasil")
    .AddRow("5|China| 9999999999999")
    .AddRow("6| EU | 7777777| qwerty")
    .AddRow("7| Portugal | 333x7777777| qwerty")
    .Write
End With
