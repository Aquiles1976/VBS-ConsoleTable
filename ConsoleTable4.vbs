Option Explicit
  ' Default window size should be 80x25. Table width should not exceed that number.
  ' If the content of a row is bigger than 80, it should be splitted in more than one row.
Class ConsoleTable  'Defining the Class
    Private intConsoleWidth ' Width of the actual console window
    Private strSeparator'This is a special character to delimit table cells
    Private strHeads    'This specifies the head of each column
    Private strWidths   'This specifies the len of each cell
    Private strSpacing  'Select 1=Left, 2=Center, 3=Rigth
    Private intRow      'Integer dedicated to rows 
    Private intCol      'Integer dedicated to columns
    Private objDic
    Private Sub Class_Initialize( )
        intConsoleWidth = 80
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
    Public Sub SetHeaders(strInput0)
        If UBound(Split(strInput0,strSeparator)) > 0 Then
            strHeads = strInput0
            strWidths = strInput0
            UpdateWidths(strInput0) 
            SetSpacing(1)
        End If
    End Sub
    Public Sub AddRow(strInput1)
        Dim tmpArray1
            tmpArray1 = Split(strInput1,strSeparator)
        Redim Preserve tmpArray1(GetTableCols-1)
        'Format Array to fit inside console
        Dim strNewRow, intIndex
        For intIndex = 0 to UBound(tmpArray1)
            If intIndex = 0 Then
                strNewRow = tmpArray1(intIndex)
            Else
                strNewRow = strNewRow & strSeparator & tmpArray1(intIndex)
            End If
            If Len(strNewRow) > intConsoleWidth then 
                strNewRow = Left(strNewRow,intConsoleWidth-3) + "..."
                Exit For
            End If
        Next
        objDic.Add GetTableRows, strNewRow
        UpdateWidths(strNewRow)
    End Sub
    Public Sub SetSpacing(intInput2)
        Dim strTemp
        For intCol = 0 to GetTableCols - 1
            strTemp = strTemp & intInput2 & strSeparator
        Next
        strSpacing = strTemp
    End Sub
    Private Sub PrintRow(strInput3)
        Dim tmpArray3
            tmpArray3 = Split(strInput3,strSeparator)
        For intCol = 0 to UBound(tmpArray3)
            WScript.StdOut.Write GetSpacedCell(tmpArray3(intCol),intCol) & " "
        Next
        WScript.StdOut.Write vbCrLf
    End Sub
    Private Function GetSpacingMode(intInput4)
        Dim tmpArray4
            tmpArray4 = Split(strSpacing,strSeparator)
        If (intInput4 => LBound(tmpArray4)) AND (intInput4 <= UBound(tmpArray4)) Then
            GetSpacingMode = CInt(tmpArray4(intInput4))
        End If
    End Function
    Private Function GetSpacedCell(strInput5,intColumnNumber)
        Dim intLSpaces
        Dim intSpaces
        Dim tmpArray5
            tmpArray5 = Split(strWidths,strSeparator)
            strInput5 = Trim(strInput5)
        intSpaces  = Len(tmpArray5(intColumnNumber)) - Len(strInput5)
        intLSpaces = CInt(intSpaces/2)
        Select Case GetSpacingMode(intColumnNumber)
            Case 1 'Left
              GetSpacedCell = strInput5 & Space(intSpaces)
            Case 3 'Rigth
              GetSpacedCell = Space(intSpaces) & strInput5 
            Case Else 'Center
              GetSpacedCell = Space(intLSpaces) & strInput5 & Space(intSpaces - intLSpaces)
        End Select 
    End Function
    Private Function GetHorizontalRules(strInput6)
        Dim tmpArray6
            tmpArray6 = Split(strInput6,strSeparator)
        For intCol = 0 to UBound(tmpArray6)
            tmpArray6(intCol) = String(Len(tmpArray6(intCol)),"-")
        Next
        GetHorizontalRules = Join(tmpArray6,strSeparator)
    End Function
    Private Sub UpdateWidths(strInput7)
        Dim tmpArrayWidths 
            tmpArrayWidths = Split(strWidths,strSeparator)
        Dim tmpArray7      
            tmpArray7      = Split(strInput7,strSeparator)
        For intCol = 0 to UBound(tmpArrayWidths)
            If intCol > UBound(tmpArray7) Then Exit For
            If Len(tmpArray7(intCol)) > Len(tmpArrayWidths(intCol)) Then 
                tmpArrayWidths(intCol) = Trim(tmpArray7(intCol))
            End If
        Next
        strWidths = GetHorizontalRules(Join(tmpArrayWidths,strSeparator)) 
    End Sub
    Public Sub Write
        Dim objKey
        PrintRow strHeads                       'Print heads
        PrintRow GetHorizontalRules(strHeads)   'Print horizontal rules
        For Each objKey In objDic               'Explore rows one by one
            PrintRow objDic(objKey)             'Print row
        Next
    End Sub

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
    .AddRow("77| Polonia | El pequeño murcielago hindú volaba errante por los cielos hasta que escuchó una voz que lo llamaba desde lejos.")
    .Write
End With