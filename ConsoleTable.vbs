Option Explicit

Class ConsoleTable  'Defining the Class

    Private VerticalMark
    Private HorizontalMark 
    Private objDic
    Private strHeads
    Private strWidths
    Private intSpacingMode
    Private ShowInnerBorders
    
    Private Sub Class_Initialize(  )
	    'Initalization code goes here
        VerticalMark   = "|"
        HorizontalMark = "-"
        intSpacingMode = 1 ' 0=left 1=center 2=rigth
        ShowInnerBorders = True
        Set objDic = WScript.CreateObject("Scripting.Dictionary")
    End Sub

    Private Function GetTableWidth()    
        Dim TableCells
            TableCells = Split(strWidths,",")
        Dim intTableWidth
            intTableWidth = 1
        If UBound(TableCells) > 0 Then
            Dim intCol
            For intCol = 0 to UBound(TableCells)
                intTableWidth = intTableWidth + Len(TableCells(intCol)) + 1
            Next
        End If
        GetTableWidth = intTableWidth
    End Function

    Sub SetHeaders(strHeaders)
        strHeads = strHeaders
        UpdateWidths(strHeaders)
    End Sub

    Sub SetSpacingMode(intMode)
        Select Case CInt(intMode)
            Case 1
                intSpacingMode = 1
            Case 2
                intSpacingMode = 2
            Case Else
                intSpacingMode = 0
        End Select
    End Sub

    Sub DisableBorders
        ShowInnerBorders = False
    End Sub

    Private Sub UpdateWidths(strNewRowContent)      
        Dim NewRowArray
            NewRowArray = Split(strNewRowContent,",")

        Dim WidthsArray 
            WidthsArray = Split(strWidths, ",")

        Dim intCounter, intNewWidth, intWidth

        If UBound(WidthsArray) > 0 Then
            If UBound(NewRowArray) => UBound(WidthsArray) Then
                For intCounter = 0 to UBound(WidthsArray)
                    intWidth = Len(WidthsArray(intCounter))
                    intNewWidth = Len(NewRowArray(intCounter))
                    If intNewWidth > intWidth Then
                        intWidth = intNewWidth
                    End If
                    Select Case intCounter
                        Case 0
                            strWidths = String(intWidth,"X") & ","
                        Case UBound(WidthsArray)
                            strWidths = strWidths & String(intWidth,"X")
                        Case Else
                            strWidths = strWidths & String(intWidth,"X") & ","
                    End Select
                Next
                
            End if
        Else    ' The first time
            For intCounter = 0 to UBound(NewRowArray)
                intWidth = Len(NewRowArray(intCounter))
                Select Case intCounter
                    Case 0
                        strWidths = String(intWidth,"X") & ","
                    Case UBound(NewRowArray)
                        strWidths = strWidths & String(intWidth,"X")
                    Case Else
                        strWidths = strWidths & String(intWidth,"X") & ","
                End Select
            Next
        End If
        

    End Sub

    Sub AddRow(strRowContent)
        objDic.Add objDic.Count, Trim( strRowContent )
        UpdateWidths(strRowContent)
    End Sub

    Private Function GetVerticalBorder
        If ShowInnerBorders Then
            GetVerticalBorder = VerticalMark
        Else
            GetVerticalBorder = " "
        End If
    End Function

    Private Sub PrintTableRow(strRowContent)
        Dim strTemp
            strTemp = GetVerticalBorder 'Start of Row
        ' Retrieve columns' widths
        Dim strWidthArray
            strWidthArray = Split(strWidths, ",")
        ' Parse content of each cell
        Dim strContentArray
            strContentArray = Split(strRowContent, ",")
        
        Dim intCol, intSpaces, intLSpaces, intRSpaces, strCellContent
        If UBound(strWidthArray) > 0 Then
            'Loop through the width's array
            For intCol = 0 to UBound(strWidthArray)
                ' Get the cell content
                If intCol > UBound(strContentArray) Then
                    strCellContent = ""
                Else
                    strCellContent = strContentArray(intCol)
                End If
                ' calculate spacing inside each cell
                intSpaces = Len(strWidthArray(intCol)) - Len(strCellContent)
                Select Case intSpacingMode
                    Case 0 ' Left
                        strTemp = strTemp & strCellContent
                        strTemp = strTemp & Space(intSpaces) & GetVerticalBorder
                    Case 2 ' Rigth
                        strTemp = strTemp & Space(intSpaces)
                        strTemp = strTemp & strCellContent & GetVerticalBorder
                    Case Else ' Center
                        intLSpaces = Int(intSpaces/2)
                        intRSpaces = intSpaces - intLSpaces
                        strTemp = strTemp & Space(intLSpaces) & strCellContent
                        strTemp = strTemp & Space(intRSpaces) & GetVerticalBorder
                End Select
            Next
        End If
        ' End of Row
        WScript.Echo(strTemp)   ' Write Row
    End Sub

    Sub Write 
        Dim Items
            Items = objDic.Items
        
        'Print start of table
        WScript.Echo(String(GetTableWidth,HorizontalMark))

        If Len(strHeads) > 0 Then                
            'Print headers
            PrintTableRow(strHeads)

            'Print separator
            WScript.Echo(String(GetTableWidth,HorizontalMark))
        End If

        If UBound(Items) > 0 Then
            'Print rows
            Dim intRow
            For intRow = 0 to UBound(Items)
                PrintTableRow(Items(intRow))        
            Next
        End If

        'Print end of table
        WScript.Echo(String(GetTableWidth,HorizontalMark))
    End Sub

    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
        Set objDic = Nothing
    End Sub

End Class

' Instantiation of the Class
Dim objTable
Set objTable = New ConsoleTable

With objTable
    .SetHeaders("Numero,Pais,Fondos")
    .DisableBorders
    .AddRow("1,Cuba,1000")
    .AddRow("2,Rusia,10000")
    .AddRow("3,Estados Unidos,1000000")
    .AddRow("4,Brasil")
    .AddRow("5,China, 9999999999999")
    .AddRow("6, EU , 7777777, qwerty")
    .AddRow("7, Portugal , 333x7777777, qwerty")
    .Write
End With
