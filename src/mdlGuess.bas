Attribute VB_Name = "mdlGuess"
'Author: David Nissim
Option Explicit


'===============================MAIN FUNCTIONS==========================
    Function GuessValue(strSearch As String, _
                        arrSearch As Variant, _
                        arrResult As Variant, _
                        Optional arrFilter As Variant) _
            As Guess
            
        'Compares strSearch to all values in arrSearch.
        'Tallies how many matches are seen and the corresponding Result from arrResult
        'Returns a Guess with the number of matches, proportion of matches, and the guessed Result itself

        'DATA PULLING. ARRAY AND COLLECTION INITIALIZATION
        'arrSearch contains values from the Range being searched for a match with strSearch
        'arrResult contains values from the Range providing the corresponding result for a match
        'As matching is done, populate a collection that holds all matching results and an array
        '(with the same order as the collection) order holding the match counts
        Dim collResultList As Collection
        Set collResultList = New Collection

        Dim arrResultMatches() As Long
        ReDim arrResultMatches(1 To 1)

        Dim i As Long   'Index used throughout the function

        'Remove special characters from the search string so it matches the training data
        'Note this function is applied to all the data when RangeToArray is called
        strSearch = RemoveSpecialCharacters(strSearch)

        'Check whether filter array has been provided.  If not generate one that is all True (so no filtering occurs)
        If IsMissing(arrFilter) Then arrFilter = GenerateFilter(arrSearch, "FilterNothing")

        'POPULATE MATCH ARRAY
        'i is the index for each row of training data
        Dim currentResult As String
        Dim currentResultIndex As Long

        For i = LBound(arrSearch, 1) To UBound(arrSearch, 1)
            
            If InStr(1, arrSearch(i, 1), strSearch, vbTextCompare) <> 0 And _
            arrSearch(i, 1) <> vbNullString And _
            arrFilter(i, 1) Then
            
                currentResult = arrResult(i, 1)
                If currentResult <> vbNullString Then
                    
                    '***** Added 1/15/2020 to compile result list during evaluation rather than before, more efficient
                    With FindInCollection(currentResult, collResultList)
                    
                        If (.Found) Then
                            currentResultIndex = .Index
                        Else 'Result isn't in the current result collection
                            collResultList.Add currentResult
                            currentResultIndex = FindInCollection(currentResult, collResultList).Index
                            ReDim Preserve arrResultMatches(1 To collResultList.count)
                        End If
                    
                    End With
                    '*****
                    arrResultMatches(currentResultIndex) = arrResultMatches(currentResultIndex) + 1
                End If
            End If
        Next i
            
        'EVALUATE RESULTS
        Dim strGuess As String
        Dim HighestMatchCount As Long
        Dim HighestProportion As Single

        HighestMatchCount = Application.WorksheetFunction.Max(arrResultMatches)

        If HighestMatchCount > 0 Then
            HighestProportion = HighestMatchCount / Application.WorksheetFunction.Sum(arrResultMatches)
            
            For i = 1 To collResultList.count
                If arrResultMatches(i) = HighestMatchCount Then
                    strGuess = collResultList(i)
                    Exit For
                End If
            Next i
            
            GuessValue.Matches = HighestMatchCount
            GuessValue.strGuess = strGuess
            GuessValue.Proportion = HighestProportion
        End If


        'UNCOMMENT TO PRINT RESULTS
        'outputGuessResults collResultList, arrResultMatches, strSearch, GuessValue

        Set collResultList = Nothing

    End Function


    Function CascadeGuessValue(ByVal strSearch As String, _
                            arrSearch As Variant, _
                            arrResult As Variant, _
                    Optional minMatches As Integer, _
                    Optional minTolerance As Single, _
                    Optional arrFilter As Variant) _
            As Guess

        'This function takes strSearch and breaks it into progressively smaller pieces, until guess value returns an acceptable guess.
        'The number of pieces in strSearch is based delimiting by spaces.
        'If multiple acceptable guesses are found at the same level of splitting,
        ' the code takes the guess with the higher number of matches for the guessed category = [# of Matches that had the same guess answer]x[Proportion Of Matches that had the same guessed answer]
        'If this function fails to guess, it outputs a blank.

        Dim codeGuess As Guess, currentGuess As Guess, lastGuess As Guess
        Dim NumOfSubStrings As Integer, i As Integer
        Dim collStrings As Collection
        Dim subString As String
        Dim CountOfAcceptableGuesses As Integer

        strSearch = RemoveSpecialCharacters(strSearch)  'Replace special characters with spaces

        'Check whether filter array has been provided.  If not generate one that is all True (so no filtering occurs)
        If IsMissing(arrFilter) Then arrFilter = GenerateFilter(arrSearch, "FilterNothing")


        For NumOfSubStrings = 1 To UBound(SplitMultiDelims(strSearch, " ", True)) + 1
            
            CountOfAcceptableGuesses = 0    'Reset the counter
            Set collStrings = CascadeString(strSearch, NumOfSubStrings)  'Create a collection of all cascaded substrings
            
            For i = 1 To collStrings.count 'Don't use for each in this place as it requires substring to be a variant
                subString = collStrings(i)
                
                codeGuess = GuessValue(subString, arrSearch, arrResult, arrFilter)
                
                'Debug.Print subString, codeGuess.strGuess, Format(codeGuess.Proportion, "#%"), codeGuess.Matches
                
                If GuessAcceptable(codeGuess, minTolerance, minMatches) Then
                    CountOfAcceptableGuesses = CountOfAcceptableGuesses + 1
                    
                    'If there are multiple guesses are found that pass the necessary criteria, then pick the best one.
                    'Whichever guess has the highest Matches*Proportion is chosen
                    If CountOfAcceptableGuesses > 1 And codeGuess.strGuess <> currentGuess.strGuess Then
                        If codeGuess.Matches * codeGuess.Proportion > currentGuess.Matches * currentGuess.Proportion Then
                            currentGuess = codeGuess
                        End If
                    Else
                        currentGuess = codeGuess
                    End If
                End If
                
            Next i
            
            Set collStrings = Nothing
            
            'If there is an acceptable guess then output it.
            If currentGuess.strGuess <> vbNullString Then
                CascadeGuessValue = currentGuess
                'Debug.Print "String was split into " & NumOfSubStrings
                Exit Function
            End If
            
        Next NumOfSubStrings

    End Function

    Function GuessAcceptable(inGuess As Guess, _
                            minTolerance As Single, _
                            minMatches As Integer) _
            As Boolean
        'Returns TRUE/FALSE about whether the guess has passed the two tests
        '1. The number of matches exceeds the minimum
        '2. The proportion of matches exceeds the tolerance

        GuessAcceptable = inGuess.Matches >= minMatches And inGuess.Proportion >= minTolerance

    End Function

'===============================APPLICATION OF GUESS==========================
    Function GuessCategory(ByVal strLocation As String, _
                Optional strSource As String, _
                Optional minMatches As Integer = 1, _
                Optional minTolerance As Single = 0.4) _
            As String
            
        'Estimates what the category should be for a transaction
        'Regarding the guessing parameters:
        '   Location needs to show up at least minMatches times before suggesting a category
        '   A category needs to be chosen for a given location at least minTolerance proportion (0 to 1) of the time to be chosen

        Dim codeGuess As Guess
        Dim arrLocation() As Variant, arrCode() As Variant
        Dim arrSource() As Variant
        Dim arrFilter() As Boolean
        Dim tblTrans As ListObject

        Dim FilterSkip As Byte

        strLocation = RemoveSpecialCharacters(strLocation)
        FilterSkip = 0

        Set tblTrans = GetTable("tblTrans")

        arrLocation = RangeToArray(TableColumnRange(tblTrans, "Location"))
        arrCode = RangeToArray(TableColumnRange(tblTrans, "Code"))

        If strSource = vbNullString Then
            FilterSkip = FilterSkip + 1
        Else
            arrSource = RangeToArray(TableColumnRange(tblTrans, "Source"))
        End If


        Set tblTrans = Nothing


        'Check the full input string for an acceptable guess.
        'If this doesn't exist, cascade the string into two smaller ones and repeat
        'If one or both of these guesses are acceptable, then return that.
        'If both guesses are acceptable but different then exit function with no return
        'If neither are acceptable, continue cascading the string into smaller pieces until
        'either an acceptable match is found, different matches are found, or no matches are found.
            
        Dim FilterLevel As Byte
        'First time through, use only training data from that account (SOURCE)
        'If an answer is not found, go through a second time using all the data

        For FilterLevel = 1 + FilterSkip To 2
            Select Case FilterLevel
                Case 1
                    arrFilter = GenerateFilter(arrSource, "FilterByString", strSource)
                Case Else
                    arrFilter = GenerateFilter(arrLocation, "FilterNothing")
            End Select
            
            codeGuess = CascadeGuessValue(strLocation, arrLocation, arrCode, minMatches, minTolerance, arrFilter)
            If codeGuess.strGuess <> vbNullString Then
                GuessCategory = codeGuess.strGuess
                'Debug.Print "Guess: " & codeGuess.strGuess & "; Matches=" & codeGuess.Matches & "; Proportion=" & codeGuess.Proportion
                'Debug.Print "FilterLevel: " & FilterLevel
                Exit Function
            End If

        Next FilterLevel

    End Function

    Function GuessAccount(strFileName As String, _
                Optional minMatches As Integer = 1, _
                Optional minTolerance As Single = 0.9) _
            As String
        'This function checks whether a specific file name has been associated with importing a particular account's data

        Dim accGuess As Guess
        Dim arrSearch(), arrResult()
        Dim tblAccounts As ListObject, tblImportHistory As ListObject

        Set tblAccounts = GetTable("tblAccounts")
        Set tblImportHistory = GetTable("tblImportHistory")


        arrSearch = RangeToArray(TableColumnRange(tblImportHistory, "File Name"))
        arrResult = RangeToArray(TableColumnRange(tblImportHistory, "Account"))

        GuessAccount = CascadeGuessValue(strFileName, arrSearch, arrResult, minMatches, minTolerance).strGuess

    End Function

'===============================AUXILIARY FUNCTIONS==========================

    Function LogicArray(strLogicGate As String, _
                        array1 As Variant, _
            Optional array2 As Variant) _
            As Variant
        'This function applies logic gates to arrays and returns the resulting logic array
        'The returned array is always given with 2 dimensions even if it's nx1 (so it plays nicely with other functions)
        'User provides desired operator as string
        'If the operator is NOT, then only provide one array containing Booleans
        'Otherwise, provide two same size arrays containing Booleans

        Dim arrOut() As Boolean
        Dim i As Long, j As Long

        strLogicGate = UCase(strLogicGate) 'Set input to uppercase for easier comparison

        ReDim arrOut(LBound(array1, 1) To UBound(array1, 1), LBound(array1, 2) To UBound(array1, 2))

        If strLogicGate = "NOT" Then
            For j = LBound(array1, 2) To UBound(array1, 2)
                For i = LBound(array1, 1) To UBound(array1, 1)
                    arrOut(i, j) = Not (array1(i, j))
                Next i
            Next j

        Else
        'Check for 2 arrays with the same dimensions
            If IsMissing(array2) Then
                MsgBox "Two input arrays required for all logic gates other than NOT.", _
                vbOKOnly + vbCritical, "Missing Input"
                
                Exit Function
            ElseIf (UBound(array1, 1) - LBound(array1, 1)) <> (UBound(array2, 1) - LBound(array2, 1)) Or _
                (UBound(array1, 1) - LBound(array1, 2)) <> (UBound(array2, 1) - LBound(array2, 2)) Then
                
                MsgBox "The two arrays provided to the LogicArray function are not the same size.", _
                    vbOKOnly + vbCritical, "Cannot apply logic gate"
                
                Exit Function
            End If
            
            'If both checks pass, create new array
            Select Case strLogicGate
                Case "AND"
                    For j = LBound(array1, 2) To UBound(array1, 2)
                        For i = LBound(array1, 1) To UBound(array1, 1)
                            arrOut(i, j) = array1(i, j) And array2(i, j)
                        Next i
                    Next j
                
                Case "OR"
                    For j = LBound(array1, 2) To UBound(array1, 2)
                        For i = LBound(array1, 1) To UBound(array1, 1)
                            arrOut(i, j) = array1(i, j) Or array2(i, j)
                        Next i
                    Next j
                
                Case "XOR"
                    For j = LBound(array1, 2) To UBound(array1, 2)
                        For i = LBound(array1, 1) To UBound(array1, 1)
                            arrOut(i, j) = array1(i, j) Xor array2(i, j)
                        Next i
                    Next j
                    
                Case Else
                    MsgBox "Invalid logic gate provided to LogicArray function. Function will return the first array provided.", _
                    vbOKOnly + vbExclamation, "Invalid Logic Gate"

                    arrOut = array1
            End Select
        End If

        LogicArray = arrOut

    End Function

'===============================FILTER FUNCTIONS=======================================
    'ABOUT FILTER FUNCTIONS
    'A filter function is used by GenerateFilter to create an array of TRUE/FALSE values
    'that can be used by GuessValue (and maybe others in the future).
    'A filter function must have the input value being evaluated as the first argument
    'A filter function may have a second argument for configuration or whatever, but it's not necessary
    'As written GenerateFilter cannot handle a filter function with more than 2 arguments

    Function GenerateFilter(arrInput As Variant, _
                            strFilterFunctionName As String, _
                Optional varFilterFunctionArgument As Variant) _
            As Variant
        'This function takes a 1D input array and the name of a filter function.
        'It outputs one of the same size containing TRUE or FALSE for each element of the input array.
        'Each element of the input array is evaluated on by the Filter Function

        Dim iElement As Long
        ReDim arrOut(LBound(arrInput, 1) To UBound(arrInput, 1), 1 To 1) As Boolean
        'Note, if you want to write these values to a table for testing, you have to define the second dimension

        'Different function call whether parameter is present or not.
        'If statement outside the loop so it's only evaluated once
        If IsEmpty(varFilterFunctionArgument) Then
            For iElement = LBound(arrInput, 1) To UBound(arrInput, 1)
            
                arrOut(iElement, 1) = Application.Run(strFilterFunctionName, arrInput(iElement, 1))
                
            Next iElement
        Else
            For iElement = LBound(arrInput, 1) To UBound(arrInput, 1)
                arrOut(iElement, 1) = Application.Run(strFilterFunctionName, arrInput(iElement, 1), varFilterFunctionArgument)
            Next iElement
        End If

        GenerateFilter = arrOut

    End Function

    Function FilterNothing(ByVal fInput As Variant) As Boolean
        'Returns true no matter what.  fInput is a dummy parameter so that this can be used by GenerateFilter
        FilterNothing = True
    End Function


    Function FilterByString(ByVal strInput As String, ByVal strTarget As String) As Boolean
        'Returns true if input string matches target string
        'Case insensitive

        FilterByString = LCase(strInput) = LCase(strTarget)

    End Function

    Function FilterByValueSign(ByVal valInput As Single, ByVal strValueSign As String) As Boolean
        'Returns True if valInput's sign is the same as indicated by strValueSign
        'Returns everything as False if an invalid target value is provided

        Select Case strValueSign
            Case "Positive"
                FilterByValueSign = valInput > 0
            Case "Negative"
                FilterByValueSign = valInput < 0
            Case Else
                FilterByValueSign = False
        End Select

    End Function

    Function FilterByCollection(ByVal fInput As Variant, collTarget As Collection) As Boolean
        '# Incomplete
        'Returns true if the input value is contained within the collection

        FilterByCollection = FindInCollection(fInput, collTarget).Found
    End Function
