' ========================================
' GLOSSA CODE INDENTER - ΕΣΟΧΕΣ ΚΩΔΙΚΑ
' ========================================

' Ρυθμίσεις εσοχής
Const INDENT_SPACES As Integer = 3          ' Πόσα spaces ανά επίπεδο
Const KEEP_SINGLE_EMPTY_LINES As Boolean = True ' Κράτα μόνο 1 κενή γραμμή

' Global μεταβλητές για δομικά keywords
Global INDENT_OPENERS As Variant           ' Keywords που ανοίγουν block
Global INDENT_CLOSERS As Variant           ' Keywords που κλείνουν block  
Global INDENT_MIDDLE As Variant            ' Keywords μεσαίας θέσης (ΑΛΛΙΩΣ)
Global VARIABLE_TYPES As Variant           ' Τύποι μεταβλητών

' ========================================
' ΑΡΧΙΚΟΠΟΙΗΣΗ ΔΟΜΙΚΩΝ KEYWORDS
' ========================================

Sub InitializeIndentKeywords()
    ' Keywords που ανοίγουν block (αυξάνουν εσοχή)
    INDENT_OPENERS = Array( _
        "ΠΡΟΓΡΑΜΜΑ", "ΑΡΧΗ", _
        "ΑΝ", "ΑΛΛΙΩΣ", "ΑΛΛΙΩΣ_ΑΝ", _
        "ΓΙΑ", "ΟΣΟ", "ΑΡΧΗ_ΕΠΑΝΑΛΗΨΗΣ", _
        "ΕΠΙΛΕΞΕ", "ΠΕΡΙΠΤΩΣΗ", _
        "ΣΥΝΑΡΤΗΣΗ", "ΔΙΑΔΙΚΑΣΙΑ" _
    )
    
    ' Keywords που κλείνουν block (μειώνουν εσοχή)
    INDENT_CLOSERS = Array( _
        "ΤΕΛΟΣ_ΠΡΟΓΡΑΜΜΑΤΟΣ", _
        "ΤΕΛΟΣ_ΑΝ", _
        "ΤΕΛΟΣ_ΕΠΑΝΑΛΗΨΗΣ", "ΜΕΧΡΙΣ_ΟΤΟΥ", _
        "ΤΕΛΟΣ_ΕΠΙΛΟΓΩΝ", _
        "ΤΕΛΟΣ_ΣΥΝΑΡΤΗΣΗΣ", "ΤΕΛΟΣ_ΔΙΑΔΙΚΑΣΙΑΣ" _
    )
    
    ' Keywords μεσαίας θέσης (ίδιο επίπεδο με opener, αλλά ανοίγουν νέο block)
    INDENT_MIDDLE = Array( _
        "ΑΛΛΙΩΣ", "ΑΛΛΙΩΣ_ΑΝ", _
        "ΠΕΡΙΠΤΩΣΗ" _
    )
    
    ' Τύποι μεταβλητών (indent μετά το ΜΕΤΑΒΛΗΤΕΣ)
    VARIABLE_TYPES = Array( _
        "ΑΚΕΡΑΙΕΣ:", "ΧΑΡΑΚΤΗΡΕΣ:", "ΠΡΑΓΜΑΤΙΚΕΣ:", "ΛΟΓΙΚΕΣ:" _
    )
End Sub

' ========================================
' ΚΥΡΙΑ ΣΥΝΑΡΤΗΣΗ ΕΣΟΧΩΝ
' ========================================

Sub IndentGlossaCode()
    ' Εσοχές κώδικα ΓΛΩΣΣΑ σε επιλεγμένα text boxes
    Dim slide As slide
    Dim shape As shape
    Dim txtRange As TextRange
    
    ' Αρχικοποίηση keywords
    Call InitializeIndentKeywords
    
    ' Έλεγχος επιλογής
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Παρακαλώ επιλέξτε ένα ή περισσότερα πλαίσια κειμένου", vbExclamation
        Exit Sub
    End If
    
    ' Επεξεργασία κάθε επιλεγμένου shape
    For Each shape In ActiveWindow.Selection.ShapeRange
        If shape.HasTextFrame And shape.TextFrame.HasText Then
            Set txtRange = shape.TextFrame.TextRange
            
            ' Εφαρμογή εσοχών
            Call ProcessIndentation(txtRange)
        End If
    Next shape
    
    MsgBox "Οι εσοχές ολοκληρώθηκαν!", vbInformation
End Sub

' ========================================
' ΕΠΕΞΕΡΓΑΣΙΑ ΕΣΟΧΩΝ
' ========================================

Private Sub ProcessIndentation(txtRange As TextRange)
    Dim originalText As String
    Dim lines As Variant
    Dim processedLines() As String
    Dim i As Long
    Dim currentIndent As Integer
    Dim inVariablesSection As Boolean
    
    ' Διάβασε το κείμενο
    originalText = txtRange.Text
    lines = Split(originalText, vbCr)
    
    ' Προετοιμασία array για επεξεργασμένες γραμμές
    ReDim processedLines(0 To UBound(lines))
    
    ' Αρχικές τιμές
    currentIndent = 0
    inVariablesSection = False
    
    ' Επεξεργασία κάθε γραμμής
    For i = 0 To UBound(lines)
        Dim cleanLine As String
        Dim trimmedLine As String
        
        cleanLine = lines(i)
        trimmedLine = Trim(cleanLine)
        
        ' Αγνόησε εντελώς κενές γραμμές - θα τις διαχειριστούμε μετά
        If Len(trimmedLine) = 0 Then
            processedLines(i) = ""
            GoTo NextLine
        End If
        
        ' Έλεγχος για variables section
        If UCase(trimmedLine) = "ΜΕΤΑΒΛΗΤΕΣ" Then
            inVariablesSection = True
        ElseIf IsLineKeyword(trimmedLine, INDENT_CLOSERS) Or _
               (IsLineKeyword(trimmedLine, INDENT_OPENERS) And UCase(trimmedLine) <> "ΜΕΤΑΒΛΗΤΕΣ") Then
            inVariablesSection = False
        End If
        
        ' Υπολογισμός εσοχής για αυτή τη γραμμή
        Dim lineIndent As Integer
        lineIndent = CalculateLineIndent(trimmedLine, currentIndent, inVariablesSection)
        
        ' Εφαρμογή εσοχής
        processedLines(i) = String(lineIndent * INDENT_SPACES, " ") & trimmedLine
        
        ' Ενημέρωση currentIndent για την επόμενη γραμμή
        currentIndent = CalculateNextIndent(trimmedLine, currentIndent, inVariablesSection)
        
NextLine:
    Next i
    
    ' Διαχείριση κενών γραμμών και ανακατασκευή κειμένου
    Dim finalText As String
    finalText = CleanupEmptyLines(processedLines)
    
    ' Αντικατάσταση κειμένου
    txtRange.Text = finalText
End Sub

' ========================================
' ΥΠΟΛΟΓΙΣΜΟΣ ΕΣΟΧΗΣ ΓΡΑΜΜΗΣ
' ========================================

Private Function CalculateLineIndent(line As String, currentIndent As Integer, inVariablesSection As Boolean) As Integer
    Dim upperLine As String
    upperLine = UCase(Trim(line))
    
    ' Έλεγχος για closers - μειώνουν την εσοχή ΠΡΙΝ τη γραμμή τους
    If IsLineKeyword(upperLine, INDENT_CLOSERS) Then
        CalculateLineIndent = IIf(currentIndent > 0, currentIndent - 1, 0)
        Exit Function
    End If
    
    ' Έλεγχος για middle keywords - ίδιο επίπεδο με το opener
    If IsLineKeyword(upperLine, INDENT_MIDDLE) Then
        CalculateLineIndent = IIf(currentIndent > 0, currentIndent - 1, 0)
        Exit Function
    End If
    
    ' Αν είμαστε στο variables section και είναι τύπος μεταβλητής
    If inVariablesSection And IsLineKeyword(upperLine, VARIABLE_TYPES) Then
        CalculateLineIndent = currentIndent + 1
        Exit Function
    End If
    
    ' Κανονική εσοχή
    CalculateLineIndent = currentIndent
End Function

' ========================================
' ΥΠΟΛΟΓΙΣΜΟΣ ΕΠΟΜΕΝΗΣ ΕΣΟΧΗΣ
' ========================================

Private Function CalculateNextIndent(line As String, currentIndent As Integer, inVariablesSection As Boolean) As Integer
    Dim upperLine As String
    upperLine = UCase(Trim(line))
    
    ' Έλεγχος για closers - μειώνουν την εσοχή
    If IsLineKeyword(upperLine, INDENT_CLOSERS) Then
        CalculateNextIndent = IIf(currentIndent > 0, currentIndent - 1, 0)
        Exit Function
    End If
    
    ' Έλεγχος για openers - αυξάνουν την εσοχή
    If IsLineKeyword(upperLine, INDENT_OPENERS) Then
        CalculateNextIndent = currentIndent + 1
        Exit Function
    End If
    
    ' Έλεγχος για middle keywords - διατηρούν την εσοχή
    If IsLineKeyword(upperLine, INDENT_MIDDLE) Then
        CalculateNextIndent = currentIndent
        Exit Function
    End If
    
    ' Κανονική περίπτωση
    CalculateNextIndent = currentIndent
End Function

' ========================================
' HELPER FUNCTIONS
' ========================================

' Έλεγχος αν γραμμή ξεκινά με keyword από λίστα
Private Function IsLineKeyword(line As String, keywordList As Variant) As Boolean
    Dim i As Integer
    Dim upperLine As String
    upperLine = UCase(Trim(line))
    
    For i = 0 To UBound(keywordList)
        If Left(upperLine, Len(keywordList(i))) = UCase(keywordList(i)) Then
            ' Έλεγχος ότι είναι ολόκληρη λέξη (word boundary)
            Dim nextCharPos As Integer
            nextCharPos = Len(keywordList(i)) + 1
            
            If nextCharPos > Len(upperLine) Then
                ' Τέλος γραμμής
                IsLineKeyword = True
                Exit Function
            Else
                ' Έλεγχος word boundary
                Dim nextChar As String
                nextChar = Mid(upperLine, nextCharPos, 1)
                If nextChar = " " Or nextChar = vbTab Or nextChar = ":" Or nextChar = "_" Then
                    IsLineKeyword = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    IsLineKeyword = False
End Function

' Καθαρισμός κενών γραμμών (κράτα μόνο 1 συνεχόμενη)
Private Function CleanupEmptyLines(lines() As String) As String
    Dim result As String
    Dim i As Long
    Dim lastWasEmpty As Boolean
    
    lastWasEmpty = False
    
    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then
            ' Κενή γραμμή
            If Not lastWasEmpty And KEEP_SINGLE_EMPTY_LINES Then
                result = result & vbCr
                lastWasEmpty = True
            End If
        Else
            ' Μη κενή γραμμή
            If i > 0 Then result = result & vbCr
            result = result & lines(i)
            lastWasEmpty = False
        End If
    Next i
    
    CleanupEmptyLines = result
End Function