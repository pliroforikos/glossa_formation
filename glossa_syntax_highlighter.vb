' ========================================
' ΠΑΡΑΜΕΤΡΟΙ ΜΟΡΦΟΠΟΙΗΣΗΣ - ΑΛΛΑΞΤΕ ΕΔΩ
' ========================================

' Γενικές ρυθμίσεις
Const FONT_NAME As String = "Courier New"
Const DEFAULT_FONT_SIZE As Integer = 0  ' 0 = διατήρηση υπάρχοντος μεγέθους

' ΛΕΞΕΙΣ-ΚΛΕΙΔΙΑ - Προσθέστε ή αφαιρέστε όσες θέλετε
Global GLOSSA_KEYWORDS As Variant

' ΤΕΛΕΣΤΕΣ - Προσθέστε ή αφαιρέστε όσους θέλετε
Global GLOSSA_OPERATORS As Variant

' ΕΙΔΙΚΕΣ ΣΥΝΑΡΤΗΣΕΙΣ - Προσθέστε ή αφαιρέστε όσες θέλετε
Global GLOSSA_FUNCTIONS As Variant

' Χρώματα (σε RGB format - χρησιμοποιήστε RGB(R, G, B) για υπολογισμό)
' Παραδείγματα: RGB(255,0,0)=κόκκινο, RGB(0,255,0)=πράσινο, RGB(0,0,255)=μπλε
Const COLOR_KEYWORDS_R As Integer = 0      ' Κόκκινο component για keywords
Const COLOR_KEYWORDS_G As Integer = 0      ' Πράσινο component για keywords
Const COLOR_KEYWORDS_B As Integer = 255    ' Μπλε component για keywords (μπλε)

Const COLOR_STRINGS_R As Integer = 163     ' Κόκκινο σκούρο για strings
Const COLOR_STRINGS_G As Integer = 21
Const COLOR_STRINGS_B As Integer = 21

Const COLOR_COMMENTS_R As Integer = 0      ' Πράσινο για comments
Const COLOR_COMMENTS_G As Integer = 128
Const COLOR_COMMENTS_B As Integer = 0

Const COLOR_OPERATORS_R As Integer = 128   ' Μωβ για operators
Const COLOR_OPERATORS_G As Integer = 0
Const COLOR_OPERATORS_B As Integer = 128

Const COLOR_FUNCTIONS_R As Integer = 255   ' Πορτοκαλί για functions
Const COLOR_FUNCTIONS_G As Integer = 140
Const COLOR_FUNCTIONS_B As Integer = 0

' Στυλ γραμματοσειράς
Const KEYWORDS_BOLD As Boolean = True
Const KEYWORDS_ITALIC As Boolean = False

Const STRINGS_BOLD As Boolean = False
Const STRINGS_ITALIC As Boolean = False

Const COMMENTS_BOLD As Boolean = False
Const COMMENTS_ITALIC As Boolean = True

Const OPERATORS_BOLD As Boolean = False
Const OPERATORS_ITALIC As Boolean = False

Const FUNCTIONS_BOLD As Boolean = False
Const FUNCTIONS_ITALIC As Boolean = False

' ========================================
' ΑΡΧΙΚΟΠΟΙΗΣΗ ΛΕΞΕΩΝ-ΚΛΕΙΔΙΩΝ
' ========================================

Sub InitializeKeywords()
    ' Καλέστε αυτή τη συνάρτηση για να αρχικοποιήσετε τις λέξεις-κλειδιά
    GLOSSA_KEYWORDS = Array( _
        "ΠΡΟΓΡΑΜΜΑ", "ΑΡΧΗ", "ΤΕΛΟΣ_ΠΡΟΓΡΑΜΜΑΤΟΣ", _
        "ΑΝ", "ΤΕΛΟΣ_ΑΝ", "ΑΛΛΙΩΣ", "ΑΛΛΙΩΣ_ΑΝ", "ΤΟΤΕ", _
        "ΓΙΑ", "ΑΠΟ", "ΜΕΧΡΙ", "ΜΕ", "ΒΗΜΑ", "ΤΕΛΟΣ_ΕΠΑΝΑΛΗΨΗΣ", _
        "ΟΣΟ", "ΕΠΑΝΑΛΑΒΕ", "ΑΡΧΗ_ΕΠΑΝΑΛΗΨΗΣ", "ΜΕΧΡΙΣ_ΟΤΟΥ", _
        "ΕΠΙΛΕΞΕ", "ΤΕΛΟΣ_ΕΠΙΛΟΓΩΝ", "ΠΕΡΙΠΤΩΣΗ", _
        "ΣΥΝΑΡΤΗΣΗ", "ΤΕΛΟΣ_ΣΥΝΑΡΤΗΣΗΣ", _
        "ΔΙΑΔΙΚΑΣΙΑ", "ΤΕΛΟΣ_ΔΙΑΔΙΚΑΣΙΑΣ", _
        "ΜΕΤΑΒΛΗΤΕΣ", "ΑΚΕΡΑΙΕΣ", "ΧΑΡΑΚΤΗΡΕΣ", "ΛΟΓΙΚΕΣ", "ΠΡΑΓΜΑΤΙΚΕΣ", _
        "ΔΙΑΒΑΣΕ", "ΓΡΑΨΕ" _
    )
End Sub

' ========================================
' ΑΡΧΙΚΟΠΟΙΗΣΗ ΤΕΛΕΣΤΩΝ
' ========================================

Sub InitializeOperators()
    ' Προσοχή: Η σειρά έχει σημασία! Βάλτε τους μεγαλύτερους πρώτα
    ' για να αποφύγετε conflicts (π.χ. "<--" πριν από "<")
    GLOSSA_OPERATORS = Array( _
        "<--", "<=", ">=", "<>", _
        "<", ">", "=" _
    )
End Sub

' ========================================
' ΑΡΧΙΚΟΠΟΙΗΣΗ ΕΙΔΙΚΩΝ ΣΥΝΑΡΤΗΣΕΩΝ
' ========================================

Sub InitializeFunctions()
    ' Μόνο το όνομα της συνάρτησης - οι παρενθέσεις ελέγχονται αυτόματα
    GLOSSA_FUNCTIONS = Array( _
        "Α_Τ", "Α_Μ", "Τ_Ρ" _
    )
End Sub

' ========================================
' ΚΥΡΙΟΣ ΚΩΔΙΚΑΣ - ΜΗΝ ΑΛΛΑΞΕΤΕ ΠΑΡΑΚΑΤΩ
' ========================================

Sub FormatGlossaCode()
    ' Μορφοποίηση κώδικα ΓΛΩΣΣΑ σε επιλεγμένα text boxes
    Dim slide As slide
    Dim shape As shape
    Dim txtRange As TextRange
    
    ' Αρχικοποίηση όλων των λιστών
    Call InitializeKeywords
    Call InitializeOperators
    Call InitializeFunctions
    
    ' Έλεγχος αν υπάρχει επιλογή
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Παρακαλώ επιλέξτε ένα ή περισσότερα πλαίσια κειμένου", vbExclamation
        Exit Sub
    End If
    
    ' Επεξεργασία κάθε επιλεγμένου shape
    For Each shape In ActiveWindow.Selection.ShapeRange
        If shape.HasTextFrame And shape.TextFrame.HasText Then
            Set txtRange = shape.TextFrame.TextRange
            
            ' 1. Εφαρμογή γενικών ρυθμίσεων γραμματοσειράς
            txtRange.Font.Name = FONT_NAME
            If DEFAULT_FONT_SIZE > 0 Then
                txtRange.Font.Size = DEFAULT_FONT_SIZE
            End If
            
            ' 2. Μορφοποίηση με σειρά προτεραιότητας
            Call FormatStrings(txtRange)
            Call FormatComments(txtRange)
            Call FormatKeywords(txtRange)
            Call FormatOperators(txtRange)
            Call FormatSpecialFunctions(txtRange)
        End If
    Next shape
    
    MsgBox "Η μορφοποίηση ολοκληρώθηκε!", vbInformation
End Sub

' 1. Μορφοποίηση συμβολοσειρών (strings)
Private Sub FormatStrings(txtRange As TextRange)
    Dim i As Long
    Dim inString As Boolean
    Dim startPos As Long
    
    inString = False
    For i = 1 To txtRange.Length
        If txtRange.Characters(i, 1).Text = "'" Then
            If Not inString Then
                startPos = i
                inString = True
            Else
                ' Τέλος string - μορφοποίηση
                With txtRange.Characters(startPos, i - startPos + 1).Font
                    .color.RGB = RGB(COLOR_STRINGS_R, COLOR_STRINGS_G, COLOR_STRINGS_B)
                    .Bold = STRINGS_BOLD
                    .Italic = STRINGS_ITALIC
                End With
                inString = False
            End If
        End If
    Next i
End Sub

' 2. Μορφοποίηση σχολίων
Private Sub FormatComments(txtRange As TextRange)
    Dim lines As Variant
    Dim i As Long, j As Long
    Dim currentPos As Long
    Dim lineText As String
    Dim commentPos As Long
    
    lines = Split(txtRange.Text, vbCr)
    currentPos = 1
    
    For i = 0 To UBound(lines)
        lineText = lines(i)
        commentPos = InStr(lineText, "!")
        
        If commentPos > 0 Then
            ' Μορφοποίηση από ! μέχρι τέλος γραμμής
            Dim commentStart As Long
            commentStart = currentPos + commentPos - 1
            Dim commentLength As Long
            commentLength = Len(lineText) - commentPos + 1
            
            With txtRange.Characters(commentStart, commentLength).Font
                .color.RGB = RGB(COLOR_COMMENTS_R, COLOR_COMMENTS_G, COLOR_COMMENTS_B)
                .Bold = COMMENTS_BOLD
                .Italic = COMMENTS_ITALIC
            End With
        End If
        
        currentPos = currentPos + Len(lineText) + 1  ' +1 για το vbCr
    Next i
End Sub

' 3. Μορφοποίηση λέξεων-κλειδιών
Private Sub FormatKeywords(txtRange As TextRange)
    ' Χρησιμοποιεί την global μεταβλητή GLOSSA_KEYWORDS
    Dim keyword As Variant
    For Each keyword In GLOSSA_KEYWORDS
        Call FindAndFormatKeyword(txtRange, CStr(keyword), RGB(COLOR_KEYWORDS_R, COLOR_KEYWORDS_G, COLOR_KEYWORDS_B), KEYWORDS_BOLD, KEYWORDS_ITALIC)
    Next keyword
End Sub

' 4. Μορφοποίηση τελεστών
Private Sub FormatOperators(txtRange As TextRange)
    ' Χρησιμοποιεί την global μεταβλητή GLOSSA_OPERATORS
    Dim operator As Variant
    For Each operator In GLOSSA_OPERATORS
        Call FindAndFormatExact(txtRange, CStr(operator), RGB(COLOR_OPERATORS_R, COLOR_OPERATORS_G, COLOR_OPERATORS_B), OPERATORS_BOLD, OPERATORS_ITALIC)
    Next operator
End Sub

' 5. Μορφοποίηση ειδικών συναρτήσεων
Private Sub FormatSpecialFunctions(txtRange As TextRange)
    ' Χρησιμοποιεί την global μεταβλητή GLOSSA_FUNCTIONS
    Dim pattern As Variant
    For Each pattern In GLOSSA_FUNCTIONS
        Call FindAndFormatFunction(txtRange, CStr(pattern), RGB(COLOR_FUNCTIONS_R, COLOR_FUNCTIONS_G, COLOR_FUNCTIONS_B), FUNCTIONS_BOLD, FUNCTIONS_ITALIC)
    Next pattern
End Sub

' Helper: Εύρεση και μορφοποίηση λέξεων-κλειδιών (case-insensitive, word boundaries)
Private Sub FindAndFormatKeyword(txtRange As TextRange, keyword As String, color As Long, isBold As Boolean, isItalic As Boolean)
    Dim i As Long
    Dim textUpper As String
    Dim keywordUpper As String
    Dim keywordLen As Long
    
    textUpper = UCase(txtRange.Text)
    keywordUpper = UCase(keyword)
    keywordLen = Len(keyword)
    
    i = 1
    Do While i <= Len(textUpper) - keywordLen + 1
        If Mid(textUpper, i, keywordLen) = keywordUpper Then
            ' Έλεγχος word boundaries
            Dim prevChar As String, nextChar As String
            If i > 1 Then
                prevChar = Mid(textUpper, i - 1, 1)
            Else
                prevChar = " "
            End If
            
            If i + keywordLen <= Len(textUpper) Then
                nextChar = Mid(textUpper, i + keywordLen, 1)
            Else
                nextChar = " "
            End If
            
            If IsWordBoundary(prevChar) And IsWordBoundary(nextChar) Then
                ' Έλεγχος αν είναι μέσα σε string ή comment
                If Not IsInsideStringOrComment(txtRange, i) Then
                    With txtRange.Characters(i, keywordLen).Font
                        .color.RGB = color
                        .Bold = isBold
                        .Italic = isItalic
                    End With
                End If
                i = i + keywordLen
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

' Helper: Εύρεση και μορφοποίηση ακριβών matches (για operators)
Private Sub FindAndFormatExact(txtRange As TextRange, searchText As String, color As Long, isBold As Boolean, isItalic As Boolean)
    Dim i As Long
    Dim textContent As String
    Dim searchLen As Long
    
    textContent = txtRange.Text
    searchLen = Len(searchText)
    
    i = 1
    Do While i <= Len(textContent) - searchLen + 1
        If Mid(textContent, i, searchLen) = searchText Then
            If Not IsInsideStringOrComment(txtRange, i) Then
                With txtRange.Characters(i, searchLen).Font
                    .color.RGB = color
                    .Bold = isBold
                    .Italic = isItalic
                End With
            End If
            i = i + searchLen
        Else
            i = i + 1
        End If
    Loop
End Sub

' Helper: Εύρεση συναρτήσεων με παρενθέσεις
Private Sub FindAndFormatFunction(txtRange As TextRange, funcName As String, color As Long, isBold As Boolean, isItalic As Boolean)
    Dim i As Long
    Dim textUpper As String
    Dim funcUpper As String
    Dim funcLen As Long
    
    textUpper = UCase(txtRange.Text)
    funcUpper = UCase(funcName)
    funcLen = Len(funcName)
    
    i = 1
    Do While i <= Len(textUpper) - funcLen + 1
        If Mid(textUpper, i, funcLen) = funcUpper Then
            ' Ψάχνουμε για άνοιγμα παρένθεσης
            Dim j As Long
            j = i + funcLen
            Do While j <= Len(textUpper) And Mid(textUpper, j, 1) = " "
                j = j + 1
            Loop
            
            If j <= Len(textUpper) And Mid(textUpper, j, 1) = "(" Then
                ' Βρίσκουμε το κλείσιμο παρένθεσης
                Dim parenCount As Long
                Dim endPos As Long
                parenCount = 1
                endPos = j + 1
                
                Do While endPos <= Len(textUpper) And parenCount > 0
                    If Mid(textUpper, endPos, 1) = "(" Then
                        parenCount = parenCount + 1
                    ElseIf Mid(textUpper, endPos, 1) = ")" Then
                        parenCount = parenCount - 1
                    End If
                    endPos = endPos + 1
                Loop
                
                If parenCount = 0 Then
                    ' Μορφοποίηση ολόκληρης της συνάρτησης
                    If Not IsInsideStringOrComment(txtRange, i) Then
                        With txtRange.Characters(i, endPos - i).Font
                            .color.RGB = color
                            .Bold = isBold
                            .Italic = isItalic
                        End With
                    End If
                End If
            End If
            i = i + funcLen
        Else
            i = i + 1
        End If
    Loop
End Sub

' Helper: Έλεγχος αν character είναι word boundary
Private Function IsWordBoundary(char As String) As Boolean
    IsWordBoundary = (char = " " Or char = vbTab Or char = vbCr Or char = vbLf Or _
                     char = "(" Or char = ")" Or char = "," Or char = ";" Or _
                     char = "+" Or char = "-" Or char = "*" Or char = "/" Or char = "=" Or _
                     char = "[" Or char = "]" Or char = "{" Or char = "}" Or _
                     char = "<" Or char = ">" Or char = ":" Or char = "." Or char = "!")
End Function

' Helper: Έλεγχος αν θέση είναι μέσα σε string ή comment
Private Function IsInsideStringOrComment(txtRange As TextRange, pos As Long) As Boolean
    Dim i As Long
    Dim inString As Boolean
    Dim lineStart As Long
    Dim char As String
    
    ' Έλεγχος για string
    inString = False
    For i = 1 To pos - 1
        char = txtRange.Characters(i, 1).Text
        If char = "'" Then
            inString = Not inString
        End If
    Next i
    
    If inString Then
        IsInsideStringOrComment = True
        Exit Function
    End If
    
    ' Έλεγχος για comment
    ' Βρίσκουμε την αρχή της γραμμής
    lineStart = pos
    Do While lineStart > 1
        char = txtRange.Characters(lineStart - 1, 1).Text
        If char = vbCr Or char = vbLf Then
            Exit Do
        End If
        lineStart = lineStart - 1
    Loop
    
    ' Ψάχνουμε για ! στη γραμμή πριν από τη θέση μας
    For i = lineStart To pos - 1
        char = txtRange.Characters(i, 1).Text
        If char = "!" Then
            IsInsideStringOrComment = True
            Exit Function
        End If
    Next i
    
    IsInsideStringOrComment = False
End Function

