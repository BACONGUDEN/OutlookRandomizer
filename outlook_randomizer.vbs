' OutlookRandomizer - Written by BACONGUDEN, github.com/BACONGUDEN

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim signatureFilePath As String
    signatureFilePath = "C:\path\to\signature.html"

    ' Load the signature HTML content
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim signatureFile As Object
    Set signatureFile = fso.OpenTextFile(signatureFilePath, 1)
    Dim signatureContent As String
    signatureContent = signatureFile.ReadAll
    signatureFile.Close

    ' Replace the placeholder with a random greeting
    Dim greetingsFilePath As String
    greetingsFilePath = "C:\path\to\greetings.txt"
    Dim greetingsFile As Object
    Set greetingsFile = fso.OpenTextFile(greetingsFilePath, 1)
    Dim greetings As String
    greetings = greetingsFile.ReadAll
    greetingsFile.Close
    Dim randomGreeting As String
    Randomize
    randomGreeting = Split(greetings, vbCrLf)(Int((UBound(Split(greetings, vbCrLf)) + 1) * Rnd))

    ' Replace the placeholder in the signature with the random greeting
    signatureContent = Replace(signatureContent, "[[RANDOM_GREETING]]", randomGreeting)

    ' Load the quotes content
    Dim quotesFilePath As String
    quotesFilePath = "C:\path\to\quotes.txt"
    Dim quotesFile As Object
    Set quotesFile = fso.OpenTextFile(quotesFilePath, 1)
    Dim quotes As String
    quotes = quotesFile.ReadAll
    quotesFile.Close
    Dim randomQuote As String
    Randomize
    randomQuote = Split(quotes, vbCrLf)(Int((UBound(Split(quotes, vbCrLf)) + 1) * Rnd))

    ' Replace the placeholder in the signature with the random quote
    signatureContent = Replace(signatureContent, "[[RANDOM_QUOTE]]", randomQuote)

    ' Set the modified HTML as the signature
    Item.HTMLBody = Item.HTMLBody & "<br><br>" & signatureContent & vbCrLf
End Sub

