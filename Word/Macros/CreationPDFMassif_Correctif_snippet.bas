'==============================================================================
' CORRECTIF CreationPDFMassif — Win11
'
' IMPORTANT: outputFolder / excelFilePath sont DÉJÀ déclarés en tête de Sub.
' Ne PAS remettre "Dim outputFolder As String" ici
' (erreur: Déclaration existante dans la portée en cours).
'
' 1) En tête de Sub, garder UNE SEULE fois :
'      Dim outputFolder As String
'      Dim excelFilePath As String
'      Dim fd As FileDialog
'
' 2) SUPPRIMER le bloc UserName + chemins UNC Isilon.
'
' 3) Coller le code ci-dessous À LA PLACE (sans nouveaux Dim en double).
'==============================================================================

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Sélectionnez le fichier Excel MACCS (Modèle tableau Maccs.xlsx)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
        .Filters.Add "Tous les fichiers", "*.*"
        On Error Resume Next
        .InitialFileName = Environ$("USERPROFILE") & "\Desktop\"
        On Error GoTo 0
        If .Show <> -1 Then
            MsgBox "Opération annulée : aucun fichier Excel sélectionné.", vbExclamation
            Exit Sub
        End If
        excelFilePath = .SelectedItems(1)
    End With
    Set fd = Nothing
    
    outputFolder = Left$(excelFilePath, InStrRev(excelFilePath, "\"))
    
    If Dir(excelFilePath) = "" Then
        MsgBox "Le fichier sélectionné est introuvable :" & vbCrLf & excelFilePath, vbCritical
        Exit Sub
    End If
    
    ' Ensuite: MsgBox + Workbooks.Open(excelFilePath) comme avant
