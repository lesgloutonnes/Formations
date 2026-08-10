Attribute VB_Name = "CreationPDFMassif"
'==============================================================================
' Génération PDF en masse (MACCS)
' Sous Win11, le chemin UNC Bureau Isilon n'est plus fiable (OneDrive, etc.).
' L'utilisateur choisit le fichier Excel via l'explorateur Windows.
' Les PDF sont enregistrés dans le même dossier que le fichier Excel sélectionné.
'==============================================================================

Sub CreationPDFMassif()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim lastRow As Integer
    Dim nom As String
    Dim prenom As String
    Dim penta As String
    Dim mdp As String
    Dim mailchc As String
    Dim omni As String
    Dim mdp2 As String
    Dim pdfFileName As String
    Dim outputFolder As String
    Dim excelFilePath As String
    Dim cc As ContentControl
    Dim fd As FileDialog
    
    ' --- Sélection du fichier Excel via l'explorateur Windows ---
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Sélectionnez le fichier Excel MACCS (Modèle tableau Maccs.xlsx)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
        .Filters.Add "Tous les fichiers", "*.*"
        ' Démarrer sur le Bureau local si possible
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
    
    ' Dossier de sortie = dossier du fichier Excel choisi
    outputFolder = Left$(excelFilePath, InStrRev(excelFilePath, "\"))
    
    ' Vérification que le fichier existe bien
    If Dir(excelFilePath) = "" Then
        MsgBox "Le fichier sélectionné est introuvable :" & vbCrLf & excelFilePath, vbCritical
        Exit Sub
    End If
    
    MsgBox "Génération des PDF des MACCS, appuyez sur OK pour commencer et patientez !", vbInformation
    
    ' Ouvrir Excel et le fichier
    On Error GoTo ErrHandler
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set xlSheet = xlBook.Sheets(1)
    
    ' Trouver la dernière ligne avec des données
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(-4162).Row ' -4162 = xlUp
    
    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans le fichier Excel (à partir de la ligne 2).", vbExclamation
        GoTo CleanUp
    End If
    
    ' Boucle à travers chaque ligne (commençant à la ligne 2)
    For i = 2 To lastRow
        nom = xlSheet.Cells(i, 1).Value       ' Colonne A
        prenom = xlSheet.Cells(i, 2).Value    ' Colonne B
        penta = xlSheet.Cells(i, 17).Value    ' Colonne Q
        mdp = xlSheet.Cells(i, 18).Value      ' Colonne R
        mailchc = xlSheet.Cells(i, 19).Value  ' Colonne S
        omni = xlSheet.Cells(i, 13).Value     ' Colonne M
        mdp2 = xlSheet.Cells(i, 14).Value     ' Colonne N
        
        ' Ignorer les lignes vides
        If Trim$(nom) = "" And Trim$(prenom) = "" Then GoTo NextRow
        
        ' Remplir les champs du document Word
        For Each cc In ActiveDocument.ContentControls
            Select Case cc.Title
                Case "NOM"
                    cc.Range.Text = nom & " " & prenom
                Case "PENTA"
                    cc.Range.Text = penta
                Case "MDP"
                    cc.Range.Text = mdp
                Case "OWA"
                    cc.Range.Text = penta
                Case "MAILCHC"
                    cc.Range.Text = mailchc
                Case "OMNI"
                    cc.Range.Text = omni
                Case "MDP2"
                    cc.Range.Text = mdp2
            End Select
        Next cc
        
        ' Colonne O (PMI)
        If LCase$(CStr(xlSheet.Cells(i, 15).Value)) = "ok" Then
            Set cc = GetContentControlByTitle("PMI")
            If Not cc Is Nothing Then cc.Checked = True
        End If
        
        ' Colonne P (Cyberlab)
        If LCase$(CStr(xlSheet.Cells(i, 16).Value)) = "ok" Then
            Set cc = GetContentControlByTitle("Cyberlab")
            If Not cc Is Nothing Then cc.Checked = True
        End If
        
        ' Colonne T (Orline)
        If LCase$(CStr(xlSheet.Cells(i, 20).Value)) = "ok" Then
            Set cc = GetContentControlByTitle("Orline")
            If Not cc Is Nothing Then cc.Checked = True
        End If
        
        ' Enregistrer en PDF dans le même dossier que l'Excel
        pdfFileName = outputFolder & "Accès - " & nom & " " & prenom & ".pdf"
        ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfFileName, ExportFormat:=wdExportFormatPDF
        
        ResetRichTextContentControls
NextRow:
    Next i
    
    MsgBox "Génération de PDF terminée avec succès !" & vbCrLf & _
           "Dossier : " & outputFolder, vbInformation
    GoTo CleanUp
    
ErrHandler:
    MsgBox "Erreur lors de la génération :" & vbCrLf & Err.Description, vbCritical
    
CleanUp:
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    ResetRichTextContentControls
End Sub

' Fonction pour obtenir un ContentControl par son titre
Function GetContentControlByTitle(ByVal title As String) As ContentControl
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If cc.Title = title Then
            Set GetContentControlByTitle = cc
            Exit Function
        End If
    Next cc
    Set GetContentControlByTitle = Nothing
End Function
