'==============================================================================
' CORRECTIF CreationPDFMassif — Win11 / chemin Bureau Isilon
' Remplace le bloc UserName + chemins UNC en dur par une boîte Parcourir.
'
' Dans votre module existant, SUPPRIMER ces lignes :
'
'   Dim UserName As String
'   UserName = Environ("USERNAME")
'   excelFilePath = "\\FS.ISILON.CHC.BE\Datas\usersdata\" & UserName & "\Desktop\MACCS\Modèle tableau Maccs.xlsx"
'   outputFolder = "\\FS.ISILON.CHC.BE\Datas\usersdata\" & UserName & "\Desktop\MACCS\"
'
' et les REMPLACER par le bloc ci-dessous (avant l'ouverture d'Excel).
'==============================================================================

    Dim fd As FileDialog
    Dim excelFilePath As String
    Dim outputFolder As String
    
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
    
    ' PDF générés dans le même dossier que l'Excel choisi
    outputFolder = Left$(excelFilePath, InStrRev(excelFilePath, "\"))
    
    If Dir(excelFilePath) = "" Then
        MsgBox "Le fichier sélectionné est introuvable :" & vbCrLf & excelFilePath, vbCritical
        Exit Sub
    End If
    
    ' Ensuite : Set xlBook = xlApp.Workbooks.Open(excelFilePath)  (inchangé)
