Attribute VB_Name = "Module1"
Option Explicit

Sub ClotureDuMois()
    ' Automatise la procédure documentée dans HISTORIQUE (lignes 5 à 9) :
    ' fige JOURNAL dans le bloc du mois en cours, marque le mois comme
    ' Clos dans le tableau de suivi, puis fait avancer PARAMÈTRES vers
    ' le mois suivant. Ne touche jamais un bloc déjà Clos.
    Dim wsParam As Worksheet, wsJournal As Worksheet, wsHist As Worksheet
    Dim anneeCourante As Long, moisNumero As Long
    Dim trackRow As Long, blocRow As Long, dstRow As Long, srcRow As Long
    Dim agentsArchives As Long
    Dim nomsMois(1 To 12) As String
    Dim i As Long
    Dim reponse As VbMsgBoxResult

    nomsMois(1) = "Janvier": nomsMois(2) = "Février": nomsMois(3) = "Mars"
    nomsMois(4) = "Avril": nomsMois(5) = "Mai": nomsMois(6) = "Juin"
    nomsMois(7) = "Juillet": nomsMois(8) = "Août": nomsMois(9) = "Septembre"
    nomsMois(10) = "Octobre": nomsMois(11) = "Novembre": nomsMois(12) = "Décembre"

    On Error GoTo ErreurConfig
    Set wsParam = ThisWorkbook.Worksheets("PARAMETRES")
    Set wsJournal = ThisWorkbook.Worksheets("JOURNAL")
    Set wsHist = ThisWorkbook.Worksheets("HISTORIQUE")
    On Error GoTo 0

    Application.Calculate

    anneeCourante = wsParam.Range("C29").Value

    moisNumero = 0
    For i = 1 To 12
        If wsParam.Range("C30").Value = nomsMois(i) Then
            moisNumero = i
            Exit For
        End If
    Next i
    If moisNumero = 0 Then
        MsgBox "Mois en cours (PARAMÈTRES!C30) non reconnu : " & wsParam.Range("C30").Value & _
               ". Clôture annulée.", vbCritical, "Clôture du mois"
        Exit Sub
    End If

    trackRow = 13 + moisNumero        ' lignes 14 à 25 du tableau de suivi
    blocRow = 28 + (moisNumero - 1) * 64

    If wsHist.Cells(trackRow, "D").Value = "Clos" Then
        MsgBox "Le bloc " & wsHist.Cells(trackRow, "B").Value & " de HISTORIQUE est déjà clôturé" & _
               " (année " & wsHist.Cells(trackRow, "C").Value & ")." & vbCrLf & _
               "Cette maquette ne gère qu'un seul jeu de 12 blocs (une année). Pour archiver" & _
               " une deuxième année, HISTORIQUE doit d'abord être étendu manuellement" & _
               " (nouveaux blocs ajoutés en fin de feuille)." & vbCrLf & _
               "Clôture annulée, aucune donnée modifiée.", vbCritical, "Clôture impossible"
        Exit Sub
    End If

    reponse = MsgBox("Clôturer " & nomsMois(moisNumero) & " " & anneeCourante & " ?" & vbCrLf & _
              "Les 60 lignes de JOURNAL seront figées (valeurs) dans le bloc " & nomsMois(moisNumero) & _
              " de HISTORIQUE, le mois sera marqué Clos, puis le mois en cours de PARAMÈTRES avancera.", _
              vbYesNo + vbQuestion, "Clôture du mois")
    If reponse <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    agentsArchives = 0
    dstRow = blocRow + 2
    For srcRow = 7 To 66
        If wsJournal.Cells(srcRow, "C").Value <> "" Then
            wsHist.Cells(dstRow, "B").Value = anneeCourante                        ' Année
            wsHist.Cells(dstRow, "C").Value = nomsMois(moisNumero)                 ' Mois
            wsHist.Cells(dstRow, "D").Value = wsJournal.Cells(srcRow, "C").Value   ' Matricule
            wsHist.Cells(dstRow, "E").Value = wsJournal.Cells(srcRow, "D").Value   ' Nom
            wsHist.Cells(dstRow, "F").Value = wsJournal.Cells(srcRow, "Q").Value   ' BRUT
            wsHist.Cells(dstRow, "G").Value = wsJournal.Cells(srcRow, "R").Value   ' CNSS salariale
            wsHist.Cells(dstRow, "H").Value = wsJournal.Cells(srcRow, "T").Value   ' IPR retenu
            wsHist.Cells(dstRow, "I").Value = wsJournal.Cells(srcRow, "U").Value   ' NET avant retenues
            wsHist.Cells(dstRow, "J").Value = wsJournal.Cells(srcRow, "V").Value   ' Avance
            wsHist.Cells(dstRow, "K").Value = wsJournal.Cells(srcRow, "W").Value   ' Retenues diverses
            wsHist.Cells(dstRow, "L").Value = wsJournal.Cells(srcRow, "X").Value   ' NET À PAYER
            wsHist.Cells(dstRow, "M").Value = wsJournal.Cells(srcRow, "Y").Value   ' CNSS patronale
            wsHist.Cells(dstRow, "N").Value = wsJournal.Cells(srcRow, "Z").Value   ' INPP
            wsHist.Cells(dstRow, "O").Value = wsJournal.Cells(srcRow, "AA").Value  ' ONEM
            wsHist.Cells(dstRow, "P").Value = wsJournal.Cells(srcRow, "AB").Value  ' Coût total employeur
            agentsArchives = agentsArchives + 1
        Else
            wsHist.Range(wsHist.Cells(dstRow, "B"), wsHist.Cells(dstRow, "P")).ClearContents
        End If
        dstRow = dstRow + 1
    Next srcRow

    wsHist.Cells(trackRow, "C").Value = anneeCourante   ' fige l'année (était une formule)
    wsHist.Cells(trackRow, "D").Value = "Clos"
    wsHist.Cells(trackRow, "E").Value = Date
    wsHist.Cells(trackRow, "F").Value = Application.UserName

    If moisNumero = 12 Then
        wsParam.Range("C30").Value = "Janvier"
        wsParam.Range("C29").Value = anneeCourante + 1
    Else
        wsParam.Range("C30").Value = nomsMois(moisNumero + 1)
    End If

    Application.Calculate
    Application.ScreenUpdating = True

    MsgBox agentsArchives & " agent(s) archivé(s) pour " & nomsMois(moisNumero) & " " & anneeCourante & "." & vbCrLf & _
           "Mois en cours désormais : " & wsParam.Range("C30").Value & " " & wsParam.Range("C29").Value, _
           vbInformation, "Clôture terminée"
    Exit Sub

ErreurConfig:
    Application.ScreenUpdating = True
    MsgBox "Impossible de trouver les feuilles PARAMETRES / JOURNAL / HISTORIQUE dans ce classeur." & vbCrLf & _
           "Erreur : " & Err.Description, vbCritical, "Clôture du mois"
End Sub
