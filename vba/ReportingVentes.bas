Attribute VB_Name = "ReportingVentes"
'==============================================================================
' ReportingVentes - nettoyage d'un export de ventes + tableau croise dynamique
' Version 100 % Excel (aucune installation), compatible Windows et Mac.
'
' Usage : ouvrir l'export, se placer sur l'onglet des donnees (en-tetes en
' ligne 1), puis lancer la macro NettoyerEtCreerTCD (Alt+F8).
'
' Ce que fait la macro :
'   1. copie les donnees dans un nouvel onglet (l'onglet d'origine n'est jamais
'      modifie) ;
'   2. supprime les espaces parasites et les lignes vides ;
'   3. harmonise la casse (ex. "BRETAGNE", "bretagne" -> "Bretagne") ;
'   4. convertit les dates et les montants saisis en texte ;
'   5. supprime les doublons ;
'   6. signale les dates / montants illisibles dans une colonne "Anomalie" ;
'   7. cree un tableau croise : CA HT par mois (lignes) et par region (colonnes).
'
' Fichier volontairement en ASCII : les accents sont produits par la fonction
' T() pour que l'import fonctionne quel que soit l'encodage de l'editeur VBA.
'==============================================================================
Option Explicit

' --- Intitules acceptes pour chaque colonne (sans accents, minuscules, "|") ---
Private Const COL_DATE As String = "date|date commande|date facture|date de vente"
Private Const COL_CLIENT As String = "client|raison sociale|nom client"
Private Const COL_PRODUIT As String = "produit|article|designation|libelle article"
Private Const COL_REGION As String = "region|zone|secteur"
Private Const COL_COMMERCIAL As String = "commercial|vendeur|representant"
Private Const COL_HT As String = "montant ht|total ht|ca ht"
Private Const COL_TVA As String = "tva|montant tva"
Private Const COL_TTC As String = "montant ttc|total ttc"

Public Sub NettoyerEtCreerTCD()
    Dim src As Worksheet, dst As Worksheet, tcd As Worksheet
    Dim data As Variant, outArr() As Variant
    Dim nRows As Long, nCols As Long, r As Long, c As Long, k As Long, nOut As Long
    Dim cDate As Long, cClient As Long, cProduit As Long, cRegion As Long
    Dim cCommercial As Long, cHT As Long, cTVA As Long, cTTC As Long
    Dim nVides As Long, nDoublons As Long, nAnomalies As Long, ok As Boolean
    Dim v As Variant, anomalie As String, d As Date

    Set src = ActiveSheet
    If src.Name = T("Donn[e]es nettoy[e]es") Or src.Name = "TCD" Then
        MsgBox T("Placez-vous sur l'onglet de l'export brut avant de lancer la macro."), vbExclamation
        Exit Sub
    End If
    cDate = TrouverColonne(src, COL_DATE)
    cClient = TrouverColonne(src, COL_CLIENT)
    cProduit = TrouverColonne(src, COL_PRODUIT)
    cRegion = TrouverColonne(src, COL_REGION)
    cCommercial = TrouverColonne(src, COL_COMMERCIAL)
    cHT = TrouverColonne(src, COL_HT)
    cTVA = TrouverColonne(src, COL_TVA)
    cTTC = TrouverColonne(src, COL_TTC)
    If cDate = 0 Or cClient = 0 Or cHT = 0 Or cRegion = 0 Then
        MsgBox T("Colonnes Date, Client, R[e]gion ou Montant HT introuvables en ligne 1." & vbLf & _
                 "Adaptez les constantes COL_... en haut du module."), vbExclamation
        Exit Sub
    End If

    nRows = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For c = 2 To src.Cells(1, src.Columns.Count).End(xlToLeft).Column
        If src.Cells(src.Rows.Count, c).End(xlUp).Row > nRows Then nRows = src.Cells(src.Rows.Count, c).End(xlUp).Row
    Next c
    nCols = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    If nRows < 2 Then
        MsgBox T("Aucune donn[e]e sous la ligne d'en-t[e]te."), vbExclamation
        Exit Sub
    End If
    data = src.Range(src.Cells(1, 1), src.Cells(nRows, nCols)).Value

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- 1. Espaces parasites (toutes les cellules texte) ---
    For r = 2 To nRows
        For c = 1 To nCols
            If VarType(data(r, c)) = vbString Then
                data(r, c) = Application.WorksheetFunction.Trim( _
                    Replace(Replace(data(r, c), ChrW(160), " "), ChrW(8239), " "))
            End If
        Next c
    Next r

    ' --- 2. Casse harmonisee sur les colonnes de libelles ---
    HarmoniserCasse data, cClient, nRows
    HarmoniserCasse data, cProduit, nRows
    HarmoniserCasse data, cRegion, nRows
    HarmoniserCasse data, cCommercial, nRows

    ' --- 3. Lignes vides, dates, montants ; ajout des colonnes Mois et Anomalie ---
    ReDim outArr(1 To nRows, 1 To nCols + 2)
    For c = 1 To nCols
        outArr(1, c) = data(1, c)
    Next c
    outArr(1, nCols + 1) = "Mois"
    outArr(1, nCols + 2) = "Anomalie"
    nOut = 1
    For r = 2 To nRows
        If LigneVide(data, r, nCols) Then
            nVides = nVides + 1
        Else
            nOut = nOut + 1
            anomalie = ""
            For c = 1 To nCols
                outArr(nOut, c) = data(r, c)
            Next c
            ' Date
            v = data(r, cDate)
            ok = False
            If VarType(v) = vbDate Then
                d = v: ok = True
            ElseIf VarType(v) = vbString Then
                d = LireDate(CStr(v), ok)
            ElseIf VarType(v) = vbDouble Then     ' numero de serie Excel non formate
                If v >= 20000 And v <= 80000 Then d = CDate(v): ok = True
            End If
            If ok Then
                outArr(nOut, cDate) = d
                outArr(nOut, nCols + 1) = Format(d, "yyyy-mm")
            Else
                anomalie = anomalie & T("Date illisible ; ")
            End If
            ' Montants
            For Each v In Array(cHT, cTVA, cTTC)
                k = CLng(v)
                If k > 0 Then
                    If VarType(data(r, k)) = vbString Then
                        outArr(nOut, k) = LireMontant(CStr(data(r, k)), ok)
                        If Not ok Then
                            outArr(nOut, k) = data(r, k)
                            anomalie = anomalie & CStr(data(1, k)) & " illisible ; "
                        End If
                    ElseIf IsEmpty(data(r, k)) And k = cHT Then
                        anomalie = anomalie & "Montant HT vide ; "
                    End If
                End If
            Next v
            If Len(CStr(data(r, cClient))) = 0 Then anomalie = anomalie & "Client manquant ; "
            outArr(nOut, nCols + 2) = anomalie
        End If
    Next r

    ' --- 4. Ecriture dans un nouvel onglet ---
    Set dst = NouvelOnglet(T("Donn[e]es nettoy[e]es"))
    dst.Range("A1").Resize(nOut, nCols + 2).Value = outArr

    ' --- 5. Doublons (comparaison sur les colonnes d'origine) ---
    Dim cols() As Variant
    ReDim cols(0 To nCols - 1)
    For c = 1 To nCols
        cols(c - 1) = c
    Next c
    dst.Range("A1").Resize(nOut, nCols + 2).RemoveDuplicates Columns:=(cols), Header:=xlYes
    nDoublons = nOut - dst.Range("A1").CurrentRegion.Rows.Count
    nOut = nOut - nDoublons

    ' --- 6. Mise en forme : tableau Excel, formats, anomalies en orange ---
    Dim lo As ListObject
    Set lo = dst.ListObjects.Add(xlSrcRange, dst.Range("A1").Resize(nOut, nCols + 2), , xlYes)
    lo.Name = "T_Ventes"
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns(cDate).DataBodyRange.NumberFormat = "dd/mm/yyyy"
    For Each v In Array(cHT, cTVA, cTTC)
        If CLng(v) > 0 Then lo.ListColumns(CLng(v)).DataBodyRange.NumberFormat = "#,##0.00"
    Next v
    For r = 1 To lo.ListRows.Count
        If Len(lo.DataBodyRange.Cells(r, nCols + 2).Value) > 0 Then
            lo.ListRows(r).Range.Interior.Color = RGB(255, 229, 204)
            nAnomalies = nAnomalies + 1
        End If
    Next r
    dst.Columns.AutoFit
    dst.Activate
    ActiveWindow.FreezePanes = False
    dst.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' --- 7. Tableau croise dynamique : CA HT par mois x region ---
    Set tcd = NouvelOnglet("TCD")
    Dim pc As PivotCache, pt As PivotTable
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo.Name)
    Set pt = pc.CreatePivotTable(TableDestination:=tcd.Range("A3"), TableName:="TCD_Ventes")
    With pt
        .PivotFields("Mois").Orientation = xlRowField
        .PivotFields(CStr(data(1, cRegion))).Orientation = xlColumnField
        With .AddDataField(.PivotFields(CStr(data(1, cHT))), "CA HT", xlSum)
            .NumberFormat = "#,##0"
        End With
        .RowAxisLayout xlTabularRow
    End With
    tcd.Range("A1").Value = T("Chiffre d'affaires HT par mois et par r[e]gion")
    tcd.Range("A1").Font.Bold = True
    tcd.Range("A1").Font.Size = 14

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    tcd.Activate
    MsgBox T("Termin[e]." & vbLf & _
             "Lignes lues : " & (nRows - 1) & vbLf & _
             "Lignes vides supprim[e]es : " & nVides & vbLf & _
             "Doublons supprim[e]s : " & nDoublons & vbLf & _
             "Lignes [a] v[e]rifier (en orange) : " & nAnomalies), vbInformation
    Exit Sub

Erreur:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
End Sub

'--- Remplace [e] [E] [a] par e accent aigu, E accent aigu, a accent grave ---
Private Function T(ByVal s As String) As String
    s = Replace(s, "[e]", ChrW(233))
    s = Replace(s, "[E]", ChrW(201))
    s = Replace(s, "[a]", ChrW(224))
    T = s
End Function

'--- Texte sans accents ni majuscules, pour comparer les intitules ---
Private Function Cle(ByVal s As String) As String
    Dim i As Long, a As Variant, b As Variant
    a = Array(224, 226, 228, 231, 232, 233, 234, 235, 238, 239, 244, 246, 249, 251, 252, 176)
    b = Array("a", "a", "a", "c", "e", "e", "e", "e", "i", "i", "o", "o", "u", "u", "u", "")
    s = LCase(Application.WorksheetFunction.Trim(s))
    For i = LBound(a) To UBound(a)
        s = Replace(s, ChrW(a(i)), b(i))
    Next i
    Cle = s
End Function

Private Function TrouverColonne(ws As Worksheet, ByVal alias As String) As Long
    Dim c As Long, noms As Variant, i As Long
    noms = Split(alias, "|")
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For i = LBound(noms) To UBound(noms)
            If Cle(CStr(ws.Cells(1, c).Value)) = noms(i) Then
                TrouverColonne = c
                Exit Function
            End If
        Next i
    Next c
    TrouverColonne = 0
End Function

Private Function LigneVide(data As Variant, ByVal r As Long, ByVal nCols As Long) As Boolean
    Dim c As Long
    For c = 1 To nCols
        If Len(CStr(data(r, c))) > 0 Then
            LigneVide = False
            Exit Function
        End If
    Next c
    LigneVide = True
End Function

'--- Pour chaque libelle, garde la premiere graphie en casse mixte rencontree ---
'    (Collection plutot que Scripting.Dictionary : fonctionne aussi sur Mac)
Private Sub HarmoniserCasse(data As Variant, ByVal col As Long, ByVal nRows As Long)
    Dim formes As New Collection, r As Long, s As String, f As String
    If col = 0 Then Exit Sub
    For r = 2 To nRows
        s = CStr(data(r, col))
        If Len(s) > 0 And s <> UCase(s) And s <> LCase(s) Then
            On Error Resume Next
            formes.Add s, Cle(s)          ' ignore si la cle existe deja
            On Error GoTo 0
        End If
    Next r
    For r = 2 To nRows
        s = CStr(data(r, col))
        If Len(s) > 0 Then
            f = ""
            On Error Resume Next
            f = formes(Cle(s))
            On Error GoTo 0
            If Len(f) > 0 Then data(r, col) = f
        End If
    Next r
End Sub

'--- Dates texte : jj/mm/aaaa, jj-mm-aaaa, jj.mm.aaaa, jj/mm/aa, aaaa-mm-jj ---
Private Function LireDate(ByVal s As String, ByRef ok As Boolean) As Date
    Dim j As Integer, m As Integer, a As Integer
    ok = False
    s = Trim(s)
    If InStr(s, " ") > 0 Then s = Left(s, InStr(s, " ") - 1)
    s = Replace(Replace(s, "-", "/"), ".", "/")
    If s Like "####/##/##" Then
        a = CInt(Left(s, 4)): m = CInt(Mid(s, 6, 2)): j = CInt(Mid(s, 9, 2))
    ElseIf s Like "##/##/####" Then
        j = CInt(Left(s, 2)): m = CInt(Mid(s, 4, 2)): a = CInt(Mid(s, 7, 4))
    ElseIf s Like "##/##/##" Then
        j = CInt(Left(s, 2)): m = CInt(Mid(s, 4, 2)): a = 2000 + CInt(Mid(s, 7, 2))
    Else
        Exit Function
    End If
    If m < 1 Or m > 12 Or j < 1 Or j > 31 Then Exit Function
    LireDate = DateSerial(a, m, j)
    ok = (Day(LireDate) = j)              ' rejette le 31/02, le 31/04, etc.
End Function

'--- Montants texte : "1 234,50 EUR", "1234.5", "1.234,50", "(120,00)" ---
Private Function LireMontant(ByVal s As String, ByRef ok As Boolean) As Double
    Dim negatif As Boolean
    ok = False
    s = Replace(s, ChrW(8364), "")
    s = Replace(s, "EUR", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, ChrW(8239), "")
    s = Replace(s, " ", "")
    If Left(s, 1) = "(" And Right(s, 1) = ")" Then
        negatif = True
        s = Mid(s, 2, Len(s) - 2)
    End If
    If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(Replace(s, ".", ""), ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    Else
        s = Replace(s, ",", ".")
    End If
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9.-]*" Then Exit Function
    If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
    LireMontant = Val(s)                  ' Val lit toujours le point comme decimal
    If negatif Then LireMontant = -LireMontant
    ok = True
End Function

'--- Cree (ou recree) un onglet en fin de classeur ---
Private Function NouvelOnglet(ByVal nom As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(nom).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set NouvelOnglet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    NouvelOnglet.Name = nom
End Function
