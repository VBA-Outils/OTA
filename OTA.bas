Attribute VB_Name = "OTA"
Option Explicit
Option Compare Text

Public Sub OTApgAdmin()

    Dim sTexte As String, sTabLig() As String, iIcLig As Integer, sTabCol() As String, iIcCol As Integer, sValeur As String, bEstChaine As Boolean
    Dim iColNbCar() As Integer
   
    On Error GoTo Dysfonctionnement
   
    ' Lire le PP
    sTexte = LirePressePapiers()
    If sTexte = "" Then
        MsgBox "Le presse-papiers est vide.", vbInformation
        Exit Sub
    End If
   
    ' Calculer la longueur max de chaque colonne
    For iIcLig = LBound(sTabLig) To UBound(sTabLig)
        sTabCol = Split(sTabLig(iIcLig), vbTab)
        If iIcLig = 0 Then
            ReDim iColNbCar(UBound(sTabCol)) As Integer
        End If
        ' Ligne vide ?
        If UBound(sTabCol) <> -1 Then
            ' Pour chaque colonne
            For iIcCol = LBound(sTabCol) To UBound(sTabCol)
                sValeur = sTabCol(iIcCol)
                ' Supprimer les guillemets " qui délimitent les textes
                If Left(sValeur, 1) = """" And Right(sValeur, 1) = """" Then
                    sValeur = Mid(sValeur, 2, Len(sValeur) - 2)
                    bEstChaine = True
                Else
                    bEstChaine = False
                    ' Si la valeur est non renseignée et n'est pas une chaine de caractères vide alors la remplacer par Null
                    If sValeur = "" Then sValeur = "Null"
                End If
                iColNbCar(iIcCol) = Application.WorksheetFunction.Max(iColNbCar(iIcCol), Len(sValeur))
            Next iIcCol
        End If
    Next iIcLig
   
    ' Découper par ligne
    sTabLig = Split(sTexte, vbLf)
    sTexte = ""
    For iIcLig = LBound(sTabLig) To UBound(sTabLig)
        ' Découper par colonne
        sTabCol = Split(sTabLig(iIcLig), vbTab)
        ' Ligne vide ?
        If UBound(sTabCol) <> -1 Then
            ' Si nouvelle ligne
            sTexte = sTexte & IIf(sTexte = "", "", vbCrLf)
            ' Pour chaque colonne
            For iIcCol = LBound(sTabCol) To UBound(sTabCol)
                sValeur = sTabCol(iIcCol)
                ' Supprimer les guillemets " qui délimitent les textes
                If Left(sValeur, 1) = """" And Right(sValeur, 1) = """" Then
                    sValeur = Mid(sValeur, 2, Len(sValeur) - 2)
                    bEstChaine = True
                Else
                    bEstChaine = False
                End If
                ' Si la valeur est non renseignée et n'est pas une chaine de caractères vide alors la remplacer par Null
                sTexte = sTexte & "|" & IIf(sTabCol(iIcCol) = "" And bEstChaine = False, "Null", sValeur)
            Next iIcCol
            ' Délimiteur de fin
            sTexte = sTexte & "|"
        End If
    Next iIcLig
    ' Copier le résultat dans le presse-papiers
    Call EcrirePressePapiers(sTexte)

    Exit Sub
   
Dysfonctionnement:
    MsgBox "Un grain de sable s'est glissé dans la mécanique bien huilée du traitement." & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description : " & Err.Description, vbCritical, "Fin anormale du traitement ""Reinitialiser"""

End Sub

Public Sub OTAsquirrel()

    Dim sTexte As String, sTabLig() As String, iIcLig As Integer, sTabCol() As String, iIcCol As Integer, sValeur As String, bEstChaine As Boolean
    Dim iColNbCar() As Integer
   
    On Error GoTo Dysfonctionnement
   
    ' Lire le PP
    sTexte = LirePressePapiers()
    If sTexte = "" Then
        MsgBox "Le presse-papiers est vide.", vbInformation
        Exit Sub
    End If
   
    ' Découper par ligne
    sTabLig = Split(sTexte, vbCrLf)
   
    ' Calculer la longueur max de chaque colonne
    For iIcLig = LBound(sTabLig) To UBound(sTabLig)
        sTabCol = Split(sTabLig(iIcLig), vbTab)
        If iIcLig = 0 Then
            ReDim iColNbCar(UBound(sTabCol)) As Integer
        End If
        ' Ligne vide ?
        If UBound(sTabCol) <> -1 Then
            ' Pour chaque colonne
            For iIcCol = LBound(sTabCol) To UBound(sTabCol)
                iColNbCar(iIcCol) = Application.WorksheetFunction.Max(iColNbCar(iIcCol), Len(sTabCol(iIcCol)))
            Next iIcCol
        End If
    Next iIcLig
   
    sTexte = ""
    For iIcLig = LBound(sTabLig) To UBound(sTabLig)
        ' Découper par colonne
        sTabCol = Split(sTabLig(iIcLig), vbTab)
        ' Ligne vide ?
        If UBound(sTabCol) <> -1 Then
            ' Si nouvelle ligne
            sTexte = sTexte & IIf(sTexte = "", "", vbCrLf)
            ' Pour chaque colonne
            For iIcCol = LBound(sTabCol) To UBound(sTabCol)
                sValeur = sTabCol(iIcCol)
                If IsNumeric(sValeur) Then
                    sValeur = Replace(Replace(sValeur, Chr$(160), "", , , vbTextCompare), ",", ".", , , vbTextCompare)
                End If
                If LCase$(sValeur) = "<null>" Then sValeur = "Null"
                sTexte = sTexte & "|" & CadreGauche(sValeur, iColNbCar(iIcCol))
            Next iIcCol
            ' Délimiteur de fin
            sTexte = sTexte & "|"
        End If
    Next iIcLig
    ' Copier le résultat dans le presse-papiers
    Call EcrirePressePapiers(sTexte)

    Exit Sub
   
Dysfonctionnement:
    MsgBox "Un grain de sable s'est glissé dans la mécanique bien huilée du traitement." & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description : " & Err.Description, vbCritical, "Fin anormale du traitement ""Reinitialiser"""

End Sub

Private Function CadreGauche(sChaine As String, iLongueur As Integer) As String

    If Len(sChaine) > iLongueur Then iLongueur = Len(sChaine)
    CadreGauche = Space$(iLongueur)
    Mid$(CadreGauche, 1, Len(sChaine)) = sChaine
   
End Function
