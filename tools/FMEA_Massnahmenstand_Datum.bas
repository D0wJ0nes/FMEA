Attribute VB_Name = "FMEA_Massnahmenstand_Datum"
'===============================================================================
' FMEA_Massnahmenstand_Datum
'
' Setzt in einem APIS IQ-FMEA XML-Export (MSR-FMEA) das Datum aller
' Massnahmenstaende auf ein eingegebenes Datum.
' Optional werden auch die Termine der Massnahmen in den Folgestaenden
' (alle Staende nach dem ersten Stand einer Ursache) auf dieses Datum gesetzt.
'
' - Die Historie der Massnahmen wird nicht veraendert.
' - Das Datum wird wie von APIS gespeichert: 23:59:59 Uhr, Zeitzone
'   Europe/Berlin (+02:00 Sommerzeit, +01:00 Winterzeit).
' - Ausser den Datumswerten bleibt die Datei Byte fuer Byte gleich.
' - Das Ergebnis wird als neue Datei neben dem Original gespeichert
'   (<Name>_JJJJ-MM-TT.xml). Das Original wird nie veraendert.
'
' Start: Makro "Massnahmenstaende_Datum_aendern" ausfuehren.
'===============================================================================
Option Explicit

Private Const DATUM_TAG As String = "<DATE-1 SYSCOND=""Europe/Berlin"">"
Private Const DATUM_LAENGE As Long = 25          ' z.B. 2027-04-30T23:59:59+02:00
Private Const FEHLER_NR As Long = 1513           ' eigene Fehlernummer

Private P_HIST As Variant, P_HIST_ENDE As Variant
Private P_SETS As Variant, P_SETS_ENDE As Variant
Private P_SET1 As Variant, P_SET2 As Variant
Private P_DATUM As Variant, P_DATUM_TAG As Variant, P_DATUM_ENDE As Variant
Private P_OCC As Variant, P_OCC_ENDE As Variant
Private P_DET As Variant, P_DET_ENDE As Variant
Private P_REF As Variant
Private P_ACTIONS As Variant, P_ACTIONS_ENDE As Variant
Private P_ACTION As Variant, P_ACTION_ENDE As Variant
Private P_STATE As Variant, P_HISTSTATES As Variant, P_SDGS As Variant
Private P_SCHEDULE As Variant


'-------------------------------------------------------------------------------
' Start mit Dialogen (Excel)
'-------------------------------------------------------------------------------
Public Sub Massnahmenstaende_Datum_aendern()
    Dim inPfad As Variant, eingabe As String, neuesDatum As Date
    Dim antwort As VbMsgBoxResult, mitTerminen As Boolean
    Dim outPfad As String, meldung As String

    ' 1. Datei waehlen
    inPfad = Application.GetOpenFilename("FMEA XML-Export (*.xml),*.xml", , _
             Txt("APIS FMEA XML-Export ausw{ae}hlen"))
    If VarType(inPfad) = vbBoolean Then Exit Sub            ' abgebrochen

    ' 2. Datum eingeben
    Do
        eingabe = InputBox(Txt("Neues Datum f{ue}r alle Ma{ss}nahmenst{ae}nde (TT.MM.JJJJ):"), _
                           Txt("Ma{ss}nahmenst{ae}nde: Datum {ae}ndern"))
        If Len(Trim$(eingabe)) = 0 Then Exit Sub            ' abgebrochen oder leer
        If ParseDatum(eingabe, neuesDatum) Then Exit Do
        MsgBox Txt("Ung{ue}ltiges Datum: """ & eingabe & """" & vbCrLf & _
                   "Bitte im Format TT.MM.JJJJ eingeben, z.B. 30.04.2027."), vbExclamation
    Loop

    ' 3. Termine der Massnahmen in den Folgestaenden?
    antwort = MsgBox(Txt("Sollen auch die Termine der Ma{ss}nahmen in den Folgest{ae}nden " & _
                         "auf den " & DeutschesDatum(neuesDatum) & " gesetzt werden?" & vbCrLf & vbCrLf & _
                         "Folgest{ae}nde = alle Ma{ss}nahmenst{ae}nde nach dem ersten Stand einer Ursache." & vbCrLf & _
                         "Die Historie wird in keinem Fall ver{ae}ndert."), _
                     vbYesNoCancel + vbQuestion, Txt("Termine der Ma{ss}nahmen"))
    If antwort = vbCancel Then Exit Sub
    mitTerminen = (antwort = vbYes)

    ' 4. Ausgabedatei neben dem Original
    outPfad = AusgabePfad(CStr(inPfad), neuesDatum)
    If Len(Dir(outPfad)) > 0 Then
        If MsgBox(Txt("Die Datei existiert bereits:" & vbCrLf & outPfad & vbCrLf & vbCrLf & _
                      "{Ue}berschreiben?"), vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If

    ' 5. Ausfuehren
    Application.Cursor = xlWait
    If DatumSetzen(CStr(inPfad), outPfad, neuesDatum, mitTerminen, meldung) Then
        Application.Cursor = xlDefault
        MsgBox meldung, vbInformation, "Fertig"
    Else
        Application.Cursor = xlDefault
        MsgBox meldung, vbCritical, "Fehler"
    End If
End Sub


'-------------------------------------------------------------------------------
' Kernfunktion (ohne Dialoge)
' Gibt True zurueck, wenn die Ausgabedatei geschrieben wurde. In meldung steht
' die Zusammenfassung bzw. der Fehler. Bei einem Fehler wird keine Datei
' geschrieben.
'-------------------------------------------------------------------------------
Public Function DatumSetzen(ByVal inPfad As String, ByVal outPfad As String, _
                            ByVal neuesDatum As Date, ByVal mitTerminen As Boolean, _
                            ByRef meldung As String) As Boolean
    Dim buf() As Byte, pruef() As Byte, n As Long, f As Integer
    Dim wert() As Byte, wertText As String
    Dim pos As Long, h As Long, he As Long, sp As Long, se As Long, q As Long
    Dim grenze As Long, x As Long, dp As Long, k As Long
    Dim nStaendeGesamt As Long, nAnfang As Long, nFolge As Long
    Dim ids As String, nIds As Long
    Dim a0 As Long, a1 As Long, ap As Long, ae As Long, id As String
    Dim nGefunden As Long, nTermine As Long, nOhneTermin As Long, ts As Long
    Dim schreibenBegonnen As Boolean, i As Long

    On Error GoTo fehler
    InitMuster
    If StrComp(inPfad, outPfad, vbTextCompare) = 0 Then _
        Fehler "Eingabe- und Ausgabedatei sind identisch."

    If Len(Dir(inPfad)) = 0 Then Fehler "Die Datei wurde nicht gefunden: " & inPfad

    ' Datei als Bytes lesen (keine Zeichensatz-Umwandlung)
    f = FreeFile
    Open inPfad For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then Close #f: Fehler "Die Datei ist leer."
    ReDim buf(0 To n - 1)
    Get #f, , buf
    Close #f
    f = 0

    If FindB(buf, ToBytes("<MSRFMEA"), 0, IIf(n < 4096, n, 4096)) < 0 Then _
        Fehler "Die Datei ist kein APIS MSR-FMEA XML-Export (<MSRFMEA> fehlt)."

    wertText = IsoDatum(neuesDatum) & "T23:59:59" & BerlinOffset(neuesDatum)
    wert = ToBytes(wertText)

    ' Anzahl aller Massnahmenstaende in der Datei (zur Kontrolle)
    pos = 0
    Do
        x = FindTag(buf, P_SETS, pos, n)
        If x < 0 Then Exit Do
        nStaendeGesamt = nStaendeGesamt + 1
        pos = x + 1
    Loop
    If nStaendeGesamt = 0 Then Fehler Txt("In der Datei wurden keine Ma{ss}nahmenst{ae}nde gefunden.")

    ' Massnahmenstaende je Ursache (FM-TASK-HISTORY) durchgehen
    ids = "|"
    pos = 0
    Do
        h = FindTag(buf, P_HIST, pos, n)
        If h < 0 Then Exit Do
        he = FindB(buf, P_HIST_ENDE, h, n)
        If he < 0 Then Fehler "Ende von <FM-TASK-HISTORY> nicht gefunden (Byte " & h & ")."
        k = 0
        q = h
        Do
            sp = FindTag(buf, P_SETS, q, he)
            If sp < 0 Then Exit Do
            se = FindB(buf, P_SETS_ENDE, sp, he)
            If se < 0 Then Fehler Txt("Ende eines Ma{ss}nahmenstands nicht gefunden (Byte " & sp & ").")
            ' Datum des Stands: das erste <DATE-1 vor der ersten Massnahmengruppe
            grenze = se
            x = FindB(buf, P_SET1, sp, se): If x >= 0 And x < grenze Then grenze = x
            x = FindB(buf, P_SET2, sp, se): If x >= 0 And x < grenze Then grenze = x
            dp = FindB(buf, P_DATUM, sp, grenze)
            If dp < 0 Then Fehler Txt("Ma{ss}nahmenstand ohne Datum gefunden (Byte " & sp & ").")
            DatumErsetzen buf, dp, grenze, wert
            If k = 0 Then nAnfang = nAnfang + 1 Else nFolge = nFolge + 1
            ' Massnahmen der Folgestaende merken (nur eigene Vermeidung/Entdeckung)
            If k >= 1 Then
                SammleRefs buf, P_OCC, P_OCC_ENDE, sp, se, ids, nIds
                SammleRefs buf, P_DET, P_DET_ENDE, sp, se, ids, nIds
            End If
            k = k + 1
            q = se + 1
        Loop
        pos = he + 1
    Loop
    If nAnfang + nFolge <> nStaendeGesamt Then _
        Fehler Txt(nStaendeGesamt & " Ma{ss}nahmenst{ae}nde in der Datei, aber nur " & (nAnfang + nFolge) & _
                   " davon bei Ursachen gefunden. Unerwarteter Dateiaufbau.")

    ' Termine der Massnahmen in den Folgestaenden
    If mitTerminen And nIds > 0 Then
        a0 = FindB(buf, P_ACTIONS, 0, n)
        If a0 < 0 Then Fehler "<FM-ACTIONS> nicht gefunden."
        a1 = FindB(buf, P_ACTIONS_ENDE, a0, n)
        If a1 < 0 Then Fehler "</FM-ACTIONS> nicht gefunden."
        pos = a0
        Do
            ap = FindB(buf, P_ACTION, pos, a1)
            If ap < 0 Then Exit Do
            id = LeseBisAnfuehrung(buf, ap + UBound(P_ACTION) + 1)
            ae = FindB(buf, P_ACTION_ENDE, ap, a1)
            If ae < 0 Then Fehler Txt("Ende der Ma{ss}nahme " & id & " nicht gefunden.")
            If InStr(1, ids, "|" & id & "|", vbBinaryCompare) > 0 Then
                nGefunden = nGefunden + 1
                ' nur der direkte Termin der Massnahme, nicht die Historie
                grenze = ae
                x = FindB(buf, P_STATE, ap, ae): If x >= 0 And x < grenze Then grenze = x
                x = FindB(buf, P_HISTSTATES, ap, ae): If x >= 0 And x < grenze Then grenze = x
                x = FindB(buf, P_SDGS, ap, ae): If x >= 0 And x < grenze Then grenze = x
                ts = FindB(buf, P_SCHEDULE, ap, grenze)
                If ts >= 0 Then
                    dp = FindB(buf, P_DATUM, ts, grenze)
                    If dp < 0 Then Fehler Txt("Termin der Ma{ss}nahme " & id & " ohne Datum.")
                    DatumErsetzen buf, dp, grenze, wert
                    nTermine = nTermine + 1
                Else
                    nOhneTermin = nOhneTermin + 1
                End If
            End If
            pos = ae + 1
        Loop
        If nGefunden <> nIds Then _
            Fehler Txt(nIds & " Ma{ss}nahmen in den Folgest{ae}nden, aber nur " & nGefunden & _
                       " davon in <FM-ACTIONS> gefunden.")
    End If

    ' Neue Datei schreiben und kontrollieren
    If Len(Dir(outPfad)) > 0 Then Kill outPfad
    schreibenBegonnen = True
    f = FreeFile
    Open outPfad For Binary Access Write As #f
    Put #f, , buf
    Close #f
    f = 0
    If FileLen(outPfad) <> n Then Fehler "Die geschriebene Datei hat eine falsche Gr" & ChrW(246) & ChrW(223) & "e."
    ReDim pruef(0 To n - 1)
    f = FreeFile
    Open outPfad For Binary Access Read As #f
    Get #f, , pruef
    Close #f
    f = 0
    For i = 0 To n - 1
        If pruef(i) <> buf(i) Then Fehler "Die geschriebene Datei weicht vom erwarteten Inhalt ab."
    Next i

    meldung = Txt("Neue Datei:" & vbCrLf & outPfad & vbCrLf & vbCrLf & _
              "Neues Datum: " & DeutschesDatum(neuesDatum) & "  (" & wertText & ")" & vbCrLf & vbCrLf & _
              "Ma{ss}nahmenst{ae}nde auf das neue Datum gesetzt: " & (nAnfang + nFolge) & vbCrLf & _
              "   davon erster Stand (Anfang): " & nAnfang & vbCrLf & _
              "   davon Folgest{ae}nde: " & nFolge & vbCrLf & vbCrLf)
    If mitTerminen Then
        meldung = meldung & Txt("Termine der Ma{ss}nahmen in Folgest{ae}nden auf das neue Datum gesetzt: " & nTermine)
        If nOhneTermin > 0 Then meldung = meldung & Txt(vbCrLf & "   " & nOhneTermin & _
                                         " Ma{ss}nahmen ohne Termin blieben ohne Termin.")
    Else
        meldung = meldung & Txt("Termine der Ma{ss}nahmen: nicht ge{ae}ndert")
    End If
    meldung = meldung & Txt(vbCrLf & "Historie: nicht ver{ae}ndert" & vbCrLf & vbCrLf & _
              "Das Original wurde nicht ver{ae}ndert.")
    DatumSetzen = True
    Exit Function

fehler:
    meldung = Err.Description
    On Error Resume Next
    If f > 0 Then Close #f
    If schreibenBegonnen Then Kill outPfad
    meldung = Txt("Abbruch, es wurde keine Datei geschrieben.") & vbCrLf & vbCrLf & meldung
    DatumSetzen = False
End Function


'-------------------------------------------------------------------------------
' Hilfsfunktionen
'-------------------------------------------------------------------------------

' Ersetzt den Datumswert ab <DATE-1 ...> an Position dp. Prueft vorher das Format.
Private Sub DatumErsetzen(buf() As Byte, ByVal dp As Long, ByVal grenze As Long, wert() As Byte)
    Dim v As Long, alt As String, i As Long
    If Not MatchAt(buf, dp, P_DATUM_TAG) Then _
        Fehler "Unerwartetes Datumsformat an Byte " & dp & ": " & LeseText(buf, dp, 60)
    v = dp + UBound(P_DATUM_TAG) + 1
    If v + DATUM_LAENGE > grenze Or Not MatchAt(buf, v + DATUM_LAENGE, P_DATUM_ENDE) Then _
        Fehler "Unerwartetes Datumsformat an Byte " & dp & ": " & LeseText(buf, dp, 60)
    alt = LeseText(buf, v, DATUM_LAENGE)
    If Not (alt Like "####-##-##T##:##:##[+-]##:##") Then _
        Fehler "Unerwartetes Datumsformat an Byte " & dp & ": " & alt
    For i = 0 To DATUM_LAENGE - 1
        buf(v + i) = wert(i)
    Next i
End Sub

' Sammelt die IDs aller <FM-ACTION-REF> innerhalb der Bloecke startTag...endeTag
' im Bereich [von, bis).
Private Sub SammleRefs(buf() As Byte, startTag As Variant, endeTag As Variant, _
                       ByVal von As Long, ByVal bis As Long, ByRef ids As String, ByRef nIds As Long)
    Dim b As Long, be As Long, r As Long, id As String
    b = von
    Do
        b = FindB(buf, startTag, b, bis)
        If b < 0 Then Exit Do
        be = FindB(buf, endeTag, b, bis)
        If be < 0 Then Fehler "Unerwarteter Dateiaufbau an Byte " & b & "."
        r = b
        Do
            r = FindB(buf, P_REF, r, be)
            If r < 0 Then Exit Do
            id = LeseBisAnfuehrung(buf, r + UBound(P_REF) + 1)
            If InStr(1, ids, "|" & id & "|", vbBinaryCompare) = 0 Then
                ids = ids & id & "|"
                nIds = nIds + 1
            End If
            r = r + 1
        Loop
        b = be + 1
    Loop
End Sub

' Sucht pat in buf ab Index von. Der Treffer muss ganz vor bis liegen (bis exklusiv).
' Gibt den Index oder -1 zurueck.
Private Function FindB(buf() As Byte, pat As Variant, ByVal von As Long, ByVal bis As Long) As Long
    Dim i As Long, j As Long, m As Long, c0 As Byte, letzter As Long
    m = UBound(pat) + 1
    If bis > UBound(buf) + 1 Then bis = UBound(buf) + 1
    letzter = bis - m
    c0 = pat(0)
    For i = von To letzter
        If buf(i) = c0 Then
            For j = 1 To m - 1
                If buf(i + j) <> pat(j) Then Exit For
            Next j
            If j >= m Then
                FindB = i
                Exit Function
            End If
        End If
    Next i
    FindB = -1
End Function

' Wie FindB, aber nur Treffer, auf die ein Leerzeichen oder ">" folgt (ganzer Tag-Name).
Private Function FindTag(buf() As Byte, pat As Variant, ByVal von As Long, ByVal bis As Long) As Long
    Dim x As Long, nach As Long
    Do
        x = FindB(buf, pat, von, bis)
        If x < 0 Then FindTag = -1: Exit Function
        nach = x + UBound(pat) + 1
        If nach <= UBound(buf) Then
            If buf(nach) = 32 Or buf(nach) = 62 Or buf(nach) = 13 Or buf(nach) = 10 Or buf(nach) = 9 Then
                FindTag = x
                Exit Function
            End If
        End If
        von = x + 1
    Loop
End Function

Private Function MatchAt(buf() As Byte, ByVal p As Long, pat As Variant) As Boolean
    Dim j As Long
    If p < 0 Or p + UBound(pat) > UBound(buf) Then Exit Function
    For j = 0 To UBound(pat)
        If buf(p + j) <> pat(j) Then Exit Function
    Next j
    MatchAt = True
End Function

Private Function LeseBisAnfuehrung(buf() As Byte, ByVal p As Long) As String
    Dim e As Long
    e = p
    Do While e <= UBound(buf)
        If buf(e) = 34 Then Exit Do
        If e - p > 200 Then Fehler "Unerwarteter Dateiaufbau an Byte " & p & "."
        e = e + 1
    Loop
    If e > UBound(buf) Then Fehler "Unerwartetes Dateiende."
    LeseBisAnfuehrung = LeseText(buf, p, e - p)
End Function

Private Function LeseText(buf() As Byte, ByVal p As Long, ByVal anzahl As Long) As String
    Dim i As Long, s As String
    For i = p To p + anzahl - 1
        If i > UBound(buf) Then Exit For
        s = s & ChrW(buf(i))
    Next i
    LeseText = s
End Function

' Text (nur ASCII) in ein Byte-Array (Variant) umwandeln
Private Function ToBytes(ByVal t As String) As Variant
    Dim b() As Byte, i As Long
    ReDim b(0 To Len(t) - 1)
    For i = 1 To Len(t)
        b(i - 1) = AscW(Mid$(t, i, 1)) And &HFF
    Next i
    ToBytes = b
End Function

Private Sub InitMuster()
    P_HIST = ToBytes("<FM-TASK-HISTORY")
    P_HIST_ENDE = ToBytes("</FM-TASK-HISTORY>")
    P_SETS = ToBytes("<FM-TASK-SETS")
    P_SETS_ENDE = ToBytes("</FM-TASK-SETS>")
    P_SET1 = ToBytes("<FM-TASK-SET ")
    P_SET2 = ToBytes("<FM-TASK-SET>")
    P_DATUM = ToBytes("<DATE-1")
    P_DATUM_TAG = ToBytes(DATUM_TAG)
    P_DATUM_ENDE = ToBytes("</DATE-1>")
    P_OCC = ToBytes("<FM-OCCURRENCE-TASKS>")
    P_OCC_ENDE = ToBytes("</FM-OCCURRENCE-TASKS>")
    P_DET = ToBytes("<FM-DETECTION-TASKS>")
    P_DET_ENDE = ToBytes("</FM-DETECTION-TASKS>")
    P_REF = ToBytes("<FM-ACTION-REF ID-REF=""")
    P_ACTIONS = ToBytes("<FM-ACTIONS>")
    P_ACTIONS_ENDE = ToBytes("</FM-ACTIONS>")
    P_ACTION = ToBytes("<FM-ACTION ID=""")
    P_ACTION_ENDE = ToBytes("</FM-ACTION>")
    P_STATE = ToBytes("<FM-ACTION-STATE>")
    P_HISTSTATES = ToBytes("<FM-HISTORY-STATES>")
    P_SDGS = ToBytes("<SDGS>")
    P_SCHEDULE = ToBytes("<FM-TASK-SCHEDULE>")
End Sub

Private Sub Fehler(ByVal nachricht As String)
    Err.Raise FEHLER_NR, "FMEA_Massnahmenstand_Datum", nachricht
End Sub

' "30.04.2027" -> Datum. Gibt False bei ungueltiger Eingabe zurueck.
Private Function ParseDatum(ByVal t As String, ByRef d As Date) As Boolean
    Dim teile() As String, i As Long, tg As Long, mo As Long, jr As Long
    t = Trim$(t)
    teile = Split(t, ".")
    If UBound(teile) <> 2 Then Exit Function
    For i = 0 To 2
        If Not NurZiffern(teile(i)) Then Exit Function
    Next i
    If Len(teile(0)) > 2 Or Len(teile(1)) > 2 Or Len(teile(2)) <> 4 Then Exit Function
    tg = CLng(teile(0)): mo = CLng(teile(1)): jr = CLng(teile(2))
    If jr < 2000 Or jr > 2099 Or mo < 1 Or mo > 12 Or tg < 1 Or tg > 31 Then Exit Function
    d = DateSerial(jr, mo, tg)
    If Day(d) <> tg Or Month(d) <> mo Or Year(d) <> jr Then Exit Function
    ParseDatum = True
End Function

Private Function NurZiffern(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        If Not (Mid$(t, i, 1) Like "[0-9]") Then Exit Function
    Next i
    NurZiffern = True
End Function

' Zeitzone Europe/Berlin um 23:59:59 Uhr des Tages:
' Sommerzeit vom letzten Sonntag im Maerz bis zum Tag vor dem letzten Sonntag im Oktober.
Private Function BerlinOffset(ByVal d As Date) As String
    Dim jr As Long, maerz As Date, okt As Date
    jr = Year(d)
    maerz = DateSerial(jr, 3, 31)
    maerz = maerz - (Weekday(maerz, vbSunday) - 1)
    okt = DateSerial(jr, 10, 31)
    okt = okt - (Weekday(okt, vbSunday) - 1)
    If d >= maerz And d < okt Then BerlinOffset = "+02:00" Else BerlinOffset = "+01:00"
End Function

Private Function IsoDatum(ByVal d As Date) As String
    IsoDatum = Year(d) & "-" & Right$("0" & Month(d), 2) & "-" & Right$("0" & Day(d), 2)
End Function

Private Function DeutschesDatum(ByVal d As Date) As String
    DeutschesDatum = Right$("0" & Day(d), 2) & "." & Right$("0" & Month(d), 2) & "." & Year(d)
End Function

Private Function AusgabePfad(ByVal inPfad As String, ByVal d As Date) As String
    Dim punkt As Long, slash As Long
    punkt = InStrRev(inPfad, ".")
    slash = InStrRev(inPfad, "\")
    If InStrRev(inPfad, "/") > slash Then slash = InStrRev(inPfad, "/")
    If punkt > slash Then
        AusgabePfad = Left$(inPfad, punkt - 1) & "_" & IsoDatum(d) & Mid$(inPfad, punkt)
    Else
        AusgabePfad = inPfad & "_" & IsoDatum(d) & ".xml"
    End If
End Function

' Umlaute ohne Sonderzeichen im Quelltext: {ae} {oe} {ue} {Ae} {Oe} {Ue} {ss}
Private Function Txt(ByVal s As String) As String
    s = Replace(s, "{ae}", ChrW(228))
    s = Replace(s, "{oe}", ChrW(246))
    s = Replace(s, "{ue}", ChrW(252))
    s = Replace(s, "{Ae}", ChrW(196))
    s = Replace(s, "{Oe}", ChrW(214))
    s = Replace(s, "{Ue}", ChrW(220))
    s = Replace(s, "{ss}", ChrW(223))
    Txt = s
End Function
