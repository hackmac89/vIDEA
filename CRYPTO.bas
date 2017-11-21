Attribute VB_Name = "CRYPTO"
'  _________________________________________________________________________________________________
'//                                  ________..:: CRYPTO ::..________                               \\
'\\_________________________________________________________________________________________________//

'========================================================================
'########################################################################
'#  VBA Access Modul, welches einen                                     #
'#  Cryptoalgorithmus inkl. Anwendungsfunktion implementiert hat.       #
'#----------------------------------------------------------------------#
'#    _________________________________________                         #
'#   |*****************************************|                        #
'#   | Neue Ver- und Entschlüsselungsroutinen  |                        #
'#   | Algorithmus : IDEA (128-Bit)            |                        #
'#   |*****************************************|                        #
'#   |version : 0.3                            |                        #
'#   |_________________________________________|                        #
'#                                                                      #
'#----------------------------------------------------------------------#
'########################################################################
'#  TODO : - Bug bei Binärcodierten Dateien in der Funktion             #
'#           "FileEncryptDecrypt" beheben (Textdateien funktionieren)   #
'########################################################################
'# //..:: Aufbau des Moduls ::..\\                                      #
'# -------------------------------                                      #
'# --- StrEncryptDecrypt(...)                                           #
'# --- FileEncryptDecrypt(...)                                          #
'#       ||                                                             #
'#        ==> Init() - Prüfen, ob ver- oder entschl. werden soll        #
'#             ||      (in Abhängigk. davon evtl. Schlüssel invertieren)#
'#             ||                                                       #
'#              ==> CreateKey() - Schlüssel aufbereiten                 #
'#              ==> invertKey() - Schlüssel invertieren (entschlüsseln) #
'#                     ||                                               #
'#                      ==> inv(...)   - Hilfsfunktion von "invertKey"  #
'#       ||                                                             #
'#        ==> coreCrypt(...) - Hauptfunktion der symmetrischen Verschl. #
'#             ||                                                       #
'#              ==> ideaMul(...) - Multiplikation Modulo (2^16)+1       #
'#        ==> Der Rest sind ANDs, ORs, XORs und die Hilfsfunktionen...  #
'#                                                                      #
'########################################################################
'========================================================================

Option Compare Database

' GLOBALE VARIABLEN
' Teilschlüssel für VER-UND ENTSCHLÜSSELUNG (Global, damit d. Array nicht durch die ganzen Hilfsfunktionen durchgereicht werden muss.)
Private ks(51) As Integer
Private Const RUNDEN = 8
Private Const BLOCK_SIZE = 8
Private Const KEY_LENGTH = 16
Private Const INTERNAL_KEY_LENGTH = 51
Private key(KEY_LENGTH) As Byte
Private noCritialError As Boolean
         
' -##################################################-
'|   "Shortcuts" für die Datentyp-Konvertierungen     |
'|   --------------------------------------------     |
'|                                                    |
'|    Integer %                                       |
'|    Long &                                          |
'|    Currency @                                      |
'|    Single !                                        |
'|    Double #                                        |
'|    String $                                        |
'|   --------------------------------------------     |
'|                                                    |
' -##################################################-

' ========================================
'||__--:: HILFSFUNKTIONEN - ANFANG ::--__||
' ========================================

' Bitweiser Rechts-Shift (versch. Datentypen)
Private Function shrByte(ByVal Value As Byte, ByVal Shift As Byte) As Byte
    shrByte = Value

    If Shift > 0 Then
        shrByte = Int(shrByte / (2 ^ Shift))
    End If
End Function

Private Function shrInteger(ByVal Value As Integer, ByVal Shift As Byte) As Integer
    shrInteger = Value
    
    If Shift > 0 Then
        If Value > 0 Then
            shrInteger = Int(shrInteger / (2 ^ Shift))
        Else
            If Shift > 15 Then
                shrInteger = 0
            Else
                shrInteger = shrInteger And &H7FFF
                shrInteger = Int(shrInteger / (2 ^ Shift))
                shrInteger = shrInteger Or 2 ^ (15 - Shift)
            End If
        End If
    End If
End Function

Private Function shrLong(ByVal Value As Long, ByVal Shift As Byte) As Long
    shrLong = Value
    
    If Shift > 0 Then
        If Value > 0 Then
            shrLong = Int(shrLong / (2 ^ Shift))
        Else
            If Shift > 31 Then
                shrLong = 0
            Else
                shrLong = shrLong And &H7FFFFFFF
                shrLong = Int(shrLong / (2 ^ Shift))
                shrLong = shrLong Or 2 ^ (31 - Shift)
            End If
        End If
    End If
End Function

'======================================================================================================================

' Bitweiser Links-Shift (versch. Datentypen)
Private Function shlByte(ByVal Value As Byte, ByVal Shift As Byte) As Byte
    Dim i As Byte
    Dim m As Byte

    shlByte = Value

    If Shift > 0 Then
        For i = 1 To Shift
            shlByte = (shlByte And &H7F) * 2
        Next i
    End If
End Function

Private Function shlInteger(ByVal Value As Integer, ByVal Shift As Byte) As Integer
    Dim i As Byte
    Dim m As Integer

    shlInteger = Value

    If Shift > 0 Then
        For i = 1 To Shift
            ' Das 14. Bit speichern
            m = shlInteger And &H4000
            ' Das 14. und 15. Bit löschen
            shlInteger = (shlInteger And &H3FFF) * 2
            If m <> 0 Then
                ' Das 14. Bit setzen
                shlInteger = shlInteger Or &H8000
            End If
        Next i
    End If
End Function

Private Function shlLong(ByVal Value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    Dim m As Long
    
    shlLong = Value
    
    If Shift > 0 Then
        For i = 1 To Shift
            ' Das 30. Bit speichern
            m = shlLong And &H40000000
            ' Das 30. und 31. Bit löschen
            shlLong = (shlLong And &H3FFFFFFF) * 2
            If m <> 0 Then
                ' Das 31. Bit setzen
                shlLong = shlLong Or &H80000000
            End If
        Next i
    End If
End Function

'======================================================================================================================

Private Function inc(ByRef val As Integer) As Integer
    val = val + 1
End Function

Private Function dec(ByRef val As Long) As Long
    val = val - 1
End Function

Private Function decInt(ByRef val As Integer) As Integer
    val = val - 1
End Function

' Hilfsfunktion für die Prozedur "FileEncryptDecrypt"
Private Function FilePathExists(ByVal strPath As String) As Boolean
  On Error Resume Next
  GetAttr strPath
  FilePathExists = (Err = 0)
End Function

' ======================================
'||__--:: HILFSFUNKTIONEN - ENDE ::--__||
' ======================================

'=========================================================================

' Vorbereitungsfunktion
Private Sub Init(decrypt As Boolean)
    noCritialError = True
    CreateKey
    If (decrypt) Then
        invertKey
    End If
End Sub

' Aufbereiten der Teilschlüssel
Private Sub CreateKey()
    ' Deklarationen / Initialisierungen
    Dim zOff As Integer: zOff = 0
    Dim i As Integer: i = 0
    
    ' DEBUG
    Debug.Print "[+] Aufbereiten der 8 Teilschlüssel ks..."
    
    On Error GoTo CreateSubKeyError
        ' Die 8 Teilschlüssel erzeugen
        ks(0) = CInt((shlInteger(key(0) And &HFF, 8)) Or (key(1) And &HFF))
        ks(1) = CInt((shlInteger(key(2) And &HFF, 8)) Or (key(3) And &HFF))
        ks(2) = CInt((shlInteger(key(4) And &HFF, 8)) Or (key(5) And &HFF))
        ks(3) = CInt((shlInteger(key(6) And &HFF, 8)) Or (key(7) And &HFF))
        ks(4) = CInt((shlInteger(key(8) And &HFF, 8)) Or (key(9) And &HFF))
        ks(5) = CInt((shlInteger(key(10) And &HFF, 8)) Or (key(11) And &HFF))
        ks(6) = CInt((shlInteger(key(12) And &HFF, 8)) Or (key(13) And &HFF))
        ks(7) = CInt((shlInteger(key(14) And &HFF, 8)) Or (key(15) And &HFF))
        
        For j = 8 To (INTERNAL_KEY_LENGTH)
            inc i
                                
            ks(i + 7 + zOff) = ((shlInteger(ks((i And 7) + zOff), 9) And &H7FFF&) - _
                               (shlInteger(ks((i And 7) + zOff), 9) And &H8000&)) Or _
                               (shrInteger(ks((i + 1 And 7) + zOff), 7) And &H1FF)
            zOff = zOff + (i And 8)
            i = (i And 7)
        Next
        
        Debug.Print "[+] Erzeugen/Aufbereiten der Teilschlüssel erfolgreich abgeschlossen !"
        
        ' DEBUG
        Debug.Print vbTab & "[*] i = " & i & " / 52 Teilschlüssel erzeugt"
        Debug.Print "[+] Generieren der nötigen Subkeys erfolgreich abgeschlossen !"
        
        GoTo Finalize
CreateSubKeyError:
    Debug.Print "[!] Es trat ein Fehler bei der Aufbereitung der Subkeys " & _
                "in der Funktion ""createKey"" auf. Breche ab..." & vbCrLf & vbTab & _
                "(MELDUNG: " & Error$ & ")"
    noCritialError = False
    GoTo Finalize
Finalize:
    ' Freigaben
    '...
    Debug.Print "[+] Destruktoren von ""createKey"" wurden erfolgreich ausgeführt."
End Sub

' Schlüssel zur Entschlüsselung invertieren
Private Sub invertKey()
    ' Deklarationen / Initialisierungen
    Dim i As Long, j As Long, k As Long
    Dim temp(INTERNAL_KEY_LENGTH) As Integer
    
    j = 4: k = INTERNAL_KEY_LENGTH
    
    Debug.Print "[+] Invertiere Key zum entschlüsseln."
    temp(k) = inv(ks(3)): dec k
    temp(k) = -((ks(2) And &H7FFF&) - (ks(2) And &H8000&)): dec k
    temp(k) = -((ks(1) And &H7FFF&) - (ks(1) And &H8000&)): dec k
    temp(k) = inv(ks(0)): dec k
    
    For i = 1 To RUNDEN - 1
        temp(k) = ks(j + 1): dec k
        temp(k) = ks(j): dec k
        temp(k) = inv(ks(j + 5)): dec k
        temp(k) = -((ks(j + 3) And &H7FFF&) - (ks(j + 3) And &H8000&)): dec k
        temp(k) = -((ks(j + 4) And &H7FFF&) - (ks(j + 4) And &H8000&)): dec k
        temp(k) = inv(ks(j + 2)): dec k
        j = j + 6
    Next i
    
    temp(k) = ks(j + 1): dec k
    temp(k) = ks(j): dec k
    temp(k) = inv(ks(j + 5)): dec k
    temp(k) = -((ks(j + 4) And &H7FFF&) - (ks(j + 4) And &H8000&)): dec k
    temp(k) = -((ks(j + 3) And &H7FFF&) - (ks(j + 3) And &H8000&)): dec k
    temp(k) = inv(ks(j + 2)): dec k
    
    For i = 0 To INTERNAL_KEY_LENGTH
        ks(i) = temp(i)
    Next i
    Debug.Print "[+] Key ist erfolgreich zum entschlüsseln invertiert worden."
End Sub

' Hilfsfunktion für "invertKey()"
Private Function inv(xx As Integer) As Integer
    ' Deklarationen / Initialisierungen
    Dim x As Long: x = xx And 65535
    Dim t0 As Long, t1 As Long, y As Long, q As Long
    
    If (x <= 1) Then
        inv = (x And &H7FFF&) - (x And &H8000&)
        Exit Function
    End If
    
    t1 = &H10001 \ x
    y = &H10001 Mod x
    
    If (y = 1) Then
        inv = ((1 - t1) And 32767) - ((1 - t1) And 32768)
        Exit Function
    End If
    
    t0 = 1
    
    While (y <> 1)
        q = x \ y
        x = x Mod y
        t0 = t0 + (q * t1)
        If (x = 1) Then
            inv = (t0 And 32767) - (t0 And 32768)
            Exit Function
        End If
        q = y \ x
        y = y Mod x
        t1 = t1 + (q * t0)
    Wend
    
    inv = ((1 - t1) And 32767) - ((1 - t1) And 32768)
End Function

' Multiplikation Modulo (2^16)+1
Private Function ideaMul(ByVal a As Long, ByVal b As Long) As Integer
    ' DEKLARATIONEN / INITIALISIERUNGEN
    a = a And &HFFFF&   ' Nur die unteren 16 Bits (short-Format)
    b = b And &HFFFF&
    
    Dim p As Long
    
    If (a <> 0) Then
        If (b <> 0) Then
            On Error GoTo NextDType
                p = (CLng(a) * CLng(b))
                GoTo Normal
NextDType:
            p = (CDbl(CDbl(a) * CDbl(b))) - 4294967296#
Normal:
            b = p And &HFFFF&   '65535
            a = shrLong(p, 16)
            ideaMul = CInt((b - a + (IIf(b < a, 1, 0))) And &H7FFF&) - ((b - a + (IIf(b < a, 1, 0))) And &H8000&)
        Else
            ideaMul = CInt(((1 - a) And &H7FFF&) - ((1 - a) And &H8000&))     ' dito
        End If
    Else
        ideaMul = CInt(((1 - b) And &H7FFF&) - ((1 - b) And &H8000&))      ' dito
    End If
End Function

' Kern der sym. Ver- und Entschlüsselung
Private Sub coreCrypt(v() As Byte, inOffset As Integer, ByRef out() As Byte, outOffset As Integer)
    ' DEKLARATIONEN / INITIALISIERUNGEN
    Dim x1 As Integer: x1 = (((shlInteger(v(inOffset + 0) And &HFF, 8)) Or (v(inOffset + 1) And &HFF) And &H7FFF&))
    Dim x2 As Integer: x2 = (((shlInteger(v(inOffset + 2) And &HFF, 8)) Or (v(inOffset + 3) And &HFF) And &H7FFF&))
    Dim x3 As Integer: x3 = (((shlInteger(v(inOffset + 4) And &HFF, 8)) Or (v(inOffset + 5) And &HFF) And &H7FFF&))
    Dim x4 As Integer: x4 = (((shlInteger(v(inOffset + 6) And &HFF, 8)) Or (v(inOffset + 7) And &HFF) And &H7FFF&))
    Dim s2 As Integer, s3 As Integer: s2 = 0: s3 = 0
    Dim i As Integer: i = 0
    Dim j As Integer: j = 0
    Dim runde As Integer: runde = RUNDEN
     
    ' DEBUG
    Debug.Print "[*] Datenblock 1 = " & Replace(Space(4 - Len(Hex(x1))), " ", "0") & Hex(x1)
    Debug.Print "[*] Datenblock 2 = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2)
    Debug.Print "[*] Datenblock 3 = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3)
    Debug.Print "[*] Datenblock 4 = " & Replace(Space(4 - Len(Hex(x4))), " ", "0") & Hex(x4)
    
    On Error GoTo coreCryptError
        While (runde > 0)
            Debug.Print (RUNDEN - runde) + 1 & ". Runde"
            x1 = ideaMul(x1, ks(i)): inc i
            x2 = ((CLng(x2) + ks(i)) And &H7FFF&) - ((CLng(x2) + ks(i)) And &H8000&): inc i
            x3 = ((CLng(x3) + ks(i)) And &H7FFF&) - ((CLng(x3) + ks(i)) And &H8000&): inc i
            x4 = ideaMul(x4, ks(i)): inc i
            
            'DEBUG
            Debug.Print "Ebene 1: Links = " & Replace(Space(4 - Len(Hex(x1))), " ", "0") & Hex(x1) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x4))), " ", "0") & Hex(x4)
            Debug.Print "Ebene 2: Links = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3)
                        
            s3 = x3
            x3 = CInt(x1 Xor x3)
            s2 = x2
            x2 = CInt(x2 Xor x4)
            
            'DEBUG
            Debug.Print "Ebene 3: Links = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2)
                        
            x3 = ideaMul(x3, ks(i)): inc i
            x2 = ((CLng(x3) + CLng(x2)) And &H7FFF&) - ((CLng(x3) + CLng(x2)) And &H8000&)
            
            'DEBUG
            Debug.Print "Ebene 4: Links = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2)
            
            x2 = ideaMul(x2, ks(i)): inc i
            x3 = ((CLng(x3) + CLng(x2)) And &H7FFF&) - ((CLng(x3) + CLng(x2)) And &H8000&)
            
            'DEBUG
            Debug.Print "Ebene 5: Links = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2)
                        
            ' XOR-STUFF
            x1 = x1 Xor x2: x4 = x4 Xor x3: x2 = x2 Xor s3: x3 = x3 Xor s2
            'DEBUG
            Debug.Print "Ebene 6: Links = " & Replace(Space(4 - Len(Hex(x2))), " ", "0") & Hex(x2) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x3))), " ", "0") & Hex(x3)
            Debug.Print "Ebene 7: Links = " & Replace(Space(4 - Len(Hex(x1))), " ", "0") & Hex(x1) & _
                        " Rechts = " & Replace(Space(4 - Len(Hex(x4))), " ", "0") & Hex(x4)
            
            decInt runde
        Wend
        
        ' Die letzte Teilrunde
        s2 = ideaMul(x1, ks(i)): inc i
        out(outOffset) = CByte(shrInteger(s2, 8)): inc outOffset
        out(outOffset) = CByte(Abs(s2 And &HFF)): inc outOffset
        s2 = ((CLng(x3) + ks(i)) And &H7FFF&) - ((CLng(x3) + ks(i)) And &H8000&): inc i
        
        out(outOffset) = CByte(shrInteger(s2, 8)): inc outOffset
        out(outOffset) = CByte(Abs(s2 And &HFF)): inc outOffset
        s2 = ((CLng(x2) + ks(i)) And &H7FFF&) - ((CLng(x2) + ks(i)) And &H8000&): inc i
        
        out(outOffset) = CByte(shrInteger(s2, 8)): inc outOffset
        out(outOffset) = CByte(Abs(s2 And &HFF)): inc outOffset
        s2 = ideaMul(x4, ks(i))
        
        out(outOffset) = CByte(shrInteger(s2, 8)): inc outOffset
        out(outOffset) = CByte(Abs(s2 And &HFF))
        
        GoTo Finalize
coreCryptError:
    Debug.Print "[!] Es trat ein Fehler bei der Berechnung der Ver- bzw. Entschlüsselung " & _
                "in der Funktion ""coreCrypt"" auf. Breche ab..." & vbCrLf & vbTab & _
                "(MELDUNG: " & Error$ & ")"
    noCritialError = False
    GoTo Finalize
Finalize:
    ' Freigaben
    '...
    Debug.Print "[+] Destruktoren von ""coreCrypt"" wurden erfolgreich ausgeführt."
End Sub

'################################################################################
'# WRAPPER FUNKTION, WELCHE ALS EINZIG VON AUSSEN SICHTBARE FUNKTION DIE        #
'# VER- UND ENTSCHLÜSSELUNG REGELT.                                             #
'# --------------------------------                                             #
'# Falls beim Kryptotext ein "-Zeichen auftritt, so muss dieses                 #
'# "speziell" behandelt werden, ansonsten wird an                               #
'# dieser/diesen Stelle/n der String abgeschnitten. Dafür muss VOR DEM AUFRUF   #
'# der Funktion Sorge getragen werden (GILT NUR FÜR DIE "DIREKTZEILE") !        #
'# --------------------------------                                             #
'# Zur besseren Übersicht KÖNNTE man die versch. Ausgabekodierungen             #
'# ("HEX" oder "TEXT") auch nochmal in 2 separate Funktionen                    #
'# ("StrEncryptDecryptHEX" und "StrEncryptDecryptTEXT") packen.                 #
'# --------------------------------                                             #
'# Params:                                                                      #
'#    @STRING                                                                   #
'#         Data        : Text/PW, welches ver- oder entschlüsselt werden soll   #
'#         InputModus  : "HEX" oder "TEXT"                                      #
'#         Schluessel  : Schlüssel zur Ver- oder Entschlüsselung                #
'#         Modus       : "ENCRYPT" oder "DECRYPT"                               #
'#         OutputModus : "HEX" oder "TEXT"                                      #
'#______________________________________________________________________________#
'################################################################################

' HAUPTFUNKTION (WRAPPER)
Public Function StrEncryptDecrypt(ByRef Data As String, ByRef InputModus As String, _
                                  ByRef Schluessel As String, ByVal Modus As String, _
                                  Optional ByVal OutputModus As String = "HEX") As String
    ' Deklarationen
    Dim v() As Byte
    Dim i As Integer, j As Integer
    Dim output() As Byte
    Dim testOutput As String: testOutput = ""
    
    ' Hier dann später auch noch padden, um variable Passwörter zu ermöglichen
    ' (ABER ACHTUNG : Passwörter mit MEHR als 16 Zeichen dürfen natürlich NICHT ERLAUBT sein !!!)
    If Len(Schluessel) <> 16 Then
        GoTo KeyError
    End If
    
    If Len(Data) = 0 Then
        GoTo DataError
    End If
    
    Select Case UCase(InputModus)
        Case "HEX":
                        ' Der Code fürs Padding ist leider etwas doppelt,
                        ' da bei den untersch. Modi auch untersch. Werte "gepaddet" werden
                        
                        ' Nur bei VERSCHLÜSSELUNG muss "gepaddet" werden, bei ENTSCHLÜSSELUNG ist das Padding ja mit drin.
                        If UCase(Modus) <> "DECRYPT" Then
                            If (Len(Data) Mod 8) <> 0 Then   ' PADDING
                                Debug.Print "[!] Muss den Plaintext padden..."
                                'For i = 0 To ((Len(Data) Mod 8) - 1) \ 2   ' "\" = Ganzzahlige Division
                                For i = 0 To 14 - (Len(Data) Mod 8) Step 2
                                    Data = Data & "20"
                                Next i
                                Debug.Print "[+] Padding erfolgreich." & vbCrLf & vbTab & "[*] data = " & Data
                            End If
                        End If
                        
                        Debug.Print "[+] Erstelle BYTE Array für Daten aus Hexadezimaltext..."
                        ReDim Preserve v(((Len(Data)) / 2) - 1)
                        
                        ' "Mid$" fängt die Indizierung bei 1 an -.-
                        j = 0
                        For i = 1 To Len(Data) Step 2
                            v(j) = CByte("&H" & Mid$(Data, i, 2))
                            inc j
                        Next i
        Case "TEXT":
                        ' Nur bei VERSCHLÜSSELUNG muss "gepaddet" werden, bei ENTSCHLÜSSELUNG ist das Padding ja mit drin.
                        If UCase(Modus) <> "DECRYPT" Then
                            If (Len(Data) Mod 8) <> 0 Then   ' PADDING
                                Debug.Print "[!] Muss den Plaintext padden..."
                                'For i = 0 To ((Len(Data) Mod 8) - 1) \ 2   ' "\" = Ganzzahlige Division
                                For i = 0 To 7 - (Len(Data) Mod 8)
                                    Data = Data & " "
                                Next i
                                Debug.Print "[+] Padding erfolgreich." & vbCrLf & vbTab & "[*] data = " & Data
                            End If
                        End If
                        
                        Debug.Print "[+] Erstelle BYTE Array für Daten aus Klartext..."
                        v = StrConv(Data, vbFromUnicode)
                        
                        ' HIER MÜSSTE FÜR DEN FEHLERHAFTEN SACHVERHALT BEI BINÄRCODIERTEN DATEIEN KORRIGIERT WERDEN
        Case Else:
                        Debug.Print "[!] FEHLER: Es wurde ein unbekannter EINGABEMODUS angegeben. Breche ab..."
                        Exit Function
    End Select
    
    ' UPDATE 09.02.
    ReDim Preserve output(((((UBound(v)) + 7) \ 8) * 8) - 1)   ' Array mit adäquater Blockgröße dyn. anpassen
    
    Debug.Print "[+] Erstelle BYTE Array für Schlüssel..."
    For i = 0 To Len(Schluessel) - 1 'Step 2
        key(i) = Asc(Mid$(Schluessel, i + 1, 1))
    Next i
    Debug.Print "[+] Die BYTE Arrays wurden erfolgreich initialisiert."
    
    Select Case UCase(Modus)
            Case "ENCRYPT":
                        Debug.Print "[+] Verschlüssele String..."
                        Debug.Print "[+] ..:: Verschlüsselungsmodus ::.."
                        Init (False)
            Case "DECRYPT":
                        Debug.Print "[+] Entschlüssele String..."
                        Debug.Print "[+] ..:: Entschlüsselungsmodus ::.."
                        Init (True)
            Case Else:
                        Debug.Print "[!] FEHLER: Es wurde ein unbekannter AUSGABEMODUS angegeben. Breche ab..."
                        Exit Function
    End Select
    
    ' Symmetrischen Algorithmus aufrufen
    ' ANMERKUNG: Falls mal große Dateien/Datenmengen (schon ab > 1MB) ver-/entschlüsselt werden sollen, muss
    '            das hier anders geregelt werden, denn durch das splitten auf ein komplettes Array
    '            (statt auf jeweils immer nur 8 Bytes zu arbeiten und gleich ver/-entschlüsselt schreiben)
    '            wird der Zähler i überlaufen bzw. ist in der jeweiligen Funktion zum Dateihandling vorher ein
    '            aufteilen (zeilenweises lesen etc.) von nöten.
    For i = 0 To (UBound(output) \ 8)
        coreCrypt v, i * 8, output, i * 8
    Next i
    
    If noCritialError Then
        ' Ausgabe
        Debug.Print " ======================="
        Debug.Print "||_____::AUSGABE::_____||"
        Debug.Print " ======================="
        
        Select Case UCase(OutputModus)
            Case "HEX":
                        For Each x In output
                            testOutput = testOutput & Replace(Space(2 - Len(Hex(CByte(x)))), " ", "0") & Hex(CByte(x))
                        Next
            Case "TEXT":
                        For Each x In output
                            testOutput = testOutput & Chr(CByte(x))
                        Next
            Case Else:
                        Debug.Print "[!] FEHLER: Es wurde ein unbekannter AUSGABEMODUS angegeben. Breche ab..."
                        Exit Function
        End Select

        Debug.Print "ERGEBNIS = " & RTrim$(testOutput)   ' Evtl. zuvor "gepaddete" Daten abschneiden
        StrEncryptDecrypt = RTrim$(testOutput)   ' Evtl. zuvor "gepaddete" Daten abschneiden
    End If
    
    GoTo Finalize
        
KeyError:
    Debug.Print "[!] FEHLER : Der Schlüssel muss aus 16 Zeichen bestehen !!!"
    End
DataError:
    Debug.Print "[!] FEHLER : Der Funktion wurde kein zu verschlüsselnder Klartext übergeben !!!"
    End
Finalize:
    Debug.Print "[+] Der Verschlüsselungs-Wrapper wurde erfolgreich ausgeführt."
End Function

' TODO (FILE-ENCRYPTION)
Public Sub FileEncryptDecrypt(ByRef FileInput As String, ByRef FileOutput As String, _
                              ByRef CryptMode As String, ByRef InputMode As String, _
                              ByRef Password As String, ByRef OutputMode As String)
    Dim File As String   ' Die Datei
    Dim Line As String   ' Zeilenweises einlesen ?
    Dim Str As String   ' Der Dateiinhalt
    Dim i As Long
    
    Str = ""
    
    ' Dialog anzeigen
    'File = Application.GetOpenFilename
    
    ' ODER Direktzuweisung
    'File = "C:\..."
    
    File = FileInput
    If FilePathExists(File) Then
        Open File For Binary As #1
        Open FileOutput For Binary As #2
            Str = Space(LOF(1))
            
            Get #1, , Str
            
            ' in 8 KB-Blöcken abarbeiten
            For i = 0 To FileLen(File) - 1 Step 8192
                If ((i + 8192) < FileLen(File)) Then
                    ' Ver- oder Entschlüsselte Daten in Ausgabedatei speichern
                    Put #2, , CRYPTO.StrEncryptDecrypt(Mid$(Str, i + 1, 8192), InputMode, Password, CryptMode, OutputMode)
                    ' Das Semikolon ";" unterbindet es "Print" eine abschliessende Leerzeile zu speichern (man könnte auch "Put" nehmen)
                    'Print #2, CRYPTO.StrEncryptDecrypt(Mid$(Str, i + 1, 128), InputMode, Password, CryptMode, OutputMode);
                Else
                    Put #2, , CRYPTO.StrEncryptDecrypt(Mid$(Str, i + 1, FileLen(File) - i), InputMode, Password, CryptMode, OutputMode)
                    ' Das Semikolon ";" unterbindet es "Print" eine abschliessende Leerzeile zu speichern (man könnte auch "Put" nehmen)
                    'Print #2, CRYPTO.StrEncryptDecrypt(Mid$(Str, i + 1, FileLen(File) - i), InputMode, Password, CryptMode, OutputMode);
                End If
            Next i
            Debug.Print "[+] D0NE !"
        Close #1   ' Datei(-handles) schließen
        Close #2
     
        Debug.Print "[+] Speichern der ver- oder entschlüsselten Datei abgeschlossen."
    Else
        Debug.Print "[!] FEHLER : Konnte die Datei/das Verzeichnis nicht finden."
    End If
End Sub

'  _________________________________________________________________________________________________
'//                                  ________..:: CRYPTO ::..________                               \\
'\\_________________________________________________________________________________________________//
