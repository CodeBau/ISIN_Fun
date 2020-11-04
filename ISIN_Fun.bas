Attribute VB_Name = "ISIN_Fun"
Function isin(c As String, b As String, d As String) As String
      
      Dim a As String
      Dim dlugosc_tekstu As Integer
      
      a = b & " " & c & " " & d
      
      dlugosc_tekstu = Len(a)
      
      Dim tablica_znakow(99) As String
      Dim ii As Integer
      Dim jj As Integer
      ii = 0
      Dim c1 As String
      Dim c2 As String
      Dim c3 As String
      Dim c4 As String
      Dim c5 As String
      Dim c6 As String
      Dim c7 As String
      Dim c8 As String
      Dim c9 As String
      Dim c10 As String
      Dim c11 As String
      Dim c12 As String
      Dim c1c2 As String
      Dim i_isin As String
      
      Dim c1_int As String
      Dim c2_int As String
      Dim c3_int As String
      Dim c4_int As String
      Dim c5_int As String
      Dim c6_int As String
      Dim c7_int As String
      Dim c8_int As String
      Dim c9_int As String
      Dim c10_int As String
      Dim c11_int As String
      
      Dim liczby1 As String
      Dim rev_liczby1 As String
      Dim dl_rev_liczby1 As String
      
      Dim parzyste_suma As Integer
      Dim nieparzyste As String
      Dim nieparzyste_suma As Integer
      
      Dim dl_nieparzyste As String
      Dim suma As Integer
      Dim mod10 As Integer
      Dim check As Integer
      Dim dupa As Integer
For i = 1 To dlugosc_tekstu
        
        c1 = Mid(a, i, 1)
        c2 = Mid(a, i + 1, 1)
        c3 = Mid(a, i + 2, 1)
        c4 = Mid(a, i + 3, 1)
        c5 = Mid(a, i + 4, 1)
        c6 = Mid(a, i + 5, 1)
        c7 = Mid(a, i + 6, 1)
        c8 = Mid(a, i + 7, 1)
        c9 = Mid(a, i + 8, 1)
        c10 = Mid(a, i + 9, 1)
        c11 = Mid(a, i + 10, 1)
        c12 = Mid(a, i + 11, 1)
        c1c2 = c1 + c2
        i_isin = c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10 + c11 + c12
        
        
        
If (c1c2 = "AD" Or c1c2 = "AE" Or c1c2 = "AF" Or c1c2 = "AG" Or c1c2 = "AI" Or c1c2 = "AL" Or c1c2 = "AM" Or c1c2 = "AN" Or c1c2 = "AO" Or c1c2 = "AQ" Or c1c2 = "AR" Or c1c2 = "AS" Or c1c2 = "AT" Or c1c2 = "AU" Or c1c2 = "AW" Or c1c2 = "AZ" Or c1c2 = "BA" Or c1c2 = "BB" Or c1c2 = "BD" Or c1c2 = "BE" Or c1c2 = "BF" Or c1c2 = "BG" Or c1c2 = "BH" Or c1c2 = "BI" Or c1c2 = "BJ" Or c1c2 = "BM" Or c1c2 = "BN" Or c1c2 = "BO" Or c1c2 = "BR" Or c1c2 = "BS" Or c1c2 = "BT" Or c1c2 = "BV" Or c1c2 = "BW" Or c1c2 = "BY" Or c1c2 = "BZ" Or c1c2 = "CA" Or c1c2 = "CC" Or c1c2 = "CD" Or c1c2 = "CF" Or c1c2 = "CG" Or c1c2 = "CH" Or c1c2 = "CI" Or c1c2 = "CK" Or c1c2 = "CL" Or c1c2 = "CM" Or c1c2 = "CN" Or c1c2 = "CO" Or c1c2 = "CR" Or c1c2 = "CU" Or c1c2 = "CV" Or c1c2 = "CX" Or c1c2 = "CY" Or c1c2 = "CZ" Or c1c2 = "DE" Or c1c2 = "DJ" Or c1c2 = "DK" Or c1c2 = "DM" Or c1c2 = "DO" Or c1c2 = "DZ" Or c1c2 = "EC" Or c1c2 = "EE" Or c1c2 = "EG" Or c1c2 = "ER" Or c1c2 = "ES" Or c1c2 = "ET" Or c1c2 = "FI" Or c1c2 = "FJ" Or c1c2 = "FK" _
 Or c1c2 = "FO" Or c1c2 = "FR" Or c1c2 = "GA" Or c1c2 = "GB" Or c1c2 = "GD" Or c1c2 = "GE" Or c1c2 = "GH" Or c1c2 = "GI" Or c1c2 = "GL" Or c1c2 = "GM" Or c1c2 = "GN" Or c1c2 = "GQ" Or c1c2 = "GR" Or c1c2 = "GS" Or c1c2 = "GT" Or c1c2 = "GU" Or c1c2 = "GW" Or c1c2 = "GY" Or c1c2 = "HK" Or c1c2 = "HM" Or c1c2 = "HN" Or c1c2 = "HR" Or c1c2 = "HT" Or c1c2 = "HU" Or c1c2 = "ID" Or c1c2 = "IE" Or c1c2 = "IL" Or c1c2 = "IN" Or c1c2 = "IO" Or c1c2 = "IQ" Or c1c2 = "IR" Or c1c2 = "IS" Or c1c2 = "IT" Or c1c2 = "JM" Or c1c2 = "JO" Or c1c2 = "JP" Or c1c2 = "KE" Or c1c2 = "KG" Or c1c2 = "KH" Or c1c2 = "KI" Or c1c2 = "KM" Or c1c2 = "KN" Or c1c2 = "KP" Or c1c2 = "KR" Or c1c2 = "KW" Or c1c2 = "KY" Or c1c2 = "KZ" Or c1c2 = "LA" Or c1c2 = "LB" Or c1c2 = "LC" Or c1c2 = "LI" Or c1c2 = "LK" Or c1c2 = "LR" Or c1c2 = "LS" Or c1c2 = "LT" Or c1c2 = "LU" Or c1c2 = "LV" Or c1c2 = "LY" Or c1c2 = "MA" Or c1c2 = "MD" Or c1c2 = "ME" Or c1c2 = "MG" Or c1c2 = "MH" Or c1c2 = "MK" Or c1c2 = "ML" Or c1c2 = "MM" Or c1c2 = "MN" Or c1c2 = "MO" _
 Or c1c2 = "MP" Or c1c2 = "MR" Or c1c2 = "MS" Or c1c2 = "MT" Or c1c2 = "MU" Or c1c2 = "MV" Or c1c2 = "MW" Or c1c2 = "MX" Or c1c2 = "MY" Or c1c2 = "MZ" Or c1c2 = "NA" Or c1c2 = "NC" Or c1c2 = "NE" Or c1c2 = "NF" Or c1c2 = "NG" Or c1c2 = "NI" Or c1c2 = "NL" Or c1c2 = "NO" Or c1c2 = "NP" Or c1c2 = "NR" Or c1c2 = "NU" Or c1c2 = "NZ" Or c1c2 = "OM" Or c1c2 = "PA" Or c1c2 = "PE" Or c1c2 = "PF" Or c1c2 = "PG" Or c1c2 = "PH" Or c1c2 = "PK" Or c1c2 = "PL" Or c1c2 = "PM" Or c1c2 = "PN" Or c1c2 = "PS" Or c1c2 = "PT" Or c1c2 = "PW" Or c1c2 = "PY" Or c1c2 = "QA" Or c1c2 = "QR" Or c1c2 = "QV" Or c1c2 = "RO" Or c1c2 = "RU" Or c1c2 = "RW" Or c1c2 = "SA" Or c1c2 = "SB" Or c1c2 = "SC" Or c1c2 = "SD" Or c1c2 = "SE" Or c1c2 = "SG" Or c1c2 = "SH" Or c1c2 = "SI" Or c1c2 = "SK" Or c1c2 = "SL" Or c1c2 = "SM" Or c1c2 = "SN" Or c1c2 = "SO" Or c1c2 = "SR" Or c1c2 = "ST" Or c1c2 = "SV" Or c1c2 = "SY" Or c1c2 = "SZ" Or c1c2 = "TC" Or c1c2 = "TD" Or c1c2 = "TF" Or c1c2 = "TG" Or c1c2 = "TH" Or c1c2 = "TJ" Or c1c2 = "TK" Or c1c2 = "TL" _
 Or c1c2 = "TM" Or c1c2 = "TN" Or c1c2 = "TO" Or c1c2 = "TR" Or c1c2 = "TT" Or c1c2 = "TV" Or c1c2 = "TW" Or c1c2 = "TZ" Or c1c2 = "UA" Or c1c2 = "UG" Or c1c2 = "UM" Or c1c2 = "US" Or c1c2 = "UY" Or c1c2 = "UZ" Or c1c2 = "VA" Or c1c2 = "VC" Or c1c2 = "VE" Or c1c2 = "VG" Or c1c2 = "VI" Or c1c2 = "VN" Or c1c2 = "VU" Or c1c2 = "WF" Or c1c2 = "WS" Or c1c2 = "XC" Or c1c2 = "XK" Or c1c2 = "XL" Or c1c2 = "XS" Or c1c2 = "YE" Or c1c2 = "YT" Or c1c2 = "ZA" Or c1c2 = "ZM" Or c1c2 = "ZW") Then

    If c12 = "0" Or c12 = "1" Or c12 = "2" Or c12 = "3" Or c12 = "4" Or c12 = "5" Or c12 = "6" Or c12 = "7" Or c12 = "8" Or c12 = "9" Then
      
      If c3 = "0" Or c3 = "1" Or c3 = "2" Or c3 = "3" Or c3 = "4" Or c3 = "5" Or c3 = "6" Or c3 = "7" Or c3 = "8" Or c3 = "9" Or c3 = "A" Or c3 = "B" Or c3 = "C" Or c3 = "D" Or c3 = "E" Or c3 = "F" Or c3 = "G" Or c3 = "H" Or c3 = "I" Or c3 = "J" Or c3 = "K" Or c3 = "L" Or c3 = "M" Or c3 = "N" Or c3 = "O" Or c3 = "P" Or c3 = "Q" Or c3 = "R" Or c3 = "S" Or c3 = "T" Or c3 = "U" Or c3 = "V" Or c3 = "W" Or c3 = "X" Or c3 = "Y" Or c3 = "Z" Then
      If c4 = "0" Or c4 = "1" Or c4 = "2" Or c4 = "3" Or c4 = "4" Or c4 = "5" Or c4 = "6" Or c4 = "7" Or c4 = "8" Or c4 = "9" Or c4 = "A" Or c4 = "B" Or c4 = "C" Or c4 = "D" Or c4 = "E" Or c4 = "F" Or c4 = "G" Or c4 = "H" Or c4 = "I" Or c4 = "J" Or c4 = "K" Or c4 = "L" Or c4 = "M" Or c4 = "N" Or c4 = "O" Or c4 = "P" Or c4 = "Q" Or c4 = "R" Or c4 = "S" Or c4 = "T" Or c4 = "U" Or c4 = "V" Or c4 = "W" Or c4 = "X" Or c4 = "Y" Or c4 = "Z" Then
      If c5 = "0" Or c5 = "1" Or c5 = "2" Or c5 = "3" Or c5 = "4" Or c5 = "5" Or c5 = "6" Or c5 = "7" Or c5 = "8" Or c5 = "9" Or c5 = "A" Or c5 = "B" Or c5 = "C" Or c5 = "D" Or c5 = "E" Or c5 = "F" Or c5 = "G" Or c5 = "H" Or c5 = "I" Or c5 = "J" Or c5 = "K" Or c5 = "L" Or c5 = "M" Or c5 = "N" Or c5 = "O" Or c5 = "P" Or c5 = "Q" Or c5 = "R" Or c5 = "S" Or c5 = "T" Or c5 = "U" Or c5 = "V" Or c5 = "W" Or c5 = "X" Or c5 = "Y" Or c5 = "Z" Then
      If c6 = "0" Or c6 = "1" Or c6 = "2" Or c6 = "3" Or c6 = "4" Or c6 = "5" Or c6 = "6" Or c6 = "7" Or c6 = "8" Or c6 = "9" Or c6 = "A" Or c6 = "B" Or c6 = "C" Or c6 = "D" Or c6 = "E" Or c6 = "F" Or c6 = "G" Or c6 = "H" Or c6 = "I" Or c6 = "J" Or c6 = "K" Or c6 = "L" Or c6 = "M" Or c6 = "N" Or c6 = "O" Or c6 = "P" Or c6 = "Q" Or c6 = "R" Or c6 = "S" Or c6 = "T" Or c6 = "U" Or c6 = "V" Or c6 = "W" Or c6 = "X" Or c6 = "Y" Or c6 = "Z" Then
      If c7 = "0" Or c7 = "1" Or c7 = "2" Or c7 = "3" Or c7 = "4" Or c7 = "5" Or c7 = "6" Or c7 = "7" Or c7 = "8" Or c7 = "9" Or c7 = "A" Or c7 = "B" Or c7 = "C" Or c7 = "D" Or c7 = "E" Or c7 = "F" Or c7 = "G" Or c7 = "H" Or c7 = "I" Or c7 = "J" Or c7 = "K" Or c7 = "L" Or c7 = "M" Or c7 = "N" Or c7 = "O" Or c7 = "P" Or c7 = "Q" Or c7 = "R" Or c7 = "S" Or c7 = "T" Or c7 = "U" Or c7 = "V" Or c7 = "W" Or c7 = "X" Or c7 = "Y" Or c7 = "Z" Then
      If c8 = "0" Or c8 = "1" Or c8 = "2" Or c8 = "3" Or c8 = "4" Or c8 = "5" Or c8 = "6" Or c8 = "7" Or c8 = "8" Or c8 = "9" Or c8 = "A" Or c8 = "B" Or c8 = "C" Or c8 = "D" Or c8 = "E" Or c8 = "F" Or c8 = "G" Or c8 = "H" Or c8 = "I" Or c8 = "J" Or c8 = "K" Or c8 = "L" Or c8 = "M" Or c8 = "N" Or c8 = "O" Or c8 = "P" Or c8 = "Q" Or c8 = "R" Or c8 = "S" Or c8 = "T" Or c8 = "U" Or c8 = "V" Or c8 = "W" Or c8 = "X" Or c8 = "Y" Or c8 = "Z" Then
      If c9 = "0" Or c9 = "1" Or c9 = "2" Or c9 = "3" Or c9 = "4" Or c9 = "5" Or c9 = "6" Or c9 = "7" Or c9 = "8" Or c9 = "9" Or c9 = "A" Or c9 = "B" Or c9 = "C" Or c9 = "D" Or c9 = "E" Or c9 = "F" Or c9 = "G" Or c9 = "H" Or c9 = "I" Or c9 = "J" Or c9 = "K" Or c9 = "L" Or c9 = "M" Or c9 = "N" Or c9 = "O" Or c9 = "P" Or c9 = "Q" Or c9 = "R" Or c9 = "S" Or c9 = "T" Or c9 = "U" Or c9 = "V" Or c9 = "W" Or c9 = "X" Or c9 = "Y" Or c9 = "Z" Then
      If c10 = "0" Or c10 = "1" Or c10 = "2" Or c10 = "3" Or c10 = "4" Or c10 = "5" Or c10 = "6" Or c10 = "7" Or c10 = "8" Or c10 = "9" Or c10 = "A" Or c10 = "B" Or c10 = "C" Or c10 = "D" Or c10 = "E" Or c10 = "F" Or c10 = "G" Or c10 = "H" Or c10 = "I" Or c10 = "J" Or c10 = "K" Or c10 = "L" Or c10 = "M" Or c10 = "N" Or c10 = "O" Or c10 = "P" Or c10 = "Q" Or c10 = "R" Or c10 = "S" Or c10 = "T" Or c10 = "U" Or c10 = "V" Or c10 = "W" Or c10 = "X" Or c10 = "Y" Or c10 = "Z" Then
      If c11 = "0" Or c11 = "1" Or c11 = "2" Or c11 = "3" Or c11 = "4" Or c11 = "5" Or c11 = "6" Or c11 = "7" Or c11 = "8" Or c11 = "9" Or c11 = "A" Or c11 = "B" Or c11 = "C" Or c11 = "D" Or c11 = "E" Or c11 = "F" Or c11 = "G" Or c11 = "H" Or c11 = "I" Or c11 = "J" Or c11 = "K" Or c11 = "L" Or c11 = "M" Or c11 = "N" Or c11 = "O" Or c11 = "P" Or c11 = "Q" Or c11 = "R" Or c11 = "S" Or c11 = "T" Or c11 = "U" Or c11 = "V" Or c11 = "W" Or c11 = "X" Or c11 = "Y" Or c11 = "Z" Then
         
         c1_int = Asc(c1) - 55
         
         c2_int = Asc(c2) - 55
         
         If c3 = "0" Or c3 = "1" Or c3 = "2" Or c3 = "3" Or c3 = "4" Or c3 = "5" Or c3 = "6" Or c3 = "7" Or c3 = "8" Or c3 = "9" Then
            c3_int = c3
            
           Else
            c3_int = Asc(c3) - 55
         'c3int
         End If
         
         If c4 = "0" Or c4 = "1" Or c4 = "2" Or c4 = "3" Or c4 = "4" Or c4 = "5" Or c4 = "6" Or c4 = "7" Or c4 = "8" Or c4 = "9" Then
            c4_int = c4
            
           Else
            c4_int = Asc(c4) - 55
         'c4int
         End If
         
         If c5 = "0" Or c5 = "1" Or c5 = "2" Or c5 = "3" Or c5 = "4" Or c5 = "5" Or c5 = "6" Or c5 = "7" Or c5 = "8" Or c5 = "9" Then
            c5_int = c5
            
           Else
            c5_int = Asc(c5) - 55
         'c5int
         End If
         
         If c6 = "0" Or c6 = "1" Or c6 = "2" Or c6 = "3" Or c6 = "4" Or c6 = "5" Or c6 = "6" Or c6 = "7" Or c6 = "8" Or c6 = "9" Then
            c6_int = c6
            
           Else
            c6_int = Asc(c6) - 55
         'c6int
         End If
         
         If c7 = "0" Or c7 = "1" Or c7 = "2" Or c7 = "3" Or c7 = "4" Or c7 = "5" Or c7 = "6" Or c7 = "7" Or c7 = "8" Or c7 = "9" Then
            c7_int = c7
           Else
            c7_int = Asc(c7) - 55
         'c7int
         End If
         
         If c8 = "0" Or c8 = "1" Or c8 = "2" Or c8 = "3" Or c8 = "4" Or c8 = "5" Or c8 = "6" Or c8 = "7" Or c8 = "8" Or c8 = "9" Then
            c8_int = c8

           Else
            c8_int = Asc(c8) - 55
         'c8int
         End If
         
         If c9 = "0" Or c9 = "1" Or c9 = "2" Or c9 = "3" Or c9 = "4" Or c9 = "5" Or c9 = "6" Or c9 = "7" Or c9 = "8" Or c9 = "9" Then
            c9_int = c9

           Else
            c9_int = Asc(c9) - 55
         'c9int
         End If
         
         If c10 = "0" Or c10 = "1" Or c10 = "2" Or c10 = "3" Or c10 = "4" Or c10 = "5" Or c10 = "6" Or c10 = "7" Or c10 = "8" Or c10 = "9" Then
            c10_int = c10

           Else
            c10_int = Asc(c10) - 55
         'c10int
         End If
         
         If c11 = "0" Or c11 = "1" Or c11 = "2" Or c11 = "3" Or c11 = "4" Or c11 = "5" Or c11 = "6" Or c11 = "7" Or c11 = "8" Or c11 = "9" Then
            c11_int = c11

           Else
            c11_int = Asc(c11) - 55
         'c11int
         End If
         
         
         liczby1 = c1_int + c2_int + c3_int + c4_int + c5_int + c6_int + c7_int + c8_int + c9_int + c10_int + c11_int
         rev_liczby1 = StrReverse(liczby1)
         dl_rev_liczby1 = Len(rev_liczby1)
         
         parzyste_suma = 0
         nieparzyste = ""

         For j = 1 To dl_rev_liczby1 Step 2
         jj = 2 * Mid(rev_liczby1, j, 1)
         nieparzyste = nieparzyste & jj
         Next j

         nieparzyste_suma = 0
         dl_nieparzyste = Len(nieparzyste)
         
         For e = 1 To dl_nieparzyste
         nieparzyste_suma = nieparzyste_suma + Mid(nieparzyste, e, 1)
         Next e
        
         For f = 2 To dl_rev_liczby1 Step 2
         parzyste_suma = parzyste_suma + Mid(rev_liczby1, f, 1)
         Next f
        
         suma = nieparzyste_suma + parzyste_suma
         mod10 = Application.WorksheetFunction.RoundUp(suma / 10, 0) * 10 - suma
         
         check = mod10 - c12
         
         
            If check = 0 Then
            
            tablica_znakow(ii) = i_isin
            ii = ii + 1
            
           End If
         
         
      'c11
      End If
      'c10
      End If
      'c9
      End If
      'c8
      End If
      'c7
      End If
      'c6
      End If
      'c5
      End If
      'c4
      End If
      'c3
      End If
    End If
End If
 
Next i
      
      If tablica_znakow(1) <> "" Then
      tablica_znakow(0) = ""
      End If
      
     isin = tablica_znakow(0)
     
End Function
