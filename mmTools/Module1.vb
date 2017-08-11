REM ********************************************************************************************************************************************************************
REM * 2015-03-17 Beginn Rework
REM * 2015-03-20 Übernehmen bereinigte Subs/Functionen aus mm_CorDskTrk
REM *            Bildschrimaufbereitung vereinheitlichen
REM *            Fehlerbehebungen FixLngInt|Str
REM *            Update Bilder
REM * 2015-03-22 Update Bilder
REM *            Mehrfachdefinition sd. bereinigen
REM *            Beginnen Update-Logik
REM * 2015-04-12 LstAllSng implementieren
REM * 2015-04-15 Dateinamenendungskorrekturfunktion einbauen
REM * 2015-05-24 Datenexport: Dateilokation/Name ausgeben
REM * 2015-07-18 FixLngInt: Format-String und formatieren String in zwei Variablen vorhalten
REM * 2016-03-17 Format-String in CurDat von yyyy-mm-dd auf yyyy-MM-dd umstellen
REM *            Format-String in CurTim von hh:mm:ss auf HH:mm:ss umstellen
REM * 2017-08-11 CSV-Listing für HandBase ausgeben
REM ********************************************************************************************************************************************************************

REM ********************************************************************************************************************************************************************
Imports System.Text
REM ********************************************************************************************************************************************************************

REM ********************************************************************************************************************************************************************
Imports SongsDB
REM ********************************************************************************************************************************************************************

REM ********************************************************************************************************************************************************************
Module Module1
   REM ****************************************************************************************************************************************************************

   REM ****************************************************************************************************************************************************************
   Public cnscLn As Integer = -1

   Const AlbNamLng = 40
   Const MusNamLng = 40
   Const MusCprLng = 40
   Const TrkNamLng = 40
   Const AlbArtLng = 2
   Const ArtNamLng = 40
   Const TrkNbrLng = 3
   Const TrkTimLng = 8
   Const FilExtLng = 4

   Public mm As SongsDB.SDBApplication = New SongsDB.SDBApplication()
   Public sl As New SDBSongList
   Public slCnt As Integer
   Public sd As New SDBSongData

   Public AlbArtSDB As New SDBAlbumArtList
   Public AlbArtItm As New SDBAlbumArtItem
   Public AlbArtCnt As Integer = Nothing
   Public AlbArtNam As String
   Public AlbArtAry() As SDBImage

   Public FilDrv As String = Nothing
   Public FilExt As String = Nothing
   REM ****************************************************************************************************************************************************************

   REM ****************************************************************************************************************************************************************
   Sub Main()

      mm.ShutdownAfterDisconnect = False
      Dim MnuSel As String = Nothing

      While MnuSel <> "99"

         MnuSel = Menu()

         sl = mm.AllVisibleSongList
         slCnt = sl.Count - 1

         Select Case MnuSel

            Case "q" : MnuSel = "99" : Exit Select
            Case "99" : Exit Select

            Case "01" : If Not ChgTrkDsk() Then MnuSel = "99"
            Case "02" : If Not CnsMscChr() Then MnuSel = "99"
            Case "03" : If Not CnsMscSyl() Then MnuSel = "99"
            Case "04" : If Not CorFilExt() Then MnuSel = "99"

            Case "10" : If Not CnsAlbMus() Then MnuSel = "99"
            Case "11" : If Not CnsTrkAlbMus() Then MnuSel = "99"
            Case "12" : If Not CnsTrkMus() Then MnuSel = "99"
            Case "13" : If Not CordbpCpr() Then MnuSel = "99"
            Case "14" : If Not CnsTrkCpr() Then MnuSel = "99"
            Case "15" : If Not CnsAlbTrkArt() Then MnuSel = "99"

            Case "20" : If Not Exp2HndBas() Then MnuSel = "99"

            Case "80" : If Not LstAllSng() Then MnuSel = "99"

            Case Else
               Console.WriteLine(" Unbekannte Auswahl, weiter mit Anykey")
               Console.ReadLine()

         End Select

      End While

   End Sub
   REM ****************************************************************************************************************************************************************

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Two digits for Disk- and Track-Number, where applicable
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function ChgTrkDsk() As Boolean

      Dim cslCnt As Long
      Dim Prcssd As Long = 0
      Dim lcPrlg As String = Prolog("ChgTrkDsk")

      Console.WriteLine()
      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         If (sd.TrackOrder < 10 And Left(sd.TrackOrderStr, 1) <> "0") _
         Or (sd.DiscNumber > 0 And sd.DiscNumber < 10 And Left(sd.DiscNumberStr, 1) <> "0") _
             Then
            Console.Write(pfx(cslCnt, slCnt) _
                        & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                        & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                        & "|" & FixLngStr(sd.Title, TrkNamLng) _
                        & "|" & FixLngStr(sd.AlbumArtistName, MusNamLng) _
                        & "|" & FixLngStr(sd.AlbumName, AlbNamLng)
                          )
            If (sd.TrackOrder < 10 And Left(sd.TrackOrderStr, 1) <> "0") Then sd.TrackOrderStr = LdgZ(sd.TrackOrder)
            If (sd.DiscNumber > 0 And sd.DiscNumber < 10 And Left(sd.DiscNumberStr, 1) <> "0") Then sd.DiscNumberStr = LdgZ(sd.DiscNumber)
            Console.WriteLine("|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                            & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                            & "|"
                             )
            sd.UpdateDB()
            sd.WriteTags()
            Prcssd += 1
         End If

         HrtBet(cslCnt, slCnt)

      Next

      Epilog(lcPrlg, Prcssd, slCnt)

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Clean up funny chars
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsMscChr() As Boolean

      Dim cslCnt As Long
      Dim sdSrch As String()
      Dim sdRplc As String()
      Dim sdElem As Integer = -1
      Dim csdTit As String

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "’" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "'"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "  " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Ã¤" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "ä"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Ã¶" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "ö"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Ã³" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "ó"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Ã©" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "é"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Ã¼" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "ü"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "ÃŸ" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "ß"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Â°" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "°"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "â€™" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "'"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "    " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " "
      '
      Dim sdSrchFound As Boolean = False

      Dim Prcssd As Long = 0

      Dim lcPrlg As String = Prolog("CnsMscChr")

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdTit = sd.Title
         Dim _1p As Boolean = False
         For cIndex As Integer = 0 To sdElem
            If csdTit.IndexOf(sdSrch(cIndex)) > 0 Then
               If Not _1p Then
                  Console.Write("|" & pfx(cslCnt, slCnt) _
                              & "|" & FixLngStr(sd.AlbumArtistName, MusNamLng) _
                              & "|" & FixLngStr(sd.AlbumName, AlbNamLng) _
                              & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                              & "|" & FixLngStr(csdTit, TrkNamLng)
                               )
                  _1p = True
               End If
               csdTit = Replace(csdTit, sdSrch(cIndex), sdRplc(cIndex))
            End If
         Next

         Console.WriteLine("|" & FixLngStr(csdTit, TrkNamLng) & "|")
         sd.Title = csdTit
         sd.UpdateDB()
         sd.WriteTags()
         Prcssd += 1

         HrtBet(cslCnt, slCnt)

      Next

      Epilog(lcPrlg, Prcssd, slCnt)

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Clean Up Capitalization and words commonly misspelled
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsMscSyl() As Boolean

      Dim cslCnt As Long
      Dim sdSrch As String()
      Dim sdRplc As String()
      Dim sdElem As Integer = -1
      Dim csdTit As String

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " A " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " a "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " An " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " an "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " At " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " at "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Be " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " be "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " By " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " by "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Es " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " es "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Im " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " im "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " In " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " in "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Is " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " is "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " It " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " it "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " my " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " My "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " No " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " no "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Of " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " of "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " On " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " on "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Or " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " or "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Up " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " up "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " So " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " so "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " us " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Us "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " To " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " to "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Zu " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " zu "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " All " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " all "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " And " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " and "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Bei " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " bei "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " can " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Can "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Das " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " das "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Der " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " der "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " For " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " for "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " her " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Her "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Ist " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " ist "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " I'M " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " I'm "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Not " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " not "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " The " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " the "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Too " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " too "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Und " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " und "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " No. " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " N°. "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " It's " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " it's "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " he's " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " He's "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " From " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " from "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " dont " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " don't "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " can'T " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " can't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Can'T " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " can't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Can't " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " can't "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " We'Ll " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " We'll "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " we'Ll " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " We'll "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " we'll " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " We'll "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " don't " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Don't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " won'T " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Won't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " won't " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Won't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Ain'T " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Ain't "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " ain'T " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " Ain't "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " Ersten " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " ersten "

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " U.S.A" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " U.S.A."
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " U.S.A " : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " U.S.A. "
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = " U.S.A.." : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = " U.S.A."

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Er hod graucht" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Er hod g'raucht"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Er Hod G'raucht" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Er hod g'raucht"

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Glamour & Pain" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Glamour and Pain"

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Live)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Live]"

      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Bonus)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Bonus Track]"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Bonustrack)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Bonus Track]"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Bonus track)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Bonus Track]"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Bonus Track)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Bonus Track]"

      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "([Instrumental])" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Instr." : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "[Inst.]" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Instr.)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Instrumental" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "(Instrumental)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      'sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "([Instrumental] MnucIndexx)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "(Instrumental MnucIndexx)"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "([[Instrumental]] Rough MnucIndexx)" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "(Instrumental Rough MnucIndexx)"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "([Instrumental])" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "[Instrumental]"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "[[" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "["
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "]]" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "]"


      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Beach Boy Stomp" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Beach Boys Stomp"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Amazing Grace , Nearer My God to Thee" : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Amazing Grace (Closer My God to Thee)"
      sdElem += 1 : ReDim Preserve sdSrch(sdElem) : sdSrch(sdElem) = "Brandenburgisches Konzert N°. 3 in G-Dur," : ReDim Preserve sdRplc(sdElem) : sdRplc(sdElem) = "Brandenburgisches Konzert N°. 3, G-Dur,"

      Dim sdSrchFound As Boolean = False

      Dim Prcssd As Long = 0

      Dim lcPrlg As String = Prolog("CnsMscSyl")

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdTit = sd.Title
         Dim _1p As Boolean = False
         For cIndex As Integer = 0 To sdElem
            If csdTit.IndexOf(sdSrch(cIndex)) > 0 Then
               If Not _1p Then
                  Console.Write("|" & pfx(cslCnt, slCnt) _
                              & "|" & FixLngStr(sd.AlbumArtistName, MusNamLng) _
                              & "|" & FixLngStr(sd.AlbumName, AlbNamLng) _
                              & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                              & "|" & FixLngStr(csdTit, TrkNamLng)
                               )
                  _1p = True
               End If
               csdTit = Replace(csdTit, sdSrch(cIndex), sdRplc(cIndex))
            End If
         Next

         Console.WriteLine("|" & FixLngStr(csdTit, TrkNamLng) & "|")
         sd.Title = csdTit
         sd.UpdateDB()
         sd.WriteTags()
         Prcssd += 1

         HrtBet(cslCnt, slCnt)

      Next

      Epilog(lcPrlg, Prcssd, cslCnt)

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Clean Up File extensions
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CorFilExt() As Boolean

      Dim cslCnt As Long
      Dim Prcssd As Long = 0
      Dim lcPrlg As String = Prolog("CorFilExt")

      Dim sdPathParts() As String
      Dim sdPathElems As Integer
      Dim sdPath As String
      Dim TxtStr As String

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         FilDrvExt(sd.Path)
         sdPathParts = Split(sd.Path, ".") : sdPathElems = sdPathParts.GetUpperBound(0)

         If sdPathParts(sdPathElems) = sdPathParts(sdPathElems - 1) Then

            sdPath = sdPathParts(0)
            For idx As Integer = 1 To sdPathElems - 1 : sdPath = sdPath & "." & sdPathParts(idx) : Next

            TxtStr = FixLngInt(cslCnt, slCnt) _
                   & "|" & Trim(sd.Path) _
                   & "|" & Trim(sdPath) _
                   & "|"
            Console.WriteLine(TxtStr)

            sd.Path = sdPath

            sd.UpdateDB()
            sd.WriteTags()
            Prcssd += 1

         End If

         HrtBet(cslCnt, slCnt)

      Next

      Epilog(lcPrlg, Prcssd, slCnt)

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Condolidate/Clean-Up Album-Artist (Driver)
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsAlbMus() As Boolean

      Dim cslCnt As Long

      Dim lcPrlg As String = Prolog("CnsAlbMus")

      Dim sdIndx As Integer()
      Dim cIndex As Integer = -1

      Dim csdAlb As String = Nothing
      Dim psdAlb As String = Nothing
      Dim psdTit As String = Nothing
      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing
      Dim cAlbMs As String = Nothing
      Dim pAlbMs As String = Nothing
      Dim dAlbMs As Boolean = False

      cIndex = -1
      ReDim sdIndx(cIndex)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdAlb = sd.AlbumName
         cAlbMs = sd.AlbumArtistName

         If csdAlb = psdAlb Then

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt
            If pAlbMs <> Nothing And pAlbMs <> cAlbMs Then dAlbMs = True

         Else

            If cIndex > -1 And dAlbMs Then
               If Not CnsAlbMusExc(slCnt, sdIndx) Then Return False
            Else
               cnscLn += 1 : Console.WriteLine(AlbHdr(cslCnt, slCnt, 0, 0) & "|")
            End If

            psdAlb = csdAlb
            pAlbMs = cAlbMs

            cIndex = -1
            ReDim sdIndx(cIndex)
            dAlbMs = False

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt

         End If

      Next

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Condolidate/Clean-Up Album-Artist (Executor)
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsAlbMusExc(slLstCt As Integer, sdIndx As Integer()) As Boolean

      CnsAlbMusExc = False

      Dim MaxcIndex As Integer = sdIndx.GetUpperBound(0)
      Dim wmxcIndex As Integer : If MaxcIndex < 99 Then wmxcIndex = 99 Else wmxcIndex = MaxcIndex

      Dim cIndex As Integer

      Header("Albumartisten richtig stellen")

      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing

      For cIndex = 0 To MaxcIndex

         sd = sl.Item(sdIndx(cIndex))

         cAlbNm = sd.AlbumName
         If cAlbNm <> pAlbNm And pAlbNm <> Nothing Then
            cnscLn += 1 : Console.WriteLine()
         End If
         pAlbNm = cAlbNm

         cnscLn += 1 : Console.WriteLine(AlbHdr(cIndex, wmxcIndex, sdIndx(cIndex), slLstCt) _
                                       & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                       & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                       & "|"
                                        )

      Next

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      REM  Datenlieferant&Empfänger
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim DtaSrcInt As Integer = EntDatSrc()
      If DtaSrcInt < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaSrcInt < -99 Then Return False Else Return True
      End If

      Dim DtaDrnInt() As Integer = EntDatDrn("A")
      If DtaDrnInt(0) < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaDrnInt(0) < -99 Then Return False Else Return True
      End If

      If DtaDrnInt(0) = -1 Then
         Dim z As Integer = -1
         For y = 0 To sdIndx.GetUpperBound(0)
            If y <> DtaSrcInt Then
               z += 1 : ReDim Preserve DtaDrnInt(z) : DtaDrnInt(z) = sdIndx(y)
            End If
         Next
      End If
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim SrcDta As New SDBSongData
      Dim DrnDta As New SDBSongData

      SrcDta = sl.Item(sdIndx(DtaSrcInt))

      For x = 0 To DtaDrnInt.GetUpperBound(0)

         Console.Write("|" & FixLngInt(DtaSrcInt, 999) _
                     & "|" & FixLngInt(sdIndx(DtaSrcInt), slLstCt) _
                     & "|" & SrcDta.AlbumArtistName _
                     & "|" & FixLngInt(x, 999) _
                     & "|" & FixLngInt(DtaDrnInt(x), slLstCt)
                     )

         DrnDta = sl.Item(DtaDrnInt(x))
         Console.Write("|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng))

         DrnDta.AlbumArtistName = ""
         DrnDta.AlbumArtistName = SrcDta.AlbumArtistName
         Console.WriteLine("|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng) & "|")

         DrnDta.UpdateDB()
         DrnDta.WriteTags()

      Next
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      cnscLn += 1 : Console.WriteLine("AnyKey")
      Console.ReadLine()
      Console.Clear()

      SrcDta = Nothing
      DrnDta = Nothing

      Return True
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Consolidate Track-Artist (Driver)
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkAlbMus() As Boolean

      Dim cslCnt As Long

      Dim lcPrlg As String = Prolog("CnsTrkAlbMus")

      Dim sdIndx As Integer()
      Dim cIndex As Integer = -1

      Dim csdTit As String = Nothing
      Dim psdTit As String = Nothing
      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing
      Dim cAlbMs As String = Nothing
      Dim pAlbMs As String = Nothing
      Dim dAlbMs As Boolean = False

      cIndex = -1
      ReDim sdIndx(cIndex)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdTit = sd.Title
         cAlbMs = sd.AlbumArtistName

         If csdTit = psdTit Then

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt
            If pAlbMs <> Nothing And pAlbMs <> cAlbMs Then dAlbMs = True
            pAlbMs = cAlbMs

         Else

            If cIndex > -1 And dAlbMs Then
               If Not CnsTrkAlbMusExc(cslCnt, sdIndx) Then Return False
            Else
               cnscLn += 1 : Console.WriteLine(AlbHdr(cslCnt, slCnt, 0, 0) _
                                             & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                             & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                             & "|"
                                              )
            End If

            psdTit = csdTit
            pAlbMs = cAlbMs

            cIndex = -1
            ReDim sdIndx(cIndex)
            dAlbMs = False

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt

         End If

      Next

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Consolidate Track-Artist (Executor)
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkAlbMusExc(slLstCt As Integer, sdIndx As Integer()) As Boolean

      CnsTrkAlbMusExc = False

      Dim MaxcIndex As Integer = sdIndx.GetUpperBound(0)
      Dim wmxcIndex As Integer : If MaxcIndex < 99 Then wmxcIndex = 99 Else wmxcIndex = MaxcIndex

      Dim sd As New SDBSongData
      Dim cIndex As Integer

      Header("Albumartisten für Track richtig stellen")

      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing

      For cIndex = 0 To MaxcIndex

         sd = sl.Item(sdIndx(cIndex))

         cAlbNm = sd.AlbumName
         If cAlbNm <> pAlbNm And pAlbNm <> Nothing Then
            cnscLn += 1 : Console.WriteLine()
         End If
         pAlbNm = cAlbNm

         cnscLn += 1 : Console.WriteLine(AlbHdr(cIndex, wmxcIndex, sdIndx(cIndex), slLstCt) _
                                       & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                       & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                       & "|" & FixLngStr(sd.ArtistName, MusNamLng) _
                                       & "|"
                                        )

      Next

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      REM  Datenlieferant&Empfänger
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim DtaSrcInt As Integer = EntDatSrc()
      If DtaSrcInt < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaSrcInt < -99 Then Return False Else Return True
      End If

      Dim DtaDrnInt() As Integer = EntDatDrn("A")
      If DtaDrnInt(0) < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaDrnInt(0) < -99 Then Return False Else Return True
      End If

      If DtaDrnInt(0) = -1 Then
         Dim z As Integer = -1
         For y = 0 To sdIndx.GetUpperBound(0)
            If y <> DtaSrcInt Then
               z += 1 : ReDim Preserve DtaDrnInt(z) : DtaDrnInt(z) = sdIndx(y)
            End If
         Next
      End If
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim SrcDta As New SDBSongData
      Dim DrnDta As New SDBSongData

      SrcDta = sl.Item(sdIndx(DtaSrcInt))

      For x = 0 To DtaDrnInt.GetUpperBound(0)

         DrnDta = sl.Item(sdIndx(DtaDrnInt(x)))

         Console.Write("|" & FixLngInt(DtaSrcInt, 999) _
                     & "|" & FixLngInt(sdIndx(DtaSrcInt), slLstCt) _
                     & "|" & FixLngStr(SrcDta.AlbumArtistName, MusNamLng) _
                     & "|" & FixLngInt(x, 999) _
                     & "|" & FixLngInt(sdIndx(DtaDrnInt(x)), slLstCt) _
                     & "|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng)
                      )

         DrnDta.ArtistName = ""
         DrnDta.ArtistName = SrcDta.ArtistName
         Console.WriteLine("|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng) & "|")

         DrnDta.UpdateDB()
         DrnDta.WriteTags()

      Next

      Console.ReadLine()

      SrcDta = Nothing
      DrnDta = Nothing

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkMus() As Boolean

      Dim cslCnt As Long

      Dim lcPrlg As String = Prolog("CnsTrkMus")

      Dim sdIndx As Integer()
      Dim cIndex As Integer = -1

      Dim csdTit As String = Nothing
      Dim psdTit As String = Nothing
      Dim csdMus As String = Nothing
      Dim psdMus As String = Nothing
      Dim dTrkMs As Boolean = False

      cIndex = -1
      ReDim sdIndx(cIndex)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdTit = sd.Title
         csdMus = sd.ArtistName

         If csdTit = psdTit Then

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt
            If psdMus <> Nothing And psdMus <> csdMus Then dTrkMs = True
            psdMus = csdMus

         Else

            If cIndex > -1 And dTrkMs Then
               If Not CnsTrkMusExc(cslCnt, sdIndx) Then Return False

            End If
            cnscLn += 1 : Console.WriteLine(AlbHdr(cslCnt, slCnt, 0, 0) _
                          & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                          & "|" & FixLngStr(sd.Title, TrkNamLng) _
                          & "|"
                           )

            psdTit = csdTit
            psdMus = csdMus

            cIndex = -1
            ReDim sdIndx(cIndex)
            dTrkMs = False

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt

         End If

      Next

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkMusExc(slLstCt As Integer, sdIndx As Integer()) As Boolean

      CnsTrkMusExc = False

      Dim MaxcIndex As Integer = sdIndx.GetUpperBound(0)
      Dim wmxcIndex As Integer : If MaxcIndex < 99 Then wmxcIndex = 99 Else wmxcIndex = MaxcIndex

      Dim sd As New SDBSongData
      Dim cIndex As Integer

      Header("Trackartisten für Track richtig stellen")

      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing

      For cIndex = 0 To MaxcIndex

         sd = sl.Item(sdIndx(cIndex))

         cAlbNm = sd.AlbumName
         If cAlbNm <> pAlbNm And pAlbNm <> Nothing Then
            cnscLn += 1 : Console.WriteLine()
         End If
         pAlbNm = cAlbNm

         cnscLn += 1 : Console.WriteLine(AlbHdr(cIndex, wmxcIndex, sdIndx(cIndex), slLstCt) _
                                       & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                       & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                       & "|" & FixLngStr(sd.ArtistName, MusNamLng) _
                                       & "|"
                                        )

      Next

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      REM  Datenlieferant&Empfänger
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim DtaSrcInt As Integer = EntDatSrc()
      If DtaSrcInt < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaSrcInt < -99 Then Return False Else Return True
      End If

      Dim DtaDrnInt() As Integer = EntDatDrn("A")
      If DtaDrnInt(0) < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaDrnInt(0) < -99 Then Return False Else Return True
      End If

      If DtaDrnInt(0) = -1 Then
         Dim z As Integer = -1
         For y = 0 To sdIndx.GetUpperBound(0)
            If y <> DtaSrcInt Then
               z += 1 : ReDim Preserve DtaDrnInt(z) : DtaDrnInt(z) = sdIndx(y)
            End If
         Next
      End If
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim SrcDta As New SDBSongData
      Dim DrnDta As New SDBSongData

      SrcDta = sl.Item(sdIndx(DtaSrcInt))

      For x = 0 To DtaDrnInt.GetUpperBound(0)

         DrnDta = sl.Item(sdIndx(DtaDrnInt(x)))

         Console.Write("|" & FixLngInt(DtaSrcInt, 999) _
                     & "|" & FixLngInt(sdIndx(DtaSrcInt), slLstCt) _
                     & "|" & FixLngStr(SrcDta.AlbumArtistName, MusNamLng) _
                     & "|" & FixLngInt(x, 999) _
                     & "|" & FixLngInt(sdIndx(DtaDrnInt(x)), slLstCt) _
                     & "|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng)
                      )

         DrnDta.ArtistName = ""
         DrnDta.ArtistName = SrcDta.ArtistName
         Console.WriteLine("|" & FixLngStr(DrnDta.AlbumArtistName, MusNamLng) & "|")

         DrnDta.UpdateDB()
         DrnDta.WriteTags()

      Next

      Console.ReadLine()

      SrcDta = Nothing
      DrnDta = Nothing

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CordbpCpr() As Boolean

      Dim cslCnt As Long
      Dim Prcssd As Long = 0

      Dim lcPrlg As String = Prolog("CordbpCpr")

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         If sd.MusicComposer = "" Then Continue For
         Dim io As Long = sd.MusicComposer.IndexOf("/")
         If io > 0 Then

            Console.Write("|" & pfx(cslCnt, slCnt) _
                        & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                        & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                        & "|" & FixLngStr(sd.Title, TrkNamLng) _
                        & "|" & FixLngStr(sd.AlbumArtistName, MusNamLng) _
                        & "|" & FixLngStr(sd.MusicComposer, AlbNamLng)
                         )
            sd.MusicComposer = Replace(sd.MusicComposer, "/", ";")
            Console.WriteLine("|" & "|" & FixLngStr(sd.MusicComposer, AlbNamLng) & "|")

            sd.UpdateDB()
            sd.WriteTags()
            Prcssd += 1
            HrtBet(cslCnt, slCnt)

         End If

      Next

      Epilog(lcPrlg, Prcssd, slCnt)

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkCpr() As Boolean

      Dim cslCnt As Long

      Dim lcPrlg As String = Prolog("CnsTrkCpr")

      Dim sdIndx As Integer()
      Dim cIndex As Integer = -1

      Dim csdTit As String = Nothing
      Dim psdTit As String = Nothing
      Dim csdCpr As String = Nothing
      Dim psdCpr As String = Nothing
      Dim dTrkCp As Boolean = False

      ReDim sdIndx(cIndex)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)
         csdTit = sd.Title
         csdCpr = sd.MusicComposer

         If csdTit = psdTit Then
            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt
            If psdCpr <> Nothing And psdCpr <> csdCpr Then dTrkCp = True
            psdCpr = csdCpr

         Else

            If cIndex > -1 And dTrkCp Then
               If Not CnsTrkCprExc(cslCnt, sdIndx) Then Return False
            End If

            psdTit = csdTit
            psdCpr = csdCpr

            cIndex = -1
            ReDim sdIndx(cIndex)
            dTrkCp = False

            cIndex += 1
            ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt

         End If

      Next

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsTrkCprExc(cslCnt, sdIndx) As Boolean

      CnsTrkCprExc = False

      Dim MaxcIndex As Integer = sdIndx.GetUpperBound(0)
      Dim wmxcIndex As Integer : If MaxcIndex < 99 Then wmxcIndex = 99 Else wmxcIndex = MaxcIndex

      Dim sd As New SDBSongData
      Dim cIndex As Integer

      Header("Komponist(en) für Track(s) richtig stellen")

      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing

      Dim cTrkNm As String = Nothing
      Dim pTrkNm As String = Nothing

      For cIndex = 0 To MaxcIndex

         sd = sl.Item(sdIndx(cIndex))

         cAlbNm = sd.AlbumName
         cTrkNm = sd.Title
         If cAlbNm <> pAlbNm And pAlbNm <> Nothing _
         Or cTrkNm <> pTrkNm And pTrkNm <> Nothing _
         Then
            cnscLn += 1 : Console.WriteLine()
         End If
         pAlbNm = cAlbNm
         pTrkNm = cTrkNm

         cnscLn += 1 : Console.WriteLine(AlbHdr(cIndex, wmxcIndex, sdIndx(cIndex), cslCnt) _
                                       & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                       & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                       & "|" & FixLngStr(sd.MusicComposer, MusCprLng) _
                                       & "|"
                                        )

      Next

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      REM  Datenlieferant&Empfänger
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim DtaSrcInt As Integer = EntDatSrc()
      If DtaSrcInt < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaSrcInt < -99 Then Return False Else Return True
      End If

      Dim DtaDrnInt() As Integer = EntDatDrn("A")
      If DtaDrnInt(0) < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaDrnInt(0) < -99 Then Return False Else Return True
      End If

      If DtaDrnInt(0) = -1 Then
         Dim z As Integer = -1
         For y = 0 To sdIndx.GetUpperBound(0)
            If y <> DtaSrcInt Then
               z += 1 : ReDim Preserve DtaDrnInt(z) : DtaDrnInt(z) = sdIndx(y)
            End If
         Next
      End If
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim SrcDta As New SDBSongData
      Dim DrnDta As New SDBSongData

      SrcDta = sl.Item(sdIndx(DtaSrcInt))

      For x = 0 To DtaDrnInt.GetUpperBound(0)

         DrnDta = sl.Item(sdIndx(DtaDrnInt(x)))

         Console.Write("|" & FixLngInt(EntDatSrc, 999) _
                     & "|" & FixLngInt(sdIndx(DtaSrcInt), cslCnt) _
                     & "|" & FixLngStr(SrcDta.MusicComposer, MusCprLng) _
                     & "|" & FixLngInt(DtaDrnInt(x), 999) _
                     & "|" & FixLngInt(sdIndx(DtaDrnInt(x)), cslCnt) _
                     & "|" & FixLngStr(DrnDta.MusicComposer, MusCprLng)
                      )

         DrnDta.MusicComposer = SrcDta.MusicComposer
         Console.WriteLine("|" & FixLngStr(DrnDta.MusicComposer, MusCprLng) & "|")

         DrnDta.UpdateDB()
         DrnDta.WriteTags()

      Next

      Console.ReadLine()

      SrcDta = Nothing
      DrnDta = Nothing

      Return True

   End Function
   REM *--------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   REM * Bilder konsolidieren
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CnsAlbTrkArt() As Boolean

      Dim cslCnt As Long

      Dim lcPrlg As String = Prolog("CnsAlbTrkArt")

      Dim sdIndx As Integer()
      Dim cIndex As Integer = -1
      ReDim sdIndx(cIndex)

      Dim csdAlb As String = Nothing
      Dim psdAlb As String = Nothing

      Dim AlbArtIdx As Long = -1

      Dim _1p As Boolean = False

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)

         csdAlb = sd.AlbumName
         If Not _1p Then
            psdAlb = csdAlb : _1p = True
            AlbArtIdx = -1 : ReDim AlbArtAry(AlbArtIdx)
            cIndex = -1 : ReDim Preserve sdIndx(cIndex)
            cnscLn = -1 : Console.Clear()
            cnscLn += 1 : Console.WriteLine(AlbHdr(cslCnt, slCnt, 0, 0) & "|")
         End If

         If csdAlb <> psdAlb Then

            If AlbArtIdx > -1 Then If Not EndOfAlbum(cslCnt, sdIndx) Then Return True
            psdAlb = csdAlb
            AlbArtIdx = -1 : ReDim AlbArtAry(AlbArtIdx)
            cIndex = -1 : ReDim Preserve sdIndx(cIndex)
            cnscLn = -1 : Console.Clear()
            cnscLn += 1 : Console.WriteLine(AlbHdr(cslCnt, slCnt, 0, 0) & "|")

         End If

         AlbArtSDB = sd.AlbumArt
         AlbArtCnt = AlbArtSDB.Count - 1

         FilDrvExt(sd.Path)
         cnscLn += 1 : Console.Write("|" & FixLngInt(cslCnt, slCnt) _
                                   & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                                   & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                   & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                   & "|" & FixLngStr(FilExt, FilExtLng)
                                        )

         For aIndx = 0 To AlbArtCnt

            AlbArtItm = AlbArtSDB.Item(aIndx)
            AlbArtNam = AlbArtItm.PicturePath

            If aIndx = 0 Then

               Console.WriteLine("|" & FixLngInt(aIndx, AlbArtLng) _
                               & "|" & trsltAlbArtItmTyp(AlbArtItm.ItemType) _
                               & "|" & FixLngInt(AlbArtItm.ItemStorage, 1) _
                               & "|" & FixLngStr(AlbArtItm.PicturePath, ArtNamLng) _
                               & "|"
                                )

            Else
               cnscLn += 1 : Console.WriteLine("|" & FixLngInt(cslCnt, slCnt) _
                                             & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                                             & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                             & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                             & "|" & FixLngStr(FilExt, FilExtLng) _
                                             & "|" & FixLngInt(aIndx, AlbArtLng) _
                                             & "|" & trsltAlbArtItmTyp(AlbArtItm.ItemType) _
                                             & "|" & FixLngInt(AlbArtItm.ItemStorage, 1) _
                                             & "|" & FixLngStr(AlbArtItm.PicturePath, ArtNamLng) _
                                             & "|"
                                              )

            End If

            AlbArtIdx += 1 : ReDim Preserve AlbArtAry(AlbArtIdx) : AlbArtAry(AlbArtIdx) = AlbArtItm.Image

         Next

         cIndex += 1 : ReDim Preserve sdIndx(cIndex) : sdIndx(cIndex) = cslCnt

         If AlbArtCnt < 0 Then Console.WriteLine("|")

      Next

      Return True

   End Function

   Function EndOfAlbum(cslCnt, sdIndx) As Boolean
      cnscLn += 1 : Console.WriteLine()
      Dim DtaSrc = "Album Fertig, weiter mit AnyKey (upd:Update all pictures del:Delete all pictures q:Quit) "
      cnscLn += 1 : Console.WriteLine(DtaSrc)
      Console.SetCursorPosition(Len(DtaSrc), cnscLn)
      DtaSrc = Console.ReadLine()
      cnscLn = -1 : Console.Clear()

      If DtaSrc = "q" Or DtaSrc = "q" Then Return False

      Select Case DtaSrc

         Case "upd" : If Not CnsAlbTrkArtUpd(cslCnt, sdIndx) Then DtaSrc = "q"

      End Select

      Return True

   End Function

   Function CnsAlbTrkArtUpd(cslCnt, sdIndx) As Boolean

      CnsAlbTrkArtUpd = False

      Dim MaxcIndex As Integer = sdIndx.GetUpperBound(0)
      Dim wmxcIndex As Integer : If MaxcIndex < 99 Then wmxcIndex = 99 Else wmxcIndex = MaxcIndex

      REM        Dim sd As New SDBSongData
      Dim cIndex As Integer

      Header("Bilder für Track(s) richtig stellen")

      Dim cAlbNm As String = Nothing
      Dim pAlbNm As String = Nothing

      Dim cTrkNm As String = Nothing
      Dim pTrkNm As String = Nothing

      Dim AlbArtIdx As Long = -1

      For cIndex = 0 To MaxcIndex

         sd = sl.Item(sdIndx(cIndex))
         FilDrvExt(sd.Path)

         cAlbNm = sd.AlbumName
         cTrkNm = sd.Title

         If cAlbNm <> pAlbNm And pAlbNm <> Nothing _
         Or cTrkNm <> pTrkNm And pTrkNm <> Nothing _
         Then
            cnscLn += 1 : Console.WriteLine()
         End If

         pAlbNm = cAlbNm
         pTrkNm = cTrkNm

         AlbArtSDB = sd.AlbumArt
         AlbArtCnt = AlbArtSDB.Count - 1

         For aIndx = 0 To AlbArtCnt

            AlbArtItm = AlbArtSDB.Item(aIndx)
            AlbArtNam = AlbArtItm.PicturePath

            cnscLn += 1 : Console.WriteLine("|" & FixLngInt(cslCnt, slCnt) _
                                          & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                                          & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                          & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                          & "|" & FixLngStr(FilExt, FilExtLng) _
                                          & "|" & FixLngInt(aIndx, AlbArtLng) _
                                          & "|" & trsltAlbArtItmTyp(AlbArtItm.ItemType) _
                                          & "|" & FixLngInt(AlbArtItm.ItemStorage, 1) _
                                          & "|" & FixLngStr(AlbArtItm.PicturePath, ArtNamLng) _
                                          & "|"
                                           )

         Next

      Next

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      REM  Datenlieferant&Empfänger
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim DtaSrcInt As Integer = EntDatSrc()
      If DtaSrcInt < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaSrcInt < -99 Then Return False Else Return True
      End If

      Dim DtaDrnInt() As Integer = EntDatDrn("A")
      If DtaDrnInt(0) < -98 Then
         ReDim sdIndx(-1)
         sd = Nothing
         Console.Clear()
         If DtaDrnInt(0) < -99 Then Return False Else Return True
      End If

      If DtaDrnInt(0) = -1 Then
         Dim z As Integer = -1
         For y = 0 To sdIndx.GetUpperBound(0)
            If y <> DtaSrcInt Then
               z += 1 : ReDim Preserve DtaDrnInt(z) : DtaDrnInt(z) = sdIndx(y)
            End If
         Next
      End If
      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------

      REM *-----------------------------------------------------------------------------------------------------------------------------------------------------------
      Dim SrcDta As New SDBSongData
      Dim DrnDta As New SDBSongData

      SrcDta = sl.Item(sdIndx(DtaSrcInt))

      For x = 0 To DtaDrnInt.GetUpperBound(0)

         DrnDta = sl.Item(sdIndx(DtaDrnInt(x)))

         For aIndx = 0 To AlbArtCnt

            AlbArtItm = AlbArtSDB.Item(aIndx)
            AlbArtNam = AlbArtItm.PicturePath

            If aIndx = 0 Then

               Console.WriteLine("|" & FixLngInt(aIndx, AlbArtLng) _
                               & "|" & trsltAlbArtItmTyp(AlbArtItm.ItemType) _
                               & "|" & FixLngInt(AlbArtItm.ItemStorage, 1) _
                               & "|" & FixLngStr(AlbArtItm.PicturePath, ArtNamLng) _
                               & "|"
                                )

            Else
               cnscLn += 1 : Console.WriteLine("|" & FixLngInt(cslCnt, slCnt) _
                                             & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng) _
                                             & "|" & FixLngStr(sd.TrackOrderStr, TrkNbrLng) _
                                             & "|" & FixLngStr(sd.Title, TrkNamLng) _
                                             & "|" & FixLngStr(FilExt, FilExtLng) _
                                             & "|" & FixLngInt(aIndx, AlbArtLng) _
                                             & "|" & trsltAlbArtItmTyp(AlbArtItm.ItemType) _
                                             & "|" & FixLngInt(AlbArtItm.ItemStorage, 1) _
                                             & "|" & FixLngStr(AlbArtItm.PicturePath, ArtNamLng) _
                                             & "|"
                                              )

            End If

         Next

         REM DrnDta.UpdateDB()
         REM DrnDta.WriteTags()

      Next

      Console.ReadLine()

      SrcDta = Nothing
      DrnDta = Nothing

      Return True

   End Function

   Function trsltAlbArtItmTyp(ItemType As Long) As String
      Dim wkStr As String

      Select Case ItemType

         Case 0 : wkStr = "n/a"
         Case 1 : wkStr = "32x32 file icon"
         Case 2 : wkStr = "other file icon"
         Case 3 : wkStr = "Front-Cover"
         Case 4 : wkStr = "Back-Cover"
         Case 5 : wkStr = "Leaflet page"
         Case 6 : wkStr = "Media"
         Case 7 : wkStr = "Lead artist"
         Case 8 : wkStr = "Performer"
         Case 9 : wkStr = "Conductor"
         Case 10 : wkStr = "Orchestra"
         Case 11 : wkStr = "Composer"
         Case 12 : wkStr = "n/a"
         Case 13 : wkStr = "n/a"
         Case 14 : wkStr = "n/a"
         Case 15 : wkStr = "n/a"
         Case 16 : wkStr = "n/a"
         Case 17 : wkStr = "n/a"
         Case 18 : wkStr = "n/a"
         Case 19 : wkStr = "n/a"
         Case 20 : wkStr = "n/a"
         Case Else : wkStr = "unknown"

      End Select

      Return FixLngInt(ItemType, 2) & " " & FixLngStr(wkStr, 15)

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function Exp2HndBas() As Boolean

      Dim cslCnt As Long

      Dim CvsFil = My.Computer.FileSystem.OpenTextFileWriter("mmTools.LstAllSng.4.HndBas.csv", False)
      Dim TxtStr As String
      Dim AlbMus As String
      Dim AlbNam As String
      Dim TrcTit As String
      Dim TrcMus As String
      Dim Compos As String

      TxtStr = "Albumartist" _
       & ";" & "Year" _
       & ";" & "Album" _
       & ";" & "Trackartist" _
       & ";" & "Disc" _
       & ";" & "Track" _
       & ";" & "Running Time" _
       & ";" & "Title" _
       & ";" & "|"
      Console.WriteLine(TxtStr)
      CvsFil.WriteLine(TxtStr)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)

         FilDrvExt(sd.Path)

         AlbMus = Trim(Replace(sd.AlbumArtistName, ";", ":"))
         AlbNam = Trim(Replace(sd.AlbumName, ";", ":"))
         TrcTit = Trim(Replace(sd.Title, ";", ":"))
         TrcMus = Trim(Replace(sd.ArtistName, ";", ":"))
         Compos = Trim(Replace(sd.MusicComposer, ";", ":"))

         TxtStr = AlbMus _
          & ";" & FixLngStr(sd.Year, 4, "R") _
          & ";" & AlbNam _
          & ";" & TrcMus _
          & ";" & FixLngStr(sd.DiscNumberStr, TrkNbrLng, "R") _
          & ";" & FixLngStr(sd.TrackOrderStr, TrkNbrLng, "R") _
          & ";" & FixLngStr(sd.SongLengthString, TrkTimLng, "R") _
          & ";" & TrcTit

         WrtCnsLin(TxtStr)
         CvsFil.WriteLine(TxtStr)

      Next

      CvsFil.Flush()
      CvsFil.Close()

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function LstAllSng() As Boolean

      Dim cslCnt As Long

      Dim CvsFil = My.Computer.FileSystem.OpenTextFileWriter("mmTools.LstAllSng.csv", False)
      Dim TxtStr As String
      Dim AlbMus As String
      Dim AlbNam As String
      Dim TrcTit As String
      Dim TrcMus As String
      Dim Compos As String

      TxtStr = "|" & "Lfnr" _
             & ";" & "Albumartist" _
             & ";" & "Year" _
             & ";" & "Album" _
             & ";" & "Trackartist" _
             & ";" & "Disc" _
             & ";" & "Track" _
             & ";" & "Running Time" _
             & ";" & "Title" _
             & ";" & "Komponist(en) & Texter" _
             & ";" & "Drive&Extension" _
             & ";" & "Vollständiger Pfad" _
             & "|" _
             & ";" & "BildLfnr" _
             & ";" & "AlbArtItm.ItemType" _
             & ";" & "AlbArtItm.Storage" _
             & "|"
      Console.WriteLine(TxtStr)
      CvsFil.WriteLine(TxtStr)

      For cslCnt = 0 To (slCnt)

         sd = sl.Item(cslCnt)

         FilDrvExt(sd.Path)

         AlbMus = Trim(Replace(sd.AlbumArtistName, ";", ":"))
         AlbNam = Trim(Replace(sd.AlbumName, ";", ":"))
         TrcTit = Trim(Replace(sd.Title, ";", ":"))
         TrcMus = Trim(Replace(sd.ArtistName, ";", ":"))
         Compos = Trim(Replace(sd.MusicComposer, ";", ":"))

         AlbArtSDB = sd.AlbumArt

         TxtStr = FixLngInt(cslCnt, slCnt, "Z") _
                & ";" & AlbMus _
                & ";" & FixLngStr(sd.Year, 4, "R") _
                & ";" & AlbNam _
                & ";" & TrcMus _
                & ";" & FixLngStr(sd.DiscNumberStr, TrkNbrLng, "R") _
                & ";" & FixLngStr(sd.TrackOrderStr, TrkNbrLng, "R") _
                & ";" & FixLngStr(sd.SongLengthString, TrkTimLng, "R") _
                & ";" & TrcTit _
                & ";" & Compos _
                & ";" & FilDrv & "." & FilExt _
                & ";" & sd.Path

         AlbArtCnt = AlbArtSDB.Count - 1
         For aIndx = 0 To AlbArtCnt
            AlbArtItm = AlbArtSDB.Item(aIndx)
            TxtStr = TxtStr _
                   & ";" & FixLngInt(aIndx, AlbArtLng) _
                   & ";" & Trim(trsltAlbArtItmTyp(AlbArtItm.ItemType)) _
                   & ";" & FixLngInt(AlbArtItm.ItemStorage, 1)

         Next

         WrtCnsLin(TxtStr)
         CvsFil.WriteLine(TxtStr)

      Next

      CvsFil.Flush()
      CvsFil.Close()

      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function WrtCnsLin(wrkInpStr As String) As Boolean

      Dim InpStr As New StringBuilder
      InpStr.Clear()
      InpStr.Append(wrkInpStr)

      If Console.BufferWidth <= InpStr.Length Then Console.BufferWidth = (InpStr.Length + 1)
      Console.WriteLine(InpStr.ToString)
      Return True

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function EntDatSrc() As Integer

      Dim MnuLin(0) As String
      Dim MnucIndex As Integer = -1
      Dim MnuSel As String = Nothing

      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = ""
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " Datenspender: "

      cnscLn += 1 : Console.WriteLine(MnuLin(0))
      cnscLn += 1 : Console.WriteLine(MnuLin(1))
      Console.SetCursorPosition(Len(MnuLin(1)), cnscLn)

      Dim DtaSrc = Console.ReadLine()
      If DtaSrc = "99" Or DtaSrc = "" Or DtaSrc = "q" Or DtaSrc = "Q" Then
         EntDatSrc = -99
         Console.Clear()
         If DtaSrc = "q" Then EntDatSrc -= 1
         Return EntDatSrc
      End If

      If Not IsNumeric(DtaSrc) Then DtaSrc = ""
      Return DtaSrc

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Sub FilDrvExt(sdPath As String)
      FilDrv = Left(sdPath, 1)
      FilExt = Right(sdPath, Len(sdPath) - (InStrRev(sdPath, ".", -1, CompareMethod.Text)))
   End Sub
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function EntDatDrn(Variante As String) As Integer()

      Dim MnuLin(0) As String
      Dim MnucIndex As Integer = -1
      Dim MnuSel As String = Nothing
      Dim DtaDrnDta As Integer()
      Dim DtaDrnIdx As Integer = -1

      ReDim DtaDrnDta(DtaDrnIdx)

      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = ""
      If Variante = "a" Or Variante = "A" Then
         MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " Datenempfänger (a/A : Alle ausser Datenspender) : "
      Else
         MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " Datenempfänger : "
      End If

      cnscLn += 1 : Console.WriteLine(MnuLin(0))

      Do While 1 = 1

         cnscLn += 1 : Console.WriteLine(MnuLin(1))
         Console.SetCursorPosition(Len(MnuLin(1)), cnscLn)
         Dim DtaDrn = Console.ReadLine

         If DtaDrn = "q" Or DtaDrn = "r" Then
            DtaDrnIdx = 0 : ReDim DtaDrnDta(DtaDrnIdx)
            DtaDrnDta(DtaDrnIdx) = -99
            Console.Clear()
            If DtaDrn = "q" Then DtaDrnDta(DtaDrnIdx) -= 1
            Return DtaDrnDta
         End If

         If DtaDrn = "" Then Return DtaDrnDta
         If DtaDrn = "a" Or DtaDrn = "A" Then
            DtaDrnIdx = 0 : ReDim DtaDrnDta(DtaDrnIdx)
            DtaDrnDta(DtaDrnIdx) = -1
            Return DtaDrnDta
         End If
         DtaDrnIdx += 1 : ReDim Preserve DtaDrnDta(DtaDrnIdx) : DtaDrnDta(DtaDrnIdx) = DtaDrn
      Loop

      Return DtaDrnDta

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function Menu() As String

      Dim MnuLin(0) As String
      Dim MnucIndex As Integer = -1
      Dim MnuSel As String = Nothing

      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = ""
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " ******************************************************************"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * (c)&(p) 2015ff ThomasW                                         *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 01 - Track- und Disc-Nummern 2-stellig machen                  *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 02 - Sonderzeichen richtig stellen                             *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 03 - on/the/is usw. richtig stellen                            *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 04 - Datei-Endungen korrigieren                                *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 10 - Album-Künstler korrigieren                                *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 11 - Track-Album-Künstler korrigieren                          *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 12 - Track-Künstler korrigieren                                *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 13 - dbPoweramp Komponisten richtig stellen                    *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 14 - Track-Komponist(en) korrigieren                           *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 15 - Album/Track-Bild korrigieren                              *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 20 - HandBase-Basis ausgeben                                   *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * 80 - Liste aktueller Titel ausgeben                            *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *                                             99 (04+: r) - Ende *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " *----------------------------------------------------------------*"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " * Generell: 99 oder Eingabe: keine Änderung          q : Abbruch *"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " ******************************************************************"
      MnucIndex += 1 : ReDim Preserve MnuLin(MnucIndex) : MnuLin(MnucIndex) = " Auswahl: "

      Dim cIndex As Integer
      Console.Clear()
      For cIndex = 0 To MnucIndex : Console.WriteLine(MnuLin(cIndex)) : Next

      Console.SetCursorPosition(Len(MnuLin(UBound(MnuLin))), MnucIndex)
      MnuSel = Console.ReadLine()
      Return MnuSel

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function Header(HeaderText As String) As Boolean

      Header = True

      cnscLn = -1 : Console.Clear()
      cnscLn += 1 : Console.WriteLine()
      cnscLn += 1 : Console.WriteLine(HeaderText)
      cnscLn += 1 : Console.WriteLine()

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function AlbHdr(Pos As Integer, mPos As Integer, sdidxcIdx As Integer, slLstCt As Integer) As String

      Return "|" & FixLngInt(Pos, mPos) _
           & "|" & FixLngInt(sdidxcIdx, slLstCt) _
           & "|" & FixLngStr(sd.AlbumName, AlbNamLng) _
           & "|" & FixLngStr(sd.AlbumArtistName, MusNamLng) _
           & "|" & FixLngStr(sd.DiscNumberStr, TrkNbrLng)

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function Prolog(Optional Caller As String = "") As String

      Dim wkstrstr As New StringBuilder
      wkstrstr.Append("***")
      If Caller <> "" Then
         wkstrstr.Append(" ")
         wkstrstr.Append(Caller)
         wkstrstr.Append(" ***")
      End If
      wkstrstr.Append(" Str: " & CurDatTim() & " to Process: " & FixLngInt(slCnt, slCnt))
      Dim strstr As String = wkstrstr.ToString
      wkstrstr = Nothing

      Console.Clear()
      Console.WriteLine(strstr)
      Return strstr

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Sub Epilog(Prolog As String, Prcssd As Integer, slCnt As Integer)

      slCnt = slCnt - 1
      Console.WriteLine(Prolog)
      Console.WriteLine("*** read: " & FixLngInt(slCnt, slCnt))
      Console.WriteLine("*** prc : " & FixLngInt(Prcssd, slCnt))
      Console.WriteLine("*** End : " & CurDatTim())
      Console.WriteLine("AnyKey")
      Console.ReadLine()

   End Sub
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------


   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function LdgZ(N As Integer) As String
      If (N >= 0) And (N < 10) Then LdgZ = "0" & N Else LdgZ = "" & N
   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Sub HrtBet(x As Integer, y As Integer)
      If (x Mod 1000) = 0 Then Console.WriteLine(pfx(x, y))
   End Sub
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function pfx(x As Integer, y As Integer) As String
      Return "*** " & CurDatTim() & " " & (FixLngInt(x, y))
   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function FixLngInt(n As Integer, maxn As Integer, Optional Edt As String = "K") As String

      Dim FmtStr As String
      Select Case Edt
         Case "Z"
            FmtStr = "0"
         Case Else
            FmtStr = "#,##0"
      End Select
      Dim StrFmt = Format(n, FmtStr)

      Dim lStr As Integer = Len(StrFmt)
      Dim lMax As Integer = Len(Format(maxn, FmtStr))
      Dim clen = lMax - lStr
      Dim wkFixLng As New StringBuilder
      Dim strFixLng As String

      If n = 0 And maxn = 0 Then Return ""
      If clen < 0 Then clen = 0

      wkFixLng.Insert(0, " ", (clen))
      wkFixLng.Append(StrFmt)
      strFixLng = wkFixLng.ToString
      wkFixLng = Nothing

      Return strFixLng

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function FixLngStr(s As String, maxn As Integer, Optional LoR As String = "L") As String

      Dim slen As Integer = Len(s)
      Dim clen As Integer = maxn - slen
      Dim strFixLng As String
      Dim wkFixLng As New StringBuilder

      If clen < 0 Then
         wkFixLng.Append(Left(s, (maxn - 3)))
         wkFixLng.Append("...")
      Else

         Select Case LoR
            Case "L"                                  'Left
               wkFixLng.Append(Left(s, slen))
               wkFixLng.Insert(slen, " ", (clen))
            Case "R"                                  'Right
               wkFixLng.Append(" ", clen)
               wkFixLng.Append(Left(s, slen))
            Case "C"                                  'Center
               clen = Math.Ceiling((maxn - slen) / 2)
               wkFixLng.Append(" ", 0, clen)
               wkFixLng.Append(Left(s, slen))
         End Select

      End If

      strFixLng = wkFixLng.ToString
      wkFixLng = Nothing

      Return strFixLng

   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------

   REM ****************************************************************************************************************************************************************
   Function CurDatTim() As String
      Return CurDat() & " " & CurTim()
   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CurDat() As String
      Return Format(Now, "yyyy-MM-dd")
   End Function
   REM *---------------------------------------------------------------------------------------------------------------------------------------------------------------
   Function CurTim() As String
      Return Format(Now, "HH:mm:ss")
   End Function
   REM ****************************************************************************************************************************************************************

   REM ****************************************************************************************************************************************************************
End Module
REM ********************************************************************************************************************************************************************