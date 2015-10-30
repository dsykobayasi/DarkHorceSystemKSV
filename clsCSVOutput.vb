Friend Class clsCSVOutput

    Public cSqlConnection As System.Data.SqlClient.SqlConnection
    Public tran As SqlClient.SqlTransaction
    Private strRAFields As String = "RA_KEY,YMD,KEIBAJYO,KAISAI_KAI,KAISAI_DAYS,RACE_NO,YOUBI,WEIGHT_SYUBETU,KYORI,TRACK_CODE,START_TIME,HORCE_COUNT,RACE_NAME_DETAIL1,RACE_NAME_DETAIL2,RACE_NAME_DETAIL3"
    Private strSEFields As String = "SE_KEY,HORCE_NO,HORCE_NAME,JOKEY_NAME"
    Private strO1Fields As String = "O1_KEY,GET_DATETIME,HORCE_NO1,ODDS1,LOW_ODDS1,HORCE_NO2,ODDS2,LOW_ODDS2,HORCE_NO3,ODDS3,LOW_ODDS3,HORCE_NO4,ODDS4,LOW_ODDS4,HORCE_NO5,ODDS5,LOW_ODDS5,HORCE_NO6,ODDS6,LOW_ODDS6,HORCE_NO7,ODDS7,LOW_ODDS7,HORCE_NO8,ODDS8,LOW_ODDS8,HORCE_NO9,ODDS9,LOW_ODDS9,HORCE_NO10,ODDS10,LOW_ODDS10,HORCE_NO11,ODDS11,LOW_ODDS11,HORCE_NO12,ODDS12,LOW_ODDS12,HORCE_NO13,ODDS13,LOW_ODDS13,HORCE_NO14,ODDS14,LOW_ODDS14,HORCE_NO15,ODDS15,LOW_ODDS15,HORCE_NO16,ODDS16,LOW_ODDS16,HORCE_NO17,ODDS17,LOW_ODDS17,HORCE_NO18,ODDS18,LOW_ODDS18,HORCE_NO19,ODDS19,LOW_ODDS19,HORCE_NO20,ODDS20,LOW_ODDS20,HORCE_NO21,ODDS21,LOW_ODDS21,HORCE_NO22,ODDS22,LOW_ODDS22,HORCE_NO23,ODDS23,LOW_ODDS23,HORCE_NO24,ODDS24,LOW_ODDS24,HORCE_NO25,ODDS25,LOW_ODDS25,HORCE_NO26,ODDS26,LOW_ODDS26,HORCE_NO27,ODDS27,LOW_ODDS27,HORCE_NO28,ODDS28,LOW_ODDS28,HORCE_COUNT,FILLER"
    Private strO2Fields As String = "O2_KEY,GET_DATETIME,PATTERN1,ODDS1,PATTERN2,ODDS2,PATTERN3,ODDS3,PATTERN4,ODDS4,PATTERN5,ODDS5,PATTERN6,ODDS6,PATTERN7,ODDS7,PATTERN8,ODDS8,PATTERN9,ODDS9,PATTERN10,ODDS10,PATTERN11,ODDS11,PATTERN12,ODDS12,PATTERN13,ODDS13,PATTERN14,ODDS14,PATTERN15,ODDS15,PATTERN16,ODDS16,PATTERN17,ODDS17,PATTERN18,ODDS18,PATTERN19,ODDS19,PATTERN20,ODDS20,PATTERN21,ODDS21,PATTERN22,ODDS22,PATTERN23,ODDS23,PATTERN24,ODDS24,PATTERN25,ODDS25,PATTERN26,ODDS26,PATTERN27,ODDS27,PATTERN28,ODDS28,PATTERN29,ODDS29,PATTERN30,ODDS30,PATTERN31,ODDS31,PATTERN32,ODDS32,PATTERN33,ODDS33,PATTERN34,ODDS34,PATTERN35,ODDS35,PATTERN36,ODDS36,PATTERN37,ODDS37,PATTERN38,ODDS38,PATTERN39,ODDS39,PATTERN40,ODDS40,PATTERN41,ODDS41,PATTERN42,ODDS42,PATTERN43,ODDS43,PATTERN44,ODDS44,PATTERN45,ODDS45,PATTERN46,ODDS46,PATTERN47,ODDS47,PATTERN48,ODDS48,PATTERN49,ODDS49,PATTERN50,ODDS50,PATTERN51,ODDS51,PATTERN52,ODDS52,PATTERN53,ODDS53,PATTERN54,ODDS54,PATTERN55,ODDS55,PATTERN56,ODDS56,PATTERN57,ODDS57,PATTERN58,ODDS58,PATTERN59,ODDS59,PATTERN60,ODDS60,PATTERN61,ODDS61,PATTERN62,ODDS62,PATTERN63,ODDS63,PATTERN64,ODDS64,PATTERN65,ODDS65,PATTERN66,ODDS66,PATTERN67,ODDS67,PATTERN68,ODDS68,PATTERN69,ODDS69,PATTERN70,ODDS70,PATTERN71,ODDS71,PATTERN72,ODDS72,PATTERN73,ODDS73,PATTERN74,ODDS74,PATTERN75,ODDS75,PATTERN76,ODDS76,PATTERN77,ODDS77,PATTERN78,ODDS78,PATTERN79,ODDS79,PATTERN80,ODDS80,PATTERN81,ODDS81,PATTERN82,ODDS82,PATTERN83,ODDS83,PATTERN84,ODDS84,PATTERN85,ODDS85,PATTERN86,ODDS86,PATTERN87,ODDS87,PATTERN88,ODDS88,PATTERN89,ODDS89,PATTERN90,ODDS90,PATTERN91,ODDS91,PATTERN92,ODDS92,PATTERN93,ODDS93,PATTERN94,ODDS94,PATTERN95,ODDS95,PATTERN96,ODDS96,PATTERN97,ODDS97,PATTERN98,ODDS98,PATTERN99,ODDS99,PATTERN100,ODDS100,PATTERN101,ODDS101,PATTERN102,ODDS102,PATTERN103,ODDS103,PATTERN104,ODDS104,PATTERN105,ODDS105,PATTERN106,ODDS106,PATTERN107,ODDS107,PATTERN108,ODDS108,PATTERN109,ODDS109,PATTERN110,ODDS110,PATTERN111,ODDS111,PATTERN112,ODDS112,PATTERN113,ODDS113,PATTERN114,ODDS114,PATTERN115,ODDS115,PATTERN116,ODDS116,PATTERN117,ODDS117,PATTERN118,ODDS118,PATTERN119,ODDS119,PATTERN120,ODDS120,PATTERN121,ODDS121,PATTERN122,ODDS122,PATTERN123,ODDS123,PATTERN124,ODDS124,PATTERN125,ODDS125,PATTERN126,ODDS126,PATTERN127,ODDS127,PATTERN128,ODDS128,PATTERN129,ODDS129,PATTERN130,ODDS130,PATTERN131,ODDS131,PATTERN132,ODDS132,PATTERN133,ODDS133,PATTERN134,ODDS134,PATTERN135,ODDS135,PATTERN136,ODDS136,PATTERN137,ODDS137,PATTERN138,ODDS138,PATTERN139,ODDS139,PATTERN140,ODDS140,PATTERN141,ODDS141,PATTERN142,ODDS142,PATTERN143,ODDS143,PATTERN144,ODDS144,PATTERN145,ODDS145,PATTERN146,ODDS146,PATTERN147,ODDS147,PATTERN148,ODDS148,PATTERN149,ODDS149,PATTERN150,ODDS150,PATTERN151,ODDS151,PATTERN152,ODDS152,PATTERN153,ODDS153,FILLER1,FILLER2"
    Private strO4Fields As String = "O4_KEY,GET_DATE_TIME,PATTERN1,ODDS1,PATTERN2,ODDS2,PATTERN3,ODDS3,PATTERN4,ODDS4,PATTERN5,ODDS5,PATTERN6,ODDS6,PATTERN7,ODDS7,PATTERN8,ODDS8,PATTERN9,ODDS9,PATTERN10,ODDS10,PATTERN11,ODDS11,PATTERN12,ODDS12,PATTERN13,ODDS13,PATTERN14,ODDS14,PATTERN15,ODDS15,PATTERN16,ODDS16,PATTERN17,ODDS17,PATTERN18,ODDS18,PATTERN19,ODDS19,PATTERN20,ODDS20,PATTERN21,ODDS21,PATTERN22,ODDS22,PATTERN23,ODDS23,PATTERN24,ODDS24,PATTERN25,ODDS25,PATTERN26,ODDS26,PATTERN27,ODDS27,PATTERN28,ODDS28,PATTERN29,ODDS29,PATTERN30,ODDS30,PATTERN31,ODDS31,PATTERN32,ODDS32,PATTERN33,ODDS33,PATTERN34,ODDS34,PATTERN35,ODDS35,PATTERN36,ODDS36,PATTERN37,ODDS37,PATTERN38,ODDS38,PATTERN39,ODDS39,PATTERN40,ODDS40,PATTERN41,ODDS41,PATTERN42,ODDS42,PATTERN43,ODDS43,PATTERN44,ODDS44,PATTERN45,ODDS45,PATTERN46,ODDS46,PATTERN47,ODDS47,PATTERN48,ODDS48,PATTERN49,ODDS49,PATTERN50,ODDS50,PATTERN51,ODDS51,PATTERN52,ODDS52,PATTERN53,ODDS53,PATTERN54,ODDS54,PATTERN55,ODDS55,PATTERN56,ODDS56,PATTERN57,ODDS57,PATTERN58,ODDS58,PATTERN59,ODDS59,PATTERN60,ODDS60,PATTERN61,ODDS61,PATTERN62,ODDS62,PATTERN63,ODDS63,PATTERN64,ODDS64,PATTERN65,ODDS65,PATTERN66,ODDS66,PATTERN67,ODDS67,PATTERN68,ODDS68,PATTERN69,ODDS69,PATTERN70,ODDS70,PATTERN71,ODDS71,PATTERN72,ODDS72,PATTERN73,ODDS73,PATTERN74,ODDS74,PATTERN75,ODDS75,PATTERN76,ODDS76,PATTERN77,ODDS77,PATTERN78,ODDS78,PATTERN79,ODDS79,PATTERN80,ODDS80,PATTERN81,ODDS81,PATTERN82,ODDS82,PATTERN83,ODDS83,PATTERN84,ODDS84,PATTERN85,ODDS85,PATTERN86,ODDS86,PATTERN87,ODDS87,PATTERN88,ODDS88,PATTERN89,ODDS89,PATTERN90,ODDS90,PATTERN91,ODDS91,PATTERN92,ODDS92,PATTERN93,ODDS93,PATTERN94,ODDS94,PATTERN95,ODDS95,PATTERN96,ODDS96,PATTERN97,ODDS97,PATTERN98,ODDS98,PATTERN99,ODDS99,PATTERN100,ODDS100,PATTERN101,ODDS101,PATTERN102,ODDS102,PATTERN103,ODDS103,PATTERN104,ODDS104,PATTERN105,ODDS105,PATTERN106,ODDS106,PATTERN107,ODDS107,PATTERN108,ODDS108,PATTERN109,ODDS109,PATTERN110,ODDS110,PATTERN111,ODDS111,PATTERN112,ODDS112,PATTERN113,ODDS113,PATTERN114,ODDS114,PATTERN115,ODDS115,PATTERN116,ODDS116,PATTERN117,ODDS117,PATTERN118,ODDS118,PATTERN119,ODDS119,PATTERN120,ODDS120,PATTERN121,ODDS121,PATTERN122,ODDS122,PATTERN123,ODDS123,PATTERN124,ODDS124,PATTERN125,ODDS125,PATTERN126,ODDS126,PATTERN127,ODDS127,PATTERN128,ODDS128,PATTERN129,ODDS129,PATTERN130,ODDS130,PATTERN131,ODDS131,PATTERN132,ODDS132,PATTERN133,ODDS133,PATTERN134,ODDS134,PATTERN135,ODDS135,PATTERN136,ODDS136,PATTERN137,ODDS137,PATTERN138,ODDS138,PATTERN139,ODDS139,PATTERN140,ODDS140,PATTERN141,ODDS141,PATTERN142,ODDS142,PATTERN143,ODDS143,PATTERN144,ODDS144,PATTERN145,ODDS145,PATTERN146,ODDS146,PATTERN147,ODDS147,PATTERN148,ODDS148,PATTERN149,ODDS149,PATTERN150,ODDS150,PATTERN151,ODDS151,PATTERN152,ODDS152,PATTERN153,ODDS153,PATTERN154,ODDS154,PATTERN155,ODDS155,PATTERN156,ODDS156,PATTERN157,ODDS157,PATTERN158,ODDS158,PATTERN159,ODDS159,PATTERN160,ODDS160,PATTERN161,ODDS161,PATTERN162,ODDS162,PATTERN163,ODDS163,PATTERN164,ODDS164,PATTERN165,ODDS165,PATTERN166,ODDS166,PATTERN167,ODDS167,PATTERN168,ODDS168,PATTERN169,ODDS169,PATTERN170,ODDS170,PATTERN171,ODDS171,PATTERN172,ODDS172,PATTERN173,ODDS173,PATTERN174,ODDS174,PATTERN175,ODDS175,PATTERN176,ODDS176,PATTERN177,ODDS177,PATTERN178,ODDS178,PATTERN179,ODDS179,PATTERN180,ODDS180,PATTERN181,ODDS181,PATTERN182,ODDS182,PATTERN183,ODDS183,PATTERN184,ODDS184,PATTERN185,ODDS185,PATTERN186,ODDS186,PATTERN187,ODDS187,PATTERN188,ODDS188,PATTERN189,ODDS189,PATTERN190,ODDS190,PATTERN191,ODDS191,PATTERN192,ODDS192,PATTERN193,ODDS193,PATTERN194,ODDS194,PATTERN195,ODDS195,PATTERN196,ODDS196,PATTERN197,ODDS197,PATTERN198,ODDS198,PATTERN199,ODDS199,PATTERN200,ODDS200,PATTERN201,ODDS201,PATTERN202,ODDS202,PATTERN203,ODDS203,PATTERN204,ODDS204,PATTERN205,ODDS205,PATTERN206,ODDS206,PATTERN207,ODDS207,PATTERN208,ODDS208,PATTERN209,ODDS209,PATTERN210,ODDS210,PATTERN211,ODDS211,PATTERN212,ODDS212,PATTERN213,ODDS213,PATTERN214,ODDS214,PATTERN215,ODDS215,PATTERN216,ODDS216,PATTERN217,ODDS217,PATTERN218,ODDS218,PATTERN219,ODDS219,PATTERN220,ODDS220,PATTERN221,ODDS221,PATTERN222,ODDS222,PATTERN223,ODDS223,PATTERN224,ODDS224,PATTERN225,ODDS225,PATTERN226,ODDS226,PATTERN227,ODDS227,PATTERN228,ODDS228,PATTERN229,ODDS229,PATTERN230,ODDS230,PATTERN231,ODDS231,PATTERN232,ODDS232,PATTERN233,ODDS233,PATTERN234,ODDS234,PATTERN235,ODDS235,PATTERN236,ODDS236,PATTERN237,ODDS237,PATTERN238,ODDS238,PATTERN239,ODDS239,PATTERN240,ODDS240,PATTERN241,ODDS241,PATTERN242,ODDS242,PATTERN243,ODDS243,PATTERN244,ODDS244,PATTERN245,ODDS245,PATTERN246,ODDS246,PATTERN247,ODDS247,PATTERN248,ODDS248,PATTERN249,ODDS249,PATTERN250,ODDS250,PATTERN251,ODDS251,PATTERN252,ODDS252,PATTERN253,ODDS253,PATTERN254,ODDS254,PATTERN255,ODDS255,PATTERN256,ODDS256,PATTERN257,ODDS257,PATTERN258,ODDS258,PATTERN259,ODDS259,PATTERN260,ODDS260,PATTERN261,ODDS261,PATTERN262,ODDS262,PATTERN263,ODDS263,PATTERN264,ODDS264,PATTERN265,ODDS265,PATTERN266,ODDS266,PATTERN267,ODDS267,PATTERN268,ODDS268,PATTERN269,ODDS269,PATTERN270,ODDS270,PATTERN271,ODDS271,PATTERN272,ODDS272,PATTERN273,ODDS273,PATTERN274,ODDS274,PATTERN275,ODDS275,PATTERN276,ODDS276,PATTERN277,ODDS277,PATTERN278,ODDS278,PATTERN279,ODDS279,PATTERN280,ODDS280,PATTERN281,ODDS281,PATTERN282,ODDS282,PATTERN283,ODDS283,PATTERN284,ODDS284,PATTERN285,ODDS285,PATTERN286,ODDS286,PATTERN287,ODDS287,PATTERN288,ODDS288,PATTERN289,ODDS289,PATTERN290,ODDS290,PATTERN291,ODDS291,PATTERN292,ODDS292,PATTERN293,ODDS293,PATTERN294,ODDS294,PATTERN295,ODDS295,PATTERN296,ODDS296,PATTERN297,ODDS297,PATTERN298,ODDS298,PATTERN299,ODDS299,PATTERN300,ODDS300,PATTERN301,ODDS301,PATTERN302,ODDS302,PATTERN303,ODDS303,PATTERN304,ODDS304,PATTERN305,ODDS305,PATTERN306,ODDS306,FILLER1,FILLER2"
    'DB接続
    Public Function connectDb() As Boolean
        Try
            ' 接続文字列を生成する
            Dim stConnectionString As String = String.Empty
            stConnectionString = "Data Source=218.251.113.60;Initial Catalog=ANASOFT;Persist Security Info=True;User ID=sa;Password=HBeiK3wR"

            ' SqlConnection の新しいインスタンスを生成する (接続文字列を指定)
            cSqlConnection = New System.Data.SqlClient.SqlConnection(stConnectionString)

            ' データベース接続を開く
            cSqlConnection.Open()

            Return True
        Catch ex As Exception
            MessageBox.Show("接続失敗")

            Return False
        End Try
    End Function

    'DB切断
    Public Function closeDB() As Boolean
        Try
            If Not cSqlConnection Is Nothing Then
                cSqlConnection.Close()
                cSqlConnection.Dispose()
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' 機能　　 : レース詳細のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したレース詳細レコード
    '            aRaceList - レース詳細リスト
    ' 返り値　 : なし
    ' 機能説明 : レース詳細レコードの情報を編集し、レース詳細リストに設定する
    '
    Public Sub RaceInfoMakeList(ByVal strBuff As String, ByRef aRaceList As ArrayList)
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim strRace As String
        Dim strRace2 As String
        Dim strRaceInfo() As String
        Dim strFromKey As String
        Dim strNewKey As String

        Try
            bSize = 1272
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strRace = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' SQLServer用
            strRace2 = "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & "',"

            ' <競走識別情報>
            strRace = strRace & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & "," & _
                    objCodeConv.GetCodeName("2001", MidB2S(bBuff, 20, 2), "3") & "," & _
                    MidB2S(bBuff, 22, 2) & "," & _
                    MidB2S(bBuff, 24, 2) & "," & _
                    MidB2S(bBuff, 26, 2) & "," & _
                    objCodeConv.GetCodeName("2002", MidB2S(bBuff, 28, 1), "2") & "," & _
                    objCodeConv.GetCodeName("2008", MidB2S(bBuff, 622, 1), "1") & "," & _
                    MidB2S(bBuff, 698, 4) & "," & _
                    objCodeConv.GetCodeName("2009", MidB2S(bBuff, 706, 2), "2") & "," & _
                    MidB2S(bBuff, 874, 4) & "," & _
                    MidB2S(bBuff, 882, 2) & "," & _
                    Trim(MidB2S(bBuff, 33, 60)) & "," & _
                    Trim(MidB2S(bBuff, 93, 60)) & "," & _
                    Trim(MidB2S(bBuff, 153, 60)) 

            ' SQLServer用
            strRace2 = strRace2 & "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & "','" & _
                    objCodeConv.GetCodeName("2001", MidB2S(bBuff, 20, 2), "3") & "','" & _
                    MidB2S(bBuff, 22, 2) & "','" & _
                    MidB2S(bBuff, 24, 2) & "','" & _
                    MidB2S(bBuff, 26, 2) & "','" & _
                    objCodeConv.GetCodeName("2002", MidB2S(bBuff, 28, 1), "2") & "','" & _
                    objCodeConv.GetCodeName("2008", MidB2S(bBuff, 622, 1), "1") & "'," & _
                    MidB2S(bBuff, 698, 4) & ",'" & _
                    objCodeConv.GetCodeName("2009", MidB2S(bBuff, 706, 2), "2") & "','" & _
                    MidB2S(bBuff, 874, 4) & "'," & _
                    MidB2S(bBuff, 882, 2) & ",'" & _
                    Trim(MidB2S(bBuff, 33, 60)) & "','" & _
                    Trim(MidB2S(bBuff, 93, 60)) & "','" & _
                    Trim(MidB2S(bBuff, 153, 60)) & "'"
            '出走頭数をやめて登録頭数を取得するよう修正
            'MidB2S(bBuff, 884, 2) & "," & _

            ' レース情報を判定し、更新があった場合はリストを更新する
            For i = 0 To RaceInfo.Count - 1
                ' 各レース情報から先頭25文字をキーとして取得
                strRaceInfo = RaceInfo(i)
                strFromKey = strRaceInfo(CommonConstant.Const0) & "," & _
                             strRaceInfo(CommonConstant.Const1)
                strNewKey = Strings.Left(strRace, CommonConstant.Const25)
                ' 現在レース情報と取得したレース情報のキーを比較する
                If strFromKey = strNewKey Then
                    ' レース情報を比較して異なる場合、現在レース情報を更新する
                    If Join(strRaceInfo, ",") <> strRace Then
                        RaceInfo.Item(i) = strRace.Split(","c)
                        Exit For
                    End If
                End If
            Next

            ' レース情報がすでに追加済みかチェック
            If Not aRaceList.Contains(strRace) Then
                aRaceList.Add(strRace)

                ExecSQL("RA_INFODATA_WORK", strRAFields, strRace2)
            End If
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : レース詳細のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aRaceList - レース詳細リスト
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにレース詳細のCSVファイルを出力する
    '
    Public Sub RaceInfoOutput(ByVal strFileName As String, ByVal strFilePath As String, ByVal aRaceList As ArrayList)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名
            Dim workRaceList As New ArrayList

            ' 引数fileNameからCSVファイル名を作成
            'strCsvName = strFileName
            strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            If System.IO.File.Exists(strCsvPath) Then
                ' ファイルを削除
                Kill(strCsvPath)
            End If

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            ' 取得した全レース詳細をCSVに出力する（20100912修正）
            For i = 0 To aRaceList.Count - 1
                ' <競走識別情報>
                sr.Write(aRaceList.Item(i).ToString)
                sr.Write(vbCrLf)
            Next

            '' レース詳細をArrayList<String>に変換する
            'For i = 0 To RaceInfo.Count - 1
            '    strWorkRace = Join(RaceInfo(i), ",")
            '    workRaceList.Add(strWorkRace)
            'Next

            'For i = 0 To workRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' 取得したレース詳細リストにデータが存在する場合はリストに追加しない
            '    If Not aRaceList.Contains(workRaceList.Item(i)) Then
            '        ' <競走識別情報>
            '        sr.Write(workRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            'For i = 0 To aRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            '    If Not workRaceList.Contains(aRaceList.Item(i)) Then
            '        ' <競走識別情報>
            '        sr.Write(aRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            ''For i = 0 To aRaceList.Count - 1
            ''    ' CSVファイルにレコードが存在するかチェックする
            ''    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            ''    If ((RaceInfo.Count = 0) Or _
            ''         (Not workRaceList.Contains(aRaceList.Item(i)))) Then
            ''        '' <競走識別情報>
            ''        sr.Write(aRaceList.Item(i).ToString)
            ''        sr.Write(vbCrLf)

            ''        ' ファイルオープンフラグをFalseにする
            ''        gFileOpenFlg = False
            ''    End If
            ''    '' <競走識別情報>
            ''    'sr.Write(aRaceList.Item(i).ToString)
            ''    'sr.Write(vbCrLf)
            ''Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : 馬毎レース情報のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得した馬毎レース情報レコード
    '            aHorseList - 馬毎レース情報リスト
    ' 返り値　 : なし
    ' 機能説明 : 馬毎レース情報レコードの情報を編集し、馬毎レース情報リストに設定する
    '
    Public Sub HorseMakeList(ByVal strBuff As String, ByRef aHorseList As ArrayList)
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim strHorse As String          ' レース情報
        Dim strHorse2 As String
        Dim strHorseInfo() As String
        Dim strFromKey As String
        Dim strNewKey As String

        Try
            bSize = 2042
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strHorse = ""

            ' 馬番が未決定の場合は処理終了
            If MidB2S(bBuff, 29, 2) = CommonConstant.No_HorseNo Then
                Exit Sub
            End If

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strHorse = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            'SQLServer用
            strHorse2 = "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & "',"
            ' <馬毎レース情報>
            strHorse = strHorse & _
                    MidB2S(bBuff, 29, 2) & "," & _
                    Trim(MidB2S(bBuff, 41, 36)) & "," & _
                    Trim(MidB2S(bBuff, 307, 8))

            'SQLServer用
            strHorse2 = strHorse2 & _
                    "'" & MidB2S(bBuff, 29, 2) & "','" & _
                    Trim(MidB2S(bBuff, 41, 36)) & "','" & _
                    Trim(MidB2S(bBuff, 307, 8)) & "'"
            ' 馬毎レース情報を判定し、更新があった場合はリストを更新する
            For i = 0 To AllHorseInfo.Count - 1
                ' 各レース情報から先頭19文字をキーとして取得
                strHorseInfo = AllHorseInfo(i)
                strFromKey = strHorseInfo(CommonConstant.Const0) & "," & _
                             strHorseInfo(CommonConstant.Const1)
                strNewKey = Strings.Left(strHorse, CommonConstant.Const19)
                ' 現在レース情報と取得したレース情報のキーを比較する
                If strFromKey = strNewKey Then
                    ' レース情報を比較して異なる場合、現在レース情報を更新する
                    If Join(strHorseInfo, ",") <> strHorse Then
                        AllHorseInfo.Item(i) = strHorse.Split(","c)
                        Exit For
                    End If
                End If
            Next

            ' レース情報がすでに追加済みかチェック
            If Not aHorseList.Contains(strHorse) Then
                aHorseList.Add(strHorse)

                ExecSQL("SE_INFODATA_WORK", strSEFields, strHorse2)
            End If
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ExecSQL(ByVal strTableName As String, ByVal strFields As String, ByVal strValues As String)
        Dim strSQL As String = ""

        strSQL = "INSERT INTO " & strTableName & "(" & strFields & ") VALUES (" & strValues & ")"

        Dim cmd As New System.Data.SqlClient.SqlCommand
        cmd.Connection = cSqlConnection
        cmd.CommandText = strSQL
        cmd.Transaction = tran
        cmd.ExecuteNonQuery()
        cmd.Dispose()
    End Sub

    ' 機能　　 : オッズ（単複枠）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ（単複枠）レコード
    '            aOdds1List - オッズ（単複枠）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ（単複枠）レコードの情報を編集し、オッズ（単複枠）リストに設定する
    '
    Public Sub Odds1MakeList(ByVal strBuff As String, ByRef aOdds1List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報
        Dim strOdds2 As String

        Try
            bSize = 962
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' SQLServer用
            strOdds2 = "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & "',"
            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            ' SQLServer用
            strOdds2 = strOdds2 & "'" & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & "',"
            ' 馬番（28頭立て）分繰り返す
            For i = 0 To 27
                ' <単勝オッズ>
                ' 馬番、オッズ
                bOddsInfo = MidB2B(bBuff, 44 + (8 * i), 8)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 2) & "," & _
                            MidB2S(bOddsInfo, 3, 4) & ","

                'SQLServer用
                strOdds2 = strOdds2 & "'" & _
                            MidB2S(bOddsInfo, 1, 2) & "','" & _
                            MidB2S(bOddsInfo, 3, 4) & "',"
                ' <複勝オッズ>
                ' 最低オッズ
                bOddsInfo = MidB2B(bBuff, 268 + (12 * i), 12)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 3, 4) & ","
                'SQLServer用
                strOdds2 = strOdds2 & "'" & _
                            MidB2S(bOddsInfo, 3, 4) & "',"
            Next

            ' <登録頭数>
            strOdds = strOdds & MidB2S(bBuff, 36, 2)

            'SQLServer用
            strOdds2 = strOdds2 & MidB2S(bBuff, 36, 2) & ",''"
            '出走頭数をやめて登録頭数を取得するよう修正
            '' <出走頭数>
            'strOdds = strOdds & MidB2S(bBuff, 38, 2)

            aOdds1List.Add(strOdds)

            ExecSQL("O1_INFODATA_WORK", strO1Fields, strOdds2)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ2（馬連）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ2（馬連）レコード
    '            aOdds2List - オッズ2（馬連）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ2（馬連）レコードの情報を編集し、オッズ2（馬連）リストに設定する
    '
    Public Sub Odds2MakeList(ByVal strBuff As String, ByRef aOdds2List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報
        Dim strOdds2 As String

        Try
            bSize = 2042
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""
            strOdds2 = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            'SQLSerer
            strOdds2 = "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & "',"
            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            'SQLServer
            strOdds2 = strOdds2 & "'" & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & "',"
            For i = 0 To 152
                ' <馬連オッズ>
                ' 組番、オッズ
                bOddsInfo = MidB2B(bBuff, 41 + (13 * i), 13)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 4) & "," & _
                            MidB2S(bOddsInfo, 5, 6) & ","
                'SQLServer
                strOdds2 = strOdds2 & "'" & _
                            MidB2S(bOddsInfo, 1, 4) & "','" & _
                            MidB2S(bOddsInfo, 5, 6) & "',"
            Next
            aOdds2List.Add(strOdds)

            strOdds2 = strOdds2 & "'',''"
            ExecSQL("O2_INFODATA_WORK", strO2Fields, strOdds2)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ4（馬単）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ4（馬単）レコード
    '            aOdds4List - オッズ4（馬単）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ4（馬単）レコードの情報を編集し、オッズ4（馬単）リストに設定する
    '
    Public Sub Odds4MakeList(ByVal strBuff As String, ByRef aOdds4List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報
        Dim strOdds2 As String

        Try
            bSize = 4031
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""
            strOdds2 = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            'SQLServer
            strOdds2 = "'" & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & "',"
            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            'SQLServer
            strOdds2 = strOdds2 & "'" & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & "',"
            For i = 0 To 305
                ' <馬連オッズ>
                bOddsInfo = MidB2B(bBuff, 41 + (13 * i), 13)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 4) & "," & _
                            MidB2S(bOddsInfo, 5, 6) & ","

                'SQLServer
                strOdds2 = strOdds2 & "'" & _
                            MidB2S(bOddsInfo, 1, 4) & "','" & _
                            MidB2S(bOddsInfo, 5, 6) & "',"
            Next
            aOdds4List.Add(strOdds)
            strOdds2 = strOdds2 & "'',''"
            ExecSQL("O4_INFODATA_WORK", strO4Fields, strOdds2)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : 馬毎レース情報のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aHorseList - 馬毎レース情報のリスト
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにCSVファイルを出力する
    '
    Public Sub SeInfoOutput(ByVal strFileName As String, ByVal strFilePath As String, ByVal aHorseList As ArrayList)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名
            Dim workRaceList As New ArrayList

            ' 引数fileNameからCSVファイル名を作成
            'strCsvName = strFileName
            strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            If System.IO.File.Exists(strCsvPath) Then
                ' ファイルを削除
                Kill(strCsvPath)
            End If

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            ' 取得した全馬毎レース情報をCSVに出力する（20100912修正）
            For i = 0 To aHorseList.Count - 1
                ' <競走識別情報>
                sr.Write(aHorseList.Item(i).ToString)
                sr.Write(vbCrLf)
            Next

            '' レース詳細をArrayList<String>に変換する
            'For i = 0 To AllHorseInfo.Count - 1
            '    strWorkRace = Join(AllHorseInfo(i), ",")
            '    workRaceList.Add(strWorkRace)
            'Next

            'For i = 0 To workRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' 取得したレース詳細リストにデータが存在する場合はリストに追加しない
            '    If Not aHorseList.Contains(workRaceList.Item(i)) Then
            '        ' <馬毎レース情報>
            '        sr.Write(workRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            'For i = 0 To aHorseList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            '    If Not workRaceList.Contains(aHorseList.Item(i)) Then
            '        ' <馬毎レース情報>
            '        sr.Write(aHorseList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            ''For i = 0 To aHorseList.Count - 1
            ''    ' CSVファイルにレコードが存在するかチェックする
            ''    ' 馬毎レース情報リストに取得データが存在する場合はリストに追加しない
            ''    If ((AllHorseInfo.Count = 0) Or _
            ''         (Not workRaceList.Contains(aHorseList.Item(i)))) Then
            ''        ' <馬毎レース情報>
            ''        sr.Write(aHorseList.Item(i).ToString)
            ''        sr.Write(","c)
            ''        sr.Write(vbCrLf)
            ''    End If
            ''    '' <馬毎レース情報>
            ''    'sr.Write(aHorseList.Item(i).ToString)
            ''    'sr.Write(","c)
            ''    'sr.Write(vbCrLf)
            ''Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ情報のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aOddsList  - オッズ情報のリスト
    '            strDataKbn - データ区分（1：単複枠、2：馬連、3：馬単）
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにCSVファイルを出力する
    '
    Public Sub OddsOutput(ByVal strFileName As String, ByVal strFilePath As String, _
                          ByVal aOddsList As ArrayList, ByVal strDataKbn As String)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名

            ' 引数fileNameからCSVファイル名を作成
            strCsvName = strFileName
            'strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            For i = 0 To aOddsList.Count - 1
                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報1リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.TanshouOddsKbn) And _
                    (TanFukuAllOddsInfo.Count = 0 Or _
                     Not TanFukuAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If

                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報2リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.UmarenOddsKbn) And _
                    (UmarenAllOddsInfo.Count = 0 Or _
                     Not UmarenAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If

                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報4リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.UmatanOddsKbn) And _
                    (UmatanAllOddsInfo.Count = 0 Or _
                     Not UmatanAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If
            Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
