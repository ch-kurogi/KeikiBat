' **************************************************************************
' 機能名称：MJ4103
' 機能概要：ダウンロードファイル作成
' 使用方法：
' 前提条件：
' 更新履歴：2022.09.26 - EC Y.Ohnishi 新規作成
' 更新履歴：Rev002 - 2023.07.18 - EC H.Morii 高圧他テーブル対応
' Rev038  ：2024.07.12 - EC Y.Ohnishi 懸127_環境文字不正混入対応
' Rev040  ：2024.12.03 - EC Y.Ohnishi 託送情報切り替え対応
' Rev056  ：2025.07.01 - EC Y.Hiragi  託送情報切り替え対応（高圧他）
' **************************************************************************
Friend Class MJ4103

#Region "（変数）グローバル"

    '-------------------------------------------------------------------------
    ' （変数）グローバル
    '-------------------------------------------------------------------------

    'グローバルクラス変数
    Public myUser As New HdSysBasePg.Cmn.UserInfo   'ユーザ情報
    Public myPath As New HdSysBasePg.Cmn.PathInfo   'パス情報
    Public myDB As New HdPostgre.HdPostgreDb        'ＤＢ接続
    Public KinoCD As String = ""                    '機能コード
    Public SyoriYMD As String = ""                  '処理年月日
    Public ZGSYO_CD_List As New List(Of String)     '対象の事業所(先頭文字・指定なしは全て対象)
    Public myPID As String = ""                     'プロセスID

#End Region


#Region "（変数）クラス内"

    '-------------------------------------------------------------------------
    ' （変数）クラス内
    '-------------------------------------------------------------------------
    ''システム日時
    'Private _sysDateYmdh As String = ""
    '
    ''業務処理日時
    'Private _syoriYmd As String = ""

    '持出用テンポラリフォルダ
    Private _mtdsTemp As String = ""
    Const C_MTDS_dir = "MTDS"
    '添付書類フォルダ名
    Const C_TNPDOC_dir As String = "TNPDOC"

    'ファイル名定義
    '施工
    Const C_FS_TRKE_TRKEHY_TIAT As String = "TRKE_TRKEHY_TIAT.csv"        '取替_取替票（低圧）
    Const C_FS_TRKE_TRKEHY_PRC As String = "TRKE_TRKEHY_PRC.csv"          '取替_取替票行程
    Const C_FS_KEIKI_SPOT_TKZK As String = "KEIKI_SPOT_TKZK.csv"          '計器_地点_特記事項
    Const C_FS_TRKE_TRKEHY_TUIKAKH As String = "TRKE_TRKEHY_TUIKAKH.csv"  '取替_取替票追加工費
    Const C_FS_TRKE_CHECK_SHEET As String = "TRKE_CHECK_SHEET.csv"        '取替_自主点検チェックシート
    Const C_FS_COM_TNPDOC_DTL As String = "COM_TNPDOC_DTL.csv"            '共通_添付書類詳細

    'Rev002-Start
    Const C_FS_TRKE_TRKEHY_KAHK As String = "TRKE_TRKEHY_KAHK.csv"        '取替_取替票（高圧他）
    Const C_FS_TRKE_TRKEHY_PRC_KAHK As String = "TRKE_TRKEHY_PRC_KAHK.csv"   '取替_取替票行程（高圧他）
    Const C_FS_KEIKI_SPOT_TKZK_KAHK As String = "KEIKI_SPOT_TKZK_KAHK.csv"   '計器_地点_特記事項（高圧他）
    Const C_FS_TRKE_TRKEHY_TUIKAKH_KAHK As String = "TRKE_TRKEHY_TUIKAKH_KAHK.csv"  '取替_取替票追加工費（高圧他）
    Const C_FS_TRKE_CHECK_SHEET_KAHK As String = "TRKE_CHECK_SHEET_KAHK.csv"        '取替_自主点検チェックシート（高圧他）
    'Rev002-End


    '抜取検査
    Const C_FN_TRKE_TRKEHY_TIAT As String = "TRKE_TRKEHY_TIAT.csv"        '取替_取替票（低圧）
    Const C_FN_TRKE_TRKEHY_PRC As String = "TRKE_TRKEHY_PRC.csv"          '取替_取替票行程
    Const C_FN_TRKE_CHECK_SHEET As String = "TRKE_CHECK_SHEET.csv"        '取替_自主点検チェックシート

    'Rev002-Start
    Const C_FN_TRKE_TRKEHY_KAHK As String = "TRKE_TRKEHY_KAHK.csv"        '取替_取替票（高圧他）
    Const C_FN_TRKE_TRKEHY_PRC_KAHK As String = "TRKE_TRKEHY_PRC_KAHK.csv"          '取替_取替票行程（高圧他）
    Const C_FN_TRKE_CHECK_SHEET_KAHK As String = "TRKE_CHECK_SHEET_KAHK.csv"        '取替_自主点検チェックシート（高圧他）
    'Rev002-End

    '共通
    Const C_FC_MST_MSG As String = "MST_MSG.csv"                            '共通_メッセージマスタ
    Const C_FC_MST_KORYO As String = "MST_KORYO.csv"                        'マスタ_工量
    Const C_FC_COM_PARAMATER_CD_MST As String = "COM_PARAMATER_CD_MST.csv"  '共通_パラメータコードマスタ
    Const C_FC_MST_HNMK As String = "MST_HNMK.csv"                          'マスタ_品目
    Const C_FC_MST_KEIKI_KTSK As String = "MST_KEIKI_KTSK.csv"              'マスタ_計器型式
    Const C_FC_MST_TYAZA As String = "MST_TYAZA.csv"                        'マスタ_町字
    Const C_FC_COM_ZYOUSU_MST As String = "COM_ZYOUSU_MST.csv"              '定数マスタ

    'Rev002-Start
    Const C_PHASE_KBN_NON As String = " "
    Const C_PHASE_KBN_1 As String = "1"
    Const C_PHASE_KBN_2 As String = "2"

    Const C_KIYK_KBN_CD_1 As String = "1"   '低圧契約
    Const C_KEIKI_KBN_CD_1 As String = "1"  '低圧計器
    Const C_SM_ZYRI_KBN_CD_1 As String = "1"    'スマートメータ
    'Rev002-End
#End Region

#Region "コンストラクタ"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

#End Region



#Region "メイン処理"

    ''' <summary>
    ''' メイン処理
    ''' </summary>
    ''' <remarks>
    ''' メイン処理を実行する。
    ''' </remarks>
    Public Function MainLogic() As HdSysBase.Cmn.ExitCode

        Dim wHaitaSigyoNo As Decimal = 0                                    ' 排他制御番号
        Dim wWarningFlg As Boolean = False                  '警告有無フラグ
        Dim wRet As Boolean = False                         'Private メソッド判定用
        Dim wMtdsMngNoPadZero As String = ""                '持出管理番号(年度4桁＋持出単位番号6桁を10桁前0埋め)
        Dim dbr As New HdPostgre.HdPostgreReader            'DB Reader
        Dim dao As New MJ4103_DAO(Me)           'DAO
        Dim wHS_NENDO As String = ""
        Dim wKOZI_CORP_CD As String = ""
        Dim wKZKIS_ZGSYO_CD As String = ""
        Dim wZGSYO_CD As String = ""
        Dim wMDSYA_USER_ID As String = ""
        Dim wMTDS_SIZI_MNG_ID As String = ""
        Dim wMTDS_TANI_NO As String = ""
        Dim wRow As System.Data.DataRow = Nothing
        Dim wPidStr As String = CLng(myPID).ToString
        Dim recCount As Double = 0
        Dim msg As String = ""
        Dim exitCode As String = ""

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'プロセスIDをログ出力
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}を開始しました。PID={1}", KinoCD, myPID))

            ' 持出用テンポラリフォルダ(myPath.BAT.K1TEMP & "MTDS")の設定
            Me._mtdsTemp = System.IO.Path.Combine(myPath.BAT.K1TEMP, "MTDS")
            If Not System.IO.Directory.Exists(Me._mtdsTemp) Then
                System.IO.Directory.CreateDirectory(Me._mtdsTemp)
            End If

            ' バッチ終了ファイルのチェックフォルダ
            Dim wEndCheckFolder As String = System.IO.Path.Combine(Me.myPath.BAT.K1TEMP, "BAT")


            '不要な日付は取得しない。mdlStart.vb内で取得した日付(SyoriYmd)を使用する。
            ''HdSysbase.Cmn.Function.GetSystemYmdh を呼出し，取得した日時を，内部変数.システム日時に設定する。
            'Me._sysDateYmdh = HdSysBasePg.Cmn.GetSystemYmdh(Me.myDB)
            '
            ''業務処理日を取得する。
            'Me._syoriYmd = HdSysBasePg.Cmn.GetSyoriYmd(myDB)

            wWarningFlg = False
            While HdSysBasePg.Cmn.CanUseKino(myDB, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD)
                ' 機能が使用可能な間（使用可能な時間帯）処理を繰り返す
                ' バッチ終了ファイルをチェックする
                If System.IO.File.Exists(System.IO.Path.Combine(wEndCheckFolder, Me.KinoCD)) OrElse
                    System.IO.File.Exists(System.IO.Path.Combine(wEndCheckFolder, Me.KinoCD + "_" + Me.myUser.USER_ID)) Then
                    ' 終了結果ファイル（空ファイル）を作成して，繰り返しを抜ける
                    System.IO.File.Create(System.IO.Path.Combine(wEndCheckFolder, "END_" + Me.KinoCD + "_" & Me.myPID)).Close()
                    Exit While
                End If

                '施工の処理を行う。
                '「持出ファイル作成処理」を行う。
                If Not Me.CreateMtdsFileSEKO(ZGSYO_CD_List, dao) Then
                    msg = String.Format("持出ファイル作成処理エラー(施工)")
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, msg)
                    wWarningFlg = True

                    'Zabbixにメッセージ送信する
                    exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4103:施工)=" & exitCode)

                    '20240712 Rev038 DB再接続追加 ADD Start
                    mdlStart.ReConnectDB()
                    '20240712 Rev038 DB再接続追加 ADD End
                End If

                '抜取検査の処理を行う。
                If Not Me.CreateMtdsFileNKTRKNS(ZGSYO_CD_List, dao) Then
                    msg = String.Format("持出ファイル作成処理エラー(抜取検査)")
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, msg)
                    wWarningFlg = True

                    'Zabbixにメッセージ送信する
                    exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4103:抜取検査)=" & exitCode)

                    '20240712 Rev038 DB再接続追加 ADD Start
                    mdlStart.ReConnectDB()
                    '20240712 Rev038 DB再接続追加 ADD End
                End If

                '処理間隔のためwait処理を行う(ミリ秒指定)。
                System.Threading.Thread.Sleep([Const].C_SLEEP_MTDS_INTERVAL)
                System.Windows.Forms.Application.DoEvents()
            End While

            'zabbixメッセージ送信で監視に対応するため警告終了でのJP1対応は不要 2023/6/16
            ''警告終了をJP1に返却する
            'If wWarningFlg = True Then
            '    Return HdSysBase.Cmn.ExitCode.Warning
            'End If

            '処理を終了する。
            Return HdSysBase.Cmn.ExitCode.Normal

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))

            'ロールバックする。
            myDB.Rollback()
            Return HdSysBase.Cmn.ExitCode.Abnormal
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If

        End Try

    End Function
#End Region


#Region "ダウンロードファイル作成（施工）"

    ''' <summary>
    ''' ダウンロードファイル作成（施工）
    ''' </summary>
    ''' <param name="pZGSYO_CD_List">対象事業所CD</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMtdsFileSEKO(ByVal pZGSYO_CD_List As List(Of String), ByVal dao As MJ4103_DAO) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader            'DB Reader
        Dim recCount As Double = 0

        Dim wDownldRenno As String = ""
        Dim wKTTNMT_NO As String = ""
        Dim wKeikiUserId As String = ""
        Dim wKeikiDwldSbtCd As String = ""
        Dim wMtdsFileZtiCd As String = ""

        Dim wDig4ZgsyoCd As String = ""
        Dim wNendo As String = ""
        Dim wTrkhyKbn As String = ""
        Dim wTrkhyNo As String = ""


        Dim wTempMtdsFolder As String = ""          '持出一時
        Dim wTempMtdsRennoFolder As String = ""     '持出一時ダウンロード番号フォルダ
        Dim wMtdsFolder As String = ""              'NAS側持出フォルダ

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名
        Dim wZipFileName As String = ""

        Dim firstFlag As Boolean = True

        Dim wTempTopDir As String = ""

        'Rev002-Start
        Dim wPhaseKbn = C_PHASE_KBN_NON     '取替票を検出し存在しなければ" ",低圧は"TIAT",高圧他は"KAHK"
        Dim wTrkhyExistFlag As Boolean      '取替票（高圧他）の有無
        'Rev002-End

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            ' ダウンロード管理データ一覧を取得する。
            If Not dao.S001(dbr, MjK1.Cmn.Const.KeikiDwldSbtCd.Seko, pZGSYO_CD_List) Then
                Throw New Exception(String.Format("施工：ダウンロード管理データ一覧取得処理でエラーが発生しました。"))
            End If
            '検索結果０件の場合、処理を終了する。
            If Not dbr.DataStruct.HasRows Then
                '対象レコードがない場合、正常終了する
                Return True
            End If

            'PostgreSQLは、selectの多重ができないため、結果をDataTableに設定し、dbrを開放する
            wDataTable = New DataTable
            'カラム名設定
            Dim i As Integer
            For i = 0 To dbr.DataStruct.FieldCount - 1
                wDataTable.Columns.Add(dbr.DataStruct.GetName(i))
            Next
            'DataTableにデータ保存
            While (dbr.DataStruct.Read)
                wRow = wDataTable.NewRow
                For i = 0 To dbr.DataStruct.FieldCount - 1
                    wRow(dbr.DataStruct.GetName(i)) = dbr.DataStruct.Item(i).ToString
                Next
                wDataTable.Rows.Add(wRow)
            End While
            dbr.Close()

            '20240717 Rev038 DB再接続追加 DEL Start  場所移動
            ''検索結果のキー情報を構築する。
            'Dim wKey As New List(Of String)
            'Dim wKeyIN As String = ""

            'firstFlag = True
            'For Each d As DataRow In wDataTable.Rows
            '    If firstFlag Then
            '        firstFlag = False
            '    Else
            '        wKeyIN = wKeyIN & ","
            '    End If

            '    wKeyIN = wKeyIN & CStr(Me.myDB.ConvDbNull(d("downld_renno"), ""))

            '    'リストにも保存
            '    wKey.Add(CStr(Me.myDB.ConvDbNull(d("downld_renno"), "")))
            'Next
            'Dim wWhere As String = "downld_renno IN (" & wKeyIN & ")"


            ''ダウンロード管理の検索したレコードの持出ファイル状態コードを"持出データ作成中"に更新する
            ''トランザクションを開始する。
            'Me.myDB.BeginTransaction()
            'If Not dao.U001(wWhere, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileChu, recCount) Then
            '    myDB.Rollback()
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。持出種別コード：{0}", MjK1.Cmn.Const.KeikiDwldSbtCd.Seko))
            '    Return False
            'End If
            'myDB.Commit()

            ''更新結果Trueの場合，更新件数をチェックする
            'If recCount = 0 Then
            '    '対象レコードがない場合、正常終了する
            '    Return True
            'End If
            '20240717 Rev038 DB再接続追加 DEL End  場所移動


            '/*'持出用テンポラリフォルダ/【内部変数.事業所コード】/MTDS
            'wTempMtdsFolder = System.IO.Path.Combine(Me._mtdsTemp, wDig4ZgsyoCd, C_MTDS_dir)
            '持出用テンポラリフォルダ（ダウンロード番号が一意のため、事業所コードでフォルダを分けない。）
            '*/
            '作業フォルダTOPを設定する
            wTempTopDir = System.IO.Path.Combine(Me._mtdsTemp, myPID)
            '持出用テンポラリフォルダ/MTDS
            wTempMtdsFolder = System.IO.Path.Combine(wTempTopDir, C_MTDS_dir)


            'フォルダが存在する場合，フォルダを削除する。
            If System.IO.Directory.Exists(wTempMtdsFolder) Then
                System.IO.Directory.Delete(wTempMtdsFolder, True)
            End If
            '「持出管理一時フォルダ」フォルダを作成する。
            System.IO.Directory.CreateDirectory(wTempMtdsFolder)


            'FETCH
            Dim wWhere As String = ""
            For Each dRow As DataRow In wDataTable.Rows
                'ダウンロード管理を読み込む
                wDownldRenno = CStr(Me.myDB.ConvDbNull(dRow("downld_renno"), ""))
                wKTTNMT_NO = CStr(Me.myDB.ConvDbNull(dRow("KTTNMT_NO"), ""))
                wKeikiUserId = CStr(Me.myDB.ConvDbNull(dRow("keiki_user_id"), ""))
                wKeikiDwldSbtCd = CStr(Me.myDB.ConvDbNull(dRow("keiki_dwld_sbt_cd"), ""))
                wDig4ZgsyoCd = CStr(Me.myDB.ConvDbNull(dRow("dig4_zgsyo_cd"), ""))


                '20240712 Rev038 DB再接続追加 ADD Start
                wWhere = "downld_renno=" & wDownldRenno
                'ダウンロード管理の検索したレコードの持出ファイル状態コードを"持出データ作成中"に更新する
                'トランザクションを開始する。
                Me.myDB.BeginTransaction()
                If Not dao.U001(wWhere, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileChu, recCount) Then
                    myDB.Rollback()
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。ダウンロード管理番号：{0}", wDownldRenno))
                    Return False
                End If
                myDB.Commit()

                '更新結果Trueの場合，更新件数をチェックする
                If recCount = 0 Then
                    '対象レコードがない場合、次のレコードに進める
                    Continue For
                End If
                '20240712 Rev038 DB再接続追加 ADD End

                '「持出ダウンロード番号一時フォルダ」フォルダを作成する。
                wTempMtdsRennoFolder = System.IO.Path.Combine(wTempMtdsFolder, wDownldRenno)
                System.IO.Directory.CreateDirectory(wTempMtdsRennoFolder)


                'ダウンロード番号に連係するレコードを検索しファイルに保存する。

                '取替_取替票(低圧)を作成する。
                'Rev002.1 MOD Start ライスセンター対応
                If Not CreateTRKE_TRKEHY_TIAT(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao, wPhaseKbn) Then
                    'Rev002.1 MOD End
                    Return False
                End If

                ' wPhaseKbn ' '：なし、'1'：１期、'2'：２期
                If wPhaseKbn = C_PHASE_KBN_1 Then
                    '取替票（低圧）が存在する
                    '取替_取替票行程を作成する。
                    If Not CreateTRKE_TRKEHY_PRC(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '特記事項を作成する。
                    If Not CreateKEIKI_SPOT_TKZK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '追加工費を作成する。
                    If Not CreateTRKE_TRKEHY_TUIKAKH(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '自主点検を作成する。
                    If Not CreateTRKE_CHECK_SHEET(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '添付書類詳細を作成する。
                    If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If
                ElseIf wPhaseKbn = C_PHASE_KBN_2 Then
                    '取替票（高圧他）低圧電圧が存在する（ライスセンター）

                    '取替_取替票行程（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateTRKE_TRKEHY_PRC_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '特記事項（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateKEIKI_SPOT_TKZK_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '追加工費（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateTRKE_TRKEHY_TUIKAKH_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '自主点検（高圧他）（低圧契約）を作成する。
                    If Not CreateTRKE_CHECK_SHEET_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '添付書類詳細を作成する。
                    If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                Else
                    '取替_取替票(高圧他)を作成する。　上記の切替によりここの処理はすべて高圧、特高になる
                    If Not CreateTRKE_TRKEHY_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao, wTrkhyExistFlag) Then
                        Return False
                    End If

                    If wTrkhyExistFlag Then
                        '取替票（高圧他）（高圧・特高契約）が存在する

                        '取替_取替票行程（高圧他）（高圧・特高契約）を作成する。
                        If Not CreateTRKE_TRKEHY_PRC_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '特記事項（高圧他）（高圧・特高契約）を作成する。
                        If Not CreateKEIKI_SPOT_TKZK_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '追加工費（高圧他）（高圧・特高契約）を作成する。
                        If Not CreateTRKE_TRKEHY_TUIKAKH_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '自主点検（高圧他）（低圧契約）を作成する。
                        If Not CreateTRKE_CHECK_SHEET_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '添付書類詳細を作成する。
                        If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If
                    Else
                        '低圧、高圧他（低圧契約）、高圧他のいずれもないのでエラー
                        Return False
                    End If
                End If


                '******************************************************
                '* ダウンロード用に作成したファイルを１つにまとめる。 *
                '******************************************************
                Dim wFname As String = wKTTNMT_NO & "_" & wKeikiUserId & "_" & wKeikiDwldSbtCd & "_" & wDownldRenno

                '【圧縮ファイル後ファイル名】：【持出グループ一時フォルダ】/【携帯端末番号】_【ユーザーID】_【ダウンロード識別コード】_【ダウンロード番号】.zip
                wZipFileName = System.IO.Path.Combine(wTempMtdsFolder, wFname + ".zip")
                '持出一時フォルダを圧縮する。
                HdSysBase.Cmn.CreateArchive(wZipFileName, wTempMtdsRennoFolder)

                'HdSysBase.Cmn.EncryptFile を呼出し，ファイルを暗号化する。
                If Not HdSysBase.Cmn.EncryptFile(wZipFileName, HdSysBase.Cmn.EncryptionKey, wZipFileName) Then
                    'ログ出力を行う。
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "暗号化に失敗しました：" & wZipFileName)
                    'Falseを返却して処理を終了する。
                    Return False
                End If

                '* NASへのコピー処理
                '【myPath.NAS.K1DAT】/to_tablet/【内部変数.事業所コード】/MTDS/
                wMtdsFolder = System.IO.Path.Combine(Me.myPath.NAS.K1DAT, "to_tablet", wDig4ZgsyoCd, C_MTDS_dir)

                If Not System.IO.Directory.Exists(wMtdsFolder) Then
                    '「持出フォルダ」がない場合，フォルダを作成する。
                    System.IO.Directory.CreateDirectory(wMtdsFolder)
                End If

                '持出圧縮ファイルを，コピーする。
                EcOrgIO.EcOrgFileIO.CopyFile(wZipFileName, System.IO.Path.Combine(wMtdsFolder, System.IO.Path.GetFileName(wZipFileName)))

                '持出ダウンロード番号一時フォルダを削除する。
                If System.IO.Directory.Exists(wTempMtdsRennoFolder) Then
                    System.IO.Directory.Delete(wTempMtdsRennoFolder, True)
                    System.IO.File.Delete(wZipFileName)
                End If
                '持出圧縮ファイルを削除する
                If System.IO.File.Exists(wZipFileName) Then
                    System.IO.File.Delete(wZipFileName)
                End If

                '１件ずつ、ダウンロード管理のレコードの持出ファイル状態コードを"持出データ作成完了"に更新する
                'トランザクションを開始する。
                Me.myDB.BeginTransaction()
                If Not dao.U002(wDownldRenno, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileCmp, recCount) Then
                    myDB.Rollback()
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。持出種別コード：{0}", MjK1.Cmn.Const.KeikiDwldSbtCd.Seko))
                    'Rollbackしたので、中止せず、次のレコード処理に進める
                    'Return False
                Else
                    myDB.Commit()
                End If
            Next

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を終了しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If

            ' 作業フォルダTOPを削除する
            If System.IO.Directory.Exists(wTempTopDir) Then
                System.IO.Directory.Delete(wTempTopDir, True)
                'EcOrgIO.EcOrgFileIO.DeleteDirectory(wTempTopDir, True)
            End If
        End Try

    End Function
#End Region


#Region "ダウンロードファイル作成（抜取検査）"

    ''' <summary>
    ''' ダウンロードファイル作成（抜取検査）
    ''' </summary>
    ''' <param name="pZGSYO_CD_List">対象事業所CD</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMtdsFileNKTRKNS(ByVal pZGSYO_CD_List As List(Of String), ByVal dao As MJ4103_DAO) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader            'DB Reader
        Dim recCount As Double = 0

        Dim wDownldRenno As String = ""
        Dim wKTTNMT_NO As String = ""
        Dim wKeikiUserId As String = ""
        Dim wKeikiDwldSbtCd As String = ""
        Dim wMtdsFileZtiCd As String = ""

        Dim wDig4ZgsyoCd As String = ""
        Dim wNendo As String = ""
        Dim wTrkhyKbn As String = ""
        Dim wTrkhyNo As String = ""


        Dim wTempMtdsFolder As String = ""          '持出一時
        Dim wTempMtdsRennoFolder As String = ""     '持出一時ダウンロード番号フォルダ
        Dim wMtdsFolder As String = ""              'NAS側持出フォルダ

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名
        Dim wZipFileName As String = ""

        Dim firstFlag As Boolean = True

        Dim wTempTopDir As String = ""

        'Dim wTrkhyExistFlag = False     '取替票を検出し存在すればTrue,なければFalse

        'Rev002-Start
        Dim wPhaseKbn = C_PHASE_KBN_NON     '取替票を検出し存在しなければ" ",低圧は"TIAT",高圧他は"KAHK"
        Dim wTrkhyExistFlag As Boolean      '取替票（高圧他）の有無
        'Rev002-End


        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            ' ダウンロード管理データ一覧を取得する。
            If Not dao.S001(dbr, MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns, pZGSYO_CD_List) Then
                Throw New Exception(String.Format("抜取検査：ダウンロード管理データ一覧取得処理でエラーが発生しました。"))
            End If
            '検索結果０件の場合、処理を終了する。
            If Not dbr.DataStruct.HasRows Then
                '対象レコードがない場合、正常終了する
                Return True
            End If

            'PostgreSQLは、selectの多重ができないため、結果をDataTableに設定し、dbrを開放する
            wDataTable = New DataTable
            'カラム名設定
            Dim i As Integer
            For i = 0 To dbr.DataStruct.FieldCount - 1
                wDataTable.Columns.Add(dbr.DataStruct.GetName(i))
            Next
            'DataTableにデータ保存
            While (dbr.DataStruct.Read)
                wRow = wDataTable.NewRow
                For i = 0 To dbr.DataStruct.FieldCount - 1
                    wRow(dbr.DataStruct.GetName(i)) = dbr.DataStruct.Item(i).ToString
                Next
                wDataTable.Rows.Add(wRow)
            End While
            dbr.Close()

            '20240717 Rev038 DB再接続追加 DEL Start  場所移動
            ''検索結果のキー情報を構築する。
            'Dim wKey As New List(Of String)
            'Dim wKeyIN As String = ""

            'firstFlag = True
            'For Each d As DataRow In wDataTable.Rows
            '    If firstFlag Then
            '        firstFlag = False
            '    Else
            '        wKeyIN = wKeyIN & ","
            '    End If

            '    wKeyIN = wKeyIN & CStr(Me.myDB.ConvDbNull(d("downld_renno"), ""))

            '    'リストにも保存
            '    wKey.Add(CStr(Me.myDB.ConvDbNull(d("downld_renno"), "")))
            'Next
            'Dim wWhere As String = "downld_renno IN (" & wKeyIN & ")"


            ''ダウンロード管理の検索したレコードの持出ファイル状態コードを"持出データ作成中"に更新する
            ''トランザクションを開始する。
            'Me.myDB.BeginTransaction()
            'If Not dao.U001(wWhere, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileChu, recCount) Then
            '    myDB.Rollback()
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。持出種別コード：{0}", MjK1.Cmn.Const.KeikiDwldSbtCd.Seko))
            '    Return False
            'End If
            'myDB.Commit()

            ''更新結果Trueの場合，更新件数をチェックする
            'If recCount = 0 Then
            '    '対象レコードがない場合、正常終了する
            '    Return True
            'End If
            '20240717 Rev038 DB再接続追加 DEL Start  場所移動


            '/*'持出用テンポラリフォルダ/【内部変数.事業所コード】/MTDS
            'wTempMtdsFolder = System.IO.Path.Combine(Me._mtdsTemp, wDig4ZgsyoCd, C_MTDS_dir)
            '持出用テンポラリフォルダ（ダウンロード番号が一意のため、事業所コードでフォルダを分けない。）
            'wTempMtdsFolder = Me._mtdsTemp
            '*/
            '作業フォルダTOPを設定する
            wTempTopDir = System.IO.Path.Combine(Me._mtdsTemp, myPID)
            '持出用テンポラリフォルダ/MTDS
            wTempMtdsFolder = System.IO.Path.Combine(wTempTopDir, C_MTDS_dir)

            'ファイルが存在する場合，フォルダを削除する。
            If System.IO.Directory.Exists(wTempMtdsFolder) Then
                System.IO.Directory.Delete(wTempMtdsFolder, True)
            End If
            '「持出管理一時フォルダ」フォルダを作成する。
            System.IO.Directory.CreateDirectory(wTempMtdsFolder)


            'FETCH
            Dim wWhere As String = ""
            For Each dRow As DataRow In wDataTable.Rows
                'ダウンロード管理を読み込む
                wDownldRenno = CStr(Me.myDB.ConvDbNull(dRow("downld_renno"), ""))
                wKTTNMT_NO = CStr(Me.myDB.ConvDbNull(dRow("KTTNMT_NO"), ""))
                wKeikiUserId = CStr(Me.myDB.ConvDbNull(dRow("keiki_user_id"), ""))
                wKeikiDwldSbtCd = CStr(Me.myDB.ConvDbNull(dRow("keiki_dwld_sbt_cd"), ""))
                wDig4ZgsyoCd = CStr(Me.myDB.ConvDbNull(dRow("dig4_zgsyo_cd"), ""))

                '20240712 Rev038 DB再接続追加 ADD Start
                wWhere = "downld_renno=" & wDownldRenno
                'ダウンロード管理の検索したレコードの持出ファイル状態コードを"持出データ作成中"に更新する
                'トランザクションを開始する。
                Me.myDB.BeginTransaction()
                If Not dao.U001(wWhere, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileChu, recCount) Then
                    myDB.Rollback()
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。ダウンロード管理番号：{0}", wDownldRenno))
                    Return False
                End If
                myDB.Commit()

                '更新結果Trueの場合，更新件数をチェックする
                If recCount = 0 Then
                    '対象レコードがない場合、次のレコードに進める
                    Continue For
                End If
                '20240712 Rev038 DB再接続追加 ADD End

                '「持出ダウンロード番号一時フォルダ」フォルダを作成する。
                wTempMtdsRennoFolder = System.IO.Path.Combine(wTempMtdsFolder, wDownldRenno)
                System.IO.Directory.CreateDirectory(wTempMtdsRennoFolder)


                'ダウンロード番号に連係するレコードを検索しファイルに保存する。

                'Rev002-Start
                '取替_取替票(低圧)を作成する。
                If Not CreateTRKE_TRKEHY_TIAT(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao, wPhaseKbn) Then
                    Return False
                End If

                ' wPhaseKbn ' '：なし、'1'：１期、'2'：２期
                If wPhaseKbn = C_PHASE_KBN_1 Then
                    'If wTrkhyExistFlag Then
                    '取替票（低圧）が存在する
                    '取替_取替票行程を作成する。
                    If Not CreateTRKE_TRKEHY_PRC(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '特記事項を作成する。
                    If Not CreateKEIKI_SPOT_TKZK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '追加工費を作成する。
                    If Not CreateTRKE_TRKEHY_TUIKAKH(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '自主点検を作成する。
                    If Not CreateTRKE_CHECK_SHEET(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '添付書類詳細を作成する。
                    If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                ElseIf wPhaseKbn = C_PHASE_KBN_2 Then
                    '取替票（高圧他）低圧電圧が存在する（ライスセンター）

                    '取替_取替票行程（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateTRKE_TRKEHY_PRC_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '特記事項（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateKEIKI_SPOT_TKZK_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '追加工費（高圧他）（高圧・特高契約）を作成する。
                    If Not CreateTRKE_TRKEHY_TUIKAKH_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '自主点検（高圧他）（低圧契約）を作成する。
                    If Not CreateTRKE_CHECK_SHEET_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If

                    '添付書類詳細を作成する。
                    If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                        Return False
                    End If


                Else
                    '取替_取替票(高圧他)（高圧）を作成する。
                    If Not CreateTRKE_TRKEHY_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao, wTrkhyExistFlag) Then
                        Return False
                    End If

                    If wTrkhyExistFlag Then
                        '取替票（高圧他）（高圧）が存在する

                        '取替_取替票行程を作成する。
                        If Not CreateTRKE_TRKEHY_PRC_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '特記事項を作成する。
                        If Not CreateKEIKI_SPOT_TKZK_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '追加工費を作成する。
                        If Not CreateTRKE_TRKEHY_TUIKAKH_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '自主点検を作成する。
                        If Not CreateTRKE_CHECK_SHEET_KAHK(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If

                        '添付書類詳細を作成する。
                        If Not CreateCOM_TNPDOC_DTL(wTempMtdsRennoFolder, wDownldRenno, wKeikiDwldSbtCd, dao) Then
                            Return False
                        End If
                    Else
                        '低圧、高圧他（低圧契約）、高圧他のいずれもないのでエラー
                        Return False
                    End If

                End If


                'Rev002-End

                '******************************************************
                '* ダウンロード用に作成したファイルを１つにまとめる。 *
                '******************************************************
                Dim wFname As String = wKTTNMT_NO & "_" & wKeikiUserId & "_" & wKeikiDwldSbtCd & "_" & wDownldRenno

                '【圧縮ファイル後ファイル名】：【持出グループ一時フォルダ】/【携帯端末番号】_【ユーザーID】_【ダウンロード識別コード】_【ダウンロード番号】.zip
                wZipFileName = System.IO.Path.Combine(wTempMtdsFolder, wFname + ".zip")
                '持出一時フォルダを圧縮する。
                HdSysBase.Cmn.CreateArchive(wZipFileName, wTempMtdsRennoFolder)

                'HdSysBase.Cmn.EncryptFile を呼出し，ファイルを暗号化する。
                If Not HdSysBase.Cmn.EncryptFile(wZipFileName, HdSysBase.Cmn.EncryptionKey, wZipFileName) Then
                    'ログ出力を行う。
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "暗号化に失敗しました：" & wZipFileName)
                    'Falseを返却して処理を終了する。
                    Return False
                End If

                '* NASへのコピー処理
                '【myPath.NAS.K1DAT】/to_tablet/【内部変数.事業所コード】/MTDS/
                wMtdsFolder = System.IO.Path.Combine(Me.myPath.NAS.K1DAT, "to_tablet", wDig4ZgsyoCd, C_MTDS_dir)

                If Not System.IO.Directory.Exists(wMtdsFolder) Then
                    '「持出フォルダ」がない場合，フォルダを作成する。
                    System.IO.Directory.CreateDirectory(wMtdsFolder)
                End If

                '持出圧縮ファイルを，コピーする。
                EcOrgIO.EcOrgFileIO.CopyFile(wZipFileName, System.IO.Path.Combine(wMtdsFolder, System.IO.Path.GetFileName(wZipFileName)))

                '持出ダウンロード番号一時フォルダを削除する。
                If System.IO.Directory.Exists(wTempMtdsRennoFolder) Then
                    System.IO.Directory.Delete(wTempMtdsRennoFolder, True)
                    System.IO.File.Delete(wZipFileName)
                End If
                '持出圧縮ファイルを削除する
                If System.IO.File.Exists(wZipFileName) Then
                    System.IO.File.Delete(wZipFileName)
                End If

                '１件ずつ、ダウンロード管理のレコードの持出ファイル状態コードを"持出データ作成完了"に更新する
                'トランザクションを開始する。
                Me.myDB.BeginTransaction()
                If Not dao.U002(wDownldRenno, MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileCmp, recCount) Then
                    myDB.Rollback()
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("'持出処理中'に更新する処理でエラーが発生しました。持出種別コード：{0}", MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns))
                    'Rollbackしたので、中止せず、次のレコード処理に進める
                    'Return False
                Else
                    myDB.Commit()
                End If
            Next

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を終了しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If

            ' 作業フォルダTOPを削除する
            If System.IO.Directory.Exists(wTempTopDir) Then
                System.IO.Directory.Delete(wTempTopDir, True)
            End If
        End Try

    End Function
#End Region


#Region "取替_取替票（低圧）ファイル作成処理"
    ''' <summary>
    ''' 取替_取替票（低圧）ファイル作成処理(施工)
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_TIAT(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByRef pPhaseKbn As String) As Boolean
        '施工：抜取検査フラグにFalseを設定して実行する
        Return CreateTRKE_TRKEHY_TIAT(pMtdsMngNoFolder, pDownldRenno, pKeikiDwldSbtCd, dao, False, pPhaseKbn)
    End Function


    ''' <summary>
    ''' 取替_取替票（低圧）ファイル作成処理(抜取)
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_TIAT_Nktr(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByRef pPhaseKbn As String) As Boolean
        '施工：抜取検査フラグにFalseを設定して実行する
        Return CreateTRKE_TRKEHY_TIAT(pMtdsMngNoFolder, pDownldRenno, pKeikiDwldSbtCd, dao, True, pPhaseKbn)
    End Function

    ''' <summary>
    ''' 取替_取替票（低圧）ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <param name="NktrKnsFlag">抜取検査フラグ</param>
    ''' <param name="PhaseKbn">低圧か高圧他の区分</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_TIAT(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByVal NktrKnsFlag As Boolean, ByRef PhaseKbn As String) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・取替_取替票（低圧）を取得する。
            If Not dao.S002(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・取替_取替票（低圧）検索でエラーが発生しました。"))
            End If

            'Rev002-Start

            '低圧の取替票データが無い場合はTrueで戻るが、テーブルが存在しない（ExistFlag＝False）で上位に知らせる
            If Not dbr.DataStruct.HasRows Then
                PhaseKbn = C_PHASE_KBN_NON    'テーブルなし
                Return True
            End If

            'テーブルはなしにセットする
            PhaseKbn = C_PHASE_KBN_NON

            'Rev002-End

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If

            'FETCH
            While (dbr.DataStruct.Read)

                'Rev002.1 ADD Start
                '１期か２期を格納する
                PhaseKbn = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("phase_kbn"), ""))
                '取替票ファイル用Datatableの初期化
                wDataTable = New DataTable

                If PhaseKbn.Equals(C_PHASE_KBN_1) Then
                    initDtblTRKE_TRKEHY_TIAT(wDataTable)
                    PhaseKbn = C_PHASE_KBN_1
                Else
                    initDtblTRKE_TRKEHY_KAHK(wDataTable)
                    PhaseKbn = C_PHASE_KBN_2
                End If
                'Datatableの初期化
                wDataTable.Rows.Clear()
                'Rev002.1 ADD End

                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("file_sbt") = wFileSbt
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("prcmg_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prcmg_cd"), ""))
                wRow("trke_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_sbt_cd"), ""))
                wRow("kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kzkis_mdgt_cd"), ""))
                wRow("kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kozitn_cd"), ""))
                wRow("skosya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skosya_cd"), ""))
                wRow("skosya_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skosya_ms"), ""))
                wRow("kiyk_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_no"), ""))
                wRow("kiyk_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_sbt_cd"), ""))
                wRow("kiyk_cs_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_cs_mskn"), ""))
                wRow("kiyk_cs_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_cs_ms"), ""))
                wRow("dig14_cs_tel") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig14_cs_tel_str"), ""))
                wRow("cs_tdhkn_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_tdhkn_add_cd"), ""))
                wRow("cs_siku_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_siku_add_cd"), ""))
                wRow("cs_oazat_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_oazat_add_cd"), ""))
                wRow("cs_azatm_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_azatm_add_cd"), ""))
                wRow("tdhkn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_tdhkn_add_cd_ms"), ""))
                wRow("siku_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_siku_add_cd_ms"), ""))
                wRow("oazat_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_oazat_add_cd_ms"), ""))
                wRow("azatm_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_azatm_add_cd_ms"), ""))
                wRow("cs_addhsk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_addhsk_naiyo"), ""))
                wRow("mkhy_senro_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_senro_cd"), ""))
                wRow("mkhy_kansn_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_kansn_no"), ""))
                wRow("mkhy_bunk1_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk1_no"), ""))
                wRow("mkhy_bunk2_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk2_no"), ""))
                wRow("mkhy_bunk3_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk3_no"), ""))
                wRow("stbit_xzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stbit_xzahyo"), ""))
                wRow("stbit_yzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stbit_yzahyo"), ""))
                wRow("iewku_stdp_xzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("iewku_stdp_xzahyo"), ""))
                wRow("iewku_stdp_yzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("iewku_stdp_yzahyo"), ""))
                wRow("cs_kiyk_dnryk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_kiyk_dnryk"), ""))
                wRow("std_kensn_dd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("std_kensn_dd"), ""))
                wRow("kensn_yotei_togt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_togt_md"), ""))
                wRow("kensn_yotei_ykgt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_ykgt_md"), ""))
                wRow("kensn_yotei_yygt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_yygt_md"), ""))
                wRow("kyokyu_spot_tokti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kyokyu_spot_tokti_no"), ""))
                wRow("hist_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hist_flg"), ""))
                wRow("brot_sett_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("brot_sett_flg"), ""))
                wRow("knti_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("knti_sbt_cd"), ""))
                wRow("surge_yksi_siyo_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("surge_yksi_siyo_umu_flg"), ""))
                wRow("tidn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tidn_kbn"), ""))
                wRow("tnsb_reuse_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsb_reuse_kbn"), ""))
                wRow("trke_tis_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_tis_abms"), ""))
                wRow("kohu_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kohu_ymd"), ""))
                wRow("skohu_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skohu_ymd"), ""))
                wRow("hiky_yotei_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hiky_yotei_ymd"), ""))
                wRow("sksti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sksti_ymd"), ""))
                wRow("skssti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skssti_ymd"), ""))
                wRow("skyti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skyti_ymd"), ""))
                wRow("syun_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("syun_ymd"), ""))
                wRow("keiki_sitei_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sitei_no"), ""))
                wRow("nktr_kensa_zyokyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktr_kensa_zyokyo_cd"), ""))
                If NktrKnsFlag Then
                    '抜取検査の場合、抜取検査タブレット受信年月日・抜取検査タブレット受信_担当者コードを設定する
                    wRow("ntkns_tblt_recv_ymd") = SyoriYMD
                    wRow("ntkns_tblt_recv_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("D_keiki_user_id"), ""))
                Else
                    wRow("ntkns_tblt_recv_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_tblt_recv_ymd"), ""))
                    wRow("ntkns_tblt_recv_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_tblt_recv_tntsy_cd"), ""))
                End If
                wRow("ntkns_zissi_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_zissi_ymd"), ""))
                wRow("ntkns_zissi_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_zissi_tntsy_cd"), ""))
                wRow("ntkns_kka_upld_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_kka_upld_ymd"), ""))
                wRow("ntkns_kka_upld_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ntkns_kka_upld_tntsy_cd"), ""))
                wRow("nktr_kensa_zyokyo_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktr_kensa_zyokyo_abms"), ""))
                wRow("tkkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_id"), ""))
                wRow("tkkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kesbt_kbn"), ""))
                wRow("tkkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_krhsk_cd"), ""))
                wRow("tkkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_yr"), ""))
                wRow("tkkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tshsk_cd"), ""))
                wRow("tkkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ktsk_ms"), ""))
                wRow("tkkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ktsk_cd"), ""))
                wRow("tkkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_taiko_kbn"), ""))
                wRow("tkkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_knthk_cd"), ""))
                wRow("tkkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_ts_kino_umu_flg"), ""))
                wRow("tkkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_khk_kino_umu_flg"), ""))
                wRow("tkkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gaibu_output_umu_flg"), ""))
                wRow("tkkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_yuko_kigen_ym"), ""))
                wRow("tkkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_seizo_yy"), ""))
                wRow("tkkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_zyrt"), ""))
                wRow("tkkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_digsu"), ""))
                wRow("tkkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_szsya_cd"), ""))
                wRow("tkkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_yy"), ""))
                wRow("tkkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_zytr_szsu"), ""))
                wRow("tkkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gytr_szsu"), ""))
                wRow("no1_tkkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_tkkk_hnsik_seizo_no"), ""))
                wRow("no2_tkkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_tkkk_hnsik_seizo_no"), ""))
                wRow("tkkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_yuko_kigen_ym"), ""))
                wRow("tkkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_gknti_no"), ""))
                wRow("tkkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_mado_su"), ""))
                wRow("tkkk_smkey1") = ""
                wRow("tkkk_smkey2") = ""
                wRow("tkkk_smkey_startdate") = ""
                wRow("tkkk_smkey_enddate") = ""
                wRow("tkkk_mac_address") = ""
                wRow("tkkk_tsn_id") = ""
                wRow("ttkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_id"), ""))
                wRow("ttkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_kesbt_kbn"), ""))
                wRow("ttkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_krhsk_cd"), ""))
                wRow("ttkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_yr"), ""))
                wRow("ttkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tshsk_cd"), ""))
                wRow("ttkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ktsk_ms"), ""))
                wRow("ttkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ktsk_cd"), ""))
                wRow("ttkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_taiko_kbn"), ""))
                wRow("ttkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_knthk_cd"), ""))
                wRow("ttkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_ts_kino_umu_flg"), ""))
                wRow("ttkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_khk_kino_umu_flg"), ""))
                wRow("ttkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gaibu_output_umu_flg"), ""))
                wRow("ttkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_yuko_kigen_ym"), ""))
                wRow("ttkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_seizo_yy"), ""))
                wRow("ttkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_zyrt"), ""))
                wRow("ttkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_digsu"), ""))
                wRow("ttkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_szsya_cd"), ""))
                wRow("ttkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_yy"), ""))
                wRow("ttkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_zytr_szsu"), ""))
                wRow("ttkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gytr_szsu"), ""))
                wRow("no1_ttkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_ttkk_hnsik_seizo_no"), ""))
                wRow("no2_ttkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_ttkk_hnsik_seizo_no"), ""))
                wRow("ttkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_yuko_kigen_ym"), ""))
                wRow("ttkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_gknti_no"), ""))
                wRow("ttkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_mado_su"), ""))
                wRow("kdkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_id"), ""))
                wRow("kdkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_kesbt_kbn"), ""))
                wRow("kdkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_krhsk_cd"), ""))
                wRow("kdkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_yr"), ""))
                wRow("kdkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tshsk_cd"), ""))
                wRow("kdkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ktsk_ms"), ""))
                wRow("kdkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ktsk_cd"), ""))
                wRow("kdkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_taiko_kbn"), ""))
                wRow("kdkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_knthk_cd"), ""))
                wRow("kdkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_ts_kino_umu_flg"), ""))
                wRow("kdkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_khk_kino_umu_flg"), ""))
                wRow("kdkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_gaibu_output_umu_flg"), ""))
                wRow("kdkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_yuko_kigen_ym"), ""))
                wRow("kdkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_seizo_yy"), ""))
                wRow("kdkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_zyrt"), ""))
                wRow("kdkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_digsu"), ""))
                wRow("kdkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_szsya_cd"), ""))
                wRow("kdkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_yy"), ""))
                wRow("kdkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_zytr_szsu"), ""))
                wRow("kdkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_gytr_szsu"), ""))
                wRow("no1_kdkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_kdkk_hnsik_seizo_no"), ""))
                wRow("no2_kdkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_kdkk_hnsik_seizo_no"), ""))
                wRow("kdkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_yuko_kigen_ym"), ""))
                wRow("kdkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_gknti_no"), ""))
                wRow("kdkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_mado_su"), ""))
                wRow("tdnstr_ts_dosa_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_ts_dosa_kbn"), ""))
                wRow("tdnstr_tdnt_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_tdnt_su"), ""))
                wRow("tdnstr_tdnstr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_tdnstr_time"), ""))
                wRow("hkgds_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_kbn"), ""))
                wRow("hkgds_start_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_start_md"), ""))
                wRow("hkgds_end_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_end_md"), ""))
                wRow("no1_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_hkgds_tdn_hms"), ""))
                wRow("no1_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_hkgds_sydan_hms"), ""))
                wRow("no2_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_hkgds_tdn_hms"), ""))
                wRow("no2_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_hkgds_sydan_hms"), ""))
                wRow("no3_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_hkgds_tdn_hms"), ""))
                wRow("no3_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_hkgds_sydan_hms"), ""))
                wRow("no4_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_hkgds_tdn_hms"), ""))
                wRow("no4_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_hkgds_sydan_hms"), ""))
                wRow("no5_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_hkgds_tdn_hms"), ""))
                wRow("no5_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_hkgds_sydan_hms"), ""))
                wRow("no6_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no6_hkgds_tdn_hms"), ""))
                wRow("no6_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no6_hkgds_sydan_hms"), ""))
                wRow("no7_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no7_hkgds_tdn_hms"), ""))
                wRow("no7_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no7_hkgds_sydan_hms"), ""))
                wRow("no8_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no8_hkgds_tdn_hms"), ""))
                wRow("no8_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no8_hkgds_sydan_hms"), ""))
                wRow("no9_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no9_hkgds_tdn_hms"), ""))
                wRow("no9_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no9_hkgds_sydan_hms"), ""))
                wRow("no10_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no10_hkgds_tdn_hms"), ""))
                wRow("no10_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no10_hkgds_sydan_hms"), ""))
                wRow("kbttd_start_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kbttd_start_date"), ""))
                wRow("kbttd_end_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kbttd_end_date"), ""))
                wRow("hoka_kihi_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_kihi_kbn_cd"), ""))
                wRow("hoka_gmn_flckr_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_gmn_flckr_kbn"), ""))
                wRow("hoka_evnt_krk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_evnt_krk_kbn"), ""))
                wRow("hoka_hoka_hyz_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_hoka_hyz_kbn"), ""))
                wRow("hoka_hoka_hyz_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_hoka_hyz_value"), ""))
                wRow("hksgk_hksgn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_hksgn_kbn"), ""))
                wRow("hksgk_hka_dnr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_hka_dnr"), ""))
                wRow("hksgk_attny_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_time"), ""))
                wRow("hksgk_attny_kaisu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_kaisu"), ""))
                wRow("hksgk_attny_kaisu_clr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_kaisu_clr_time"), ""))
                wRow("hksgr_hksgn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_hksgn_kbn"), ""))
                wRow("hksgr_hka_dnr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_hka_dnr"), ""))
                wRow("hksgr_attny_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_time"), ""))
                wRow("hksgr_attny_kaisu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_kaisu"), ""))
                wRow("hksgr_attny_kaisu_clr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_kaisu_clr_time"), ""))
                wRow("khkmk_keiki_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_keiki_kbn"), ""))
                wRow("khkmk_mg_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_kbn"), ""))
                wRow("khkmk_hnsik_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_hnsik_kbn"), ""))
                wRow("khkmk_wrms_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_wrms_kbn"), ""))
                wRow("khkmk_mg_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_umu_flg"), ""))
                wRow("khkmk_mg_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_yr"), ""))
                wRow("tuika_kohi_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_umu_flg"), ""))
                wRow("zippi_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zippi_umu_flg"), ""))
                wRow("no1_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_no"), ""))
                wRow("no1_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kbn"), ""))
                wRow("no1_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kngk"), ""))
                wRow("no2_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_no"), ""))
                wRow("no2_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kbn"), ""))
                wRow("no2_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kngk"), ""))
                wRow("no3_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_no"), ""))
                wRow("no3_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kbn"), ""))
                wRow("no3_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kngk"), ""))
                wRow("tenp_file_mng_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tenp_file_mng_no"), ""))
                wRow("ryssy_tenp_file_mng_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ryssy_tenp_file_mng_no"), ""))
                wRow("rrzk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("rrzk_naiyo"), ""))
                wRow("prcmg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prcmg_cd_abms"), ""))
                wRow("trke_sbt_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_sbt_abms"), ""))
                wRow("brot_sett_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("brot_sett_flg_abms"), ""))
                wRow("tkkk_kesbt_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kesbt_kbn_abms"), ""))
                wRow("tkkk_krhsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_krhsk_cd_abms"), ""))
                wRow("tkkk_tshsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tshsk_cd_abms"), ""))
                wRow("tkkk_knthk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_knthk_cd_abms"), ""))
                wRow("tkkk_ts_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_ts_kino_umu_flg_abms"), ""))
                wRow("tkkk_khk_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_khk_kino_umu_flg_abms"), ""))
                wRow("tkkk_gaibu_output_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gaibu_output_umu_flg_abms"), ""))
                wRow("tkkk_taiko_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_taiko_kbn_abms"), ""))
                wRow("ttkk_kesbt_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_kesbt_kbn_abms"), ""))
                wRow("ttkk_krhsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_krhsk_cd_abms"), ""))
                wRow("ttkk_tshsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tshsk_cd_abms"), ""))
                wRow("ttkk_knthk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_knthk_abms"), ""))
                wRow("ttkk_ts_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_ts_kino_umu_flg_abms"), ""))
                wRow("ttkk_khk_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_khk_kino_umu_flg_abms"), ""))
                wRow("ttkk_gaibu_output_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gaibu_output_umu_flg_ms"), ""))
                wRow("ttkk_taiko_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_taiko_kbn_abms"), ""))
                wRow("knti_sbt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("knti_sbt_cd_ms"), ""))
                wRow("surge_yksi_siyo_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("surge_yksi_siyo_umu_flg_abms"), ""))
                wRow("tidn_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tidn_kbn_abms"), ""))
                wRow("tnsb_reuse_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsb_reuse_kbn_abms"), ""))
                wRow("khkmk_keiki_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_keiki_kbn_abms"), ""))
                wRow("khkmk_mg_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_kbn_abms"), ""))
                wRow("khkmk_hnsik_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_hnsik_kbn_abms"), ""))
                wRow("khkmk_wrms_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_wrms_kbn_abms"), ""))
                wRow("khkmk_mg_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_umu_flg_abms"), ""))
                wRow("no1_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kbn_abms"), ""))
                wRow("no2_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kbn_abms"), ""))
                wRow("no3_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kbn_abms"), ""))
                wRow("seko_kka_send_ymd") = ""
                wRow("seko_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("hktg_yohi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("hktg_zumi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mdgt_ms"), ""))
                wRow("kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kozitn_ms"), ""))
                wRow("senro_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("senro_ms"), ""))
                wRow("tykei_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tykei_mdgt_ms"), ""))
                wRow("tykei_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tykei_kozitn_ms"), ""))
                wRow("yukoWhZytrTime") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yukoWhZytrTime"), ""))
                wRow("yukoWhGytrTime") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yukoWhGytrTime"), ""))
                wRow("nktrk_sizsk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizsk_kbn"), ""))
                wRow("smj_send_zumi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff

                'Rev002.1 ADD Start
                '２期専用の項目を追加
                If PhaseKbn.Equals(C_PHASE_KBN_2) Then
                    wRow("kiyk_kbn_cd") = C_KIYK_KBN_CD_1
                    wRow("keiki_kbn_cd") = C_KEIKI_KBN_CD_1
                    wRow("sm_zyri_kbn_cd") = C_SM_ZYRI_KBN_CD_1
                End If
                'Rev002.1 ADD End
                'Rev002.1 ADD Start 20241204 移動
                wRow("trke_tis_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_tis_kbn"), ""))
                'Rev002.1 ADD End

                'Rev040 2024/12/03 物理分割 次世代QR対応 ADD Start
                wRow("tkkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_ym"), ""))
                wRow("tkkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_sbt_cd"), ""))
                wRow("tkkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("tkkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_yr"), ""))
                wRow("tkkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_no"), ""))
                wRow("tkkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_ktsk_ms"), ""))
                wRow("tkkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_kzskbt_kbn"), ""))
                wRow("tkkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_szsya_mng_value"), ""))
                wRow("ttkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_ym"), ""))
                wRow("ttkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_sbt_cd"), ""))
                wRow("ttkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("ttkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_yr"), ""))
                wRow("ttkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_no"), ""))
                wRow("ttkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_ktsk_ms"), ""))
                wRow("ttkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_kzskbt_kbn"), ""))
                wRow("ttkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_szsya_mng_value"), ""))
                wRow("kdkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_ym"), ""))
                wRow("kdkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_sbt_cd"), ""))
                wRow("kdkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("kdkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_yr"), ""))
                wRow("kdkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_no"), ""))
                wRow("kdkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_ktsk_ms"), ""))
                wRow("kdkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_kzskbt_kbn"), ""))
                wRow("kdkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_szsya_mng_value"), ""))
                wRow("stzk_sdnsv_menu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stzk_sdnsv_menu_cd"), ""))
                'Rev040 2024/12/03 物理分割 次世代QR対応 ADD End


                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            'Rev002.1 UPD Start
            If PhaseKbn.Equals(C_PHASE_KBN_1) Then
                wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_TIAT)
            Else
                wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_KAHK)
            End If
            'Rev002.1 UPD End

            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("取替票ファイル出力に失敗しました。"))
            End If

            'アップロード時にアップロード情報と併せて登録するので、ここでは設定しない
            'If NktrKnsFlag Then
            '    '抜取検査の場合、タブレット受信年月日を設定する
            '    dbr.Close()
            '    dbr = Nothing
            '    updNtknsTbltRecvYmd(wDataTable)
            'End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_TRKEHY_TIAT(ByVal pDataTable As System.Data.DataTable)
        ' 空テーブルを作る
        pDataTable.Columns.Add("saksi_date")
        pDataTable.Columns.Add("upd_date")
        pDataTable.Columns.Add("file_sbt")
        pDataTable.Columns.Add("dig4_zgsyo_cd")
        pDataTable.Columns.Add("trkhy_hakko_nendo")
        pDataTable.Columns.Add("trkhy_kbn")
        pDataTable.Columns.Add("trkhy_no")
        pDataTable.Columns.Add("prcmg_cd")
        pDataTable.Columns.Add("trke_sbt_cd")
        pDataTable.Columns.Add("kzkis_mdgt_cd")
        pDataTable.Columns.Add("kozitn_cd")
        pDataTable.Columns.Add("skosya_cd")
        pDataTable.Columns.Add("skosya_ms")
        pDataTable.Columns.Add("kiyk_no")
        pDataTable.Columns.Add("kiyk_sbt_cd")
        pDataTable.Columns.Add("kiyk_cs_mskn")
        pDataTable.Columns.Add("kiyk_cs_ms")
        pDataTable.Columns.Add("dig14_cs_tel")
        pDataTable.Columns.Add("cs_tdhkn_add_cd")
        pDataTable.Columns.Add("cs_siku_add_cd")
        pDataTable.Columns.Add("cs_oazat_add_cd")
        pDataTable.Columns.Add("cs_azatm_add_cd")
        pDataTable.Columns.Add("tdhkn_ms")
        pDataTable.Columns.Add("siku_ms")
        pDataTable.Columns.Add("oazat_ms")
        pDataTable.Columns.Add("azatm_ms")
        pDataTable.Columns.Add("cs_addhsk_naiyo")
        pDataTable.Columns.Add("mkhy_senro_cd")
        pDataTable.Columns.Add("mkhy_kansn_no")
        pDataTable.Columns.Add("mkhy_bunk1_no")
        pDataTable.Columns.Add("mkhy_bunk2_no")
        pDataTable.Columns.Add("mkhy_bunk3_no")
        pDataTable.Columns.Add("stbit_xzahyo")
        pDataTable.Columns.Add("stbit_yzahyo")
        pDataTable.Columns.Add("iewku_stdp_xzahyo")
        pDataTable.Columns.Add("iewku_stdp_yzahyo")
        pDataTable.Columns.Add("cs_kiyk_dnryk")
        pDataTable.Columns.Add("std_kensn_dd")
        pDataTable.Columns.Add("kensn_yotei_togt_md")
        pDataTable.Columns.Add("kensn_yotei_ykgt_md")
        pDataTable.Columns.Add("kensn_yotei_yygt_md")
        pDataTable.Columns.Add("kyokyu_spot_tokti_no")
        pDataTable.Columns.Add("hist_flg")
        pDataTable.Columns.Add("brot_sett_flg")
        pDataTable.Columns.Add("knti_sbt_cd")
        pDataTable.Columns.Add("surge_yksi_siyo_umu_flg")
        pDataTable.Columns.Add("tidn_kbn")
        pDataTable.Columns.Add("tnsb_reuse_kbn")
        pDataTable.Columns.Add("trke_tis_abms")
        pDataTable.Columns.Add("kohu_ymd")
        pDataTable.Columns.Add("skohu_ymd")
        pDataTable.Columns.Add("hiky_yotei_ymd")
        pDataTable.Columns.Add("sksti_ymd")
        pDataTable.Columns.Add("skssti_ymd")
        pDataTable.Columns.Add("skyti_ymd")
        pDataTable.Columns.Add("syun_ymd")
        pDataTable.Columns.Add("keiki_sitei_no")
        pDataTable.Columns.Add("nktr_kensa_zyokyo_cd")
        pDataTable.Columns.Add("ntkns_tblt_recv_ymd")
        pDataTable.Columns.Add("ntkns_tblt_recv_tntsy_cd")
        pDataTable.Columns.Add("ntkns_zissi_ymd")
        pDataTable.Columns.Add("ntkns_zissi_tntsy_cd")
        pDataTable.Columns.Add("ntkns_kka_upld_ymd")
        pDataTable.Columns.Add("ntkns_kka_upld_tntsy_cd")
        pDataTable.Columns.Add("nktr_kensa_zyokyo_abms")
        pDataTable.Columns.Add("tkkk_keiki_id")
        pDataTable.Columns.Add("tkkk_kesbt_kbn")
        pDataTable.Columns.Add("tkkk_krhsk_cd")
        pDataTable.Columns.Add("tkkk_keiki_yr")
        pDataTable.Columns.Add("tkkk_tshsk_cd")
        pDataTable.Columns.Add("tkkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("tkkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("tkkk_taiko_kbn")
        pDataTable.Columns.Add("tkkk_knthk_cd")
        pDataTable.Columns.Add("tkkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("tkkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("tkkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("tkkk_yuko_kigen_ym")
        pDataTable.Columns.Add("tkkk_seizo_yy")
        pDataTable.Columns.Add("tkkk_zyrt")
        pDataTable.Columns.Add("tkkk_digsu")
        pDataTable.Columns.Add("tkkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("tkkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("tkkk_zytr_szsu")
        pDataTable.Columns.Add("tkkk_gytr_szsu")
        pDataTable.Columns.Add("no1_tkkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_tkkk_hnsik_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("tkkk_hnsik_gknti_no")
        pDataTable.Columns.Add("tkkk_mado_su")
        pDataTable.Columns.Add("tkkk_smkey1")
        pDataTable.Columns.Add("tkkk_smkey2")
        pDataTable.Columns.Add("tkkk_smkey_startdate")
        pDataTable.Columns.Add("tkkk_smkey_enddate")
        pDataTable.Columns.Add("tkkk_mac_address")
        pDataTable.Columns.Add("tkkk_tsn_id")
        pDataTable.Columns.Add("ttkk_keiki_id")
        pDataTable.Columns.Add("ttkk_kesbt_kbn")
        pDataTable.Columns.Add("ttkk_krhsk_cd")
        pDataTable.Columns.Add("ttkk_keiki_yr")
        pDataTable.Columns.Add("ttkk_tshsk_cd")
        pDataTable.Columns.Add("ttkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("ttkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("ttkk_taiko_kbn")
        pDataTable.Columns.Add("ttkk_knthk_cd")
        pDataTable.Columns.Add("ttkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("ttkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("ttkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("ttkk_yuko_kigen_ym")
        pDataTable.Columns.Add("ttkk_seizo_yy")
        pDataTable.Columns.Add("ttkk_zyrt")
        pDataTable.Columns.Add("ttkk_digsu")
        pDataTable.Columns.Add("ttkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("ttkk_zytr_szsu")
        pDataTable.Columns.Add("ttkk_gytr_szsu")
        pDataTable.Columns.Add("no1_ttkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_ttkk_hnsik_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("ttkk_hnsik_gknti_no")
        pDataTable.Columns.Add("ttkk_mado_su")
        pDataTable.Columns.Add("kdkk_keiki_id")
        pDataTable.Columns.Add("kdkk_kesbt_kbn")
        pDataTable.Columns.Add("kdkk_krhsk_cd")
        pDataTable.Columns.Add("kdkk_keiki_yr")
        pDataTable.Columns.Add("kdkk_tshsk_cd")
        pDataTable.Columns.Add("kdkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("kdkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("kdkk_taiko_kbn")
        pDataTable.Columns.Add("kdkk_knthk_cd")
        pDataTable.Columns.Add("kdkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("kdkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("kdkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("kdkk_yuko_kigen_ym")
        pDataTable.Columns.Add("kdkk_seizo_yy")
        pDataTable.Columns.Add("kdkk_zyrt")
        pDataTable.Columns.Add("kdkk_digsu")
        pDataTable.Columns.Add("kdkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("kdkk_zytr_szsu")
        pDataTable.Columns.Add("kdkk_gytr_szsu")
        pDataTable.Columns.Add("no1_kdkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_kdkk_hnsik_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("kdkk_hnsik_gknti_no")
        pDataTable.Columns.Add("kdkk_mado_su")
        pDataTable.Columns.Add("tdnstr_ts_dosa_kbn")
        pDataTable.Columns.Add("tdnstr_tdnt_su")
        pDataTable.Columns.Add("tdnstr_tdnstr_time")
        pDataTable.Columns.Add("hkgds_kbn")
        pDataTable.Columns.Add("hkgds_start_md")
        pDataTable.Columns.Add("hkgds_end_md")
        pDataTable.Columns.Add("no1_hkgds_tdn_hms")
        pDataTable.Columns.Add("no1_hkgds_sydan_hms")
        pDataTable.Columns.Add("no2_hkgds_tdn_hms")
        pDataTable.Columns.Add("no2_hkgds_sydan_hms")
        pDataTable.Columns.Add("no3_hkgds_tdn_hms")
        pDataTable.Columns.Add("no3_hkgds_sydan_hms")
        pDataTable.Columns.Add("no4_hkgds_tdn_hms")
        pDataTable.Columns.Add("no4_hkgds_sydan_hms")
        pDataTable.Columns.Add("no5_hkgds_tdn_hms")
        pDataTable.Columns.Add("no5_hkgds_sydan_hms")
        pDataTable.Columns.Add("no6_hkgds_tdn_hms")
        pDataTable.Columns.Add("no6_hkgds_sydan_hms")
        pDataTable.Columns.Add("no7_hkgds_tdn_hms")
        pDataTable.Columns.Add("no7_hkgds_sydan_hms")
        pDataTable.Columns.Add("no8_hkgds_tdn_hms")
        pDataTable.Columns.Add("no8_hkgds_sydan_hms")
        pDataTable.Columns.Add("no9_hkgds_tdn_hms")
        pDataTable.Columns.Add("no9_hkgds_sydan_hms")
        pDataTable.Columns.Add("no10_hkgds_tdn_hms")
        pDataTable.Columns.Add("no10_hkgds_sydan_hms")
        pDataTable.Columns.Add("kbttd_start_date")
        pDataTable.Columns.Add("kbttd_end_date")
        pDataTable.Columns.Add("hoka_kihi_kbn_cd")
        pDataTable.Columns.Add("hoka_gmn_flckr_kbn")
        pDataTable.Columns.Add("hoka_evnt_krk_kbn")
        pDataTable.Columns.Add("hoka_hoka_hyz_kbn")
        pDataTable.Columns.Add("hoka_hoka_hyz_value")
        pDataTable.Columns.Add("hksgk_hksgn_kbn")
        pDataTable.Columns.Add("hksgk_hka_dnr")
        pDataTable.Columns.Add("hksgk_attny_time")
        pDataTable.Columns.Add("hksgk_attny_kaisu")
        pDataTable.Columns.Add("hksgk_attny_kaisu_clr_time")
        pDataTable.Columns.Add("hksgr_hksgn_kbn")
        pDataTable.Columns.Add("hksgr_hka_dnr")
        pDataTable.Columns.Add("hksgr_attny_time")
        pDataTable.Columns.Add("hksgr_attny_kaisu")
        pDataTable.Columns.Add("hksgr_attny_kaisu_clr_time")
        pDataTable.Columns.Add("khkmk_keiki_kbn")
        pDataTable.Columns.Add("khkmk_mg_kbn")
        pDataTable.Columns.Add("khkmk_hnsik_kbn")
        pDataTable.Columns.Add("khkmk_wrms_kbn")
        pDataTable.Columns.Add("khkmk_mg_umu_flg")
        pDataTable.Columns.Add("khkmk_mg_yr")
        pDataTable.Columns.Add("tuika_kohi_umu_flg")
        pDataTable.Columns.Add("zippi_umu_flg")
        pDataTable.Columns.Add("no1_zippi_no")
        pDataTable.Columns.Add("no1_zippi_kbn")
        pDataTable.Columns.Add("no1_zippi_kngk")
        pDataTable.Columns.Add("no2_zippi_no")
        pDataTable.Columns.Add("no2_zippi_kbn")
        pDataTable.Columns.Add("no2_zippi_kngk")
        pDataTable.Columns.Add("no3_zippi_no")
        pDataTable.Columns.Add("no3_zippi_kbn")
        pDataTable.Columns.Add("no3_zippi_kngk")
        pDataTable.Columns.Add("tenp_file_mng_no")
        pDataTable.Columns.Add("ryssy_tenp_file_mng_no")
        pDataTable.Columns.Add("rrzk_naiyo")
        pDataTable.Columns.Add("prcmg_abms")
        pDataTable.Columns.Add("trke_sbt_abms")
        pDataTable.Columns.Add("brot_sett_flg_abms")
        pDataTable.Columns.Add("tkkk_kesbt_kbn_abms")
        pDataTable.Columns.Add("tkkk_krhsk_abms")
        pDataTable.Columns.Add("tkkk_tshsk_abms")
        pDataTable.Columns.Add("tkkk_knthk_abms")
        pDataTable.Columns.Add("tkkk_ts_kino_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_khk_kino_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_gaibu_output_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_taiko_kbn_abms")
        pDataTable.Columns.Add("ttkk_kesbt_kbn_abms")
        pDataTable.Columns.Add("ttkk_krhsk_abms")
        pDataTable.Columns.Add("ttkk_tshsk_abms")
        pDataTable.Columns.Add("ttkk_knthk_abms")
        pDataTable.Columns.Add("ttkk_ts_kino_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_khk_kino_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_gaibu_output_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_taiko_kbn_abms")
        pDataTable.Columns.Add("knti_sbt_ms")
        pDataTable.Columns.Add("surge_yksi_siyo_umu_flg_abms")
        pDataTable.Columns.Add("tidn_kbn_abms")
        pDataTable.Columns.Add("tnsb_reuse_kbn_abms")
        pDataTable.Columns.Add("khkmk_keiki_kbn_abms")
        pDataTable.Columns.Add("khkmk_mg_kbn_abms")
        pDataTable.Columns.Add("khkmk_hnsik_kbn_abms")
        pDataTable.Columns.Add("khkmk_wrms_kbn_abms")
        pDataTable.Columns.Add("khkmk_mg_umu_flg_abms")
        pDataTable.Columns.Add("no1_zippi_kbn_abms")
        pDataTable.Columns.Add("no2_zippi_kbn_abms")
        pDataTable.Columns.Add("no3_zippi_kbn_abms")
        pDataTable.Columns.Add("seko_kka_send_ymd")
        pDataTable.Columns.Add("seko_flg")
        pDataTable.Columns.Add("hktg_yohi_flg")
        pDataTable.Columns.Add("hktg_zumi_flg")
        pDataTable.Columns.Add("mdgt_ms")
        pDataTable.Columns.Add("kozitn_ms")
        pDataTable.Columns.Add("senro_ms")
        pDataTable.Columns.Add("tykei_mdgt_ms")
        pDataTable.Columns.Add("tykei_kozitn_ms")
        pDataTable.Columns.Add("yukoWhZytrTime")
        pDataTable.Columns.Add("yukoWhGytrTime")
        pDataTable.Columns.Add("nktrk_sizsk_kbn")
        pDataTable.Columns.Add("smj_send_zumi_flg")
        'Rev002.1 ADD Start 20241204 移動
        pDataTable.Columns.Add("trke_tis_kbn")
        'Rev002.1 ADD End
        'Rev040 2024/12/03 物理分割 次世代QR対応 ADD Start
        pDataTable.Columns.Add("tkkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("tkkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("tkkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("tkkk_tnsb_yr")
        pDataTable.Columns.Add("tkkk_tnsb_seizo_no")
        pDataTable.Columns.Add("tkkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("tkkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("tkkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("ttkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("ttkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("ttkk_tnsb_yr")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_no")
        pDataTable.Columns.Add("ttkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("ttkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("ttkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("kdkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("kdkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("kdkk_tnsb_yr")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_no")
        pDataTable.Columns.Add("kdkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("kdkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("kdkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("stzk_sdnsv_menu_cd")
        'Rev040 2024/12/03 物理分割 次世代QR対応 ADD End

    End Sub
#End Region


#Region "取替_取替票行程ファイル作成処理"

    ''' <summary>
    ''' 取替_取替票行程ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_PRC(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・取替_取替票行程を取得する。
            If Not dao.S003(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・取替_取替票行程検索でエラーが発生しました。"))
            End If

            '取替_取替票行程ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_TRKEHY_PRC(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("seko_tanto_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syori_ymd"), ""))
                wRow("seko_tanto_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syors_cd"), ""))
                wRow("seko_tanto_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syors_ms"), ""))
                wRow("seko_tanto_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_ka_cd"), ""))
                wRow("seko_tanto_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_ka_ms"), ""))
                wRow("seko_tanto_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kzkis_mdgt_cd"), ""))
                wRow("seko_tanto_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kzkis_mdgt_ms"), ""))
                wRow("seko_tanto_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kozitn_cd"), ""))
                wRow("seko_tanto_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kozitn_ms"), ""))
                wRow("seko_elder_tnknk_kknn_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnknk_kknn_flg"), ""))
                wRow("seko_elder_tnkn_kknn_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_kknn_ymd"), ""))
                wRow("seko_elder_tnkn_knsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_knsya_cd"), ""))
                wRow("seko_elder_tnkn_knsya_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_knsya_ms"), ""))
                wRow("seko_elder_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_ka_cd"), ""))
                wRow("seko_elder_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_ka_ms"), ""))
                wRow("seko_elder_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kzkis_mdgt_cd"), ""))
                wRow("seko_elder_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kzkis_mdgt_ms"), ""))
                wRow("seko_elder_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kozitn_cd"), ""))
                wRow("seko_elder_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kozitn_ms"), ""))
                wRow("nktrk_sizi_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syori_ymd"), ""))
                wRow("nktrk_sizi_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syors_cd"), ""))
                wRow("nktrk_sizi_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syors_ms"), ""))
                wRow("nktrk_sizi_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_ka_cd"), ""))
                wRow("nktrk_sizi_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_ka_ms"), ""))
                wRow("nktrk_sizi_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kzkis_mdgt_cd"), ""))
                wRow("nktrk_sizi_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kzkis_mdgt_ms"), ""))
                wRow("nktrk_sizi_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kozitn_cd"), ""))
                wRow("nktrk_sizi_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kozitn_ms"), ""))
                wRow("nktrk_tanto_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syori_ymd"), ""))
                wRow("nktrk_tanto_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syors_cd"), ""))
                wRow("nktrk_tanto_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syors_ms"), ""))
                wRow("nktrk_tanto_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ka_cd"), ""))
                wRow("nktrk_tanto_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ka_ms"), ""))
                wRow("nktrk_tanto_ces_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ces_cd"), ""))
                wRow("nktrk_tanto_ces_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ces_ms"), ""))
                wRow("nktrk_elder_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syori_ymd"), ""))
                wRow("nktrk_elder_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syors_cd"), ""))
                wRow("nktrk_elder_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syors_ms"), ""))
                wRow("nktrk_elder_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ka_cd"), ""))
                wRow("nktrk_elder_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ka_ms"), ""))
                wRow("nktrk_elder_ces_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ces_cd"), ""))
                wRow("nktrk_elder_ces_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ces_ms"), ""))
                wRow("nktrk_hnn_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_hnn_flg"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_PRC)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("取替票行程ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_TRKEHY_PRC(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("seko_tanto_syori_ymd")
        wDestDt.Columns.Add("seko_tanto_syors_cd")
        wDestDt.Columns.Add("seko_tanto_syors_ms")
        wDestDt.Columns.Add("seko_tanto_ka_cd")
        wDestDt.Columns.Add("seko_tanto_ka_ms")
        wDestDt.Columns.Add("seko_tanto_kzkis_mdgt_cd")
        wDestDt.Columns.Add("seko_tanto_kzkis_mdgt_ms")
        wDestDt.Columns.Add("seko_tanto_kozitn_cd")
        wDestDt.Columns.Add("seko_tanto_kozitn_ms")
        wDestDt.Columns.Add("seko_elder_tnknk_kknn_flg")
        wDestDt.Columns.Add("seko_elder_tnkn_kknn_ymd")
        wDestDt.Columns.Add("seko_elder_tnkn_knsya_cd")
        wDestDt.Columns.Add("seko_elder_tnkn_knsya_ms")
        wDestDt.Columns.Add("seko_elder_ka_cd")
        wDestDt.Columns.Add("seko_elder_ka_ms")
        wDestDt.Columns.Add("seko_elder_kzkis_mdgt_cd")
        wDestDt.Columns.Add("seko_elder_kzkis_mdgt_ms")
        wDestDt.Columns.Add("seko_elder_kozitn_cd")
        wDestDt.Columns.Add("seko_elder_kozitn_ms")
        wDestDt.Columns.Add("nktrk_sizi_syori_ymd")
        wDestDt.Columns.Add("nktrk_sizi_syors_cd")
        wDestDt.Columns.Add("nktrk_sizi_syors_ms")
        wDestDt.Columns.Add("nktrk_sizi_ka_cd")
        wDestDt.Columns.Add("nktrk_sizi_ka_ms")
        wDestDt.Columns.Add("nktrk_sizi_kzkis_mdgt_cd")
        wDestDt.Columns.Add("nktrk_sizi_kzkis_mdgt_ms")
        wDestDt.Columns.Add("nktrk_sizi_kozitn_cd")
        wDestDt.Columns.Add("nktrk_sizi_kozitn_ms")
        wDestDt.Columns.Add("nktrk_tanto_syori_ymd")
        wDestDt.Columns.Add("nktrk_tanto_syors_cd")
        wDestDt.Columns.Add("nktrk_tanto_syors_ms")
        wDestDt.Columns.Add("nktrk_tanto_ka_cd")
        wDestDt.Columns.Add("nktrk_tanto_ka_ms")
        wDestDt.Columns.Add("nktrk_tanto_ces_cd")
        wDestDt.Columns.Add("nktrk_tanto_ces_ms")
        wDestDt.Columns.Add("nktrk_elder_syori_ymd")
        wDestDt.Columns.Add("nktrk_elder_syors_cd")
        wDestDt.Columns.Add("nktrk_elder_syors_ms")
        wDestDt.Columns.Add("nktrk_elder_ka_cd")
        wDestDt.Columns.Add("nktrk_elder_ka_ms")
        wDestDt.Columns.Add("nktrk_elder_ces_cd")
        wDestDt.Columns.Add("nktrk_elder_ces_ms")
        wDestDt.Columns.Add("nktrk_hnn_flg")

    End Sub
#End Region


#Region "特記事項ファイル作成処理"

    ''' <summary>
    ''' 特記事項ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateKEIKI_SPOT_TKZK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・特記事項を取得する。
            If Not dao.S004(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・特記事項検索でエラーが発生しました。"))
            End If

            '特記事項ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblKEIKI_SPOT_TKZK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("kyokyu_spot_tokti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kyokyu_spot_tokti_no"), ""))
                wRow("tkzk_renno") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_renno"), ""))
                wRow("tkzk_bnrui_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_bnrui_cd"), ""))
                wRow("tkzk_kizi_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_kizi_naiyo"), ""))
                wRow("itaks_kaizi_kizi_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("itaks_kaizi_kizi_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_KEIKI_SPOT_TKZK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("特記事項ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblKEIKI_SPOT_TKZK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("kyokyu_spot_tokti_no")
        wDestDt.Columns.Add("tkzk_renno")
        wDestDt.Columns.Add("tkzk_bnrui_cd")
        wDestDt.Columns.Add("tkzk_kizi_naiyo")
        wDestDt.Columns.Add("itaks_kaizi_kizi_naiyo")
    End Sub
#End Region


    'Rev002-Start
#Region "特記事項（高圧他）ファイル作成処理"

    ''' <summary>
    ''' 特記事項ファイル作成処理（高圧用）
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateKEIKI_SPOT_TKZK_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・特記事項（高圧他）を取得する。
            If Not dao.S018(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・特記事項（高圧他）検索でエラーが発生しました。"))
            End If

            '特記事項ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblKEIKI_SPOT_TKZK_KAHK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("kyokyu_spot_tokti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kyokyu_spot_tokti_no"), ""))
                wRow("tkzk_renno") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_renno"), ""))
                wRow("tkzk_bnrui_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_bnrui_cd"), ""))
                wRow("tkzk_kizi_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkzk_kizi_naiyo"), ""))
                wRow("itaks_kaizi_kizi_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("itaks_kaizi_kizi_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，高圧用ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_KEIKI_SPOT_TKZK_KAHK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("特記事項（高圧他）ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblKEIKI_SPOT_TKZK_KAHK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("kyokyu_spot_tokti_no")
        wDestDt.Columns.Add("tkzk_renno")
        wDestDt.Columns.Add("tkzk_bnrui_cd")
        wDestDt.Columns.Add("tkzk_kizi_naiyo")
        wDestDt.Columns.Add("itaks_kaizi_kizi_naiyo")
    End Sub
#End Region
    'Rev002-End

#Region "追加工費ファイル作成処理"

    ''' <summary>
    ''' 追加工費ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_TUIKAKH(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・追加工費を取得する。
            If Not dao.S005(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・追加工費検索でエラーが発生しました。"))
            End If

            '追加工費ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_TRKEHY_TUIKAKH(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("tuika_kohi_row_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_row_no"), ""))
                wRow("tuika_kohi_kozi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_kozi_kbn"), ""))
                wRow("tuika_kohi_kozi_kbn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_kozi_kbn_ms"), ""))
                wRow("tuika_kohi_trtk_suryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_trtk_suryo"), ""))
                wRow("tuika_kohi_tekyo_suryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_tekyo_suryo"), ""))
                wRow("trtk_kozih_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_kozih_kngk"), ""))
                wRow("tekyo_kozih_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_kozih_kngk"), ""))
                wRow("tuika_kohi_riyu_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_riyu_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_TUIKAKH)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("追加工費ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_TRKEHY_TUIKAKH(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("tuika_kohi_row_no")
        wDestDt.Columns.Add("tuika_kohi_kozi_kbn")
        wDestDt.Columns.Add("tuika_kohi_kozi_kbn_ms")
        wDestDt.Columns.Add("tuika_kohi_trtk_suryo")
        wDestDt.Columns.Add("tuika_kohi_tekyo_suryo")
        wDestDt.Columns.Add("trtk_kozih_kngk")
        wDestDt.Columns.Add("tekyo_kozih_kngk")
        wDestDt.Columns.Add("tuika_kohi_riyu_naiyo")
    End Sub
#End Region




    'Rev002-Start
#Region "追加工費（高圧他）ファイル作成処理"

    ''' <summary>
    ''' 追加工費（高圧他）ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_TUIKAKH_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・追加工費（高圧他）を取得する。
            If Not dao.S019(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・追加工費（高圧他）検索でエラーが発生しました。"))
            End If

            '追加工費（高圧他）ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_TRKEHY_TUIKAKH_KAHK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("tuika_kohi_row_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_row_no"), ""))
                wRow("tuika_kohi_kozi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_kozi_kbn"), ""))
                wRow("tuika_kohi_kozi_kbn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_kozi_kbn_ms"), ""))
                wRow("tuika_kohi_trtk_suryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_trtk_suryo"), ""))
                wRow("tuika_kohi_tekyo_suryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_tekyo_suryo"), ""))
                wRow("trtk_kozih_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_kozih_kngk"), ""))
                wRow("tekyo_kozih_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_kozih_kngk"), ""))
                wRow("tuika_kohi_riyu_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_riyu_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，高圧用ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_TUIKAKH_KAHK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("追加工費ファイル（高圧）出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_TRKEHY_TUIKAKH_KAHK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("tuika_kohi_row_no")
        wDestDt.Columns.Add("tuika_kohi_kozi_kbn")
        wDestDt.Columns.Add("tuika_kohi_kozi_kbn_ms")
        wDestDt.Columns.Add("tuika_kohi_trtk_suryo")
        wDestDt.Columns.Add("tuika_kohi_tekyo_suryo")
        wDestDt.Columns.Add("trtk_kozih_kngk")
        wDestDt.Columns.Add("tekyo_kozih_kngk")
        wDestDt.Columns.Add("tuika_kohi_riyu_naiyo")
    End Sub
#End Region
    'Rev002-End


#Region "自主点検ファイル作成処理"

    ''' <summary>
    ''' 自主点検ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_CHECK_SHEET(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・自主点検を取得する。
            If Not dao.S006(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・自主点検検索でエラーが発生しました。"))
            End If

            '自主点検ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_CHECK_SHEET(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("takcd_ttt_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_skosya_ryohi_kbn"), ""))
                wRow("takcd_ttt_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_ntktt_ryohi_kbn"), ""))
                wRow("takcd_ttt_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_kstts_ryohi_kbn"), ""))
                wRow("takcd_ttt_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_tio_naiyo"), ""))
                wRow("takcd_ttt_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_tio_cmp_ymd"), ""))
                wRow("takcd_szsu_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_skosya_ryohi_kbn"), ""))
                wRow("takcd_szsu_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_ntktt_ryohi_kbn"), ""))
                wRow("takcd_szsu_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_kstts_ryohi_kbn"), ""))
                wRow("takcd_szsu_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_tio_naiyo"), ""))
                wRow("takcd_szsu_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_tio_cmp_ymd"), ""))
                wRow("takck_ssda_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_skosya_ryohi_kbn"), ""))
                wRow("takck_ssda_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_ntktt_ryohi_kbn"), ""))
                wRow("takck_ssda_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_kstts_ryohi_kbn"), ""))
                wRow("takck_ssda_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_tio_naiyo"), ""))
                wRow("takck_ssda_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_tio_cmp_ymd"), ""))
                wRow("takck_zyrt_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_skosya_ryohi_kbn"), ""))
                wRow("takck_zyrt_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_ntktt_ryohi_kbn"), ""))
                wRow("takck_zyrt_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_kstts_ryohi_kbn"), ""))
                wRow("takck_zyrt_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_tio_naiyo"), ""))
                wRow("takck_zyrt_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_tio_cmp_ymd"), ""))
                wRow("takck_yr_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_skosya_ryohi_kbn"), ""))
                wRow("takck_yr_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_ntktt_ryohi_kbn"), ""))
                wRow("takck_yr_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_kstts_ryohi_kbn"), ""))
                wRow("takck_yr_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_tio_naiyo"), ""))
                wRow("takck_yr_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_tio_cmp_ymd"), ""))
                wRow("takck_sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_skosya_ryohi_kbn"), ""))
                wRow("takck_sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_ntktt_ryohi_kbn"), ""))
                wRow("takck_sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_kstts_ryohi_kbn"), ""))
                wRow("takck_sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_tio_naiyo"), ""))
                wRow("takck_sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_tio_cmp_ymd"), ""))
                wRow("takck_kkano_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_skosya_ryohi_kbn"), ""))
                wRow("takck_kkano_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_ntktt_ryohi_kbn"), ""))
                wRow("takck_kkano_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_kstts_ryohi_kbn"), ""))
                wRow("takck_kkano_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_tio_naiyo"), ""))
                wRow("takck_kkano_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_tio_cmp_ymd"), ""))
                wRow("takck_hksno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_skosya_ryohi_kbn"), ""))
                wRow("takck_hksno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_ntktt_ryohi_kbn"), ""))
                wRow("takck_hksno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_kstts_ryohi_kbn"), ""))
                wRow("takck_hksno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_tio_naiyo"), ""))
                wRow("takck_hksno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_tio_cmp_ymd"), ""))
                wRow("takck_hkano_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_skosya_ryohi_kbn"), ""))
                wRow("takck_hkano_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_ntktt_ryohi_kbn"), ""))
                wRow("takck_hkano_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_kstts_ryohi_kbn"), ""))
                wRow("takck_hkano_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_tio_naiyo"), ""))
                wRow("takck_hkano_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_tio_cmp_ymd"), ""))
                wRow("takckk_krkkd_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_skosya_ryohi_kbn"), ""))
                wRow("takckk_krkkd_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_ntktt_ryohi_kbn"), ""))
                wRow("takckk_krkkd_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_kstts_ryohi_kbn"), ""))
                wRow("takckk_krkkd_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_tio_naiyo"), ""))
                wRow("takckk_krkkd_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_tio_cmp_ymd"), ""))
                wRow("takckk_krkb_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_skosya_ryohi_kbn"), ""))
                wRow("takckk_krkb_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_ntktt_ryohi_kbn"), ""))
                wRow("takckk_krkb_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_kstts_ryohi_kbn"), ""))
                wRow("takckk_krkb_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_tio_naiyo"), ""))
                wRow("takckk_krkb_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_tio_cmp_ymd"), ""))
                wRow("takcks_kkzen_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkzen_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkzen_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkzen_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_tio_naiyo"), ""))
                wRow("takcks_kkzen_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_tio_cmp_ymd"), ""))
                wRow("takcks_kksl_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_skosya_ryohi_kbn"), ""))
                wRow("takcks_kksl_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kksl_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_kstts_ryohi_kbn"), ""))
                wRow("takcks_kksl_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_tio_naiyo"), ""))
                wRow("takcks_kksl_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_tio_cmp_ymd"), ""))
                wRow("takcks_kkhks_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkhks_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkhks_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkhks_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_tio_naiyo"), ""))
                wRow("takcks_kkhks_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_tio_cmp_ymd"), ""))
                wRow("takcks_kkgsz_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_tio_naiyo"), ""))
                wRow("takcks_kkgsz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_tio_cmp_ymd"), ""))
                wRow("takcks_kkaki_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkaki_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkaki_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkaki_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_tio_naiyo"), ""))
                wRow("takcks_kkaki_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_tio_cmp_ymd"), ""))
                wRow("takcks_kkyrs_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_tio_naiyo"), ""))
                wRow("takcks_kkyrs_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_tio_cmp_ymd"), ""))
                wRow("takcks_kkrst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkrst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkrst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkrst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_tio_naiyo"), ""))
                wRow("takcks_kkrst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_tio_cmp_ymd"), ""))
                wRow("takcks_kkbst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkbst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkbst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkbst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_tio_naiyo"), ""))
                wRow("takcks_kkbst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_tio_cmp_ymd"), ""))
                wRow("takcks_kikib_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_skosya_ryohi_kbn"), ""))
                wRow("takcks_kikib_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kikib_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_kstts_ryohi_kbn"), ""))
                wRow("takcks_kikib_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_tio_naiyo"), ""))
                wRow("takcks_kikib_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_tio_cmp_ymd"), ""))
                wRow("takcks_kkssu_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkssu_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkssu_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkssu_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_tio_naiyo"), ""))
                wRow("takcks_kkssu_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_tio_cmp_ymd"), ""))
                wRow("takcks_hrkkl_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_skosya_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_kstts_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_tio_naiyo"), ""))
                wRow("takcks_hrkkl_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_tio_cmp_ymd"), ""))
                wRow("takcks_hp123_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_skosya_ryohi_kbn"), ""))
                wRow("takcks_hp123_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hp123_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_kstts_ryohi_kbn"), ""))
                wRow("takcks_hp123_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_tio_naiyo"), ""))
                wRow("takcks_hp123_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_tio_cmp_ymd"), ""))
                wRow("takcks_hkaki_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_skosya_ryohi_kbn"), ""))
                wRow("takcks_hkaki_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hkaki_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_kstts_ryohi_kbn"), ""))
                wRow("takcks_hkaki_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_tio_naiyo"), ""))
                wRow("takcks_hkaki_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_tio_cmp_ymd"), ""))
                wRow("takckrk_hnrk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_skosya_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_ntktt_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_kstts_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_tio_naiyo"), ""))
                wRow("takckrk_hnrk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_tio_cmp_ymd"), ""))
                wRow("takckrr_sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_skosya_ryohi_kbn"), ""))
                wRow("takckrr_sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_kstts_ryohi_kbn"), ""))
                wRow("takckrr_sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_tio_naiyo"), ""))
                wRow("takckrr_sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_tio_cmp_ymd"), ""))
                wRow("takcc_led_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_skosya_ryohi_kbn"), ""))
                wRow("takcc_led_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_ntktt_ryohi_kbn"), ""))
                wRow("takcc_led_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_kstts_ryohi_kbn"), ""))
                wRow("takcc_led_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_tio_naiyo"), ""))
                wRow("takcc_led_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_tio_cmp_ymd"), ""))
                wRow("takcc_souck_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_skosya_ryohi_kbn"), ""))
                wRow("takcc_souck_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_ntktt_ryohi_kbn"), ""))
                wRow("takcc_souck_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_kstts_ryohi_kbn"), ""))
                wRow("takcc_souck_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_tio_naiyo"), ""))
                wRow("takcc_souck_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_tio_cmp_ymd"), ""))
                wRow("takct_szno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_skosya_ryohi_kbn"), ""))
                wRow("takct_szno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_ntktt_ryohi_kbn"), ""))
                wRow("takct_szno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_kstts_ryohi_kbn"), ""))
                wRow("takct_szno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_tio_naiyo"), ""))
                wRow("takct_szno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_tio_cmp_ymd"), ""))
                wRow("takct_aino_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_skosya_ryohi_kbn"), ""))
                wRow("takct_aino_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_ntktt_ryohi_kbn"), ""))
                wRow("takct_aino_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_kstts_ryohi_kbn"), ""))
                wRow("takct_aino_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_tio_naiyo"), ""))
                wRow("takct_aino_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_tio_cmp_ymd"), ""))
                wRow("takcs_photo_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_skosya_ryohi_kbn"), ""))
                wRow("takcs_photo_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_ntktt_ryohi_kbn"), ""))
                wRow("takcs_photo_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_kstts_ryohi_kbn"), ""))
                wRow("takcs_photo_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_tio_naiyo"), ""))
                wRow("takcs_photo_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_tio_cmp_ymd"), ""))
                wRow("takcs_tktrk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_skosya_ryohi_kbn"), ""))
                wRow("takcs_tktrk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_ntktt_ryohi_kbn"), ""))
                wRow("takcs_tktrk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_kstts_ryohi_kbn"), ""))
                wRow("takcs_tktrk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_tio_naiyo"), ""))
                wRow("takcs_tktrk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_tio_cmp_ymd"), ""))
                wRow("takcz_ntktt_seigo_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_ntktt_seigo_kknn_kbn"), ""))
                wRow("takcz_kstts_seigo_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_kstts_seigo_kknn_kbn"), ""))
                wRow("takcz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_tio_naiyo"), ""))
                wRow("takcz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_tio_cmp_ymd"), ""))
                wRow("takca_ntktt_itti_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_ntktt_itti_kknn_kbn"), ""))
                wRow("takca_kstts_itti_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_kstts_itti_kknn_kbn"), ""))
                wRow("takca_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_tio_naiyo"), ""))
                wRow("takca_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_tio_cmp_ymd"), ""))
                wRow("kizi_skosya_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_skosya_naiyo"), ""))
                wRow("kizi_ntktt_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_ntktt_naiyo"), ""))
                wRow("kizi_kstts_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_kstts_naiyo"), ""))
                wRow("hrktkk_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_1s_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p1_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p3_kstckk_flg"), ""))
                wRow("hrktkk_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_3s_kstckk_flg"), ""))
                wRow("hrktkk_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_3l_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p2_kstckk_flg"), ""))
                wRow("hrktkk_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_1l_kstckk_flg"), ""))
                wRow("p1_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_1s_kstckk_flg"), ""))
                wRow("p1_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_p1_kstckk_flg"), ""))
                wRow("p1_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_1l_kstckk_flg"), ""))
                wRow("p2_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_skosya_p2_kstckk_flg"), ""))
                wRow("p3_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_3s_kstckk_flg"), ""))
                wRow("p3_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_p3_kstckk_flg"), ""))
                wRow("p3_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_3l_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_1s_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p1_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p3_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_3s_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_3l_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p2_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_1l_kstckk_flg"), ""))
                wRow("p1_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_1s_kstckk_flg"), ""))
                wRow("p1_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_p1_kstckk_flg"), ""))
                wRow("p1_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_1l_kstckk_flg"), ""))
                wRow("p2_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_ntktt_p2_kstckk_flg"), ""))
                wRow("p3_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_3s_kstckk_flg"), ""))
                wRow("p3_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_p3_kstckk_flg"), ""))
                wRow("p3_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_3l_kstckk_flg"), ""))
                wRow("hrktkk_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_1s_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p1_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p3_kstckk_flg"), ""))
                wRow("hrktkk_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_3s_kstckk_flg"), ""))
                wRow("hrktkk_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_3l_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p2_kstckk_flg"), ""))
                wRow("hrktkk_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_1l_kstckk_flg"), ""))
                wRow("p1_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_1s_kstckk_flg"), ""))
                wRow("p1_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_p1_kstckk_flg"), ""))
                wRow("p1_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_1l_kstckk_flg"), ""))
                wRow("p2_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_kstts_p2_kstckk_flg"), ""))
                wRow("p3_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_3s_kstckk_flg"), ""))
                wRow("p3_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_p3_kstckk_flg"), ""))
                wRow("p3_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_3l_kstckk_flg"), ""))
                'Rev040 2024/11/20 物理分割 画面変更対応 ADD Start
                wRow("takckrr_smhyz_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_skosya_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_kstts_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_tio_naiyo"), ""))
                wRow("takckrr_smhyz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_tio_cmp_ymd"), ""))
                wRow("takckrr_d2sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_skosya_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_kstts_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_tio_naiyo"), ""))
                wRow("takckrr_d2sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_tio_cmp_ymd"), ""))
                wRow("takckrr_khk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_skosya_ryohi_kbn"), ""))
                wRow("takckrr_khk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_khk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_kstts_ryohi_kbn"), ""))
                wRow("takckrr_khk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_tio_naiyo"), ""))
                wRow("takckrr_khk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_tio_cmp_ymd"), ""))
                wRow("takckk_okgkb_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_skosya_ryohi_kbn"), ""))
                wRow("takckk_okgkb_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_ntktt_ryohi_kbn"), ""))
                wRow("takckk_okgkb_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_kstts_ryohi_kbn"), ""))
                wRow("takckk_okgkb_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_tio_naiyo"), ""))
                wRow("takckk_okgkb_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_tio_cmp_ymd"), ""))
                wRow("takckk_dsipcv_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_skosya_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_ntktt_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_kstts_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_tio_naiyo"), ""))
                wRow("takckk_dsipcv_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_tio_cmp_ymd"), ""))
                'Rev040 2024/11/20 物理分割 画面変更対応 ADD End

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_CHECK_SHEET)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("自主点検ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_CHECK_SHEET(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("takcd_ttt_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_tio_naiyo")
        wDestDt.Columns.Add("takcd_ttt_tio_cmp_ymd")
        wDestDt.Columns.Add("takcd_szsu_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_tio_naiyo")
        wDestDt.Columns.Add("takcd_szsu_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_ssda_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_tio_naiyo")
        wDestDt.Columns.Add("takck_ssda_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_zyrt_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_tio_naiyo")
        wDestDt.Columns.Add("takck_zyrt_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_yr_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_tio_naiyo")
        wDestDt.Columns.Add("takck_yr_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_tio_naiyo")
        wDestDt.Columns.Add("takck_sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_kkano_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_tio_naiyo")
        wDestDt.Columns.Add("takck_kkano_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_hksno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_tio_naiyo")
        wDestDt.Columns.Add("takck_hksno_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_hkano_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_tio_naiyo")
        wDestDt.Columns.Add("takck_hkano_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_krkkd_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_tio_naiyo")
        wDestDt.Columns.Add("takckk_krkkd_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_krkb_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_tio_naiyo")
        wDestDt.Columns.Add("takckk_krkb_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkzen_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkzen_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kksl_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_tio_naiyo")
        wDestDt.Columns.Add("takcks_kksl_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkhks_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkhks_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkgsz_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkgsz_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkaki_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkaki_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkyrs_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkyrs_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkrst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkrst_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkbst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkbst_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kikib_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_tio_naiyo")
        wDestDt.Columns.Add("takcks_kikib_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkssu_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkssu_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hrkkl_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_tio_naiyo")
        wDestDt.Columns.Add("takcks_hrkkl_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hp123_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_tio_naiyo")
        wDestDt.Columns.Add("takcks_hp123_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hkaki_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_tio_naiyo")
        wDestDt.Columns.Add("takcks_hkaki_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrk_hnrk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_tio_naiyo")
        wDestDt.Columns.Add("takckrk_hnrk_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_tio_naiyo")
        wDestDt.Columns.Add("takckrr_sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takcc_led_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_tio_naiyo")
        wDestDt.Columns.Add("takcc_led_tio_cmp_ymd")
        wDestDt.Columns.Add("takcc_souck_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_tio_naiyo")
        wDestDt.Columns.Add("takcc_souck_tio_cmp_ymd")
        wDestDt.Columns.Add("takct_szno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_tio_naiyo")
        wDestDt.Columns.Add("takct_szno_tio_cmp_ymd")
        wDestDt.Columns.Add("takct_aino_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_tio_naiyo")
        wDestDt.Columns.Add("takct_aino_tio_cmp_ymd")
        wDestDt.Columns.Add("takcs_photo_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_tio_naiyo")
        wDestDt.Columns.Add("takcs_photo_tio_cmp_ymd")
        wDestDt.Columns.Add("takcs_tktrk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_tio_naiyo")
        wDestDt.Columns.Add("takcs_tktrk_tio_cmp_ymd")
        wDestDt.Columns.Add("takcz_ntktt_seigo_kknn_kbn")
        wDestDt.Columns.Add("takcz_kstts_seigo_kknn_kbn")
        wDestDt.Columns.Add("takcz_tio_naiyo")
        wDestDt.Columns.Add("takcz_tio_cmp_ymd")
        wDestDt.Columns.Add("takca_ntktt_itti_kknn_kbn")
        wDestDt.Columns.Add("takca_kstts_itti_kknn_kbn")
        wDestDt.Columns.Add("takca_tio_naiyo")
        wDestDt.Columns.Add("takca_tio_cmp_ymd")
        wDestDt.Columns.Add("kizi_skosya_naiyo")
        wDestDt.Columns.Add("kizi_ntktt_naiyo")
        wDestDt.Columns.Add("kizi_kstts_naiyo")
        wDestDt.Columns.Add("hrktkk_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_3l_kstckk_flg")
        'Rev040 2024/11/20 物理分割 画面表示変更対応 ADD Start
        wDestDt.Columns.Add("takckrr_smhyz_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_tio_naiyo")
        wDestDt.Columns.Add("takckrr_smhyz_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_d2sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_tio_naiyo")
        wDestDt.Columns.Add("takckrr_d2sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_khk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_tio_naiyo")
        wDestDt.Columns.Add("takckrr_khk_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_okgkb_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_tio_naiyo")
        wDestDt.Columns.Add("takckk_okgkb_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_dsipcv_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_tio_naiyo")
        wDestDt.Columns.Add("takckk_dsipcv_tio_cmp_ymd")
        'Rev040 2024/11/20 物理分割 画面表示変更対応 ADD End
    End Sub
#End Region


#Region "添付書類詳細ファイル作成処理"

    ''' <summary>
    ''' 添付書類詳細ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateCOM_TNPDOC_DTL(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim sDataTable As System.Data.DataTable = Nothing   '検索データテーブル
        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・添付書類詳細を取得する。
            If Not dao.S007(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・添付書類詳細検索でエラーが発生しました。"))
            End If

            'PostgreSQLは、selectの多重ができないため、結果をDataTableに設定し、dbrを開放する
            sDataTable = New DataTable
            'カラム名設定
            For i = 0 To dbr.DataStruct.FieldCount - 1
                sDataTable.Columns.Add(dbr.DataStruct.GetName(i))
            Next
            'DataTableにデータ保存
            While (dbr.DataStruct.Read)
                wRow = sDataTable.NewRow
                For i = 0 To dbr.DataStruct.FieldCount - 1
                    wRow(dbr.DataStruct.GetName(i)) = dbr.DataStruct.Item(i).ToString
                Next
                sDataTable.Rows.Add(wRow)
            End While
            dbr.Close()


            '添付書類詳細ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblCOM_TNPDOC_DTL(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            For Each dRow As DataRow In sDataTable.Rows
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dRow("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dRow("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dRow("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dRow("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dRow("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dRow("trkhy_no"), ""))
                wRow("file_renno") = CStr(Me.myDB.ConvDbNull(dRow("file_renno"), ""))
                wRow("no1_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no1_tnpdoc_hosok_naiyo"), ""))

                'B-ST-0325起因　検査の場合、実費・領収書の調整が必要（施工の場合は、タブレットで登録なので考慮不要）Start
                '写真区分：実費(04)の場合、写真種類が空白なのでタブレット内構成に合わせて領収書:06に補正する
                Dim wkstr As String = CStr(Me.myDB.ConvDbNull(dRow("no1_tnpdoc_hosok_naiyo"), ""))
                If wkstr.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Jippi) Then
                    wRow("no2_tnpdoc_hosok_naiyo") = MjK1.Cmn.Const.KeikiPhotoSbtCd.RyoushuuSho
                Else
                    wRow("no2_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no2_tnpdoc_hosok_naiyo"), ""))
                End If
                'B-ST-0325起因　End
                wRow("no3_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no3_tnpdoc_hosok_naiyo"), ""))
                wRow("no4_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no4_tnpdoc_hosok_naiyo"), ""))
                wRow("no5_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no5_tnpdoc_hosok_naiyo"), ""))
                wRow("no6_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no6_tnpdoc_hosok_naiyo"), ""))
                wRow("no7_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no7_tnpdoc_hosok_naiyo"), ""))
                wRow("no8_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no8_tnpdoc_hosok_naiyo"), ""))
                wRow("no9_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no9_tnpdoc_hosok_naiyo"), ""))
                wRow("no10_tnpdoc_hosok_naiyo") = CStr(Me.myDB.ConvDbNull(dRow("no10_tnpdoc_hosok_naiyo"), ""))
                wRow("photo_sbt_ms") = CStr(Me.myDB.ConvDbNull(dRow("photo_sbt_ms"), ""))
                wRow("file_ms") = CStr(Me.myDB.ConvDbNull(dRow("file_ms"), ""))

                '出力データテーブルに１レコード追加する。
                wDataTable.Rows.Add(wRow)

                '添付書類ファイルをコピーする
                copyFileFromNAS(pMtdsMngNoFolder, pKeikiDwldSbtCd, dRow)
            Next

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_COM_TNPDOC_DTL)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("添付書類ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblCOM_TNPDOC_DTL(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("file_renno")
        wDestDt.Columns.Add("no1_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no2_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no3_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no4_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no5_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no6_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no7_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no8_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no9_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("no10_tnpdoc_hosok_naiyo")
        wDestDt.Columns.Add("photo_sbt_ms")
        wDestDt.Columns.Add("file_ms")
    End Sub

    Sub copyFileFromNAS(ByVal pMtdsMngNoFolder As String, ByVal pKeikiDwldSbtCd As String, ByVal dRow As DataRow)

        Dim wdig4_zgsyo_cd As String = ""
        Dim wtrkhy_hakko_nendo As String = ""
        Dim wtrkhy_kbn As String = ""
        Dim wtrkhy_no As String = ""
        Dim wfile_renno As String = ""
        Dim wno1_tnpdoc_hosok_naiyo As String = ""
        Dim wno2_tnpdoc_hosok_naiyo As String = ""
        Dim wno3_tnpdoc_hosok_naiyo As String = ""
        Dim wno4_tnpdoc_hosok_naiyo As String = ""
        Dim wno5_tnpdoc_hosok_naiyo As String = ""
        Dim wno6_tnpdoc_hosok_naiyo As String = ""
        Dim wno7_tnpdoc_hosok_naiyo As String = ""
        Dim wno8_tnpdoc_hosok_naiyo As String = ""
        Dim wno9_tnpdoc_hosok_naiyo As String = ""
        Dim wno10_tnpdoc_hosok_naiyo As String = ""
        Dim wphoto_sbt_ms As String = ""
        Dim wfile_ms As String = ""
        Dim wnasfilepath As String = ""
        Dim trkhyno_key As String = ""
        Dim tenp_file_mng_no As Decimal = 0
        Dim ryssy_tenp_file_mng_no As Decimal = 0
        Dim mng_no As Decimal = 0
        Dim from_path As String = ""
        Dim to_path As String = ""
        Dim sbtFolderMs As String = ""

        wdig4_zgsyo_cd = CStr(Me.myDB.ConvDbNull(dRow("dig4_zgsyo_cd"), ""))
        wtrkhy_hakko_nendo = CStr(Me.myDB.ConvDbNull(dRow("trkhy_hakko_nendo"), ""))
        wtrkhy_kbn = CStr(Me.myDB.ConvDbNull(dRow("trkhy_kbn"), ""))
        wtrkhy_no = CStr(Me.myDB.ConvDbNull(dRow("trkhy_no"), ""))
        wfile_renno = CStr(Me.myDB.ConvDbNull(dRow("file_renno"), ""))
        wno1_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no1_tnpdoc_hosok_naiyo"), ""))
        wno2_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no2_tnpdoc_hosok_naiyo"), ""))
        wno3_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no3_tnpdoc_hosok_naiyo"), ""))
        wno4_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no4_tnpdoc_hosok_naiyo"), ""))
        wno5_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no5_tnpdoc_hosok_naiyo"), ""))
        wno6_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no6_tnpdoc_hosok_naiyo"), ""))
        wno7_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no7_tnpdoc_hosok_naiyo"), ""))
        wno8_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no8_tnpdoc_hosok_naiyo"), ""))
        wno9_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no9_tnpdoc_hosok_naiyo"), ""))
        wno10_tnpdoc_hosok_naiyo = CStr(Me.myDB.ConvDbNull(dRow("no10_tnpdoc_hosok_naiyo"), ""))
        wphoto_sbt_ms = CStr(Me.myDB.ConvDbNull(dRow("photo_sbt_ms"), ""))
        wfile_ms = CStr(Me.myDB.ConvDbNull(dRow("file_ms"), ""))

        '取替票番号キーを取得する
        trkhyno_key = MjK1.Cmn.Func.GetTRKHYNO(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no)


        '添付書類管理番号取得
        tenp_file_mng_no = MyCDec(Me.myDB.ConvDbNull(dRow("tenp_file_mng_no"), "").ToString)
        ryssy_tenp_file_mng_no = MyCDec(Me.myDB.ConvDbNull(dRow("ryssy_tenp_file_mng_no"), "").ToString)

        Select Case wno1_tnpdoc_hosok_naiyo
            Case MjK1.Cmn.Const.KeikiPhotoKbn.Tekkyo, MjK1.Cmn.Const.KeikiPhotoKbn.Torituke, MjK1.Cmn.Const.KeikiPhotoKbn.Kuradasi
                mng_no = tenp_file_mng_no
            Case MjK1.Cmn.Const.KeikiPhotoKbn.Jippi
                mng_no = ryssy_tenp_file_mng_no
                '実費の写真種類はなしだが、タブレットで使用するので割り当てておく。
                wno2_tnpdoc_hosok_naiyo = MjK1.Cmn.Const.KeikiPhotoSbtCd.RyoushuuSho
            Case Else
                mng_no = 0
        End Select

        ''フォルダ名称取得・・・SEKO/KENSのどちらか
        'Select Case pKeikiDwldSbtCd
        '    Case MjK1.Cmn.Const.KeikiDwldSbtCd.Seko
        '        sbtFolderMs = MjK1.Cmn.Const.KeikiFolder.Seko
        '    Case MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns
        '        sbtFolderMs = MjK1.Cmn.Const.KeikiFolder.NktrKns
        '    Case Else
        '        sbtFolderMs = ""
        'End Select

        '添付書類管理番号がある場合のみ添付書類ファイルをコピーする
        If mng_no <> 0 Then
            'NAS上のファイルパス
            from_path = System.IO.Path.Combine(HdSysBasePg.TnpDoc.GetTnpDocPath(myDB, myPath, mng_no), wfile_ms)

            'コピー先
            to_path = System.IO.Path.Combine(pMtdsMngNoFolder, C_TNPDOC_dir, trkhyno_key, wno1_tnpdoc_hosok_naiyo, wno2_tnpdoc_hosok_naiyo, mng_no.ToString)
            'ディレクトリがない場合、作成する
            If Not System.IO.Directory.Exists(to_path) Then
                System.IO.Directory.CreateDirectory(to_path)
            End If
            'ファイル名を構築する
            to_path = System.IO.Path.Combine(to_path, wfile_ms)

            'NAS上のファイルをコピーする
            EcOrgIO.EcOrgFileIO.CopyFile(from_path, to_path)
        End If

    End Sub

    Function MyCDec(ByVal s As String) As Decimal
        If String.IsNullOrEmpty(s) Then
            Return 0
        Else
            Return CDec(s)
        End If
    End Function
#End Region




#Region "メッセージファイル作成処理"

    '''' <summary>
    '''' メッセージファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateCOM_MSG_MST(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        'メッセージを取得する。
    '        If Not dao.S008(dbr) Then
    '            Throw New Exception(String.Format("メッセージ検索でエラーが発生しました。"))
    '        End If

    '        'メッセージファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblCOM_MSG_MST(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("msg_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("msg_id"), ""))
    '            wRow("msg_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("msg_naiyo"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_MSG)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("メッセージファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblCOM_MSG_MST(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("msg_id")
    '    wDestDt.Columns.Add("msg_naiyo")
    'End Sub
#End Region

#Region "マスタ_工量ファイル作成処理"

    '''' <summary>
    '''' マスタ_工量ファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateMST_KORYO(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        'マスタ_工量を取得する。
    '        If Not dao.S009(dbr) Then
    '            Throw New Exception(String.Format("マスタ_工量検索でエラーが発生しました。"))
    '        End If

    '        'マスタ_工量ファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblMST_KORYO(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("koryo_zunit_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_zunit_cd"), ""))
    '            wRow("tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_start_ymd"), ""))
    '            wRow("tky_end_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_end_ymd"), ""))
    '            wRow("trtk_koryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_koryo"), ""))
    '            wRow("tekyo_koryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_koryo"), ""))
    '            wRow("koryo_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_tanka_kngk"), ""))
    '            wRow("kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kohi_tanka_kngk"), ""))
    '            wRow("koryo_kohi_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_kohi_sbt_cd"), ""))
    '            wRow("koryo_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_mskn"), ""))
    '            wRow("koryo_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_ms"), ""))
    '            wRow("koryo_tanka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_tanka_cd"), ""))
    '            wRow("wrms_ritu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("wrms_ritu"), ""))
    '            wRow("trtk_kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_kohi_tanka_kngk"), ""))
    '            wRow("tekyo_kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_kohi_tanka_kngk"), ""))
    '            wRow("koryo_idou_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_idou_kbn"), ""))
    '            wRow("koryo_cd_tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_cd_tky_start_ymd"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_KORYO)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("マスタ_工量ファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblMST_KORYO(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("koryo_zunit_cd")
    '    wDestDt.Columns.Add("tky_start_ymd")
    '    wDestDt.Columns.Add("tky_end_ymd")
    '    wDestDt.Columns.Add("trtk_koryo")
    '    wDestDt.Columns.Add("tekyo_koryo")
    '    wDestDt.Columns.Add("koryo_tanka_kngk")
    '    wDestDt.Columns.Add("kohi_tanka_kngk")
    '    wDestDt.Columns.Add("koryo_kohi_sbt_cd")
    '    wDestDt.Columns.Add("koryo_mskn")
    '    wDestDt.Columns.Add("koryo_ms")
    '    wDestDt.Columns.Add("koryo_tanka_cd")
    '    wDestDt.Columns.Add("wrms_ritu")
    '    wDestDt.Columns.Add("trtk_kohi_tanka_kngk")
    '    wDestDt.Columns.Add("tekyo_kohi_tanka_kngk")
    '    wDestDt.Columns.Add("koryo_idou_kbn")
    '    wDestDt.Columns.Add("koryo_cd_tky_start_ymd")
    'End Sub
#End Region

#Region "パラメータコードマスタファイル作成処理"

    '''' <summary>
    '''' パラメータコードマスタファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateCOM_PARAMATER_CD_MST(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        'パラメータコードマスタを取得する。
    '        If Not dao.S010(dbr) Then
    '            Throw New Exception(String.Format("パラメータコードマスタ検索でエラーが発生しました。"))
    '        End If

    '        'パラメータコードマスタファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblCOM_PARAMATER_CD_MST(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("prm_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_id"), ""))
    '            wRow("prm_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd"), ""))
    '            wRow("prm_cd_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd_ms"), ""))
    '            wRow("prm_cd_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd_abms"), ""))
    '            wRow("str1_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str1_hksu_value"), ""))
    '            wRow("str2_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str2_hksu_value"), ""))
    '            wRow("value1_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_hksu_value"), ""))
    '            wRow("value2_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value2_hksu_value"), ""))
    '            wRow("prm_hyz_zyun") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_hyz_zyun"), ""))
    '            wRow("HYZ_FLG") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("HYZ_FLG"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_COM_PARAMATER_CD_MST)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("パラメータコードマスタファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblCOM_PARAMATER_CD_MST(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("prm_id")
    '    wDestDt.Columns.Add("prm_cd")
    '    wDestDt.Columns.Add("prm_cd_ms")
    '    wDestDt.Columns.Add("prm_cd_abms")
    '    wDestDt.Columns.Add("str1_hksu_value")
    '    wDestDt.Columns.Add("str2_hksu_value")
    '    wDestDt.Columns.Add("value1_hksu_value")
    '    wDestDt.Columns.Add("value2_hksu_value")
    '    wDestDt.Columns.Add("prm_hyz_zyun")
    '    wDestDt.Columns.Add("HYZ_FLG")
    'End Sub
#End Region

#Region "品目マスタファイル作成処理"

    '''' <summary>
    '''' 品目マスタファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateMST_HNMK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        '品目マスタを取得する。
    '        If Not dao.S011(dbr) Then
    '            Throw New Exception(String.Format("品目マスタ検索でエラーが発生しました。"))
    '        End If

    '        '品目マスタファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblMST_HNMK(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("hnmk_nnsk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hnmk_nnsk_kbn"), ""))
    '            wRow("krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krhsk_cd"), ""))
    '            wRow("keiki_yr_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_yr_cd"), ""))
    '            wRow("keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_ktsk_ms"), ""))
    '            wRow("mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mado_su"), ""))
    '            wRow("ts_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_ktsk_cd"), ""))
    '            wRow("dnzsk_maker_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dnzsk_maker_cd"), ""))
    '            wRow("keiki_htknt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_htknt_kbn"), ""))
    '            wRow("hzk_ziry_hnbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzk_ziry_hnbt_cd"), ""))
    '            wRow("hzk_ziry_hts") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzk_ziry_hts"), ""))
    '            wRow("cktrm_d") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrm_d"), ""))
    '            wRow("huing_color_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("huing_color_kbn"), ""))
    '            wRow("tnsdai_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsdai_umu_flg"), ""))
    '            wRow("sm_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_tshsk_cd"), ""))
    '            wRow("tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_start_ymd"), ""))
    '            wRow("hnmk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hnmk_cd"), ""))
    '            wRow("keiki_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sbt_cd"), ""))
    '            wRow("keiki_hryhn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_hryhn_kbn"), ""))
    '            wRow("ryohu_kyki_yysu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ryohu_kyki_yysu"), ""))
    '            wRow("yosyrhn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yosyrhn_cd"), ""))
    '            wRow("yuko_kigen_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_kigen_ymd"), ""))
    '            wRow("mof_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mof_kbn"), ""))
    '            wRow("sm_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_taiko_kbn"), ""))
    '            wRow("sm_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_knthk_cd"), ""))
    '            wRow("ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_kino_umu_flg"), ""))
    '            wRow("khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khk_kino_umu_flg"), ""))
    '            wRow("gbdg_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("gbdg_umu_flg"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_HNMK)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("品目マスタファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblMST_HNMK(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("hnmk_nnsk_kbn")
    '    wDestDt.Columns.Add("krhsk_cd")
    '    wDestDt.Columns.Add("keiki_yr_cd")
    '    wDestDt.Columns.Add("keiki_ktsk_ms")
    '    wDestDt.Columns.Add("mado_su")
    '    wDestDt.Columns.Add("ts_ktsk_cd")
    '    wDestDt.Columns.Add("dnzsk_maker_cd")
    '    wDestDt.Columns.Add("keiki_htknt_kbn")
    '    wDestDt.Columns.Add("hzk_ziry_hnbt_cd")
    '    wDestDt.Columns.Add("hzk_ziry_hts")
    '    wDestDt.Columns.Add("cktrm_d")
    '    wDestDt.Columns.Add("huing_color_kbn")
    '    wDestDt.Columns.Add("tnsdai_umu_flg")
    '    wDestDt.Columns.Add("sm_tshsk_cd")
    '    wDestDt.Columns.Add("tky_start_ymd")
    '    wDestDt.Columns.Add("hnmk_cd")
    '    wDestDt.Columns.Add("keiki_sbt_cd")
    '    wDestDt.Columns.Add("keiki_hryhn_kbn")
    '    wDestDt.Columns.Add("ryohu_kyki_yysu")
    '    wDestDt.Columns.Add("yosyrhn_cd")
    '    wDestDt.Columns.Add("yuko_kigen_ymd")
    '    wDestDt.Columns.Add("mof_kbn")
    '    wDestDt.Columns.Add("sm_taiko_kbn")
    '    wDestDt.Columns.Add("sm_knthk_cd")
    '    wDestDt.Columns.Add("ts_kino_umu_flg")
    '    wDestDt.Columns.Add("khk_kino_umu_flg")
    '    wDestDt.Columns.Add("gbdg_umu_flg")
    'End Sub
#End Region

#Region "計器型式マスタファイル作成処理"

    '''' <summary>
    '''' 計器型式マスタファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateMST_KEIKI_KTSK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        '計器型式マスタを取得する。
    '        If Not dao.S012(dbr) Then
    '            Throw New Exception(String.Format("計器型式マスタ検索でエラーが発生しました。"))
    '        End If

    '        '計器型式マスタファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblMST_KEIKI_KTSK(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("kotea_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kotea_kbn"), ""))
    '            wRow("ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ktsk_cd"), ""))
    '            wRow("ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ksyu_cd"), ""))
    '            wRow("krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krhsk_cd"), ""))
    '            wRow("ts_kotea_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_kotea_kbn"), ""))
    '            wRow("keiki_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sbt_cd"), ""))
    '            wRow("keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_ktsk_ms"), ""))
    '            wRow("nozok_keiki_reuse_hnti_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nozok_keiki_reuse_hnti_cd"), ""))
    '            wRow("tk_keiki_reuse_hnti_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tk_keiki_reuse_hnti_cd"), ""))
    '            wRow("nozok_ts_trtk_krksbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nozok_ts_trtk_krksbt_cd"), ""))
    '            wRow("tk_ts_trtk_krksbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tk_ts_trtk_krksbt_cd"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_KEIKI_KTSK)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("計器型式マスタファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblMST_KEIKI_KTSK(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("kotea_kbn")
    '    wDestDt.Columns.Add("ktsk_cd")
    '    wDestDt.Columns.Add("ksyu_cd")
    '    wDestDt.Columns.Add("krhsk_cd")
    '    wDestDt.Columns.Add("ts_kotea_kbn")
    '    wDestDt.Columns.Add("keiki_sbt_cd")
    '    wDestDt.Columns.Add("keiki_ktsk_ms")
    '    wDestDt.Columns.Add("nozok_keiki_reuse_hnti_cd")
    '    wDestDt.Columns.Add("tk_keiki_reuse_hnti_cd")
    '    wDestDt.Columns.Add("nozok_ts_trtk_krksbt_cd")
    '    wDestDt.Columns.Add("tk_ts_trtk_krksbt_cd")
    'End Sub
#End Region

#Region "マスタ_町字ファイル作成処理"

    '''' <summary>
    '''' マスタ_町字ファイル作成処理
    '''' </summary>
    '''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '''' <param name="pDownldRenno">ダウンロード番号</param>
    '''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '''' <param name="dao">MJ4103_DAO</param>
    '''' <returns>true・falseが返却される</returns>
    '''' <remarks></remarks>
    'Private Function CreateMST_TYAZA(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '    Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '    Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '    Dim wRow As System.Data.DataRow = Nothing

    '    Dim wFileName As String = ""                'ファイル名

    '    Dim wFileSbt As String = " "
    '    Dim wItemCd As String = ""
    '    Dim wItemNaiyo As String = ""
    '    Dim i As Integer = 0

    '    Try
    '        'Functionの開始ログ作成
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '        'マスタ_町字を取得する。
    '        If Not dao.S013(dbr) Then
    '            Throw New Exception(String.Format("マスタ_町字検索でエラーが発生しました。"))
    '        End If

    '        'マスタ_町字ファイル用Datatableの初期化
    '        wDataTable = New DataTable
    '        initDtblMST_TYAZA(wDataTable)

    '        'Datatableの初期化
    '        wDataTable.Rows.Clear()

    '        'FETCH
    '        While (dbr.DataStruct.Read)
    '            '持出対象データを設定する
    '            wRow = wDataTable.NewRow

    '            wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '            wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '            wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '            wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '            wRow("tdhkn_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_add_cd"), ""))
    '            wRow("siku_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_add_cd"), ""))
    '            wRow("oazat_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_add_cd"), ""))
    '            wRow("azatm_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_add_cd"), ""))
    '            wRow("tdhkn_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_mskn"), ""))
    '            wRow("siku_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_mskn"), ""))
    '            wRow("oazat_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_mskn"), ""))
    '            wRow("azatm_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_mskn"), ""))
    '            wRow("tdhkn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_ms"), ""))
    '            wRow("siku_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_ms"), ""))
    '            wRow("oazat_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_ms"), ""))
    '            wRow("azatm_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_ms"), ""))
    '            wRow("add_zipcd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("add_zipcd"), ""))
    '            wRow("new_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("new_add_cd"), ""))
    '            wRow("sikou_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sikou_ym"), ""))
    '            wRow("haisi_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("haisi_ym"), ""))
    '            wRow("new_add_cd_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("new_add_cd_ym"), ""))
    '            wRow("name_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("name_henko_ym"), ""))
    '            wRow("zipcd_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zipcd_henko_ym"), ""))
    '            wRow("tino_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tino_henko_ym"), ""))
    '            wRow("barcd_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("barcd_naiyo"), ""))
    '            wRow("oyako_knk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oyako_knk_naiyo"), ""))
    '            wRow("cs_barcd_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_barcd_henko_ym"), ""))
    '            wRow("oyako_knk_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oyako_knk_henko_ym"), ""))
    '            wRow("tusyo_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tusyo_flg"), ""))

    '            wDataTable.Rows.Add(wRow)
    '        End While

    '        'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '        wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_TYAZA)
    '        If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '            Throw New Exception(String.Format("マスタ_町字ファイル出力に失敗しました。"))
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '        Return False
    '    Finally
    '        If dbr IsNot Nothing Then
    '            dbr.Close()
    '            dbr = Nothing
    '        End If
    '    End Try

    'End Function

    'Sub initDtblMST_TYAZA(ByVal wDestDt As System.Data.DataTable)
    '    ' 空テーブルを作る
    '    wDestDt.Columns.Add("saksi_date")
    '    wDestDt.Columns.Add("upd_date")
    '    wDestDt.Columns.Add("sousa_user_id")
    '    wDestDt.Columns.Add("sousa_appli_cd")
    '    wDestDt.Columns.Add("tdhkn_add_cd")
    '    wDestDt.Columns.Add("siku_add_cd")
    '    wDestDt.Columns.Add("oazat_add_cd")
    '    wDestDt.Columns.Add("azatm_add_cd")
    '    wDestDt.Columns.Add("tdhkn_mskn")
    '    wDestDt.Columns.Add("siku_mskn")
    '    wDestDt.Columns.Add("oazat_mskn")
    '    wDestDt.Columns.Add("azatm_mskn")
    '    wDestDt.Columns.Add("tdhkn_ms")
    '    wDestDt.Columns.Add("siku_ms")
    '    wDestDt.Columns.Add("oazat_ms")
    '    wDestDt.Columns.Add("azatm_ms")
    '    wDestDt.Columns.Add("add_zipcd")
    '    wDestDt.Columns.Add("new_add_cd")
    '    wDestDt.Columns.Add("sikou_ym")
    '    wDestDt.Columns.Add("haisi_ym")
    '    wDestDt.Columns.Add("new_add_cd_ym")
    '    wDestDt.Columns.Add("name_henko_ym")
    '    wDestDt.Columns.Add("zipcd_henko_ym")
    '    wDestDt.Columns.Add("tino_henko_ym")
    '    wDestDt.Columns.Add("barcd_naiyo")
    '    wDestDt.Columns.Add("oyako_knk_naiyo")
    '    wDestDt.Columns.Add("cs_barcd_henko_ym")
    '    wDestDt.Columns.Add("oyako_knk_henko_ym")
    '    wDestDt.Columns.Add("tusyo_flg")
    'End Sub
#End Region

#Region "定数マスタファイル作成処理"

    '    ''' <summary>
    '    ''' 定数マスタファイル作成処理
    '    ''' </summary>
    '    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    '    ''' <param name="pDownldRenno">ダウンロード番号</param>
    '    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    '    ''' <param name="dao">MJ4103_DAO</param>
    '    ''' <returns>true・falseが返却される</returns>
    '    ''' <remarks></remarks>
    '    Private Function CreateCOM_ZYOUSU_MST(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

    '        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

    '        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
    '        Dim wRow As System.Data.DataRow = Nothing

    '        Dim wFileName As String = ""                'ファイル名

    '        Dim wFileSbt As String = " "
    '        Dim wItemCd As String = ""
    '        Dim wItemNaiyo As String = ""
    '        Dim i As Integer = 0

    '        Try
    '            'Functionの開始ログ作成
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

    '            '定数マスタを取得する。
    '            If Not dao.S014(dbr) Then
    '                Throw New Exception(String.Format("定数マスタ検索でエラーが発生しました。"))
    '            End If

    '            'メッセージファイル用Datatableの初期化
    '            wDataTable = New DataTable
    '            initDtblCOM_ZYOUSU_MST(wDataTable)

    '            'Datatableの初期化
    '            wDataTable.Rows.Clear()

    '            'FETCH
    '            While (dbr.DataStruct.Read)
    '                '持出対象データを設定する
    '                wRow = wDataTable.NewRow

    '                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
    '                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
    '                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
    '                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
    '                wRow("zyousu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zyousu_cd"), ""))
    '                wRow("yuko_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_start_ymd"), ""))
    '                wRow("yuko_end_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_end_ymd"), ""))
    '                wRow("str1_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str1_zyousu_value"), ""))
    '                wRow("str2_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str2_zyousu_value"), ""))
    '                wRow("value1_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_zyousu_value"), ""))
    '                wRow("value2_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value2_zyousu_value"), ""))

    '                wDataTable.Rows.Add(wRow)
    '            End While

    '            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
    '            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_COM_ZYOUSU_MST)
    '            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
    '                Throw New Exception(String.Format("定数マスタファイル出力に失敗しました。"))
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '            If dbr IsNot Nothing Then
    '                dbr.Close()
    '                dbr = Nothing
    '            End If
    '        End Try

    '    End Function

    '    Sub initDtblCOM_ZYOUSU_MST(ByVal wDestDt As System.Data.DataTable)
    '        ' 空テーブルを作る
    '        wDestDt.Columns.Add("saksi_date")
    '        wDestDt.Columns.Add("upd_date")
    '        wDestDt.Columns.Add("sousa_user_id")
    '        wDestDt.Columns.Add("sousa_appli_cd")
    '        wDestDt.Columns.Add("zyousu_cd")
    '        wDestDt.Columns.Add("yuko_start_ymd")
    '        wDestDt.Columns.Add("yuko_end_ymd")
    '        wDestDt.Columns.Add("str1_zyousu_value")
    '        wDestDt.Columns.Add("str2_zyousu_value")
    '        wDestDt.Columns.Add("value1_zyousu_value")
    '        wDestDt.Columns.Add("value2_zyousu_value")
    '    End Sub
#End Region

#Region "取替_取替票（高圧他）ファイル作成処理"

    'Rev002-Start

    ''' <summary>
    ''' 取替_取替票（高圧他）ファイル作成処理(施工)
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByRef pExistFlag As Boolean) As Boolean
        '施工：抜取検査フラグにFalseを設定して実行する
        Return CreateTRKE_TRKEHY_KAHK(pMtdsMngNoFolder, pDownldRenno, pKeikiDwldSbtCd, dao, False, pExistFlag)
    End Function


    ''' <summary>
    ''' 取替_取替票（高圧他）ファイル作成処理(抜取)
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_KAHK_Nktr(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByRef pExistFlag As Boolean) As Boolean
        '施工：抜取検査フラグにFalseを設定して実行する
        Return CreateTRKE_TRKEHY_KAHK(pMtdsMngNoFolder, pDownldRenno, pKeikiDwldSbtCd, dao, True, pExistFlag)
    End Function


    ''' <summary>
    ''' 取替_取替票（高圧他）ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <param name="NktrKnsFlag">抜取検査フラグ</param>
    ''' <param name="ExistFlag">テーブル存在フラグ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO, ByVal NktrKnsFlag As Boolean, ByRef ExistFlag As Boolean) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・取替_取替票（高圧他）を取得する。
            If Not dao.S016(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・取替_取替票（高圧他）検索でエラーが発生しました。"))
            End If


            '高圧の取替票データが無い場合はTrueで戻るが、テーブルが存在しない（ExistFlag＝False）で上位に知らせる
            If Not dbr.DataStruct.HasRows Then
                ExistFlag = False
                Return True
            End If
            'テーブルは存在する
            ExistFlag = True


            '取替票ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_TRKEHY_KAHK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("file_sbt") = wFileSbt
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("prcmg_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prcmg_cd"), ""))
                wRow("trke_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_sbt_cd"), ""))
                wRow("kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kzkis_mdgt_cd"), ""))
                wRow("kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kozitn_cd"), ""))
                wRow("skosya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skosya_cd"), ""))
                wRow("skosya_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skosya_ms"), ""))
                wRow("kiyk_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_no"), ""))
                wRow("kiyk_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_sbt_cd"), ""))
                wRow("kiyk_cs_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_cs_mskn"), ""))
                wRow("kiyk_cs_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_cs_ms"), ""))
                wRow("dig14_cs_tel") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig14_cs_tel_str"), ""))
                wRow("cs_tdhkn_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_tdhkn_add_cd"), ""))
                wRow("cs_siku_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_siku_add_cd"), ""))
                wRow("cs_oazat_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_oazat_add_cd"), ""))
                wRow("cs_azatm_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_azatm_add_cd"), ""))
                wRow("tdhkn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_tdhkn_add_cd_ms"), ""))
                wRow("siku_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_siku_add_cd_ms"), ""))
                wRow("oazat_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_oazat_add_cd_ms"), ""))
                wRow("azatm_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_azatm_add_cd_ms"), ""))
                wRow("cs_addhsk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_addhsk_naiyo"), ""))
                wRow("mkhy_senro_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_senro_cd"), ""))
                wRow("mkhy_kansn_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_kansn_no"), ""))
                wRow("mkhy_bunk1_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk1_no"), ""))
                wRow("mkhy_bunk2_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk2_no"), ""))
                wRow("mkhy_bunk3_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mkhy_bunk3_no"), ""))
                wRow("stbit_xzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stbit_xzahyo"), ""))
                wRow("stbit_yzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stbit_yzahyo"), ""))
                wRow("iewku_stdp_xzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("iewku_stdp_xzahyo"), ""))
                wRow("iewku_stdp_yzahyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("iewku_stdp_yzahyo"), ""))
                wRow("cs_kiyk_dnryk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_kiyk_dnryk"), ""))
                wRow("std_kensn_dd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("std_kensn_dd"), ""))
                wRow("kensn_yotei_togt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_togt_md"), ""))
                wRow("kensn_yotei_ykgt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_ykgt_md"), ""))
                wRow("kensn_yotei_yygt_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kensn_yotei_yygt_md"), ""))
                wRow("kyokyu_spot_tokti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kyokyu_spot_tokti_no"), ""))
                wRow("zyuky_ukky_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zyuky_ukky_kbn"), ""))
                wRow("hist_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hist_flg"), ""))
                wRow("brot_sett_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("brot_sett_flg"), ""))
                wRow("knti_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("knti_sbt_cd"), ""))
                wRow("surge_yksi_siyo_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("surge_yksi_siyo_umu_flg"), ""))
                wRow("tidn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tidn_kbn"), ""))
                wRow("tidn_riyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tidn_riyu_cd"), ""))
                wRow("sagyo_time_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sagyo_time_kbn_cd"), ""))
                wRow("sagyo_start_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sagyo_start_hms"), ""))
                wRow("sagyo_end_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sagyo_end_hms"), ""))
                wRow("tnsb_reuse_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsb_reuse_kbn"), ""))
                wRow("trke_tis_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_tis_kbn"), ""))
                wRow("trke_tis_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_tis_abms"), ""))
                wRow("hnsik_kykbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hnsik_kykbn_cd"), ""))
                wRow("kohu_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kohu_ymd"), ""))
                wRow("skohu_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skohu_ymd"), ""))
                wRow("hiky_yotei_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hiky_yotei_ymd"), ""))
                wRow("sksti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sksti_ymd"), ""))
                wRow("skssti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skssti_ymd"), ""))
                wRow("skyti_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skyti_ymd"), ""))
                wRow("syun_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("syun_ymd"), ""))
                wRow("keiki_kmaws_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_kmaws_no"), ""))
                wRow("keiki_sitei_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sitei_no"), ""))
                wRow("kiyk_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kiyk_kbn_cd"), ""))
                wRow("keiki_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_kbn_cd"), ""))
                wRow("sm_zyri_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_zyri_kbn_cd"), ""))
                wRow("ftikk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ftikk_flg"), ""))
                wRow("dtkk_sougo_drkei_tekyo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_sougo_drkei_tekyo_szsu"), ""))
                wRow("dtkk_sougo_drkei_trtk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_sougo_drkei_trtk_szsu"), ""))
                wRow("dtkk_max_jydrrkei_tekyo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_max_jydrrkei_tekyo_szsu"), ""))
                wRow("dtkk_max_jydrrkei_trtk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_max_jydrrkei_trtk_szsu"), ""))
                wRow("dtkk_rsyk_tekyo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_rsyk_tekyo_szsu"), ""))
                wRow("dtkk_rsyk_trtk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_rsyk_trtk_szsu"), ""))
                wRow("dtkk_rsmk_tekyo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_rsmk_tekyo_szsu"), ""))
                wRow("dtkk_rsmk_trtk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_rsmk_trtk_szsu"), ""))
                wRow("dtkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dtkk_zyrt"), ""))
                wRow("nktr_kensa_zyokyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktr_kensa_zyokyo_cd"), ""))
                If NktrKnsFlag Then
                    '抜取検査の場合、抜取検査タブレット受信年月日・抜取検査タブレット受信_担当者コードを設定する
                    wRow("ntkns_tblt_recv_ymd") = SyoriYMD
                    wRow("ntkns_tblt_recv_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("D_keiki_user_id"), ""))
                Else
                    wRow("ntkns_tblt_recv_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tblt_recv_ymd"), ""))
                    wRow("ntkns_tblt_recv_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tblt_recv_tntsy_cd"), ""))
                End If
                wRow("ntkns_zissi_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_zissi_ymd"), ""))
                wRow("ntkns_zissi_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_zissi_tntsy_cd"), ""))
                wRow("ntkns_kka_upld_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_kka_upld_ymd"), ""))
                wRow("ntkns_kka_upld_tntsy_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_kka_upld_tntsy_cd"), ""))
                wRow("nktr_kensa_zyokyo_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktr_kensa_zyokyo_abms"), ""))
                wRow("tkkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_id"), ""))
                wRow("tkkk_kkinf_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kkinf_id"), ""))
                wRow("tkkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kesbt_kbn"), ""))
                wRow("tkkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_krhsk_cd"), ""))
                wRow("tkkk_keiki_ssnsk_so_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ssnsk_so_cd"), ""))
                wRow("tkkk_keiki_ssnsk_line_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ssnsk_line_cd"), ""))
                wRow("tkkk_kiry_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kiry_dnat_cd"), ""))
                wRow("tkkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_yr"), ""))
                wRow("tkkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tshsk_cd"), ""))
                wRow("tkkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ktsk_ms"), ""))

                '個別の場合、計器型式名称が空のためマスタ計器型式より抽出する  クエリ側で対処　削除予定
                'If String.IsNullOrWhiteSpace(wRow("tkkk_keiki_ktsk_ms").ToString) Then
                '  wRow("tkkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_ktsk_ms"), ""))
                'End If

                wRow("tkkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_keiki_ktsk_cd"), ""))
                wRow("tkkk_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_ksyu_cd"), ""))
                '撤去計器_製造番号がないので削除予定
                'wRow("") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item(""), ""))
                wRow("tkkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_taiko_kbn"), ""))
                wRow("tkkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_knthk_cd"), ""))
                wRow("tkkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_ts_kino_umu_flg"), ""))
                wRow("tkkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_khk_kino_umu_flg"), ""))
                wRow("tkkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gaibu_output_umu_flg"), ""))
                wRow("tkkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_yuko_kigen_ym"), ""))
                wRow("tkkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_seizo_yy"), ""))
                wRow("tkkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_zyrt"), ""))
                wRow("tkkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_digsu"), ""))
                wRow("tkkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_szsya_cd"), ""))
                wRow("tkkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_yy"), ""))
                wRow("tkkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_zytr_szsu"), ""))
                wRow("tkkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gytr_szsu"), ""))
                wRow("tkkk_sougo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_sougo_szsu"), ""))
                wRow("tkkk_max_zyyo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_max_zyyo_szsu"), ""))
                wRow("tkkk_rsyk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_rsyk_szsu"), ""))
                wRow("tkkk_rsmk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_rsmk_szsu"), ""))
                wRow("tkkk_hnsik_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_ksyu_cd"), ""))
                wRow("no1_tkkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_tkkk_hnsik_seizo_no"), ""))
                wRow("no2_tkkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_tkkk_hnsik_seizo_no"), ""))
                wRow("tkkk_hnsik_vctct_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vctct_ktsk_cd"), ""))
                wRow("tkkk_hnsik_vctct_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vctct_seizo_no"), ""))
                wRow("tkkk_hnsik_vctct_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vctct_seizo_yy"), ""))
                wRow("tkkk_hnsik_ct2_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_ct2_ktsk_cd"), ""))
                wRow("tkkk_hnsik_ct2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_ct2_seizo_no"), ""))
                wRow("tkkk_hnsik_ct2_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_ct2_seizo_yy"), ""))
                wRow("tkkk_hnsik_vt1_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt1_ktsk_cd"), ""))
                wRow("tkkk_hnsik_vt1_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt1_seizo_no"), ""))
                wRow("tkkk_hnsik_vt1_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt1_seizo_yy"), ""))
                wRow("tkkk_hnsik_vt2_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt2_ktsk_cd"), ""))
                wRow("tkkk_hnsik_vt2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt2_seizo_no"), ""))
                wRow("tkkk_hnsik_vt2_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_vt2_seizo_yy"), ""))
                wRow("tkkk_hyzki_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hyzki_seizo_no"), ""))
                wRow("tkkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_yuko_kigen_ym"), ""))
                wRow("tkkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_hnsik_gknti_no"), ""))
                wRow("tkkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_mado_su"), ""))
                wRow("tkkk_smkey1") = ""
                wRow("tkkk_smkey2") = ""
                wRow("tkkk_smkey_startdate") = ""
                wRow("tkkk_smkey_enddate") = ""
                wRow("tkkk_mac_address") = ""
                wRow("tkkk_tsn_id") = ""
                wRow("ttkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_id"), ""))
                wRow("ttkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_kesbt_kbn"), ""))
                wRow("ttkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_krhsk_cd"), ""))
                wRow("ttkk_keiki_ssnsk_so_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ssnsk_so_cd"), ""))
                wRow("ttkk_keiki_ssnsk_line_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ssnsk_line_cd"), ""))
                wRow("ttkk_kiry_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_kiry_dnat_cd"), ""))
                wRow("ttkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_yr"), ""))
                wRow("ttkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tshsk_cd"), ""))
                wRow("ttkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ktsk_ms"), ""))
                wRow("ttkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_keiki_ktsk_cd"), ""))
                wRow("ttkk_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_ksyu_cd"), ""))
                wRow("ttkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_taiko_kbn"), ""))
                wRow("ttkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_knthk_cd"), ""))
                wRow("ttkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_ts_kino_umu_flg"), ""))
                wRow("ttkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_khk_kino_umu_flg"), ""))
                wRow("ttkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gaibu_output_umu_flg"), ""))
                wRow("ttkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_yuko_kigen_ym"), ""))
                wRow("ttkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_seizo_yy"), ""))
                wRow("ttkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_zyrt"), ""))
                wRow("ttkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_digsu"), ""))
                wRow("ttkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_szsya_cd"), ""))
                wRow("ttkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_yy"), ""))
                wRow("ttkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_zytr_szsu"), ""))
                wRow("ttkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gytr_szsu"), ""))
                wRow("ttkk_sougo_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_sougo_szsu"), ""))
                wRow("ttkk_rsyk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_rsyk_szsu"), ""))
                wRow("ttkk_rsmk_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_rsmk_szsu"), ""))
                wRow("ttkk_hnsik_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_ksyu_cd"), ""))
                wRow("no1_ttkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_ttkk_hnsik_seizo_no"), ""))
                wRow("no2_ttkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_ttkk_hnsik_seizo_no"), ""))
                wRow("ttkk_hnsik_vctct_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vctct_ktsk_cd"), ""))
                wRow("ttkk_hnsik_vctct_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vctct_seizo_no"), ""))
                wRow("ttkk_hnsik_vctct_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vctct_seizo_yy"), ""))
                wRow("ttkk_hnsik_ct2_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_ct2_ktsk_cd"), ""))
                wRow("ttkk_hnsik_ct2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_ct2_seizo_no"), ""))
                wRow("ttkk_hnsik_ct2_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_ct2_seizo_yy"), ""))
                wRow("ttkk_hnsik_vt1_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt1_ktsk_cd"), ""))
                wRow("ttkk_hnsik_vt1_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt1_seizo_no"), ""))
                wRow("ttkk_hnsik_vt1_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt1_seizo_yy"), ""))
                wRow("ttkk_hnsik_vt2_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt2_ktsk_cd"), ""))
                wRow("ttkk_hnsik_vt2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt2_seizo_no"), ""))
                wRow("ttkk_hnsik_vt2_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_vt2_seizo_yy"), ""))
                wRow("ttkk_hyzki_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hyzki_seizo_no"), ""))
                wRow("ttkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_yuko_kigen_ym"), ""))
                wRow("ttkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_hnsik_gknti_no"), ""))
                wRow("ttkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_mado_su"), ""))
                wRow("hzkst_tstnmt_info_tsn_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_tstnmt_info_tsn_id"), ""))
                wRow("hzkst_tstnmt_info_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_tstnmt_info_tshsk_cd"), ""))
                wRow("hzkst_htan_info_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_htan_info_ksyu_cd"), ""))
                wRow("hzkst_htan_info_kesbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_htan_info_kesbt_cd"), ""))
                wRow("hzkst_htan_info_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_htan_info_ktsk_cd"), ""))
                wRow("hzkst_htan_info_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_htan_info_seizo_no"), ""))
                wRow("hzkst_htan_info_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_htan_info_seizo_yy"), ""))
                wRow("hzkst_ts_info_ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_ts_info_ksyu_cd"), ""))
                wRow("hzkst_ts_info_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_ts_info_ktsk_cd"), ""))
                wRow("hzkst_ts_info_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_ts_info_seizo_no"), ""))
                wRow("hzkst_ts_info_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzkst_ts_info_seizo_yy"), ""))
                wRow("kdkk_keiki_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_id"), ""))
                wRow("kdkk_kesbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_kesbt_kbn"), ""))
                wRow("kdkk_krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_krhsk_cd"), ""))
                wRow("kdkk_keiki_ssnsk_so_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ssnsk_so_cd"), ""))
                wRow("kdkk_keiki_ssnsk_line_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ssnsk_line_cd"), ""))
                wRow("kdkk_kiry_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_kiry_dnat_cd"), ""))
                wRow("kdkk_keiki_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_yr"), ""))
                wRow("kdkk_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tshsk_cd"), ""))
                wRow("kdkk_keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ktsk_ms"), ""))
                wRow("kdkk_keiki_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_keiki_ktsk_cd"), ""))
                wRow("kdkk_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_taiko_kbn"), ""))
                wRow("kdkk_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_knthk_cd"), ""))
                wRow("kdkk_ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_ts_kino_umu_flg"), ""))
                wRow("kdkk_khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_khk_kino_umu_flg"), ""))
                wRow("kdkk_gaibu_output_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_gaibu_output_umu_flg"), ""))
                wRow("kdkk_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_yuko_kigen_ym"), ""))
                wRow("kdkk_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_seizo_yy"), ""))
                wRow("kdkk_zyrt") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_zyrt"), ""))
                wRow("kdkk_digsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_digsu"), ""))
                wRow("kdkk_tnsb_szsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_szsya_cd"), ""))
                wRow("kdkk_tnsb_seizo_yy") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_yy"), ""))
                wRow("kdkk_zytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_zytr_szsu"), ""))
                wRow("kdkk_gytr_szsu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_gytr_szsu"), ""))
                wRow("no1_kdkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_kdkk_hnsik_seizo_no"), ""))
                wRow("no2_kdkk_hnsik_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_kdkk_hnsik_seizo_no"), ""))
                wRow("kdkk_hnsik_vctct_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_vctct_seizo_no"), ""))
                wRow("kdkk_hnsik_ct2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_ct2_seizo_no"), ""))
                wRow("kdkk_hnsik_vt1_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_vt1_seizo_no"), ""))
                wRow("kdkk_hnsik_vt2_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_vt2_seizo_no"), ""))
                wRow("kdkk_hnsik_yuko_kigen_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_yuko_kigen_ym"), ""))
                wRow("kdkk_hnsik_gknti_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_hnsik_gknti_no"), ""))
                wRow("kdkk_mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_mado_su"), ""))
                wRow("tdnstr_ts_dosa_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_ts_dosa_kbn"), ""))
                wRow("tdnstr_tdnt_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_tdnt_su"), ""))
                wRow("tdnstr_tdnstr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdnstr_tdnstr_time"), ""))
                wRow("hkgds_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_kbn"), ""))
                wRow("hkgds_start_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_start_md"), ""))
                wRow("hkgds_end_md") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hkgds_end_md"), ""))
                wRow("no1_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_hkgds_tdn_hms"), ""))
                wRow("no1_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_hkgds_sydan_hms"), ""))
                wRow("no2_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_hkgds_tdn_hms"), ""))
                wRow("no2_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_hkgds_sydan_hms"), ""))
                wRow("no3_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_hkgds_tdn_hms"), ""))
                wRow("no3_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_hkgds_sydan_hms"), ""))
                wRow("no4_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_hkgds_tdn_hms"), ""))
                wRow("no4_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_hkgds_sydan_hms"), ""))
                wRow("no5_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_hkgds_tdn_hms"), ""))
                wRow("no5_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_hkgds_sydan_hms"), ""))
                wRow("no6_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no6_hkgds_tdn_hms"), ""))
                wRow("no6_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no6_hkgds_sydan_hms"), ""))
                wRow("no7_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no7_hkgds_tdn_hms"), ""))
                wRow("no7_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no7_hkgds_sydan_hms"), ""))
                wRow("no8_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no8_hkgds_tdn_hms"), ""))
                wRow("no8_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no8_hkgds_sydan_hms"), ""))
                wRow("no9_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no9_hkgds_tdn_hms"), ""))
                wRow("no9_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no9_hkgds_sydan_hms"), ""))
                wRow("no10_hkgds_tdn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no10_hkgds_tdn_hms"), ""))
                wRow("no10_hkgds_sydan_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no10_hkgds_sydan_hms"), ""))
                wRow("kbttd_start_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kbttd_start_date"), ""))
                wRow("kbttd_end_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kbttd_end_date"), ""))
                wRow("hoka_kihi_kbn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_kihi_kbn_cd"), ""))
                wRow("hoka_gmn_flckr_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_gmn_flckr_kbn"), ""))
                wRow("hoka_evnt_krk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_evnt_krk_kbn"), ""))
                wRow("hoka_hoka_hyz_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_hoka_hyz_kbn"), ""))
                wRow("hoka_hoka_hyz_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hoka_hoka_hyz_value"), ""))
                wRow("hksgk_hksgn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_hksgn_kbn"), ""))
                wRow("hksgk_hka_dnr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_hka_dnr"), ""))
                wRow("hksgk_attny_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_time"), ""))
                wRow("hksgk_attny_kaisu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_kaisu"), ""))
                wRow("hksgk_attny_kaisu_clr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgk_attny_kaisu_clr_time"), ""))
                wRow("hksgr_hksgn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_hksgn_kbn"), ""))
                wRow("hksgr_hka_dnr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_hka_dnr"), ""))
                wRow("hksgr_attny_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_time"), ""))
                wRow("hksgr_attny_kaisu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_kaisu"), ""))
                wRow("hksgr_attny_kaisu_clr_time") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hksgr_attny_kaisu_clr_time"), ""))
                wRow("khkmk_keiki_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_keiki_kbn"), ""))
                wRow("khkmk_mg_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_kbn"), ""))
                wRow("khkmk_hnsik_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_hnsik_kbn"), ""))
                wRow("khkmk_hnsik_st_iti_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_hnsik_st_iti_kbn"), ""))
                wRow("khkmk_wrms_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_wrms_kbn"), ""))
                wRow("khkmk_mg_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_umu_flg"), ""))
                wRow("khkmk_mg_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_yr"), ""))
                wRow("tuika_kohi_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tuika_kohi_umu_flg"), ""))
                wRow("zippi_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zippi_umu_flg"), ""))
                wRow("no1_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_no"), ""))
                wRow("no1_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kbn"), ""))
                wRow("no1_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kngk"), ""))
                wRow("no2_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_no"), ""))
                wRow("no2_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kbn"), ""))
                wRow("no2_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kngk"), ""))
                wRow("no3_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_no"), ""))
                wRow("no3_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kbn"), ""))
                wRow("no3_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kngk"), ""))
                wRow("no4_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_zippi_no"), ""))
                wRow("no4_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_zippi_kbn"), ""))
                wRow("no4_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_zippi_kngk"), ""))
                wRow("no5_zippi_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_zippi_no"), ""))
                wRow("no5_zippi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_zippi_kbn"), ""))
                wRow("no5_zippi_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_zippi_kngk"), ""))
                wRow("tenp_file_mng_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tenp_file_mng_no"), ""))
                wRow("ryssy_tenp_file_mng_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ryssy_tenp_file_mng_no"), ""))
                wRow("rrzk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("rrzk_naiyo"), ""))
                wRow("prcmg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prcmg_cd_abms"), ""))
                wRow("trke_sbt_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trke_sbt_abms"), ""))
                wRow("brot_sett_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("brot_sett_flg_abms"), ""))
                wRow("tkkk_kesbt_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_kesbt_kbn_abms"), ""))
                wRow("tkkk_krhsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_krhsk_cd_abms"), ""))
                wRow("tkkk_tshsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tshsk_cd_abms"), ""))
                wRow("tkkk_knthk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_knthk_cd_abms"), ""))
                wRow("tkkk_ts_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_ts_kino_umu_flg_abms"), ""))
                wRow("tkkk_khk_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_khk_kino_umu_flg_abms"), ""))
                wRow("tkkk_gaibu_output_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_gaibu_output_umu_flg_abms"), ""))
                wRow("tkkk_taiko_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_taiko_kbn_abms"), ""))
                wRow("ttkk_kesbt_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_kesbt_kbn_abms"), ""))
                wRow("ttkk_krhsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_krhsk_cd_abms"), ""))
                wRow("ttkk_tshsk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tshsk_cd_abms"), ""))
                wRow("ttkk_knthk_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_knthk_abms"), ""))
                wRow("ttkk_ts_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_ts_kino_umu_flg_abms"), ""))
                wRow("ttkk_khk_kino_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_khk_kino_umu_flg_abms"), ""))
                wRow("ttkk_gaibu_output_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_gaibu_output_umu_flg_ms"), ""))
                wRow("ttkk_taiko_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_taiko_kbn_abms"), ""))
                wRow("knti_sbt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("knti_sbt_cd_ms"), ""))
                wRow("surge_yksi_siyo_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("surge_yksi_siyo_umu_flg_abms"), ""))
                wRow("tidn_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tidn_kbn_abms"), ""))
                wRow("tnsb_reuse_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsb_reuse_kbn_abms"), ""))
                wRow("khkmk_keiki_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_keiki_kbn_abms"), ""))
                wRow("khkmk_mg_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_kbn_abms"), ""))
                wRow("khkmk_hnsik_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_hnsik_kbn_abms"), ""))
                wRow("khkmk_wrms_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_wrms_kbn_abms"), ""))
                wRow("khkmk_mg_umu_flg_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khkmk_mg_umu_flg_abms"), ""))
                wRow("no1_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no1_zippi_kbn_abms"), ""))
                wRow("no2_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no2_zippi_kbn_abms"), ""))
                wRow("no3_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no3_zippi_kbn_abms"), ""))
                wRow("no4_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no4_zippi_kbn_abms"), ""))
                wRow("no5_zippi_kbn_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("no5_zippi_kbn_abms"), ""))
                wRow("seko_kka_send_ymd") = ""
                wRow("seko_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("hktg_yohi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("hktg_zumi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff
                wRow("mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mdgt_ms"), ""))
                wRow("kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kozitn_ms"), ""))
                wRow("senro_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("senro_ms"), ""))
                wRow("tykei_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tykei_mdgt_ms"), ""))
                wRow("tykei_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tykei_kozitn_ms"), ""))
                wRow("yukoWhZytrTime") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yukoWhZytrTime"), ""))
                wRow("yukoWhGytrTime") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yukoWhGytrTime"), ""))
                wRow("nktrk_sizsk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizsk_kbn"), ""))
                wRow("smj_send_zumi_flg") = MjK1.Cmn.Const.UmuFlg.FlgOff

                'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD Start
                wRow("tkkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_ym"), ""))
                wRow("tkkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_sbt_cd"), ""))
                wRow("tkkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("tkkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_yr"), ""))
                wRow("tkkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_seizo_no"), ""))
                wRow("tkkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_ktsk_ms"), ""))
                wRow("tkkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_kzskbt_kbn"), ""))
                wRow("tkkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_tnsb_szsya_mng_value"), ""))
                wRow("ttkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_ym"), ""))
                wRow("ttkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_sbt_cd"), ""))
                wRow("ttkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("ttkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_yr"), ""))
                wRow("ttkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_seizo_no"), ""))
                wRow("ttkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_ktsk_ms"), ""))
                wRow("ttkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_kzskbt_kbn"), ""))
                wRow("ttkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ttkk_tnsb_szsya_mng_value"), ""))
                wRow("kdkk_tnsb_seizo_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_ym"), ""))
                wRow("kdkk_tnsb_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_sbt_cd"), ""))
                wRow("kdkk_tnsb_ssnsk_dnat_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_ssnsk_dnat_cd"), ""))
                wRow("kdkk_tnsb_yr") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_yr"), ""))
                wRow("kdkk_tnsb_seizo_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_seizo_no"), ""))
                wRow("kdkk_tnsb_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_ktsk_ms"), ""))
                wRow("kdkk_tnsb_kzskbt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_kzskbt_kbn"), ""))
                wRow("kdkk_tnsb_szsya_mng_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kdkk_tnsb_szsya_mng_value"), ""))
                wRow("stzk_sdnsv_menu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("stzk_sdnsv_menu_cd"), ""))
                'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD End

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，高圧用ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_KAHK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("取替票ファイル出力に失敗しました。"))
            End If

            'アップロード時にアップロード情報と併せて登録するので、ここでは設定しない
            'If NktrKnsFlag Then
            '    '抜取検査の場合、タブレット受信年月日を設定する
            '    dbr.Close()
            '    dbr = Nothing
            '    updNtknsTbltRecvYmd(wDataTable)
            'End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    ''' <summary>
    ''' 高圧他用DataTableのカラム作成
    ''' </summary>
    ''' <param name="pDataTable"></param>
    Sub initDtblTRKE_TRKEHY_KAHK(ByVal pDataTable As System.Data.DataTable)
        ' 空テーブルを作る
        pDataTable.Columns.Add("saksi_date")
        pDataTable.Columns.Add("upd_date")
        pDataTable.Columns.Add("file_sbt")
        pDataTable.Columns.Add("dig4_zgsyo_cd")
        pDataTable.Columns.Add("trkhy_hakko_nendo")
        pDataTable.Columns.Add("trkhy_kbn")
        pDataTable.Columns.Add("trkhy_no")
        pDataTable.Columns.Add("prcmg_cd")
        pDataTable.Columns.Add("trke_sbt_cd")
        pDataTable.Columns.Add("kzkis_mdgt_cd")
        pDataTable.Columns.Add("kozitn_cd")
        pDataTable.Columns.Add("skosya_cd")
        pDataTable.Columns.Add("skosya_ms")
        pDataTable.Columns.Add("kiyk_no")
        pDataTable.Columns.Add("kiyk_sbt_cd")
        pDataTable.Columns.Add("kiyk_cs_mskn")
        pDataTable.Columns.Add("kiyk_cs_ms")
        pDataTable.Columns.Add("dig14_cs_tel")
        pDataTable.Columns.Add("cs_tdhkn_add_cd")
        pDataTable.Columns.Add("cs_siku_add_cd")
        pDataTable.Columns.Add("cs_oazat_add_cd")
        pDataTable.Columns.Add("cs_azatm_add_cd")
        pDataTable.Columns.Add("tdhkn_ms")
        pDataTable.Columns.Add("siku_ms")
        pDataTable.Columns.Add("oazat_ms")
        pDataTable.Columns.Add("azatm_ms")
        pDataTable.Columns.Add("cs_addhsk_naiyo")
        pDataTable.Columns.Add("mkhy_senro_cd")
        pDataTable.Columns.Add("mkhy_kansn_no")
        pDataTable.Columns.Add("mkhy_bunk1_no")
        pDataTable.Columns.Add("mkhy_bunk2_no")
        pDataTable.Columns.Add("mkhy_bunk3_no")
        pDataTable.Columns.Add("stbit_xzahyo")
        pDataTable.Columns.Add("stbit_yzahyo")
        pDataTable.Columns.Add("iewku_stdp_xzahyo")
        pDataTable.Columns.Add("iewku_stdp_yzahyo")
        pDataTable.Columns.Add("cs_kiyk_dnryk")
        pDataTable.Columns.Add("std_kensn_dd")
        pDataTable.Columns.Add("kensn_yotei_togt_md")
        pDataTable.Columns.Add("kensn_yotei_ykgt_md")
        pDataTable.Columns.Add("kensn_yotei_yygt_md")
        pDataTable.Columns.Add("kyokyu_spot_tokti_no")
        pDataTable.Columns.Add("zyuky_ukky_kbn")
        pDataTable.Columns.Add("hist_flg")
        pDataTable.Columns.Add("brot_sett_flg")
        pDataTable.Columns.Add("knti_sbt_cd")
        pDataTable.Columns.Add("surge_yksi_siyo_umu_flg")
        pDataTable.Columns.Add("tidn_kbn")
        pDataTable.Columns.Add("tidn_riyu_cd")
        pDataTable.Columns.Add("sagyo_time_kbn_cd")
        pDataTable.Columns.Add("sagyo_start_hms")
        pDataTable.Columns.Add("sagyo_end_hms")
        pDataTable.Columns.Add("tnsb_reuse_kbn")
        pDataTable.Columns.Add("trke_tis_kbn")
        pDataTable.Columns.Add("trke_tis_abms")
        pDataTable.Columns.Add("hnsik_kykbn_cd")
        pDataTable.Columns.Add("kohu_ymd")
        pDataTable.Columns.Add("skohu_ymd")
        pDataTable.Columns.Add("hiky_yotei_ymd")
        pDataTable.Columns.Add("sksti_ymd")
        pDataTable.Columns.Add("skssti_ymd")
        pDataTable.Columns.Add("skyti_ymd")
        pDataTable.Columns.Add("syun_ymd")
        pDataTable.Columns.Add("keiki_kmaws_no")
        pDataTable.Columns.Add("keiki_sitei_no")
        pDataTable.Columns.Add("kiyk_kbn_cd")
        pDataTable.Columns.Add("keiki_kbn_cd")
        pDataTable.Columns.Add("sm_zyri_kbn_cd")
        pDataTable.Columns.Add("ftikk_flg")
        pDataTable.Columns.Add("dtkk_sougo_drkei_tekyo_szsu")
        pDataTable.Columns.Add("dtkk_sougo_drkei_trtk_szsu")
        pDataTable.Columns.Add("dtkk_max_jydrrkei_tekyo_szsu")
        pDataTable.Columns.Add("dtkk_max_jydrrkei_trtk_szsu")
        pDataTable.Columns.Add("dtkk_rsyk_tekyo_szsu")
        pDataTable.Columns.Add("dtkk_rsyk_trtk_szsu")
        pDataTable.Columns.Add("dtkk_rsmk_tekyo_szsu")
        pDataTable.Columns.Add("dtkk_rsmk_trtk_szsu")
        pDataTable.Columns.Add("dtkk_zyrt")
        pDataTable.Columns.Add("nktr_kensa_zyokyo_cd")
        pDataTable.Columns.Add("ntkns_tblt_recv_ymd")
        pDataTable.Columns.Add("ntkns_tblt_recv_tntsy_cd")
        pDataTable.Columns.Add("ntkns_zissi_ymd")
        pDataTable.Columns.Add("ntkns_zissi_tntsy_cd")
        pDataTable.Columns.Add("ntkns_kka_upld_ymd")
        pDataTable.Columns.Add("ntkns_kka_upld_tntsy_cd")
        pDataTable.Columns.Add("nktr_kensa_zyokyo_abms")
        pDataTable.Columns.Add("tkkk_keiki_id")
        pDataTable.Columns.Add("tkkk_kkinf_id")
        pDataTable.Columns.Add("tkkk_kesbt_kbn")
        pDataTable.Columns.Add("tkkk_krhsk_cd")
        pDataTable.Columns.Add("tkkk_keiki_ssnsk_so_cd")
        pDataTable.Columns.Add("tkkk_keiki_ssnsk_line_cd")
        pDataTable.Columns.Add("tkkk_kiry_dnat_cd")
        pDataTable.Columns.Add("tkkk_keiki_yr")
        pDataTable.Columns.Add("tkkk_tshsk_cd")
        pDataTable.Columns.Add("tkkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("tkkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("tkkk_ksyu_cd")
        pDataTable.Columns.Add("tkkk_taiko_kbn")
        pDataTable.Columns.Add("tkkk_knthk_cd")
        pDataTable.Columns.Add("tkkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("tkkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("tkkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("tkkk_yuko_kigen_ym")
        pDataTable.Columns.Add("tkkk_seizo_yy")
        pDataTable.Columns.Add("tkkk_zyrt")
        pDataTable.Columns.Add("tkkk_digsu")
        pDataTable.Columns.Add("tkkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("tkkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("tkkk_zytr_szsu")
        pDataTable.Columns.Add("tkkk_gytr_szsu")
        pDataTable.Columns.Add("tkkk_sougo_szsu")
        pDataTable.Columns.Add("tkkk_max_zyyo_szsu")
        pDataTable.Columns.Add("tkkk_rsyk_szsu")
        pDataTable.Columns.Add("tkkk_rsmk_szsu")
        pDataTable.Columns.Add("tkkk_hnsik_ksyu_cd")
        pDataTable.Columns.Add("no1_tkkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_tkkk_hnsik_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_vctct_ktsk_cd")
        pDataTable.Columns.Add("tkkk_hnsik_vctct_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_vctct_seizo_yy")
        pDataTable.Columns.Add("tkkk_hnsik_ct2_ktsk_cd")
        pDataTable.Columns.Add("tkkk_hnsik_ct2_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_ct2_seizo_yy")
        pDataTable.Columns.Add("tkkk_hnsik_vt1_ktsk_cd")
        pDataTable.Columns.Add("tkkk_hnsik_vt1_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_vt1_seizo_yy")
        pDataTable.Columns.Add("tkkk_hnsik_vt2_ktsk_cd")
        pDataTable.Columns.Add("tkkk_hnsik_vt2_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_vt2_seizo_yy")
        pDataTable.Columns.Add("tkkk_hyzki_seizo_no")
        pDataTable.Columns.Add("tkkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("tkkk_hnsik_gknti_no")
        pDataTable.Columns.Add("tkkk_mado_su")
        pDataTable.Columns.Add("tkkk_smkey1")
        pDataTable.Columns.Add("tkkk_smkey2")
        pDataTable.Columns.Add("tkkk_smkey_startdate")
        pDataTable.Columns.Add("tkkk_smkey_enddate")
        pDataTable.Columns.Add("tkkk_mac_address")
        pDataTable.Columns.Add("tkkk_tsn_id")
        pDataTable.Columns.Add("ttkk_keiki_id")
        pDataTable.Columns.Add("ttkk_kesbt_kbn")
        pDataTable.Columns.Add("ttkk_krhsk_cd")
        pDataTable.Columns.Add("ttkk_keiki_ssnsk_so_cd")
        pDataTable.Columns.Add("ttkk_keiki_ssnsk_line_cd")
        pDataTable.Columns.Add("ttkk_kiry_dnat_cd")
        pDataTable.Columns.Add("ttkk_keiki_yr")
        pDataTable.Columns.Add("ttkk_tshsk_cd")
        pDataTable.Columns.Add("ttkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("ttkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("ttkk_ksyu_cd")
        pDataTable.Columns.Add("ttkk_taiko_kbn")
        pDataTable.Columns.Add("ttkk_knthk_cd")
        pDataTable.Columns.Add("ttkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("ttkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("ttkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("ttkk_yuko_kigen_ym")
        pDataTable.Columns.Add("ttkk_seizo_yy")
        pDataTable.Columns.Add("ttkk_zyrt")
        pDataTable.Columns.Add("ttkk_digsu")
        pDataTable.Columns.Add("ttkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("ttkk_zytr_szsu")
        pDataTable.Columns.Add("ttkk_gytr_szsu")
        pDataTable.Columns.Add("ttkk_sougo_szsu")
        pDataTable.Columns.Add("ttkk_rsyk_szsu")
        pDataTable.Columns.Add("ttkk_rsmk_szsu")
        pDataTable.Columns.Add("ttkk_hnsik_ksyu_cd")
        pDataTable.Columns.Add("no1_ttkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_ttkk_hnsik_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_vctct_ktsk_cd")
        pDataTable.Columns.Add("ttkk_hnsik_vctct_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_vctct_seizo_yy")
        pDataTable.Columns.Add("ttkk_hnsik_ct2_ktsk_cd")
        pDataTable.Columns.Add("ttkk_hnsik_ct2_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_ct2_seizo_yy")
        pDataTable.Columns.Add("ttkk_hnsik_vt1_ktsk_cd")
        pDataTable.Columns.Add("ttkk_hnsik_vt1_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_vt1_seizo_yy")
        pDataTable.Columns.Add("ttkk_hnsik_vt2_ktsk_cd")
        pDataTable.Columns.Add("ttkk_hnsik_vt2_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_vt2_seizo_yy")
        pDataTable.Columns.Add("ttkk_hyzki_seizo_no")
        pDataTable.Columns.Add("ttkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("ttkk_hnsik_gknti_no")
        pDataTable.Columns.Add("ttkk_mado_su")
        pDataTable.Columns.Add("hzkst_tstnmt_info_tsn_id")
        pDataTable.Columns.Add("hzkst_tstnmt_info_tshsk_cd")
        pDataTable.Columns.Add("hzkst_htan_info_ksyu_cd")
        pDataTable.Columns.Add("hzkst_htan_info_kesbt_cd")
        pDataTable.Columns.Add("hzkst_htan_info_ktsk_cd")
        pDataTable.Columns.Add("hzkst_htan_info_seizo_no")
        pDataTable.Columns.Add("hzkst_htan_info_seizo_yy")
        pDataTable.Columns.Add("hzkst_ts_info_ksyu_cd")
        pDataTable.Columns.Add("hzkst_ts_info_ktsk_cd")
        pDataTable.Columns.Add("hzkst_ts_info_seizo_no")
        pDataTable.Columns.Add("hzkst_ts_info_seizo_yy")
        pDataTable.Columns.Add("kdkk_keiki_id")
        pDataTable.Columns.Add("kdkk_kesbt_kbn")
        pDataTable.Columns.Add("kdkk_krhsk_cd")
        pDataTable.Columns.Add("kdkk_keiki_ssnsk_so_cd")
        pDataTable.Columns.Add("kdkk_keiki_ssnsk_line_cd")
        pDataTable.Columns.Add("kdkk_kiry_dnat_cd")
        pDataTable.Columns.Add("kdkk_keiki_yr")
        pDataTable.Columns.Add("kdkk_tshsk_cd")
        pDataTable.Columns.Add("kdkk_keiki_ktsk_ms")
        pDataTable.Columns.Add("kdkk_keiki_ktsk_cd")
        pDataTable.Columns.Add("kdkk_taiko_kbn")
        pDataTable.Columns.Add("kdkk_knthk_cd")
        pDataTable.Columns.Add("kdkk_ts_kino_umu_flg")
        pDataTable.Columns.Add("kdkk_khk_kino_umu_flg")
        pDataTable.Columns.Add("kdkk_gaibu_output_umu_flg")
        pDataTable.Columns.Add("kdkk_yuko_kigen_ym")
        pDataTable.Columns.Add("kdkk_seizo_yy")
        pDataTable.Columns.Add("kdkk_zyrt")
        pDataTable.Columns.Add("kdkk_digsu")
        pDataTable.Columns.Add("kdkk_tnsb_szsya_cd")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_yy")
        pDataTable.Columns.Add("kdkk_zytr_szsu")
        pDataTable.Columns.Add("kdkk_gytr_szsu")
        pDataTable.Columns.Add("no1_kdkk_hnsik_seizo_no")
        pDataTable.Columns.Add("no2_kdkk_hnsik_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_vctct_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_ct2_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_vt1_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_vt2_seizo_no")
        pDataTable.Columns.Add("kdkk_hnsik_yuko_kigen_ym")
        pDataTable.Columns.Add("kdkk_hnsik_gknti_no")
        pDataTable.Columns.Add("kdkk_mado_su")
        pDataTable.Columns.Add("tdnstr_ts_dosa_kbn")
        pDataTable.Columns.Add("tdnstr_tdnt_su")
        pDataTable.Columns.Add("tdnstr_tdnstr_time")
        pDataTable.Columns.Add("hkgds_kbn")
        pDataTable.Columns.Add("hkgds_start_md")
        pDataTable.Columns.Add("hkgds_end_md")
        pDataTable.Columns.Add("no1_hkgds_tdn_hms")
        pDataTable.Columns.Add("no1_hkgds_sydan_hms")
        pDataTable.Columns.Add("no2_hkgds_tdn_hms")
        pDataTable.Columns.Add("no2_hkgds_sydan_hms")
        pDataTable.Columns.Add("no3_hkgds_tdn_hms")
        pDataTable.Columns.Add("no3_hkgds_sydan_hms")
        pDataTable.Columns.Add("no4_hkgds_tdn_hms")
        pDataTable.Columns.Add("no4_hkgds_sydan_hms")
        pDataTable.Columns.Add("no5_hkgds_tdn_hms")
        pDataTable.Columns.Add("no5_hkgds_sydan_hms")
        pDataTable.Columns.Add("no6_hkgds_tdn_hms")
        pDataTable.Columns.Add("no6_hkgds_sydan_hms")
        pDataTable.Columns.Add("no7_hkgds_tdn_hms")
        pDataTable.Columns.Add("no7_hkgds_sydan_hms")
        pDataTable.Columns.Add("no8_hkgds_tdn_hms")
        pDataTable.Columns.Add("no8_hkgds_sydan_hms")
        pDataTable.Columns.Add("no9_hkgds_tdn_hms")
        pDataTable.Columns.Add("no9_hkgds_sydan_hms")
        pDataTable.Columns.Add("no10_hkgds_tdn_hms")
        pDataTable.Columns.Add("no10_hkgds_sydan_hms")
        pDataTable.Columns.Add("kbttd_start_date")
        pDataTable.Columns.Add("kbttd_end_date")
        pDataTable.Columns.Add("hoka_kihi_kbn_cd")
        pDataTable.Columns.Add("hoka_gmn_flckr_kbn")
        pDataTable.Columns.Add("hoka_evnt_krk_kbn")
        pDataTable.Columns.Add("hoka_hoka_hyz_kbn")
        pDataTable.Columns.Add("hoka_hoka_hyz_value")
        pDataTable.Columns.Add("hksgk_hksgn_kbn")
        pDataTable.Columns.Add("hksgk_hka_dnr")
        pDataTable.Columns.Add("hksgk_attny_time")
        pDataTable.Columns.Add("hksgk_attny_kaisu")
        pDataTable.Columns.Add("hksgk_attny_kaisu_clr_time")
        pDataTable.Columns.Add("hksgr_hksgn_kbn")
        pDataTable.Columns.Add("hksgr_hka_dnr")
        pDataTable.Columns.Add("hksgr_attny_time")
        pDataTable.Columns.Add("hksgr_attny_kaisu")
        pDataTable.Columns.Add("hksgr_attny_kaisu_clr_time")
        pDataTable.Columns.Add("khkmk_keiki_kbn")
        pDataTable.Columns.Add("khkmk_mg_kbn")
        pDataTable.Columns.Add("khkmk_hnsik_kbn")
        pDataTable.Columns.Add("khkmk_hnsik_st_iti_kbn")
        pDataTable.Columns.Add("khkmk_wrms_kbn")
        pDataTable.Columns.Add("khkmk_mg_umu_flg")
        pDataTable.Columns.Add("khkmk_mg_yr")
        pDataTable.Columns.Add("tuika_kohi_umu_flg")
        pDataTable.Columns.Add("zippi_umu_flg")
        pDataTable.Columns.Add("no1_zippi_no")
        pDataTable.Columns.Add("no1_zippi_kbn")
        pDataTable.Columns.Add("no1_zippi_kngk")
        pDataTable.Columns.Add("no2_zippi_no")
        pDataTable.Columns.Add("no2_zippi_kbn")
        pDataTable.Columns.Add("no2_zippi_kngk")
        pDataTable.Columns.Add("no3_zippi_no")
        pDataTable.Columns.Add("no3_zippi_kbn")
        pDataTable.Columns.Add("no3_zippi_kngk")
        pDataTable.Columns.Add("no4_zippi_no")
        pDataTable.Columns.Add("no4_zippi_kbn")
        pDataTable.Columns.Add("no4_zippi_kngk")
        pDataTable.Columns.Add("no5_zippi_no")
        pDataTable.Columns.Add("no5_zippi_kbn")
        pDataTable.Columns.Add("no5_zippi_kngk")
        pDataTable.Columns.Add("tenp_file_mng_no")
        pDataTable.Columns.Add("ryssy_tenp_file_mng_no")
        pDataTable.Columns.Add("rrzk_naiyo")
        pDataTable.Columns.Add("prcmg_abms")
        pDataTable.Columns.Add("trke_sbt_abms")
        pDataTable.Columns.Add("brot_sett_flg_abms")
        pDataTable.Columns.Add("tkkk_kesbt_kbn_abms")
        pDataTable.Columns.Add("tkkk_krhsk_abms")
        pDataTable.Columns.Add("tkkk_tshsk_abms")
        pDataTable.Columns.Add("tkkk_knthk_abms")
        pDataTable.Columns.Add("tkkk_ts_kino_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_khk_kino_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_gaibu_output_umu_flg_abms")
        pDataTable.Columns.Add("tkkk_taiko_kbn_abms")
        pDataTable.Columns.Add("ttkk_kesbt_kbn_abms")
        pDataTable.Columns.Add("ttkk_krhsk_abms")
        pDataTable.Columns.Add("ttkk_tshsk_abms")
        pDataTable.Columns.Add("ttkk_knthk_abms")
        pDataTable.Columns.Add("ttkk_ts_kino_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_khk_kino_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_gaibu_output_umu_flg_abms")
        pDataTable.Columns.Add("ttkk_taiko_kbn_abms")
        pDataTable.Columns.Add("knti_sbt_ms")
        pDataTable.Columns.Add("surge_yksi_siyo_umu_flg_abms")
        pDataTable.Columns.Add("tidn_kbn_abms")
        pDataTable.Columns.Add("tnsb_reuse_kbn_abms")
        pDataTable.Columns.Add("khkmk_keiki_kbn_abms")
        pDataTable.Columns.Add("khkmk_mg_kbn_abms")
        pDataTable.Columns.Add("khkmk_hnsik_kbn_abms")
        pDataTable.Columns.Add("khkmk_wrms_kbn_abms")
        pDataTable.Columns.Add("khkmk_mg_umu_flg_abms")
        pDataTable.Columns.Add("no1_zippi_kbn_abms")
        pDataTable.Columns.Add("no2_zippi_kbn_abms")
        pDataTable.Columns.Add("no3_zippi_kbn_abms")
        pDataTable.Columns.Add("no4_zippi_kbn_abms")
        pDataTable.Columns.Add("no5_zippi_kbn_abms")
        pDataTable.Columns.Add("seko_kka_send_ymd")
        pDataTable.Columns.Add("seko_flg")
        pDataTable.Columns.Add("hktg_yohi_flg")
        pDataTable.Columns.Add("hktg_zumi_flg")
        pDataTable.Columns.Add("mdgt_ms")
        pDataTable.Columns.Add("kozitn_ms")
        pDataTable.Columns.Add("senro_ms")
        pDataTable.Columns.Add("tykei_mdgt_ms")
        pDataTable.Columns.Add("tykei_kozitn_ms")
        pDataTable.Columns.Add("yukoWhZytrTime")
        pDataTable.Columns.Add("yukoWhGytrTime")
        pDataTable.Columns.Add("nktrk_sizsk_kbn")
        pDataTable.Columns.Add("smj_send_zumi_flg")

        'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD Start
        pDataTable.Columns.Add("tkkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("tkkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("tkkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("tkkk_tnsb_yr")
        pDataTable.Columns.Add("tkkk_tnsb_seizo_no")
        pDataTable.Columns.Add("tkkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("tkkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("tkkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("ttkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("ttkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("ttkk_tnsb_yr")
        pDataTable.Columns.Add("ttkk_tnsb_seizo_no")
        pDataTable.Columns.Add("ttkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("ttkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("ttkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_ym")
        pDataTable.Columns.Add("kdkk_tnsb_sbt_cd")
        pDataTable.Columns.Add("kdkk_tnsb_ssnsk_dnat_cd")
        pDataTable.Columns.Add("kdkk_tnsb_yr")
        pDataTable.Columns.Add("kdkk_tnsb_seizo_no")
        pDataTable.Columns.Add("kdkk_tnsb_ktsk_ms")
        pDataTable.Columns.Add("kdkk_tnsb_kzskbt_kbn")
        pDataTable.Columns.Add("kdkk_tnsb_szsya_mng_value")
        pDataTable.Columns.Add("stzk_sdnsv_menu_cd")
        'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD End

    End Sub

    'Rev002-End

#End Region


    'Rev002-Start
#Region "取替_取替票行程（高圧他）ファイル作成処理"

    ''' <summary>
    ''' 取替_取替票行程（高圧他）ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_TRKEHY_PRC_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・取替_取替票行程（高圧他）を取得する。
            If Not dao.S017(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・取替_取替票行程（高圧他）検索でエラーが発生しました。"))
            End If

            '取替_取替票行程ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_TRKEHY_PRC_KAHK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))
                wRow("seko_tanto_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syori_ymd"), ""))
                wRow("seko_tanto_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syors_cd"), ""))
                wRow("seko_tanto_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_syors_ms"), ""))
                wRow("seko_tanto_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_ka_cd"), ""))
                wRow("seko_tanto_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_ka_ms"), ""))
                wRow("seko_tanto_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kzkis_mdgt_cd"), ""))
                wRow("seko_tanto_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kzkis_mdgt_ms"), ""))
                wRow("seko_tanto_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kozitn_cd"), ""))
                wRow("seko_tanto_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_tanto_kozitn_ms"), ""))
                wRow("seko_elder_tnknk_kknn_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnknk_kknn_flg"), ""))
                wRow("seko_elder_tnkn_kknn_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_kknn_ymd"), ""))
                wRow("seko_elder_tnkn_knsya_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_knsya_cd"), ""))
                wRow("seko_elder_tnkn_knsya_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_tnkn_knsya_ms"), ""))
                wRow("seko_elder_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_ka_cd"), ""))
                wRow("seko_elder_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_ka_ms"), ""))
                wRow("seko_elder_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kzkis_mdgt_cd"), ""))
                wRow("seko_elder_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kzkis_mdgt_ms"), ""))
                wRow("seko_elder_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kozitn_cd"), ""))
                wRow("seko_elder_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("seko_elder_kozitn_ms"), ""))
                wRow("nktrk_sizi_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syori_ymd"), ""))
                wRow("nktrk_sizi_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syors_cd"), ""))
                wRow("nktrk_sizi_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_syors_ms"), ""))
                wRow("nktrk_sizi_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_ka_cd"), ""))
                wRow("nktrk_sizi_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_ka_ms"), ""))
                wRow("nktrk_sizi_kzkis_mdgt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kzkis_mdgt_cd"), ""))
                wRow("nktrk_sizi_kzkis_mdgt_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kzkis_mdgt_ms"), ""))
                wRow("nktrk_sizi_kozitn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kozitn_cd"), ""))
                wRow("nktrk_sizi_kozitn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_sizi_kozitn_ms"), ""))
                wRow("nktrk_tanto_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syori_ymd"), ""))
                wRow("nktrk_tanto_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syors_cd"), ""))
                wRow("nktrk_tanto_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_syors_ms"), ""))
                wRow("nktrk_tanto_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ka_cd"), ""))
                wRow("nktrk_tanto_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ka_ms"), ""))
                wRow("nktrk_tanto_ces_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ces_cd"), ""))
                wRow("nktrk_tanto_ces_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_tanto_ces_ms"), ""))
                wRow("nktrk_elder_syori_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syori_ymd"), ""))
                wRow("nktrk_elder_syors_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syors_cd"), ""))
                wRow("nktrk_elder_syors_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_syors_ms"), ""))
                wRow("nktrk_elder_ka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ka_cd"), ""))
                wRow("nktrk_elder_ka_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ka_ms"), ""))
                wRow("nktrk_elder_ces_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ces_cd"), ""))
                wRow("nktrk_elder_ces_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_elder_ces_ms"), ""))
                wRow("nktrk_hnn_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nktrk_hnn_flg"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_TRKEHY_PRC_KAHK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("取替票行程（高圧他）ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_TRKEHY_PRC_KAHK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")
        wDestDt.Columns.Add("seko_tanto_syori_ymd")
        wDestDt.Columns.Add("seko_tanto_syors_cd")
        wDestDt.Columns.Add("seko_tanto_syors_ms")
        wDestDt.Columns.Add("seko_tanto_ka_cd")
        wDestDt.Columns.Add("seko_tanto_ka_ms")
        wDestDt.Columns.Add("seko_tanto_kzkis_mdgt_cd")
        wDestDt.Columns.Add("seko_tanto_kzkis_mdgt_ms")
        wDestDt.Columns.Add("seko_tanto_kozitn_cd")
        wDestDt.Columns.Add("seko_tanto_kozitn_ms")
        wDestDt.Columns.Add("seko_elder_tnknk_kknn_flg")
        wDestDt.Columns.Add("seko_elder_tnkn_kknn_ymd")
        wDestDt.Columns.Add("seko_elder_tnkn_knsya_cd")
        wDestDt.Columns.Add("seko_elder_tnkn_knsya_ms")
        wDestDt.Columns.Add("seko_elder_ka_cd")
        wDestDt.Columns.Add("seko_elder_ka_ms")
        wDestDt.Columns.Add("seko_elder_kzkis_mdgt_cd")
        wDestDt.Columns.Add("seko_elder_kzkis_mdgt_ms")
        wDestDt.Columns.Add("seko_elder_kozitn_cd")
        wDestDt.Columns.Add("seko_elder_kozitn_ms")
        wDestDt.Columns.Add("nktrk_sizi_syori_ymd")
        wDestDt.Columns.Add("nktrk_sizi_syors_cd")
        wDestDt.Columns.Add("nktrk_sizi_syors_ms")
        wDestDt.Columns.Add("nktrk_sizi_ka_cd")
        wDestDt.Columns.Add("nktrk_sizi_ka_ms")
        wDestDt.Columns.Add("nktrk_sizi_kzkis_mdgt_cd")
        wDestDt.Columns.Add("nktrk_sizi_kzkis_mdgt_ms")
        wDestDt.Columns.Add("nktrk_sizi_kozitn_cd")
        wDestDt.Columns.Add("nktrk_sizi_kozitn_ms")
        wDestDt.Columns.Add("nktrk_tanto_syori_ymd")
        wDestDt.Columns.Add("nktrk_tanto_syors_cd")
        wDestDt.Columns.Add("nktrk_tanto_syors_ms")
        wDestDt.Columns.Add("nktrk_tanto_ka_cd")
        wDestDt.Columns.Add("nktrk_tanto_ka_ms")
        wDestDt.Columns.Add("nktrk_tanto_ces_cd")
        wDestDt.Columns.Add("nktrk_tanto_ces_ms")
        wDestDt.Columns.Add("nktrk_elder_syori_ymd")
        wDestDt.Columns.Add("nktrk_elder_syors_cd")
        wDestDt.Columns.Add("nktrk_elder_syors_ms")
        wDestDt.Columns.Add("nktrk_elder_ka_cd")
        wDestDt.Columns.Add("nktrk_elder_ka_ms")
        wDestDt.Columns.Add("nktrk_elder_ces_cd")
        wDestDt.Columns.Add("nktrk_elder_ces_ms")
        wDestDt.Columns.Add("nktrk_hnn_flg")

    End Sub
#End Region
    'Rev002-End

    'Rev002-Start





#Region "自主点検（高圧他）ファイル作成処理"

    ''' <summary>
    ''' 自主点検（高圧他）ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="dao">MJ4103_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateTRKE_CHECK_SHEET_KAHK(ByVal pMtdsMngNoFolder As String, ByVal pDownldRenno As String, ByVal pKeikiDwldSbtCd As String, ByVal dao As MJ4103_DAO) As Boolean

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名

        Dim wFileSbt As String = " "
        Dim wItemCd As String = ""
        Dim wItemNaiyo As String = ""
        Dim i As Integer = 0

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            'ダウンロード管理詳細・自主点検（高圧他）を取得する。
            If Not dao.S020(dbr, pDownldRenno) Then
                Throw New Exception(String.Format("ダウンロード管理詳細・自主点検（高圧他）検索でエラーが発生しました。"))
            End If

            '自主点検（高圧他）ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblTRKE_CHECK_SHEET_KAHK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'ファイル種別の設定
            If pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.Seko) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.Seko
            ElseIf pKeikiDwldSbtCd.Equals(MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns) Then
                wFileSbt = MjK1.Cmn.Const.FileSbtCd.NktrKns
            End If


            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("dig4_zgsyo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))
                wRow("trkhy_hakko_nendo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_hakko_nendo"), ""))
                wRow("trkhy_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_kbn"), ""))
                wRow("trkhy_no") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trkhy_no"), ""))

                wRow("skosya_tnkn_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("skosya_tnkn_ymd"), ""))
                wRow("kstts_tnkn_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kstts_tnkn_ymd"), ""))
                wRow("kkchyr_snst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_snst_skosya_ryohi_kbn"), ""))
                wRow("kkchyr_snst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_snst_ntktt_ryohi_kbn"), ""))
                wRow("kkchyr_snst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_snst_kstts_ryohi_kbn"), ""))
                wRow("kkchyr_snst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_snst_tio_naiyo"), ""))
                wRow("kkchyr_snst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_snst_tio_cmp_ymd"), ""))

                wRow("kkchyr_kst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_kst_skosya_ryohi_kbn"), ""))
                wRow("kkchyr_kst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_kst_ntktt_ryohi_kbn"), ""))
                wRow("kkchyr_kst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_kst_kstts_ryohi_kbn"), ""))
                wRow("kkchyr_kst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_kst_tio_naiyo"), ""))
                wRow("kkchyr_kst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchyr_kst_tio_cmp_ymd"), ""))

                wRow("kkchano_snst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_snst_skosya_ryohi_kbn"), ""))
                wRow("kkchano_snst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_snst_ntktt_ryohi_kbn"), ""))
                wRow("kkchano_snst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_snst_kstts_ryohi_kbn"), ""))
                wRow("kkchano_snst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_snst_tio_naiyo"), ""))
                wRow("kkchano_snst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_snst_tio_cmp_ymd"), ""))

                wRow("kkchano_kst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_kst_skosya_ryohi_kbn"), ""))
                wRow("kkchano_kst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_kst_ntktt_ryohi_kbn"), ""))
                wRow("kkchano_kst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_kst_kstts_ryohi_kbn"), ""))
                wRow("kkchano_kst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_kst_tio_naiyo"), ""))
                wRow("kkchano_kst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchano_kst_tio_cmp_ymd"), ""))


                wRow("kkchsno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchsno_skosya_ryohi_kbn"), ""))
                wRow("kkchsno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchsno_ntktt_ryohi_kbn"), ""))
                wRow("kkchsno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchsno_kstts_ryohi_kbn"), ""))
                wRow("kkchsno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchsno_tio_naiyo"), ""))
                wRow("kkchsno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchsno_tio_cmp_ymd"), ""))


                wRow("kkchkkl_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkkl_skosya_ryohi_kbn"), ""))
                wRow("kkchkkl_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkkl_ntktt_ryohi_kbn"), ""))
                wRow("kkchkkl_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkkl_kstts_ryohi_kbn"), ""))
                wRow("kkchkkl_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkkl_tio_naiyo"), ""))
                wRow("kkchkkl_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkkl_tio_cmp_ymd"), ""))

                wRow("kkchkss_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkss_skosya_ryohi_kbn"), ""))
                wRow("kkchkss_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkss_ntktt_ryohi_kbn"), ""))
                wRow("kkchkss_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkss_kstts_ryohi_kbn"), ""))
                wRow("kkchkss_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkss_tio_naiyo"), ""))
                wRow("kkchkss_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkss_tio_cmp_ymd"), ""))

                wRow("kkchkib_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkib_skosya_ryohi_kbn"), ""))
                wRow("kkchkib_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkib_ntktt_ryohi_kbn"), ""))
                wRow("kkchkib_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkib_kstts_ryohi_kbn"), ""))
                wRow("kkchkib_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkib_tio_naiyo"), ""))
                wRow("kkchkib_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkib_tio_cmp_ymd"), ""))

                wRow("kkchkhd_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkhd_skosya_ryohi_kbn"), ""))
                wRow("kkchkhd_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkhd_ntktt_ryohi_kbn"), ""))
                wRow("kkchkhd_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkhd_kstts_ryohi_kbn"), ""))
                wRow("kkchkhd_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkhd_tio_naiyo"), ""))
                wRow("kkchkhd_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkhd_tio_cmp_ymd"), ""))

                wRow("kkchkbs_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkbs_skosya_ryohi_kbn"), ""))
                wRow("kkchkbs_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkbs_ntktt_ryohi_kbn"), ""))
                wRow("kkchkbs_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkbs_kstts_ryohi_kbn"), ""))
                wRow("kkchkbs_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkbs_tio_naiyo"), ""))
                wRow("kkchkbs_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchkbs_tio_cmp_ymd"), ""))

                wRow("kkckzr_snst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_snst_skosya_ryohi_kbn"), ""))
                wRow("kkckzr_snst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_snst_ntktt_ryohi_kbn"), ""))
                wRow("kkckzr_snst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_snst_kstts_ryohi_kbn"), ""))
                wRow("kkckzr_snst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_snst_tio_naiyo"), ""))
                wRow("kkckzr_snst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_snst_tio_cmp_ymd"), ""))

                wRow("kkckzr_kst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_kst_skosya_ryohi_kbn"), ""))
                wRow("kkckzr_kst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_kst_ntktt_ryohi_kbn"), ""))
                wRow("kkckzr_kst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_kst_kstts_ryohi_kbn"), ""))
                wRow("kkckzr_kst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_kst_tio_naiyo"), ""))
                wRow("kkckzr_kst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckzr_kst_tio_cmp_ymd"), ""))

                wRow("kkckano_snst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_snst_skosya_ryohi_kbn"), ""))
                wRow("kkckano_snst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_snst_ntktt_ryohi_kbn"), ""))
                wRow("kkckano_snst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_snst_kstts_ryohi_kbn"), ""))
                wRow("kkckano_snst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_snst_tio_naiyo"), ""))
                wRow("kkckano_snst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_snst_tio_cmp_ymd"), ""))

                wRow("kkckano_kst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_kst_skosya_ryohi_kbn"), ""))
                wRow("kkckano_kst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_kst_ntktt_ryohi_kbn"), ""))
                wRow("kkckano_kst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_kst_kstts_ryohi_kbn"), ""))
                wRow("kkckano_kst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_kst_tio_naiyo"), ""))
                wRow("kkckano_kst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckano_kst_tio_cmp_ymd"), ""))

                wRow("kkcksno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkcksno_skosya_ryohi_kbn"), ""))
                wRow("kkcksno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkcksno_ntktt_ryohi_kbn"), ""))
                wRow("kkcksno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkcksno_kstts_ryohi_kbn"), ""))
                wRow("kkcksno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkcksno_tio_naiyo"), ""))
                wRow("kkcksno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkcksno_tio_cmp_ymd"), ""))

                wRow("kkckdj_dosa_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_dosa_skosya_ryohi_kbn"), ""))
                wRow("kkckdj_dosa_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_dosa_ntktt_ryohi_kbn"), ""))
                wRow("kkckdj_dosa_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_dosa_kstts_ryohi_kbn"), ""))
                wRow("kkckdj_dosa_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_dosa_tio_naiyo"), ""))
                wRow("kkckdj_dosa_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_dosa_tio_cmp_ymd"), ""))

                wRow("kkckdj_settjk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_settjk_skosya_ryohi_kbn"), ""))
                wRow("kkckdj_settjk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_settjk_ntktt_ryohi_kbn"), ""))
                wRow("kkckdj_settjk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_settjk_kstts_ryohi_kbn"), ""))
                wRow("kkckdj_settjk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_settjk_tio_naiyo"), ""))
                wRow("kkckdj_settjk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_settjk_tio_cmp_ymd"), ""))

                wRow("kkckdj_skosya_tnkn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_skosya_tnkn_hms"), ""))
                wRow("kkckdj_ntktt_tnkn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_ntktt_tnkn_hms"), ""))
                wRow("kkckdj_kstts_tnkn_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_kstts_tnkn_hms"), ""))
                wRow("kkckdj_skosya_kikhyz_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_skosya_kikhyz_hms"), ""))
                wRow("kkckdj_ntktt_kikhyz_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_ntktt_kikhyz_hms"), ""))
                wRow("kkckdj_kstts_kikhyz_hms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckdj_kstts_kikhyz_hms"), ""))

                wRow("kkckks_ib_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_ib_skosya_ryohi_kbn"), ""))
                wRow("kkckks_ib_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_ib_ntktt_ryohi_kbn"), ""))
                wRow("kkckks_ib_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_ib_kstts_ryohi_kbn"), ""))
                wRow("kkckks_ib_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_ib_tio_naiyo"), ""))
                wRow("kkckks_ib_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_ib_tio_cmp_ymd"), ""))

                wRow("kkckks_hndag_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_hndag_skosya_ryohi_kbn"), ""))
                wRow("kkckks_hndag_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_hndag_ntktt_ryohi_kbn"), ""))
                wRow("kkckks_hndag_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_hndag_kstts_ryohi_kbn"), ""))
                wRow("kkckks_hndag_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_hndag_tio_naiyo"), ""))
                wRow("kkckks_hndag_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_hndag_tio_cmp_ymd"), ""))

                wRow("kkckks_rosyt_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_rosyt_skosya_ryohi_kbn"), ""))
                wRow("kkckks_rosyt_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_rosyt_ntktt_ryohi_kbn"), ""))
                wRow("kkckks_rosyt_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_rosyt_kstts_ryohi_kbn"), ""))
                wRow("kkckks_rosyt_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_rosyt_tio_naiyo"), ""))
                wRow("kkckks_rosyt_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_rosyt_tio_cmp_ymd"), ""))

                wRow("kkckks_bsst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_bsst_skosya_ryohi_kbn"), ""))
                wRow("kkckks_bsst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_bsst_ntktt_ryohi_kbn"), ""))
                wRow("kkckks_bsst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_bsst_kstts_ryohi_kbn"), ""))
                wRow("kkckks_bsst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_bsst_tio_naiyo"), ""))
                wRow("kkckks_bsst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckks_bsst_tio_cmp_ymd"), ""))

                wRow("kkckkd_tansi_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckkd_tansi_skosya_ryohi_kbn"), ""))
                wRow("kkckkd_tansi_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckkd_tansi_ntktt_ryohi_kbn"), ""))
                wRow("kkckkd_tansi_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckkd_tansi_kstts_ryohi_kbn"), ""))
                wRow("kkckkd_tansi_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckkd_tansi_tio_naiyo"), ""))
                wRow("kkckkd_tansi_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkckkd_tansi_tio_cmp_ymd"), ""))

                wRow("kkchtsn_tekyo_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchtsn_tekyo_skosya_ryohi_kbn"), ""))
                wRow("kkchtsn_tekyo_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchtsn_tekyo_ntktt_ryohi_kbn"), ""))
                wRow("kkchtsn_tekyo_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchtsn_tekyo_kstts_ryohi_kbn"), ""))
                wRow("kkchtsn_tekyo_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchtsn_tekyo_tio_naiyo"), ""))
                wRow("kkchtsn_tekyo_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kkchtsn_tekyo_tio_cmp_ymd"), ""))

                wRow("kakctkvct_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkvct_skosya_ryohi_kbn"), ""))
                wRow("kakctkvct_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkvct_ntktt_ryohi_kbn"), ""))
                wRow("kakctkvct_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkvct_kstts_ryohi_kbn"), ""))
                wRow("kakctkvct_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkvct_tio_naiyo"), ""))
                wRow("kakctkvct_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkvct_tio_cmp_ymd"), ""))

                wRow("kakctkib_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkib_skosya_ryohi_kbn"), ""))
                wRow("kakctkib_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkib_ntktt_ryohi_kbn"), ""))
                wRow("kakctkib_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkib_kstts_ryohi_kbn"), ""))
                wRow("kakctkib_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkib_tio_naiyo"), ""))
                wRow("kakctkib_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkib_tio_cmp_ymd"), ""))

                wRow("kakctkhd_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkhd_skosya_ryohi_kbn"), ""))
                wRow("kakctkhd_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkhd_ntktt_ryohi_kbn"), ""))
                wRow("kakctkhd_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkhd_kstts_ryohi_kbn"), ""))
                wRow("kakctkhd_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkhd_tio_naiyo"), ""))
                wRow("kakctkhd_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkhd_tio_cmp_ymd"), ""))

                wRow("kakctkbs_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkbs_skosya_ryohi_kbn"), ""))
                wRow("kakctkbs_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkbs_ntktt_ryohi_kbn"), ""))
                wRow("kakctkbs_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkbs_kstts_ryohi_kbn"), ""))
                wRow("kakctkbs_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkbs_tio_naiyo"), ""))
                wRow("kakctkbs_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkbs_tio_cmp_ymd"), ""))

                wRow("kakctkss_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkss_skosya_ryohi_kbn"), ""))
                wRow("kakctkss_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkss_ntktt_ryohi_kbn"), ""))
                wRow("kakctkss_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkss_kstts_ryohi_kbn"), ""))
                wRow("kakctkss_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkss_tio_naiyo"), ""))
                wRow("kakctkss_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakctkss_tio_cmp_ymd"), ""))

                wRow("kakct_szno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakct_szno_skosya_ryohi_kbn"), ""))
                wRow("kakct_szno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakct_szno_ntktt_ryohi_kbn"), ""))
                wRow("kakct_szno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakct_szno_kstts_ryohi_kbn"), ""))
                wRow("kakct_szno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakct_szno_tio_naiyo"), ""))
                wRow("kakct_szno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakct_szno_tio_cmp_ymd"), ""))

                wRow("kakcs_ksn_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_ksn_skosya_ryohi_kbn"), ""))
                wRow("kakcs_ksn_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_ksn_ntktt_ryohi_kbn"), ""))
                wRow("kakcs_ksn_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_ksn_kstts_ryohi_kbn"), ""))
                wRow("kakcs_ksn_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_ksn_tio_naiyo"), ""))
                wRow("kakcs_ksn_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_ksn_tio_cmp_ymd"), ""))

                wRow("kakcs_kns_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kns_skosya_ryohi_kbn"), ""))
                wRow("kakcs_kns_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kns_ntktt_ryohi_kbn"), ""))
                wRow("kakcs_kns_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kns_kstts_ryohi_kbn"), ""))
                wRow("kakcs_kns_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kns_tio_naiyo"), ""))
                wRow("kakcs_kns_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kns_tio_cmp_ymd"), ""))

                wRow("kakcs_skphoto_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_skphoto_skosya_ryohi_kbn"), ""))
                wRow("kakcs_skphoto_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_skphoto_ntktt_ryohi_kbn"), ""))
                wRow("kakcs_skphoto_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_skphoto_kstts_ryohi_kbn"), ""))
                wRow("kakcs_skphoto_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_skphoto_tio_naiyo"), ""))
                wRow("kakcs_skphoto_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_skphoto_tio_cmp_ymd"), ""))

                wRow("kakcs_kknn_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kknn_skosya_ryohi_kbn"), ""))
                wRow("kakcs_kknn_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kknn_ntktt_ryohi_kbn"), ""))
                wRow("kakcs_kknn_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kknn_kstts_ryohi_kbn"), ""))
                wRow("kakcs_kknn_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kknn_tio_naiyo"), ""))
                wRow("kakcs_kknn_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcs_kknn_tio_cmp_ymd"), ""))

                wRow("kakcz_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcz_skosya_ryohi_kbn"), ""))
                wRow("kakcz_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcz_ntktt_ryohi_kbn"), ""))
                wRow("kakcz_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcz_kstts_ryohi_kbn"), ""))
                wRow("kakcz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcz_tio_naiyo"), ""))
                wRow("kakcz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakcz_tio_cmp_ymd"), ""))

                wRow("kakca_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakca_skosya_ryohi_kbn"), ""))
                wRow("kakca_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakca_ntktt_ryohi_kbn"), ""))
                wRow("kakca_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakca_kstts_ryohi_kbn"), ""))
                wRow("kakca_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakca_tio_naiyo"), ""))
                wRow("kakca_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kakca_tio_cmp_ymd"), ""))

                wRow("zstk_kka_bk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zstk_kka_bk_naiyo"), ""))

                wRow("kahsk_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_1s_kstckk_flg"), ""))
                wRow("kahsk_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_p1_kstckk_flg"), ""))
                wRow("kahsk_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_p3_kstckk_flg"), ""))
                wRow("kahsk_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_3s_kstckk_flg"), ""))
                wRow("kahsk_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_3l_kstckk_flg"), ""))
                wRow("kahsk_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_p2_kstckk_flg"), ""))
                wRow("kahsk_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_skosya_1l_kstckk_flg"), ""))

                wRow("cktrmvct_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_1s_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_p1_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_p3_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_3s_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_3l_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_p2_kstckk_flg"), ""))
                wRow("cktrmvct_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_skosya_1l_kstckk_flg"), ""))

                wRow("cktrmwh_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_1s_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_p1_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_p3_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_3s_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_3l_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_p2_kstckk_flg"), ""))
                wRow("cktrmwh_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_skosya_1l_kstckk_flg"), ""))

                wRow("djfk_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_1s_kstckk_flg"), ""))
                wRow("djfk_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_p1_kstckk_flg"), ""))
                wRow("djfk_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_p3_kstckk_flg"), ""))
                wRow("djfk_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_3s_kstckk_flg"), ""))
                wRow("djfk_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_3l_kstckk_flg"), ""))
                wRow("djfk_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_p2_kstckk_flg"), ""))
                wRow("djfk_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_1l_kstckk_flg"), ""))

                wRow("djfk_skosya_dt_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_dt_kstckk_flg"), ""))
                wRow("djfk_skosya_sg_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_skosya_sg_kstckk_flg"), ""))
                wRow("sdyyht_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_skosya_p3_kstckk_flg"), ""))
                wRow("sdyyht_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_skosya_3s_kstckk_flg"), ""))
                wRow("sdyyht_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_skosya_3l_kstckk_flg"), ""))
                wRow("sdyyht_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_skosya_p2_kstckk_flg"), ""))

                wRow("jkykrk_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_1s_kstckk_flg"), ""))
                wRow("jkykrk_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_p1_kstckk_flg"), ""))
                wRow("jkykrk_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_p3_kstckk_flg"), ""))
                wRow("jkykrk_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_3s_kstckk_flg"), ""))
                wRow("jkykrk_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_3l_kstckk_flg"), ""))
                wRow("jkykrk_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_p2_kstckk_flg"), ""))
                wRow("jkykrk_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_skosya_1l_kstckk_flg"), ""))

                wRow("kahsk_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_1s_kstckk_flg"), ""))
                wRow("kahsk_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_p1_kstckk_flg"), ""))
                wRow("kahsk_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_p3_kstckk_flg"), ""))
                wRow("kahsk_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_3s_kstckk_flg"), ""))
                wRow("kahsk_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_3l_kstckk_flg"), ""))
                wRow("kahsk_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_p2_kstckk_flg"), ""))
                wRow("kahsk_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_ntktt_1l_kstckk_flg"), ""))


                wRow("cktrmvct_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_1s_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_p1_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_p3_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_3s_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_3l_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_p2_kstckk_flg"), ""))
                wRow("cktrmvct_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_ntktt_1l_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_1s_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_p1_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_p3_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_3s_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_3l_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_p2_kstckk_flg"), ""))
                wRow("cktrmwh_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_ntktt_1l_kstckk_flg"), ""))
                wRow("djfk_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_1s_kstckk_flg"), ""))
                wRow("djfk_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_p1_kstckk_flg"), ""))
                wRow("djfk_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_p3_kstckk_flg"), ""))
                wRow("djfk_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_3s_kstckk_flg"), ""))
                wRow("djfk_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_3l_kstckk_flg"), ""))
                wRow("djfk_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_p2_kstckk_flg"), ""))
                wRow("djfk_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_1l_kstckk_flg"), ""))
                wRow("djfk_ntktt_dt_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_dt_kstckk_flg"), ""))
                wRow("djfk_ntktt_sg_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_ntktt_sg_kstckk_flg"), ""))
                wRow("sdyyht_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_ntktt_p3_kstckk_flg"), ""))
                wRow("sdyyht_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_ntktt_3s_kstckk_flg"), ""))
                wRow("sdyyht_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_ntktt_3l_kstckk_flg"), ""))
                wRow("sdyyht_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_ntktt_p2_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_1s_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_p1_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_p3_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_3s_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_3l_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_p2_kstckk_flg"), ""))
                wRow("jkykrk_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_ntktt_1l_kstckk_flg"), ""))

                wRow("kahsk_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_1s_kstckk_flg"), ""))
                wRow("kahsk_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_p1_kstckk_flg"), ""))
                wRow("kahsk_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_p3_kstckk_flg"), ""))
                wRow("kahsk_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_3s_kstckk_flg"), ""))
                wRow("kahsk_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_3l_kstckk_flg"), ""))
                wRow("kahsk_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_p2_kstckk_flg"), ""))
                wRow("kahsk_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kahsk_kstts_1l_kstckk_flg"), ""))







                wRow("cktrmvct_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_1s_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_p1_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_p3_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_3s_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_3l_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_p2_kstckk_flg"), ""))
                wRow("cktrmvct_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmvct_kstts_1l_kstckk_flg"), ""))

                wRow("cktrmwh_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_1s_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_p1_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_p3_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_3s_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_3l_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_p2_kstckk_flg"), ""))
                wRow("cktrmwh_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrmwh_kstts_1l_kstckk_flg"), ""))

                wRow("djfk_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_1s_kstckk_flg"), ""))
                wRow("djfk_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_p1_kstckk_flg"), ""))
                wRow("djfk_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_p3_kstckk_flg"), ""))
                wRow("djfk_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_3s_kstckk_flg"), ""))
                wRow("djfk_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_3l_kstckk_flg"), ""))
                wRow("djfk_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_p2_kstckk_flg"), ""))
                wRow("djfk_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_1l_kstckk_flg"), ""))

                wRow("djfk_kstts_dt_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_dt_kstckk_flg"), ""))
                wRow("djfk_kstts_sg_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("djfk_kstts_sg_kstckk_flg"), ""))
                wRow("sdyyht_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_kstts_p3_kstckk_flg"), ""))
                wRow("sdyyht_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_kstts_3s_kstckk_flg"), ""))
                wRow("sdyyht_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_kstts_3l_kstckk_flg"), ""))
                wRow("sdyyht_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sdyyht_kstts_p2_kstckk_flg"), ""))



                wRow("jkykrk_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_1s_kstckk_flg"), ""))
                wRow("jkykrk_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_p1_kstckk_flg"), ""))
                wRow("jkykrk_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_p3_kstckk_flg"), ""))
                wRow("jkykrk_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_3s_kstckk_flg"), ""))
                wRow("jkykrk_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_3l_kstckk_flg"), ""))
                wRow("jkykrk_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_p2_kstckk_flg"), ""))
                wRow("jkykrk_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("jkykrk_kstts_1l_kstckk_flg"), ""))






                wRow("takcd_ttt_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_skosya_ryohi_kbn"), ""))
                wRow("takcd_ttt_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_ntktt_ryohi_kbn"), ""))
                wRow("takcd_ttt_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_kstts_ryohi_kbn"), ""))
                wRow("takcd_ttt_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_tio_naiyo"), ""))
                wRow("takcd_ttt_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_ttt_tio_cmp_ymd"), ""))
                wRow("takcd_szsu_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_skosya_ryohi_kbn"), ""))
                wRow("takcd_szsu_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_ntktt_ryohi_kbn"), ""))
                wRow("takcd_szsu_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_kstts_ryohi_kbn"), ""))
                wRow("takcd_szsu_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_tio_naiyo"), ""))
                wRow("takcd_szsu_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcd_szsu_tio_cmp_ymd"), ""))
                wRow("takck_ssda_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_skosya_ryohi_kbn"), ""))
                wRow("takck_ssda_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_ntktt_ryohi_kbn"), ""))
                wRow("takck_ssda_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_kstts_ryohi_kbn"), ""))
                wRow("takck_ssda_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_tio_naiyo"), ""))
                wRow("takck_ssda_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_ssda_tio_cmp_ymd"), ""))
                wRow("takck_zyrt_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_skosya_ryohi_kbn"), ""))
                wRow("takck_zyrt_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_ntktt_ryohi_kbn"), ""))
                wRow("takck_zyrt_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_kstts_ryohi_kbn"), ""))
                wRow("takck_zyrt_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_tio_naiyo"), ""))
                wRow("takck_zyrt_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_zyrt_tio_cmp_ymd"), ""))
                wRow("takck_yr_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_skosya_ryohi_kbn"), ""))
                wRow("takck_yr_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_ntktt_ryohi_kbn"), ""))
                wRow("takck_yr_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_kstts_ryohi_kbn"), ""))
                wRow("takck_yr_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_tio_naiyo"), ""))
                wRow("takck_yr_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_yr_tio_cmp_ymd"), ""))
                wRow("takck_sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_skosya_ryohi_kbn"), ""))
                wRow("takck_sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_ntktt_ryohi_kbn"), ""))
                wRow("takck_sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_kstts_ryohi_kbn"), ""))
                wRow("takck_sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_tio_naiyo"), ""))
                wRow("takck_sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_sm_tio_cmp_ymd"), ""))
                wRow("takck_kkano_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_skosya_ryohi_kbn"), ""))
                wRow("takck_kkano_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_ntktt_ryohi_kbn"), ""))
                wRow("takck_kkano_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_kstts_ryohi_kbn"), ""))
                wRow("takck_kkano_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_tio_naiyo"), ""))
                wRow("takck_kkano_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_kkano_tio_cmp_ymd"), ""))
                wRow("takck_hksno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_skosya_ryohi_kbn"), ""))
                wRow("takck_hksno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_ntktt_ryohi_kbn"), ""))
                wRow("takck_hksno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_kstts_ryohi_kbn"), ""))
                wRow("takck_hksno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_tio_naiyo"), ""))
                wRow("takck_hksno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hksno_tio_cmp_ymd"), ""))
                wRow("takck_hkano_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_skosya_ryohi_kbn"), ""))
                wRow("takck_hkano_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_ntktt_ryohi_kbn"), ""))
                wRow("takck_hkano_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_kstts_ryohi_kbn"), ""))
                wRow("takck_hkano_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_tio_naiyo"), ""))
                wRow("takck_hkano_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takck_hkano_tio_cmp_ymd"), ""))
                wRow("takckk_krkkd_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_skosya_ryohi_kbn"), ""))
                wRow("takckk_krkkd_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_ntktt_ryohi_kbn"), ""))
                wRow("takckk_krkkd_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_kstts_ryohi_kbn"), ""))
                wRow("takckk_krkkd_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_tio_naiyo"), ""))
                wRow("takckk_krkkd_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkkd_tio_cmp_ymd"), ""))
                wRow("takckk_krkb_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_skosya_ryohi_kbn"), ""))
                wRow("takckk_krkb_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_ntktt_ryohi_kbn"), ""))
                wRow("takckk_krkb_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_kstts_ryohi_kbn"), ""))
                wRow("takckk_krkb_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_tio_naiyo"), ""))
                wRow("takckk_krkb_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_krkb_tio_cmp_ymd"), ""))
                wRow("takcks_kkzen_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkzen_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkzen_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkzen_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_tio_naiyo"), ""))
                wRow("takcks_kkzen_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkzen_tio_cmp_ymd"), ""))
                wRow("takcks_kksl_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_skosya_ryohi_kbn"), ""))
                wRow("takcks_kksl_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kksl_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_kstts_ryohi_kbn"), ""))
                wRow("takcks_kksl_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_tio_naiyo"), ""))
                wRow("takcks_kksl_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kksl_tio_cmp_ymd"), ""))
                wRow("takcks_kkhks_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkhks_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkhks_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkhks_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_tio_naiyo"), ""))
                wRow("takcks_kkhks_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkhks_tio_cmp_ymd"), ""))
                wRow("takcks_kkgsz_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkgsz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_tio_naiyo"), ""))
                wRow("takcks_kkgsz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkgsz_tio_cmp_ymd"), ""))
                wRow("takcks_kkaki_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkaki_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkaki_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkaki_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_tio_naiyo"), ""))
                wRow("takcks_kkaki_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkaki_tio_cmp_ymd"), ""))
                wRow("takcks_kkyrs_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkyrs_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_tio_naiyo"), ""))
                wRow("takcks_kkyrs_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkyrs_tio_cmp_ymd"), ""))
                wRow("takcks_kkrst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkrst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkrst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkrst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_tio_naiyo"), ""))
                wRow("takcks_kkrst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkrst_tio_cmp_ymd"), ""))
                wRow("takcks_kkbst_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkbst_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkbst_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkbst_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_tio_naiyo"), ""))
                wRow("takcks_kkbst_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkbst_tio_cmp_ymd"), ""))
                wRow("takcks_kikib_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_skosya_ryohi_kbn"), ""))
                wRow("takcks_kikib_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kikib_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_kstts_ryohi_kbn"), ""))
                wRow("takcks_kikib_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_tio_naiyo"), ""))
                wRow("takcks_kikib_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kikib_tio_cmp_ymd"), ""))
                wRow("takcks_kkssu_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_skosya_ryohi_kbn"), ""))
                wRow("takcks_kkssu_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_ntktt_ryohi_kbn"), ""))
                wRow("takcks_kkssu_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_kstts_ryohi_kbn"), ""))
                wRow("takcks_kkssu_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_tio_naiyo"), ""))
                wRow("takcks_kkssu_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_kkssu_tio_cmp_ymd"), ""))
                wRow("takcks_hrkkl_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_skosya_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_kstts_ryohi_kbn"), ""))
                wRow("takcks_hrkkl_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_tio_naiyo"), ""))
                wRow("takcks_hrkkl_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hrkkl_tio_cmp_ymd"), ""))
                wRow("takcks_hp123_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_skosya_ryohi_kbn"), ""))
                wRow("takcks_hp123_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hp123_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_kstts_ryohi_kbn"), ""))
                wRow("takcks_hp123_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_tio_naiyo"), ""))
                wRow("takcks_hp123_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hp123_tio_cmp_ymd"), ""))
                wRow("takcks_hkaki_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_skosya_ryohi_kbn"), ""))
                wRow("takcks_hkaki_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_ntktt_ryohi_kbn"), ""))
                wRow("takcks_hkaki_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_kstts_ryohi_kbn"), ""))
                wRow("takcks_hkaki_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_tio_naiyo"), ""))
                wRow("takcks_hkaki_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcks_hkaki_tio_cmp_ymd"), ""))
                wRow("takckrk_hnrk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_skosya_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_ntktt_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_kstts_ryohi_kbn"), ""))
                wRow("takckrk_hnrk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_tio_naiyo"), ""))
                wRow("takckrk_hnrk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrk_hnrk_tio_cmp_ymd"), ""))
                wRow("takckrr_sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_skosya_ryohi_kbn"), ""))
                wRow("takckrr_sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_kstts_ryohi_kbn"), ""))
                wRow("takckrr_sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_tio_naiyo"), ""))
                wRow("takckrr_sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_sm_tio_cmp_ymd"), ""))
                wRow("takcc_led_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_skosya_ryohi_kbn"), ""))
                wRow("takcc_led_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_ntktt_ryohi_kbn"), ""))
                wRow("takcc_led_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_kstts_ryohi_kbn"), ""))
                wRow("takcc_led_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_tio_naiyo"), ""))
                wRow("takcc_led_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_led_tio_cmp_ymd"), ""))
                wRow("takcc_souck_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_skosya_ryohi_kbn"), ""))
                wRow("takcc_souck_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_ntktt_ryohi_kbn"), ""))
                wRow("takcc_souck_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_kstts_ryohi_kbn"), ""))
                wRow("takcc_souck_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_tio_naiyo"), ""))
                wRow("takcc_souck_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcc_souck_tio_cmp_ymd"), ""))
                wRow("takct_szno_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_skosya_ryohi_kbn"), ""))
                wRow("takct_szno_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_ntktt_ryohi_kbn"), ""))
                wRow("takct_szno_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_kstts_ryohi_kbn"), ""))
                wRow("takct_szno_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_tio_naiyo"), ""))
                wRow("takct_szno_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_szno_tio_cmp_ymd"), ""))
                wRow("takct_aino_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_skosya_ryohi_kbn"), ""))
                wRow("takct_aino_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_ntktt_ryohi_kbn"), ""))
                wRow("takct_aino_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_kstts_ryohi_kbn"), ""))
                wRow("takct_aino_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_tio_naiyo"), ""))
                wRow("takct_aino_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takct_aino_tio_cmp_ymd"), ""))
                wRow("takcs_photo_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_skosya_ryohi_kbn"), ""))
                wRow("takcs_photo_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_ntktt_ryohi_kbn"), ""))
                wRow("takcs_photo_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_kstts_ryohi_kbn"), ""))
                wRow("takcs_photo_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_tio_naiyo"), ""))
                wRow("takcs_photo_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_photo_tio_cmp_ymd"), ""))
                wRow("takcs_tktrk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_skosya_ryohi_kbn"), ""))
                wRow("takcs_tktrk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_ntktt_ryohi_kbn"), ""))
                wRow("takcs_tktrk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_kstts_ryohi_kbn"), ""))
                wRow("takcs_tktrk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_tio_naiyo"), ""))
                wRow("takcs_tktrk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcs_tktrk_tio_cmp_ymd"), ""))
                wRow("takcz_ntktt_seigo_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_ntktt_seigo_kknn_kbn"), ""))
                wRow("takcz_kstts_seigo_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_kstts_seigo_kknn_kbn"), ""))
                wRow("takcz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_tio_naiyo"), ""))
                wRow("takcz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takcz_tio_cmp_ymd"), ""))
                wRow("takca_ntktt_itti_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_ntktt_itti_kknn_kbn"), ""))
                wRow("takca_kstts_itti_kknn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_kstts_itti_kknn_kbn"), ""))
                wRow("takca_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_tio_naiyo"), ""))
                wRow("takca_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takca_tio_cmp_ymd"), ""))
                wRow("kizi_skosya_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_skosya_naiyo"), ""))
                wRow("kizi_ntktt_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_ntktt_naiyo"), ""))
                wRow("kizi_kstts_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kizi_kstts_naiyo"), ""))
                wRow("hrktkk_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_1s_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p1_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p3_kstckk_flg"), ""))
                wRow("hrktkk_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_3s_kstckk_flg"), ""))
                wRow("hrktkk_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_3l_kstckk_flg"), ""))
                wRow("hrktkk_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_p2_kstckk_flg"), ""))
                wRow("hrktkk_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_skosya_1l_kstckk_flg"), ""))
                wRow("p1_skosya_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_1s_kstckk_flg"), ""))
                wRow("p1_skosya_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_p1_kstckk_flg"), ""))
                wRow("p1_skosya_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_skosya_1l_kstckk_flg"), ""))
                wRow("p2_skosya_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_skosya_p2_kstckk_flg"), ""))
                wRow("p3_skosya_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_3s_kstckk_flg"), ""))
                wRow("p3_skosya_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_p3_kstckk_flg"), ""))
                wRow("p3_skosya_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_skosya_3l_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_1s_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p1_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p3_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_3s_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_3l_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_p2_kstckk_flg"), ""))
                wRow("hrktkk_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_ntktt_1l_kstckk_flg"), ""))
                wRow("p1_ntktt_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_1s_kstckk_flg"), ""))
                wRow("p1_ntktt_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_p1_kstckk_flg"), ""))
                wRow("p1_ntktt_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_ntktt_1l_kstckk_flg"), ""))
                wRow("p2_ntktt_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_ntktt_p2_kstckk_flg"), ""))
                wRow("p3_ntktt_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_3s_kstckk_flg"), ""))
                wRow("p3_ntktt_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_p3_kstckk_flg"), ""))
                wRow("p3_ntktt_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_ntktt_3l_kstckk_flg"), ""))
                wRow("hrktkk_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_1s_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p1_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p3_kstckk_flg"), ""))
                wRow("hrktkk_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_3s_kstckk_flg"), ""))
                wRow("hrktkk_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_3l_kstckk_flg"), ""))
                wRow("hrktkk_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_p2_kstckk_flg"), ""))
                wRow("hrktkk_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hrktkk_kstts_1l_kstckk_flg"), ""))
                wRow("p1_kstts_1s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_1s_kstckk_flg"), ""))
                wRow("p1_kstts_p1_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_p1_kstckk_flg"), ""))
                wRow("p1_kstts_1l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p1_kstts_1l_kstckk_flg"), ""))
                wRow("p2_kstts_p2_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p2_kstts_p2_kstckk_flg"), ""))
                wRow("p3_kstts_3s_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_3s_kstckk_flg"), ""))
                wRow("p3_kstts_p3_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_p3_kstckk_flg"), ""))
                wRow("p3_kstts_3l_kstckk_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("p3_kstts_3l_kstckk_flg"), ""))

                'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD Start
                wRow("takckrr_smhyz_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_skosya_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_kstts_ryohi_kbn"), ""))
                wRow("takckrr_smhyz_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_tio_naiyo"), ""))
                wRow("takckrr_smhyz_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_smhyz_tio_cmp_ymd"), ""))
                wRow("takckrr_d2sm_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_skosya_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_kstts_ryohi_kbn"), ""))
                wRow("takckrr_d2sm_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_tio_naiyo"), ""))
                wRow("takckrr_d2sm_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_d2sm_tio_cmp_ymd"), ""))
                wRow("takckrr_khk_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_skosya_ryohi_kbn"), ""))
                wRow("takckrr_khk_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_ntktt_ryohi_kbn"), ""))
                wRow("takckrr_khk_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_kstts_ryohi_kbn"), ""))
                wRow("takckrr_khk_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_tio_naiyo"), ""))
                wRow("takckrr_khk_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckrr_khk_tio_cmp_ymd"), ""))
                wRow("takckk_okgkb_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_skosya_ryohi_kbn"), ""))
                wRow("takckk_okgkb_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_ntktt_ryohi_kbn"), ""))
                wRow("takckk_okgkb_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_kstts_ryohi_kbn"), ""))
                wRow("takckk_okgkb_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_tio_naiyo"), ""))
                wRow("takckk_okgkb_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_okgkb_tio_cmp_ymd"), ""))
                wRow("takckk_dsipcv_skosya_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_skosya_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_ntktt_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_ntktt_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_kstts_ryohi_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_kstts_ryohi_kbn"), ""))
                wRow("takckk_dsipcv_tio_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_tio_naiyo"), ""))
                wRow("takckk_dsipcv_tio_cmp_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("takckk_dsipcv_tio_cmp_ymd"), ""))
                'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD End

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FS_TRKE_CHECK_SHEET_KAHK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("自主点検（高圧他）ファイル出力に失敗しました。"))
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If
        End Try

    End Function

    Sub initDtblTRKE_CHECK_SHEET_KAHK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("dig4_zgsyo_cd")
        wDestDt.Columns.Add("trkhy_hakko_nendo")
        wDestDt.Columns.Add("trkhy_kbn")
        wDestDt.Columns.Add("trkhy_no")

        wDestDt.Columns.Add("skosya_tnkn_ymd")
        wDestDt.Columns.Add("kstts_tnkn_ymd")
        wDestDt.Columns.Add("kkchyr_snst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_snst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_snst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_snst_tio_naiyo")
        wDestDt.Columns.Add("kkchyr_snst_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchyr_kst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_kst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_kst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchyr_kst_tio_naiyo")
        wDestDt.Columns.Add("kkchyr_kst_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchano_snst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_snst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_snst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_snst_tio_naiyo")
        wDestDt.Columns.Add("kkchano_snst_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchano_kst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_kst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_kst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchano_kst_tio_naiyo")
        wDestDt.Columns.Add("kkchano_kst_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchsno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchsno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchsno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchsno_tio_naiyo")
        wDestDt.Columns.Add("kkchsno_tio_cmp_ymd")

        wDestDt.Columns.Add("kkchkkl_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchkkl_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchkkl_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchkkl_tio_naiyo")
        wDestDt.Columns.Add("kkchkkl_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchkss_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchkss_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchkss_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchkss_tio_naiyo")
        wDestDt.Columns.Add("kkchkss_tio_cmp_ymd")
        wDestDt.Columns.Add("kkchkib_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchkib_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchkib_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchkib_tio_naiyo")
        wDestDt.Columns.Add("kkchkib_tio_cmp_ymd")

        wDestDt.Columns.Add("kkchkhd_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchkhd_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchkhd_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchkhd_tio_naiyo")
        wDestDt.Columns.Add("kkchkhd_tio_cmp_ymd")

        wDestDt.Columns.Add("kkchkbs_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchkbs_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchkbs_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchkbs_tio_naiyo")
        wDestDt.Columns.Add("kkchkbs_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckzr_snst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_snst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_snst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_snst_tio_naiyo")
        wDestDt.Columns.Add("kkckzr_snst_tio_cmp_ymd")
        wDestDt.Columns.Add("kkckzr_kst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_kst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_kst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckzr_kst_tio_naiyo")
        wDestDt.Columns.Add("kkckzr_kst_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckano_snst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_snst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_snst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_snst_tio_naiyo")
        wDestDt.Columns.Add("kkckano_snst_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckano_kst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_kst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_kst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckano_kst_tio_naiyo")
        wDestDt.Columns.Add("kkckano_kst_tio_cmp_ymd")


        wDestDt.Columns.Add("kkcksno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkcksno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkcksno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkcksno_tio_naiyo")
        wDestDt.Columns.Add("kkcksno_tio_cmp_ymd")
        wDestDt.Columns.Add("kkckdj_dosa_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_dosa_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_dosa_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_dosa_tio_naiyo")
        wDestDt.Columns.Add("kkckdj_dosa_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckdj_settjk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_settjk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_settjk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckdj_settjk_tio_naiyo")
        wDestDt.Columns.Add("kkckdj_settjk_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckdj_skosya_tnkn_hms")
        wDestDt.Columns.Add("kkckdj_ntktt_tnkn_hms")
        wDestDt.Columns.Add("kkckdj_kstts_tnkn_hms")
        wDestDt.Columns.Add("kkckdj_skosya_kikhyz_hms")
        wDestDt.Columns.Add("kkckdj_ntktt_kikhyz_hms")
        wDestDt.Columns.Add("kkckdj_kstts_kikhyz_hms")

        wDestDt.Columns.Add("kkckks_ib_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_ib_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_ib_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_ib_tio_naiyo")
        wDestDt.Columns.Add("kkckks_ib_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckks_hndag_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_hndag_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_hndag_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_hndag_tio_naiyo")
        wDestDt.Columns.Add("kkckks_hndag_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckks_rosyt_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_rosyt_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_rosyt_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_rosyt_tio_naiyo")
        wDestDt.Columns.Add("kkckks_rosyt_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckks_bsst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_bsst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_bsst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckks_bsst_tio_naiyo")
        wDestDt.Columns.Add("kkckks_bsst_tio_cmp_ymd")

        wDestDt.Columns.Add("kkckkd_tansi_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkckkd_tansi_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkckkd_tansi_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkckkd_tansi_tio_naiyo")
        wDestDt.Columns.Add("kkckkd_tansi_tio_cmp_ymd")

        wDestDt.Columns.Add("kkchtsn_tekyo_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kkchtsn_tekyo_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kkchtsn_tekyo_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kkchtsn_tekyo_tio_naiyo")
        wDestDt.Columns.Add("kkchtsn_tekyo_tio_cmp_ymd")


        wDestDt.Columns.Add("kakctkvct_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakctkvct_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakctkvct_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakctkvct_tio_naiyo")
        wDestDt.Columns.Add("kakctkvct_tio_cmp_ymd")
        wDestDt.Columns.Add("kakctkib_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakctkib_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakctkib_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakctkib_tio_naiyo")
        wDestDt.Columns.Add("kakctkib_tio_cmp_ymd")

        wDestDt.Columns.Add("kakctkhd_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakctkhd_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakctkhd_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakctkhd_tio_naiyo")
        wDestDt.Columns.Add("kakctkhd_tio_cmp_ymd")

        wDestDt.Columns.Add("kakctkbs_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakctkbs_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakctkbs_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakctkbs_tio_naiyo")
        wDestDt.Columns.Add("kakctkbs_tio_cmp_ymd")

        wDestDt.Columns.Add("kakctkss_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakctkss_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakctkss_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakctkss_tio_naiyo")
        wDestDt.Columns.Add("kakctkss_tio_cmp_ymd")

        wDestDt.Columns.Add("kakct_szno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakct_szno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakct_szno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakct_szno_tio_naiyo")
        wDestDt.Columns.Add("kakct_szno_tio_cmp_ymd")

        wDestDt.Columns.Add("kakcs_ksn_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_ksn_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_ksn_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_ksn_tio_naiyo")
        wDestDt.Columns.Add("kakcs_ksn_tio_cmp_ymd")

        wDestDt.Columns.Add("kakcs_kns_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kns_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kns_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kns_tio_naiyo")
        wDestDt.Columns.Add("kakcs_kns_tio_cmp_ymd")

        wDestDt.Columns.Add("kakcs_skphoto_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_skphoto_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_skphoto_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_skphoto_tio_naiyo")
        wDestDt.Columns.Add("kakcs_skphoto_tio_cmp_ymd")

        wDestDt.Columns.Add("kakcs_kknn_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kknn_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kknn_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakcs_kknn_tio_naiyo")
        wDestDt.Columns.Add("kakcs_kknn_tio_cmp_ymd")

        wDestDt.Columns.Add("kakcz_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakcz_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakcz_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakcz_tio_naiyo")
        wDestDt.Columns.Add("kakcz_tio_cmp_ymd")

        wDestDt.Columns.Add("kakca_skosya_ryohi_kbn")
        wDestDt.Columns.Add("kakca_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("kakca_kstts_ryohi_kbn")
        wDestDt.Columns.Add("kakca_tio_naiyo")
        wDestDt.Columns.Add("kakca_tio_cmp_ymd")

        wDestDt.Columns.Add("zstk_kka_bk_naiyo")

        wDestDt.Columns.Add("kahsk_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("kahsk_skosya_1l_kstckk_flg")

        wDestDt.Columns.Add("cktrmvct_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_skosya_1l_kstckk_flg")

        wDestDt.Columns.Add("cktrmwh_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_skosya_1l_kstckk_flg")

        wDestDt.Columns.Add("djfk_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_1l_kstckk_flg")

        wDestDt.Columns.Add("djfk_skosya_dt_kstckk_flg")
        wDestDt.Columns.Add("djfk_skosya_sg_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_skosya_p2_kstckk_flg")

        wDestDt.Columns.Add("jkykrk_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_skosya_1l_kstckk_flg")

        wDestDt.Columns.Add("kahsk_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("kahsk_ntktt_1l_kstckk_flg")



        wDestDt.Columns.Add("cktrmvct_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_dt_kstckk_flg")
        wDestDt.Columns.Add("djfk_ntktt_sg_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_ntktt_p2_kstckk_flg")

        wDestDt.Columns.Add("jkykrk_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("kahsk_kstts_1l_kstckk_flg")


        wDestDt.Columns.Add("jkykrk_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("jkykrk_kstts_1l_kstckk_flg")








        wDestDt.Columns.Add("cktrmvct_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmvct_kstts_1l_kstckk_flg")

        wDestDt.Columns.Add("cktrmwh_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("cktrmwh_kstts_1l_kstckk_flg")

        wDestDt.Columns.Add("djfk_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_1l_kstckk_flg")

        wDestDt.Columns.Add("djfk_kstts_dt_kstckk_flg")
        wDestDt.Columns.Add("djfk_kstts_sg_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("sdyyht_kstts_p2_kstckk_flg")

        '低圧分の追加

        wDestDt.Columns.Add("takcd_ttt_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcd_ttt_tio_naiyo")
        wDestDt.Columns.Add("takcd_ttt_tio_cmp_ymd")
        wDestDt.Columns.Add("takcd_szsu_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcd_szsu_tio_naiyo")
        wDestDt.Columns.Add("takcd_szsu_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_ssda_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_ssda_tio_naiyo")
        wDestDt.Columns.Add("takck_ssda_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_zyrt_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_zyrt_tio_naiyo")
        wDestDt.Columns.Add("takck_zyrt_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_yr_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_yr_tio_naiyo")
        wDestDt.Columns.Add("takck_yr_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_sm_tio_naiyo")
        wDestDt.Columns.Add("takck_sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_kkano_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_kkano_tio_naiyo")
        wDestDt.Columns.Add("takck_kkano_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_hksno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_hksno_tio_naiyo")
        wDestDt.Columns.Add("takck_hksno_tio_cmp_ymd")
        wDestDt.Columns.Add("takck_hkano_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takck_hkano_tio_naiyo")
        wDestDt.Columns.Add("takck_hkano_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_krkkd_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkkd_tio_naiyo")
        wDestDt.Columns.Add("takckk_krkkd_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_krkb_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_krkb_tio_naiyo")
        wDestDt.Columns.Add("takckk_krkb_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkzen_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkzen_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkzen_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kksl_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kksl_tio_naiyo")
        wDestDt.Columns.Add("takcks_kksl_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkhks_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkhks_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkhks_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkgsz_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkgsz_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkgsz_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkaki_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkaki_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkaki_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkyrs_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkyrs_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkyrs_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkrst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkrst_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkrst_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkbst_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkbst_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkbst_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kikib_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kikib_tio_naiyo")
        wDestDt.Columns.Add("takcks_kikib_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_kkssu_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_kkssu_tio_naiyo")
        wDestDt.Columns.Add("takcks_kkssu_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hrkkl_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hrkkl_tio_naiyo")
        wDestDt.Columns.Add("takcks_hrkkl_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hp123_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hp123_tio_naiyo")
        wDestDt.Columns.Add("takcks_hp123_tio_cmp_ymd")
        wDestDt.Columns.Add("takcks_hkaki_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcks_hkaki_tio_naiyo")
        wDestDt.Columns.Add("takcks_hkaki_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrk_hnrk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrk_hnrk_tio_naiyo")
        wDestDt.Columns.Add("takckrk_hnrk_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_sm_tio_naiyo")
        wDestDt.Columns.Add("takckrr_sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takcc_led_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcc_led_tio_naiyo")
        wDestDt.Columns.Add("takcc_led_tio_cmp_ymd")
        wDestDt.Columns.Add("takcc_souck_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcc_souck_tio_naiyo")
        wDestDt.Columns.Add("takcc_souck_tio_cmp_ymd")
        wDestDt.Columns.Add("takct_szno_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takct_szno_tio_naiyo")
        wDestDt.Columns.Add("takct_szno_tio_cmp_ymd")
        wDestDt.Columns.Add("takct_aino_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takct_aino_tio_naiyo")
        wDestDt.Columns.Add("takct_aino_tio_cmp_ymd")
        wDestDt.Columns.Add("takcs_photo_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcs_photo_tio_naiyo")
        wDestDt.Columns.Add("takcs_photo_tio_cmp_ymd")
        wDestDt.Columns.Add("takcs_tktrk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takcs_tktrk_tio_naiyo")
        wDestDt.Columns.Add("takcs_tktrk_tio_cmp_ymd")
        wDestDt.Columns.Add("takcz_ntktt_seigo_kknn_kbn")
        wDestDt.Columns.Add("takcz_kstts_seigo_kknn_kbn")
        wDestDt.Columns.Add("takcz_tio_naiyo")
        wDestDt.Columns.Add("takcz_tio_cmp_ymd")
        wDestDt.Columns.Add("takca_ntktt_itti_kknn_kbn")
        wDestDt.Columns.Add("takca_kstts_itti_kknn_kbn")
        wDestDt.Columns.Add("takca_tio_naiyo")
        wDestDt.Columns.Add("takca_tio_cmp_ymd")
        wDestDt.Columns.Add("kizi_skosya_naiyo")
        wDestDt.Columns.Add("kizi_ntktt_naiyo")
        wDestDt.Columns.Add("kizi_kstts_naiyo")
        wDestDt.Columns.Add("hrktkk_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_skosya_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_skosya_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_skosya_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_skosya_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_ntktt_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_ntktt_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_ntktt_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_3l_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("hrktkk_kstts_1l_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_1s_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_p1_kstckk_flg")
        wDestDt.Columns.Add("p1_kstts_1l_kstckk_flg")
        wDestDt.Columns.Add("p2_kstts_p2_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_3s_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_p3_kstckk_flg")
        wDestDt.Columns.Add("p3_kstts_3l_kstckk_flg")

        'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD Start
        wDestDt.Columns.Add("takckrr_smhyz_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_smhyz_tio_naiyo")
        wDestDt.Columns.Add("takckrr_smhyz_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_d2sm_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_d2sm_tio_naiyo")
        wDestDt.Columns.Add("takckrr_d2sm_tio_cmp_ymd")
        wDestDt.Columns.Add("takckrr_khk_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckrr_khk_tio_naiyo")
        wDestDt.Columns.Add("takckrr_khk_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_okgkb_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_okgkb_tio_naiyo")
        wDestDt.Columns.Add("takckk_okgkb_tio_cmp_ymd")
        wDestDt.Columns.Add("takckk_dsipcv_skosya_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_ntktt_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_kstts_ryohi_kbn")
        wDestDt.Columns.Add("takckk_dsipcv_tio_naiyo")
        wDestDt.Columns.Add("takckk_dsipcv_tio_cmp_ymd")
        'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD End

    End Sub
#End Region

    'Rev002-End


End Class
