' **************************************************************************
' 機能名称：MJ4102
' 機能概要：アップロード完了監視
' 使用方法：
' 前提条件：
' 更新履歴：2022.10.05 - EC Y.Ohnishi 新規作成
' 更新履歴：Rev002 - 2023.07.11 - EC H.Morii 高圧他テーブル対応
' 　　　　　Rev052 - 2024.10.10 - EC Y.Ohnishi 1期_工費計算時の端数処理対応
' 　　　　　Rev076 - 2025.04.18 - EC Y.Ohnishi 庫入情報送信エラーの対応
'           Rev040 - 2024.12.26 - EC Y.Ohnishi 託送情報切り替え対応
'           Rev056 - 2025.06.30 - EC Y.Hiragi  託送情報切り替え対応（高圧他）
'           Rev107 - 2026.05.13 - EC Y.Ohnishi 取付計器が第２世代の場合の統合QR読取時の指示数取得不具合への対応
' **************************************************************************
Friend Class MJ4102

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
    Public NowAppVer As String = ""                 'アプリバージョン

    'ファイル名定義
    Const C_F_TRKE_TRKEHY_TIAT As String = "TRKE_TRKEHY_TIAT.csv"        '取替_取替票（低圧）
    Const C_F_TRKE_TRKEHY_PRC As String = "TRKE_TRKEHY_PRC.csv"          '取替_取替票行程
    Const C_F_KEIKI_SPOT_TKZK As String = "KEIKI_SPOT_TKZK.csv"          '計器_地点_特記事項
    Const C_F_TRKE_TRKEHY_TUIKAKH As String = "TRKE_TRKEHY_TUIKAKH.csv"  '取替_取替票追加工費
    Const C_F_TRKE_CHECK_SHEET As String = "TRKE_CHECK_SHEET.csv"        '取替_自主点検チェックシート
    Const C_F_COM_TNPDOC_DTL As String = "COM_TNPDOC_DTL.csv"            '共通_添付書類詳細
    'Rev002-Start
    Const C_F_TRKE_TRKEHY_KAHK As String = "TRKE_TRKEHY_KAHK.csv"                   '取替_取替票（高圧他）
    Const C_F_TRKE_TRKEHY_PRC_KAHK As String = "TRKE_TRKEHY_PRC_KAHK.csv"           '取替_取替票行程（高圧他）
    Const C_F_KEIKI_SPOT_TKZK_KAHK As String = "KEIKI_SPOT_TKZK_KAHK.csv"           '計器_地点_特記事項（高圧他）
    Const C_F_TRKE_TRKEHY_TUIKAKH_KAHK As String = "TRKE_TRKEHY_TUIKAKH_KAHK.csv"   '取替_取替票追加工費（高圧他）
    Const C_F_TRKE_CHECK_SHEET_KAHK As String = "TRKE_CHECK_SHEET_KAHK.csv"         '取替_自主点検チェックシート（高圧他）
    'Rev002-End

    'ファイル・テーブル名称
    Const C_MS_TRKE_TRKEHY_TIAT As String = "取替_取替票（低圧）"
    Const C_MS_TRKE_TRKEHY_PRC As String = "取替_取替票行程"
    Const C_MS_KEIKI_SPOT_TKZK As String = "計器_地点_特記事項"
    Const C_MS_TRKE_TRKEHY_TUIKAKH As String = "取替_取替票追加工費"
    Const C_MS_TRKE_CHECK_SHEET As String = "取替_自主点検チェックシート"
    Const C_MS_COM_TNPDOC_DTL As String = "共通_添付書類詳細"
    Const C_MS_KEIKI_JUJU_RIREKI As String = "計器_授受履歴"
    'Rev002-Start
    Const C_MS_TRKE_TRKEHY_KAHK As String = "取替_取替票（高圧他）"
    Const C_MS_TRKE_TRKEHY_PRC_KAHK As String = "取替_取替票行程（高圧他）"
    Const C_MS_KEIKI_SPOT_TKZK_KAHK As String = "計器_地点_特記事項（高圧他）"
    Const C_MS_TRKE_TRKEHY_TUIKAKH_KAHK As String = "取替_取替票追加工費（高圧他）"
    Const C_MS_TRKE_CHECK_SHEET_KAHK As String = "取替_自主点検チェックシート（高圧他）"
    'Rev002-End

    'テーブル名
    Const C_T_TRKE_TRKEHY_TIAT As String = "TRKE_TRKEHY_TIAT"            '取替_取替票（低圧）
    Const C_T_KEIKI_JUJU_RIREKI As String = "KEIKI_JUJU_RIREKI"          '計器_授受履歴
    Const C_T_TRKE_TRKEHY_TUIKAKH As String = "TRKE_TRKEHY_TUIKAKH"      '取替_取替票追加工費
    Const C_T_TRKE_CHECK_SHEET As String = "TRKE_CHECK_SHEET"            '取替_自主点検チェックシート
    'Rev002-Start
    Const C_T_TRKE_TRKEHY_KAHK As String = "TRKE_TRKEHY_KAHK"                   '取替_取替票（高圧他）
    Const C_T_TRKE_TRKEHY_TUIKAKH_KAHK As String = "TRKE_TRKEHY_TUIKAKH_KAHK"   '取替_取替票追加工費（高圧他）
    Const C_T_TRKE_CHECK_SHEET_KAHK As String = "TRKE_CHECK_SHEET_KAHK"         '取替_自主点検チェックシート（高圧他）
    'Rev002-End

    'ログ　ファイル名定義
    Const C_F_COM_KIFRFLG As String = "COM_KIFRFLG.csv"                      '共通_個人情報アクセスログ
    Const C_F_COM_LOG_SYUSYU As String = "COM_LOG_SYUSYU.csv"                '共通_ログ収集
    Const C_F_COM_LOGIN_NNSY_RK As String = "COM_LOGIN_NNSY_RK.csv"          '共通_ログイン認証履歴
    Const C_F_COM_LOGIN_RIREKI_DATA As String = "COM_LOGIN_RIREKI_DATA.csv"  '共通_ログイン履歴データ
    'ログ　ファイル・テーブル名称
    Const C_MS_COM_KIFRFLG As String = "共通_個人情報アクセスログ"
    Const C_MS_COM_LOG_SYUSYU As String = "共通_ログ収集"
    Const C_MS_COM_LOGIN_NNSY_RK As String = "共通_ログイン認証履歴"
    Const C_MS_COM_LOGIN_RIREKI_DATA As String = "共通_ログイン履歴データ"

    'フォルダ名定義
    Const C_DIR_LOG As String = "log"
    Const C_DIR_TNPDOC As String = "TNPDOC"

    '桁（年月日時分秒:yyyyMMddHHmmss）
    Const C_KETA_YMDHMS As Integer = 14

    '名称定義
    Const C_MS_NKTRKNS As String = "抜取検査"
    Const C_MS_NKTRKNS_KK As String = "(" + C_MS_NKTRKNS + ")"
#End Region

#Region "メイン処理"

    ''' <summary>
    ''' メイン処理
    ''' </summary>
    ''' <remarks>
    ''' メイン処理を実行する。
    ''' </remarks>
    Public Function MainLogic() As HdSysBase.Cmn.ExitCode

        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim dao As New MJ4102_DAO(Me)       'DAO

        Dim wZgsyoCdList As List(Of String) = New List(Of String)               '処理対象事業所リスト
        Dim wZgsyoCd As String = ""
        Dim wWarningFlg As Boolean = False                                      '異常有無フラグにFalseを設定する。
        Dim wAbnormalFlg As Boolean = False

        Dim wZgsyoTopPath As String = ""                                        '持出結果パス
        Dim wExitCode As HdSysBase.Cmn.ExitCode = HdSysBase.Cmn.ExitCode.Normal '終了コード
        Dim wPidStr As String = CLng(myPID).ToString

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))
            'プロセスIDログ作成
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}を開始しました。PID={1}", KinoCD, myPID))

            '共通_事業所マスタテーブルを取得し，内部変数.事業所コードリストに設定する。
            If Not dao.S001(dbr) Then
                Throw New Exception("事業所マスタを取得できません。")
            End If
            'FETCH
            '計器定期取替システムでは、引数に指定した事業所のみを対象とする仕様に対応
            Dim wkZgStr As String = ""
            While (dbr.DataStruct.Read)
                wkZgStr = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), ""))

                If ZGSYO_CD_List.Count > 0 Then
                    For i = 0 To ZGSYO_CD_List.Count - 1
                        '事業所の先頭１文字が、起動引数に含まれているかチェックする。含まれていれば対象とする。
                        If wkZgStr.Substring(1, 1).Equals(ZGSYO_CD_List(i)) Then
                            wZgsyoCdList.Add(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), "")))
                            Exit For
                        End If
                    Next
                Else
                    wZgsyoCdList.Add(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dig4_zgsyo_cd"), "")))
                End If
            End While
            dbr.Close()

            '内部変数.異常有無フラグにFalseを設定する。
            wWarningFlg = False

            ' バッチ終了ファイルのチェックフォルダ
            Dim wEndCheckFolder As String = System.IO.Path.Combine(Me.myPath.BAT.K1TEMP, "BAT")


            While HdSysBasePg.Cmn.CanUseKino(myDB, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD)
                ' 機能が使用可能な間（使用可能な時間帯）処理を繰り返す
                ' バッチ終了ファイルをチェックする
                If System.IO.File.Exists(System.IO.Path.Combine(wEndCheckFolder, Me.KinoCD)) OrElse
                    System.IO.File.Exists(System.IO.Path.Combine(wEndCheckFolder, Me.KinoCD + "_" + Me.myUser.USER_ID)) Then
                    ' 終了結果ファイル（空ファイル）を作成して，繰り返しを抜ける
                    'System.IO.File.Create(System.IO.Path.Combine(wEndCheckFolder, "END_" + Me.KinoCD + "_" & Me.myUser.USER_ID)).Close()
                    System.IO.File.Create(System.IO.Path.Combine(wEndCheckFolder, "END_" + Me.KinoCD + "_" & Me.myPID)).Close()
                    Exit While
                End If

                'mdlStart.vb内で取得した日付を使用する。
                ''内部変数.処理年月日を設定する。
                'Me.SyoriYMD = HdSysBasePg.Cmn.GetSyoriYmd(Me.myDB)

                '内部変数.処理対象事業所リストのすべての値（事業所）に対して，以下の処理を行う。
                For Each wZgsyoCd In wZgsyoCdList
                    'アップロード用事業所TOPパス =【myPath.NAS.K1DAT】/from_tablet/【内部変数.事業所コード】
                    wZgsyoTopPath = System.IO.Path.Combine(myPath.NAS.K1DAT, MjK1.Cmn.Const.KeikiFolder.FromTablet, wZgsyoCd)

                    '対象事業所フォルダが存在しない場合，処理を行わない
                    If Not System.IO.Directory.Exists(wZgsyoTopPath) Then
                        Continue For
                    End If

                    ''事業所毎開始ログ作成
                    'HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}（アップロード） 事業所コード:{1}を開始しました。", KinoCD, wZgsyoCd))

                    '施工
                    UpldTrkmKekka(wZgsyoCd, wZgsyoTopPath, dao, wWarningFlg)

                    '抜取検査
                    UpldTrkmKekka(wZgsyoCd, wZgsyoTopPath, dao, wWarningFlg, True)

                    ''事業所毎終了ログ作成
                    'HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}（アップロード） 事業所コード:{1}を終了しました。", KinoCD, wZgsyoCd))
                Next

                '処理間隔のためwait処理を行う(ミリ秒指定)。
                System.Threading.Thread.Sleep([Const].C_SLEEP_UPLD_TRKM_INTERVAL)
                System.Windows.Forms.Application.DoEvents()
            End While

            'Finallyで設定するのだが、デバッグ時のブレークポイント用に残しておく
            wExitCode = HdSysBase.Cmn.ExitCode.Normal

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))

            'ロールバックする。
            'If Me.myDB.TranState Then
            '    MjK1.Cmn.Func.RollBack(Me.myDB)
            'End If
            myDB.Rollback()

            '内部変数.異常有無フラグにTrueを設定する。
            wAbnormalFlg = True
        Finally
            '終了処理
            If dbr IsNot Nothing Then
                dbr.Close()
                dbr = Nothing
            End If

            'zabbixメッセージ送信で監視に対応するため警告終了でのJP1対応は不要 2023/6/16
            '/*
            'If wAbnormalFlg Then
            '    '異常コード(HdSysBase.Cmn.ExitCode.Abnormal)を返却する。
            '    wExitCode = HdSysBase.Cmn.ExitCode.Abnormal
            'ElseIf wWarningFlg Then
            '    '警告コード(HdSysBase.Cmn.ExitCode.Warning)を返却する。
            '    wExitCode = HdSysBase.Cmn.ExitCode.Warning
            'Else
            '    '正常コード(HdSysBase.Cmn.ExitCode.Normal)を返却する。
            '    wExitCode = HdSysBase.Cmn.ExitCode.Normal
            'End If
            '*/
            If wAbnormalFlg Then
                '異常コード(HdSysBase.Cmn.ExitCode.Abnormal)を返却する。
                wExitCode = HdSysBase.Cmn.ExitCode.Abnormal
            Else
                '正常コード(HdSysBase.Cmn.ExitCode.Normal)を返却する。
                wExitCode = HdSysBase.Cmn.ExitCode.Normal
            End If
        End Try

        Return wExitCode

    End Function
#End Region


#Region "アップロード完了監視"
    ''' <summary>
    ''' アップロード完了監視（施工）
    ''' </summary>
    ''' <param name="pZgsyoCd">事業所コード</param>
    ''' <param name="pZgsyoTopPath">事業所パス</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pWarningFlg">異常有無フラグ</param>
    ''' <remarks></remarks>
    Private Sub UpldTrkmKekka(ByVal pZgsyoCd As String, ByVal pZgsyoTopPath As String, ByVal dao As MJ4102_DAO, ByRef pWarningFlg As Boolean)
        UpldTrkmKekka(pZgsyoCd, pZgsyoTopPath, dao, pWarningFlg, False)
    End Sub

    ''' <summary>
    ''' アップロード完了監視（施工／抜取検査）
    ''' </summary>
    ''' <param name="pZgsyoCd">事業所コード</param>
    ''' <param name="pZgsyoTopPath">事業所パス</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pWarningFlg">異常有無フラグ</param>
    ''' <param name="NktrKnsFlag">抜取検査フラグ</param>
    ''' <remarks></remarks>
    Private Sub UpldTrkmKekka(ByVal pZgsyoCd As String, ByVal pZgsyoTopPath As String, ByVal dao As MJ4102_DAO, ByRef pWarningFlg As Boolean, ByVal NktrKnsFlag As Boolean)

        Dim wUploadPath As String = ""                              'アップロードパス
        Dim wCtrPath As String = ""                                 '監視パス
        Dim wTmpPath As String = ""                                 '作業パス
        Dim wExpandPath As String = ""                              '展開パス
        Dim wBkupPath As String = ""                                'バックアップパス
        Dim wCtrFileList As List(Of String) = Nothing               '監視ファイル名一覧 
        Dim wFinishFileName As String = ""                          '完了ファイル名
        Dim wUploadFileId As String = ""                            'アップロードファイルのID部（完了ファイル名の５文字目～）
        Dim wUploadFileName As String = ""                          'アップロードファイル名
        Dim wDestFileName As String = ""                            '移動先ファイル名
        Dim wSourceFile As String = ""                              '解凍元ファイル
        Dim wDestFolder As String = ""                              '解凍先フォルダ
        Dim wFileName As String = ""                                'ファイル名
        Dim wBkupFileName As String = ""                            'バックアップファイル名

        Dim wExpUploadFileDir As String = ""                        '展開後のアップロードファイルIDフォルダ

        Dim wErrImportFlg As Boolean = False
        Dim array() As String
        Dim wUpldKttmNo As String = ""
        Dim wUpldUserId As String = ""
        Dim wUpldVersion As String = ""
        Dim wUpldDate As String = ""
        'Dim wUpldUserInfo As HdSysBase.Cmn.UserInfo = New HdSysBase.Cmn.UserInfo

        Dim wOldVerPath As String = ""
        Dim wOldVerFileName As String = ""
        Dim wKinoMs As String = ""

        Dim wKeikiDwldSbtCd As String = ""

        Dim nfiles As Integer = 0

        Dim wTempTopDir As String = ""
        Dim myPID As String = ""                       'プロセスID
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
        myPID = p.Id.ToString().PadLeft(10, "0"c)

        Dim msg As String = ""
        Dim exitCode As String = ""

        '20240115 警告フラグがクリアされずエラーが続行する件の修正 ADD Start
        pWarningFlg = False
        '20240115 警告フラグがクリアされずエラーが続行する件の修正 ADD End

        If NktrKnsFlag Then
            'アップロードパス =【myPath.NAS.K1DAT】/from_tablet/【事業所コード】/KENS
            wUploadPath = System.IO.Path.Combine(pZgsyoTopPath, MjK1.Cmn.Const.KeikiFolder.NktrKns)  'アップロードパス
            wBkupPath = System.IO.Path.Combine(myPath.NAS.K1DAT, MjK1.Cmn.Const.KeikiFolder.Bkup,
                                           MjK1.Cmn.Const.KeikiFolder.FromTablet, pZgsyoCd, MjK1.Cmn.Const.KeikiFolder.NktrKns)
            wKinoMs = "抜取検査"
            wKeikiDwldSbtCd = MjK1.Cmn.Const.KeikiDwldSbtCd.NktrKns
        Else
            'アップロードパス =【myPath.NAS.K1DAT】/from_tablet/【事業所コード】/SEKO
            wUploadPath = System.IO.Path.Combine(pZgsyoTopPath, MjK1.Cmn.Const.KeikiFolder.Seko)  'アップロードパス
            wBkupPath = System.IO.Path.Combine(myPath.NAS.K1DAT, MjK1.Cmn.Const.KeikiFolder.Bkup,
                                           MjK1.Cmn.Const.KeikiFolder.FromTablet, pZgsyoCd, MjK1.Cmn.Const.KeikiFolder.Seko)
            wKinoMs = "施工"
            wKeikiDwldSbtCd = MjK1.Cmn.Const.KeikiDwldSbtCd.Seko
        End If
        wCtrPath = System.IO.Path.Combine(wUploadPath, "ctr")                                '監視パス
        wTmpPath = System.IO.Path.Combine(wUploadPath, "tmp")                                '作業パス

        wTempTopDir = System.IO.Path.Combine(myPath.BAT.K1TEMP, myPID)                       'バッチサーバ作業パスTOP
        wExpandPath = System.IO.Path.Combine(wTempTopDir, "UPLT", pZgsyoCd, "expand")        'バッチサーバ展開パス

        '「監視パス」フォルダが無い場合はデータなしで終了する
        If Not System.IO.Directory.Exists(wCtrPath) Then
            Exit Sub
        End If
        '「作業パス」フォルダが無い場合は作成する
        If Not System.IO.Directory.Exists(wTmpPath) Then
            System.IO.Directory.CreateDirectory(wTmpPath)
        End If

        '事業所毎開始ログ作成
        HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}：{2} 事業所コード:{1}を開始しました。", KinoCD, pZgsyoCd, wKinoMs))

        '内部変数.展開パスの配下に，ファイル，および，フォルダが存在した場合，すべて削除する。
        If System.IO.Directory.Exists(wExpandPath) Then
            'ファイル削除
            For Each tempFile As String In System.IO.Directory.GetFiles(wExpandPath)
                EcOrgIO.EcOrgFileIO.DeleteFile(tempFile)
            Next

            'フォルダ削除
            For Each tempDirectory As String In System.IO.Directory.GetDirectories(wExpandPath)
                EcOrgIO.EcOrgFileIO.DeleteDirectory(tempDirectory)
            Next
        End If

        '内部変数.監視パスのフォルダに存在するファイル名が"END_*"ファイルの一覧を取得する。
        wCtrFileList = New List(Of String)

        wCtrFileList.AddRange(System.IO.Directory.GetFiles(wCtrPath, "END_*"))

        nfiles = 0
        For Each wCtrFileItm As String In wCtrFileList

            '１回の処理回数を設定し、waitさせてCPU空き時間を設ける
            nfiles = nfiles + 1
            If [Const].C_PROC_MAX_FILES < nfiles Then
                Exit For
            End If

            wErrImportFlg = False

            '内部変数を設定する。
            '完了ファイル名　        END_【端末番号】_【ユーザID】_【バージョン】_【YYYYMMDDHHMMSS】
            wFinishFileName = System.IO.Path.GetFileName(wCtrFileItm)                                 '完了ファイル名
            'アップロードファイルID      【端末番号】_【ユーザID】_【バージョン】_【YYYYMMDDHHMMSS】  END_を削除したファイル名
            wUploadFileId = wFinishFileName.Substring(4)  'アップロードファイルID
            'アップロードファイル名      【端末番号】_【ユーザID】_【バージョン】_【YYYYMMDDHHMMSS】.zip
            wUploadFileName = wUploadFileId + ".zip"                                                  'アップロードファイル名

            'アップロードファイルIDから端末番号・ユーザID・バージョン・アップロード日時を取得する
            array = Split(wUploadFileId, "_")
            If array.Count >= 4 Then
                wUpldKttmNo = array(0).ToString
                wUpldUserId = array(1).ToString
                wUpldVersion = array(2).ToString
                wUpldDate = array(3).ToString
            Else
                'ファイル名フォーマットが異なる場合は処理をスキップする
                Continue For
            End If


            '完了ファイルの存在をチェックする。
            wSourceFile = System.IO.Path.Combine(wCtrPath, wFinishFileName)
            wDestFileName = System.IO.Path.Combine(wTmpPath, wFinishFileName)

            If Not System.IO.File.Exists(wSourceFile) Then
                msg = "アップロード完了ファイルが存在しません"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, msg & ":" & pZgsyoCd & ":" & wSourceFile)
                pWarningFlg = True
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If

            'アップロードファイル名を構築する。
            wFileName = System.IO.Path.Combine(wUploadPath, wUploadFileName)


            'アプリバージョンを取得
            NowAppVer = getAppVer()
            If String.IsNullOrEmpty(NowAppVer.Trim) Then
                msg = "アップロード時アプリバージョン取得エラー"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, msg & ":" & pZgsyoCd & ":" & wUploadFileName)
                pWarningFlg = True
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If

            'バージョンが最新バージョンと異なる場合（メジャーバージョンアップ後に旧バージョンをアップロードした場合）
            If isMajorVerUp(NowAppVer, wUpldVersion) Then
                '該当バージョンのディレクトリにアップロードファイルを移動し、完了ファイルを削除する。

                If NktrKnsFlag Then
                    '旧バージョンパス =【myPath.NAS.K1DAT】/from_tablet/【事業所コード】/KENS_OLD_VER/旧AppVer
                    wOldVerPath = System.IO.Path.Combine(pZgsyoTopPath, MjK1.Cmn.Const.KeikiFolder.NktrKnsOldVer, wUpldVersion)
                Else
                    '旧バージョンパス =【myPath.NAS.K1DAT】/from_tablet/【事業所コード】/SEKO_OLD_VER/旧AppVer
                    wOldVerPath = System.IO.Path.Combine(pZgsyoTopPath, MjK1.Cmn.Const.KeikiFolder.SekoOldVer, wUpldVersion)
                End If
                If Not System.IO.Directory.Exists(wOldVerPath) Then
                    '旧バージョンディレクトリが無い場合作成する
                    System.IO.Directory.CreateDirectory(wOldVerPath)
                End If

                '旧バージョンディレクトリにアップロードファイルを移動する。
                wOldVerFileName = System.IO.Path.Combine(wOldVerPath, wUploadFileName)
                If System.IO.File.Exists(wOldVerFileName) Then
                    '旧バージョンディレクトリにアップロードファイルが既存の場合削除後、移動する。
                    System.IO.File.Delete(wOldVerFileName)
                End If
                System.IO.File.Move(wFileName, wOldVerFileName)

                '完了ファイルを削除する。
                System.IO.File.Delete(wSourceFile)

                '処理を終了し、次のファイルに進める。
                msg = "アップロードファイルは旧バージョンです"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, msg & ":" & pZgsyoCd & ":" & wUploadFileName)
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(MjK1.Cmn.Const.ZabbixMsgKbn.YokuEigyoBi, msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If


            '完了ファイルを作業パスに移動する。
            If System.IO.File.Exists(wDestFileName) Then
                System.IO.File.Delete(wDestFileName)
            End If
            System.IO.File.Move(wSourceFile, wDestFileName)


            'アップロードファイルに対し、以下の処理を行う。
            If Not System.IO.File.Exists(wFileName) Then
                msg = "アップロードファイルが存在しません"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, msg & ":" & pZgsyoCd & ":" & wUploadFileName)
                pWarningFlg = True
                wErrImportFlg = True

                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If

            'バックアップファイルが存在しない場合，バックアップする。
            wBkupFileName = System.IO.Path.Combine(wBkupPath, wUploadFileName)
            If Not System.IO.File.Exists(wBkupFileName) Then
                If Not System.IO.Directory.Exists(wBkupPath) Then
                    'バックアップディレクトリが無い場合作成する
                    System.IO.Directory.CreateDirectory(wBkupPath)
                End If
                EcOrgIO.EcOrgFileIO.CopyFile(wFileName, wBkupFileName)
            End If


            '展開用フォルダが無い場合は作成する。
            If Not System.IO.Directory.Exists(wExpandPath) Then
                System.IO.Directory.CreateDirectory(wExpandPath)
            End If

            ' 展開フォルダ（バッチサーバ）に，対象ファイルをコピーする
            Dim wExpandFileName As String = System.IO.Path.Combine(wExpandPath, wUploadFileName)
            EcOrgIO.EcOrgFileIO.CopyFile(wFileName, wExpandFileName)
            '復号化
            If Not HdSysBase.Cmn.DecryptFile(wExpandFileName, HdSysBase.Cmn.EncryptionKey, wExpandFileName) Then
                msg = "アップロードファイル復号化エラー"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, msg & ":" & pZgsyoCd & ":" & wUploadFileName)
                pWarningFlg = True
                wErrImportFlg = True
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If

            'アップロードファイルを解凍する。
            If Not HdSysBase.Cmn.ExpandArchive(System.IO.Path.Combine(wExpandPath, wUploadFileName), wExpandPath) Then
                msg = "アップロードファイル展開エラー"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, msg & ":" & pZgsyoCd & ":" & wUploadFileName)
                pWarningFlg = True
                wErrImportFlg = True
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
                Continue For
            End If

            '展開フォルダ・アップロードファイルIDのフォルダ名を取得する。
            wExpUploadFileDir = System.IO.Path.Combine(wExpandPath, wUploadFileId)

            'DBトランザクション開始
            'If Not Me.myDB.TranState() Then
            '    Me.myDB.BeginTransaction()
            'End If
            Me.myDB.BeginTransaction()

            Dim retval As Boolean
            If NktrKnsFlag Then
                'テーブル更新（抜取検査）
                retval = updateDataNktrKns(dao, wUpldKttmNo, wKeikiDwldSbtCd, wUpldUserId, wExpUploadFileDir, pWarningFlg, wErrImportFlg)
            Else
                'テーブル更新（施工）
                retval = updateDataSeko(dao, wUpldKttmNo, wKeikiDwldSbtCd, wUpldUserId, wExpUploadFileDir, pWarningFlg, wErrImportFlg)
            End If

            If retval Then
                '取込みが正常終了した場合
                myDB.Commit()

                'アップロードファイルを削除する。
                System.IO.File.Delete(System.IO.Path.Combine(wUploadPath, wUploadFileName))     '【内部変数.アップロードパス】/【内部変数.アップロードファイル名】
                'アップロード完了ファイルを削除する。
                System.IO.File.Delete(System.IO.Path.Combine(wTmpPath, wFinishFileName))        '【内部変数.作業パス】/【内部変数.完了ファイル名】
            Else
                '取込エラーの場合
                myDB.Rollback()

                msg = "アップロードファイル取込エラー"
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, msg & ":" & wUploadFileName)
                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4102)=" & exitCode)
            End If


            '展開パスのzipファイル・展開フォルダ(アップロードファイルID)を削除する。
            System.IO.File.Delete(System.IO.Path.Combine(wExpandPath, wUploadFileName))
            EcOrgIO.EcOrgFileIO.DeleteDirectory(wExpUploadFileDir, True)
            ' 作業フォルダTOPを削除する
            If System.IO.Directory.Exists(wTempTopDir) Then
                System.IO.Directory.Delete(wTempTopDir, True)
            End If
        Next

    End Sub

    '-------------------------------------------------------------------------
    ' （内部処理）アプリバージョン取得
    '-------------------------------------------------------------------------
    Private Function getAppVer() As String
        Dim wFilePath As String = System.IO.Path.Combine("N:\haiden", [Const].C_APPVER_DESTPATH, [Const].C_APPVER_SUBSYS, [Const].C_APPVER_FILE)
        Dim appver As String = ""

        Try
            'バージョンファイルを読み込む
            If Not EcOrgIO.EcOrgFileIO.ReadFileAll(wFilePath, appver) Then
                Return ""
            End If

            'バージョン番号以外を削除
            appver = System.Text.RegularExpressions.Regex.Replace(appver, "[^0-9.]", "")

            Return appver

        Catch ex As Exception
            Return ""
        End Try

    End Function

    Private Function isMajorVerUp(ByVal pAppVer As String, ByVal pUploadVer As String) As Boolean
        Dim arrayNow() As String
        Dim arrayUpload() As String

        'バージョンを分割する
        arrayNow = Split(pAppVer, ".")
        arrayUpload = Split(pUploadVer, ".")

        If arrayNow.Count >= 1 And arrayUpload.Count >= 1 Then
            '最新バージョンの１桁目がアップロードバージョンの１桁目より大きい場合
            If CInt(arrayNow(0)) > CInt(arrayUpload(0)) Then
                Return True
            End If
        End If

        Return False
    End Function

    'zabbix送信用メッセージを準備する
    Private Sub setMsgForZabbix(ByVal msg As String, ByRef zabbixMsg As String)
        'メッセージ未設定の場合のみ設定する
        If String.IsNullOrEmpty(zabbixMsg) Then
            zabbixMsg = msg
        End If
    End Sub

#End Region

#Region "テーブル更新（施工）"
    ''' <summary>
    ''' テーブル更新（施工）
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir">アップロードファイル展開フォルダ</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataSeko(ByVal dao As MJ4102_DAO,
                                    ByVal pUpldKttmNo As String,
                                    ByVal pKeikiDwldSbtCd As String,
                                    ByVal pUpldUserId As String,
                                    ByVal pExpUploadFileDir As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow


        Try
            'Rev002-Start

            '取替票（低圧）ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_TIAT
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_TIAT)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then

                '取替票（高圧他）ファイルを指定。
                wFileTableMS = C_MS_TRKE_TRKEHY_KAHK
                wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_KAHK)

                '取替票（高圧他）ファイルを読み込む
                If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                    '低圧も高圧他ファイルの両方が読み取れなかった場合はエラーで抜ける
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                    pErrImportFlg = True
                    Return False
                Else

                    '取替票（高圧他）テーブルを更新する。
                    If Not updateTRKE_TRKEHY_KAHK(dao, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId,
                                               pExpUploadFileDir, wDtblUpload, pWarningFlg, pErrImportFlg) Then
                        '高圧他（低圧契約）も高圧他も両方とも更新できなかったのでエラー
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                        pErrImportFlg = True
                        Return False
                    End If

                End If

            Else
                '取替票（低圧）テーブルを更新する。
                If Not updateTRKE_TRKEHY_TIAT(dao, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId,
                                               pExpUploadFileDir, wDtblUpload, pWarningFlg, pErrImportFlg) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pErrImportFlg = True
                    Return False
                End If

            End If

            'Rev002-End

            '共通_個人情報アクセスログに登録する。
            wFileTableMS = C_MS_COM_KIFRFLG
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_KIFRFLG)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_KIFRFLG(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pErrImportFlg = True
                    Return False
                End If
            End If


            '共通_ログ収集に登録する。
            wFileTableMS = C_MS_COM_LOG_SYUSYU
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOG_SYUSYU)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOG_SYUSYU(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pErrImportFlg = True
                    Return False
                End If
            End If

            '共通_ログイン認証履歴に登録する。
            wFileTableMS = C_MS_COM_LOGIN_NNSY_RK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOGIN_NNSY_RK)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOGIN_NNSY_RK(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pErrImportFlg = True
                    Return False
                End If
            End If

            '共通_ログイン履歴データに登録する。
            wFileTableMS = C_MS_COM_LOGIN_RIREKI_DATA
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOGIN_RIREKI_DATA)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOGIN_RIREKI_DATA(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pErrImportFlg = True
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region

#Region "テーブル更新（施工）　取替票（低圧）連係"
    ''' <summary>
    ''' テーブル更新（施工）　取替票（低圧）連係
    ''' </summary>
    ''' <param name="pRow">ファイル取替票（低圧）レコード</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="pDtblUpload">取替票（低圧）データ</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataSekoOthers(ByVal pRow As DataRow,
                                    ByVal dao As MJ4102_DAO,
                                    ByVal pExpUploadFileDir As String,
                                    ByVal pDtblUpload As DataTable,
                                    ByVal k1 As String,
                                    ByVal k2 As String,
                                    ByVal k3 As String,
                                    ByVal k4 As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow

        'k1:dig4_zgsyo_cd      事業所コード（4桁）
        'k2:trkhy_hakko_nendo  取替票発行年度
        'k3:trkhy_kbn          取替票区分
        'k4:trkhy_no           取替票番号

        Try
            '計器_授受履歴テーブルを更新する。（取替票（低圧）データを連係する）
            wFileTableMS = C_MS_KEIKI_JUJU_RIREKI
            If Not updateKEIKI_JUJU_RIREKI(dao, pDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '取替票行程情報ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_PRC
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_PRC)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '取替票行程情報テーブルを更新する。
            If Not updateTRKE_TRKEHY_PRC(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '取替_取替票追加工費ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_TUIKAKH
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_TUIKAKH)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '取替_取替票追加工費テーブルを更新する。（工事費合計の更新も行う）
            If Not updateTRKE_TRKEHY_TUIKAKH(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '自主点検チェックシートファイルを読み込む。
            wFileTableMS = C_MS_TRKE_CHECK_SHEET
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_CHECK_SHEET)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '自主点検チェックシートテーブルを更新する。
            If Not updateTRKE_CHECK_SHEET(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '共通_添付書類ファイルを読み込む。
            wFileTableMS = C_MS_COM_TNPDOC_DTL
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_COM_TNPDOC_DTL)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '配下の画像を共通_添付書類に登録する。
            If Not updateCOM_TNPDOC_DTL(pRow, dao, wDtblUpload, pExpUploadFileDir, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "err:" & ex.Message & " " & wFileTableMS & "ファイル取込エラー： " & wFileName)
            Return False
        End Try

    End Function
#End Region

#Region "テーブル更新（施工）　取替票（高圧他）連係"
    'Rev002-Start
    ''' <summary>
    ''' テーブル更新（施工）　取替票（高圧他）連係
    ''' </summary>
    ''' <param name="pRow">ファイル取替票（高圧他）レコード</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="pDtblUpload">取替票（高圧他）データ</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataSekoOthersKahk(ByVal pRow As DataRow,
                                    ByVal dao As MJ4102_DAO,
                                    ByVal pExpUploadFileDir As String,
                                    ByVal pDtblUpload As DataTable,
                                    ByVal k1 As String,
                                    ByVal k2 As String,
                                    ByVal k3 As String,
                                    ByVal k4 As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow

        'k1:dig4_zgsyo_cd      事業所コード（4桁）
        'k2:trkhy_hakko_nendo  取替票発行年度
        'k3:trkhy_kbn          取替票区分
        'k4:trkhy_no           取替票番号

        Try
            '取替票行程情報ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_PRC_KAHK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_PRC_KAHK)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '取替票行程情報テーブルを更新する。
            If Not updateTRKE_TRKEHY_PRC_KAHK(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '取替_取替票追加工費ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_TUIKAKH_KAHK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_TUIKAKH_KAHK)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '取替_取替票追加工費（高圧他）テーブルを更新する。（工事費合計の更新も行う）
            If Not updateTRKE_TRKEHY_TUIKAKH_KAHK(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '自主点検チェックシート（高圧他）ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_CHECK_SHEET_KAHK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_CHECK_SHEET_KAHK)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '自主点検チェックシート（高圧他）テーブルを更新する。
            If Not updateTRKE_CHECK_SHEET_KAHK(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If


            '共通_添付書類ファイルを読み込む。
            wFileTableMS = C_MS_COM_TNPDOC_DTL
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_COM_TNPDOC_DTL)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pErrImportFlg = True
                Return False
            End If
            '配下の画像を共通_添付書類に登録する。
            If Not updateCOM_TNPDOC_DTL_KAHK(pRow, dao, wDtblUpload, pExpUploadFileDir, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pErrImportFlg = True
                Return False
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "err:" & ex.Message & " " & wFileTableMS & "ファイル取込エラー： " & wFileName)
            Return False
        End Try

    End Function
    'Rev002-End
#End Region


#Region "取替_取替票（低圧）テーブル"
    ''' <summary>
    ''' 取替_取替票（低圧）テーブル
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_TIAT(ByVal dao As MJ4102_DAO,
                                            ByVal pUpldKttmNo As String,
                                            ByVal pKeikiDwldSbtCd As String,
                                            ByVal pUpldUserId As String,
                                            ByVal pExpUploadFileDir As String,
                                            ByVal pDataTable As DataTable, ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票（低圧）テーブルが更新可能かチェックする。
            Dim k1 As String = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            Dim k2 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            Dim k3 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            Dim k4 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))
            Dim jouken1 As String = "dig4_zgsyo_cd = " & EcOrgIO.EcOrgString.GetSqlText(k1)
            Dim jouken2 As String = "trkhy_hakko_nendo = " & EcOrgIO.EcOrgString.GetSqlText(k2)
            Dim jouken3 As String = "trkhy_kbn = " & EcOrgIO.EcOrgString.GetSqlText(k3)
            Dim jouken4 As String = "trkhy_no = " & EcOrgIO.EcOrgString.GetSqlText(k4)
            Dim s1 As String = ""
            Dim s2 As String = ""
            Dim s3 As String = ""
            Dim s4 As String = ""
            Dim field1 As String = "prcmg_cd"
            Dim field2 As String = "kesbt_bnrui_cd" '計器種別分類コード←検定種別コード"knti_sbt_cd"
            Dim field3 As String = "upd_date"
            Dim field4 As String = "syun_ymd"
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる ADD Start
            Dim s5 As String = ""
            Dim s6 As String = "" 'SMJ計器情報との撤去計器指示数比較で6桁指示数が必要なため
            Dim field5 As String = "trkhy_hinpt_flg"
            Dim field6 As String = "tkkk_digsu" 'SMJ計器情報との撤去計器指示数比較で6桁指示数が必要なため
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる ADD End
            Dim retval As Boolean = False
            Dim nextProc As String = ""
            Dim isTaidou As Boolean = False
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる MOD Start
            'retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_TIAT,
            '                                     field1, field2, field3, field4, "", "", "", "", "", "",
            '                                     s1, s2, s3, s4, "", "", "", "", "", "",
            '                                     jouken1, jouken2, jouken3, jouken4)
            retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_TIAT,
                                                 field1, field2, field3, field4, field5, field6, "", "", "", "",
                                                 s1, s2, s3, s4, s5, s6, "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる MOD End
            If Not retval Then
                'テーブルの取替票項目値の取得エラーの場合、エラー終了
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票検索エラーです:" + k1 + "-" + k2 + k3 + k4)
                Return False
            End If

            '行程が他異動かチェックする(行程の先頭1文字が'F'であること)
            If s1(0) = MjK1.Cmn.Const.ProcMngCd.FTaIdou(0) Then
                isTaidou = True
            End If

            '取替票の存在チェックと行程チェック
            If String.IsNullOrEmpty(s1.Trim) Then
                'テーブルに該当取替票キーのレコードがない場合、エラー終了
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票がありません:" + k1 + "-" + k2 + k3 + k4)
                Return False
            ElseIf Not s1.Equals(MjK1.Cmn.Const.ProcMngCd.CJusin) And isTaIdou = False Then
                'テーブルの行程管理コードが「C040:取替票受信済」「Fxxx(F010:他異動中止,F020:他異動中止確認)」以外の場合、スキップする
                'Rev076 庫入情報送信エラー対応　行程不正の場合、「手入力フラグがONでない場合」のみ取込エラーとする MOD Start
                'HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票の行程が不正です:" + k1 + "-" + k2 + k3 + k4 + "[" + s1 + "]")
                'pWarningFlg = True
                If Not s5.Trim.Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票の行程が不正です:" + k1 + "-" + k2 + k3 + k4 + "[" + s1 + "]")
                    pWarningFlg = True
                End If
                'Rev076 庫入情報送信エラー対応　行程不正の場合、「手入力フラグがONでない場合」のみ取込エラーとする MOD End
                Continue For
            End If

            '次の行程管理コードを設定する。他異動以外の場合、竣工年月日・更新年月日のチェックを行う。
            If isTaidou Then
                '他異動中止の場合、変更しない。（SQLのupdate文に今と同じ行程を指定することで対応）
                nextProc = s1
            Else
                '取替票受信済の場合
                'CT付かどうか判定する（計器種別分類コードで判定。検定種別コードはやめる）
                If s2.Trim.Equals(MjK1.Cmn.Const.KesbtBriCd.CT) Then
                    'CT付    D010：竣工結果確認待
                    nextProc = MjK1.Cmn.Const.ProcMngCd.DShunkoKekka
                Else
                    'CTなし　D020：竣工報告待
                    nextProc = MjK1.Cmn.Const.ProcMngCd.DShunkoHokoku
                End If

                'タブレットで施工結果送信時にチェックしているので、ここでは行わない　2023/6/6
                ''テーブルの竣工年月日が登録済みならエラー終了
                'If Not String.IsNullOrEmpty(s4.Trim) Then
                '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "竣工年月日が登録済みです:" + k1 + "-" + k2 + k3 + k4)
                '    Return False
                'End If

                'アップロードの排他確認は不要。多重ダウンロードが可能のため。
                ''排他確認　テーブルの更新年月日が更新済みならエラー終了
                'Dim fileUpdDate As String = getDateStrWithKeta(wRow("upd_date").ToString, C_KETA_YMDHMS)
                'Dim dbUpdDate As String = getDateStrWithKeta(s3, C_KETA_YMDHMS)
                'If fileUpdDate <= dbUpdDate Then
                '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "更新年月日が更新済みです:" + k1 + "-" + k2 + k3 + k4)
                '    Return False
                'End If
            End If



            retval = False
            Dim wttkk_keiki_id = CStr(Me.myDB.ConvDbNull(wRow("ttkk_keiki_id"), "")).Trim
            'Dim jouken5 As String = "ttkk_keiki_id = " & EcOrgIO.EcOrgString.GetSqlText(wttkk_keiki_id)
            'retval = myDB.CheckExistRec(C_T_TRKE_TRKEHY_TIAT, jouken2, jouken5)

            'B-ST-0196 取付計器の多重チェックは不要となったため処理を削除 2023/5/26
            ''同一取付計器の多重チェック（同一年度の取替票を検索）
            'Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
            'If Not dao.S003(dbr, wRow) Then
            '    Return False
            'End If
            'If dbr.DataStruct.HasRows Then
            '    retval = True
            'End If
            'dbr.Close()
            '
            'If retval Then
            '    '多重している場合（自取替票は未更新なので、除外する取替票は無い。）
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器IDは使用済みです:" + k1 + "-" + k2 + k3 + k4 + "(" + wttkk_keiki_id + ")")
            '    Return False
            'End If

            'Rev076 庫入情報送信結果が「廃滅登録重複」の場合、SMJ撤去計器チェックを行う ADD Start
            Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
            Try
                Dim smj_send_zumi_flg = CStr(Me.myDB.ConvDbNull(wRow("smj_send_zumi_flg"), "")).Trim
                Dim wValue1 = MjK1.Cmn.Func.GetZyousu(myDB, [Const].C_ZYOUSUCD_HAIMETU_DUP_VALUE, SyoriYMD)
                If smj_send_zumi_flg.Equals(wValue1) Then
                    '廃滅登録重複の場合、撤去計器IDを条件にSMJから計器情報を取得する
                    Dim wtkkk_keiki_id = CStr(Me.myDB.ConvDbNull(wRow("tkkk_keiki_id"), "")).Trim
                    Dim wkstr1 As String = ""
                    Dim wkstr2 As String = ""

                    If Not dao.S007(dbr, wtkkk_keiki_id) Then
                        'DB検索エラー
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック DBエラーです:" + k1 + "-" + k2 + k3 + k4)
                        Return False
                    End If

                    If Not dbr.DataStruct.HasRows Then
                        '撤去計器情報なし
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 撤去計器情報なし:" + k1 + "-" + k2 + k3 + k4)
                        Return False
                    End If

                    While (dbr.DataStruct.Read)
                        '撤去計器_指示数_順潮流, 撤去計器_指示数_逆潮流
                        wkstr1 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krir_zytr_yuko_wh"), "")).Trim 'SMJ計器情報　庫入指示数(6桁)
                        wkstr2 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krir_gytr_yuko_wh"), "")).Trim 'SMJ計器情報　庫入指示数(6桁)
                        If String.IsNullOrEmpty(wkstr1) Or String.IsNullOrEmpty(wkstr2) Then
                            '指示数がない
                            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 撤去計器 指示数なし:" + k1 + "-" + k2 + k3 + k4)
                            Return False
                        End If

                        '指示数がタブレット取得値と異なる場合エラー
                        Dim tkkk_zytr_szsu = CStr(Me.myDB.ConvDbNull(wRow("tkkk_zytr_szsu"), "")).Trim '撤去計器_指示数_順潮流
                        Dim tkkk_gytr_szsu = CStr(Me.myDB.ConvDbNull(wRow("tkkk_gytr_szsu"), "")).Trim '撤去計器_指示数_逆潮流
                        '庫入情報送信のxmlフォーマットで撤去計器の指示数は6桁。SMJはそのままテーブルに保存しているので6桁合わせする Start
                        tkkk_zytr_szsu = Edit_Szsu(tkkk_zytr_szsu, s6)
                        tkkk_gytr_szsu = Edit_Szsu(tkkk_gytr_szsu, s6)
                        '庫入情報送信のxmlフォーマットで撤去計器の指示数は6桁。SMJはそのままテーブルに保存しているので7桁合わせする End
                        If Not (tkkk_zytr_szsu.Equals(wkstr1) And tkkk_gytr_szsu.Equals(wkstr2)) Then
                            '指示数がタブレット取得値と異なる場合エラー
                            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 撤去計器 指示数不一致:" + k1 + "-" + k2 + k3 + k4)
                            Return False
                        End If

                        '撤去計器_庫出日が設定されており、撤去計器_庫入日が空の場合
                        wkstr1 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krds_ymd"), "")).Trim
                        wkstr2 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krir_ymd"), "")).Trim
                        If Not String.IsNullOrEmpty(wkstr1) And String.IsNullOrEmpty(wkstr2) Then
                            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 庫出日あり・庫入日なし:" + k1 + "-" + k2 + k3 + k4)
                            Return False
                        End If

                        '撤去_庫出日 >= 撤去_庫入日 の場合エラーとする
                        If wkstr1 >= wkstr2 Then
                            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 庫出日>=庫入日:" + k1 + "-" + k2 + k3 + k4)
                            Return False
                        End If

                        '１件のみ処理
                        Exit While
                    End While
                End If
            Catch ex As Exception
            Finally
                dbr.Close()
            End Try
            'Rev076 庫入情報送信結果が「廃滅登録重複」の場合、SMJ撤去計器チェックを行う ADD End


            'Rev107 取付計器が第２世代の場合、統合QR読取時の指示数取得不具合への対応でMAMデータと突合する ADD Start
            Try
                If IsTnsbQR2G(CStr(Me.myDB.ConvDbNull(wRow("ttkk_tnsb_sbt_cd"), ""))) Then
                    '第２世代
                    Dim sisn_1 As String = ""
                    Dim sisn_2 As String = ""
                    Dim wSisn_1 As String = ""
                    Dim wSisn_2 As String = ""
                    Dim wTtkkZytrSzsu As String = CStr(Me.myDB.ConvDbNull(wRow("ttkk_zytr_szsu"), "")) '取付計器_指示数_順潮流
                    Dim wTtkkGytrSzsu As String = CStr(Me.myDB.ConvDbNull(wRow("ttkk_gytr_szsu"), "")) '取付計器_指示数_逆潮流

                    'MAMから取付計器IDに対応する指針を検索する。
                    If Not dao.S010(dbr, wttkk_keiki_id) Then
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器(第２世代) MAM計器設定(公開)参照エラー:" + k1 + "-" + k2 + k3 + k4)
                        Return False
                    End If
                    If dbr.DataStruct.HasRows Then
                        While (dbr.DataStruct.Read)
                            sisn_1 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sisn_1"), ""))
                            sisn_2 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sisn_2"), ""))
                            '１件のみ処理
                            Exit While
                        End While
                    Else
                        '対象取付計器の指示数がMAM計器設定（公開）から取得できない場合、エラー
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器(第２世代) 取付計器の情報なし:" + k1 + "-" + k2 + k3 + k4 + " 計器ID:" + wttkk_keiki_id)
                        Return False
                    End If
                    dbr.Close()

                    If CStr(Me.myDB.ConvDbNull(wRow("ttkk_digsu"), "")).Equals([Const].C_DIGSU_5) Then
                        '第２世代かつ計器桁数「5」の場合のMAMの指針は『整数5桁.小数1桁』の6桁構成。よって、小数点削除＋右0埋し、計器業務の7桁に合わせる
                        wSisn_1 = sisn_1.Replace(".", "").PadRight([Const].C_SZSU_KETA, "0"c)
                        wSisn_2 = sisn_2.Replace(".", "").PadRight([Const].C_SZSU_KETA, "0"c)
                    Else
                        '第２世代かつ計器桁数「4」の場合のMAMの指針は『整数4桁.小数2桁』の6桁構成。よって、小数点削除＋左0埋し、計器業務の7桁に合わせる
                        wSisn_1 = sisn_1.Replace(".", "").PadLeft([Const].C_SZSU_KETA, "0"c)
                        wSisn_2 = sisn_2.Replace(".", "").PadLeft([Const].C_SZSU_KETA, "0"c)
                    End If

                    'タブレット統合QR読取結果の取付計器_指示数_順潮流と比較
                    If Not wTtkkZytrSzsu.Equals(wSisn_1) Then
                        '差異があるためエラー終了
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器(第２世代) 順潮流差異:" + k1 + "-" + k2 + k3 + k4 + "  QR:" + wTtkkZytrSzsu + "  指針1:" + wSisn_1)
                        Return False
                    End If
                    'タブレット統合QR読取結果の取付計器_指示数_逆潮流と比較
                    If Not wTtkkGytrSzsu.Equals(wSisn_2) Then
                        '差異があるためエラー終了
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器(第２世代) 逆潮流差異:" + k1 + "-" + k2 + k3 + k4 + "  QR:" + wTtkkGytrSzsu + "  指針2:" + wSisn_2)
                        Return False
                    End If
                End If

            Catch ex As Exception
            Finally
                dbr.Close()
            End Try
            'Rev107 取付計器が第２世代の場合、統合QR読取時の指示数取得不具合への対応でMAMデータと突合する ADD End


            '取替_取替票（低圧）テーブルを更新する。
            If Not dao.U001(wRow, nextProc, SyoriYMD) Then
                Return False
            End If


            '更新した取替_取替票（低圧）に連係するテーブルを取替票単位で更新する。
            '（更新対象の取替票についてのみ、連係するテーブルも更新する。）
            If Not updateDataSekoOthers(wRow, dao, pExpUploadFileDir, pDataTable, k1, k2, k3, k4, pWarningFlg, pErrImportFlg) Then
                Return False
            End If


            'ダウンロード取替票管理の持出ファイル状態コードを「5:アップロード済」に更新する。
            If Not dao.U011(wRow, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId) Then
                Return False
            End If

            'Rev040 2024/12/25 物理分割 計器マスタ連係対応 ADD Start
            'WEBAPI実行管理テーブルへの登録を行う。
            'IF023_取付完了通知
            If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF023) Then
                Return False
            End If
            'IF024_撤去完了通知
            If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF024) Then
                Return False
            End If
            'IF025_設備情報更新（端子部再用の場合）
            Dim tnsb_reuse_kbn As String = CStr(Me.myDB.ConvDbNull(wRow("tnsb_reuse_kbn"), ""))
            Dim kbn As Integer = 0
            If tnsb_reuse_kbn.Equals([Const].C_TNSB_REUSE) Then
                '端子部再用の場合
                If IsTnsbQR2G(CStr(Me.myDB.ConvDbNull(wRow("tkkk_tnsb_sbt_cd"), ""))) Then
                    '第２世代
                    kbn = [Const].C_KEIKI_MASTER_RENKEI_IF025G2
                Else
                    '第１世代
                    kbn = [Const].C_KEIKI_MASTER_RENKEI_IF025G1
                End If
                If Not dao.I011(wRow, kbn) Then
                    Return False
                End If
            End If
            'Rev040 2024/12/25 物理分割 計器マスタ連係対応 ADD End
        Next

        '１件でもWarningがあればFalseを返却する
        If pWarningFlg Then
            Return False
        End If

        Return True

    End Function

    'Rev076 庫入情報送信エラー対応　MjSysTableBaseのメソッドを取り込む ADD Start
    Private Function Edit_Szsu(ByVal pSzsu As String, ByVal pDigsu As String) As String
        Dim wSzsu As String = ""
        Try
            If pSzsu.Length = [Const].C_SZSU_KETA Then
                '7桁の場合、桁数に従う
                Select Case CInt(pDigsu.ToString)
                    Case 4
                        ' 先頭1桁を除外し、編集
                        wSzsu = pSzsu.Substring(1)
                    Case 5
                        ' 末尾1桁を除外し、編集
                        wSzsu = pSzsu.Substring(0, [Const].C_SZSU_KETA - 1)
                    Case Else
                        wSzsu = pSzsu
                End Select
            ElseIf pSzsu.Length = [Const].C_SZSU_KETA - 1 Then
                '6桁の場合はそのまま返却
            Else
            End If

            Return wSzsu
        Catch ex As Exception
            Return wSzsu
        End Try

    End Function
    'Rev076 庫入情報送信エラー対応　MjSysTableBaseのメソッドを取り込む ADD End
    
    'Rev040 2024/12/25 物理分割 計器マスタ連係対応 ADD Start
    ''' <summary>
    ''' 端子部QR第２世代判定
    ''' </summary>
    ''' <param name="pTnsbSbtCd">端子部_種別コード</param>
    ''' <returns>True:第２世代，False:第１世代</returns>
    Public Shared Function IsTnsbQR2G(ByVal pTnsbSbtCd As String) As Boolean

        If String.IsNullOrEmpty(pTnsbSbtCd) Then
            Return False
        End If

        ' 端子部種別コードが第２世代かチェックする
        If pTnsbSbtCd.Trim.Equals([Const].C_TNSB_SBT_CD_2G) Then
            Return True
        End If

        Return False
    End Function
    'Rev040 2024/12/25 物理分割 計器マスタ連係対応 ADD End
#End Region

#Region "取替_取替票行程テーブル(取替票単位)"
    ''' <summary>
    ''' 取替_取替票行程テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票行程)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_PRC(ByVal dao As MJ4102_DAO, ByRef pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""
        '取替票キー条件設定
        jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(k1)
        jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(k2)
        jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(k3)
        jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(k4)

        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim s4 As String = ""
        Dim field1 As String = "no1_zippi_no"
        Dim field2 As String = "no2_zippi_no"
        Dim field3 As String = "no3_zippi_no"
        Dim field4 As String = "zippi_umu_flg"
        Dim wtable As String = C_T_TRKE_TRKEHY_TIAT
        Dim retval As Boolean = False

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If


        '実費確認フラグを追加する
        pDataTable.Columns.Add("zippi_umu_flg")
        pDataTable.Columns.Add("no1_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no2_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no3_skrepo_zippi_kknn_flg")

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替票（低圧）の実費ありの場合、竣工報告_実費確認フラグ1～3を設定する
            retval = myDB.GetRecValueStringMulti(wtable,
                                                 field1, field2, field3, field4, "", "", "", "", "", "",
                                                 s1, s2, s3, s4, "", "", "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            wRow("zippi_umu_flg") = s4.Trim
            'Rev002.1 MOD 20241210 フラグオフ時のクリアが抜けていたので追加
            If retval Then
                If Not String.IsNullOrEmpty(s1.Trim) Then
                    wRow("no1_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no1_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s2.Trim) Then
                    wRow("no2_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no2_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s3.Trim) Then
                    wRow("no3_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no3_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
            End If

            '取替_取替票行程テーブルを更新する。
            If Not dao.U002(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region

#Region "計器_授受履歴テーブル(取替票単位)"
    ''' <summary>
    ''' 計器_授受履歴テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替票（低圧）)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateKEIKI_JUJU_RIREKI(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean
        Dim isExist As Boolean = False
        Dim tblname As String = C_T_KEIKI_JUJU_RIREKI
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)
        Dim jouken1 As String = ""

        Dim wDataRows() As DataRow = Nothing
        Dim clmname As String = ""

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows

            '撤去計器　既存チェック
            clmname = [Const].C_CLM_TKKK_KEIKI_ID
            '計器_授受履歴テーブルに登録する（新規追加のみ 2023/5/31）。
            If Not dao.I001(wRow, MjK1.Cmn.Const.zyzy_kbn.Tekyo, MjK1.Cmn.Const.keiki_zyzy_status_cd.TekyoMtds, clmname) Then
                Return False
            End If

            '取付計器　既存チェック
            clmname = [Const].C_CLM_TTKK_KEIKI_ID
            '計器_授受履歴テーブルに登録する（新規追加のみ 2023/5/31）。
            If Not dao.I001(wRow, MjK1.Cmn.Const.zyzy_kbn.Trtk, MjK1.Cmn.Const.keiki_zyzy_status_cd.Trtk, clmname) Then
                Return False
            End If

        Next

        Return True

    End Function
#End Region

#Region "取替票（低圧）条件設定"
    ''' <summary>
    ''' 取替票（低圧）条件設定
    ''' </summary>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    Public Function getTrkhyWhereStr(ByVal k1 As String,
                                      ByVal k2 As String,
                                      ByVal k3 As String,
                                      ByVal k4 As String) As String
        Dim s As String = "1=1 "
        s = s & " AND dig4_zgsyo_cd = " & EcOrgIO.EcOrgString.GetSqlText(k1)
        s = s & " AND trkhy_hakko_nendo = " & EcOrgIO.EcOrgString.GetSqlText(k2)
        s = s & " AND trkhy_kbn = " & EcOrgIO.EcOrgString.GetSqlText(k3)
        s = s & " AND trkhy_no = " & EcOrgIO.EcOrgString.GetSqlText(k4)
        Return s
    End Function


    ''' <summary>
    ''' 取替票番号キー取得
    ''' </summary>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    Public Function getTrkhyNoKey(ByVal k1 As String,
                                  ByVal k2 As String,
                                  ByVal k3 As String,
                                  ByVal k4 As String) As String
        Dim s As String = ""
        s = s & Right("0000" & k1, 4)
        s = s & "-" & Right("00" & k2, 2)
        s = s & Right(" " & k3, 1)
        s = s & Right("00000" & k4, 5)
        Return s
    End Function
#End Region


#Region "取替_取替票追加工費テーブル(取替票単位)"
    ''' <summary>
    ''' 取替_取替票追加工費テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票追加工費)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_TUIKAKH(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean
        Dim isExist As Boolean = False
        Dim tblname As String = C_T_TRKE_TRKEHY_TUIKAKH
        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        Dim trtkTanka As Decimal = 0
        Dim tekyoTanka As Decimal = 0
        Dim wtrtkTanka As String = ""
        Dim wtekyoTanka As String = ""
        Dim trtkKojiHi As Decimal = 0         '追加工費テーブルの取付工事費
        Dim tekyoKojiHi As Decimal = 0        '追加工費テーブルの撤去工事費
        Dim trkhy_trtkKojiHi As Decimal = 0   '取替票単位の追加工費取付の合計
        Dim trkhy_tekyoKojiHi As Decimal = 0  '取替票単位の追加工費撤去の合計
        Dim TUIKAKH_umu_flg As String = MjK1.Cmn.Const.UmuFlg.FlgOff


        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim s4 As String = ""
        Dim s5 As String = ""
        Dim syun_ymd As String = ""
        Dim wrms_kbn As String = ""
        Dim zippi1 As String = ""
        Dim zippi2 As String = ""
        Dim zippi3 As String = ""
        Dim field1 As String = "khkmk_trtk_kozih_kngk"
        Dim field2 As String = "khkmk_tekyo_kozih_kngk"
        Dim field3 As String = "kozitn_cd"
        Dim field4 As String = "kzkmk_hnsik_trtk_kozih_kngk"
        Dim field5 As String = "kzkmk_hnsik_tekyo_kozih_kngk"
        Dim field6 As String = "syun_ymd"
        Dim field7 As String = "khkmk_wrms_kbn"
        Dim field8 As String = "no1_zippi_kngk"
        Dim field9 As String = "no2_zippi_kngk"
        Dim field10 As String = "no3_zippi_kngk"
        Dim wtable As String = C_T_TRKE_TRKEHY_TIAT
        Dim retval As Boolean
        Dim trkhy_trtk_kozih_kngk As Decimal = 0        '取替票　取付工事費
        Dim trkhy_tekyo_kozih_kngk As Decimal = 0       '取替票　撤去工事費
        Dim kzkmk_hnsik_trtk_kozih_kngk As Decimal = 0  '工事項目_変成器_取付工事費 
        Dim kzkmk_hnsik_tekyo_kozih_kngk As Decimal = 0 '工事項目_変成器_撤去工事費
        Dim zippi1_kngk As Decimal = 0                  '実費1
        Dim zippi2_kngk As Decimal = 0                  '実費2
        Dim zippi3_kngk As Decimal = 0                  '実費3
        Dim gokei_kngk As Decimal = 0                   '取替票　工事費合計
        Dim tuikaKohiKoziKbn As String = ""
        Dim wTuikakhSytkMstKbn As String = ""
        Dim jouken5 As String = ""

        '取替票キー条件設定
        jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(k1)
        jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(k2)
        jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(k3)
        jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(k4)


        '取替票の取付工事費・撤去工事費・竣工年月日・割増区分を取得する。
        retval = myDB.GetRecValueStringMulti(wtable,
                                                 field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                                                 s1, s2, s3, s4, s5, syun_ymd, wrms_kbn, zippi1, zippi2, zippi3,
                                                 jouken1, jouken2, jouken3, jouken4)
        trkhy_trtk_kozih_kngk = MyCDec(s1)   '取替票　取付工事費
        trkhy_tekyo_kozih_kngk = MyCDec(s2)  '取替票　撤去工事費
        kzkmk_hnsik_trtk_kozih_kngk = MyCDec(s4)  '工事項目_変成器_取付工事費 
        kzkmk_hnsik_tekyo_kozih_kngk = MyCDec(s5) '工事項目_変成器_撤去工事費
        zippi1_kngk = MyCDec(zippi1)              '実費1
        zippi2_kngk = MyCDec(zippi2)              '実費2
        zippi3_kngk = MyCDec(zippi3)              '実費3

        '割増率取得
        Dim wrms_ritu As String = dao.getKoryoWrmsRitu(wrms_kbn, syun_ymd)


        '取得したデータ件数 > 0 の場合、追加工費を登録し、追加工費金額を算出する。
        If (pDataTable.Rows.Count > 0) Then
            wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,tuika_kohi_row_no")

            '取替票番号キーは確定済み。k1,k2,k3,k4

            '取替_取替票追加工費　既存チェック
            isExist = myDB.CheckExistRec(tblname, jouken1, jouken2, jouken3, jouken4)
            '既存の場合、取替_取替票追加工費テーブルのレコードを削除する。
            If isExist Then
                '該当取替票番号のレコード一式を削除する。
                If Not dao.D001(k1, k2, k3, k4) Then
                    Return False
                End If
            End If

            '取替_取替票追加工費の処理を行う。
            For Each wRow As DataRow In wDataRows
                '追加工費区分を取得する。
                tuikaKohiKoziKbn = CStr(Me.myDB.ConvDbNull(wRow("tuika_kohi_kozi_kbn"), ""))
                tuikaKohiKoziKbn = tuikaKohiKoziKbn.PadLeft([Const].C_KORYO_CD_KETA, "0"c)

                '追加工費_取得マスタ区分取得
                jouken5 = "koryo_zunit_cd = " & EcOrgIO.EcOrgString.GetSqlText(tuikaKohiKoziKbn)
                wTuikakhSytkMstKbn = myDB.GetRecValueString("MST_USE_KANO_TUIKAKH", "tuikakh_sytk_mst_kbn", jouken5)

                '追加工費_工事区分の単価を取得する
                getTankaT(dao, tuikaKohiKoziKbn, wTuikakhSytkMstKbn, syun_ymd, wtrtkTanka, wtekyoTanka)
                trtkTanka = MyCDec(wtrtkTanka)
                tekyoTanka = MyCDec(wtekyoTanka)

                '工量マスタ・材料ユニットで処理を分ける
                If wTuikakhSytkMstKbn.Equals(MjK1.Cmn.Const.TuikakhSytkMstKbn.ZunitMst) Then
                    '材料ユニットの場合（割増なし）
                    '追加工費_取付工事費を算出する（単価*数量）
                    trtkKojiHi = trtkTanka * MyCDec(wRow("tuika_kohi_trtk_suryo").ToString)
                    '追加工費_撤去工事費を算出する（単価*数量）
                    tekyoKojiHi = tekyoTanka * MyCDec(wRow("tuika_kohi_tekyo_suryo").ToString)
                    'マイナス値となる
                    tekyoKojiHi = tekyoKojiHi * -1
                Else
                    '工量ユニットマスタの場合（割増あり）
                    '追加工費_取付工事費を算出する（単価*数量）
                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD Start
                    'trtkKojiHi = trtkTanka * MyCDec(wRow("tuika_kohi_trtk_suryo").ToString)
                    'trtkKojiHi = dao.CalcWrmsKngk(trtkKojiHi, wrms_ritu)
                    trtkKojiHi = dao.CalcKozihKngk(wtrtkTanka, MyCDec(wRow("tuika_kohi_trtk_suryo").ToString), wrms_ritu)
                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 算出順も変更する MOD End

                    '追加工費_撤去工事費を算出する（単価*数量）
                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD Start
                    'tekyoKojiHi = tekyoTanka * MyCDec(wRow("tuika_kohi_tekyo_suryo").ToString)
                    'tekyoKojiHi = dao.CalcWrmsKngk(tekyoKojiHi, wrms_ritu)
                    tekyoKojiHi = dao.CalcKozihKngk(wtekyoTanka, MyCDec(wRow("tuika_kohi_tekyo_suryo").ToString), wrms_ritu)
                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 算出順も変更する MOD End
                End If

                '取替_取替票追加工費を追加する
                If Not dao.I003(wRow, trtkKojiHi.ToString, tekyoKojiHi.ToString) Then
                    Return False
                End If

                '取替票単位の追加工費の合計を算出する
                trkhy_trtkKojiHi = trkhy_trtkKojiHi + trtkKojiHi     '取替票　追加工費取付の合計
                trkhy_tekyoKojiHi = trkhy_tekyoKojiHi + tekyoKojiHi  '取替票　追加工費撤去の合計

                TUIKAKH_umu_flg = MjK1.Cmn.Const.UmuFlg.FlgOn
            Next
        End If


        '******************************************************
        '計算結果を取替票番号キーのデータに更新する。
        '******************************************************

        '取替票　工事費合計
        gokei_kngk = 0
        gokei_kngk = gokei_kngk + trkhy_trtk_kozih_kngk + trkhy_tekyo_kozih_kngk             '取付工費 + 撤去工費
        gokei_kngk = gokei_kngk + kzkmk_hnsik_trtk_kozih_kngk + kzkmk_hnsik_tekyo_kozih_kngk '変成器_取付工費 + 変成器_撤去工費
        gokei_kngk = gokei_kngk + trkhy_trtkKojiHi + trkhy_tekyoKojiHi                       '追加工費_取付工費 + 追加工費_撤去工費
        gokei_kngk = gokei_kngk + zippi1_kngk + zippi2_kngk + zippi3_kngk                    '実費1,2,3

        '直営工事の場合、金額に０円を設定する。
        Dim wChokueiKztnCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, SyoriYMD)
        If s3.Trim.Equals(wChokueiKztnCd) Then
            '０円設定
            '工費項目_取付工事費 
            '工費項目_撤去工事費 
            '工事項目_変成器_取付工事費 
            '工事項目_変成器_撤去工事費 
            '→取替票の処理時に設定済

            '追加工費_取付工事費 
            trkhy_trtkKojiHi = 0
            '追加工費_撤去工事費
            trkhy_tekyoKojiHi = 0
            '工事費合計 
            gokei_kngk = 0
        End If


        '取替票　追加工費取付・追加工費撤去および工事費合計を更新する
        If Not dao.U010(k1, k2, k3, k4, TUIKAKH_umu_flg, trkhy_trtkKojiHi.ToString, trkhy_tekyoKojiHi.ToString, gokei_kngk.ToString) Then
            Return False
        End If


        Return True
    End Function


    ''' <summary>
    ''' 単価の取得（追加工費）
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="tuikaKohiKoziKbn">追加工費区分(桁合わせ済み)</param>
    ''' <param name="tuikakhSytkMstKbn">追加工費_取得マスタ区分</param>
    ''' <param name="syunYmd">竣工年月日</param>
    ''' <param name="trtkTanka">取付単価</param>
    ''' <param name="tekyoTanka">撤去単価</param>
    ''' <remarks></remarks>
    Function getTankaT(ByVal dao As MJ4102_DAO, ByVal tuikaKohiKoziKbn As String, ByVal tuikakhSytkMstKbn As String,
                       ByVal syunYmd As String, ByRef trtkTanka As String, ByRef tekyoTanka As String) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        Dim koryo_kohi_sbt_cd As String = ""
        Dim trtk_tanka_kngkD As Decimal = 0
        Dim trtk_koryoD As Decimal = 0
        Dim tekyo_tanka_kngkD As Decimal = 0
        Dim tekyo_koryoD As Decimal = 0

        trtkTanka = "0"
        tekyoTanka = "0"

        '追加工費の単価情報を取得する。
        If Not dao.S004(dbr, tuikakhSytkMstKbn, tuikaKohiKoziKbn, syunYmd, "", "") Then
            Return False
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)

                koryo_kohi_sbt_cd = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_kohi_sbt_cd"), ""))

                '工量マスタ　工量工費種別コードを判定
                If koryo_kohi_sbt_cd.Equals("1") Then
                    '追加工費　工量マスタ
                    trtk_tanka_kngkD = MyCDec(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_tanka_kngk"), "")))
                    tekyo_tanka_kngkD = MyCDec(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_tanka_kngk"), "")))
                    trtk_koryoD = MyCDec(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_koryo"), "")))
                    tekyo_koryoD = MyCDec(CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_koryo"), "")))

                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD Start
                    'trtkTanka = CStr(trtk_tanka_kngkD * trtk_koryoD)
                    'tekyoTanka = CStr(tekyo_tanka_kngkD * tekyo_koryoD)
                    Dim wtrtkTankaD As Decimal = trtk_tanka_kngkD * trtk_koryoD
                    Dim wtekyoTankaD As Decimal = tekyo_tanka_kngkD * tekyo_koryoD
                    trtkTanka = CStr(dao.TruncateDec(wtrtkTankaD, [Const].C_YUKO_SHOSU_KETA))
                    tekyoTanka = CStr(dao.TruncateDec(wtekyoTankaD, [Const].C_YUKO_SHOSU_KETA))
                    'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD End
                Else
                    '追加工費　材料ユニット
                    trtkTanka = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_tanka_kngk"), ""))
                    tekyoTanka = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_tanka_kngk"), ""))
                End If

                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return True
    End Function
    'Rev002-End
#End Region


#Region "取替_自主点検チェックシートテーブル(取替票単位)"
    ''' <summary>
    ''' 取替_自主点検チェックシートテーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_自主点検チェックシート)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_CHECK_SHEET(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)
        Dim tblname As String = C_T_TRKE_CHECK_SHEET
        Dim isExist As Boolean = False

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'テーブルの処理を行う。
        For Each wRow As DataRow In wDataRows
            '自主点検チェックシートレコードの既存チェック
            isExist = myDB.CheckExistRec(tblname, whereStr)
            If isExist Then
                '取替票に対するレコードが既存の場合
                If Not dao.U004(wRow) Then
                    Return False
                End If
            Else
                '取替_自主点検チェックシートテーブルに追加する。
                If Not dao.I004(wRow) Then
                    Return False
                End If
            End If
        Next

        Return True

    End Function
#End Region


#Region "共通_添付書類詳細(取替票単位)"
    ''' <summary>
    ''' 共通_添付書類詳細(取替票単位)
    ''' </summary>
    ''' <param name="pRow">ファイル取替票（低圧）レコード</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(共通_添付書類詳細)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_TNPDOC_DTL(ByVal pRow As DataRow,
                                          ByVal dao As MJ4102_DAO,
                                          ByVal pDataTable As DataTable, ByVal pExpUploadFileDir As String,
                                          ByVal k1 As String,
                                          ByVal k2 As String,
                                          ByVal k3 As String,
                                          ByVal k4 As String,
                                          ByRef pWarningFlg As Boolean) As Boolean
        Dim wdig4_zgsyo_cd As String = ""
        Dim wtrkhy_hakko_nendo As String = ""
        Dim wtrkhy_kbn As String = ""
        Dim wtrkhy_no As String = ""
        Dim wtrkhy_no_key As String = "" '取替票番号キー

        Dim wDataTable As DataTable = Nothing
        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        Dim isExist As Boolean = False
        Dim tblname As String = C_T_TRKE_TRKEHY_TIAT         '取替_取替票（低圧）
        Dim field_tenp As String = "tenp_file_mng_no"        '添付ファイル管理番号 項目名
        Dim field_ryss As String = "ryssy_tenp_file_mng_no"  '領収書添付ファイル管理番号 項目名
        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""
        Dim wtenp_file_mng_no As String = ""          '添付ファイル管理番号
        Dim wryssy_tenp_file_mng_no As String = ""    '領収書添付ファイル管理番号
        Dim dir_tenp_file_mng_no As String = ""       '展開フォルダ用添付ファイル管理番号
        Dim dir_ryssy_tenp_file_mng_no As String = "" '展開フォルダ用領収書添付ファイル管理番号

        Dim wTnpDocKyes As HdSysBasePg.TnpDoc.TnpDocKeys = New HdSysBasePg.TnpDoc.TnpDocKeys
        Dim wTenpFileMngNo As Decimal = 0
        Dim wRyssyTenpFileMngNo As Decimal = 0
        Dim retval As Boolean = False

        Dim wPhotoKbn As String = ""
        Dim wPhotoSbt As String = ""
        Dim ht As Hashtable = New Hashtable
        Dim oYaFname As String = ""

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        'ファイル取替票（低圧）レコードの添付書類管理番号情報を取得する
        '→添付書類管理番号のフォルダを作成するのだが、値なし(null or 0)のフォルダは"0"１文字とする
        dir_tenp_file_mng_no = getDataRowItemCStr(pRow, field_tenp, "0")
        dir_ryssy_tenp_file_mng_no = getDataRowItemCStr(pRow, field_ryss, "0")
        'ファイル取替票（低圧）レコードの実費_No1～3の情報を取得する
        Dim wNo1_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_no"), ""))
        Dim wNo2_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_no"), ""))
        Dim wNo3_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_no"), ""))
        Dim wFile_renno As String = ""
        Dim wUpdZippiClm As String = ""
        Dim jippi_file_renno As Integer = 1 '添付書類詳細に登録する実費・領収書は1オリジンでカウントアップする

        '実費以外の写真用連番（共通部品を使用するので１～連番となる。ただし、庫出している場合は開始がずれる）
        Dim seq_file_renno As Integer = 0
        Dim Max_file_renno As String = ""
        dao.getMaxFileRenno(Max_file_renno, k1, k2, k3, k4)
        seq_file_renno = MyCInt(Max_file_renno) + 1

        '共通_添付書類の処理を行う。
        'wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,file_renno")
        'file_rennoが文字型のため「1,10,2,...」の並びになる。Integer列を追加し数値ソートする
        wDataTable = pDataTable.Copy()
        wDataTable.Columns.add("file_renno_int", GetType(Integer), "Convert(file_renno,'System.Int32')")
        wDataRows = wDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,file_renno_int")

        '添付書類を登録する。
        For Each wRow As DataRow In wDataRows
            wdig4_zgsyo_cd = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            wtrkhy_hakko_nendo = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            wtrkhy_kbn = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            wtrkhy_no = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))
            wPhotoKbn = CStr(Me.myDB.ConvDbNull(wRow("no1_tnpdoc_hosok_naiyo"), ""))
            wPhotoSbt = CStr(Me.myDB.ConvDbNull(wRow("no2_tnpdoc_hosok_naiyo"), ""))
            wtrkhy_no_key = getTrkhyNoKey(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no)

            jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(wdig4_zgsyo_cd)
            jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_hakko_nendo)
            jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_kbn)
            jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_no)

            'file_rennoとfile_msのハッシュテーブル登録
            ht.Add(CStr(Me.myDB.ConvDbNull(wRow("file_renno"), "")), CStr(Me.myDB.ConvDbNull(wRow("file_ms"), "")))

            '写真区分および写真種類をチェックし処理を分ける。
            If wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Jippi) Then
                '写真区分：実費の場合（実費の場合、写真種類はなし）
                '領収書添付書類管理番号が未取得の場合、取得する。
                wryssy_tenp_file_mng_no = myDB.GetRecValueString(tblname, field_ryss, jouken1, jouken2, jouken3, jouken4)
                If String.IsNullOrEmpty(wryssy_tenp_file_mng_no.Trim) Then
                    '領収書添付ファイル管理番号を新規取得する。
                    wTnpDocKyes.KEY1_VALUE = wtrkhy_no_key
                    retval = HdSysBasePg.TnpDoc.CreateTnpDocInfo(myDB, myUser, myPath, MjK1.Cmn.Const.TnpDocTypeCd.PhotoRyoushuu, wTnpDocKyes, wRyssyTenpFileMngNo)
                    If Not retval Then
                        Return False
                    End If
                    wryssy_tenp_file_mng_no = wRyssyTenpFileMngNo.ToString

                    '取替_取替票(低圧)の領収書添付ファイル管理番号を更新する。
                    If Not dao.U006(wRow, wryssy_tenp_file_mng_no) Then
                        Return False
                    End If
                Else
                    '検索した領収書添付書類管理番号をDecimal変換する。
                    wRyssyTenpFileMngNo = MyCDec(wryssy_tenp_file_mng_no)
                End If

                retval = addTnpDocKeiki(jippi_file_renno, wRow, wryssy_tenp_file_mng_no, pExpUploadFileDir, dir_ryssy_tenp_file_mng_no)
                If Not retval Then
                    Return False
                End If

                '実費・領収書を取替_取替票(低圧)の実費No1～3に登録する
                wFile_renno = CStr(Me.myDB.ConvDbNull(wRow("file_renno"), ""))
                retval = False
                If wFile_renno.Equals(wNo1_zippi_no) Then
                    retval = dao.U0061(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo2_zippi_no) Then
                    retval = dao.U0062(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo3_zippi_no) Then
                    retval = dao.U0063(wRow, CStr(jippi_file_renno))
                End If
                If Not retval Then
                    Return False
                End If
                jippi_file_renno = jippi_file_renno + 1

            ElseIf wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Tekkyo) Or wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Torituke) Then
                '撤去・取付の場合
                '添付書類管理番号を取得する。
                wtenp_file_mng_no = myDB.GetRecValueString(tblname, field_tenp, jouken1, jouken2, jouken3, jouken4)
                '添付ファイル管理番号未取得の場合、取得する。
                If String.IsNullOrEmpty(wtenp_file_mng_no.Trim) Then
                    '添付ファイル管理番号を新規取得する
                    wTnpDocKyes.KEY1_VALUE = wtrkhy_no_key
                    retval = HdSysBasePg.TnpDoc.CreateTnpDocInfo(myDB, myUser, myPath, MjK1.Cmn.Const.TnpDocTypeCd.PhotoKozi, wTnpDocKyes, wTenpFileMngNo)
                    If Not retval Then
                        Return False
                    End If
                    wtenp_file_mng_no = wTenpFileMngNo.ToString

                    '取替_取替票(低圧)に更新する。
                    If Not dao.U005(wRow, wtenp_file_mng_no) Then
                        Return False
                    End If
                Else
                    '検索した添付ファイル管理番号をDecimal変換する。
                    wTenpFileMngNo = MyCDec(wtenp_file_mng_no)
                End If

                'サムネイルの場合、親画像を参照し作成する。
                If CStr(wRow("no3_tnpdoc_hosok_naiyo")).Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                    'サムネイルの場合、親画像を参照し作成する。
                    oYaFname = CType(ht(CStr(Me.myDB.ConvDbNull(wRow("no4_tnpdoc_hosok_naiyo"), ""))), String)
                    retval = createThumbNailFile(oYaFname, wRow, wtenp_file_mng_no, pExpUploadFileDir, dir_tenp_file_mng_no)
                    If Not retval Then
                        Return False
                    End If
                End If

                '添付書類管理番号をキーに、画像を登録する。
                retval = addTnpDocKeiki(seq_file_renno, wRow, wtenp_file_mng_no, pExpUploadFileDir, dir_tenp_file_mng_no)
                If Not retval Then
                    Return False
                End If

                seq_file_renno = seq_file_renno + 1
            Else
                '庫出設定はアップロードの対象でないので取込なし
            End If
        Next

        Return True

    End Function

    Private Function getDataRowItemCStr(ByVal pRow As DataRow, ByVal itemName As String, ByVal defval As String) As String
        Dim str As String = defval
        Try
            str = CStr(Me.myDB.ConvDbNull(pRow(itemName), defval))
            If String.IsNullOrEmpty(str.Trim) Then
                str = defval
            End If
        Catch ex As Exception
            str = defval
        End Try

        Return str
    End Function

    ''' <summary>
    ''' 共通_添付書類詳細登録
    ''' </summary>
    ''' <param name="seq_file_renno">連番（登録の想定番号）</param>
    ''' <param name="pRow">ファイル添付書類詳細レコード</param>
    ''' <param name="pTenpFileMngNo">添付書類管理番号</param>
    ''' <param name="pExpUploadFileDir">アップロードデータ展開フォルダ</param>
    ''' <param name="pDirFileMngNo">展開フォルダ用添付書類管理番号フォルダ</param>
    ''' <remarks></remarks>
    Private Function addTnpDocKeiki(ByVal seq_file_renno As Integer, ByVal pRow As DataRow, ByVal pTenpFileMngNo As String, ByVal pExpUploadFileDir As String, ByVal pDirFileMngNo As String) As Boolean
        Dim wTenpFileMngNo As Decimal = MyCDec(pTenpFileMngNo)
        Dim wFileName As String = ""
        Dim wHosokuNaiyo As List(Of String) = New List(Of String)
        Dim retval As Boolean = False

        Try
            Dim wdig4_zgsyo_cd As String = CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))
            Dim wtrkhy_hakko_nendo As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))
            Dim wtrkhy_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))
            Dim wtrkhy_no As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))
            Dim wPhotoKbn As String = CStr(Me.myDB.ConvDbNull(pRow("no1_tnpdoc_hosok_naiyo"), ""))
            Dim wPhotoSbt As String = CStr(Me.myDB.ConvDbNull(pRow("no2_tnpdoc_hosok_naiyo"), ""))
            Dim wfile_ms As String = CStr(Me.myDB.ConvDbNull(pRow("file_ms"), ""))
            Dim wTNflag As String = CStr(Me.myDB.ConvDbNull(pRow("no3_tnpdoc_hosok_naiyo"), ""))
            Dim wNo4HosokNaiyo As String = ""

            '写真区分が実費の場合、写真種類は指定なしだがタブレットのフォルダ構成に使用しているので取得する
            If wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Jippi) Then
                wPhotoSbt = MjK1.Cmn.Const.KeikiPhotoSbtCd.RyoushuuSho
            End If

            'ファイルパス
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_TNPDOC, getTrkhyNoKey(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no))
            wFileName = System.IO.Path.Combine(wFileName, wPhotoKbn, wPhotoSbt, pDirFileMngNo, wfile_ms)

            '補足情報登録
            wHosokuNaiyo.Add(CStr(Me.myDB.ConvDbNull(pRow("no1_tnpdoc_hosok_naiyo"), "")))

            '写真区分が実費の場合、写真種類は指定なしだがタブレットで使用しているので削除する。
            If wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Jippi) Then
                wHosokuNaiyo.Add(" ")
            Else
                wHosokuNaiyo.Add(CStr(Me.myDB.ConvDbNull(pRow("no2_tnpdoc_hosok_naiyo"), "")))
            End If

            'no4_tnpdoc_hosok_naiyoの調整(欠番があっても
            If wTNflag.Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                'サムネイルの場合、参照元ファイル連番を決定する
                wNo4HosokNaiyo = (seq_file_renno - 1).ToString
            Else
                '以外の場合はファイル内情報を設定する
                wNo4HosokNaiyo = CStr(Me.myDB.ConvDbNull(pRow("no4_tnpdoc_hosok_naiyo"), ""))
            End If

            wHosokuNaiyo.Add(CStr(Me.myDB.ConvDbNull(pRow("no3_tnpdoc_hosok_naiyo"), "")))
            wHosokuNaiyo.Add(wNo4HosokNaiyo)


            retval = HdSysBasePg.TnpDoc.AddTnpDoc(myDB, myUser, myPath, wTenpFileMngNo, wFileName, wHosokuNaiyo)

        Catch ex As Exception

        End Try

        Return retval

    End Function

    Private Function createThumbNailFile(ByVal oYaFname As String, ByVal pRow As DataRow, ByVal pTenpFileMngNo As String, ByVal pExpUploadFileDir As String, ByVal pDirFileMngNo As String) As Boolean
        'サムネイル画像を作成する。
        Dim srcFname As String = ""
        Dim dstFname As String = ""
        Dim srcFullPath As String = ""
        Dim dstPath As String = ""
        Dim ratio As Integer = 20
        Dim retval As Boolean = True

        Try
            If Not CStr(pRow("no3_tnpdoc_hosok_naiyo")).Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                'サムネイル画像でない場合
                Return True
            End If

            'サムネイルの場合
            '設定されているファイル名を取得する
            dstFname = CStr(pRow("file_ms"))

            '参照元からサムネイル画像を作成する
            Dim wdig4_zgsyo_cd As String = CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))
            Dim wtrkhy_hakko_nendo As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))
            Dim wtrkhy_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))
            Dim wtrkhy_no As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))
            Dim wPhotoKbn As String = CStr(Me.myDB.ConvDbNull(pRow("no1_tnpdoc_hosok_naiyo"), ""))
            Dim wPhotoSbt As String = CStr(Me.myDB.ConvDbNull(pRow("no2_tnpdoc_hosok_naiyo"), ""))
            srcFname = oYaFname
            srcFullPath = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_TNPDOC, getTrkhyNoKey(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no))
            srcFullPath = System.IO.Path.Combine(srcFullPath, wPhotoKbn, wPhotoSbt, pDirFileMngNo, srcFname)

            'サムネイルのファイル名
            dstFname = CStr(Me.myDB.ConvDbNull(pRow("file_ms"), ""))
            dstPath = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_TNPDOC, getTrkhyNoKey(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no))
            dstPath = System.IO.Path.Combine(dstPath, wPhotoKbn, wPhotoSbt, pDirFileMngNo)

            'サムネイルファイル作成
            retval = HdSysBase.Cmn.CreateThumbnailImageFile(srcFullPath, dstPath, dstFname, ratio)
            If Not retval Then
                Return False
            End If

            Return True

        Catch ex As Exception
            Return False

        End Try
    End Function
#End Region


#Region "テーブル更新（抜取検査）"
    ''' <summary>
    ''' テーブル更新（抜取検査）
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataNktrKns(ByVal dao As MJ4102_DAO,
                                    ByVal pUpldKttmNo As String,
                                    ByVal pKeikiDwldSbtCd As String,
                                    ByVal pUpldUserId As String,
                                    ByVal pExpUploadFileDir As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow

        Try
            '取替票（低圧）ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_TIAT + C_MS_NKTRKNS_KK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_TIAT)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
				'Rev002-Start
                '高圧他ファイルを指定
                wFileTableMS = C_MS_TRKE_TRKEHY_KAHK + C_MS_NKTRKNS_KK
                wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_KAHK)

                '取替票（高圧他）ファイルを読み込む
                If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                    '低圧も高圧他も両方読み取れなかったのでエラー
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False

                Else
                    '取替票（高圧他）テーブルを更新する(抜取検査)。
                    If Not updateTRKE_TRKEHY_KAHK_NKTRKNS(dao, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId, pExpUploadFileDir, wDtblUpload, pWarningFlg, pErrImportFlg) Then
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                        pWarningFlg = True
                        pErrImportFlg = True
                        Return False
                    End If
                End If

            Else

                '取替票（低圧）テーブルを更新する(抜取検査)。
                If Not updateTRKE_TRKEHY_TIAT_NKTRKNS(dao, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId, pExpUploadFileDir, wDtblUpload, pWarningFlg, pErrImportFlg) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False
                End If
			'Rev002-End
            End If


            '共通_個人情報アクセスログに登録する。
            wFileTableMS = C_MS_COM_KIFRFLG
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_KIFRFLG)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_KIFRFLG(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False
                End If
            End If


            '共通_ログ収集に登録する。
            wFileTableMS = C_MS_COM_LOG_SYUSYU
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOG_SYUSYU)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOG_SYUSYU(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False
                End If
            End If

            '共通_ログイン認証履歴に登録する。
            wFileTableMS = C_MS_COM_LOGIN_NNSY_RK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOGIN_NNSY_RK)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOGIN_NNSY_RK(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False
                End If
            End If

            '共通_ログイン履歴データに登録する。
            wFileTableMS = C_MS_COM_LOGIN_RIREKI_DATA
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_DIR_LOG, C_F_COM_LOGIN_RIREKI_DATA)
            If HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                'ファイル読込できたら登録する。
                If Not updateCOM_LOGIN_RIREKI_DATA(dao, wDtblUpload) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                    pWarningFlg = True
                    pErrImportFlg = True
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region
#Region "テーブル更新（抜取検査）　取替票（低圧）連係"
    ''' <summary>
    ''' テーブル更新（抜取検査）　取替票（低圧）連係
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataNktrKnsOthers(ByVal dao As MJ4102_DAO,
                                    ByVal pExpUploadFileDir As String,
                                    ByVal k1 As String,
                                    ByVal k2 As String,
                                    ByVal k3 As String,
                                    ByVal k4 As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean

        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow

        Try
            '取替票行程情報ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_PRC + C_MS_NKTRKNS_KK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_PRC)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If
            '取替票行程情報テーブルを更新する(抜取検査)。
            If Not updateTRKE_TRKEHY_PRC_NKTRKNS(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If



            '自主点検チェックシートファイルを読み込む。
            wFileTableMS = C_MS_TRKE_CHECK_SHEET + C_MS_NKTRKNS_KK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_CHECK_SHEET)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If
            '自主点検チェックシートテーブルを更新する(抜取検査)。
            If Not updateTRKE_CHECK_SHEET_NKTRKNS(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region


#Region "取替_取替票（低圧）テーブル(抜取検査)"
    ''' <summary>
    ''' 取替_取替票（低圧）テーブル(抜取検査)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir">アップロードファイル展開フォルダ</param>
    ''' <param name="pDataTable">アップロードデータ取替_取替票（低圧）</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_TIAT_NKTRKNS(ByVal dao As MJ4102_DAO,
                                                    ByVal pUpldKttmNo As String,
                                                    ByVal pKeikiDwldSbtCd As String,
                                                    ByVal pUpldUserId As String,
                                                    ByVal pExpUploadFileDir As String,
                                                    ByVal pDataTable As DataTable, ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票（低圧）テーブルが更新可能かチェックする。
            Dim k1 As String = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            Dim k2 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            Dim k3 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            Dim k4 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))
            Dim jouken1 As String = "dig4_zgsyo_cd = " & EcOrgIO.EcOrgString.GetSqlText(k1)
            Dim jouken2 As String = "trkhy_hakko_nendo = " & EcOrgIO.EcOrgString.GetSqlText(k2)
            Dim jouken3 As String = "trkhy_kbn = " & EcOrgIO.EcOrgString.GetSqlText(k3)
            Dim jouken4 As String = "trkhy_no = " & EcOrgIO.EcOrgString.GetSqlText(k4)
            Dim s1 As String = ""
            Dim s2 As String = ""
            Dim s3 As String = ""
            Dim s4 As String = ""
            Dim field1 As String = "prcmg_cd"
            Dim field2 As String = "nktr_kensa_zyokyo_cd"
            Dim field3 As String = "upd_date"
            Dim field4 As String = "ntkns_zissi_ymd"
            Dim retval As Boolean = False
            Dim nextProc As String = ""
            retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_TIAT,
                                                 field1, field2, field3, field4, "", "", "", "", "", "",
                                                 s1, s2, s3, s4, "", "", "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            If Not retval Then
                '取替票項目値の取得エラーの場合
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票検索エラーです(抜取検査):" + k1 + "-" + k2 + k3 + k4)
                pWarningFlg = True
                Continue For
            End If
            If String.IsNullOrEmpty(s1.Trim) Then
                '該当レコードがない場合
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票がありません(抜取検査):" + k1 + "-" + k2 + k3 + k4)
                pWarningFlg = True
                Continue For
            End If

            '施工と異なり物理的NGではないのでスキップしない。2023/4/22
            ''抜取検査実施年月日が登録済みならスキップする
            'If Not String.IsNullOrEmpty(s4.Trim) Then
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "抜取検査実施年月日が登録済みです:" + k1 + "-" + k2 + k3 + k4)
            '    Continue For
            'End If

            'アップロードの排他確認は不要。多重ダウンロードが可能のため。
            ''排他確認　更新年月日が更新済みならスキップする
            'Dim fileUpdDate As String = getDateStrWithKeta(wRow("upd_date").ToString, C_KETA_YMDHMS)
            'Dim dbUpdDate As String = getDateStrWithKeta(s3, C_KETA_YMDHMS)
            'If fileUpdDate <= dbUpdDate Then
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "更新年月日が更新済みです(抜取検査):" + k1 + "-" + k2 + k3 + k4)
            '    Continue For
            'End If


            '取替_取替票（低圧）テーブルを更新する(抜取検査)。
            If Not dao.U007(wRow, pUpldUserId) Then
                Return False
            End If

            '更新した取替_取替票（低圧）に連係するテーブルを取替票単位で更新する。
            '（更新対象の取替票についてのみ、連係するテーブルも更新する。）
            If Not updateDataNktrKnsOthers(dao, pExpUploadFileDir, k1, k2, k3, k4, pWarningFlg, pErrImportFlg) Then
                Return False
            End If


            'ダウンロード取替票管理の持出ファイル状態コードを「5:アップロード済」に更新する。
            If Not dao.U011(wRow, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId) Then
                Return False
            End If

        Next

        Return True

    End Function

#End Region

#Region "取替_取替票行程テーブル(抜取検査)(取替票単位)"
    ''' <summary>
    ''' 取替_取替票行程テーブル(抜取検査)(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票行程)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_PRC_NKTRKNS(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票行程テーブルを更新する。(抜取検査)
            If Not dao.U008(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region

#Region "取替_自主点検チェックシートテーブル(抜取検査)(取替票単位)"
    ''' <summary>
    ''' 取替_自主点検チェックシートテーブル(抜取検査)(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_自主点検チェックシート)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_CHECK_SHEET_NKTRKNS(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_自主点検チェックシートテーブルを更新する(抜取検査)。
            If Not dao.U009(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region



#Region "共通_個人情報アクセスログテーブル"
    ''' <summary>
    ''' 共通_個人情報アクセスログテーブル
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_KIFRFLG(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable) As Boolean

        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "saksi_date, kifrflg_mng_renno")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '新規登録する。
            If Not dao.I005(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region
'Rev002-Start
#Region "取替_取替票（高圧他）テーブル"
    'Rev002-Start
    ''' <summary>
    ''' 取替_取替票（高圧他）テーブル
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_KAHK(ByVal dao As MJ4102_DAO,
                                            ByVal pUpldKttmNo As String,
                                            ByVal pKeikiDwldSbtCd As String,
                                            ByVal pUpldUserId As String,
                                            ByVal pExpUploadFileDir As String,
                                            ByVal pDataTable As DataTable, ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票（高圧他）テーブルが更新可能かチェックする。
            Dim k1 As String = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            Dim k2 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            Dim k3 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            Dim k4 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))

            Dim jouken1 As String = "dig4_zgsyo_cd = " & EcOrgIO.EcOrgString.GetSqlText(k1)
            Dim jouken2 As String = "trkhy_hakko_nendo = " & EcOrgIO.EcOrgString.GetSqlText(k2)
            Dim jouken3 As String = "trkhy_kbn = " & EcOrgIO.EcOrgString.GetSqlText(k3)
            Dim jouken4 As String = "trkhy_no = " & EcOrgIO.EcOrgString.GetSqlText(k4)

            Dim s1 As String = ""
            Dim s2 As String = ""
            Dim s3 As String = ""
            Dim s4 As String = ""
            Dim field1 As String = "prcmg_cd"
            Dim field2 As String = "kesbt_bnrui_cd" '計器種別分類コード←検定種別コード"knti_sbt_cd"
            Dim field3 As String = "upd_date"
            Dim field4 As String = "syun_ymd"
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる ADD Start
            Dim s5 As String = ""
            Dim field5 As String = "trkhy_hinpt_flg"
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる ADD End
            Dim retval As Boolean = False
            Dim nextProc As String = ""
            Dim isTaidou As Boolean = False
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる MOD Start
            'retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_KAHK,
            '                                     field1, field2, field3, field4, "", "", "", "", "", "",
            '                                     s1, s2, s3, s4, "", "", "", "", "", "",
            '                                     jouken1, jouken2, jouken3, jouken4)
            retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_KAHK,
                                                 field1, field2, field3, field4, field5, "", "", "", "", "",
                                                 s1, s2, s3, s4, s5, "", "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            'Rev076 庫入情報送信エラー対応　行程不正で「手入力」の場合、取替票取込は正常終了させる MOD End
            If Not retval Then
                'テーブルの取替票項目値の取得エラーの場合、エラー終了
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票検索エラーです:" + k1 + "-" + k2 + k3 + k4)
                Return False
            End If

            '行程が他異動かチェックする(行程の先頭1文字が'F'であること)
            If s1(0) = MjK1.Cmn.Const.ProcMngCd.FTaIdou(0) Then
                isTaidou = True
            End If

            '取替票の存在チェックと行程チェック
            If String.IsNullOrEmpty(s1.Trim) Then
                'テーブルに該当取替票キーのレコードがない場合、エラー終了
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票がありません:" + k1 + "-" + k2 + k3 + k4)
                Return False
            ElseIf Not s1.Equals(MjK1.Cmn.Const.ProcMngCd.CJusin) And isTaidou = False Then
                'テーブルの行程管理コードが「C040:取替票受信済」「Fxxx(F010:他異動中止,F020:他異動中止確認)」以外の場合、スキップする
                'Rev076 庫入情報送信エラー対応　行程不正の場合、「手入力フラグがONでない場合」のみ取込エラーとする MOD Start
                'HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票の行程が不正です:" + k1 + "-" + k2 + k3 + k4 + "[" + s1 + "]")
                'pWarningFlg = True
                If Not s5.Trim.Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票の行程が不正です:" + k1 + "-" + k2 + k3 + k4 + "[" + s1 + "]")
                    pWarningFlg = True
                End If
                'Rev076 庫入情報送信エラー対応　行程不正の場合、「手入力フラグがONでない場合」のみ取込エラーとする MOD End
                Continue For
            End If

            '次の行程管理コードを設定する。他異動以外の場合、竣工年月日・更新年月日のチェックを行う。
            If isTaidou Then
                '他異動中止の場合、変更しない。（SQLのupdate文に今と同じ行程を指定することで対応）
                nextProc = s1
            Else
                '取替票受信済の場合
                'CT付かどうか判定する（計器種別分類コードで判定。検定種別コードはやめる）
                If s2.Trim.Equals(MjK1.Cmn.Const.KesbtBriCd.CT) Then
                    'CT付    D010：竣工結果確認待
                    nextProc = MjK1.Cmn.Const.ProcMngCd.DShunkoKekka
                Else
                    'CTなし　D020：竣工報告待
                    nextProc = MjK1.Cmn.Const.ProcMngCd.DShunkoHokoku
                End If

                'タブレットで施工結果送信時にチェックしているので、ここでは行わない　2023/6/6
                ''テーブルの竣工年月日が登録済みならエラー終了
                'If Not String.IsNullOrEmpty(s4.Trim) Then
                '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "竣工年月日が登録済みです:" + k1 + "-" + k2 + k3 + k4)
                '    Return False
                'End If

                'アップロードの排他確認は不要。多重ダウンロードが可能のため。
                ''排他確認　テーブルの更新年月日が更新済みならエラー終了
                'Dim fileUpdDate As String = getDateStrWithKeta(wRow("upd_date").ToString, C_KETA_YMDHMS)
                'Dim dbUpdDate As String = getDateStrWithKeta(s3, C_KETA_YMDHMS)
                'If fileUpdDate <= dbUpdDate Then
                '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "更新年月日が更新済みです:" + k1 + "-" + k2 + k3 + k4)
                '    Return False
                'End If
            End If



            retval = False
            Dim wttkk_keiki_id = CStr(Me.myDB.ConvDbNull(wRow("ttkk_keiki_id"), "")).Trim
            'Dim jouken5 As String = "ttkk_keiki_id = " & EcOrgIO.EcOrgString.GetSqlText(wttkk_keiki_id)
            'retval = myDB.CheckExistRec(C_T_TRKE_TRKEHY_TIAT, jouken2, jouken5)

            'B-ST-0196 取付計器の多重チェックは不要となったため処理を削除 2023/5/26
            ''同一取付計器の多重チェック（同一年度の取替票を検索）
            'Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
            'If Not dao.S003(dbr, wRow) Then
            '    Return False
            'End If
            'If dbr.DataStruct.HasRows Then
            '    retval = True
            'End If
            'dbr.Close()
            '
            'If retval Then
            '    '多重している場合（自取替票は未更新なので、除外する取替票は無い。）
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取付計器IDは使用済みです:" + k1 + "-" + k2 + k3 + k4 + "(" + wttkk_keiki_id + ")")
            '    Return False
            'End If

            'Rev076 庫入情報送信結果が「廃滅登録重複」の場合、SMJ撤去計器チェックを行う ADD Start
            Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
            Try
                Dim smj_send_zumi_flg = CStr(Me.myDB.ConvDbNull(wRow("smj_send_zumi_flg"), "")).Trim
                Dim wValue1 = MjK1.Cmn.Func.GetZyousu(myDB, [Const].C_ZYOUSUCD_HAIMETU_DUP_VALUE, SyoriYMD)
                If smj_send_zumi_flg.Equals(wValue1) Then
                    '廃滅登録重複の場合、撤去計器IDを条件にSMJから計器情報を取得する
                    Dim wtkkk_keiki_id = CStr(Me.myDB.ConvDbNull(wRow("tkkk_keiki_id"), "")).Trim
                    Dim wkstr1 As String = ""
                    Dim wkstr2 As String = ""
                    Dim wkstr3 As String = ""
                    Dim wkstr4 As String = ""

                    If Not dao.S008(dbr, wtkkk_keiki_id) Then
                        'DB検索エラー
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック DBエラーです(高圧):" + k1 + "-" + k2 + k3 + k4)
                        Return False
                    End If

                    If Not dbr.DataStruct.HasRows Then
                        '撤去計器情報なし
                        HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 撤去計器情報なし(高圧):" + k1 + "-" + k2 + k3 + k4)
                        Return False
                    End If

                    While (dbr.DataStruct.Read)
                        '庫入全日有効電力量(現在値) = 撤去計器_指示数_総合
                        wkstr1 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_sougo_szsu"), "")).Trim
                        '庫入最大DM(現在値) = 撤去計器_指示数_最大需要
                        wkstr2 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_max_zyyo_szsu"), "")).Trim
                        '庫入力測有効電力量(現在値) = 撤去計器_指示数_力測有効
                        wkstr3 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_rsyk_szsu"), "")).Trim
                        '庫入力測無効電力量(現在値)遅れ = 撤去計器_指示数_力測無効
                        wkstr4 = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tkkk_rsmk_szsu"), "")).Trim

                        If String.IsNullOrEmpty(wkstr1) Or String.IsNullOrEmpty(wkstr2) Or String.IsNullOrEmpty(wkstr3) Or String.IsNullOrEmpty(wkstr4) Then
                            '指示数がない
                            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "廃滅登録重複チェック 撤去計器 指示数なし(高圧):" + k1 + "-" + k2 + k3 + k4)
                            Return False
                        End If

                        '１件のみ処理
                        Exit While
                    End While
                End If
            Catch ex As Exception
            Finally
                dbr.Close()
            End Try
            'Rev076 庫入情報送信結果が「廃滅登録重複」の場合、SMJ撤去計器チェックを行う ADD End

            '取替_取替票（高圧他）テーブルを更新する。
            If Not dao.U012(wRow, nextProc, SyoriYMD) Then
                Return False
            End If


            '更新した取替_取替票（高圧他）に連係するテーブルを取替票単位で更新する。
            '（更新対象の取替票についてのみ、連係するテーブルも更新する。）
            If Not updateDataSekoOthersKahk(wRow, dao, pExpUploadFileDir, pDataTable, k1, k2, k3, k4, pWarningFlg, pErrImportFlg) Then
                Return False
            End If



            'ダウンロード取替票管理の持出ファイル状態コードを「5:アップロード済」に更新する。
            If Not dao.U011(wRow, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId) Then
                Return False
            End If


            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD Start
            '個別管理以外(計器指定番号が数字のみ)の場合、計器マスタに連係する
            If Not System.Text.RegularExpressions.Regex.IsMatch(CStr(Me.myDB.ConvDbNull(wRow("keiki_sitei_no"), "")).Trim, "[^\d]") Then
                '低圧・高圧を判定し、各々の連係を行う
                If CStr(Me.myDB.ConvDbNull(wRow("keiki_kbn_cd"), "")).Trim.Equals([Const].C_KEIKI_KBN_CD_TIAT) Then
                    '撤去計器_相線式_相コードが空白の場合は、低圧
                    '計器マスタ連係　低圧

                    'WEBAPI実行管理テーブルへの登録を行う。
                    'IF023_取付完了通知
                    If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF023) Then
                        Return False
                    End If
                    'IF024_撤去完了通知
                    If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF024) Then
                        Return False
                    End If
                    'IF025_設備情報更新（端子部再用の場合）
                    Dim tnsb_reuse_kbn As String = CStr(Me.myDB.ConvDbNull(wRow("tnsb_reuse_kbn"), ""))
                    Dim kbn As Integer = 0
                    If tnsb_reuse_kbn.Equals([Const].C_TNSB_REUSE) Then
                        '端子部再用の場合
                        If IsTnsbQR2G(CStr(Me.myDB.ConvDbNull(wRow("tkkk_tnsb_sbt_cd"), ""))) Then
                            '第２世代
                            kbn = [Const].C_KEIKI_MASTER_RENKEI_IF025G2
                        Else
                            '第１世代
                            kbn = [Const].C_KEIKI_MASTER_RENKEI_IF025G1
                        End If
                        If Not dao.I011(wRow, kbn) Then
                            Return False
                        End If
                    End If
                Else
                    '計器マスタ連係　高圧

                    'WEBAPI実行管理テーブルへの登録を行う。
                    'IF023_取付完了通知
                    If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF023KH) Then
                        Return False
                    End If
                    'IF024_撤去完了通知
                    If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF024KH) Then
                        Return False
                    End If

                    'IF025_設備情報更新（変成器の連係）
                    If Not dao.I011(wRow, [Const].C_KEIKI_MASTER_RENKEI_IF025KH) Then
                        Return False
                    End If
                End If
            End If
            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD End
        Next

        '１件でもWarningがあればFalseを返却する
        If pWarningFlg Then
            Return False
        End If

        Return True

    End Function
    'Rev002-End
#End Region



#Region "取替_取替票行程（高圧他）テーブル(取替票単位)"
    'Rev002-Start
    ''' <summary>
    ''' 取替_取替票行程（高圧他）テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票行程)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_PRC_KAHK(ByVal dao As MJ4102_DAO, ByRef pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""
        '取替票キー条件設定
        jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(k1)
        jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(k2)
        jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(k3)
        jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(k4)

        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim s4 As String = ""
        Dim s5 As String = ""
        Dim s6 As String = ""
        Dim field1 As String = "no1_zippi_no"
        Dim field2 As String = "no2_zippi_no"
        Dim field3 As String = "no3_zippi_no"
        Dim field4 As String = "no4_zippi_no"
        Dim field5 As String = "no5_zippi_no"
        Dim field6 As String = "zippi_umu_flg"
        Dim wtable As String = C_T_TRKE_TRKEHY_KAHK
        Dim retval As Boolean = False

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If


        '実費確認フラグを追加する
        pDataTable.Columns.Add("zippi_umu_flg")
        pDataTable.Columns.Add("no1_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no2_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no3_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no4_skrepo_zippi_kknn_flg")
        pDataTable.Columns.Add("no5_skrepo_zippi_kknn_flg")

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替票（高圧他）の実費ありの場合、竣工報告_実費確認フラグ1～5を設定する
            retval = myDB.GetRecValueStringMulti(wtable,
                                                 field1, field2, field3, field4, field5, field6, "", "", "", "",
                                                 s1, s2, s3, s4, s5, s6, "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            wRow("zippi_umu_flg") = s6.Trim
            'Rev002.1 MOD 20241210 フラグオフ時のクリアが抜けていたので追加
            If retval Then
                If Not String.IsNullOrEmpty(s1.Trim) Then
                    wRow("no1_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no1_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s2.Trim) Then
                    wRow("no2_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no2_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s3.Trim) Then
                    wRow("no3_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no3_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s4.Trim) Then
                    wRow("no4_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no4_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
                If Not String.IsNullOrEmpty(s5.Trim) Then
                    wRow("no5_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_NOT_CONFIRM
                Else
                    wRow("no5_skrepo_zippi_kknn_flg") = [Const].C_ZIPPIKKNN_FLG_OFF
                End If
            End If

            '取替_取替票行程（高圧他）テーブルを更新する。
            If Not dao.U013(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
    'Rev002-End
#End Region




#Region "取替_取替票追加工費（高圧他）テーブル(取替票単位)"
    'Rev002-Start
    ''' <summary>
    ''' 取替_取替票追加工費（高圧他）テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票追加工費（高圧他）)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_TUIKAKH_KAHK(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean
        Dim isExist As Boolean = False
        Dim tblname As String = C_T_TRKE_TRKEHY_TUIKAKH_KAHK
        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""
        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)
        Dim trtkTanka As Decimal = 0
        Dim tekyoTanka As Decimal = 0
        Dim wtrtkTanka As String = ""
        Dim wtekyoTanka As String = ""
        Dim trtkKojiHi As Decimal = 0         '追加工費テーブルの取付工事費
        Dim tekyoKojiHi As Decimal = 0        '追加工費テーブルの撤去工事費
        Dim trkhy_trtkKojiHi As Decimal = 0   '取替票単位の追加工費取付の合計
        Dim trkhy_tekyoKojiHi As Decimal = 0  '取替票単位の追加工費撤去の合計
        Dim TUIKAKH_umu_flg As String = MjK1.Cmn.Const.UmuFlg.FlgOff

        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim s4 As String = ""
        Dim s5 As String = ""
        Dim syun_ymd As String = ""
        Dim wrms_kbn As String = ""
        Dim zippi1 As String = ""
        Dim zippi2 As String = ""
        Dim zippi3 As String = ""
        Dim zippi4 As String = ""
        Dim zippi5 As String = ""
        Dim field1 As String = "khkmk_trtk_kozih_kngk"
        Dim field2 As String = "khkmk_tekyo_kozih_kngk"
        Dim field3 As String = "kozitn_cd"
        Dim field4 As String = "kzkmk_hnsik_trtk_kozih_kngk"
        Dim field5 As String = "kzkmk_hnsik_tekyo_kozih_kngk"
        Dim field6 As String = "syun_ymd"
        Dim field7 As String = "khkmk_wrms_kbn"
        Dim field8 As String = "no1_zippi_kngk"
        Dim field9 As String = "no2_zippi_kngk"
        Dim field10 As String = "no3_zippi_kngk"
        Dim field11 As String = "no4_zippi_kngk"
        Dim field12 As String = "no5_zippi_kngk"
        Dim wtable As String = C_T_TRKE_TRKEHY_KAHK
        Dim retval As Boolean
        Dim trkhy_trtk_kozih_kngk As Decimal = 0        '取替票　取付工事費
        Dim trkhy_tekyo_kozih_kngk As Decimal = 0       '取替票　撤去工事費
        Dim kzkmk_hnsik_trtk_kozih_kngk As Decimal = 0  '工事項目_変成器_取付工事費 
        Dim kzkmk_hnsik_tekyo_kozih_kngk As Decimal = 0 '工事項目_変成器_撤去工事費
        Dim zippi1_kngk As Decimal = 0                  '実費1
        Dim zippi2_kngk As Decimal = 0                  '実費2
        Dim zippi3_kngk As Decimal = 0                  '実費3
        Dim zippi4_kngk As Decimal = 0                  '実費4
        Dim zippi5_kngk As Decimal = 0                  '実費5
        Dim gokei_kngk As Decimal = 0                   '取替票　工事費合計
        Dim tuikaKohiKoziKbn As String = ""
        Dim wTuikakhSytkMstKbn As String = ""
        Dim jouken5 As String = ""

        '取替票キー条件設定
        jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(k1)
        jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(k2)
        jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(k3)
        jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(k4)


        '取替票の取付工事費・撤去工事費・竣工年月日・割増区分を取得する。（高圧他用に拡張版を利用）
        retval = myDB.GetRecValueStringMultiEx(wtable,
                                                 field1, field2, field3, field4, field5, field6, field7, field8, field9, field10, field11, field12,
                                                 s1, s2, s3, s4, s5, syun_ymd, wrms_kbn, zippi1, zippi2, zippi3, zippi4, zippi5,
                                                 jouken1, jouken2, jouken3, jouken4)
        If Not retval Then
            Return False
        End If
        trkhy_trtk_kozih_kngk = MyCDec(s1)   '取替票　取付工事費
        trkhy_tekyo_kozih_kngk = MyCDec(s2)  '取替票　撤去工事費
        kzkmk_hnsik_trtk_kozih_kngk = MyCDec(s4)  '工事項目_変成器_取付工事費 
        kzkmk_hnsik_tekyo_kozih_kngk = MyCDec(s5) '工事項目_変成器_撤去工事費
        zippi1_kngk = MyCDec(zippi1)              '実費1
        zippi2_kngk = MyCDec(zippi2)              '実費2
        zippi3_kngk = MyCDec(zippi3)              '実費3
        zippi4_kngk = MyCDec(zippi4)              '実費4
        zippi5_kngk = MyCDec(zippi5)              '実費5

        '割増率取得
        Dim wrms_ritu As String = dao.getKoryoWrmsRitu(wrms_kbn, syun_ymd)


        '取得したデータ件数 > 0 の場合、追加工費を登録し、追加工費金額を算出する。
        If (pDataTable.Rows.Count > 0) Then
            wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,tuika_kohi_row_no")

            '取替票番号キーは確定済み。k1,k2,k3,k4

            '取替_取替票追加工費　既存チェック
            isExist = myDB.CheckExistRec(tblname, jouken1, jouken2, jouken3, jouken4)
            '既存の場合、取替_取替票追加工費テーブルのレコードを削除する。
            If isExist Then
                '該当取替票番号のレコード一式を削除する。
                If Not dao.D002(k1, k2, k3, k4) Then
                    Return False
                End If
            End If

            '取替_取替票追加工費の処理を行う。
            For Each wRow As DataRow In wDataRows
                '追加工費区分を取得する。
                tuikaKohiKoziKbn = CStr(Me.myDB.ConvDbNull(wRow("tuika_kohi_kozi_kbn"), ""))
                tuikaKohiKoziKbn = tuikaKohiKoziKbn.PadLeft([Const].C_KORYO_CD_KETA, "0"c)

                '追加工費_取得マスタ区分取得
                jouken5 = "koryo_zunit_cd = " & EcOrgIO.EcOrgString.GetSqlText(tuikaKohiKoziKbn)
                wTuikakhSytkMstKbn = myDB.GetRecValueString("MST_USE_KANO_TUIKAKH", "tuikakh_sytk_mst_kbn", jouken5)

                '追加工費_工事区分の単価を取得する
                getTankaT(dao, tuikaKohiKoziKbn, wTuikakhSytkMstKbn, syun_ymd, wtrtkTanka, wtekyoTanka)
                trtkTanka = MyCDec(wtrtkTanka)
                tekyoTanka = MyCDec(wtekyoTanka)

                '工量マスタ・材料ユニットで処理を分ける
                If wTuikakhSytkMstKbn.Equals(MjK1.Cmn.Const.TuikakhSytkMstKbn.ZunitMst) Then
                    '材料ユニットの場合（割増なし）
                    '追加工費_取付工事費を算出する（単価*数量）
                    trtkKojiHi = trtkTanka * MyCDec(wRow("tuika_kohi_trtk_suryo").ToString)
                    '追加工費_撤去工事費を算出する（単価*数量）
                    tekyoKojiHi = tekyoTanka * MyCDec(wRow("tuika_kohi_tekyo_suryo").ToString)
                    'マイナス値となる
                    tekyoKojiHi = tekyoKojiHi * -1
                Else
                    '工量ユニットマスタの場合（割増あり）
                    '追加工費_取付工事費を算出する（単価*数量）
                    trtkKojiHi = trtkTanka * MyCDec(wRow("tuika_kohi_trtk_suryo").ToString)
                    trtkKojiHi = dao.CalcWrmsKngk(trtkKojiHi, wrms_ritu)
                    '追加工費_撤去工事費を算出する（単価*数量）
                    tekyoKojiHi = tekyoTanka * MyCDec(wRow("tuika_kohi_tekyo_suryo").ToString)
                    tekyoKojiHi = dao.CalcWrmsKngk(tekyoKojiHi, wrms_ritu)
                End If

                '取替_取替票追加工費（高圧他）を追加する
                If Not dao.I009(wRow, trtkKojiHi.ToString, tekyoKojiHi.ToString) Then
                    Return False
                End If

                '取替票単位の追加工費の合計を算出する
                trkhy_trtkKojiHi = trkhy_trtkKojiHi + trtkKojiHi     '取替票　追加工費取付の合計
                trkhy_tekyoKojiHi = trkhy_tekyoKojiHi + tekyoKojiHi  '取替票　追加工費撤去の合計

                TUIKAKH_umu_flg = MjK1.Cmn.Const.UmuFlg.FlgOn
            Next
        End If


        '******************************************************
        '計算結果を取替票番号キーのデータに更新する。
        '******************************************************

        '取替票　工事費合計
        gokei_kngk = 0
        gokei_kngk = gokei_kngk + trkhy_trtk_kozih_kngk + trkhy_tekyo_kozih_kngk             '取付工費 + 撤去工費
        gokei_kngk = gokei_kngk + kzkmk_hnsik_trtk_kozih_kngk + kzkmk_hnsik_tekyo_kozih_kngk '変成器_取付工費 + 変成器_撤去工費
        gokei_kngk = gokei_kngk + trkhy_trtkKojiHi + trkhy_tekyoKojiHi                       '追加工費_取付工費 + 追加工費_撤去工費
        gokei_kngk = gokei_kngk + zippi1_kngk + zippi2_kngk + zippi3_kngk + zippi4_kngk + zippi5_kngk  '実費1,2,3,4,5

        '直営工事の場合、金額に０円を設定する。
        Dim wChokueiKztnCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, SyoriYMD)
        If s3.Trim.Equals(wChokueiKztnCd) Then
            '０円設定
            '工費項目_取付工事費 
            '工費項目_撤去工事費 
            '工事項目_変成器_取付工事費 
            '工事項目_変成器_撤去工事費 
            '→取替票の処理時に設定済

            '追加工費_取付工事費 
            trkhy_trtkKojiHi = 0
            '追加工費_撤去工事費
            trkhy_tekyoKojiHi = 0
            '工事費合計 
            gokei_kngk = 0
        End If


        '取替票（高圧他）　追加工費取付・追加工費撤去および工事費合計を更新する
        If Not dao.U016(k1, k2, k3, k4, TUIKAKH_umu_flg, trkhy_trtkKojiHi.ToString, trkhy_tekyoKojiHi.ToString, gokei_kngk.ToString) Then
            Return False
        End If


        Return True
    End Function
#End Region



#Region "取替_自主点検チェックシート（高圧他）テーブル(取替票単位)"
    'Rev002-Start
    ''' <summary>
    ''' 取替_自主点検チェックシート（高圧他）テーブル(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_自主点検チェックシート)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_CHECK_SHEET_KAHK(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)
        Dim tblname As String = C_T_TRKE_CHECK_SHEET_KAHK
        Dim isExist As Boolean = False

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'テーブルの処理を行う。
        For Each wRow As DataRow In wDataRows
            '自主点検チェックシート（高圧他）レコードの既存チェック
            isExist = myDB.CheckExistRec(tblname, whereStr)
            If isExist Then
                '取替票に対するレコードが既存の場合
                If Not dao.U017(wRow) Then
                    Return False
                End If
            Else
                '取替_自主点検チェックシート（高圧他）テーブルに追加する。
                If Not dao.I010(wRow) Then
                    Return False
                End If
            End If
        Next

        Return True

    End Function
    'Rev002-End
#End Region



    'Rev002-Start
    ''' <summary>
    ''' 共通_添付書類詳細(取替票単位)（高圧他）
    ''' </summary>
    ''' <param name="pRow">ファイル取替票（高圧他）レコード</param>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(共通_添付書類詳細)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_TNPDOC_DTL_KAHK(ByVal pRow As DataRow,
                                          ByVal dao As MJ4102_DAO,
                                          ByVal pDataTable As DataTable, ByVal pExpUploadFileDir As String,
                                          ByVal k1 As String,
                                          ByVal k2 As String,
                                          ByVal k3 As String,
                                          ByVal k4 As String,
                                          ByRef pWarningFlg As Boolean) As Boolean
        Dim wdig4_zgsyo_cd As String = ""
        Dim wtrkhy_hakko_nendo As String = ""
        Dim wtrkhy_kbn As String = ""
        Dim wtrkhy_no As String = ""
        Dim wtrkhy_no_key As String = "" '取替票番号キー

        Dim wDataTable As DataTable = Nothing
        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        Dim isExist As Boolean = False
        Dim tblname As String = C_T_TRKE_TRKEHY_KAHK         '取替_取替票（高圧他）
        Dim field_tenp As String = "tenp_file_mng_no"        '添付ファイル管理番号 項目名
        Dim field_ryss As String = "ryssy_tenp_file_mng_no"  '領収書添付ファイル管理番号 項目名
        Dim jouken1 As String = ""
        Dim jouken2 As String = ""
        Dim jouken3 As String = ""
        Dim jouken4 As String = ""
        Dim wtenp_file_mng_no As String = ""          '添付ファイル管理番号
        Dim wryssy_tenp_file_mng_no As String = ""    '領収書添付ファイル管理番号
        Dim dir_tenp_file_mng_no As String = ""       '保存フォルダ用添付ファイル管理番号
        Dim dir_ryssy_tenp_file_mng_no As String = "" '保存フォルダ用領収書添付ファイル管理番号

        Dim wTnpDocKyes As HdSysBasePg.TnpDoc.TnpDocKeys = New HdSysBasePg.TnpDoc.TnpDocKeys
        Dim wTenpFileMngNo As Decimal = 0
        Dim wRyssyTenpFileMngNo As Decimal = 0
        Dim retval As Boolean = False

        Dim wPhotoKbn As String = ""
        Dim wPhotoSbt As String = ""
        Dim ht As Hashtable = New Hashtable
        Dim oYaFname As String = ""

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        'ファイル取替票（高圧他）レコードの添付書類管理番号情報を取得する
        '→添付書類管理番号のフォルダを作成するのだが、値なし(null or 0)のフォルダは"0"１文字とする
        dir_tenp_file_mng_no = getDataRowItemCStr(pRow, field_tenp, "0")
        dir_ryssy_tenp_file_mng_no = getDataRowItemCStr(pRow, field_ryss, "0")
        'ファイル取替票（高圧他）レコードの実費_No1～5の情報を取得する
        Dim wNo1_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_no"), ""))
        Dim wNo2_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_no"), ""))
        Dim wNo3_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_no"), ""))
        Dim wNo4_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no4_zippi_no"), ""))
        Dim wNo5_zippi_no As String = CStr(Me.myDB.ConvDbNull(pRow("no5_zippi_no"), ""))
        Dim wFile_renno As String = ""
        Dim wUpdZippiClm As String = ""
        Dim jippi_file_renno As Integer = 1 '添付書類詳細に登録する実費・領収書は1オリジンでカウントアップする

        '実費以外の写真用連番（共通部品を使用するので１～連番となる。ただし、庫出している場合は開始がずれる）
        Dim seq_file_renno As Integer = 0
        Dim Max_file_renno As String = ""
        dao.getMaxFileRennoKahk(Max_file_renno, k1, k2, k3, k4)
        seq_file_renno = MyCInt(Max_file_renno) + 1

        '共通_添付書類の処理を行う。
        'wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,file_renno")
        'file_rennoが文字型のため「1,10,2,...」の並びになる。Integer列を追加し数値ソートする
        wDataTable = pDataTable.Copy()
        wDataTable.Columns.Add("file_renno_int", GetType(Integer), "Convert(file_renno,'System.Int32')")
        wDataRows = wDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no,file_renno_int")

        '添付書類を登録する。
        For Each wRow As DataRow In wDataRows
            wdig4_zgsyo_cd = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            wtrkhy_hakko_nendo = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            wtrkhy_kbn = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            wtrkhy_no = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))
            wPhotoKbn = CStr(Me.myDB.ConvDbNull(wRow("no1_tnpdoc_hosok_naiyo"), ""))
            wPhotoSbt = CStr(Me.myDB.ConvDbNull(wRow("no2_tnpdoc_hosok_naiyo"), ""))
            wtrkhy_no_key = getTrkhyNoKey(wdig4_zgsyo_cd, wtrkhy_hakko_nendo, wtrkhy_kbn, wtrkhy_no)

            jouken1 = "dig4_zgsyo_cd =" & EcOrgIO.EcOrgString.GetSqlText(wdig4_zgsyo_cd)
            jouken2 = "trkhy_hakko_nendo =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_hakko_nendo)
            jouken3 = "trkhy_kbn =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_kbn)
            jouken4 = "trkhy_no =" & EcOrgIO.EcOrgString.GetSqlText(wtrkhy_no)

            'file_rennoとfile_msのハッシュテーブル登録
            ht.Add(CStr(Me.myDB.ConvDbNull(wRow("file_renno"), "")), CStr(Me.myDB.ConvDbNull(wRow("file_ms"), "")))

            '写真区分および写真種類をチェックし処理を分ける。
            If wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Jippi) Then
                '写真区分：実費の場合（実費の場合、写真種類はなし）
                '領収書添付書類管理番号が未取得の場合、取得する。
                wryssy_tenp_file_mng_no = myDB.GetRecValueString(tblname, field_ryss, jouken1, jouken2, jouken3, jouken4)
                If String.IsNullOrEmpty(wryssy_tenp_file_mng_no.Trim) Then
                    '領収書添付ファイル管理番号を新規取得する。
                    wTnpDocKyes.KEY1_VALUE = wtrkhy_no_key
                    retval = HdSysBasePg.TnpDoc.CreateTnpDocInfo(myDB, myUser, myPath, MjK1.Cmn.Const.TnpDocTypeCd.PhotoRyoushuu, wTnpDocKyes, wRyssyTenpFileMngNo)
                    If Not retval Then
                        Return False
                    End If
                    wryssy_tenp_file_mng_no = wRyssyTenpFileMngNo.ToString

                    '取替_取替票(高圧他)の領収書添付ファイル管理番号を更新する。
                    If Not dao.U015(wRow, wryssy_tenp_file_mng_no) Then
                        Return False
                    End If
                Else
                    '検索した領収書添付書類管理番号をDecimal変換する。
                    wRyssyTenpFileMngNo = MyCDec(wryssy_tenp_file_mng_no)
                End If

                retval = addTnpDocKeiki(seq_file_renno, wRow, wryssy_tenp_file_mng_no, pExpUploadFileDir, dir_ryssy_tenp_file_mng_no)
                If Not retval Then
                    Return False
                End If

                '実費・領収書を取替_取替票(高圧他)の実費No1～5に登録する
                wFile_renno = CStr(Me.myDB.ConvDbNull(wRow("file_renno"), ""))
                retval = False
                If wFile_renno.Equals(wNo1_zippi_no) Then
                    retval = dao.U0151(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo2_zippi_no) Then
                    retval = dao.U0152(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo3_zippi_no) Then
                    retval = dao.U0153(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo4_zippi_no) Then
                    retval = dao.U0154(wRow, CStr(jippi_file_renno))
                ElseIf wFile_renno.Equals(wNo5_zippi_no) Then
                    retval = dao.U0155(wRow, CStr(jippi_file_renno))
                End If
                If Not retval Then
                    Return False
                End If
                jippi_file_renno = jippi_file_renno + 1

            ElseIf wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Tekkyo) Or
                wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Torituke) Or
                wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Sign) Or
                wPhotoKbn.Equals(MjK1.Cmn.Const.KeikiPhotoKbn.Dige) Then
                '撤去・取付・サイン・代替の場合
                '添付書類管理番号を取得する。
                wtenp_file_mng_no = myDB.GetRecValueString(tblname, field_tenp, jouken1, jouken2, jouken3, jouken4)
                '添付ファイル管理番号未取得の場合、取得する。
                If String.IsNullOrEmpty(wtenp_file_mng_no.Trim) Then
                    '添付ファイル管理番号を新規取得する
                    wTnpDocKyes.KEY1_VALUE = wtrkhy_no_key
                    retval = HdSysBasePg.TnpDoc.CreateTnpDocInfo(myDB, myUser, myPath, MjK1.Cmn.Const.TnpDocTypeCd.PhotoKozi, wTnpDocKyes, wTenpFileMngNo)
                    If Not retval Then
                        Return False
                    End If
                    wtenp_file_mng_no = wTenpFileMngNo.ToString

                    '取替_取替票(高圧他)に更新する。
                    If Not dao.U014(wRow, wtenp_file_mng_no) Then
                        Return False
                    End If
                Else
                    '検索した添付ファイル管理番号をDecimal変換する。
                    wTenpFileMngNo = MyCDec(wtenp_file_mng_no)
                End If

                'サムネイルの場合、親画像を参照し作成する。
                If CStr(wRow("no3_tnpdoc_hosok_naiyo")).Equals(MjK1.Cmn.Const.UmuFlg.FlgOn) Then
                    'サムネイルの場合、親画像を参照し作成する。
                    oYaFname = CType(ht(CStr(Me.myDB.ConvDbNull(wRow("no4_tnpdoc_hosok_naiyo"), ""))), String)
                    retval = createThumbNailFile(oYaFname, wRow, wtenp_file_mng_no, pExpUploadFileDir, dir_tenp_file_mng_no)
                    If Not retval Then
                        Return False
                    End If
                End If

                '添付書類管理番号をキーに、画像を登録する。
                retval = addTnpDocKeiki(seq_file_renno, wRow, wtenp_file_mng_no, pExpUploadFileDir, dir_tenp_file_mng_no)
                If Not retval Then
                    Return False
                End If

                seq_file_renno = seq_file_renno + 1
            Else
                '庫出設定はアップロードの対象でないので取込なし
            End If
        Next

        Return True

    End Function
    'Rev002-End




#Region "テーブル更新（抜取検査）　取替票（高圧他）連係"
    ''' <summary>
    ''' テーブル更新（抜取検査）　取替票（高圧他）連係
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pExpUploadFileDir"></param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateDataNktrKnsOthersKahk(ByVal dao As MJ4102_DAO,
                                    ByVal pExpUploadFileDir As String,
                                    ByVal k1 As String,
                                    ByVal k2 As String,
                                    ByVal k3 As String,
                                    ByVal k4 As String,
                                    ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean

        Dim wFileTableMS As String = ""
        Dim wFileName As String = ""
        Dim wDtblUpload As DataTable = Nothing                      'アップロードDataTable
        Dim wDrUpload() As DataRow = Nothing                        'アップロードDataRow

        Try
            '取替票行程情報ファイルを読み込む。
            wFileTableMS = C_MS_TRKE_TRKEHY_PRC_KAHK + C_MS_NKTRKNS_KK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_TRKEHY_PRC_KAHK)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If
            '取替票行程（高圧他）情報テーブルを更新する(抜取検査)。
            If Not updateTRKE_TRKEHY_PRC_KAHK_NKTRKNS(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If



            '自主点検チェックシートファイルを読み込む。
            wFileTableMS = C_MS_TRKE_CHECK_SHEET_KAHK + C_MS_NKTRKNS_KK
            wFileName = System.IO.Path.Combine(pExpUploadFileDir, C_F_TRKE_CHECK_SHEET_KAHK)
            If Not HdSysBase.Cmn.ReadFile(wFileName, wDtblUpload) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル読込エラー：" & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If
            '自主点検チェックシートテーブルを更新する(抜取検査)。
            If Not updateTRKE_CHECK_SHEET_KAHK_NKTRKNS(dao, wDtblUpload, k1, k2, k3, k4, pWarningFlg) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, wFileTableMS & "ファイル取込エラー： " & wFileName)
                pWarningFlg = True
                pErrImportFlg = True
                Return False
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region



#Region "取替_取替票（高圧他）テーブル(抜取検査)"
    ''' <summary>
    ''' 取替_取替票（高圧他）テーブル(抜取検査)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pUpldKttmNo">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードユーザーID</param>
    ''' <param name="pExpUploadFileDir">アップロードファイル展開フォルダ</param>
    ''' <param name="pDataTable">アップロードデータ取替_取替票（低圧）</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <param name="pErrImportFlg">エラーフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_KAHK_NKTRKNS(ByVal dao As MJ4102_DAO,
                                                    ByVal pUpldKttmNo As String,
                                                    ByVal pKeikiDwldSbtCd As String,
                                                    ByVal pUpldUserId As String,
                                                    ByVal pExpUploadFileDir As String,
                                                    ByVal pDataTable As DataTable, ByRef pWarningFlg As Boolean, ByRef pErrImportFlg As Boolean) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票（高圧他）テーブルが更新可能かチェックする。
            Dim k1 As String = CStr(Me.myDB.ConvDbNull(wRow("dig4_zgsyo_cd"), ""))
            Dim k2 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_hakko_nendo"), ""))
            Dim k3 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_kbn"), ""))
            Dim k4 As String = CStr(Me.myDB.ConvDbNull(wRow("trkhy_no"), ""))
            Dim jouken1 As String = "dig4_zgsyo_cd = " & EcOrgIO.EcOrgString.GetSqlText(k1)
            Dim jouken2 As String = "trkhy_hakko_nendo = " & EcOrgIO.EcOrgString.GetSqlText(k2)
            Dim jouken3 As String = "trkhy_kbn = " & EcOrgIO.EcOrgString.GetSqlText(k3)
            Dim jouken4 As String = "trkhy_no = " & EcOrgIO.EcOrgString.GetSqlText(k4)
            Dim s1 As String = ""
            Dim s2 As String = ""
            Dim s3 As String = ""
            Dim s4 As String = ""
            Dim field1 As String = "prcmg_cd"
            Dim field2 As String = "nktr_kensa_zyokyo_cd"
            Dim field3 As String = "upd_date"
            Dim field4 As String = "nktrk_zissi_ymd"
            Dim retval As Boolean = False
            Dim nextProc As String = ""
            retval = myDB.GetRecValueStringMulti(C_T_TRKE_TRKEHY_KAHK,
                                                 field1, field2, field3, field4, "", "", "", "", "", "",
                                                 s1, s2, s3, s4, "", "", "", "", "", "",
                                                 jouken1, jouken2, jouken3, jouken4)
            If Not retval Then
                '取替票項目値の取得エラーの場合
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票検索エラーです(抜取検査):" + k1 + "-" + k2 + k3 + k4)
                pWarningFlg = True
                Continue For
            End If
            If String.IsNullOrEmpty(s1.Trim) Then
                '該当レコードがない場合
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "取替票がありません(抜取検査):" + k1 + "-" + k2 + k3 + k4)
                pWarningFlg = True
                Continue For
            End If

            '施工と異なり物理的NGではないのでスキップしない。2023/4/22
            ''抜取検査実施年月日が登録済みならスキップする
            'If Not String.IsNullOrEmpty(s4.Trim) Then
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "抜取検査実施年月日が登録済みです:" + k1 + "-" + k2 + k3 + k4)
            '    Continue For
            'End If

            'アップロードの排他確認は不要。多重ダウンロードが可能のため。
            ''排他確認　更新年月日が更新済みならスキップする
            'Dim fileUpdDate As String = getDateStrWithKeta(wRow("upd_date").ToString, C_KETA_YMDHMS)
            'Dim dbUpdDate As String = getDateStrWithKeta(s3, C_KETA_YMDHMS)
            'If fileUpdDate <= dbUpdDate Then
            '    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "更新年月日が更新済みです(抜取検査):" + k1 + "-" + k2 + k3 + k4)
            '    Continue For
            'End If


            '取替_取替票（高圧他）テーブルを更新する(抜取検査)。
            If Not dao.U018(wRow, pUpldUserId) Then
                Return False
            End If

            '更新した取替_取替票（高圧他）に連係するテーブルを取替票単位で更新する。
            '（更新対象の取替票についてのみ、連係するテーブルも更新する。）
            If Not updateDataNktrKnsOthersKahk(dao, pExpUploadFileDir, k1, k2, k3, k4, pWarningFlg, pErrImportFlg) Then
                Return False
            End If


            'ダウンロード取替票管理の持出ファイル状態コードを「5:アップロード済」に更新する。
            If Not dao.U011(wRow, pUpldKttmNo, pKeikiDwldSbtCd, pUpldUserId) Then
                Return False
            End If

        Next

        Return True

    End Function

#End Region




#Region "取替_取替票行程（高圧他）テーブル(抜取検査)(取替票単位)"
    ''' <summary>
    ''' 取替_取替票行程（高圧他）テーブル(抜取検査)(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_取替票行程)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_TRKEHY_PRC_KAHK_NKTRKNS(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_取替票行程（高圧他）テーブルを更新する。(抜取検査)
            If Not dao.U019(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region


#Region "取替_自主点検チェックシート（高圧他）テーブル(抜取検査)(取替票単位)"
    ''' <summary>
    ''' 取替_自主点検チェックシート（高圧他）テーブル(抜取検査)(取替票単位)
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ(取替_自主点検チェックシート)</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <param name="pWarningFlg">ワーニングフラグ</param>
    ''' <remarks></remarks>
    Private Function updateTRKE_CHECK_SHEET_KAHK_NKTRKNS(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable,
                                             ByVal k1 As String,
                                             ByVal k2 As String,
                                             ByVal k3 As String,
                                             ByVal k4 As String,
                                             ByRef pWarningFlg As Boolean) As Boolean

        Dim wDataRows() As DataRow = Nothing
        Dim whereStr As String = getTrkhyWhereStr(k1, k2, k3, k4)

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select(whereStr, "dig4_zgsyo_cd,trkhy_hakko_nendo,trkhy_kbn,trkhy_no")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '取替_自主点検チェックシート（高圧他）テーブルを更新する(抜取検査)。
            If Not dao.U020(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region
'Rev002-End

#Region "共通_ログ収集"
    ''' <summary>
    ''' 共通_ログ収集
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_LOG_SYUSYU(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "saksi_date, KTTNMT_LOG_MNG_RENNO")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '新規登録する。
            If Not dao.I006(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region

#Region "共通_ログイン認証履歴"
    ''' <summary>
    ''' 共通_ログイン認証履歴
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_LOGIN_NNSY_RK(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "saksi_date, login_nnsy_rireki_renno")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '新規登録する。
            If Not dao.I007(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region

#Region "共通_ログイン履歴データ"
    ''' <summary>
    ''' 共通_ログイン履歴データ
    ''' </summary>
    ''' <param name="dao">MJ4102_DAO</param>
    ''' <param name="pDataTable">アップロードデータ</param>
    ''' <remarks></remarks>
    Private Function updateCOM_LOGIN_RIREKI_DATA(ByVal dao As MJ4102_DAO, ByVal pDataTable As DataTable) As Boolean
        Dim wDataRows() As DataRow = Nothing

        '取得したデータ件数 = 0 の場合
        If (pDataTable.Rows.Count = 0) Then
            Return True
        End If

        wDataRows = pDataTable.Select("", "saksi_date, log_mng_renno")

        'アップロードデータのループ
        For Each wRow As DataRow In wDataRows
            '新規登録する。
            If Not dao.I008(wRow) Then
                Return False
            End If
        Next

        Return True

    End Function
#End Region



#Region "ファイル試しread"
    ''' <summary>
    ''' ファイル試しread
    ''' </summary>
    ''' <param name="pFileName">ファイル名</param>
    ''' <param name="pDtbl">データ設定領域</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function TryReadFile(ByVal pFileName As String, ByRef pDtbl As DataTable) As Boolean
        Try
            'まず暗号化ファイルとして読む
            If Not TryReadEncryptFile(pFileName, pDtbl) Then
                '通常ファイルとして読む
                If Not HdSysBase.Cmn.ReadFile(pFileName, pDtbl) Then
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 暗号化ファイル試しread
    ''' </summary>
    ''' <param name="pFileName">ファイル名</param>
    ''' <param name="pDtbl">データ設定領域</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function TryReadEncryptFile(ByVal pFileName As String, ByRef pDtbl As DataTable) As Boolean
        Try
            If Not HdSysBase.Cmn.ReadEncryptFile(pFileName, HdSysBase.Cmn.EncryptionKey, pDtbl) Then
                Return False
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "ファイル移動関数"
    ''' <summary>
    ''' ファイル移動関数
    ''' </summary>
    ''' <param name="pSourcePath">移動元</param>
    ''' <param name="pSearchPattern">pSourcePath 内のファイル名と対応させる検索文字列</param>
    ''' <param name="pDestinationPath">移動先</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function MoveAllFileInFolder(ByVal pSourcePath As String, ByVal pSearchPattern As String, ByVal pDestinationPath As String) As Boolean

        Dim wFileList() As String = Nothing     'ファイルリスト
        Dim wDestFilePath As String = ""        '移動先パス

        Try
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}メソッド開始", System.Reflection.MethodBase.GetCurrentMethod.Name))

            If Not System.IO.Directory.Exists(pSourcePath) Then
                Return True
            End If

            wFileList = System.IO.Directory.GetFiles(pSourcePath, pSearchPattern)
            'ファイルが存在しない場合もエラーとしない。
            If (wFileList Is Nothing) OrElse (wFileList.Count = 0) Then
                Return True
            End If

            If Not System.IO.Directory.Exists(pDestinationPath) Then
                System.IO.Directory.CreateDirectory(pDestinationPath)
            End If

            For Each wItm As String In wFileList
                wDestFilePath = System.IO.Path.Combine(pDestinationPath, System.IO.Path.GetFileName(wItm))

                '移動先フォルダの古いファイルを削除する。 
                If System.IO.File.Exists(wDestFilePath) Then
                    EcOrgIO.EcOrgFileIO.DeleteFile(wDestFilePath)
                End If

                'ファイルを移動する。
                System.IO.File.Move(wItm, wDestFilePath)

            Next

            Return True

        Catch ex As Exception
            'ログ出力】(HdSysBasePg.Cmn.InsertLog)を呼び出す。
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function
#End Region

#Region "ファイル名リスト取得関数"
    ''' <summary>
    ''' 指定したディレクトリ内の指定した検索パターンに一致するファイル名を返します。
    ''' </summary>
    ''' <param name="pPath">検索するディレクトリ</param>
    ''' <param name="pSearchPattern">path 内のファイル名と対応させる検索文字列</param>
    ''' <param name="pFileList">ファイル一覧</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function GetFiles(ByVal pPath As String, ByVal pSearchPattern As String, ByRef pFileList() As String) As Boolean

        Try
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}メソッド開始", System.Reflection.MethodBase.GetCurrentMethod.Name))

            If System.IO.Directory.Exists(pPath) Then
                pFileList = System.IO.Directory.GetFiles(pPath, pSearchPattern)
            End If

            Return True

        Catch ex As Exception
            'ログ出力】(HdSysBasePg.Cmn.InsertLog)を呼び出す。
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function
#End Region

#Region "共通部品"
    ''' <summary>
    ''' 文字列をDecimal型に変換する
    ''' </summary>
    ''' <param name="s">元文字列</param>
    ''' <returns>Decimal型値</returns>
    ''' <remarks></remarks>
    Function MyCDec(ByVal s As String) As Decimal
        If String.IsNullOrEmpty(s.Trim) Then
            Return 0
        Else
            Return CDec(s)
        End If
    End Function

    ''' <summary>
    ''' 文字列をInteger型に変換する
    ''' </summary>
    ''' <param name="s">元文字列</param>
    ''' <returns>Integer型値</returns>
    ''' <remarks></remarks>
    Function MyCInt(ByVal s As String) As Integer
        If String.IsNullOrEmpty(s.Trim) Then
            Return 0
        Else
            Return CInt(s)
        End If
    End Function

    ''' <summary>
    ''' 固定長文字列を取得する。
    ''' </summary>
    ''' <param name="s">元文字列</param>
    ''' <param name="keta">桁数</param>
    ''' <returns>固定長文字列('0'右埋め)</returns>
    ''' <remarks></remarks>
    Private Function getDateStrWithKeta(ByVal s As String, ByVal keta As Integer) As String
        Dim wkstr = New String("0"c, keta)
        Return Left(s.Trim + wkstr, keta)
    End Function
#End Region


End Class
