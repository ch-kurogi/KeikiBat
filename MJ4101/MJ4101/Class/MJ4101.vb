' **************************************************************************
' 機能名称：MJ4101
' 機能概要：ダウンロードファイル作成
' 使用方法：
' 前提条件：
' 更新履歴：2022.09.26 - EC Y.Ohnishi 新規作成
' 更新履歴：2023.10.31 - EC H.Morii   Rev002（２期対応）
'           2025.03.13 - EC Y.Ohnishi Rev040　託送情報切り替え対応
' **************************************************************************
Friend Class MJ4101

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
    Public myPID As String = ""                     'プロセスID

#End Region


#Region "（変数）クラス内"

    '-------------------------------------------------------------------------
    ' （変数）クラス内
    '-------------------------------------------------------------------------
    ''システム日時
    'Private _sysDateYmdh As String = ""

    ''業務処理日時
    'Private _syoriYmd As String = ""

    '持出用テンポラリフォルダ
    Private _mtdsTemp As String = ""
    Const C_MTDS_dir = "MTDS"

    'ファイル名定義
    '共通
    Const C_FC_MST_MSG As String = "MST_MSG.csv"                            '共通_メッセージマスタ
    Const C_FC_MST_KORYO As String = "MST_KORYO.csv"                        'マスタ_工量
    Const C_FC_COM_PARAMATER_CD_MST As String = "COM_PARAMATER_CD_MST.csv"  '共通_パラメータコードマスタ
    Const C_FC_MST_HNMK As String = "MST_HNMK.csv"                          'マスタ_品目
    Const C_FC_MST_KEIKI_KTSK As String = "MST_KEIKI_KTSK.csv"              'マスタ_計器型式
    Const C_FC_MST_TYAZA As String = "MST_TYAZA.csv"                        'マスタ_町字
    Const C_FC_COM_ZYOUSU_MST As String = "COM_ZYOUSU_MST.csv"              '定数マスタ
    Const C_FC_MST_USE_KANO_TUIKAKH As String = "MST_USE_KANO_TUIKAKH.csv"  'マスタ_利用可能追加工費
    Const C_FC_COM_YMD_MNG_MST As String = "COM_YMD_MNG_MST.csv"            '共通_日付管理マスタ

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

        Dim wWarningFlg As Boolean = False                  '警告有無フラグ
        Dim dbr As New HdPostgre.HdPostgreReader            'DB Reader
        Dim dao As New MJ4101_DAO(Me)           'DAO

        Dim wTempTopDir As String = ""
        Dim myPID As String = ""                       'プロセスID
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
        myPID = p.Id.ToString().PadLeft(10, "0"c)

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            ' 作業フォルダTOPを設定する
            wTempTopDir = System.IO.Path.Combine(myPath.BAT.K1TEMP, myPID)

            ' 持出用テンポラリフォルダ(myPath.BAT.K1TEMP & プロセスID & "MTDS")の設定
            Me._mtdsTemp = System.IO.Path.Combine(wTempTopDir, "MTDS")
            If Not System.IO.Directory.Exists(Me._mtdsTemp) Then
                System.IO.Directory.CreateDirectory(Me._mtdsTemp)
            End If

            '不要な日付は取得しない。mdlStart.vb内で取得した日付を使用する。
            ''HdSysbase.Cmn.Function.GetSystemYmdh を呼出し，取得した日時を，内部変数.システム日時に設定する。
            'Me._sysDateYmdh = HdSysBasePg.Cmn.GetSystemYmdh(Me.myDB)
            '
            ''業務処理日を取得する。
            'Me._syoriYmd = HdSysBasePg.Cmn.GetSyoriYmd(myDB)

            '共通の処理を行う。
            If Not Me.CreateMtdsFileCMN(dao) Then
                Dim msg As String = "共通ダウンロードファイル作成処理エラー"
                Dim exitCode As String = ""
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, msg)
                wWarningFlg = True

                'Zabbixにメッセージ送信する
                exitCode = MjK1.Cmn.Func.sendMsgZabbix(msg)
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "sendMsgZabbix(MJ4101)=" & exitCode)
            End If

            'zabbixメッセージ送信で監視に対応するため警告終了でのJP1対応は不要 2023/6/16
            ''警告終了をJP1に返却する
            'If wWarningFlg = True Then
            '    Return HdSysBase.Cmn.ExitCode.Warning
            'End If

            '処理を終了する。
            Return HdSysBase.Cmn.ExitCode.Normal

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))

            myDB.Rollback()

            Return HdSysBase.Cmn.ExitCode.Abnormal
        Finally
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


#Region "ダウンロードファイル作成（共通）"

    ''' <summary>
    ''' ダウンロードファイル作成（共通）
    ''' </summary>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMtdsFileCMN(ByVal dao As MJ4101_DAO) As Boolean
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
        Dim wMtdsFolder As String = ""              'NAS側持出フォルダ

        Dim wDataTable As System.Data.DataTable = Nothing   'データテーブル
        Dim wRow As System.Data.DataRow = Nothing

        Dim wFileName As String = ""                'ファイル名
        Dim wZipFileName As String = ""

        Dim firstFlag As Boolean = True

        Try
            'Functionの開始ログ作成
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L3, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("{0}を開始しました。", System.Reflection.MethodBase.GetCurrentMethod.Name))

            '持出用テンポラリフォルダを設定する。
            wTempMtdsFolder = Me._mtdsTemp

            'ファイルが存在する場合，フォルダを削除する。
            If System.IO.Directory.Exists(wTempMtdsFolder) Then
                System.IO.Directory.Delete(wTempMtdsFolder, True)
            End If
            '「持出管理一時フォルダ」フォルダを作成する。
            System.IO.Directory.CreateDirectory(wTempMtdsFolder)

            'メッセージファイルを作成する。
            If Not CreateCOM_MSG_MST(wTempMtdsFolder, dao) Then
                Return False
            End If

            ''工量マスタファイルを作成する。　対象外のため削除
            'If Not CreateMST_KORYO(wTempMtdsFolder, dao) Then
            '    Return False
            'End If

            'パラメータコードマスタファイルを作成する。
            If Not CreateCOM_PARAMATER_CD_MST(wTempMtdsFolder, dao) Then
                Return False
            End If

            '品目マスタファイルを作成する。
            If Not CreateMST_HNMK(wTempMtdsFolder, dao) Then
                Return False
            End If

            '計器型式マスタファイルを作成する。
            If Not CreateMST_KEIKI_KTSK(wTempMtdsFolder, dao) Then
                Return False
            End If

            '町字マスタファイルを作成する。
            If Not CreateMST_TYAZA(wTempMtdsFolder, dao) Then
                Return False
            End If

            '定数マスタファイルを作成する。
            If Not CreateCOM_ZYOUSU_MST(wTempMtdsFolder, dao) Then
                Return False
            End If

            'マスタ_利用可能追加工費ファイルを作成する。
            If Not CreateMST_USE_KANO_TUIKAKH(wTempMtdsFolder, dao) Then
                Return False
            End If

            'タブレットのシステム日付を使用するので不要
            ''共通_日付管理マスタファイルを作成する。
            'If Not CreateCOM_YMD_MNG_MST(wTempMtdsFolder, dao) Then
            '    Return False
            'End If

            '******************************************************
            '* ダウンロード用に作成したファイルを１つにまとめる。 *
            '******************************************************
            Dim wFname As String = "KEIKI_CMN"
            wZipFileName = System.IO.Path.Combine(wTempMtdsFolder, wFname + ".zip")
            '持出一時フォルダを圧縮する。
            HdSysBase.Cmn.CreateArchive(wZipFileName, wTempMtdsFolder)

            'HdSysBase.Cmn.EncryptFile を呼出し，ファイルを暗号化する。
            If Not HdSysBase.Cmn.EncryptFile(wZipFileName, HdSysBase.Cmn.EncryptionKey, wZipFileName) Then
            'ログ出力を行う。
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, "暗号化に失敗しました：" & wZipFileName)
                'Falseを返却して処理を終了する。
                Return False
            End If

            '* NASへのコピー処理
            '【myPath.NAS.K1DAT】/Cmn/MST/
            wMtdsFolder = System.IO.Path.Combine(Me.myPath.NAS.K1DAT, "Cmn", "MST")

            If Not System.IO.Directory.Exists(wMtdsFolder) Then
                '「共通の持出フォルダ」がない場合，フォルダを作成する。
                System.IO.Directory.CreateDirectory(wMtdsFolder)
            End If

            '持出圧縮ファイルを，コピーする。
            EcOrgIO.EcOrgFileIO.CopyFile(wZipFileName, System.IO.Path.Combine(wMtdsFolder, System.IO.Path.GetFileName(wZipFileName)))

            '持出圧縮ファイルを削除する
            If System.IO.File.Exists(wZipFileName) Then
                System.IO.File.Delete(wZipFileName)
            End If
            '持出一時フォルダを削除する。
            If System.IO.Directory.Exists(wTempMtdsFolder) Then
                System.IO.Directory.Delete(wTempMtdsFolder, True)
            End If


            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.KinoCD, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region


#Region "メッセージファイル作成処理"

    ''' <summary>
    ''' メッセージファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateCOM_MSG_MST(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            'メッセージを取得する。
            If Not dao.S001(dbr) Then
                Throw New Exception(String.Format("メッセージ検索でエラーが発生しました。"))
            End If

            'メッセージファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblCOM_MSG_MST(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("msg_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("msg_id"), ""))
                wRow("msg_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("msg_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_MSG)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("メッセージファイル出力に失敗しました。"))
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

    Sub initDtblCOM_MSG_MST(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("msg_id")
        wDestDt.Columns.Add("msg_naiyo")
    End Sub
#End Region

#Region "マスタ_工量ファイル作成処理"

    ''' <summary>
    ''' マスタ_工量ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMST_KORYO(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            'マスタ_工量を取得する。
            If Not dao.S002(dbr, SyoriYMD) Then
                Throw New Exception(String.Format("マスタ_工量検索でエラーが発生しました。"))
            End If

            'マスタ_工量ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblMST_KORYO(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("koryo_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_cd"), ""))
                wRow("tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_start_ymd"), ""))
                wRow("tky_end_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_end_ymd"), ""))
                wRow("trtk_koryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_koryo"), ""))
                wRow("tekyo_koryo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_koryo"), ""))
                wRow("koryo_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_tanka_kngk"), ""))
                wRow("kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kohi_tanka_kngk"), ""))
                wRow("koryo_kohi_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_kohi_sbt_cd"), ""))
                wRow("koryo_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_mskn"), ""))
                wRow("koryo_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_ms"), ""))
                wRow("koryo_tanka_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_tanka_cd"), ""))
                wRow("wrms_ritu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("wrms_ritu"), ""))
                wRow("trtk_kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_kohi_tanka_kngk"), ""))
                wRow("tekyo_kohi_tanka_kngk") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_kohi_tanka_kngk"), ""))
                wRow("koryo_idou_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_idou_kbn"), ""))
                wRow("koryo_cd_tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_cd_tky_start_ymd"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_KORYO)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("マスタ_工量ファイル出力に失敗しました。"))
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

    Sub initDtblMST_KORYO(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("koryo_cd")
        wDestDt.Columns.Add("tky_start_ymd")
        wDestDt.Columns.Add("tky_end_ymd")
        wDestDt.Columns.Add("trtk_koryo")
        wDestDt.Columns.Add("tekyo_koryo")
        wDestDt.Columns.Add("koryo_tanka_kngk")
        wDestDt.Columns.Add("kohi_tanka_kngk")
        wDestDt.Columns.Add("koryo_kohi_sbt_cd")
        wDestDt.Columns.Add("koryo_mskn")
        wDestDt.Columns.Add("koryo_ms")
        wDestDt.Columns.Add("koryo_tanka_cd")
        wDestDt.Columns.Add("wrms_ritu")
        wDestDt.Columns.Add("trtk_kohi_tanka_kngk")
        wDestDt.Columns.Add("tekyo_kohi_tanka_kngk")
        wDestDt.Columns.Add("koryo_idou_kbn")
        wDestDt.Columns.Add("koryo_cd_tky_start_ymd")
    End Sub
#End Region

#Region "パラメータコードマスタファイル作成処理"

    ''' <summary>
    ''' パラメータコードマスタファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateCOM_PARAMATER_CD_MST(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            'パラメータコードマスタを取得する。
            If Not dao.S003(dbr) Then
                Throw New Exception(String.Format("パラメータコードマスタ検索でエラーが発生しました。"))
            End If

            'パラメータコードマスタファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblCOM_PARAMATER_CD_MST(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("prm_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_id"), ""))
                wRow("prm_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd"), ""))
                wRow("prm_cd_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd_ms"), ""))
                wRow("prm_cd_abms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_cd_abms"), ""))
                wRow("str1_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str1_hksu_value"), ""))
                wRow("str2_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str2_hksu_value"), ""))
                wRow("value1_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_hksu_value"), ""))
                wRow("value2_hksu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value2_hksu_value"), ""))
                wRow("prm_hyz_zyun") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("prm_hyz_zyun"), ""))
                wRow("HYZ_FLG") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("HYZ_FLG"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_COM_PARAMATER_CD_MST)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("パラメータコードマスタファイル出力に失敗しました。"))
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

    Sub initDtblCOM_PARAMATER_CD_MST(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("prm_id")
        wDestDt.Columns.Add("prm_cd")
        wDestDt.Columns.Add("prm_cd_ms")
        wDestDt.Columns.Add("prm_cd_abms")
        wDestDt.Columns.Add("str1_hksu_value")
        wDestDt.Columns.Add("str2_hksu_value")
        wDestDt.Columns.Add("value1_hksu_value")
        wDestDt.Columns.Add("value2_hksu_value")
        wDestDt.Columns.Add("prm_hyz_zyun")
        wDestDt.Columns.Add("HYZ_FLG")
    End Sub
#End Region

#Region "品目マスタファイル作成処理"

    ''' <summary>
    ''' 品目マスタファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMST_HNMK(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            '品目マスタを取得する。
            If Not dao.S004(dbr) Then
                Throw New Exception(String.Format("品目マスタ検索でエラーが発生しました。"))
            End If

            '品目マスタファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblMST_HNMK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("hnmk_nnsk_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hnmk_nnsk_kbn"), ""))
                wRow("krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krhsk_cd"), ""))
                wRow("keiki_yr_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_yr_cd"), ""))
                wRow("keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_ktsk_ms"), ""))
                wRow("mado_su") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mado_su"), ""))
                wRow("ts_ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_ktsk_cd"), ""))
                wRow("dnzsk_maker_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("dnzsk_maker_cd"), ""))
                wRow("keiki_htknt_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_htknt_kbn"), ""))
                wRow("hzk_ziry_hnbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzk_ziry_hnbt_cd"), ""))
                wRow("hzk_ziry_hts") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hzk_ziry_hts"), ""))
                wRow("cktrm_d") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cktrm_d"), ""))
                wRow("huing_color_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("huing_color_kbn"), ""))
                wRow("tnsdai_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tnsdai_umu_flg"), ""))
                wRow("sm_tshsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_tshsk_cd"), ""))
                wRow("tky_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tky_start_ymd"), ""))
                wRow("hnmk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("hnmk_cd"), ""))
                wRow("keiki_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sbt_cd"), ""))
                wRow("keiki_hryhn_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_hryhn_kbn"), ""))
                wRow("ryohu_kyki_yysu") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ryohu_kyki_yysu"), ""))
                wRow("yosyrhn_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yosyrhn_cd"), ""))
                wRow("yuko_kigen_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_kigen_ymd"), ""))
                wRow("mof_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mof_kbn"), ""))
                wRow("sm_taiko_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_taiko_kbn"), ""))
                wRow("sm_knthk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sm_knthk_cd"), ""))
                wRow("ts_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_kino_umu_flg"), ""))
                wRow("khk_kino_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("khk_kino_umu_flg"), ""))
                wRow("gbdg_umu_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("gbdg_umu_flg"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_HNMK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("品目マスタファイル出力に失敗しました。"))
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

    Sub initDtblMST_HNMK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("hnmk_nnsk_kbn")
        wDestDt.Columns.Add("krhsk_cd")
        wDestDt.Columns.Add("keiki_yr_cd")
        wDestDt.Columns.Add("keiki_ktsk_ms")
        wDestDt.Columns.Add("mado_su")
        wDestDt.Columns.Add("ts_ktsk_cd")
        wDestDt.Columns.Add("dnzsk_maker_cd")
        wDestDt.Columns.Add("keiki_htknt_kbn")
        wDestDt.Columns.Add("hzk_ziry_hnbt_cd")
        wDestDt.Columns.Add("hzk_ziry_hts")
        wDestDt.Columns.Add("cktrm_d")
        wDestDt.Columns.Add("huing_color_kbn")
        wDestDt.Columns.Add("tnsdai_umu_flg")
        wDestDt.Columns.Add("sm_tshsk_cd")
        wDestDt.Columns.Add("tky_start_ymd")
        wDestDt.Columns.Add("hnmk_cd")
        wDestDt.Columns.Add("keiki_sbt_cd")
        wDestDt.Columns.Add("keiki_hryhn_kbn")
        wDestDt.Columns.Add("ryohu_kyki_yysu")
        wDestDt.Columns.Add("yosyrhn_cd")
        wDestDt.Columns.Add("yuko_kigen_ymd")
        wDestDt.Columns.Add("mof_kbn")
        wDestDt.Columns.Add("sm_taiko_kbn")
        wDestDt.Columns.Add("sm_knthk_cd")
        wDestDt.Columns.Add("ts_kino_umu_flg")
        wDestDt.Columns.Add("khk_kino_umu_flg")
        wDestDt.Columns.Add("gbdg_umu_flg")
    End Sub
#End Region

#Region "計器型式マスタファイル作成処理"

    ''' <summary>
    ''' 計器型式マスタファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMST_KEIKI_KTSK(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            '計器型式マスタを取得する。
            If Not dao.S005(dbr) Then
                Throw New Exception(String.Format("計器型式マスタ検索でエラーが発生しました。"))
            End If

            '計器型式マスタファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblMST_KEIKI_KTSK(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("kotea_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kotea_kbn"), ""))
                wRow("ktsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ktsk_cd"), ""))
                wRow("ksyu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ksyu_cd"), ""))
                wRow("krhsk_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("krhsk_cd"), ""))
                wRow("ts_kotea_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("ts_kotea_kbn"), ""))
                wRow("keiki_sbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_sbt_cd"), ""))
                wRow("keiki_ktsk_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_ktsk_ms"), ""))
                wRow("nozok_keiki_reuse_hnti_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nozok_keiki_reuse_hnti_cd"), ""))
                wRow("tk_keiki_reuse_hnti_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tk_keiki_reuse_hnti_cd"), ""))
                wRow("nozok_ts_trtk_krksbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("nozok_ts_trtk_krksbt_cd"), ""))
                wRow("tk_ts_trtk_krksbt_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tk_ts_trtk_krksbt_cd"), ""))
                'Rev040 2025/03/13 物理分割 計器マスタ連係対応 ADD Start
                wRow("keiki_yr_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("keiki_yr_cd"), ""))
                wRow("yr_cd_kbn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yr_cd_kbn"), ""))
                'Rev040 2025/03/13 物理分割 計器マスタ連係対応 ADD End

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_KEIKI_KTSK)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("計器型式マスタファイル出力に失敗しました。"))
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

    Sub initDtblMST_KEIKI_KTSK(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("kotea_kbn")
        wDestDt.Columns.Add("ktsk_cd")
        wDestDt.Columns.Add("ksyu_cd")
        wDestDt.Columns.Add("krhsk_cd")
        wDestDt.Columns.Add("ts_kotea_kbn")
        wDestDt.Columns.Add("keiki_sbt_cd")
        wDestDt.Columns.Add("keiki_ktsk_ms")
        wDestDt.Columns.Add("nozok_keiki_reuse_hnti_cd")
        wDestDt.Columns.Add("tk_keiki_reuse_hnti_cd")
        wDestDt.Columns.Add("nozok_ts_trtk_krksbt_cd")
        wDestDt.Columns.Add("tk_ts_trtk_krksbt_cd")
        'Rev040 2025/03/13 物理分割 計器マスタ連係対応 ADD Start
        wDestDt.Columns.Add("keiki_yr_cd")
        wDestDt.Columns.Add("yr_cd_kbn")
        'Rev040 2025/03/13 物理分割 計器マスタ連係対応 ADD End
    End Sub
#End Region

#Region "マスタ_町字ファイル作成処理"

    ''' <summary>
    ''' マスタ_町字ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMST_TYAZA(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            'マスタ_町字を取得する。
            If Not dao.S006(dbr) Then
                Throw New Exception(String.Format("マスタ_町字検索でエラーが発生しました。"))
            End If

            'マスタ_町字ファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblMST_TYAZA(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("tdhkn_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_add_cd"), ""))
                wRow("siku_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_add_cd"), ""))
                wRow("oazat_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_add_cd"), ""))
                wRow("azatm_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_add_cd"), ""))
                wRow("tdhkn_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_mskn"), ""))
                wRow("siku_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_mskn"), ""))
                wRow("oazat_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_mskn"), ""))
                wRow("azatm_mskn") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_mskn"), ""))
                wRow("tdhkn_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tdhkn_ms"), ""))
                wRow("siku_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("siku_ms"), ""))
                wRow("oazat_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oazat_ms"), ""))
                wRow("azatm_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("azatm_ms"), ""))
                wRow("add_zipcd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("add_zipcd"), ""))
                wRow("new_add_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("new_add_cd"), ""))
                wRow("sikou_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sikou_ym"), ""))
                wRow("haisi_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("haisi_ym"), ""))
                wRow("new_add_cd_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("new_add_cd_ym"), ""))
                wRow("name_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("name_henko_ym"), ""))
                wRow("zipcd_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zipcd_henko_ym"), ""))
                wRow("tino_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tino_henko_ym"), ""))
                wRow("barcd_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("barcd_naiyo"), ""))
                wRow("oyako_knk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oyako_knk_naiyo"), ""))
                wRow("cs_barcd_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("cs_barcd_henko_ym"), ""))
                wRow("oyako_knk_henko_ym") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("oyako_knk_henko_ym"), ""))
                wRow("tusyo_flg") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tusyo_flg"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_TYAZA)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("マスタ_町字ファイル出力に失敗しました。"))
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

    Sub initDtblMST_TYAZA(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("tdhkn_add_cd")
        wDestDt.Columns.Add("siku_add_cd")
        wDestDt.Columns.Add("oazat_add_cd")
        wDestDt.Columns.Add("azatm_add_cd")
        wDestDt.Columns.Add("tdhkn_mskn")
        wDestDt.Columns.Add("siku_mskn")
        wDestDt.Columns.Add("oazat_mskn")
        wDestDt.Columns.Add("azatm_mskn")
        wDestDt.Columns.Add("tdhkn_ms")
        wDestDt.Columns.Add("siku_ms")
        wDestDt.Columns.Add("oazat_ms")
        wDestDt.Columns.Add("azatm_ms")
        wDestDt.Columns.Add("add_zipcd")
        wDestDt.Columns.Add("new_add_cd")
        wDestDt.Columns.Add("sikou_ym")
        wDestDt.Columns.Add("haisi_ym")
        wDestDt.Columns.Add("new_add_cd_ym")
        wDestDt.Columns.Add("name_henko_ym")
        wDestDt.Columns.Add("zipcd_henko_ym")
        wDestDt.Columns.Add("tino_henko_ym")
        wDestDt.Columns.Add("barcd_naiyo")
        wDestDt.Columns.Add("oyako_knk_naiyo")
        wDestDt.Columns.Add("cs_barcd_henko_ym")
        wDestDt.Columns.Add("oyako_knk_henko_ym")
        wDestDt.Columns.Add("tusyo_flg")
    End Sub
#End Region

#Region "定数マスタファイル作成処理"

    ''' <summary>
    ''' 定数マスタファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateCOM_ZYOUSU_MST(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            '定数マスタを取得する。
            If Not dao.S007(dbr) Then
                Throw New Exception(String.Format("定数マスタ検索でエラーが発生しました。"))
            End If

            'メッセージファイル用Datatableの初期化
            wDataTable = New DataTable
            initDtblCOM_ZYOUSU_MST(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("zyousu_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("zyousu_cd"), ""))
                wRow("yuko_start_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_start_ymd"), ""))
                wRow("yuko_end_ymd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("yuko_end_ymd"), ""))
                wRow("str1_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str1_zyousu_value"), ""))
                wRow("str2_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("str2_zyousu_value"), ""))
                wRow("value1_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_zyousu_value"), ""))
                wRow("value2_zyousu_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value2_zyousu_value"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_COM_ZYOUSU_MST)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("定数マスタファイル出力に失敗しました。"))
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

    Sub initDtblCOM_ZYOUSU_MST(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("zyousu_cd")
        wDestDt.Columns.Add("yuko_start_ymd")
        wDestDt.Columns.Add("yuko_end_ymd")
        wDestDt.Columns.Add("str1_zyousu_value")
        wDestDt.Columns.Add("str2_zyousu_value")
        wDestDt.Columns.Add("value1_zyousu_value")
        wDestDt.Columns.Add("value2_zyousu_value")
    End Sub
#End Region

#Region "マスタ_利用可能追加工費ファイル作成処理"

    ''' <summary>
    ''' マスタ_利用可能追加工費ファイル作成処理
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateMST_USE_KANO_TUIKAKH(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            'マスタ_利用可能追加工費を取得する。
            If Not dao.S008(dbr, SyoriYMD) Then
                Throw New Exception(String.Format("マスタ_利用可能追加工費検索でエラーが発生しました。"))
            End If

            'Datatableの初期化
            wDataTable = New DataTable
            initDtblMST_USE_KANO_TUIKAKH(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("koryo_zunit_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_zunit_cd"), ""))
                wRow("TUIKAKH_SYTK_MST_KBN") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("TUIKAKH_SYTK_MST_KBN"), ""))
                wRow("tekyo_suryo_uplim_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_suryo_uplim_value"), ""))
                wRow("trtk_suryo_uplim_value") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_suryo_uplim_value"), ""))
                wRow("koryo_zunit_cd_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("koryo_zunit_cd_ms"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_MST_USE_KANO_TUIKAKH)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("マスタ_利用可能追加工費ファイル出力に失敗しました。"))
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

    Sub initDtblMST_USE_KANO_TUIKAKH(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("koryo_zunit_cd")
        wDestDt.Columns.Add("TUIKAKH_SYTK_MST_KBN")
        wDestDt.Columns.Add("tekyo_suryo_uplim_value")
        wDestDt.Columns.Add("trtk_suryo_uplim_value")
        wDestDt.Columns.Add("koryo_zunit_cd_ms")
    End Sub
#End Region

#Region "共通_日付管理マスタ"

    ''' <summary>
    ''' 共通_日付管理マスタ
    ''' </summary>
    ''' <param name="pMtdsMngNoFolder">持出単位フォルダパス</param>
    ''' <param name="dao">MJ4101_DAO</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Private Function CreateCOM_YMD_MNG_MST(ByVal pMtdsMngNoFolder As String, ByVal dao As MJ4101_DAO) As Boolean

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

            '共通_日付管理マスタを取得する。
            If Not dao.S009(dbr) Then
                Throw New Exception(String.Format("共通_日付管理マスタ検索でエラーが発生しました。"))
            End If

            'Datatableの初期化
            wDataTable = New DataTable
            initDtblCOM_YMD_MNG_MST(wDataTable)

            'Datatableの初期化
            wDataTable.Rows.Clear()

            'FETCH
            While (dbr.DataStruct.Read)
                '持出対象データを設定する
                wRow = wDataTable.NewRow

                wRow("saksi_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("saksi_date"), ""))
                wRow("upd_date") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("upd_date"), ""))
                wRow("sousa_user_id") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_user_id"), ""))
                wRow("sousa_appli_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("sousa_appli_cd"), ""))
                wRow("mst_mng_cd") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mst_mng_cd"), ""))
                wRow("mst_mng_ms") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mst_mng_ms"), ""))
                wRow("mst_mng_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mst_mng_naiyo"), ""))
                wRow("mst_bk_naiyo") = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("mst_bk_naiyo"), ""))

                wDataTable.Rows.Add(wRow)
            End While

            'HdSysBase.Cmn.WriteFile を呼出し，ファイルを出力する。
            wFileName = System.IO.Path.Combine(pMtdsMngNoFolder, C_FC_COM_YMD_MNG_MST)
            If Not HdSysBase.Cmn.WriteFile(wFileName, wDataTable) Then
                Throw New Exception(String.Format("共通_日付管理マスタファイル出力に失敗しました。"))
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

    Sub initDtblCOM_YMD_MNG_MST(ByVal wDestDt As System.Data.DataTable)
        ' 空テーブルを作る
        wDestDt.Columns.Add("saksi_date")
        wDestDt.Columns.Add("upd_date")
        wDestDt.Columns.Add("sousa_user_id")
        wDestDt.Columns.Add("sousa_appli_cd")
        wDestDt.Columns.Add("mst_mng_cd")
        wDestDt.Columns.Add("mst_mng_ms")
        wDestDt.Columns.Add("mst_mng_naiyo")
        wDestDt.Columns.Add("mst_bk_naiyo")
    End Sub
#End Region


End Class
