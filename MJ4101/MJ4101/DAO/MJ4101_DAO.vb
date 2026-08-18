''' <summary>
''' 特定巡視ダウンロードデータ作成（防護管）データベース処理
''' </summary>
Friend Class MJ4101_DAO

#Region "内部パラメータ"

    Private myDB As HdPostgre.HdPostgreDb       'DB接続
    Private myUser As HdSysBasePg.Cmn.UserInfo  'ユーザ情報
    Private myPath As HdSysBasePg.Cmn.PathInfo  'パス情報
    Private Kino_Cd As String = ""              '機能コード

    '処理用定義
    Const C_SELECT_MAX_REC As String = "100"    '周期最大処理件数

#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="pBat">特定巡視ダウンロードデータ作成（防護管）バッチ</param>
    Public Sub New(ByRef pBat As MJ4101)

        Me.myDB = pBat.myDB
        Me.myUser = pBat.myUser
        Me.myPath = pBat.myPath
        Me.Kino_Cd = pBat.KinoCD

    End Sub

#End Region


#Region "(SQL)共通_メッセージマスタ　検索"
    ''' <summary>
    ''' 共通_メッセージマスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>共通_メッセージマスタを検索します</remarks>
    Friend Function S001(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,msg_id
 ,msg_naiyo
  FROM COM_MSG_MST 
ORDER BY 
  msg_id
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)マスタ_工量　検索"
    ''' <summary>
    ''' マスタ_工量　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pSyoriYMD">処理年月日</param>
    ''' <remarks>マスタ_工量を検索します</remarks>
    Friend Function S002(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pSyoriYMD As String) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,koryo_cd
 ,tky_start_ymd
 ,tky_end_ymd
 ,trtk_koryo
 ,tekyo_koryo
 ,koryo_tanka_kngk
 ,kohi_tanka_kngk
 ,koryo_kohi_sbt_cd
 ,koryo_mskn
 ,koryo_ms
 ,koryo_tanka_cd
 ,wrms_ritu
 ,trtk_kohi_tanka_kngk
 ,tekyo_kohi_tanka_kngk
 ,koryo_idou_kbn
 ,koryo_cd_tky_start_ymd
  FROM MST_KORYO 
  WHERE 1=1
    and tky_start_ymd <= {0}
    and tky_end_ymd >= {0}
 order by koryo_cd
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString, EcOrgIO.EcOrgString.GetSqlText(pSyoriYMD))
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)パラメータコードマスタ　検索"
    ''' <summary>
    ''' パラメータコードマスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>パラメータコードマスタを検索します</remarks>
    Friend Function S003(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            'Rev044 SJIS変換問題をUTF8化で対応・除外終了 /*'Rev002 外字を含む２期用データの除外　and prm_id <> 'KYSBT_CD'*/
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,prm_id
 ,prm_cd
 ,prm_cd_ms
 ,prm_cd_abms
 ,str1_hksu_value
 ,str2_hksu_value
 ,value1_hksu_value
 ,value2_hksu_value
 ,prm_hyz_zyun
 ,HYZ_FLG
  FROM COM_PARAMATER_CD_MST
 WHERE 1=1
  and HYZ_FLG='1'
 order by prm_id,prm_hyz_zyun
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)品目マスタ　検索"
    ''' <summary>
    ''' 品目マスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>品目マスタを検索します</remarks>
    Friend Function S004(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,hnmk_nnsk_kbn
 ,krhsk_cd
 ,keiki_yr_cd
 ,keiki_ktsk_ms
 ,mado_su
 ,ts_ktsk_cd
 ,dnzsk_maker_cd
 ,keiki_htknt_kbn
 ,hzk_ziry_hnbt_cd
 ,hzk_ziry_hts
 ,cktrm_d
 ,huing_color_kbn
 ,tnsdai_umu_flg
 ,sm_tshsk_cd
 ,tky_start_ymd
 ,hnmk_cd
 ,keiki_sbt_cd
 ,keiki_hryhn_kbn
 ,ryohu_kyki_yysu
 ,yosyrhn_cd
 ,yuko_kigen_ymd
 ,mof_kbn
 ,sm_taiko_kbn
 ,sm_knthk_cd
 ,ts_kino_umu_flg
 ,khk_kino_umu_flg
 ,gbdg_umu_flg

  FROM MST_HNMK 
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)計器型式マスタ　検索"
    ''' <summary>
    ''' 計器型式マスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>計器型式マスタを検索します</remarks>
    Friend Function S005(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            'Rev040 2025/03/13 物理分割 計器マスタ連係対応 項目追加 keiki_yr_cd,yr_cd_kbn

            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,kotea_kbn
 ,ktsk_cd
 ,ksyu_cd
 ,krhsk_cd
 ,ts_kotea_kbn
 ,keiki_sbt_cd
 ,keiki_ktsk_ms
 ,nozok_keiki_reuse_hnti_cd
 ,tk_keiki_reuse_hnti_cd
 ,nozok_ts_trtk_krksbt_cd
 ,tk_ts_trtk_krksbt_cd
 ,keiki_yr_cd
 ,yr_cd_kbn
 FROM MST_KEIKI_KTSK 
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)マスタ_町字　検索"
    ''' <summary>
    ''' マスタ_町字　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>マスタ_町字を検索します</remarks>
    Friend Function S006(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        '対象：中国５県＋香川＋愛媛＋兵庫

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,tdhkn_add_cd
 ,siku_add_cd
 ,oazat_add_cd
 ,azatm_add_cd
 ,' ' as tdhkn_mskn
 ,' ' as siku_mskn
 ,' ' as oazat_mskn
 ,' ' as azatm_mskn
 ,tdhkn_ms
 ,siku_ms
 ,oazat_ms
 ,azatm_ms
 ,add_zipcd
 ,new_add_cd
 ,sikou_ym
 ,haisi_ym
 ,new_add_cd_ym
 ,name_henko_ym
 ,zipcd_henko_ym
 ,tino_henko_ym
 ,barcd_naiyo
 ,oyako_knk_naiyo
 ,cs_barcd_henko_ym
 ,oyako_knk_henko_ym
 ,tusyo_flg
  FROM MST_TYAZA 
  WHERE 1=1
  AND tdhkn_add_cd IN ('28','31','32' ,'33' ,'34' ,'35' ,'37' ,'38')
  AND haisi_ym = ' '
  order by tdhkn_add_cd,siku_add_cd,oazat_add_cd,azatm_add_cd
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)定数マスタ　検索"
    ''' <summary>
    ''' 定数マスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>定数マスタを検索します</remarks>
    Friend Function S007(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,zyousu_cd
 ,yuko_start_ymd
 ,yuko_end_ymd
 ,str1_zyousu_value
 ,str2_zyousu_value
 ,value1_zyousu_value
 ,value2_zyousu_value
  FROM COM_ZYOUSU_MST 
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)マスタ_利用可能追加工費　検索"
    ''' <summary>
    ''' マスタ_利用可能追加工費　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pSyoriYMD">処理年月日</param>
    ''' <remarks>マスタ_利用可能追加工費を検索します</remarks>
    Friend Function S008(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pSyoriYMD As String) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  M1.saksi_date
 ,M1.upd_date
 ,M1.sousa_user_id
 ,M1.sousa_appli_cd
 ,M1.koryo_zunit_cd
 ,M1.TUIKAKH_SYTK_MST_KBN
 ,M1.tekyo_suryo_uplim_value
 ,M1.trtk_suryo_uplim_value
 ,M2.koryo_cd
 ,M2.koryo_ms
 ,M3.zunit_cd
 ,M3.zunit_ms
 ,case M1.tuikakh_sytk_mst_kbn
    when '1' then M2.koryo_ms
    when '2' then M3.zunit_ms
    else ' ' 
  END As koryo_zunit_cd_ms
FROM MST_USE_KANO_TUIKAKH M1
--マスター工量
LEFT  JOIN  mst_koryo M2 ON (M2.koryo_cd=M1.koryo_zunit_cd and M2.tky_start_ymd<={0} and {0}<=M2.tky_end_ymd)
--マスター材料ユニット
LEFT  JOIN  mst_zunit M3 ON (M3.zunit_cd=M1.koryo_zunit_cd and M3.tky_start_ymd<={0} and {0}<=M3.tky_end_ymd)
ORDER BY M1.TUIKAKH_SYTK_MST_KBN,M1.koryo_zunit_cd
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString, EcOrgIO.EcOrgString.GetSqlText(pSyoriYMD))
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region

#Region "(SQL)共通_日付管理マスタ　検索"
    ''' <summary>
    ''' 共通_日付管理マスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>共通_日付管理マスタを検索します</remarks>
    Friend Function S009(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
  M1.saksi_date
 ,M1.upd_date
 ,M1.sousa_user_id
 ,M1.sousa_appli_cd
 ,M1.mst_mng_cd
 ,M1.mst_mng_ms
 ,M1.mst_mng_naiyo
 ,M1.mst_bk_naiyo
 FROM COM_YMD_MNG_MST M1
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString)
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(dbr.ErrDescription)
            End If

            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        Finally
        End Try
    End Function
#End Region




#Region "インテリセンス非表示"
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Public Shadows Function Equals() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Public Shadows Function GetHashCode() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Public Shadows Function [GetType]() As Type
        Return GetType(String)
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Public Shadows Function ReferenceEquals() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Public Shadows Function ToString() As Object
        Return New Object
    End Function
#End Region

End Class
