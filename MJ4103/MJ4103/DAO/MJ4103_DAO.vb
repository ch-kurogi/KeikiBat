''' <summary>
''' 計器業務ダウンロードデータ作成データベース処理
''' </summary>
Friend Class MJ4103_DAO

#Region "内部パラメータ"

    Private myDB As HdPostgre.HdPostgreDb       'DB接続
    Private myUser As HdSysBasePg.Cmn.UserInfo  'ユーザ情報
    Private myPath As HdSysBasePg.Cmn.PathInfo  'パス情報
    Private Kino_Cd As String = ""              '機能コード
    Private SyoriYMD As String = ""             '処理年月日
#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="pBat">特定巡視ダウンロードデータ作成（防護管）バッチ</param>
    Public Sub New(ByRef pBat As MJ4103)

        Me.myDB = pBat.myDB
        Me.myUser = pBat.myUser
        Me.myPath = pBat.myPath
        Me.Kino_Cd = pBat.KinoCD
        Me.SyoriYMD = pBat.SyoriYMD

    End Sub

#End Region

#Region "(SQL)ダウンロード管理　検索"
    ''' <summary>
    ''' ダウンロード管理　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pKeikiDwldSbtCd">計器ダウンロード種別コード</param>
    ''' <param name="pZGSYO_CD_List">対象事業所CD</param>
    ''' <remarks>ダウンロード管理データ取得を取得します</remarks>
    Friend Function S001(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pKeikiDwldSbtCd As String, ByVal pZGSYO_CD_List As List(Of String)) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String
        Dim whereZgsyo As String = " "
        Dim i As Integer
        If pZGSYO_CD_List.Count > 0 Then
            whereZgsyo = " AND ("
            For i = 0 To pZGSYO_CD_List.Count - 1
                If i > 0 Then
                    whereZgsyo = whereZgsyo & " Or "
                End If
                whereZgsyo = whereZgsyo & "A.dig4_zgsyo_cd like '" & pZGSYO_CD_List(i) & "%'"
            Next
            whereZgsyo = whereZgsyo & " )"
        End If

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            sb.Append(<sql><![CDATA[
select
    A.downld_renno
   ,A.KTTNMT_NO
   ,A.keiki_user_id
   ,A.keiki_dwld_sbt_cd
   ,A.mtds_file_zti_cd
   ,A.dig4_zgsyo_cd
FROM trke_downld_mng A
WHERE 1 = 1
and A.keiki_dwld_sbt_cd={0}
and A.mtds_file_zti_cd={1}
{2}
order by
    A.downld_renno
limit {3} offset 0
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , EcOrgIO.EcOrgString.GetSqlText(pKeikiDwldSbtCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileSizi) _
                               , whereZgsyo _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText([Const].C_SELECT_MAX_REC))
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

#Region "(SQL)ダウンロード管理詳細・取替_取替票(低圧)　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票(低圧)　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票(低圧)を検索します</remarks>
    Friend Function S002(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
            'Rev044 UTF-8対応でCONVER_SJIS()不要
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            '持出ファイル状態コード='1'(持出指示)を検索する。
            'Rev040 2024/12/03 物理分割 次世代QR対応 ADD tkkk_tnsb_seizo_ym ～ kdkk_tnsb_maker_mng
            'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD tkkk_tnsb_seizo_ym ～ kdkk_tnsb_maker_mng
            sb.Append(<sql><![CDATA[
select
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,D2.keiki_user_id As D_keiki_user_id
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.prcmg_cd
,T1.kiyk_no
,T1.kyokyu_spot_tokti_no
,T1.yuko_kigen_ym
,T1.trke_sbt_cd
,T1.cs_tdhkn_add_cd
,T1.cs_siku_add_cd
,T1.cs_oazat_add_cd
,T1.cs_azatm_add_cd
,T1.kozi_corp_cd
,T1.kzkis_mdgt_cd
,T1.kozitn_cd
,T1.skosya_cd
,T1.skosya_ms
,T1.kohu_ymd
,T1.skohu_ymd
,T1.hiky_yotei_ymd
,T1.sksti_ymd
,T1.skssti_ymd
,T1.skyti_ymd
,T1.syun_ymd
,T1.trkhy_hinpt_flg
,T1.seko_kka_tenso_ymd
,T1.skrepo_kka_kbn
,T1.knsgk_ymd
,T1.keiki_sitei_no
,T1.eigyo_rnrk_flg
,T1.brot_sett_flg
--,T1.kzkis_sett_sett_ymd
--,T1.kzkis_sett_tntsy_cd
--,T1.trtk_pln_sett_ymd
--,T1.trtk_pln_tntsy_cd
,T1.krds_sett_zyokyo_cd
--,T1.krds_kizyo_ymd
--,T1.krds_kizyo_tntsy_cd
,T1.hory_flg
--,T1.kensa_sizi_ymd
--,T1.kensa_sizi_tntsy_cd
,T1.nktr_kensa_zyokyo_cd
,T1.ntkns_tblt_recv_ymd
,T1.ntkns_tblt_recv_tntsy_cd
,T1.ntkns_zissi_ymd
,T1.ntkns_zissi_tntsy_cd
,T1.ntkns_kka_upld_ymd
,T1.ntkns_kka_upld_tntsy_cd
,T1.thyhkn_syori_ymd
,T1.skohu_mae_prcmg_cd
,T1.skohu_kzkis_mdgt_cd
,T1.skohu_skssti_ymd
,T1.skohu_hiky_yotei_ymd
,T1.skohu_knti_sbt_cd
,T1.tidu_cclme_prcmg_cd
,T1.tidu_cancl_ymd
,T1.keiki_kmaws_no
,T1.ksyu_cd
,T1.gnp_no
,T1.tkkk_keiki_id
,T1.tkkk_kesbt_kbn
,T1.tkkk_krhsk_cd
,T1.tkkk_keiki_yr
,T1.tkkk_tshsk_cd
,CASE T1.tkkk_keiki_ktsk_ms
   WHEN ' ' THEN (SELECT keiki_ktsk_ms FROM mst_keiki_ktsk WHERE 1 = 1 AND ktsk_cd = T1.tkkk_keiki_ktsk_cd AND ksyu_cd = T1.ksyu_cd LIMIT 1)
   ELSE T1.tkkk_keiki_ktsk_ms
   END AS tkkk_keiki_ktsk_ms
,T1.tkkk_keiki_ktsk_cd
,T1.tkkk_taiko_kbn
,T1.tkkk_knthk_cd
,T1.tkkk_ts_kino_umu_flg
,T1.tkkk_khk_kino_umu_flg
,T1.tkkk_gaibu_output_umu_flg
,T1.tkkk_yuko_kigen_ym
,T1.tkkk_seizo_yy
,T1.tkkk_zyrt
,T1.tkkk_digsu
,T1.tkkk_tnsb_szsya_cd
,T1.tkkk_tnsb_seizo_yy
,T1.tkkk_zytr_szsu
,T1.tkkk_gytr_szsu
,T1.tkkk_szsu_sytk_zumi_flg
,T1.no1_tkkk_hnsik_seizo_no
,T1.no2_tkkk_hnsik_seizo_no
,T1.tkkk_hnsik_yuko_kigen_ym
,T1.tkkk_hnsik_gknti_no
,T1.tkkk_mado_su
,T1.ttkk_keiki_id
,T1.ttkk_kesbt_kbn
,T1.ttkk_krhsk_cd
,T1.ttkk_keiki_yr
,T1.ttkk_tshsk_cd
,T1.ttkk_keiki_ktsk_ms
,T1.ttkk_keiki_ktsk_cd
,T1.ttkk_taiko_kbn
,T1.ttkk_knthk_cd
,T1.ttkk_ts_kino_umu_flg
,T1.ttkk_khk_kino_umu_flg
,T1.ttkk_gaibu_output_umu_flg
,T1.ttkk_yuko_kigen_ym
,T1.ttkk_seizo_yy
,T1.ttkk_zyrt
,T1.ttkk_digsu
,T1.ttkk_tnsb_szsya_cd
,T1.ttkk_tnsb_seizo_yy
,T1.ttkk_zytr_szsu
,T1.ttkk_gytr_szsu
,T1.no1_ttkk_hnsik_seizo_no
,T1.no2_ttkk_hnsik_seizo_no
,T1.ttkk_hnsik_yuko_kigen_ym
,T1.ttkk_hnsik_gknti_no
,T1.ttkk_mado_su
,T1.kdkk_keiki_id
,T1.kdkk_kesbt_kbn
,T1.kdkk_krhsk_cd
,T1.kdkk_keiki_yr
,T1.kdkk_tshsk_cd
,T1.kdkk_keiki_ktsk_ms
,T1.kdkk_keiki_ktsk_cd
,T1.kdkk_taiko_kbn
,T1.kdkk_knthk_cd
,T1.kdkk_ts_kino_umu_flg
,T1.kdkk_khk_kino_umu_flg
,T1.kdkk_gaibu_output_umu_flg
,T1.kdkk_yuko_kigen_ym
,T1.kdkk_seizo_yy
,T1.kdkk_zyrt
,T1.kdkk_digsu
,T1.kdkk_tnsb_szsya_cd
,T1.kdkk_tnsb_seizo_yy
,T1.kdkk_zytr_szsu
,T1.kdkk_gytr_szsu
,T1.no1_kdkk_hnsik_seizo_no
,T1.no2_kdkk_hnsik_seizo_no
,T1.kdkk_hnsik_yuko_kigen_ym
,T1.kdkk_hnsik_gknti_no
,T1.kdkk_mado_su
,T1.tdnstr_ts_dosa_kbn
,T1.tdnstr_tdnt_su
,T1.tdnstr_tdnstr_time
,T1.hkgds_kbn
,T1.hkgds_start_md
,T1.hkgds_end_md
,T1.no1_hkgds_tdn_hms
,T1.no1_hkgds_sydan_hms
,T1.no2_hkgds_tdn_hms
,T1.no2_hkgds_sydan_hms
,T1.no3_hkgds_tdn_hms
,T1.no3_hkgds_sydan_hms
,T1.no4_hkgds_tdn_hms
,T1.no4_hkgds_sydan_hms
,T1.no5_hkgds_tdn_hms
,T1.no5_hkgds_sydan_hms
,T1.no6_hkgds_tdn_hms
,T1.no6_hkgds_sydan_hms
,T1.no7_hkgds_tdn_hms
,T1.no7_hkgds_sydan_hms
,T1.no8_hkgds_tdn_hms
,T1.no8_hkgds_sydan_hms
,T1.no9_hkgds_tdn_hms
,T1.no9_hkgds_sydan_hms
,T1.no10_hkgds_tdn_hms
,T1.no10_hkgds_sydan_hms
,T1.kbttd_start_date
,T1.kbttd_end_date
,T1.hoka_kihi_kbn_cd
,T1.hoka_gmn_flckr_kbn
,T1.hoka_evnt_krk_kbn
,T1.hoka_hoka_hyz_kbn
,T1.hoka_hoka_hyz_value
,T1.hksgk_hksgn_kbn
,T1.hksgk_hka_dnr
,T1.hksgk_attny_time
,T1.hksgk_attny_kaisu
,T1.hksgk_attny_kaisu_clr_time
,T1.hksgr_hksgn_kbn
,T1.hksgr_hka_dnr
,T1.hksgr_attny_time
,T1.hksgr_attny_kaisu
,T1.hksgr_attny_kaisu_clr_time
,T1.knti_sbt_cd
,T1.surge_yksi_siyo_umu_flg
,T1.tidn_kbn
,T1.tnsb_reuse_kbn
,T1.khkmk_keiki_kbn
,T1.khkmk_mg_kbn
,T1.khkmk_hnsik_kbn
,T1.khkmk_wrms_kbn
,T1.khkmk_mg_umu_flg
,T1.khkmk_mg_yr
,T1.khkmk_kozi_kbn
,T1.khkmk_kozi_suryo
,T1.khkmk_trtk_kozih_kngk
,T1.khkmk_tekyo_kozih_kngk
,T1.kzkmk_hnsik_kozi_kbn
,T1.kzkmk_hnsik_kozi_suryo
,T1.kzkmk_hnsik_trtk_kozih_kngk
,T1.kzkmk_hnsik_tekyo_kozih_kngk
,T1.kzkmk_mg_kozi_kbn
,T1.kzkmk_mg_kozi_suryo
,T1.khkmk_mg_trtk_kozih_kngk
,T1.khkmk_mg_tekyo_kozih_kngk
,T1.tuika_kohi_trtk_kozih_kngk
,T1.tuika_kohi_tekyo_kozih_kngk
,T1.tuika_kohi_umu_flg
,T1.zippi_umu_flg
,T1.no1_zippi_no
,T1.no1_zippi_kbn
,T1.no1_zippi_kngk
,T1.no2_zippi_no
,T1.no2_zippi_kbn
,T1.no2_zippi_kngk
,T1.no3_zippi_no
,T1.no3_zippi_kbn
,T1.no3_zippi_kngk
,T1.kozih_gokei_kngk
,T1.ryssy_tenp_file_mng_no
,T1.tenp_file_mng_no
,T1.rrzk_naiyo
,T1.brtst_htiri_saksi_flg
,T1.kktkt_syun_list_saksi_flg
,T1.srcy_kiyk_no_renno
,T1.srcy_kiyk_no_edab_no
,T1.kesbt_bnrui_cd
,SUBSTR(K1.kiyk_no,5,8) || ' -' || SUBSTR(K1.kiyk_no,13,1) AS kiyk_no_str 
 ,K1.kiyk_sbt_cd                 AS kiyk_sbt_cd 
 ,K1.kiyk_cs_mskn                AS kiyk_cs_mskn 
 ,K1.kiyk_cs_ms                  AS kiyk_cs_ms 
 ,dig14_cs_tel                   AS dig14_cs_tel_str 
 ,J1.tdhkn_ms                    AS cs_tdhkn_add_cd_ms 
 ,J1.siku_ms                     AS cs_siku_add_cd_ms 
 ,J1.oazat_ms                    AS cs_oazat_add_cd_ms 
 ,J1.azatm_ms                    AS cs_azatm_add_cd_ms 
 ,K1.cs_addhsk_naiyo             AS cs_addhsk_naiyo 
 ,K1.mkhy_senro_cd               AS mkhy_senro_cd 
 ,K1.mkhy_kansn_no               AS mkhy_kansn_no 
 ,K1.mkhy_bunk1_no               AS mkhy_bunk1_no 
 ,K1.mkhy_bunk2_no               AS mkhy_bunk2_no 
 ,K1.mkhy_bunk3_no               AS mkhy_bunk3_no 
 ,K1.cs_kiyk_dnryk               AS cs_kiyk_dnryk 
 ,K1.std_kensn_dd                AS std_kensn_dd 
 ,K1.kensn_yotei_togt_md         AS kensn_yotei_togt_md 
 ,K1.kensn_yotei_ykgt_md         AS kensn_yotei_ykgt_md 
 ,K1.kensn_yotei_yygt_md         AS kensn_yotei_yygt_md 
 ,K1.hist_flg                    AS hist_flg 
 ,K1.stbit_xzahyo                AS stbit_xzahyo 
 ,K1.stbit_yzahyo                AS stbit_yzahyo 
 ,K1.iewku_stdp_xzahyo           AS iewku_stdp_xzahyo 
 ,K1.iewku_stdp_yzahyo           AS iewku_stdp_yzahyo 
 ,M19.prm_cd_abms                AS prcmg_cd_abms 
 ,M2.prm_cd_abms                 AS brot_sett_flg_abms 
 ,M3.prm_cd_abms                 AS tkkk_kesbt_kbn_abms 
 ,M4.prm_cd_abms                 AS tkkk_krhsk_cd_abms 
 ,M5.prm_cd_abms                 AS tkkk_tshsk_cd_abms 
 ,M6.prm_cd_abms                 AS tkkk_knthk_cd_abms 
 ,M7.prm_cd_abms                 AS tkkk_ts_kino_umu_flg_abms 
 ,M11.prm_cd_abms                AS ttkk_kesbt_kbn_abms 
 ,M12.prm_cd_abms                AS ttkk_krhsk_cd_abms 
 ,M13.prm_cd_abms                AS ttkk_tshsk_cd_abms 
 ,M14.prm_cd_abms                AS ttkk_knthk_cd_abms 
 ,M15.prm_cd_abms                AS ttkk_ts_kino_umu_flg_abms 
 ,M16.prm_cd_abms                AS ttkk_khk_kino_umu_flg_abms 
 ,M17.prm_cd_abms                AS ttkk_gaibu_output_umu_flg_ms 
 ,M10.prm_cd_abms                AS ttkk_taiko_kbn_abms 
 ,M37.prm_cd_ms                  AS knti_sbt_cd_ms
 ,M1.prm_cd_ms                   AS trke_tis_abms
 ,M20.prm_cd_abms                AS nktr_kensa_zyokyo_abms
 ,M21.prm_cd_abms                AS trke_sbt_abms
 ,M22.prm_cd_abms                AS tkkk_khk_kino_umu_flg_abms
 ,M23.prm_cd_abms                AS tkkk_gaibu_output_umu_flg_abms
 ,M24.prm_cd_abms                AS tkkk_taiko_kbn_abms
 ,M25.prm_cd_abms                AS ttkk_knthk_abms
 ,M26.prm_cd_abms                AS surge_yksi_siyo_umu_flg_abms
 ,M27.prm_cd_abms                AS tidn_kbn_abms
 ,M28.prm_cd_abms                AS tnsb_reuse_kbn_abms
 ,M29.prm_cd_abms                AS khkmk_keiki_kbn_abms
 ,M30.prm_cd_abms                AS khkmk_mg_kbn_abms
 ,M31.prm_cd_abms                AS khkmk_hnsik_kbn_abms
 ,M32.prm_cd_abms                AS khkmk_wrms_kbn_abms
 ,M33.prm_cd_abms                AS khkmk_mg_umu_flg_abms
 ,M34.prm_cd_abms                AS no1_zippi_kbn_abms
 ,M35.prm_cd_abms                AS no2_zippi_kbn_abms
 ,M36.prm_cd_abms                AS no3_zippi_kbn_abms
 ,J2.mdgt_ms                     AS mdgt_ms
 ,J3.kozitn_ms                   AS kozitn_ms
 ,J4.senro_ms                    AS senro_ms
 ,J5.str2_zyousu_value           AS tykei_mdgt_ms
 ,J6.str2_zyousu_value           AS tykei_kozitn_ms
 ,' '                            AS yukoWhZytrTime 
 ,' '                            AS yukoWhGytrTime
 ,T1.nktrk_sizsk_kbn             AS nktrk_sizsk_kbn
 ,'1'                            AS phase_kbn
 ,' '                            AS trke_tis_kbn
 ,T1.tkkk_tnsb_seizo_ym          AS tkkk_tnsb_seizo_ym
 ,T1.tkkk_tnsb_sbt_cd            AS tkkk_tnsb_sbt_cd
 ,T1.tkkk_tnsb_ssnsk_dnat_cd     AS tkkk_tnsb_ssnsk_dnat_cd
 ,T1.tkkk_tnsb_yr                AS tkkk_tnsb_yr
 ,T1.tkkk_tnsb_seizo_no          AS tkkk_tnsb_seizo_no
 ,T1.tkkk_tnsb_ktsk_ms           AS tkkk_tnsb_ktsk_ms
 ,T1.tkkk_tnsb_kzskbt_kbn        AS tkkk_tnsb_kzskbt_kbn
 ,T1.tkkk_tnsb_szsya_mng_value   AS tkkk_tnsb_szsya_mng_value
 ,T1.ttkk_tnsb_seizo_ym          AS ttkk_tnsb_seizo_ym
 ,T1.ttkk_tnsb_sbt_cd            AS ttkk_tnsb_sbt_cd
 ,T1.ttkk_tnsb_ssnsk_dnat_cd     AS ttkk_tnsb_ssnsk_dnat_cd
 ,T1.ttkk_tnsb_yr                AS ttkk_tnsb_yr
 ,T1.ttkk_tnsb_seizo_no          AS ttkk_tnsb_seizo_no
 ,T1.ttkk_tnsb_ktsk_ms           AS ttkk_tnsb_ktsk_ms
 ,T1.ttkk_tnsb_kzskbt_kbn        AS ttkk_tnsb_kzskbt_kbn
 ,T1.ttkk_tnsb_szsya_mng_value   AS ttkk_tnsb_szsya_mng_value
 ,T1.kdkk_tnsb_seizo_ym          AS kdkk_tnsb_seizo_ym
 ,T1.kdkk_tnsb_sbt_cd            AS kdkk_tnsb_sbt_cd
 ,T1.kdkk_tnsb_ssnsk_dnat_cd     AS kdkk_tnsb_ssnsk_dnat_cd
 ,T1.kdkk_tnsb_yr                AS kdkk_tnsb_yr
 ,T1.kdkk_tnsb_seizo_no          AS kdkk_tnsb_seizo_no
 ,T1.kdkk_tnsb_ktsk_ms           AS kdkk_tnsb_ktsk_ms
 ,T1.kdkk_tnsb_kzskbt_kbn        AS kdkk_tnsb_kzskbt_kbn
 ,T1.kdkk_tnsb_szsya_mng_value   AS kdkk_tnsb_szsya_mng_value
 ,K1.stzk_sdnsv_menu_cd          AS stzk_sdnsv_menu_cd
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--ダウンロード管理
  INNER JOIN trke_downld_mng D2 
ON D2.downld_renno=D1.downld_renno 
--取替_取替票（低圧）
  INNER JOIN trke_trkehy_tiat T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--契約_お客さま情報
  INNER JOIN kiyk_customer_info K1 
ON T1.dig4_zgsyo_cd = K1.dig4_zgsyo_cd 
   AND T1.trkhy_hakko_nendo = K1.hk_nendo 
   AND T1.kyokyu_spot_tokti_no = K1.kyokyu_spot_tokti_no 
   AND T1.trkhy_kbn = K1.trkhy_kbn 
--マスタ_町字
  LEFT JOIN mst_tyaza J1  
ON J1.tdhkn_add_cd = T1.cs_tdhkn_add_cd 
   AND J1.siku_add_cd = T1.cs_siku_add_cd  
   AND J1.oazat_add_cd = T1.cs_oazat_add_cd  
   AND J1.azatm_add_cd = T1.cs_azatm_add_cd 
--取替対象名称
  INNER JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'TRKE_TIS_CD' 
   AND M1.prm_cd = T1.knti_sbt_cd 
--Ｂルート設定 
  INNER JOIN  com_paramater_cd_mst M2 
ON M2.prm_id = 'BROOT_SETT_FLG' 
   AND M2.prm_cd = T1.brot_sett_flg 
--撤去計器_計器種別区分
  INNER JOIN  com_paramater_cd_mst M3 
ON M3.prm_id = 'KESBT_KBN' 
   AND M3.prm_cd = T1.tkkk_kesbt_kbn 
--撤去計器_計量方式コード
  INNER JOIN  com_paramater_cd_mst M4 
ON M4.prm_id = 'KRHSK_CD' 
   AND M4.prm_cd = T1.tkkk_krhsk_cd 
--撤去計器_通信方式コード
  INNER JOIN  com_paramater_cd_mst M5 
   ON M5.prm_id = 'TSHSK_CD' 
  AND M5.prm_cd = T1.tkkk_tshsk_cd 
--撤去計器_検定方向コード
  INNER JOIN  com_paramater_cd_mst M6 
ON M6.prm_id = 'KNTHK_CD' 
   AND M6.prm_cd = T1.tkkk_knthk_cd 
--撤去計器_ＴＳ機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M7 
ON M7.prm_id = 'TS_KINOUMU_FLG' 
   AND M7.prm_cd = T1.tkkk_ts_kino_umu_flg 
--取付計器_計器_耐候区分
  LEFT JOIN  com_paramater_cd_mst M10 
ON M10.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M10.prm_cd = T1.ttkk_taiko_kbn 
--取付計器_計器種別コード
  LEFT JOIN  com_paramater_cd_mst M11 
ON M11.prm_id = 'KESBT_KBN' 
   AND M11.prm_cd = T1.ttkk_kesbt_kbn 
--取付計器_計量方式コード
  LEFT JOIN  com_paramater_cd_mst M12 
ON M12.prm_id = 'KRHSK_CD' 
   AND M12.prm_cd = T1.ttkk_krhsk_cd 
--取付計器_通信方式コード
  LEFT JOIN  com_paramater_cd_mst M13 
ON M13.prm_id = 'TSHSK_CD' 
   AND M13.prm_cd = T1.ttkk_tshsk_cd 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M14 
ON M14.prm_id = 'KNTHK_CD' 
   AND M14.prm_cd = T1.ttkk_knthk_cd 
--取付計器_ＴＳ機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M15 
ON M15.prm_id = 'TS_KINOUMU_FLG' 
   AND M15.prm_cd = T1.ttkk_ts_kino_umu_flg 
--取付計器_開閉器機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M16 
ON M16.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M16.prm_cd = T1.ttkk_khk_kino_umu_flg 
--取付計器_外部出力有無フラグ
  LEFT JOIN  com_paramater_cd_mst M17 
ON M17.prm_id = 'OUTPUTUMU_FLG' 
   AND M17.prm_cd = T1.ttkk_gaibu_output_umu_flg 
--行程管理コード
  INNER JOIN  com_paramater_cd_mst M19 
ON M19.prm_id = 'PROC_MNG_CD' 
   AND M19.prm_cd = T1.prcmg_cd 
--抜取検査状況取替対象
  LEFT JOIN  com_paramater_cd_mst M20 
ON M20.prm_id = 'NKTRKENSA_ZYOKYO_CD' 
   AND M20.prm_cd = T1.nktr_kensa_zyokyo_cd 
--取替種別
  INNER JOIN  com_paramater_cd_mst M21 
ON M21.prm_id = 'TRKE_SBT_CD' 
   AND M21.prm_cd = T1.trke_sbt_cd 
--撤去計器_開閉器機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M22 
ON M22.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M22.prm_cd = T1.tkkk_khk_kino_umu_flg 
--撤去計器_外部出力有無フラグ
  INNER JOIN  com_paramater_cd_mst M23 
ON M23.prm_id = 'OUTPUTUMU_FLG' 
   AND M23.prm_cd = T1.tkkk_gaibu_output_umu_flg 
--耐候区分
  INNER JOIN  com_paramater_cd_mst M24 
ON M24.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M24.prm_cd = T1.tkkk_taiko_kbn 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M25 
ON M25.prm_id = 'KNTHK_CD' 
   AND M25.prm_cd = T1.ttkk_knthk_cd 
--サージ抑制使用有無フラグ
  LEFT JOIN  com_paramater_cd_mst M26 
ON M26.prm_id = 'SURG_YKSISIYOUMU_FLG' 
   AND M26.prm_cd = T1.surge_yksi_siyo_umu_flg 
--停電区分
  LEFT JOIN  com_paramater_cd_mst M27 
ON M27.prm_id = 'TIDN_KBN' 
   AND M27.prm_cd = T1.tidn_kbn 
--端子部再用区分
  LEFT JOIN  com_paramater_cd_mst M28 
ON M28.prm_id = 'TNSB_REUSE_KBN' 
   AND M28.prm_cd = T1.tnsb_reuse_kbn 
--工費項目_計器区分
  INNER JOIN  com_paramater_cd_mst M29 
ON M29.prm_id = 'KOHIKOMK_KEIKI_KBN' 
   AND M29.prm_cd = T1.khkmk_keiki_kbn 
--工費項目_ＭＧ
  LEFT JOIN  com_paramater_cd_mst M30 
ON M30.prm_id = 'KOHIKOMK_MG_KBN' 
   AND M30.prm_cd = T1.khkmk_mg_kbn 
--工費項目_変成器区分
  LEFT JOIN  com_paramater_cd_mst M31 
ON M31.prm_id = 'KOHIKOMK_HNSK_KBN' 
   AND M31.prm_cd = T1.khkmk_hnsik_kbn 
--工費項目_割増区分
  LEFT JOIN  com_paramater_cd_mst M32 
ON M32.prm_id = 'KOHIKOMK_WRMS_CD' 
   AND M32.prm_cd = T1.khkmk_wrms_kbn 
--工費項目_ＭＧ有無フラグ
  INNER JOIN  com_paramater_cd_mst M33 
ON M33.prm_id = 'KOHIKOMK_MGUMU_FLG' 
   AND M33.prm_cd = T1.khkmk_mg_umu_flg 
--実費_区分１
  LEFT JOIN  com_paramater_cd_mst M34 
ON M34.prm_id = 'ZIPPI_KBN' 
   AND M34.prm_cd = T1.no1_zippi_kbn 
--実費_区分２
  LEFT JOIN  com_paramater_cd_mst M35 
ON M35.prm_id = 'ZIPPI_KBN' 
   AND M35.prm_cd = T1.no2_zippi_kbn 
--実費_区分３
  LEFT JOIN  com_paramater_cd_mst M36 
ON M36.prm_id = 'ZIPPI_KBN' 
   AND M36.prm_cd = T1.no3_zippi_kbn 
--検定種別名称
  LEFT JOIN  com_paramater_cd_mst M37
ON M37.prm_id = 'KNTI_SBT_CD'
   AND M37.prm_cd = T1.knti_sbt_cd
--工事会社窓口名
 LEFT JOIN mst_mdgt J2
    ON J2.mdgt_cd = T1.kzkis_mdgt_cd
   AND J2.latest_flg = '1'
   AND J2.haisi_flg = '0'
 --工事店名
 LEFT JOIN mst_kozitn J3
    ON J3.kozitn_cd = T1.kozitn_cd
   AND J3.latest_flg = '1'
   AND J3.haisi_flg = '0'
 --線路名称
 LEFT JOIN mst_senro J4
    ON J4.senro_cd = K1.mkhy_senro_cd
--直営　窓口名
 LEFT JOIN com_zyousu_mst J5
    ON J5.zyousu_cd = {1} and {0} BETWEEN J5.yuko_start_ymd AND J5.yuko_end_ymd
    --直営　工事店名
 LEFT JOIN com_zyousu_mst J6
    ON J6.zyousu_cd = {2} and {0} BETWEEN J6.yuko_start_ymd AND J6.yuko_end_ymd
WHERE 1 = 1 
AND D1.downld_renno={3}

UNION

SELECT
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,D2.keiki_user_id As D_keiki_user_id
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.prcmg_cd
,T1.kiyk_no
,T1.kyokyu_spot_tokti_no
,T1.yuko_kigen_ym
,T1.trke_sbt_cd
,T1.cs_tdhkn_add_cd
,T1.cs_siku_add_cd
,T1.cs_oazat_add_cd
,T1.cs_azatm_add_cd
,T1.kozi_corp_cd
,T1.kzkis_mdgt_cd
,T1.kozitn_cd
,T1.skosya_cd
,T1.skosya_ms
,T1.kohu_ymd
,T1.skohu_ymd
,T1.hiky_yotei_ymd
,T1.sksti_ymd
,T1.skssti_ymd
,T1.skyti_ymd
,T1.syun_ymd
,T1.trkhy_hinpt_flg
,T1.seko_kka_tenso_ymd
,T1.skrepo_kka_kbn
,T1.knsgk_ymd
,T1.keiki_sitei_no
,T1.eigyo_rnrk_flg
,T1.brot_sett_flg
--,T1.kzkis_sett_sett_ymd
--,T1.kzkis_sett_tntsy_cd
--,T1.trtk_pln_sett_ymd
--,T1.trtk_pln_tntsy_cd
,T1.krds_sett_zyokyo_cd
--,T1.krds_kizyo_ymd
--,T1.krds_kizyo_tntsy_cd
,T1.hory_flg
--,T1.kensa_sizi_ymd
--,T1.kensa_sizi_tntsy_cd
,T1.nktr_kensa_zyokyo_cd
,T1.nktrk_tblt_recv_ymd
,T1.nktrk_tblt_recv_tntsy_cd
,T1.nktrk_zissi_ymd
,T1.nktrk_zissi_tntsy_cd
,T1.nktrk_kka_upld_ymd
,T1.nktrk_kka_upld_tntsy_cd
,T1.thyhkn_syori_ymd
,T1.skohu_mae_prcmg_cd
,T1.skohu_kzkis_mdgt_cd
,T1.skohu_skssti_ymd
,T1.skohu_hiky_yotei_ymd
,T1.skohu_knti_sbt_cd
,T1.tidu_cclme_prcmg_cd
,T1.tidu_cancl_ymd
,T1.keiki_kmaws_no
,' ' AS ksyu_cd
,T1.gnp_no
,T1.tkkk_keiki_id
,T1.tkkk_kesbt_kbn
,T1.tkkk_krhsk_cd
,T1.tkkk_keiki_yr
,T1.tkkk_tshsk_cd
,CASE T1.tkkk_keiki_ktsk_ms
   WHEN ' ' THEN (SELECT keiki_ktsk_ms FROM mst_keiki_ktsk WHERE 1 = 1 AND ktsk_cd = T1.tkkk_keiki_ktsk_cd AND ksyu_cd = T1.tkkk_ksyu_cd LIMIT 1)
   ELSE T1.tkkk_keiki_ktsk_ms
   END AS tkkk_keiki_ktsk_ms
,T1.tkkk_keiki_ktsk_cd
,T1.tkkk_taiko_kbn
,T1.tkkk_knthk_cd
,T1.tkkk_ts_kino_umu_flg
,T1.tkkk_khk_kino_umu_flg
,T1.tkkk_gaibu_output_umu_flg
,T1.tkkk_yuko_kigen_ym
,T1.tkkk_seizo_yy
,T1.tkkk_zyrt
,T1.tkkk_digsu
,T1.tkkk_tnsb_szsya_cd
,T1.tkkk_tnsb_seizo_yy
,T1.tkkk_zytr_szsu
,T1.tkkk_gytr_szsu
,T1.tkkk_szsu_sytk_zumi_flg
,T1.no1_tkkk_hnsik_seizo_no
,T1.no2_tkkk_hnsik_seizo_no
,T1.tkkk_hnsik_yuko_kigen_ym
,T1.tkkk_hnsik_gknti_no
,T1.tkkk_mado_su
,T1.ttkk_keiki_id
,T1.ttkk_kesbt_kbn
,T1.ttkk_krhsk_cd
,T1.ttkk_keiki_yr
,T1.ttkk_tshsk_cd
,T1.ttkk_keiki_ktsk_ms
,T1.ttkk_keiki_ktsk_cd
,T1.ttkk_taiko_kbn
,T1.ttkk_knthk_cd
,T1.ttkk_ts_kino_umu_flg
,T1.ttkk_khk_kino_umu_flg
,T1.ttkk_gaibu_output_umu_flg
,T1.ttkk_yuko_kigen_ym
,T1.ttkk_seizo_yy
,T1.ttkk_zyrt
,T1.ttkk_digsu
,T1.ttkk_tnsb_szsya_cd
,T1.ttkk_tnsb_seizo_yy
,T1.ttkk_zytr_szsu
,T1.ttkk_gytr_szsu
,T1.no1_ttkk_hnsik_seizo_no
,T1.no2_ttkk_hnsik_seizo_no
,T1.ttkk_hnsik_yuko_kigen_ym
,T1.ttkk_hnsik_gknti_no
,T1.ttkk_mado_su
,T1.kdkk_keiki_id
,T1.kdkk_kesbt_kbn
,T1.kdkk_krhsk_cd
,T1.kdkk_keiki_yr
,T1.kdkk_tshsk_cd
,T1.kdkk_keiki_ktsk_ms
,T1.kdkk_keiki_ktsk_cd
,T1.kdkk_taiko_kbn
,T1.kdkk_knthk_cd
,T1.kdkk_ts_kino_umu_flg
,T1.kdkk_khk_kino_umu_flg
,T1.kdkk_gaibu_output_umu_flg
,T1.kdkk_yuko_kigen_ym
,T1.kdkk_seizo_yy
,T1.kdkk_zyrt
,T1.kdkk_digsu
,T1.kdkk_tnsb_szsya_cd
,T1.kdkk_tnsb_seizo_yy
,T1.kdkk_zytr_szsu
,T1.kdkk_gytr_szsu
,T1.no1_kdkk_hnsik_seizo_no
,T1.no2_kdkk_hnsik_seizo_no
,T1.kdkk_hnsik_yuko_kigen_ym
,T1.kdkk_hnsik_gknti_no
,T1.kdkk_mado_su
,T1.tdnstr_ts_dosa_kbn
,T1.tdnstr_tdnt_su
,T1.tdnstr_tdnstr_time
,T1.hkgds_kbn
,T1.hkgds_start_md
,T1.hkgds_end_md
,T1.no1_hkgds_tdn_hms
,T1.no1_hkgds_sydan_hms
,T1.no2_hkgds_tdn_hms
,T1.no2_hkgds_sydan_hms
,T1.no3_hkgds_tdn_hms
,T1.no3_hkgds_sydan_hms
,T1.no4_hkgds_tdn_hms
,T1.no4_hkgds_sydan_hms
,T1.no5_hkgds_tdn_hms
,T1.no5_hkgds_sydan_hms
,T1.no6_hkgds_tdn_hms
,T1.no6_hkgds_sydan_hms
,T1.no7_hkgds_tdn_hms
,T1.no7_hkgds_sydan_hms
,T1.no8_hkgds_tdn_hms
,T1.no8_hkgds_sydan_hms
,T1.no9_hkgds_tdn_hms
,T1.no9_hkgds_sydan_hms
,T1.no10_hkgds_tdn_hms
,T1.no10_hkgds_sydan_hms
,T1.kbttd_start_date
,T1.kbttd_end_date
,T1.hoka_kihi_kbn_cd
,T1.hoka_gmn_flckr_kbn
,T1.hoka_evnt_krk_kbn
,T1.hoka_hoka_hyz_kbn
,T1.hoka_hoka_hyz_value
,T1.hksgk_hksgn_kbn
,T1.hksgk_hka_dnr
,T1.hksgk_attny_time
,T1.hksgk_attny_kaisu
,T1.hksgk_attny_kaisu_clr_time
,T1.hksgr_hksgn_kbn
,T1.hksgr_hka_dnr
,T1.hksgr_attny_time
,T1.hksgr_attny_kaisu
,T1.hksgr_attny_kaisu_clr_time
,T1.knti_sbt_cd
,T1.surge_yksi_siyo_umu_flg
,T1.tidn_kbn
,T1.tnsb_reuse_kbn
,T1.khkmk_keiki_kbn
,T1.khkmk_mg_kbn
,T1.khkmk_hnsik_kbn
,T1.khkmk_wrms_kbn
,T1.khkmk_mg_umu_flg
,T1.khkmk_mg_yr
,T1.khkmk_kozi_kbn
,T1.khkmk_kozi_suryo
,T1.khkmk_trtk_kozih_kngk
,T1.khkmk_tekyo_kozih_kngk
,T1.kzkmk_hnsik_kozi_kbn
,T1.kzkmk_hnsik_kozi_suryo
,T1.kzkmk_hnsik_trtk_kozih_kngk
,T1.kzkmk_hnsik_tekyo_kozih_kngk
,T1.kzkmk_mg_kozi_kbn
,T1.kzkmk_mg_kozi_suryo
,T1.khkmk_mg_trtk_kozih_kngk
,T1.khkmk_mg_tekyo_kozih_kngk
,T1.tuika_kohi_trtk_kozih_kngk
,T1.tuika_kohi_tekyo_kozih_kngk
,T1.tuika_kohi_umu_flg
,T1.zippi_umu_flg
,T1.no1_zippi_no
,T1.no1_zippi_kbn
,T1.no1_zippi_kngk
,T1.no2_zippi_no
,T1.no2_zippi_kbn
,T1.no2_zippi_kngk
,T1.no3_zippi_no
,T1.no3_zippi_kbn
,T1.no3_zippi_kngk
,T1.kozih_gokei_kngk
,T1.ryssy_tenp_file_mng_no
,T1.tenp_file_mng_no
,T1.rrzk_naiyo
,T1.brtst_htiri_saksi_flg
,T1.kktkt_syun_list_saksi_flg
,T1.srcy_kiyk_no_renno
,T1.srcy_kiyk_no_edab_no
,T1.kesbt_bnrui_cd
,SUBSTR(K1.kiyk_no,5,8) || ' -' || SUBSTR(K1.kiyk_no,13,1) AS kiyk_no_str 
 ,K1.kiyk_sbt_cd                 AS kiyk_sbt_cd 
 ,K1.kiyk_cs_mskn                AS kiyk_cs_mskn 
 ,K1.kiyk_cs_ms                  AS kiyk_cs_ms 
 ,dig14_cs_tel                   AS dig14_cs_tel_str 
 ,J1.tdhkn_ms                    AS cs_tdhkn_add_cd_ms 
 ,J1.siku_ms                     AS cs_siku_add_cd_ms 
 ,J1.oazat_ms                    AS cs_oazat_add_cd_ms 
 ,J1.azatm_ms                    AS cs_azatm_add_cd_ms 
 ,K1.cs_addhsk_naiyo             AS cs_addhsk_naiyo 
 ,K1.mkhy_senro_cd               AS mkhy_senro_cd 
 ,K1.mkhy_kansn_no               AS mkhy_kansn_no 
 ,K1.mkhy_bunk1_no               AS mkhy_bunk1_no 
 ,K1.mkhy_bunk2_no               AS mkhy_bunk2_no 
 ,K1.mkhy_bunk3_no               AS mkhy_bunk3_no 
 ,K1.cs_kiyk_dnryk               AS cs_kiyk_dnryk 
 ,K1.std_kensn_dd                AS std_kensn_dd 
 ,K1.kensn_yotei_togt_md         AS kensn_yotei_togt_md 
 ,K1.kensn_yotei_ykgt_md         AS kensn_yotei_ykgt_md 
 ,K1.kensn_yotei_yygt_md         AS kensn_yotei_yygt_md 
 ,K1.hist_flg                    AS hist_flg 
 ,K1.stbit_xzahyo                AS stbit_xzahyo 
 ,K1.stbit_yzahyo                AS stbit_yzahyo 
 ,K1.iewku_stdp_xzahyo           AS iewku_stdp_xzahyo 
 ,K1.iewku_stdp_yzahyo           AS iewku_stdp_yzahyo 
 ,M19.prm_cd_abms                AS prcmg_cd_abms 
 ,M2.prm_cd_abms                 AS brot_sett_flg_abms 
 ,M3.prm_cd_abms                 AS tkkk_kesbt_kbn_abms 
 ,M4.prm_cd_abms                 AS tkkk_krhsk_cd_abms 
 ,M5.prm_cd_abms                 AS tkkk_tshsk_cd_abms 
 ,M6.prm_cd_abms                 AS tkkk_knthk_cd_abms 
 ,M7.prm_cd_abms                 AS tkkk_ts_kino_umu_flg_abms 
 ,M11.prm_cd_abms                AS ttkk_kesbt_kbn_abms 
 ,M12.prm_cd_abms                AS ttkk_krhsk_cd_abms 
 ,M13.prm_cd_abms                AS ttkk_tshsk_cd_abms
 ,M14.prm_cd_abms                AS ttkk_knthk_cd_abms 
 ,M15.prm_cd_abms                AS ttkk_ts_kino_umu_flg_abms 
 ,M16.prm_cd_abms                AS ttkk_khk_kino_umu_flg_abms 
 ,M17.prm_cd_abms                AS ttkk_gaibu_output_umu_flg_ms 
 ,M10.prm_cd_abms                AS ttkk_taiko_kbn_abms 
 ,M37.prm_cd_ms                  AS knti_sbt_cd_ms
 ,M1.prm_cd_ms                   AS trke_tis_abms
 ,M20.prm_cd_abms                AS nktr_kensa_zyokyo_abms
 ,M21.prm_cd_abms                AS trke_sbt_abms
 ,M22.prm_cd_abms                AS tkkk_khk_kino_umu_flg_abms
 ,M23.prm_cd_abms                AS tkkk_gaibu_output_umu_flg_abms
 ,M24.prm_cd_abms                AS tkkk_taiko_kbn_abms
 ,M25.prm_cd_abms                AS ttkk_knthk_abms
 ,M26.prm_cd_abms                AS surge_yksi_siyo_umu_flg_abms
 ,M27.prm_cd_abms                AS tidn_kbn_abms
 ,M28.prm_cd_abms                AS tnsb_reuse_kbn_abms
 ,M29.prm_cd_abms                AS khkmk_keiki_kbn_abms
 ,M30.prm_cd_abms                AS khkmk_mg_kbn_abms
 ,M31.prm_cd_abms                AS khkmk_hnsik_kbn_abms
 ,M32.prm_cd_abms                AS khkmk_wrms_kbn_abms
 ,M33.prm_cd_abms                AS khkmk_mg_umu_flg_abms
 ,M34.prm_cd_abms                AS no1_zippi_kbn_abms
 ,M35.prm_cd_abms                AS no2_zippi_kbn_abms
 ,M36.prm_cd_abms                AS no3_zippi_kbn_abms
 ,J2.mdgt_ms                     AS mdgt_ms
 ,J3.kozitn_ms                   AS kozitn_ms
 ,J4.senro_ms                    AS senro_ms
 ,J5.str2_zyousu_value           AS tykei_mdgt_ms
 ,J6.str2_zyousu_value           AS tykei_kozitn_ms
 ,' '                            AS yukoWhZytrTime 
 ,' '                            AS yukoWhGytrTime
 ,T1.nktrk_sizsk_kbn             AS nktrk_sizsk_kbn       --抜取検査指示先区分 　高圧他なのでない　T1.nktrk_sizsk_kbn
 ,'2'                            AS phase_kbn
 ,T1.trke_tis_kbn                AS trke_tis_kbn
 ,T1.tkkk_tnsb_seizo_ym          AS tkkk_tnsb_seizo_ym
 ,T1.tkkk_tnsb_sbt_cd            AS tkkk_tnsb_sbt_cd
 ,T1.tkkk_tnsb_ssnsk_dnat_cd     AS tkkk_tnsb_ssnsk_dnat_cd
 ,T1.tkkk_tnsb_yr                AS tkkk_tnsb_yr
 ,T1.tkkk_tnsb_seizo_no          AS tkkk_tnsb_seizo_no
 ,T1.tkkk_tnsb_ktsk_ms           AS tkkk_tnsb_ktsk_ms
 ,T1.tkkk_tnsb_kzskbt_kbn        AS tkkk_tnsb_kzskbt_kbn
 ,T1.tkkk_tnsb_szsya_mng_value   AS tkkk_tnsb_szsya_mng_value
 ,T1.ttkk_tnsb_seizo_ym          AS ttkk_tnsb_seizo_ym
 ,T1.ttkk_tnsb_sbt_cd            AS ttkk_tnsb_sbt_cd
 ,T1.ttkk_tnsb_ssnsk_dnat_cd     AS ttkk_tnsb_ssnsk_dnat_cd
 ,T1.ttkk_tnsb_yr                AS ttkk_tnsb_yr
 ,T1.ttkk_tnsb_seizo_no          AS ttkk_tnsb_seizo_no
 ,T1.ttkk_tnsb_ktsk_ms           AS ttkk_tnsb_ktsk_ms
 ,T1.ttkk_tnsb_kzskbt_kbn        AS ttkk_tnsb_kzskbt_kbn
 ,T1.ttkk_tnsb_szsya_mng_value   AS ttkk_tnsb_szsya_mng_value
 ,T1.kdkk_tnsb_seizo_ym          AS kdkk_tnsb_seizo_ym
 ,T1.kdkk_tnsb_sbt_cd            AS kdkk_tnsb_sbt_cd
 ,T1.kdkk_tnsb_ssnsk_dnat_cd     AS kdkk_tnsb_ssnsk_dnat_cd
 ,T1.kdkk_tnsb_yr                AS kdkk_tnsb_yr
 ,T1.kdkk_tnsb_seizo_no          AS kdkk_tnsb_seizo_no
 ,T1.kdkk_tnsb_ktsk_ms           AS kdkk_tnsb_ktsk_ms
 ,T1.kdkk_tnsb_kzskbt_kbn        AS kdkk_tnsb_kzskbt_kbn
 ,T1.kdkk_tnsb_szsya_mng_value   AS kdkk_tnsb_szsya_mng_value
 ,K1.stzk_sdnsv_menu_cd          AS stzk_sdnsv_menu_cd
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--ダウンロード管理
  INNER JOIN trke_downld_mng D2 
ON D2.downld_renno=D1.downld_renno 
--取替_取替票（高圧他）
  INNER JOIN trke_trkehy_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--契約_お客さま情報（高圧他）
  INNER JOIN kiyk_customer_info_kahk K1 
ON T1.dig4_zgsyo_cd = K1.dig4_zgsyo_cd 
   AND T1.trkhy_hakko_nendo = K1.hakko_nendo 
   AND T1.kyokyu_spot_tokti_no = K1.kyokyu_spot_tokti_no 
   AND T1.trkhy_kbn = K1.trkhy_kbn 
--マスタ_町字
  LEFT JOIN mst_tyaza J1  
ON J1.tdhkn_add_cd = T1.cs_tdhkn_add_cd 
   AND J1.siku_add_cd = T1.cs_siku_add_cd  
   AND J1.oazat_add_cd = T1.cs_oazat_add_cd  
   AND J1.azatm_add_cd = T1.cs_azatm_add_cd 
--計器_地点_特記事項
  LEFT JOIN keiki_spot_tkzk_kahk T2  
ON T2.kyokyu_spot_tokti_no = T1.kyokyu_spot_tokti_no 
--計器_計器情報（高圧他）_管理
  INNER JOIN  keiki_keiki_info_kahk_mng T3
ON T3.kkinf_id = T1.tkkk_kkinf_id 
   AND T3.kiyk_kbn_cd = '1'
   AND T3.keiki_kbn_cd = '1'
   AND T3.sm_zyri_kbn_cd = '1'
--取替対象名称
  INNER JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'TRKE_TIS_CD' 
   AND M1.prm_cd = T1.knti_sbt_cd 
--Ｂルート設定 
  INNER JOIN  com_paramater_cd_mst M2 
ON M2.prm_id = 'BROOT_SETT_FLG' 
   AND M2.prm_cd = T1.brot_sett_flg 
--撤去計器_計器種別区分
  INNER JOIN  com_paramater_cd_mst M3 
ON M3.prm_id = 'KESBT_KBN' 
   AND M3.prm_cd = T1.tkkk_kesbt_kbn 
--撤去計器_計量方式コード
  INNER JOIN  com_paramater_cd_mst M4 
ON M4.prm_id = 'KRHSK_CD' 
   AND M4.prm_cd = T1.tkkk_krhsk_cd 
--撤去計器_通信方式コード
  INNER JOIN  com_paramater_cd_mst M5 
   ON M5.prm_id = 'TSHSK_CD' 
  AND M5.prm_cd = T1.tkkk_tshsk_cd 
--撤去計器_検定方向コード
  INNER JOIN  com_paramater_cd_mst M6 
ON M6.prm_id = 'KNTHK_CD' 
   AND M6.prm_cd = T1.tkkk_knthk_cd 
--撤去計器_ＴＳ機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M7 
ON M7.prm_id = 'TS_KINOUMU_FLG' 
   AND M7.prm_cd = T1.tkkk_ts_kino_umu_flg 
--撤去計器_開閉器機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M8 
ON M8.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M8.prm_cd = T1.tkkk_khk_kino_umu_flg 
--撤去計器_外部出力有無フラグ
  INNER JOIN  com_paramater_cd_mst M9 
ON M9.prm_id = 'OUTPUTUMU_FLG' 
   AND M9.prm_cd = T1.tkkk_gaibu_output_umu_flg 
--取付計器_計器_耐候区分
  LEFT JOIN  com_paramater_cd_mst M10 
ON M10.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M10.prm_cd = T1.ttkk_taiko_kbn 
--取付計器_計器種別コード
  LEFT JOIN  com_paramater_cd_mst M11 
ON M11.prm_id = 'KESBT_KBN' 
   AND M11.prm_cd = T1.ttkk_kesbt_kbn 
--取付計器_計量方式コード
  LEFT JOIN  com_paramater_cd_mst M12 
ON M12.prm_id = 'KRHSK_CD' 
   AND M12.prm_cd = T1.ttkk_krhsk_cd 
--取付計器_通信方式コード
  LEFT JOIN  com_paramater_cd_mst M13 
ON M13.prm_id = 'TSHSK_CD' 
   AND M13.prm_cd = T1.ttkk_tshsk_cd 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M14 
ON M14.prm_id = 'KNTHK_CD' 
   AND M14.prm_cd = T1.ttkk_knthk_cd 
--取付計器_ＴＳ機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M15 
ON M15.prm_id = 'TS_KINOUMU_FLG' 
   AND M15.prm_cd = T1.ttkk_ts_kino_umu_flg 
--取付計器_開閉器機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M16 
ON M16.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M16.prm_cd = T1.ttkk_khk_kino_umu_flg 
--取付計器_外部出力有無フラグ
  LEFT JOIN  com_paramater_cd_mst M17 
ON M17.prm_id = 'OUTPUTUMU_FLG' 
   AND M17.prm_cd = T1.ttkk_gaibu_output_umu_flg 
--取付計器_計器_耐候区分
  LEFT JOIN  com_paramater_cd_mst M18 
ON M18.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M18.prm_cd = T1.ttkk_taiko_kbn 
--行程管理コード
  INNER JOIN  com_paramater_cd_mst M19 
ON M19.prm_id = 'PROC_MNG_CD' 
   AND M19.prm_cd = T1.prcmg_cd 
--抜取検査状況取替対象
  LEFT JOIN  com_paramater_cd_mst M20 
ON M20.prm_id = 'NKTRKENSA_ZYOKYO_CD' 
   AND M20.prm_cd = T1.nktr_kensa_zyokyo_cd 
--取替種別
  INNER JOIN  com_paramater_cd_mst M21 
ON M21.prm_id = 'TRKE_SBT_CD' 
   AND M21.prm_cd = T1.trke_sbt_cd 
--開閉器機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M22 
ON M22.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M22.prm_cd = T1.tkkk_khk_kino_umu_flg 
--外部出力有無フラグ
  INNER JOIN  com_paramater_cd_mst M23 
ON M23.prm_id = 'OUTPUTUMU_FLG' 
   AND M23.prm_cd = T1.tkkk_gaibu_output_umu_flg 
--耐候区分
  INNER JOIN  com_paramater_cd_mst M24 
ON M24.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M24.prm_cd = T1.tkkk_taiko_kbn 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M25 
ON M25.prm_id = 'KNTHK_CD' 
   AND M25.prm_cd = T1.ttkk_knthk_cd 
--サージ抑制使用有無フラグ
  LEFT JOIN  com_paramater_cd_mst M26 
ON M26.prm_id = 'SURG_YKSISIYOUMU_FLG' 
   AND M26.prm_cd = T1.surge_yksi_siyo_umu_flg 
--停電区分
  LEFT JOIN  com_paramater_cd_mst M27 
ON M27.prm_id = 'TIDN_KBN' 
   AND M27.prm_cd = T1.tidn_kbn 
--端子部再用区分
  LEFT JOIN  com_paramater_cd_mst M28 
ON M28.prm_id = 'TNSB_REUSE_KBN' 
   AND M28.prm_cd = T1.tnsb_reuse_kbn 
--工費項目_計器区分
  INNER JOIN  com_paramater_cd_mst M29 
ON M29.prm_id = 'KOHIKOMK_KEIKI_KBN' 
   AND M29.prm_cd = T1.khkmk_keiki_kbn 
--工費項目_ＭＧ
  LEFT JOIN  com_paramater_cd_mst M30 
ON M30.prm_id = 'KOHIKOMK_MG_KBN' 
   AND M30.prm_cd = T1.khkmk_mg_kbn 
--工費項目_変成器区分
  LEFT JOIN  com_paramater_cd_mst M31 
ON M31.prm_id = 'KOHIKOMK_HNSK_KBN' 
   AND M31.prm_cd = T1.khkmk_hnsik_kbn 
--工費項目_割増区分
  LEFT JOIN  com_paramater_cd_mst M32 
ON M32.prm_id = 'KOHIKOMK_WRMS_CD' 
   AND M32.prm_cd = T1.khkmk_wrms_kbn 
--工費項目_ＭＧ有無フラグ
  INNER JOIN  com_paramater_cd_mst M33 
ON M33.prm_id = 'KOHIKOMK_MGUMU_FLG' 
   AND M33.prm_cd = T1.khkmk_mg_umu_flg 
--実費_区分１
  LEFT JOIN  com_paramater_cd_mst M34 
ON M34.prm_id = 'ZIPPI_KBN' 
   AND M34.prm_cd = T1.no1_zippi_kbn 
--実費_区分２
  LEFT JOIN  com_paramater_cd_mst M35 
ON M35.prm_id = 'ZIPPI_KBN' 
   AND M35.prm_cd = T1.no2_zippi_kbn 
--実費_区分３
  LEFT JOIN  com_paramater_cd_mst M36 
ON M36.prm_id = 'ZIPPI_KBN' 
   AND M36.prm_cd = T1.no3_zippi_kbn 
--取替対象名称
  LEFT JOIN  com_paramater_cd_mst M37
ON M37.prm_id = 'KNTI_SBT_CD'
   AND M37.prm_cd = T1.knti_sbt_cd
--工事会社窓口名
 LEFT JOIN mst_mdgt J2
    ON J2.mdgt_cd = T1.kzkis_mdgt_cd
   AND J2.latest_flg = '1'
   AND J2.haisi_flg = '0'
 --工事店名
 LEFT JOIN mst_kozitn J3
    ON J3.kozitn_cd = T1.kozitn_cd
   AND J3.latest_flg = '1'
   AND J3.haisi_flg = '0'
 --線路名称
 LEFT JOIN mst_senro J4
    ON J4.senro_cd = K1.mkhy_senro_cd
--直営　窓口名
 LEFT JOIN com_zyousu_mst J5
    ON J5.zyousu_cd = {1} and {0} BETWEEN J5.yuko_start_ymd AND J5.yuko_end_ymd
    --直営　工事店名
 LEFT JOIN com_zyousu_mst J6
    ON J6.zyousu_cd = {2} and {0} BETWEEN J6.yuko_start_ymd AND J6.yuko_end_ymd
WHERE 1 = 1 
AND D1.downld_renno={3}

        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.ZyousuCd.CHOKUEI_MDGT_INFO) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理テーブル　持出状態更新"
    ''' <summary>
    ''' ダウンロード管理テーブル　持出状態更新
    ''' </summary>
    ''' <remarks>ダウンロード管理テーブルの持出状態を更新します</remarks>
    Friend Function U001(ByVal pWhereIN As String, ByVal pMtdsZtiCd As String, ByRef pRecCount As Double) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '*************************************************************************

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
            UPDATE trke_downld_mng SET         -- ダウンロード管理
              {0}                              -- 共通ヘッダフィールド
             ,mtds_file_zti_cd = {1}           -- 持出ファイル状態コード
            WHERE 1 = 1
             AND {2}                           -- ダウンロード番号
                     ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString, HdSysBasePg.Cmn.Func.GetSQLUpdateSet(Me.myDB, Me.myUser, Me.Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pMtdsZtiCd) _
                               , pWhereIN
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr, pRecCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)ダウンロード管理テーブル　持出状態更新(個別)"
    ''' <summary>
    ''' ダウンロード管理テーブル　持出状態更新(個別)
    ''' </summary>
    ''' <remarks>ダウンロード管理テーブルの持出状態を更新します</remarks>
    Friend Function U002(ByVal pDownldRenno As String, ByVal pMtdsZtiCd As String, ByRef pRecCount As Double) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '*************************************************************************

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
            UPDATE trke_downld_mng SET         -- ダウンロード管理
              {0}                              -- 共通ヘッダフィールド
             ,mtds_file_zti_cd = {1}           -- 持出ファイル状態コード
            WHERE 1 = 1
             AND downld_renno={2}              -- ダウンロード番号
                     ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString, HdSysBasePg.Cmn.Func.GetSQLUpdateSet(Me.myDB, Me.myUser, Me.Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pMtdsZtiCd) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno)
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr, pRecCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region


#Region "(SQL)ダウンロード管理詳細・取替_取替票行程　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票行程　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票行程を検索します</remarks>
    Friend Function S003(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.kkst_tanto_kaih_ymd
,T1.kkst_tanto_kaih_syors_cd
,T1.kkst_tanto_kaih_syors_ms
,T1.kkst_tanto_ka_cd
,T1.kkst_tanto_ka_ms
,T1.kkst_ht_snn_ymd
,T1.kkst_ht_snsya_cd
,T1.kkst_ht_snsya_ms
,T1.kkst_ht_ka_cd
,T1.kkst_ht_ka_ms
,T1.kkst_hnn_flg
,T1.tkpln_tanto_khiri_ymd
,T1.tkpln_tanto_khiri_syors_cd
,T1.tkpln_tanto_khiri_syors_ms
,T1.tkpln_tanto_ka_cd
,T1.tkpln_tanto_ka_ms
,T1.tkpln_tanto_kzkis_mdgt_cd
,T1.tkpln_tanto_kzkis_mdgt_ms
,T1.tkpln_tanto_kozitn_cd
,T1.tkpln_tanto_kozitn_ms
,T1.tkpln_tanto_khiri_hnn_flg
,T1.khsz_tanto_kaih_ymd
,T1.khsz_tanto_kaih_syors_cd
,T1.khsz_tanto_kaih_syors_ms
,T1.khsz_tanto_ka_cd
,T1.khsz_tanto_ka_ms
,T1.khsz_hnn_flg
,T1.seko_kti_ht_snn_ymd
,T1.seko_kti_ht_snsya_cd
,T1.seko_kti_ht_snsya_ms
,T1.seko_kti_ht_ka_cd
,T1.seko_kti_ht_ka_ms
,T1.seko_kti_kt_snn_ymd
,T1.seko_kti_kt_snsya_cd
,T1.seko_kti_kt_snsya_ms
,T1.seko_kti_kt_ka_cd
,T1.seko_kti_kt_ka_ms
,T1.seko_kti_hnn_flg
,T1.skohu_tanto_kaih_syori_ymd
,T1.skohu_tanto_kaih_syors_cd
,T1.skohu_tanto_kaih_syors_ms
,T1.skohu_tanto_ka_cd
,T1.skohu_tanto_ka_ms
,T1.skohu_ht_snn_ymd
,T1.skohu_ht_snsya_cd
,T1.skohu_ht_snsya_ms
,T1.skohu_ht_ka_cd
,T1.skohu_ht_ka_ms
,T1.skohu_kt_snn_ymd
,T1.skohu_kt_snsya_cd
,T1.skohu_kt_snsya_ms
,T1.skohu_kt_ka_cd
,T1.skohu_kt_ka_ms
,T1.skohu_hnn_flg
,T1.zkskh_tanto_kaih_syori_ymd
,T1.zkskh_tanto_kaih_syors_cd
,T1.zkskh_tanto_kaih_syors_ms
,T1.zkskh_tanto_ka_cd
,T1.zkskh_tanto_ka_ms
,T1.zkskh_ht_snn_ymd
,T1.zkskh_ht_snsya_cd
,T1.zkskh_ht_snsya_ms
,T1.zkskh_ht_ka_cd
,T1.zkskh_ht_ka_ms
,T1.zkskh_kt_snn_ymd
,T1.zkskh_kt_snsya_cd
,T1.zkskh_kt_snsya_ms
,T1.zkskh_kt_ka_cd
,T1.zkskh_kt_ka_ms
,T1.zkskh_hnn_flg
,T1.skyti_ktst_syori_ymd
,T1.skyti_ktst_syors_cd
,T1.skyti_ktst_syors_ms
,T1.skyti_ktst_ka_cd
,T1.skyti_ktst_ka_ms
,T1.skyti_ktst_kzkis_mdgt_cd
,T1.skyti_ktst_kzkis_mdgt_ms
,T1.skyti_ktst_kozitn_cd
,T1.skyti_ktst_kozitn_ms
,T1.skyti_skyti_inpt_syori_ymd
,T1.skyti_skyti_inpt_syors_cd
,T1.skyti_skyti_inpt_syors_ms
,T1.skyti_skyti_inpt_ka_cd
,T1.skyti_skyti_inpt_ka_ms
,T1.skyti_skyti_inpt_kkmg_cd
,T1.skyti_skyti_inpt_kkmg_ms
,T1.skyti_skyti_inpt_kozitn_cd
,T1.skyti_skyti_inpt_kozitn_ms
,T1.seko_tanto_syori_ymd
,T1.seko_tanto_syors_cd
,T1.seko_tanto_syors_ms
,T1.seko_tanto_ka_cd
,T1.seko_tanto_ka_ms
,T1.seko_tanto_kzkis_mdgt_cd
,T1.seko_tanto_kzkis_mdgt_ms
,T1.seko_tanto_kozitn_cd
,T1.seko_tanto_kozitn_ms
,T1.seko_elder_tnknk_kknn_flg
,T1.seko_elder_tnkn_kknn_ymd
,T1.seko_elder_tnkn_knsya_cd
,T1.seko_elder_tnkn_knsya_ms
,T1.seko_elder_ka_cd
,T1.seko_elder_ka_ms
,T1.seko_elder_kzkis_mdgt_cd
,T1.seko_elder_kzkis_mdgt_ms
,T1.seko_elder_kozitn_cd
,T1.seko_elder_kozitn_ms
,T1.skrepo_syori_ymd
,T1.skrepo_syors_cd
,T1.skrepo_syors_ms
,T1.skrepo_ka_cd
,T1.skrepo_ka_ms
,T1.skrepo_kzkis_mdgt_cd
,T1.skrepo_kzkis_mdgt_ms
,T1.skrepo_kozitn_cd
,T1.skrepo_kozitn_ms
,T1.no1_skrepo_zippi_kknn_flg
,T1.no2_skrepo_zippi_kknn_flg
,T1.no3_skrepo_zippi_kknn_flg
,T1.skkns_kiktt_kaih_ymd
,T1.skkns_kiktt_kaih_syors_cd
,T1.skkns_kiktt_kaih_syors_ms
,T1.skkns_kiktt_ka_cd
,T1.skkns_kiktt_ka_ms
,T1.skkns_kiktt_tnknk_kknn_flg
,T1.no1_skkns_kiktt_zpkkn_flg
,T1.no2_skkns_kiktt_zpkkn_flg
,T1.no3_skkns_kiktt_zpkkn_flg
,T1.skkns_kikht_snn_ymd
,T1.skkns_kikht_snsya_cd
,T1.skkns_kikht_snsya_ms
,T1.skkns_kikht_ka_cd
,T1.skkns_kikht_ka_ms
,T1.skkns_kikht_tnknk_kknn_flg
,T1.no1_skkns_kikht_zpkkn_flg
,T1.no2_skkns_kikht_zpkkn_flg
,T1.no3_skkns_kikht_zpkkn_flg
,T1.skkns_kikkt_snn_ymd
,T1.skkns_kikkt_snsya_cd
,T1.skkns_kikkt_snsya_ms
,T1.skkns_kikkt_ka_cd
,T1.skkns_kikkt_ka_ms
,T1.skkns_kikkt_tnknk_kknn_flg
,T1.no1_skkns_kikkt_zpkkn_flg
,T1.no2_skkns_kikkt_zpkkn_flg
,T1.no3_skkns_kikkt_zpkkn_flg
,T1.skkns_knstt_kaih_ymd
,T1.skkns_knstt_kaih_syors_cd
,T1.skkns_knstt_kaih_syors_ms
,T1.skkns_knstt_ka_cd
,T1.skkns_knstt_ka_ms
,T1.skkns_knstt_tnknk_kknn_flg
,T1.no1_skkns_knstt_zpkkn_flg
,T1.no2_skkns_knstt_zpkkn_flg
,T1.no3_skkns_knstt_zpkkn_flg
,T1.skkns_knsht_snn_ymd
,T1.skkns_knsht_snsya_cd
,T1.skkns_knsht_snsya_ms
,T1.skkns_knsht_ka_cd
,T1.skkns_knsht_ka_ms
,T1.skkns_knsht_tnknk_kknn_flg
,T1.no1_skkns_knsht_zpkkn_flg
,T1.no2_skkns_knsht_zpkkn_flg
,T1.no3_skkns_knsht_zpkkn_flg
,T1.skkns_knskt_snn_ymd
,T1.skkns_knskt_snsya_cd
,T1.skkns_knskt_snsya_ms
,T1.skkns_knskt_ka_cd
,T1.skkns_knskt_ka_ms
,T1.skkns_knskt_tnknk_kknn_flg
,T1.no1_skkns_knskt_zpkkn_flg
,T1.no2_skkns_knskt_zpkkn_flg
,T1.no3_skkns_knskt_zpkkn_flg
,T1.skkns_hnn_flg
,T1.skkns_hnn_riyu_naiyo
,T1.krds_sett_tanto_syori_ymd
,T1.krds_sett_tanto_syors_cd
,T1.krds_sett_tanto_syors_ms
,T1.krds_sett_tanto_ka_cd
,T1.krds_sett_tanto_ka_ms
,T1.krds_sett_ht_snn_ymd
,T1.krds_sett_ht_snsya_cd
,T1.krds_sett_ht_snsya_ms
,T1.krds_sett_ht_ka_cd
,T1.krds_sett_ht_ka_ms
,T1.krds_sett_hnn_flg
,T1.nktrk_sizi_syori_ymd
,T1.nktrk_sizi_syors_cd
,T1.nktrk_sizi_syors_ms
,T1.nktrk_sizi_ka_cd
,T1.nktrk_sizi_ka_ms
,T1.nktrk_sizi_kzkis_mdgt_cd
,T1.nktrk_sizi_kzkis_mdgt_ms
,T1.nktrk_sizi_kozitn_cd
,T1.nktrk_sizi_kozitn_ms
,T1.nktrk_tanto_syori_ymd
,T1.nktrk_tanto_syors_cd
,T1.nktrk_tanto_syors_ms
,T1.nktrk_tanto_ka_cd
,T1.nktrk_tanto_ka_ms
,T1.nktrk_tanto_ces_cd
,T1.nktrk_tanto_ces_ms
,T1.nktrk_elder_syori_ymd
,T1.nktrk_elder_syors_cd
,T1.nktrk_elder_syors_ms
,T1.nktrk_elder_ka_cd
,T1.nktrk_elder_ka_ms
,T1.nktrk_elder_ces_cd
,T1.nktrk_elder_ces_ms
,T1.nktrk_hnn_flg
,T1.tidu_ccsz_syori_ymd
,T1.tidu_ccsz_syors_cd
,T1.tidu_ccsz_syors_ms
,T1.tidu_ccsz_ka_cd
,T1.tidu_ccsz_ka_ms
,T1.tidu_ccsz_kkmg_cd
,T1.tidu_ccsz_kkmg_ms
,T1.tidu_ccsz_kozitn_cd
,T1.tidu_ccsz_kozitn_ms
,T1.tidu_cckkn_tanto_syori_ymd
,T1.tidu_cckkn_tanto_syors_cd
,T1.tidu_cckkn_tanto_syors_ms
,T1.tidu_cckkn_tanto_ka_cd
,T1.tidu_cckkn_tanto_ka_ms
,T1.tidu_cckkn_tanto_kkmg_cd
,T1.tidu_cckkn_tanto_kkmg_ms
,T1.tidu_cckkn_tanto_kozitn_cd
,T1.tidu_cckkn_tanto_kozitn_ms
,T1.tidu_cckkn_seko_syori_ymd
,T1.tidu_cckkn_seko_syors_cd
,T1.tidu_cckkn_seko_syors_ms
,T1.tidu_cckkn_seko_ka_cd
,T1.tidu_cckkn_seko_ka_ms
,T1.tidu_cckkn_seko_kkmg_cd
,T1.tidu_cckkn_seko_kkmg_ms
,T1.tidu_cckkn_seko_kozitn_cd
,T1.tidu_cckkn_seko_kozitn_ms
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票行程
  INNER JOIN trke_trkehy_prc T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理詳細・計器_地点_特記事項　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・計器_地点_特記事項　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・計器_地点_特記事項を検索します</remarks>
    Friend Function S004(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.kyokyu_spot_tokti_no
 ,K1.tkzk_renno
 ,K1.tkzk_bnrui_cd
 ,K1.tkzk_kizi_naiyo
 ,K1.itaks_kaizi_kizi_naiyo
 ,K1.kiyk_cs_ms
 ,K1.tenp_file_mng_no
 ,K1.TOROK_YMD
 ,K1.UPD_YMD
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（低圧）
  INNER JOIN trke_trkehy_tiat T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--計器_地点_特記事項
  INNER JOIN keiki_spot_tkzk K1 
ON T1.kyokyu_spot_tokti_no = K1.kyokyu_spot_tokti_no 
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理詳細・取替_取替票追加工費　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票追加工費　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票追加工費を検索します</remarks>
    Friend Function S005(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,T1.saksi_date
 ,T1.upd_date
 ,T1.sousa_user_id
 ,T1.sousa_appli_cd
 ,T1.dig4_zgsyo_cd
 ,T1.trkhy_hakko_nendo
 ,T1.trkhy_kbn
 ,T1.trkhy_no
 ,T1.tuika_kohi_row_no
 ,T1.tuika_kohi_kozi_kbn
 ,T1.tuika_kohi_trtk_suryo
 ,T1.tuika_kohi_tekyo_suryo
 ,T1.trtk_kozih_kngk
 ,T1.tekyo_kozih_kngk
 ,T1.tuika_kohi_riyu_naiyo
 ,case M1.tuikakh_sytk_mst_kbn
    when '1' then M2.koryo_ms
    when '2' then M3.zunit_ms
    else ' ' 
  END As tuika_kohi_kozi_kbn_ms
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票追加工費
  INNER JOIN trke_trkehy_tuikakh T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--マスター利用可能追加工費
  LEFT JOIN mst_use_kano_tuikakh M1
ON T1.tuika_kohi_kozi_kbn = SUBSTRING(M1.koryo_zunit_cd, 4,4)
--マスター工量
  LEFT  JOIN  mst_koryo M2 
ON M2.koryo_cd=M1.koryo_zunit_cd
and M2.tky_start_ymd <= {1} and {1} <= M2.tky_end_ymd
--マスター材料ユニット
  LEFT  JOIN  mst_zunit M3 
ON M3.zunit_cd=M1.koryo_zunit_cd
and M3.tky_start_ymd <= {1} and {1} <= M3.tky_end_ymd
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno) _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD))
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

#Region "(SQL)ダウンロード管理詳細・取替_自主点検チェックシート　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_自主点検チェックシート　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_自主点検チェックシートを検索します</remarks>
    Friend Function S006(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
            'Rev040 2024/11/20 物理分割 次世代QR対応 ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_tio_cmp_ymd
            sb.Append(<sql><![CDATA[
select
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,T1.saksi_date
 ,T1.upd_date
 ,T1.sousa_user_id
 ,T1.sousa_appli_cd
 ,T1.dig4_zgsyo_cd
 ,T1.trkhy_hakko_nendo
 ,T1.trkhy_kbn
 ,T1.trkhy_no
 ,T1.takcd_ttt_skosya_ryohi_kbn
 ,T1.takcd_ttt_ntktt_ryohi_kbn
 ,T1.takcd_ttt_kstts_ryohi_kbn
 ,T1.takcd_ttt_tio_naiyo
 ,T1.takcd_ttt_tio_cmp_ymd
 ,T1.takcd_szsu_skosya_ryohi_kbn
 ,T1.takcd_szsu_ntktt_ryohi_kbn
 ,T1.takcd_szsu_kstts_ryohi_kbn
 ,T1.takcd_szsu_tio_naiyo
 ,T1.takcd_szsu_tio_cmp_ymd
 ,T1.takck_ssda_skosya_ryohi_kbn
 ,T1.takck_ssda_ntktt_ryohi_kbn
 ,T1.takck_ssda_kstts_ryohi_kbn
 ,T1.takck_ssda_tio_naiyo
 ,T1.takck_ssda_tio_cmp_ymd
 ,T1.takck_zyrt_skosya_ryohi_kbn
 ,T1.takck_zyrt_ntktt_ryohi_kbn
 ,T1.takck_zyrt_kstts_ryohi_kbn
 ,T1.takck_zyrt_tio_naiyo
 ,T1.takck_zyrt_tio_cmp_ymd
 ,T1.takck_yr_skosya_ryohi_kbn
 ,T1.takck_yr_ntktt_ryohi_kbn
 ,T1.takck_yr_kstts_ryohi_kbn
 ,T1.takck_yr_tio_naiyo
 ,T1.takck_yr_tio_cmp_ymd
 ,T1.takck_sm_skosya_ryohi_kbn
 ,T1.takck_sm_ntktt_ryohi_kbn
 ,T1.takck_sm_kstts_ryohi_kbn
 ,T1.takck_sm_tio_naiyo
 ,T1.takck_sm_tio_cmp_ymd
 ,T1.takck_kkano_skosya_ryohi_kbn
 ,T1.takck_kkano_ntktt_ryohi_kbn
 ,T1.takck_kkano_kstts_ryohi_kbn
 ,T1.takck_kkano_tio_naiyo
 ,T1.takck_kkano_tio_cmp_ymd
 ,T1.takck_hksno_skosya_ryohi_kbn
 ,T1.takck_hksno_ntktt_ryohi_kbn
 ,T1.takck_hksno_kstts_ryohi_kbn
 ,T1.takck_hksno_tio_naiyo
 ,T1.takck_hksno_tio_cmp_ymd
 ,T1.takck_hkano_skosya_ryohi_kbn
 ,T1.takck_hkano_ntktt_ryohi_kbn
 ,T1.takck_hkano_kstts_ryohi_kbn
 ,T1.takck_hkano_tio_naiyo
 ,T1.takck_hkano_tio_cmp_ymd
 ,T1.takckk_krkkd_skosya_ryohi_kbn
 ,T1.takckk_krkkd_ntktt_ryohi_kbn
 ,T1.takckk_krkkd_kstts_ryohi_kbn
 ,T1.takckk_krkkd_tio_naiyo
 ,T1.takckk_krkkd_tio_cmp_ymd
 ,T1.takckk_krkb_skosya_ryohi_kbn
 ,T1.takckk_krkb_ntktt_ryohi_kbn
 ,T1.takckk_krkb_kstts_ryohi_kbn
 ,T1.takckk_krkb_tio_naiyo
 ,T1.takckk_krkb_tio_cmp_ymd
 ,T1.takcks_kkzen_skosya_ryohi_kbn
 ,T1.takcks_kkzen_ntktt_ryohi_kbn
 ,T1.takcks_kkzen_kstts_ryohi_kbn
 ,T1.takcks_kkzen_tio_naiyo
 ,T1.takcks_kkzen_tio_cmp_ymd
 ,T1.takcks_kksl_skosya_ryohi_kbn
 ,T1.takcks_kksl_ntktt_ryohi_kbn
 ,T1.takcks_kksl_kstts_ryohi_kbn
 ,T1.takcks_kksl_tio_naiyo
 ,T1.takcks_kksl_tio_cmp_ymd
 ,T1.takcks_kkhks_skosya_ryohi_kbn
 ,T1.takcks_kkhks_ntktt_ryohi_kbn
 ,T1.takcks_kkhks_kstts_ryohi_kbn
 ,T1.takcks_kkhks_tio_naiyo
 ,T1.takcks_kkhks_tio_cmp_ymd
 ,T1.takcks_kkgsz_skosya_ryohi_kbn
 ,T1.takcks_kkgsz_ntktt_ryohi_kbn
 ,T1.takcks_kkgsz_kstts_ryohi_kbn
 ,T1.takcks_kkgsz_tio_naiyo
 ,T1.takcks_kkgsz_tio_cmp_ymd
 ,T1.takcks_kkaki_skosya_ryohi_kbn
 ,T1.takcks_kkaki_ntktt_ryohi_kbn
 ,T1.takcks_kkaki_kstts_ryohi_kbn
 ,T1.takcks_kkaki_tio_naiyo
 ,T1.takcks_kkaki_tio_cmp_ymd
 ,T1.takcks_kkyrs_skosya_ryohi_kbn
 ,T1.takcks_kkyrs_ntktt_ryohi_kbn
 ,T1.takcks_kkyrs_kstts_ryohi_kbn
 ,T1.takcks_kkyrs_tio_naiyo
 ,T1.takcks_kkyrs_tio_cmp_ymd
 ,T1.takcks_kkrst_skosya_ryohi_kbn
 ,T1.takcks_kkrst_ntktt_ryohi_kbn
 ,T1.takcks_kkrst_kstts_ryohi_kbn
 ,T1.takcks_kkrst_tio_naiyo
 ,T1.takcks_kkrst_tio_cmp_ymd
 ,T1.takcks_kkbst_skosya_ryohi_kbn
 ,T1.takcks_kkbst_ntktt_ryohi_kbn
 ,T1.takcks_kkbst_kstts_ryohi_kbn
 ,T1.takcks_kkbst_tio_naiyo
 ,T1.takcks_kkbst_tio_cmp_ymd
 ,T1.takcks_kikib_skosya_ryohi_kbn
 ,T1.takcks_kikib_ntktt_ryohi_kbn
 ,T1.takcks_kikib_kstts_ryohi_kbn
 ,T1.takcks_kikib_tio_naiyo
 ,T1.takcks_kikib_tio_cmp_ymd
 ,T1.takcks_kkssu_skosya_ryohi_kbn
 ,T1.takcks_kkssu_ntktt_ryohi_kbn
 ,T1.takcks_kkssu_kstts_ryohi_kbn
 ,T1.takcks_kkssu_tio_naiyo
 ,T1.takcks_kkssu_tio_cmp_ymd
 ,T1.takcks_hrkkl_skosya_ryohi_kbn
 ,T1.takcks_hrkkl_ntktt_ryohi_kbn
 ,T1.takcks_hrkkl_kstts_ryohi_kbn
 ,T1.takcks_hrkkl_tio_naiyo
 ,T1.takcks_hrkkl_tio_cmp_ymd
 ,T1.takcks_hp123_skosya_ryohi_kbn
 ,T1.takcks_hp123_ntktt_ryohi_kbn
 ,T1.takcks_hp123_kstts_ryohi_kbn
 ,T1.takcks_hp123_tio_naiyo
 ,T1.takcks_hp123_tio_cmp_ymd
 ,T1.takcks_hkaki_skosya_ryohi_kbn
 ,T1.takcks_hkaki_ntktt_ryohi_kbn
 ,T1.takcks_hkaki_kstts_ryohi_kbn
 ,T1.takcks_hkaki_tio_naiyo
 ,T1.takcks_hkaki_tio_cmp_ymd
 ,T1.takckrk_hnrk_skosya_ryohi_kbn
 ,T1.takckrk_hnrk_ntktt_ryohi_kbn
 ,T1.takckrk_hnrk_kstts_ryohi_kbn
 ,T1.takckrk_hnrk_tio_naiyo
 ,T1.takckrk_hnrk_tio_cmp_ymd
 ,T1.takckrr_sm_skosya_ryohi_kbn
 ,T1.takckrr_sm_ntktt_ryohi_kbn
 ,T1.takckrr_sm_kstts_ryohi_kbn
 ,T1.takckrr_sm_tio_naiyo
 ,T1.takckrr_sm_tio_cmp_ymd
 ,T1.takcc_led_skosya_ryohi_kbn
 ,T1.takcc_led_ntktt_ryohi_kbn
 ,T1.takcc_led_kstts_ryohi_kbn
 ,T1.takcc_led_tio_naiyo
 ,T1.takcc_led_tio_cmp_ymd
 ,T1.takcc_souck_skosya_ryohi_kbn
 ,T1.takcc_souck_ntktt_ryohi_kbn
 ,T1.takcc_souck_kstts_ryohi_kbn
 ,T1.takcc_souck_tio_naiyo
 ,T1.takcc_souck_tio_cmp_ymd
 ,T1.takct_szno_skosya_ryohi_kbn
 ,T1.takct_szno_ntktt_ryohi_kbn
 ,T1.takct_szno_kstts_ryohi_kbn
 ,T1.takct_szno_tio_naiyo
 ,T1.takct_szno_tio_cmp_ymd
 ,T1.takct_aino_skosya_ryohi_kbn
 ,T1.takct_aino_ntktt_ryohi_kbn
 ,T1.takct_aino_kstts_ryohi_kbn
 ,T1.takct_aino_tio_naiyo
 ,T1.takct_aino_tio_cmp_ymd
 ,T1.takcs_photo_skosya_ryohi_kbn
 ,T1.takcs_photo_ntktt_ryohi_kbn
 ,T1.takcs_photo_kstts_ryohi_kbn
 ,T1.takcs_photo_tio_naiyo
 ,T1.takcs_photo_tio_cmp_ymd
 ,T1.takcs_tktrk_skosya_ryohi_kbn
 ,T1.takcs_tktrk_ntktt_ryohi_kbn
 ,T1.takcs_tktrk_kstts_ryohi_kbn
 ,T1.takcs_tktrk_tio_naiyo
 ,T1.takcs_tktrk_tio_cmp_ymd
 ,T1.takcz_ntktt_seigo_kknn_kbn
 ,T1.takcz_kstts_seigo_kknn_kbn
 ,T1.takcz_tio_naiyo
 ,T1.takcz_tio_cmp_ymd
 ,T1.takca_ntktt_itti_kknn_kbn
 ,T1.takca_kstts_itti_kknn_kbn
 ,T1.takca_tio_naiyo
 ,T1.takca_tio_cmp_ymd
 ,T1.kizi_skosya_naiyo
 ,T1.kizi_ntktt_naiyo
 ,T1.kizi_kstts_naiyo
 ,T1.hrktkk_skosya_1s_kstckk_flg
 ,T1.hrktkk_skosya_p1_kstckk_flg
 ,T1.hrktkk_skosya_p3_kstckk_flg
 ,T1.hrktkk_skosya_3s_kstckk_flg
 ,T1.hrktkk_skosya_3l_kstckk_flg
 ,T1.hrktkk_skosya_p2_kstckk_flg
 ,T1.hrktkk_skosya_1l_kstckk_flg
 ,T1.p1_skosya_1s_kstckk_flg
 ,T1.p1_skosya_p1_kstckk_flg
 ,T1.p1_skosya_1l_kstckk_flg
 ,T1.p2_skosya_p2_kstckk_flg
 ,T1.p3_skosya_3s_kstckk_flg
 ,T1.p3_skosya_p3_kstckk_flg
 ,T1.p3_skosya_3l_kstckk_flg
 ,T1.hrktkk_ntktt_1s_kstckk_flg
 ,T1.hrktkk_ntktt_p1_kstckk_flg
 ,T1.hrktkk_ntktt_p3_kstckk_flg
 ,T1.hrktkk_ntktt_3s_kstckk_flg
 ,T1.hrktkk_ntktt_3l_kstckk_flg
 ,T1.hrktkk_ntktt_p2_kstckk_flg
 ,T1.hrktkk_ntktt_1l_kstckk_flg
 ,T1.p1_ntktt_1s_kstckk_flg
 ,T1.p1_ntktt_p1_kstckk_flg
 ,T1.p1_ntktt_1l_kstckk_flg
 ,T1.p2_ntktt_p2_kstckk_flg
 ,T1.p3_ntktt_3s_kstckk_flg
 ,T1.p3_ntktt_p3_kstckk_flg
 ,T1.p3_ntktt_3l_kstckk_flg
 ,T1.hrktkk_kstts_1s_kstckk_flg
 ,T1.hrktkk_kstts_p1_kstckk_flg
 ,T1.hrktkk_kstts_p3_kstckk_flg
 ,T1.hrktkk_kstts_3s_kstckk_flg
 ,T1.hrktkk_kstts_3l_kstckk_flg
 ,T1.hrktkk_kstts_p2_kstckk_flg
 ,T1.hrktkk_kstts_1l_kstckk_flg
 ,T1.p1_kstts_1s_kstckk_flg
 ,T1.p1_kstts_p1_kstckk_flg
 ,T1.p1_kstts_1l_kstckk_flg
 ,T1.p2_kstts_p2_kstckk_flg
 ,T1.p3_kstts_3s_kstckk_flg
 ,T1.p3_kstts_p3_kstckk_flg
 ,T1.p3_kstts_3l_kstckk_flg
 ,T1.takckrr_smhyz_skosya_ryohi_kbn 
 ,T1.takckrr_smhyz_ntktt_ryohi_kbn 
 ,T1.takckrr_smhyz_kstts_ryohi_kbn 
 ,T1.takckrr_smhyz_tio_naiyo 
 ,T1.takckrr_smhyz_tio_cmp_ymd 
 ,T1.takckrr_d2sm_skosya_ryohi_kbn 
 ,T1.takckrr_d2sm_ntktt_ryohi_kbn
 ,T1.takckrr_d2sm_kstts_ryohi_kbn
 ,T1.takckrr_d2sm_tio_naiyo
 ,T1.takckrr_d2sm_tio_cmp_ymd
 ,T1.takckrr_khk_skosya_ryohi_kbn
 ,T1.takckrr_khk_ntktt_ryohi_kbn 
 ,T1.takckrr_khk_kstts_ryohi_kbn 
 ,T1.takckrr_khk_tio_naiyo 
 ,T1.takckrr_khk_tio_cmp_ymd 
 ,T1.takckk_okgkb_skosya_ryohi_kbn 
 ,T1.takckk_okgkb_ntktt_ryohi_kbn
 ,T1.takckk_okgkb_kstts_ryohi_kbn
 ,T1.takckk_okgkb_tio_naiyo
 ,T1.takckk_okgkb_tio_cmp_ymd
 ,T1.takckk_dsipcv_skosya_ryohi_kbn
 ,T1.takckk_dsipcv_ntktt_ryohi_kbn 
 ,T1.takckk_dsipcv_kstts_ryohi_kbn 
 ,T1.takckk_dsipcv_tio_naiyo 
 ,T1.takckk_dsipcv_tio_cmp_ymd 
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_自主点検チェックシート
  INNER JOIN trke_check_sheet T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理詳細・共通_添付書類詳細　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・共通_添付書類詳細　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・共通_添付書類詳細を検索します</remarks>
    Friend Function S007(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
            'Rev002.1 MOD Start 20241128 高圧添付書類の修正
            sb.Append(<sql><![CDATA[
select TP.*
FROM (
(
select
  D1.dig4_zgsyo_cd AS dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS trkhy_hakko_nendo
 ,D1.trkhy_kbn AS trkhy_kbn
 ,D1.trkhy_no As trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.tnpdoc_mng_renno
 ,K1.file_renno
 ,K1.file_ms
 ,K1.no1_tnpdoc_hosok_naiyo
 ,K1.no2_tnpdoc_hosok_naiyo
 ,K1.no3_tnpdoc_hosok_naiyo
 ,K1.no4_tnpdoc_hosok_naiyo
 ,K1.no5_tnpdoc_hosok_naiyo
 ,K1.no6_tnpdoc_hosok_naiyo
 ,K1.no7_tnpdoc_hosok_naiyo
 ,K1.no8_tnpdoc_hosok_naiyo
 ,K1.no9_tnpdoc_hosok_naiyo
 ,K1.no10_tnpdoc_hosok_naiyo
 ,M1.prm_cd_ms AS photo_sbt_ms
 ,T1.tenp_file_mng_no
 ,T1.ryssy_tenp_file_mng_no
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（低圧）
  INNER JOIN trke_trkehy_tiat T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--共通_添付書類詳細
  INNER JOIN com_tnpdoc_dtl K1 
ON T1.tenp_file_mng_no = K1.tnpdoc_mng_renno 
--パラメータコードマスタ(写真種別)
  LEFT  JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'PHOTO_SBT_CD' 
   AND M1.prm_cd = K1.no2_tnpdoc_hosok_naiyo
WHERE 1 = 1 
AND D1.downld_renno={0}
)
UNION
(
select
  D1.dig4_zgsyo_cd AS dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS trkhy_hakko_nendo
 ,D1.trkhy_kbn AS trkhy_kbn
 ,D1.trkhy_no As trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.tnpdoc_mng_renno
 ,K1.file_renno
 ,K1.file_ms
 ,K1.no1_tnpdoc_hosok_naiyo
 ,K1.no2_tnpdoc_hosok_naiyo
 ,K1.no3_tnpdoc_hosok_naiyo
 ,K1.no4_tnpdoc_hosok_naiyo
 ,K1.no5_tnpdoc_hosok_naiyo
 ,K1.no6_tnpdoc_hosok_naiyo
 ,K1.no7_tnpdoc_hosok_naiyo
 ,K1.no8_tnpdoc_hosok_naiyo
 ,K1.no9_tnpdoc_hosok_naiyo
 ,K1.no10_tnpdoc_hosok_naiyo
 ,M1.prm_cd_ms AS photo_sbt_ms
 ,T1.tenp_file_mng_no
 ,T1.ryssy_tenp_file_mng_no
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（高圧他）
  INNER JOIN trke_trkehy_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--共通_添付書類詳細
  INNER JOIN com_tnpdoc_dtl K1 
ON T1.tenp_file_mng_no = K1.tnpdoc_mng_renno 
--パラメータコードマスタ(写真種別)
  LEFT  JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'PHOTO_SBT_CD' 
   AND M1.prm_cd = K1.no2_tnpdoc_hosok_naiyo
WHERE 1 = 1 
AND D1.downld_renno={0}
)
) TP
ORDER BY 
  TP.dig4_zgsyo_cd
 ,TP.trkhy_hakko_nendo
 ,TP.trkhy_kbn
 ,TP.trkhy_no
 ,TP.file_renno
        ]]></sql>.Value)
            'Rev002.1 MOD End 20241128

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

    'Rev002-Start 


#Region "(SQL)ダウンロード管理詳細・取替_取替票行程（高圧他）（低圧契約）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票行程（高圧他）（低圧契約）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票行程を検索します</remarks>
    Friend Function S023(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.kkst_tanto_kaih_ymd
,T1.kkst_tanto_kaih_syors_cd
,T1.kkst_tanto_kaih_syors_ms
,T1.kkst_tanto_ka_cd
,T1.kkst_tanto_ka_ms
,T1.kkst_ht_snn_ymd
,T1.kkst_ht_snsya_cd
,T1.kkst_ht_snsya_ms
,T1.kkst_ht_ka_cd
,T1.kkst_ht_ka_ms
,T1.kkst_hnn_flg
,T1.tkpln_tanto_khiri_ymd
,T1.tkpln_tanto_khiri_syors_cd
,T1.tkpln_tanto_khiri_syors_ms
,T1.tkpln_tanto_ka_cd
,T1.tkpln_tanto_ka_ms
,T1.tkpln_tanto_kzkis_mdgt_cd
,T1.tkpln_tanto_kzkis_mdgt_ms
,T1.tkpln_tanto_kozitn_cd
,T1.tkpln_tanto_kozitn_ms
,T1.tkpln_tanto_khiri_hnn_flg
,T1.khsz_tanto_kaih_ymd
,T1.khsz_tanto_kaih_syors_cd
,T1.khsz_tanto_kaih_syors_ms
,T1.khsz_tanto_ka_cd
,T1.khsz_tanto_ka_ms
,T1.khsz_hnn_flg
,T1.seko_kti_ht_snn_ymd
,T1.seko_kti_ht_snsya_cd
,T1.seko_kti_ht_snsya_ms
,T1.seko_kti_ht_ka_cd
,T1.seko_kti_ht_ka_ms
,T1.seko_kti_kt_snn_ymd
,T1.seko_kti_kt_snsya_cd
,T1.seko_kti_kt_snsya_ms
,T1.seko_kti_kt_ka_cd
,T1.seko_kti_kt_ka_ms
,T1.seko_kti_hnn_flg
,T1.skohu_tanto_kaih_syori_ymd
,T1.skohu_tanto_kaih_syors_cd
,T1.skohu_tanto_kaih_syors_ms
,T1.skohu_tanto_ka_cd
,T1.skohu_tanto_ka_ms
,T1.skohu_ht_snn_ymd
,T1.skohu_ht_snsya_cd
,T1.skohu_ht_snsya_ms
,T1.skohu_ht_ka_cd
,T1.skohu_ht_ka_ms
,T1.skohu_kt_snn_ymd
,T1.skohu_kt_snsya_cd
,T1.skohu_kt_snsya_ms
,T1.skohu_kt_ka_cd
,T1.skohu_kt_ka_ms
,T1.skohu_hnn_flg
,T1.zkskh_tanto_kaih_syori_ymd
,T1.zkskh_tanto_kaih_syors_cd
,T1.zkskh_tanto_kaih_syors_ms
,T1.zkskh_tanto_ka_cd
,T1.zkskh_tanto_ka_ms
,T1.zkskh_ht_snn_ymd
,T1.zkskh_ht_snsya_cd
,T1.zkskh_ht_snsya_ms
,T1.zkskh_ht_ka_cd
,T1.zkskh_ht_ka_ms
,T1.zkskh_kt_snn_ymd
,T1.zkskh_kt_snsya_cd
,T1.zkskh_kt_snsya_ms
,T1.zkskh_kt_ka_cd
,T1.zkskh_kt_ka_ms
,T1.zkskh_hnn_flg
,T1.skyti_ktst_syori_ymd
,T1.skyti_ktst_syors_cd
,T1.skyti_ktst_syors_ms
,T1.skyti_ktst_ka_cd
,T1.skyti_ktst_ka_ms
,T1.skyti_ktst_kzkis_mdgt_cd
,T1.skyti_ktst_kzkis_mdgt_ms
,T1.skyti_ktst_kozitn_cd
,T1.skyti_ktst_kozitn_ms
,T1.skyti_skyti_inpt_syori_ymd
,T1.skyti_skyti_inpt_syors_cd
,T1.skyti_skyti_inpt_syors_ms
,T1.skyti_skyti_inpt_ka_cd
,T1.skyti_skyti_inpt_ka_ms
,T1.skyti_skyti_inpt_kkmg_cd
,T1.skyti_skyti_inpt_kkmg_ms
,T1.skyti_skyti_inpt_kozitn_cd
,T1.skyti_skyti_inpt_kozitn_ms
,T1.seko_tanto_syori_ymd
,T1.seko_tanto_syors_cd
,T1.seko_tanto_syors_ms
,T1.seko_tanto_ka_cd
,T1.seko_tanto_ka_ms
,T1.seko_tanto_kzkis_mdgt_cd
,T1.seko_tanto_kzkis_mdgt_ms
,T1.seko_tanto_kozitn_cd
,T1.seko_tanto_kozitn_ms
,T1.seko_elder_tnknk_kknn_flg
,T1.seko_elder_tnkn_kknn_ymd
,T1.seko_elder_tnkn_knsya_cd
,T1.seko_elder_tnkn_knsya_ms
,T1.seko_elder_ka_cd
,T1.seko_elder_ka_ms
,T1.seko_elder_kzkis_mdgt_cd
,T1.seko_elder_kzkis_mdgt_ms
,T1.seko_elder_kozitn_cd
,T1.seko_elder_kozitn_ms
,T1.skrepo_syori_ymd
,T1.skrepo_syors_cd
,T1.skrepo_syors_ms
,T1.skrepo_ka_cd
,T1.skrepo_ka_ms
,T1.skrepo_kzkis_mdgt_cd
,T1.skrepo_kzkis_mdgt_ms
,T1.skrepo_kozitn_cd
,T1.skrepo_kozitn_ms
,T1.no1_skrepo_zippi_kknn_flg
,T1.no2_skrepo_zippi_kknn_flg
,T1.no3_skrepo_zippi_kknn_flg
,T1.skkns_kiktt_kaih_ymd
,T1.skkns_kiktt_kaih_syors_cd
,T1.skkns_kiktt_kaih_syors_ms
,T1.skkns_kiktt_ka_cd
,T1.skkns_kiktt_ka_ms
,T1.skkns_kiktt_tnknk_kknn_flg
,T1.no1_skkns_kiktt_zpkkn_flg
,T1.no2_skkns_kiktt_zpkkn_flg
,T1.no3_skkns_kiktt_zpkkn_flg
,T1.skkns_kikht_snn_ymd
,T1.skkns_kikht_snsya_cd
,T1.skkns_kikht_snsya_ms
,T1.skkns_kikht_ka_cd
,T1.skkns_kikht_ka_ms
,T1.skkns_kikht_tnknk_kknn_flg
,T1.no1_skkns_kikht_zpkkn_flg
,T1.no2_skkns_kikht_zpkkn_flg
,T1.no3_skkns_kikht_zpkkn_flg
,T1.skkns_kikkt_snn_ymd
,T1.skkns_kikkt_snsya_cd
,T1.skkns_kikkt_snsya_ms
,T1.skkns_kikkt_ka_cd
,T1.skkns_kikkt_ka_ms
,T1.skkns_kikkt_tnknk_kknn_flg
,T1.no1_skkns_kikkt_zpkkn_flg
,T1.no2_skkns_kikkt_zpkkn_flg
,T1.no3_skkns_kikkt_zpkkn_flg
,T1.skkns_knstt_kaih_ymd
,T1.skkns_knstt_kaih_syors_cd
,T1.skkns_knstt_kaih_syors_ms
,T1.skkns_knstt_ka_cd
,T1.skkns_knstt_ka_ms
,T1.skkns_knstt_tnknk_kknn_flg
,T1.no1_skkns_knstt_zpkkn_flg
,T1.no2_skkns_knstt_zpkkn_flg
,T1.no3_skkns_knstt_zpkkn_flg
,T1.skkns_knsht_snn_ymd
,T1.skkns_knsht_snsya_cd
,T1.skkns_knsht_snsya_ms
,T1.skkns_knsht_ka_cd
,T1.skkns_knsht_ka_ms
,T1.skkns_knsht_tnknk_kknn_flg
,T1.no1_skkns_knsht_zpkkn_flg
,T1.no2_skkns_knsht_zpkkn_flg
,T1.no3_skkns_knsht_zpkkn_flg
,T1.skkns_knskt_snn_ymd
,T1.skkns_knskt_snsya_cd
,T1.skkns_knskt_snsya_ms
,T1.skkns_knskt_ka_cd
,T1.skkns_knskt_ka_ms
,T1.skkns_knskt_tnknk_kknn_flg
,T1.no1_skkns_knskt_zpkkn_flg
,T1.no2_skkns_knskt_zpkkn_flg
,T1.no3_skkns_knskt_zpkkn_flg
,T1.skkns_hnn_flg
,T1.skkns_hnn_riyu_naiyo
,T1.krds_sett_tanto_syori_ymd
,T1.krds_sett_tanto_syors_cd
,T1.krds_sett_tanto_syors_ms
,T1.krds_sett_tanto_ka_cd
,T1.krds_sett_tanto_ka_ms
,T1.krds_sett_ht_snn_ymd
,T1.krds_sett_ht_snsya_cd
,T1.krds_sett_ht_snsya_ms
,T1.krds_sett_ht_ka_cd
,T1.krds_sett_ht_ka_ms
,T1.krds_sett_hnn_flg
,T1.nktrk_sizi_syori_ymd
,T1.nktrk_sizi_syors_cd
,T1.nktrk_sizi_syors_ms
,T1.nktrk_sizi_ka_cd
,T1.nktrk_sizi_ka_ms
,T1.nktrk_sizi_kzkis_mdgt_cd
,T1.nktrk_sizi_kzkis_mdgt_ms
,T1.nktrk_sizi_kozitn_cd
,T1.nktrk_sizi_kozitn_ms
,T1.nktrk_tanto_syori_ymd
,T1.nktrk_tanto_syors_cd
,T1.nktrk_tanto_syors_ms
,T1.nktrk_tanto_ka_cd
,T1.nktrk_tanto_ka_ms
,T1.nktrk_tanto_ces_cd
,T1.nktrk_tanto_ces_ms
,T1.nktrk_elder_syori_ymd
,T1.nktrk_elder_syors_cd
,T1.nktrk_elder_syors_ms
,T1.nktrk_elder_ka_cd
,T1.nktrk_elder_ka_ms
,T1.nktrk_elder_ces_cd
,T1.nktrk_elder_ces_ms
,T1.nktrk_hnn_flg
,T1.tidu_ccsz_syori_ymd
,T1.tidu_ccsz_syors_cd
,T1.tidu_ccsz_syors_ms
,T1.tidu_ccsz_ka_cd
,T1.tidu_ccsz_ka_ms
,T1.tidu_ccsz_kkmg_cd
,T1.tidu_ccsz_kkmg_ms
,T1.tidu_ccsz_kozitn_cd
,T1.tidu_ccsz_kozitn_ms
,T1.tidu_cckkn_tanto_syori_ymd
,T1.tidu_cckkn_tanto_syors_cd
,T1.tidu_cckkn_tanto_syors_ms
,T1.tidu_cckkn_tanto_ka_cd
,T1.tidu_cckkn_tanto_ka_ms
,T1.tidu_cckkn_tanto_kkmg_cd
,T1.tidu_cckkn_tanto_kkmg_ms
,T1.tidu_cckkn_tanto_kozitn_cd
,T1.tidu_cckkn_tanto_kozitn_ms
,T1.tidu_cckkn_seko_syori_ymd
,T1.tidu_cckkn_seko_syors_cd
,T1.tidu_cckkn_seko_syors_ms
,T1.tidu_cckkn_seko_ka_cd
,T1.tidu_cckkn_seko_ka_ms
,T1.tidu_cckkn_seko_kkmg_cd
,T1.tidu_cckkn_seko_kkmg_ms
,T1.tidu_cckkn_seko_kozitn_cd
,T1.tidu_cckkn_seko_kozitn_ms
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票行程（高圧他）
  INNER JOIN trke_trkehy_prc_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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


#Region "(SQL)ダウンロード管理詳細・取替_取替票追加工費（高圧他）（低圧契約）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票追加工費（高圧他）（低圧契約）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票追加工費（高圧他）（低圧契約）を検索します</remarks>
    Friend Function S025(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,T1.saksi_date
 ,T1.upd_date
 ,T1.sousa_user_id
 ,T1.sousa_appli_cd
 ,T1.dig4_zgsyo_cd
 ,T1.trkhy_hakko_nendo
 ,T1.trkhy_kbn
 ,T1.trkhy_no
 ,T1.tuika_kohi_row_no
 ,T1.tuika_kohi_kozi_kbn
 ,T1.tuika_kohi_trtk_suryo
 ,T1.tuika_kohi_tekyo_suryo
 ,T1.trtk_kozih_kngk
 ,T1.tekyo_kozih_kngk
 ,T1.tuika_kohi_riyu_naiyo
 ,case M1.tuikakh_sytk_mst_kbn
    when '1' then M2.koryo_ms
    when '2' then M3.zunit_ms
    else ' ' 
  END As tuika_kohi_kozi_kbn_ms
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票追加工費（高圧他）（低圧契約）
  INNER JOIN trke_trkehy_tuikakh_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--マスター利用可能追加工費
  LEFT JOIN mst_use_kano_tuikakh M1
ON T1.tuika_kohi_kozi_kbn = SUBSTRING(M1.koryo_zunit_cd, 4,4)
--マスター工量
  LEFT  JOIN  mst_koryo M2 
ON M2.koryo_cd=M1.koryo_zunit_cd
and M2.tky_start_ymd <= {1} and {1} <= M2.tky_end_ymd
--マスター材料ユニット
  LEFT  JOIN  mst_zunit M3 
ON M3.zunit_cd=M1.koryo_zunit_cd
and M3.tky_start_ymd <= {1} and {1} <= M3.tky_end_ymd
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno) _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD))
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



#Region "(SQL)ダウンロード管理詳細・取替_取替票(高圧他)　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票(高圧他)　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票(高圧他)を検索します</remarks>
    Friend Function S016(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
            'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD tkkk_tnsb_seizo_ym ～ kdkk_tnsb_maker_mng
            '持出ファイル状態コード='1'(持出指示)を検索する。
            'Rev084 Rev084_取替票(高圧)ダウンロードエラー対応 計器_特記事項(高圧)は別ファイルのためJOINを削除する
            sb.Append(<sql><![CDATA[
select
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.prcmg_cd
,T1.kiyk_no
,T1.kyokyu_spot_tokti_no
,T1.yuko_kigen_ym
,T1.trke_sbt_cd
,T1.cs_tdhkn_add_cd
,T1.cs_siku_add_cd
,T1.cs_oazat_add_cd
,T1.cs_azatm_add_cd
,T1.kozi_corp_cd
,T1.kzkis_mdgt_cd
,T1.kozitn_cd
,T1.skosya_cd
,T1.skosya_ms
,T1.kohu_ymd
,T1.skohu_ymd
,T1.hiky_yotei_ymd
,T1.sksti_ymd
,T1.skssti_ymd
,T1.skyti_ymd
,T1.syun_ymd
,T1.trkhy_hinpt_flg
,T1.seko_kka_tenso_ymd
,T1.skrepo_kka_kbn
,T1.knsgk_ymd
,T1.keiki_sitei_no
,T1.eigyo_rnrk_flg
,T1.brot_sett_flg
,T1.krds_sett_zyokyo_cd
,T1.hory_flg
,T1.nktr_kensa_zyokyo_cd
,T1.NKTRK_TBLT_RECV_YMD
,T1.NKTRK_TBLT_RECV_TNTSY_CD
,T1.NKTRK_zissi_ymd
,T1.NKTRK_zissi_tntsy_cd
,T1.NKTRK_kka_upld_ymd
,T1.NKTRK_kka_upld_tntsy_cd
,T1.thyhkn_syori_ymd
,T1.skohu_mae_prcmg_cd
,T1.skohu_kzkis_mdgt_cd
,T1.skohu_skssti_ymd
,T1.skohu_hiky_yotei_ymd
,T1.skohu_knti_sbt_cd
,T1.tidu_cclme_prcmg_cd
,T1.tidu_cancl_ymd
,T1.keiki_kmaws_no
,T1.gnp_no
,T1.tkkk_keiki_id
,T1.tkkk_kkinf_id
,T1.tkkk_kesbt_kbn
,T1.tkkk_krhsk_cd
,T1.tkkk_keiki_SSNSK_SO_cd
,T1.tkkk_keiki_SSNSK_LINE_cd
,T1.tkkk_kiry_dnat_cd
,T1.tkkk_keiki_yr
,T1.tkkk_tshsk_cd
,CASE T1.tkkk_keiki_ktsk_ms
   WHEN ' ' THEN (SELECT keiki_ktsk_ms FROM mst_keiki_ktsk WHERE 1 = 1 AND ktsk_cd = T1.tkkk_keiki_ktsk_cd AND ksyu_cd = T1.tkkk_ksyu_cd LIMIT 1)
   ELSE T1.tkkk_keiki_ktsk_ms
   END AS tkkk_keiki_ktsk_ms
,T1.tkkk_keiki_ktsk_cd
,T1.tkkk_ksyu_cd
,T1.tkkk_taiko_kbn
,T1.tkkk_knthk_cd
,T1.tkkk_ts_kino_umu_flg
,T1.tkkk_khk_kino_umu_flg
,T1.tkkk_gaibu_output_umu_flg
,T1.tkkk_yuko_kigen_ym
,T1.tkkk_seizo_yy
,T1.tkkk_zyrt
,T1.tkkk_digsu
,T1.tkkk_tnsb_szsya_cd
,T1.tkkk_tnsb_seizo_yy
,T1.tkkk_zytr_szsu
,T1.tkkk_gytr_szsu
,T1.tkkk_sougo_szsu
,T1.tkkk_max_zyyo_szsu
,T1.tkkk_rsyk_szsu
,T1.tkkk_rsmk_szsu
,T1.tkkk_szsu_sytk_zumi_flg
,T1.tkkk_hnsik_ksyu_cd
,T1.no1_tkkk_hnsik_seizo_no
,T1.no2_tkkk_hnsik_seizo_no
,T1.tkkk_hnsik_VCTCT_ktsk_cd
,T1.tkkk_hnsik_VCTCT_seizo_no
,T1.tkkk_hnsik_VCTCT_seizo_yy
,T1.tkkk_hnsik_ct2_ktsk_cd
,T1.tkkk_hnsik_ct2_seizo_no
,T1.tkkk_hnsik_ct2_seizo_yy
,T1.tkkk_hnsik_VT1_ktsk_cd
,T1.tkkk_hnsik_VT1_seizo_no
,T1.tkkk_hnsik_VT1_seizo_yy
,T1.tkkk_hnsik_VT2_ktsk_cd
,T1.tkkk_hnsik_VT2_seizo_no
,T1.tkkk_hnsik_VT2_seizo_yy
,T1.tkkk_HYZKI_seizo_no
,T1.tkkk_hnsik_yuko_kigen_ym
,T1.tkkk_hnsik_gknti_no
,T1.tkkk_mado_su
,T1.ttkk_keiki_id
,T1.ttkk_kesbt_kbn
,T1.ttkk_krhsk_cd
,T1.ttkk_keiki_yr
,T1.ttkk_keiki_SSNSK_SO_cd
,T1.ttkk_keiki_SSNSK_LINE_cd
,T1.ttkk_kiry_dnat_cd
,T1.ttkk_tshsk_cd
,T1.ttkk_keiki_ktsk_ms
,T1.ttkk_keiki_ktsk_cd
,T1.ttkk_ksyu_cd
,T1.ttkk_taiko_kbn
,T1.ttkk_knthk_cd
,T1.ttkk_ts_kino_umu_flg
,T1.ttkk_khk_kino_umu_flg
,T1.ttkk_gaibu_output_umu_flg
,T1.ttkk_yuko_kigen_ym
,T1.ttkk_seizo_yy
,T1.ttkk_zyrt
,T1.ttkk_digsu
,T1.ttkk_tnsb_szsya_cd
,T1.ttkk_tnsb_seizo_yy
,T1.ttkk_zytr_szsu
,T1.ttkk_gytr_szsu
,T1.ttkk_sougo_szsu
,T1.ttkk_rsyk_szsu
,T1.ttkk_rsmk_szsu
,T1.ttkk_hnsik_ksyu_cd
,T1.no1_ttkk_hnsik_seizo_no
,T1.no2_ttkk_hnsik_seizo_no
,T1.ttkk_hnsik_VCTCT_ktsk_cd
,T1.ttkk_hnsik_VCTCT_seizo_no
,T1.ttkk_hnsik_VCTCT_seizo_yy
,T1.ttkk_hnsik_ct2_ktsk_cd
,T1.ttkk_hnsik_ct2_seizo_no
,T1.ttkk_hnsik_ct2_seizo_yy
,T1.ttkk_hnsik_VT1_ktsk_cd
,T1.ttkk_hnsik_VT1_seizo_no
,T1.ttkk_hnsik_VT1_seizo_yy
,T1.ttkk_hnsik_VT2_ktsk_cd
,T1.ttkk_hnsik_VT2_seizo_no
,T1.ttkk_hnsik_VT2_seizo_yy
,T1.TTkk_HYZKI_seizo_no
,T1.ttkk_hnsik_yuko_kigen_ym
,T1.ttkk_hnsik_gknti_no
,T1.ttkk_mado_su
,T1.HZKST_TSTNMT_info_TSN_id
,T1.HZKST_TSTNMT_info_TSHSK_CD
,T1.HZKST_HTAN_info_ksyu_cd
,T1.hzkst_htan_info_kesbt_cd
,T1.HZKST_HTAN_info_ktsk_cd
,T1.HZKST_HTAN_info_seizo_no
,T1.HZKST_HTAN_info_seizo_yy
,T1.HZKST_ts_info_ksyu_cd
,T1.HZKST_ts_info_ktsk_cd
,T1.HZKST_ts_info_seizo_no
,T1.HZKST_ts_info_seizo_yy
,T1.kdkk_keiki_id
,T1.kdkk_kesbt_kbn
,T1.kdkk_krhsk_cd
,T1.kdkk_keiki_SSNSK_SO_cd
,T1.kdkk_keiki_SSNSK_LINE_cd
,T1.kdkk_kiry_dnat_cd
,T1.kdkk_keiki_yr
,T1.kdkk_tshsk_cd
,T1.kdkk_keiki_ktsk_ms
,T1.kdkk_keiki_ktsk_cd
,T1.kdkk_taiko_kbn
,T1.kdkk_knthk_cd
,T1.kdkk_ts_kino_umu_flg
,T1.kdkk_khk_kino_umu_flg
,T1.kdkk_gaibu_output_umu_flg
,T1.kdkk_yuko_kigen_ym
,T1.kdkk_seizo_yy
,T1.kdkk_zyrt
,T1.kdkk_digsu
,T1.kdkk_tnsb_szsya_cd
,T1.kdkk_tnsb_seizo_yy
,T1.kdkk_zytr_szsu
,T1.kdkk_gytr_szsu
,T1.no1_kdkk_hnsik_seizo_no
,T1.no2_kdkk_hnsik_seizo_no
,T1.kdkk_hnsik_VCTCT_seizo_no
,T1.kdkk_hnsik_ct2_seizo_no
,T1.kdkk_hnsik_VT1_seizo_no
,T1.kdkk_hnsik_VT2_seizo_no
,T1.kdkk_hnsik_yuko_kigen_ym
,T1.kdkk_hnsik_gknti_no
,T1.kdkk_mado_su
,T1.tdnstr_ts_dosa_kbn
,T1.tdnstr_tdnt_su
,T1.tdnstr_tdnstr_time
,T1.hkgds_kbn
,T1.hkgds_start_md
,T1.hkgds_end_md
,T1.no1_hkgds_tdn_hms
,T1.no1_hkgds_sydan_hms
,T1.no2_hkgds_tdn_hms
,T1.no2_hkgds_sydan_hms
,T1.no3_hkgds_tdn_hms
,T1.no3_hkgds_sydan_hms
,T1.no4_hkgds_tdn_hms
,T1.no4_hkgds_sydan_hms
,T1.no5_hkgds_tdn_hms
,T1.no5_hkgds_sydan_hms
,T1.no6_hkgds_tdn_hms
,T1.no6_hkgds_sydan_hms
,T1.no7_hkgds_tdn_hms
,T1.no7_hkgds_sydan_hms
,T1.no8_hkgds_tdn_hms
,T1.no8_hkgds_sydan_hms
,T1.no9_hkgds_tdn_hms
,T1.no9_hkgds_sydan_hms
,T1.no10_hkgds_tdn_hms
,T1.no10_hkgds_sydan_hms
,T1.kbttd_start_date
,T1.kbttd_end_date
,T1.hoka_kihi_kbn_cd
,T1.hoka_gmn_flckr_kbn
,T1.hoka_evnt_krk_kbn
,T1.hoka_hoka_hyz_kbn
,T1.hoka_hoka_hyz_value
,T1.hksgk_hksgn_kbn
,T1.hksgk_hka_dnr
,T1.hksgk_attny_time
,T1.hksgk_attny_kaisu
,T1.hksgk_attny_kaisu_clr_time
,T1.hksgr_hksgn_kbn
,T1.hksgr_hka_dnr
,T1.hksgr_attny_time
,T1.hksgr_attny_kaisu
,T1.hksgr_attny_kaisu_clr_time
,T1.knti_sbt_cd
,T1.trke_tis_kbn
,T1.hnsik_KYKBN_CD
,T1.surge_yksi_siyo_umu_flg
,T1.tidn_kbn
,T1.tidn_riyu_cd
,T1.sagyo_time_kbn_cd
,T1.sagyo_start_HMS
,T1.sagyo_end_HMS
,T1.dtkk_sougo_DRKEI_TEKYO_szsu
,T1.dtkk_sougo_DRKEI_TRTK_szsu
,T1.dtkk_max_JYDRRKEI_TEKYO_szsu
,T1.dtkk_max_JYDRRKEI_TRTK_szsu
,T1.dtkk_rsyk_TEKYO_szsu
,T1.dtkk_rsyk_TRTK_szsu
,T1.dtkk_rsmk_TEKYO_szsu
,T1.dtkk_rsmk_TRTK_szsu
,T1.dtkk_zyrt
,T1.tnsb_reuse_kbn
,T1.khkmk_keiki_kbn
,T1.khkmk_mg_kbn
,T1.khkmk_hnsik_kbn
,T1.khkmk_hnsik_st_iti_kbn
,T1.khkmk_wrms_kbn
,T1.khkmk_mg_umu_flg
,T1.khkmk_mg_yr
,T1.khkmk_kozi_kbn
,T1.khkmk_kozi_suryo
,T1.khkmk_trtk_kozih_kngk
,T1.khkmk_tekyo_kozih_kngk
,T1.kzkmk_hnsik_kozi_kbn
,T1.kzkmk_hnsik_kozi_suryo
,T1.kzkmk_hnsik_trtk_kozih_kngk
,T1.kzkmk_hnsik_tekyo_kozih_kngk
,T1.kzkmk_mg_kozi_kbn
,T1.kzkmk_mg_kozi_suryo
,T1.khkmk_mg_trtk_kozih_kngk
,T1.khkmk_mg_tekyo_kozih_kngk
,T1.tuika_kohi_trtk_kozih_kngk
,T1.tuika_kohi_tekyo_kozih_kngk
,T1.tuika_kohi_umu_flg
,T1.zippi_umu_flg
,T1.no1_zippi_no
,T1.no1_zippi_kbn
,T1.no1_zippi_kngk
,T1.no2_zippi_no
,T1.no2_zippi_kbn
,T1.no2_zippi_kngk
,T1.no3_zippi_no
,T1.no3_zippi_kbn
,T1.no3_zippi_kngk
,T1.no4_zippi_no
,T1.no4_zippi_kbn
,T1.no4_zippi_kngk
,T1.no5_zippi_no
,T1.no5_zippi_kbn
,T1.no5_zippi_kngk
,T1.kozih_gokei_kngk
,T1.ryssy_tenp_file_mng_no
,T1.tenp_file_mng_no
,T1.rrzk_naiyo
,T1.brtst_htiri_saksi_flg
,T1.kktkt_syun_list_saksi_flg
,T1.srcy_kiyk_no_renno
,T1.srcy_kiyk_no_edab_no
,T1.kesbt_bnrui_cd
,T3.kiyk_kbn_cd
,T3.keiki_kbn_cd
,T3.sm_zyri_kbn_cd
,T3.ftikk_flg
,SUBSTR(K1.kiyk_no,5,8) || ' -' || SUBSTR(K1.kiyk_no,13,1) AS kiyk_no_str 
 ,K1.kiyk_sbt_cd                 AS kiyk_sbt_cd 
 ,K1.kiyk_cs_mskn                AS kiyk_cs_mskn 
 ,K1.kiyk_cs_ms                  AS kiyk_cs_ms 
 ,dig14_cs_tel                   AS dig14_cs_tel_str 
 ,J1.tdhkn_ms                    AS cs_tdhkn_add_cd_ms 
 ,J1.siku_ms                     AS cs_siku_add_cd_ms 
 ,J1.oazat_ms                    AS cs_oazat_add_cd_ms 
 ,J1.azatm_ms                    AS cs_azatm_add_cd_ms 
 ,K1.cs_addhsk_naiyo             AS cs_addhsk_naiyo 
 ,K1.mkhy_senro_cd               AS mkhy_senro_cd 
 ,K1.mkhy_kansn_no               AS mkhy_kansn_no 
 ,K1.mkhy_bunk1_no               AS mkhy_bunk1_no 
 ,K1.mkhy_bunk2_no               AS mkhy_bunk2_no 
 ,K1.mkhy_bunk3_no               AS mkhy_bunk3_no 
 ,K1.cs_kiyk_dnryk               AS cs_kiyk_dnryk 
 ,K1.std_kensn_dd                AS std_kensn_dd 
 ,K1.kensn_yotei_togt_md         AS kensn_yotei_togt_md 
 ,K1.kensn_yotei_ykgt_md         AS kensn_yotei_ykgt_md 
 ,K1.kensn_yotei_yygt_md         AS kensn_yotei_yygt_md 
 ,K1.hist_flg                    AS hist_flg 
 ,K1.stbit_xzahyo                AS stbit_xzahyo 
 ,K1.stbit_yzahyo                AS stbit_yzahyo 
 ,K1.iewku_stdp_xzahyo           AS iewku_stdp_xzahyo 
 ,K1.iewku_stdp_yzahyo           AS iewku_stdp_yzahyo 
 ,' '                            AS ts_tdn_time             --Rev56 タイムスイッチ通電時間　項目削除への対応
 ,K1.zyuky_ukky_kbn              AS zyuky_ukky_kbn
 ,M19.prm_cd_abms                AS prcmg_cd_abms 
 ,M2.prm_cd_abms                 AS brot_sett_flg_abms 
 ,M3.prm_cd_abms                 AS tkkk_kesbt_kbn_abms 
 ,M4.prm_cd_abms                 AS tkkk_krhsk_cd_abms 
 ,M5.prm_cd_abms                 AS tkkk_tshsk_cd_abms 
 ,M6.prm_cd_abms                 AS tkkk_knthk_cd_abms 
 ,M7.prm_cd_abms                 AS tkkk_ts_kino_umu_flg_abms 
 ,M11.prm_cd_abms                AS ttkk_kesbt_kbn_abms 
 ,M12.prm_cd_abms                AS ttkk_krhsk_cd_abms 
 ,M14.prm_cd_abms                AS ttkk_tshsk_cd_abms 
 ,M15.prm_cd_abms                AS ttkk_ts_kino_umu_flg_abms 
 ,M8.prm_cd_abms                 AS ttkk_khk_kino_umu_flg_abms 
 ,M9.prm_cd_abms                 AS ttkk_gaibu_output_umu_flg_ms 
 ,M10.prm_cd_abms                AS ttkk_taiko_kbn_abms 
 ,M1.prm_cd_ms                   AS knti_sbt_cd_ms
 ,M20.prm_cd_abms                AS nktr_kensa_zyokyo_abms
 ,M21.prm_cd_abms                AS trke_sbt_abms
 ,M22.prm_cd_abms                AS tkkk_khk_kino_umu_flg_abms
 ,M23.prm_cd_abms                AS tkkk_gaibu_output_umu_flg_abms
 ,M24.prm_cd_abms                AS tkkk_taiko_kbn_abms
 ,M25.prm_cd_abms                AS ttkk_knthk_abms
 ,M26.prm_cd_abms                AS surge_yksi_siyo_umu_flg_abms
 ,M27.prm_cd_abms                AS tidn_kbn_abms
 ,M28.prm_cd_abms                AS tnsb_reuse_kbn_abms
 ,M29.prm_cd_abms                AS khkmk_keiki_kbn_abms
 ,M30.prm_cd_abms                AS khkmk_mg_kbn_abms
 ,M31.prm_cd_abms                AS khkmk_hnsik_kbn_abms
 ,M32.prm_cd_abms                AS khkmk_wrms_kbn_abms
 ,M33.prm_cd_abms                AS khkmk_mg_umu_flg_abms
 ,M34.prm_cd_abms                AS no1_zippi_kbn_abms
 ,M35.prm_cd_abms                AS no2_zippi_kbn_abms
 ,M36.prm_cd_abms                AS no3_zippi_kbn_abms
 ,M37.prm_cd_abms                AS no4_zippi_kbn_abms
 ,M38.prm_cd_abms                AS no5_zippi_kbn_abms
 ,M39.prm_cd_abms                AS trke_tis_abms
 ,J2.mdgt_ms                     AS mdgt_ms
 ,J3.kozitn_ms                   AS kozitn_ms
 ,J4.senro_ms                    AS senro_ms
 ,J5.str2_zyousu_value           AS tykei_mdgt_ms
 ,J6.str2_zyousu_value           AS tykei_kozitn_ms
 ,' '                            AS yukoWhZytrTime 
 ,' '                            AS yukoWhGytrTime
 ,T1.nktrk_sizsk_kbn             AS nktrk_sizsk_kbn
 ,'2'                            AS phase_kbn
 ,T1.tkkk_tnsb_seizo_ym          AS tkkk_tnsb_seizo_ym
 ,T1.tkkk_tnsb_sbt_cd            AS tkkk_tnsb_sbt_cd
 ,T1.tkkk_tnsb_ssnsk_dnat_cd     AS tkkk_tnsb_ssnsk_dnat_cd
 ,T1.tkkk_tnsb_yr                AS tkkk_tnsb_yr
 ,T1.tkkk_tnsb_seizo_no          AS tkkk_tnsb_seizo_no
 ,T1.tkkk_tnsb_ktsk_ms           AS tkkk_tnsb_ktsk_ms
 ,T1.tkkk_tnsb_kzskbt_kbn        AS tkkk_tnsb_kzskbt_kbn
 ,T1.tkkk_tnsb_szsya_mng_value   AS tkkk_tnsb_szsya_mng_value
 ,T1.ttkk_tnsb_seizo_ym          AS ttkk_tnsb_seizo_ym
 ,T1.ttkk_tnsb_sbt_cd            AS ttkk_tnsb_sbt_cd
 ,T1.ttkk_tnsb_ssnsk_dnat_cd     AS ttkk_tnsb_ssnsk_dnat_cd
 ,T1.ttkk_tnsb_yr                AS ttkk_tnsb_yr
 ,T1.ttkk_tnsb_seizo_no          AS ttkk_tnsb_seizo_no
 ,T1.ttkk_tnsb_ktsk_ms           AS ttkk_tnsb_ktsk_ms
 ,T1.ttkk_tnsb_kzskbt_kbn        AS ttkk_tnsb_kzskbt_kbn
 ,T1.ttkk_tnsb_szsya_mng_value   AS ttkk_tnsb_szsya_mng_value
 ,T1.kdkk_tnsb_seizo_ym          AS kdkk_tnsb_seizo_ym
 ,T1.kdkk_tnsb_sbt_cd            AS kdkk_tnsb_sbt_cd
 ,T1.kdkk_tnsb_ssnsk_dnat_cd     AS kdkk_tnsb_ssnsk_dnat_cd
 ,T1.kdkk_tnsb_yr                AS kdkk_tnsb_yr
 ,T1.kdkk_tnsb_seizo_no          AS kdkk_tnsb_seizo_no
 ,T1.kdkk_tnsb_ktsk_ms           AS kdkk_tnsb_ktsk_ms
 ,T1.kdkk_tnsb_kzskbt_kbn        AS kdkk_tnsb_kzskbt_kbn
 ,T1.kdkk_tnsb_szsya_mng_value   AS kdkk_tnsb_szsya_mng_value
 ,K1.stzk_sdnsv_menu_cd          AS stzk_sdnsv_menu_cd
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（高圧他）
  INNER JOIN trke_trkehy_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--契約_お客さま情報（高圧他）
  INNER JOIN kiyk_customer_info_kahk K1 
ON T1.dig4_zgsyo_cd = K1.dig4_zgsyo_cd 
   AND T1.trkhy_hakko_nendo = K1.hakko_nendo 
   AND T1.kyokyu_spot_tokti_no = K1.kyokyu_spot_tokti_no 
   AND T1.trkhy_kbn = K1.trkhy_kbn 
--   AND K1.hist_flg = '0'   Rev072 DEL 20250404 持出ファイル作成エラー対応 
--マスタ_町字
  LEFT JOIN mst_tyaza J1  
ON J1.tdhkn_add_cd = T1.cs_tdhkn_add_cd 
   AND J1.siku_add_cd = T1.cs_siku_add_cd  
   AND J1.oazat_add_cd = T1.cs_oazat_add_cd  
   AND J1.azatm_add_cd = T1.cs_azatm_add_cd 
--Rev084削除--計器_地点_特記事項（高圧他）
--Rev084削除  LEFT JOIN keiki_spot_tkzk_kahk T2  
--Rev084削除ON T2.kyokyu_spot_tokti_no = T1.kyokyu_spot_tokti_no 
--計器_計器情報（高圧他）_管理
  INNER JOIN  keiki_keiki_info_kahk_mng T3
ON T3.kkinf_id = T1.tkkk_kkinf_id 
   AND (T3.kiyk_kbn_cd = '2' OR T3.kiyk_kbn_cd = '3')
   AND (T3.keiki_kbn_cd = '1' OR T3.keiki_kbn_cd = '2')
   AND T3.sm_zyri_kbn_cd = '1'
--取替対象名称
  INNER JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'KNTI_SBT_CD' 
   AND M1.prm_cd = T1.knti_sbt_cd 
--Ｂルート設定 
  INNER JOIN  com_paramater_cd_mst M2 
ON M2.prm_id = 'BROOT_SETT_FLG' 
   AND M2.prm_cd = T1.brot_sett_flg 
--撤去計器_計器種別区分
  INNER JOIN  com_paramater_cd_mst M3 
ON M3.prm_id = 'KESBT_KBN' 
   AND M3.prm_cd = T1.tkkk_kesbt_kbn 
--撤去計器_計量方式コード
  LEFT JOIN  com_paramater_cd_mst M4 
ON M4.prm_id = 'KRHSK_CD' 
   AND M4.prm_cd = T1.tkkk_krhsk_cd 
--撤去計器_通信方式コード
  LEFT JOIN  com_paramater_cd_mst M5 
   ON M5.prm_id = 'TSHSK_CD' 
  AND M5.prm_cd = T1.tkkk_tshsk_cd 
--撤去計器_検定方向コード
  INNER JOIN  com_paramater_cd_mst M6 
ON M6.prm_id = 'KNTHK_CD' 
   AND M6.prm_cd = T1.tkkk_knthk_cd 
--撤去計器_ＴＳ機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M7 
ON M7.prm_id = 'TS_KINOUMU_FLG' 
   AND M7.prm_cd = T1.tkkk_ts_kino_umu_flg 
--撤去計器_開閉器機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M8 
ON M8.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M8.prm_cd = T1.tkkk_khk_kino_umu_flg 
--撤去計器_外部出力有無フラグ
  INNER JOIN  com_paramater_cd_mst M9 
ON M9.prm_id = 'OUTPUTUMU_FLG' 
   AND M9.prm_cd = T1.tkkk_gaibu_output_umu_flg 
--撤去計器_計器_耐候区分
  INNER JOIN  com_paramater_cd_mst M10 
ON M10.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M10.prm_cd = T1.tkkk_taiko_kbn 
--取付計器_計器種別コード
  LEFT JOIN  com_paramater_cd_mst M11 
ON M11.prm_id = 'KESBT_KBN' 
   AND M11.prm_cd = T1.ttkk_kesbt_kbn 
--取付計器_計量方式コード
  LEFT JOIN  com_paramater_cd_mst M12 
ON M12.prm_id = 'KRHSK_CD' 
   AND M12.prm_cd = T1.ttkk_krhsk_cd 
--取付計器_通信方式コード
  LEFT JOIN  com_paramater_cd_mst M13 
ON M13.prm_id = 'TSHSK_CD' 
   AND M13.prm_cd = T1.ttkk_tshsk_cd 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M14 
ON M14.prm_id = 'KNTHK_CD' 
   AND M14.prm_cd = T1.ttkk_knthk_cd 
--取付計器_ＴＳ機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M15 
ON M15.prm_id = 'TS_KINOUMU_FLG' 
   AND M15.prm_cd = T1.ttkk_ts_kino_umu_flg 
--取付計器_開閉器機能有無フラグ
  LEFT JOIN  com_paramater_cd_mst M16 
ON M16.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M16.prm_cd = T1.ttkk_khk_kino_umu_flg 
--取付計器_外部出力有無フラグ
  LEFT JOIN  com_paramater_cd_mst M17 
ON M17.prm_id = 'OUTPUTUMU_FLG' 
   AND M17.prm_cd = T1.ttkk_gaibu_output_umu_flg 
--取付計器_計器_耐候区分
  LEFT JOIN  com_paramater_cd_mst M18 
ON M18.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M18.prm_cd = T1.ttkk_taiko_kbn 
--行程管理コード
  INNER JOIN  com_paramater_cd_mst M19 
ON M19.prm_id = 'PROC_MNG_CD' 
   AND M19.prm_cd = T1.prcmg_cd 
--抜取検査状況取替対象
  LEFT JOIN  com_paramater_cd_mst M20 
ON M20.prm_id = 'NKTRKENSA_ZYOKYO_CD' 
   AND M20.prm_cd = T1.nktr_kensa_zyokyo_cd 
--取替種別
  INNER JOIN  com_paramater_cd_mst M21 
ON M21.prm_id = 'TRKE_SBT_CD' 
   AND M21.prm_cd = T1.trke_sbt_cd 
--開閉器機能有無フラグ
  INNER JOIN  com_paramater_cd_mst M22 
ON M22.prm_id = 'KEIKI_KHKKINOUMU_FLG' 
   AND M22.prm_cd = T1.tkkk_khk_kino_umu_flg 
--外部出力有無フラグ
  INNER JOIN  com_paramater_cd_mst M23 
ON M23.prm_id = 'OUTPUTUMU_FLG' 
   AND M23.prm_cd = T1.tkkk_gaibu_output_umu_flg 
--耐候区分
  INNER JOIN  com_paramater_cd_mst M24 
ON M24.prm_id = 'KEIKI_TAIKO_KBN' 
   AND M24.prm_cd = T1.tkkk_taiko_kbn 
--取付計器_検定方向コード
  LEFT JOIN  com_paramater_cd_mst M25 
ON M25.prm_id = 'KNTHK_CD' 
   AND M25.prm_cd = T1.ttkk_knthk_cd 
--サージ抑制使用有無フラグ
  LEFT JOIN  com_paramater_cd_mst M26 
ON M26.prm_id = 'SURG_YKSISIYOUMU_FLG' 
   AND M26.prm_cd = T1.surge_yksi_siyo_umu_flg 
--停電区分
  LEFT JOIN  com_paramater_cd_mst M27 
ON M27.prm_id = 'TIDN_KBN' 
   AND M27.prm_cd = T1.tidn_kbn 
--端子部再用区分
  LEFT JOIN  com_paramater_cd_mst M28 
ON M28.prm_id = 'TNSB_REUSE_KBN' 
   AND M28.prm_cd = T1.tnsb_reuse_kbn 
--工費項目_計器区分
  INNER JOIN  com_paramater_cd_mst M29 
ON M29.prm_id = 'KOHIKOMK_KEIKI_KBN' 
   AND M29.prm_cd = T1.khkmk_keiki_kbn 
--工費項目_ＭＧ
  LEFT JOIN  com_paramater_cd_mst M30 
ON M30.prm_id = 'KOHIKOMK_MG_KBN' 
   AND M30.prm_cd = T1.khkmk_mg_kbn 
--工費項目_変成器区分
  LEFT JOIN  com_paramater_cd_mst M31 
ON M31.prm_id = 'KOHIKOMK_HNSK_KBN' 
   AND M31.prm_cd = T1.khkmk_hnsik_kbn 
--工費項目_割増区分
  LEFT JOIN  com_paramater_cd_mst M32 
ON M32.prm_id = 'KOHIKOMK_WRMS_CD' 
   AND M32.prm_cd = T1.khkmk_wrms_kbn 
--工費項目_ＭＧ有無フラグ
  INNER JOIN  com_paramater_cd_mst M33 
ON M33.prm_id = 'KOHIKOMK_MGUMU_FLG' 
   AND M33.prm_cd = T1.khkmk_mg_umu_flg 
--実費_区分１
  LEFT JOIN  com_paramater_cd_mst M34 
ON M34.prm_id = 'ZIPPI_KBN' 
   AND M34.prm_cd = T1.no1_zippi_kbn 
--実費_区分２
  LEFT JOIN  com_paramater_cd_mst M35 
ON M35.prm_id = 'ZIPPI_KBN' 
   AND M35.prm_cd = T1.no2_zippi_kbn 
--実費_区分３
  LEFT JOIN  com_paramater_cd_mst M36 
ON M36.prm_id = 'ZIPPI_KBN' 
   AND M36.prm_cd = T1.no3_zippi_kbn 
--実費_区分４
  LEFT JOIN  com_paramater_cd_mst M37 
ON M37.prm_id = 'ZIPPI_KBN' 
   AND M37.prm_cd = T1.no4_zippi_kbn 
--実費_区分５
  LEFT JOIN  com_paramater_cd_mst M38 
ON M38.prm_id = 'ZIPPI_KBN' 
   AND M38.prm_cd = T1.no5_zippi_kbn 
--取替対象区分
  LEFT JOIN  com_paramater_cd_mst M39 
ON M39.prm_id = 'TRKE_TIS_KBN' 
   AND M39.prm_cd = T1.trke_tis_kbn

--工事会社窓口名
 LEFT JOIN mst_mdgt J2
    ON J2.mdgt_cd = T1.kzkis_mdgt_cd
   AND J2.latest_flg = '1'
   AND J2.haisi_flg = '0'
 --工事店名
 LEFT JOIN mst_kozitn J3
    ON J3.kozitn_cd = T1.kozitn_cd
   AND J3.latest_flg = '1'
   AND J3.haisi_flg = '0'
 --線路名称
 LEFT JOIN mst_senro J4
    ON J4.senro_cd = K1.mkhy_senro_cd
--直営　窓口名
 LEFT JOIN com_zyousu_mst J5
    ON J5.zyousu_cd = {1} and {0} BETWEEN J5.yuko_start_ymd AND J5.yuko_end_ymd
    --直営　工事店名
 LEFT JOIN com_zyousu_mst J6
    ON J6.zyousu_cd = {2} and {0} BETWEEN J6.yuko_start_ymd AND J6.yuko_end_ymd
 WHERE 1 = 1 
  AND D1.downld_renno={3}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.ZyousuCd.CHOKUEI_MDGT_INFO) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))

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


#Region "(SQL)ダウンロード管理詳細・取替_取替票行程（高圧他）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票行程（高圧他）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票行程（高圧他）を検索します</remarks>
    Friend Function S017(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
,T1.saksi_date
,T1.upd_date
,T1.sousa_user_id
,T1.sousa_appli_cd
,T1.dig4_zgsyo_cd
,T1.trkhy_hakko_nendo
,T1.trkhy_kbn
,T1.trkhy_no
,T1.kkst_tanto_kaih_ymd
,T1.kkst_tanto_kaih_syors_cd
,T1.kkst_tanto_kaih_syors_ms
,T1.kkst_tanto_ka_cd
,T1.kkst_tanto_ka_ms
,T1.kkst_ht_snn_ymd
,T1.kkst_ht_snsya_cd
,T1.kkst_ht_snsya_ms
,T1.kkst_ht_ka_cd
,T1.kkst_ht_ka_ms
,T1.kkst_hnn_flg
,T1.KZHTNKN_tanto_kaih_ymd
,T1.KZHTNKN_tanto_kaih_syors_cd
,T1.KZHTNKN_tanto_kaih_syors_ms
,T1.KZHTNKN_tanto_ka_cd
,T1.KZHTNKN_tanto_ka_ms
,T1.KZHTNKN_ht_snn_ymd
,T1.KZHTNKN_ht_snsya_cd
,T1.KZHTNKN_ht_snsya_ms
,T1.KZHTNKN_ht_ka_cd
,T1.KZHTNKN_ht_ka_ms
,T1.KZHTNKN_kt_snn_ymd
,T1.KZHTNKN_kt_snsya_cd
,T1.KZHTNKN_kt_snsya_ms
,T1.KZHTNKN_kt_ka_cd
,T1.KZHTNKN_kt_ka_ms
,T1.KZHTNKN_hnn_flg
,T1.tkpln_tanto_khiri_ymd
,T1.tkpln_tanto_khiri_syors_cd
,T1.tkpln_tanto_khiri_syors_ms
,T1.tkpln_tanto_ka_cd
,T1.tkpln_tanto_ka_ms
,T1.tkpln_tanto_kzkis_mdgt_cd
,T1.tkpln_tanto_kzkis_mdgt_ms
,T1.tkpln_tanto_kozitn_cd
,T1.tkpln_tanto_kozitn_ms
,T1.tkpln_tanto_khiri_hnn_flg
,T1.khsz_tanto_kaih_ymd
,T1.khsz_tanto_kaih_syors_cd
,T1.khsz_tanto_kaih_syors_ms
,T1.khsz_tanto_ka_cd
,T1.khsz_tanto_ka_ms
,T1.khsz_hnn_flg
,T1.seko_kti_ht_snn_ymd
,T1.seko_kti_ht_snsya_cd
,T1.seko_kti_ht_snsya_ms
,T1.seko_kti_ht_ka_cd
,T1.seko_kti_ht_ka_ms
,T1.seko_kti_kt_snn_ymd
,T1.seko_kti_kt_snsya_cd
,T1.seko_kti_kt_snsya_ms
,T1.seko_kti_kt_ka_cd
,T1.seko_kti_kt_ka_ms
,T1.seko_kti_hnn_flg
,T1.skohu_tanto_kaih_syori_ymd
,T1.skohu_tanto_kaih_syors_cd
,T1.skohu_tanto_kaih_syors_ms
,T1.skohu_tanto_ka_cd
,T1.skohu_tanto_ka_ms
,T1.skohu_ht_snn_ymd
,T1.skohu_ht_snsya_cd
,T1.skohu_ht_snsya_ms
,T1.skohu_ht_ka_cd
,T1.skohu_ht_ka_ms
,T1.skohu_kt_snn_ymd
,T1.skohu_kt_snsya_cd
,T1.skohu_kt_snsya_ms
,T1.skohu_kt_ka_cd
,T1.skohu_kt_ka_ms
,T1.skohu_hnn_flg
,T1.zkskh_tanto_kaih_syori_ymd
,T1.zkskh_tanto_kaih_syors_cd
,T1.zkskh_tanto_kaih_syors_ms
,T1.zkskh_tanto_ka_cd
,T1.zkskh_tanto_ka_ms
,T1.zkskh_ht_snn_ymd
,T1.zkskh_ht_snsya_cd
,T1.zkskh_ht_snsya_ms
,T1.zkskh_ht_ka_cd
,T1.zkskh_ht_ka_ms
,T1.zkskh_kt_snn_ymd
,T1.zkskh_kt_snsya_cd
,T1.zkskh_kt_snsya_ms
,T1.zkskh_kt_ka_cd
,T1.zkskh_kt_ka_ms
,T1.zkskh_hnn_flg
,T1.skyti_ktst_syori_ymd
,T1.skyti_ktst_syors_cd
,T1.skyti_ktst_syors_ms
,T1.skyti_ktst_ka_cd
,T1.skyti_ktst_ka_ms
,T1.skyti_ktst_kzkis_mdgt_cd
,T1.skyti_ktst_kzkis_mdgt_ms
,T1.skyti_ktst_kozitn_cd
,T1.skyti_ktst_kozitn_ms
,T1.skyti_skyti_inpt_syori_ymd
,T1.skyti_skyti_inpt_syors_cd
,T1.skyti_skyti_inpt_syors_ms
,T1.skyti_skyti_inpt_ka_cd
,T1.skyti_skyti_inpt_ka_ms
,T1.skyti_skyti_inpt_kkmg_cd
,T1.skyti_skyti_inpt_kkmg_ms
,T1.skyti_skyti_inpt_kozitn_cd
,T1.skyti_skyti_inpt_kozitn_ms
,T1.skyti_skyti_snn_syori_ymd
,T1.skyti_skyti_snn_syors_cd
,T1.skyti_skyti_snn_syors_ms
,T1.skyti_skyti_snn_ka_cd
,T1.skyti_skyti_snn_ka_ms
,T1.skyti_skyti_snn_kkmg_cd
,T1.skyti_skyti_snn_kkmg_ms
,T1.skyti_skyti_snn_kozitn_cd
,T1.skyti_skyti_snn_kozitn_ms
,T1.seko_tanto_syori_ymd
,T1.seko_tanto_syors_cd
,T1.seko_tanto_syors_ms
,T1.seko_tanto_ka_cd
,T1.seko_tanto_ka_ms
,T1.seko_tanto_kzkis_mdgt_cd
,T1.seko_tanto_kzkis_mdgt_ms
,T1.seko_tanto_kozitn_cd
,T1.seko_tanto_kozitn_ms
,T1.seko_elder_tnknk_kknn_flg
,T1.seko_elder_tnkn_kknn_ymd
,T1.seko_elder_tnkn_knsya_cd
,T1.seko_elder_tnkn_knsya_ms
,T1.seko_elder_ka_cd
,T1.seko_elder_ka_ms
,T1.seko_elder_kzkis_mdgt_cd
,T1.seko_elder_kzkis_mdgt_ms
,T1.seko_elder_kozitn_cd
,T1.seko_elder_kozitn_ms
,T1.skrepo_syori_ymd
,T1.skrepo_syors_cd
,T1.skrepo_syors_ms
,T1.skrepo_ka_cd
,T1.skrepo_ka_ms
,T1.skrepo_kzkis_mdgt_cd
,T1.skrepo_kzkis_mdgt_ms
,T1.skrepo_kozitn_cd
,T1.skrepo_kozitn_ms
,T1.no1_skrepo_zippi_kknn_flg
,T1.no2_skrepo_zippi_kknn_flg
,T1.no3_skrepo_zippi_kknn_flg
,T1.skkns_kiktt_kaih_ymd
,T1.skkns_kiktt_kaih_syors_cd
,T1.skkns_kiktt_kaih_syors_ms
,T1.skkns_kiktt_ka_cd
,T1.skkns_kiktt_ka_ms
,T1.skkns_kiktt_tnknk_kknn_flg
,T1.no1_skkns_kiktt_zpkkn_flg
,T1.no2_skkns_kiktt_zpkkn_flg
,T1.no3_skkns_kiktt_zpkkn_flg
,T1.skkns_kikht_snn_ymd
,T1.skkns_kikht_snsya_cd
,T1.skkns_kikht_snsya_ms
,T1.skkns_kikht_ka_cd
,T1.skkns_kikht_ka_ms
,T1.skkns_kikht_tnknk_kknn_flg
,T1.no1_skkns_kikht_zpkkn_flg
,T1.no2_skkns_kikht_zpkkn_flg
,T1.no3_skkns_kikht_zpkkn_flg
,T1.skkns_kikkt_snn_ymd
,T1.skkns_kikkt_snsya_cd
,T1.skkns_kikkt_snsya_ms
,T1.skkns_kikkt_ka_cd
,T1.skkns_kikkt_ka_ms
,T1.skkns_kikkt_tnknk_kknn_flg
,T1.no1_skkns_kikkt_zpkkn_flg
,T1.no2_skkns_kikkt_zpkkn_flg
,T1.no3_skkns_kikkt_zpkkn_flg
,T1.skkns_knstt_kaih_ymd
,T1.skkns_knstt_kaih_syors_cd
,T1.skkns_knstt_kaih_syors_ms
,T1.skkns_knstt_ka_cd
,T1.skkns_knstt_ka_ms
,T1.skkns_knstt_tnknk_kknn_flg
,T1.no1_skkns_knstt_zpkkn_flg
,T1.no2_skkns_knstt_zpkkn_flg
,T1.no3_skkns_knstt_zpkkn_flg
,T1.skkns_knsht_snn_ymd
,T1.skkns_knsht_snsya_cd
,T1.skkns_knsht_snsya_ms
,T1.skkns_knsht_ka_cd
,T1.skkns_knsht_ka_ms
,T1.skkns_knsht_tnknk_kknn_flg
,T1.no1_skkns_knsht_zpkkn_flg
,T1.no2_skkns_knsht_zpkkn_flg
,T1.no3_skkns_knsht_zpkkn_flg
,T1.skkns_knskt_snn_ymd
,T1.skkns_knskt_snsya_cd
,T1.skkns_knskt_snsya_ms
,T1.skkns_knskt_ka_cd
,T1.skkns_knskt_ka_ms
,T1.skkns_knskt_tnknk_kknn_flg
,T1.no1_skkns_knskt_zpkkn_flg
,T1.no2_skkns_knskt_zpkkn_flg
,T1.no3_skkns_knskt_zpkkn_flg
,T1.skkns_hnn_flg
,T1.skkns_hnn_riyu_naiyo
,T1.krds_sett_tanto_syori_ymd
,T1.krds_sett_tanto_syors_cd
,T1.krds_sett_tanto_syors_ms
,T1.krds_sett_tanto_ka_cd
,T1.krds_sett_tanto_ka_ms
,T1.krds_sett_ht_snn_ymd
,T1.krds_sett_ht_snsya_cd
,T1.krds_sett_ht_snsya_ms
,T1.krds_sett_ht_ka_cd
,T1.krds_sett_ht_ka_ms
,T1.krds_sett_hnn_flg
,T1.nktrk_sizi_syori_ymd
,T1.nktrk_sizi_syors_cd
,T1.nktrk_sizi_syors_ms
,T1.nktrk_sizi_ka_cd
,T1.nktrk_sizi_ka_ms
,T1.nktrk_sizi_kzkis_mdgt_cd
,T1.nktrk_sizi_kzkis_mdgt_ms
,T1.nktrk_sizi_kozitn_cd
,T1.nktrk_sizi_kozitn_ms
,T1.nktrk_tanto_syori_ymd
,T1.nktrk_tanto_syors_cd
,T1.nktrk_tanto_syors_ms
,T1.nktrk_tanto_ka_cd
,T1.nktrk_tanto_ka_ms
,T1.nktrk_tanto_ces_cd
,T1.nktrk_tanto_ces_ms
,T1.nktrk_elder_syori_ymd
,T1.nktrk_elder_syors_cd
,T1.nktrk_elder_syors_ms
,T1.nktrk_elder_ka_cd
,T1.nktrk_elder_ka_ms
,T1.nktrk_elder_ces_cd
,T1.nktrk_elder_ces_ms
,T1.nktrk_hnn_flg
,T1.tidu_ccsz_syori_ymd
,T1.tidu_ccsz_syors_cd
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票行程（高圧他）
  INNER JOIN trke_trkehy_prc_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理詳細・計器_地点_特記事項（高圧他）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・計器_地点_特記事項（高圧他）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・計器_地点_特記事項（高圧他）を検索します</remarks>
    Friend Function S018(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.kyokyu_spot_tokti_no
 ,K1.tkzk_renno
 ,K1.tkzk_bnrui_cd
 ,K1.tkzk_kizi_naiyo
 ,K1.itaks_kaizi_kizi_naiyo
 ,K1.kiyk_cs_ms
 ,K1.tenp_file_mng_no
 ,K1.TOROK_YMD
 ,K1.UPD_YMD
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（高圧他）
  INNER JOIN trke_trkehy_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--計器_地点_特記事項
  INNER JOIN keiki_spot_tkzk_kahk K1 
ON T1.kyokyu_spot_tokti_no = K1.kyokyu_spot_tokti_no 
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

#Region "(SQL)ダウンロード管理詳細・取替_取替票追加工費（高圧他）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_取替票追加工費（高圧他）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_取替票追加工費（高圧他）を検索します</remarks>
    Friend Function S019(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,T1.saksi_date
 ,T1.upd_date
 ,T1.sousa_user_id
 ,T1.sousa_appli_cd
 ,T1.dig4_zgsyo_cd
 ,T1.trkhy_hakko_nendo
 ,T1.trkhy_kbn
 ,T1.trkhy_no
 ,T1.tuika_kohi_row_no
 ,T1.tuika_kohi_kozi_kbn
 ,T1.tuika_kohi_trtk_suryo
 ,T1.tuika_kohi_tekyo_suryo
 ,T1.trtk_kozih_kngk
 ,T1.tekyo_kozih_kngk
 ,T1.tuika_kohi_riyu_naiyo
 ,case M1.tuikakh_sytk_mst_kbn
    when '1' then M2.koryo_ms
    when '2' then M3.zunit_ms
    else ' ' 
  END As tuika_kohi_kozi_kbn_ms
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票追加工費（高圧他）
  INNER JOIN trke_trkehy_tuikakh_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--マスター利用可能追加工費
  LEFT JOIN mst_use_kano_tuikakh M1
ON T1.tuika_kohi_kozi_kbn = SUBSTRING(M1.koryo_zunit_cd, 4,4)
--マスター工量
  LEFT  JOIN  mst_koryo M2 
ON M2.koryo_cd=M1.koryo_zunit_cd
and M2.tky_start_ymd <= {1} and {1} <= M2.tky_end_ymd
--マスター材料ユニット
  LEFT  JOIN  mst_zunit M3 
ON M3.zunit_cd=M1.koryo_zunit_cd
and M3.tky_start_ymd <= {1} and {1} <= M3.tky_end_ymd
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno) _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD))
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

#Region "(SQL)ダウンロード管理詳細・取替_自主点検チェックシート（高圧他）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・取替_自主点検チェックシート（高圧他）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・取替_自主点検チェックシート（高圧他）を検索します</remarks>
    Friend Function S020(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
            'Rev056 2025/07/01 託送情報切り替え対応（高圧他） ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_tio_cmp_ymd
            sb.Append(<sql><![CDATA[
select
  D1.dig4_zgsyo_cd AS D_dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS D_trkhy_hakko_nendo
 ,D1.trkhy_kbn AS D_trkhy_kbn
 ,D1.trkhy_no As D_trkhy_no
 ,T1.saksi_date
 ,T1.upd_date
 ,T1.sousa_user_id
 ,T1.sousa_appli_cd
 ,T1.dig4_zgsyo_cd
 ,T1.trkhy_hakko_nendo
 ,T1.trkhy_kbn
 ,T1.trkhy_no
 ,T1.skosya_tnkn_ymd
 ,T1.kstts_tnkn_ymd
 ,T1.kkchyr_snst_skosya_ryohi_kbn
 ,T1.kkchyr_snst_ntktt_ryohi_kbn
 ,T1.kkchyr_snst_kstts_ryohi_kbn
 ,T1.kkchyr_snst_tio_naiyo
 ,T1.kkchyr_snst_tio_cmp_ymd
 ,T1.kkchyr_kst_skosya_ryohi_kbn
 ,T1.kkchyr_kst_ntktt_ryohi_kbn
 ,T1.kkchyr_kst_kstts_ryohi_kbn
 ,T1.kkchyr_kst_tio_naiyo
 ,T1.kkchyr_kst_tio_cmp_ymd
 ,T1.kkchano_snst_skosya_ryohi_kbn
 ,T1.kkchano_snst_ntktt_ryohi_kbn
 ,T1.kkchano_snst_kstts_ryohi_kbn
 ,T1.kkchano_snst_tio_naiyo
 ,T1.kkchano_snst_tio_cmp_ymd
 ,T1.kkchano_kst_skosya_ryohi_kbn
 ,T1.kkchano_kst_ntktt_ryohi_kbn
 ,T1.kkchano_kst_kstts_ryohi_kbn
 ,T1.kkchano_kst_tio_naiyo
 ,T1.kkchano_kst_tio_cmp_ymd
 ,T1.kkchsno_skosya_ryohi_kbn
 ,T1.kkchsno_ntktt_ryohi_kbn
 ,T1.kkchsno_kstts_ryohi_kbn
 ,T1.kkchsno_tio_naiyo
 ,T1.kkchsno_tio_cmp_ymd
 ,T1.kkchkkl_skosya_ryohi_kbn
 ,T1.kkchkkl_ntktt_ryohi_kbn
 ,T1.kkchkkl_kstts_ryohi_kbn
 ,T1.kkchkkl_tio_naiyo
 ,T1.kkchkkl_tio_cmp_ymd
 ,T1.kkchkss_skosya_ryohi_kbn
 ,T1.kkchkss_ntktt_ryohi_kbn
 ,T1.kkchkss_kstts_ryohi_kbn
 ,T1.kkchkss_tio_naiyo
 ,T1.kkchkss_tio_cmp_ymd
 ,T1.kkchkib_skosya_ryohi_kbn
 ,T1.kkchkib_ntktt_ryohi_kbn
 ,T1.kkchkib_kstts_ryohi_kbn
 ,T1.kkchkib_tio_naiyo
 ,T1.kkchkib_tio_cmp_ymd
 ,T1.kkchkhd_skosya_ryohi_kbn
 ,T1.kkchkhd_ntktt_ryohi_kbn
 ,T1.kkchkhd_kstts_ryohi_kbn
 ,T1.kkchkhd_tio_naiyo
 ,T1.kkchkhd_tio_cmp_ymd
 ,T1.kkchkbs_skosya_ryohi_kbn
 ,T1.kkchkbs_ntktt_ryohi_kbn
 ,T1.kkchkbs_kstts_ryohi_kbn
 ,T1.kkchkbs_tio_naiyo
 ,T1.kkchkbs_tio_cmp_ymd
 ,T1.kkckzr_snst_skosya_ryohi_kbn
 ,T1.kkckzr_snst_ntktt_ryohi_kbn
 ,T1.kkckzr_snst_kstts_ryohi_kbn
 ,T1.kkckzr_snst_tio_naiyo
 ,T1.kkckzr_snst_tio_cmp_ymd
 ,T1.kkckzr_kst_skosya_ryohi_kbn
 ,T1.kkckzr_kst_ntktt_ryohi_kbn
 ,T1.kkckzr_kst_kstts_ryohi_kbn
 ,T1.kkckzr_kst_tio_naiyo
 ,T1.kkckzr_kst_tio_cmp_ymd
 ,T1.kkckano_snst_skosya_ryohi_kbn
 ,T1.kkckano_snst_ntktt_ryohi_kbn
 ,T1.kkckano_snst_kstts_ryohi_kbn
 ,T1.kkckano_snst_tio_naiyo
 ,T1.kkckano_snst_tio_cmp_ymd
 ,T1.kkckano_kst_skosya_ryohi_kbn
 ,T1.kkckano_kst_ntktt_ryohi_kbn
 ,T1.kkckano_kst_kstts_ryohi_kbn
 ,T1.kkckano_kst_tio_naiyo
 ,T1.kkckano_kst_tio_cmp_ymd
 ,T1.kkcksno_skosya_ryohi_kbn
 ,T1.kkcksno_ntktt_ryohi_kbn
 ,T1.kkcksno_kstts_ryohi_kbn
 ,T1.kkcksno_tio_naiyo
 ,T1.kkcksno_tio_cmp_ymd
 ,T1.kkckdj_dosa_skosya_ryohi_kbn
 ,T1.kkckdj_dosa_ntktt_ryohi_kbn
 ,T1.kkckdj_dosa_kstts_ryohi_kbn
 ,T1.kkckdj_dosa_tio_naiyo
 ,T1.kkckdj_dosa_tio_cmp_ymd
 ,T1.kkckdj_settjk_skosya_ryohi_kbn
 ,T1.kkckdj_settjk_ntktt_ryohi_kbn
 ,T1.kkckdj_settjk_kstts_ryohi_kbn
 ,T1.kkckdj_settjk_tio_naiyo
 ,T1.kkckdj_settjk_tio_cmp_ymd
 ,T1.kkckdj_skosya_tnkn_hms
 ,T1.kkckdj_ntktt_tnkn_hms
 ,T1.kkckdj_kstts_tnkn_hms
 ,T1.kkckdj_skosya_kikhyz_hms
 ,T1.kkckdj_ntktt_kikhyz_hms
 ,T1.kkckdj_kstts_kikhyz_hms
 ,T1.kkckks_ib_skosya_ryohi_kbn
 ,T1.kkckks_ib_ntktt_ryohi_kbn
 ,T1.kkckks_ib_kstts_ryohi_kbn
 ,T1.kkckks_ib_tio_naiyo
 ,T1.kkckks_ib_tio_cmp_ymd
 ,T1.kkckks_hndag_skosya_ryohi_kbn
 ,T1.kkckks_hndag_ntktt_ryohi_kbn
 ,T1.kkckks_hndag_kstts_ryohi_kbn
 ,T1.kkckks_hndag_tio_naiyo
 ,T1.kkckks_hndag_tio_cmp_ymd
 ,T1.kkckks_rosyt_skosya_ryohi_kbn
 ,T1.kkckks_rosyt_ntktt_ryohi_kbn
 ,T1.kkckks_rosyt_kstts_ryohi_kbn
 ,T1.kkckks_rosyt_tio_naiyo
 ,T1.kkckks_rosyt_tio_cmp_ymd
 ,T1.kkckks_bsst_skosya_ryohi_kbn
 ,T1.kkckks_bsst_ntktt_ryohi_kbn
 ,T1.kkckks_bsst_kstts_ryohi_kbn
 ,T1.kkckks_bsst_tio_naiyo
 ,T1.kkckks_bsst_tio_cmp_ymd
 ,T1.kkckkd_tansi_skosya_ryohi_kbn
 ,T1.kkckkd_tansi_ntktt_ryohi_kbn
 ,T1.kkckkd_tansi_kstts_ryohi_kbn
 ,T1.kkckkd_tansi_tio_naiyo
 ,T1.kkckkd_tansi_tio_cmp_ymd
 ,T1.kkchtsn_tekyo_skosya_ryohi_kbn
 ,T1.kkchtsn_tekyo_ntktt_ryohi_kbn
 ,T1.kkchtsn_tekyo_kstts_ryohi_kbn
 ,T1.kkchtsn_tekyo_tio_naiyo
 ,T1.kkchtsn_tekyo_tio_cmp_ymd
 ,T1.kakctkvct_skosya_ryohi_kbn
 ,T1.kakctkvct_ntktt_ryohi_kbn
 ,T1.kakctkvct_kstts_ryohi_kbn
 ,T1.kakctkvct_tio_naiyo
 ,T1.kakctkvct_tio_cmp_ymd
 ,T1.kakctkib_skosya_ryohi_kbn
 ,T1.kakctkib_ntktt_ryohi_kbn
 ,T1.kakctkib_kstts_ryohi_kbn
 ,T1.kakctkib_tio_naiyo
 ,T1.kakctkib_tio_cmp_ymd
 ,T1.kakctkhd_skosya_ryohi_kbn
 ,T1.kakctkhd_ntktt_ryohi_kbn
 ,T1.kakctkhd_kstts_ryohi_kbn
 ,T1.kakctkhd_tio_naiyo
 ,T1.kakctkhd_tio_cmp_ymd
 ,T1.kakctkbs_skosya_ryohi_kbn
 ,T1.kakctkbs_ntktt_ryohi_kbn
 ,T1.kakctkbs_kstts_ryohi_kbn
 ,T1.kakctkbs_tio_naiyo
 ,T1.kakctkbs_tio_cmp_ymd
 ,T1.kakctkss_skosya_ryohi_kbn
 ,T1.kakctkss_ntktt_ryohi_kbn
 ,T1.kakctkss_kstts_ryohi_kbn
 ,T1.kakctkss_tio_naiyo
 ,T1.kakctkss_tio_cmp_ymd
 ,T1.kakct_szno_skosya_ryohi_kbn
 ,T1.kakct_szno_ntktt_ryohi_kbn
 ,T1.kakct_szno_kstts_ryohi_kbn
 ,T1.kakct_szno_tio_naiyo
 ,T1.kakct_szno_tio_cmp_ymd
 ,T1.kakcs_ksn_skosya_ryohi_kbn
 ,T1.kakcs_ksn_ntktt_ryohi_kbn
 ,T1.kakcs_ksn_kstts_ryohi_kbn
 ,T1.kakcs_ksn_tio_naiyo
 ,T1.kakcs_ksn_tio_cmp_ymd
 ,T1.kakcs_kns_skosya_ryohi_kbn
 ,T1.kakcs_kns_ntktt_ryohi_kbn
 ,T1.kakcs_kns_kstts_ryohi_kbn
 ,T1.kakcs_kns_tio_naiyo
 ,T1.kakcs_kns_tio_cmp_ymd
 ,T1.kakcs_skphoto_skosya_ryohi_kbn
 ,T1.kakcs_skphoto_ntktt_ryohi_kbn
 ,T1.kakcs_skphoto_kstts_ryohi_kbn
 ,T1.kakcs_skphoto_tio_naiyo
 ,T1.kakcs_skphoto_tio_cmp_ymd
 ,T1.kakcs_kknn_skosya_ryohi_kbn
 ,T1.kakcs_kknn_ntktt_ryohi_kbn
 ,T1.kakcs_kknn_kstts_ryohi_kbn
 ,T1.kakcs_kknn_tio_naiyo
 ,T1.kakcs_kknn_tio_cmp_ymd
 ,T1.kakcz_skosya_ryohi_kbn
 ,T1.kakcz_ntktt_ryohi_kbn
 ,T1.kakcz_kstts_ryohi_kbn
 ,T1.kakcz_tio_naiyo
 ,T1.kakcz_tio_cmp_ymd
 ,T1.kakca_skosya_ryohi_kbn
 ,T1.kakca_ntktt_ryohi_kbn
 ,T1.kakca_kstts_ryohi_kbn
 ,T1.kakca_tio_naiyo
 ,T1.kakca_tio_cmp_ymd
 ,T1.zstk_kka_bk_naiyo
 ,T1.kahsk_skosya_1s_kstckk_flg
 ,T1.kahsk_skosya_p1_kstckk_flg
 ,T1.kahsk_skosya_p3_kstckk_flg
 ,T1.kahsk_skosya_3s_kstckk_flg
 ,T1.kahsk_skosya_3l_kstckk_flg
 ,T1.kahsk_skosya_p2_kstckk_flg
 ,T1.kahsk_skosya_1l_kstckk_flg
 ,T1.cktrmvct_skosya_1s_kstckk_flg
 ,T1.cktrmvct_skosya_p1_kstckk_flg
 ,T1.cktrmvct_skosya_p3_kstckk_flg
 ,T1.cktrmvct_skosya_3s_kstckk_flg
 ,T1.cktrmvct_skosya_3l_kstckk_flg
 ,T1.cktrmvct_skosya_p2_kstckk_flg
 ,T1.cktrmvct_skosya_1l_kstckk_flg
 ,T1.cktrmwh_skosya_1s_kstckk_flg
 ,T1.cktrmwh_skosya_p1_kstckk_flg
 ,T1.cktrmwh_skosya_p3_kstckk_flg
 ,T1.cktrmwh_skosya_3s_kstckk_flg
 ,T1.cktrmwh_skosya_3l_kstckk_flg
 ,T1.cktrmwh_skosya_p2_kstckk_flg
 ,T1.cktrmwh_skosya_1l_kstckk_flg
 ,T1.djfk_skosya_1s_kstckk_flg
 ,T1.djfk_skosya_p1_kstckk_flg
 ,T1.djfk_skosya_p3_kstckk_flg
 ,T1.djfk_skosya_3s_kstckk_flg
 ,T1.djfk_skosya_3l_kstckk_flg
 ,T1.djfk_skosya_p2_kstckk_flg
 ,T1.djfk_skosya_1l_kstckk_flg
 ,T1.djfk_skosya_dt_kstckk_flg
 ,T1.djfk_skosya_sg_kstckk_flg
 ,T1.sdyyht_skosya_p3_kstckk_flg
 ,T1.sdyyht_skosya_3s_kstckk_flg
 ,T1.sdyyht_skosya_3l_kstckk_flg
 ,T1.sdyyht_skosya_p2_kstckk_flg
 ,T1.jkykrk_skosya_1s_kstckk_flg
 ,T1.jkykrk_skosya_p1_kstckk_flg
 ,T1.jkykrk_skosya_p3_kstckk_flg
 ,T1.jkykrk_skosya_3s_kstckk_flg
 ,T1.jkykrk_skosya_3l_kstckk_flg
 ,T1.jkykrk_skosya_p2_kstckk_flg
 ,T1.jkykrk_skosya_1l_kstckk_flg
 ,T1.kahsk_ntktt_1s_kstckk_flg
 ,T1.kahsk_ntktt_p1_kstckk_flg
 ,T1.kahsk_ntktt_p3_kstckk_flg
 ,T1.kahsk_ntktt_3s_kstckk_flg
 ,T1.kahsk_ntktt_3l_kstckk_flg
 ,T1.kahsk_ntktt_p2_kstckk_flg
 ,T1.kahsk_ntktt_1l_kstckk_flg



 ,T1.cktrmvct_ntktt_1s_kstckk_flg
 ,T1.cktrmvct_ntktt_p1_kstckk_flg
 ,T1.cktrmvct_ntktt_p3_kstckk_flg
 ,T1.cktrmvct_ntktt_3s_kstckk_flg
 ,T1.cktrmvct_ntktt_3l_kstckk_flg
 ,T1.cktrmvct_ntktt_p2_kstckk_flg
 ,T1.cktrmvct_ntktt_1l_kstckk_flg
 ,T1.cktrmwh_ntktt_1s_kstckk_flg
 ,T1.cktrmwh_ntktt_p1_kstckk_flg
 ,T1.cktrmwh_ntktt_p3_kstckk_flg
 ,T1.cktrmwh_ntktt_3s_kstckk_flg
 ,T1.cktrmwh_ntktt_3l_kstckk_flg
 ,T1.cktrmwh_ntktt_p2_kstckk_flg
 ,T1.cktrmwh_ntktt_1l_kstckk_flg
 ,T1.djfk_ntktt_1s_kstckk_flg
 ,T1.djfk_ntktt_p1_kstckk_flg
 ,T1.djfk_ntktt_p3_kstckk_flg
 ,T1.djfk_ntktt_3s_kstckk_flg
 ,T1.djfk_ntktt_3l_kstckk_flg
 ,T1.djfk_ntktt_p2_kstckk_flg

 ,T1.djfk_ntktt_1l_kstckk_flg
 ,T1.djfk_ntktt_dt_kstckk_flg
 ,T1.djfk_ntktt_sg_kstckk_flg
 ,T1.sdyyht_ntktt_p3_kstckk_flg
 ,T1.sdyyht_ntktt_3s_kstckk_flg
 ,T1.sdyyht_ntktt_3l_kstckk_flg
 ,T1.sdyyht_ntktt_p2_kstckk_flg

 ,T1.jkykrk_ntktt_1s_kstckk_flg
 ,T1.jkykrk_ntktt_p1_kstckk_flg
 ,T1.jkykrk_ntktt_p3_kstckk_flg
 ,T1.jkykrk_ntktt_3s_kstckk_flg
 ,T1.jkykrk_ntktt_3l_kstckk_flg
 ,T1.jkykrk_ntktt_p2_kstckk_flg
 ,T1.jkykrk_ntktt_1l_kstckk_flg
 ,T1.kahsk_kstts_1s_kstckk_flg
 ,T1.kahsk_kstts_p1_kstckk_flg
 ,T1.kahsk_kstts_p3_kstckk_flg
 ,T1.kahsk_kstts_3s_kstckk_flg
 ,T1.kahsk_kstts_3l_kstckk_flg
 ,T1.kahsk_kstts_p2_kstckk_flg
 ,T1.kahsk_kstts_1l_kstckk_flg


 ,T1.jkykrk_kstts_1s_kstckk_flg
 ,T1.jkykrk_kstts_p1_kstckk_flg
 ,T1.jkykrk_kstts_p3_kstckk_flg
 ,T1.jkykrk_kstts_3s_kstckk_flg
 ,T1.jkykrk_kstts_3l_kstckk_flg
 ,T1.jkykrk_kstts_p2_kstckk_flg
 ,T1.jkykrk_kstts_1l_kstckk_flg


 ,T1.cktrmvct_ntktt_1s_kstckk_flg
 ,T1.cktrmvct_ntktt_p1_kstckk_flg
 ,T1.cktrmvct_ntktt_p3_kstckk_flg
 ,T1.cktrmvct_ntktt_3s_kstckk_flg
 ,T1.cktrmvct_ntktt_3l_kstckk_flg
 ,T1.cktrmvct_ntktt_p2_kstckk_flg
 ,T1.cktrmvct_ntktt_1l_kstckk_flg
 ,T1.cktrmwh_ntktt_1s_kstckk_flg
 ,T1.cktrmwh_ntktt_p1_kstckk_flg
 ,T1.cktrmwh_ntktt_P3_kstckk_flg
 ,T1.cktrmwh_ntktt_3S_kstckk_flg
 ,T1.cktrmwh_ntktt_3L_kstckk_flg
 ,T1.cktrmwh_ntktt_P2_kstckk_flg
 ,T1.cktrmwh_ntktt_1L_kstckk_flg
 ,T1.djfk_ntktt_1s_kstckk_flg
 ,T1.djfk_ntktt_p1_kstckk_flg
 ,T1.djfk_ntktt_p3_kstckk_flg
 ,T1.djfk_ntktt_3s_kstckk_flg
 ,T1.djfk_ntktt_3l_kstckk_flg
 ,T1.djfk_ntktt_p2_kstckk_flg
 ,T1.djfk_ntktt_1l_kstckk_flg
 ,T1.djfk_ntktt_dt_kstckk_flg
 ,T1.djfk_ntktt_sg_kstckk_flg
 ,T1.sdyyht_ntktt_p3_kstckk_flg
 ,T1.sdyyht_ntktt_3s_kstckk_flg
 ,T1.sdyyht_ntktt_3l_kstckk_flg
 ,T1.sdyyht_ntktt_p2_kstckk_flg
 ,T1.jkykrk_ntktt_1s_kstckk_flg
 ,T1.jkykrk_ntktt_p1_kstckk_flg
 ,T1.jkykrk_ntktt_p3_kstckk_flg
 ,T1.jkykrk_ntktt_3s_kstckk_flg
 ,T1.jkykrk_ntktt_3l_kstckk_flg
 ,T1.jkykrk_ntktt_p2_kstckk_flg
 ,T1.jkykrk_ntktt_1l_kstckk_flg
 ,T1.kahsk_kstts_1s_kstckk_flg
 ,T1.kahsk_kstts_p1_kstckk_flg
 ,T1.kahsk_kstts_p3_kstckk_flg
 ,T1.kahsk_kstts_3s_kstckk_flg
 ,T1.kahsk_kstts_3l_kstckk_flg
 ,T1.kahsk_kstts_p2_kstckk_flg
 ,T1.kahsk_kstts_1l_kstckk_flg
 ,T1.cktrmvct_kstts_1s_kstckk_flg
 ,T1.cktrmvct_kstts_p1_kstckk_flg
 ,T1.cktrmvct_kstts_p3_kstckk_flg
 ,T1.cktrmvct_kstts_3s_kstckk_flg
 ,T1.cktrmvct_kstts_3l_kstckk_flg
 ,T1.cktrmvct_kstts_p2_kstckk_flg
 ,T1.cktrmvct_kstts_1l_kstckk_flg
 ,T1.cktrmwh_kstts_1s_kstckk_flg
 ,T1.cktrmwh_kstts_p1_kstckk_flg
 ,T1.cktrmwh_kstts_P3_kstckk_flg
 ,T1.cktrmwh_kstts_3S_kstckk_flg
 ,T1.cktrmwh_kstts_3L_kstckk_flg
 ,T1.cktrmwh_kstts_P2_kstckk_flg
 ,T1.cktrmwh_kstts_1L_kstckk_flg
 ,T1.djfk_kstts_1s_kstckk_flg
 ,T1.djfk_kstts_p1_kstckk_flg
 ,T1.djfk_kstts_p3_kstckk_flg
 ,T1.djfk_kstts_3s_kstckk_flg
 ,T1.djfk_kstts_3l_kstckk_flg
 ,T1.djfk_kstts_p2_kstckk_flg
 ,T1.djfk_kstts_1l_kstckk_flg
 ,T1.djfk_kstts_dt_kstckk_flg
 ,T1.djfk_kstts_sg_kstckk_flg
 ,T1.sdyyht_kstts_p3_kstckk_flg
 ,T1.sdyyht_kstts_3s_kstckk_flg
 ,T1.sdyyht_kstts_3l_kstckk_flg
 ,T1.sdyyht_kstts_p2_kstckk_flg
 ,T1.jkykrk_kstts_1s_kstckk_flg
 ,T1.jkykrk_kstts_p1_kstckk_flg
 ,T1.jkykrk_kstts_p3_kstckk_flg
 ,T1.jkykrk_kstts_3s_kstckk_flg
 ,T1.jkykrk_kstts_3l_kstckk_flg
 ,T1.jkykrk_kstts_p2_kstckk_flg
 ,T1.jkykrk_kstts_1l_kstckk_flg
 ,T1.takck_yr_skosya_ryohi_kbn
 ,T1.takck_yr_ntktt_ryohi_kbn
 ,T1.takck_yr_kstts_ryohi_kbn
 ,T1.takck_yr_tio_naiyo
 ,T1.takck_yr_tio_cmp_ymd
 ,T1.takck_sm_skosya_ryohi_kbn
 ,T1.takck_sm_ntktt_ryohi_kbn
 ,T1.takck_sm_kstts_ryohi_kbn
 ,T1.takck_sm_tio_naiyo
 ,T1.takck_sm_tio_cmp_ymd
 ,T1.takck_kkano_skosya_ryohi_kbn
 ,T1.takck_kkano_ntktt_ryohi_kbn
 ,T1.takck_kkano_kstts_ryohi_kbn
 ,T1.takck_kkano_tio_naiyo
 ,T1.takck_kkano_tio_cmp_ymd
 ,T1.takck_hksno_skosya_ryohi_kbn
 ,T1.takck_hksno_ntktt_ryohi_kbn
 ,T1.takck_hksno_kstts_ryohi_kbn
 ,T1.takck_hksno_tio_naiyo
 ,T1.takck_hksno_tio_cmp_ymd
 ,T1.takck_hkano_skosya_ryohi_kbn
 ,T1.takck_hkano_ntktt_ryohi_kbn
 ,T1.takck_hkano_kstts_ryohi_kbn
 ,T1.takck_hkano_tio_naiyo
 ,T1.takck_hkano_tio_cmp_ymd
 ,T1.takckk_krkkd_skosya_ryohi_kbn
 ,T1.takckk_krkkd_ntktt_ryohi_kbn
 ,T1.takckk_krkkd_kstts_ryohi_kbn
 ,T1.takckk_krkkd_tio_naiyo
 ,T1.takckk_krkkd_tio_cmp_ymd
 ,T1.takckk_krkb_skosya_ryohi_kbn
 ,T1.takckk_krkb_ntktt_ryohi_kbn
 ,T1.takckk_krkb_kstts_ryohi_kbn
 ,T1.takckk_krkb_tio_naiyo
 ,T1.takckk_krkb_tio_cmp_ymd
 ,T1.takcks_kkzen_skosya_ryohi_kbn
 ,T1.takcks_kkzen_ntktt_ryohi_kbn
 ,T1.takcks_kkzen_kstts_ryohi_kbn
 ,T1.takcks_kkzen_tio_naiyo
 ,T1.takcks_kkzen_tio_cmp_ymd
 ,T1.takcks_kksl_skosya_ryohi_kbn
 ,T1.takcks_kksl_ntktt_ryohi_kbn
 ,T1.takcks_kksl_kstts_ryohi_kbn
 ,T1.takcks_kksl_tio_naiyo
 ,T1.takcks_kksl_tio_cmp_ymd
 ,T1.takcks_kkhks_skosya_ryohi_kbn
 ,T1.takcks_kkhks_ntktt_ryohi_kbn
 ,T1.takcks_kkhks_kstts_ryohi_kbn
 ,T1.takcks_kkhks_tio_naiyo
 ,T1.takcks_kkhks_tio_cmp_ymd
 ,T1.takcks_kkgsz_skosya_ryohi_kbn
 ,T1.takcks_kkgsz_ntktt_ryohi_kbn
 ,T1.takcks_kkgsz_kstts_ryohi_kbn
 ,T1.takcks_kkgsz_tio_naiyo
 ,T1.takcks_kkgsz_tio_cmp_ymd
 ,T1.takcks_kkaki_skosya_ryohi_kbn
 ,T1.takcks_kkaki_ntktt_ryohi_kbn
 ,T1.takcks_kkaki_kstts_ryohi_kbn
 ,T1.takcks_kkaki_tio_naiyo
 ,T1.takcks_kkaki_tio_cmp_ymd
 ,T1.takcks_kkyrs_skosya_ryohi_kbn
 ,T1.takcks_kkyrs_ntktt_ryohi_kbn
 ,T1.takcks_kkyrs_kstts_ryohi_kbn
 ,T1.takcks_kkyrs_tio_naiyo
 ,T1.takcks_kkyrs_tio_cmp_ymd
 ,T1.takcks_kkrst_skosya_ryohi_kbn
 ,T1.takcks_kkrst_ntktt_ryohi_kbn
 ,T1.takcks_kkrst_kstts_ryohi_kbn
 ,T1.takcks_kkrst_tio_naiyo
 ,T1.takcks_kkrst_tio_cmp_ymd
 ,T1.takcks_kkbst_skosya_ryohi_kbn
 ,T1.takcks_kkbst_ntktt_ryohi_kbn
 ,T1.takcks_kkbst_kstts_ryohi_kbn
 ,T1.takcks_kkbst_tio_naiyo
 ,T1.takcks_kkbst_tio_cmp_ymd
 ,T1.takcks_kikib_skosya_ryohi_kbn
 ,T1.takcks_kikib_ntktt_ryohi_kbn
 ,T1.takcks_kikib_kstts_ryohi_kbn
 ,T1.takcks_kikib_tio_naiyo
 ,T1.takcks_kikib_tio_cmp_ymd
 ,T1.takcks_kkssu_skosya_ryohi_kbn
 ,T1.takcks_kkssu_ntktt_ryohi_kbn
 ,T1.takcks_kkssu_kstts_ryohi_kbn
 ,T1.takcks_kkssu_tio_naiyo
 ,T1.takcks_kkssu_tio_cmp_ymd
 ,T1.takcks_hrkkl_skosya_ryohi_kbn
 ,T1.takcks_hrkkl_ntktt_ryohi_kbn
 ,T1.takcks_hrkkl_kstts_ryohi_kbn
 ,T1.takcks_hrkkl_tio_naiyo
 ,T1.takcks_hrkkl_tio_cmp_ymd
 ,T1.takcks_hp123_skosya_ryohi_kbn
 ,T1.takcks_hp123_ntktt_ryohi_kbn
 ,T1.takcks_hp123_kstts_ryohi_kbn
 ,T1.takcks_hp123_tio_naiyo
 ,T1.takcks_hp123_tio_cmp_ymd
 ,T1.takcks_hkaki_skosya_ryohi_kbn
 ,T1.takcks_hkaki_ntktt_ryohi_kbn
 ,T1.takcks_hkaki_kstts_ryohi_kbn
 ,T1.takcks_hkaki_tio_naiyo
 ,T1.takcks_hkaki_tio_cmp_ymd
 ,T1.takckrk_hnrk_skosya_ryohi_kbn
 ,T1.takckrk_hnrk_ntktt_ryohi_kbn
 ,T1.takckrk_hnrk_kstts_ryohi_kbn
 ,T1.takckrk_hnrk_tio_naiyo
 ,T1.takckrk_hnrk_tio_cmp_ymd
 ,T1.takckrr_sm_skosya_ryohi_kbn
 ,T1.takckrr_sm_ntktt_ryohi_kbn
 ,T1.takckrr_sm_kstts_ryohi_kbn
 ,T1.takckrr_sm_tio_naiyo
 ,T1.takckrr_sm_tio_cmp_ymd
 ,T1.takcc_led_skosya_ryohi_kbn
 ,T1.takcc_led_ntktt_ryohi_kbn
 ,T1.takcc_led_kstts_ryohi_kbn
 ,T1.takcc_led_tio_naiyo
 ,T1.takcc_led_tio_cmp_ymd
 ,T1.takcc_souck_skosya_ryohi_kbn
 ,T1.takcc_souck_ntktt_ryohi_kbn
 ,T1.takcc_souck_kstts_ryohi_kbn
 ,T1.takcc_souck_tio_naiyo
 ,T1.takcc_souck_tio_cmp_ymd
 ,T1.takct_szno_skosya_ryohi_kbn
 ,T1.takct_szno_ntktt_ryohi_kbn
 ,T1.takct_szno_kstts_ryohi_kbn
 ,T1.takct_szno_tio_naiyo
 ,T1.takct_szno_tio_cmp_ymd
 ,T1.takct_aino_skosya_ryohi_kbn
 ,T1.takct_aino_ntktt_ryohi_kbn
 ,T1.takct_aino_kstts_ryohi_kbn
 ,T1.takct_aino_tio_naiyo
 ,T1.takct_aino_tio_cmp_ymd
 ,T1.takcs_photo_skosya_ryohi_kbn
 ,T1.takcs_photo_ntktt_ryohi_kbn
 ,T1.takcs_photo_kstts_ryohi_kbn
 ,T1.takcs_photo_tio_naiyo
 ,T1.takcs_photo_tio_cmp_ymd
 ,T1.takcs_tktrk_skosya_ryohi_kbn
 ,T1.takcs_tktrk_ntktt_ryohi_kbn
 ,T1.takcs_tktrk_kstts_ryohi_kbn
 ,T1.takcs_tktrk_tio_naiyo
 ,T1.takcs_tktrk_tio_cmp_ymd
 ,T1.takcz_ntktt_seigo_kknn_kbn
 ,T1.takcz_kstts_seigo_kknn_kbn
 ,T1.takcz_tio_naiyo
 ,T1.takcz_tio_cmp_ymd
 ,T1.takca_ntktt_itti_kknn_kbn
 ,T1.takca_kstts_itti_kknn_kbn
 ,T1.takca_tio_naiyo
 ,T1.takca_tio_cmp_ymd
 ,T1.kizi_skosya_naiyo
 ,T1.kizi_ntktt_naiyo
 ,T1.kizi_kstts_naiyo
 ,T1.hrktkk_skosya_1s_kstckk_flg
 ,T1.hrktkk_skosya_p1_kstckk_flg
 ,T1.hrktkk_skosya_p3_kstckk_flg
 ,T1.hrktkk_skosya_3s_kstckk_flg
 ,T1.hrktkk_skosya_3l_kstckk_flg
 ,T1.hrktkk_skosya_p2_kstckk_flg
 ,T1.hrktkk_skosya_1l_kstckk_flg
 ,T1.p1_skosya_1s_kstckk_flg
 ,T1.p1_skosya_p1_kstckk_flg
 ,T1.p1_skosya_1l_kstckk_flg
 ,T1.p2_skosya_p2_kstckk_flg
 ,T1.p3_skosya_3s_kstckk_flg
 ,T1.p3_skosya_p3_kstckk_flg
 ,T1.p3_skosya_3l_kstckk_flg
 ,T1.hrktkk_ntktt_1s_kstckk_flg
 ,T1.hrktkk_ntktt_p1_kstckk_flg
 ,T1.hrktkk_ntktt_p3_kstckk_flg
 ,T1.hrktkk_ntktt_3s_kstckk_flg
 ,T1.hrktkk_ntktt_3l_kstckk_flg
 ,T1.hrktkk_ntktt_p2_kstckk_flg
 ,T1.hrktkk_ntktt_1l_kstckk_flg
 ,T1.p1_ntktt_1s_kstckk_flg
 ,T1.p1_ntktt_p1_kstckk_flg
 ,T1.p1_ntktt_1l_kstckk_flg
 ,T1.p2_ntktt_p2_kstckk_flg
 ,T1.p3_ntktt_3s_kstckk_flg
 ,T1.p3_ntktt_p3_kstckk_flg
 ,T1.p3_ntktt_3l_kstckk_flg
 ,T1.hrktkk_kstts_1s_kstckk_flg
 ,T1.hrktkk_kstts_p1_kstckk_flg
 ,T1.hrktkk_kstts_p3_kstckk_flg
 ,T1.hrktkk_kstts_3s_kstckk_flg
 ,T1.hrktkk_kstts_3l_kstckk_flg
 ,T1.hrktkk_kstts_p2_kstckk_flg
 ,T1.hrktkk_kstts_1l_kstckk_flg
 ,T1.p1_kstts_1s_kstckk_flg
 ,T1.p1_kstts_p1_kstckk_flg
 ,T1.p1_kstts_1l_kstckk_flg
 ,T1.p2_kstts_p2_kstckk_flg
 ,T1.p3_kstts_3s_kstckk_flg
 ,T1.p3_kstts_p3_kstckk_flg
 ,T1.p3_kstts_3l_kstckk_flg

-- 低圧分
 ,T1.takcd_ttt_skosya_ryohi_kbn
 ,T1.takcd_ttt_ntktt_ryohi_kbn
 ,T1.takcd_ttt_kstts_ryohi_kbn
 ,T1.takcd_ttt_tio_naiyo
 ,T1.takcd_ttt_tio_cmp_ymd
 ,T1.takcd_szsu_skosya_ryohi_kbn
 ,T1.takcd_szsu_ntktt_ryohi_kbn
 ,T1.takcd_szsu_kstts_ryohi_kbn
 ,T1.takcd_szsu_tio_naiyo
 ,T1.takcd_szsu_tio_cmp_ymd
 ,T1.takck_ssda_skosya_ryohi_kbn
 ,T1.takck_ssda_ntktt_ryohi_kbn
 ,T1.takck_ssda_kstts_ryohi_kbn
 ,T1.takck_ssda_tio_naiyo
 ,T1.takck_ssda_tio_cmp_ymd
 ,T1.takck_zyrt_skosya_ryohi_kbn
 ,T1.takck_zyrt_ntktt_ryohi_kbn
 ,T1.takck_zyrt_kstts_ryohi_kbn
 ,T1.takck_zyrt_tio_naiyo
 ,T1.takck_zyrt_tio_cmp_ymd
 ,T1.takck_yr_skosya_ryohi_kbn
 ,T1.takck_yr_ntktt_ryohi_kbn
 ,T1.takck_yr_kstts_ryohi_kbn
 ,T1.takck_yr_tio_naiyo
 ,T1.takck_yr_tio_cmp_ymd
 ,T1.takck_sm_skosya_ryohi_kbn
 ,T1.takck_sm_ntktt_ryohi_kbn
 ,T1.takck_sm_kstts_ryohi_kbn
 ,T1.takck_sm_tio_naiyo
 ,T1.takck_sm_tio_cmp_ymd
 ,T1.takck_kkano_skosya_ryohi_kbn
 ,T1.takck_kkano_ntktt_ryohi_kbn
 ,T1.takck_kkano_kstts_ryohi_kbn
 ,T1.takck_kkano_tio_naiyo
 ,T1.takck_kkano_tio_cmp_ymd
 ,T1.takck_hksno_skosya_ryohi_kbn
 ,T1.takck_hksno_ntktt_ryohi_kbn
 ,T1.takck_hksno_kstts_ryohi_kbn
 ,T1.takck_hksno_tio_naiyo
 ,T1.takck_hksno_tio_cmp_ymd
 ,T1.takck_hkano_skosya_ryohi_kbn
 ,T1.takck_hkano_ntktt_ryohi_kbn
 ,T1.takck_hkano_kstts_ryohi_kbn
 ,T1.takck_hkano_tio_naiyo
 ,T1.takck_hkano_tio_cmp_ymd
 ,T1.takckk_krkkd_skosya_ryohi_kbn
 ,T1.takckk_krkkd_ntktt_ryohi_kbn
 ,T1.takckk_krkkd_kstts_ryohi_kbn
 ,T1.takckk_krkkd_tio_naiyo
 ,T1.takckk_krkkd_tio_cmp_ymd
 ,T1.takckk_krkb_skosya_ryohi_kbn
 ,T1.takckk_krkb_ntktt_ryohi_kbn
 ,T1.takckk_krkb_kstts_ryohi_kbn
 ,T1.takckk_krkb_tio_naiyo
 ,T1.takckk_krkb_tio_cmp_ymd
 ,T1.takcks_kkzen_skosya_ryohi_kbn
 ,T1.takcks_kkzen_ntktt_ryohi_kbn
 ,T1.takcks_kkzen_kstts_ryohi_kbn
 ,T1.takcks_kkzen_tio_naiyo
 ,T1.takcks_kkzen_tio_cmp_ymd
 ,T1.takcks_kksl_skosya_ryohi_kbn
 ,T1.takcks_kksl_ntktt_ryohi_kbn
 ,T1.takcks_kksl_kstts_ryohi_kbn
 ,T1.takcks_kksl_tio_naiyo
 ,T1.takcks_kksl_tio_cmp_ymd
 ,T1.takcks_kkhks_skosya_ryohi_kbn
 ,T1.takcks_kkhks_ntktt_ryohi_kbn
 ,T1.takcks_kkhks_kstts_ryohi_kbn
 ,T1.takcks_kkhks_tio_naiyo
 ,T1.takcks_kkhks_tio_cmp_ymd
 ,T1.takcks_kkgsz_skosya_ryohi_kbn
 ,T1.takcks_kkgsz_ntktt_ryohi_kbn
 ,T1.takcks_kkgsz_kstts_ryohi_kbn
 ,T1.takcks_kkgsz_tio_naiyo
 ,T1.takcks_kkgsz_tio_cmp_ymd
 ,T1.takcks_kkaki_skosya_ryohi_kbn
 ,T1.takcks_kkaki_ntktt_ryohi_kbn
 ,T1.takcks_kkaki_kstts_ryohi_kbn
 ,T1.takcks_kkaki_tio_naiyo
 ,T1.takcks_kkaki_tio_cmp_ymd
 ,T1.takcks_kkyrs_skosya_ryohi_kbn
 ,T1.takcks_kkyrs_ntktt_ryohi_kbn
 ,T1.takcks_kkyrs_kstts_ryohi_kbn
 ,T1.takcks_kkyrs_tio_naiyo
 ,T1.takcks_kkyrs_tio_cmp_ymd
 ,T1.takcks_kkrst_skosya_ryohi_kbn
 ,T1.takcks_kkrst_ntktt_ryohi_kbn
 ,T1.takcks_kkrst_kstts_ryohi_kbn
 ,T1.takcks_kkrst_tio_naiyo
 ,T1.takcks_kkrst_tio_cmp_ymd
 ,T1.takcks_kkbst_skosya_ryohi_kbn
 ,T1.takcks_kkbst_ntktt_ryohi_kbn
 ,T1.takcks_kkbst_kstts_ryohi_kbn
 ,T1.takcks_kkbst_tio_naiyo
 ,T1.takcks_kkbst_tio_cmp_ymd
 ,T1.takcks_kikib_skosya_ryohi_kbn
 ,T1.takcks_kikib_ntktt_ryohi_kbn
 ,T1.takcks_kikib_kstts_ryohi_kbn
 ,T1.takcks_kikib_tio_naiyo
 ,T1.takcks_kikib_tio_cmp_ymd
 ,T1.takcks_kkssu_skosya_ryohi_kbn
 ,T1.takcks_kkssu_ntktt_ryohi_kbn
 ,T1.takcks_kkssu_kstts_ryohi_kbn
 ,T1.takcks_kkssu_tio_naiyo
 ,T1.takcks_kkssu_tio_cmp_ymd
 ,T1.takcks_hrkkl_skosya_ryohi_kbn
 ,T1.takcks_hrkkl_ntktt_ryohi_kbn
 ,T1.takcks_hrkkl_kstts_ryohi_kbn
 ,T1.takcks_hrkkl_tio_naiyo
 ,T1.takcks_hrkkl_tio_cmp_ymd
 ,T1.takcks_hp123_skosya_ryohi_kbn
 ,T1.takcks_hp123_ntktt_ryohi_kbn
 ,T1.takcks_hp123_kstts_ryohi_kbn
 ,T1.takcks_hp123_tio_naiyo
 ,T1.takcks_hp123_tio_cmp_ymd
 ,T1.takcks_hkaki_skosya_ryohi_kbn
 ,T1.takcks_hkaki_ntktt_ryohi_kbn
 ,T1.takcks_hkaki_kstts_ryohi_kbn
 ,T1.takcks_hkaki_tio_naiyo
 ,T1.takcks_hkaki_tio_cmp_ymd
 ,T1.takckrk_hnrk_skosya_ryohi_kbn
 ,T1.takckrk_hnrk_ntktt_ryohi_kbn
 ,T1.takckrk_hnrk_kstts_ryohi_kbn
 ,T1.takckrk_hnrk_tio_naiyo
 ,T1.takckrk_hnrk_tio_cmp_ymd
 ,T1.takckrr_sm_skosya_ryohi_kbn
 ,T1.takckrr_sm_ntktt_ryohi_kbn
 ,T1.takckrr_sm_kstts_ryohi_kbn
 ,T1.takckrr_sm_tio_naiyo
 ,T1.takckrr_sm_tio_cmp_ymd
 ,T1.takcc_led_skosya_ryohi_kbn
 ,T1.takcc_led_ntktt_ryohi_kbn
 ,T1.takcc_led_kstts_ryohi_kbn
 ,T1.takcc_led_tio_naiyo
 ,T1.takcc_led_tio_cmp_ymd
 ,T1.takcc_souck_skosya_ryohi_kbn
 ,T1.takcc_souck_ntktt_ryohi_kbn
 ,T1.takcc_souck_kstts_ryohi_kbn
 ,T1.takcc_souck_tio_naiyo
 ,T1.takcc_souck_tio_cmp_ymd
 ,T1.takct_szno_skosya_ryohi_kbn
 ,T1.takct_szno_ntktt_ryohi_kbn
 ,T1.takct_szno_kstts_ryohi_kbn
 ,T1.takct_szno_tio_naiyo
 ,T1.takct_szno_tio_cmp_ymd
 ,T1.takct_aino_skosya_ryohi_kbn
 ,T1.takct_aino_ntktt_ryohi_kbn
 ,T1.takct_aino_kstts_ryohi_kbn
 ,T1.takct_aino_tio_naiyo
 ,T1.takct_aino_tio_cmp_ymd
 ,T1.takcs_photo_skosya_ryohi_kbn
 ,T1.takcs_photo_ntktt_ryohi_kbn
 ,T1.takcs_photo_kstts_ryohi_kbn
 ,T1.takcs_photo_tio_naiyo
 ,T1.takcs_photo_tio_cmp_ymd
 ,T1.takcs_tktrk_skosya_ryohi_kbn
 ,T1.takcs_tktrk_ntktt_ryohi_kbn
 ,T1.takcs_tktrk_kstts_ryohi_kbn
 ,T1.takcs_tktrk_tio_naiyo
 ,T1.takcs_tktrk_tio_cmp_ymd
 ,T1.takcz_ntktt_seigo_kknn_kbn
 ,T1.takcz_kstts_seigo_kknn_kbn
 ,T1.takcz_tio_naiyo
 ,T1.takcz_tio_cmp_ymd
 ,T1.takca_ntktt_itti_kknn_kbn
 ,T1.takca_kstts_itti_kknn_kbn
 ,T1.takca_tio_naiyo
 ,T1.takca_tio_cmp_ymd
 ,T1.kizi_skosya_naiyo
 ,T1.kizi_ntktt_naiyo
 ,T1.kizi_kstts_naiyo
 ,T1.hrktkk_skosya_1s_kstckk_flg
 ,T1.hrktkk_skosya_p1_kstckk_flg
 ,T1.hrktkk_skosya_p3_kstckk_flg
 ,T1.hrktkk_skosya_3s_kstckk_flg
 ,T1.hrktkk_skosya_3l_kstckk_flg
 ,T1.hrktkk_skosya_p2_kstckk_flg
 ,T1.hrktkk_skosya_1l_kstckk_flg
 ,T1.p1_skosya_1s_kstckk_flg
 ,T1.p1_skosya_p1_kstckk_flg
 ,T1.p1_skosya_1l_kstckk_flg
 ,T1.p2_skosya_p2_kstckk_flg
 ,T1.p3_skosya_3s_kstckk_flg
 ,T1.p3_skosya_p3_kstckk_flg
 ,T1.p3_skosya_3l_kstckk_flg
 ,T1.hrktkk_ntktt_1s_kstckk_flg
 ,T1.hrktkk_ntktt_p1_kstckk_flg
 ,T1.hrktkk_ntktt_p3_kstckk_flg
 ,T1.hrktkk_ntktt_3s_kstckk_flg
 ,T1.hrktkk_ntktt_3l_kstckk_flg
 ,T1.hrktkk_ntktt_p2_kstckk_flg
 ,T1.hrktkk_ntktt_1l_kstckk_flg
 ,T1.p1_ntktt_1s_kstckk_flg
 ,T1.p1_ntktt_p1_kstckk_flg
 ,T1.p1_ntktt_1l_kstckk_flg
 ,T1.p2_ntktt_p2_kstckk_flg
 ,T1.p3_ntktt_3s_kstckk_flg
 ,T1.p3_ntktt_p3_kstckk_flg
 ,T1.p3_ntktt_3l_kstckk_flg
 ,T1.hrktkk_kstts_1s_kstckk_flg
 ,T1.hrktkk_kstts_p1_kstckk_flg
 ,T1.hrktkk_kstts_p3_kstckk_flg
 ,T1.hrktkk_kstts_3s_kstckk_flg
 ,T1.hrktkk_kstts_3l_kstckk_flg
 ,T1.hrktkk_kstts_p2_kstckk_flg
 ,T1.hrktkk_kstts_1l_kstckk_flg
 ,T1.p1_kstts_1s_kstckk_flg
 ,T1.p1_kstts_p1_kstckk_flg
 ,T1.p1_kstts_1l_kstckk_flg
 ,T1.p2_kstts_p2_kstckk_flg
 ,T1.p3_kstts_3s_kstckk_flg
 ,T1.p3_kstts_p3_kstckk_flg
 ,T1.p3_kstts_3l_kstckk_flg

 ,T1.takckrr_smhyz_skosya_ryohi_kbn 
 ,T1.takckrr_smhyz_ntktt_ryohi_kbn 
 ,T1.takckrr_smhyz_kstts_ryohi_kbn 
 ,T1.takckrr_smhyz_tio_naiyo 
 ,T1.takckrr_smhyz_tio_cmp_ymd 
 ,T1.takckrr_d2sm_skosya_ryohi_kbn 
 ,T1.takckrr_d2sm_ntktt_ryohi_kbn
 ,T1.takckrr_d2sm_kstts_ryohi_kbn
 ,T1.takckrr_d2sm_tio_naiyo
 ,T1.takckrr_d2sm_tio_cmp_ymd
 ,T1.takckrr_khk_skosya_ryohi_kbn
 ,T1.takckrr_khk_ntktt_ryohi_kbn 
 ,T1.takckrr_khk_kstts_ryohi_kbn 
 ,T1.takckrr_khk_tio_naiyo 
 ,T1.takckrr_khk_tio_cmp_ymd 
 ,T1.takckk_okgkb_skosya_ryohi_kbn 
 ,T1.takckk_okgkb_ntktt_ryohi_kbn
 ,T1.takckk_okgkb_kstts_ryohi_kbn
 ,T1.takckk_okgkb_tio_naiyo
 ,T1.takckk_okgkb_tio_cmp_ymd
 ,T1.takckk_dsipcv_skosya_ryohi_kbn
 ,T1.takckk_dsipcv_ntktt_ryohi_kbn 
 ,T1.takckk_dsipcv_kstts_ryohi_kbn 
 ,T1.takckk_dsipcv_tio_naiyo 
 ,T1.takckk_dsipcv_tio_cmp_ymd 
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_自主点検チェックシート（高圧他）
  INNER JOIN trke_check_sheet_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
WHERE 1 = 1 
AND D1.downld_renno={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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

    'Rev002 Start
#Region "(SQL)ダウンロード管理詳細・共通_添付書類詳細（高圧他）　検索"
    ''' <summary>
    ''' ダウンロード管理詳細・共通_添付書類詳細（高圧他）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pDownldRenno">ダウンロード番号</param>
    ''' <remarks>ダウンロード管理詳細・共通_添付書類詳細（高圧他）を検索します</remarks>
    Friend Function S021(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pDownldRenno As String) As Boolean

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
select TP.*
FROM (
(
select
  D1.dig4_zgsyo_cd AS dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS trkhy_hakko_nendo
 ,D1.trkhy_kbn AS trkhy_kbn
 ,D1.trkhy_no As trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.tnpdoc_mng_renno
 ,K1.file_renno
 ,K1.file_ms
 ,K1.no1_tnpdoc_hosok_naiyo
 ,K1.no2_tnpdoc_hosok_naiyo
 ,K1.no3_tnpdoc_hosok_naiyo
 ,K1.no4_tnpdoc_hosok_naiyo
 ,K1.no5_tnpdoc_hosok_naiyo
 ,K1.no6_tnpdoc_hosok_naiyo
 ,K1.no7_tnpdoc_hosok_naiyo
 ,K1.no8_tnpdoc_hosok_naiyo
 ,K1.no9_tnpdoc_hosok_naiyo
 ,K1.no10_tnpdoc_hosok_naiyo
 ,M1.prm_cd_ms AS photo_sbt_ms
 ,T1.tenp_file_mng_no
 ,T1.ryssy_tenp_file_mng_no
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（高圧他）
  INNER JOIN trke_trkehy_kahk T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--共通_添付書類詳細
  INNER JOIN com_tnpdoc_dtl K1 
ON T1.tenp_file_mng_no = K1.tnpdoc_mng_renno 
--パラメータコードマスタ(写真種別)
  LEFT  JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'PHOTO_SBT_CD' 
   AND M1.prm_cd = K1.no2_tnpdoc_hosok_naiyo
WHERE 1 = 1 
AND D1.downld_renno={0}
)
UNION
(
select
  D1.dig4_zgsyo_cd AS dig4_zgsyo_cd
 ,D1.trkhy_hakko_nendo AS trkhy_hakko_nendo
 ,D1.trkhy_kbn AS trkhy_kbn
 ,D1.trkhy_no As trkhy_no
 ,K1.saksi_date
 ,K1.upd_date
 ,K1.sousa_user_id
 ,K1.sousa_appli_cd
 ,K1.tnpdoc_mng_renno
 ,K1.file_renno
 ,K1.file_ms
 ,K1.no1_tnpdoc_hosok_naiyo
 ,K1.no2_tnpdoc_hosok_naiyo
 ,K1.no3_tnpdoc_hosok_naiyo
 ,K1.no4_tnpdoc_hosok_naiyo
 ,K1.no5_tnpdoc_hosok_naiyo
 ,K1.no6_tnpdoc_hosok_naiyo
 ,K1.no7_tnpdoc_hosok_naiyo
 ,K1.no8_tnpdoc_hosok_naiyo
 ,K1.no9_tnpdoc_hosok_naiyo
 ,K1.no10_tnpdoc_hosok_naiyo
 ,M1.prm_cd_ms AS photo_sbt_ms
 ,T1.tenp_file_mng_no
 ,T1.ryssy_tenp_file_mng_no
--ダウンロード管理明細
  FROM trke_downld_dtl D1 
--取替_取替票（低圧）
  INNER JOIN trke_trkehy_tiat T1 
ON D1.dig4_zgsyo_cd = T1.dig4_zgsyo_cd 
   AND D1.trkhy_hakko_nendo = T1.trkhy_hakko_nendo
   AND D1.trkhy_kbn = T1.trkhy_kbn
   AND D1.trkhy_no = T1.trkhy_no
--共通_添付書類詳細
  INNER JOIN com_tnpdoc_dtl K1 
ON T1.tenp_file_mng_no = K1.tnpdoc_mng_renno 
--パラメータコードマスタ(写真種別)
  LEFT  JOIN  com_paramater_cd_mst M1 
ON M1.prm_id = 'PHOTO_SBT_CD' 
   AND M1.prm_cd = K1.no2_tnpdoc_hosok_naiyo
WHERE 1 = 1 
AND D1.downld_renno={0}
)
) TP
ORDER BY 
  TP.dig4_zgsyo_cd
 ,TP.trkhy_hakko_nendo
 ,TP.trkhy_kbn
 ,TP.trkhy_no
 ,TP.file_renno
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pDownldRenno))
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


    'Rev002-End


#Region "(SQL)共通_メッセージマスタ　検索"
    '    ''' <summary>
    '    ''' 共通_メッセージマスタ　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>共通_メッセージマスタを検索します</remarks>
    '    Friend Function S008(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,msg_id
    ' ,msg_naiyo
    '  FROM COM_MSG_MST 
    'ORDER BY 
    '  msg_id
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
#End Region

#Region "(SQL)マスタ_工量　検索"
    '    ''' <summary>
    '    ''' マスタ_工量　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>マスタ_工量を検索します</remarks>
    '    Friend Function S009(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,koryo_zunit_cd
    ' ,tky_start_ymd
    ' ,tky_end_ymd
    ' ,trtk_koryo
    ' ,tekyo_koryo
    ' ,koryo_tanka_kngk
    ' ,kohi_tanka_kngk
    ' ,koryo_kohi_sbt_cd
    ' ,koryo_mskn
    ' ,koryo_ms
    ' ,koryo_tanka_cd
    ' ,wrms_ritu
    ' ,trtk_kohi_tanka_kngk
    ' ,tekyo_kohi_tanka_kngk
    ' ,koryo_idou_kbn
    ' ,koryo_cd_tky_start_ymd
    '  FROM MST_KORYO 
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
#End Region

#Region "(SQL)パラメータコードマスタ　検索"
    '    ''' <summary>
    '    ''' パラメータコードマスタ　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>パラメータコードマスタを検索します</remarks>
    '    Friend Function S010(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,prm_id
    ' ,prm_cd
    ' ,prm_cd_ms
    ' ,prm_cd_abms
    ' ,str1_hksu_value
    ' ,str2_hksu_value
    ' ,value1_hksu_value
    ' ,value2_hksu_value
    ' ,prm_hyz_zyun
    ' ,HYZ_FLG
    '  FROM COM_PARAMATER_CD_MST
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
#End Region

#Region "(SQL)品目マスタ　検索"
    '    ''' <summary>
    '    ''' 品目マスタ　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>品目マスタを検索します</remarks>
    '    Friend Function S011(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,hnmk_nnsk_kbn
    ' ,krhsk_cd
    ' ,keiki_yr_cd
    ' ,keiki_ktsk_ms
    ' ,mado_su
    ' ,ts_ktsk_cd
    ' ,dnzsk_maker_cd
    ' ,keiki_htknt_kbn
    ' ,hzk_ziry_hnbt_cd
    ' ,hzk_ziry_hts
    ' ,cktrm_d
    ' ,huing_color_kbn
    ' ,tnsdai_umu_flg
    ' ,sm_tshsk_cd
    ' ,tky_start_ymd
    ' ,hnmk_cd
    ' ,keiki_sbt_cd
    ' ,keiki_hryhn_kbn
    ' ,ryohu_kyki_yysu
    ' ,yosyrhn_cd
    ' ,yuko_kigen_ymd
    ' ,mof_kbn
    ' ,sm_taiko_kbn
    ' ,sm_knthk_cd
    ' ,ts_kino_umu_flg
    ' ,khk_kino_umu_flg
    ' ,gbdg_umu_flg

    '  FROM MST_HNMK 
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
#End Region

#Region "(SQL)計器型式マスタ　検索"
    '    ''' <summary>
    '    ''' 計器型式マスタ　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>計器型式マスタを検索します</remarks>
    '    Friend Function S012(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,kotea_kbn
    ' ,ktsk_cd
    ' ,ksyu_cd
    ' ,krhsk_cd
    ' ,ts_kotea_kbn
    ' ,keiki_sbt_cd
    ' ,keiki_ktsk_ms
    ' ,nozok_keiki_reuse_hnti_cd
    ' ,tk_keiki_reuse_hnti_cd
    ' ,nozok_ts_trtk_krksbt_cd
    ' ,tk_ts_trtk_krksbt_cd
    '   FROM MST_KEIKI_KTSK 
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
#End Region

#Region "(SQL)マスタ_町字　検索"
    '    ''' <summary>
    '    ''' マスタ_町字　検索
    '    ''' </summary>
    '    ''' <param name="dbr">DB Reader</param>
    '    ''' <remarks>マスタ_町字を検索します</remarks>
    '    Friend Function S013(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

    '        '*************************************************************************
    '        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
    '        '※このブロックコメントは削除して下さい。
    '        '
    '        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
    '        '*************************************************************************

    '        dbr.Close()

    '        Dim sb As New StringBuilder                     'StringBuilder
    '        Dim sqlstr As String = ""                       'SQL String

    '        '対象：中国５県＋香川＋愛媛＋兵庫

    '        Try
    '            'SQL文字列編集
    '            '*************************************************************************
    '            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
    '            '持出ファイル状態コード='1'(持出指示)を検索する。
    '            sb.Append(<sql><![CDATA[
    'select
    '  saksi_date
    ' ,upd_date
    ' ,sousa_user_id
    ' ,sousa_appli_cd
    ' ,tdhkn_add_cd
    ' ,siku_add_cd
    ' ,oazat_add_cd
    ' ,azatm_add_cd
    ' ,tdhkn_mskn
    ' ,siku_mskn
    ' ,oazat_mskn
    ' ,azatm_mskn
    ' ,tdhkn_ms
    ' ,siku_ms
    ' ,oazat_ms
    ' ,azatm_ms
    ' ,add_zipcd
    ' ,new_add_cd
    ' ,sikou_ym
    ' ,haisi_ym
    ' ,new_add_cd_ym
    ' ,name_henko_ym
    ' ,zipcd_henko_ym
    ' ,tino_henko_ym
    ' ,barcd_naiyo
    ' ,oyako_knk_naiyo
    ' ,cs_barcd_henko_ym
    ' ,oyako_knk_henko_ym
    ' ,tusyo_flg
    '  FROM MST_TYAZA 
    '  WHERE 1=1
    '  AND tdhkn_add_cd IN ('28','31','32' ,'33' ,'34' ,'35' ,'37' ,'38')
    '  AND haisi_ym = ' '
    '  order by tdhkn_add_cd,siku_add_cd,oazat_add_cd,azatm_add_cd
    '        ]]></sql>.Value)

    '            sqlstr = String.Format(sb.ToString)
    '            '*************************************************************************

    '            'SQLログ出力
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

    '            'SQL実行
    '            If Not dbr.ExecSelect(Me.myDB, sqlstr) Then
    '                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
    '                Throw New Exception(dbr.ErrDescription)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}，内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
    '            Return False
    '        Finally
    '        End Try
    '    End Function
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
