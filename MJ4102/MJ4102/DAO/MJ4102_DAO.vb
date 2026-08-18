''' <summary>
''' アップロード完了監視データベース処理
''' </summary>
Friend Class MJ4102_DAO

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
    ''' <param name="pBat">アップロード完了監視バッチ</param>
    Public Sub New(ByRef pBat As MJ4102)

        Me.myDB = pBat.myDB
        Me.myUser = pBat.myUser
        Me.myPath = pBat.myPath
        Me.Kino_Cd = pBat.KinoCD
        Me.SyoriYMD = pBat.SyoriYMD

    End Sub

#End Region


#Region "マスタ_事業所テーブルより，事業所情報を取得する。"
    ''' <summary>
    ''' マスタ_事業所テーブルより，事業所情報を取得する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S001(ByRef dbr As HdPostgre.HdPostgreReader) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
SELECT
     ZGSYO_CD
    ,dig4_zgsyo_cd
FROM
    COM_ZGSYO_MST
WHERE 1 = 1
    AND ZGSYO_HYZ_ZYUN <> '0'
ORDER BY
    ZGSYO_HYZ_ZYUN
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


#Region "(SQL)取替票（低圧）更新"
    ''' <summary>
    ''' 取替票（低圧）更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pNextProc">行程管理コード</param>
    ''' <param name="pSyoriYMD">処理年月日</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）を更新します</remarks>
    Friend Function U001(ByVal pRow As DataRow, ByVal pNextProc As String, ByVal pSyoriYMD As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev040 2024/12/25 物理分割 託送切替対応 ADD ttkk_tnsb_seizo_ym ～ ttkk_tnsb_szsya_mng_value
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,prcmg_cd={3}      --行程管理コード
 ,skosya_cd={4}      --施工者コード
 ,skosya_ms={5}      --施工者名
 ,syun_ymd={6}      --竣工年月日
 ,seko_kka_tenso_ymd={7}      --施工結果転送年月日
 ,tkkk_zytr_szsu={8}      --撤去計器_指示数_順潮流
 ,tkkk_gytr_szsu={9}      --撤去計器_指示数_逆潮流
 ,tkkk_szsu_sytk_zumi_flg={10}      --撤去計器_指示数_取得済フラグ
 ,ttkk_keiki_id={11}      --取付計器_計器ＩＤ
 ,ttkk_kesbt_kbn={12}      --取付計器_計器種別区分
 ,ttkk_krhsk_cd={13}      --取付計器_計量方式コード
 ,ttkk_keiki_yr={14}      --取付計器_計器容量
 ,ttkk_tshsk_cd={15}      --取付計器_通信方式コード
 ,ttkk_keiki_ktsk_ms={16}      --取付計器_計器型式名称
 ,ttkk_keiki_ktsk_cd={17}      --取付計器_計器型式コード
 ,ttkk_taiko_kbn={18}      --取付計器_耐候区分
 ,ttkk_knthk_cd={19}      --取付計器_検定方向コード
 ,ttkk_ts_kino_umu_flg={20}      --取付計器_タイムスイッチ機能有無フラグ
 ,ttkk_khk_kino_umu_flg={21}      --取付計器_開閉器機能有無フラグ
 ,ttkk_gaibu_output_umu_flg={22}      --取付計器_外部出力有無フラグ
 ,ttkk_yuko_kigen_ym={23}      --取付計器_有効期限
 ,ttkk_seizo_yy={24}      --取付計器_製造年
 ,ttkk_zyrt={25}      --取付計器_乗率
 ,ttkk_digsu={26}      --取付計器_桁数
 ,ttkk_tnsb_szsya_cd={27}      --取付計器_端子部_製造者コード
 ,ttkk_tnsb_seizo_yy={28}      --取付計器_端子部_製造年
 ,ttkk_zytr_szsu={29}      --取付計器_指示数_順潮流
 ,ttkk_gytr_szsu={30}      --取付計器_指示数_逆潮流
 ,no1_ttkk_hnsik_seizo_no={31}      --取付計器_変成器_製造ＮＯ１
 ,no2_ttkk_hnsik_seizo_no={32}      --取付計器_変成器_製造ＮＯ２
 ,ttkk_hnsik_yuko_kigen_ym={33}      --取付計器_変成器_有効期限
 ,ttkk_hnsik_gknti_no={34}      --取付計器_変成器_原検定番号
 ,ttkk_mado_su={35}      --取付計器_窓数
 ,tdnstr_ts_dosa_kbn={36}      --通電開始_タイムスイッチ動作区分
 ,tdnstr_tdnt_su={37}      --通電開始_通電時間数
 ,tdnstr_tdnstr_time={38}      --通電開始_通電開始時刻
 ,hkgds_kbn={39}      --複合動作_区分
 ,hkgds_start_md={40}      --複合動作_開始月日
 ,hkgds_end_md={41}      --複合動作_終了月日
 ,no1_hkgds_tdn_hms={42}      --複合動作_通電時刻０１
 ,no1_hkgds_sydan_hms={43}      --複合動作_遮断時刻０１
 ,no2_hkgds_tdn_hms={44}      --複合動作_通電時刻０２
 ,no2_hkgds_sydan_hms={45}      --複合動作_遮断時刻０２
 ,no3_hkgds_tdn_hms={46}      --複合動作_通電時刻０３
 ,no3_hkgds_sydan_hms={47}      --複合動作_遮断時刻０３
 ,no4_hkgds_tdn_hms={48}      --複合動作_通電時刻０４
 ,no4_hkgds_sydan_hms={49}      --複合動作_遮断時刻０４
 ,no5_hkgds_tdn_hms={50}      --複合動作_通電時刻０５
 ,no5_hkgds_sydan_hms={51}      --複合動作_遮断時刻０５
 ,no6_hkgds_tdn_hms={52}      --複合動作_通電時刻０６
 ,no6_hkgds_sydan_hms={53}      --複合動作_遮断時刻０６
 ,no7_hkgds_tdn_hms={54}      --複合動作_通電時刻０７
 ,no7_hkgds_sydan_hms={55}      --複合動作_遮断時刻０７
 ,no8_hkgds_tdn_hms={56}      --複合動作_通電時刻０８
 ,no8_hkgds_sydan_hms={57}      --複合動作_遮断時刻０８
 ,no9_hkgds_tdn_hms={58}      --複合動作_通電時刻０９
 ,no9_hkgds_sydan_hms={59}      --複合動作_遮断時刻０９
 ,no10_hkgds_tdn_hms={60}      --複合動作_通電時刻１０
 ,no10_hkgds_sydan_hms={61}      --複合動作_遮断時刻１０
 ,kbttd_start_date={62}      --個別通電_開始日時
 ,kbttd_end_date={63}      --個別通電_終了日時
 ,hoka_kihi_kbn_cd={64}      --その他_開閉区分コード
 ,hoka_gmn_flckr_kbn={65}      --その他_画面フリッカ区分
 ,hoka_evnt_krk_kbn={66}      --その他_イベント記録区分
 ,hoka_hoka_hyz_kbn={67}      --その他_その他表示区分
 ,hoka_hoka_hyz_value={68}      --その他_その他表示値
 ,hksgk_hksgn_kbn={69}      --負荷制限基本_負荷制限区分
 ,hksgk_hka_dnr={70}      --負荷制限基本_負荷電流
 ,hksgk_attny_time={71}      --負荷制限基本_自動投入時間
 ,hksgk_attny_kaisu={72}      --負荷制限基本_自動投入回数
 ,hksgk_attny_kaisu_clr_time={73}      --負荷制限基本_自動投入回数クリア時間
 ,hksgr_hksgn_kbn={74}      --負荷制限臨時_負荷制限区分
 ,hksgr_hka_dnr={75}      --負荷制限臨時_負荷電流
 ,hksgr_attny_time={76}      --負荷制限臨時_自動投入時間
 ,hksgr_attny_kaisu={77}      --負荷制限臨時_自動投入回数
 ,hksgr_attny_kaisu_clr_time={78}      --負荷制限臨時_自動投入回数クリア時間
 ,surge_yksi_siyo_umu_flg={79}      --サージ抑制使用有無フラグ
 ,tidn_kbn={80}      --停電区分
 ,tnsb_reuse_kbn={81}      --端子部再用区分
 ,khkmk_mg_kbn={82}      --工費項目_ＭＧ
 ,khkmk_hnsik_kbn={83}      --工費項目_変成器区分
 ,khkmk_wrms_kbn={84}      --工費項目_割増区分
 ,khkmk_mg_umu_flg={85}      --工費項目_ＭＧ有無フラグ
 ,khkmk_mg_yr={86}      --工費項目_ＭＧ容量
 ,khkmk_kozi_kbn={87}      --工費項目_工事区分
 ,khkmk_kozi_suryo={88}      --工費項目_工事数量
 ,khkmk_trtk_kozih_kngk={89}      --工費項目_取付工事費
 ,khkmk_tekyo_kozih_kngk={90}      --工費項目_撤去工事費
 ,kzkmk_hnsik_kozi_kbn={91}      --工事項目_変成器_工事区分
 ,kzkmk_hnsik_kozi_suryo={92}      --工事項目_変成器_工事数量
 ,kzkmk_hnsik_trtk_kozih_kngk={93}      --工事項目_変成器_取付工事費
 ,kzkmk_hnsik_tekyo_kozih_kngk={94}      --工事項目_変成器_撤去工事費
 ,kzkmk_mg_kozi_kbn={95}      --工事項目_MG_工事区分
 ,kzkmk_mg_kozi_suryo={96}      --工事項目_MG_工事数量
 ,khkmk_mg_trtk_kozih_kngk={97}      --工費項目_MG_取付工事費
 ,khkmk_mg_tekyo_kozih_kngk={98}      --工費項目_MG_撤去工事費
 ,tuika_kohi_trtk_kozih_kngk={99}      --追加工費_取付工事費
 ,tuika_kohi_tekyo_kozih_kngk={100}      --追加工費_撤去工事費
 ,tuika_kohi_umu_flg={101}      --追加工費有無フラグ
 ,zippi_umu_flg={102}      --実費有無フラグ
 ,no1_zippi_no={103}      --実費_ＮＯ１
 ,no1_zippi_kbn={104}      --実費_区分１
 ,no1_zippi_kngk={105}      --実費_金額１
 ,no2_zippi_no={106}      --実費_ＮＯ２
 ,no2_zippi_kbn={107}      --実費_区分２
 ,no2_zippi_kngk={108}      --実費_金額２
 ,no3_zippi_no={109}      --実費_ＮＯ３
 ,no3_zippi_kbn={110}      --実費_区分３
 ,no3_zippi_kngk={111}      --実費_金額３
 ,kozih_gokei_kngk={112}      --工事費合計
 ,rrzk_naiyo={113}      --連絡事項内容
 ,skrepo_kka_kbn={114}      --竣工報告結果区分
 ,ttkk_tnsb_seizo_ym={115}      --取付計器_端子部_製造年月
 ,ttkk_tnsb_sbt_cd={116}      --取付計器_端子部_種別コード
 ,ttkk_tnsb_ssnsk_dnat_cd={117}      --取付計器_端子部_相線式／電圧コード
 ,ttkk_tnsb_yr={118}      --取付計器_端子部_容量
 ,ttkk_tnsb_seizo_no={119}      --取付計器_端子部_製造番号 
 ,ttkk_tnsb_ktsk_ms={120}      --取付計器_端子部_型式名称
 ,ttkk_tnsb_kzskbt_kbn={121}      --取付計器_端子部_構造識別区分
 ,ttkk_tnsb_szsya_mng_value={122}      --取付計器_端子部_製造者管理値
WHERE 1 = 1
  AND  dig4_zgsyo_cd={123}      --事業所コード（4桁）
  AND  trkhy_hakko_nendo={124}      --取替票発行年度
  AND  trkhy_kbn={125}      --取替票区分
  AND  trkhy_no={126}      --取替票番号
         ]]></sql>.Value)

            '指示数　桁合わせ
            Dim wtkkk_zytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("tkkk_zytr_szsu"), ""))) '撤去計器_指示数_順潮流
            Dim wtkkk_gytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("tkkk_gytr_szsu"), ""))) '撤去計器_指示数_逆潮流
            Dim wttkk_zytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zytr_szsu"), ""))) '取付計器_指示数_順潮流
            Dim wttkk_gytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gytr_szsu"), ""))) '取付計器_指示数_逆潮流
            '撤去計器　指示数設定有とする。撤去計器から情報取得できない場合は結果送信不可。送信可==取得済
            Dim wtkkk_szsu_sytk_zumi_flg As String = MjK1.Cmn.Const.UmuFlg.FlgOn

            '検定種別コード(1:特別検定 取替対象-計器、2:提出検定 取替対象-計器・変成器)
            Dim wknti_sbt_cd As String = CStr(Me.myDB.ConvDbNull(pRow("knti_sbt_cd"), ""))

            '工事費の算出
            Dim khkmk_kozi_kbn As String = ""         '工費項目_工事区分・・・工量コード
            Dim khkmk_kozi_suryo As Decimal = 0       '工費項目_工事数量
            Dim khkmk_trtk_kozih_kngk As Decimal = 0  '工費項目_取付工事費
            Dim khkmk_tekyo_kozih_kngk As Decimal = 0 '工費項目_撤去工事費

            Dim kzkmk_hnsik_kozi_kbn As String = ""         '工事項目_変成器_工事区分・・・工量コード
            Dim kzkmk_hnsik_kozi_suryo As Decimal = 0       '工事項目_変成器_工事数量
            Dim kzkmk_hnsik_trtk_kozih_kngk As Decimal = 0  '工事項目_変成器_取付工事費
            Dim kzkmk_hnsik_tekyo_kozih_kngk As Decimal = 0 '工事項目_変成器_撤去工事費

            Dim kzkmk_mg_kozi_kbn As String = ""         '工事項目_MG_工事区分・・・工量コード
            Dim kzkmk_mg_kozi_suryo As Decimal = 0       '工事項目_MG_工事数量
            Dim khkmk_mg_trtk_kozih_kngk As Decimal = 0  '工費項目_MG_取付工事費
            Dim khkmk_mg_tekyo_kozih_kngk As Decimal = 0 '工費項目_MG_撤去工事費

            '各項目の値を取得
            Dim wtrke_sbt_cd As String = CStr(Me.myDB.ConvDbNull(pRow("trke_sbt_cd"), "")) '取替種別コード 0:定期 1:不良 2:その他
            Dim wttkk_kesbt_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_kesbt_kbn"), "")) '取付計器_計器種別区分 SM/SMTS/CTSM/CTSMTS
            Dim wttkk_krhsk_cd As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_krhsk_cd"), "")) '取付計器_計量方式コード
            Dim wsyun_ymd As String = CStr(Me.myDB.ConvDbNull(pRow("syun_ymd"), "")) '竣工年月日
            Dim wkhkmk_wrms_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_wrms_kbn"), "")) '工費項目_割増区分

            Dim wkhkmk_mg_umu_flg As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_umu_flg"), "")) '工費項目_ＭＧ有無フラグ
            If wkhkmk_mg_umu_flg.Equals(MjK1.Cmn.Const.KohikomkMGUmuFlg.None) Or String.IsNullOrEmpty(wkhkmk_mg_umu_flg) Then
                wkhkmk_mg_umu_flg = MjK1.Cmn.Const.KohikomkMGUmuFlg.Nashi
            End If

            '割増率取得
            Dim wrms_ritu As String = getKoryoWrmsRitu(wkhkmk_wrms_kbn, wsyun_ymd)


            '工費項目_計器
            Dim wkhkmk_keiki_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_keiki_kbn"), ""))

            '工費項目_変成器
            Dim wkhkmk_hnsik_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_hnsik_kbn"), ""))
            '未設定の場合でもここで設定することはない 2023/4/6
            '(ただし、更新項目になってしまっているので、現在値を取得し、その値で更新している)
            'If String.IsNullOrEmpty(wkhkmk_hnsik_kbn.Trim) Then
            '    '未設定の場合、ここで設定する。（庫出時に設定するよう変更した。2023/3/17）
            '    If wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.SM) Or wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.SMTS) Then
            '        wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.None  'なし
            '    ElseIf wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSM) Or wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSMTS) Then
            '        If wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T2_100V) Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T2_200V) Then
            '            wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.Tan2   '単２
            '        ElseIf wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T3_100V) Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.S3_200V) _
            '                                                                 Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.S3_400V) Then
            '            wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.Tan3   '単３
            '        End If
            '    End If
            'End If

            '工費項目_MG
            Dim wkhkmk_mg_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_kbn"), ""))

            '/* 工量コードはパラメータコードマスタに定義があり、金額計算内の単価取得共通メソッド内で取得する   2023/4/14 */
            '/* しかし、それでは取替票を更新できないため、やはり取得する。４桁で取得するので元の処理を利用する 2023/4/15 */
            '工量コード取得(取付計器ベース)
            khkmk_kozi_kbn = GetKoryoCd(wtrke_sbt_cd, wttkk_kesbt_kbn, wttkk_krhsk_cd)
            '変成器　工量コード取得(取付計器ベース)
            kzkmk_hnsik_kozi_kbn = GetKoryoCdHnsik(wtrke_sbt_cd, wknti_sbt_cd, wttkk_kesbt_kbn, wttkk_krhsk_cd)
            'MG_工量コード取得(取付計器ベース)
            kzkmk_mg_kozi_kbn = GetKoryoCdMG(wkhkmk_mg_umu_flg, wttkk_krhsk_cd)

            '数量取得
            'khkmk_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("khkmk_kozi_suryo"), "")) '工費項目_工事数量
            'kzkmk_hnsik_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("kzkmk_hnsik_kozi_suryo"), "")) '工事項目_変成器_工事数量
            'kzkmk_mg_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("kzkmk_mg_kozi_suryo"), "")) '工事項目_MG_工事数量
            khkmk_kozi_suryo = 1       'ファイルに項目がないので設定
            kzkmk_mg_kozi_suryo = 0    'ファイルに項目がない。I期はMGなし。0 設定

            '工事項目_変成器_工事数量は、CT付かつ取替対象:計器・変成器の場合のみ「1」とする。 2023/4/6
            '検定種別コードが2:提出検定の場合、取替対象は2:計器・変成器
            If wknti_sbt_cd.Equals(MjK1.Cmn.Const.KntiSbtCd.Teishutu) Then
                '変成器工事あり
                kzkmk_hnsik_kozi_suryo = 1
            End If
            '/* 変成器製造番号の有無ではない→
            ''変成器工事数量（1台または2台でも、工事数量としては「1」）
            'Dim no1_ttkk_hnsik_seizo_no As String = CStr(Me.myDB.ConvDbNull(pRow("no1_ttkk_hnsik_seizo_no"), ""))
            'Dim no2_ttkk_hnsik_seizo_no As String = CStr(Me.myDB.ConvDbNull(pRow("no2_ttkk_hnsik_seizo_no"), ""))
            'If Not String.IsNullOrEmpty(no1_ttkk_hnsik_seizo_no.Trim) Then
            '    kzkmk_hnsik_kozi_suryo = 1
            'End If
            'If Not String.IsNullOrEmpty(no2_ttkk_hnsik_seizo_no.Trim) Then
            '    kzkmk_hnsik_kozi_suryo = 1
            'End If
            '←変成器製造番号の有無ではない */

            Dim trtkTanka As String = ""
            Dim tekyoTanka As String = ""
            '工費項目_計器　単価取得
            getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_KEIKI_KBN, wkhkmk_keiki_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
            '取付工事費　金額算出
            khkmk_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, khkmk_kozi_suryo, wrms_ritu)
            '撤去工事費　金額算出
            khkmk_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, khkmk_kozi_suryo, wrms_ritu)

            '工費項目_変成器　単価取得
            getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_HNSK_KBN, wkhkmk_hnsik_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
            '変成器_取付工事費　金額算出
            kzkmk_hnsik_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, kzkmk_hnsik_kozi_suryo, wrms_ritu)
            '変成器_撤去工事費　金額算出
            kzkmk_hnsik_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, kzkmk_hnsik_kozi_suryo, wrms_ritu)

            If Not String.IsNullOrEmpty(kzkmk_mg_kozi_kbn.Trim) Then
                '工費項目_MG　単価取得
                getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_MG_KBN, wkhkmk_mg_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
                'MG_取付工事費　金額算出
                khkmk_mg_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, kzkmk_mg_kozi_suryo, wrms_ritu)
                'MG_撤去工事費　金額算出
                khkmk_mg_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, kzkmk_mg_kozi_suryo, wrms_ritu)
            End If

            '工費項目_MG　「0 固定」
            Dim khkmk_mg_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_kbn"), ""))
            khkmk_mg_kbn = "0"

            '直営工事の場合、金額に０円を設定する。
            Dim wChokueiKztnCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, pSyoriYMD)
            If CStr(Me.myDB.ConvDbNull(pRow("kozitn_cd"), "")).Trim.Equals(wChokueiKztnCd) Then
                '０円設定
                '工費項目_取付工事費 
                khkmk_trtk_kozih_kngk = 0
                '工費項目_撤去工事費 
                khkmk_tekyo_kozih_kngk = 0
                '工事項目_変成器_取付工事費 
                kzkmk_hnsik_trtk_kozih_kngk = 0
                '工事項目_変成器_撤去工事費 
                kzkmk_hnsik_tekyo_kozih_kngk = 0

                kzkmk_mg_kozi_suryo = 0       '工事項目_MG_工事数量
                khkmk_mg_trtk_kozih_kngk = 0  '工費項目_MG_取付工事費
                khkmk_mg_tekyo_kozih_kngk = 0 '工費項目_MG_撤去工事費

                'MG_取付工事費　金額算出
                khkmk_mg_trtk_kozih_kngk = 0
                'MG_撤去工事費　金額算出
                khkmk_mg_tekyo_kozih_kngk = 0

                '追加工費_取付工事費 
                '追加工費_撤去工事費 
                '工事費合計         
                '→追加工事費の処理時に対応
            End If


            '取付計器　窓数　調整（撤去計器は取替票作成時、庫出計器は庫出時に設定済み）
            Dim wttkk_mado_su = CStr(Me.myDB.ConvDbNull(pRow("ttkk_mado_su"), "")).Trim
            '値設定済みでも「1 」を設定する。
            wttkk_mado_su = CStr(MjK1.Cmn.Const.C_MADO_SU)


            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pNextProc) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("syun_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(pSyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_zytr_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_gytr_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_szsu_sytk_zumi_flg) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_kesbt_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_krhsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tshsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ktsk_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_taiko_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_knthk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_ts_kino_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_khk_kino_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gaibu_output_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_yuko_kigen_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zyrt"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_digsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_szsya_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zytr_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gytr_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_ttkk_hnsik_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_ttkk_hnsik_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_yuko_kigen_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_gknti_no"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wttkk_mado_su) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_ts_dosa_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_tdnt_su"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_tdnstr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_start_md"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_end_md"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no6_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no6_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no7_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no7_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no8_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no8_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no9_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no9_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no10_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no10_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kbttd_start_date"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kbttd_end_date"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_kihi_kbn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_gmn_flckr_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_evnt_krk_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_hoka_hyz_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_hoka_hyz_value"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_hksgn_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_hka_dnr"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_time"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_kaisu"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_kaisu_clr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_hksgn_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_hka_dnr"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_time"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_kaisu"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_kaisu_clr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("surge_yksi_siyo_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tidn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tnsb_reuse_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(khkmk_mg_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(wkhkmk_hnsik_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_wrms_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(khkmk_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_tekyo_kozih_kngk)) _
                               , EcOrgIO.EcOrgString.GetSqlText(kzkmk_hnsik_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_tekyo_kozih_kngk)) _
                               , EcOrgIO.EcOrgString.GetSqlText(kzkmk_mg_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_mg_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_mg_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_mg_tekyo_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.UmuFlg.FlgOff) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("zippi_umu_flg"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("rrzk_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.SkRepoHokkKka.Normal) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_sbt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_ssnsk_dnat_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_ktsk_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_kzskbt_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_szsya_mng_value"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************


            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

    'Rev002-Start
#Region "(SQL)取替票（高圧他）更新"
    ''' <summary>
    ''' 取替票（高圧他）更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pNextProc">行程管理コード</param>
    ''' <param name="pSyoriYMD">処理年月日</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）を更新します</remarks>
    Friend Function U012(ByVal pRow As DataRow, ByVal pNextProc As String, ByVal pSyoriYMD As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD ttkk_tnsb_seizo_ym ～ ttkk_tnsb_szsya_mng_value
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET         -- 取替票（高圧他）
  upd_date={0}                      --更新日時
 ,sousa_user_id={1}                 --更新ユーザ
 ,sousa_appli_cd={2}                --更新プログラム
 ,prcmg_cd={3}                      --行程管理コード
 ,skosya_cd={4}                     --施工者コード
 ,skosya_ms={5}                     --施工者名
 ,syun_ymd={6}                      --竣工年月日
 ,seko_kka_tenso_ymd={7}            --施工結果転送年月日
 ,tkkk_zytr_szsu={8}                --撤去計器_指示数_順潮流
 ,tkkk_gytr_szsu={9}                --撤去計器_指示数_逆潮流
 ,tkkk_sougo_szsu={10}              --撤去計器_指示数_総合
 ,tkkk_max_zyyo_szsu={11}           --撤去計器_指示数_最大需要
 ,tkkk_rsyk_szsu={12}               --撤去計器_指示数_力測有効
 ,tkkk_rsmk_szsu={13}               --撤去計器_指示数_力測無効
 ,tkkk_szsu_sytk_zumi_flg={14}      --撤去計器_指示数_取得済フラグ
 ,ttkk_keiki_id={15}                --取付計器_計器ＩＤ
 ,ttkk_kesbt_kbn={16}               --取付計器_計器種別区分
 ,ttkk_krhsk_cd={17}                --取付計器_計量方式コード
 ,ttkk_keiki_ssnsk_so_cd={18}       --取付計器_相線式_相コード
 ,ttkk_keiki_ssnsk_line_cd={19}     --取付計器_相線式_線コード
 ,ttkk_kiry_dnat_cd={20}            --取付計器_供給計量電圧コード
 ,ttkk_keiki_yr={21}                --取付計器_計器容量
 ,ttkk_tshsk_cd={22}                --取付計器_通信方式コード
 ,ttkk_keiki_ktsk_ms={23}           --取付計器_計器型式名称
 ,ttkk_keiki_ktsk_cd={24}           --取付計器_計器型式コード
 ,ttkk_ksyu_cd={25}                 --取付計器_器種コード
 ,ttkk_taiko_kbn={26}               --取付計器_耐候区分
 ,ttkk_knthk_cd={27}                --取付計器_検定方向コード
 ,ttkk_ts_kino_umu_flg={28}         --取付計器_タイムスイッチ機能有無フラグ
 ,ttkk_khk_kino_umu_flg={29}        --取付計器_開閉器機能有無フラグ
 ,ttkk_gaibu_output_umu_flg={30}    --取付計器_外部出力有無フラグ
 ,ttkk_yuko_kigen_ym={31}           --取付計器_有効期限
 ,ttkk_seizo_yy={32}                --取付計器_製造年
 ,ttkk_zyrt={33}                    --取付計器_乗率
 ,ttkk_digsu={34}                   --取付計器_桁数
 ,ttkk_tnsb_szsya_cd={35}           --取付計器_端子部_製造者コード
 ,ttkk_tnsb_seizo_yy={36}           --取付計器_端子部_製造年
 ,ttkk_zytr_szsu={37}               --取付計器_指示数_順潮流
 ,ttkk_gytr_szsu={38}               --取付計器_指示数_逆潮流
 ,no1_ttkk_hnsik_seizo_no={39}      --取付計器_変成器_製造ＮＯ１
 ,no2_ttkk_hnsik_seizo_no={40}      --取付計器_変成器_製造ＮＯ２
 ,ttkk_hnsik_VCTCT_ktsk_cd={41}     --取付計器_変成器_ＶＣＴ・ＣＴ１_型式コード
 ,ttkk_hnsik_VCTCT_seizo_no={42}    --取付計器_変成器_ＶＣＴ・ＣＴ１_製造番号
 ,ttkk_hnsik_VCTCT_seizo_yy={43}    --取付計器_変成器_ＶＣＴ・ＣＴ１_製造年
 ,ttkk_hnsik_CT2_ktsk_cd={44}       --取付計器_変成器_ＣＴ２_型式コード
 ,ttkk_hnsik_CT2_seizo_no={45}      --取付計器_変成器_ＣＴ２_製造番号
 ,ttkk_hnsik_CT2_seizo_yy={46}      --取付計器_変成器_ＣＴ２_製造年
 ,ttkk_hnsik_VT1_ktsk_cd={47}       --取付計器_変成器_ⅤＴ１_型式コード
 ,ttkk_hnsik_VT1_seizo_no={48}      --取付計器_変成器_ⅤＴ１_製造番号
 ,ttkk_hnsik_VT1_seizo_yy={49}      --取付計器_変成器_ⅤＴ１_製造年
 ,ttkk_hnsik_VT2_ktsk_cd={50}       --取付計器_変成器_ⅤＴ２_型式コード
 ,ttkk_hnsik_VT2_seizo_no={51}      --取付計器_変成器_ⅤＴ２_製造番号
 ,ttkk_hnsik_VT2_seizo_yy={52}      --取付計器_変成器_ⅤＴ２_製造年
 ,TTkk_HYZKI_seizo_no={53}          --取付計器_表示器_製造番号
 ,ttkk_hnsik_yuko_kigen_ym={54}     --取付計器_変成器_有効期限
 ,ttkk_hnsik_gknti_no={55}          --取付計器_変成器_原検定番号
 ,ttkk_mado_su={56}                 --取付計器_窓数
 ,tdnstr_ts_dosa_kbn={57}           --通電開始_タイムスイッチ動作区分
 ,tdnstr_tdnt_su={58}               --通電開始_通電時間数
 ,tdnstr_tdnstr_time={59}           --通電開始_通電開始時刻
 ,hkgds_kbn={60}                    --複合動作_区分
 ,hkgds_start_md={61}               --複合動作_開始月日
 ,hkgds_end_md={62}                 --複合動作_終了月日
 ,no1_hkgds_tdn_hms={63}            --複合動作_通電時刻０１
 ,no1_hkgds_sydan_hms={64}          --複合動作_遮断時刻０１
 ,no2_hkgds_tdn_hms={65}            --複合動作_通電時刻０２
 ,no2_hkgds_sydan_hms={66}          --複合動作_遮断時刻０２
 ,no3_hkgds_tdn_hms={67}            --複合動作_通電時刻０３
 ,no3_hkgds_sydan_hms={68}          --複合動作_遮断時刻０３
 ,no4_hkgds_tdn_hms={69}            --複合動作_通電時刻０４
 ,no4_hkgds_sydan_hms={70}          --複合動作_遮断時刻０４
 ,no5_hkgds_tdn_hms={71}            --複合動作_通電時刻０５
 ,no5_hkgds_sydan_hms={72}          --複合動作_遮断時刻０５
 ,no6_hkgds_tdn_hms={73}            --複合動作_通電時刻０６
 ,no6_hkgds_sydan_hms={74}          --複合動作_遮断時刻０６
 ,no7_hkgds_tdn_hms={75}            --複合動作_通電時刻０７
 ,no7_hkgds_sydan_hms={76}          --複合動作_遮断時刻０７
 ,no8_hkgds_tdn_hms={77}            --複合動作_通電時刻０８
 ,no8_hkgds_sydan_hms={78}          --複合動作_遮断時刻０８
 ,no9_hkgds_tdn_hms={79}            --複合動作_通電時刻０９
 ,no9_hkgds_sydan_hms={80}          --複合動作_遮断時刻０９
 ,no10_hkgds_tdn_hms={81}           --複合動作_通電時刻１０
 ,no10_hkgds_sydan_hms={82}         --複合動作_遮断時刻１０
 ,kbttd_start_date={83}             --個別通電_開始日時
 ,kbttd_end_date={84}               --個別通電_終了日時
 ,hoka_kihi_kbn_cd={85}             --その他_開閉区分コード
 ,hoka_gmn_flckr_kbn={86}           --その他_画面フリッカ区分
 ,hoka_evnt_krk_kbn={87}            --その他_イベント記録区分
 ,hoka_hoka_hyz_kbn={88}            --その他_その他表示区分
 ,hoka_hoka_hyz_value={89}          --その他_その他表示値
 ,hksgk_hksgn_kbn={90}              --負荷制限基本_負荷制限区分
 ,hksgk_hka_dnr={91}                --負荷制限基本_負荷電流
 ,hksgk_attny_time={92}             --負荷制限基本_自動投入時間
 ,hksgk_attny_kaisu={93}            --負荷制限基本_自動投入回数
 ,hksgk_attny_kaisu_clr_time={94}   --負荷制限基本_自動投入回数クリア時間
 ,hksgr_hksgn_kbn={95}              --負荷制限臨時_負荷制限区分
 ,hksgr_hka_dnr={96}                --負荷制限臨時_負荷電流
 ,hksgr_attny_time={97}             --負荷制限臨時_自動投入時間
 ,hksgr_attny_kaisu={98}            --負荷制限臨時_自動投入回数
 ,hksgr_attny_kaisu_clr_time={99}   --負荷制限臨時_自動投入回数クリア時間
 ,surge_yksi_siyo_umu_flg={100}     --サージ抑制使用有無フラグ
 ,tidn_kbn={101}                    --停電区分
 ,tidn_riyu_cd={102}                --停電理由コード
 ,sagyo_time_kbn_cd={103}           --作業時間区分コード
 ,sagyo_start_hms={104}             --作業開始時刻
 ,sagyo_end_hms={105}               --作業終了時刻
 ,dtkk_sougo_drkei_tekyo_szsu={106} --代替計器_総合電力計_撤去指示数
 ,dtkk_sougo_drkei_trtk_szsu={107}  --代替計器_総合電力計_取付指示数
 ,dtkk_max_jydrrkei_tekyo_szsu={108}--代替計器_最大需要電力計_撤去指示数
 ,dtkk_max_jydrrkei_trtk_szsu={109} --代替計器_最大需要電力計_取付指示数
 ,dtkk_rsyk_tekyo_szsu={110}        --代替計器_力測有効_撤去指示数
 ,dtkk_rsyk_trtk_szsu={111}         --代替計器_力測有効_取付指示数
 ,dtkk_rsmk_tekyo_szsu={112}        --代替計器_力測無効_撤去指示数
 ,dtkk_rsmk_trtk_szsu={113}         --代替計器_力測無効_取付指示数
 ,dtkk_zyrt={114}                   --代替計器_力測無効_乗率
 ,tnsb_reuse_kbn={115}              --端子部再用区分
 ,khkmk_mg_kbn={116}                --工費項目_ＭＧ
 ,khkmk_hnsik_kbn={117}             --工費項目_変成器区分
 ,khkmk_hnsik_st_iti_kbn={118}      --工費項目_変成器_設置_位置_区分
 ,khkmk_wrms_kbn={119}              --工費項目_割増区分
 ,khkmk_mg_umu_flg={120}            --工費項目_ＭＧ有無フラグ
 ,khkmk_mg_yr={121}                 --工費項目_ＭＧ容量
 ,khkmk_kozi_kbn={122}              --工費項目_工事区分
 ,khkmk_kozi_suryo={123}            --工費項目_工事数量
 ,khkmk_trtk_kozih_kngk={124}       --工費項目_取付工事費
 ,khkmk_tekyo_kozih_kngk={125}      --工費項目_撤去工事費
 ,kzkmk_hnsik_kozi_kbn={126}        --工事項目_変成器_工事区分
 ,kzkmk_hnsik_kozi_suryo={127}      --工事項目_変成器_工事数量
 ,kzkmk_hnsik_trtk_kozih_kngk={128} --工事項目_変成器_取付工事費
 ,kzkmk_hnsik_tekyo_kozih_kngk={129}--工事項目_変成器_撤去工事費
 ,kzkmk_mg_kozi_kbn={130}           --工事項目_MG_工事区分
 ,kzkmk_mg_kozi_suryo={131}         --工事項目_MG_工事数量
 ,khkmk_mg_trtk_kozih_kngk={132}    --工費項目_MG_取付工事費
 ,khkmk_mg_tekyo_kozih_kngk={133}   --工費項目_MG_撤去工事費
 ,tuika_kohi_trtk_kozih_kngk={134}  --追加工費_取付工事費
 ,tuika_kohi_tekyo_kozih_kngk={135} --追加工費_撤去工事費
 ,tuika_kohi_umu_flg={136}          --追加工費有無フラグ
 ,zippi_umu_flg={137}               --実費有無フラグ
 ,no1_zippi_no={138}                --実費_ＮＯ１
 ,no1_zippi_kbn={139}               --実費_区分１
 ,no1_zippi_kngk={140}              --実費_金額１
 ,no2_zippi_no={141}                --実費_ＮＯ２
 ,no2_zippi_kbn={142}               --実費_区分２
 ,no2_zippi_kngk={143}              --実費_金額２
 ,no3_zippi_no={144}                --実費_ＮＯ３
 ,no3_zippi_kbn={145}               --実費_区分３
 ,no3_zippi_kngk={146}              --実費_金額３
 ,no4_zippi_no={147}                --実費_ＮＯ４
 ,no4_zippi_kbn={148}               --実費_区分４
 ,no4_zippi_kngk={149}              --実費_金額４
 ,no5_zippi_no={150}                --実費_ＮＯ５
 ,no5_zippi_kbn={151}               --実費_区分５
 ,no5_zippi_kngk={152}              --実費_金額５
 ,kozih_gokei_kngk={153}            --工事費合計
 ,rrzk_naiyo={154}                  --連絡事項内容
 ,skrepo_kka_kbn={155}              --竣工報告結果区分
 ,ttkk_tnsb_seizo_ym={156}          --取付計器_端子部_製造年月
 ,ttkk_tnsb_sbt_cd={157}            --取付計器_端子部_種別コード
 ,ttkk_tnsb_ssnsk_dnat_cd={158}     --取付計器_端子部_相線式／電圧コード
 ,ttkk_tnsb_yr={159}                --取付計器_端子部_容量
 ,ttkk_tnsb_seizo_no={160}          --取付計器_端子部_製造番号 
 ,ttkk_tnsb_ktsk_ms={161}           --取付計器_端子部_型式名称
 ,ttkk_tnsb_kzskbt_kbn={162}        --取付計器_端子部_構造識別区分
 ,ttkk_tnsb_szsya_mng_value={163}   --取付計器_端子部_製造者管理値
 WHERE 1 = 1
  AND  dig4_zgsyo_cd={164}          --事業所コード（4桁）
  AND  trkhy_hakko_nendo={165}      --取替票発行年度
  AND  trkhy_kbn={166}              --取替票区分
  AND  trkhy_no={167}               --取替票番号
         ]]></sql>.Value)

            '指示数　桁合わせ
            Dim wtkkk_zytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("tkkk_zytr_szsu"), ""))) '撤去計器_指示数_順潮流
            Dim wtkkk_gytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("tkkk_gytr_szsu"), ""))) '撤去計器_指示数_逆潮流
            'Rev002.1 MOD Start 20241112 指示数の桁変更
            Dim wtkkk_sougo_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("tkkk_sougo_szsu"), ""))          '撤去計器_指示数_総合
            Dim wtkkk_max_zyyo_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("tkkk_max_zyyo_szsu"), ""))    '撤去計器_指示数_最大需要
            Dim wtkkk_rsyk_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("tkkk_rsyk_szsu"), ""))            '撤去計器_指示数_力測有効
            Dim wtkkk_rsmk_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("tkkk_rsmk_szsu"), ""))            '撤去計器_指示数_力測無効
            'Rev002.1 MOD End 20241112

            Dim wttkk_zytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zytr_szsu"), ""))) '取付計器_指示数_順潮流
            Dim wttkk_gytr_szsu As String = getSzsuStr(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gytr_szsu"), ""))) '取付計器_指示数_逆潮流
            'Rev002.1 MOD Start 20241112 指示数の桁変更
            Dim wttkk_sougo_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_sougo_szsu"), ""))          '取付計器_指示数_総合
            Dim wttkk_rsyk_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_rsyk_szsu"), ""))            '取付計器_指示数_力測有効
            Dim wttkk_rsmk_szsu As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_rsmk_szsu"), ""))            '取付計器_指示数_力測無効
            'Rev002.1 MOD End 20241112

            '撤去計器　指示数設定有とする。撤去計器から情報取得できない場合は結果送信不可。送信可==取得済
            Dim wtkkk_szsu_sytk_zumi_flg As String = MjK1.Cmn.Const.UmuFlg.FlgOn

            '検定種別コード(1:特別検定 取替対象-計器、2:提出検定 取替対象-計器・変成器)
            Dim wknti_sbt_cd As String = CStr(Me.myDB.ConvDbNull(pRow("knti_sbt_cd"), ""))

            '工事費の算出
            Dim khkmk_kozi_kbn As String = ""         '工費項目_工事区分・・・工量コード
            Dim khkmk_kozi_suryo As Decimal = 0       '工費項目_工事数量
            Dim khkmk_trtk_kozih_kngk As Decimal = 0  '工費項目_取付工事費
            Dim khkmk_tekyo_kozih_kngk As Decimal = 0 '工費項目_撤去工事費

            Dim kzkmk_hnsik_kozi_kbn As String = ""         '工事項目_変成器_工事区分・・・工量コード
            Dim kzkmk_hnsik_kozi_suryo As Decimal = 0       '工事項目_変成器_工事数量
            Dim kzkmk_hnsik_trtk_kozih_kngk As Decimal = 0  '工事項目_変成器_取付工事費
            Dim kzkmk_hnsik_tekyo_kozih_kngk As Decimal = 0 '工事項目_変成器_撤去工事費

            Dim kzkmk_mg_kozi_kbn As String = ""         '工事項目_MG_工事区分・・・工量コード
            Dim kzkmk_mg_kozi_suryo As Decimal = 0       '工事項目_MG_工事数量
            Dim khkmk_mg_trtk_kozih_kngk As Decimal = 0  '工費項目_MG_取付工事費
            Dim khkmk_mg_tekyo_kozih_kngk As Decimal = 0 '工費項目_MG_撤去工事費

            '各項目の値を取得
            Dim wtrke_sbt_cd As String = CStr(Me.myDB.ConvDbNull(pRow("trke_sbt_cd"), "")) '取替種別コード 0:定期 1:不良 2:その他
            Dim wttkk_kesbt_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_kesbt_kbn"), "")) '取付計器_計器種別区分 SM/SMTS/CTSM/CTSMTS
            Dim wtrke_tis_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("trke_tis_kbn"), ""))    '取替対象区分


            Dim wttkk_ssnsk_so_cd As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ssnsk_so_cd"), ""))      '取付計器_相線式_相コード
            Dim wttkk_ssnsk_line_cd As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ssnsk_line_cd"), ""))  '取付計器_相線式_線コード
            Dim wttkk_kiry_dnat_cd As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_kiry_dnat_cd"), ""))              '取付計器_計量電圧コード

            '計器が低圧の場合は格納される
            Dim wttkk_krhsk_cd As String = CStr(Me.myDB.ConvDbNull(pRow("ttkk_krhsk_cd"), ""))                  '低圧計器用計量方式コード

            '取付計器_相線式_相コードが空白でない場合は高圧計器
            If Not String.IsNullOrWhiteSpace(wttkk_ssnsk_so_cd) Then
                'MG_工量コード取得を得るために高圧用の１バイトの計量方式コードをセットする
                wttkk_krhsk_cd = MjK1.Cmn.Func.ConvKkKrhskCdSoSenDnat(wttkk_ssnsk_so_cd, wttkk_ssnsk_line_cd, wttkk_kiry_dnat_cd)
            End If

            Dim wsyun_ymd As String = CStr(Me.myDB.ConvDbNull(pRow("syun_ymd"), "")) '竣工年月日
            Dim wkhkmk_wrms_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_wrms_kbn"), "")) '工費項目_割増区分

            Dim wkhkmk_mg_umu_flg As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_umu_flg"), "")) '工費項目_ＭＧ有無フラグ
            If wkhkmk_mg_umu_flg.Equals(MjK1.Cmn.Const.KohikomkMGUmuFlg.None) Or String.IsNullOrEmpty(wkhkmk_mg_umu_flg) Then
                wkhkmk_mg_umu_flg = MjK1.Cmn.Const.KohikomkMGUmuFlg.Nashi
            End If

            '割増率取得
            Dim wrms_ritu As String = getKoryoWrmsRitu(wkhkmk_wrms_kbn, wsyun_ymd)


            '工費項目_計器
            Dim wkhkmk_keiki_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_keiki_kbn"), ""))

            '工費項目_変成器
            Dim wkhkmk_hnsik_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_hnsik_kbn"), ""))
            '未設定の場合でもここで設定することはない 2023/4/6
            '(ただし、更新項目になってしまっているので、現在値を取得し、その値で更新している)
            'If String.IsNullOrEmpty(wkhkmk_hnsik_kbn.Trim) Then
            '    '未設定の場合、ここで設定する。（庫出時に設定するよう変更した。2023/3/17）
            '    If wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.SM) Or wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.SMTS) Then
            '        wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.None  'なし
            '    ElseIf wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSM) Or wttkk_kesbt_kbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSMTS) Then
            '        If wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T2_100V) Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T2_200V) Then
            '            wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.Tan2   '単２
            '        ElseIf wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.T3_100V) Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.S3_200V) _
            '                                                                 Or wttkk_krhsk_cd.Equals(MjK1.Cmn.Const.KrHskCd.S3_400V) Then
            '            wkhkmk_hnsik_kbn = MjK1.Cmn.Const.KohikomkHnskKbn.Tan3   '単３
            '        End If
            '    End If
            'End If

            '工費項目_変成器_設置_位置
            Dim wkhkmk_hnsik_st_iti_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_hnsik_st_iti_kbn"), ""))

            '工費項目_MG
            Dim wkhkmk_mg_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_kbn"), ""))

            Dim wtkkk_krhsk_cd As String = CStr(Me.myDB.ConvDbNull(pRow("tkkk_krhsk_cd"), ""))
            '相線式コードが" "でない場合、高圧
            If Not String.IsNullOrWhiteSpace(wttkk_ssnsk_so_cd) Then
                '高圧計器
                '工量コード取得(取付計器ベース)（高圧用）
                khkmk_kozi_kbn = MjK1.Cmn.Const.KoryoCd.KkHnsikKeikiKbn

                '変成器　工量コード取得(取付計器ベース)　（高圧用）
                kzkmk_hnsik_kozi_kbn = GetKoryoCdHnsikKoat(wkhkmk_hnsik_st_iti_kbn)

                'MG_工量コード取得(取付計器ベース)（高圧用）
                kzkmk_mg_kozi_kbn = GetKoryoCdMGKoat(wkhkmk_mg_umu_flg, wttkk_krhsk_cd)

            Else
                '低圧計器
                '/* 工量コードはパラメータコードマスタに定義があり、金額計算内の単価取得共通メソッド内で取得する   2023/4/14 */
                '/* しかし、それでは取替票を更新できないため、やはり取得する。４桁で取得するので元の処理を利用する 2023/4/15 */
                khkmk_kozi_kbn = GetKoryoCd(wtrke_sbt_cd, wttkk_kesbt_kbn, wtkkk_krhsk_cd)

                '変成器　工量コード取得(取付計器ベース)
                kzkmk_hnsik_kozi_kbn = GetKoryoCdHnsik(wtrke_sbt_cd, wknti_sbt_cd, wttkk_kesbt_kbn, wttkk_krhsk_cd)

                'MG_工量コード取得(取付計器ベース)
                kzkmk_mg_kozi_kbn = GetKoryoCdMG(wkhkmk_mg_umu_flg, wttkk_krhsk_cd)

            End If


            '数量取得
            'khkmk_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("khkmk_kozi_suryo"), "")) '工費項目_工事数量
            'kzkmk_hnsik_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("kzkmk_hnsik_kozi_suryo"), "")) '工事項目_変成器_工事数量
            'kzkmk_mg_kozi_suryo = CDec(Me.myDB.ConvDbNull(pRow("kzkmk_mg_kozi_suryo"), "")) '工事項目_MG_工事数量
            khkmk_kozi_suryo = 1       'ファイルに項目がないので設定
            kzkmk_mg_kozi_suryo = 0    'ファイルに項目がない。I期はMGなし。0 設定

            '工事項目_変成器_工事数量は、CT付かつ取替対象:計器・変成器の場合のみ「1」とする。 2023/4/6
            '取替対象区分が、2:計器・変成器の場合、変成器工事あり
            If wtrke_tis_kbn.Equals(MjK1.Cmn.Const.TrkeTisKbn.KeikiHnsik) Then
                '変成器工事あり
                kzkmk_hnsik_kozi_suryo = 1
            End If
            '/* 変成器製造番号の有無ではない→
            ''変成器工事数量（1台または2台でも、工事数量としては「1」）
            'Dim no1_ttkk_hnsik_seizo_no As String = CStr(Me.myDB.ConvDbNull(pRow("no1_ttkk_hnsik_seizo_no"), ""))
            'Dim no2_ttkk_hnsik_seizo_no As String = CStr(Me.myDB.ConvDbNull(pRow("no2_ttkk_hnsik_seizo_no"), ""))
            'If Not String.IsNullOrEmpty(no1_ttkk_hnsik_seizo_no.Trim) Then
            '    kzkmk_hnsik_kozi_suryo = 1
            'End If
            'If Not String.IsNullOrEmpty(no2_ttkk_hnsik_seizo_no.Trim) Then
            '    kzkmk_hnsik_kozi_suryo = 1
            'End If
            '←変成器製造番号の有無ではない */

            Dim trtkTanka As String = ""
            Dim tekyoTanka As String = ""
            '工費項目_計器　単価取得
            getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_KEIKI_KBN, wkhkmk_keiki_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
            '取付工事費　金額算出
            khkmk_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, khkmk_kozi_suryo, wrms_ritu)
            '撤去工事費　金額算出
            khkmk_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, khkmk_kozi_suryo, wrms_ritu)

            '工費項目_変成器　単価取得  2024/08/23 修正
            If String.IsNullOrWhiteSpace(wtkkk_krhsk_cd) Then
                '高圧計器
                getTankaH(MjK1.Cmn.Const.PrmId.HNSIK_STIT_KBN, wkhkmk_hnsik_st_iti_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
            Else
                '高圧計器
                getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_HNSK_KBN, wkhkmk_hnsik_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
            End If
            '変成器_取付工事費　金額算出
            kzkmk_hnsik_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, kzkmk_hnsik_kozi_suryo, wrms_ritu)
            '変成器_撤去工事費　金額算出
            kzkmk_hnsik_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, kzkmk_hnsik_kozi_suryo, wrms_ritu)

            If Not String.IsNullOrEmpty(kzkmk_mg_kozi_kbn.Trim) Then
                '工費項目_MG　単価取得
                getTankaH(MjK1.Cmn.Const.PrmId.KOHIKOMK_MG_KBN, wkhkmk_mg_kbn, wsyun_ymd, trtkTanka, tekyoTanka)
                'MG_取付工事費　金額算出
                khkmk_mg_trtk_kozih_kngk = CalcKozihKngk(trtkTanka, kzkmk_mg_kozi_suryo, wrms_ritu)
                'MG_撤去工事費　金額算出
                khkmk_mg_tekyo_kozih_kngk = CalcKozihKngk(tekyoTanka, kzkmk_mg_kozi_suryo, wrms_ritu)
            End If

            '工費項目_MG　「0 固定」
            Dim khkmk_mg_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_kbn"), ""))
            khkmk_mg_kbn = "0"

            '直営工事の場合、金額に０円を設定する。
            Dim wChokueiKztnCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, pSyoriYMD)
            If CStr(Me.myDB.ConvDbNull(pRow("kozitn_cd"), "")).Trim.Equals(wChokueiKztnCd) Then
                '０円設定
                '工費項目_取付工事費 
                khkmk_trtk_kozih_kngk = 0
                '工費項目_撤去工事費 
                khkmk_tekyo_kozih_kngk = 0
                '工事項目_変成器_取付工事費 
                kzkmk_hnsik_trtk_kozih_kngk = 0
                '工事項目_変成器_撤去工事費 
                kzkmk_hnsik_tekyo_kozih_kngk = 0

                kzkmk_mg_kozi_suryo = 0       '工事項目_MG_工事数量
                khkmk_mg_trtk_kozih_kngk = 0  '工費項目_MG_取付工事費
                khkmk_mg_tekyo_kozih_kngk = 0 '工費項目_MG_撤去工事費

                'MG_取付工事費　金額算出
                khkmk_mg_trtk_kozih_kngk = 0
                'MG_撤去工事費　金額算出
                khkmk_mg_tekyo_kozih_kngk = 0

                '追加工費_取付工事費 
                '追加工費_撤去工事費 
                '工事費合計         
                '→追加工事費の処理時に対応
            End If


            '取付計器　窓数　調整（撤去計器は取替票作成時、庫出計器は庫出時に設定済み）
            Dim wttkk_mado_su = CStr(Me.myDB.ConvDbNull(pRow("ttkk_mado_su"), "")).Trim
            '値設定済みでも「1 」を設定する。
            wttkk_mado_su = CStr(MjK1.Cmn.Const.C_MADO_SU)


            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pNextProc) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("syun_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(pSyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_zytr_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_gytr_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_sougo_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_max_zyyo_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_rsyk_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_rsmk_szsu) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtkkk_szsu_sytk_zumi_flg) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_kesbt_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_krhsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ssnsk_so_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ssnsk_line_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_kiry_dnat_cd"), "")), EcOrgIO.SearchType.normal, False, EcOrgIO.TrimType.non) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tshsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ktsk_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_ksyu_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_taiko_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_knthk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_ts_kino_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_khk_kino_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gaibu_output_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_yuko_kigen_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zyrt"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_digsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_szsya_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_zytr_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_gytr_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_ttkk_hnsik_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_ttkk_hnsik_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VCTCT_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VCTCT_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VCTCT_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_CT2_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_CT2_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_CT2_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT1_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT1_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT1_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT2_ktsk_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT2_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_VT2_seizo_yy"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hyzki_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_yuko_kigen_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_hnsik_gknti_no"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wttkk_mado_su) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_ts_dosa_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_tdnt_su"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tdnstr_tdnstr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_start_md"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hkgds_end_md"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no6_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no6_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no7_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no7_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no8_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no8_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no9_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no9_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no10_hkgds_tdn_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no10_hkgds_sydan_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kbttd_start_date"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kbttd_end_date"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_kihi_kbn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_gmn_flckr_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_evnt_krk_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_hoka_hyz_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hoka_hoka_hyz_value"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_hksgn_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_hka_dnr"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_time"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_kaisu"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgk_attny_kaisu_clr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_hksgn_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_hka_dnr"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_time"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_kaisu"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("hksgr_attny_kaisu_clr_time"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("surge_yksi_siyo_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tidn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tidn_riyu_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sagyo_time_kbn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sagyo_start_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sagyo_end_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_sougo_drkei_tekyo_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_sougo_drkei_trtk_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_max_jydrrkei_tekyo_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_max_jydrrkei_trtk_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_rsyk_tekyo_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_rsyk_trtk_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_rsmk_tekyo_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_rsmk_trtk_szsu"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dtkk_zyrt"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tnsb_reuse_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(khkmk_mg_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(wkhkmk_hnsik_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(wkhkmk_hnsik_st_iti_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_wrms_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_umu_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("khkmk_mg_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(khkmk_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_tekyo_kozih_kngk)) _
                               , EcOrgIO.EcOrgString.GetSqlText(kzkmk_hnsik_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_hnsik_tekyo_kozih_kngk)) _
                               , EcOrgIO.EcOrgString.GetSqlText(kzkmk_mg_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(kzkmk_mg_kozi_suryo)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_mg_trtk_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(khkmk_mg_tekyo_kozih_kngk)) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.UmuFlg.FlgOff) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("zippi_umu_flg"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no1_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no2_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no3_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no4_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no4_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no5_zippi_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_zippi_kbn"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("no5_zippi_kngk"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr("0")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("rrzk_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.SkRepoHokkKka.Normal) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_ym"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_sbt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_ssnsk_dnat_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_yr"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_seizo_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_ktsk_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_kzskbt_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ttkk_tnsb_szsya_mng_value"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************


            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region
    'Rev002-End


#Region "共通　工量コード取得(4桁)"
    ''' <summary>
    ''' 工量コード取得
    ''' </summary>
    ''' <param name="pTrkeSbtCd">取替種別コード</param>
    ''' <param name="pTtkkKesbtKbn">取付計器_計器種別区分</param>
    ''' <param name="pTtkkKrhskCd">取付計器_計量方式コード</param>
    ''' <returns>割引率</returns>
    Friend Function GetKoryoCd(ByVal pTrkeSbtCd As String, ByVal pTtkkKesbtKbn As String, ByVal pTtkkKrhskCd As String) As String
        Dim wStr As String = ""

        '取付種別コードチェック
        If pTrkeSbtCd.Equals(MjK1.Cmn.Const.TrkeSbtCd.Teiki) Then
            '定期の場合
            '計器種別区分チェック
            If pTtkkKesbtKbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSM) Or pTtkkKesbtKbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSMTS) Then
                'CT付
                wStr = MjK1.Cmn.Const.KoryoCd.KkHnsik
            Else
                'CTなし
                '計量方式チェック
                If pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_100V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_200V) Then
                    wStr = MjK1.Cmn.Const.KoryoCd.KkTns2
                ElseIf pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T3_100V) Or
                       pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_200V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_400V) Then
                    wStr = MjK1.Cmn.Const.KoryoCd.KkTns3
                End If
            End If
        ElseIf pTrkeSbtCd.Equals(MjK1.Cmn.Const.TrkeSbtCd.Sonota) Then
            'その他の場合（システム対象外）
        End If

        Return wStr
    End Function

    ''' <summary>
    ''' 工量コード取得（変成器）
    ''' </summary>
    ''' <param name="pTrkeSbtCd">取替種別コード</param>
    ''' <param name="pKntiSbtCd">検定種別コード</param>
    ''' <param name="pTtkkKesbtKbn">取付計器_計器種別区分</param>
    ''' <param name="pTtkkKrhskCd">取付計器_計量方式コード</param>
    ''' <returns>割引率</returns>
    Friend Function GetKoryoCdHnsik(ByVal pTrkeSbtCd As String, ByVal pKntiSbtCd As String, ByVal pTtkkKesbtKbn As String, ByVal pTtkkKrhskCd As String) As String
        Dim wStr As String = " "

        '取付種別コードチェック
        If pTrkeSbtCd.Equals(MjK1.Cmn.Const.TrkeSbtCd.Teiki) Then
            '定期の場合
            '検定種別コードが「2:提出検定（取替対象：計器・変成器）」の場合のみ設定
            If pKntiSbtCd.Equals(MjK1.Cmn.Const.KntiSbtCd.Teishutu) Then
                '計器種別区分チェック
                If pTtkkKesbtKbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSM) Or pTtkkKesbtKbn.Equals(MjK1.Cmn.Const.KeSbtKbn.CTSMTS) Then
                    'CT付
                    '計量方式チェック
                    If pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_100V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_200V) Then
                        wStr = MjK1.Cmn.Const.KoryoCd.HnsikTns2
                    ElseIf pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T3_100V) Or
                       pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_200V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_400V) Then
                        wStr = MjK1.Cmn.Const.KoryoCd.HnsikTns3
                    End If
                End If
            End If
        ElseIf pTrkeSbtCd.Equals(MjK1.Cmn.Const.TrkeSbtCd.Sonota) Then
            'その他の場合（システム対象外）
        End If

        Return wStr
    End Function

    ''' <summary>
    ''' 工量コード取得（MG）
    ''' </summary>
    ''' <param name="pKhkmkMgUmuFlg">ＭＧ有無フラグ</param>
    ''' <param name="pTtkkKrhskCd">取付計器_計量方式コード</param>
    ''' <returns>割引率</returns>
    Friend Function GetKoryoCdMG(ByVal pKhkmkMgUmuFlg As String, ByVal pTtkkKrhskCd As String) As String
        Dim wStr As String = ""

        'MG有無フラグチェック
        If pKhkmkMgUmuFlg.Equals(MjK1.Cmn.Const.KohikomkMGUmuFlg.Ari) Then
            '計量方式チェック
            If pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_100V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_200V) Or
               pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T3_100V) Or
               pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_200V) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_400V) Then
                wStr = MjK1.Cmn.Const.KoryoCd.KkMG
            End If
        End If

        Return wStr
    End Function

    ''' <summary>
    ''' 工量コード取得（変成器）（高圧）
    ''' </summary>
    ''' <param name="pkhkmkHnsikStItiKbn">工費項目_変成器_設置_位置</param>
    ''' <returns>工量コード</returns>
    Friend Function GetKoryoCdHnsikKoat(ByVal pkhkmkHnsikStItiKbn As String) As String
        Dim wStr As String = " "

        '計量方式チェック
        If pkhkmkHnsikStItiKbn.Equals(MjK1.Cmn.Const.HnsikStiit.indoor) Then
            wStr = MjK1.Cmn.Const.KoryoCd.KkHnsikHensikOkuni
        Else
            wStr = MjK1.Cmn.Const.KoryoCd.KkHnsikHensikHashira
        End If

        Return wStr
    End Function

    ''' <summary>
    ''' 工量コード取得（MG）（高圧）
    ''' </summary>
    ''' <param name="pKhkmkMgUmuFlg">ＭＧ有無フラグ</param>
    ''' <param name="pTtkkKrhskCd">取付計器_計量方式コード</param>
    ''' <returns>割引率</returns>
    Friend Function GetKoryoCdMGKoat(ByVal pKhkmkMgUmuFlg As String, ByVal pTtkkKrhskCd As String) As String
        Dim wStr As String = ""

        'MG有無フラグチェック
        If pKhkmkMgUmuFlg.Equals(MjK1.Cmn.Const.KohikomkMGUmuFlg.Ari) Then
            '計量方式チェック
            If pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_3kV) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.T2_6kV) Or
               pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_3kV) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_6kV) Or
               pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S4_6kV) Or pTtkkKrhskCd.Equals(MjK1.Cmn.Const.KrHskCd.S3_10kV) Then
                wStr = MjK1.Cmn.Const.KoryoCd.KkMG
            End If
        End If

        Return wStr
    End Function


#End Region




#Region "共通　割増率取得"

    ''' <summary>
    ''' マスター工量　割増率取得
    ''' </summary>
    ''' <param name="pKhkmkWrmsKbn">工費項目_割増区分</param>
    ''' <param name="pSyunYmd">竣工日</param>
    ''' <returns>割増率</returns>
    Friend Function getKoryoWrmsRitu(ByVal pKhkmkWrmsKbn As String, ByVal pSyunYmd As String) As String
        Dim wStr As String = ""
        Dim wKoryoCd As String = ""

        If pKhkmkWrmsKbn.Equals(MjK1.Cmn.Const.KohiKomkWrmsCd.Ritou) Then
            '工費項目_割増区分==離島

            wKoryoCd = MjK1.Cmn.Func.GetParamaterCdMst(myDB, MjK1.Cmn.Const.PrmId.KOHIKOMK_WRMS_CD, pKhkmkWrmsKbn)
            'wKoryoCd = [Const].C_WRMS_KORYO_CD   パラメータコードマスタの定義を使用する  2023/4/14

            Dim jouken1 As String = "koryo_cd = " & EcOrgIO.EcOrgString.GetSqlText(wKoryoCd)
            Dim jouken2 As String = EcOrgIO.EcOrgString.GetSqlText(pSyunYmd) & " BETWEEN tky_start_ymd AND tky_end_ymd "
            wStr = myDB.GetRecValueString("mst_koryo", "wrms_ritu", jouken1, jouken2)
            If String.IsNullOrEmpty(wStr.Trim) Then
                wStr = "0"
            End If
        Else
            '工費項目_割増区分==なし
            wStr = [Const].C_WRMS_NONE
        End If

        Return wStr
    End Function
#End Region

#Region "単価情報を取得する。"
    ''' <summary>
    ''' 単価の取得（本体工費）
    ''' </summary>
    ''' <param name="pPrmId">工費項目の種別</param>
    ''' <param name="pPrmCd">工費項目の値</param>
    ''' <param name="syunYmd">竣工年月日</param>
    ''' <param name="trtkTanka">取付単価</param>
    ''' <param name="tekyoTanka">撤去単価</param>
    ''' <remarks></remarks>
    Function getTankaH(ByVal pPrmId As String, ByVal pPrmCd As String,
                       ByVal syunYmd As String, ByRef trtkTanka As String, ByRef tekyoTanka As String) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        trtkTanka = "0"
        tekyoTanka = "0"

        '工費の単価情報を取得する。
        If Not S004(dbr, MjK1.Cmn.Const.TuikakhSytkMstKbn.KoryoMst, "", syunYmd, pPrmId, pPrmCd) Then
            Return False
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                trtkTanka = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("trtk_tanka_kngk"), ""))
                tekyoTanka = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("tekyo_tanka_kngk"), ""))

                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return True
    End Function


    ''' <summary>
    ''' 単価情報を取得する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pTuikakhSytkMstKbn">追加工費_取得マスタ区分</param>
    ''' <param name="pKoryoZunitCd">工量_材料コード(7桁)</param>
    ''' <param name="pSyunYmd">竣工年月日</param>
    ''' <param name="pPrmId">工費項目パラメータID</param>
    ''' <param name="pPrmCd">工費項目値</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S004(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pTuikakhSytkMstKbn As String,
                         ByVal pKoryoZunitCd As String, ByVal pSyunYmd As String, ByVal pPrmId As String, ByVal pPrmCd As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            If pTuikakhSytkMstKbn = MjK1.Cmn.Const.TuikakhSytkMstKbn.KoryoMst And (Not String.IsNullOrEmpty(pPrmId)) Then
                '工量マスタ
                sb.Append(<sql><![CDATA[
SELECT
  koryo_cd                                    -- 工量コード
 ,koryo_kohi_sbt_cd                           -- 工量工費種別コード
 ,(CASE koryo_kohi_sbt_cd
      WHEN '1' THEN     koryo_tanka_kngk      -- 工量工費種別コード＝'1'の場合、工量単価
      WHEN '2' THEN     trtk_kohi_tanka_kngk  -- 工量工費種別コード＝'2'の場合、取付工費単価
      END ) AS trtk_tanka_kngk                -- 取付単価
 ,(CASE koryo_kohi_sbt_cd
      WHEN '1' THEN     koryo_tanka_kngk      -- 工量工費種別コード＝'1'の場合、工量単価
      WHEN '2' THEN     tekyo_kohi_tanka_kngk -- 工量工費種別コード＝'2'の場合、撤去工費単価
      END ) AS tekyo_tanka_kngk               -- 撤去単価
 ,trtk_koryo                                  -- 取付工量
 ,tekyo_koryo                                 -- 撤去工量
FROM
 mst_koryo 
WHERE
 1 = 1 
     AND koryo_cd = (SELECT str1_hksu_value FROM com_paramater_cd_mst WHERE prm_id = {0} AND prm_cd = {1}) 
 AND {2} BETWEEN tky_start_ymd AND tky_end_ymd 
        ]]></sql>.Value)

                sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(pPrmId) _
                               , EcOrgIO.EcOrgString.GetSqlText(pPrmCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pSyunYmd)
                            )
            ElseIf pTuikakhSytkMstKbn = MjK1.Cmn.Const.TuikakhSytkMstKbn.KoryoMst And String.IsNullOrEmpty(pPrmId) Then
                '追加工費_取得マスタ区分が工量マスタの場合
                sb.Append(<sql><![CDATA[
SELECT
  koryo_cd                                    -- 工量コード
 ,koryo_kohi_sbt_cd                          -- 工量工費種別コード    ※データ確認時に有効にする
 ,(CASE koryo_kohi_sbt_cd
      WHEN '1' THEN     koryo_tanka_kngk      -- 工量工費種別コード＝'1'の場合、工量単価
      WHEN '2' THEN     trtk_kohi_tanka_kngk  -- 工量工費種別コード＝'2'の場合、取付工費単価
      END ) AS trtk_tanka_kngk                -- 取付単価
 ,(CASE koryo_kohi_sbt_cd
      WHEN '1' THEN     koryo_tanka_kngk      -- 工量工費種別コード＝'1'の場合、工量単価
      WHEN '2' THEN     tekyo_kohi_tanka_kngk -- 工量工費種別コード＝'2'の場合、撤去工費単価
      END ) AS tekyo_tanka_kngk               -- 撤去単価
 ,trtk_koryo                                  -- 取付工量
 ,tekyo_koryo                                 -- 撤去工量
FROM
 mst_koryo 
WHERE
 1 = 1 
 AND koryo_cd = {0} 
 AND {1} BETWEEN tky_start_ymd AND tky_end_ymd 
        ]]></sql>.Value)

                sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(pKoryoZunitCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pSyunYmd)
                            )
            Else
                '追加工費_取得マスタ区分が材料ユニットマスタの場合
                sb.Append(<sql><![CDATA[
SELECT
  zunit_cd AS koryo_cd                         -- 工量コード
 ,'2' AS koryo_kohi_sbt_cd
 ,trtk_zunit_tanka_kngk  AS trtk_tanka_kngk    -- 取付材料ユニット単価 AS trtk_tanka_kngk
 ,tekyo_zunit_tanka_kngk AS tekyo_tanka_kngk   -- 撤去材料ユニット単価 AS tekyo_tanka_kngk
,'0' AS trtk_koryo
,'0' AS tekyo_koryo
FROM
 mst_zunit 
WHERE
 1 = 1 
 AND zunit_cd = {0} 
 AND {1} BETWEEN tky_start_ymd AND tky_end_ymd;
        ]]></sql>.Value)

                sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(pKoryoZunitCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pSyunYmd)
                            )
            End If

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

    ''' <summary>
    ''' 金額計算
    ''' </summary>
    ''' <param name="pTanka">単価</param>
    ''' <param name="pSuryo">数量</param>
    ''' <param name="pWrmsRitu">割増率</param>
    ''' <returns>取付工事費金額</returns>
    Friend Function CalcKozihKngk(ByVal pTanka As String, ByVal pSuryo As Decimal, pWrmsRitu As String) As Decimal
        Dim tankaD As Decimal = MyCDec(pTanka)
        Dim wrmsrituD As Decimal = MyCDec(pWrmsRitu)
        Dim wrmsTankaD As Decimal = tankaD * wrmsrituD

        'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD Start
        Dim wrmsTankaTruncD As Decimal = TruncateDec(wrmsTankaD, [Const].C_YUKO_SHOSU_KETA)
        'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD End

        '金額算出後、小数切り捨て
        'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD Start
        'Return CDec(Math.Truncate(wrmsTankaD * pSuryo))
        Return CDec(Math.Truncate(wrmsTankaTruncD * pSuryo))
        'Rev052 20241010 工費計算の小数第２位切り捨て対応 MOD Start
    End Function


    ''' <summary>
    ''' 割増金額算出
    ''' </summary>
    ''' <param name="pKngk">金額</param>
    ''' <param name="pWrmsRitu">割増率</param>
    ''' <returns>割増金額算出</returns>
    Friend Function CalcWrmsKngk(ByVal pKngk As Decimal, pWrmsRitu As String) As Decimal
        Dim wrmsrituD As Decimal = MyCDec(pWrmsRitu)
        Dim wrmsKngkD As Decimal = pKngk * wrmsrituD

        '金額算出後、小数切り捨て
        Dim wrmsKngk As Decimal = CDec(Math.Truncate(wrmsKngkD))
        Return wrmsKngk
    End Function
#End Region

    '/* 単価取得方法変更　金額計算変更 2023/4/15
    '#Region "取付工事金額算出"
    '    ''' <summary>
    '    ''' マスター工量　取付工費単価取得
    '    ''' </summary>
    '    ''' <param name="pKoryoCd">工量コード</param>
    '    ''' <param name="pSyunYmd">竣工日</param>
    '    ''' <returns>取付工費単価</returns>
    '    Friend Function getKoryoTrtkTanka(ByVal pKoryoCd As String, ByVal pSyunYmd As String) As String
    '        Dim wStr As String = ""
    '        Dim jouken1 As String = "koryo_cd = " & EcOrgIO.EcOrgString.GetSqlText(pKoryoCd.PadLeft([Const].C_KORYO_CD_KETA, "0"c))
    '        Dim jouken2 As String = EcOrgIO.EcOrgString.GetSqlText(pSyunYmd) & " BETWEEN tky_start_ymd AND tky_end_ymd "
    '        wStr = myDB.GetRecValueString("mst_koryo", "trtk_kohi_tanka_kngk", jouken1, jouken2)
    '        If String.IsNullOrEmpty(wStr.Trim) Then
    '            wStr = "0"
    '        End If
    '        Return wStr
    '    End Function

    '    ''' <summary>
    '    ''' 取付工事金額
    '    ''' </summary>
    '    ''' <param name="pKoryoCd">工量コード</param>
    '    ''' <param name="pSuryo">数量</param>
    '    ''' <param name="pSyunYmd">竣工年月日</param>
    '    ''' <param name="pWrmsRitu">割増率</param>
    '    ''' <returns>取付工事費金額</returns>
    '    Friend Function CalcTrtkKozihKngk(ByVal pKoryoCd As String, ByVal pSuryo As Decimal, ByVal pSyunYmd As String, pWrmsRitu As String) As Decimal
    '        Dim tanka As String = getKoryoTrtkTanka(pKoryoCd, pSyunYmd)
    '        Dim tankaD As Double = CDbl(tanka)
    '        Dim wrmsrituD As Double = CDbl(pWrmsRitu)

    '        Dim wrmsTankaD As Double = tankaD * wrmsrituD
    '        Dim wrmsTanka As Decimal = CDec(Math.Truncate(wrmsTankaD))

    '        Return wrmsTanka * pSuryo
    '    End Function
    '#End Region

    '#Region "撤去工事金額算出"
    '    ''' <summary>
    '    ''' マスター工量　撤去工費単価取得
    '    ''' </summary>
    '    ''' <param name="pKoryoCd">工量コード</param>
    '    ''' <param name="pSyunYmd">竣工日</param>
    '    ''' <returns>撤去工費単価</returns>
    '    Friend Function getKoryoTekyoTanka(ByVal pKoryoCd As String, ByVal pSyunYmd As String) As String
    '        Dim wStr As String = ""
    '        Dim jouken1 As String = "koryo_cd = " & EcOrgIO.EcOrgString.GetSqlText(pKoryoCd.PadLeft([Const].C_KORYO_CD_KETA, "0"c))
    '        Dim jouken2 As String = EcOrgIO.EcOrgString.GetSqlText(pSyunYmd) & " BETWEEN tky_start_ymd AND tky_end_ymd "
    '        wStr = myDB.GetRecValueString("mst_koryo", "tekyo_kohi_tanka_kngk", jouken1, jouken2)
    '        If String.IsNullOrEmpty(wStr.Trim) Then
    '            wStr = "0"
    '        End If
    '        Return wStr
    '    End Function

    '    ''' <summary>
    '    ''' 撤去工事金額
    '    ''' </summary>
    '    ''' <param name="pKoryoCd">工量コード</param>
    '    ''' <param name="pSuryo">数量</param>
    '    ''' <param name="pSyunYmd">竣工年月日</param>
    '    ''' <returns>撤去工事費金額</returns>
    '    Friend Function CalcTekyoKozihKngk(ByVal pKoryoCd As String, ByVal pSuryo As Decimal, ByVal pSyunYmd As String, pWrmsRitu As String) As Decimal
    '        Dim tanka As String = getKoryoTekyoTanka(pKoryoCd, pSyunYmd)
    '        Dim tankaD As Double = CDbl(tanka)
    '        Dim wrmsrituD As Double = CDbl(pWrmsRitu)

    '        Dim wrmsTankaD As Double = tankaD * wrmsrituD
    '        Dim wrmsTanka As Decimal = CDec(Math.Truncate(wrmsTankaD))

    '        Return wrmsTanka * pSuryo
    '    End Function
    '#End Region
    '*/

#Region "指示数　桁合わせ"
    ''' <summary>
    ''' 指示数　桁調整後取得
    ''' </summary>
    ''' <param name="pSzsu">指示数</param>
    ''' <returns>桁調整後指示数</returns>
    Friend Function getSzsuStr(ByVal pSzsu As String) As String
        Dim str As String = pSzsu.Trim.PadLeft(MjK1.Cmn.Const.C_SIZISU_KETA, "0"c)
        Return str
    End Function
#End Region

#Region "(SQL)取替_取替票行程更新"
    ''' <summary>
    ''' 取替_取替票行程更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票行程を更新します</remarks>
    Friend Function U002(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sb2 As New StringBuilder                    'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC SET         -- 取替行程
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,seko_tanto_syori_ymd={3}      --施工（担当）_処理年月日
 ,seko_tanto_syors_cd={4}      --施工（担当）_処理者コード
 ,seko_tanto_syors_ms={5}      --施工（担当）_処理者名
 ,seko_tanto_ka_cd={6}      --施工（担当）_課コード
 ,seko_tanto_ka_ms={7}      --施工（担当）_課名
 ,seko_tanto_kzkis_mdgt_cd={8}      --施工（担当）_工事会社窓口コード
 ,seko_tanto_kzkis_mdgt_ms={9}      --施工（担当）_工事会社窓口名
 ,seko_tanto_kozitn_cd={10}      --施工（担当）_工事店コード
 ,seko_tanto_kozitn_ms={11}      --施工（担当）_工事店名
WHERE 1 = 1
 AND  dig4_zgsyo_cd={12}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={13}      --取替票発行年度
 AND  trkhy_kbn={14}      --取替票区分
 AND  trkhy_no={15}      --取替票番号

         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kzkis_mdgt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kzkis_mdgt_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kozitn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kozitn_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If


            '実費確認フラグの更新
            Dim zippi_umu_flg As String = CStr(Me.myDB.ConvDbNull(pRow("zippi_umu_flg"), ""))
            If zippi_umu_flg = MjK1.Cmn.Const.UmuFlg.FlgOn Then
                sb2.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC SET         -- 取替行程
  no1_skrepo_zippi_kknn_flg={0}      --竣工報告_実費確認1フラグ
 ,no2_skrepo_zippi_kknn_flg={1}      --竣工報告_実費確認2フラグ
 ,no3_skrepo_zippi_kknn_flg={2}      --竣工報告_実費確認3フラグ
WHERE 1 = 1
 AND  dig4_zgsyo_cd={3}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={4}      --取替票発行年度
 AND  trkhy_kbn={5}      --取替票区分
 AND  trkhy_no={6}      --取替票番号
         ]]></sql>.Value)

                sqlstr = String.Format(sb2.ToString,
                                     EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                                   )
                '*************************************************************************

                'SQLログ出力
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

                'SQL実行
                If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                    Throw New Exception(Me.myDB.ErrDescription)
                End If
                If wCount = 0 Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                    Throw New Exception("竣工報告_実費確認フラグ 更新結果0件エラー")
                End If
            End If


            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)取替_取替票行程（高圧他）更新"
    ''' <summary>
    ''' 取替_取替票行程（高圧他）更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票行程を更新します</remarks>
    Friend Function U013(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sb2 As New StringBuilder                    'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC_KAHK SET         -- 取替行程
  upd_date={0}                          --更新日時
 ,sousa_user_id={1}                     --更新ユーザ
 ,sousa_appli_cd={2}                    --更新プログラム
 ,seko_tanto_syori_ymd={3}              --施工（担当）_処理年月日
 ,seko_tanto_syors_cd={4}               --施工（担当）_処理者コード
 ,seko_tanto_syors_ms={5}               --施工（担当）_処理者名
 ,seko_tanto_ka_cd={6}                  --施工（担当）_課コード
 ,seko_tanto_ka_ms={7}                  --施工（担当）_課名
 ,seko_tanto_kzkis_mdgt_cd={8}          --施工（担当）_工事会社窓口コード
 ,seko_tanto_kzkis_mdgt_ms={9}          --施工（担当）_工事会社窓口名
 ,seko_tanto_kozitn_cd={10}             --施工（担当）_工事店コード
 ,seko_tanto_kozitn_ms={11}             --施工（担当）_工事店名
WHERE 1 = 1
 AND  dig4_zgsyo_cd={12}                --事業所コード（4桁）
 AND  trkhy_hakko_nendo={13}            --取替票発行年度
 AND  trkhy_kbn={14}                    --取替票区分
 AND  trkhy_no={15}                     --取替票番号

         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kzkis_mdgt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kzkis_mdgt_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kozitn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("seko_tanto_kozitn_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If


            '実費確認フラグの更新
            Dim zippi_umu_flg As String = CStr(Me.myDB.ConvDbNull(pRow("zippi_umu_flg"), ""))
            If zippi_umu_flg = MjK1.Cmn.Const.UmuFlg.FlgOn Then
                sb2.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC_KAHK SET     -- 取替行程
  no1_skrepo_zippi_kknn_flg={0}     --竣工報告_実費確認1フラグ
 ,no2_skrepo_zippi_kknn_flg={1}     --竣工報告_実費確認2フラグ
 ,no3_skrepo_zippi_kknn_flg={2}     --竣工報告_実費確認3フラグ
 ,no4_skrepo_zippi_kknn_flg={3}     --竣工報告_実費確認3フラグ
 ,no5_skrepo_zippi_kknn_flg={4}     --竣工報告_実費確認3フラグ
WHERE 1 = 1
 AND  dig4_zgsyo_cd={5}             --事業所コード（4桁）
 AND  trkhy_hakko_nendo={6}         --取替票発行年度
 AND  trkhy_kbn={7}                 --取替票区分
 AND  trkhy_no={8}                  --取替票番号
         ]]></sql>.Value)

                sqlstr = String.Format(sb2.ToString,
                                     EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no1_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no2_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no3_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no4_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("no5_skrepo_zippi_kknn_flg"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                                   , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                                   )
                '*************************************************************************

                'SQLログ出力
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

                'SQL実行
                If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                    Throw New Exception(Me.myDB.ErrDescription)
                End If
                If wCount = 0 Then
                    HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                    Throw New Exception("竣工報告_実費確認フラグ 更新結果0件エラー")
                End If
            End If


            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region



#Region "(SQL)計器_授受履歴追加"
    ''' <summary>
    ''' 計器_授受履歴追加
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pZyzyKbn">授受区分</param>
    ''' <param name="pKeikiZyzyStatusCd">計器授受ステータスコード</param>
    ''' <param name="pClmName">計器IDカラム名</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>計器_授受履歴（撤去計器）を追加します</remarks>
    Friend Function I001(ByVal pRow As DataRow, ByVal pZyzyKbn As String, ByVal pKeikiZyzyStatusCd As String, ByVal pClmName As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO KEIKI_JUJU_RIREKI        -- 計器_授受履歴
VALUES (
  {0}              --作成日時
 ,{1}              --更新日時
 ,{2}              --更新ユーザ
 ,{3}              --更新プログラム
 ,{4}              --SM計器ID
 ,{5}              --授受区分
 ,{6}              --計器授受ステータスコード
 ,{7}              --事業所コード（4桁）
 ,{8}              --保管工事会社窓口コード
 ,{9}              --保管工事会社窓口名
 ,{10}              --保管工事店コード
 ,{11}              --保管工事店名
 ,{12}              --登録担当者コード
 ,{13}              --登録担当者名
 ,{14}              --登録日時
 ,{15}              --確認状況フラグ
 ,{16}              --確認年月日
)
         ]]></sql>.Value)

            Dim wkstr As String = CStr(Me.myDB.ConvDbNull(pRow("kzkis_mdgt_cd"), ""))
            Dim whokan_kzkis_mdgt_ms As String = getKzkisMdgtMs(wkstr)
            wkstr = CStr(Me.myDB.ConvDbNull(pRow("kozitn_cd"), ""))
            Dim whokan_kozitn_ms As String = getKozitnMs(wkstr)
            Dim procDate As String = getDbDateYmd()
            Dim procDateTime As String = getDbDateTime()

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow(pClmName), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(pZyzyKbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(pKeikiZyzyStatusCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kzkis_mdgt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(whokan_kzkis_mdgt_ms) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kozitn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(whokan_kozitn_ms) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.UmuFlg.FlgOff) _
                               , EcOrgIO.EcOrgString.GetSqlText("")
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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


    ''' <summary>
    ''' マスター工事窓口　窓口名称取得(直営含む)
    ''' </summary>
    ''' <param name="pMdgtCd">窓口コード</param>
    ''' <returns>窓口名称</returns>
    Friend Function getKzkisMdgtMs(ByVal pMdgtCd As String) As String
        Dim wMS As String = ""
        Dim jouken1 As String = "mdgt_cd = " & EcOrgIO.EcOrgString.GetSqlText(pMdgtCd)
        Dim jouken2 As String = "latest_flg = " & EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.UmuFlg.FlgOn)

        Dim wChokueiCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_MDGT_INFO, SyoriYMD)
        If pMdgtCd.Equals(wChokueiCd) Then
            wMS = MjK1.Cmn.Func.GetZyousu2(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_MDGT_INFO, SyoriYMD)
        Else
            wMS = myDB.GetRecValueString("mst_mdgt", "mdgt_ms", jouken1, jouken2)
        End If

        Return wMS
    End Function

    ''' <summary>
    ''' マスター工事店　工事店名取得(直営含む)
    ''' </summary>
    ''' <param name="pKozitnCd">工事店コード</param>
    ''' <returns>工事店名</returns>
    Friend Function getKozitnMs(ByVal pKozitnCd As String) As String
        Dim wMS As String = ""
        Dim jouken1 As String = "kozitn_cd = " & EcOrgIO.EcOrgString.GetSqlText(pKozitnCd)
        Dim jouken2 As String = "latest_flg = " & EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.UmuFlg.FlgOn)

        Dim wChokueiCd = MjK1.Cmn.Func.GetZyousu(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, SyoriYMD)
        If pKozitnCd.Equals(wChokueiCd) Then
            wMS = MjK1.Cmn.Func.GetZyousu2(myDB, MjK1.Cmn.Const.ZyousuCd.CHOKUEI_KOZITN_INFO, SyoriYMD)
        Else
            wMS = myDB.GetRecValueString("mst_kozitn", "kozitn_ms", jouken1, jouken2)
        End If

        Return wMS
    End Function

    ''' <summary>
    ''' 処理年月日を取得する
    ''' </summary>
    ''' <returns>処理年月日</returns>
    Friend Function getDbDateYmd() As String
        Dim wDate As Date
        If myDB.GetDbDate(wDate) Then
            Return wDate.ToString("yyyyMMdd")
        Else
            Return Now.ToString("yyyyMMdd")
        End If
    End Function

    ''' <summary>
    ''' 処理日時を取得する
    ''' </summary>
    ''' <returns>処理日時</returns>
    Friend Function getDbDateTime() As String
        Dim wDate As Date
        If myDB.GetDbDateTime(wDate) Then
            Return wDate.ToString("yyyyMMddHHmmss")
        Else
            Return Now.ToString("yyyyMMddHHmmss")
        End If
    End Function


#Region "(SQL)計取替_取替票追加工費追加"
    ''' <summary>
    ''' 取替_取替票追加工費追加
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pTrtkKojiHi">取付工事費</param>
    ''' <param name="pTekyoKojiHi">撤去工事費</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を追加します</remarks>
    Friend Function I003(ByVal pRow As DataRow, ByVal pTrtkKojiHi As String, ByVal pTekyoKojiHi As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO trke_trkehy_tuikakh        -- 取替_取替票追加工費
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --事業所コード（4桁）
 ,{5}          --取替票発行年度
 ,{6}          --取替票区分
 ,{7}          --取替票番号
 ,{8}          --追加工費_行NO
 ,{9}          --追加工費_工事区分
 ,{10}          --追加工費_取付数量
 ,{11}          --追加工費_撤去数量
 ,{12}          --取付工事費
 ,{13}          --撤去工事費
 ,{14}          --追加工費理由内容
)
         ]]></sql>.Value)

            '追加工費工事区分が4桁超の場合、先頭の「000」を削除し4桁にする。
            Dim wtuika_kohi_kozi_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_kozi_kbn"), ""))
            If wtuika_kohi_kozi_kbn.Length > [Const].C_TUIKAKOHI_KOZI_KBN_KETA Then
                wtuika_kohi_kozi_kbn = Right(wtuika_kohi_kozi_kbn, [Const].C_TUIKAKOHI_KOZI_KBN_KETA)
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_row_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtuika_kohi_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_trtk_suryo"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_tekyo_suryo"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pTrtkKojiHi) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pTekyoKojiHi) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_riyu_naiyo"), "")))
                               )
            '*************************************************************************


            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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



#Region "(SQL)計取替_取替票追加工費（高圧他）追加"
    ''' <summary>
    ''' 取替_取替票追加工費（高圧他）追加
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pTrtkKojiHi">取付工事費</param>
    ''' <param name="pTekyoKojiHi">撤去工事費</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を追加します</remarks>
    Friend Function I009(ByVal pRow As DataRow, ByVal pTrtkKojiHi As String, ByVal pTekyoKojiHi As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO trke_trkehy_tuikakh_kahk        -- 取替_取替票追加工費（高圧他）
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --事業所コード（4桁）
 ,{5}          --取替票発行年度
 ,{6}          --取替票区分
 ,{7}          --取替票番号
 ,{8}          --追加工費_行NO
 ,{9}          --追加工費_工事区分
 ,{10}          --追加工費_取付数量
 ,{11}          --追加工費_撤去数量
 ,{12}          --取付工事費
 ,{13}          --撤去工事費
 ,{14}          --追加工費理由内容
)
         ]]></sql>.Value)

            '追加工費工事区分が4桁超の場合、先頭の「000」を削除し4桁にする。
            Dim wtuika_kohi_kozi_kbn As String = CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_kozi_kbn"), ""))
            If wtuika_kohi_kozi_kbn.Length > [Const].C_TUIKAKOHI_KOZI_KBN_KETA Then
                wtuika_kohi_kozi_kbn = Right(wtuika_kohi_kozi_kbn, [Const].C_TUIKAKOHI_KOZI_KBN_KETA)
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_row_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(wtuika_kohi_kozi_kbn) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_trtk_suryo"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_tekyo_suryo"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pTrtkKojiHi) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(pTekyoKojiHi) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tuika_kohi_riyu_naiyo"), "")))
                               )
            '*************************************************************************


            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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

#Region "(SQL)計取替_取替票追加工費削除"
    ''' <summary>
    ''' 取替_取替票追加工費削除(DataRow)
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を削除します</remarks>
    Friend Function D001(ByVal pRow As DataRow) As Boolean
        Return D001(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), "")), CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), "")),
                    CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), "")), CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
    End Function

    ''' <summary>
    ''' 取替_取替票追加工費削除
    ''' </summary>
    ''' <param name="dig4_zgsyo_cd">事業所コード（4桁）</param>
    ''' <param name="trkhy_hakko_nendo">取替票発行年度</param>
    ''' <param name="trkhy_kbn">取替票区分</param>
    ''' <param name="trkhy_no">取替票番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を削除します</remarks>
    Friend Function D001(ByVal dig4_zgsyo_cd As String, ByVal trkhy_hakko_nendo As String, ByVal trkhy_kbn As String, ByVal trkhy_no As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
DELETE from trke_trkehy_tuikakh        -- 取替_取替票追加工費
WHERE 1 = 1
 AND  dig4_zgsyo_cd={0}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={1}      --取替票発行年度
 AND  trkhy_kbn={2}      --取替票区分
 AND  trkhy_no={3}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(dig4_zgsyo_cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_hakko_nendo) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_no)
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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

#Region "単価情報を取得する。"
    ''' <summary>
    ''' 単価情報を取得する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pTuikaKohiKoziKbn">工事費区分</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S002(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pTuikaKohiKoziKbn As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
select 
 A.tuikakh_sytk_mst_kbn
,B.koryo_kohi_sbt_cd
,B.trtk_kohi_tanka_kngk
,B.tekyo_kohi_tanka_kngk
,C.trtk_zunit_tanka_kngk
,C.tekyo_zunit_tanka_kngk
from
MST_USE_KANO_TUIKAKH A 
left join mst_koryo B On (B.koryo_cd=A.koryo_zunit_cd And {0} BETWEEN B.tky_start_ymd and B.tky_end_ymd)
left join mst_zunit C ON (C.zunit_cd=A.koryo_zunit_cd and {0} BETWEEN C.tky_start_ymd and C.tky_end_ymd)
WHERE 1=1
AND A.koryo_zunit_cd={1}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(SyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(pTuikaKohiKoziKbn)
                            )
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


#Region "(SQL)取替_自主点検チェックシート追加(全項目)"
    ''' <summary>
    ''' 取替_自主点検チェックシート追加
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシートを追加します</remarks>
    Friend Function I004(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev040 2024/12/25 物理分割 託送切替対応 ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_tio_cmp_ymd
            sb.Append(<sql><![CDATA[
INSERT INTO trke_check_sheet        -- 取替_自主点検チェックシート
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --事業所コード（4桁）
 ,{5}          --取替票発行年度
 ,{6}          --取替票区分
 ,{7}          --取替票番号
 ,{8}          --低圧計器チェック・電力量計・取付取替等・施工者・良否区分
 ,{9}          --低圧計器チェック・電力量計・取付取替等・抜取検査担当者・良否区分
 ,{10}          --低圧計器チェック・電力量計・取付取替等・検査担当者・良否区分
 ,{11}          --低圧計器チェック・電力量計・取付取替等・対応内容
 ,{12}          --低圧計器チェック・電力量計・取付取替等・対応完了日
 ,{13}          --低圧計器チェック・電力量計・指示数・施工者・良否区分
 ,{14}          --低圧計器チェック・電力量計・指示数・抜取検査担当者・良否区分
 ,{15}          --低圧計器チェック・電力量計・指示数・検査担当者・良否区分
 ,{16}          --低圧計器チェック・電力量計・指示数・対応内容
 ,{17}          --低圧計器チェック・電力量計・指示数・対応完了日
 ,{18}          --低圧計器チェック・計器項目・相線電圧・施工者・良否区分
 ,{19}          --低圧計器チェック・計器項目・相線電圧・抜取検査担当者・良否区分
 ,{20}          --低圧計器チェック・計器項目・相線電圧・検査担当者・良否区分
 ,{21}          --低圧計器チェック・計器項目・相線電圧・対応内容
 ,{22}          --低圧計器チェック・計器項目・相線電圧・対応完了日
 ,{23}          --低圧計器チェック・計器項目・乗率・施工者・良否区分
 ,{24}          --低圧計器チェック・計器項目・乗率・抜取検査担当者・良否区分
 ,{25}          --低圧計器チェック・計器項目・乗率・検査担当者・良否区分
 ,{26}          --低圧計器チェック・計器項目・乗率・対応内容
 ,{27}          --低圧計器チェック・計器項目・乗率・対応完了日
 ,{28}          --低圧計器チェック・計器項目・容量・施工者・良否区分
 ,{29}          --低圧計器チェック・計器項目・容量・抜取検査担当者・良否区分
 ,{30}          --低圧計器チェック・計器項目・容量・検査担当者・良否区分
 ,{31}          --低圧計器チェック・計器項目・容量・対応内容
 ,{32}          --低圧計器チェック・計器項目・容量・対応完了日
 ,{33}          --低圧計器チェック・計器項目・SM・施工者・良否区分
 ,{34}          --低圧計器チェック・計器項目・SM・抜取検査担当者・良否区分
 ,{35}          --低圧計器チェック・計器項目・SM・検査担当者・良否区分
 ,{36}          --低圧計器チェック・計器項目・SM・対応内容
 ,{37}          --低圧計器チェック・計器項目・SM・対応完了日
 ,{38}          --低圧計器チェック・計器項目・計器・合番号・施工者・良否区分
 ,{39}          --低圧計器チェック・計器項目・計器・合番号・抜取検査担当者・良否区分
 ,{40}          --低圧計器チェック・計器項目・計器・合番号・検査担当者・良否区分
 ,{41}          --低圧計器チェック・計器項目・計器・合番号・対応内容
 ,{42}          --低圧計器チェック・計器項目・計器・合番号・対応完了日
 ,{43}          --低圧計器チェック・計器項目・変流器・製造番号・施工者・良否区分
 ,{44}          --低圧計器チェック・計器項目・変流器・製造番号・抜取検査担当者・良否区分
 ,{45}          --低圧計器チェック・計器項目・変流器・製造番号・検査担当者・良否区分
 ,{46}          --低圧計器チェック・計器項目・変流器・製造番号・対応内容
 ,{47}          --低圧計器チェック・計器項目・変流器・製造番号・対応完了日
 ,{48}          --低圧計器チェック・計器項目・変流器・合番号・施工者・良否区分
 ,{49}          --低圧計器チェック・計器項目・変流器・合番号・抜取検査担当者・良否区分
 ,{50}          --低圧計器チェック・計器項目・変流器・合番号・検査担当者・良否区分
 ,{51}          --低圧計器チェック・計器項目・変流器・合番号・対応内容
 ,{52}          --低圧計器チェック・計器項目・変流器・合番号・対応完了日
 ,{53}          --低圧計器チェック・構造検査・計量器・検電・施工者・良否区分
 ,{54}          --低圧計器チェック・構造検査・計量器・検電・抜取検査担当者・良否区分
 ,{55}          --低圧計器チェック・構造検査・計量器・検電・検査担当者・良否区分
 ,{56}          --低圧計器チェック・構造検査・計量器・検電・対応内容
 ,{57}          --低圧計器チェック・構造検査・計量器・検電・対応完了日
 ,{58}          --低圧計器チェック・構造検査・計量器箱・施工者・良否区分
 ,{59}          --低圧計器チェック・構造検査・計量器箱・抜取検査担当者・良否区分
 ,{60}          --低圧計器チェック・構造検査・計量器箱・検査担当者・良否区分
 ,{61}          --低圧計器チェック・構造検査・計量器箱・対応内容
 ,{62}          --低圧計器チェック・構造検査・計量器箱・対応完了日
 ,{63}          --低圧計器チェック・結線等・計器・絶縁・施工者・良否区分
 ,{64}          --低圧計器チェック・結線等・計器・絶縁・抜取検査担当者・良否区分
 ,{65}          --低圧計器チェック・結線等・計器・絶縁・検査担当者・良否区分
 ,{66}          --低圧計器チェック・結線等・計器・絶縁・対応内容
 ,{67}          --低圧計器チェック・結線等・計器・絶縁・対応完了日
 ,{68}          --低圧計器チェック・結線等・計器・SL・施工者・良否区分
 ,{69}          --低圧計器チェック・結線等・計器・SL・抜取検査担当者・良否区分
 ,{70}          --低圧計器チェック・結線等・計器・SL・検査担当者・良否区分
 ,{71}          --低圧計器チェック・結線等・計器・SL・対応内容
 ,{72}          --低圧計器チェック・結線等・計器・SL・対応完了日
 ,{73}          --低圧計器チェック・結線等・計器・複数・施工者・良否区分
 ,{74}          --低圧計器チェック・結線等・計器・複数・抜取検査担当者・良否区分
 ,{75}          --低圧計器チェック・結線等・計器・複数・検査担当者・良否区分
 ,{76}          --低圧計器チェック・結線等・計器・複数・対応内容
 ,{77}          --低圧計器チェック・結線等・計器・複数・対応完了日
 ,{78}          --低圧計器チェック・結線等・計器・誤接続・施工者・良否区分
 ,{79}          --低圧計器チェック・結線等・計器・誤接続・抜取検査担当者・良否区分
 ,{80}          --低圧計器チェック・結線等・計器・誤接続・検査担当者・良否区分
 ,{81}          --低圧計器チェック・結線等・計器・誤接続・対応内容
 ,{82}          --低圧計器チェック・結線等・計器・誤接続・対応完了日
 ,{83}          --低圧計器チェック・結線等・計器・赤色・施工者・良否区分
 ,{84}          --低圧計器チェック・結線等・計器・赤色・抜取検査担当者・良否区分
 ,{85}          --低圧計器チェック・結線等・計器・赤色・検査担当者・良否区分
 ,{86}          --低圧計器チェック・結線等・計器・赤色・対応内容
 ,{87}          --低圧計器チェック・結線等・計器・赤色・対応完了日
 ,{88}          --低圧計器チェック・結線等・計器・より線・施工者・良否区分
 ,{89}          --低圧計器チェック・結線等・計器・より線・抜取検査担当者・良否区分
 ,{90}          --低圧計器チェック・結線等・計器・より線・検査担当者・良否区分
 ,{91}          --低圧計器チェック・結線等・計器・より線・対応内容
 ,{92}          --低圧計器チェック・結線等・計器・より線・対応完了日
 ,{93}          --低圧計器チェック・結線等・計器・露出・施工者・良否区分
 ,{94}          --低圧計器チェック・結線等・計器・露出・抜取検査担当者・良否区分
 ,{95}          --低圧計器チェック・結線等・計器・露出・検査担当者・良否区分
 ,{96}          --低圧計器チェック・結線等・計器・露出・対応内容
 ,{97}          --低圧計器チェック・結線等・計器・露出・対応完了日
 ,{98}          --低圧計器チェック・結線等・計器・ビス締付・施工者・良否区分
 ,{99}          --低圧計器チェック・結線等・計器・ビス締付・抜取検査担当者・良否区分
 ,{100}          --低圧計器チェック・結線等・計器・ビス締付・検査担当者・良否区分
 ,{101}          --低圧計器チェック・結線等・計器・ビス締付・対応内容
 ,{102}          --低圧計器チェック・結線等・計器・ビス締付・対応完了日
 ,{103}          --低圧計器チェック・結線等・計器・色別・施工者・良否区分
 ,{104}          --低圧計器チェック・結線等・計器・色別・抜取検査担当者・良否区分
 ,{105}          --低圧計器チェック・結線等・計器・色別・検査担当者・良否区分
 ,{106}          --低圧計器チェック・結線等・計器・色別・対応内容
 ,{107}          --低圧計器チェック・結線等・計器・色別・対応完了日
 ,{108}          --低圧計器チェック・結線等・計器・正相・施工者・良否区分
 ,{109}          --低圧計器チェック・結線等・計器・正相・抜取検査担当者・良否区分
 ,{110}          --低圧計器チェック・結線等・計器・正相・検査担当者・良否区分
 ,{111}          --低圧計器チェック・結線等・計器・正相・対応内容
 ,{112}          --低圧計器チェック・結線等・計器・正相・対応完了日
 ,{113}          --低圧計器チェック・結線等・変流器・ＫＬ・施工者・良否区分
 ,{114}          --低圧計器チェック・結線等・変流器・ＫＬ・抜取検査担当者・良否区分
 ,{115}          --低圧計器チェック・結線等・変流器・ＫＬ・検査担当者・良否区分
 ,{116}          --低圧計器チェック・結線等・変流器・ＫＬ・対応内容
 ,{117}          --低圧計器チェック・結線等・変流器・ＫＬ・対応完了日
 ,{118}          --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・施工者・良否区分
 ,{119}          --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・抜取検査担当者・良否区分
 ,{120}          --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・検査担当者・良否区分
 ,{121}          --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応内容
 ,{122}          --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応完了日
 ,{123}          --低圧計器チェック・結線等・変流器・赤色・施工者・良否区分
 ,{124}          --低圧計器チェック・結線等・変流器・赤色・抜取検査担当者・良否区分
 ,{125}          --低圧計器チェック・結線等・変流器・赤色・検査担当者・良否区分
 ,{126}          --低圧計器チェック・結線等・変流器・赤色・対応内容
 ,{127}          --低圧計器チェック・結線等・変流器・赤色・対応完了日
 ,{128}          --低圧計器チェック・計器回転・検電・変流器・施工者・良否区分
 ,{129}          --低圧計器チェック・計器回転・検電・変流器・抜取検査担当者・良否区分
 ,{130}          --低圧計器チェック・計器回転・検電・変流器・検査担当者・良否区分
 ,{131}          --低圧計器チェック・計器回転・検電・変流器・対応内容
 ,{132}          --低圧計器チェック・計器回転・検電・変流器・対応完了日
 ,{133}          --低圧計器チェック・計器回転・回転・SM・施工者・良否区分
 ,{134}          --低圧計器チェック・計器回転・回転・SM・抜取検査担当者・良否区分
 ,{135}          --低圧計器チェック・計器回転・回転・SM・検査担当者・良否区分
 ,{136}          --低圧計器チェック・計器回転・回転・SM・対応内容
 ,{137}          --低圧計器チェック・計器回転・回転・SM・対応完了日
 ,{138}          --低圧計器チェック・通信部・ＬＥＤ・施工者・良否区分
 ,{139}          --低圧計器チェック・通信部・ＬＥＤ・抜取検査担当者・良否区分
 ,{140}          --低圧計器チェック・通信部・ＬＥＤ・検査担当者・良否区分
 ,{141}          --低圧計器チェック・通信部・ＬＥＤ・対応内容
 ,{142}          --低圧計器チェック・通信部・ＬＥＤ・対応完了日
 ,{143}          --低圧計器チェック・通信部・装着・施工者・良否区分
 ,{144}          --低圧計器チェック・通信部・装着・抜取検査担当者・良否区分
 ,{145}          --低圧計器チェック・通信部・装着・検査担当者・良否区分
 ,{146}          --低圧計器チェック・通信部・装着・対応内容
 ,{147}          --低圧計器チェック・通信部・装着・対応完了日
 ,{148}          --低圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,{149}          --低圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,{150}          --低圧計器チェック・撤去計器等・製造番号・検査担当者・良否区分
 ,{151}          --低圧計器チェック・撤去計器等・製造番号・対応内容
 ,{152}          --低圧計器チェック・撤去計器等・製造番号・対応完了日
 ,{153}          --低圧計器チェック・撤去計器等・合番号・施工者・良否区分
 ,{154}          --低圧計器チェック・撤去計器等・合番号・抜取検査担当者・良否区分
 ,{155}          --低圧計器チェック・撤去計器等・合番号・検査担当者・良否区分
 ,{156}          --低圧計器チェック・撤去計器等・合番号・対応内容
 ,{157}          --低圧計器チェック・撤去計器等・合番号・対応完了日
 ,{158}          --低圧計器チェック・その他・写真・施工者・良否区分
 ,{159}          --低圧計器チェック・その他・写真・抜取検査担当者・良否区分
 ,{160}          --低圧計器チェック・その他・写真・検査担当者・良否区分
 ,{161}          --低圧計器チェック・その他・写真・対応内容
 ,{162}          --低圧計器チェック・その他・写真・対応完了日
 ,{163}          --低圧計器チェック・その他・撤去取替・施工者・良否区分
 ,{164}          --低圧計器チェック・その他・撤去取替・抜取検査担当者・良否区分
 ,{165}          --低圧計器チェック・その他・撤去取替・検査担当者・良否区分
 ,{166}          --低圧計器チェック・その他・撤去取替・対応内容
 ,{167}          --低圧計器チェック・その他・撤去取替・対応完了日
 ,{168}          --低圧計器チェック・乗率チェック・抜取検査担当者・整合確認区分
 ,{169}          --低圧計器チェック・乗率チェック・検査担当者・整合確認区分
 ,{170}          --低圧計器チェック・乗率チェック・対応内容
 ,{171}          --低圧計器チェック・乗率チェック・対応完了日
 ,{172}          --低圧計器チェック・合番号チェック・抜取検査担当者・一致確認区分
 ,{173}          --低圧計器チェック・合番号チェック・検査担当者・一致確認区分
 ,{174}          --低圧計器チェック・合番号チェック・対応内容
 ,{175}          --低圧計器チェック・合番号チェック・対応完了日
 ,{176}          --記事（施工者）内容
 ,{177}          --記事（抜取検査担当者）内容
 ,{178}          --記事（検査担当者）内容
 ,{179}          --変流器付計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,{180}          --変流器付計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,{181}          --変流器付計器・施工者・青（P3）・結線等チェック結果フラグ
 ,{182}          --変流器付計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,{183}          --変流器付計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,{184}          --変流器付計器・施工者・白（P2）・結線等チェック結果フラグ
 ,{185}          --変流器付計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,{186}          --P1・施工者・黒（1S）・結線等チェック結果フラグ
 ,{187}          --P1・施工者・赤（P1）・結線等チェック結果フラグ
 ,{188}          --P1・施工者・緑（1L）・結線等チェック結果フラグ
 ,{189}          --P2・施工者・白（P2）・結線等チェック結果フラグ
 ,{190}          --P3・施工者・茶（3S）・結線等チェック結果フラグ
 ,{191}          --P3・施工者・青（P3）・結線等チェック結果フラグ
 ,{192}          --P3・施工者・黄（3L）・結線等チェック結果フラグ
 ,{193}          --変流器付計器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{194}          --変流器付計器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{195}          --変流器付計器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{196}          --変流器付計器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{197}          --変流器付計器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{198}          --変流器付計器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{199}          --変流器付計器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{200}          --P1・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{201}          --P1・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{202}          --P1・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{203}          --P2・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{204}          --P3・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{205}          --P3・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{206}          --P3・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{207}          --変流器付計器・検査担当者・黒(1S)・結線等チェック結果フラグ
 ,{208}          --変流器付計器・検査担当者・赤(P1)・結線等チェック結果フラグ
 ,{209}          --変流器付計器・検査担当者・青(P3)・結線等チェック結果フラグ
 ,{210}          --変流器付計器・検査担当者・茶(3S)・結線等チェック結果フラグ
 ,{211}          --変流器付計器・検査担当者・黄(3L)・結線等チェック結果フラグ
 ,{212}          --変流器付計器・検査担当者・白(P2)・結線等チェック結果フラグ
 ,{213}          --変流器付計器・検査担当者・緑(1L)・結線等チェック結果フラグ
 ,{214}          --P1・検査担当者・黒(1S)・結線等チェック結果フラグ
 ,{215}          --P1・検査担当者・赤(P1)・結線等チェック結果フラグ
 ,{216}          --P1・検査担当者・緑(1L)・結線等チェック結果フラグ
 ,{217}          --P2・検査担当者・白(P2)・結線等チェック結果フラグ
 ,{218}          --P3・検査担当者・茶(3S)・結線等チェック結果フラグ
 ,{219}          --P3・検査担当者・青(P3)・結線等チェック結果フラグ
 ,{220}          --P3・検査担当者・黄(3L)・結線等チェック結果フラグ
 ,{221}          --低圧計器チェック・計器回転・回転・SM表示・施工者・良否区分
 ,{222}          --低圧計器チェック・計器回転・回転・SM表示・抜取検査担当者・良否区分
 ,{223}          --低圧計器チェック・計器回転・回転・SM表示・検査担当者・良否区分
 ,{224}          --低圧計器チェック・計器回転・回転・SM表示・対応内容
 ,{225}          --低圧計器チェック・計器回転・回転・SM表示・対応完了日
 ,{226}          --低圧計器チェック・計器回転・回転・D2SM・施工者・良否区分
 ,{227}          --低圧計器チェック・計器回転・回転・D2SM・抜取検査担当者・良否区分
 ,{228}          --低圧計器チェック・計器回転・回転・D2SM・検査担当者・良否区分
 ,{229}          --低圧計器チェック・計器回転・回転・D2SM・対応内容
 ,{230}          --低圧計器チェック・計器回転・回転・D2SM・対応完了日
 ,{231}          --低圧計器チェック・計器回転・回転・開閉器・施工者・良否区分
 ,{232}          --低圧計器チェック・計器回転・回転・開閉器・抜取検査担当者・良否区分
 ,{233}          --低圧計器チェック・計器回転・回転・開閉器・検査担当者・良否区分
 ,{234}          --低圧計器チェック・計器回転・回転・開閉器・対応内容
 ,{235}          --低圧計器チェック・計器回転・回転・開閉器・対応完了日
 ,{236}          --低圧計器チェック・構造検査・計量器箱・屋外計器箱・施工者・良否区分
 ,{237}          --低圧計器チェック・構造検査・計量器箱・屋外計器箱・抜取検査担当者・良否区分
 ,{238}          --低圧計器チェック・構造検査・計量器箱・屋外計器箱・検査担当者・良否区分
 ,{239}          --低圧計器チェック・構造検査・計量器箱・屋外計器箱・対応内容
 ,{240}          --低圧計器チェック・構造検査・計量器箱・屋外計器箱・対応完了日
 ,{241}          --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・施工者・良否区分
 ,{242}          --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・抜取検査担当者・良否区分
 ,{243}          --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・検査担当者・良否区分
 ,{244}          --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・対応内容
 ,{245}          --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・対応完了日
)
         ]]></sql>.Value)

            Dim procDateTime As String = getDbDateTime()

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_ntktt_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_kstts_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_ntktt_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_kstts_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_skosya_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_ntktt_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_kstts_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_tio_cmp_ymd"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
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

#Region "(SQL)取替_自主点検チェックシートテーブル(施工)更新"
    ''' <summary>
    ''' 取替_自主点検チェックシートテーブル(施工)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシートテーブル(施工)を更新します</remarks>
    Friend Function U004(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev040 2024/12/25 物理分割 託送切替対応 ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_skosya_ryohi_kbn
            sb.Append(<sql><![CDATA[
UPDATE trke_check_sheet SET         -- 取替_自主点検チェックシート
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,takcd_ttt_skosya_ryohi_kbn={7}      --低圧計器チェック・電力量計・取付取替等・施工者・良否区分
 ,takcd_ttt_tio_naiyo={8}      --低圧計器チェック・電力量計・取付取替等・対応内容
 ,takcd_ttt_tio_cmp_ymd={9}      --低圧計器チェック・電力量計・取付取替等・対応完了日
 ,takcd_szsu_skosya_ryohi_kbn={10}      --低圧計器チェック・電力量計・指示数・施工者・良否区分
 ,takcd_szsu_tio_naiyo={11}      --低圧計器チェック・電力量計・指示数・対応内容
 ,takcd_szsu_tio_cmp_ymd={12}      --低圧計器チェック・電力量計・指示数・対応完了日
 ,takck_ssda_skosya_ryohi_kbn={13}      --低圧計器チェック・計器項目・相線電圧・施工者・良否区分
 ,takck_ssda_tio_naiyo={14}      --低圧計器チェック・計器項目・相線電圧・対応内容
 ,takck_ssda_tio_cmp_ymd={15}      --低圧計器チェック・計器項目・相線電圧・対応完了日
 ,takck_zyrt_skosya_ryohi_kbn={16}      --低圧計器チェック・計器項目・乗率・施工者・良否区分
 ,takck_zyrt_tio_naiyo={17}      --低圧計器チェック・計器項目・乗率・対応内容
 ,takck_zyrt_tio_cmp_ymd={18}      --低圧計器チェック・計器項目・乗率・対応完了日
 ,takck_yr_skosya_ryohi_kbn={19}      --低圧計器チェック・計器項目・容量・施工者・良否区分
 ,takck_yr_tio_naiyo={20}      --低圧計器チェック・計器項目・容量・対応内容
 ,takck_yr_tio_cmp_ymd={21}      --低圧計器チェック・計器項目・容量・対応完了日
 ,takck_sm_skosya_ryohi_kbn={22}      --低圧計器チェック・計器項目・SM・施工者・良否区分
 ,takck_sm_tio_naiyo={23}      --低圧計器チェック・計器項目・SM・対応内容
 ,takck_sm_tio_cmp_ymd={24}      --低圧計器チェック・計器項目・SM・対応完了日
 ,takck_kkano_skosya_ryohi_kbn={25}      --低圧計器チェック・計器項目・計器・合番号・施工者・良否区分
 ,takck_kkano_tio_naiyo={26}      --低圧計器チェック・計器項目・計器・合番号・対応内容
 ,takck_kkano_tio_cmp_ymd={27}      --低圧計器チェック・計器項目・計器・合番号・対応完了日
 ,takck_hksno_skosya_ryohi_kbn={28}      --低圧計器チェック・計器項目・変流器・製造番号・施工者・良否区分
 ,takck_hksno_tio_naiyo={29}      --低圧計器チェック・計器項目・変流器・製造番号・対応内容
 ,takck_hksno_tio_cmp_ymd={30}      --低圧計器チェック・計器項目・変流器・製造番号・対応完了日
 ,takck_hkano_skosya_ryohi_kbn={31}      --低圧計器チェック・計器項目・変流器・合番号・施工者・良否区分
 ,takck_hkano_tio_naiyo={32}      --低圧計器チェック・計器項目・変流器・合番号・対応内容
 ,takck_hkano_tio_cmp_ymd={33}      --低圧計器チェック・計器項目・変流器・合番号・対応完了日
 ,takckk_krkkd_skosya_ryohi_kbn={34}      --低圧計器チェック・構造検査・計量器・検電・施工者・良否区分
 ,takckk_krkkd_tio_naiyo={35}      --低圧計器チェック・構造検査・計量器・検電・対応内容
 ,takckk_krkkd_tio_cmp_ymd={36}      --低圧計器チェック・構造検査・計量器・検電・対応完了日
 ,takckk_krkb_skosya_ryohi_kbn={37}      --低圧計器チェック・構造検査・計量器箱・施工者・良否区分
 ,takckk_krkb_tio_naiyo={38}      --低圧計器チェック・構造検査・計量器箱・対応内容
 ,takckk_krkb_tio_cmp_ymd={39}      --低圧計器チェック・構造検査・計量器箱・対応完了日
 ,takcks_kkzen_skosya_ryohi_kbn={40}      --低圧計器チェック・結線等・計器・絶縁・施工者・良否区分
 ,takcks_kkzen_tio_naiyo={41}      --低圧計器チェック・結線等・計器・絶縁・対応内容
 ,takcks_kkzen_tio_cmp_ymd={42}      --低圧計器チェック・結線等・計器・絶縁・対応完了日
 ,takcks_kksl_skosya_ryohi_kbn={43}      --低圧計器チェック・結線等・計器・SL・施工者・良否区分
 ,takcks_kksl_tio_naiyo={44}      --低圧計器チェック・結線等・計器・SL・対応内容
 ,takcks_kksl_tio_cmp_ymd={45}      --低圧計器チェック・結線等・計器・SL・対応完了日
 ,takcks_kkhks_skosya_ryohi_kbn={46}      --低圧計器チェック・結線等・計器・複数・施工者・良否区分
 ,takcks_kkhks_tio_naiyo={47}      --低圧計器チェック・結線等・計器・複数・対応内容
 ,takcks_kkhks_tio_cmp_ymd={48}      --低圧計器チェック・結線等・計器・複数・対応完了日
 ,takcks_kkgsz_skosya_ryohi_kbn={49}      --低圧計器チェック・結線等・計器・誤接続・施工者・良否区分
 ,takcks_kkgsz_tio_naiyo={50}      --低圧計器チェック・結線等・計器・誤接続・対応内容
 ,takcks_kkgsz_tio_cmp_ymd={51}      --低圧計器チェック・結線等・計器・誤接続・対応完了日
 ,takcks_kkaki_skosya_ryohi_kbn={52}      --低圧計器チェック・結線等・計器・赤色・施工者・良否区分
 ,takcks_kkaki_tio_naiyo={53}      --低圧計器チェック・結線等・計器・赤色・対応内容
 ,takcks_kkaki_tio_cmp_ymd={54}      --低圧計器チェック・結線等・計器・赤色・対応完了日
 ,takcks_kkyrs_skosya_ryohi_kbn={55}      --低圧計器チェック・結線等・計器・より線・施工者・良否区分
 ,takcks_kkyrs_tio_naiyo={56}      --低圧計器チェック・結線等・計器・より線・対応内容
 ,takcks_kkyrs_tio_cmp_ymd={57}      --低圧計器チェック・結線等・計器・より線・対応完了日
 ,takcks_kkrst_skosya_ryohi_kbn={58}      --低圧計器チェック・結線等・計器・露出・施工者・良否区分
 ,takcks_kkrst_tio_naiyo={59}      --低圧計器チェック・結線等・計器・露出・対応内容
 ,takcks_kkrst_tio_cmp_ymd={60}      --低圧計器チェック・結線等・計器・露出・対応完了日
 ,takcks_kkbst_skosya_ryohi_kbn={61}      --低圧計器チェック・結線等・計器・ビス締付・施工者・良否区分
 ,takcks_kkbst_tio_naiyo={62}      --低圧計器チェック・結線等・計器・ビス締付・対応内容
 ,takcks_kkbst_tio_cmp_ymd={63}      --低圧計器チェック・結線等・計器・ビス締付・対応完了日
 ,takcks_kikib_skosya_ryohi_kbn={64}      --低圧計器チェック・結線等・計器・色別・施工者・良否区分
 ,takcks_kikib_tio_naiyo={65}      --低圧計器チェック・結線等・計器・色別・対応内容
 ,takcks_kikib_tio_cmp_ymd={66}      --低圧計器チェック・結線等・計器・色別・対応完了日
 ,takcks_kkssu_skosya_ryohi_kbn={67}      --低圧計器チェック・結線等・計器・正相・施工者・良否区分
 ,takcks_kkssu_tio_naiyo={68}      --低圧計器チェック・結線等・計器・正相・対応内容
 ,takcks_kkssu_tio_cmp_ymd={69}      --低圧計器チェック・結線等・計器・正相・対応完了日
 ,takcks_hrkkl_skosya_ryohi_kbn={70}      --低圧計器チェック・結線等・変流器・ＫＬ・施工者・良否区分
 ,takcks_hrkkl_tio_naiyo={71}      --低圧計器チェック・結線等・変流器・ＫＬ・対応内容
 ,takcks_hrkkl_tio_cmp_ymd={72}      --低圧計器チェック・結線等・変流器・ＫＬ・対応完了日
 ,takcks_hp123_skosya_ryohi_kbn={73}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・施工者・良否区分
 ,takcks_hp123_tio_naiyo={74}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応内容
 ,takcks_hp123_tio_cmp_ymd={75}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応完了日
 ,takcks_hkaki_skosya_ryohi_kbn={76}      --低圧計器チェック・結線等・変流器・赤色・施工者・良否区分
 ,takcks_hkaki_tio_naiyo={77}      --低圧計器チェック・結線等・変流器・赤色・対応内容
 ,takcks_hkaki_tio_cmp_ymd={78}      --低圧計器チェック・結線等・変流器・赤色・対応完了日
 ,takckrk_hnrk_skosya_ryohi_kbn={79}      --低圧計器チェック・計器回転・検電・変流器・施工者・良否区分
 ,takckrk_hnrk_tio_naiyo={80}      --低圧計器チェック・計器回転・検電・変流器・対応内容
 ,takckrk_hnrk_tio_cmp_ymd={81}      --低圧計器チェック・計器回転・検電・変流器・対応完了日
 ,takckrr_sm_skosya_ryohi_kbn={82}      --低圧計器チェック・計器回転・回転・SM・施工者・良否区分
 ,takckrr_sm_tio_naiyo={83}      --低圧計器チェック・計器回転・回転・SM・対応内容
 ,takckrr_sm_tio_cmp_ymd={84}      --低圧計器チェック・計器回転・回転・SM・対応完了日
 ,takcc_led_skosya_ryohi_kbn={85}      --低圧計器チェック・通信部・ＬＥＤ・施工者・良否区分
 ,takcc_led_tio_naiyo={86}      --低圧計器チェック・通信部・ＬＥＤ・対応内容
 ,takcc_led_tio_cmp_ymd={87}      --低圧計器チェック・通信部・ＬＥＤ・対応完了日
 ,takcc_souck_skosya_ryohi_kbn={88}      --低圧計器チェック・通信部・装着・施工者・良否区分
 ,takcc_souck_tio_naiyo={89}      --低圧計器チェック・通信部・装着・対応内容
 ,takcc_souck_tio_cmp_ymd={90}      --低圧計器チェック・通信部・装着・対応完了日
 ,takct_szno_skosya_ryohi_kbn={91}      --低圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,takct_szno_tio_naiyo={92}      --低圧計器チェック・撤去計器等・製造番号・対応内容
 ,takct_szno_tio_cmp_ymd={93}      --低圧計器チェック・撤去計器等・製造番号・対応完了日
 ,takct_aino_skosya_ryohi_kbn={94}      --低圧計器チェック・撤去計器等・合番号・施工者・良否区分
 ,takct_aino_tio_naiyo={95}      --低圧計器チェック・撤去計器等・合番号・対応内容
 ,takct_aino_tio_cmp_ymd={96}      --低圧計器チェック・撤去計器等・合番号・対応完了日
 ,takcs_photo_skosya_ryohi_kbn={97}      --低圧計器チェック・その他・写真・施工者・良否区分
 ,takcs_photo_tio_naiyo={98}      --低圧計器チェック・その他・写真・対応内容
 ,takcs_photo_tio_cmp_ymd={99}      --低圧計器チェック・その他・写真・対応完了日
 ,takcs_tktrk_skosya_ryohi_kbn={100}      --低圧計器チェック・その他・撤去取替・施工者・良否区分
 ,takcs_tktrk_tio_naiyo={101}      --低圧計器チェック・その他・撤去取替・対応内容
 ,takcs_tktrk_tio_cmp_ymd={102}      --低圧計器チェック・その他・撤去取替・対応完了日
 ,takcz_ntktt_seigo_kknn_kbn={103}      --低圧計器チェック・乗率チェック・抜取検査担当者・整合確認区分
 ,takcz_tio_cmp_ymd={104}      --低圧計器チェック・乗率チェック・対応完了日
 ,takca_ntktt_itti_kknn_kbn={105}      --低圧計器チェック・合番号チェック・抜取検査担当者・一致確認区分
 ,takca_kstts_itti_kknn_kbn={106}      --低圧計器チェック・合番号チェック・検査担当者・一致確認区分
 ,takca_tio_naiyo={107}      --低圧計器チェック・合番号チェック・対応内容
 ,takca_tio_cmp_ymd={108}      --低圧計器チェック・合番号チェック・対応完了日
 ,kizi_skosya_naiyo={109}      --記事（施工者）内容
 ,hrktkk_skosya_1s_kstckk_flg={110}      --変流器付計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,hrktkk_skosya_p1_kstckk_flg={111}      --変流器付計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,hrktkk_skosya_p3_kstckk_flg={112}      --変流器付計器・施工者・青（P3）・結線等チェック結果フラグ
 ,hrktkk_skosya_3s_kstckk_flg={113}      --変流器付計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,hrktkk_skosya_3l_kstckk_flg={114}      --変流器付計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,hrktkk_skosya_p2_kstckk_flg={115}      --変流器付計器・施工者・白（P2）・結線等チェック結果フラグ
 ,hrktkk_skosya_1l_kstckk_flg={116}      --変流器付計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,p1_skosya_1s_kstckk_flg={117}      --P1・施工者・黒（1S）・結線等チェック結果フラグ
 ,p1_skosya_p1_kstckk_flg={118}      --P1・施工者・赤（P1）・結線等チェック結果フラグ
 ,p1_skosya_1l_kstckk_flg={119}      --P1・施工者・緑（1L）・結線等チェック結果フラグ
 ,p2_skosya_p2_kstckk_flg={120}      --P2・施工者・白（P2）・結線等チェック結果フラグ
 ,p3_skosya_3s_kstckk_flg={121}      --P3・施工者・茶（3S）・結線等チェック結果フラグ
 ,p3_skosya_p3_kstckk_flg={122}      --P3・施工者・青（P3）・結線等チェック結果フラグ
 ,p3_skosya_3l_kstckk_flg={123}      --P3・施工者・黄（3L）・結線等チェック結果フラグ
 ,takckrr_smhyz_skosya_ryohi_kbn={124}   --低圧計器チェック・計器回転・回転・SM表示・施工者・良否区分
 ,takckrr_d2sm_skosya_ryohi_kbn={125}    --低圧計器チェック・計器回転・回転・D2SM・施工者・良否区分
 ,takckrr_khk_skosya_ryohi_kbn={126}     --低圧計器チェック・計器回転・回転・開閉器・施工者・良否区分
 ,takckk_okgkb_skosya_ryohi_kbn={127}    --低圧計器チェック・構造検査・計量器箱・屋外計器箱・施工者・良否区分
 ,takckk_dsipcv_skosya_ryohi_kbn={128}   --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・施工者・良否区分
WHERE 1 = 1
 AND  dig4_zgsyo_cd={3}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={4}      --取替票発行年度
 AND  trkhy_kbn={5}      --取替票区分
 AND  trkhy_no={6}      --取替票番号         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                               EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_ntktt_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_ntktt_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_kstts_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_skosya_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_skosya_ryohi_kbn"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "添付書類詳細　ファイル連番の最大値検索"
    ''' <summary>
    ''' 添付書類詳細　ファイル連番の最大値を検索する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S005(ByRef dbr As HdPostgre.HdPostgreReader, ByVal k1 As String, ByVal k2 As String, ByVal k3 As String, ByVal k4 As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
SELECT
    MAX(file_renno) as max_file_renno
FROM
    trke_trkehy_tiat A
    left join COM_TNPDOC_DTL B ON (B.tnpdoc_mng_renno=A.tenp_file_mng_no)
WHERE 1 = 1
    AND dig4_zgsyo_cd = {0}
    AND trkhy_hakko_nendo = {1}
    AND trkhy_kbn =  {2}
    AND trkhy_no = {3}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , EcOrgIO.EcOrgString.GetSqlText(k1) _
                               , EcOrgIO.EcOrgString.GetSqlText(k2) _
                               , EcOrgIO.EcOrgString.GetSqlText(k3) _
                               , EcOrgIO.EcOrgString.GetSqlText(k4))
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

    ''' <summary>
    ''' 添付書類詳細　最大ファイル連番取得
    ''' </summary>
    ''' <param name="pFileRenno">添付書類詳細　最大ファイル連番</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <remarks></remarks>
    Function getMaxFileRenno(ByRef pFileRenno As String, ByVal k1 As String, ByVal k2 As String, ByVal k3 As String, ByVal k4 As String) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        pFileRenno = "0"

        '添付書類詳細　最大ファイル連番を検索する。
        If Not S005(dbr, k1, k2, k3, k4) Then
            Return False
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                pFileRenno = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("max_file_renno"), ""))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return True
    End Function
#End Region

#Region "(SQL)取替票（低圧）添付ファイル管理番号更新"
    ''' <summary>
    ''' 取替票（低圧）添付ファイル管理番号更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pTtenpFileMngNo">添付ファイル管理番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）添付ファイル管理番号を更新します</remarks>
    Friend Function U005(ByVal pRow As DataRow, ByVal pTtenpFileMngNo As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  upd_date = {0}                   -- 更新日時
 ,sousa_user_id = {1}              -- 更新ユーザ
 ,sousa_appli_cd = {2}             -- 更新プログラム
 ,tenp_file_mng_no={3}      --添付ファイル管理番号
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}      --取替票発行年度
 AND  trkhy_kbn={6}      --取替票区分
 AND  trkhy_no={7}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pTtenpFileMngNo) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)取替票（低圧）領収書添付ファイル管理番号更新"
    ''' <summary>
    ''' 取替票（低圧）領収書添付ファイル管理番号更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pRyssyTenpFileMngNo">領収書添付ファイル管理番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）領収書添付ファイル管理番号を更新します</remarks>
    Friend Function U006(ByVal pRow As DataRow, ByVal pRyssyTenpFileMngNo As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  upd_date = {0}                   -- 更新日時
 ,sousa_user_id = {1}              -- 更新ユーザ
 ,sousa_appli_cd = {2}             -- 更新プログラム
 ,ryssy_tenp_file_mng_no={3}      --領収書添付ファイル管理番号
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}      --取替票発行年度
 AND  trkhy_kbn={6}      --取替票区分
 AND  trkhy_no={7}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pRyssyTenpFileMngNo) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 取替票（低圧）実費_No1更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）実費_No1を更新します</remarks>
    Friend Function U0061(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  no1_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取替票（低圧）実費_No2更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）実費_No2を更新します</remarks>
    Friend Function U0062(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  no2_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取替票（低圧）実費_No3更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）実費_No3を更新します</remarks>
    Friend Function U0063(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  no3_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function
#End Region



#Region "(SQL)取替票（低圧）(抜取検査)更新"
    ''' <summary>
    ''' 取替票（低圧）(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pUpldUserId">アップロードUSER_ID</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）(抜取検査)を更新します</remarks>
    Friend Function U007(ByVal pRow As DataRow, ByVal pUpldUserId As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,nktr_kensa_zyokyo_cd={3}      --抜取検査状況コード
 ,ntkns_tblt_recv_ymd={4}      --抜取検査タブレット受信年月日
 ,ntkns_tblt_recv_tntsy_cd={5}      --抜取検査タブレット受信_担当者コード
 ,ntkns_zissi_ymd={6}      --抜取検査実施年月日
 ,ntkns_zissi_tntsy_cd={7}      --抜取検査実施_担当者コード
 ,ntkns_kka_upld_ymd={8}      --抜取検査結果アップロード年月日
 ,ntkns_kka_upld_tntsy_cd={9}      --抜取検査結果アップロード_担当者コード
WHERE 1 = 1
 AND  dig4_zgsyo_cd={10}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={11}      --取替票発行年度
 AND  trkhy_kbn={12}      --取替票区分
 AND  trkhy_no={13}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktr_kensa_zyokyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_tblt_recv_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_tblt_recv_tntsy_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_zissi_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_zissi_tntsy_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(pUpldUserId) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)取替_取替票行程(抜取検査)更新"
    ''' <summary>
    ''' 取替_取替票行程(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票行程(抜取検査)を更新します</remarks>
    Friend Function U008(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC SET         -- 取替行程
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,nktrk_sizi_syori_ymd={3}      --抜取検査指示_処理年月日
 ,nktrk_sizi_syors_cd={4}      --抜取検査指示_処理者コード
 ,nktrk_sizi_syors_ms={5}      --抜取検査指示_処理者名
 ,nktrk_sizi_ka_cd={6}      --抜取検査指示_課コード
 ,nktrk_sizi_ka_ms={7}      --抜取検査指示_課名
 ,nktrk_sizi_kzkis_mdgt_cd={8}      --抜取検査指示_工事会社窓口コード
 ,nktrk_sizi_kzkis_mdgt_ms={9}      --抜取検査指示_工事会社窓口名
 ,nktrk_sizi_kozitn_cd={10}      --抜取検査指示_工事店コード
 ,nktrk_sizi_kozitn_ms={11}      --抜取検査指示_工事店名
 ,nktrk_tanto_syori_ymd={12}      --抜取検査（担当）_処理年月日
 ,nktrk_tanto_syors_cd={13}      --抜取検査（担当）_処理者コード
 ,nktrk_tanto_syors_ms={14}      --抜取検査（担当）_処理者名
 ,nktrk_tanto_ka_cd={15}      --抜取検査（担当）_課コード
 ,nktrk_tanto_ka_ms={16}      --抜取検査（担当）_課名
 ,nktrk_tanto_ces_cd={17}      --抜取検査（担当）_保安協会コード
 ,nktrk_tanto_ces_ms={18}      --抜取検査（担当）_保安協会名
 ,nktrk_elder_syori_ymd={19}      --抜取検査（上長）_処理年月日
 ,nktrk_elder_syors_cd={20}      --抜取検査（上長）_処理者コード
 ,nktrk_elder_syors_ms={21}      --抜取検査（上長）_処理者名
 ,nktrk_elder_ka_cd={22}      --抜取検査（上長）_課コード
 ,nktrk_elder_ka_ms={23}      --抜取検査（上長）_課名
 ,nktrk_elder_ces_cd={24}      --抜取検査（上長）_保安協会コード
 ,nktrk_elder_ces_ms={25}      --抜取検査（上長）_保安協会名
 ,nktrk_hnn_flg={26}      --抜取検査_否認フラグ
WHERE 1 = 1
 AND  dig4_zgsyo_cd={27}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={28}      --取替票発行年度
 AND  trkhy_kbn={29}      --取替票区分
 AND  trkhy_no={30}      --取替票番号

         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kzkis_mdgt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kzkis_mdgt_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kozitn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kozitn_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ces_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ces_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ces_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ces_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_hnn_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)取替_自主点検チェックシートテーブル(抜取検査)更新"
    ''' <summary>
    ''' 取替_自主点検チェックシートテーブル(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシートテーブル(抜取検査)を更新します</remarks>
    Friend Function U009(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev040 2024/12/25 物理分割 託送切替対応 ADD takckrr_smhyz_ntktt_ryohi_kbn ～ takckk_dsipcv_ntktt_ryohi_kbn
            sb.Append(<sql><![CDATA[
UPDATE trke_check_sheet SET         -- 取替_自主点検チェックシート
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,takcd_ttt_ntktt_ryohi_kbn={3}      --低圧計器チェック・電力量計・取付取替等・抜取検査担当者・良否区分
 ,takcd_szsu_ntktt_ryohi_kbn={4}      --低圧計器チェック・電力量計・指示数・抜取検査担当者・良否区分
 ,takck_ssda_ntktt_ryohi_kbn={5}      --低圧計器チェック・計器項目・相線電圧・抜取検査担当者・良否区分
 ,takck_zyrt_ntktt_ryohi_kbn={6}      --低圧計器チェック・計器項目・乗率・抜取検査担当者・良否区分
 ,takck_yr_ntktt_ryohi_kbn={7}      --低圧計器チェック・計器項目・容量・抜取検査担当者・良否区分
 ,takck_sm_ntktt_ryohi_kbn={8}      --低圧計器チェック・計器項目・SM・抜取検査担当者・良否区分
 ,takck_kkano_ntktt_ryohi_kbn={9}      --低圧計器チェック・計器項目・計器・合番号・抜取検査担当者・良否区分
 ,takck_hksno_ntktt_ryohi_kbn={10}      --低圧計器チェック・計器項目・変流器・製造番号・抜取検査担当者・良否区分
 ,takck_hkano_ntktt_ryohi_kbn={11}      --低圧計器チェック・計器項目・変流器・合番号・抜取検査担当者・良否区分
 ,takckk_krkkd_ntktt_ryohi_kbn={12}      --低圧計器チェック・構造検査・計量器・検電・抜取検査担当者・良否区分
 ,takckk_krkb_ntktt_ryohi_kbn={13}      --低圧計器チェック・構造検査・計量器箱・抜取検査担当者・良否区分
 ,takcks_kkzen_ntktt_ryohi_kbn={14}      --低圧計器チェック・結線等・計器・絶縁・抜取検査担当者・良否区分
 ,takcks_kksl_ntktt_ryohi_kbn={15}      --低圧計器チェック・結線等・計器・SL・抜取検査担当者・良否区分
 ,takcks_kkhks_ntktt_ryohi_kbn={16}      --低圧計器チェック・結線等・計器・複数・抜取検査担当者・良否区分
 ,takcks_kkgsz_ntktt_ryohi_kbn={17}      --低圧計器チェック・結線等・計器・誤接続・抜取検査担当者・良否区分
 ,takcks_kkaki_ntktt_ryohi_kbn={18}      --低圧計器チェック・結線等・計器・赤色・抜取検査担当者・良否区分
 ,takcks_kkyrs_ntktt_ryohi_kbn={19}      --低圧計器チェック・結線等・計器・より線・抜取検査担当者・良否区分
 ,takcks_kkrst_ntktt_ryohi_kbn={20}      --低圧計器チェック・結線等・計器・露出・抜取検査担当者・良否区分
 ,takcks_kkbst_ntktt_ryohi_kbn={21}      --低圧計器チェック・結線等・計器・ビス締付・抜取検査担当者・良否区分
 ,takcks_kikib_ntktt_ryohi_kbn={22}      --低圧計器チェック・結線等・計器・色別・抜取検査担当者・良否区分
 ,takcks_kkssu_ntktt_ryohi_kbn={23}      --低圧計器チェック・結線等・計器・正相・抜取検査担当者・良否区分
 ,takcks_hrkkl_ntktt_ryohi_kbn={24}      --低圧計器チェック・結線等・変流器・ＫＬ・抜取検査担当者・良否区分
 ,takcks_hp123_ntktt_ryohi_kbn={25}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・抜取検査担当者・良否区分
 ,takcks_hkaki_ntktt_ryohi_kbn={26}      --低圧計器チェック・結線等・変流器・赤色・抜取検査担当者・良否区分
 ,takckrk_hnrk_ntktt_ryohi_kbn={27}      --低圧計器チェック・計器回転・検電・変流器・抜取検査担当者・良否区分
 ,takckrr_sm_ntktt_ryohi_kbn={28}      --低圧計器チェック・計器回転・回転・SM・抜取検査担当者・良否区分
 ,takcc_led_ntktt_ryohi_kbn={29}      --低圧計器チェック・通信部・ＬＥＤ・抜取検査担当者・良否区分
 ,takcc_souck_ntktt_ryohi_kbn={30}      --低圧計器チェック・通信部・装着・抜取検査担当者・良否区分
 ,takct_szno_ntktt_ryohi_kbn={31}      --低圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,takct_aino_ntktt_ryohi_kbn={32}      --低圧計器チェック・撤去計器等・合番号・抜取検査担当者・良否区分
 ,takcs_photo_ntktt_ryohi_kbn={33}      --低圧計器チェック・その他・写真・抜取検査担当者・良否区分
 ,takcs_tktrk_ntktt_ryohi_kbn={34}      --低圧計器チェック・その他・撤去取替・抜取検査担当者・良否区分
 ,takcz_ntktt_seigo_kknn_kbn={35}      --低圧計器チェック・乗率チェック・抜取検査担当者・整合確認区分
 ,takca_ntktt_itti_kknn_kbn={36}      --低圧計器チェック・合番号チェック・抜取検査担当者・一致確認区分
 ,kizi_ntktt_naiyo={37}      --記事（抜取検査担当者）内容
 ,hrktkk_ntktt_1s_kstckk_flg={38}      --変流器付計器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p1_kstckk_flg={39}      --変流器付計器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p3_kstckk_flg={40}      --変流器付計器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,hrktkk_ntktt_3s_kstckk_flg={41}      --変流器付計器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,hrktkk_ntktt_3l_kstckk_flg={42}      --変流器付計器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p2_kstckk_flg={43}      --変流器付計器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,hrktkk_ntktt_1l_kstckk_flg={44}      --変流器付計器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,p1_ntktt_1s_kstckk_flg={45}      --P1・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,p1_ntktt_p1_kstckk_flg={46}      --P1・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,p1_ntktt_1l_kstckk_flg={47}      --P1・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,p2_ntktt_p2_kstckk_flg={48}      --P2・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,p3_ntktt_3s_kstckk_flg={49}      --P3・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,p3_ntktt_p3_kstckk_flg={50}      --P3・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,p3_ntktt_3l_kstckk_flg={51}      --P3・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,takckrr_smhyz_ntktt_ryohi_kbn={52}   --低圧計器チェック・計器回転・回転・SM表示・抜取検査担当者・良否区分
 ,takckrr_d2sm_ntktt_ryohi_kbn={53}    --低圧計器チェック・計器回転・回転・D2SM・抜取検査担当者・良否区分
 ,takckrr_khk_ntktt_ryohi_kbn={54}     --低圧計器チェック・計器回転・回転・開閉器・抜取検査担当者・良否区分
 ,takckk_okgkb_ntktt_ryohi_kbn={55}    --低圧計器チェック・構造検査・計量器箱・屋外計器箱・抜取検査担当者・良否区分
 ,takckk_dsipcv_ntktt_ryohi_kbn={56}   --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・抜取検査担当者・良否区分
WHERE 1 = 1
 AND  dig4_zgsyo_cd={57}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={58}      --取替票発行年度
 AND  trkhy_kbn={59}      --取替票区分
 AND  trkhy_no={60}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_ntktt_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_ntktt_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_ntktt_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region


#Region "(SQL)取替票（低圧）(追加工費および工事費合計)更新"
    ''' <summary>
    ''' 取替票（低圧）(追加工費および工事費合計)更新
    ''' </summary>
    ''' <param name="dig4_zgsyo_cd">事業所コード（4桁）</param>
    ''' <param name="trkhy_hakko_nendo">取替票発行年度</param>
    ''' <param name="trkhy_kbn">取替票区分</param>
    ''' <param name="trkhy_no">取替票番号</param>
    ''' <param name="TUIKAKH_umu_flg">追加工費有無フラグ</param>
    ''' <param name="trkhy_trtkKojiHi">追加工費　取付工事費</param>
    ''' <param name="trkhy_tekyoKojiHi">追加工費　撤去工事費</param>
    ''' <param name="gokei_kngk">工事費合計</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）(追加工費および工事費合計)を更新します</remarks>
    Friend Function U010(ByVal dig4_zgsyo_cd As String, ByVal trkhy_hakko_nendo As String, ByVal trkhy_kbn As String, ByVal trkhy_no As String,
                         ByVal TUIKAKH_umu_flg As String,
                         ByVal trkhy_trtkKojiHi As String, ByVal trkhy_tekyoKojiHi As String, ByVal gokei_kngk As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_tiat SET        -- 取替票（低圧）
  tuika_kohi_trtk_kozih_kngk={0}      --追加工費_取付工事費
 ,tuika_kohi_tekyo_kozih_kngk={1}      --追加工費_撤去工事費
 ,tuika_kohi_umu_flg={2}      --追加工費有無フラグ
 ,kozih_gokei_kngk={3}      --工事費合計
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}      --取替票発行年度
 AND  trkhy_kbn={6}      --取替票区分
 AND  trkhy_no={7}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(trkhy_trtkKojiHi) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(trkhy_tekyoKojiHi) _
                               , EcOrgIO.EcOrgString.GetSqlText(TUIKAKH_umu_flg) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(gokei_kngk) _
                               , EcOrgIO.EcOrgString.GetSqlText(dig4_zgsyo_cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_hakko_nendo) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_no)
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)取替票（高圧他）添付ファイル管理番号更新"
    ''' <summary>
    ''' 取替票（高圧他）添付ファイル管理番号更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pTtenpFileMngNo">添付ファイル管理番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）添付ファイル管理番号を更新します</remarks>
    Friend Function U014(ByVal pRow As DataRow, ByVal pTtenpFileMngNo As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  upd_date = {0}                   -- 更新日時
 ,sousa_user_id = {1}              -- 更新ユーザ
 ,sousa_appli_cd = {2}             -- 更新プログラム
 ,tenp_file_mng_no={3}      --添付ファイル管理番号
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}      --取替票発行年度
 AND  trkhy_kbn={6}      --取替票区分
 AND  trkhy_no={7}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pTtenpFileMngNo) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region

#Region "(SQL)計取替_取替票追加工費（高圧他）削除"
    ''' <summary>
    ''' 取替_取替票追加工費（高圧他）削除(DataRow)
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を削除します</remarks>
    Friend Function D002(ByVal pRow As DataRow) As Boolean
        Return D002(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), "")), CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), "")),
                    CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), "")), CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
    End Function

    ''' <summary>
    ''' 取替_取替票追加工費（高圧他）削除
    ''' </summary>
    ''' <param name="dig4_zgsyo_cd">事業所コード（4桁）</param>
    ''' <param name="trkhy_hakko_nendo">取替票発行年度</param>
    ''' <param name="trkhy_kbn">取替票区分</param>
    ''' <param name="trkhy_no">取替票番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票追加工費を削除します</remarks>
    Friend Function D002(ByVal dig4_zgsyo_cd As String, ByVal trkhy_hakko_nendo As String, ByVal trkhy_kbn As String, ByVal trkhy_no As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
DELETE from trke_trkehy_tuikakh_kahk        -- 取替_取替票追加工費（高圧他）
WHERE 1 = 1
 AND  dig4_zgsyo_cd={0}         --事業所コード（4桁）
 AND  trkhy_hakko_nendo={1}     --取替票発行年度
 AND  trkhy_kbn={2}             --取替票区分
 AND  trkhy_no={3}              --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(dig4_zgsyo_cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_hakko_nendo) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_no)
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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




#Region "(SQL)取替_自主点検チェックシート（高圧他）追加(全項目)"
    ''' <summary>
    ''' 取替_自主点検チェックシート（高圧他）追加
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシート（高圧他）を追加します</remarks>
    Friend Function I010(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_tio_cmp_ymd
            sb.Append(<sql><![CDATA[
INSERT INTO trke_check_sheet_kahk        -- 取替_自主点検チェックシート（高圧他）
VALUES (
  {0}       --作成日時
 ,{1}       --更新日時
 ,{2}       --更新ユーザ
 ,{3}       --更新プログラム
 ,{4}       --事業所コード（4桁）
 ,{5}       --取替票発行年度
 ,{6}       --取替票区分
 ,{7}       --取替票番号
 ,{8}       --点検（検査）年月日・施工者
 ,{9}       --点検（検査）年月日・検査担当者
 ,{10}      --高圧計器チェック・変成器・容量・新設・施工者・良否区分
 ,{11}      --高圧計器チェック・変成器・容量・新設・抜取検査担当者・良否区分
 ,{12}      --高圧計器チェック・変成器・容量・新設・検査担当者・良否区分
 ,{13}      --高圧計器チェック・変成器・容量・新設・対応内容
 ,{14}      --高圧計器チェック・変成器・容量・新設・対応完了日
 ,{15}      --高圧計器チェック・変成器・容量・既設・施工者・良否区分
 ,{16}      --高圧計器チェック・変成器・容量・既設・抜取検査担当者・良否区分
 ,{17}      --高圧計器チェック・変成器・容量・既設・検査担当者・良否区分
 ,{18}      --高圧計器チェック・変成器・容量・既設・対応内容
 ,{19}      --高圧計器チェック・変成器・容量・既設・対応完了日
 ,{20}      --高圧計器チェック・変成器・合番号・新設・施工者・良否区分
 ,{21}      --高圧計器チェック・変成器・合番号・新設・抜取検査担当者・良否区分
 ,{22}      --高圧計器チェック・変成器・合番号・新設・検査担当者・良否区分
 ,{23}      --高圧計器チェック・変成器・合番号・新設・対応内容
 ,{24}      --高圧計器チェック・変成器・合番号・新設・対応完了日
 ,{25}      --高圧計器チェック・変成器・合番号・既設・施工者・良否区分
 ,{26}      --高圧計器チェック・変成器・合番号・既設・抜取検査担当者・良否区分
 ,{27}      --高圧計器チェック・変成器・合番号・既設・検査担当者・良否区分
 ,{28}      --高圧計器チェック・変成器・合番号・既設・対応内容
 ,{29}      --高圧計器チェック・変成器・合番号・既設・対応完了日
 ,{30}      --高圧計器チェック・変成器・製造番号・施工者・良否区分
 ,{31}      --高圧計器チェック・変成器・製造番号・抜取検査担当者・良否区分
 ,{32}      --高圧計器チェック・変成器・製造番号・検査担当者・良否区分
 ,{33}      --高圧計器チェック・変成器・製造番号・対応内容
 ,{34}      --高圧計器チェック・変成器・製造番号・対応完了日
 ,{35}      --高圧計器チェック・変成器・結線・KL・施工者・良否区分
 ,{36}      --高圧計器チェック・変成器・結線・KL・抜取検査担当者・良否区分
 ,{37}      --高圧計器チェック・変成器・結線・KL・検査担当者・良否区分
 ,{38}      --高圧計器チェック・変成器・結線・KL・対応内容
 ,{39}      --高圧計器チェック・変成器・結線・KL・対応完了日
 ,{40}      --高圧計器チェック・変成器・結線・正相・施工者・良否区分
 ,{41}      --高圧計器チェック・変成器・結線・正相・抜取検査担当者・良否区分
 ,{42}      --高圧計器チェック・変成器・結線・正相・検査担当者・良否区分
 ,{43}      --高圧計器チェック・変成器・結線・正相・対応内容
 ,{44}      --高圧計器チェック・変成器・結線・正相・対応完了日
 ,{45}      --高圧計器チェック・変成器・結線・色別・施工者・良否区分
 ,{46}      --高圧計器チェック・変成器・結線・色別・抜取検査担当者・良否区分
 ,{47}      --高圧計器チェック・変成器・結線・色別・検査担当者・良否区分
 ,{48}      --高圧計器チェック・変成器・結線・色別・対応内容
 ,{49}      --高圧計器チェック・変成器・結線・色別・対応完了日
 ,{50}      --高圧計器チェック・変成器・結線・半田・施工者・良否区分
 ,{51}      --高圧計器チェック・変成器・結線・半田・抜取検査担当者・良否区分
 ,{52}      --高圧計器チェック・変成器・結線・半田・検査担当者・良否区分
 ,{53}      --高圧計器チェック・変成器・結線・半田・対応内容
 ,{54}      --高圧計器チェック・変成器・結線・半田・対応完了日
 ,{55}      --高圧計器チェック・変成器・結線・ビス締付・施工者・良否区分
 ,{56}      --高圧計器チェック・変成器・結線・ビス締付・抜取検査担当者・良否区分
 ,{57}      --高圧計器チェック・変成器・結線・ビス締付・検査担当者・良否区分
 ,{58}      --高圧計器チェック・変成器・結線・ビス締付・対応内容
 ,{59}      --高圧計器チェック・変成器・結線・ビス締付・対応完了日
 ,{60}      --高圧計器チェック・計器・乗率・新設・施工者・良否区分
 ,{61}      --高圧計器チェック・計器・乗率・新設・抜取検査担当者・良否区分
 ,{62}      --高圧計器チェック・計器・乗率・新設・検査担当者・良否区分
 ,{63}      --高圧計器チェック・計器・乗率・新設・対応内容
 ,{64}      --高圧計器チェック・計器・乗率・新設・対応完了日
 ,{65}      --高圧計器チェック・計器・乗率・既設・施工者・良否区分
 ,{66}      --高圧計器チェック・計器・乗率・既設・抜取検査担当者・良否区分
 ,{67}      --高圧計器チェック・計器・乗率・既設・検査担当者・良否区分
 ,{68}      --高圧計器チェック・計器・乗率・既設・対応内容
 ,{69}      --高圧計器チェック・計器・乗率・既設・対応完了日
 ,{70}      --高圧計器チェック・計器・合番号・新設・施工者・良否区分
 ,{71}      --高圧計器チェック・計器・合番号・新設・抜取検査担当者・良否区分
 ,{72}      --高圧計器チェック・計器・合番号・新設・検査担当者・良否区分
 ,{73}      --高圧計器チェック・計器・合番号・新設・対応内容
 ,{74}      --高圧計器チェック・計器・合番号・新設・対応完了日
 ,{75}      --高圧計器チェック・計器・合番号・既設・施工者・良否区分
 ,{76}      --高圧計器チェック・計器・合番号・既設・抜取検査担当者・良否区分
 ,{77}      --高圧計器チェック・計器・合番号・既設・検査担当者・良否区分
 ,{78}      --高圧計器チェック・計器・合番号・既設・対応内容
 ,{79}      --高圧計器チェック・計器・合番号・既設・対応完了日
 ,{80}      --高圧計器チェック・計器・製造番号・施工者・良否区分
 ,{81}      --高圧計器チェック・計器・製造番号・抜取検査担当者・良否区分
 ,{82}      --高圧計器チェック・計器・製造番号・検査担当者・良否区分
 ,{83}      --高圧計器チェック・計器・製造番号・対応内容
 ,{84}      --高圧計器チェック・計器・製造番号・対応完了日
 ,{85}      --高圧計器チェック・計器・動作状況・動作・施工者・良否区分
 ,{86}      --高圧計器チェック・計器・動作状況・動作・抜取検査担当者・良否区分
 ,{87}      --高圧計器チェック・計器・動作状況・動作・検査担当者・良否区分
 ,{88}      --高圧計器チェック・計器・動作状況・動作・対応内容
 ,{89}      --高圧計器チェック・計器・動作状況・動作・対応完了日
 ,{90}      --高圧計器チェック・計器・動作状況・設定時刻・施工者・良否区分
 ,{91}      --高圧計器チェック・計器・動作状況・設定時刻・抜取検査担当者・良否区分
 ,{92}      --高圧計器チェック・計器・動作状況・設定時刻・検査担当者・良否区分
 ,{93}      --高圧計器チェック・計器・動作状況・設定時刻・対応内容
 ,{94}      --高圧計器チェック・計器・動作状況・設定時刻・対応完了日
 ,{95}      --高圧計器チェック・計器・動作状況・施工者・点検時刻
 ,{96}      --高圧計器チェック・計器・動作状況・検査担当者・点検時刻
 ,{97}      --高圧計器チェック・計器・動作状況・施工者・計器表示時刻
 ,{98}      --高圧計器チェック・計器・動作状況・検査担当者・計器表示時刻
 ,{99}      --高圧計器チェック・計器・結線・色別・施工者・良否区分
 ,{100}     --高圧計器チェック・計器・結線・色別・抜取検査担当者・良否区分
 ,{101}     --高圧計器チェック・計器・結線・色別・検査担当者・良否区分
 ,{102}     --高圧計器チェック・計器・結線・色別・対応内容
 ,{103}     --高圧計器チェック・計器・結線・色別・対応完了日
 ,{104}     --高圧計器チェック・計器・結線・半田揚げ・施工者・良否区分
 ,{105}     --高圧計器チェック・計器・結線・半田揚げ・抜取検査担当者・良否区分
 ,{106}     --高圧計器チェック・計器・結線・半田揚げ・検査担当者・良否区分
 ,{107}     --計器・結線・半田揚げ・対応内容
 ,{108}     --高圧計器チェック・計器・結線・半田揚げ・対応完了日
 ,{109}     --高圧計器チェック・計器・結線・露出・施工者・良否区分
 ,{110}     --高圧計器チェック・計器・結線・露出・抜取検査担当者・良否区分
 ,{111}     --高圧計器チェック・計器・結線・露出・検査担当者・良否区分
 ,{112}     --高圧計器チェック・計器・結線・露出・対応内容
 ,{113}     --高圧計器チェック・計器・結線・露出・対応完了日
 ,{114}     --高圧計器チェック・計器・結線・ビス締付・施工者・良否区分
 ,{115}     --高圧計器チェック・計器・結線・ビス締付・抜取検査担当者・良否区分
 ,{116}     --高圧計器チェック・計器・結線・ビス締付・検査担当者・良否区分
 ,{117}     --高圧計器チェック・計器・結線・ビス締付・対応内容
 ,{118}     --高圧計器チェック・計器・結線・ビス締付・対応完了日
 ,{119}     --高圧計器チェック・計器・検電・端子・施工者・良否区分
 ,{120}     --高圧計器チェック・計器・検電・端子・抜取検査担当者・良否区分
 ,{121}     --高圧計器チェック・計器・検電・端子・検査担当者・良否区分
 ,{122}     --高圧計器チェック・計器・検電・端子・対応内容
 ,{123}     --高圧計器チェック・計器・検電・端子・対応完了日
 ,{124}     --高圧計器チェック・表示端末・製造番号・撤去・施工者・良否区分
 ,{125}     --高圧計器チェック・表示端末・製造番号・撤去・抜取検査担当者・良否区分
 ,{126}     --高圧計器チェック・表示端末・製造番号・撤去・検査担当者・良否区分
 ,{127}     --高圧計器チェック・表示端末・製造番号・撤去・対応内容
 ,{128}     --高圧計器チェック・表示端末・製造番号・撤去・対応完了日
 ,{129}     --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・施工者・良否区分
 ,{130}     --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・抜取検査担当者・良否区分
 ,{131}     --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・検査担当者・良否区分
 ,{132}     --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・対応内容
 ,{133}     --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・対応完了日
 ,{134}     --高圧計器チェック・チェックターミナル・結線・色別・施工者・良否区分
 ,{135}     --高圧計器チェック・チェックターミナル・結線・色別・抜取検査担当者・良否区分
 ,{136}     --高圧計器チェック・チェックターミナル・結線・色別・検査担当者・良否区分
 ,{137}     --高圧計器チェック・チェックターミナル・結線・色別・対応内容
 ,{138}     --高圧計器チェック・チェックターミナル・結線・色別・対応完了日
 ,{139}     --高圧計器チェック・チェックターミナル・結線・半田・施工者・良否区分
 ,{140}     --高圧計器チェック・チェックターミナル・結線・半田・抜取検査担当者・良否区分
 ,{141}     --高圧計器チェック・チェックターミナル・結線・半田・検査担当者・良否区分
 ,{142}     --高圧計器チェック・チェックターミナル・結線・半田・対応内容
 ,{143}     --高圧計器チェック・チェックターミナル・結線・半田・対応完了日
 ,{144}     --高圧計器チェック・チェックターミナル・結線・ビス締付・施工者・良否区分
 ,{145}     --高圧計器チェック・チェックターミナル・結線・ビス締付・抜取検査担当者・良否区分
 ,{146}     --高圧計器チェック・チェックターミナル・結線・ビス締付・検査担当者・良否区分
 ,{147}     --高圧計器チェック・チェックターミナル・結線・ビス締付・対応内容
 ,{148}     --高圧計器チェック・チェックターミナル・結線・ビス締付・対応完了日
 ,{149}     --高圧計器チェック・チェックターミナル・結線・接栓締付・施工者・良否区分
 ,{150}     --高圧計器チェック・チェックターミナル・結線・接栓締付・抜取検査担当者・良否区分
 ,{151}     --高圧計器チェック・チェックターミナル・結線・接栓締付・検査担当者・良否区分
 ,{152}     --高圧計器チェック・チェックターミナル・結線・接栓締付・対応内容
 ,{153}     --高圧計器チェック・チェックターミナル・結線・接栓締付・対応完了日
 ,{154}     --高圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,{155}     --高圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,{156}     --高圧計器チェック・撤去計器等・製造番号・検査担当者・良否区分
 ,{157}     --高圧計器チェック・撤去計器等・製造番号・対応内容
 ,{158}     --高圧計器チェック・撤去計器等・製造番号・対応完了日
 ,{159}     --高圧計器チェック・その他・結線・施工者・良否区分
 ,{160}     --高圧計器チェック・その他・結線・抜取検査担当者・良否区分
 ,{161}     --高圧計器チェック・その他・結線・検査担当者・良否区分
 ,{162}     --高圧計器チェック・その他・結線・対応内容
 ,{163}     --高圧計器チェック・その他・結線・対応完了日
 ,{164}     --高圧計器チェック・その他・検相・施工者・良否区分
 ,{165}     --高圧計器チェック・その他・検相・抜取検査担当者・良否区分
 ,{166}     --高圧計器チェック・その他・検相・検査担当者・良否区分
 ,{167}     --高圧計器チェック・その他・検相・対応内容
 ,{168}     --高圧計器チェック・その他・検相・対応完了日
 ,{169}     --高圧計器チェック・その他・竣工写真・施工者・良否区分
 ,{170}     --高圧計器チェック・その他・竣工写真・抜取検査担当者・良否区分
 ,{171}     --高圧計器チェック・その他・竣工写真・検査担当者・良否区分
 ,{172}     --高圧計器チェック・その他・竣工写真・対応内容
 ,{173}     --高圧計器チェック・その他・竣工写真・対応完了日
 ,{174}     --高圧計器チェック・その他・最終確認・施工者・良否区分
 ,{175}     --高圧計器チェック・その他・最終確認・抜取検査担当者・良否区分
 ,{176}     --高圧計器チェック・その他・最終確認・検査担当者・良否区分
 ,{177}     --高圧計器チェック・その他・最終確認・対応内容
 ,{178}     --高圧計器チェック・その他・最終確認・対応完了日
 ,{179}     --高圧計器チェック・乗率チェック・施工者・良否区分
 ,{180}     --高圧計器チェック・乗率チェック・抜取検査担当者・良否区分
 ,{181}     --高圧計器チェック・乗率チェック・検査担当者・良否区分
 ,{182}     --高圧計器チェック・乗率チェック・対応内容
 ,{183}     --高圧計器チェック・乗率チェック・対応完了日
 ,{184}     --高圧計器チェック・合番号チェック・施工者・良否区分
 ,{185}     --高圧計器チェック・合番号チェック・抜取検査担当者・良否区分
 ,{186}     --高圧計器チェック・合番号チェック・検査担当者・良否区分
 ,{187}     --高圧計器チェック・合番号チェック・対応内容
 ,{188}     --高圧計器チェック・合番号チェック・対応完了日
 ,{189}     --自主点検結果_備考内容
 ,{190}     --高圧変成器・施工者・黒（1S）・結線等チェック結果フラグ
 ,{191}     --高圧変成器・施工者・赤（P1）・結線等チェック結果フラグ
 ,{192}     --高圧変成器・施工者・青（P3）・結線等チェック結果フラグ
 ,{193}     --高圧変成器・施工者・茶（3S）・結線等チェック結果フラグ
 ,{194}     --高圧変成器・施工者・黄（3L）・結線等チェック結果フラグ
 ,{195}     --高圧変成器・施工者・白（P2）・結線等チェック結果フラグ
 ,{196}     --高圧変成器・施工者・緑（1L）・結線等チェック結果フラグ
 ,{197}     --チェックターミナル・VCT・施工者・黒（1S）・結線等チェック結果フラグ
 ,{198}     --チェックターミナル・VCT・施工者・赤（P1）・結線等チェック結果フラグ
 ,{199}     --チェックターミナル・VCT・施工者・青（P3）・結線等チェック結果フラグ
 ,{200}     --チェックターミナル・VCT・施工者・茶（3S）・結線等チェック結果フラグ
 ,{201}     --チェックターミナル・VCT・施工者・黄（3L）・結線等チェック結果フラグ
 ,{202}     --チェックターミナル・VCT・施工者・白（P2）・結線等チェック結果フラグ
 ,{203}     --チェックターミナル・VCT・施工者・緑（1L）・結線等チェック結果フラグ
 ,{204}     --チェックターミナル・WH・施工者・黒（1S）・結線等チェック結果フラグ
 ,{205}     --チェックターミナル・WH・施工者・赤（P1）・結線等チェック結果フラグ
 ,{206}     --チェックターミナル・WH・施工者・青（P3）・結線等チェック結果フラグ
 ,{207}     --チェックターミナル・WH・施工者・茶（3S）・結線等チェック結果フラグ
 ,{208}     --チェックターミナル・WH・黄（3L）・施工者・結線等チェック結果フラグ
 ,{209}     --チェックターミナル・WH・施工者・白（P2）・結線等チェック結果フラグ
 ,{210}     --チェックターミナル・WH・施工者・緑（1L）・結線等チェック結果フラグ
 ,{211}     --電力需給用複合計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,{212}     --電力需給用複合計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,{213}     --電力需給用複合計器・施工者・青（P3）・結線等チェック結果フラグ
 ,{214}     --電力需給用複合計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,{215}     --電力需給用複合計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,{216}     --電力需給用複合計器・施工者・白（P2）・結線等チェック結果フラグ
 ,{217}     --電力需給用複合計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,{218}     --電力需給用複合計器・施工者・茶（DT）・結線等チェック結果フラグ
 ,{219}     --電力需給用複合計器・施工者・黄（SG）・結線等チェック結果フラグ
 ,{220}     --使用電力量表示端末・施工者・青（P3）・結線等チェック結果フラグ
 ,{221}     --使用電力量表示端末・施工者・茶（3S）・結線等チェック結果フラグ
 ,{222}     --使用電力量表示端末・施工者・黄（3L）・結線等チェック結果フラグ
 ,{223}     --使用電力量表示端末・施工者・白（P2）・結線等チェック結果フラグ
 ,{224}     --受給用計量器・施工者・黒（1S）・結線等チェック結果フラグ
 ,{225}     --受給用計量器・施工者・赤（P1）・結線等チェック結果フラグ
 ,{226}     --受給用計量器・施工者・青（P3）・結線等チェック結果フラグ
 ,{227}     --受給用計量器・施工者・茶（3S）・結線等チェック結果フラグ
 ,{228}     --受給用計量器・施工者・黄（3L）・結線等チェック結果フラグ
 ,{229}     --受給用計量器・施工者・白（P2）・結線等チェック結果フラグ
 ,{230}     --受給用計量器・施工者・緑（1L）・結線等チェック結果フラグ
 ,{231}     --高圧変成器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{232}     --高圧変成器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{233}     --高圧変成器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{234}     --高圧変成器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{235}     --高圧変成器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{236}     --高圧変成器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{237}     --高圧変成器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{238}     --チェックターミナル・VCT・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{239}     --チェックターミナル・VCT・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{240}     --チェックターミナル・VCT・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{241}     --チェックターミナル・VCT・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{242}     --チェックターミナル・VCT・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{243}     --チェックターミナル・VCT・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{244}     --チェックターミナル・VCT・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{245}     --チェックターミナル・WH・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{246}     --チェックターミナル・WH・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{247}     --チェックターミナル・WH・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{248}     --チェックターミナル・WH・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{249}     --チェックターミナル・WH・黄（3L）・抜取検査担当者・結線等チェック結果フラグ
 ,{250}     --チェックターミナル・WH・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{251}     --チェックターミナル・WH・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{252}     --電力需給用複合計器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{253}     --電力需給用複合計器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{254}     --電力需給用複合計器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{255}     --電力需給用複合計器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{256}     --電力需給用複合計器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{257}     --電力需給用複合計器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{258}     --電力需給用複合計器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{259}     --電力需給用複合計器・抜取検査担当者・茶（DT）・結線等チェック結果フラグ
 ,{260}     --電力需給用複合計器・抜取検査担当者・黄（SG）・結線等チェック結果フラグ
 ,{261}     --使用電力量表示端末・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{262}     --使用電力量表示端末・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{263}     --使用電力量表示端末・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{264}     --使用電力量表示端末・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{265}     --受給用計量器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{266}     --受給用計量器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{267}     --受給用計量器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{268}     --受給用計量器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{269}     --受給用計量器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{270}     --受給用計量器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{271}     --受給用計量器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{272}     --高圧変成器・検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{273}     --高圧変成器・検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{274}     --高圧変成器・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{275}     --高圧変成器・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{276}     --高圧変成器・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{277}     --高圧変成器・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{278}     --高圧変成器・検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{279}     --チェックターミナル・VCT・検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{280}     --チェックターミナル・VCT・検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{281}     --チェックターミナル・VCT・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{282}     --チェックターミナル・VCT・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{283}     --チェックターミナル・VCT・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{284}     --チェックターミナル・VCT・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{285}     --チェックターミナル・VCT・検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{286}     --チェックターミナル・WH・検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{287}     --チェックターミナル・WH・検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{288}     --チェックターミナル・WH・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{289}     --チェックターミナル・WH・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{290}     --チェックターミナル・WH・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{291}     --チェックターミナル・WH・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{292}     --チェックターミナル・WH・検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{293}     --電力需給用複合計器・検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{294}     --電力需給用複合計器・検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{295}     --電力需給用複合計器・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{296}     --電力需給用複合計器・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{297}     --電力需給用複合計器・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{298}     --電力需給用複合計器・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{299}     --電力需給用複合計器・検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{300}     --電力需給用複合計器・検査担当者・茶（DT）・結線等チェック結果フラグ
 ,{301}     --電力需給用複合計器・検査担当者・黄（SG）・結線等チェック結果フラグ
 ,{302}     --使用電力量表示端末・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{303}     --使用電力量表示端末・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{304}     --使用電力量表示端末・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{305}     --使用電力量表示端末・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{306}     --受給用計量器・検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{307}     --受給用計量器・検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{308}     --受給用計量器・検査担当者・青（P3）・結線等チェック結果フラグ
 ,{309}     --受給用計量器・検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{310}     --受給用計量器・検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{311}     --受給用計量器・検査担当者・白（P2）・結線等チェック結果フラグ
 ,{312}     --受給用計量器・検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{313}      --低圧計器チェック・電力量計・取付取替等・施工者・良否区分
 ,{314}      --低圧計器チェック・電力量計・取付取替等・抜取検査担当者・良否区分
 ,{315}      --低圧計器チェック・電力量計・取付取替等・検査担当者・良否区分
 ,{316}      --低圧計器チェック・電力量計・取付取替等・対応内容
 ,{317}      --低圧計器チェック・電力量計・取付取替等・対応完了日
 ,{318}      --低圧計器チェック・電力量計・指示数・施工者・良否区分
 ,{319}      --低圧計器チェック・電力量計・指示数・抜取検査担当者・良否区分
 ,{320}      --低圧計器チェック・電力量計・指示数・検査担当者・良否区分
 ,{321}      --低圧計器チェック・電力量計・指示数・対応内容
 ,{322}      --低圧計器チェック・電力量計・指示数・対応完了日
 ,{323}      --低圧計器チェック・計器項目・相線電圧・施工者・良否区分
 ,{324}      --低圧計器チェック・計器項目・相線電圧・抜取検査担当者・良否区分
 ,{325}      --低圧計器チェック・計器項目・相線電圧・検査担当者・良否区分
 ,{326}      --低圧計器チェック・計器項目・相線電圧・対応内容
 ,{327}      --低圧計器チェック・計器項目・相線電圧・対応完了日
 ,{328}      --低圧計器チェック・計器項目・乗率・施工者・良否区分
 ,{329}      --低圧計器チェック・計器項目・乗率・抜取検査担当者・良否区分
 ,{330}      --低圧計器チェック・計器項目・乗率・検査担当者・良否区分
 ,{331}      --低圧計器チェック・計器項目・乗率・対応内容
 ,{332}      --低圧計器チェック・計器項目・乗率・対応完了日
 ,{333}      --低圧計器チェック・計器項目・容量・施工者・良否区分
 ,{334}      --低圧計器チェック・計器項目・容量・抜取検査担当者・良否区分
 ,{335}      --低圧計器チェック・計器項目・容量・検査担当者・良否区分
 ,{336}      --低圧計器チェック・計器項目・容量・対応内容
 ,{337}      --低圧計器チェック・計器項目・容量・対応完了日
 ,{338}      --低圧計器チェック・計器項目・SM・施工者・良否区分
 ,{339}      --低圧計器チェック・計器項目・SM・抜取検査担当者・良否区分
 ,{340}      --低圧計器チェック・計器項目・SM・検査担当者・良否区分
 ,{341}      --低圧計器チェック・計器項目・SM・対応内容
 ,{342}      --低圧計器チェック・計器項目・SM・対応完了日
 ,{343}      --低圧計器チェック・計器項目・計器・合番号・施工者・良否区分
 ,{344}      --低圧計器チェック・計器項目・計器・合番号・抜取検査担当者・良否区分
 ,{345}      --低圧計器チェック・計器項目・計器・合番号・検査担当者・良否区分
 ,{346}      --低圧計器チェック・計器項目・計器・合番号・対応内容
 ,{347}      --低圧計器チェック・計器項目・計器・合番号・対応完了日
 ,{348}      --低圧計器チェック・計器項目・変流器・製造番号・施工者・良否区分
 ,{349}      --低圧計器チェック・計器項目・変流器・製造番号・抜取検査担当者・良否区分
 ,{350}      --低圧計器チェック・計器項目・変流器・製造番号・検査担当者・良否区分
 ,{351}      --低圧計器チェック・計器項目・変流器・製造番号・対応内容
 ,{352}      --低圧計器チェック・計器項目・変流器・製造番号・対応完了日
 ,{353}      --低圧計器チェック・計器項目・変流器・合番号・施工者・良否区分
 ,{354}      --低圧計器チェック・計器項目・変流器・合番号・抜取検査担当者・良否区分
 ,{355}      --低圧計器チェック・計器項目・変流器・合番号・検査担当者・良否区分
 ,{356}      --低圧計器チェック・計器項目・変流器・合番号・対応内容
 ,{357}      --低圧計器チェック・計器項目・変流器・合番号・対応完了日
 ,{358}      --低圧計器チェック・構造検査・計量器・検電・施工者・良否区分
 ,{359}      --低圧計器チェック・構造検査・計量器・検電・抜取検査担当者・良否区分
 ,{360}      --低圧計器チェック・構造検査・計量器・検電・検査担当者・良否区分
 ,{361}      --低圧計器チェック・構造検査・計量器・検電・対応内容
 ,{362}      --低圧計器チェック・構造検査・計量器・検電・対応完了日
 ,{363}      --低圧計器チェック・構造検査・計量器箱・施工者・良否区分
 ,{364}      --低圧計器チェック・構造検査・計量器箱・抜取検査担当者・良否区分
 ,{365}      --低圧計器チェック・構造検査・計量器箱・検査担当者・良否区分
 ,{366}      --低圧計器チェック・構造検査・計量器箱・対応内容
 ,{367}      --低圧計器チェック・構造検査・計量器箱・対応完了日
 ,{368}      --低圧計器チェック・結線等・計器・絶縁・施工者・良否区分
 ,{369}      --低圧計器チェック・結線等・計器・絶縁・抜取検査担当者・良否区分
 ,{370}      --低圧計器チェック・結線等・計器・絶縁・検査担当者・良否区分
 ,{371}      --低圧計器チェック・結線等・計器・絶縁・対応内容
 ,{372}      --低圧計器チェック・結線等・計器・絶縁・対応完了日
 ,{373}      --低圧計器チェック・結線等・計器・SL・施工者・良否区分
 ,{374}      --低圧計器チェック・結線等・計器・SL・抜取検査担当者・良否区分
 ,{375}      --低圧計器チェック・結線等・計器・SL・検査担当者・良否区分
 ,{376}      --低圧計器チェック・結線等・計器・SL・対応内容
 ,{377}      --低圧計器チェック・結線等・計器・SL・対応完了日
 ,{378}      --低圧計器チェック・結線等・計器・複数・施工者・良否区分
 ,{379}      --低圧計器チェック・結線等・計器・複数・抜取検査担当者・良否区分
 ,{380}      --低圧計器チェック・結線等・計器・複数・検査担当者・良否区分
 ,{381}      --低圧計器チェック・結線等・計器・複数・対応内容
 ,{382}      --低圧計器チェック・結線等・計器・複数・対応完了日
 ,{383}      --低圧計器チェック・結線等・計器・誤接続・施工者・良否区分
 ,{384}      --低圧計器チェック・結線等・計器・誤接続・抜取検査担当者・良否区分
 ,{385}      --低圧計器チェック・結線等・計器・誤接続・検査担当者・良否区分
 ,{386}      --低圧計器チェック・結線等・計器・誤接続・対応内容
 ,{387}      --低圧計器チェック・結線等・計器・誤接続・対応完了日
 ,{388}      --低圧計器チェック・結線等・計器・赤色・施工者・良否区分
 ,{389}      --低圧計器チェック・結線等・計器・赤色・抜取検査担当者・良否区分
 ,{390}      --低圧計器チェック・結線等・計器・赤色・検査担当者・良否区分
 ,{391}      --低圧計器チェック・結線等・計器・赤色・対応内容
 ,{392}      --低圧計器チェック・結線等・計器・赤色・対応完了日
 ,{393}      --低圧計器チェック・結線等・計器・より線・施工者・良否区分
 ,{394}      --低圧計器チェック・結線等・計器・より線・抜取検査担当者・良否区分
 ,{395}      --低圧計器チェック・結線等・計器・より線・検査担当者・良否区分
 ,{396}      --低圧計器チェック・結線等・計器・より線・対応内容
 ,{397}      --低圧計器チェック・結線等・計器・より線・対応完了日
 ,{398}      --低圧計器チェック・結線等・計器・露出・施工者・良否区分
 ,{399}      --低圧計器チェック・結線等・計器・露出・抜取検査担当者・良否区分
 ,{400}      --低圧計器チェック・結線等・計器・露出・検査担当者・良否区分
 ,{401}      --低圧計器チェック・結線等・計器・露出・対応内容
 ,{402}      --低圧計器チェック・結線等・計器・露出・対応完了日
 ,{403}      --低圧計器チェック・結線等・計器・ビス締付・施工者・良否区分
 ,{404}      --低圧計器チェック・結線等・計器・ビス締付・抜取検査担当者・良否区分
 ,{405}      --低圧計器チェック・結線等・計器・ビス締付・検査担当者・良否区分
 ,{406}      --低圧計器チェック・結線等・計器・ビス締付・対応内容
 ,{407}      --低圧計器チェック・結線等・計器・ビス締付・対応完了日
 ,{408}      --低圧計器チェック・結線等・計器・色別・施工者・良否区分
 ,{409}      --低圧計器チェック・結線等・計器・色別・抜取検査担当者・良否区分
 ,{410}      --低圧計器チェック・結線等・計器・色別・検査担当者・良否区分
 ,{411}      --低圧計器チェック・結線等・計器・色別・対応内容
 ,{412}      --低圧計器チェック・結線等・計器・色別・対応完了日
 ,{413}      --低圧計器チェック・結線等・計器・正相・施工者・良否区分
 ,{414}      --低圧計器チェック・結線等・計器・正相・抜取検査担当者・良否区分
 ,{415}      --低圧計器チェック・結線等・計器・正相・検査担当者・良否区分
 ,{416}      --低圧計器チェック・結線等・計器・正相・対応内容
 ,{417}      --低圧計器チェック・結線等・計器・正相・対応完了日
 ,{418}      --低圧計器チェック・結線等・変流器・ＫＬ・施工者・良否区分
 ,{419}      --低圧計器チェック・結線等・変流器・ＫＬ・抜取検査担当者・良否区分
 ,{420}      --低圧計器チェック・結線等・変流器・ＫＬ・検査担当者・良否区分
 ,{421}      --低圧計器チェック・結線等・変流器・ＫＬ・対応内容
 ,{422}      --低圧計器チェック・結線等・変流器・ＫＬ・対応完了日
 ,{423}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・施工者・良否区分
 ,{424}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・抜取検査担当者・良否区分
 ,{425}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・検査担当者・良否区分
 ,{426}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応内容
 ,{427}      --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・対応完了日
 ,{428}      --低圧計器チェック・結線等・変流器・赤色・施工者・良否区分
 ,{429}      --低圧計器チェック・結線等・変流器・赤色・抜取検査担当者・良否区分
 ,{430}      --低圧計器チェック・結線等・変流器・赤色・検査担当者・良否区分
 ,{431}      --低圧計器チェック・結線等・変流器・赤色・対応内容
 ,{432}      --低圧計器チェック・結線等・変流器・赤色・対応完了日
 ,{433}      --低圧計器チェック・計器回転・検電・変流器・施工者・良否区分
 ,{434}      --低圧計器チェック・計器回転・検電・変流器・抜取検査担当者・良否区分
 ,{435}      --低圧計器チェック・計器回転・検電・変流器・検査担当者・良否区分
 ,{436}      --低圧計器チェック・計器回転・検電・変流器・対応内容
 ,{437}      --低圧計器チェック・計器回転・検電・変流器・対応完了日
 ,{438}      --低圧計器チェック・計器回転・回転・SM・施工者・良否区分
 ,{439}      --低圧計器チェック・計器回転・回転・SM・抜取検査担当者・良否区分
 ,{440}      --低圧計器チェック・計器回転・回転・SM・検査担当者・良否区分
 ,{441}      --低圧計器チェック・計器回転・回転・SM・対応内容
 ,{442}      --低圧計器チェック・計器回転・回転・SM・対応完了日
 ,{443}      --低圧計器チェック・通信部・ＬＥＤ・施工者・良否区分
 ,{444}      --低圧計器チェック・通信部・ＬＥＤ・抜取検査担当者・良否区分
 ,{445}      --低圧計器チェック・通信部・ＬＥＤ・検査担当者・良否区分
 ,{446}      --低圧計器チェック・通信部・ＬＥＤ・対応内容
 ,{447}      --低圧計器チェック・通信部・ＬＥＤ・対応完了日
 ,{448}      --低圧計器チェック・通信部・装着・施工者・良否区分
 ,{449}      --低圧計器チェック・通信部・装着・抜取検査担当者・良否区分
 ,{450}      --低圧計器チェック・通信部・装着・検査担当者・良否区分
 ,{451}      --低圧計器チェック・通信部・装着・対応内容
 ,{452}      --低圧計器チェック・通信部・装着・対応完了日
 ,{453}      --低圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,{454}      --低圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,{455}      --低圧計器チェック・撤去計器等・製造番号・検査担当者・良否区分
 ,{456}      --低圧計器チェック・撤去計器等・製造番号・対応内容
 ,{457}      --低圧計器チェック・撤去計器等・製造番号・対応完了日
 ,{458}      --低圧計器チェック・撤去計器等・合番号・施工者・良否区分
 ,{459}      --低圧計器チェック・撤去計器等・合番号・抜取検査担当者・良否区分
 ,{460}      --低圧計器チェック・撤去計器等・合番号・検査担当者・良否区分
 ,{461}      --低圧計器チェック・撤去計器等・合番号・対応内容
 ,{462}      --低圧計器チェック・撤去計器等・合番号・対応完了日
 ,{463}      --低圧計器チェック・その他・写真・施工者・良否区分
 ,{464}      --低圧計器チェック・その他・写真・抜取検査担当者・良否区分
 ,{465}      --低圧計器チェック・その他・写真・検査担当者・良否区分
 ,{466}      --低圧計器チェック・その他・写真・対応内容
 ,{467}      --低圧計器チェック・その他・写真・対応完了日
 ,{468}      --低圧計器チェック・その他・撤去取替・施工者・良否区分
 ,{469}      --低圧計器チェック・その他・撤去取替・抜取検査担当者・良否区分
 ,{470}      --低圧計器チェック・その他・撤去取替・検査担当者・良否区分
 ,{471}      --低圧計器チェック・その他・撤去取替・対応内容
 ,{472}      --低圧計器チェック・その他・撤去取替・対応完了日
 ,{473}      --低圧計器チェック・乗率チェック・抜取検査担当者・整合確認区分
 ,{474}      --低圧計器チェック・乗率チェック・検査担当者・整合確認区分
 ,{475}      --低圧計器チェック・乗率チェック・対応内容
 ,{476}      --低圧計器チェック・乗率チェック・対応完了日
 ,{477}      --低圧計器チェック・合番号チェック・抜取検査担当者・一致確認区分
 ,{478}      --低圧計器チェック・合番号チェック・検査担当者・一致確認区分
 ,{479}      --低圧計器チェック・合番号チェック・対応内容
 ,{480}      --低圧計器チェック・合番号チェック・対応完了日
 ,{481}      --記事（施工者）内容
 ,{482}      --記事（抜取検査担当者）内容
 ,{483}      --記事（検査担当者）内容
 ,{484}      --変流器付計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,{485}      --変流器付計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,{486}      --変流器付計器・施工者・青（P3）・結線等チェック結果フラグ
 ,{487}      --変流器付計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,{488}      --変流器付計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,{489}      --変流器付計器・施工者・白（P2）・結線等チェック結果フラグ
 ,{490}      --変流器付計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,{491}      --P1・施工者・黒（1S）・結線等チェック結果フラグ
 ,{492}      --P1・施工者・赤（P1）・結線等チェック結果フラグ
 ,{493}      --P1・施工者・緑（1L）・結線等チェック結果フラグ
 ,{494}      --P2・施工者・白（P2）・結線等チェック結果フラグ
 ,{495}      --P3・施工者・茶（3S）・結線等チェック結果フラグ
 ,{496}      --P3・施工者・青（P3）・結線等チェック結果フラグ
 ,{497}      --P3・施工者・黄（3L）・結線等チェック結果フラグ
 ,{498}      --変流器付計器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{499}      --変流器付計器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{500}      --変流器付計器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{501}      --変流器付計器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{502}      --変流器付計器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{503}      --変流器付計器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{504}      --変流器付計器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{505}      --P1・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,{506}      --P1・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,{507}      --P1・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,{508}      --P2・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,{509}      --P3・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,{510}      --P3・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,{511}      --P3・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,{512}      --変流器付計器・検査担当者・黒(1S)・結線等チェック結果フラグ
 ,{513}      --変流器付計器・検査担当者・赤(P1)・結線等チェック結果フラグ
 ,{514}      --変流器付計器・検査担当者・青(P3)・結線等チェック結果フラグ
 ,{515}      --変流器付計器・検査担当者・茶(3S)・結線等チェック結果フラグ
 ,{516}      --変流器付計器・検査担当者・黄(3L)・結線等チェック結果フラグ
 ,{517}      --変流器付計器・検査担当者・白(P2)・結線等チェック結果フラグ
 ,{518}      --変流器付計器・検査担当者・緑(1L)・結線等チェック結果フラグ
 ,{519}      --P1・検査担当者・黒(1S)・結線等チェック結果フラグ
 ,{520}      --P1・検査担当者・赤(P1)・結線等チェック結果フラグ
 ,{521}      --P1・検査担当者・緑(1L)・結線等チェック結果フラグ
 ,{522}      --P2・検査担当者・白(P2)・結線等チェック結果フラグ
 ,{523}      --P3・検査担当者・茶(3S)・結線等チェック結果フラグ
 ,{524}      --P3・検査担当者・青(P3)・結線等チェック結果フラグ
 ,{525}      --P3・検査担当者・黄(3L)・結線等チェック結果フラグ
 ,{526}      --低圧計器チェック・計器回転・回転・SM表示・施工者・良否区分
 ,{527}      --低圧計器チェック・計器回転・回転・SM表示・抜取検査担当者・良否区分
 ,{528}      --低圧計器チェック・計器回転・回転・SM表示・検査担当者・良否区分
 ,{529}      --低圧計器チェック・計器回転・回転・SM表示・対応内容
 ,{530}      --低圧計器チェック・計器回転・回転・SM表示・対応完了日
 ,{531}      --低圧計器チェック・計器回転・回転・D2SM・施工者・良否区分
 ,{532}      --低圧計器チェック・計器回転・回転・D2SM・抜取検査担当者・良否区分
 ,{533}      --低圧計器チェック・計器回転・回転・D2SM・検査担当者・良否区分
 ,{534}      --低圧計器チェック・計器回転・回転・D2SM・対応内容
 ,{535}      --低圧計器チェック・計器回転・回転・D2SM・対応完了日
 ,{536}      --低圧計器チェック・計器回転・回転・開閉器・施工者・良否区分
 ,{537}      --低圧計器チェック・計器回転・回転・開閉器・抜取検査担当者・良否区分
 ,{538}      --低圧計器チェック・計器回転・回転・開閉器・検査担当者・良否区分
 ,{539}      --低圧計器チェック・計器回転・回転・開閉器・対応内容
 ,{540}      --低圧計器チェック・計器回転・回転・開閉器・対応完了日
 ,{541}      --低圧計器チェック・構造検査・計量器箱・屋外計器箱・施工者・良否区分
 ,{542}      --低圧計器チェック・構造検査・計量器箱・屋外計器箱・抜取検査担当者・良否区分
 ,{543}      --低圧計器チェック・構造検査・計量器箱・屋外計器箱・検査担当者・良否区分
 ,{544}      --低圧計器チェック・構造検査・計量器箱・屋外計器箱・対応内容
 ,{545}      --低圧計器チェック・構造検査・計量器箱・屋外計器箱・対応完了日
 ,{546}      --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・施工者・良否区分
 ,{547}      --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・抜取検査担当者・良否区分
 ,{548}      --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・検査担当者・良否区分
 ,{549}      --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・対応内容
 ,{550}      --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・対応完了日
)
         ]]></sql>.Value)

            Dim procDateTime As String = getDbDateTime()

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kstts_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_kst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_kst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_kst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_kst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_kst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_kst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_kst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_kst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_skosya_TNKN_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_kstts_tnkn_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_skosya_kikhyz_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_kstts_kikhyz_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KEIKI_ksn_hndag_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKBS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKBS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKBS_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcc_ksn_bsst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcc_ksn_bsst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ZSTK_KKA_bk_NAIYO"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_dt_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_sg_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_dt_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_ntktt_sg_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_P3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_3S_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_3L_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_P2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_kstts_1L_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_dt_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_kstts_sg_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_ntktt_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_kstts_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_ntktt_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_kstts_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_skosya_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_ntktt_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_kstts_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_kstts_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_kstts_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_kstts_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_tio_cmp_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_kstts_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_tio_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_tio_cmp_ymd"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, MjK1.Cmn.Func.DelSQLComment(sqlstr))
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


#Region "(SQL)取替_自主点検チェックシート（高圧他）テーブル(施工)更新"
    ''' <summary>
    ''' 取替_自主点検チェックシート（高圧他）テーブル(施工)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシート（高圧他）テーブル(施工)を更新します</remarks>
    Friend Function U017(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD takckrr_smhyz_skosya_ryohi_kbn ～ takckk_dsipcv_skosya_ryohi_kbn
            sb.Append(<sql><![CDATA[
UPDATE trke_check_sheet_kahk SET        -- 取替_自主点検チェックシート（高圧他）
  upd_date={0}                          --更新日時
 ,sousa_user_id={1}                     --更新ユーザ
 ,sousa_appli_cd={2}                    --更新プログラム
 ,skosya_tnkn_ymd={7}                   --点検（検査）年月日・施工者
 ,kstts_tnkn_ymd={8}                    --点検（検査）年月日・検査担当者
 ,KKCHYR_snst_skosya_ryohi_kbn={9}      --高圧計器チェック・変成器・容量・新設・施工者・良否区分
 ,KKCHANO_snst_skosya_ryohi_kbn={10}    --高圧計器チェック・変成器・合番号・新設・施工者・良否区分
 ,KKCHSNO_skosya_ryohi_kbn={11}         --高圧計器チェック・変成器・製造番号・施工者・良否区分
 ,KKCHKKL_skosya_ryohi_kbn={12}         --高圧計器チェック・変成器・結線・KL・施工者・良否区分
 ,KKCHKSS_skosya_ryohi_kbn={13}         --高圧計器チェック・変成器・結線・正相・施工者・良否区分
 ,KKCHKIB_skosya_ryohi_kbn={14}         --高圧計器チェック・変成器・結線・色別・施工者・良否区分
 ,KKCHKHD_skosya_ryohi_kbn={15}         --高圧計器チェック・変成器・結線・半田・施工者・良否区分
 ,KKCHKBS_skosya_ryohi_kbn={16}         --高圧計器チェック・変成器・結線・ビス締付・施工者・良否区分
 ,KKCKZR_snst_skosya_ryohi_kbn={17}     --高圧計器チェック・計器・乗率・新設・施工者・良否区分
 ,KKCKANO_snst_skosya_ryohi_kbn={18}    --高圧計器チェック・計器・合番号・新設・施工者・良否区分
 ,KKCKSNO_skosya_ryohi_kbn={19}         --高圧計器チェック・計器・製造番号・施工者・良否区分
 ,KKCKDJ_dosa_skosya_ryohi_kbn={20}     --高圧計器チェック・計器・動作状況・動作・施工者・良否区分
 ,KKCKDJ_settjk_skosya_ryohi_kbn={21}   --高圧計器チェック・計器・動作状況・設定時刻・施工者・良否区分
 ,KKCKDJ_skosya_TNKN_HMS={22}           --高圧計器チェック・計器・動作状況・施工者・点検時刻
 ,KKCKDJ_skosya_kikhyz_HMS={23}         --高圧計器チェック・計器・動作状況・施工者・計器表示時刻
 ,KKCKKS_ib_skosya_ryohi_kbn={24}       --高圧計器チェック・計器・結線・色別・施工者・良否区分
 ,KKCKKS_hndag_skosya_ryohi_kbn={25}    --高圧計器チェック・計器・結線・半田揚げ・施工者・良否区分
 ,KKCKKS_rosyt_skosya_ryohi_kbn={26}    --高圧計器チェック・計器・結線・露出・施工者・良否区分
 ,KKCKKS_bsst_skosya_ryohi_kbn={27}     --高圧計器チェック・計器・結線・ビス締付・施工者・良否区分
 ,KKCKKD_tansi_skosya_ryohi_kbn={28}    --高圧計器チェック・計器・検電・端子・施工者・良否区分
 ,KKCHTSN_tekyo_skosya_ryohi_kbn={29}   --高圧計器チェック・表示端末・製造番号・撤去・施工者・良否区分
 ,KAKCTKVCT_skosya_ryohi_kbn={30}       --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・施工者・良否区分
 ,KAKCTKIB_skosya_ryohi_kbn={31}        --高圧計器チェック・チェックターミナル・結線・色別・施工者・良否区分
 ,KAKCTKHD_skosya_ryohi_kbn={32}        --高圧計器チェック・チェックターミナル・結線・半田・施工者・良否区分
 ,KAKCTKBS_skosya_ryohi_kbn={33}        --高圧計器チェック・チェックターミナル・結線・ビス締付・施工者・良否区分
 ,KAKCTKSS_skosya_ryohi_kbn={34}        --高圧計器チェック・チェックターミナル・結線・接栓締付・施工者・良否区分
 ,kakct_szno_skosya_ryohi_kbn={35}      --高圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,kakcs_ksn_skosya_ryohi_kbn={36}       --高圧計器チェック・その他・結線・施工者・良否区分
 ,kakcs_kns_skosya_ryohi_kbn={37}       --高圧計器チェック・その他・検相・施工者・良否区分
 ,kakcs_skphoto_skosya_ryohi_kbn={38}   --高圧計器チェック・その他・竣工写真・施工者・良否区分
 ,kakcs_kknn_skosya_ryohi_kbn={39}      --高圧計器チェック・その他・最終確認・施工者・良否区分
 ,kakcz_skosya_ryohi_kbn={40}           --高圧計器チェック・乗率チェック・施工者・良否区分
 ,kakca_skosya_ryohi_kbn={41}           --高圧計器チェック・合番号チェック・施工者・良否区分
 ,ZSTK_KKA_bk_NAIYO={42}                --自主点検結果_備考内容
 ,KAHSK_skosya_1s_kstckk_flg={43}       --高圧変成器・施工者・黒（1S）・結線等チェック結果フラグ
 ,KAHSK_skosya_p1_kstckk_flg={44}       --高圧変成器・施工者・赤（P1）・結線等チェック結果フラグ
 ,KAHSK_skosya_p3_kstckk_flg={45}       --高圧変成器・施工者・青（P3）・結線等チェック結果フラグ
 ,KAHSK_skosya_3s_kstckk_flg={46}       --高圧変成器・施工者・茶（3S）・結線等チェック結果フラグ
 ,KAHSK_skosya_3l_kstckk_flg={47}       --高圧変成器・施工者・黄（3L）・結線等チェック結果フラグ
 ,KAHSK_skosya_p2_kstckk_flg={48}       --高圧変成器・施工者・白（P2）・結線等チェック結果フラグ
 ,KAHSK_skosya_1l_kstckk_flg={49}       --高圧変成器・施工者・緑（1L）・結線等チェック結果フラグ
 ,cktrmvct_skosya_1s_kstckk_flg={50}    --チェックターミナル・VCT・施工者・黒（1S）・結線等チェック結果フラグ
 ,cktrmvct_skosya_p1_kstckk_flg={51}    --チェックターミナル・VCT・施工者・赤（P1）・結線等チェック結果フラグ
 ,cktrmvct_skosya_p3_kstckk_flg={52}    --チェックターミナル・VCT・施工者・青（P3）・結線等チェック結果フラグ
 ,cktrmvct_skosya_3s_kstckk_flg={53}    --チェックターミナル・VCT・施工者・茶（3S）・結線等チェック結果フラグ
 ,cktrmvct_skosya_3l_kstckk_flg={54}    --チェックターミナル・VCT・施工者・黄（3L）・結線等チェック結果フラグ
 ,cktrmvct_skosya_p2_kstckk_flg={55}    --チェックターミナル・VCT・施工者・白（P2）・結線等チェック結果フラグ
 ,cktrmvct_skosya_1l_kstckk_flg={56}    --チェックターミナル・VCT・施工者・緑（1L）・結線等チェック結果フラグ
 ,cktrmwh_skosya_1s_kstckk_flg={57}     --チェックターミナル・WH・施工者・黒（1S）・結線等チェック結果フラグ
 ,cktrmwh_skosya_p1_kstckk_flg={58}     --チェックターミナル・WH・施工者・赤（P1）・結線等チェック結果フラグ
 ,cktrmwh_skosya_p3_kstckk_flg={59}     --チェックターミナル・WH・施工者・青（P3）・結線等チェック結果フラグ
 ,cktrmwh_skosya_3s_kstckk_flg={60}     --チェックターミナル・WH・施工者・茶（3S）・結線等チェック結果フラグ
 ,cktrmwh_skosya_3l_kstckk_flg={61}     --チェックターミナル・WH・黄（3L）・施工者・結線等チェック結果フラグ
 ,cktrmwh_skosya_p2_kstckk_flg={62}     --チェックターミナル・WH・施工者・白（P2）・結線等チェック結果フラグ
 ,cktrmwh_skosya_1l_kstckk_flg={63}     --チェックターミナル・WH・施工者・緑（1L）・結線等チェック結果フラグ
 ,DJFK_skosya_1s_kstckk_flg={64}        --電力需給用複合計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,DJFK_skosya_p1_kstckk_flg={65}        --電力需給用複合計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,DJFK_skosya_p3_kstckk_flg={66}        --電力需給用複合計器・施工者・青（P3）・結線等チェック結果フラグ
 ,DJFK_skosya_3s_kstckk_flg={67}        --電力需給用複合計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,DJFK_skosya_3l_kstckk_flg={68}        --電力需給用複合計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,DJFK_skosya_p2_kstckk_flg={69}        --電力需給用複合計器・施工者・白（P2）・結線等チェック結果フラグ
 ,DJFK_skosya_1l_kstckk_flg={70}        --電力需給用複合計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,DJFK_skosya_dt_kstckk_flg={71}        --電力需給用複合計器・施工者・茶（DT）・結線等チェック結果フラグ
 ,DJFK_skosya_sg_kstckk_flg={72}        --電力需給用複合計器・施工者・黄（SG）・結線等チェック結果フラグ
 ,sdyyht_skosya_p3_kstckk_flg={73}      --使用電力量表示端末・施工者・青（P3）・結線等チェック結果フラグ
 ,sdyyht_skosya_3s_kstckk_flg={74}      --使用電力量表示端末・施工者・茶（3S）・結線等チェック結果フラグ
 ,sdyyht_skosya_3l_kstckk_flg={75}      --使用電力量表示端末・施工者・黄（3L）・結線等チェック結果フラグ
 ,sdyyht_skosya_p2_kstckk_flg={76}      --使用電力量表示端末・施工者・白（P2）・結線等チェック結果フラグ
 ,jkykrk_skosya_1s_kstckk_flg={77}      --受給用計量器・施工者・黒（1S）・結線等チェック結果フラグ
 ,jkykrk_skosya_p1_kstckk_flg={78}      --受給用計量器・施工者・赤（P1）・結線等チェック結果フラグ
 ,jkykrk_skosya_p3_kstckk_flg={79}      --受給用計量器・施工者・青（P3）・結線等チェック結果フラグ
 ,jkykrk_skosya_3s_kstckk_flg={80}      --受給用計量器・施工者・茶（3S）・結線等チェック結果フラグ
 ,jkykrk_skosya_3l_kstckk_flg={81}      --受給用計量器・施工者・黄（3L）・結線等チェック結果フラグ
 ,jkykrk_skosya_p2_kstckk_flg={82}      --受給用計量器・施工者・白（P2）・結線等チェック結果フラグ
 ,jkykrk_skosya_1l_kstckk_flg={83}      --受給用計量器・施工者・緑（1L）・結線等チェック結果フラグ
 ,takcd_ttt_skosya_ryohi_kbn={84}       --低圧計器チェック・電力量計・取付取替等・施工者・良否区分
 ,takcd_szsu_skosya_ryohi_kbn={85}      --低圧計器チェック・電力量計・指示数・施工者・良否区分
 ,takck_ssda_skosya_ryohi_kbn={86}      --低圧計器チェック・計器項目・相線電圧・施工者・良否区分
 ,takck_zyrt_skosya_ryohi_kbn={87}      --低圧計器チェック・計器項目・乗率・施工者・良否区分
 ,takck_yr_skosya_ryohi_kbn={88}        --低圧計器チェック・計器項目・容量・施工者・良否区分
 ,takck_sm_skosya_ryohi_kbn={89}        --低圧計器チェック・計器項目・SM・施工者・良否区分
 ,takck_kkano_skosya_ryohi_kbn={90}     --低圧計器チェック・計器項目・計器・合番号・施工者・良否区分
 ,takck_hksno_skosya_ryohi_kbn={91}     --低圧計器チェック・計器項目・変流器・製造番号・施工者・良否区分
 ,takck_hkano_skosya_ryohi_kbn={92}     --低圧計器チェック・計器項目・変流器・合番号・施工者・良否区分
 ,takckk_krkkd_skosya_ryohi_kbn={93}    --低圧計器チェック・構造検査・計量器・検電・施工者・良否区分
 ,takckk_krkb_skosya_ryohi_kbn={94}     --低圧計器チェック・構造検査・計量器箱・施工者・良否区分
 ,takcks_kkzen_skosya_ryohi_kbn={95}    --低圧計器チェック・結線等・計器・絶縁・施工者・良否区分
 ,takcks_kksl_skosya_ryohi_kbn={96}     --低圧計器チェック・結線等・計器・SL・施工者・良否区分
 ,takcks_kkhks_skosya_ryohi_kbn={97}    --低圧計器チェック・結線等・計器・複数・施工者・良否区分
 ,takcks_kkgsz_skosya_ryohi_kbn={98}    --低圧計器チェック・結線等・計器・誤接続・施工者・良否区分
 ,takcks_kkaki_skosya_ryohi_kbn={99}    --低圧計器チェック・結線等・計器・赤色・施工者・良否区分
 ,takcks_kkyrs_skosya_ryohi_kbn={100}   --低圧計器チェック・結線等・計器・より線・施工者・良否区分
 ,takcks_kkrst_skosya_ryohi_kbn={101}   --低圧計器チェック・結線等・計器・露出・施工者・良否区分
 ,takcks_kkbst_skosya_ryohi_kbn={102}   --低圧計器チェック・結線等・計器・ビス締付・施工者・良否区分
 ,takcks_kikib_skosya_ryohi_kbn={103}   --低圧計器チェック・結線等・計器・色別・施工者・良否区分
 ,takcks_kkssu_skosya_ryohi_kbn={104}   --低圧計器チェック・結線等・計器・正相・施工者・良否区分
 ,takcks_hrkkl_skosya_ryohi_kbn={105}   --低圧計器チェック・結線等・変流器・ＫＬ・施工者・良否区分
 ,takcks_hp123_skosya_ryohi_kbn={106}   --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・施工者・良否区分
 ,takcks_hkaki_skosya_ryohi_kbn={107}   --低圧計器チェック・結線等・変流器・赤色・施工者・良否区分
 ,takckrk_hnrk_skosya_ryohi_kbn={108}   --低圧計器チェック・計器回転・検電・変流器・施工者・良否区分
 ,takckrr_sm_skosya_ryohi_kbn={109}     --低圧計器チェック・計器回転・回転・SM・施工者・良否区分
 ,takcc_led_skosya_ryohi_kbn={110}      --低圧計器チェック・通信部・ＬＥＤ・施工者・良否区分
 ,takcc_souck_skosya_ryohi_kbn={111}    --低圧計器チェック・通信部・装着・施工者・良否区分
 ,takct_szno_skosya_ryohi_kbn={112}     --低圧計器チェック・撤去計器等・製造番号・施工者・良否区分
 ,takct_aino_skosya_ryohi_kbn={113}     --低圧計器チェック・撤去計器等・合番号・施工者・良否区分
 ,takcs_photo_skosya_ryohi_kbn={114}    --低圧計器チェック・その他・写真・施工者・良否区分
 ,takcs_tktrk_skosya_ryohi_kbn={115}    --低圧計器チェック・その他・撤去取替・施工者・良否区分
 ,kizi_skosya_naiyo={116}               --記事（施工者）内容
 ,hrktkk_skosya_1s_kstckk_flg={117}      --変流器付計器・施工者・黒（1S）・結線等チェック結果フラグ
 ,hrktkk_skosya_p1_kstckk_flg={118}      --変流器付計器・施工者・赤（P1）・結線等チェック結果フラグ
 ,hrktkk_skosya_p3_kstckk_flg={119}      --変流器付計器・施工者・青（P3）・結線等チェック結果フラグ
 ,hrktkk_skosya_3s_kstckk_flg={120}      --変流器付計器・施工者・茶（3S）・結線等チェック結果フラグ
 ,hrktkk_skosya_3l_kstckk_flg={121}      --変流器付計器・施工者・黄（3L）・結線等チェック結果フラグ
 ,hrktkk_skosya_p2_kstckk_flg={122}      --変流器付計器・施工者・白（P2）・結線等チェック結果フラグ
 ,hrktkk_skosya_1l_kstckk_flg={123}      --変流器付計器・施工者・緑（1L）・結線等チェック結果フラグ
 ,p1_skosya_1s_kstckk_flg={124}      --P1・施工者・黒（1S）・結線等チェック結果フラグ
 ,p1_skosya_p1_kstckk_flg={125}      --P1・施工者・赤（P1）・結線等チェック結果フラグ
 ,p1_skosya_1l_kstckk_flg={126}      --P1・施工者・緑（1L）・結線等チェック結果フラグ
 ,p2_skosya_p2_kstckk_flg={127}      --P2・施工者・白（P2）・結線等チェック結果フラグ
 ,p3_skosya_3s_kstckk_flg={128}      --P3・施工者・茶（3S）・結線等チェック結果フラグ
 ,p3_skosya_p3_kstckk_flg={129}      --P3・施工者・青（P3）・結線等チェック結果フラグ
 ,p3_skosya_3l_kstckk_flg={130}      --P3・施工者・黄（3L）・結線等チェック結果フラグ
 ,takckrr_smhyz_skosya_ryohi_kbn={131}   --低圧計器チェック・計器回転・回転・SM表示・施工者・良否区分
 ,takckrr_d2sm_skosya_ryohi_kbn={132}    --低圧計器チェック・計器回転・回転・D2SM・施工者・良否区分
 ,takckrr_khk_skosya_ryohi_kbn={133}     --低圧計器チェック・計器回転・回転・開閉器・施工者・良否区分
 ,takckk_okgkb_skosya_ryohi_kbn={134}    --低圧計器チェック・構造検査・計量器箱・屋外計器箱・施工者・良否区分
 ,takckk_dsipcv_skosya_ryohi_kbn={135}   --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・施工者・良否区分
WHERE 1 = 1
 AND  dig4_zgsyo_cd={3}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={4}      --取替票発行年度
 AND  trkhy_kbn={5}      --取替票区分
 AND  trkhy_no={6}      --取替票番号         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                               EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kstts_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_skosya_TNKN_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_skosya_kikhyz_HMS"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKBS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ZSTK_KKA_bk_NAIYO"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_dt_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("DJFK_skosya_sg_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sdyyht_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_skosya_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_skosya_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_skosya_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_skosya_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_skosya_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_skosya_ryohi_kbn"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region



#Region "(SQL)取替票（高圧他）領収書添付ファイル管理番号更新"
    ''' <summary>
    ''' 取替票（高圧他）領収書添付ファイル管理番号更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pRyssyTenpFileMngNo">領収書添付ファイル管理番号</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）領収書添付ファイル管理番号を更新します</remarks>
    Friend Function U015(ByVal pRow As DataRow, ByVal pRyssyTenpFileMngNo As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  upd_date = {0}                   -- 更新日時
 ,sousa_user_id = {1}              -- 更新ユーザ
 ,sousa_appli_cd = {2}             -- 更新プログラム
 ,ryssy_tenp_file_mng_no={3}      --領収書添付ファイル管理番号
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}      --取替票発行年度
 AND  trkhy_kbn={6}      --取替票区分
 AND  trkhy_no={7}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(pRyssyTenpFileMngNo) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 取替票（高圧他）実費_No1更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）実費_No1を更新します</remarks>
    Friend Function U0151(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  no1_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取替票（高圧他）実費_No2更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）実費_No2を更新します</remarks>
    Friend Function U0152(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  no2_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取替票（高圧他）実費_No3更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）実費_No3を更新します</remarks>
    Friend Function U0153(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  no3_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function


    ''' <summary>
    ''' 取替票（高圧他）実費_No4更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）実費_No4を更新します</remarks>
    Friend Function U0154(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  no4_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取替票（高圧他）実費_No3更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pFile_renno">添付書類詳細　連番</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）実費_No5を更新します</remarks>
    Friend Function U0155(ByVal pRow As DataRow, ByVal pFile_renno As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET        -- 取替票（高圧他）
  no5_zippi_no = {0}
WHERE 1 = 1
 AND  dig4_zgsyo_cd={1}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={2}      --取替票発行年度
 AND  trkhy_kbn={3}      --取替票区分
 AND  trkhy_no={4}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(pFile_renno) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try
    End Function
#End Region



#Region "(SQL)取替票（高圧他）(抜取検査)更新"
    ''' <summary>
    ''' 取替票（高圧他）(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pUpldUserId">アップロードUSER_ID</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（高圧他）(抜取検査)を更新します</remarks>
    Friend Function U018(ByVal pRow As DataRow, ByVal pUpldUserId As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET      -- 取替票（高圧他）
  upd_date={0}                  --更新日時
 ,sousa_user_id={1}             --更新ユーザ
 ,sousa_appli_cd={2}            --更新プログラム
 ,nktr_kensa_zyokyo_cd={3}      --抜取検査状況コード
 ,nktrk_tblt_recv_ymd={4}       --抜取検査タブレット受信年月日
 ,nktrk_tblt_recv_tntsy_cd={5}  --抜取検査タブレット受信_担当者コード
 ,nktrk_zissi_ymd={6}           --抜取検査実施年月日
 ,nktrk_zissi_tntsy_cd={7}      --抜取検査実施_担当者コード
 ,nktrk_kka_upld_ymd={8}        --抜取検査結果アップロード年月日
 ,nktrk_kka_upld_tntsy_cd={9}   --抜取検査結果アップロード_担当者コード
WHERE 1 = 1
 AND  dig4_zgsyo_cd={10}        --事業所コード（4桁）
 AND  trkhy_hakko_nendo={11}    --取替票発行年度
 AND  trkhy_kbn={12}            --取替票区分
 AND  trkhy_no={13}             --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktr_kensa_zyokyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_tblt_recv_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_tblt_recv_tntsy_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_zissi_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ntkns_zissi_tntsy_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(SyoriYMD) _
                               , EcOrgIO.EcOrgString.GetSqlText(pUpldUserId) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region


#Region "(SQL)取替_取替票行程（高圧他）(抜取検査)更新"
    ''' <summary>
    ''' 取替_取替票行程（高圧他）(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_取替票行程(抜取検査)を更新します</remarks>
    Friend Function U019(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE TRKE_TRKEHY_PRC_KAHK SET         -- 取替行程（高圧他）
  upd_date={0}      --更新日時
 ,sousa_user_id={1}      --更新ユーザ
 ,sousa_appli_cd={2}      --更新プログラム
 ,nktrk_sizi_syori_ymd={3}      --抜取検査指示_処理年月日
 ,nktrk_sizi_syors_cd={4}      --抜取検査指示_処理者コード
 ,nktrk_sizi_syors_ms={5}      --抜取検査指示_処理者名
 ,nktrk_sizi_ka_cd={6}      --抜取検査指示_課コード
 ,nktrk_sizi_ka_ms={7}      --抜取検査指示_課名
 ,nktrk_sizi_kzkis_mdgt_cd={8}      --抜取検査指示_工事会社窓口コード
 ,nktrk_sizi_kzkis_mdgt_ms={9}      --抜取検査指示_工事会社窓口名
 ,nktrk_sizi_kozitn_cd={10}      --抜取検査指示_工事店コード
 ,nktrk_sizi_kozitn_ms={11}      --抜取検査指示_工事店名
 ,nktrk_tanto_syori_ymd={12}      --抜取検査（担当）_処理年月日
 ,nktrk_tanto_syors_cd={13}      --抜取検査（担当）_処理者コード
 ,nktrk_tanto_syors_ms={14}      --抜取検査（担当）_処理者名
 ,nktrk_tanto_ka_cd={15}      --抜取検査（担当）_課コード
 ,nktrk_tanto_ka_ms={16}      --抜取検査（担当）_課名
 ,nktrk_tanto_ces_cd={17}      --抜取検査（担当）_保安協会コード
 ,nktrk_tanto_ces_ms={18}      --抜取検査（担当）_保安協会名
 ,nktrk_elder_syori_ymd={19}      --抜取検査（上長）_処理年月日
 ,nktrk_elder_syors_cd={20}      --抜取検査（上長）_処理者コード
 ,nktrk_elder_syors_ms={21}      --抜取検査（上長）_処理者名
 ,nktrk_elder_ka_cd={22}      --抜取検査（上長）_課コード
 ,nktrk_elder_ka_ms={23}      --抜取検査（上長）_課名
 ,nktrk_elder_ces_cd={24}      --抜取検査（上長）_保安協会コード
 ,nktrk_elder_ces_ms={25}      --抜取検査（上長）_保安協会名
 ,nktrk_hnn_flg={26}      --抜取検査_否認フラグ
WHERE 1 = 1
 AND  dig4_zgsyo_cd={27}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={28}      --取替票発行年度
 AND  trkhy_kbn={29}      --取替票区分
 AND  trkhy_no={30}      --取替票番号

         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kzkis_mdgt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kzkis_mdgt_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kozitn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_sizi_kozitn_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ces_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_tanto_ces_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syori_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syors_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_syors_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ka_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ka_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ces_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_elder_ces_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("nktrk_hnn_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region



#Region "(SQL)取替_自主点検チェックシート（高圧他）テーブル(抜取検査)更新"
    ''' <summary>
    ''' 取替_自主点検チェックシート（高圧他）テーブル(抜取検査)更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_自主点検チェックシートテーブル(抜取検査)を更新します</remarks>
    Friend Function U020(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集0
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            'Rev056 2025/06/30 託送情報切り替え対応（高圧他） ADD takckrr_smhyz_ntktt_ryohi_kbn ～ takckk_dsipcv_ntktt_ryohi_kbn
            sb.Append(<sql><![CDATA[
UPDATE trke_check_sheet_kahk SET        -- 取替_自主点検チェックシート（高圧他）
  upd_date={0}                          --更新日時
 ,sousa_user_id={1}                     --更新ユーザ
 ,sousa_appli_cd={2}                    --更新プログラム
 ,skosya_tnkn_ymd={3}                   --点検（検査）年月日・施工者
 ,kstts_tnkn_ymd={4}                    --点検（検査）年月日・検査担当者
 ,KKCHYR_snst_ntktt_ryohi_kbn={5}       --高圧計器チェック・変成器・容量・新設・抜取検査担当者・良否区分
 ,KKCHYR_kst_ntktt_ryohi_kbn={6}        --高圧計器チェック・変成器・容量・既設・抜取検査担当者・良否区分
 ,KKCHANO_snst_ntktt_ryohi_kbn={7}      --高圧計器チェック・変成器・合番号・新設・抜取検査担当者・良否区分
 ,KKCHANO_kst_ntktt_ryohi_kbn={8}       --高圧計器チェック・変成器・合番号・既設・抜取検査担当者・良否区分
 ,KKCHSNO_ntktt_ryohi_kbn={9}           --高圧計器チェック・変成器・製造番号・抜取検査担当者・良否区分
 ,KKCHKKL_ntktt_ryohi_kbn={10}          --高圧計器チェック・変成器・結線・KL・抜取検査担当者・良否区分
 ,KKCHKSS_ntktt_ryohi_kbn={11}          --高圧計器チェック・変成器・結線・正相・抜取検査担当者・良否区分
 ,KKCHKIB_ntktt_ryohi_kbn={12}          --高圧計器チェック・変成器・結線・色別・抜取検査担当者・良否区分
 ,KKCHKHD_ntktt_ryohi_kbn={13}          --高圧計器チェック・変成器・結線・半田・抜取検査担当者・良否区分
 ,KKCHKBS_ntktt_ryohi_kbn={14}          --高圧計器チェック・変成器・結線・ビス締付・抜取検査担当者・良否区分
 ,KKCKZR_snst_ntktt_ryohi_kbn={15}      --高圧計器チェック・計器・乗率・新設・抜取検査担当者・良否区分
 ,KKCKANO_snst_ntktt_ryohi_kbn={16}     --高圧計器チェック・計器・合番号・新設・抜取検査担当者・良否区分
 ,KKCKSNO_ntktt_ryohi_kbn={17}          --高圧計器チェック・計器・製造番号・抜取検査担当者・良否区分
 ,KKCKDJ_dosa_ntktt_ryohi_kbn={18}      --高圧計器チェック・計器・動作状況・動作・抜取検査担当者・良否区分
 ,KKCKDJ_settjk_ntktt_ryohi_kbn={19}    --高圧計器チェック・計器・動作状況・設定時刻・抜取検査担当者・良否区分
 ,KKCKKS_ib_ntktt_ryohi_kbn={20}        --高圧計器チェック・計器・結線・色別・抜取検査担当者・良否区分
 ,KKCKKS_hndag_ntktt_ryohi_kbn={21}     --高圧計器チェック・計器・結線・半田揚げ・抜取検査担当者・良否区分
 ,KKCKKS_rosyt_ntktt_ryohi_kbn={22}     --高圧計器チェック・計器・結線・露出・抜取検査担当者・良否区分
 ,KKCKKS_bsst_ntktt_ryohi_kbn={23}      --高圧計器チェック・計器・結線・ビス締付・抜取検査担当者・良否区分
 ,KKCKKD_tansi_ntktt_ryohi_kbn={24}     --高圧計器チェック・計器・検電・端子・抜取検査担当者・良否区分
 ,KKCHTSN_tekyo_ntktt_ryohi_kbn={25}    --高圧計器チェック・表示端末・製造番号・撤去・抜取検査担当者・良否区分
 ,KAKCTKVCT_ntktt_ryohi_kbn={26}        --高圧計器チェック・チェックターミナル・結線・ＶＣＴ・抜取検査担当者・良否区分
 ,KAKCTKIB_ntktt_ryohi_kbn={27}         --高圧計器チェック・チェックターミナル・結線・色別・抜取検査担当者・良否区分
 ,KAKCTKHD_ntktt_ryohi_kbn={28}         --高圧計器チェック・チェックターミナル・結線・半田・抜取検査担当者・良否区分
 ,KAKCTKBS_ntktt_ryohi_kbn={29}         --高圧計器チェック・チェックターミナル・結線・ビス締付・抜取検査担当者・良否区分
 ,KAKCTKSS_ntktt_ryohi_kbn={30}         --高圧計器チェック・チェックターミナル・結線・接栓締付・抜取検査担当者・良否区分
 ,kakct_szno_ntktt_ryohi_kbn={31}       --高圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,kakcs_ksn_ntktt_ryohi_kbn={32}        --高圧計器チェック・その他・結線・抜取検査担当者・良否区分
 ,kakcs_kns_ntktt_ryohi_kbn={33}        --高圧計器チェック・その他・検相・抜取検査担当者・良否区分
 ,kakcs_skphoto_ntktt_ryohi_kbn={34}    --高圧計器チェック・その他・竣工写真・抜取検査担当者・良否区分
 ,kakcs_kknn_ntktt_ryohi_kbn={35}       --高圧計器チェック・その他・最終確認・抜取検査担当者・良否区分
 ,kakcz_ntktt_ryohi_kbn={36}            --高圧計器チェック・乗率チェック・抜取検査担当者・良否区分
 ,kakca_ntktt_ryohi_kbn={37}            --高圧計器チェック・合番号チェック・抜取検査担当者・良否区分
 ,ZSTK_KKA_bk_NAIYO={38}                --自主点検結果_備考内容
 ,KAHSK_ntktt_1s_kstckk_flg={39}        --高圧変成器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,KAHSK_ntktt_p1_kstckk_flg={40}        --高圧変成器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,KAHSK_ntktt_p3_kstckk_flg={41}        --高圧変成器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,KAHSK_ntktt_3s_kstckk_flg={42}        --高圧変成器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,KAHSK_ntktt_3l_kstckk_flg={43}        --高圧変成器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,KAHSK_ntktt_p2_kstckk_flg={44}        --高圧変成器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,KAHSK_ntktt_1l_kstckk_flg={45}        --高圧変成器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_1s_kstckk_flg={46}     --チェックターミナル・VCT・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_p1_kstckk_flg={47}     --チェックターミナル・VCT・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_p3_kstckk_flg={48}     --チェックターミナル・VCT・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_3s_kstckk_flg={49}     --チェックターミナル・VCT・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_3l_kstckk_flg={50}     --チェックターミナル・VCT・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_p2_kstckk_flg={51}     --チェックターミナル・VCT・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,cktrmvct_ntktt_1l_kstckk_flg={52}     --チェックターミナル・VCT・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_1s_kstckk_flg={53}      --チェックターミナル・WH・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_p1_kstckk_flg={54}      --チェックターミナル・WH・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_p3_kstckk_flg={55}      --チェックターミナル・WH・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_3s_kstckk_flg={56}      --チェックターミナル・WH・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_3l_kstckk_flg={57}      --チェックターミナル・WH・黄（3L）・抜取検査担当者・結線等チェック結果フラグ
 ,cktrmwh_ntktt_p2_kstckk_flg={58}      --チェックターミナル・WH・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,cktrmwh_ntktt_1l_kstckk_flg={59}      --チェックターミナル・WH・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,jkykrk_ntktt_1s_kstckk_flg={60}       --受給用計量器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,jkykrk_ntktt_p1_kstckk_flg={61}       --受給用計量器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,jkykrk_ntktt_p3_kstckk_flg={62}       --受給用計量器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,jkykrk_ntktt_3s_kstckk_flg={63}       --受給用計量器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,jkykrk_ntktt_3l_kstckk_flg={64}       --受給用計量器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,jkykrk_ntktt_p2_kstckk_flg={65}       --受給用計量器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,jkykrk_ntktt_1l_kstckk_flg={66}       --受給用計量器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,takcd_ttt_ntktt_ryohi_kbn={67}        --低圧計器チェック・電力量計・取付取替等・抜取検査担当者・良否区分
 ,takcd_szsu_ntktt_ryohi_kbn={68}       --低圧計器チェック・電力量計・指示数・抜取検査担当者・良否区分
 ,takck_ssda_ntktt_ryohi_kbn={69}       --低圧計器チェック・計器項目・相線電圧・抜取検査担当者・良否区分
 ,takck_zyrt_ntktt_ryohi_kbn={70}       --低圧計器チェック・計器項目・乗率・抜取検査担当者・良否区分
 ,takck_yr_ntktt_ryohi_kbn={71}         --低圧計器チェック・計器項目・容量・抜取検査担当者・良否区分
 ,takck_sm_ntktt_ryohi_kbn={72}         --低圧計器チェック・計器項目・SM・抜取検査担当者・良否区分
 ,takck_kkano_ntktt_ryohi_kbn={73}      --低圧計器チェック・計器項目・計器・合番号・抜取検査担当者・良否区分
 ,takck_hksno_ntktt_ryohi_kbn={74}      --低圧計器チェック・計器項目・変流器・製造番号・抜取検査担当者・良否区分
 ,takck_hkano_ntktt_ryohi_kbn={75}      --低圧計器チェック・計器項目・変流器・合番号・抜取検査担当者・良否区分
 ,takckk_krkkd_ntktt_ryohi_kbn={76}     --低圧計器チェック・構造検査・計量器・検電・抜取検査担当者・良否区分
 ,takckk_krkb_ntktt_ryohi_kbn={77}      --低圧計器チェック・構造検査・計量器箱・抜取検査担当者・良否区分
 ,takcks_kkzen_ntktt_ryohi_kbn={78}     --低圧計器チェック・結線等・計器・絶縁・抜取検査担当者・良否区分
 ,takcks_kksl_ntktt_ryohi_kbn={79}      --低圧計器チェック・結線等・計器・SL・抜取検査担当者・良否区分
 ,takcks_kkhks_ntktt_ryohi_kbn={80}     --低圧計器チェック・結線等・計器・複数・抜取検査担当者・良否区分
 ,takcks_kkgsz_ntktt_ryohi_kbn={81}     --低圧計器チェック・結線等・計器・誤接続・抜取検査担当者・良否区分
 ,takcks_kkaki_ntktt_ryohi_kbn={82}     --低圧計器チェック・結線等・計器・赤色・抜取検査担当者・良否区分
 ,takcks_kkyrs_ntktt_ryohi_kbn={83}     --低圧計器チェック・結線等・計器・より線・抜取検査担当者・良否区分
 ,takcks_kkrst_ntktt_ryohi_kbn={84}     --低圧計器チェック・結線等・計器・露出・抜取検査担当者・良否区分
 ,takcks_kkbst_ntktt_ryohi_kbn={85}     --低圧計器チェック・結線等・計器・ビス締付・抜取検査担当者・良否区分
 ,takcks_kikib_ntktt_ryohi_kbn={86}     --低圧計器チェック・結線等・計器・色別・抜取検査担当者・良否区分
 ,takcks_kkssu_ntktt_ryohi_kbn={87}     --低圧計器チェック・結線等・計器・正相・抜取検査担当者・良否区分
 ,takcks_hrkkl_ntktt_ryohi_kbn={88}     --低圧計器チェック・結線等・変流器・ＫＬ・抜取検査担当者・良否区分
 ,takcks_hp123_ntktt_ryohi_kbn={89}     --低圧計器チェック・結線等・変流器・Ｐ１Ｐ２Ｐ３・抜取検査担当者・良否区分
 ,takcks_hkaki_ntktt_ryohi_kbn={90}     --低圧計器チェック・結線等・変流器・赤色・抜取検査担当者・良否区分
 ,takckrk_hnrk_ntktt_ryohi_kbn={91}     --低圧計器チェック・計器回転・検電・変流器・抜取検査担当者・良否区分
 ,takckrr_sm_ntktt_ryohi_kbn={92}       --低圧計器チェック・計器回転・回転・SM・抜取検査担当者・良否区分
 ,takcc_led_ntktt_ryohi_kbn={93}        --低圧計器チェック・通信部・ＬＥＤ・抜取検査担当者・良否区分
 ,takcc_souck_ntktt_ryohi_kbn={94}      --低圧計器チェック・通信部・装着・抜取検査担当者・良否区分
 ,takct_szno_ntktt_ryohi_kbn={95}       --低圧計器チェック・撤去計器等・製造番号・抜取検査担当者・良否区分
 ,takct_aino_ntktt_ryohi_kbn={96}       --低圧計器チェック・撤去計器等・合番号・抜取検査担当者・良否区分
 ,takcs_photo_ntktt_ryohi_kbn={97}      --低圧計器チェック・その他・写真・抜取検査担当者・良否区分
 ,takcs_tktrk_ntktt_ryohi_kbn={98}      --低圧計器チェック・その他・撤去取替・抜取検査担当者・良否区分
 ,takcz_ntktt_seigo_kknn_kbn={99}       --低圧計器チェック・乗率チェック・抜取検査担当者・整合確認区分
 ,takca_ntktt_itti_kknn_kbn={100}       --低圧計器チェック・合番号チェック・抜取検査担当者・一致確認区分
 ,kizi_ntktt_naiyo={101}                --記事（抜取検査担当者）内容
 ,hrktkk_ntktt_1s_kstckk_flg={102}      --変流器付計器・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p1_kstckk_flg={103}      --変流器付計器・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p3_kstckk_flg={104}      --変流器付計器・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,hrktkk_ntktt_3s_kstckk_flg={105}      --変流器付計器・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,hrktkk_ntktt_3l_kstckk_flg={106}      --変流器付計器・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,hrktkk_ntktt_p2_kstckk_flg={107}      --変流器付計器・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,hrktkk_ntktt_1l_kstckk_flg={108}      --変流器付計器・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,p1_ntktt_1s_kstckk_flg={109}          --P1・抜取検査担当者・黒（1S）・結線等チェック結果フラグ
 ,p1_ntktt_p1_kstckk_flg={110}          --P1・抜取検査担当者・赤（P1）・結線等チェック結果フラグ
 ,p1_ntktt_1l_kstckk_flg={111}          --P1・抜取検査担当者・緑（1L）・結線等チェック結果フラグ
 ,p2_ntktt_p2_kstckk_flg={112}          --P2・抜取検査担当者・白（P2）・結線等チェック結果フラグ
 ,p3_ntktt_3s_kstckk_flg={113}          --P3・抜取検査担当者・茶（3S）・結線等チェック結果フラグ
 ,p3_ntktt_p3_kstckk_flg={114}          --P3・抜取検査担当者・青（P3）・結線等チェック結果フラグ
 ,p3_ntktt_3l_kstckk_flg={115}          --P3・抜取検査担当者・黄（3L）・結線等チェック結果フラグ
 ,takckrr_smhyz_ntktt_ryohi_kbn={116}   --低圧計器チェック・計器回転・回転・SM表示・抜取検査担当者・良否区分
 ,takckrr_d2sm_ntktt_ryohi_kbn={117}    --低圧計器チェック・計器回転・回転・D2SM・抜取検査担当者・良否区分
 ,takckrr_khk_ntktt_ryohi_kbn={118}     --低圧計器チェック・計器回転・回転・開閉器・抜取検査担当者・良否区分
 ,takckk_okgkb_ntktt_ryohi_kbn={119}    --低圧計器チェック・構造検査・計量器箱・屋外計器箱・抜取検査担当者・良否区分
 ,takckk_dsipcv_ntktt_ryohi_kbn={120}   --低圧計器チェック・構造検査・計量器箱・電線隠蔽用カバー・抜取検査担当者・良否区分
WHERE 1 = 1
 AND  dig4_zgsyo_cd={121}               --事業所コード（4桁）
 AND  trkhy_hakko_nendo={122}           --取替票発行年度
 AND  trkhy_kbn={123}                   --取替票区分
 AND  trkhy_no={124}                    --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("skosya_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kstts_tnkn_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHYR_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHANO_kst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHSNO_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKKL_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKSS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKIB_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKHD_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHKBS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKZR_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKANO_snst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKSNO_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_dosa_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKDJ_settjk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_ib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_hndag_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_rosyt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKS_bsst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCKKD_tansi_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KKCHTSN_tekyo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKVCT_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKIB_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKHD_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKBS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAKCTKSS_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_ksn_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kns_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_skphoto_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcs_kknn_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakcz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kakca_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ZSTK_KKA_bk_NAIYO"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KAHSK_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmvct_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cktrmwh_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("jkykrk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_ttt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcd_szsu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_ssda_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_zyrt_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_yr_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_kkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hksno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takck_hkano_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkkd_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_krkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkzen_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kksl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkhks_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkgsz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkyrs_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkrst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkbst_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kikib_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_kkssu_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hrkkl_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hp123_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcks_hkaki_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrk_hnrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_led_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcc_souck_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_szno_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takct_aino_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_photo_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcs_tktrk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takcz_ntktt_seigo_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takca_ntktt_itti_kknn_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kizi_ntktt_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("hrktkk_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_p1_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p1_ntktt_1l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p2_ntktt_p2_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3s_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_p3_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("p3_ntktt_3l_kstckk_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_smhyz_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_d2sm_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckrr_khk_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_okgkb_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("takckk_dsipcv_ntktt_ryohi_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region



#Region "(SQL)取替票（高圧他）(追加工費および工事費合計)更新"
    ''' <summary>
    ''' 取替票（低圧）(追加工費および工事費合計)更新
    ''' </summary>
    ''' <param name="dig4_zgsyo_cd">事業所コード（4桁）</param>
    ''' <param name="trkhy_hakko_nendo">取替票発行年度</param>
    ''' <param name="trkhy_kbn">取替票区分</param>
    ''' <param name="trkhy_no">取替票番号</param>
    ''' <param name="TUIKAKH_umu_flg">追加工費有無フラグ</param>
    ''' <param name="trkhy_trtkKojiHi">追加工費　取付工事費</param>
    ''' <param name="trkhy_tekyoKojiHi">追加工費　撤去工事費</param>
    ''' <param name="gokei_kngk">工事費合計</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替票（低圧）(追加工費および工事費合計)を更新します</remarks>
    Friend Function U016(ByVal dig4_zgsyo_cd As String, ByVal trkhy_hakko_nendo As String, ByVal trkhy_kbn As String, ByVal trkhy_no As String,
                         ByVal TUIKAKH_umu_flg As String,
                         ByVal trkhy_trtkKojiHi As String, ByVal trkhy_tekyoKojiHi As String, ByVal gokei_kngk As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_trkehy_kahk SET         -- 取替票（高圧他）
  tuika_kohi_trtk_kozih_kngk={0}    --追加工費_取付工事費
 ,tuika_kohi_tekyo_kozih_kngk={1}   --追加工費_撤去工事費
 ,tuika_kohi_umu_flg={2}            --追加工費有無フラグ
 ,kozih_gokei_kngk={3}              --工事費合計
WHERE 1 = 1
 AND  dig4_zgsyo_cd={4}             --事業所コード（4桁）
 AND  trkhy_hakko_nendo={5}         --取替票発行年度
 AND  trkhy_kbn={6}                 --取替票区分
 AND  trkhy_no={7}                  --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 HdSysBasePg.Cmn.Func.GetSqlNumberText(trkhy_trtkKojiHi) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(trkhy_tekyoKojiHi) _
                               , EcOrgIO.EcOrgString.GetSqlText(TUIKAKH_umu_flg) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(gokei_kngk) _
                               , EcOrgIO.EcOrgString.GetSqlText(dig4_zgsyo_cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_hakko_nendo) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_kbn) _
                               , EcOrgIO.EcOrgString.GetSqlText(trkhy_no)
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region



#Region "(SQL)共通_個人情報アクセスログ"
    ''' <summary>
    ''' 共通_個人情報アクセスログ
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>共通_個人情報アクセスログを追加します</remarks>
    Friend Function I005(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO COM_KIFRFLG        -- 共通_個人情報アクセスログ
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --個人情報参照ログ管理番号
 ,{5}          --個人情報参照ログ年月日
 ,{6}          --個人情報参照ログ時分秒
 ,{7}          --オンライン識別コード
 ,{8}          --機能コード
 ,{9}          --ユーザID
 ,{10}          --お客さま名カナ
 ,{11}          --お客さま名
 ,{12}          --市外局番
 ,{13}          --市内番号
 ,{14}          --番号
 ,{15}          --個人情報参照ログ金融機関コード
 ,{16}          --個人情報参照ログ店舗名称
 ,{17}          --個人情報参照ログ預金種別コード
 ,{18}          --個人情報参照ログ口座番号
 ,{19}          --住所表示フラグ
 ,{20}          --IPアドレス
 ,{21}          --システム独自項目
)
         ]]></sql>.Value)

            Dim wkstr As String = ""
            If Not myDB.GetSequenceNo("seq_kifrflg_mng_renno", wkstr, HdPostgre.HdPostgreDb.SequenceMode.nextval) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, "個人情報参照ログ管理番号　シーケンス　nextvalエラー")
                Return False
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_appli_cd"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wkstr) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("oln_skbt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kino_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cs_mskn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cs_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tel_sgikk_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tel_sinai_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tel_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_knykkn_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_tenpo_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_yokin_sbt_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kifrflg_kouza_no"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("add_hyz_flg"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ip_adrs"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("system_dkzkm"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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

#Region "添付書類詳細　高圧他のファイル連番の最大値検索"
    ''' <summary>
    ''' 添付書類詳細　高圧他のファイル連番の最大値を検索する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S005K(ByRef dbr As HdPostgre.HdPostgreReader, ByVal k1 As String, ByVal k2 As String, ByVal k3 As String, ByVal k4 As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
SELECT
    MAX(file_renno) as max_file_renno
FROM
    trke_trkehy_kahk A
    left join COM_TNPDOC_DTL B ON (B.tnpdoc_mng_renno=A.tenp_file_mng_no)
WHERE 1 = 1
    AND dig4_zgsyo_cd = {0}
    AND trkhy_hakko_nendo = {1}
    AND trkhy_kbn =  {2}
    AND trkhy_no = {3}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString _
                               , EcOrgIO.EcOrgString.GetSqlText(k1) _
                               , EcOrgIO.EcOrgString.GetSqlText(k2) _
                               , EcOrgIO.EcOrgString.GetSqlText(k3) _
                               , EcOrgIO.EcOrgString.GetSqlText(k4))
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

    ''' <summary>
    ''' 添付書類詳細　高圧他の最大ファイル連番取得
    ''' </summary>
    ''' <param name="pFileRenno">添付書類詳細　最大ファイル連番</param>
    ''' <param name="k1">事業所コード（4桁）</param>
    ''' <param name="k2">取替票発行年度</param>
    ''' <param name="k3">取替票区分</param>
    ''' <param name="k4">取替票番号</param>
    ''' <remarks></remarks>
    Function getMaxFileRennoKahk(ByRef pFileRenno As String, ByVal k1 As String, ByVal k2 As String, ByVal k3 As String, ByVal k4 As String) As Boolean
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader

        pFileRenno = "0"

        '添付書類詳細　最大ファイル連番を検索する。
        If Not S005K(dbr, k1, k2, k3, k4) Then
            Return False
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                pFileRenno = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("max_file_renno"), ""))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return True
    End Function
#End Region


#Region "(SQL)共通_ログ収集"
    ''' <summary>
    ''' 共通_ログ収集
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>共通_ログ収集を追加します</remarks>
    Friend Function I006(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO COM_LOG_SYUSYU        -- 共通_ログ収集
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --ログ管理番号
 ,{5}          --ログ作成年月日
 ,{6}          --ログ作成時刻
 ,{7}          --ログ収集レベル
 ,{8}          --機能レベル
 ,{9}          --機能コード
 ,{10}          --ログ内容
 ,{11}          --ユーザID
 ,{12}          --会社コード
 ,{13}          --事業所コード
 ,{14}          --担当コード
 ,{15}          --アプリケーションバージョン
 ,{16}          --携帯端末番号
)
         ]]></sql>.Value)

            Dim wkstr As String = ""
            If Not myDB.GetSequenceNo("SEQ_LOG_SYUSYU", wkstr, HdPostgre.HdPostgreDb.SequenceMode.nextval) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, "共通_ログ収集　ログ管理番号 シーケンス　nextvalエラー")
                Return False
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_appli_cd"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wkstr) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("log_saksi_ymd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("log_saksi_hms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("log_syusyu_lv"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kino_lv"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("kino_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("log_naiyo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("corp_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tanto_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("APPLI_VER"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("KTTNMT_NO"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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

#Region "(SQL)共通_ログイン認証履歴"
    ''' <summary>
    ''' 共通_ログイン認証履歴
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>共通_ログイン認証履歴を追加します</remarks>
    Friend Function I007(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO COM_LOGIN_NNSY_RK        -- 共通_ログイン認証履歴
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --ログイン認証履歴番号
 ,{5}          --ログイン認証年月日
 ,{6}          --ユーザID
 ,{7}          --ユーザパスワード
 ,{8}          --会社コード
 ,{9}          --事業所コード（4桁）
 ,{10}          --担当コード
 ,{11}          --IPアドレス
 ,{12}          --ログイン認証箇所コード
 ,{13}          --ログイン認証結果コード
)
         ]]></sql>.Value)

            Dim wkstr As String = ""
            If Not myDB.GetSequenceNo("seq_login_nnsy_rireki_renno", wkstr, HdPostgre.HdPostgreDb.SequenceMode.nextval) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, "ログイン認証履歴番号　シーケンス　nextvalエラー")
                Return False
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_appli_cd"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wkstr) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("login_nnsy_date"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("user_pwd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("corp_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("tanto_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ip_adrs"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("login_nnsy_kasyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("login_nnsy_kka_cd"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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

#Region "(SQL)共通_ログイン履歴データ"
    ''' <summary>
    ''' 共通_ログイン履歴データ
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>共通_ログイン履歴データを追加します</remarks>
    Friend Function I008(ByVal pRow As DataRow) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO COM_LOGIN_RIREKI_DATA        -- 共通_ログイン履歴データ
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,{4}          --ログインログ番号
 ,{5}          --コンピュータ名称
 ,{6}          --IPアドレス
 ,{7}          --検索用年月日
)
         ]]></sql>.Value)

            Dim wkstr As String = ""
            If Not myDB.GetSequenceNo("seq_log_mng_renno", wkstr, HdPostgre.HdPostgreDb.SequenceMode.nextval) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, "ログインログ番号　シーケンス　nextvalエラー")
                Return False
            End If

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_user_id"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("sousa_appli_cd"), ""))) _
                               , HdSysBasePg.Cmn.Func.GetSqlNumberText(wkstr) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("cmpt_ms"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("ip_adrs"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("src_ymd"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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


#Region "(SQL)取替_ダウンロード取替票管理更新"
    ''' <summary>
    ''' 取替_ダウンロード取替票管理更新
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pKTTNMT_NO">携帯端末番号</param>
    ''' <param name="pKeikiDwldSbtCd">計器ダウンロード種別コード</param>
    ''' <param name="pUpldUserId">アップロードUSER_ID</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>取替_ダウンロード取替票管理を更新します</remarks>
    Friend Function U011(ByVal pRow As DataRow, ByVal pKTTNMT_NO As String, ByVal pKeikiDwldSbtCd As String, ByVal pUpldUserId As String) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String
        Dim procDateTime As String = System.DateTime.Now.ToString("yyyyMMddHHmmss") 'getDbDateTime()

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
UPDATE trke_downld_trkehy_mng SET        -- 取替_ダウンロード取替票管理
  upd_date={0}      --更新日時
 ,sousa_appli_cd={1}      --更新プログラム
 ,mtds_file_zti_cd={2}      --持出ファイル状態コード
 ,keiki_user_id={3}      --アップロードユーザーID
WHERE 1 = 1
 AND  KTTNMT_NO={4}      --携帯端末番号
 AND  keiki_dwld_sbt_cd={5}      --ダウンロード種別コード
 AND  dig4_zgsyo_cd={6}      --事業所コード（4桁）
 AND  trkhy_hakko_nendo={7}      --取替票発行年度
 AND  trkhy_kbn={8}      --取替票区分
 AND  trkhy_no={9}      --取替票番号
         ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(procDateTime) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(MjK1.Cmn.Const.MtdsFileZtiCd.MtdsFileUpload) _
                               , EcOrgIO.EcOrgString.GetSqlText(pUpldUserId) _
                               , EcOrgIO.EcOrgString.GetSqlText(pKTTNMT_NO) _
                               , EcOrgIO.EcOrgString.GetSqlText(pKeikiDwldSbtCd) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))) _
                               , EcOrgIO.EcOrgString.GetSqlText(CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), "")))
                               )
            '*************************************************************************

            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            Dim wCount As Double = 0
            If Not Me.myDB.ExecDML(sqlstr, wCount) Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception(Me.myDB.ErrDescription)
            End If
            If wCount = 0 Then
                HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)
                Throw New Exception("更新結果0件エラー")
            End If

            'Trueを返却する。
            Return True

        Catch ex As Exception
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L0, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, String.Format("発生箇所:{0}, 内容:{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message))
            Return False
        End Try

    End Function
#End Region


#Region "取替票　取付計器重複検索"
    ''' <summary>
    ''' 取替票　取付計器重複検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pRow">取替票</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S003(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pRow As DataRow) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Dim k1 As String = CStr(Me.myDB.ConvDbNull(pRow("dig4_zgsyo_cd"), ""))
        Dim k2 As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_hakko_nendo"), ""))
        Dim k3 As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_kbn"), ""))
        Dim k4 As String = CStr(Me.myDB.ConvDbNull(pRow("trkhy_no"), ""))
        Dim wttkk_keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), "")).Trim

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
select 
 *
from
 TRKE_TRKEHY_TIAT A
WHERE 1=1
AND A.ttkk_keiki_id={0}
AND A.trkhy_hakko_nendo={2}
AND NOT (A.dig4_zgsyo_cd={1} and A.trkhy_hakko_nendo={2} and trkhy_kbn={3} and trkhy_no={4})
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(wttkk_keiki_id) _
                               , EcOrgIO.EcOrgString.GetSqlText(k1) _
                               , EcOrgIO.EcOrgString.GetSqlText(k2) _
                               , EcOrgIO.EcOrgString.GetSqlText(k3) _
                               , EcOrgIO.EcOrgString.GetSqlText(k4)
                            )
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
#End Region

    'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD Start
#Region "小数切り捨て"
    ''' <summary>
    ''' 小数指定桁未満切り捨て
    ''' </summary>
    ''' <param name="numb">数値</param>
    ''' <param name="keta">桁位置</param>
    ''' <returns>Double型値</returns>
    ''' <remarks></remarks>
    Function TruncateDbl(ByVal numb As Double, ByVal keta As Integer) As Double
        Dim multiplier As Double = CDbl(Math.Pow(10, keta))
        Return Math.Truncate(numb * multiplier) / multiplier
    End Function

    Function TruncateDec(ByVal numb As Decimal, ByVal keta As Integer) As Decimal
        Dim multiplier As Decimal = CDec(Math.Pow(10, keta))
        Return Math.Truncate(numb * multiplier) / multiplier
    End Function
#End Region
    'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD End

    'Rev076 20250509 庫入情報送信エラー対応 ADD Start
#Region "(SQL)SMJ計器情報を検索する。"
    ''' <summary>
    ''' SMJ計器情報テーブルより，計器情報を取得する。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pKeikiId">計器ID</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S007(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pKeikiId As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
SELECT
 keiki_id                                              -- 計器ＩＤ
 , krhsk_cd                                            -- 計量方式コード
 , RIGHT(keiki_yr, 3) AS keiki_yr                      -- 計器容量
 , tshsk_cd                                            -- 通信方式コード
 , keiki_ktsk_ms                                       -- 計器型式名称
 , taiko_kbn                                           -- 耐候区分
 , knthk_cd                                            -- 検定方向コード
 , ts_kino_umu_flg                                     -- タイムスイッチ機能有無フラグ
 , khk_kino_umu_flg                                    -- 開閉器機能有無フラグ
 , yuko_kigen_ym                                       -- 有効期限年月
 , seizo_yy                                            -- 表示器・製造年
 , RIGHT(zyrt, 3) AS zyrt                              -- 乗率
 , digsu                                               -- 桁数
 , tnsb_szsya_cd                                       -- 端子部_製造者コード
 , tnsb_seizo_yy                                       -- 端子部製造年
 , krds_zytr_yuko_wh                                   -- 庫出_順潮流（取付時の指示数にセット）
 , krds_gytr_yuko_wh                                   -- 庫出_逆潮流（取付時の指示数にセット）
 , krir_zytr_yuko_wh                                   -- 庫入_順潮流（撤去の指示数にセット）
 , krir_gytr_yuko_wh                                   -- 庫入_逆潮流（撤去の指示数にセット）
 , RIGHT(no1_hnsik_seizo_no, 2) AS no1_hnsik_seizo_no  -- 変成器_製造ＮＯ１
 , RIGHT(no2_hnsik_seizo_no, 2) AS no2_hnsik_seizo_no  -- 変成器_製造ＮＯ２
 , hnsik_yuko_kigen_ym                                 -- 変成器_有効期限
 , hnsik_gknti_no                                      -- 変成器_原検定番号
 , krds_ymd                                            -- 撤去_庫出年月日
 , krir_ymd                                            -- 撤去_庫入年月日
FROM
 view_smj_keiki_info
WHERE
 keiki_id = {0}                                        -- 計器ＩＤ
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(pKeikiId)
                            )
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

    ''' <summary>
    ''' SMJ計器情報テーブルより，計器情報を取得する（高圧）。
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <param name="pKeikiId">計器ID</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks></remarks>
    Friend Function S008(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pKeikiId As String) As Boolean

        '*************************************************************************
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try

            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
SELECT
 krir_zenit_yuko_wh_now AS tkkk_sougo_szsu             -- 庫入全日有効電力量(現在値) ＝ 撤去計器_指示数_総合
 , krir_max_dm AS tkkk_max_zyyo_szsu                   -- 庫入最大DM(現在値) ＝ 撤去計器_指示数_最大需要
 , krir_rkrt_yuko_wh_now AS tkkk_rsyk_szsu             -- 庫入力測有効電力量(現在値) ＝ 撤去計器_指示数_力測有効
 , krir_rkrt_muko_wh_okr AS tkkk_rsmk_szsu             -- 庫入力測無効電力量(現在値)遅れ ＝ 撤去計器_指示数_力測無効
FROM
 view_smj_kk_keiki_info 
WHERE
 keiki_id = {0}                                        -- 計器ＩＤ
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(pKeikiId)
                            )
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
    'Rev076 20250509 庫入情報送信エラー対応 ADD End


    'Rev040 2024/12/26 物理分割 計器マスタ対応 ADD Start
#Region "(SQL)計器マスタWebAPI実行管理登録"
    ''' <summary>
    ''' WEBAPI実行管理　登録
    ''' </summary>
    ''' <param name="pRow">レコードデータ</param>
    ''' <param name="pKbn">区分</param>
    ''' <returns>true・falseが返却される</returns>
    ''' <remarks>WEBAPI実行管理　登録</remarks>
    Friend Function I011(ByVal pRow As DataRow, ByVal pKbn As Integer) As Boolean

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        Try
            'SQL文字列編集
            '*************************************************************************
            '【サンプル】C2.3.13_SQL詳細設計書からSQLをコピーして下さい。
            sb.Append(<sql><![CDATA[
INSERT INTO keiki_webapi_exec_mng (      -- WEBAPI実行管理
  saksi_date
 ,upd_date
 ,sousa_user_id
 ,sousa_appli_cd
 ,exec_sizi_no
 ,keiki_id
 ,webapi_url_naiyo
 ,webapi_methd_naiyo
 ,webapi_req_data_naiyo
 ,exec_zyokyo_status_cd
 ,syori_time
 ,webapi_rspns_status_cd
 ,webapi_rspns_err_msg_naiyo
)
VALUES (
  {0}          --作成日時
 ,{1}          --更新日時
 ,{2}          --更新ユーザ
 ,{3}          --更新プログラム
 ,nextval('seq_exec_sizi_no')
 ,{4}          --計器ID
 ,{5}          --WEBAPIURL内容
 ,{6}          --WEBAPIメソッド内容
 ,{7}          --WEBAPIリクエストデータ内容
 ,'1'          --実行状況ステータスコード  1:実行待
 ,' '          --処理時間
 ,' '          --WEBAPIレスポンスステータスコード
 ,' '          --WEBAPIレスポンスエラーメッセージ内容
)
         ]]></sql>.Value)

            '区分ごとの処理
            Dim keiki_id As String = ""
            Dim webapi_url_naiyo As String = ""
            Dim webapi_methd_naiyo As String = "POST"
            Dim webapi_req_data_naiyo As String = ""

            Select Case pKbn
                Case [Const].C_KEIKI_MASTER_RENKEI_IF023 '取付完了通知
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/ttk_kar_tti"
                    webapi_req_data_naiyo = get_req_naiyo_IF023IF024(pRow, keiki_id)
                Case [Const].C_KEIKI_MASTER_RENKEI_IF024 '撤去完了通知
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("tkkk_keiki_id"), ""))
                    'Rev093 廃滅と撤去完了逆転の抑止(JSON作成を取付・撤去で共用していたが撤去を分離) MOD Start
                    'webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/teo_kar_tti"
                    'webapi_req_data_naiyo = get_req_naiyo_IF023IF024(pRow, keiki_id)
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/teo_kar_tti"
                    webapi_req_data_naiyo = get_req_naiyo_IF024(pRow)
                    'Rev093 廃滅と撤去完了逆転の抑止 MOD End
                Case [Const].C_KEIKI_MASTER_RENKEI_IF025G1 '設備更新通知(端子部再用　第１世代)
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stb"
                    webapi_req_data_naiyo = get_req_naiyo_IF025G1(pRow)
                Case [Const].C_KEIKI_MASTER_RENKEI_IF025G2 '設備更新通知(端子部再用　第２世代)
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stb"
                    webapi_req_data_naiyo = get_req_naiyo_IF025G2(pRow)
                Case [Const].C_KEIKI_MASTER_RENKEI_IF025KH '設備更新通知(高圧　変成器) 'Rev056 高圧対応
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))     'Rev056 高圧対応
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stb"                 'Rev056 高圧対応
                    webapi_req_data_naiyo = get_req_naiyo_IF025KH(pRow)                'Rev056 高圧対応
                Case [Const].C_KEIKI_MASTER_RENKEI_IF023KH '取付完了通知（高圧他）     'Rev056 高圧対応
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))     'Rev056 高圧対応
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/ttk_kar_tti"     'Rev056 高圧対応
                    webapi_req_data_naiyo = get_req_naiyo_IF023IF024KH(pRow, keiki_id) 'Rev056 高圧対応
                Case [Const].C_KEIKI_MASTER_RENKEI_IF024KH '撤去完了通知（高圧他）     'Rev056 高圧対応
                    keiki_id = CStr(Me.myDB.ConvDbNull(pRow("tkkk_keiki_id"), ""))     'Rev056 高圧対応
                    'Rev093 廃滅と撤去完了逆転の抑止(JSON作成を取付・撤去で共用していたが撤去を分離) MOD Start
                    'webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/teo_kar_tti"     'Rev056 高圧対応
                    'webapi_req_data_naiyo = get_req_naiyo_IF023IF024KH(pRow, keiki_id) 'Rev056 高圧対応
                    webapi_url_naiyo = getSETUBI_MNG_SYS_URL() & "stt/teo_kar_tti"
                    webapi_req_data_naiyo = get_req_naiyo_IF024KH(pRow)
                    'Rev093 廃滅と撤去完了逆転の抑止 MOD End
                Case Else
                    Return False
            End Select

            sqlstr = String.Format(sb.ToString,
                                 EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(System.DateTime.Now.ToString("yyyyMMddHHmmss")) _
                               , EcOrgIO.EcOrgString.GetSqlText(myUser.USER_ID) _
                               , EcOrgIO.EcOrgString.GetSqlText(Kino_Cd) _
                               , EcOrgIO.EcOrgString.GetSqlText(keiki_id) _
                               , EcOrgIO.EcOrgString.GetSqlText(webapi_url_naiyo) _
                               , EcOrgIO.EcOrgString.GetSqlText(webapi_methd_naiyo) _
                               , EcOrgIO.EcOrgString.GetSqlText(webapi_req_data_naiyo)
                               )
            '*************************************************************************


            'SQLログ出力
            HdSysBasePg.Cmn.InsertLog(Me.myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L4, HdSysBase.Cmn.Kino_Lv.L3, Me.Kino_Cd, sqlstr)

            'SQL実行
            If Not Me.myDB.ExecDML(sqlstr) Then
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


    '定数マスタから計器設備管理システムへのURLを取得
    Friend Function getSETUBI_MNG_SYS_URL() As String
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim wUrl As String = ""

        '定数マスタを検索し計器設備管理のURLを取得する。
        If Not S006(dbr, [Const].C_ZYOUSU_CD_SETUBI_MNG_SYS_URL) Then
            Return ""
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                wUrl = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("STR1_ZYOUSU_VALUE"), ""))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return wUrl
    End Function

    ''' <summary>
    ''' 定数マスタ　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>定数マスタを検索します</remarks>
    Friend Function S006(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pZYOUSU_CD As String) As Boolean

        '*************************************************************************
        'C2.3.13_SQL詳細設計書のSQLID，SQL名称，概要に該当するメソッド記号名称，メソッド名称，remarksを記述して下さい。
        '※このブロックコメントは削除して下さい。
        '
        '【注意】dbr.Close()はクローズ漏れを防ぐためです。dbrが呼出元メソッドで複数回使用された場合は，削除しないで下さい。
        '*************************************************************************

        dbr.Close()

        Dim sb As New StringBuilder                     'StringBuilder
        Dim sqlstr As String = ""                       'SQL String

        If String.IsNullOrEmpty(pZYOUSU_CD) Then
            '定数コードなしの場合エラー
            Return False
        End If

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
  WHERE ZYOUSU_CD={0}
    AND YUKO_START_YMD<={1} 
    AND YUKO_END_YMD>={1}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                   EcOrgIO.EcOrgString.GetSqlText(pZYOUSU_CD),
                                   EcOrgIO.EcOrgString.GetSqlText(SyoriYMD))
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

    '取付・撤去　完了通知　リクエストデータ
    Private Function get_req_naiyo_IF023IF024(ByVal pRow As DataRow, ByVal pKeikiId As String) As String
        Dim sb As New StringBuilder
        Dim str As String = ""

        sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""kkyTtnTktBng"":""{2}""｝｝"
         )

        str = String.Format(sb.ToString, pKeikiId _
                                       , System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") _
                                       , CStr(Me.myDB.ConvDbNull(pRow("kyokyu_spot_tokti_no"), "")))
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function

    'Rev093 廃滅と撤去完了日の逆転抑止 ADD Start
    '撤去　完了通知　リクエストデータ
    Private Function get_req_naiyo_IF024(ByVal pRow As DataRow) As String
        Dim sb As New StringBuilder
        Dim str As String = ""
        'tkkk_smkey_enddateに保存してある竣工年月日時分秒をDateTime型で取得する
        Dim dateValue As DateTime = DateTime.ParseExact(CStr(Me.myDB.ConvDbNull(pRow("tkkk_smkey_enddate"), "")), "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture)

        sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""kkyTtnTktBng"":""{2}""｝｝"
         )

        str = String.Format(sb.ToString, CStr(Me.myDB.ConvDbNull(pRow("tkkk_keiki_id"), "")) _
                                       , dateValue.ToString("yyyy-MM-ddTHH:mm:ss") _
                                       , CStr(Me.myDB.ConvDbNull(pRow("kyokyu_spot_tokti_no"), "")))
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function
    'Rev093 廃滅と撤去完了日の逆転抑止 ADD End


    '取付・撤去（高圧他）　完了通知　リクエストデータ
    Private Function get_req_naiyo_IF023IF024KH(ByVal pRow As DataRow, ByVal pKeikiId As String) As String
        Dim sb As New StringBuilder
        Dim str As String = ""

        If CStr(Me.myDB.ConvDbNull(pRow("zyuky_ukky_kbn"), "")).Equals([Const].C_ZYUKY_UKKY_KBN_ZYUKY) Then
            '需給受給区分　需給（供給地点特定番号）
            sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""kkyTtnTktBng"":""{2}""｝｝"
         )
        Else
            '需給受給区分　受給（受電地点特定番号）
            sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""jdnTtnTktBng"":""{2}""｝｝"
         )
        End If

        str = String.Format(sb.ToString, pKeikiId _
                                       , System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") _
                                       , CStr(Me.myDB.ConvDbNull(pRow("kyokyu_spot_tokti_no"), "")))
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function

    'Rev093 廃滅と撤去完了日の逆転抑止(JSON作成を取付・撤去で共用していたが撤去を分離) ADD Start
    '撤去（高圧他）　完了通知　リクエストデータ
    Private Function get_req_naiyo_IF024KH(ByVal pRow As DataRow) As String
        Dim sb As New StringBuilder
        Dim str As String = ""
        'tkkk_smkey_enddateに保存してある竣工年月日時分秒をDateTime型で取得する
        Dim dateValue As DateTime = DateTime.ParseExact(CStr(Me.myDB.ConvDbNull(pRow("tkkk_smkey_enddate"), "")), "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture)

        If CStr(Me.myDB.ConvDbNull(pRow("zyuky_ukky_kbn"), "")).Equals([Const].C_ZYUKY_UKKY_KBN_ZYUKY) Then
            '需給受給区分　需給（供給地点特定番号）
            sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""kkyTtnTktBng"":""{2}""｝｝"
         )
        Else
            '需給受給区分　受給（受電地点特定番号）
            sb.Append(
"｛""kikId"":""{0}"",""sttKsnDt"":""{1}"",""sttJho"":｛""jdnTtnTktBng"":""{2}""｝｝"
         )
        End If

        str = String.Format(sb.ToString, CStr(Me.myDB.ConvDbNull(pRow("tkkk_keiki_id"), "")) _
                                       , dateValue.ToString("yyyy-MM-ddTHH:mm:ss") _
                                       , CStr(Me.myDB.ConvDbNull(pRow("kyokyu_spot_tokti_no"), "")))
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function
    'Rev093 廃滅と撤去完了日の逆転抑止 ADD End


    '設備更新通知（第１世代）　リクエストデータ
    Private Function get_req_naiyo_IF025G1(ByVal pRow As DataRow) As String
        Dim sb As New StringBuilder
        Dim str As String = ""

        sb.Append(
"｛""kikId"":""{0}"",""stbJhoKsnDt"":""{1}"",""ktaKbn"":""1"",""stbJhoTia"":｛""ksnKbnKikKhnJho"":""0"",""ksnKbnKikKoiJhoTusb"":""0"",""ksnKbnKikKoiJhoTanb"":""1"",""ksnKbnKikSeeJho"":""0"",""ksnKbnHnskJho"":""0"",""ksnKbnFzkSocJho"":""0"",""ksnKbnKntJho"":""0"",""ksnKbnSntZksJho"":""0"",""kikKoiJhoTia"":｛""tanb"":｛""szosCd"":""{2}"",""szoYm"":""{3}""｝｝｝｝"
        )

        str = String.Format(sb.ToString,
                             CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), "")) _
                           , System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_szsya_cd"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_seizo_yy"), "")) + "01")   '西暦4桁だが計器マスタの要望で「西暦4桁＋"01"」に変更
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function

    '設備更新通知（第２世代）　リクエストデータ
    Private Function get_req_naiyo_IF025G2(ByVal pRow As DataRow) As String
        Dim sb As New StringBuilder
        Dim str As String = ""

        sb.Append(
"｛""kikId"":""{0}"",""stbJhoKsnDt"":""{1}"",""ktaKbn"":""1"",""stbJhoTia"":｛""ksnKbnKikKhnJho"":""0"",""ksnKbnKikKoiJhoTusb"":""0"",""ksnKbnKikKoiJhoTanb"":""1"",""ksnKbnKikSeeJho"":""0"",""ksnKbnHnskJho"":""0"",""ksnKbnFzkSocJho"":""0"",""ksnKbnKntJho"":""0"",""ksnKbnSntZksJho"":""0"",""kikKoiJhoTia"":｛""tanb"":｛""szosCd"":""{2}"",""szoYm"":""{3}"",""sbtCd"":""{4}"",""ssnsDnaCd"":""{5}"",""yryCd"":""{6}"",""szoBng"":""{7}"",""ktsMei"":""{8}"",""kuzShtKbn"":""{9}"",""szosKnr"":""{10}""｝｝｝｝"
        )

        str = String.Format(sb.ToString,
                             CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), "")) _
                           , System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_szsya_cd"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_seizo_ym"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_sbt_cd"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_ssnsk_dnat_cd"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_yr"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_seizo_no"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_ktsk_ms"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_kzskbt_kbn"), "")) _
                           , CStr(Me.myDB.ConvDbNull(pRow("tkkk_tnsb_szsya_mng_value"), "")))
        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function

    'Rev056 2025/10/01 物理分割 計器マスタ対応 ADD Start
    '設備更新通知（高圧　変成器）　リクエストデータ
    Private Function get_req_naiyo_IF025KH(ByVal pRow As DataRow) As String
        Dim sb As New StringBuilder
        Dim str As String = ""
        Dim wstr As String = ""

        '項目名用変数
        Dim keiki_yr As String = ""              '容量(一次側電流)コード　必須
        Dim hnsik_ksyu_cd As String = ""         '器種コード　　　　　　　必須
        Dim hnsik_vctct_ktsk_cd As String = ""   'VCT・CT_型式コード　　　必須
        Dim hnsik_vctct_seizo_no As String = ""  'VCT・CT_製造番号　　　　必須
        Dim hnsik_vctct_seizo_yy As String = ""  'VCT・CT_製造年
        Dim hnsik_ct2_ktsk_cd As String = ""     'CT2_型式コード
        Dim hnsik_ct2_seizo_no As String = ""    'CT2_製造番号
        Dim hnsik_ct2_seizo_yy As String = ""    'CT2_製造年
        Dim hnsik_vt1_ktsk_cd As String = ""     'VT1_型式コード
        Dim hnsik_vt1_seizo_no As String = ""    'VT1_製造番号
        Dim hnsik_vt1_seizo_yy As String = ""    'VT1_製造年
        Dim hnsik_vt2_ktsk_cd As String = ""     'VT2_型式コード
        Dim hnsik_vt2_seizo_no As String = ""    'VT2_製造番号
        Dim hnsik_vt2_seizo_yy As String = ""    'VT2_製造年

        Dim zyrt As String = ""                  '計器_乗率
        Dim yuko_kigen_ym As String = ""         '計器_検定満了有効年月
        Dim hnsik_gknti_no As String = ""        '変成器_原検定番号
        Dim hnsik_yuko_kigen_ym As String = ""   '変成器_有効期限年月

        If CStr(Me.myDB.ConvDbNull(pRow("trke_tis_kbn"), "")) = MjK1.Cmn.Const.TrkeTisKbn.KeikiHnsik Then
            '取替対象区分が計器・変成器の場合
            '取付変成器の項目名を設定
            keiki_yr = "ttkk_keiki_yr"
            hnsik_ksyu_cd = "ttkk_hnsik_ksyu_cd"
            hnsik_vctct_ktsk_cd = "ttkk_hnsik_vctct_ktsk_cd"
            hnsik_vctct_seizo_no = "ttkk_hnsik_vctct_seizo_no"
            hnsik_vctct_seizo_yy = "ttkk_hnsik_vctct_seizo_yy"
            hnsik_ct2_ktsk_cd = "ttkk_hnsik_ct2_ktsk_cd"
            hnsik_ct2_seizo_no = "ttkk_hnsik_ct2_seizo_no"
            hnsik_ct2_seizo_yy = "ttkk_hnsik_ct2_seizo_yy"
            hnsik_vt1_ktsk_cd = "ttkk_hnsik_vt1_ktsk_cd"
            hnsik_vt1_seizo_no = "ttkk_hnsik_vt1_seizo_no"
            hnsik_vt1_seizo_yy = "ttkk_hnsik_vt1_seizo_yy"
            hnsik_vt2_ktsk_cd = "ttkk_hnsik_vt2_ktsk_cd"
            hnsik_vt2_seizo_no = "ttkk_hnsik_vt2_seizo_no"
            hnsik_vt2_seizo_yy = "ttkk_hnsik_vt2_seizo_yy"
            zyrt = "ttkk_zyrt"
            yuko_kigen_ym = "ttkk_yuko_kigen_ym"
            hnsik_gknti_no = "ttkk_hnsik_gknti_no"
            hnsik_yuko_kigen_ym = "ttkk_hnsik_yuko_kigen_ym"
        Else
            '取替対象区分が「計器」または「計器のみ（変成器共用）」の場合
            '撤去変成器の項目名を設定
            keiki_yr = "ttkk_keiki_yr"            '変成器情報ではないので取付計器の値を設定する 2026/1/21
            hnsik_ksyu_cd = "tkkk_hnsik_ksyu_cd"
            hnsik_vctct_ktsk_cd = "tkkk_hnsik_vctct_ktsk_cd"
            hnsik_vctct_seizo_no = "tkkk_hnsik_vctct_seizo_no"
            hnsik_vctct_seizo_yy = "tkkk_hnsik_vctct_seizo_yy"
            hnsik_ct2_ktsk_cd = "tkkk_hnsik_ct2_ktsk_cd"
            hnsik_ct2_seizo_no = "tkkk_hnsik_ct2_seizo_no"
            hnsik_ct2_seizo_yy = "tkkk_hnsik_ct2_seizo_yy"
            hnsik_vt1_ktsk_cd = "tkkk_hnsik_vt1_ktsk_cd"
            hnsik_vt1_seizo_no = "tkkk_hnsik_vt1_seizo_no"
            hnsik_vt1_seizo_yy = "tkkk_hnsik_vt1_seizo_yy"
            hnsik_vt2_ktsk_cd = "tkkk_hnsik_vt2_ktsk_cd"
            hnsik_vt2_seizo_no = "tkkk_hnsik_vt2_seizo_no"
            hnsik_vt2_seizo_yy = "tkkk_hnsik_vt2_seizo_yy"
            zyrt = "ttkk_zyrt"                    '変成器情報ではないので取付計器の値を設定する 2026/1/21
            yuko_kigen_ym = "ttkk_yuko_kigen_ym"  '変成器情報ではないので取付計器の値を設定する 2026/1/21
            hnsik_gknti_no = "tkkk_hnsik_gknti_no"
            hnsik_yuko_kigen_ym = "tkkk_hnsik_yuko_kigen_ym"
        End If

        '計器マスタ転送文字列を構築
        sb.Append(
"｛""kikId"":""{0}"",""stbJhoKsnDt"":""{1}"",""ktaKbn"":""2"",""stbJhoKat"":｛""ksnKbnKikKhnJho"":""0"",""ksnKbnKikKoiJhoTusb"":""0"",""ksnKbnKikSeeJho"":""0"",""ksnKbnHnskJho"":""1"",""ksnKbnFzkSocJho"":""0"",""ksnKbnKntJho"":""1"",""ksnKbnSntZksJho"":""0"",""hnskJhoKat"":｛"
        )
        str = String.Format(sb.ToString, CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), "")), System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"))

        '高圧　変成器
        '容量(一次側電流)コード　必須
        wstr = getStringFormat("""yryIjgDnrCd"":""{0}""", getLPadStr(CStr(Me.myDB.ConvDbNull(pRow(keiki_yr), "")), [Const].C_KEIKI_YORYO_KETA_IF025))
        str = str + wstr
        '器種コード　必須
        wstr = getStringFormat("""ksyCd"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_ksyu_cd), "")))
        str = str + "," + wstr
        'VCT・CT_型式コード　必須
        wstr = getStringFormat("""vctCtKtsCd"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vctct_ktsk_cd), "")))
        str = str + "," + wstr
        'VCT・CT_製造番号　必須
        wstr = getStringFormat("""vctCtSzoBng"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vctct_seizo_no), "")))
        str = str + "," + wstr
        'VCT・CT_製造年 必須
        wstr = getStringFormat("""vctCtSzoY"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vctct_seizo_yy), "")))
        str = str + "," + wstr
        'CT2_型式コード
        wstr = getStringFormat("""ct2KtsCd"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_ct2_ktsk_cd), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'CT2_製造番号
        wstr = getStringFormat("""ct2SzoBng"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_ct2_seizo_no), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'CT2_製造年
        wstr = getStringFormat("""ct2SzoY"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_ct2_seizo_yy), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT1_型式コード
        wstr = getStringFormat("""vt1KtsCd"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt1_ktsk_cd), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT1_製造番号
        wstr = getStringFormat("""vt1SzoBng"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt1_seizo_no), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT1_製造年
        wstr = getStringFormat("""vt1SzoY"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt1_seizo_yy), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT2_型式コード
        wstr = getStringFormat("""vt2KtsCd"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt2_ktsk_cd), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT2_製造番号
        wstr = getStringFormat("""vt2SzoBng"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt2_seizo_no), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        'VT2_製造年
        wstr = getStringFormat("""vt2SzoY"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_vt2_seizo_yy), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        str = str + "｝,"

        '高圧　検定情報
        str = str & """kntJhoKat"":｛"
        '計器_検定番号　必須
        wstr = getStringFormat("""kikKntBng"":""{0}""", getKNTJHOKAT_KIKKNTBNG(CStr(Me.myDB.ConvDbNull(pRow("ttkk_keiki_id"), ""))))  'MAMから取得2026/1/22←固定値を設定
        str = str + wstr
        '計器_乗率　必須
        wstr = getStringFormat("""kikJrt"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(zyrt), "")))
        str = str + "," + wstr
        '計器_検定満了有効年月　必須
        wstr = getStringFormat("""kikKntMnrYkoYm"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(yuko_kigen_ym), "")))
        str = str + "," + wstr
        '計器_原検定番号(合番号)
        wstr = getStringFormat("""kikGktBngAbg"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_gknti_no), "")))         '変成器_原検定番号を設定
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        '変成器_原検定番号
        wstr = getStringFormat("""hnskGktBng"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_gknti_no), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        '変成器_原検定年月
        wstr = getStringFormat("""hnskGktYm"":""{0}""", getHnskGktYm(pRow, hnsik_gknti_no, hnsik_yuko_kigen_ym))       '変成器_有効期限年月から算出
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        '変成器_有効期限年月
        wstr = getStringFormat("""hnskYkoKgnYm"":""{0}""", CStr(Me.myDB.ConvDbNull(pRow(hnsik_yuko_kigen_ym), "")))
        If Not String.IsNullOrEmpty(wstr) Then str = str + "," + wstr
        str = str + "｝｝｝"

        Return str.Replace("｛", "{").Replace("｝", "}")
    End Function

    ''' <summary>
    ''' 固定長文字列を取得する。
    ''' </summary>
    ''' <param name="s">元文字列</param>
    ''' <param name="keta">桁数</param>
    ''' <returns>固定長文字列('0'左埋め)</returns>
    ''' <remarks></remarks>
    Private Function getLPadStr(ByVal s As String, ByVal keta As Integer) As String
        Dim wkstr = New String("0"c, keta)
        Return Right(wkstr + s.Trim, keta)
    End Function

    ''' <summary>
    ''' 文字列を編集する。
    ''' </summary>
    ''' <param name="f">Format</param>
    ''' <param name="v">value</param>
    ''' <returns>文字列を編集する。値なしは空文字を返却</returns>
    ''' <remarks></remarks>
    Private Function getStringFormat(ByVal f As String, ByVal v As String) As String
        If String.IsNullOrEmpty(v.Trim) Then
            Return ""
        Else
            Return String.Format(f, v)
        End If
    End Function

    '変成器_原検定年月を設定
    Private Function getHnskGktYm(ByVal pRow As DataRow, ByVal clm_hnsik_gknti_no As String, ByVal clm_hnsik_yuko_kigen_ym As String) As String
        Dim wkstr As String = ""
        Dim s As String = ""
        Dim yy As Integer = 0
        Dim hnsik_yuko_kigen_ym As String = ""

        '変成器＿原検定番号の頭が「4」、「5」の場合、変成器＿有効期限年月　-14年
        '変成器＿原検定番号の頭が「7」、「8」の場合、変成器＿有効期限年月　-21年

        '変成器＿原検定番号を取得し、先頭１文字をチェック
        wkstr = CStr(Me.myDB.ConvDbNull(pRow(clm_hnsik_gknti_no), ""))
        If wkstr.Length < 1 Then
            Return ""
        End If
        '先頭１文字取得
        s = wkstr.Substring(0, 1)
        If s.Equals([Const].C_KASHO_CD_4) Or s.Equals([Const].C_KASHO_CD_5) Then
            yy = getHNSIK_YUKO_YEAR()       '14年を取得
        ElseIf s.Equals([Const].C_KASHO_CD_7) Or s.Equals([Const].C_KASHO_CD_8) Then
            yy = getHNSIK_YUKO_YEAR_TOKTI() '21年を取得
        Else
            Return ""
        End If

        '変成器＿有効期限年月から変成器_原検定年月を算出
        hnsik_yuko_kigen_ym = CStr(Me.myDB.ConvDbNull(pRow(clm_hnsik_yuko_kigen_ym), ""))
        If hnsik_yuko_kigen_ym.Length < "yyyymm".Length Then
            Return ""
        End If
        wkstr = CStr(CInt(hnsik_yuko_kigen_ym.Substring(0, 4).ToString) - yy) + hnsik_yuko_kigen_ym.Substring(4, 2)

        Return wkstr
    End Function

    '定数マスタから高圧変成器有効年を取得
    Friend Function getHNSIK_YUKO_YEAR() As Integer
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim wk As Integer = 0

        '定数マスタを検索し計器設備管理のURLを取得する。
        If Not S006(dbr, [Const].C_ZYOUSUCD_HNSIK_YUKO_YEAR_) Then
            Return 0
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                wk = CInt(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_zyousu_value"), 0))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return wk
    End Function

    '定数マスタから高圧変成器有効年(特定検定)を取得
    Friend Function getHNSIK_YUKO_YEAR_TOKTI() As Integer
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim wk As Integer = 0

        '定数マスタを検索し計器設備管理のURLを取得する。
        If Not S006(dbr, [Const].C_ZYOUSUCD_HNSIK_YUKO_YEAR_TOKTI_) Then
            Return 0
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                wk = CInt(Me.myDB.ConvDbNull(dbr.DataStruct.Item("value1_zyousu_value"), 0))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return wk
    End Function

    '定数マスタから検定情報(高圧) 計器_検定番号を取得
    Friend Function getKNTJHOKAT_KIKKNTBNG() As String
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim wkstr As String = ""

        '定数マスタを検索する。
        If Not S006(dbr, [Const].C_ZYOUSUCD_KNTJHOKAT_KIKKNTBNG) Then
            Return ""
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                wkstr = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("STR1_ZYOUSU_VALUE"), ""))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return wkstr
    End Function


    'MAMから検定情報(高圧) 計器_検定番号を取得 2026/1/22
    Friend Function getKNTJHOKAT_KIKKNTBNG(ByVal pKeikiId As String) As String
        Dim dbr As New HdPostgre.HdPostgreReader  'DB Reader
        Dim wkstr As String = ""

        'MAMから計器_検定番号を検索する。
        If Not S009(dbr, pKeikiId) Then
            Return wkstr
        End If

        If dbr.DataStruct.HasRows Then
            While (dbr.DataStruct.Read)
                wkstr = CStr(Me.myDB.ConvDbNull(dbr.DataStruct.Item("kik_knt_bng"), ""))
                '１件のみ処理
                Exit While
            End While
        End If
        dbr.Close()

        Return wkstr
    End Function

    ''' <summary>
    ''' MAM 検定情報（公開）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>MAM 検定情報（公開）を検索します</remarks>
    Friend Function S009(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pKeikiId As String) As Boolean

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
            sb.Append(<sql><![CDATA[
select
  kik_knt_bng
  FROM msti_p_knt_kki
  WHERE kik_id={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                   EcOrgIO.EcOrgString.GetSqlText(pKeikiId))
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
    'Rev056 2025/10/01 物理分割 計器マスタ対応 ADD End

#End Region
    'Rev040 2024/12/26 物理分割 計器マスタ対応 ADD End


    'Rev107 取付計器が第２世代の場合に統合QR読取時の指示数取得不具合への対応でMAMデータと突合するためSQL追加 ADD Start
    ''' <summary>
    ''' MAM 計器設定（公開）　検索
    ''' </summary>
    ''' <param name="dbr">DB Reader</param>
    ''' <remarks>MAM 計器設定（公開）を検索します</remarks>
    Friend Function S010(ByRef dbr As HdPostgre.HdPostgreReader, ByVal pKeikiId As String) As Boolean

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
            sb.Append(<sql><![CDATA[
select
  sisn_1,sisn_2
  FROM msti_p_kik_see_kki
  WHERE kik_id={0}
        ]]></sql>.Value)

            sqlstr = String.Format(sb.ToString,
                                   EcOrgIO.EcOrgString.GetSqlText(pKeikiId))
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
    'Rev107 取付計器が第２世代の場合に統合QR読取時の指示数取得不具合への対応でMAMデータと突合するためSQL追加 ADD End


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
