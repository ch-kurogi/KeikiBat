''' <summary>
''' 共通定数クラス
''' </summary>
Public Class [Const]

    ' 処理間隔(ミリ秒)
    Public Const C_PROC_MAX_FILES As Integer = 1              '１回の最大処理ファイル数  2024/03/18 1←3
    Public Const C_SLEEP_UPLD_TRKM_INTERVAL As Integer = 1000

    ' バージョンファイル名
    Public Const C_APPVER_DESTPATH As String = "M00"
    Public Const C_APPVER_SUBSYS As String = "KEIKI"
    Public Const C_APPVER_FILE As String = "AppVer"

    'テーブル項目関連
    Public Const C_KORYO_CD_KETA As Integer = 7           '工量コード桁数
    Public Const C_TUIKAKOHI_KOZI_KBN_KETA As Integer = 4 '追加工費工事区分桁数
    Public Const C_WRMS_KORYO_CD As String = "9900040"    '離島
    Public Const C_WRMS_NONE As String = "1.00"
    Public Const C_CLM_TKKK_KEIKI_ID As String = "tkkk_keiki_id" '撤去計器 計器ID
    Public Const C_CLM_TTKK_KEIKI_ID As String = "ttkk_keiki_id" '取付計器 計器ID

    '取替_取替票行程　実費確認フラグ(ZIPPIKKNN_FLG)
    Public Const C_ZIPPIKKNN_FLG_OFF As String = "0" '対象外
    Public Const C_ZIPPIKKNN_FLG_NOT_CONFIRM As String = "1" '対象(未確認)
    Public Const C_ZIPPIKKNN_FLG_CONFIRM As String = "2"     '対象(確認済)

    'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD Start
    Public Const C_YUKO_SHOSU_KETA As Integer = 1            '有効小数桁１桁
    'Rev052 20241010 工費計算の小数第２位切り捨て対応 ADD End

    'Rev076 庫入情報送信エラー対応 ADD Start
    Public Const C_ZYOUSUCD_HAIMETU_DUP_VALUE As String = "HAIMETU_DUP_VALUE"
    Public Const C_SZSU_KETA As Integer = 7                  '計器業務管理　指示数のデータ桁数
    'Rev076 庫入情報送信エラー対応 ADD End

    'Rev040 2024/12/26 物理分割　計器マスタ連係対応 ADD Start
    Public Const C_TNSB_SBT_CD_2G As String = "B5"            '端子部種別 B5:第２世代
    Public Const C_TNSB_TORIKAE As String = "1"               '端子部再用　全取替
    Public Const C_TNSB_REUSE As String = "2"                 '端子部再用　再用
    Public Const C_KEIKI_MASTER_RENKEI_IF023 As Integer = 1   'IF023_取付完了通知
    Public Const C_KEIKI_MASTER_RENKEI_IF024 As Integer = 2   'IF024_撤去完了通知
    Public Const C_KEIKI_MASTER_RENKEI_IF025G1 As Integer = 3 'IF025_設備情報更新（端子部再用の場合）第１世代
    Public Const C_KEIKI_MASTER_RENKEI_IF025G2 As Integer = 4 'IF025_設備情報更新（端子部再用の場合）第２世代
    Public Const C_ZYOUSU_CD_SETUBI_MNG_SYS_URL = "SETUBI_MNG_SYS_URL" '計器設備管理URL取得　定数コード
    'Rev040 2024/12/26 物理分割　計器マスタ連係対応 ADD End
    'Rev056 2025/10/01 物理分割　計器マスタ連係対応 ADD Start
    Public Const C_KEIKI_MASTER_RENKEI_IF025KH As Integer = 5 'IF025_設備情報更新（高圧　変成器）
    Public Const C_KEIKI_YORYO_KETA_IF025 As Integer = 4      'IF025連係の計器容量桁
    Public Const C_KEIKI_YORYO_KETA As Integer = 3            '計器業務管理システム　計器容量桁
    Public Const C_KEIKI_YORYO_KETA_QR As Integer = 4         '統合QR　計器容量桁
    Public Const C_KEIKI_MASTER_RENKEI_IF023KH As Integer = 6 'IF023_取付完了通知（高圧他）　需給受給判定対応
    Public Const C_KEIKI_MASTER_RENKEI_IF024KH As Integer = 7 'IF024_撤去完了通知（高圧他）　需給受給判定対応
    Public Const C_ZYUKY_UKKY_KBN_ZYUKY As String = "1"       '需給受給区分　需給
    Public Const C_ZYUKY_UKKY_KBN_UKKY As String = "2"        '需給受給区分　受給
    Public Const C_KEIKI_KBN_CD_TIAT As String = "1"          '計器区分コード　低圧
    Public Const C_KEIKI_KBN_CD_KOAT As String = "2"          '計器区分コード　高圧
    Public Const C_KEIKI_KBN_CD_TKOAT As String = "3"         '計器区分コード　特高
    Public Const C_ZYOUSUCD_HNSIK_YUKO_YEAR_ = "HNSIK_YUKO_YEAR"              '高圧変成器有効年　定数コード
    Public Const C_ZYOUSUCD_HNSIK_YUKO_YEAR_TOKTI_ = "HNSIK_YUKO_YEAR_TOKTI"  '高圧変成器有効年(特定検定)　定数コード
    Public Const C_ZYOUSUCD_KNTJHOKAT_KIKKNTBNG = "KNTJHOKAT_KIKKNTBNG_VALUE" '検定情報(高圧) 計器_検定番号　定数コード
    Public Const C_KASHO_CD_4 As String = "4"                 '高圧変成器　原検定番号の先頭文字
    Public Const C_KASHO_CD_5 As String = "5"                 '高圧変成器　原検定番号の先頭文字
    Public Const C_KASHO_CD_7 As String = "7"                 '高圧変成器　原検定番号の先頭文字
    Public Const C_KASHO_CD_8 As String = "8"                 '高圧変成器　原検定番号の先頭文字
    'Rev056 2025/10/01 物理分割　計器マスタ連係対応 ADD End

    'Rev107 取付計器が第２世代の場合に統合QR読取時の指示数取得不具合への対応でMAMデータと突合するため定義追加 ADD Start
    Public Const C_DIGSU_4 As String = "4"                    '計器桁数　4
    Public Const C_DIGSU_5 As String = "5"                    '計器桁数　5
    'Rev107 取付計器が第２世代の場合に統合QR読取時の指示数取得不具合への対応でMAMデータと突合するため定義追加 ADD End

#Region "インテリセンス非表示"
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Shadows Function Equals() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Shadows Function GetHashCode() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Shadows Function [GetType]() As Type
        Return GetType(String)
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Shadows Function ReferenceEquals() As Object
        Return New Object
    End Function

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Shadows Function ToString() As Object
        Return New Object
    End Function
#End Region

End Class
