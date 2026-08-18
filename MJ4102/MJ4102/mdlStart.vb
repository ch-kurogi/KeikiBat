Module mdlStart

    ' **************************************************************************
    ' 機能名称：mdlStart
    ' 新規作成：(C)Energia Communications Y.Ohnishi.
    '
    ' 機能概要：起動モジュール
    ' 使用方法：
    ' 前提条件：
    ' 更新履歴：2022.10.05 - EC Y.Ohnishi 新規作成
    ' Rev044    2025.07.17 - EC Y.Ohnishi 外字のSJIS変換エラー対応
    ' **************************************************************************

    '-------------------------------------------------------------------------
    ' （変数）グローバル
    '-------------------------------------------------------------------------

    'グローバルクラス変数
    Public myUser As New HdSysBasePg.Cmn.UserInfo     'ユーザ情報
    Public myPath As New HdSysBasePg.Cmn.PathInfo     'パス情報
    Public myDB As New HdPostgre.HdPostgreDb          'ＤＢ接続
    Public myPID As String = ""                       'プロセスID
    Public myEncoding As String = "SJIS"              'PostgreSQL接続への指示
    'Rev044 20250717 外字のSJIS変換エラー対応 ADD Start
    Public dbEncoding As String = "UTF8"              'PostgreSQL接続への指示
    'Rev044 20250717 外字のSJIS変換エラー対応 ADD End

    Public C_MAX_RETRY As Integer = 10
    Public C_INTERVAL As Integer = 1000 'ms

    '-------------------------------------------------------------------------
    ' （共通処理）プログラム開始
    '-------------------------------------------------------------------------
    Sub Main()

        Const wKINO_CD As String = "MJ4102"         '機能コード（アップロード　完了監視）
        Dim wKINO_MS As String = ""                 '機能名称
        Dim wSYORI_YMD As String = ""               '処理年月日
        Dim wZGSYO_CD_List As New List(Of String)   '対象の事業所(先頭文字・指定なしは全て対象）
        Dim wLOG_NAIYO As String = ""               '機能利用可否チェック結果
        Dim cnt As Integer = 0
        Dim retval As Boolean = False

        Dim wExit_Code As HdSysBase.Cmn.ExitCode = HdSysBase.Cmn.ExitCode.Normal    '終了コード

        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
        myPID = p.Id.ToString().PadLeft(10, "0"c)

        Try
            'コマンドライン引数取得
            If Not mdlStart.GetCommandLineArgs(My.Application.CommandLineArgs, myUser.INSTANCE_CD, wZGSYO_CD_List, wSYORI_YMD) Then
                Call mdlStart.WriteApplicationLog("コマンドライン引数取得処理に失敗しました。")
                wExit_Code = HdSysBase.Cmn.ExitCode.Abnormal
                Exit Sub
            End If

            'DB接続
            'If Not mdlStart.ConnectDB(myDB, myUser, myPath) Then
            '    Call mdlStart.WriteApplicationLog("DB接続処理に失敗しました。")
            '    wExit_Code = HdSysBase.Cmn.ExitCode.Abnormal
            '    Exit Sub
            'End If
            cnt = 0
            While cnt < C_MAX_RETRY
                retval = mdlStart.ConnectDB(myDB, myUser, myPath)
                If retval Then
                    Exit While
                Else
                    System.Threading.Thread.Sleep(C_INTERVAL)
                End If
                cnt += 1
            End While
            If Not retval Then
                Call mdlStart.WriteApplicationLog("DB接続処理に失敗しました。")
                wExit_Code = HdSysBase.Cmn.ExitCode.Abnormal
                Exit Sub
            End If

            '機能名称を取得
            wKINO_MS = HdSysBasePg.Cmn.GetKino_Ms(myDB, HdSysBase.Cmn.Kino_Lv.L3, wKINO_CD)

            '機能利用可否チェック
            If Not HdSysBasePg.Cmn.CanUseKino(myDB, HdSysBase.Cmn.Kino_Lv.L3, wKINO_CD, wLOG_NAIYO) Then
                Call mdlStart.WriteApplicationLog("機能利用不可のため起動できませんでした。")
                wExit_Code = HdSysBase.Cmn.ExitCode.Warning
                Exit Sub
            End If


            '起動ログ作成
            Call HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, wKINO_CD, String.Format("{0}を起動しました。PID={1}", wKINO_MS, myPID))

            Try
                '処理年月日チェック
                If wSYORI_YMD.Trim = "" OrElse Not HdSysBase.Cmn.IsStringDate(wSYORI_YMD) Then
                    '業務処理日付取得
                    wSYORI_YMD = HdSysBasePg.Cmn.GetSystemYmd(myDB)
                End If

                '（内部処理）バッチメイン処理
                Dim cls As New MJ4102
                cls.myDB = myDB
                cls.myUser = myUser
                cls.myPath = myPath
                cls.KinoCD = wKINO_CD
                cls.SyoriYMD = wSYORI_YMD
                cls.myPID = myPID
                For i = 0 To wZGSYO_CD_List.Count - 1
                    cls.ZGSYO_CD_List.Add(wZGSYO_CD_List(i))
                Next

                wExit_Code = cls.MainLogic()

            Catch ex As Exception
                'エラーログ作成
                wExit_Code = HdSysBase.Cmn.ExitCode.Abnormal
                Exit Sub
            Finally
                '終了ログ作成
                Call HdSysBasePg.Cmn.InsertLog(myUser, HdSysBase.Cmn.Log_Syusyu_Lv.L1, HdSysBase.Cmn.Kino_Lv.L3, wKINO_CD, String.Format("{0}を終了しました。PID={1}", wKINO_MS, myPID))
            End Try

        Catch ex As Exception
            Call mdlStart.WriteApplicationLog(ex.Message)
            wExit_Code = HdSysBase.Cmn.ExitCode.Abnormal
            Exit Sub
        Finally

            'DB切断
            myDB.DisConnect()

            '終了コードをセット
            Call HdSysBase.Cmn.Exit(wExit_Code)
        End Try

    End Sub

#Region " 内部処理 "

    '-------------------------------------------------------------------------
    ' （共通処理）アプリケーションログ出力
    '-------------------------------------------------------------------------
    Public Sub WriteApplicationLog(ByVal pMsg As String)

        Dim FilePath As String = "" 'ログ出力先
        Dim WriteLine As String = ""

        'ログ出力先の設定
        FilePath = System.Windows.Forms.Application.StartupPath & "\" & HdSysBase.Cmn.ConvStringByteLength(System.Windows.Forms.Application.ProductName, 20) & ".log"

        'ログ編集
        WriteLine = ""
        WriteLine += Now.ToString("yyyy/MM/dd HH:mm:ss ")
        WriteLine += pMsg.Replace(vbCr, Space(1)).Replace(vbCrLf, Space(1))

        'ログ出力
        Call EcOrgIO.EcOrgFileIO.WriteFile(FilePath, True, WriteLine)

    End Sub

    '-------------------------------------------------------------------------
    ' （内部処理）コマンドライン引数取得
    '-------------------------------------------------------------------------
    Private Function GetCommandLineArgs(ByVal pCmd As System.Collections.ObjectModel.ReadOnlyCollection(Of String),
                                        ByRef pSTZKS_CD As String,
                                        ByRef pZGSYO_CD_List As List(Of String),
                                        ByRef pSYORI_YMD As String) As Boolean
        Dim i As Integer = 0
        Dim wks As String = ""

        Try
            'コマンドライン引数チェック
            Select Case pCmd.Count
                Case 1      '接続先コード
                    pSTZKS_CD = pCmd.Item(0).ToString()

                Case 2      '接続先コード，対象事業所CD
                    pSTZKS_CD = pCmd.Item(0).ToString()
                    wks = pCmd.Item(1).ToString()
                    For i = 0 To wks.Length - 1
                        pZGSYO_CD_List.Add(wks.Substring(i, 1))
                    Next
                Case 3      '接続先コード，対象事業所CD，処理年月日
                    pSTZKS_CD = pCmd.Item(0).ToString()
                    wks = pCmd.Item(1).ToString()
                    For i = 0 To wks.Length - 1
                        pZGSYO_CD_List.Add(wks.Substring(i, 1))
                    Next
                    pSYORI_YMD = pCmd.Item(2).ToString()

                Case Else   'その他
                    Return False

            End Select

            Return True

        Catch ex As Exception
            Call mdlStart.WriteApplicationLog(ex.Message)
            Return False
        End Try

    End Function

    '-------------------------------------------------------------------------
    ' （内部処理）DB接続
    '-------------------------------------------------------------------------
    Private Function ConnectDB(ByRef pDB As HdPostgre.HdPostgreDb,
                               ByRef pUser As HdSysBasePg.Cmn.UserInfo,
                               ByRef pPath As HdSysBasePg.Cmn.PathInfo) As Boolean

        Dim RepDB As New HdPostgre.HdPostgreDb
        Dim retval As Boolean = False

        Try
            '計器定期取替システムでは、リポジトリDBは使用しない。
            ''リポジトリDB接続
            'If Not HdSysBasePg.Cmn.ConnectRepDB(RepDB, pUser.INSTANCE_CD) Then
            '    Call mdlStart.WriteApplicationLog("リポジトリDB接続処理に失敗しました。")
            '    Return False
            'End If

            '接続先情報取得
            pUser.USER_ID = "SYSTEM"

            'パス情報セット
            Call HdSysBasePg.Cmn.GetProfile(RepDB, pUser, pPath)

            '接続先情報取得
            If Not HdSysBasePg.Cmn.GetInstance(RepDB, pUser.INSTANCE_CD, pUser.DB_ACCOUNT_NAME, pUser.DB_PASSWORD, pUser.DB_SERVER, pUser.DB_DATABASE, pUser.DB_PORT) Then
                Call mdlStart.WriteApplicationLog("接続先情報取得処理に失敗しました。")
                Return False
            End If

            'DBに接続
            'Rev044 20250717 外字のSJIS変換エラー対応 MOD Start
            ''If Not pDB.Connect(pUser.DB_ACCOUNT_NAME, pUser.DB_PASSWORD, pUser.DB_SERVER, pUser.DB_DATABASE, pUser.DB_PORT) Then
            ''    Call mdlStart.WriteApplicationLog(pDB.ErrDescription)
            ''    Return False
            ''End If
            ''共通部品内で「set client_encoding=sjis」を行っている。PostgreSQLバージョン依存でEncoding=SJIS指定が必要。
            'If Not pDB.Connect(pUser.DB_ACCOUNT_NAME, pUser.DB_PASSWORD, pUser.DB_SERVER, pUser.DB_DATABASE, pUser.DB_PORT,,,,,,, myEncoding) Then
            '    Call mdlStart.WriteApplicationLog(pDB.ErrDescription)
            '    Return False
            'End If

            'Rev044 UTF8→SJIS変換エラーとなる文字への対応のためUTF8接続とする
            '20240130 SQLタイムアウト値C_CMD_TIMEOUT_SEC 設定
            If Not pDB.Connect(pUser.DB_ACCOUNT_NAME, pUser.DB_PASSWORD, pUser.DB_SERVER, pUser.DB_DATABASE, pUser.DB_PORT,,,,,,, dbEncoding) Then
                Call mdlStart.WriteApplicationLog(pDB.ErrDescription)
                Return False
            End If
            'Rev044 共通部品内で「set client_encoding=sjis」を行っているので、「set client_encoding=UTF8」を別途再設定する
            Dim cmdstr As String = "set client_encoding = " & dbEncoding & " ;"
            Dim recCnt As Double = 0
            retval = pDB.ExecDML(cmdstr, recCnt)
            'Rev044 20250717 外字のSJIS変換エラー対応 MOD End

            'ユーザ情報セット
            retval = HdSysBasePg.Cmn.GetUserInfo(pDB, pUser.USER_ID, pUser, HdSysBasePg.Cmn.GetSyoriYmd(pDB))
            If retval = False Then
                Call mdlStart.WriteApplicationLog("error GetUserInfo() ErrCode:" & pDB.ErrCode & " " & pDB.ErrDescription)
            End If

            Return retval

        Catch ex As Exception
            Call mdlStart.WriteApplicationLog(ex.Message)
            Return False
        Finally
            '計器定期取替システムでは、リポジトリDBは使用しない。
            ''リポジトリDB切断
            'RepDB.DisConnect()
        End Try

    End Function
#End Region

End Module
