''' <summary>
''' 共通定数クラス
''' </summary>
Public Class [Const]

    '持出処理の項目数
    Public Const C_COL_COUNT As Integer = 1

    '再処理回数
    Public Const C_PROC_RETRY_COUNT As Integer = 120

    '再処理間隔(秒)
    Public Const C_SLEEP_INTERVAL As Integer = 1

    '持出再処理間隔(ミリ秒)
    Public Const C_SLEEP_MTDS_INTERVAL As Integer = 5000  '想定1000ms　とりあえず5秒とする


    '取替対象　名称（検定種別から判定　1:特別検定→計器　2:提出検定→計器・変成器
    Public Const C_TRKE_TIS_KEIKI As String = "計器"
    Public Const C_TRKE_TIS_KEIKI_HENSEIKI As String = "計器・変成器"

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
