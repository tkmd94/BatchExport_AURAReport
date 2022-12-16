#指定した期間の"DailyExposureReport"を出力するスクリプト

# 開始日付の指定（★★ここで日付を指定する）　
$startDateText="20221212"
# 終了日付の指定（★★ここで日付を指定する）
$endDateText  ="20221212"

# 日付型の開始日付の変数を定義
$startDate = [DateTime]::ParseExact($startDateText,"yyyyMMdd",$null)
# 日付型の終了日付の変数を定義
$endDate = [DateTime]::ParseExact($endDateText,"yyyyMMdd",$null)

# 期間計算
$dt=($endDate - $startDate).Days

#レポート名の定義
$ReportName="DailyExposureReport"
# レポートURLの定義
$reportURL="https://ariareport:56300/Reportserver/Pages/Report.aspx?%2f"+$ReportName
# 出力フォルダの定義
$exportPath="C:\Temp\"
# 出力ファイル名の定義
$newFileBaseName="DailyExposureReport"

for ($i =0;$i -le $dt ; $i++){
    # 日付計算
    $date=$startDate.AddDays($i)
    # 曜日取得　0:日～6:土
    $weekDayValue=$date.DayOfWeek.value__
    # 平日のみの処理に限定
    if(($weekDayValue -gt 0) -And ($weekDayValue -lt 6)){
       # 出力ファイルに付ける日付文字列を定義
        $dateText=$date.ToString("yyyy年MM月dd日")
       # 出力ファイル名を定義
        $newFilePath=$exportPath+$newFileBaseName+"_"+$dateText+".pdf"
       # ファイル名をログウィンドウに出力
        Write-Host($newFilePath)
       # SSRSのURL呼び出し文を定義
        $reportPath= $reportURL+"&Date="+$date.ToString("yyyy/MM/dd")+"&rs:Format=PDF"
       # PDF形式でレポートを所定パスに保存
        Invoke-WebRequest -Uri $reportPath -OutFile $newFilePath -UseDefaultCredentials
    }
}
