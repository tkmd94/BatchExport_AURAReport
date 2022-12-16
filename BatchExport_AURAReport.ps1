#指定した期間の"照射録_位置照合（日報）"を出力するスクリプト

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

# レポートURLの定義
$reportURL="https://ariareport:56300/Reportserver/Pages/Report.aspx?%2f%e7%85%a7%e5%b0%84%e9%8c%b2_%e4%bd%8d%e7%bd%ae%e7%85%a7%e5%90%88%ef%bc%88%e6%97%a5%e5%a0%b1%ef%bc%89"
# 出力フォルダの定義
$exportPath="\\ariasvr\va_transfer\mlc\--- 照射録 ---\照射録_位置照合（日報）\"
# 出力ファイル名の定義
$newFileBaseName="照射録_位置照合（日報）"

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
