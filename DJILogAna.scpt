-- Analizer for DJI Flight Log data
-- ファイル選択ダイアログ
set csvFile to choose file with prompt "CSVファイルを選択してください" of type {"csv", "text"}

-- ファイル内容の読み込み
set csvText to read csvFile as «class utf8»
set allLines to paragraphs of csvText

-- 空行を除去して有効な行だけを抽出
set cleanedLines to {}
repeat with aLine in allLines
	if length of aLine > 0 then
		copy aLine to end of cleanedLines
	end if
end repeat

-- 仕様確認：
-- item 1: "sep=,"
-- item 2: ヘッダー行
-- item 3: 最初のデータ行（離陸時刻用）
-- item (last): 最後のデータ行（その他の集計データ用）

set headerLine to item 2 of cleanedLines
set firstDataLine to item 3 of cleanedLines -- データ開始行
set lastDataLine to item (count of cleanedLines) of cleanedLines -- 最終データ行

-- CSVの各項目をリストに分解
set headers to splitCSV(headerLine)
set firstDataValues to splitCSV(firstDataLine)
set lastDataValues to splitCSV(lastDataLine)

-- 各項目のインデックス（列番号）を取得
set idxDate to getColumnIndex(headers, "CUSTOM.date [local]")
set idxTime to getColumnIndex(headers, "CUSTOM.updateTime [local]")
set idxFlyTime to getColumnIndex(headers, "OSD.flyTime")
set idxHeight to getColumnIndex(headers, "OSD.heightMax [ft]")
set idxSpeed to getColumnIndex(headers, "OSD.hSpeedMax [MPH]")
set idxDroneName to getColumnIndex(headers, "RECOVER.aircraftName")

-- データ抽出と変換
-- 1. ドローン名 (最終行から取得)
set droneName to item idxDroneName of lastDataValues

-- 2. 離陸時刻 (データ開始行 = 3行目から取得)
set rawStartDate to item idxDate of firstDataValues -- m/d/y
set rawStartTime to item idxTime of firstDataValues -- h:m:s.s AM/PM
set gmtStartDate to parseDateTime(rawStartDate, rawStartTime)
set jstStartDate to gmtStartDate + (9 * hours) -- 離陸時刻（JST）

-- 3. 飛行時間 (最終行から取得) と 着陸時刻の計算
set flyTimeStr to item idxFlyTime of lastDataValues -- 例: "12m 34.5s"
set totalSeconds to parseFlyTimeToSeconds(flyTimeStr)
-- 離陸時刻に総飛行時間を加算
set jstLandingDate to jstStartDate + totalSeconds -- 着陸時刻（JST）

-- 4. 最大高度 (ft -> m)
set heightFt to item idxHeight of lastDataValues as number
set heightM to (heightFt * 0.3048)
set heightM to (round (heightM * 100)) / 100

-- 5. 最大速度 (MPH -> Km/H)
set speedMPH to item idxSpeed of lastDataValues as number
set speedKMH to (speedMPH * 1.60934)
set speedKMH to (round (speedKMH * 100)) / 100

-- 結果のテキスト作成
set msg to "ドローン名: " & droneName & return & ¬
	"離陸時刻 (JST): " & (jstStartDate as string) & return & ¬
	"着陸時刻 (JST): " & (jstLandingDate as string) & return & ¬
	"総飛行時間: " & flyTimeStr & return & ¬
	"最大高度: " & heightM & " m" & return & ¬
	"最大速度: " & speedKMH & " Km/H"

-- 結果の表示
display dialog msg buttons {"OK"} default button 1 with title "データ抽出完了"

-- OKボタンが押された後、内容をクリップボードにコピー
set the clipboard to msg

---------------------------------------------------------
-- サブルーチン（ハンドラ）
---------------------------------------------------------

on splitCSV(theText)
	set AppleScript's text item delimiters to ","
	set theList to text items of theText
	set AppleScript's text item delimiters to ""
	return theList
end splitCSV

on getColumnIndex(headerList, targetName)
	repeat with i from 1 to count of headerList
		if item i of headerList is targetName then return i
	end repeat
	error "項目名 '" & targetName & "' が見つかりませんでした。"
end getColumnIndex

on parseDateTime(dateStr, timeStr)
	set AppleScript's text item delimiters to "/"
	set dParts to text items of dateStr
	set m to item 1 of dParts as integer
	set d to item 2 of dParts as integer
	set y to item 3 of dParts as integer
	
	set AppleScript's text item delimiters to " "
	set tParts to text items of timeStr
	set hms to item 1 of tParts
	set ampm to item 2 of tParts
	
	set AppleScript's text item delimiters to ":"
	set hmsParts to text items of hms
	set h to item 1 of hmsParts as integer
	set min to item 2 of hmsParts as integer
	set s to round (item 3 of hmsParts as number)
	
	if ampm is "PM" and h < 12 then set h to h + 12
	if ampm is "AM" and h is 12 then set h to 0
	
	set theDate to current date
	set day of theDate to 1
	set year of theDate to y
	set month of theDate to m
	set day of theDate to d
	set hours of theDate to h
	set minutes of theDate to min
	set seconds of theDate to s
	
	set AppleScript's text item delimiters to ""
	return theDate
end parseDateTime

on parseFlyTimeToSeconds(flyTimeStr)
	set totalSec to 0
	set AppleScript's text item delimiters to "m"
	
	if flyTimeStr contains "m" then
		set minPart to text item 1 of flyTimeStr
		set secPartRaw to text item 2 of flyTimeStr
		set totalSec to (minPart as integer) * 60
	else
		set secPartRaw to flyTimeStr
	end if
	
	set AppleScript's text item delimiters to "s"
	if secPartRaw contains "s" then
		set secPart to text item 1 of secPartRaw
		try
			set totalSec to totalSec + (secPart as number)
		end try
	end if
	
	set AppleScript's text item delimiters to ""
	return totalSec
end parseFlyTimeToSeconds
