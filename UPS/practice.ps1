$startTime = 0
$testData = Get-Content "D:\Users\Anthony\Desktop\readData.txt"
$last = 0

$test = 0
<<<<<<< HEAD

$findUsers = '[A-Z][A-Z][A-Z][A-Z][A-Z0-9][0-9][0-9][0-9][0-9][A-Z]'
$parsePage = '(?<=\<[tT][dD]\>&nbsp;)(.*?)(?=\<\/[tT][dD])'
=======
$pattern =  '(?<=\<[tT][dD]\>)'
$findusers = '[A-Z][A-Z][A-Z][A-Z][A-Z0-9][0-9][0-9][0-9][0-9][A-Z]'
>>>>>>> FETCH_HEAD
$scanners =  $testData | select-string -Pattern $findusers -AllMatches | % { $_.Matches } | % { $_.Value }

#Current Date used for dynaically creating links
$month = get-date -format "MM"
$day = get-date -format "dd"
$year = get-date -format "yy"

<<<<<<< HEAD
$scans = @()


$results = ""

$index = 0
foreach($scanner in $scanners)
{
	$link = "http://bldg-web-pri/gss/OpMnr/EmployeeMonitor.asp?Date=$month%2F$day%2F$year&Sort=07&Employee=$empID&Accept=Accept"
	$result = $testData | select-string -Pattern $parsePage -AllMatches | % { $_.Matches } | % { $_.Value }
	#$string = $web.DownloadString($link)
	if($scans[0] -ge 1)
	{
		$results += [string]::Format("$scanner {0}`n", $result[9]-scans[$index])
		$index++
	}
	
	else
	{
	$results += [string]::Format("$scanner {0}`n", $result[9])	
	}
	
	
	
}
$pph = @()
$index = 0
foreach($scanner in $scanners)
{
	$pph = $pph + $result[9]
	$index++
}

$results
$web = New-Object Net.WebClient
#$string = $web.DownloadString($link)

if((Get-Date).hour -eq $startHour -and (Get-date).minute -eq $startMinute)
=======
$scans = 0
$pph = @() #create empty array


$results =""
$index = 0
foreach($scanner in $scanners)
{
	$link = "http://bldg-web-pri/gss/OpMnr/EmployeeMonitor.asp?Date=$month%2F$day%2F$year&Sort=07&Employee=$scanner&Accept=Accept"
	$link2 = "http://bldg-web-pri/gss/OpMnr/EmployeeMonitor.asp?Date=$month%2F$day%2F$year&Sort=07&Employee=$empID&Accept=Accept"
	$result2 = $testData | select-string -Pattern $pattern -AllMatches | % { $_.Matches } | % { $_.Value }
	#$string2 = $web.DownloadString($link2)
	[int32]::TryParse($result2[14], [ref]$test)
	if($result2[14] -ge 0)
	{
	    $last15 = [int32]($test - $last15)
	}
	$results += [string]::Format("$scanner {0}`n", $result2[9])
	
}


$web = New-Object Net.WebClient
#$string = $web.DownloadString($link)

$index = 0
 

if((Get-Date).hour -ge 17)
>>>>>>> FETCH_HEAD
{

}

function sendEmail($results){
	Start-Process Outlook
	$o = New-Object -com Outlook.Application
	$mail = $o.CreateItem(0)
	$mail.importance = 2
	$mail.subject = [string]::Format("$results")
	$mail.HTMLBody = $string
	#$mail.body = ""
	$last = $result2[9]
	[String[]]$recepients = Get-Content "Data\readData.txt"
	$recepients | % { $mail.Recipients.Add($_) }
	 
	sleep 3
	 
	$mail.send()
	sleep 5
	$o.Quit()
}
