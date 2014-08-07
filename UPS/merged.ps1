$testData = Get-Content "D:\Users\Anthony\Desktop\readData.txt"
$last = 0

$test = 0

#Current Date used for dynaically creating links
$month = get-date -format "MM"
$day = get-date -format "dd"
$year = get-date -format "yy"

$scans = @()
$pph = 0
$user 
$results =""
$StartHour
$StartMinute
####################### Start FORM ###############################
function Select-Time(([ref]$startHour), ([ref]$startMinute)){

########################
$test = 0
# Edit This item to change the DropDown Values

[array]$hours = 1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24
[array]$minutes = "00", 15, 30, 45
[array]$message = "Employee Fails Expectation", "Employee Meets Expectation"

# This Function Returns the Selected Value and Closes the Form

function Return-DropDown {
 $script:startHour = $DropDown.SelectedItem.ToString()
 $script:startMinute = $DropDown2.SelectedItem.ToString()
 
 $Form.Close()
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$listBox1 = New-Object System.Windows.Forms.ListBox
$Form = New-Object System.Windows.Forms.Form

$Form.width = 800
$Form.height = 640
$Form.Text = ”Select A Start Time”

$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(100,10)
$DropDown.Size = new-object System.Drawing.Size(45,70)


ForEach ($Item in $hours) {
 [void] $DropDown.Items.Add($Item)
}
$DropDown.SelectedItem = 17
$Form.Controls.Add($DropDown)

################################################################

$DropDown2 = new-object System.Windows.Forms.ComboBox
$DropDown2.Location = new-object System.Drawing.Size(160,10)
$DropDown2.Size = new-object System.Drawing.Size(45,30)
ForEach ($Item in $minutes) {
 [void] $DropDown2.Items.Add($Item)
}
$DropDown2.SelectedItem = 45
$Form.Controls.Add($DropDown2)




##################################################################

$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,10) 
$DropDownLabel.size = new-object System.Drawing.Size(100,40) 
$DropDownLabel.Text = "Start Time: `n(Choose Closest Before Sort Start)"
$Form.Controls.Add($DropDownLabel)


$ColonLabel = new-object System.Windows.Forms.Label
$ColonLabel.Location = new-object System.Drawing.Size(150,10) 
$ColonLabel.size = new-object System.Drawing.Size(20,20) 
$ColonLabel.Text = ":"
$Form.Controls.Add($ColonLabel)

##################################################################
$DropDown3 = new-object System.Windows.Forms.ComboBox
$DropDown3.Location = new-object System.Drawing.Size(160, 40)
$DropDown3.Size = new-object System.Drawing.Size(160,30)
ForEach ($Item in $message) {
 [void] $DropDown3.Items.Add($Item)
}
$DropDown3.SelectedItem = "Employee Fails Expectation"
$Form.Controls.Add($DropDown3)

##################################################################

$DropDownLabel2 = new-object System.Windows.Forms.Label
$DropDownLabel2.Location = new-object System.Drawing.Size(10,45) 
$DropDownLabel2.size = new-object System.Drawing.Size(150,45) 
$DropDownLabel2.Text = "Send Message When"
$Form.Controls.Add($DropDownLabel2)


##################################################################

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox1 = New-Object System.Windows.Forms.CheckBox

$b1= $false
$b2= $false
$b3= $false

$handler_button1_Click= 
{
    #$listBox1.Items.Clear();    

    if ($checkBox1.Checked)     { $test = "test" } #$listBox1.Items.Add( "Checkbox 1 is checked"  )

    if ($checkBox2.Checked)    {  $test = $test + 1 } #$listBox1.Items.Add( "Checkbox 2 is checked"  )

    if ($checkBox3.Checked)    { $test = $test + 1; $test   } #$listBox1.Items.Add( "Checkbox 3 is checked") 

    #if ( !$checkBox1.Checked -and !$checkBox2.Checked -and !$checkBox3.Checked ) {   $listBox1.Items.Add("No CheckBox selected....")}

}

	
	
#$form.Controls.Add($listBox1)


$checkBox3.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 104
$System_Drawing_Size.Height = 24
$checkBox3.Size = $System_Drawing_Size
$checkBox3.TabIndex = 2
$checkBox3.Text = "CheckBox 3"
$checkBox3.Location =  new-object System.Drawing.Size(27, 200)
$checkBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox3.Name = "checkBox3"

#$form.Controls.Add($checkBox3)

$checkBox2.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 104
$System_Drawing_Size.Height = 24
$checkBox2.Size = $System_Drawing_Size
$checkBox2.TabIndex = 1
$checkBox2.Text = "CheckBox 2"
$System_Drawing_Point = New-Object System.Drawing.Point
$checkBox2.Location =  new-object System.Drawing.Size(27, 225)
$checkBox2.DataBindings.DefaultDataSourceUpdateMode = 0
#$checkBox2.Name = "checkBox2"

$form.Controls.Add($checkBox2)
    $checkBox1.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 104
    $System_Drawing_Size.Height = 24
    $checkBox1.Size = $System_Drawing_Size
    $checkBox1.TabIndex = 0
    $checkBox1.Text = "CheckBox 1"
    $checkBox1.Location = new-object System.Drawing.Size(27, 250)
    $checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox1.Name = "checkBox1"

#$form.Controls.Add($checkBox1)

$listBox1.FormattingEnabled = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 301
$System_Drawing_Size.Height = 212
$listBox1.Size = $System_Drawing_Size
$listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$listBox1.Name = "listBox1"
$listBox1.Location = new-object System.Drawing.Size(137, 213)
$listBox1.TabIndex = 3
#$form.Controls.Add($listBox1)

##################################################################

$Button = new-object System.Windows.Forms.Button
$Button.Location = new-object System.Drawing.Size(100,580)
$Button.Size = new-object System.Drawing.Size(100,20)
$Button.Text = "Confirm Selection"
$Button.Add_Click({Return-DropDown})
$form.Controls.Add($Button)

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()

$startHour
$startMinute

#########################################################################	End Form
}
Select-Time($StartHour, $StartMinute)


#########################################
# 	Send Email Funcion definition

function Send-Email($results){
	Start-Process Outlook
	$o = New-Object -com Outlook.Application
	$mail = $o.CreateItem(0)
	$mail.importance = 2
	$mail.subject = [string]::Format("Update for {0}:{1}", (get-date).hour, (get-date).minute)
	$mail.HTMLBody = $results
	#$mail.body = ""
	$last = $result2[9]
	[String[]]$recepients = "anthonymigliori@gmail.com" #Get-Content "Data\readData.txt"
	$recepients | % { $mail.Recipients.Add($_) }
	 
	sleep 3
	 
	$mail.send()
	sleep 5
	$o.Quit()
}
#################################### Main #####################################################
while(1)
{
	if((Get-date).hour -ge $startHour -and (Get-Date).minute -ge $startMinute)
	{
		sleep 900
		#######################       ONLY RUN ONCE
		$findusers = '[A-Z][A-Z][A-Z][A-Z][A-Z0-9][0-9][0-9][0-9][0-9][A-Z]'
		$scanners =  $testData | select-string -Pattern $findusers -AllMatches | % { $_.Matches } | % { $_.Value }

		foreach($scanner in $scanners)
		{
			New-Variable -Name "$scanner" -Value @()
		    Get-Variable -Name "$scanner" -ValueOnly
		}
		
		while(1)
		{
			if(((get-date).minute % 15) -eq 0)
			{
				sleep 60
				#this loop finds each employee and grabs there total scans for the night, this is stored into the scans array
				#a temporary results variable is created to store and email the amount of scans each scanner has made in the last 15 minutes
				foreach($scanner in $scanners)
				{
					$web = New-Object Net.WebClient
					$link = "http://bldg-web-pri/gss/OpMnr/EmployeeMonitor.asp?Date=$month%2F$day%2F$year&Sort=07&Employee=$empID&Accept=Accept"
					$string = $web.DownloadString($link)
					$result = $string | select-string -Pattern '(?<=\<[tT][dD]\>&nbsp;)(.*?)(?=\<\/[tT][dD])' -AllMatches | % { $_.Matches } | % { $_.Value }
	
					if($scans[0] -ge 1)
					{
						$results += [string]::Format("$scanner {0}`n", $result[9]- $scans[$index])
						$scans[$index] = $result[9]
			
						$index++
					}
	
					else
					{
					$results += [string]::Format("$scanner {0}`n", $result[9])	
					}
				
				}
				Send-Email($results)
				$results = ""
				if((get-date).hour -ge 21)
				{ break }
				sleep 800
			}
			
			else {sleep 1 }
		}
	}
	
	if((get-date).hour -ge 21)
	{ break }
	
	else{ sleep 1}
}





