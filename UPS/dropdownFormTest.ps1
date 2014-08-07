########################

# Edit This item to change the DropDown Values

[array]$hours = 1,2,3,4,5,6,7,8,9,10,11,12
[array]$minutes = "00", 15, 30, 45

# This Function Returns the Selected Value and Closes the Form

function Return-DropDown {
 $script:Choice = $DropDown.SelectedItem.ToString()
 $Form.Close()
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")


$Form = New-Object System.Windows.Forms.Form

$Form.width = 300
$Form.height = 150
$Form.Text = ”DropDown”

$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(100,10)
$DropDown.Size = new-object System.Drawing.Size(130,30)

ForEach ($Item in $hours) {
 [void] $DropDown.Items.Add($Item)
}

$Form.Controls.Add($DropDown)

$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,10) 
$DropDownLabel.size = new-object System.Drawing.Size(100,20) 
$DropDownLabel.Text = "Items"
$Form.Controls.Add($DropDownLabel)

$Button = new-object System.Windows.Forms.Button
$Button.Location = new-object System.Drawing.Size(100,50)
$Button.Size = new-object System.Drawing.Size(100,20)
$Button.Text = "Select an Item"
$Button.Add_Click({Return-DropDown})
$form.Controls.Add($Button)

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()

$FirstName = $Choice
$LastName = "Test"
$Name = $FirstName + " " + $LastName
#Get First Initial from Firstname.
#From position 0 select the first 1 character
$FirstInitial = $FirstName.Substring(0,1)


New-Mailbox -UserPrincipalName $Email -Alias $Alias -Database "NonTeam" -Name $Name -OrganizationalUnit "Test Team" -Firstname $FirstName -LastName $LastName -DisplayName $Name -ResetPasswordOnNextLogon $true

########################