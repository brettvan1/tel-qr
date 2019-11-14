#
#|: | Author:  Brett van Gennip
#| :| Email:   brett.vangennip@lhins.on.ca
#|: | Purpose: Open Internet explorer windows in Particular 
#|: |                             

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$wshell = New-Object -ComObject Wscript.Shell
# $searchbase = uncomment this and put in your searchbase within AD

function ddb(){

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Select an job:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

[void] $listBox.Items.Add('Extension')
[void] $listBox.Items.Add('Dell')
#[void] $listBox.Items.Add('SMS')
[void] $listBox.Items.Add('ADlookup-ext')
[void] $listBox.Items.Add('ADlookup-mob')
[void] $listBox.Items.Add('phonenumb')
#[void] $listBox.Items.Add('Future4')


$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem
	switch ( $x ) {
        'Extension'      { $y=2 }
		'Dell'   	     { $y=1 }
		'ADlookup-ext'   { $y=3 }
		'ADlookup-mob'   { $y=4 }
		'phonenumb'		 { $y=5 }
		}
		
	return $y
}
return $false
}
function Extension() {

$ext=[Microsoft.VisualBasic.Interaction]::InputBox("Enter in extension of person?", "Enter in extension of person you would like to call.", "")

if($ext){
	$sss="tel:905-895-1334" + ",," + $ext
	launchie($sss)
}

}
function Dell() {

$dell=[Microsoft.VisualBasic.Interaction]::InputBox("calling dell?", "Enter in the Express service tag", "")

	if($dell){
		$sss= "+1-866-362-5350" + ",,,,,,,,3,,,," + $dell
		launchie($sss)
		}
}
function ADlookup-ext() {

$getusr=get-aduser -filter * -searchbase $searchbase -properties samaccountname,officephone | out-gridview -passthru
	$getusr=$getusr.officephone
	$ext=$getusr.substring($getusr.length-4)

if($ext){
	$sss="tel:905-895-1334" + ",," + $ext
	launchie($sss)
}


	
}
function ADlookup-mob() {

$getusr=get-aduser -filter * -searchbase $searchbase -properties samaccountname,mobilephone | out-gridview -passthru
$getusr=$getusr.mobilephone
$ext=$getusr.substring($getusr.length-4)

$answer = $wshell.Popup("SMS Text user ?",0,"Alert",64+4)

if($answer -eq 6){
	$sss="sms:" + $ext
	launchie($sss)
}else{
	$sss="tel:" + $ext
	launchie($sss)
}

}
function phonenumb() {

	$tel=[Microsoft.VisualBasic.Interaction]::InputBox("Easy Telephone QR Code", "Enter in a phone number to call.","")
	if($tel){
	$sss="tel:" + $tel
	launchie($sss)
	}
}
function launchie ($sss) {

$url="https://api.qrserver.com/v1/create-qr-code/?data=" + $sss + "&size=200x200"

$ie = new-object -comobject InternetExplorer.Application 
 
$ie.visible = $true 
 
#$ie2 = $ie.Width = 200  
 
$ie.top = 200; $ie.width = 800; $ie.height = 600 ; $ie.Left = 100 
 
$ie.navigate($url)

}

$y1=ddb
$y1
if($y1){

switch ( $y1 ) {
        '2'   { Extension }
		'1'   { Dell }
		'3'   { ADlookup-ext }
		'4'   { ADlookup-mob }
		'5'	  { phonenumb }
		}


}else{

		 [Microsoft.VisualBasic.Interaction]::msgBox("No like? Try again?")

}


