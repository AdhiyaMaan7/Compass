Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Add-Type -AssemblyName System.IO.Compression.FileSystem


$grayForm = New-Object System.Windows.Forms.Form
    #$grayForm.Size = New-Object System.Drawing.Size(500,350)
    $grayForm.height = 350
    $grayForm.Width = 520
    $grayForm.Text = "Gray Package"
    $grayForm.StartPosition = "CenterScreen"
    $grayForm.MaximizeBox = $false
    $grayForm.MinimizeBox = $true
    $grayForm.BackColor = "#0e4d8d"
    $grayForm.FormBorderStyle = "FixedDialog"
    $grayForm.HelpButton = $true
    $grayForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]:: none
    


    $shadow_Panel = New-Object System.Windows.Forms.Panel
    $shadow_Panel.Width = $grayForm.Width
    $shadow_Panel.Height = $grayForm.Height
    $shadow_Panel.BackColor = [System.Drawing.Color]::FromArgb(100,0,0,0)
    $shadow_Panel.Location = New-Object System.Drawing.Point(5, 5)


    #Modified title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.BackColor = "#ffc425"
    $titleBar.Height = "35"
    $titleBar.Dock = [System.Windows.Forms.DockStyle]:: Top
    
    $titleBar.Add_MouseDown({
        
        $grayForm.Capture = $false
        $msg = [Windows.Forms.Message]::Create($grayForm.Handle, 0XA1, [System.IntPtr]::Zero)
        $grayForm.DefWndProc([ref]$msg)
    })

    $title_Label = New-Object System.Windows.Forms.Label
    $title_Label.Text = "Gray Automation"
    $title_Label.ForeColor = "Black"
    $title_Label.Width = 200
    $title_Label.Location = New-Object System.Drawing.Point(20, 9)

    $close_Button = New-Object System.Windows.Forms.Button
    $close_Button.Text = "X"
    $close_Button.ForeColor = "White"
    $close_Button.BackColor = "#e31837"
    $close_Button.Dock = [System.Windows.Forms.DockStyle]::Right
    $close_Button.Width = 60
    #$close_Button.Location = New-Object System.Drawing.Point(8,20)
    $close_Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $close_Button.FlatAppearance.BorderSize = 0
    $close_Button.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::bold )
    $close_Button.Add_Click({$grayForm.close()})
    
    $miniMize_Button = New-Object System.Windows.Forms.Button
    $miniMize_Button.Text = "-"
    $miniMize_Button.ForeColor = "#666666"
    $miniMize_Button.BackColor = "#dbdbdb"
    $miniMize_Button.Dock = [System.Windows.Forms.DockStyle]::Right
    $miniMize_Button.Width = 60
    $miniMize_Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $miniMize_Button.FlatAppearance.BorderSize = 0
    $miniMize_Button.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold )
    $miniMize_Button.Add_Click({$grayForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized})

    $titleBar.Controls.AddRange(@($title_Label, $miniMize_Button,$close_Button))

$panel_bottom = New-Object System.Windows.Forms.Panel
    $panel_bottom.BackColor = [System.Drawing.Color]::FromArgb(170,255,255,255)
    $panel_bottom.Height = 30
    $panel_bottom.Dock = [System.Windows.Forms.DockStyle]::Bottom

    $contact = New-Object System.Windows.Forms.Label
    $contact.Width = 300
    $contact.Location = New-Object System.Drawing.Point(30,8)
    $contact.Text = "Created by gopim5@deloitte.com"
    $contact.BackColor = [System.Drawing.Color]::FromArgb(0,0,0,0)
    $contact.ForeColor = [System.Drawing.Color]::FromArgb(100,255,255,255)
    $contact.Font = New-Object System.Drawing.Font("Arial", 8,
    [System.Drawing.FontStyle]::Italic
    )
    # $panel_bottom.Controls.AddRange(@($contact))

    $Feedback = New-Object System.Windows.Forms.Button
    # $Feedback.Width = 300
    $Feedback.Location = New-Object System.Drawing.Point(430,8)
    $Feedback.Text = "Help >>"
    $Feedback.BackColor = [System.Drawing.Color]::FromArgb(0,0,0,0)
    $Feedback.ForeColor = [System.Drawing.Color]::FromArgb(100,255,255,255)
    $Feedback.Font = New-Object System.Drawing.Font("Arial", 8,
    [System.Drawing.FontStyle]::Italic
    )
    $Feedback.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

    $panel_bottom.Controls.AddRange(@($contact,$Feedback))









    $grayForm.Font = New-Object System.Drawing.Font("Arial", 12,
    [System.Drawing.FontStyle]::Regular
    )

$pathLabel = New-Object System.Windows.Forms.label
    $pathLabel.Text = "Enter a Working Path: "
    $pathLabel.AutoSize = $true
    $pathLabel.ForeColor = "#ffffff"
    $pathLabel.Location = New-Object System.Drawing.Point(30,60)


$pathTextBox = New-Object System.Windows.Forms.TextBox
    #$pathTextBox.Size = New-Object System.Drawing.Point(220)
    $pathTextBox.Multiline = $true
    $pathTextBox.Width = 230
    $pathTextBox.Height = 55
    $pathTextBox.Location = New-Object System.Drawing.Point(250,60)
    $pathTextBox.Font = New-Object System.Drawing.Font("Arial", 10,
    [System.Drawing.FontStyle]::Regular 
    )

$trackNoLabel = New-Object System.Windows.Forms.label
    $trackNoLabel.Text = "Enter Tracking Number: "
    $trackNoLabel.AutoSize = $true
    $trackNoLabel.ForeColor = "#ffffff"
    $trackNoLabel.Location = New-Object System.Drawing.Point(30,137)


$trackNoTextBox = New-Object System.Windows.Forms.TextBox
    $trackNoTextBox.Multiline =$true
    $trackNoTextBox.Width = 200
    $trackNoTextBox.Height = 30
    #$trackNoTextBox.Size = New-Object System.Drawing.Point(130,20)
    $trackNoTextBox.Location = New-Object System.Drawing.Point(250,137)
    $trackNoTextBox.Font = New-Object System.Drawing.Font("Arial", 14,
    [System.Drawing.FontStyle]::Regular
    )

$packageListLabel = New-Object System.Windows.Forms.label
    $packageListLabel.Text = "Select Process Type: "
    $packageListLabel.AutoSize = $true
    $packageListLabel.ForeColor = "#ffffff"
    $packageListLabel.Location = New-Object System.Drawing.Point(30, 187)


$packageListComboBox = New-Object System.Windows.Forms.ComboBox
    $packageListComboBox.Width = 230
    $packageListComboBox.Height = 30
    $packageListComboBox.Location = New-Object System.Drawing.Point(250, 187)
    #$packageListComboBox.Value = "Pre-Package"
    $packageListComboBox.Font = New-Object System.Drawing.Font("Arial", 14,
    [System.Drawing.FontStyle]::Regular
    )
    #add package types
    @($packageListComboBox.Items.Add("Folder Structure")
    $packageListComboBox.Items.Add("Final Package")
    $packageListComboBox.Items.Add("Revision-Unzip")
    #$packageListComboBox.Items.Add("Pickup-Rename")
    ) > $null
                    


$button = New-Object System.Windows.Forms.Button
    $button.Text = "Start Process >>"
    $button.Size = New-Object System.Drawing.Point(160,45)
    $button.Location = New-Object System.Drawing.Point(250,250)
    $button.BackColor = "#ffc425"
    $button.ForeColor = "#000000"
    $button.Font = New-Object System.Drawing.Font("Arial", 13,
    [System.Drawing.FontStyle]::Regular
    )


$buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Size = New-Object System.Drawing.Point(100,30)
    $buttonCancel.Location = New-Object System.Drawing.Point(240,110)
    $buttonCancel.BackColor = "#ffffff"
    $buttonCancel.Enabled = $false


$progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20,150)
    $progressBar.Size = New-Object System.Drawing.Point(380,20)
    $progressBar.Maximum = 100
    $progressBar.Minimum = 0
    $progressBar.Step = 1
    $progressBar.Text = "Wait..."
    


$grayForm.Controls.AddRange(

    @($titleBar, $panel_bottom, $pathLabel, $pathTextBox, $trackNoLabel, $trackNoTextBox,
    $packageListLabel, $packageListComboBox, $button#,$shadow_Panel #$buttonCancel, $progressBar

))


function pre_Package{
    function rename_foldercreate{
        $newFolder = $Tracking + "_" + $_.BaseName
        $reName = $newFolder + $_.Extension
        New-Item -ItemType directory -Name ".\$newFolder"
        Move-Item -Path $_.FullName -Destination ".\$newFolder\$reName"
    }
    
    if (Test-Path -Path "*.fla"){
    
        New-Item -ItemType directory -Name ".\Source"
        if(Test-Path -Path  ".\300x250.psd"){
            Move-Item -Path 300x250.psd -Destination ".\Source"
        }else{Move-Item -Path "*.psd" -Destination ".\Source" }

        if(Test-Path -Path *.psd){
        New-Item -ItemType directory -Name ".\Source\Comp"
        Move-Item -Path *.psd -Destination ".\Source\Comp"
        }

        Get-ChildItem *.fla | ForEach-Object{ rename_foldercreate }
    
    }else{Get-ChildItem *.psd | ForEach-Object{ rename_foldercreate }}
    
    New-Item -ItemType directory -Name ".\source_$Tracking"
    Get-ChildItem -directory  -Exclude "source_$Tracking" | Move-Item  -Destination ".\Source_$Tracking"
}
#PrePackage End


# function Revision-Unzip {
   
#     # Action when this condition is true
#     if(Test-Path -Path ".\*.zip"){
#         Expand-Archive -Path ".\*.zip" -DestinationPath .\ -Force
#         # [System.Windows.Forms.MessageBox]::Show("Unzipped Successfully")
#         }
#     else{[System.Windows.Forms.MessageBox]::Show("No zip file found")}
    
# }





#final_Package Start
function final_Package{
# Animate start
if(Test-Path -Path *\*.fla){
        #For index set
        Get-ChildItem -Path ".\*\index.html" -Recurse | ForEach-Object{
            (Get-Content $_ ) -replace "TRACKING", "$Tracking" | Set-Content $_.FullName
        }


        if(Test-Path -Path ".\Source\Comp"){Move-Item -Path ".\Source\Comp" -Destination "..\" -Force}
        if(Test-Path -Path ".\*\*.fla"){Move-Item -Path ".\*\*.fla" -Destination .\Source -Force}
        if(Test-Path -Path .\TrafficFiles){ Remove-Item -Path .\TrafficFiles -Recurse -Force}

        Get-ChildItem -Directory | ForEach-Object {
            Compress-Archive -Path $_\* -DestinationPath ".\$_.zip" -Force -CompressionLevel Fastest
            # icacls ".\$_.zip" /grant Everyone:F
        }

        Get-ChildItem "*.zip" -Exclude "Source.zip" | Compress-Archive -DestinationPath .\TrafficFiles.zip  -Force
        Expand-Archive -Path .\TrafficFiles.zip -DestinationPath .\TrafficFiles -Force

        New-Item -ItemType directory -Name downloads -Force
        Get-ChildItem -Directory -Exclude Source, downloads, TrafficFiles | Copy-Item -Destination downloads -Force -Recurse

        $verDiff = "/index.html" # For eproof
}# Animate End





#Static Start
elseif(Test-Path -Path *\*.psd) {
    Add-Type -AssemblyName System.Drawing
    Get-ChildItem "*\*.jpg" | ForEach-Object{

    $name = $_.BaseName
    $image = [System.Drawing.Image]::FromFile($_.FullName)
    $width = $image.Width
    $height = $image.Height

    $htmlTemplate = @"
    <html>
    <head> 
    <title>$name</title> 
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> 
    </head> 
    <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
        <!-- Save for Web Slices ($name.psd) --> 
        <img src="$name.jpg" width="$width" height="$height" alt=""> 
        <!-- End Save for Web Slices --> 
    </body> 
    </html>
"@
    
    New-Item -ItemType File -Path $name -Name index.html -Value $htmlTemplate -Force
    $image.Dispose()

    }

    New-Item -ItemType directory -Name "TrafficFiles" -Force
    Copy-Item -Path *\*.jpg -Exclude TrafficFiles -Destination TrafficFiles -Recurse -Force

    Get-ChildItem -Directory | ForEach-Object {   
        Compress-Archive -Path $_\* -DestinationPath ".\$_.zip" -Force -CompressionLevel Fastest
        # icacls ".\$_.zip" /grant Everyone:F
    }

    #if(Test-Path -Path ".\TrafficFiles"){Remove-Item -Path ".\TrafficFiles"}
    New-Item -ItemType directory -Name downloads -Force
    Copy-Item -Path ".\*\*.jpg" -Destination downloads -Force

    $verDiff = ".jpg" # For eproof
}#Static End


$eproofTemplate = @"

<!DOCTYPE html> 
<html> 
<head> 
<link href="https://fonts.googleapis.com/css?family=Varela+Round" rel="stylesheet"> 
<meta name="viewport" content="width=device-width, initial-scale=1"> 
<meta http-equiv="Content-Security-Policy" content="upgrade-insecure-requests"> 
<title>$Tracking</title> 
<style> 
    .CSview{padding:150px 0 50px 0} 
    .CSproof{border:0px; margin:auto; min-width:360px; max-width:1070px} 
    a.DL:link{background-color:#009ac7; color:#ffffff; padding:5px; text-decoration:none; display:inline-block; border:0px solid gray; border-radius:0px; width:200px;} 
    a.DL:visited{background-color:#009ac7; color:#000000; padding:5px; text-decoration:none; display:inline-block; border:1px solid darkblue; border-radius:10px; width:200px;} 
</style> 
</head> 
<body style="font-family:Varela Round; text-align:center; background-color:#ffffff"> 
<div style="width:100%; padding-top:3px; position:fixed; left:0; top:0; background-color:#ffffff; min-width:360px; background-size: cover;"> 
    <div style="margin:20px 30px 100px 50px; float:left"> 
    <img src="https://gray.tv/uploads/redesign/HOME/GRAY-horizontal-logo-280x140-white.png" onerror="this.style.display='none'" alt="energyaus" style="border:0 !important; position:absolute; left:30px; bottom: 25px; top: 20px; width:8%;"> 
    </div> 
    <div style="width:100%; height:6px; background-color:#e31837; position:absolute; left:0; bottom:0;"></div> 
    <div style="width:100%; height:2px; background-color:#33CBFF; position:absolute; left:0; bottom:0;"></div> 
</div> 
<div class="CSview"> 
    <div class="CSproof"> 
    <!-- *************************************************************** BEGIN Animate Bundle ********************************************************************************* --> 
"@ 	
    if(Test-Path -Path .\*300x250){
    $eproofTemplate += @"

    <!-- ******** BEGIN Animate Bundle_300x250 Bundle - 5 ******** --> 
	<div style="width: 305px; height:270px; display:inline-block; margin: 10px;"> <span>300x250</span> 
    <object style="width:300px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_300x250$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='250px'"> </object></div> 
    <!-- ******** END Animate Bundle_300x250 Bundle - 5 ********* --> 
"@}
    if(Test-Path -Path .\*160x600){
    $eproofTemplate += @"

    <!-- ******** BEGIN Animate Bundle_160x 600 Bundle - 5 ******** --> 
	<div style="width: 165px; height:620px; display:inline-block; margin: 10px;"> <span>160x600</span> 
    <object style="width:160px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_160x600$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='600px'"></object></div>
    <!-- ******** END Animate Bundle_160x 600 Bundle - 5 ********* --> 
"@}
    if(Test-Path -Path .\*300x600){
    $eproofTemplate += @"

    <!-- ******** BEGIN Animate Bundle_300x600 Bundle - 5 ******** --> 
    <div style="width: 305px; height:620px; display:inline-block; margin: 10px;"> <span>300x600</span> 
  	<object style="width:300px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_300x600$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='600px'"></object></div> 
    <!-- ******** END Animate Bundle_300x600 Bundle - 5 ********* --> 
"@}
    if(Test-Path -Path .\*320x50){ 
    $eproofTemplate += @"

    <!-- ******** BEGIN Animate Bundle_320x50 Bundle - 5 ******** --> 
    <div style="width: 325px; height:70px; display:inline-block; margin: 10px;"> <span>320x50</span> 
    <object style="width:320px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_320x50$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='50px'"></object></div> 
  	<!-- <******* END Animate Bundle_320x50 Bundle - 5 ********* --> 
"@}
    if(Test-Path -Path .\*728x90){ 
    $eproofTemplate += @"

    <!-- ******** BEGIN Animate Bundle_728x90 Bundle - 5 ******** --> 
 	<div style="width: 733px; height:110px; display:inline-block; margin: 10px;"> <span>728x90</span> 
    <object style="width:728px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_728x90$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='90px'"></object></div> 
    <!-- ******** END Animate Bundle_728x90 Bundle - 5 ********* --> 
"@}  
    if(Test-Path -Path .\*1024x90){
    $eproofTemplate += @"
    
    <!-- ******** BEGIN Animate Bundle_1024x90 Bundle - 5 ******** --> 
    <div style="width: 1029px; height:110px; display:inline-block; margin: 10px;"> <span>1024x90</span> 
    <object style="width:1024px; border:0px; background-color:white; overflow:hidden;" data="downloads/$($Tracking)_1024x90$($verDiff)" onerror="this.style.display='none'" onload="this.style.height='90px'"></object></div> 
    <!-- ******** END Animate Bundle_1024x90 Bundle - 5 ********* --> 
"@}	
    $eproofTemplate += @"
    <!-- *************************************************************** END Animate Bundle ********************************************************************************* --> 
    </div> 
</div> 
</body> 
</html> 

"@

Compress-Archive -Path "*.zip" -DestinationPath ..\source_$Tracking.zip  -Force -CompressionLevel Fastest
Remove-Item -Path "*.zip" -Force
if(Test-Path -Path "..\$Tracking"){Remove-Item -Path "..\$Tracking" -Force -Recurse}

New-Item -ItemType File -Path ".\" -Name "$Tracking.html" -Value $eproofTemplate -Force

# New-Item -ItemType Directory -Name "..\$Tracking" -Force
# Move-Item -Path ".\downloads\*"  -Destination "..\$Tracking\downloads\" -Force -Recurse
# Move-Item -Path  ".\$Tracking.html" -Destination "..\$Tracking\" -Force

Compress-Archive -Path ".\downloads", ".\$Tracking.html"  -DestinationPath "..\$Tracking.zip" -Force -CompressionLevel Fastest | Out-Null

# icacls "..\source_$Tracking.zip" /grant Everyone:F
# icacls "..\$Tracking.zip" /grant Everyone:F

Remove-Item -Path ".\downloads", ".\$Tracking.html"  -Force -Recurse
Expand-Archive -Path "..\$Tracking.zip" -DestinationPath ..\$Tracking
Start-Process -filepath "..\$Tracking\$Tracking.html" -WindowStyle Normal

# Remove-Item -Path "..\$Tracking.zip" -Force


}
#FInal Package End

#Set Working Path
#Set-Location $pathTextBox.Text

#Set Tracking number
#$Tracking = $trackNoTextBox.Text



$button.add_click({
    
    #Set Working Path

    Set-Location $pathTextBox.Text

    #Set Tracking number
    $Tracking = $trackNoTextBox.Text

    $button.BackColor = "#ffffff"
    $button.Text = "Wait"
    $button.Enabled = $false
    


    if(($packageListComboBox.SelectedItem -eq "") -or ($pathTextBox.Text -eq "") -or ($trackNoTextBox.Text -eq "")){
    
    [System.Windows.Forms.MessageBox]::Show("Please Fill all Boxes" )
    
    }else{

        if($packageListComboBox.SelectedItem -eq "Folder Structure"){

            if(Test-Path -Path .\Source_$Tracking){[System.Windows.Forms.MessageBox]::Show("Already Source Folder created!!!")}
            elseif(-Not (Test-Path -Path .\*.*)){[System.Windows.Forms.MessageBox]::Show("Path is Empty. Unable to Process!")}
            elseif((Test-Path -Path .\*.fla) -and (-not(Test-Path -Path .\*.psd))){[System.Windows.Forms.MessageBox]::Show("Photoshop file missing!!")}
            else{pre_Package}
        }

        elseif($packageListComboBox.SelectedItem -eq "Final Package"){
            
        Set-Location "$($pathTextBox.Text)\Source_$Tracking"
        
        final_Package
        
        }
        
        elseif ($packageListComboBox.SelectedItem -eq "Revision-Unzip") {
                <# Action when this condition is true #>
                if(Test-Path -Path ".\source_$Tracking"){

                    

                    Set-Location "$($pathTextBox.Text)\Source_$Tracking"
                    if(Test-Path -Path ".\Source"){
        
                        Get-ChildItem -Path ".\Source\*.fla" | ForEach-Object{
                            $FileFolder = $_.BaseName
                            Move-Item -Path "$_" -Destination ".\$FileFolder" -Force
                            
                        }
                    }
                }

                else {
                    if(Test-Path -Path ".\*.zip"){

                        Get-ChildItem -Path ".\*.zip" | ForEach-Object{
                        # Read-Host $_.BaseName
                            $UnzipFolder = $_.BaseName
                            Expand-Archive -Path "$_" -DestinationPath ".\$UnzipFolder" -Force         
                        }
                        Remove-Item -Path ".\*.zip" -Force
                    }    
                    # else{[System.Windows.Forms.MessageBox]::Show("No zip file found")}
                    if (Test-Path -Path ".\source_$Tracking"){
        
                        Set-Location "$($pathTextBox.Text)\Source_$Tracking"
                        Remove-Item -Path ".\TrafficFiles.zip" -Force -Recurse
        
                        Get-ChildItem -Path ".\*.zip" | ForEach-Object{
                            # Read-Host $_.BaseName
                            $FileFolder = $_.BaseName
                            Expand-Archive -Path "$_" -DestinationPath ".\$FileFolder" -Force          
                            }
                        
                        Remove-Item -Path ".\*.zip" -Force

                        if(Test-Path -Path ".\Source"){
        
                            Get-ChildItem -Path ".\Source\*.fla" | ForEach-Object{
                                $UnzipFolder = $_.BaseName
                                Move-Item -Path "$_" -Destination ".\$UnzipFolder" -Force
                                
                            }
                        }

                    }
                    
                    
                }

        }
     

    }


    $button.Enabled = $true
    $button.Text = "Start Process >>"
    $button.BackColor = "#ffc425"
    $button.ForeColor = "#000000"
    
     
 #[System.Windows.Forms.MessageBox]::Show("Selected item : $($packageListComboBox.SelectedItem)" )
 })

 $Feedback.Add_Click({
    # $Feedback.BackColor = "#ffffff"
    $Feedback.Text = "Wait"
    $Feedback.Enabled = $false
    Start-Process -filepath "https://teams.microsoft.com/l/chat/0/0?users=gopim5@deloitte.com" -WindowStyle Normal 
    $Feedback.Enabled = $true
    $Feedback.Text = "Help >>"
    # $Feedback.BackColor = "#ffffff"
})


$grayForm.Add_Shown({$grayForm.Activate()})
[void]$grayForm.ShowDialog()
$grayForm.Dispose()
