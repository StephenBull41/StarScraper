#this script was created to report on the status of large amounts of item in transit(1000+)
#the script will take a csv list or tracking numbers with internal correlation information & produce a report on the status of the deliveries

Add-Type -AssemblyName System.Windows.Forms
#example complete URL: https://msto.startrack.com.au/track-trace/?id=WRZZ12345678
#expected con data format: sitename, camsID, siteID, callType, itemSwapped, ConNote, RetConNote

Write-Host "Expected data format: 'sitename,camsID,siteID,callType,itemSwapped,ConNote,RetConNote'"

#web page fields:
$lbl_status            = "__c1_lblStatus"           # Delivery status
$lbl_scan_depot        = "__c1_lblScanDepot"        # Location correlated to last scan time if not at destination
$lbl_service_type      = "__c1_lblService"          # type of delivery, premium etc.
$lbl_eta               = "__c1_lblETADate"          # Delivery ETA
$lbl_proof_of_delivery = "__c1_lblPOD"              # Proof of delivery(date/time)
$lbl_despatch_date     = "__c1_lblDespatchDate"     # Despatch Time
$lbl_delivery_DT       = "__c1_lblScanDateTime"     # Last scan time
$updpnl_TGrid          = "__c1_updpnlTrackingGrid"  # Grid containing tracking information / events

#get the data import
Write-Host "Press enter & select data export"
Read-Host

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = ("C:\Temp\") }
$null = $FileBrowser.ShowDialog()
$con_data = gc -Path $FileBrowser.FileName
$con_data = $con_data | select -Last ($con_data.Count -1)


Write-Host "Getting data from StarTrack"
$site_count = 0
#add a header to the CSV indicating the export fields
"camsID,ConNote,ConStatus,ConDeliveryDT,ConScanDsc,RetNote,RetStatus,Scan Date & Time,RetScanDesc" | Out-File -FilePath ($FileBrowser.FileName.Remove($FileBrowser.FileName.Length -4) + "_Output_" + (Get-Date).ToString("ddMMyyyy") + ".csv")

#for each item check for inbound & outbound deliveries.
foreach($site in $con_data){
    
    $site_arr = $site.Split(',')

    #check we're working with a valid line & request web pages for any tracking numbers provided
    if($site_arr[0] -ne "Name" -and $site_arr[0] -ne $null){

        $con_web_data = $null
        $ret_web_data = $null
        $site_con_str = $null
        $site_ret_str = $null
        if($site_arr[1] -ne $null){ $con_web_data = Invoke-WebRequest -Uri ("https://msto.startrack.com.au/track-trace/?id=" + $site_arr[1]) }
        if($site_arr[2] -ne $null){ $ret_web_data = Invoke-WebRequest -Uri ("https://msto.startrack.com.au/track-trace/?id=" + $site_arr[2]) }        


        #get the fields we care about out of the consignment page
        if($con_web_data -ne $null -and ($con_web_data.AllElements | where {$_.id -eq "__c1_lblStatus"}) -ne $null){
                    
            $con_status = $con_web_data.AllElements | where {$_.id -eq $lbl_status}
            $con_eta = $con_web_data.AllElements | where {$_.id -eq $lbl_eta}
            $con_pod = $con_web_data.AllElements | where {$_.id -eq $lbl_proof_of_delivery}
            $con_despatch_date = $con_web_data.AllElements | where {$_.id -eq $lbl_despatch_date}
            #$con_delivery_dt = $con_web_data.AllElements | where {$_.id -eq $lbl_delivery_DT}


            #can't select an elements text for grids, need to grab the HTML then clean it up, only interested in the 3rd & 5th column, index 2 & 4 (most recent scan event)
            #format individual tags into individual lines then select only <TD> tags
            $con_TGrid = ($con_web_data.AllElements | where {$_.id -eq $updpnl_TGrid}).innerHTML.Split("`r`n") | where {$_[2] -eq "D"}

            #remove tags to keep only displayed text.
            $con_grid_2 = $con_TGrid[2].Remove(0,4)
            $con_grid_2 = $con_grid_2.Remove($con_grid_2.IndexOf('<'))

            $con_grid_4 = $con_TGrid[4].Remove(0,4)
            $con_grid_4 = $con_grid_4.Remove($con_grid_4.IndexOf('<'))

            #build the first half of the output string for this item
            $site_con_str = $site_arr[0] + "," + $site_arr[1] + "," + $con_status.outerText + "," <#+ $con_delivery_dt.outerText + ","#> + $con_grid_2 + "," + $con_grid_4 + ","

        }else{ $site_con_str = $site_arr[0] + "," + $site_arr[1] + ",," + ",StarTrack consignment page unavailable,,," }

        #do it again for return tracking numbers

        if($ret_web_data -ne $null-and ($con_web_data.AllElements | where {$_.id -eq "__c1_lblStatus"}) -ne $null){

            $ret_status = $ret_web_data.AllElements | where {$_.id -eq $lbl_status}
            $ret_eta = $ret_web_data.AllElements | where {$_.id -eq $lbl_eta}
            $ret_pod = $ret_web_data.AllElements | where {$_.id -eq $lbl_proof_of_delivery}
            $ret_despatch_date = $ret_web_data.AllElements | where {$_.id -eq $lbl_despatch_date}
            $ret_delivery_dt = $ret_web_data.AllElements | where {$_.id -eq $lbl_delivery_DT}     

            $ret_grid_2 = $ret_TGrid[2].Remove(0,4)
            $ret_grid_2 = $ret_grid_2.Remove($ret_grid_2.IndexOf('<'))

            $ret_grid_4 = $ret_TGrid[4].Remove(0,4)
            $ret_grid_4 = $ret_grid_4.Remove($ret_grid_4.IndexOf('<'))

            $ret_TGrid = ($ret_web_data.AllElements | where {$_.id -eq $updpnl_TGrid}).innerHTML.Split("`r`n") | where {$_[2] -eq "D"}           
            $site_ret_str = $site_arr[2] + "," + $ret_status.outerText + ","<# + $ret_delivery_dt.outerText + ","#> + $ret_grid_2 + "," + $ret_grid_4 + ","

        }else{ $site_ret_str =",StarTrack return page unavailable,,," }

        $out_file_str = $site_con_str + $site_ret_str
        $out_file_str | Out-File -FilePath ($FileBrowser.FileName.Remove($FileBrowser.FileName.Length -4) + "_Output_" + (Get-Date).ToString("ddMMyyyy") + ".csv") -Append

        $site_count++
        cls
        Write-Host "completed" $site_count "sites"

    }
    #Start-Sleep -Milliseconds 250     
}
write-host "Complete"