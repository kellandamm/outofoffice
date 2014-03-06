$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())
 
$mbHash = @{ }

$tmValHash = @{ }
$tidx = 0
for($vsStartTime=[DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"));$vsStartTime -lt [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00")).AddDays(1);$vsStartTime = $vsStartTime.AddMinutes(30)){
	$tmValHash.add($vsStartTime.ToString("HH:mm"),$tidx)	
	$tidx++
}

Get-DistributionGroupMember -Identity "ADD-DISTRO-ALIAS" | foreach-object{
if ($mbHash.ContainsKey($_.PrimarySmtpAddress.ToString()) -eq $false){
$mbHash.Add($_.PrimarySmtpAddress.ToString(),$_.DisplayName)

	}
}
$Attendeesbatch = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.AttendeeInfo]))

$StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"))
$EndTime = $StartTime.AddDays(1)


$displayStartTime =  [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 07:30"))
$displayEndTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 18:00"))

$drDuration = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($StartTime,$EndTime)
$AvailabilityOptions = new-object Microsoft.Exchange.WebServices.Data.AvailabilityOptions
$AvailabilityOptions.RequestedFreeBusyView = [Microsoft.Exchange.WebServices.Data.FreeBusyViewType]::DetailedMerged
$fbBoard = $fbBoard + "<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.01 Transitional//EN`"`r`n `"http://www.w3.org/TR/html4/loose.dtd`">`r`n <html>`r`n <head>`r`n<title>CHANGE TO LOCATION</title> <link rel=`"stylesheet`" type=`"text/css`" href=`"menu.css`" />`r`n</head>`r`n<body>"

$fbBoard = $fbBoard + "<table>"

$fbBoard = $fbBoard + "<tr>"
$fbBoard = $fbBoard + ("<td width=`"870`" align=`"center`"><img src=`"toplogoCR.jpg`" width=`"870`" height=`"90`" border=`"0`"></td>")

  $fbBoard = $fbBoard + ("</tr>")
              $fbBoard = $fbBoard + ("</table>")


$fbBoard = $fbBoard + "<table>"
 
    $fbBoard = $fbBoard + "<tr valign='top'>"
   
              $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#CDEB8B'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td align=center>&nbsp;Tentative</td>")
               $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#6BBA70'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Free</th>")
             $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#356AA0'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Busy</td>")
               $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#330099'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Out of Office</td>")
             $fbBoard = $fbBoard + ("</tr>")
              $fbBoard = $fbBoard + ("</table>")

			  
#$frow = $true 
#if ($frow -eq $true){
		
		$fbBoard = $fbBoard + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$fbBoard = $fbBoard + "<td align=`"center`" width=`"150`" ><b>Employee</b></td>" + "`r`n"
		for($stime = $displayStartTime;$stime -lt $displayEndTime;$stime = $stime.AddMinutes(30)){
			$fbBoard = $fbBoard + "<td align=`"center`" width=`"20`" ><b>" + $stime.ToString("HH:mm") + "</b></td>" +"`r`n"
		}
		$fbBoard = $fbBoard + "</tr>" + "`r`n"
	#	$frow = $false
#	}
$counter = 0

if ($mbHash.Count -ne 0){
	$mbHash.GetEnumerator() | Sort Value | foreach-object {
		
			$Attendee1 = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($_.Key)
			if ($Attendee1) {
			$Attendeesbatch.add($Attendee1)
			$availresponse = $service.GetUserAvailability($Attendeesbatch,$drDuration,[Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusy,$AvailabilityOptions)
$usrIdx = 0
$counter++
foreach($res in $availresponse.AttendeesAvailability){
      if ($counter -ge 50) {
	  $fbBoard = $fbBoard + "</table>"  + "  "
	  $fbBoard = $fbBoard + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$fbBoard = $fbBoard + "<td align=`"center`" width=`"150`" ><b>Employee</b></td>" + "`r`n"
		for($stime = $displayStartTime;$stime -lt $displayEndTime;$stime = $stime.AddMinutes(30)){
			$fbBoard = $fbBoard + "<td align=`"center`" width=`"20`" ><b>" + $stime.ToString("HH:mm") + "</b></td>" +"`r`n"
		}
		$fbBoard = $fbBoard + "</tr>" + "`r`n"
		$counter = 0
		}
	  $oofFlag = 0
	for($stime = $displayStartTime;$stime -lt $displayEndTime;$stime = $stime.AddMinutes(30)){
		if ($stime -eq $displayStartTime){
			$fbUser = "<td bgcolor=`"#FFFFFF`"><b>" + $mbHash[$Attendeesbatch[$usrIdx].SmtpAddress] + "</b></td>"  + "`r`n"
		
		}
		$title = "title="
		if ($res.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] -ne $null){
			$gdet = $false
			$FbValu = $res.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]]
		#$bgColour = "bgcolor=`"#FFFFFF`""
			switch($FbValu.ToString()){
				"Free" {$bgColour = "bgcolor=`"#6BBA70`""}
				"Tentative" {$bgColour = "bgcolor=`"#CDEB8B`""
					     $gdet = $true
					}
				"Busy" {$bgColour = "bgcolor=`"#356AA0`""
					     $gdet = $true
					}
				"OOF" {$bgColour = "bgcolor=`"#330099`""
					     $gdet = $true
						 $oofFlag = 1 
						 
						 
					}
				#"NoData" {$bgColour = "bgcolor=`"#98AFC7`""}
				#		"N/A" {$bgColour = "bgcolor=`"#98AFC7`""}		
			
		}
			if ($gdet -eq $true){
				if ($res.CalendarEvents.Count -ne 0){
					for($ci=0;$ci -lt $res.CalendarEvents.Count;$ci++){
						if ($res.CalendarEvents[$ci].StartTime -ge $stime -band $stime -le $res.CalendarEvents[$ci].EndTime){				
							if($res.CalendarEvents[$ci].Details.IsPrivate -eq $False){
								$subject = ""
								$location = ""
								if ($res.CalendarEvents[$ci].Details.Subject -ne $null){
									$subject = $res.CalendarEvents[$ci].Details.Subject.ToString()
								}
								if ($res.CalendarEvents[$ci].Details.Location -ne $null){
									$location = $res.CalendarEvents[$ci].Details.Location.ToString()
								}
								$title = $title + "`"" + $subject + " " + $location + "`" "
							}
						}
					}
				}
			}
			
		}
		else{
			$bgColour = "bgcolor=`"#98AFC7`""
		}
		if($title -ne "title="){
			$fbUser = $fbUser + "<td " + $bgColour + " " + $title + "></td>"  + "`r`n"
		}
		else{
			$fbUser = $fbUser + "<td " + $bgColour + "></td>"  + "`r`n"
		}

	}
	$fbUser = $fbUser + "</tr>"  + "`r`n"
	if ($oofFlag -eq 1 ) { 
	  $fbBoard = $fbBoard + $fbUser
	  }
	  else
	  {
	  $counter--
	  }
	  
	$usrIdx++
}
			$Attendeesbatch = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.AttendeeInfo]))
			
			}
			
	}
} 


$fbBoard = $fbBoard + "</table>"  + "  " 
 
$fbBoard = $fbBoard + "<table>"
 
    $fbBoard = $fbBoard + "<tr valign='top'>"
   
             $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#CDEB8B'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td align=center>&nbsp;Tentative</td>")
               $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#6BBA70'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Free</th>")
             $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#356AA0'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Busy</td>")
               $fbBoard = $fbBoard + ("<td width='5%' bgcolor='#330099'>&nbsp</td>")
              $fbBoard = $fbBoard + ("<td>Out of Office</td>")
             $fbBoard = $fbBoard + ("</tr>")
              $fbBoard = $fbBoard + ("</table>")
 
$fbBoard = $fbBoard + "<table>"
 
    $fbBoard = $fbBoard + "<tr valign='top'>"
		 
            		 
              $fbBoard = $fbBoard + ("<td><A href=`"lsl.htm`">LSL</A></td>")
        $fbBoard = $fbBoard + ("<td> | </td>")
              $fbBoard = $fbBoard + ("<td><A href=`"mandan.htm`">Mandan</A></td>")
              $fbBoard = $fbBoard + ("<td> | </td>")  
              $fbBoard = $fbBoard + ("<td><A href=`"shawano.htm`">Shawano</A></td>")
                $fbBoard = $fbBoard + ("<td> | </td>")
              $fbBoard = $fbBoard + ("<td><A href=`"cr.htm`">Cedar Rapids</A></td>")
			       $fbBoard = $fbBoard + ("<td> | </td>")
			   $fbBoard = $fbBoard + ("<td><A href=`"http://community.nisc.coop`">Community</A></td>")
             $fbBoard = $fbBoard + ("</tr>")
              $fbBoard = $fbBoard + ("</table>")

   $fbBoard = $fbBoard + "</body></html>"
			  

$fbBoard | out-file "C:\Freebusy\html\cr.htm" -encoding ASCII