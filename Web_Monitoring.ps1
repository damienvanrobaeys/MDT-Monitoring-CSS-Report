$Date = get-date
$programdata = $env:programdata
$CSS_File = "$programdata\Monitoring_CSS.css" 
$Your_Host = ""

$URL = "http://" + $Your_Host + ":9801/MDTMonitorData/Computers/"
$HTML_Deployment_List = "$programdata\Monitoring_List.htm"
$Title = "<br><p><span class=titre_list>MDT deployment Status</span><br><span class=subtitle>This document has been updated on $Date</span></p><br>"

function GetMDTData { 
  $Data = Invoke-RestMethod $URL
  foreach($property in ($Data.content.properties)) 
  { 
		$Percent = $property.PercentComplete.'#text' 		
		$Current_Steps = $property.CurrentStep.'#text'			
		$Total_Steps = $property.TotalSteps.'#text'		
		
		If ($Current_Steps -eq $Total_Steps)
			{
				If ($Percent -eq $null)
					{			
						$Step_Status = "Not started"
					}
				Else
					{
						$Step_Status = "$Current_Steps / $Total_Steps"
					}					
			}
		Else
			{
				$Step_Status = "$Current_Steps / $Total_Steps"			
			}

	
		$Step_Name = $property.StepName		
		If ($Percent -eq 100)
			{
				$Global:StepName = "Deployment finished"
				$Percent_Value = $Percent + "%"				
			}
		Else
			{
				If ($Step_Name -eq "")
					{					
						If ($Percent -gt 0) 					
							{
								$Global:StepName = "Computer restarted"
								$Percent_Value = $Percent + "%"
							}	
						Else							
							{
								$Global:StepName = "Deployment not started"	
								$Percent_Value = "Not started"	
							}

					}
				Else
					{
						$Global:StepName = $property.StepName		
						$Percent_Value = $Percent + "%"					
					}					
			}

		$Deploy_Status = $property.DeploymentStatus.'#text'					
		If (($Percent -eq 100) -and ($Step_Name -eq "") -and ($Deploy_Status -eq 1))
			{
				$Global:StepName = "Running in PE"						
			}			
			
			
		$End_Time = $property.EndTime.'#text' 	
		If ($End_Time -eq $null)
			{
				If ($Percent -eq $null)
					{									
						$EndTime = "Not started"
						$Ellapsed = "Not started"												
					}
				Else
					{
						$EndTime = "Not finished"
						$Ellapsed = "Not finished"					
					}
			}
		Else
			{
				$EndTime = ([datetime]$($property.EndTime.'#text')).ToLocalTime().ToString('HH:mm:ss')  	 
				$Ellapsed = new-timespan -start ([datetime]$($property.starttime.'#text')).ToString('HH:mm:ss') -end ([datetime]$($property.endTime.'#text')).ToString('HH:mm:ss'); 				
			}

    New-Object PSObject -Property @{ 
      "Computer Name" = $($property.Name); 
      "Percent Complete" = $Percent_Value; 	  
      "Step Name" = $StepName;	  	  
      "Step status" = $Step_Status;	  
      Warnings = $($property.Warnings.'#text'); 
      Errors = $($property.Errors.'#text'); 
      "Deployment Status" = $( 
        Switch ($property.DeploymentStatus.'#text') { 
        1 { "Running" } 
        2 { "Failed" } 
        3 { "Success" } 
        4 { "Unresponsive" } 		
        Default { "Unknown" } 
        } 
      ); 	  
      "Date" = $($property.StartTime.'#text').split("T")[0]; 
      "Start time" = ([datetime]$($property.StartTime.'#text')).ToLocalTime().ToString('HH:mm:ss')  
	  "End time" = $EndTime;
      "Ellapsed time" = $Ellapsed;	  	  
    } 
  } 
} 

$MyData = GetMDTData | Select Date, "Computer Name", "Percent Complete", "Step Name", Warnings, Errors, "Start time", "End Time", "Ellapsed time", "Step status", "Deployment Status" | Sort -Property StartTime | ConvertTo-HTML -Fragment
$colorTagTable = @{Failed = ' class="Failed">Failed<';
				   Unknown = ' class="Failed">Unknown<';
				   Unresponsive = ' class="Failed">Unresponsive<';				   
				   Running = ' class="Running">Running<';
				   Success = ' class="Success">Success<'}
$colorTagTable.Keys | foreach { $MyData = $MyData -replace ">$_<",($colorTagTable.$_) }
$html_final = ConvertTo-HTML  -body "$title<br>$MyData" -CSSUri $CSS_File 		
$html_final | out-file -encoding ASCII $HTML_Deployment_List
Invoke-Item $HTML_Deployment_List



