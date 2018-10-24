#requires -version 2.0

#============================================================================================#
# Hoplite Industries, Inc.                                                                   #
# O365 Exchange Message Trace Export Tool                                                    #
# v0.1.0 --  October, 2018 acochenour, original release                                      #
#============================================================================================#

#============================================================================================#
# Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License.  You may obtain a copy of the License at;
# http://www.apache.org/licenses/LICENSE-2.0
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the License for the specific language governing permissions and limitations under the License.
#============================================================================================#

#============================================================================================#
# CONFIGURATION
#============================================================================================#
$o365User = '<O365_ADMIN_EMAIL_CHANGE_ME>'
$o365Pwd = '<O365_PASSWORD_CHANGE_ME>'
$uDir = "C:\Users\<UNAME_CHANGE_ME>\"
$lastXHours = -1

#============================================================================================#
# Connect to Office 365 using the O365 Exchange Admin API
#============================================================================================#
Write-Host "Connecting to O365 Exchange API as $o365User..."
$sPwd = $o365Pwd | ConvertTo-SecureString -AsPlainText -Force 
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $o365User, $sPwd 
$oSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $cred -Authentication Basic -AllowRedirection  
Import-PSSession $oSession -AllowClobber 

#============================================================================================#
# Collect Message Tracking Logs (These are broken into "pages" in Office 365 so we need to collect them all with a loop) 
#============================================================================================#
$Messages = $null 
$Page = 1 
do 
{ 
    Write-Host "Exporting O365 Message Trace Logs - Page $Page..." 
    $CurrMessages = Get-MessageTrace -StartDate (Get-Date).AddHours($lastXHours) -EndDate (Get-Date)  -PageSize 5000  -Page $Page | Select-Object Received,SenderAddress,RecipientAddress,Subject,Size,MessageId,MessageTraceId,FromIP,ProbeTag,Status,ToIP

    if ($CurrMessages -ne $null)
      {
          $fPrefix = "msgtrace-"
          $dString = get-date -f yyyy-MM-dd-hh
          $ext = ".csv"
          $CurrMessages | Export-Csv "$uDir$fPrefix$dString$ext" -NoTypeInformation
      }
    $Page++ 
    $Messages += $CurrMessages 


} 
until ($CurrMessages -eq $null) 

#=======================================================================================#
# Clean up the Exchange API session and exit
#=======================================================================================#
Remove-PSSession $oSession 
