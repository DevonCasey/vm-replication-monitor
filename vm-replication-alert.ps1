# Config files and sourcing of account information.
. ./source.ps1
$SMTPName = "smtp.gmail.com"
$EmailMessage = new-object Net.Mail.MailMessage
$SMTPServer = new-object Net.Mail.SmtpClient($SMTPName, "587")
$SMTPServer.EnableSsl = $true
$SMTPServer.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
$EmailMessage.From = $FromEmail
$EmailMessage.To.Add($ToEmail01)

#Build a nice file name
$date = Get-Date -Format M_d_yyyy_hh_mm_ss
$csvfile = ".\logs\Replication_Issues_" + $date + ".csv"
#Build the header row for the CSV file
$csv = "VM Name, Date, Server, Message `r`n"
#Find all problem VMs  
$VMList = Get-VM | Where-Object { $_.ReplicationHealth -eq "Critical" -or $_.ReplicationHealth -eq "Warning" }
#Loop through each VM to get the corresponding events
foreach ($VM in $VMList) {
    $VMReplStats = $VM | Measure-VMReplication
    $FromDate = $VMReplStats.LastReplicationTime 
    #This string will filter for events for the current VM only
    $FilterString = "<QueryList><Query Id='0' Path='Microsoft-Windows-Hyper-V-VMMS-Admin'><Select Path='Microsoft-Windows-Hyper-V-VMMS-Admin'>*[UserData[VmlEventLog[(VmId='" + $VM.ID + "')]]]</Select></Query></QueryList>" 
    $EventList = Get-WinEvent -FilterXML $FilterString  | Where-Object { $_.TimeCreated -ge $FromDate -and $_.LevelDisplayName -eq "Error" } | Select-Object -Last 3
    #Dump relevant information to the CSV file  
    foreach ($Event in $EventList) {
            if ($VM.ReplicationMode -eq "Primary") {
                    $Server = $VMReplStats.PrimaryServerName
                } else {
                    $Server = $VMReplStats.ReplicaServerName
                }
            $csv += $VM.Name + "," + $Event.TimeCreated + "," + $Server + "," + $Event.Message + "`r`n"
        }
} 
#Create a file and dump all information in CSV format
$fso = New-Object -comobject scripting.filesystemobject
$file = $fso.CreateTextFile($csvfile, $true)
$file.write($csv)
$file.close()  
#If there are VMs in critical health state, send an email
if ($VMList -and $csv.Length -gt 33) { 
    $Attachment = New-Object Net.Mail.Attachment($csvfile)
    $EmailMessage.Subject = "[ATTENTION] Replication requires your attention!"    
    $EmailMessage.Body = "The report is attached."
    $EmailMessage.Attachments.Add($Attachment)
    $SMTPServer.Send($EmailMessage)
    $Attachment.Dispose()
} else {
    $EmailMessage.Subject = "[NORMAL] All VMs replicating Normally!"  
    $EmailMessage.Body = "All VMs are replicating normally. No further action is required at this point."
    $SMTPServer.Send($EmailMessage)
}