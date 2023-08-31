# Creation of passwords
# 1. Open a PowerShell
# 2. Enter the following line and replace "Your password" with your real password
#      "Your password" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
# 3. Copy the given string with numbers in the right field below
# 4. Save the file

##### Begin of area to change ##### 

# Set values for SQL query
$SQLServer = "DATABASEPC"
$SQLDatabase = "ZapDatabase"
$SQLUsername = "ZapQuery"
$SQLPassword = "01000000d08c9ddf0115d1118c7a00c04fc297eb0100000071acaec3a56d644a922f95d62804cc3100000000020000000000106600000001000020000000c4880af0052ec895b0ddec04c5f53a5efe693e54386eaa08fa092660882a1c9f000000000e8000000002000020000000b09de674183e026f6fc314129db6d3e9cce927d32aaee216819ba0bbdc38505420000000239d97e60397e0e13457a0066dedec5d59de402782490c96818df79d1507ba6e400000003da84edcc1235c545d2dd8dcf3b2627da0402967e7089c7985af01f2eb892bb67abf536eb7595a35accdc9bade29a57794fcf8ea1f503070cd92ae0f650d8ab6"

# Set values for sending the mail
$SMTPUsername = "<SMTP server username"
$SMTPPassword = "<SMTP server password>"
$SMTPTo = @("<Any To mail address>, "<Another To mail address>")
$SMTPCc = @("<Any Cc mail address>")
$SMTPFrom = "<Sender mail address>"
$SMTPSite = "<Your mail sender name>"
$SMTPServer = "<Your SMTP server>"
$SMTPPort = 587

##### End of area to change #####

# Create year and month for the range
$CurrentDate = Get-Date -Hour 0 -Minute 0 -Second 0
$LastMonth = $CurrentDate.AddMonths(-1)
$FirstDayOfMonth = Get-Date $LastMonth -Day 1
$LastDayOfMonth = Get-Date $FirstDayOfMonth.AddMonths(1).AddSeconds(-1)

# Create SQL query
$SQLPassword = $SQLPassword | ConvertTo-SecureString
$SQLCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SQLUsername, $SQLPassword
$SQLQuery = "SELECT DISTINCT  FractionDate = CONVERT(VARCHAR(10), TreatmentTime, 111), MedicalId, PlanUuid 
          FROM [ZapDatabase].[dbo].[View_Plan] INNER JOIN [View_DeliveredBeam] ON View_Plan.PlanID = View_DeliveredBeam.PlanUuid 
            INNER JOIN [View_Patient] ON View_Plan.PatientID = View_Patient.Uuid 
          WHERE View_Plan.PatientType = 'Human' 
            AND (View_Plan.Status = '3' OR View_Plan.Status = '4') 
            AND [TreatmentTime] BETWEEN '" + $FirstDayOfMonth.ToString("yyyy/MM/dd") + " 00:00' AND '" + $LastDayOfMonth.ToString("yyyy/MM/dd") + " 23:59'  
            AND PlanName NOT LIKE '%Simulat%' 
            AND PlanName NOT LIKE '%DryRun%' 
            AND FractionID = 1 
          ORDER BY [MedicalId], [FractionDate]"

# Execute SQL query
$SQLResult = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Username $SQLUsername -Password $SQLCredential.GetNetworkCredential().Password -Query $SQLQuery | Out-String

# Create SQL query with full name 
$SQLQueryFull = "SELECT DISTINCT  FractionDate = CONVERT(VARCHAR(10), TreatmentTime, 111), MedicalId, LastName, FirstName, PlanName, PlanUuid 
          FROM [ZapDatabase].[dbo].[View_Plan] INNER JOIN [View_DeliveredBeam] ON View_Plan.PlanID = View_DeliveredBeam.PlanUuid 
            INNER JOIN [View_Patient] ON View_Plan.PatientID = View_Patient.Uuid 
          WHERE View_Plan.PatientType = 'Human' 
            AND (View_Plan.Status = '3' OR View_Plan.Status = '4') 
            AND [TreatmentTime] BETWEEN '" + $FirstDayOfMonth.ToString("yyyy/MM/dd") + " 00:00' AND '" + $LastDayOfMonth.ToString("yyyy/MM/dd") + " 23:59'  
            AND PlanName NOT LIKE '%Simulat%' 
            AND PlanName NOT LIKE '%DryRun%' 
            AND FractionID = 1 
          ORDER BY [MedicalId], [FractionDate]"

# Execute SQL query
$SQLResultFull = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Username $SQLUsername -Password $SQLCredential.GetNetworkCredential().Password -Query $SQLQueryFull | Out-String

# Count number of lines to get the sum of all patients with FrationID = 1 in the given date range
$NumberOfLines = $SQLResult | Measure-Object -Line
$NumberOfPatients = $NumberOfLines.Lines - 5 

# Create mail body
$MailBody = "Treated Patient Report for Month <b>" + $FirstDayOfMonth.ToString("yyyy-MM") + "</b><br>" #[Environment]::NewLine
$MailBody = $MailBody + "<span style='font-family:Courier;font-size:10pt;tabsize='16'>"
$MailBody = $MailBody + $SQLResult.Substring(0, $SQLResult.Length-1).Replace(" ", "&nbsp;").Replace([Environment]::NewLine, "<br>")
$MailBody = $MailBody + "</span>"
$MailBody = $MailBody + "Number of first treated patients: <b>" + $NumberOfPatients + "</b>"
$MailBody = $MailBody + "<span style='font-family:Courier;font-size:10pt;tabsize='16'>"
$MailBody = $MailBody + $SQLResultFull.Substring(0, $SQLResultFull.Length-1).Replace(" ", "&nbsp;").Replace([Environment]::NewLine, "<br>")
$MailBody = $MailBody + "</span>"

# Create SMTP properties
$SMTPPassword = $SMTPPassword | ConvertTo-SecureString
$SMTPCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername, $SMTPPassword
$SMTPProperties = @{
    To = $SMTPTo
    Cc = $SMTPCc
    From = $SMTPFrom
    Subject = $SMTPSite + ": Treated Patients Report for Month " + $FirstDayOfMonth.ToString("yyyy-MM")
    SMTPServer = $SMTPServer
    Port = $SMTPPort
    Credential = $SMTPCredential
    Encoding = "UTF8"
}

Send-MailMessage @SMTPProperties -Body $MailBody -BodyAsHTML
