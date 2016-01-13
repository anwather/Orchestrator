$scorchDBName = "Orchestrator"
$scorchDBServer = "AUS-SCORCH01"

$runningQuery = @"
Select RI.ID,RI.JobId,RI.RunbookId,RB.Name,RI.CreationTime
FROM [Microsoft.SystemCenter.Orchestrator.Runtime].[RunbookInstances] RI
INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Runbooks] RB ON RB.Id = RI.RunbookId
--INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Activities] RA ON RA.RunbookId = RI.RunbookId
WHERE RI.Status = 'InProgress'
"@




# Query In Progress Runbooks

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = "Server=$scorchDBServer; Database=$scorchDBName; Integrated Security = True"
$sqlcmd = New-Object System.Data.SqlClient.SqlCommand
$sqlcmd.CommandText = $runningQuery
$sqlcmd.Connection = $connection
$connection.Open()
$result = $sqlcmd.ExecuteReader()
$table = New-Object System.Data.DataTable
$table.Load($result)
#Remove-Variable -Name table -Force

$connection.Close()

$jobArray = @()

foreach ($line in $table)
    {
        $instance = @{}
        $runbookID = $line.JobId.Guid
        $instanceID = $line.ID.Guid
        $instance.Add("JobId",$runbookID)
        $instance.Add("RunbookName",$line.Name)
        $instance.Add("StartTime",$line.CreationTime)
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = "Server=$scorchDBServer; Database=$scorchDBName; Integrated Security = True"
        Remove-Variable -Name sqlcmd -Force

$activityCountQuery = @"
Select COUNT(*) AS ActivityCount
FROM [Microsoft.SystemCenter.Orchestrator.Runtime].[RunbookInstances] RI
INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Runbooks] RB ON RB.Id = RI.RunbookId
INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Activities] RA ON RA.RunbookId = RI.RunbookId
WHERE RI.JobId = '$runbookID'
"@

        $sqlcmd2 = New-Object System.Data.SqlClient.SqlCommand
        $sqlcmd2.CommandText = $activityCountQuery
        $sqlcmd2.Connection = $connection
        $connection.Open()
        $activityCountResult = $sqlcmd2.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($activityCountResult)
        
        $instance.Add("ActivityCount",$table.ActivityCount)
        $connection.Close()
        Remove-Variable -Name table -Force
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = "Server=$scorchDBServer; Database=$scorchDBName; Integrated Security = True"
        Remove-Variable -Name sqlcmd2 -Force
        $sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$activityQuery = @"
Select RAI.RunbookInstanceId,RAI.SequenceNumber,RAI.StartTime,RAI.Status, RA.Name 
FROM [Microsoft.SystemCenter.Orchestrator.Runtime].[ActivityInstances] RAI
INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Activities] RA ON RA.Id = RAI.ActivityId
WHERE RAI.RunbookInstanceId = '$instanceID'
"@
        $sqlcmd.CommandText = $activityQuery
        $sqlcmd.Connection = $connection
        $connection.Open()
        $activityResult = $sqlcmd.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($activityresult)
        
        $connection.Close()
        $instance.Add("ActivityName",($table | Sort-Object -Property SequenceNumber -Descending | Select -First 1 -ExpandProperty Name))
        $instance.Add("PercentComplete",(($table | Sort-Object -Property SequenceNumber -Descending | Select -First 1 -ExpandProperty SequenceNumber)/$instance.ActivityCount)*100)
        $obj = New-Object -TypeName PSObject -Property $instance
        $jobArray += $obj
        $obj = $null
        Remove-Variable -Name instance -Force
        Remove-Variable -Name table -Force

    }

$jobArray
