<#

DISCLAIMER
        This Sample Code is provided for the purpose of illustration only and is not intended to be 
        used in a production environment.  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED 
        "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED 
        TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant 
        You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and 
        distribute the object code form of the Sample Code, provided that You agree: (i) to not use 
        Our name, logo, or trademarks to market Your software product in which the Sample Code is 
        embedded; (ii) to include a valid copyright notice on Your software product in which the 
        Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our 
        suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise 
        or result from the use or distribution of the Sample Code.

#>

# Variables
$scorchDBName = "Orchestrator" # Orchestrator Database - Called Orchestrator by default
$scorchDBServer = "AUS-SCORCH01" # Orchestrator Database Server

# Main
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

#Loop and process each running Runbook
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
        $currentSequenceNumber = $table | Where {$_.Status -notmatch "^[sf]"} | Sort-Object -Property SequenceNumber -Descending | Select -First 1 -ExpandProperty SequenceNumber
        $lastStatus = $table | Where {$_.SequenceNumber -match ($currentSequenceNumber-1) } | Select -ExpandProperty Status
        $instance.Add("LastActivityStatus",$lastStatus)
        $instance.Add("ActivityName",($table | Sort-Object -Property SequenceNumber -Descending | Select -First 1 -ExpandProperty Name))
        $instance.Add("PercentComplete","$([math]::round($((($table | Sort-Object -Property SequenceNumber -Descending | Select -First 1 -ExpandProperty SequenceNumber)/$instance.ActivityCount)*100))) %")
        $obj = New-Object -TypeName PSObject -Property $instance
        $jobArray += $obj
        $obj = $null
        Remove-Variable -Name instance -Force
        Remove-Variable -Name table -Force

    }

# Print out results
$jobArray | ogv
