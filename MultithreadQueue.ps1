class MultithreadQueue {
    [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool = $null
    [System.Collections.ArrayList]$Jobs = $null

    MultithreadQueue([int]$ThreadCount, [HashTable]$SessionVariables) {
        [int]$NumThreads = $env:NUMBER_OF_PROCESSORS

        if ($ThreadCount -ge 1) {
            $NumThreads = $ThreadCount
        }

        $SessionState = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        foreach ($v in $SessionVariables.GetEnumerator()) {
            $SessionState.Variables.Add((New-Object Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $v.Key, $v.Value, ""))
        }

        $this.RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $NumThreads, $SessionState, $global:Host)
        $this.RunspacePool.ApartmentState = "MTA"
        $this.RunspacePool.Open()
        $this.Jobs = New-Object System.Collections.ArrayList
    }

    [void]AddJob([ScriptBlock]$JobScriptBlock, [Array]$Parameters) {
        $Job = [PowerShell]::Create()
        [void]$Job.AddScript($JobScriptBlock)

        foreach ($p in $Parameters) {
            [void]$Job.AddArgument($p)
        }

        $Job.RunspacePool = $this.RunspacePool

        [void]$this.Jobs.Add([PSCustomObject]@{
            Pipe = $Job
            Status = $Job.BeginInvoke()
        })
    }

    [void]AwaitJobs() {
        # loop while there are incomplete jobs
        while ($this.Jobs.Count -gt 0) {
            # detect completed jobs
            $CompletedJobs = $this.Jobs | Where-Object { $_.Status.IsCompleted }

            foreach ($Job in $CompletedJobs) {
                [void]$Job.Pipe.EndInvoke($Job.Status)
                $this.Jobs.Remove($Job)
            }
        }
    }

    [void]Close() {
        $IncompleteJobs = $this.Jobs | Where-Object { -not $_.Status.IsCompleted }

        foreach ($Job in $IncompleteJobs) {
            [void]$Job.Pipe.EndInvoke($Job.Status)
        }

        $this.Jobs.Clear()
        $this.RunspacePool.Close()
        $this.RunspacePool.Dispose()
    }
}

function Get-JobID {
    param(
        [int]$Count,
        [int]$Index
    )

    [int]$m = $Count.ToString().Length
    [int]$PadAmount = $m - $Index.ToString().Length
    return "$("0" * $PadAmount)$Index"
}
