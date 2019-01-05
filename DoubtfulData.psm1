function Export-DistributionGroupMember {
    [CmdletBinding()]
    param (
        # Specifies Distribution Group identity
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [string]
        $Identity,

        # Specifies path to use for storing exported data
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.IO.DirectoryInfo]
        $Path = "$PSScriptRoot\Export"
    )

    begin {
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $InvalidCharacters = "'`"!@#$%^&*()\s"
        $TimeStamp = { [datetime]::Now.ToString("MM/dd/yy hh:mm:ss tt") }
    }

    process {
        $Guid, $Name, $Tenant = & {
            $DistributionGroup = Get-DistributionGroup -Identity $Identity
            @(
                $DistributionGroup.Guid
                $DistributionGroup.Name
                ($DistributionGroup.DistinguishedName -split ",")[1].Substring(3)
            )
        }
        Write-Verbose -Message ("{0} [i] Processing distribution group: {1}" -f @(
                $TimeStamp.Invoke()
                $Name))
        $FilePath = "{0}\{1}\{2}\{3}.xml" -f @(
            $Path.FullName
            $Tenant
            $Guid
            $Name -replace "[$InvalidCharacters]"
        )

        # New-Item paramters
        $CmdParams = @{
            Path        = (Split-Path -Path $FilePath)
            ItemType    = "Directory"
            Force       = $true
            ErrorAction = "SilentlyContinue"
        }
        [void](New-Item @CmdParams)

        # Write-Progress parameters
        $CmdParams = @{
            Activity = "Exporting Distribution Group Members: $Name"
        }
        Write-Progress @CmdParams

        Get-DistributionGroupMember -Identity $Identity |
            Export-Clixml -Path $FilePath -Force
    }

    end {
        $Stopwatch.Stop()
        Write-Verbose -Message ("{0} [i] Process finished in: {1:N}s" -f @(
                $TimeStamp.Invoke()
                $Stopwatch.Elapsed.TotalSeconds))
    }
}


function Remove-MailContactFromDistributionGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        # Specifies Distribution Group identity
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [string]
        $Identity
    )

    begin {
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $TimeStamp = { [datetime]::Now.ToString("MM/dd/yy hh:mm:ss tt") }
    }

    process {
        $DistributionGroupName = (Get-DistributionGroup -Identity $Identity).Name
        Write-Verbose -Message ("{0} [i] Processing distribution group: {1}" -f @(
                $TimeStamp.Invoke()
                $DistributionGroupName))
        Get-DistributionGroupMember -Identity $Identity |
            Where-Object RecipientTypeDetails -EQ MailContact | ForEach-Object {
            if ($PSCmdlet.ShouldProcess(
                    $DistributionGroupName,
                    ("Remove distribution group member {0}" -f $_.Name))) {
                # Write-Progress parameters
                $CmdParams = @{
                    Activity = "Removing Distribution Group Members: $DistributionGroupName"
                }
                Write-Progress @CmdParams

                # Remove-DistributionGroupMember paramters
                $CmdParams = @{
                    Identity                        = $Identity
                    Member                          = $_.Guid.Guid
                    BypassSecurityGroupManagerCheck = $true
                    Confirm                         = $false
                    Verbose                         = $false
                }
                Remove-DistributionGroupMember @CmdParams
            }
        }
    }

    end {
        $Stopwatch.Stop()
        Write-Verbose -Message ("{0} [i] Process finished in: {1:N}s" -f @(
                $TimeStamp.Invoke()
                $Stopwatch.Elapsed.TotalSeconds))
    }
}


function Import-DistributionGroupMember {
    [CmdletBinding()]
    param (
        # Specifies Distribution Group identity
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [string]
        $Identity,

        # Specifies path to import data
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.IO.DirectoryInfo]
        $Path = "$PSScriptRoot\Export"
    )

    begin {
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $InvalidCharacters = "'`"!@#$%^&*()\s"
        $TimeStamp = { [datetime]::Now.ToString("MM/dd/yy hh:mm:ss tt") }
    }

    process {
        $Guid, $Name, $Tenant = & {
            $DistributionGroup = Get-DistributionGroup -Identity $Identity
            @(
                $DistributionGroup.Guid
                $DistributionGroup.Name
                ($DistributionGroup.DistinguishedName -split ",")[1].Substring(3)
            )
        }
        Write-Verbose -Message ("{0} [i] Processing distribution group: {1}" -f @(
                $TimeStamp.Invoke()
                $Name))

        $FilePath = "{0}\{1}\{2}\{3}.xml" -f @(
            $Path.FullName
            $Tenant
            $Guid
            $Name -replace "[$InvalidCharacters]"
        )
        Write-Verbose -Message ("{0} [i] Importing distribution group export file: {1}" -f @(
                $TimeStamp.Invoke()
                $FilePath))
        if (-not (Test-Path -Path $FilePath)) {
            # Write-Error paramters
            $CmdParams = @{
                Message  = "Distribution Group export file not found at: $FilePath"
                Category = "OpenError"
            }
            Write-Error @CmdParams
        }

        Import-Clixml -Path $FilePath |
            Where-Object { $_.RecipientTypeDetails -eq "MailContact" -and $_.PrimarySMTPAddress } |
            ForEach-Object {
            Write-Verbose -Message ("{0} [+] Adding distribution group member: '{1} <{2}>'" -f @(
                    $TimeStamp.Invoke()
                    $_.Name
                    $_.PrimarySMTPAddress))
            # Write-Progress parameters
            $CmdParams = @{
                Activity = "Importing Distribution Group Members: $Name"
            }
            Write-Progress @CmdParams

            # Add-DistributionGroupMember parameters
            $CmdParams = @{
                Identity                        = $Identity
                Member                          = $_.PrimarySMTPAddress
                BypassSecurityGroupManagerCheck = $true
                Verbose                         = $false
            }
            Add-DistributionGroupMember @CmdParams
        }
    }

    end {
        $Stopwatch.Stop()
        Write-Verbose -Message ("{0} [i] Process finished in: {1:N}s" -f @(
                $TimeStamp.Invoke()
                $Stopwatch.Elapsed.TotalSeconds))
    }
}


# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*
