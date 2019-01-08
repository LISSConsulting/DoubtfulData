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
        [IO.DirectoryInfo]
        $Path = "$PSScriptRoot\Export"
    )

    begin {
        $Stopwatch = [Diagnostics.Stopwatch]::StartNew()
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
                &$TimeStamp
                $Name))

        [IO.FileInfo]$FilePath = "{0}\{1}\{2}\{3}.xml" -f @(
            $Path.FullName
            $Tenant
            $Guid
            $Name -replace "[$InvalidCharacters]"
        )

        # New-Item paramters
        $CmdParams = @{
            Path     = ($FilePath | Split-Path)
            ItemType = "Directory"
            Force    = $true
        }
        [void](New-Item @CmdParams)

        # Write-Progress parameters
        $CmdParams = @{
            Activity = "Exporting Distribution Group Members: $Name"
        }
        Write-Progress @CmdParams

        Write-Verbose -Message ("{0} [i] Exporting distribution group members file: {1}" -f @(
                &$TimeStamp
                $FilePath.Name))

        Get-DistributionGroupMember -Identity $Identity |
            Export-Clixml -Path $FilePath -Force
    }

    end {
        $Stopwatch.Stop()
        Write-Verbose -Message ("{0} [i] Process finished in: {1:N}s" -f @(
                &$TimeStamp
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
        $Stopwatch = [Diagnostics.Stopwatch]::StartNew()
        $TimeStamp = { [datetime]::Now.ToString("MM/dd/yy hh:mm:ss tt") }
    }

    process {
        $DGName = (Get-DistributionGroup -Identity $Identity).Name

        Write-Verbose -Message ("{0} [i] Processing distribution group: {1}" -f @(
                &$TimeStamp
                $DGName))

        Get-DistributionGroupMember -Identity $Identity |
            Where-Object RecipientTypeDetails -EQ MailContact |
            ForEach-Object {
            if ($PSCmdlet.ShouldProcess(
                    $DGName, ("Remove distribution group member {0}" -f $_.Name))) {
                # Write-Progress parameters
                $CmdParams = @{
                    Activity = "Removing Distribution Group Members: $DGName"
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
                &$TimeStamp
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
        [IO.DirectoryInfo]
        $Path = "$PSScriptRoot\Export"
    )

    begin {
        $Stopwatch = [Diagnostics.Stopwatch]::StartNew()
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
                &$TimeStamp
                $Name))

        $FilePath = "{0}\{1}\{2}\{3}.xml" -f @(
            $Path.FullName
            $Tenant
            $Guid
            $Name -replace "[$InvalidCharacters]"
        )

        Write-Verbose -Message ("{0} [i] Importing distribution group export file: {1}" -f @(
                &$TimeStamp
                $FilePath))

        Import-Clixml -Path $FilePath |
            Where-Object $_.RecipientTypeDetails -EQ MailContact |
            ForEach-Object {
                if ($_.PrimarySMTPAddress) {
                    $PrimarySMTPAddress = $_.PrimarySMTPAddress
                } else {
                    $PrimarySMTPAddress = $_.ExternalEmailAddress -replace "SMTP:"
                }

                Write-Verbose -Message ("{0} [+] Adding distribution group member: '{1} <{2}>'" -f @(
                        &$TimeStamp
                        $_.Name
                        $PrimarySMTPAddress))

                # Write-Progress parameters
                $CmdParams = @{
                    Activity = "Importing Distribution Group Members: $Name"
                }
                Write-Progress @CmdParams

                # Add-DistributionGroupMember parameters
                $CmdParams = @{
                    Identity                        = $Identity
                    Member                          = $PrimarySMTPAddress
                    BypassSecurityGroupManagerCheck = $true
                    Verbose                         = $false
                }
                Add-DistributionGroupMember @CmdParams
            }
    }

    end {
        $Stopwatch.Stop()
        Write-Verbose -Message ("{0} [i] Process finished in: {1:N}s" -f @(
                &$TimeStamp
                $Stopwatch.Elapsed.TotalSeconds))
    }
}


# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*
