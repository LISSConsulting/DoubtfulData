# DoubtfulData
DoubtfulData is a PowerShell module built to aid in exporting Microsoft Exchange server distribution group members,
removing MailContact recipient type members from the distribution group, and importing distribution group members.


## Installation
Download release distribution from https://github.com/LISSConsulting/DoubtfulData/releases and extract the module archive to a location included in PSModulePath, e.g.

    ~\Documents\WindowsPowerShell\Modules


## Usage

Import the module into current PowerShell session:
```powershell
Import-Module DoubtfulData
```

This commands requires Exchange Online modules to be present in current PowerShell session. Before using this module, connect to Microsoft Exchange online.

### Export-DistributionGroupMember

#### Description
Use this command to export members of specified distribution group.

#### Parameters

**`-Identity`**

The Identity parameter specifies the distribution group or mail-enabled security group which members you want to export. You can use any value that uniquely identifies the group.

For example:
 - Name
 - Display name
 - Alias
 - Distinguished name (DN)
 - Canonical DN
 - Email address
 - GUID

This parameter accepts pipline input

**`-Path`**

Specifies the path to the directory where the exported data will be stored.

Default value: $PSScriptRoot\Export

#### Examples

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\> Get-DistributionGroup "Bugs" | Export-DistributionGroupMember


    -------------------------- EXAMPLE 2 --------------------------

    PS C:\> Export-DistributionGroupMember -Identity "Bugs"


    -------------------------- EXAMPLE 3 --------------------------

    PS C:\> Export-DistributionGroupMember -Identity "bugs@acme.org"


### Remove-MailContactFromDistributionGroup

#### Description
Use this command to remove **MailContact** recipient type members from specified distribution group.

#### Parameters

**`-Identity`**

The Identity parameter specifies the distribution group or mail-enabled security group from which you want to remove MailContact recipient type members. You can use any value that uniquely identifies the group.

For example:
 - Name
 - Display name
 - Alias
 - Distinguished name (DN)
 - Canonical DN
 - Email address
 - GUID

This parameter accepts pipline input

#### Examples

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\> Get-DistributionGroup "Bugs" | Remove-MailContactFromDistributionGroup -Verbose -Confirm:$false

    VERBOSE: 01/04/19 07:02:13 PM [i] Processing distribution group: Bugs
    VERBOSE: Performing the operation "Remove distribution group member Bunny" on target "Bugs".
    VERBOSE: Performing the operation "Remove distribution group member John Cussack" on target "Bugs".
    VERBOSE: Performing the operation "Remove distribution group member Katie Perry" on target "Bugs".
    VERBOSE: Performing the operation "Remove distribution group member Laura Croft " on target "Bugs".
    VERBOSE: 01/04/19 07:02:40 PM [i] Process finished in: 26.27s


    -------------------------- EXAMPLE 2 --------------------------

    PS C:\> Remove-MailContactFromDistributionGroup -Identity "Bugs"

    Confirm
    Are you sure you want to perform this action?
    Performing the operation "Remove distribution group member Bunny" on target "Bugs".
    [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):

    -------------------------- EXAMPLE 3 --------------------------

    PS C:\> Remove-MailContactFromDistributionGroup -Identity "bugs@acme.org" -Confirm:$false

### Import-DistributionGroupMember

#### Description
Use this command to import distribution group members exported with `Export-DistributionGroupMember` into specified distribution group. This command reads distribution group members from export file generated by `Export-DistributionGroupMember`

#### Parameters

**`-Identity`**

The Identity parameter specifies the distribution group or mail-enabled security group into which you want to import members. You can use any value that uniquely identifies the group.

For example:
 - Name
 - Display name
 - Alias
 - Distinguished name (DN)
 - Canonical DN
 - Email address
 - GUID

This parameter accepts pipline input

#### Examples

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\> Get-DistributionGroup "Bugs" | Import-DistributionGroupMember -Verbose

    VERBOSE: 01/04/19 07:04:10 PM [i] Processing distribution group: Bugs
    VERBOSE: 01/04/19 07:04:10 PM [i] Importing distribution group export file: C:\Users\Bunny\Documents\WindowsPowerShell\Modules\DoubtfulData\Export\acme.onmicrosoft.com\6f7f77ae-d6d4-4b03-8551-828c20af2b1d\Bugs.xml
    VERBOSE: 01/04/19 07:04:10 PM [+] Adding distribution group member: 'Bunny <bunny@acme.org>'
    VERBOSE: 01/04/19 07:04:10 PM [+] Adding distribution group member: 'John Cussack <JohnCussack@acme.org>'
    VERBOSE: 01/04/19 07:04:11 PM [+] Adding distribution group member: 'Katie Perry <KatiePerry@acme.org>'
    VERBOSE: 01/04/19 07:04:12 PM [+] Adding distribution group member: 'Laura Croft <LauraCroft@acme.org>'
    VERBOSE: 01/04/19 07:04:34 PM [i] Process finished in: 22.71s


    -------------------------- EXAMPLE 2 --------------------------

    PS C:\> Import-DistributionGroupMember -Identity "Bugs"


    -------------------------- EXAMPLE 3 --------------------------

    PS C:\> Import-DistributionGroupMember -Identity "bugs@acme.org"
