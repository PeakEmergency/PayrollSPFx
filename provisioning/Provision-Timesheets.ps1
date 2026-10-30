# Creates the Timesheet Submissions list and the Timesheet Templates library, and uploads the templates.
# Safe to rerun: anything that already exists is left alone.
param(
    [Parameter(Mandatory)] [string]$SiteUrl,
    [Parameter(Mandatory)] [string]$AccountsGroup,
    [string]$ListTitle = 'Timesheet Submissions',
    [string]$LibraryTitle = 'Timesheet Templates',
    # Entra app registration for PnP interactive login (required by PnP.PowerShell 2+).
    [string]$ClientId
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if ($ClientId) { Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId }
else { Connect-PnPOnline -Url $SiteUrl -Interactive }

# --- Submissions list ---
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if ($list) { Write-Host "List '$ListTitle' exists" }
else {
    $list = New-PnPList -Title $ListTitle -Template GenericList -EnableVersioning
    Write-Host "Created list '$ListTitle'"
}
Set-PnPList -Identity $ListTitle -EnableAttachments $true | Out-Null

$fields = @(
    @{ Name = 'TimesheetType'; Xml = '<Field Type="Choice" Name="TimesheetType" StaticName="TimesheetType" DisplayName="Timesheet Type" Required="TRUE"><CHOICES><CHOICE>Admin</CHOICE><CHOICE>Instructor</CHOICE><CHOICE>Medic</CHOICE></CHOICES></Field>' },
    @{ Name = 'Employee'; Xml = '<Field Type="User" Name="Employee" StaticName="Employee" DisplayName="Employee" UserSelectionMode="PeopleOnly" Required="TRUE" />' },
    @{ Name = 'PayPeriodStart'; Xml = '<Field Type="DateTime" Name="PayPeriodStart" StaticName="PayPeriodStart" DisplayName="Pay Period Start" Format="DateOnly" />' },
    @{ Name = 'PayPeriodEnd'; Xml = '<Field Type="DateTime" Name="PayPeriodEnd" StaticName="PayPeriodEnd" DisplayName="Pay Period End" Format="DateOnly" />' },
    @{ Name = 'TotalAmount'; Xml = '<Field Type="Currency" Name="TotalAmount" StaticName="TotalAmount" DisplayName="Total Amount" LCID="1033" Decimals="2" />' },
    @{ Name = 'LineItems'; Xml = '<Field Type="Note" Name="LineItems" StaticName="LineItems" DisplayName="Line Items" RichText="FALSE" NumLines="6" UnlimitedLengthInDocumentLibrary="TRUE" />' },
    @{ Name = 'Status'; Xml = '<Field Type="Choice" Name="Status" StaticName="Status" DisplayName="Status"><Default>Pending</Default><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Sent</CHOICE><CHOICE>Email failed</CHOICE></CHOICES></Field>' },
    @{ Name = 'StatusDetail'; Xml = '<Field Type="Note" Name="StatusDetail" StaticName="StatusDetail" DisplayName="Status Detail" RichText="FALSE" NumLines="3" />' }
)
foreach ($f in $fields) {
    if (Get-PnPField -List $ListTitle -Identity $f.Name -ErrorAction SilentlyContinue) { Write-Host "  field $($f.Name) exists" }
    else {
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $f.Xml | Out-Null
        Write-Host "  added field $($f.Name)"
    }
}

Set-PnPView -List $ListTitle -Identity 'All Items' -Fields (@('LinkTitle') + ($fields | ? Name -ne 'LineItems' | % Name)) | Out-Null

# --- Item-level security: users read and edit only their own items ---
Set-PnPList -Identity $ListTitle -ReadSecurity 2 -WriteSecurity 2 | Out-Null

# Item-level security is bypassed by anyone with Manage Lists (the Edit role).
# Unique permissions: accounts get Edit (see all items), site members drop to Contribute (own items only).
$list = Get-PnPList -Identity $ListTitle -Includes HasUniqueRoleAssignments
if (-not $list.HasUniqueRoleAssignments) {
    Set-PnPList -Identity $ListTitle -BreakRoleInheritance -CopyRoleAssignments | Out-Null
    $members = Get-PnPGroup -AssociatedMemberGroup
    Set-PnPListPermission -Identity $ListTitle -Group $members.Title -RemoveRole 'Edit' -AddRole 'Contribute'
    Write-Host "List permissions: '$($members.Title)' set to Contribute"
}
Set-PnPListPermission -Identity $ListTitle -Group $AccountsGroup -AddRole 'Edit'
Write-Host "List permissions: '$AccountsGroup' has Edit"

# --- Template library ---
$lib = Get-PnPList -Identity $LibraryTitle -ErrorAction SilentlyContinue
if ($lib) { Write-Host "Library '$LibraryTitle' exists" }
else {
    $lib = New-PnPList -Title $LibraryTitle -Template DocumentLibrary -EnableVersioning
    Write-Host "Created library '$LibraryTitle'"
}
$lib = Get-PnPList -Identity $LibraryTitle -Includes RootFolder
$folder = $lib.RootFolder.ServerRelativeUrl
foreach ($file in Get-ChildItem (Join-Path $root 'Templates\xlsx\*.xlsx')) {
    if (Get-PnPFile -Url "$folder/$($file.Name)" -ErrorAction SilentlyContinue) { Write-Host "  $($file.Name) exists" }
    else {
        Add-PnPFile -Path $file.FullName -Folder $folder | Out-Null
        Write-Host "  uploaded $($file.Name)"
    }
}
Write-Host "Template library URL for the web part: $folder"
