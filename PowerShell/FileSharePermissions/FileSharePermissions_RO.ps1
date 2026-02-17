# =============================================
# Script: Grant-ReadOnly-Access.ps1
# Purpose: Grant Read & Execute permissions to
#          specified users on E:\PA$ and all
#          subfolders and files
# =============================================

# Define the target folder
$FolderPath = "D:\Folder$"

# Define users (Use DOMAIN\Username format if domain accounts)
$Users = @(
    "domain\Howard",
    "domain\The",
    "domain\Duck"
)

# Loop through each user and apply permissions
foreach ($User in $Users) {

    Write-Host "Processing $User..." -ForegroundColor Cyan

    # Get current ACL
    $Acl = Get-Acl $FolderPath

    # Define the permission rule
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        $User,
        "ReadAndExecute",
        "ContainerInherit,ObjectInherit",
        "None",
        "Allow"
    )

    # Add the access rule
    $Acl.AddAccessRule($AccessRule)

    # Apply updated ACL to the folder
    Set-Acl -Path $FolderPath -AclObject $Acl

    Write-Host "Read-only access granted to $User" -ForegroundColor Green
}

Write-Host "Permissions successfully applied to all users." -ForegroundColor Yellow
