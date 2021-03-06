<#
.SYNOPSIS
    Adds an OpenSSH public key to an authorized_keys file.  

.DESCRIPTION
    The script takes the path to an OpenSSH public key file, such
    as id_rsa.pub, and appends it to the ~\.ssh\authorized_keys file 
    for a user at a target server.  The ~\.ssh folder and authorized_keys 
    file will be created if either one does not exist.  

.PARAMETER PublicKey
    The path to the OpenSSH public key file to be appended. Defaults
    to $env:USERPROFILE\.ssh\id_rsa.pub if none is given.  If the same
    key is appended to the authorized_keys file multiple times, this
    does not prevent successful logon.  For security reasons, the name
    of the public key file must end with the ".pub" extension.

.PARAMETER TargetHost
    The name of the target computer where the authorized_keys file 
    is located.  Any UNC-compatible computer name is acceptable, such
    as a NetBIOS name or full DNS hostname.  The script assumes that the 
    user profile folders are located under C:\Users at the target host.
    If Windows is installed into a different drive letter, then the full
    UNC path to the Users folder at the target must be given, e.g., 
    \\SomeServer\E$\Users.  Defaults to \\$env:COMPUTERNAME\C$\Users,
    which is on the local host, which would be unusual in practice.  

.PARAMETER UserProfile
    The name of the intended profile folder name, which is not always
    identical to the username of the intended user.  These folders
    are often found under C:\Users, such as C:\Users\Administrator.
    The argument to this parameter must match the name of the folder.
    Defaults to the local profile folder name of the person running
    the script.  The script does not create the folder if that folder
    does not exist.  The supplied folder cannot be named "Public."

.PARAMETER Verbose
    Switch to display verbose status information.  Otherwise, the script
    runs silently when successful.  Defaults to silent.  

.NOTES
    First Created: 25.Nov.2019
    Last Updated: 25.Nov.2019
    Legal: 0BSD.
    Author: JF@Enclave, https://www.sans.org/SEC505
    TODO: Add support for explicit SMB creds to stand-alones.
    TODO: Decide whether to check for duplicates in existing authorized_keys.
    TODO: Reproduce more ssh-copy-id behavior, but with SMB.
    TODO: Clean up the $? hacks...
#>


Param ( $PublicKey = "$env:UserProfile\.ssh\id_rsa.pub",
        $TargetHost = $env:ComputerName,
        $UserProfile = ($env:USERPROFILE | Split-Path -Leaf),
        [Switch] $Verbose 
      ) 



# Test path to $PublicKey 
if (-not (Test-Path -Path $PublicKey))
{ Throw "ERROR: Cannot find $PublicKey" ; Exit } 

# Confirm .pub file name extension:
if ($PublicKey -notlike "*.pub")
{ Throw "ERROR: The public key file name extension must be .pub" ; Exit }

# Read $PublicKey or fail:
$ThePubKey = Get-Content -Path $PublicKey -ErrorAction Stop

# Confirm that the key is not a private key:
if ($ThePubKey[0] -like "*PRIVATE KEY*")
{ Throw "ERROR: You must select a public key, not a private key" ; Exit }

# Assume C$ at $TargetHost, unless explicit path given:
if ($TargetHost -notlike "\\*\Users*")
{ $TargetHost = "\\$TargetHost\C$\Users" } 

# Test path to \\$TargetHost\C$\Users:
if (-not (Test-Path -Path $TargetHost))
{ Throw "ERROR: Cannot find $TargetHost" ; Exit } 

# Confirm that the Public profile folder was not given:
if ($UserProfile -eq "Public")
{ Throw "ERROR: The user profile cannot be Public" ; Exit }

# Test path to $UserProfile at $TargetHost
$UserProfile = Join-Path -Path $TargetHost -ChildPath $UserProfile 
if (-not (Test-Path -Path $UserProfile))
{ 
    $ExistingProfiles = @(dir $TargetHost | Where { $_.Name -notlike "Public" } | Select-Object -ExpandProperty Name) -Join ","
    Throw "ERROR: Cannot find $UserProfile. The profile folder(s) found at $TargetHost are: $ExistingProfiles."
    Exit 
} 

# If necessary, create $UserProfile\.ssh\:
$SshFolder = Join-Path -Path $UserProfile -ChildPath ".ssh" 
if (-not (Test-Path -Path $SshFolder))
{ 
    New-Item -ItemType Directory -Path $UserProfile -Name ".ssh" | Out-Null
    if ($?)
    { if ($Verbose){ Write-Verbose -Verbose -Message "Folder created: $UserProfile\.ssh\" } }
    else
    { Throw "ERROR: Could not create $SshFolder" ; Exit } 
} 

# Append $ThePubKey to the authorized_keys file, creating it if necessary.
# This can result in duplicate public keys being appended, but this does not
# prevent successful ssh.exe authentication.  Also, SMB is plaintext by default
# and other admins may have taken care to avoid exposing authorized_keys in
# plaintext on the network, so there is a reason to not read that file here in
# order to check for duplicates (need to think more about this...).
$AuthorizedKeysFile = Join-Path -Path $SshFolder -ChildPath "authorized_keys" 

$ThePubKey | Out-File -Append -Encoding utf8 -Force -FilePath $AuthorizedKeysFile

if ($?)
{ if ($Verbose){ Write-Verbose -Verbose -Message "$PublicKey successfully added to $AuthorizedKeysFile" } }
else
{ Throw "ERROR: Could not append to $AuthorizedKeysFile" ; Exit } 

