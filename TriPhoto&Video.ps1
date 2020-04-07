$ErrorActionPreference = 'SilentlyContinue'

# Script de tri, nommage et encodage des photos/vidéos automatique
# @Aohzan

# Paramètres
$CheminScan = $((New-Object -ComObject "Shell.Application" ).BrowseForFolder(0,"Choisir le dossier :",0)).Self.Path
$FormatNomFichier = "yyyy-MM-dd-HH\hmm\mss"
$HandBrakeCliPath = "C:\Program Files\Handbrake\HandBrakeCLI.exe" #"$PSScriptRoot\HandBrakeCLI.exe"

function Get-ImageMetadata
{
    param([string]$File)
    function GetTakenData($image) { try { return $image.GetPropertyItem(36867).Value }	 catch { return $null } }
    [Reflection.Assembly]::LoadFile('C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll') | Out-Null
    $image = New-Object System.Drawing.Bitmap -ArgumentList $file
    try {
        $takenData = $null
        $takenData = GetTakenData($image)
        if ($takenData -eq $null) { return $null }
        $takenValue = [System.Text.Encoding]::Default.GetString($takenData, 0, $takenData.Length - 1)
        $taken = [DateTime]::ParseExact($takenValue, 'yyyy:MM:dd HH:mm:ss', $null)
        return New-Object psobject -Property @{Height = $image.Height; Width = $image.Width; Date = $taken}
    }
    catch { }
    finally { $image.Dispose() }
}
Write-Host $("#"*110) -f Cyan
Write-Host "Bienvenue sur le script de tri, renommage et encodage des photos/vidéos automatique - by Matthieu BOURGAIN" -f Cyan
Write-Host "Le chemin de scan est : $CheminScan" -f Cyan
Write-Host "Renommage des photos selon $FormatNomFichier.ext" -f Cyan
Write-Host "Encodage des vidéos via HandBrake, preset Normal en mp4" -f Cyan
Write-Host $("#"*110) -f Cyan
Write-Host

# Renommage des photos
Get-ChildItem $CheminScan -Include @("*.jpg") -Recurse | ForEach-Object {
    $NomComplet = $_.Name
    $Nom = $_.BaseName
    $Chemin = $_.FullName
    $CheminParent = $_.Directory.FullName
    $Metadata= Get-ImageMetadata -File $Chemin
    $NouveauNom = $Metadata.Date.ToString($FormatNomFichier)
    $Extension = $_.Extension.ToLower()
    $I = 1
    # Si on arrive à récupérer la date de prise de vue
    if(-not ($Metadata.Date -eq $null))
    {
        # Si le fichier n'a pas déjà le bon nom
        if($NomComplet -ne "$NouveauNom$Extension")
        {
            $NomOriginal = $NouveauNom
            # Si le nom existe déjà, on met un chiffre à la fin
            while (Test-Path "$CheminParent\$NouveauNom$Extension") {
                $NouveauNom = "$NomOriginal-$I"
                $I++
            }
            # Renommage
            Write-Host "$Chemin ==> Renommage en $NouveauNom$Extension"
            Move-Item $Chemin "$CheminParent\$NouveauNom$Extension"
        }
        else
        {
            # Write-Host "$Chemin ==> Déjà nommé correctement"
        }
    }
    else 
    {
         Write-Host "$NomComplet ==> Impossible de traiter (Emplacement $CheminParent) " -f Red
    }
}

# Encodage et renommage des vidéos
If(Test-Path $HandBrakeCliPath) {
    Get-ChildItem $CheminScan -Recurse -Include @("*.mov", "*.3gp", "*.avi", "*.mpg", "*.wmv", "*.flv", "*.wmv", "*.mkv") | ForEach-Object {
        $LASTEXITCODE = 0
        $currentPath = $_.FullName
        $newName = $_.LastWriteTime.ToString($FormatNomFichier)
        $newPath = "$($_.Directory)\$newName.mp4"
        $str = """$HandBrakeCliPath"" --preset=""Normal"" -d av_mp4 -i ""$currentPath"" -o ""$newPath"""
        Write-Host "Encodage de la vidéo $($_.Name) vers $newName.mp4" -NoNewline
        Invoke-Expression "& $str 2>&1"
        If($LASTEXITCODE -eq 0) {  
            Write-Host "==> Etat OK, suppression de la vidéo"
            Remove-Item $currentPath
        } Else {
            Write-Host "==> Etat NOK : $LASTEXITCODE" -f Red
        }
    }
} Else {
    Write-Host "HandBrakeCLI n'a pas été trouvé ($HandBrakeCliPath), merci de l'installer" -f Yellow
}

Write-Host 
Write-Host $("#"*110) -f Cyan
Write-Host "Fin du programme, appuyez sur Entrée pour continuer"
Read-Host