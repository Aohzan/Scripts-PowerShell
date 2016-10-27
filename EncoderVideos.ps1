# Script d'encodage et nommage des vidéos automatique
# by Matthieu BOURGAIN

# Chemin du scan :
$scanPath = $((New-Object -ComObject "Shell.Application" ).BrowseForFolder(0,"Choisir le dossier :",0)).Self.Path


Write-Host $("#"*90) -f Cyan
Write-Host "Bienvenue sur le script d'encodage et renommage des vidéos automatique" -f Cyan
Write-Host "Le chemin de scan est : $scanPath" -f Cyan
Write-Host "Encodage des vidéos via HandBrake, preset Normal en mp4" -f Cyan
Write-Host $("#"*90) -f Cyan
Write-Host

If(Test-Path "c:\Program Files\Handbrake\HandBrakeCLI.exe") {
Get-ChildItem $scanPath -Recurse -Include @("*.mov", "*.3gp", "*.avi", "*.mpg") | %{ 
    Write-Host "Encodage de la vidéo : "$_.Name -f Yellow
    $currentPath = $_.FullName
    $newName = $_.LastWriteTime.ToString("yyyy-MM-dd-HH\hmm\mss")
    $newPath = "$($_.Directory)\$newName.mp4"
    $str = """c:\Program Files\Handbrake\HandBrakeCLI.exe"" --preset=""Normal"" -d av_mp4 -i ""$currentPath"" -o ""$newPath"""
    Write-Host "==> Lancement de la commande : "$str
    $log = Invoke-Expression "& $str 2>&1"
    If($LASTEXITCODE -eq 0) {  
        Write-Host "==> Etat OK, suppression de la vidéo"
        Remove-Item $currentPath
    } Else {
        Write-Host "==> Etat NOK" -f Red
    }
    Write-Host $("#"*90)
}
} Else {
    Write-Host "==> HandBrakeCLI n'a pas été trouvé, merci de l'installer" -f Red
}

Write-Host 
Write-Host $("#"*90) -f Cyan
Write-Host "Fin du programme, appuyez sur Entrée pour continuer"
Read-Host