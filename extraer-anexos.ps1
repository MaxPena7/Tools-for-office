# Obtener la ruta del escritorio del usuario actual
$desktop = [Environment]::GetFolderPath("Desktop")

# Rutas de origen y destino
$origen = Join-Path $desktop "ANEXOS V Y IV 2023"
$destino = Join-Path $desktop "anexosext"

# Crear la carpeta de destino si no existe
if (-not (Test-Path -Path $destino)) {
    New-Item -ItemType Directory -Path $destino | Out-Null
}

# Buscar todos los archivos dentro de las subcarpetas del origen
Get-ChildItem -Path $origen -Recurse -File | ForEach-Object {
    $archivoDestino = Join-Path -Path $destino -ChildPath $_.Name

    # Si el archivo ya existe, renombrar para evitar sobrescritura
    $contador = 1
    while (Test-Path $archivoDestino) {
        $nombreSinExtension = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $extension = [System.IO.Path]::GetExtension($_.Name)
        $archivoDestino = Join-Path -Path $destino -ChildPath "$nombreSinExtension`_$contador$extension"
        $contador++
    }

    # Mover el archivo
    Move-Item -Path $_.FullName -Destination $archivoDestino
}
