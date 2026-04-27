# ==========================================
# 1. ТВОИ НАСТРОЙКИ
# ==========================================
$Token = "y0__XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" # <-- Твой токен
$OrgId = "12345678"
$CsvPath = "C:\temp\update_displaynames.csv" # <-- Путь к твоему файлу



# ==========================================
# ПОДГОТОВКА И ПРОВЕРКИ
# ==========================================
# Прячем ошибку кодировки PowerShell ISE
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$CleanToken = $Token.Trim() -replace "[\r\n\t]", ""
$Headers = @{ "Authorization" = "OAuth $CleanToken" }

if (-not (Test-Path -Path $CsvPath)) {
    Write-Host "❌ Ошибка: Файл не найден по пути $CsvPath" -ForegroundColor Red
    return
}

# Сначала читаем с точкой с запятой
$UsersToUpdate = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8

if ($UsersToUpdate.Count -eq 0) {
    Write-Host "❌ Файл пуст!" -ForegroundColor Red
    return
}

# УМНАЯ ПРОВЕРКА КОЛОНОК
$FirstRowCols = $UsersToUpdate[0].psobject.properties.name
if ($FirstRowCols.Count -eq 1 -and $FirstRowCols[0] -match ",") {
    Write-Host "❌ ОШИБКА ФАЙЛА: Похоже, в твоем файле разделителем является ЗАПЯТАЯ (,) ! Скрипт перезапускает чтение с запятой..." -ForegroundColor Yellow
    # Перечитываем файл с запятой
    $UsersToUpdate = Import-Csv -Path $CsvPath -Delimiter "," -Encoding UTF8
    $FirstRowCols = $UsersToUpdate[0].psobject.properties.name
}

# Ищем нужные колонки (умный поиск, игнорирующий невидимые символы)
$UidColName = $FirstRowCols | Where-Object { $_ -match "uid" -or $_ -match "id" } | Select-Object -First 1
$NameColName = $FirstRowCols | Where-Object { $_ -match "name" -or $_ -match "display" } | Select-Object -First 1

if (-not $UidColName -or -not $NameColName) {
    Write-Host "❌ ОШИБКА: Не удалось найти колонки с UID и Именами." -ForegroundColor Red
    Write-Host "Скрипт увидел в твоем файле только вот эти колонки: $($FirstRowCols -join ' | ')" -ForegroundColor Cyan
    return
}

$Total = $UsersToUpdate.Count
$Counter = 1

Write-Host "Найдено записей: $Total" -ForegroundColor Green
Write-Host "Используем колонку для UID: '$UidColName'" -ForegroundColor DarkGray
Write-Host "Используем колонку для Имени: '$NameColName'" -ForegroundColor DarkGray
Write-Host "---------------------------------"

# ==========================================
# ЦИКЛ ОБНОВЛЕНИЯ
# ==========================================
foreach ($User in $UsersToUpdate) {
    
    # Берем данные из умных колонок
    $RawUid = $User.$UidColName
    $RawName = $User.$NameColName

    $Uid = $RawUid -replace '[^0-9]', ''
    $NewName = $RawName -replace "[\r\n\t]", ""
    $NewName = $NewName -replace '\s+', ' '
    $NewName = $NewName.Trim()

    if ([string]::IsNullOrWhiteSpace($Uid) -or [string]::IsNullOrWhiteSpace($NewName)) { 
        Write-Host "[$Counter/$Total] ⚠️ Пропущен (пустое имя или UID)" -ForegroundColor DarkYellow
        $Counter++
        continue 
    }

    $Url = "https://api360.yandex.net/directory/v1/org/$OrgId/users/$Uid"
    $BodyJson = @{ "displayName" = $NewName } | ConvertTo-Json -Compress
    $BodyBytes = [System.Text.Encoding]::UTF8.GetBytes($BodyJson)

    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Patch -Headers $Headers -ContentType "application/json; charset=utf-8" -Body $BodyBytes -ErrorAction Stop
        Write-Host "[$Counter/$Total] ✅ Успех: $Uid -> '$NewName'" -ForegroundColor Green
    }
    catch {
        Write-Host "[$Counter/$Total] ❌ Ошибка: $Uid -> '$NewName'" -ForegroundColor Red
        if ($_.ErrorDetails) { Write-Host "   Детали: $($_.ErrorDetails.Message)" -ForegroundColor DarkGray }
    }
    
    $Counter++
}

Write-Host "---------------------------------"
Write-Host "=== ВСЕ ПОЛЬЗОВАТЕЛИ ОБРАБОТАНЫ ===" -ForegroundColor Cyan
