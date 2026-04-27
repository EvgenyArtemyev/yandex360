# ==========================================
# 1. ТВОИ НАСТРОЙКИ
# ==========================================
$Token = "y0__XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" # <-- Твой токен
$OrgId = "XXXXXXXX" # <-- твой OrgID

# Проверяем папку
if (-not (Test-Path -Path "C:\temp")) { New-Item -ItemType Directory -Path "C:\temp" -Force | Out-Null }
$CurrentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputPath = "C:\temp\users_org_${OrgId}_${CurrentDate}.csv"

$CleanToken = $Token.Trim() -replace "[\r\n\t]", ""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==========================================
# ФУНКЦИЯ: Бронебойный запрос к Яндексу (UTF-8)
# ==========================================
function Get-YandexData ($Url) {
    $Request = [System.Net.WebRequest]::Create($Url)
    $Request.Method = "GET"
    $Request.Headers.Add("Authorization", "OAuth $CleanToken")
    
    $WebResponse = $Request.GetResponse()
    $Stream = $WebResponse.GetResponseStream()
    $Reader = New-Object System.IO.StreamReader($Stream, [System.Text.Encoding]::UTF8)
    $JsonText = $Reader.ReadToEnd()
    
    $Reader.Close(); $Stream.Close(); $WebResponse.Close()
    
    return ($JsonText | ConvertFrom-Json)
}

# ==========================================
# ЭТАП 1: Получаем общий список всех UID
# ==========================================
Write-Host "ЭТАП 1: Собираем список UID со всех страниц..." -ForegroundColor Cyan
$Page = 1
$PerPage = 1000
$TotalPages = 1
$BasicUsers = @()

do {
    $ListUrl = "https://api360.yandex.net/directory/v1/org/$OrgId/users?page=$Page&perPage=$PerPage"
    Write-Host "Читаем страницу $($Page) из $($TotalPages)..." -ForegroundColor DarkGray
    
    try {
        $Response = Get-YandexData -Url $ListUrl
        if ($null -ne $Response.pages) { $TotalPages = $Response.pages }
        $BasicUsers += $Response.users
    } catch {
        Write-Host "❌ Ошибка на странице $($Page): $_" -ForegroundColor Red
        break
    }
    $Page++
} while ($Page -le $TotalPages)

$TotalUsers = $BasicUsers.Count
Write-Host "Найдено пользователей: $TotalUsers" -ForegroundColor Green

# ==========================================
# ЭТАП 2: Запрашиваем детали (Displayname) по каждому
# ==========================================
Write-Host "`nЭТАП 2: Получаем DisplayName для каждого пользователя (это займет время)..." -ForegroundColor Cyan
$AllDetailedUsers = @()
$Counter = 1

foreach ($User in $BasicUsers) {
    $Uid = $User.id
    
    # Показываем прогресс-бар в PowerShell
    $Percent = [math]::Round(($Counter / $TotalUsers) * 100)
    Write-Progress -Activity "Опрос API Яндекса" -Status "Пользователь $Counter из $TotalUsers ($Percent%)" -PercentComplete $Percent
    
    try {
        # Запрашиваем карточку конкретного пользователя для получения DisplayName
        $DetailUrl = "https://api360.yandex.net/directory/v1/org/$OrgId/users/$Uid"
        $DetailedInfo = Get-YandexData -Url $DetailUrl

        # Безопасно извлекаем ФИО и Email (иногда у служебных учеток их может не быть)
        $FirstName = if ($null -ne $User.name -and $null -ne $User.name.first) { $User.name.first } else { "" }
        $LastName  = if ($null -ne $User.name -and $null -ne $User.name.last) { $User.name.last } else { "" }
        $MiddleName= if ($null -ne $User.name -and $null -ne $User.name.middle) { $User.name.middle } else { "" }
        $Email     = if ($null -ne $User.email) { $User.email } else { "" }

        # Собираем финальную строчку таблицы (добавили колонку Email)
        $AllDetailedUsers += [PSCustomObject]@{
            Firstname   = $FirstName
            LastName    = $LastName
            MiddleName  = $MiddleName
            Email       = $Email
            UID         = $Uid
            Displayname = $DetailedInfo.displayName
        }
    } catch {
        Write-Host "⚠️ Ошибка при запросе UID $Uid : $_" -ForegroundColor Yellow
    }
    
    $Counter++
}

# Закрываем прогресс-бар
Write-Progress -Activity "Опрос API Яндекса" -Completed

# ==========================================
# СОХРАНЕНИЕ
# ==========================================
# Сохраняем в строгий UTF-8
$AllDetailedUsers | Export-Csv -Path $OutputPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation

Write-Host "`n=== СБОР ЗАВЕРШЕН ===" -ForegroundColor Green
Write-Host "Файл успешно сохранен по пути: $OutputPath" -ForegroundColor Cyan
