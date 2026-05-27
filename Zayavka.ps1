# ==========================================================
# 🚀 ФОРМА ЗАЯВКИ 3.1 (ПОЛНОЕ ИСПРАВЛЕНИЕ ЛОГИСТИКИ)
# ==========================================================
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$dataJson = Join-Path $scriptDir "data.json"
$excelPath = Join-Path $scriptDir "Выгрузка_заявок.xlsx"
$uploadsDir = Join-Path $scriptDir "uploads"
$port = 9000

if (-not (Test-Path $uploadsDir)) { New-Item -ItemType Directory -Path $uploadsDir | Out-Null }

$smtpCfg = @{
    Server   = "mail-01"
    Port     = 25
    From     = "support@stroisservis.ru"
    User     = "support@stroisservis.ru"
    Password = 'RI2b}mb*A?yvrF9fEj'
}

$dbMutex = New-Object System.Threading.Mutex($false, "ZayavkaDbMutex_v31")

# ==========================================================
# 🔧 БЛОК 1: АВТО-ОЧИСТКА БАЗЫ
# ==========================================================
$defaultDb = @{
    system = @{
        version = "3.1"
        statuses = @(
            @{ name = "Отправлено на рассмотрение"; availableFrom = @(); availableFor = @("client", "procurement") }
            @{ name = "Выполняется"; availableFrom = @("Отправлено на рассмотрение"); availableFor = @("procurement") }
            @{ name = "Оплачено"; availableFrom = @("Выполняется"); availableFor = @("director") }
            @{ name = "Договорённость"; availableFrom = @("Выполняется"); availableFor = @("director") }
            @{ name = "Отклонено"; availableFrom = @("Отправлено на рассмотрение", "Выполняется"); availableFor = @("procurement", "director") }
            @{ name = "Приостановлено"; availableFrom = @("Выполняется"); availableFor = @("director", "procurement") }
            @{ name = "Возвращено на доработку"; availableFrom = @("Выполняется", "Приостановлено"); availableFor = @("director") }
        )
    }
    auth = @{ users = @() }
    projects = @{ active = @(); archive = @() }
    requests = @()
}

function Repair-Database {
    if (-not (Test-Path $dataJson)) {
        $defaultDb | ConvertTo-Json -Depth 10 | Set-Content $dataJson -Encoding UTF8
        Write-Host "✅ Создан новый data.json" -ForegroundColor Green
        return
    }
    try {
        $raw = Get-Content $dataJson -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        function Clean($item) {
            if ($item -is [PSCustomObject]) {
                $out = [PSCustomObject]@{}
                foreach ($p in $item.PSObject.Properties) {
                    $out | Add-Member -NotePropertyName $p.Name.Trim() -NotePropertyValue (Clean $p.Value) -Force
                }
                return $out
            } elseif ($item -is [System.Collections.IList]) {
                return @($item | ForEach-Object { Clean $_ })
            } elseif ($item -is [string]) {
                return $item.Trim()
            }
            return $item
        }
        $cleaned = Clean $obj
        $final = @{
            system   = if ($cleaned.system) { $cleaned.system } else { $defaultDb.system }
            auth     = @{ users = if ($cleaned.auth -and $cleaned.auth.users) { $cleaned.auth.users } else { @() } }
            projects = @{
                active  = if ($cleaned.projects -and $cleaned.projects.active) { $cleaned.projects.active } else { @() }
                archive = if ($cleaned.projects -and $cleaned.projects.archive) { $cleaned.projects.archive } else { @() }
            }
            requests = if ($cleaned.requests) { $cleaned.requests } else { @() }
        }
        if (-not $final.system.statuses) { $final.system.statuses = $defaultDb.system.statuses }
        $final | ConvertTo-Json -Depth 10 | Set-Content $dataJson -Encoding UTF8
        Write-Host "✅ data.json проверен" -ForegroundColor Cyan
    } catch {
        Write-Host "⚠️ Ошибка БД, создаём резервную копию" -ForegroundColor Yellow
        $backupFile = "$dataJson.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Copy-Item $dataJson $backupFile -ErrorAction SilentlyContinue
        $defaultDb | ConvertTo-Json -Depth 10 | Set-Content $dataJson -Encoding UTF8
    }
}

# ==========================================================
# 💾 БЛОК 2: DATABASE (С ПРИВЕДЕНИЕМ ТИПОВ)
# ==========================================================
function Get-Db {
    try {
        $raw = Get-Content $dataJson -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json
        
        # 🔥 ПРИНУДИТЕЛЬНОЕ ПРИВЕДЕНИЕ МАССИВОВ
        if ($obj.auth -and $obj.auth.users) {
            $obj.auth.users = @($obj.auth.users)
            foreach ($u in $obj.auth.users) {
                if ($u.PSObject.Properties.Name -contains 'readRequests') {
                    # Очищаем readRequests от битых объектов {"Length": X}
                    $cleanRR = @()
                    foreach ($rr in @($u.readRequests)) {
                        if ($rr -is [string]) { $cleanRR += $rr }
                    }
                    $u.readRequests = $cleanRR
                }
            }
        }
        
        if ($obj.projects) {
            $obj.projects.active = @($obj.projects.active)
            $obj.projects.archive = @($obj.projects.archive)
        }
        
        $obj.requests = @($obj.requests)
        foreach ($r in $obj.requests) {
            if ($r.PSObject.Properties.Name -contains 'files') {
                # Очищаем files от null и битых объектов
                $cleanFiles = @()
                foreach ($f in @($r.files)) {
                    if ($f -and $f -is [PSCustomObject] -and $f.path -and $f.name) {
                        $cleanFiles += $f
                    }
                }
                $r.files = $cleanFiles
            }
            if ($r.PSObject.Properties.Name -contains 'audit') {
                $r.audit = @($r.audit)
            }
        }
        
        # 🔥 КРИТИЧЕСКОЕ: Восстановление statuses если битые
        if ($obj.system -and $obj.system.statuses) {
            $statusesValid = $true
            foreach ($s in $obj.system.statuses) {
                if ($s.availableFor) {
                    foreach ($af in @($s.availableFor)) {
                        if ($af -isnot [string]) { $statusesValid = $false; break }
                    }
                }
                if (-not $statusesValid) { break }
            }
            if (-not $statusesValid) {
                Write-Host "⚠️ Statuses битые, восстанавливаю из шаблона" -ForegroundColor Yellow
                $obj.system.statuses = $defaultDb.system.statuses
                # Сохраняем исправленную версию
                $obj | ConvertTo-Json -Depth 10 | Set-Content $dataJson -Encoding UTF8
            }
        }
        
        return $obj
    } catch {
        Write-Host "⚠️ Ошибка чтения: $_" -ForegroundColor Yellow
        return $defaultDb
    }
}

function Set-Db($data) { 
    $data | ConvertTo-Json -Depth 10 | Set-Content $dataJson -Encoding UTF8 
}

# ==========================================================
# 🔧 БЛОК 3: МИГРАЦИЯ
# ==========================================================
function Migrate-Database {
    Write-Host "🔄 Проверка структуры БД..." -ForegroundColor Magenta
    $db = Get-Db
    $changed = $false
    
    # Проверка statuses
    if (-not $db.system.statuses) {
        $db.system.statuses = $defaultDb.system.statuses
        $changed = $true
    }
    
    $db.auth.users = @($db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() })
    foreach ($u in $db.auth.users) {
        if ($u.PSObject.Properties.Name -notcontains 'readRequests') {
            $u | Add-Member -NotePropertyName 'readRequests' -NotePropertyValue @() -Force
            $changed = $true
        }
        if ($u.readRequests -eq $null) { $u.readRequests = @(); $changed = $true }
        
        # 🔥 ОЧИСТКА БИТЫХ readRequests
        $cleanRR = @()
        foreach ($rr in @($u.readRequests)) {
            if ($rr -is [string] -and -not [string]::IsNullOrWhiteSpace($rr)) { 
                $cleanRR += $rr 
            } else { 
                Write-Host "⚠️ Очищен битый readRequest у $($u.username)" -ForegroundColor Yellow
                $changed = $true 
            }
        }
        if ($cleanRR.Count -ne @($u.readRequests).Count) { 
            $u.readRequests = $cleanRR
            $changed = $true
        }
        
        if ($u.PSObject.Properties.Name -notcontains 'deleted') {
            $u | Add-Member -NotePropertyName 'deleted' -NotePropertyValue $false -Force
            $changed = $true
        }
        if ($u.PSObject.Properties.Name -notcontains 'deletedAt') {
            $u | Add-Member -NotePropertyName 'deletedAt' -NotePropertyValue $null -Force
            $changed = $true
        }
    }
    
    foreach ($p in $db.projects.active) {
        if ($p -is [string]) {
            $idx = $db.projects.active.IndexOf($p)
            $db.projects.active[$idx] = [PSCustomObject]@{ name = $p; deadline = $null }
            $changed = $true
        } else {
            if ($p.PSObject.Properties.Name -notcontains 'deadline') {
                $p | Add-Member -NotePropertyName 'deadline' -NotePropertyValue $null -Force
                $changed = $true
            }
        }
    }
    foreach ($p in $db.projects.archive) {
        if ($p -is [string]) {
            $idx = $db.projects.archive.IndexOf($p)
            $db.projects.archive[$idx] = [PSCustomObject]@{ name = $p; deadline = $null }
            $changed = $true
        } else {
            if (-not $p.name) {
                Write-Host "⚠️ Удалён битый объект из archive" -ForegroundColor Yellow
                $db.projects.archive = @($db.projects.archive | Where-Object { $_ -ne $p })
                $changed = $true
            } elseif ($p.PSObject.Properties.Name -notcontains 'deadline') {
                $p | Add-Member -NotePropertyName 'deadline' -NotePropertyValue $null -Force
                $changed = $true
            }
        }
    }
    
    foreach ($r in $db.requests) {
        $defaultProps = @{
            id = $null; project = $null; createdDate = (Get-Date).ToString("yyyy-MM-dd")
            deadline = $null; body = $null; status = "Отправлено на рассмотрение"
            author = $null; authorDeleted = $false; authorDeletedAt = $null
            priority = "medium"; takenBy = $null; comment = $null; audit = @()
            lastDeadlineNotify = $null; budget = $null; files = @()
        }
        foreach ($prop in $defaultProps.Keys) {
            if ($r.PSObject.Properties.Name -notcontains $prop) {
                $r | Add-Member -NotePropertyName $prop -NotePropertyValue $defaultProps[$prop] -Force
                $changed = $true
            }
        }
        
        # 🔥 ОЧИСТКА БИТЫХ files
        $cleanFiles = @()
        foreach ($f in @($r.files)) {
            if ($f -and $f -is [PSCustomObject] -and $f.path -and $f.name) {
                $cleanFiles += $f
            } else {
                Write-Host "⚠️ Очищен битый файл из заявки $($r.id)" -ForegroundColor Yellow
                $changed = $true
            }
        }
        if ($cleanFiles.Count -ne @($r.files).Count) {
            $r.files = $cleanFiles
            $changed = $true
        }
        
        if ($r.status -eq "В работу") { $r.status = "Выполняется"; $changed = $true }
        if ($r.audit.Count -eq 0) {
            $auditEntry = [PSCustomObject]@{
                timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                actor = $r.author
                action = "Создание заявки"
                newStatus = $r.status
                comment = " "
            }
            $r.audit = @($auditEntry)
            $changed = $true
        }
    }
    
    if ($changed) {
        Set-Db $db
        Write-Host "✅ Миграция БД завершена" -ForegroundColor Green
    } else {
        Write-Host "✅ Структура БД в порядке" -ForegroundColor Green
    }
}

# ==========================================================
# 🔐 БЛОК 4: AUTH & USERS
# ==========================================================
function New-Code { (100000..999999 | Get-Random).ToString() }

function Auth-Register($username, $password, $role, $name) {
    $db = Get-Db
    $email = $username.ToLower().Trim()
    if (-not $email.EndsWith('@stroisservis.ru')) { return @{ok=$false; error="Только @stroisservis.ru"} }
    $existing = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $email -and $_.deleted -eq $false }
    if ($existing) { return @{ok=$false; error="Email занят"} }
    $code = New-Code
    $newUser = [PSCustomObject]@{
        username=$email; password=$password; role=$role; name=$name
        verified=$false; verificationCode=$code; codeExpires=(Get-Date).AddMinutes(15).ToString("o")
        readRequests = @(); deleted=$false; deletedAt=$null
    }
    $db.auth.users += $newUser
    Set-Db $db; Send-VerificationEmail $email $code
    return @{ok=$true}
}

function Auth-Verify($username, $code) {
    $db = Get-Db
    $u = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $username.Trim() -and $_.verified -eq $false -and $_.deleted -eq $false }
    if (-not $u) { return @{ok=$false; error="Не найден"} }
    if ($u.codeExpires -and [DateTime]::Parse($u.codeExpires) -lt (Get-Date)) { return @{ok=$false; error="Код истёк"} }
    if ($u.verificationCode -ne $code) { return @{ok=$false; error="Неверный код"} }
    $u.verified=$true; $u.verificationCode=$null; $u.codeExpires=$null
    Set-Db $db
    return @{ok=$true; username=$u.username}
}

function Auth-Login($username, $password) {
    $db = Get-Db
    $u = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $username.Trim() -and $_.password -eq $password -and $_.deleted -eq $false }
    if (-not $u) { return @{ok=$false; error="Неверные данные или пользователь удалён"} }
    if (-not $u.verified) { return @{ok=$false; error="Подтвердите почту"} }
    return @{ok=$true; user=@{username=$u.username; role=$u.role; name=$u.name}}
}

function Get-User($username) {
    if (-not $username) { return $null }
    $db = Get-Db
    $clean = $username.Trim().ToLower()
    foreach ($u in $db.auth.users) {
        if ($u -and $u.username -and $u.username.Trim().ToLower() -eq $clean -and $u.verified -and $u.deleted -eq $false) { return $u }
    }
    return $null
}

function Get-UserDisplayName($username, $db) {
    $user = $db.auth.users | Where-Object { $_.name -eq $username }
    if ($user -and $user.deleted -eq $true) {
        return "$($user.name) (Удалён $($user.deletedAt))"
    } elseif ($user) {
        return $user.name
    }
    return $username
}

function Action-DeleteUser($username, $actorRole) {
    if ($actorRole -ne "director") { return @{ok=$false; error="Только руководитель может удалять пользователей"} }
    $db = Get-Db
    $user = $db.auth.users | Where-Object { $_.username -eq $username -and $_.deleted -eq $false }
    if (-not $user) { return @{ok=$false; error="Пользователь не найден"} }
    $user.deleted = $true
    $user.deletedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    foreach ($r in $db.requests) {
        if ($r.author -eq $user.name) {
            $r.authorDeleted = $true
            $r.authorDeletedAt = $user.deletedAt
            AddAuditEntry $r "Система" $r.status "Автор удалён из системы" "Автор удалён"
        }
    }
    Set-Db $db
    return @{ok=$true; message="Пользователь $($user.name) удалён"}
}

# ==========================================================
# 📧 БЛОК 5: EMAIL
# ==========================================================
function Send-Mail($to, $subj, $body) {
    try {
        $cred = New-Object PSCredential($smtpCfg.User, (ConvertTo-SecureString $smtpCfg.Password -AsPlainText -Force))
        Send-MailMessage -SmtpServer $smtpCfg.Server -Port $smtpCfg.Port -From $smtpCfg.From -To $to -Subject $subj -Body $body -BodyAsHtml -Encoding UTF8 -UseSsl:$false -Credential $cred
    } catch { Write-Host "❌ SMTP: $_" -ForegroundColor Red }
}

function Send-VerificationEmail($email, $code) {
    $body = @"
<div style='font-family:Roboto,sans-serif;text-align:center'>
<h2>🔐 Код подтверждения</h2>
<p>Ваш код: <b style='font-size:2em;color:#1a73e8'>$code</b></p>
<p>Действует 15 минут.</p>
<hr>
<p style='font-size:0.9em;color:#666;'>Если вы не регистрировались, просто проигнорируйте это письмо.</p>
</div>
"@
    Send-Mail $email "Форма заявки: Код подтверждения" $body
}

function Notify-Status($req, $status, $actor, $comment="") {
    $db = Get-Db
    $authorDisplay = Get-UserDisplayName $req.author $db
    $subject = "Изменение статуса заявки $($req.id)"
    $body = @"
<p><b>$actor</b> сменил статус заявки <b>$($req.id)</b> от <b>$authorDisplay</b> по проекту <b>$($req.project)</b> на <b>$status</b>.</p>
<p><b>Комментарий:</b> $comment</p>
<a href='http://localhost:$port' style='background:#1a73e8;color:#fff;padding:8px 12px;text-decoration:none;border-radius:4px'>Открыть</a>
"@
    $authorUser = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }
    if ($authorUser) {
        Send-Mail $authorUser.username $subject $body
    }
}

function Notify-SupplierNew($req) {
    $db = Get-Db
    foreach($s in ($db.auth.users | Where-Object { $_.role -eq "procurement" -and $_.verified -and $_.deleted -eq $false })) {
        $subject = "Новая заявка $($req.id)"
        $body = @"
<p>Поступила новая заявка <b>$($req.id)</b> для проекта <b>$($req.project)</b>, от заказчика <b>$($req.author)</b>.</p>
<p><b>Описание:</b> $($req.body)</p>
<a href='http://localhost:$port'>Перейти</a>
"@
        Send-Mail $s.username $subject $body
        Write-Host "📧 Уведомление отправлено в отдел снабжения" -ForegroundColor Cyan
    }
}

function Notify-DirectorNewRequestFromProcurement($req) {
    $db = Get-Db
    $directors = $db.auth.users | Where-Object { $_.role -eq "director" -and $_.verified -and $_.deleted -eq $false }
    foreach ($d in $directors) {
        $subject = "Новая заявка от отдела снабжения $($req.id)"
        $body = @"
<p>Сотрудник отдела снабжения <b>$($req.author)</b> создал заявку <b>$($req.id)</b> для проекта <b>$($req.project)</b>.</p>
<p><b>Описание:</b> $($req.body)</p>
<p><b>Срок выполнения:</b> $($req.deadline)</p>
<a href='http://localhost:$port'>Перейти</a>
"@
        Send-Mail $d.username $subject $body
    }
}

function Notify-RequestParticipants($req, $projectName) {
    $db = Get-Db
    $participants = @()
    $authorUser = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }
    if ($authorUser) { $participants += $authorUser.username }
    if ($req.takenBy -and $req.takenBy.Trim()) {
        $executorUser = $db.auth.users | Where-Object { $_.name -eq $req.takenBy -and $_.deleted -eq $false }
        if ($executorUser) { $participants += $executorUser.username }
    }
    foreach ($email in $participants) {
        $subject = "Проект '$projectName' перемещён в архив"
        $body = @"
<p>Проект <b>$projectName</b> был перемещён в архив руководителем.</p>
<p>Ваша заявка <b>$($req.id)</b> по этому проекту больше не активна.</p>
<a href='http://localhost:$port'>Открыть систему</a>
"@
        Send-Mail $email $subject $body
    }
}

# ==========================================================
# 💼 БЛОК 6: BUSINESS LOGIC
# ==========================================================
function AddAuditEntry($req, $actor, $newStatus, $comment, $action = "Изменение статуса") {
    $entry = [PSCustomObject]@{
        timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        actor = $actor
        action = $action
        newStatus = $newStatus
        comment = $comment
    }
    if ($req.audit -eq $null) { $req.audit = @() }
    $req.audit = @($req.audit) + $entry
}

# 🔥 ИСПРАВЛЕННАЯ ПРОВЕРКА ПРАВ С FALLBACK
function CheckValidStatusTransition($currentStatus, $newStatus, $role, $db) {
    $statusConfig = $db.system.statuses | Where-Object { $_.name -eq $newStatus }
    if (-not $statusConfig) { return $false, "Статус '$newStatus' не существует" }
    if ($currentStatus -eq $newStatus) { return $false, "Статус не изменён" }
    
    # Проверка availableFrom
    if ($statusConfig.availableFrom) {
        $availableFrom = @($statusConfig.availableFrom)
        # Проверяем что массив содержит только строки
        $validArray = $true
        foreach ($item in $availableFrom) {
            if ($item -isnot [string]) { $validArray = $false; break }
        }
        
        if ($validArray -and $availableFrom.Count -gt 0 -and $availableFrom -notcontains $currentStatus) {
            return $false, "Невозможно перейти из '$currentStatus' в '$newStatus'"
        }
    }
    
    # Проверка availableFor
    if ($statusConfig.availableFor) {
        $availableFor = @($statusConfig.availableFor)
        $validArray = $true
        foreach ($item in $availableFor) {
            if ($item -isnot [string]) { 
                $validArray = $false
                Write-Host "⚠️ availableFor битый для $newStatus, использую fallback" -ForegroundColor Yellow
                break 
            }
        }
        
        if ($validArray -and $availableFor -notcontains $role) {
            $roleName = if ($role -eq "director") { "руководитель" } elseif ($role -eq "procurement") { "сотрудник отдела снабжения" } else { "заказчик" }
            return $false, "Роль '$roleName' не может установить статус '$newStatus'"
        }
        
        # 🔥 FALLBACK: если массив битый, берём права из шаблона
        if (-not $validArray) {
            $defaultStatus = $defaultDb.system.statuses | Where-Object { $_.name -eq $newStatus }
            if ($defaultStatus -and $defaultStatus.availableFor) {
                $defaultAvailableFor = @($defaultStatus.availableFor)
                if ($defaultAvailableFor -notcontains $role) {
                    $roleName = if ($role -eq "director") { "руководитель" } elseif ($role -eq "procurement") { "сотрудник отдела снабжения" } else { "заказчик" }
                    return $false, "Роль '$roleName' не может установить статус '$newStatus'"
                }
            }
        }
    }
    
    return $true, "OK"
}

function Action-CreateRequest($proj, $dead, $body, $author, $priority = "medium") {
    $db = Get-Db
    $projectExists = $db.projects.active | Where-Object { $_.name -eq $proj }
    if (-not $projectExists) {
        return @{ok=$false; error="Проект '$proj' не существует или находится в архиве"}
    }
    $id = "REQ-" + ($db.requests.Count+1).ToString("D3")
    $req = [PSCustomObject]@{
        id=$id; project=$proj; createdDate=(Get-Date).ToString("yyyy-MM-dd")
        deadline=$dead; body=$body; status="Отправлено на рассмотрение"
        author=$author; authorDeleted=$false; authorDeletedAt=$null
        priority=$priority; takenBy=$null; comment=$null; audit = @()
        lastDeadlineNotify = $null
    }
    AddAuditEntry $req $author "Отправлено на рассмотрение" " " "Создание заявки"
    $db.requests += $req
    Set-Db $db
    Notify-SupplierNew $req
    $user = $db.auth.users | Where-Object { $_.name -eq $author -and $_.verified }
    if ($user -and $user.role -eq "procurement") {
        Notify-DirectorNewRequestFromProcurement $req
    }
    return @{ok=$true; id=$id}
}

function Action-EditRequest($id, $newBody, $newDead, $newPriority, $author) {
    $db = Get-Db
    $r = $db.requests | Where-Object {$_.id -eq $id}
    if (-not $r -or $r.author -ne $author) { return @{ok=$false; error="Нет прав"} }
    if ($r.status -ne "Отправлено на рассмотрение") {
        return @{ok=$false; error="Редактирование невозможно: заявка уже обработана"}
    }
    $r.body=$newBody; $r.deadline=$newDead; $r.priority=$newPriority
    AddAuditEntry $r $author $r.status "Редактирование" "Редактирование"
    Set-Db $db
    return @{ok=$true}
}

function Action-Process($id, $status, $actor, $comment="") {
    $db = Get-Db
    $r = $db.requests | Where-Object {$_.id -eq $id}
    if (-not $r) { return @{ok=$false; error="Не найдено"} }
    $cur = $r.status
    $new = $status
    $userRole = ($db.auth.users | Where-Object { $_.name -eq $actor }).role
    $valid, $errorMsg = CheckValidStatusTransition $cur $new $userRole $db
    if (-not $valid) { return @{ok=$false; error=$errorMsg} }
    
    $r.status=$new
    if ($new -eq "Выполняется" -and -not $r.takenBy) { $r.takenBy=$actor }
    $r.comment=$comment
    AddAuditEntry $r $actor $new $comment "Изменение статуса"
    Set-Db $db
    Notify-Status $r $new $actor $comment
    return @{ok=$true}
}

function Action-Reject($id, $reason, $actor) {
    return Action-Process $id "Отклонено" $actor $reason
}

function Action-Archive($name, $role) {
    if ($role -ne "director") { return @{ok=$false; error="Только руководитель"} }
    $clean = $name.Trim(); if (-not $clean) { return @{ok=$false; error="Не указано имя"} }
    $db = Get-Db
    $proj = $db.projects.active | Where-Object { $_.name -eq $clean }
    if ($proj) {
        $projectRequests = $db.requests | Where-Object { $_.project -eq $clean }
        foreach ($req in $projectRequests) {
            Notify-RequestParticipants $req $clean
        }
        $activeArray = @($db.projects.active)
        $archiveArray = @($db.projects.archive)
        $db.projects.active = @($activeArray | Where-Object { $_.name -ne $clean })
        $db.projects.archive = @($archiveArray) + @($proj)
        Set-Db $db
        return @{ok=$true}
    }
    return @{ok=$false; error="Проект не найден"}
}

function Action-DeleteFromArchive($name, $role) {
    if ($role -ne "director") { return @{ok=$false; error="Только руководитель"} }
    $clean = $name.Trim(); if (-not $clean) { return @{ok=$false; error="Не указано имя"} }
    $db = Get-Db
    $proj = $db.projects.archive | Where-Object { $_.name -eq $clean }
    if (-not $proj) { return @{ok=$false; error="Проект не найден в архиве"} }
    $db.projects.archive = @($db.projects.archive | Where-Object { $_.name -ne $clean })
    Set-Db $db
    return @{ok=$true; message="Проект '$clean' удалён из архива"}
}

function Action-CreateProject($name, $deadline, $role) {
    if ($role -notin @('director','procurement')) {
        return @{ok=$false; error="Недостаточно прав (только руководитель или сотрудник отдела снабжения)"}
    }
    $clean = $name.Trim()
    if (-not $clean) { return @{ok=$false; error="Имя проекта не может быть пустым"} }
    if (-not $deadline) { return @{ok=$false; error="Не указан крайний срок проекта"} }
    $db = Get-Db
    $exists = $db.projects.active | Where-Object { $_.name -eq $clean }
    $existsArchive = $db.projects.archive | Where-Object { $_.name -eq $clean }
    if ($exists -or $existsArchive) { return @{ok=$false; error="Проект уже существует"} }
    $projObj = [PSCustomObject]@{ name = $clean; deadline = $deadline }
    $db.projects.active += $projObj
    Set-Db $db
    return @{ok=$true}
}

function Action-AddStatus($statusName, $availableFrom, $availableFor, $role) {
    if ($role -ne "director") { return @{ok=$false; error="Только руководитель может добавлять статусы"} }
    $db = Get-Db
    $exists = $db.system.statuses | Where-Object { $_.name -eq $statusName }
    if ($exists) { return @{ok=$false; error="Статус уже существует"} }
    $newStatus = @{
        name = $statusName
        availableFrom = $availableFrom
        availableFor = $availableFor
    }
    $db.system.statuses += $newStatus
    Set-Db $db
    return @{ok=$true; message="Статус '$statusName' добавлен"}
}

function Action-UpdateProjectDeadline($name, $newDeadline, $role) {
    if ($role -ne "director") { return @{ok=$false; error="Только руководитель может изменять срок проекта"} }
    $db = Get-Db
    $proj = $db.projects.active | Where-Object { $_.name -eq $name }
    if (-not $proj) { return @{ok=$false; error="Проект не найден в активных"} }
    $proj.deadline = $newDeadline
    Set-Db $db
    return @{ok=$true}
}

function MarkRequestAsRead($username, $requestId) {
    $db = Get-Db
    $user = $db.auth.users | Where-Object { $_.username -eq $username }
    if (-not $user) { return $false }
    if ($user.readRequests -notcontains $requestId) {
        $user.readRequests += $requestId
        Set-Db $db
    }
    return $true
}

function GetUnreadRequestsCount($username) {
    $db = Get-Db
    $user = $db.auth.users | Where-Object { $_.username -eq $username }
    if (-not $user -or $user.role -ne "procurement") { return 0 }
    $allNewRequests = $db.requests | Where-Object { $_.status -eq "Отправлено на рассмотрение" }
    $unread = $allNewRequests | Where-Object { $user.readRequests -notcontains $_.id }
    return @($unread).Count
}

function GetWorkingDaysDiff($startDate, $endDate) {
    $diff = 0
    $current = $startDate.Date
    $end = $endDate.Date
    while ($current -le $end) {
        if ($current.DayOfWeek -ne [DayOfWeek]::Saturday -and $current.DayOfWeek -ne [DayOfWeek]::Sunday) {
            $diff++
        }
        $current = $current.AddDays(1)
    }
    return $diff
}

$lastExp = [DateTime]::MinValue
function Export-Excel {
    if ((Get-Date).Hour -ne 23 -or $lastExp.Date -eq (Get-Date).Date) { return }
    try {
        $reqs = (Get-Db).requests; if (-not $reqs -or $reqs.Count -eq 0) { return }
        $xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
        $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets(1)
        @("ID", "Проект", "Дата", "Срок", "Статус", "Приоритет", "Описание", "Автор") | ForEach-Object -Begin {$i=0} -Process { $i++; $ws.Cells(1,$i).Value2 = $_; $ws.Cells(1,$i).Font.Bold = $true }
        for ($row = 0; $row -lt $reqs.Count; $row++) {
            $r = $reqs[$row]
            $authorDisplay = $r.author
            if ($r.authorDeleted) { $authorDisplay = "$($r.author) (Удалён $($r.authorDeletedAt))" }
            $ws.Cells($row+2,1).Value2 = $r.id
            $ws.Cells($row+2,2).Value2 = $r.project
            $ws.Cells($row+2,3).Value2 = $r.createdDate
            $ws.Cells($row+2,4).Value2 = $r.deadline
            $ws.Cells($row+2,5).Value2 = $r.status
            $ws.Cells($row+2,6).Value2 = $r.priority
            $ws.Cells($row+2,7).Value2 = $r.body
            $ws.Cells($row+2,8).Value2 = $authorDisplay
        }
        $ws.Columns.AutoFit() | Out-Null
        $wb.SaveAs($excelPath)
        $wb.Close($false)
        $xl.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
        $lastExp = Get-Date
    } catch { Write-Host "[EXCEL] ❌ $_" -ForegroundColor Red }
}

# ==========================================================
# 🌐 БЛОК 7: БРАУЗЕРНЫЕ УВЕДОМЛЕНИЯ
# ==========================================================
function Get-UserUpdates($username, $lastCheck) {
    $db = Get-Db
    $user = $db.auth.users | Where-Object { $_.username -eq $username -and $_.deleted -eq $false }
    if (-not $user) { return @{ error = "Пользователь не найден" } }
    $lastCheckDate = if ($lastCheck -and $lastCheck -gt 0) { 
        [DateTime]::FromFileTimeUtc($lastCheck) 
    } else { 
        (Get-Date).AddHours(-1)
    }
    
    $updates = @()
    
    $newRequests = $db.requests | Where-Object { 
        $_.status -eq "Отправлено на рассмотрение" -and 
        $_.createdDate -and 
        [DateTime]::Parse($_.createdDate) -gt $lastCheckDate
    }
    
    foreach ($req in $newRequests) {
        if ($user.role -eq "procurement") {
            $updates += @{
                type = "new_request"
                title = "📋 Новая заявка"
                message = "Заявка $($req.id) от $($req.author) по проекту $($req.project)"
                requestId = $req.id
                project = $req.project
                priority = $req.priority
            }
        }
        if ($user.role -eq "director") {
            $authorUser = $db.auth.users | Where-Object { $_.name -eq $req.author }
            if ($authorUser -and $authorUser.role -eq "procurement") {
                $updates += @{
                    type = "new_request_from_procurement"
                    title = "📋 Новая заявка от отдела снабжения"
                    message = "Сотрудник $($req.author) создал заявку $($req.id) для проекта $($req.project)"
                    requestId = $req.id
                    project = $req.project
                    priority = $req.priority
                }
            }
        }
    }
    
    foreach ($req in $db.requests) {
        if ($req.audit -and $req.audit.Count -gt 0) {
            $lastAudit = $req.audit[-1]
            $auditTime = [DateTime]::Parse($lastAudit.timestamp)
            
            if ($auditTime -gt $lastCheckDate) {
                $shouldSee = $false
                if ($req.author -eq $user.name) { $shouldSee = $true }
                if ($req.takenBy -eq $user.name) { $shouldSee = $true }
                if ($user.role -eq "director" -and $lastAudit.actor -ne $user.name) { $shouldSee = $true }
                if ($user.role -eq "procurement" -and $lastAudit.actor -ne $user.name) { $shouldSee = $true }
                
                if ($shouldSee) {
                    $priorityIcon = switch($req.priority) { "high" { "🔴" } "medium" { "🟡" } default { "🟢" } }
                    $updates += @{
                        type = "status_change"
                        title = "🔄 Изменение статуса $priorityIcon"
                        message = "Заявка $($req.id): $($req.status)"
                        requestId = $req.id
                        project = $req.project
                        newStatus = $req.status
                        actor = $lastAudit.actor
                    }
                }
            }
        }
    }
    
    $today = (Get-Date).Date
    foreach ($req in $db.requests) {
        if ($req.deadline -and $req.status -notin @("Оплачено", "Договорённость", "Отклонено")) {
            $deadlineDate = [DateTime]::Parse($req.deadline).Date
            $daysLeft = ($deadlineDate - $today).Days
            
            $lastDeadlineNotify = $req.lastDeadlineNotify 
            $shouldNotify = (-not $lastDeadlineNotify) -or ($lastDeadlineNotify -ne $today.ToString("yyyy-MM-dd"))
            
            if ($daysLeft -ge 0 -and $daysLeft -le 3 -and $shouldNotify) {
                $relatedUsers = @()
                $authorUser = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }
                if ($authorUser) { $relatedUsers += $authorUser.username }
                if ($req.takenBy) {
                    $executorUser = $db.auth.users | Where-Object { $_.name -eq $req.takenBy -and $_.deleted -eq $false }
                    if ($executorUser) { $relatedUsers += $executorUser.username }
                } 
                $directors = $db.auth.users | Where-Object { $_.role -eq "director" -and $_.deleted -eq $false } | ForEach-Object { $_.username }
                $procurements = $db.auth.users | Where-Object { $_.role -eq "procurement" -and $_.deleted -eq $false } | ForEach-Object { $_.username }
                $relatedUsers += $directors
                $relatedUsers += $procurements
                $relatedUsers = $relatedUsers | Select-Object -Unique
                
                if ($relatedUsers -contains $username) {
                    $updates += @{
                        type = "deadline_warning"
                        title = "⏰ Срок заявки истекает"
                        message = "Заявка $($req.id) по проекту $($req.project): осталось $daysLeft дня(ей)"
                        requestId = $req.id
                        project = $req.project
                        daysLeft = $daysLeft
                    }
                    $req.lastDeadlineNotify = $today.ToString("yyyy-MM-dd")
                    Set-Db $db
                }
            }
        }
    }
    
    return @{ updates = $updates; count = $updates.Count }
}

# ==========================================================
# 🎨 БЛОК 8: UI (HTML из вашей версии)
# ==========================================================
$ui = @'
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Форма заявки 3.1</title>
<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
:root { --bg: #f5f5f5; --card: white; --text: #202124; --text-secondary: #5f6368; --border: #dadce0; --border-light: #e0e0e0; --primary: #1a73e8; --primary-dark: #1557b0; --danger: #d93025; --success: #137333; --warning: #e37400; --gray: #5f6368; --gray-light: #f8f9fa; --shadow: 0 1px 2px 0 rgba(60,64,67,0.3), 0 1px 3px 1px rgba(60,64,67,0.15); --shadow-hover: 0 4px 8px rgba(0,0,0,0.1); }
body.dark { --bg: #202124; --card: #2d2e32; --text: #e8eaed; --text-secondary: #9aa0a6; --border: #5f6368; --border-light: #3c4043; --primary: #8ab4f8; --primary-dark: #aecbfa; --danger: #f28b82; --success: #81c995; --warning: #fdd663; --gray: #9aa0a6; --gray-light: #3c4043; }
body { font-family: 'Roboto', sans-serif; background: var(--bg); color: var(--text); padding: 20px; line-height: 1.5; }
.auth, header, .card, .modal-c, .table-wrapper { background: var(--card); border-radius: 12px; box-shadow: var(--shadow); }
.auth { max-width: 400px; margin: 60px auto; padding: 32px; }
header { padding: 16px 24px; margin-bottom: 24px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
h1 { font-size: 1.5rem; font-weight: 500; color: var(--primary); }
h2, h3 { font-weight: 500; margin-bottom: 16px; }
.tabs { display: flex; gap: 8px; border-bottom: 1px solid var(--border); margin-bottom: 24px; }
.tab { flex: 1; text-align: center; padding: 12px 0; font-weight: 500; cursor: pointer; color: var(--text-secondary); border-bottom: 2px solid transparent; }
.tab.active { color: var(--primary); border-bottom-color: var(--primary); }
input, select, textarea { width: 100%; padding: 12px; margin: 8px 0 16px; border: 1px solid var(--border); border-radius: 8px; font-family: 'Roboto', sans-serif; font-size: 14px; background: var(--card); color: var(--text); }
.btn { padding: 10px 20px; border: none; border-radius: 8px; font-weight: 500; font-size: 14px; cursor: pointer; background: var(--gray-light); color: var(--text); }
.btn-pri { background: var(--primary); color: white; }
.btn-dan { background: var(--danger); color: white; }
.btn-suc { background: var(--success); color: white; }
.btn-wrn { background: var(--warning); color: #202124; }
.btn-gra { background: var(--gray); color: white; }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; margin: 20px 0; }
.card { padding: 20px; cursor: pointer; border: 1px solid var(--border-light); }
.card:hover { transform: translateY(-2px); box-shadow: var(--shadow-hover); }
.card-new { background: var(--gray-light); border: 1px dashed var(--border); display: flex; align-items: center; justify-content: center; min-height: 100px; }
.table-wrapper { overflow-x: auto; margin: 20px 0; border-radius: 12px; }
table { width: 100%; border-collapse: collapse; min-width: 800px; }
th { background: var(--gray-light); padding: 14px 12px; font-size: 12px; font-weight: 500; color: var(--text-secondary); }
td { padding: 12px; border-bottom: 1px solid var(--border-light); font-size: 14px; }
.st { display: inline-block; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500; }
.st-new { background: rgba(26,115,232,0.1); color: var(--primary); }
.st-wrk { background: rgba(227,116,0,0.1); color: var(--warning); }
.st-ok { background: rgba(19,115,51,0.1); color: var(--success); }
.st-no { background: rgba(217,48,37,0.1); color: var(--danger); }
.modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
.modal-c { max-width: 500px; width: 90%; padding: 24px; border-radius: 16px; max-height: 90vh; overflow-y: auto; }
.modal-buttons { display: flex; gap: 12px; margin-top: 20px; flex-wrap: wrap; }
.back { color: var(--primary); cursor: pointer; margin-bottom: 16px; display: inline-block; font-weight: 500; }
.hidden { display: none !important; }
.uinfo { background: rgba(26,115,232,0.1); color: var(--primary); padding: 6px 12px; border-radius: 20px; font-size: 13px; }
.notification-bell { position: relative; cursor: pointer; background: var(--gray-light); border: none; border-radius: 40px; padding: 6px 12px; font-size: 18px; }
.notification-badge { position: absolute; top: -5px; right: -5px; background: var(--danger); color: white; border-radius: 20px; padding: 2px 6px; font-size: 10px; font-weight: bold; }
.notification-dropdown { position: absolute; top: 50px; right: 20px; width: 320px; max-height: 400px; overflow-y: auto; background: var(--card); border-radius: 12px; box-shadow: var(--shadow-hover); z-index: 1000; display: none; }
.notification-dropdown.show { display: block; }
.notification-item { padding: 12px 16px; border-bottom: 1px solid var(--border-light); cursor: pointer; }
.audit-log { background: var(--gray-light); border-radius: 12px; padding: 12px; max-height: 300px; overflow-y: auto; font-size: 12px; }
.audit-entry { border-bottom: 1px solid var(--border-light); padding: 8px 0; }
.audit-entry:last-child { border-bottom: none; }
.theme-toggle { background: var(--gray-light); border: none; border-radius: 40px; padding: 6px 12px; cursor: pointer; font-size: 16px; }
</style>
</head>
<body>
<div id="auth" class="auth">
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
<h2 style="margin:0">🔐 Форма заявки</h2>
<button class="theme-toggle" onclick="toggleTheme()" title="Переключить тему">🌓</button>
</div>
<div id="aerr" style="color:var(--danger);text-align:center;margin-bottom:12px"></div>
<div class="tabs" id="atabs">
<div class="tab active" data-tab="in">Вход</div>
<div class="tab" data-tab="rg">Регистрация</div>
</div>
<div id="fin"><input type="email" id="lem" placeholder="Email"><input type="password" id="lps" placeholder="Пароль"><button class="btn btn-pri" id="loginBtn" style="width:100%">Войти</button></div>
<div id="frg" class="hidden"><input type="email" id="rem" placeholder="Email"><input type="text" id="rnm" placeholder="Имя"><input type="password" id="rps" placeholder="Пароль"><select id="rrl"><option value="client">Заказчик</option><option value="procurement">Отдел снабжения</option><option value="director">Руководство</option></select><button class="btn btn-pri" id="regBtn" style="width:100%">Зарегистрироваться</button></div>
<div id="fvr" class="hidden"><p style="text-align:center">🔐 Код на <b id="vem"></b></p><input type="text" id="vcd" maxlength="6" placeholder="000000" style="text-align:center"><button class="btn btn-suc" id="vrfBtn" style="width:100%">Подтвердить</button><button class="btn btn-gra" id="rsndBtn" style="width:100%;margin-top:8px">Отправить повторно</button></div>
</div>
<div id="app" class="hidden">
<header>
<h1>📋 Форма заявки</h1>
<div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
<div class="notification-bell" id="notificationBell" style="position:relative">🔔<span id="notificationCount" class="notification-badge hidden">0</span></div>
<button class="theme-toggle" onclick="toggleTheme()" title="Переключить тему">🌓</button>
<span class="uinfo" id="uinf"></span>
<span id="unreadBadgeHeader" class="hidden">0</span>
<button class="btn btn-dan" id="logoutBtn">Выйти</button>
</div>
</header>
<div id="cnt"></div>
</div>
<div id="notificationDropdown" class="notification-dropdown"><div style="padding:12px;border-bottom:1px solid var(--border-light);font-weight:500">Уведомления</div><div id="notificationList"></div></div>
<div id="mdl" class="modal"><div class="modal-c"><h3 id="mtl">Заявка</h3><form id="frm"><input type="hidden" id="fid"><label>Проект</label><input id="fpr" readonly><label>Срок выполнения</label><input type="date" id="fdd" required><label>Приоритет</label><select id="fpriority"><option value="low">🟢 Низкий</option><option value="medium" selected>🟡 Средний</option><option value="high">🔴 Высокий</option></select><label>Описание</label><textarea id="fbd" rows="3" required></textarea><div class="modal-buttons"><button type="submit" class="btn btn-pri">Сохранить</button><button type="button" class="btn btn-gra" id="modalCancel">Отмена</button></div></form></div></div>
<div id="commentModal" class="modal"><div class="modal-c"><h3 id="commentTitle">Комментарий</h3><textarea id="commentText" rows="3" placeholder="Введите комментарий..."></textarea><div class="modal-buttons"><button class="btn btn-pri" id="submitCommentBtn">Подтвердить</button><button class="btn btn-gra" id="cancelCommentBtn">Отмена</button></div></div></div>
<div id="projectModal" class="modal"><div class="modal-c"><h3>Новый проект</h3><label>Название проекта</label><input type="text" id="projectName" placeholder="Название"><label>Крайний срок проекта</label><input type="date" id="projectDeadline" required><div class="modal-buttons"><button class="btn btn-pri" id="createProjectBtn">Создать</button><button class="btn btn-gra" id="cancelProjectBtn">Отмена</button></div></div></div>
<div id="notificationModal" class="modal"><div class="modal-c"><h3 id="notificationTitle">Уведомление</h3><p id="notificationMessage"></p><div class="modal-buttons"><button class="btn btn-pri" id="notificationOkBtn">OK</button></div></div></div>
<div id="editProjectDeadlineModal" class="modal"><div class="modal-c"><h3>Изменить срок проекта</h3><label>Проект</label><input type="text" id="editProjectName" readonly><label>Новый крайний срок</label><input type="date" id="editProjectDeadline" required><div class="modal-buttons"><button class="btn btn-pri" id="saveProjectDeadlineBtn">Сохранить</button><button class="btn btn-gra" id="cancelEditDeadlineBtn">Отмена</button></div></div></div>
<div id="auditModal" class="modal"><div class="modal-c"><h3 id="auditTitle">История изменений</h3><div id="auditContent" class="audit-log"></div><div class="modal-buttons"><button class="btn btn-gra" id="closeAuditBtn">Закрыть</button></div></div></div>
<div id="deleteUserModal" class="modal"><div class="modal-c"><h3>Удаление пользователя</h3><label>Выберите пользователя</label><select id="deleteUserSelect" style="width:100%"></select><div class="modal-buttons"><button class="btn btn-dan" id="confirmDeleteUserBtn">🗑️ Удалить</button><button class="btn btn-gra" id="cancelDeleteUserBtn">Отмена</button></div></div></div>
<div id="addStatusModal" class="modal"><div class="modal-c"><h3>Добавить статус</h3><label>Название статуса</label><input type="text" id="newStatusName" placeholder="Например: Согласование"><label>Доступен из статусов</label><input type="text" id="newStatusFrom" placeholder="Отправлено на рассмотрение, Выполняется"><label>Доступен для ролей</label><input type="text" id="newStatusFor" placeholder="director, procurement"><div class="modal-buttons"><button class="btn btn-pri" id="confirmAddStatusBtn">➕ Добавить</button><button class="btn btn-gra" id="cancelAddStatusBtn">Отмена</button></div></div></div>
<script>
let D = null, CP = null, U = null, PE = null, IsArch = false;
let notifications = [];
let lastCheckTime = null;
let checkInterval = null;

function toggleTheme() { document.body.classList.toggle('dark'); localStorage.setItem('theme', document.body.classList.contains('dark') ? 'dark' : 'light'); }
if (localStorage.getItem('theme') === 'dark') { document.body.classList.add('dark'); }

function requestNotificationPermission() { if ('Notification' in window && Notification.permission !== 'granted' && Notification.permission !== 'denied') { Notification.requestPermission(); } }

function showBrowserNotification(title, message, requestId = null) {
  if (!('Notification' in window)) return;
  if (Notification.permission !== 'granted') return;
  const notification = new Notification(title, { body: message, silent: false });
  notification.onclick = function() { window.focus(); notification.close(); };
  setTimeout(() => notification.close(), 5000);
}

function addNotification(notif) { notifications.unshift({ ...notif, id: Date.now() + Math.random(), timestamp: new Date(), read: false }); if (notifications.length > 50) notifications.pop(); updateNotificationUI(); showBrowserNotification(notif.title, notif.message, notif.requestId); }

function updateNotificationUI() {
  const unreadCount = notifications.filter(n => !n.read).length;
  const countEl = document.getElementById('notificationCount');
  const listEl = document.getElementById('notificationList');
  if (unreadCount > 0) { countEl.textContent = unreadCount > 99 ? '99+' : unreadCount; countEl.classList.remove('hidden'); } else { countEl.classList.add('hidden'); }
  if (listEl) {
    if (notifications.length === 0) { listEl.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-secondary);">Нет уведомлений</div>'; }
    else { listEl.innerHTML = notifications.slice(0, 20).map(n => `<div class="notification-item ${n.read ? '' : 'unread'}" onclick="markNotificationRead('${n.id}', '${n.requestId || ''}')"><div style="font-weight:500">${n.title}</div><div style="font-size:12px;color:var(--text-secondary)">${n.message}</div></div>`).join(''); }
  }
}

function markNotificationRead(id, requestId) { const notif = notifications.find(n => n.id == id); if (notif) { notif.read = true; updateNotificationUI(); } }

function toggleNotificationDropdown() { const dropdown = document.getElementById('notificationDropdown'); dropdown.classList.toggle('show'); notifications.forEach(n => n.read = true); updateNotificationUI(); }

async function checkForUpdates() { if (!U) return; const checkTime = lastCheckTime || Date.now(); try { const res = await api('/updates', { lastCheck: checkTime }, 'POST'); if (res && res.updates && res.updates.length > 0) { for (const update of res.updates) { addNotification({ title: update.title, message: update.message, requestId: update.requestId }); } } lastCheckTime = Date.now(); } catch (e) { console.error('Ошибка проверки обновлений:', e); } }

function startNotificationChecker() { if (checkInterval) clearInterval(checkInterval); checkInterval = setInterval(checkForUpdates, 30000); setTimeout(checkForUpdates, 5000); }
function stopNotificationChecker() { if (checkInterval) { clearInterval(checkInterval); checkInterval = null; } }

function escapeHtml(str) { if (str === null || str === undefined) return ''; str = String(str); return str.replace(/[&<>]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m])); }
function showNotification(message, title = "Уведомление") { document.getElementById('notificationTitle').innerText = title; document.getElementById('notificationMessage').innerHTML = message; document.getElementById('notificationModal').style.display = 'flex'; }

const api = async (p, b = {}, m = 'POST') => { const h = { 'Content-Type': 'application/json' }; if (U) h['X-Username'] = U.username; try { const r = await fetch(p, { method: m, headers: h, body: Object.keys(b).length ? JSON.stringify(b) : null }); const t = await r.text(); return t ? JSON.parse(t) : { error: 'Empty' }; } catch (e) { return { error: 'Net' }; } };

function getWorkingDaysDiff(start, end) { let current = new Date(start); let endDate = new Date(end); let days = 0; while (current <= endDate) { if (current.getDay() !== 0 && current.getDay() !== 6) days++; current.setDate(current.getDate() + 1); } return days; }

async function updateUnreadCount() { if (U && U.role === 'procurement') { const res = await api('/unread-count', {}, 'GET'); if (res && typeof res.count === 'number') { const badge = document.getElementById('unreadBadgeHeader'); if (res.count > 0) { badge.textContent = res.count; badge.classList.remove('hidden'); } else { badge.classList.add('hidden'); } } } }

async function markRequestsAsRead(requestIds) { for (let id of requestIds) await api('/mark-read', { requestId: id }); await updateUnreadCount(); }

function switchTab(tabId) { document.getElementById('aerr').textContent = ''; document.getElementById('fin').classList.add('hidden'); document.getElementById('frg').classList.add('hidden'); document.getElementById('fvr').classList.add('hidden'); if (tabId === 'in') document.getElementById('fin').classList.remove('hidden'); else if (tabId === 'rg') document.getElementById('frg').classList.remove('hidden'); else if (tabId === 'vr') document.getElementById('fvr').classList.remove('hidden'); document.querySelectorAll('#atabs .tab').forEach(tab => tab.classList.toggle('active', tab.getAttribute('data-tab') === tabId)); }

async function reg() { const e = document.getElementById('rem').value.trim().toLowerCase(); const n = document.getElementById('rnm').value.trim(); const p = document.getElementById('rps').value.trim(); const r = document.getElementById('rrl').value; const er = document.getElementById('aerr'); if (!e.endsWith('@stroisservis.ru')) { er.textContent = '❌ Только @stroisservis.ru'; return; } const res = await api('/reg', { username: e, password: p, role: r, name: n }); if (res.ok) { PE = e; document.getElementById('vem').textContent = e; switchTab('vr'); er.textContent = ''; } else er.textContent = res.error; }

async function vrf() { const c = document.getElementById('vcd').value.trim(); const er = document.getElementById('aerr'); if (!c) return; const res = await api('/vrf', { username: PE, code: c }); if (res.ok) { PE = null; switchTab('in'); document.getElementById('lem').value = res.username; } else er.textContent = res.error; }

async function rsnd() { if (PE) await api('/rsnd', { username: PE }); }

async function login() { const e = document.getElementById('lem').value.trim().toLowerCase(); const p = document.getElementById('lps').value.trim(); const er = document.getElementById('aerr'); const res = await api('/log', { username: e, password: p }); if (res.ok) { U = res.user; localStorage.setItem('z1u', JSON.stringify(U)); showApp(); await loadData(); await updateUnreadCount(); requestNotificationPermission(); startNotificationChecker(); setInterval(updateUnreadCount, 30000); } else er.textContent = res.error; }

function logout() { stopNotificationChecker(); localStorage.removeItem('z1u'); U = null; CP = null; IsArch = false; notifications = []; document.getElementById('auth').classList.remove('hidden'); document.getElementById('app').classList.add('hidden'); }

function showApp() { document.getElementById('auth').classList.add('hidden'); document.getElementById('app').classList.remove('hidden'); document.getElementById('uinf').textContent = U.name + ' (' + (U.role === 'client' ? 'Заказчик' : U.role === 'procurement' ? 'Отдел снабжения' : 'Руководство') + ')'; }

async function loadData() { const res = await api('/dat', {}, 'GET'); if (res && res.error === 'Auth') { logout(); return; } if (!res || res.error) { showNotification("Ошибка загрузки данных", "Ошибка"); return; } D = res; if (!D.projects) D.projects = { active: [], archive: [] }; if (!Array.isArray(D.projects.active)) D.projects.active = []; if (!Array.isArray(D.projects.archive)) D.projects.archive = []; if (!Array.isArray(D.requests)) D.requests = []; if (!D.system) D.system = { statuses: [] }; render(); }

function getPriorityIcon(priority) { switch(priority) { case 'low': return '<span style="color:var(--success)">🟢 Низкий</span>'; case 'medium': return '<span style="color:var(--warning)">🟡 Средний</span>'; case 'high': return '<span style="color:var(--danger)">🔴 Высокий</span>'; default: return '<span style="color:var(--warning)">🟡 Средний</span>'; } }

function getDisplayStatus(request) { if (U.role !== 'client') return request.status; switch (request.status) { case 'Отправлено на рассмотрение': return 'Ожидает рассмотрения'; case 'Выполняется': return 'Принято в работу'; case 'Отклонено': return 'Отклонено'; default: return request.status; } }

function showAudit(requestId) { const req = D.requests.find(r => r.id === requestId); if (!req || !req.audit) return; let html = '<div class="audit-log">'; req.audit.slice().reverse().forEach(entry => { let commentText = entry.comment ? '<br>📝 ' + escapeHtml(entry.comment) : ''; html += `<div class="audit-entry"><b>${escapeHtml(entry.timestamp)}</b> – ${escapeHtml(entry.actor)}<br>➜ ${escapeHtml(entry.action)}: <b>${escapeHtml(entry.newStatus)}</b>${commentText}</div>`; }); html += '</div>'; document.getElementById('auditContent').innerHTML = html; document.getElementById('auditTitle').innerText = `История заявки ${requestId}`; document.getElementById('auditModal').style.display = 'flex'; }

async function deleteFromArchive(projectName) { if (!confirm(`Удалить проект "${projectName}" из архива?`)) return; const res = await api('/delete-archive', { name: projectName }); if (res.error) showNotification(res.error, "Ошибка"); else { showNotification(res.message || `Проект удалён`, "Успешно"); await loadData(); } }

async function showDeleteUserModal() { const select = document.getElementById('deleteUserSelect'); select.innerHTML = '<option value="">-- Выберите пользователя --</option>'; D.auth.users.forEach(user => { if (!user.deleted) { select.innerHTML += `<option value="${escapeHtml(user.username)}">${escapeHtml(user.name)} (${user.role === 'client' ? 'Заказчик' : user.role === 'procurement' ? 'Отдел снабжения' : 'Руководство'})</option>`; } }); document.getElementById('deleteUserModal').style.display = 'flex'; }

async function confirmDeleteUser() { const username = document.getElementById('deleteUserSelect').value; if (!username) { showNotification("Выберите пользователя", "Ошибка"); return; } if (!confirm(`Удалить пользователя?`)) return; const res = await api('/delete-user', { username: username }); if (res.error) showNotification(res.error, "Ошибка"); else { showNotification(res.message, "Успешно"); document.getElementById('deleteUserModal').style.display = 'none'; await loadData(); } }

async function showAddStatusModal() { document.getElementById('newStatusName').value = ''; document.getElementById('newStatusFrom').value = ''; document.getElementById('newStatusFor').value = ''; document.getElementById('addStatusModal').style.display = 'flex'; }

async function confirmAddStatus() { const name = document.getElementById('newStatusName').value.trim(); const fromStr = document.getElementById('newStatusFrom').value.trim(); const forStr = document.getElementById('newStatusFor').value.trim(); if (!name) { showNotification("Введите название статуса", "Ошибка"); return; } const availableFrom = fromStr ? fromStr.split(',').map(s => s.trim()) : []; const availableFor = forStr ? forStr.split(',').map(s => s.trim()) : []; const res = await api('/add-status', { name: name, availableFrom: availableFrom, availableFor: availableFor }); if (res.error) showNotification(res.error, "Ошибка"); else { showNotification(res.message, "Успешно"); document.getElementById('addStatusModal').style.display = 'none'; await loadData(); } }

async function render() {
  const container = document.getElementById('cnt');
  if (!container || !U) return;
  try {
    container.innerHTML = '';
    if (!CP) {
      let html = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px"><h3 style="margin:0">📂 Активные проекты</h3>';
      if (U.role === 'director') { html += `<div style="display:flex;gap:8px"><button class="btn btn-gra" id="deleteUserBtn">🗑️ Удалить пользователя</button><button class="btn btn-gra" id="addStatusBtn">➕ Добавить статус</button></div>`; }
      html += '</div><div class="grid">';
      if (U.role === 'procurement') { html += `<div class="card" id="allRequestsCard" style="border-left:4px solid var(--warning)"><h3 style="margin:0">📋 Все активные заявки</h3></div>`; }
      if (U.role === 'director') { const pendingCount = (D.requests || []).filter(r => r.status === 'Выполняется').length; html += `<div class="card" id="pendingRequestsCard" style="border-left:4px solid var(--warning)"><h3>⏳ Ожидают решения (${pendingCount})</h3></div>`; }
      if (U.role !== 'client') { html += '<div class="card card-new" id="newProjectBtn"><h3>+ Создать проект</h3></div>'; }
      const active = D.projects.active || [];
      active.forEach(p => { const projName = p.name || p; html += `<div class="card"><h3 class="project-name" data-project="${escapeHtml(projName)}" style="cursor:pointer;margin:0">📁 ${escapeHtml(projName)}</h3></div>`; });
      if (active.length === 0 && U.role !== 'procurement') html += '<p>Нет активных проектов</p>';
      html += '</div><h3 style="margin: 20px 0 12px 0">🗄️ Архив проектов</h3><div class="grid">';
      const archive = D.projects.archive || [];
      archive.forEach(p => { const projName = p.name || p; html += `<div class="card" style="opacity:0.7"><h3 class="project-name" data-project="${escapeHtml(projName)}" style="cursor:pointer;margin:0">📁 ${escapeHtml(projName)}</h3></div>`; });
      if (archive.length === 0) html += '<p>Архив пуст</p>';
      container.innerHTML = html + '</div>';
      document.querySelectorAll('.project-name').forEach(el => el.addEventListener('click', () => openProject(el.getAttribute('data-project'))));
      if (U.role !== 'client') document.getElementById('newProjectBtn')?.addEventListener('click', showProjectModal);
      document.getElementById('allRequestsCard')?.addEventListener('click', () => openProject('ALL_REQUESTS'));
      document.getElementById('pendingRequestsCard')?.addEventListener('click', () => openProject('PENDING_REQUESTS'));
      if (U.role === 'director') { document.getElementById('deleteUserBtn')?.addEventListener('click', showDeleteUserModal); document.getElementById('addStatusBtn')?.addEventListener('click', showAddStatusModal); }
      await updateUnreadCount();
      return;
    }
    if (CP === 'ALL_REQUESTS') {
      const allNew = D.requests.filter(r => r.status === 'Отправлено на рассмотрение');
      let html = `<div class="back" id="backBtn">← Назад</div><h2>📋 Все активные заявки</h2>`;
      if (allNew.length === 0) html += '<p>Нет новых заявок</p>';
      else {
        html += '<div class="table-wrapper"><table><thead><tr><th>Информация</th><th></th></tr></thead><tbody>';
        allNew.forEach(r => { const authorDisplay = r.authorDeleted ? `${escapeHtml(r.author)} (Удалён)` : escapeHtml(r.author); const info = `Заявка ${escapeHtml(r.id)} от ${authorDisplay} по проекту ${escapeHtml(r.project)} ${getPriorityIcon(r.priority)}`; html += `<tr><td>${info}</td><td><button class="btn btn-wrn action-take-from-all" data-id="${r.id}" data-status="Выполняется" style="width:110px;margin-right:8px">Принять</button><button class="btn btn-dan action-reject-from-all" data-id="${r.id}" data-status="Отклонено" style="width:110px">Отклонить</button><button class="btn btn-gra audit-btn" data-id="${r.id}" style="width:70px">📋</button></td></tr>`; });
        html += '</tbody></table></div>';
        await markRequestsAsRead(allNew.map(r => r.id));
      }
      container.innerHTML = html;
      document.getElementById('backBtn')?.addEventListener('click', () => { CP = null; render(); });
      document.querySelectorAll('.action-take-from-all').forEach(btn => btn.addEventListener('click', () => showCommentModal(btn.getAttribute('data-id'), 'Выполняется')));
      document.querySelectorAll('.action-reject-from-all').forEach(btn => btn.addEventListener('click', () => showCommentModal(btn.getAttribute('data-id'), 'Отклонено')));
      document.querySelectorAll('.audit-btn').forEach(btn => btn.addEventListener('click', () => showAudit(btn.getAttribute('data-id'))));
      return;
    }
    if (CP === 'PENDING_REQUESTS') {
      const pendingReqs = D.requests.filter(r => r.status === 'Выполняется');
      let html = `<div class="back" id="backBtn">← Назад</div><h2>⏳ Ожидают решения</h2>`;
      if (pendingReqs.length === 0) html += '<p>Нет заявок.</p>';
      else {
        html += '<div class="table-wrapper"><table><thead><tr><th>ID</th><th>Проект</th><th>Срок</th><th>Описание</th><th>Автор</th><th></th></tr></thead><tbody>';
        pendingReqs.forEach(r => { const authorDisplay = r.authorDeleted ? `${escapeHtml(r.author)} (Удалён)` : escapeHtml(r.author); html += `<tr><td><b>${escapeHtml(r.id)}</b></td><td>${escapeHtml(r.project)}</td><td>${escapeHtml(r.deadline || '')}</td><td>${escapeHtml(r.body || '')}</td><td>${authorDisplay}</td><td><button class="btn btn-suc action-btn" data-id="${r.id}" data-status="Оплачено" style="width:100px;margin-right:8px">Оплачено</button><button class="btn btn-gra action-btn" data-id="${r.id}" data-status="Договорённость" style="width:105px;margin-right:8px">Договорённость</button><button class="btn btn-dan action-btn" data-id="${r.id}" data-status="Отклонено" style="width:100px;margin-right:8px">Отклонить</button><button class="btn btn-gra audit-btn" data-id="${r.id}" style="width:70px">📋</button></td></tr>`; });
        html += '</tbody></table></div>';
      }
      container.innerHTML = html;
      document.getElementById('backBtn')?.addEventListener('click', () => { CP = null; render(); });
      document.querySelectorAll('.action-btn').forEach(btn => btn.addEventListener('click', () => showCommentModal(btn.getAttribute('data-id'), btn.getAttribute('data-status'))));
      document.querySelectorAll('.audit-btn').forEach(btn => btn.addEventListener('click', () => showAudit(btn.getAttribute('data-id'))));
      return;
    }
    const reqs = D.requests.filter(r => r && r.project === CP);
    const archBadge = IsArch ? ' 🔒 Архив' : '';
    const addBtn = (U.role === 'client' || U.role === 'procurement') && !IsArch ? `<button class="btn btn-suc" id="addRequestBtn" style="float:right">+ Новая заявка</button>` : '';
    let html = `<div class="back" id="backBtn">← Назад</div><h2>${escapeHtml(CP)}${archBadge} ${addBtn}</h2>`;
    if (reqs.length === 0) html += '<p>Нет заявок</p>';
    else {
      html += '<div class="table-wrapper"><table><thead><tr><th>ID</th><th>Срок</th><th>Статус</th><th>Описание</th><th>Автор</th><th></th></tr></thead><tbody>';
      reqs.forEach(r => {
        let displayStatus = getDisplayStatus(r);
        let sc = (displayStatus === 'Ожидает рассмотрения') ? 'st-new' : (displayStatus === 'Принято в работу') ? 'st-wrk' : (displayStatus === 'Отклонено') ? 'st-no' : (displayStatus === 'Оплачено' || displayStatus === 'Договорённость') ? 'st-ok' : 'st-new';
        let btns = '';
        const authorDisplay = r.authorDeleted ? `${escapeHtml(r.author)} (Удалён)` : escapeHtml(r.author);
        if (!IsArch) {
          if (U.role === 'procurement' && r.status === 'Отправлено на рассмотрение') { btns = `<button class="btn btn-wrn action-btn" data-id="${r.id}" data-status="Выполняется" style="width:100px;margin-right:8px">Принять</button><button class="btn btn-dan action-btn" data-id="${r.id}" data-status="Отклонено" style="width:100px;margin-right:8px">Отклонить</button>`; }
          if (U.role === 'director' && r.status === 'Выполняется') { btns = `<button class="btn btn-suc action-btn" data-id="${r.id}" data-status="Оплачено" style="width:100px;margin-right:8px">Оплачено</button><button class="btn btn-gra action-btn" data-id="${r.id}" data-status="Договорённость" style="width:105px;margin-right:8px">Договорённость</button><button class="btn btn-dan action-btn" data-id="${r.id}" data-status="Отклонено" style="width:100px;margin-right:8px">Отклонить</button>`; }
        }
        html += `<tr><td><b>${escapeHtml(r.id)}</b></td><td>${escapeHtml(r.deadline || '')}</td><td><span class="st ${sc}">${escapeHtml(displayStatus)}</span></td><td>${escapeHtml(r.body || '')}</td><td>${authorDisplay}</td><td>${btns}<button class="btn btn-gra audit-btn" data-id="${r.id}" style="width:70px">📋</button></td></tr>`;
      });
      html += '</tbody></table></div>';
      if (U.role === 'procurement') { const newReqs = reqs.filter(r => r.status === 'Отправлено на рассмотрение').map(r => r.id); if (newReqs.length) await markRequestsAsRead(newReqs); }
    }
    container.innerHTML = html;
    document.getElementById('backBtn')?.addEventListener('click', () => { CP = null; render(); });
    document.getElementById('addRequestBtn')?.addEventListener('click', () => openModal());
    document.querySelectorAll('.action-btn').forEach(btn => btn.addEventListener('click', () => showCommentModal(btn.getAttribute('data-id'), btn.getAttribute('data-status'))));
    document.querySelectorAll('.audit-btn').forEach(btn => btn.addEventListener('click', () => showAudit(btn.getAttribute('data-id'))));
  } catch(e) { showNotification("Ошибка: " + e.message, "Ошибка"); }
}

let currentActionId = null, currentActionStatus = null;
function showCommentModal(id, status) { currentActionId = id; currentActionStatus = status; document.getElementById('commentText').value = ''; document.getElementById('commentTitle').innerText = { 'Выполняется':'Принять в работу', 'Оплачено':'Подтверждение оплаты', 'Договорённость':'Договорённость', 'Отклонено':'Причина отклонения' }[status] || 'Комментарий'; document.getElementById('commentModal').style.display = 'flex'; }
async function submitComment() { const comment = document.getElementById('commentText').value.trim(); let res; if (currentActionStatus === 'Отклонено') res = await api('/rjt', { id: currentActionId, reason: comment }); else res = await api('/sts', { id: currentActionId, status: currentActionStatus, comment }); if (res.error) showNotification(res.error, "Ошибка"); else { closeCommentModal(); await loadData(); await updateUnreadCount(); } }
function closeCommentModal() { document.getElementById('commentModal').style.display = 'none'; currentActionId = null; currentActionStatus = null; }

function showProjectModal() { document.getElementById('projectName').value = ''; document.getElementById('projectDeadline').value = ''; document.getElementById('projectModal').style.display = 'flex'; }
async function createProjectHandler() { const name = document.getElementById('projectName').value.trim(); const deadline = document.getElementById('projectDeadline').value; if (!name) { showNotification("Введите название", "Ошибка"); return; } if (!deadline) { showNotification("Укажите срок", "Ошибка"); return; } const res = await api('/prj', { name: name, deadline: deadline }); if (res.error) showNotification(res.error, "Ошибка"); else { closeProjectModal(); await loadData(); } }
function closeProjectModal() { document.getElementById('projectModal').style.display = 'none'; }

function showEditDeadlineModal(projectName) { const proj = D.projects.active.find(p => p.name === projectName); if (!proj) { showNotification("Проект не найден", "Ошибка"); return; } document.getElementById('editProjectName').value = projectName; document.getElementById('editProjectDeadline').value = proj.deadline || ''; document.getElementById('editProjectDeadlineModal').style.display = 'flex'; }
async function saveProjectDeadline() { const name = document.getElementById('editProjectName').value; const newDeadline = document.getElementById('editProjectDeadline').value; if (!newDeadline) { showNotification("Укажите срок", "Ошибка"); return; } const res = await api('/update-project-deadline', { name: name, deadline: newDeadline }); if (res.error) showNotification(res.error, "Ошибка"); else { closeEditDeadlineModal(); await loadData(); } }
function closeEditDeadlineModal() { document.getElementById('editProjectDeadlineModal').style.display = 'none'; }

async function openProject(p) { CP = p; IsArch = (p !== 'ALL_REQUESTS' && p !== 'PENDING_REQUESTS' && D.projects.archive && D.projects.archive.some(proj => (proj.name || proj) === p)); await render(); }
async function archiveProject(n) { if (U.role !== 'director') { showNotification("Только руководитель", "Недостаточно прав"); return; } if (confirm('В архив "' + n + '"?')) { const res = await api('/arc', { name: n }); if (res.error) showNotification(res.error, "Ошибка"); else await loadData(); } }

function openModal(id = null) {
  document.getElementById('mdl').style.display = 'flex';
  const fdd = document.getElementById('fdd'); const fpr = document.getElementById('fpr');
  if (fdd) fdd.min = new Date().toISOString().split('T')[0];
  if (id) { const r = D.requests.find(x => x.id === id); if (!r) return; document.getElementById('mtl').innerText = 'Редактирование ' + id; document.getElementById('fid').value = r.id; fpr.value = r.project; fdd.value = r.deadline || ''; document.getElementById('fpriority').value = r.priority || 'medium'; document.getElementById('fbd').value = r.body || ''; }
  else { document.getElementById('mtl').innerText = 'Новая заявка'; document.getElementById('fid').value = ''; fpr.value = CP; fdd.value = ''; document.getElementById('fpriority').value = 'medium'; document.getElementById('fbd').value = ''; }
  const projectObj = D.projects.active.find(p => p.name === CP); if (projectObj && projectObj.deadline) fdd.max = projectObj.deadline; else fdd.removeAttribute('max');
}
function closeModal() { document.getElementById('mdl').style.display = 'none'; }

window.onload = () => {
  document.querySelectorAll('#atabs .tab').forEach(tab => tab.addEventListener('click', () => switchTab(tab.getAttribute('data-tab'))));
  document.getElementById('loginBtn')?.addEventListener('click', login);
  document.getElementById('regBtn')?.addEventListener('click', reg);
  document.getElementById('vrfBtn')?.addEventListener('click', vrf);
  document.getElementById('rsndBtn')?.addEventListener('click', rsnd);
  document.getElementById('logoutBtn')?.addEventListener('click', logout);
  document.getElementById('modalCancel')?.addEventListener('click', closeModal);
  document.getElementById('cancelCommentBtn')?.addEventListener('click', closeCommentModal);
  document.getElementById('submitCommentBtn')?.addEventListener('click', submitComment);
  document.getElementById('createProjectBtn')?.addEventListener('click', createProjectHandler);
  document.getElementById('cancelProjectBtn')?.addEventListener('click', closeProjectModal);
  document.getElementById('saveProjectDeadlineBtn')?.addEventListener('click', saveProjectDeadline);
  document.getElementById('cancelEditDeadlineBtn')?.addEventListener('click', closeEditDeadlineModal);
  document.getElementById('notificationOkBtn')?.addEventListener('click', () => document.getElementById('notificationModal').style.display = 'none');
  document.getElementById('closeAuditBtn')?.addEventListener('click', () => document.getElementById('auditModal').style.display = 'none');
  document.getElementById('confirmDeleteUserBtn')?.addEventListener('click', confirmDeleteUser);
  document.getElementById('cancelDeleteUserBtn')?.addEventListener('click', () => document.getElementById('deleteUserModal').style.display = 'none');
  document.getElementById('confirmAddStatusBtn')?.addEventListener('click', confirmAddStatus);
  document.getElementById('cancelAddStatusBtn')?.addEventListener('click', () => document.getElementById('addStatusModal').style.display = 'none');
  document.getElementById('notificationBell')?.addEventListener('click', toggleNotificationDropdown);
  document.getElementById('frm')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const fid = document.getElementById('fid').value;
    const fpr = document.getElementById('fpr').value;
    const fdd = document.getElementById('fdd').value;
    const fpriority = document.getElementById('fpriority').value;
    const fbd = document.getElementById('fbd').value;
    if (!fpr || !fdd || !fbd) { showNotification("Заполните все поля", "Ошибка"); return; }
    const projectObj = D.projects.active.find(p => p.name === fpr);
    if (projectObj && projectObj.deadline && new Date(fdd) > new Date(projectObj.deadline)) { showNotification("Срок не может превышать крайний срок проекта!", "Ошибка"); return; }
    let res;
    if (fid) res = await api('/edt', { id: fid, body: fbd, deadline: fdd, priority: fpriority });
    else res = await api('/req', { project: fpr, deadline: fdd, body: fbd, priority: fpriority, author: U.name });
    if (res.error) showNotification(res.error, "Ошибка");
    else { closeModal(); await loadData(); await updateUnreadCount(); }
  });
  try { const saved = localStorage.getItem('z1u'); if (saved) { U = JSON.parse(saved); showApp(); loadData(); updateUnreadCount(); requestNotificationPermission(); startNotificationChecker(); } } catch(e) {}
};
</script>
</body>
</html>
'@

# ==========================================================
# 🌐 БЛОК 9: СЕРВЕР
# ==========================================================
function Add-UrlAcl {
    param([string]$url)
    try {
        $existing = netsh http show urlacl | Select-String $url
        if (-not $existing) {
            netsh http add urlacl url=$url user=Everyone listen=yes delegate=yes | Out-Null
        }
    } catch {
        Write-Host "⚠️ Не удалось добавить разрешение" -ForegroundColor Yellow
    }
}

Repair-Database
Migrate-Database

$localIps = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "Loopback" -and $_.IPAddress -notlike "169.254*" }).IPAddress
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$port/")
$listener.Prefixes.Add("http://localhost:$port/")
foreach ($ip in $localIps) {
    $listener.Prefixes.Add("http://$ip`:$port/")
}
Add-UrlAcl -url "http://+:$port/"
foreach ($ip in $localIps) {
    Add-UrlAcl -url "http://$ip`:$port/"
}

try {
    $listener.Start()
    Write-Host "`n🟢 ФОРМА ЗАЯВКИ 3.1 ЗАПУЩЕНА!`n" -ForegroundColor Green
    Write-Host "📍 http://localhost:$port" -ForegroundColor Cyan
    foreach ($ip in $localIps) {
        Write-Host "📍 http://$ip`:$port" -ForegroundColor Yellow
    }
    Write-Host "`n💡 Firewall: New-NetFirewallRule -DisplayName 'Zayavka31' -Direction Inbound -Protocol TCP -LocalPort $port -Action Allow`n" -ForegroundColor White
}
catch {
    Write-Host "`n❌ Ошибка запуска: $_" -ForegroundColor Red
    Write-Host "💡 Запустите PowerShell от имени Администратора" -ForegroundColor Yellow
    exit 1
}

function Send-Json($obj, $code=200) {
    $ctx.Response.StatusCode = $code
    $json = if ($obj) { $obj | ConvertTo-Json -Depth 10 } else { '{"error":"null"}' }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $ctx.Response.ContentType = "application/json"
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Read-Body {
    $reader = New-Object System.IO.StreamReader($ctx.Request.InputStream, [System.Text.Encoding]::UTF8)
    $body = $reader.ReadToEnd(); $reader.Close()
    return $body | ConvertFrom-Json
}

while ($true) {
    Export-Excel
    try {
        $ctx = $listener.GetContext()
        $m = $ctx.Request.HttpMethod; $p = $ctx.Request.Url.AbsolutePath
        Write-Host "[HTTP] $m $p" -ForegroundColor DarkGray
        
        if ($m -eq "GET" -and $p -match "^/$") {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($ui)
            $ctx.Response.ContentType = "text/html; charset=utf-8"
            $ctx.Response.ContentLength64 = $bytes.Length
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        }
        elseif ($p -eq "/reg" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Register $b.username $b.password $b.role $b.name) }
        elseif ($p -eq "/vrf" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Verify $b.username $b.code) }
        elseif ($p -eq "/log" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Login $b.username $b.password) }
        elseif ($p -eq "/dat" -and $m -eq "GET") { $u = Get-User $ctx.Request.Headers["X-Username"]; if($u){Send-Json (Get-Db)}else{Send-Json @{error="Auth"} 401} }
        elseif ($p -eq "/updates" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if ($u) { $lastCheck = if ($b.lastCheck) { $b.lastCheck } else { 0 }; Send-Json (Get-UserUpdates $u.username $lastCheck) } else { Send-Json @{error="Auth"} 401 } }
        elseif ($p -eq "/prj" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json (Action-CreateProject $b.name $b.deadline $u.role) }
        elseif ($p -eq "/update-project-deadline" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json (Action-UpdateProjectDeadline $b.name $b.deadline $u.role) }
        elseif ($p -eq "/arc" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; Send-Json (Action-Archive $b.name $u.role) }
        elseif ($p -eq "/delete-archive" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; Send-Json (Action-DeleteFromArchive $b.name $u.role) }
        elseif ($p -eq "/delete-user" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json (Action-DeleteUser $b.username $u.role) }
        elseif ($p -eq "/add-status" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json (Action-AddStatus $b.name $b.availableFrom $b.availableFor $u.role) }
        elseif ($p -eq "/req" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if ($u.role -in @('client','procurement')) { Send-Json (Action-CreateRequest $b.project $b.deadline $b.body $u.name $b.priority) } else { Send-Json @{error="Access Denied"} 403 } }
        elseif ($p -eq "/sts" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; $comment = if ($b.comment -ne $null) { $b.comment.ToString() } else { " " }; Send-Json (Action-Process $b.id $b.status $u.name $comment) }
        elseif ($p -eq "/rjt" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; $reason = if ($b.reason -ne $null) { $b.reason.ToString() } else { " " }; Send-Json (Action-Reject $b.id $reason $u.name) }
        elseif ($p -eq "/edt" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; Send-Json (Action-EditRequest $b.id $b.body $b.deadline $b.priority $u.name) }
        elseif ($p -eq "/mark-read" -and $m -eq "POST") { $b = Read-Body; $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json @{ok = MarkRequestAsRead $u.username $b.requestId} }
        elseif ($p -eq "/unread-count" -and $m -eq "GET") { $u = Get-User $ctx.Request.Headers["X-Username"]; if (-not $u) { Send-Json @{error="Unauthorized"} 401; continue }; Send-Json @{count = GetUnreadRequestsCount $u.username} }
        else { Send-Json @{error="Not Found"} 404 }
    } catch {
        Write-Host "❌ Ошибка: $_" -ForegroundColor Red
        Send-Json @{error="Server"} 500
    } finally {
        $ctx.Response.Close()
    }
    Start-Sleep -Milliseconds 50
}