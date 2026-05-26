# ==========================================================
# 🚀 ФОРМА ЗАЯВКИ 3.7.9 (DASHBOARD PROJECTS FIX)
# ==========================================================
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$dataJson = Join-Path $scriptDir "data.json"
$excelPath = Join-Path $scriptDir "Выгрузка_заявок.xlsx"
$uploadsDir = Join-Path $scriptDir "uploads"
$port = 9000

if (-not (Test-Path $uploadsDir)) { New-Item -ItemType Directory -Path $uploadsDir | Out-Null }

$smtpCfg = @{
    Server   = "mail"
    Port     = 25
    From     = ""
    User     = ""
    Password = 
}

$dbMutex = New-Object System.Threading.Mutex($false, "ZayavkaDbMutex_v3")
$script:Sessions = @{}

# ==========================================================
# 🔧 БЛОК 1: АВТО-ОЧИСТКА БАЗЫ
# ==========================================================
$defaultDb = @{
    system = @{
        version = "3.7.9"
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
        $defaultDb | ConvertTo-Json -Depth 15 | Set-Content $dataJson -Encoding UTF8
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
                active  = if ($cleaned.projects) { if ($cleaned.projects.active) { $cleaned.projects.active } elseif ($cleaned.projects -is [array]) { $cleaned.projects } else { @() } } else { @() }
                archive = if ($cleaned.projects -and $cleaned.projects.archive) { $cleaned.projects.archive } else { @() }
            }
            requests = if ($cleaned.requests) { $cleaned.requests } else { @() }
        }
        if (-not $final.system.statuses) { $final.system.statuses = $defaultDb.system.statuses }
        $final | ConvertTo-Json -Depth 15 | Set-Content $dataJson -Encoding UTF8
        Write-Host "✅ data.json проверен" -ForegroundColor Cyan
    } catch {
        Write-Host "⚠️ Ошибка чтения БД, создаём резервную копию" -ForegroundColor Yellow
        $backupFile = "$dataJson.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Copy-Item $dataJson $backupFile -ErrorAction SilentlyContinue
        $defaultDb | ConvertTo-Json -Depth 15 | Set-Content $dataJson -Encoding UTF8
    }
}

# ==========================================================
# 💾 БЛОК 2: DATABASE
# ==========================================================
function Get-Db {
    $dbMutex.WaitOne() | Out-Null
    try {
        $raw = Get-Content $dataJson -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json

        if (-not $obj.auth) { $obj | Add-Member -NotePropertyName 'auth' -NotePropertyValue @{} -Force }
        $obj.auth.users = @($obj.auth.users)

        if (-not $obj.projects) { $obj | Add-Member -NotePropertyName 'projects' -NotePropertyValue @{} -Force }
        $obj.projects.active  = @($obj.projects.active)
        $obj.projects.archive = @($obj.projects.archive)

        $obj.requests = @($obj.requests)

        if (-not $obj.system) { $obj | Add-Member -NotePropertyName 'system' -NotePropertyValue @{} -Force }
        $obj.system.statuses = @($obj.system.statuses)

        foreach ($r in $obj.requests) {
            if ($r -is [PSCustomObject]) {
                $r.audit = @($r.audit)
                $r.files = @($r.files)
            }
        }
        return $obj
    } catch {
        Write-Host "⚠️ Ошибка чтения БД: $_" -ForegroundColor Yellow
        return ($defaultDb | ConvertTo-Json -Depth 15 | ConvertFrom-Json)
    } finally {
        $dbMutex.ReleaseMutex()
    }
}

function Set-Db($data) {
    $locked = $false
    try {
        if (-not $dbMutex.WaitOne(0)) { $dbMutex.WaitOne() | Out-Null }
        $locked = $true
        $data.auth.users       = @($data.auth.users)
        $data.projects.active  = @($data.projects.active)
        $data.projects.archive = @($data.projects.archive)
        $data.requests         = @($data.requests)
        $data.system.statuses  = @($data.system.statuses)
        $data | ConvertTo-Json -Depth 15 | Set-Content $dataJson -Encoding UTF8
    } finally {
        if ($locked) { try { $dbMutex.ReleaseMutex() } catch {} }
    }
}

# ==========================================================
# 🔧 БЛОК 3: МИГРАЦИЯ
# ==========================================================
function Get-PasswordHash($password, $salt) {
    $combined = $salt + $password
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($combined)
    $hash = $sha256.ComputeHash($bytes)
    return [Convert]::ToBase64String($hash)
}

function Test-StatusesValid($statuses) {
    if (-not $statuses -or $statuses.Count -eq 0) { return $false }
    foreach ($s in $statuses) {
        if (-not ($s -is [PSCustomObject] -or $s -is [System.Collections.IDictionary])) { return $false }
        if (-not $s.name -or $s.name -isnot [string]) { return $false }
        if ($s.availableFor) {
            foreach ($item in @($s.availableFor)) { if ($item -isnot [string]) { return $false } }
        }
        if ($s.availableFrom) {
            foreach ($item in @($s.availableFrom)) { if ($item -isnot [string]) { return $false } }
        }
    }
    return $true
}

function Migrate-Database {
    Write-Host "🔄 Миграция БД v3.7.9..." -ForegroundColor Magenta
    $db = Get-Db
    $changed = $false

    if (-not (Test-StatusesValid $db.system.statuses)) {
        Write-Host "⚠️  Statuses повреждены, восстанавливаю из шаблона" -ForegroundColor Yellow
        $db.system.statuses = $defaultDb.system.statuses
        $changed = $true
    }

    $db.auth.users = @($db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() })
    foreach ($u in $db.auth.users) {
        if ($u.PSObject.Properties.Name -notcontains 'readRequests') { $u | Add-Member -NotePropertyName 'readRequests' -NotePropertyValue @() -Force; $changed = $true }
        if ($u.readRequests -eq $null) { $u.readRequests = @(); $changed = $true }
        
        $cleanRR = @()
        foreach ($rr in @($u.readRequests)) {
            if ($rr -is [string] -and -not [string]::IsNullOrWhiteSpace($rr)) { $cleanRR += $rr }
            else { $changed = $true }
        }
        if ($cleanRR.Count -ne @($u.readRequests).Count) { $u.readRequests = $cleanRR }

        if ($u.PSObject.Properties.Name -notcontains 'deleted') { $u | Add-Member -NotePropertyName 'deleted' -NotePropertyValue $false -Force; $changed = $true }
        if ($u.PSObject.Properties.Name -notcontains 'deletedAt') { $u | Add-Member -NotePropertyName 'deletedAt' -NotePropertyValue $null -Force; $changed = $true }
        if ($u.PSObject.Properties.Name -notcontains 'passwordHash') {
            if ($u.password -and $u.PSObject.Properties.Name -notcontains 'salt') {
                $salt = [guid]::NewGuid().ToString("N")
                $hash = Get-PasswordHash $u.password $salt
                $u | Add-Member -NotePropertyName 'salt' -NotePropertyValue $salt -Force
                $u | Add-Member -NotePropertyName 'passwordHash' -NotePropertyValue $hash -Force
                $u | Add-Member -NotePropertyName 'password' -NotePropertyValue $null -Force
                $changed = $true
            } elseif (-not $u.password) {
                $u | Add-Member -NotePropertyName 'salt' -NotePropertyValue $null -Force
                $u | Add-Member -NotePropertyName 'passwordHash' -NotePropertyValue $null -Force
                $changed = $true
            }
        }
        if ($u.PSObject.Properties.Name -notcontains 'recoveryCode') { $u | Add-Member -NotePropertyName 'recoveryCode' -NotePropertyValue $null -Force; $changed = $true }
        if ($u.PSObject.Properties.Name -notcontains 'recoveryExpires') { $u | Add-Member -NotePropertyName 'recoveryExpires' -NotePropertyValue $null -Force; $changed = $true }
    }

    $cleanActive = @()
    foreach ($p in $db.projects.active) {
        if ($p -and $p.name -and $p.name -is [string]) {
            if ($p.PSObject.Properties.Name -notcontains 'deadline') { $p | Add-Member -NotePropertyName 'deadline' -NotePropertyValue $null -Force; $changed = $true }
            $cleanActive += $p
        } else { Write-Host "⚠️  Удалён битый объект из active" -ForegroundColor Yellow; $changed = $true }
    }
    if ($cleanActive.Count -ne @($db.projects.active).Count) { $db.projects.active = $cleanActive }

    $cleanArchive = @()
    foreach ($p in $db.projects.archive) {
        if ($p -and $p.name -and $p.name -is [string]) {
            if ($p.PSObject.Properties.Name -notcontains 'deadline') { $p | Add-Member -NotePropertyName 'deadline' -NotePropertyValue $null -Force; $changed = $true }
            $cleanArchive += $p
        } else { Write-Host "⚠️  Удалён битый объект из archive" -ForegroundColor Yellow; $changed = $true }
    }
    if ($cleanArchive.Count -ne @($db.projects.archive).Count) { $db.projects.archive = $cleanArchive }

    foreach ($r in $db.requests) {
        $defaultProps = @{
            id = $null; project = $null; createdDate = (Get-Date).ToString("yyyy-MM-dd")
            deadline = $null; body = $null; status = "Отправлено на рассмотрение"
            author = $null; authorDeleted = $false; authorDeletedAt = $null
            priority = "medium"; takenBy = $null; comment = $null; audit = @()
            lastDeadlineNotify = $null; budget = $null; files = @(); isArchived = $false
        }
        foreach ($prop in $defaultProps.Keys) {
            if ($r.PSObject.Properties.Name -notcontains $prop) {
                $r | Add-Member -NotePropertyName $prop -NotePropertyValue $defaultProps[$prop] -Force
                $changed = $true
            }
        }
        
        if ($r.PSObject.Properties.Name -notcontains 'isArchived') {
            $isInArchive = $db.projects.archive | Where-Object { $_.name -eq $r.project }
            $r | Add-Member -NotePropertyName 'isArchived' -NotePropertyValue ($null -ne $isInArchive) -Force
            $changed = $true
        }
        
        $cleanFiles = @()
        foreach ($file in @($r.files)) {
            if ($file -and $file -is [PSCustomObject] -and $file.path -and $file.name) { $cleanFiles += $file }
            else { Write-Host "⚠️  Удалён битый файл из заявки $($r.id)" -ForegroundColor Yellow; $changed = $true }
        }
        if ($cleanFiles.Count -ne @($r.files).Count) { $r.files = $cleanFiles }
        
        if ($r.authorDeleted -eq $true -and $r.author) {
            $activeUser = $db.auth.users | Where-Object { $_.name -eq $r.author -and $_.deleted -eq $false } | Select-Object -First 1
            if ($activeUser) { $r.authorDeleted = $false; $r.authorDeletedAt = $null; $changed = $true; Write-Host "✅ Сброшен authorDeleted у $($r.id)" -ForegroundColor Green }
        }
        
        if ($r.status -eq "В работу") { $r.status = "Выполняется"; $changed = $true }
        if ($r.audit.Count -eq 0) {
            $auditEntry = [PSCustomObject]@{ timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); actor = $r.author; action = "Создание заявки"; newStatus = $r.status; comment = " " }
            $r.audit = @($auditEntry); $changed = $true
        }
    }

    if ($changed) { Set-Db $db; Write-Host "✅ Миграция завершена" -ForegroundColor Green }
    else { Write-Host "✅ Структура в порядке" -ForegroundColor Green }
}

# ==========================================================
# 🔐 БЛОК 4: AUTH & USERS
# ==========================================================
function New-Code { (100000..999999 | Get-Random).ToString() }
function New-SessionToken { [guid]::NewGuid().ToString("N") }

function Auth-Register($username, $password, $role, $name) {
    $db = Get-Db
    $email = $username.ToLower().Trim()
    if (-not $email.EndsWith('@stroisservis.ru')) { return @{ok=$false; error="Только @stroisservis.ru"} }
    $existing = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $email -and $_.deleted -eq $false }
    if ($existing) { return @{ok=$false; error="Email занят"} }
    $code = New-Code
    $salt = [guid]::NewGuid().ToString("N")
    $hash = Get-PasswordHash $password $salt
    $newUser = [PSCustomObject]@{
        username=$email; passwordHash=$hash; salt=$salt; password=$null
        role=$role; name=$name.Trim(); verified=$false
        verificationCode=$code; codeExpires=(Get-Date).AddMinutes(15).ToString("o")
        readRequests=@(); deleted=$false; deletedAt=$null; recoveryCode=$null; recoveryExpires=$null
    }
    $db.auth.users = @($db.auth.users) + @($newUser)
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
    $u = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $username.Trim() -and $_.deleted -eq $false }
    if (-not $u) { return @{ok=$false; error="Неверные данные или пользователь удалён"} }
    if (-not $u.verified) { return @{ok=$false; error="Подтвердите почту"} }
    $hash = Get-PasswordHash $password $u.salt
    if ($u.passwordHash -ne $hash) { return @{ok=$false; error="Неверный пароль"} }
    $token = New-SessionToken
    $script:Sessions[$token] = @{ username = $u.username; role = $u.role; name = $u.name; createdAt = (Get-Date) }
    return @{ok=$true; token=$token; user=@{username=$u.username; role=$u.role; name=$u.name}}
}

function Get-UserByToken($token) {
    if (-not $token) { return $null }
    $sess = $script:Sessions[$token]
    if (-not $sess) { return $null }
    if (((Get-Date) - $sess.createdAt).TotalHours -gt 24) { $script:Sessions.Remove($token); return $null }
    return (Get-User $sess.username)
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

function Auth-RecoverRequest($username) {
    $db = Get-Db
    $u = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $username.Trim() -and $_.verified -and $_.deleted -eq $false }
    if (-not $u) { return @{ok=$true} }
    $code = New-Code
    $u.recoveryCode = $code; $u.recoveryExpires = (Get-Date).AddMinutes(15).ToString("o")
    Set-Db $db; Send-RecoveryEmail $u.username $code
    return @{ok=$true}
}

function Auth-RecoverReset($username, $code, $newPassword) {
    $db = Get-Db
    $u = $db.auth.users | Where-Object { $_ -and $_.username -and $_.username.Trim() -eq $username.Trim() -and $_.verified -and $_.deleted -eq $false }
    if (-not $u) { return @{ok=$false; error="Не найден"} }
    if (-not $u.recoveryCode -or $u.recoveryCode -ne $code) { return @{ok=$false; error="Неверный код"} }
    if ($u.recoveryExpires -and [DateTime]::Parse($u.recoveryExpires) -lt (Get-Date)) { return @{ok=$false; error="Код истёк"} }
    $salt = [guid]::NewGuid().ToString("N")
    $u.salt = $salt; $u.passwordHash = Get-PasswordHash $newPassword $salt
    $u.recoveryCode = $null; $u.recoveryExpires = $null
    Set-Db $db
    return @{ok=$true}
}

function Get-UserDisplayName($nameOrUsername, $db) {
    if (-not $nameOrUsername) { return "[Неизвестно]" }
    if ($nameOrUsername -match '\(Удалён') { return $nameOrUsername }
    $activeUser = $db.auth.users | Where-Object { $_.name -eq $nameOrUsername -and $_.deleted -eq $false } | Select-Object -First 1
    if (-not $activeUser) { $activeUser = $db.auth.users | Where-Object { $_.username -eq $nameOrUsername -and $_.deleted -eq $false } | Select-Object -First 1 }
    if ($activeUser) { return $activeUser.name }
    $deletedUser = $db.auth.users | Where-Object { $_.name -eq $nameOrUsername -and $_.deleted -eq $true } | Select-Object -First 1
    if ($deletedUser) { return "$nameOrUsername (Удалён $(if($deletedUser.deletedAt){$deletedUser.deletedAt}else{'дата неизвестна'}))" }
    return $nameOrUsername
}

function Action-DeleteUser($username, $actorRole) {
    if ($actorRole -ne "director") { return @{ok=$false; error="Только руководитель"} }
    $db = Get-Db
    $user = $db.auth.users | Where-Object { $_.username -eq $username -and $_.deleted -eq $false }
    if (-not $user) { return @{ok=$false; error="Не найден"} }
    $user.deleted = $true; $user.deletedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    foreach ($r in $db.requests) {
        if ($r.author -eq $user.name) { $r.authorDeleted = $true; $r.authorDeletedAt = $user.deletedAt; AddAuditEntry $r "Система" $r.status "Автор удалён" "Автор удалён" }
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
function Send-VerificationEmail($email, $code) { Send-Mail $email "Форма заявки: Код подтверждения" "<div style='font-family:Roboto,sans-serif;text-align:center'><h2>🔐 Код</h2><p><b style='font-size:2em;color:#1a73e8'>$code</b></p><p>Действует 15 минут.</p></div>" }
function Send-RecoveryEmail($email, $code) { Send-Mail $email "Форма заявки: Восстановление" "<div style='font-family:Roboto,sans-serif;text-align:center'><h2>🔑 Код сброса</h2><p><b style='font-size:2em;color:#d93025'>$code</b></p><p>Действует 15 минут.</p></div>" }

function Notify-SupplierNew($req) {
    $db = Get-Db
    foreach($s in ($db.auth.users | Where-Object { $_.role -eq "procurement" -and $_.verified -and $_.deleted -eq $false })) {
        if ($s.name -eq $req.author) { continue }
        Send-Mail $s.username "Новая заявка $($req.id)" "<p>Новая заявка <b>$($req.id)</b> для проекта <b>$($req.project)</b>.</p><a href='http://localhost:$port'>Перейти</a>"
    }
}

function Notify-AllParticipants($req, $newStatus, $actor, $comment, $eventType = "status_change") {
    $db = Get-Db
    $authorDisplay = Get-UserDisplayName $req.author $db
    $actorDisplay  = Get-UserDisplayName $actor $db
    switch ($eventType) {
        "status_change" {
            switch ($newStatus) {
                "Выполняется" {
                    $directors = $db.auth.users | Where-Object { $_.role -eq "director" -and $_.verified -and $_.deleted -eq $false -and $_.name -ne $actor }
                    foreach ($d in $directors) {
                        Send-Mail $d.username "⬆ Заявка $($req.id) ожидает визирования" "<p><b>$actorDisplay</b> принял в работу заявку <b>$($req.id)</b>.</p><p><b>Заявка ожидает вашего визирования.</b></p><p><a href='http://localhost:$port'>Перейти</a></p>"
                    }
                    $authorUser = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }
                    if ($authorUser -and $authorUser.name -ne $actor) {
                        Send-Mail $authorUser.username "✅ Ваша заявка $($req.id) принята в работу" "<p>✅ <b>Ваша заявка $($req.id)</b> принята в работу.</p><p><b>Обработчик:</b> $actorDisplay</p><p><a href='http://localhost:$port'>Открыть</a></p>"
                    }
                }
                { $_ -in @("Оплачено","Договорённость","Отклонено","Приостановлено","Возвращено на доработку") } {
                    $emoji = switch($newStatus) { "Оплачено"{"💰"} "Договорённость"{"🤝"} "Отклонено"{"❌"} "Приостановлено"{"⏸"} "Возвращено на доработку"{"🔁"} default{"📋"} }
                    $recipients = @()
                    $au = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }; if ($au) { $recipients += $au.username }
                    if ($req.takenBy) { $eu = $db.auth.users | Where-Object { $_.name -eq $req.takenBy -and $_.deleted -eq $false }; if ($eu) { $recipients += $eu.username } }
                    $db.auth.users | Where-Object { $_.role -eq "director" -and $_.verified -and $_.deleted -eq $false -and $_.name -ne $actor } | ForEach-Object { $recipients += $_.username }
                    $db.auth.users | Where-Object { $_.role -eq "procurement" -and $_.verified -and $_.deleted -eq $false -and $_.name -ne $actor } | ForEach-Object { $recipients += $_.username }
                    foreach ($email in ($recipients | Select-Object -Unique)) {
                        Send-Mail $email "$emoji Заявка $($req.id) — $newStatus" "<p>$emoji <b>$actorDisplay</b> изменил статус на <b>$newStatus</b>.</p><p><a href='http://localhost:$port'>Открыть</a></p>"
                    }
                }
            }
        }
    }
}

# ==========================================================
# 💼 БЛОК 6: BUSINESS LOGIC
# ==========================================================
function AddAuditEntry($req, $actor, $newStatus, $comment, $action = "Изменение статуса") {
    $entry = [PSCustomObject]@{ timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); actor = $actor; action = $action; newStatus = $newStatus; comment = $comment }
    $req.audit = @($req.audit) + @($entry)
}

function CheckValidStatusTransition($currentStatus, $newStatus, $role, $db) {
    $cs = if ($currentStatus -is [string]) { $currentStatus.Trim() } else { "" }
    $ns = if ($newStatus -is [string]) { $newStatus.Trim() } else { "" }
    $rl = if ($role -is [string]) { $role.Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($rl)) { return $false, "Роль не определена" }
    $sc = $db.system.statuses | Where-Object { $_.name -and $_.name.Trim() -eq $ns }
    if (-not $sc) { return $false, "Статус '$ns' не существует" }
    if ($cs -eq $ns) { return $false, "Статус не изменён" }
    $af = @($sc.availableFrom | ForEach-Object { if ($_ -is [string]) { $_.Trim() } })
    if ($af.Count -gt 0 -and $af -notcontains $cs) { return $false, "Невозможно перейти из '$cs' в '$ns'" }
    $afor = @($sc.availableFor | ForEach-Object { if ($_ -is [string]) { $_.Trim() } })
    if ($afor -notcontains $rl) { return $false, "Нет прав для статуса '$ns'" }
    return $true, "OK"
}

function Action-CreateRequest($proj, $dead, $body, $author, $priority = "medium", $budget = $null, $authorRole = $null) {
    $db = Get-Db
    if (-not ($db.projects.active | Where-Object { $_.name -eq $proj })) { return @{ok=$false; error="Проект не существует"} }
    $id = "REQ-" + ($db.requests.Count+1).ToString("D4")
    $req = [PSCustomObject]@{ id=$id; project=$proj; createdDate=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); deadline=$dead; body=$body; status="Отправлено на рассмотрение"; author=$author; authorDeleted=$false; authorDeletedAt=$null; priority=$priority; takenBy=$null; comment=$null; audit=@(); lastDeadlineNotify=$null; budget=$budget; files=@(); isArchived=$false }
    AddAuditEntry $req $author "Отправлено на рассмотрение" " " "Создание заявки"
    $db.requests = @($db.requests) + @($req)
    Set-Db $db
    Notify-SupplierNew $req
    if ($authorRole -eq "procurement") {
        $db.auth.users | Where-Object { $_.role -eq "director" -and $_.verified -and $_.deleted -eq $false } | ForEach-Object { Send-Mail $_.username "📋 Новая заявка от снабжения: $id" "<p>Снабжение создало заявку <b>$id</b>.</p><a href='http://localhost:$port'>Перейти</a>" }
    }
    return @{ok=$true; id=$id}
}

function Action-EditRequest($id, $newBody, $newDead, $newPriority, $actor, $budget) {
    $db = Get-Db; $r = $db.requests | Where-Object {$_.id -eq $id}; if (-not $r) { return @{ok=$false; error="Не найдено"} }
    $user = $db.auth.users | Where-Object { $_.name -eq $actor }; if (-not $user) { return @{ok=$false; error="Пользователь не найден"} }
    if ($r.author -eq $actor -and $r.status -ne "Отправлено на рассмотрение") { return @{ok=$false; error="Редактирование невозможно"} }
    if ($user.role -eq "client" -and $r.author -ne $actor) { return @{ok=$false; error="Нет прав"} }
    $changes = @()
    if ($r.body -ne $newBody) { $changes += "Описание"; $r.body = $newBody }
    if ($r.deadline -ne $newDead) { $changes += "Срок"; $r.deadline = $newDead }
    if ($r.priority -ne $newPriority) { $changes += "Приоритет"; $r.priority = $newPriority }
    if ($budget -ne $null -and $r.budget -ne $budget) { $changes += "Бюджет"; $r.budget = $budget }
    if ($changes.Count -gt 0) { AddAuditEntry $r $actor $r.status ("Изменено: " + ($changes -join ", ")) "Редактирование"; Set-Db $db; Notify-AllParticipants $r $r.status $actor ($changes -join ", ") "edit_request" }
    return @{ok=$true}
}

function Action-Process($id, $status, $actor, $userRole, $comment="") {
    $db = Get-Db; $r = $db.requests | Where-Object {$_.id -eq $id}; if (-not $r) { return @{ok=$false; error="Не найдено"} }
    if ([string]::IsNullOrWhiteSpace($userRole)) { $fu = $db.auth.users | Where-Object { $_.name -eq $actor } | Select-Object -First 1; if ($fu) { $userRole = $fu.role } else { return @{ok=$false; error="Роль не определена"} } }
    $valid, $err = CheckValidStatusTransition $r.status $status $userRole $db; if (-not $valid) { return @{ok=$false; error=$err} }
    $r.status=$status; if ($status -eq "Выполняется" -and -not $r.takenBy) { $r.takenBy=$actor }; $r.comment=$comment
    AddAuditEntry $r $actor $status $comment "Изменение статуса"; Set-Db $db; Notify-AllParticipants $r $status $actor $comment "status_change"
    return @{ok=$true}
}

function Action-BulkProcess($ids, $status, $actor, $userRole, $comment="") { $results = @(); foreach ($id in $ids) { $res = Action-Process $id $status $actor $userRole $comment; $results += @{ id=$id; ok=$res.ok; error=$res.error } }; return @{ok=$true; results=$results} }
function Action-Reject($id, $reason, $actor, $userRole) { return Action-Process $id "Отклонено" $actor $userRole $reason }

function Action-Archive($name, $role) {
    if ($role -ne "director") { return @{ok=$false; error="Только руководитель"} }
    $db = Get-Db; $proj = $db.projects.active | Where-Object { $_.name -eq $name.Trim() }
    if ($proj) { foreach ($r in $db.requests) { if ($r.project -eq $name.Trim()) { $r | Add-Member -NotePropertyName 'isArchived' -NotePropertyValue $true -Force } }; $db.projects.active = @($db.projects.active | Where-Object { $_.name -ne $name.Trim() }); $db.projects.archive = @($db.projects.archive) + @($proj); Set-Db $db; return @{ok=$true} }
    return @{ok=$false; error="Проект не найден"}
}

function Action-DeleteFromArchive($name, $role) { if ($role -ne "director") { return @{ok=$false; error="Только руководитель"} }; $db = Get-Db; $db.projects.archive = @($db.projects.archive | Where-Object { $_.name -ne $name.Trim() }); Set-Db $db; return @{ok=$true; message="Проект удалён"} }
function Action-CreateProject($name, $deadline, $role) { if ($role -notin @('director','procurement')) { return @{ok=$false; error="Недостаточно прав"} }; $c = $name.Trim(); if (-not $c) { return @{ok=$false; error="Пустое имя"} }; if (-not $deadline) { return @{ok=$false; error="Не указан срок"} }; $db = Get-Db; if (($db.projects.active | Where-Object { $_.name -eq $c }) -or ($db.projects.archive | Where-Object { $_.name -eq $c })) { return @{ok=$false; error="Проект уже существует"} }; $db.projects.active = @($db.projects.active) + @([PSCustomObject]@{ name = $c; deadline = $deadline }); Set-Db $db; return @{ok=$true} }
function Action-UpdateProjectDeadline($name, $newDeadline, $role) { if ($role -ne "director") { return @{ok=$false; error="Только руководитель"} }; $db = Get-Db; $p = $db.projects.active | Where-Object { $_.name -eq $name }; if (-not $p) { return @{ok=$false; error="Не найден"} }; $p.deadline = $newDeadline; Set-Db $db; return @{ok=$true} }

function Action-UploadFile($requestId, $fileName, $fileBytes, $actor) {
    if (-not $fileBytes -or $fileBytes.Length -eq 0) { return @{ok=$false; error="Файл пустой"} }
    $db = Get-Db; $r = $db.requests | Where-Object { $_.id -eq $requestId }; if (-not $r) { return @{ok=$false; error="Заявка не найдена"} }
    $safeName = $fileName -replace '[^\w\.\-а-яА-ЯёЁ]', '_'
    $uniqueName = "${requestId}_$(Get-Date -Format 'yyyyMMddHHmmss')_$safeName"
    if ($uniqueName.Length -gt 200) { $ext = [System.IO.Path]::GetExtension($safeName); $uniqueName = $uniqueName.Substring(0, 195 - $ext.Length) + $ext }
    $filePath = Join-Path $uploadsDir $uniqueName
    $counter = 1; $baseName = [System.IO.Path]::GetFileNameWithoutExtension($uniqueName); $extension = [System.IO.Path]::GetExtension($uniqueName)
    while (Test-Path $filePath) { $uniqueName = "${baseName}_${counter}${extension}"; $filePath = Join-Path $uploadsDir $uniqueName; $counter++ }
    [System.IO.File]::WriteAllBytes($filePath, $fileBytes)
    $r.files = @($r.files) + @([PSCustomObject]@{ name=$fileName; path=$uniqueName; uploadedBy=$actor; uploadedAt=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") })
    AddAuditEntry $r $actor $r.status "Загружен файл: $fileName" "Загрузка файла"; Set-Db $db
    Notify-AllParticipants $r $r.status $actor $fileName "file_uploaded"
    return @{ok=$true; file=$uniqueName}
}

function MarkRequestAsRead($username, $requestId) { $db = Get-Db; $u = $db.auth.users | Where-Object { $_.username -eq $username }; if (-not $u) { return $false }; if ($u.readRequests -notcontains $requestId) { $u.readRequests = @($u.readRequests) + @($requestId); Set-Db $db }; return $true }
function GetUnreadRequestsCount($username) { $db = Get-Db; $u = $db.auth.users | Where-Object { $_.username -eq $username }; if (-not $u -or $u.role -ne "procurement") { return 0 }; return @($db.requests | Where-Object { $_.status -eq "Отправлено на рассмотрение" -and $u.readRequests -notcontains $_.id }).Count }

# ==========================================================
# 📊 БЛОК 6.1: ДАШБОРД
# ==========================================================
function Get-DashboardData($role) {
    if ($role -ne "director") { return @{error="Нет доступа"} }
    $db = Get-Db; $now = Get-Date; $monthAgo = $now.AddMonths(-1)
    $total = $db.requests.Count
    $thisMonth = @($db.requests | Where-Object { $_.createdDate -and [DateTime]::Parse($_.createdDate) -gt $monthAgo }).Count
    $byStatus = @{}; foreach ($r in $db.requests) { if ($byStatus.ContainsKey($r.status)) { $byStatus[$r.status]++ } else { $byStatus[$r.status] = 1 } }
    $overdue = @($db.requests | Where-Object { $_.deadline -and $_.status -notin @("Оплачено","Договорённость","Отклонено") -and [DateTime]::Parse($_.deadline) -lt $now }).Count
    $totalBudget = ($db.requests | Where-Object { $_.budget } | Measure-Object -Property budget -Sum).Sum; if (-not $totalBudget) { $totalBudget = 0 }
    return @{ total = $total; thisMonth = $thisMonth; byStatus = $byStatus; overdue = $overdue; totalBudget = $totalBudget; activeProjects = $db.projects.active.Count; archivedProjects = $db.projects.archive.Count; users = @($db.auth.users | Where-Object { $_.deleted -eq $false }).Count }
}

# ==========================================================
# 🌐 БЛОК 7: БРАУЗЕРНЫЕ УВЕДОМЛЕНИЯ
# ==========================================================
function Get-UserUpdates($username, $lastCheck) {
    $db = Get-Db; $user = $db.auth.users | Where-Object { $_.username -eq $username -and $_.deleted -eq $false }
    if (-not $user) { return @{ error = "Пользователь не найден" } }
    $lastCheckDate = if ($lastCheck -and $lastCheck -ne "0" -and $lastCheck -ne "null") { try { [DateTime]::Parse($lastCheck).ToUniversalTime() } catch { (Get-Date).AddHours(-1).ToUniversalTime() } } else { (Get-Date).AddHours(-1).ToUniversalTime() }
    $updates = @()
    $db.requests | Where-Object { $_.status -eq "Отправлено на рассмотрение" -and $_.createdDate -and [DateTime]::Parse($_.createdDate).ToUniversalTime() -gt $lastCheckDate } | ForEach-Object {
        if ($user.role -eq "procurement" -and $_.author -ne $user.name) { $updates += @{ type="new_request"; title="📋 Новая заявка"; message="Заявка $($_.id) от $($_.author)"; requestId=$_.id; project=$_.project; priority=$_.priority } }
        if ($user.role -eq "director") { $updates += @{ type="new_request_any"; title="📋 Новая заявка"; message="$($_.id) от $($_.author) — $($_.project)"; requestId=$_.id; project=$_.project; priority=$_.priority } }
    }
    foreach ($req in $db.requests) {
        if ($req.audit -and $req.audit.Count -gt 0) {
            try {
                $at = [DateTime]::Parse($req.audit[-1].timestamp).ToUniversalTime()
                if ($at -gt $lastCheckDate) {
                    $see = ($req.author -eq $user.name) -or ($req.takenBy -eq $user.name) -or ($user.role -eq "director" -and $req.audit[-1].actor -ne $user.name) -or ($user.role -eq "procurement" -and $req.audit[-1].actor -ne $user.name)
                    if ($see) { $pi = switch($req.priority) { "high"{"🔴"} "medium"{"🟡"} default{"🟢"} }; $updates += @{ type="status_change"; title="🔄 $pi $($req.id)"; message="$($req.status) — $($req.audit[-1].actor)"; requestId=$req.id; project=$req.project; newStatus=$req.status; actor=$req.audit[-1].actor } }
                }
            } catch {}
        }
    }
    $today = (Get-Date).Date
    foreach ($req in $db.requests) {
        if ($req.deadline -and $req.status -notin @("Оплачено","Договорённость","Отклонено")) {
            try {
                $dl = ([DateTime]::Parse($req.deadline).Date - $today).Days
                if ($dl -ge 0 -and $dl -le 3 -and ((-not $req.lastDeadlineNotify) -or ($req.lastDeadlineNotify -ne $today.ToString("yyyy-MM-dd")))) {
                    $ru = @(); $au = $db.auth.users | Where-Object { $_.name -eq $req.author -and $_.deleted -eq $false }; if ($au) { $ru += $au.username }
                    if ($req.takenBy) { $eu = $db.auth.users | Where-Object { $_.name -eq $req.takenBy -and $_.deleted -eq $false }; if ($eu) { $ru += $eu.username } }
                    if ($user.role -in @("director","procurement")) { $ru += $username }
                    if (($ru | Select-Object -Unique) -contains $username) { $updates += @{ type="deadline_warning"; title="⏰ Срок истекает"; message="Заявка $($req.id): осталось $dl дня(ей)"; requestId=$req.id; project=$req.project; daysLeft=$dl }; $req.lastDeadlineNotify = $today.ToString("yyyy-MM-dd"); Set-Db $db }
                }
            } catch {}
        }
    }
    return @{ updates = $updates; count = $updates.Count }
}

# ==========================================================
# 📁 БЛОК 8: ЭКСПОРТ EXCEL
# ==========================================================
$lastExp = [DateTime]::MinValue
function Export-Excel($force = $false) {
    if (-not $force -and ((Get-Date).Hour -ne 23 -or $lastExp.Date -eq (Get-Date).Date)) { return }
    try {
        $reqs = (Get-Db).requests; if (-not $reqs -or $reqs.Count -eq 0) { return }
        $xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
        $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets(1)
        @("ID","Проект","Дата","Срок","Статус","Приоритет","Бюджет","Описание","Автор","Файлы") | ForEach-Object -Begin {$i=0} -Process { $i++; $ws.Cells(1,$i).Value2 = $_; $ws.Cells(1,$i).Font.Bold = $true }
        for ($row = 0; $row -lt $reqs.Count; $row++) {
            $r = $reqs[$row]; $ad = if ($r.authorDeleted) { "$($r.author) (Удалён)" } else { $r.author }
            $fl = if ($r.files -and $r.files.Count -gt 0) { ($r.files | Where-Object { $_ -and $_.name } | ForEach-Object { $_.name }) -join ", " } else { "" }
            $ws.Cells($row+2,1).Value2 = $r.id; $ws.Cells($row+2,2).Value2 = $r.project; $ws.Cells($row+2,3).Value2 = $r.createdDate; $ws.Cells($row+2,4).Value2 = $r.deadline
            $ws.Cells($row+2,5).Value2 = $r.status; $ws.Cells($row+2,6).Value2 = $r.priority; $ws.Cells($row+2,7).Value2 = $r.budget; $ws.Cells($row+2,8).Value2 = $r.body; $ws.Cells($row+2,9).Value2 = $ad; $ws.Cells($row+2,10).Value2 = $fl
        }
        $ws.Columns.AutoFit() | Out-Null; $wb.SaveAs($excelPath); $wb.Close($false); $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null; $script:lastExp = Get-Date
    } catch { Write-Host "[EXCEL] ❌ $_" -ForegroundColor Red }
}

# ==========================================================
# 🎨 БЛОК 9: UI (ИСПРАВЛЕННЫЙ GANTT + FULL APP)
# ==========================================================
$ui = @'
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Форма заявки 3.7.9</title>
<link rel="icon" href="data:,">
<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#f5f5f5;--card:#fff;--text:#202124;--text-secondary:#5f6368;--border:#dadce0;--border-light:#e0e0e0;--primary:#1a73e8;--primary-dark:#1557b0;--danger:#d93025;--success:#137333;--warning:#e37400;--gray:#5f6368;--gray-light:#f8f9fa;--shadow:0 1px 2px 0 rgba(60,64,67,.3),0 1px 3px 1px rgba(60,64,67,.15);--shadow-hover:0 4px 8px rgba(0,0,0,.1)}
body.dark{--bg:#202124;--card:#2d2e32;--text:#e8eaed;--text-secondary:#9aa0a6;--border:#5f6368;--border-light:#3c4043;--primary:#8ab4f8;--primary-dark:#aecbfa;--danger:#f28b82;--success:#81c995;--warning:#fdd663;--gray:#9aa0a6;--gray-light:#3c4043;--shadow:0 1px 2px 0 rgba(0,0,0,.3),0 1px 3px 1px rgba(0,0,0,.15);--shadow-hover:0 4px 8px rgba(0,0,0,.3)}
body{font-family:'Roboto',sans-serif;background:var(--bg);color:var(--text);padding:20px;line-height:1.5;transition:background .3s,color .3s}
.auth,header,.card,.modal-c,.table-wrapper,.dashboard-card,.gantt-card{background:var(--card);border-radius:12px;box-shadow:var(--shadow);transition:background .3s,box-shadow .3s}
.auth{max-width:420px;margin:40px auto;padding:32px}
header{padding:16px 24px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px}
h1{font-size:1.4rem;font-weight:500;color:var(--primary)}
h2,h3{font-weight:500;margin-bottom:16px;color:var(--text)}
.tabs{display:flex;gap:8px;border-bottom:1px solid var(--border);margin-bottom:24px}
.tab{flex:1;text-align:center;padding:12px 0;font-weight:500;cursor:pointer;color:var(--text-secondary);border-bottom:2px solid transparent;transition:all .2s}
.tab.active{color:var(--primary);border-bottom-color:var(--primary)}
input,select,textarea{width:100%;padding:12px;margin:6px 0 14px;border:1px solid var(--border);border-radius:8px;font-family:'Roboto',sans-serif;font-size:14px;background:var(--card);color:var(--text);transition:all .2s}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 2px rgba(26,115,232,.2)}
label{font-weight:500;font-size:13px;color:var(--text-secondary)}
.btn{padding:9px 18px;border:none;border-radius:8px;font-weight:500;font-size:13px;cursor:pointer;transition:all .2s;background:var(--gray-light);color:var(--text);display:inline-flex;align-items:center;gap:6px}
.btn:hover{transform:translateY(-1px);box-shadow:var(--shadow-hover)}
.btn-pri{background:var(--primary);color:#fff}.btn-pri:hover{background:var(--primary-dark)}
.btn-dan{background:var(--danger);color:#fff}.btn-suc{background:var(--success);color:#fff}.btn-wrn{background:var(--warning);color:#202124}.btn-gra{background:var(--gray);color:#fff}
.btn-sm{padding:6px 10px;font-size:12px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:20px;margin:20px 0}
.card{padding:20px;cursor:pointer;transition:all .2s;border:1px solid var(--border-light)}
.card:hover{transform:translateY(-2px);box-shadow:var(--shadow-hover);border-color:var(--primary)}
.card-new{background:var(--gray-light);border:2px dashed var(--border);display:flex;align-items:center;justify-content:center;min-height:100px}
.card-new h3{margin:0;color:var(--primary)}
.table-wrapper{overflow-x:auto;margin:20px 0;border-radius:12px;border:1px solid var(--border-light)}
table{width:100%;border-collapse:collapse;min-width:900px}
thead{background:var(--gray-light)}
th{padding:14px 12px;font-size:12px;font-weight:600;color:var(--text-secondary);border-bottom:2px solid var(--border-light);text-align:left;text-transform:uppercase;letter-spacing:.5px}
td{padding:12px;border-bottom:1px solid var(--border-light);font-size:14px;color:var(--text);vertical-align:middle}
tr:hover td{background:rgba(26,115,232,.04)}
tr.overdue td{background:rgba(217,48,37,.06);animation:pulse-red 2s infinite}
tr.highlight td{background:rgba(26,115,232,.15);transition:background .5s}
@keyframes pulse-red{0%,100%{background:rgba(217,48,37,.06)}50%{background:rgba(217,48,37,.15)}}
td .btn{margin-right:4px;margin-bottom:4px;white-space:nowrap}
.st{display:inline-block;padding:4px 12px;border-radius:16px;font-size:12px;font-weight:500}
.st-new{background:rgba(26,115,232,.12);color:var(--primary)}.st-wrk{background:rgba(227,116,0,.12);color:var(--warning)}.st-ok{background:rgba(19,115,51,.12);color:var(--success)}.st-no{background:rgba(217,48,37,.12);color:var(--danger)}.st-pause{background:rgba(95,99,104,.12);color:var(--gray)}.st-back{background:rgba(162,89,247,.12);color:#a259f7}
.priority-low{color:var(--success);font-weight:500}.priority-medium{color:var(--warning);font-weight:500}.priority-high{color:var(--danger);font-weight:600}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);backdrop-filter:blur(4px);justify-content:center;align-items:center;z-index:1000}
.modal-c{max-width:600px;width:92%;padding:24px;border-radius:16px;max-height:90vh;overflow-y:auto;background:var(--card)}
.modal-buttons{display:flex;gap:12px;margin-top:20px;flex-wrap:wrap;justify-content:flex-end}
.back{color:var(--primary);cursor:pointer;margin-bottom:16px;display:inline-block;font-weight:500}.back:hover{text-decoration:underline}
.hidden{display:none!important}
.uinfo{background:rgba(26,115,232,.1);color:var(--primary);padding:6px 12px;border-radius:20px;font-size:13px;font-weight:500}
.unread-badge{background:var(--danger);color:#fff;border-radius:20px;padding:2px 8px;font-size:11px;font-weight:600}
.notification-bell{position:relative;cursor:pointer;background:var(--gray-light);border:none;border-radius:40px;padding:8px 14px;font-size:18px;transition:all .2s;color:var(--text)}
.notification-bell:hover{transform:scale(1.05);background:var(--border-light)}
.notification-badge{position:absolute;top:-5px;right:-5px;background:var(--danger);color:#fff;border-radius:20px;padding:2px 6px;font-size:10px;font-weight:bold;min-width:18px;text-align:center}
.notification-dropdown{position:absolute;top:50px;right:20px;width:360px;max-height:450px;overflow-y:auto;background:var(--card);border-radius:12px;box-shadow:var(--shadow-hover);z-index:1000;display:none;border:1px solid var(--border-light)}
.notification-dropdown.show{display:block}
.notification-item{padding:12px 16px;border-bottom:1px solid var(--border-light);cursor:pointer;transition:background .2s}
.notification-item:hover{background:var(--gray-light)}
.notification-item.unread{background:rgba(26,115,232,.06);border-left:3px solid var(--primary)}
.notification-title{font-weight:500;font-size:13px;color:var(--text)}
.notification-message{font-size:12px;color:var(--text-secondary);margin-top:4px}
.notification-time{font-size:10px;color:var(--text-secondary);margin-top:4px}
.warning{background:rgba(227,116,0,.12);padding:3px 8px;border-radius:12px;font-size:11px;color:var(--warning);margin-left:6px;display:inline-block;font-weight:500}
.overdue-badge{background:var(--danger);color:#fff;padding:3px 8px;border-radius:12px;font-size:11px;margin-left:6px;display:inline-block;font-weight:500;animation:pulse-red 2s infinite}
.arch-c{opacity:.7}
.deleted-user{color:var(--text-secondary);font-style:italic;font-size:11px}
.theme-toggle{background:var(--gray-light);border:none;border-radius:40px;padding:8px 14px;cursor:pointer;font-size:16px;transition:all .2s;color:var(--text)}
.theme-toggle:hover{transform:scale(1.05)}
.timeline{position:relative;padding-left:28px;max-height:400px;overflow-y:auto}
.timeline::before{content:'';position:absolute;left:10px;top:0;bottom:0;width:2px;background:var(--border)}
.timeline-item{position:relative;padding:10px 0 16px 0}
.timeline-item::before{content:'';position:absolute;left:-22px;top:14px;width:12px;height:12px;border-radius:50%;background:var(--primary);border:2px solid var(--card);box-shadow:0 0 0 2px var(--primary)}
.timeline-item.status-change::before{background:var(--warning)}.timeline-item.edit::before{background:var(--primary)}.timeline-item.create::before{background:var(--success)}.timeline-item.reject::before{background:var(--danger)}
.timeline-time{font-size:11px;color:var(--text-secondary);font-weight:500}
.timeline-actor{font-size:12px;font-weight:600;color:var(--text);margin-top:2px}
.timeline-action{font-size:13px;color:var(--text);margin-top:4px}
.timeline-comment{font-size:12px;color:var(--text-secondary);margin-top:4px;padding:8px;background:var(--gray-light);border-radius:8px;border-left:3px solid var(--primary)}
.dashboard-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:24px}
.dashboard-card{padding:20px;text-align:center}
.dashboard-card .big-number{font-size:2.5rem;font-weight:700;color:var(--primary);margin:8px 0}
.dashboard-card .label{font-size:13px;color:var(--text-secondary);font-weight:500}
.dashboard-card.danger .big-number{color:var(--danger)}.dashboard-card.success .big-number{color:var(--success)}.dashboard-card.warning .big-number{color:var(--warning)}
.search-bar{display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;align-items:center}
.search-bar input{flex:1;min-width:200px;margin:0}.search-bar select{width:auto;min-width:140px;margin:0}
.file-list{margin-top:10px;display:flex;flex-wrap:wrap;gap:8px}
.file-item{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;background:var(--gray-light);border-radius:8px;font-size:12px;border:1px solid var(--border-light);cursor:pointer;transition:all .2s;text-decoration:none;color:var(--text)}
.file-item:hover{background:var(--primary);color:#fff;border-color:var(--primary)}
.file-upload-area{border:2px dashed var(--border);border-radius:12px;padding:20px;text-align:center;cursor:pointer;transition:all .2s;margin:10px 0}
.file-upload-area:hover{border-color:var(--primary);background:rgba(26,115,232,.04)}
.file-upload-area.dragover{border-color:var(--primary);background:rgba(26,115,232,.1)}
.mass-checkbox{width:18px;height:18px;cursor:pointer;accent-color:var(--primary)}
.mass-actions-bar{position:sticky;bottom:0;background:var(--primary);color:#fff;padding:12px 20px;border-radius:12px;margin-top:16px;display:flex;justify-content:space-between;align-items:center;gap:12px;box-shadow:0 -2px 10px rgba(0,0,0,.2);z-index:100}
.mass-actions-bar .btn{background:rgba(255,255,255,.2);color:#fff;border:1px solid rgba(255,255,255,.3)}
.mass-actions-bar .btn:hover{background:rgba(255,255,255,.3)}
.project-link{color:var(--primary);cursor:pointer;text-decoration:underline}.project-link:hover{color:var(--primary-dark)}

/* ===== GANTT STYLES ===== */
.gantt-card{padding:24px;margin-bottom:20px;position:relative;overflow:hidden}
.gantt-card::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:linear-gradient(180deg,#a8d5ff 0%,#1a73e8 100%)}
.gantt-card.overdue::before{background:linear-gradient(180deg,#ffcccc 0%,#d93025 100%)}
.gantt-card.completed::before{background:linear-gradient(180deg,#c8e6c9 0%,#137333 100%)}
.gantt-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px}
.gantt-info{flex:1}
.gantt-title{font-size:1.3rem;font-weight:500;margin-bottom:8px;display:flex;align-items:center;gap:8px}
.gantt-meta{display:flex;gap:16px;font-size:13px;color:var(--text-secondary);flex-wrap:wrap}
.countdown{text-align:right;padding:12px 20px;background:linear-gradient(135deg,#f0f7ff 0%,#e3f2fd 100%);border-radius:12px;min-width:220px}
.countdown.overdue{background:linear-gradient(135deg,#ffebee 0%,#ffcdd2 100%)}
.countdown.completed{background:linear-gradient(135deg,#e8f5e9 0%,#c8e6c9 100%)}
.countdown-label{font-size:11px;color:var(--text-secondary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.countdown-time{font-size:1.4rem;font-weight:600;color:var(--primary);font-family:'Courier New',monospace;letter-spacing:1px}
.countdown.overdue .countdown-time{color:var(--danger)}
.countdown.completed .countdown-time{color:var(--success);font-size:1.1rem}
.countdown-ds{font-size:.8rem;color:#7baaf7;font-weight:500}
.countdown.overdue .countdown-ds{color:#e57373}
.countdown.completed .countdown-ds{color:#66bb6a}

.gantt-track{position:relative;height:80px;background:var(--gray-light);border-radius:12px;overflow:hidden;border:1px solid var(--border-light);margin-top:16px}
.gantt-bar{position:absolute;top:10px;height:60px;border-radius:8px;background:linear-gradient(90deg,#bbdefb 0%,#64b5f6 40%,#1a73e8 100%);box-shadow:0 2px 8px rgba(26,115,232,.3);display:flex;align-items:center;justify-content:space-around;padding:0 8px;transition:all .3s}
.gantt-bar.overdue{background:linear-gradient(90deg,#ffcdd2 0%,#ef5350 40%,#d93025 100%);box-shadow:0 2px 8px rgba(217,48,37,.3)}
.gantt-bar.completed{background:linear-gradient(90deg,#c8e6c9 0%,#66bb6a 40%,#137333 100%);box-shadow:0 2px 8px rgba(19,115,51,.3)}
.gantt-bar:hover{transform:scaleY(1.08);filter:brightness(1.05)}
.gantt-stage{color:white;font-size:10px;font-weight:500;text-align:center;text-shadow:0 1px 2px rgba(0,0,0,.2);flex:1;padding:2px}
.gantt-stage-icon{font-size:16px;margin-bottom:1px}
.gantt-marker{position:absolute;top:0;width:2px;height:100%;z-index:2;cursor:pointer;transition:all .2s}
.gantt-marker:hover{width:4px;filter:brightness(1.2)}
.gantt-marker.create{background:#137333}.gantt-marker.status{background:#e37400}.gantt-marker.edit{background:#1a73e8}.gantt-marker.file{background:#9334e6}.gantt-marker.reject{background:#d93025}
.gantt-tooltip{display:none;position:absolute;bottom:calc(100% + 4px);left:50%;transform:translateX(-50%);background:#202124;color:white;padding:6px 10px;border-radius:6px;font-size:11px;white-space:nowrap;z-index:100;pointer-events:none}
.gantt-marker:hover .gantt-tooltip{display:block}
.gantt-now{position:absolute;top:0;width:2px;height:100%;background:#d93025;z-index:3;animation:pulse-now 2s infinite}
@keyframes pulse-now{0%,100%{opacity:1}50%{opacity:.4}}
.gantt-now::after{content:'▼ Сейчас';position:absolute;top:-18px;left:50%;transform:translateX(-50%);color:#d93025;font-size:9px;font-weight:600;white-space:nowrap}
.gantt-events{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap}
.gantt-chip{padding:5px 10px;background:var(--gray-light);border-radius:8px;font-size:11px;border:1px solid var(--border-light);display:flex;align-items:center;gap:4px;transition:all .2s;cursor:default}
.gantt-chip:hover{transform:translateY(-1px);box-shadow:var(--shadow)}
.gantt-chip.create{background:#e8f5e9;color:#137333;border-color:#a5d6a7}
.gantt-chip.status{background:#fff3e0;color:#e37400;border-color:#ffcc80}
.gantt-chip.edit{background:#e3f2fd;color:#1a73e8;border-color:#90caf9}
.gantt-chip.file{background:#f3e5f5;color:#9334e6;border-color:#ce93d8}
.gantt-legend{display:flex;gap:16px;margin-top:16px;padding-top:16px;border-top:1px solid var(--border-light);font-size:12px;flex-wrap:wrap;align-items:center}
.gantt-legend-item{display:flex;align-items:center;gap:6px}
.gantt-legend-dot{width:10px;height:10px;border-radius:50%}
.gantt-legend-bar{width:40px;height:12px;border-radius:6px}
.countdown-time{white-space:nowrap;font-size:clamp(0.9rem,2.5vw,1.4rem)}

@media(max-width:700px){body{padding:10px}.grid{grid-template-columns:1fr}header{flex-direction:column;align-items:stretch}.notification-dropdown{width:92%;right:4%}.dashboard-grid{grid-template-columns:1fr 1fr}.gantt-header{flex-direction:column;gap:12px}.countdown{min-width:auto;text-align:left}}
</style>
</head>
<body>
<div id="auth" class="auth">
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px"><h2 style="margin:0">🔐 Форма заявки</h2><button class="theme-toggle" onclick="toggleTheme()" title="Тема">🌓</button></div>
<div id="aerr" style="color:var(--danger);text-align:center;margin-bottom:12px;font-size:13px"></div>
<div class="tabs" id="atabs"><div class="tab active" data-tab="in">Вход</div><div class="tab" data-tab="rg">Регистрация</div><div class="tab" data-tab="rc">Восстановить</div></div>
<div id="fin"><input type="email" id="lem" placeholder="Email @stroisservis.ru"><input type="password" id="lps" placeholder="Пароль"><button class="btn btn-pri" id="loginBtn" style="width:100%">Войти</button></div>
<div id="frg" class="hidden"><input type="email" id="rem" placeholder="Email"><input type="text" id="rnm" placeholder="Имя Фамилия"><input type="password" id="rps" placeholder="Пароль"><select id="rrl"><option value="client">Заказчик</option><option value="procurement">Отдел снабжения</option><option value="director">Руководство</option></select><button class="btn btn-pri" id="regBtn" style="width:100%">Зарегистрироваться</button></div>
<div id="frc" class="hidden"><input type="email" id="rcem" placeholder="Email для восстановления"><button class="btn btn-wrn" id="recoverBtn" style="width:100%">Отправить код</button><div id="recoverCodeArea" class="hidden" style="margin-top:12px"><input type="text" id="rccode" maxlength="6" placeholder="Код из письма" style="text-align:center"><input type="password" id="rcnew" placeholder="Новый пароль"><button class="btn btn-suc" id="recoverResetBtn" style="width:100%">Сбросить пароль</button></div></div>
<div id="fvr" class="hidden"><p style="text-align:center">🔐 Код на <b id="vem"></b></p><input type="text" id="vcd" maxlength="6" placeholder="000000" style="text-align:center"><button class="btn btn-suc" id="vrfBtn" style="width:100%">Подтвердить</button><button class="btn btn-gra" id="rsndBtn" style="width:100%;margin-top:8px">Отправить повторно</button></div>
</div>

<div id="app" class="hidden">
<header><h1>📋 Форма заявки</h1><div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap"><div class="notification-bell" id="notificationBell" style="position:relative">🔔<span id="notificationCount" class="notification-badge hidden">0</span></div><button class="theme-toggle" onclick="toggleTheme()">🌓</button><span class="uinfo" id="uinf"></span><span id="unreadBadgeHeader" class="hidden unread-badge">0</span><button class="btn btn-dan btn-sm" id="logoutBtn">Выйти</button></div></header>
<div id="cnt"></div>
</div>

<div id="notificationDropdown" class="notification-dropdown"><div style="padding:12px 16px;border-bottom:1px solid var(--border-light);font-weight:600;display:flex;justify-content:space-between;align-items:center"><span>🔔 Уведомления</span><button class="btn btn-sm btn-gra" onclick="event.stopPropagation();clearNotifications()">Очистить</button></div><div id="notificationList"></div></div>

<!-- Modals -->
<div id="mdl" class="modal"><div class="modal-c"><h3 id="mtl">Заявка</h3><form id="frm"><input type="hidden" id="fid"><label>Проект</label><input id="fpr" readonly><label>Срок выполнения</label><input type="date" id="fdd" required><div style="display:grid;grid-template-columns:1fr 1fr;gap:12px"><div><label>Приоритет</label><select id="fpriority"><option value="low">🟢 Низкий</option><option value="medium" selected>🟡 Средний</option><option value="high">🔴 Высокий</option></select></div><div><label>💰 Бюджет (₽)</label><input type="number" id="fbudget" placeholder="0" min="0"></div></div><label>Описание</label><textarea id="fbd" rows="4" required></textarea><div class="file-upload-area" id="fileUploadArea"><p>📎 Перетащите файлы сюда или <b>нажмите для выбора</b></p><input type="file" id="fileInput" multiple style="display:none"><div id="selectedFiles" class="file-list"></div></div><div class="modal-buttons"><button type="submit" class="btn btn-pri">💾 Сохранить</button><button type="button" class="btn btn-gra" id="modalCancel">Отмена</button></div></form></div></div>
<div id="commentModal" class="modal"><div class="modal-c"><h3 id="commentTitle">Комментарий</h3><textarea id="commentText" rows="3" placeholder="Введите комментарий..."></textarea><div class="modal-buttons"><button class="btn btn-pri" id="submitCommentBtn">Подтвердить</button><button class="btn btn-gra" id="cancelCommentBtn">Отмена</button></div></div></div>
<div id="projectModal" class="modal"><div class="modal-c"><h3>Новый проект</h3><label>Название</label><input type="text" id="projectName" placeholder="Название"><label>Крайний срок</label><input type="date" id="projectDeadline" required><div class="modal-buttons"><button class="btn btn-pri" id="createProjectBtn">Создать</button><button class="btn btn-gra" id="cancelProjectBtn">Отмена</button></div></div></div>
<div id="editProjectDeadlineModal" class="modal"><div class="modal-c"><h3>Изменить срок проекта</h3><input type="text" id="editProjectName" readonly><label>Новый срок</label><input type="date" id="editProjectDeadline" required><div class="modal-buttons"><button class="btn btn-pri" id="saveProjectDeadlineBtn">Сохранить</button><button class="btn btn-gra" id="cancelEditDeadlineBtn">Отмена</button></div></div></div>
<div id="auditModal" class="modal"><div class="modal-c" style="max-width:700px"><h3 id="auditTitle">История</h3><div id="auditContent"></div><div class="modal-buttons"><button class="btn btn-gra" id="closeAuditBtn">Закрыть</button></div></div></div>
<div id="deleteUserModal" class="modal"><div class="modal-c"><h3>Удаление пользователя</h3><select id="deleteUserSelect" style="width:100%"></select><div class="modal-buttons"><button class="btn btn-dan" id="confirmDeleteUserBtn">🗑️ Удалить</button><button class="btn btn-gra" id="cancelDeleteUserBtn">Отмена</button></div></div></div>
<div id="notificationModal" class="modal"><div class="modal-c"><h3 id="notificationTitle">Уведомление</h3><p id="notificationMessage"></p><div class="modal-buttons"><button class="btn btn-pri" id="notificationOkBtn">OK</button></div></div></div>

<script>
const ROLE_NAMES={'client':'👤 Заказчик','procurement':'📦 Отдел снабжения','director':'👔 Руководство'};
function getRoleName(r){return ROLE_NAMES[r]||r}
function getUserDisplayName(name){
    if(!name)return'[Неизвестно]';if(String(name).includes('(Удалён'))return name;
    if(!D||!D.auth||!D.auth.users)return name;
    const au=D.auth.users.find(u=>(u.name===name||u.username===name)&&!u.deleted);if(au)return au.name;
    const du=D.auth.users.find(u=>u.name===name&&u.deleted);if(du)return`${name} (Удалён ${du.deletedAt||''})`;
    return name;
}

let D=null,CP=null,U=null,PE=null,IsArch=false,TOKEN=null;
let notifications=[],lastCheckTime=null,checkInterval=null,ganttInterval=null;
let selectedMassIds=[],selectedFilesForUpload=[];

// ===== GANTT TIMER ENGINE =====
function startGanttTimers(){
    if(ganttInterval)cancelAnimationFrame(ganttInterval);
    function tick(){
    const now=Date.now();
    document.querySelectorAll('[data-deadline]').forEach(el=>{
        const dl=parseInt(el.dataset.deadline);
        const diff=dl-now;
        const timeEl=el.querySelector('.countdown-time');
        const dsEl=el.querySelector('.countdown-ds');
        if(!timeEl||!dsEl)return;
        if(diff>0){
            const d=Math.floor(diff/86400000);
            const h=Math.floor((diff%86400000)/3600000);
            const m=Math.floor((diff%3600000)/60000);
            const s=Math.floor((diff%60000)/1000);
            const ds=Math.floor((diff%1000)/100);
            // ✅ Одна строка: 85д 12:34:56.7
            timeEl.textContent=`${String(d).padStart(2,'0')}д ${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}.${ds}`;
            dsEl.textContent='';
            el.classList.remove('overdue');
        }else{
            const ad=Math.abs(diff);
            const d=Math.floor(ad/86400000);
            const h=Math.floor((ad%86400000)/3600000);
            const m=Math.floor((ad%3600000)/60000);
            const s=Math.floor((ad%60000)/1000);
            const ds=Math.floor((ad%1000)/100);
            // ✅ Одна строка: -05д 03:22:11.4
            timeEl.textContent=`-${String(d).padStart(2,'0')}д ${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}.${ds}`;
            dsEl.textContent='';
            el.classList.add('overdue');
        }
    });
    ganttInterval=requestAnimationFrame(tick);
}
    tick();
}
function stopGanttTimers(){if(ganttInterval){cancelAnimationFrame(ganttInterval);ganttInterval=null}}

function playSound(type='notify'){try{const ctx=new(window.AudioContext||window.webkitAudioContext)();const osc=ctx.createOscillator();const gain=ctx.createGain();osc.connect(gain);gain.connect(ctx.destination);if(type==='notify'){osc.frequency.value=880;osc.type='sine'}else if(type==='overdue'){osc.frequency.value=440;osc.type='sawtooth'}else if(type==='success'){osc.frequency.value=1200;osc.type='sine'}gain.gain.setValueAtTime(0.3,ctx.currentTime);gain.gain.exponentialRampToValueAtTime(0.01,ctx.currentTime+0.3);osc.start();osc.stop(ctx.currentTime+0.3)}catch(e){}}
function toggleTheme(){document.body.classList.toggle('dark');localStorage.setItem('theme',document.body.classList.contains('dark')?'dark':'light')}
if(localStorage.getItem('theme')==='dark')document.body.classList.add('dark');
function requestNotificationPermission(){if('Notification'in window&&Notification.permission!=='granted'&&Notification.permission!=='denied')Notification.requestPermission()}
function showBrowserNotification(title,message,requestId=null){if(!('Notification'in window)||Notification.permission!=='granted')return;const n=new Notification(title,{body:message,silent:false});n.onclick=function(){window.focus();if(requestId)highlightRequestRow(requestId);n.close()};setTimeout(()=>n.close(),5000)}
function highlightRequestRow(id){const row=document.querySelector(`tr[data-id="${id}"]`);if(row){row.scrollIntoView({behavior:'smooth',block:'center'});row.classList.add('highlight');setTimeout(()=>row.classList.remove('highlight'),3000)}}
function addNotification(notif){notifications.unshift({...notif,id:Date.now()+Math.random(),timestamp:new Date(),read:false});if(notifications.length>50)notifications.pop();updateNotificationUI();showBrowserNotification(notif.title,notif.message,notif.requestId);playSound(notif.type==='deadline_warning'?'overdue':'notify')}
function updateNotificationUI(){const unread=notifications.filter(n=>!n.read).length;const cE=document.getElementById('notificationCount');const lE=document.getElementById('notificationList');if(unread>0){cE.textContent=unread>99?'99+':unread;cE.classList.remove('hidden')}else{cE.classList.add('hidden')}if(lE){if(notifications.length===0){lE.innerHTML='<div style="padding:20px;text-align:center;color:var(--text-secondary)">Нет уведомлений</div>'}else{lE.innerHTML=notifications.slice(0,20).map(n=>`<div class="notification-item ${n.read?'':'unread'}" onclick="markNotificationRead('${n.id}','${n.requestId||''}')"><div class="notification-title">${escapeHtml(n.title)}</div><div class="notification-message">${escapeHtml(n.message)}</div><div class="notification-time">${formatTime(n.timestamp)}</div></div>`).join('')}}}
function formatTime(date){const d=new Date(date),now=new Date(),diff=now-d;if(diff<60000)return'только что';if(diff<3600000)return`${Math.floor(diff/60000)} мин`;if(diff<86400000)return`${Math.floor(diff/3600000)} ч`;return d.toLocaleDateString()}
function markNotificationRead(id,requestId){const n=notifications.find(x=>x.id==id);if(n){n.read=true;updateNotificationUI()}if(requestId&&D){const req=D.requests.find(r=>r.id===requestId);if(req){document.getElementById('notificationDropdown').classList.remove('show');openProject(req.project);setTimeout(()=>highlightRequestRow(requestId),300)}}}
function toggleNotificationDropdown(e){if(e)e.stopPropagation();const d=document.getElementById('notificationDropdown');d.classList.toggle('show');if(d.classList.contains('show')){notifications.forEach(n=>n.read=true);updateNotificationUI()}}
function clearNotifications(){notifications=[];updateNotificationUI()}
async function checkForUpdates(){if(!U||!TOKEN)return;const checkTime=lastCheckTime||new Date(Date.now()-3600000).toISOString();try{const res=await api('/updates',{lastCheck:checkTime},'POST');if(res&&res.updates&&res.updates.length>0){for(const u of res.updates)addNotification({title:u.title,message:u.message,requestId:u.requestId,type:u.type})}lastCheckTime=new Date().toISOString()}catch(e){console.error('Updates error:',e)}}
function startNotificationChecker(){if(checkInterval)clearInterval(checkInterval);checkInterval=setInterval(checkForUpdates,30000);setTimeout(checkForUpdates,3000)}
function stopNotificationChecker(){if(checkInterval){clearInterval(checkInterval);checkInterval=null}}
function escapeHtml(s){if(s==null)return'';return String(s).replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]))}
function showNotification(msg,title="Уведомление"){document.getElementById('notificationTitle').innerText=title;document.getElementById('notificationMessage').innerHTML=msg;document.getElementById('notificationModal').style.display='flex'}
function formatMoney(n){if(!n)return'';return Number(n).toLocaleString('ru-RU')+'₽'}
const api=async(p,b={},m='POST')=>{const h={'Content-Type':'application/json'};if(TOKEN)h['X-Token']=TOKEN;try{const r=await fetch(p,{method:m,headers:h,body:Object.keys(b).length?JSON.stringify(b):null});if(r.status===401){logout();return{error:'SessionExpired'}}const t=await r.text();return t?JSON.parse(t):{error:'Empty'}}catch(e){return{error:'Net'}}};
function getWorkingDaysDiff(start,end){let c=new Date(start),e=new Date(end),d=0;while(c<=e){if(c.getDay()!==0&&c.getDay()!==6)d++;c.setDate(c.getDate()+1)}return d}
function isOverdue(r){if(!r.deadline||r.status==='Оплачено'||r.status==='Договорённость'||r.status==='Отклонено')return false;return new Date(r.deadline)<new Date(new Date().toDateString())}
function getStatusClass(status){const map={'Отправлено на рассмотрение':'st-new','Ожидает рассмотрения':'st-new','Выполняется':'st-wrk','Принято в работу':'st-wrk','Оплачено':'st-ok','Договорённость':'st-ok','Отклонено':'st-no','Приостановлено':'st-pause','Возвращено на доработку':'st-back'};return map[status]||'st-new'}
async function updateUnreadCount(){if(U&&U.role==='procurement'){const res=await api('/unread-count',{},'GET');if(res&&res.error==='SessionExpired')return;if(res&&typeof res.count==='number'){const b=document.getElementById('unreadBadgeHeader');const bc=document.getElementById('unreadBadgeCard');if(res.count>0){b.textContent=res.count;b.classList.remove('hidden');if(bc){bc.textContent=res.count;bc.classList.remove('hidden')}}else{b.classList.add('hidden');if(bc)bc.classList.add('hidden')}}}}
async function markRequestsAsRead(ids){for(let id of ids)await api('/mark-read',{requestId:id});await updateUnreadCount()}
function switchTab(tabId){document.getElementById('aerr').textContent='';document.getElementById('aerr').style.color='var(--danger)';['fin','frg','fvr','frc'].forEach(id=>document.getElementById(id).classList.add('hidden'));if(tabId==='in')document.getElementById('fin').classList.remove('hidden');else if(tabId==='rg')document.getElementById('frg').classList.remove('hidden');else if(tabId==='vr')document.getElementById('fvr').classList.remove('hidden');else if(tabId==='rc')document.getElementById('frc').classList.remove('hidden');document.querySelectorAll('#atabs .tab').forEach(t=>t.classList.toggle('active',t.getAttribute('data-tab')===tabId))}
async function reg(){const e=document.getElementById('rem').value.trim().toLowerCase();const n=document.getElementById('rnm').value.trim();const p=document.getElementById('rps').value.trim();const r=document.getElementById('rrl').value;const er=document.getElementById('aerr');if(!e.endsWith('@stroisservis.ru')){er.textContent='❌ Только @stroisservis.ru';return}if(!n){er.textContent='❌ Укажите имя';return}if(p.length<4){er.textContent='❌ Пароль минимум 4 символа';return}const res=await api('/reg',{username:e,password:p,role:r,name:n});if(res.ok){PE=e;document.getElementById('vem').textContent=e;switchTab('vr');er.textContent=''}else er.textContent=res.error}
async function vrf(){const c=document.getElementById('vcd').value.trim();const er=document.getElementById('aerr');if(!c)return;const res=await api('/vrf',{username:PE,code:c});if(res.ok){PE=null;switchTab('in');document.getElementById('lem').value=res.username}else er.textContent=res.error}
async function rsnd(){if(PE)await api('/rsnd',{username:PE})}
async function login(){const e=document.getElementById('lem').value.trim().toLowerCase();const p=document.getElementById('lps').value.trim();const er=document.getElementById('aerr');const res=await api('/log',{username:e,password:p});if(res.ok){U=res.user;TOKEN=res.token;localStorage.setItem('z1u',JSON.stringify(U));localStorage.setItem('z1t',TOKEN);showApp();await loadData();await updateUnreadCount();requestNotificationPermission();startNotificationChecker();setInterval(updateUnreadCount,30000)}else er.textContent=res.error}
async function recoverRequest(){const e=document.getElementById('rcem').value.trim().toLowerCase();const er=document.getElementById('aerr');if(!e){er.textContent='Введите email';return}const res=await api('/recover-request',{username:e});if(res.ok){document.getElementById('recoverCodeArea').classList.remove('hidden');er.textContent='✅ Код отправлен на почту';er.style.color='var(--success)'}else er.textContent=res.error}
async function recoverReset(){const e=document.getElementById('rcem').value.trim().toLowerCase();const c=document.getElementById('rccode').value.trim();const p=document.getElementById('rcnew').value.trim();const er=document.getElementById('aerr');if(!c||!p){er.textContent='Введите код и новый пароль';return}const res=await api('/recover-reset',{username:e,code:c,newPassword:p});if(res.ok){er.textContent='✅ Пароль изменён! Войдите';er.style.color='var(--success)';document.getElementById('recoverCodeArea').classList.add('hidden');switchTab('in')}else{er.textContent=res.error;er.style.color='var(--danger)'}}
function logout(){stopNotificationChecker();stopGanttTimers();localStorage.removeItem('z1u');localStorage.removeItem('z1t');U=null;TOKEN=null;CP=null;IsArch=false;notifications=[];document.getElementById('auth').classList.remove('hidden');document.getElementById('app').classList.add('hidden')}
function showApp(){document.getElementById('auth').classList.add('hidden');document.getElementById('app').classList.remove('hidden');document.getElementById('uinf').textContent=U.name+' ('+getRoleName(U.role)+')'}
async function loadData(){const res=await api('/dat',{},'GET');if(res&&(res.error==='Auth'||res.error==='SessionExpired')){logout();return}if(!res||res.error){showNotification("Ошибка загрузки","Ошибка");return}D=res;if(!D.projects)D.projects={active:[],archive:[]};if(!Array.isArray(D.projects.active))D.projects.active=[];if(!Array.isArray(D.projects.archive))D.projects.archive=[];if(!Array.isArray(D.requests))D.requests=[];if(!D.system)D.system={statuses:[]};render()}
function getPriorityIcon(p){switch(p){case'low':return'<span class="priority-low">🟢 Низкий</span>';case'medium':return'<span class="priority-medium">🟡 Средний</span>';case'high':return'<span class="priority-high">🔴 Высокий</span>';default:return'<span class="priority-medium">🟡 Средний</span>'}}
function getDisplayStatus(r){if(U.role!=='client')return r.status;switch(r.status){case'Отправлено на рассмотрение':return'Ожидает рассмотрения';case'Выполняется':return'Принято в работу';case'Отклонено':return'Отклонено';default:return r.status}}

function showAudit(requestId){const req=D.requests.find(r=>r.id===requestId);if(!req||!req.audit)return;let html='<div class="timeline">';req.audit.slice().reverse().forEach(entry=>{let tc='status-change';if(entry.action==='Создание заявки')tc='create';else if(entry.action==='Редактирование')tc='edit';else if(entry.action==='Загрузка файла')tc='edit';else if(entry.newStatus==='Отклонено')tc='reject';const ch=entry.comment&&entry.comment.trim()&&entry.comment.trim()!==' '?`<div class="timeline-comment">💬 ${escapeHtml(entry.comment)}</div>`:'';html+=`<div class="timeline-item ${tc}"><div class="timeline-time">${escapeHtml(entry.timestamp)}</div><div class="timeline-actor">👤 ${escapeHtml(getUserDisplayName(entry.actor))}</div><div class="timeline-action">${escapeHtml(entry.action)} → <b>${escapeHtml(entry.newStatus)}</b></div>${ch}</div>`});html+='</div>';document.getElementById('auditContent').innerHTML=html;document.getElementById('auditTitle').innerText=`📜 История ${requestId}`;document.getElementById('auditModal').style.display='flex'}
async function deleteFromArchive(name){if(!confirm(`Удалить "${name}" из архива?`))return;const res=await api('/delete-archive',{name});if(res.error)showNotification(res.error,"Ошибка");else{showNotification("Удалено","Успешно");await loadData()}}
async function showDeleteUserModal(){const s=document.getElementById('deleteUserSelect');s.innerHTML='<option value="">-- Выберите --</option>';D.auth.users.forEach(u=>{if(!u.deleted)s.innerHTML+=`<option value="${escapeHtml(u.username)}">${escapeHtml(u.name)} (${getRoleName(u.role)})</option>`});document.getElementById('deleteUserModal').style.display='flex'}
async function confirmDeleteUser(){const u=document.getElementById('deleteUserSelect').value;if(!u)return;if(!confirm('Удалить пользователя?'))return;const res=await api('/delete-user',{username:u});if(res.error)showNotification(res.error);else{showNotification(res.message,"Успешно");document.getElementById('deleteUserModal').style.display='none';await loadData()}}

// ===== ИСПРАВЛЕННЫЙ GANTT RENDERER =====
function renderGanttDashboard(){
    stopGanttTimers();
    const container=document.getElementById('cnt');
    const projects=D.projects.active||[];
    const allReqs=D.requests||[];
    
    let statsHtml=`<div class="dashboard-grid">
        <div class="dashboard-card"><div class="label">Активных проектов</div><div class="big-number">${projects.length}</div></div>
        <div class="dashboard-card warning"><div class="label">Заявок в работе</div><div class="big-number">${allReqs.filter(r=>r.status==='Выполняется').length}</div></div>
        <div class="dashboard-card danger"><div class="label">Просрочено</div><div class="big-number">${allReqs.filter(r=>isOverdue(r)).length}</div></div>
        <div class="dashboard-card success"><div class="label">Завершено</div><div class="big-number">${allReqs.filter(r=>r.status==='Оплачено'||r.status==='Договорённость').length}</div></div>
    </div>`;
    
    let cardsHtml='';
    projects.forEach(proj=>{
        const reqs=allReqs.filter(r=>r.project===proj.name);
        const hasOverdue=reqs.some(r=>isOverdue(r));
        const allDone=reqs.length>0&&reqs.every(r=>['Оплачено','Договорённость','Отклонено'].includes(r.status));
        const cardClass=hasOverdue?'overdue':(allDone?'completed':'');
        
        let startDate=new Date();
        if(reqs.length>0){
            let earliest=null;
            reqs.forEach(r=>{
                if(r.createdDate){
                    const cd=new Date(r.createdDate);
                    if(!earliest||cd<earliest)earliest=cd;
                }
            });
            if(earliest)startDate=earliest;
        }
        
        let endDate=proj.deadline?new Date(proj.deadline):new Date(startDate.getTime()+86400000*30);
        if(startDate>endDate)endDate=new Date(startDate.getTime()+86400000*30);
        
        const now=new Date();
        const totalMs=endDate.getTime()-startDate.getTime();
        const nowPct=totalMs>0?Math.max(0,Math.min(100,((now.getTime()-startDate.getTime())/totalMs)*100)):0;
        const barWidth=Math.max(10,Math.min(100,100));
        
        let markersHtml='';
        reqs.forEach(r=>{
            if(r.audit){
                r.audit.forEach(a=>{
                    const at=new Date(a.timestamp);
                    const pct=totalMs>0?Math.max(0,Math.min(100,((at.getTime()-startDate.getTime())/totalMs)*100)):0;
                    let cls='status';
                    if(a.action==='Создание заявки')cls='create';
                    else if(a.action==='Редактирование')cls='edit';
                    else if(a.action==='Загрузка файла')cls='file';
                    else if(a.newStatus==='Отклонено')cls='reject';
                    markersHtml+=`<div class="gantt-marker ${cls}" style="left:${pct}%"><div class="gantt-tooltip">${escapeHtml(a.timestamp.substring(5,16))} — ${escapeHtml(getUserDisplayName(a.actor))}: ${escapeHtml(a.action)}</div></div>`;
                });
            }
        });
        
        let allEvents=[];
        reqs.forEach(r=>{if(r.audit)r.audit.forEach(a=>allEvents.push({...a,reqId:r.id}))});
        allEvents.sort((a,b)=>new Date(b.timestamp)-new Date(a.timestamp));
        let chipsHtml=allEvents.slice(0,5).map(e=>{
            let cls='status';if(e.action==='Создание заявки')cls='create';else if(e.action==='Редактирование')cls='edit';else if(e.action==='Загрузка файла')cls='file';
            const icon=cls==='create'?'📝':cls==='edit'?'✏️':cls==='file'?'📎':'⚙️';
            return`<div class="gantt-chip ${cls}"><span>${icon}</span><span>${escapeHtml(e.timestamp.substring(5,16))} ${escapeHtml(getUserDisplayName(e.actor))} — ${escapeHtml(e.reqId)}</span></div>`;
        }).join('');
        
        const dlMs=endDate.getTime();
        const countdownClass=hasOverdue?'overdue':(allDone?'completed':'');
        const countdownLabel=allDone?'Статус':'До дедлайна';
        const countdownContent=allDone?`<div class="countdown-time" style="color:var(--success);font-size:1.1rem">✓ ГОТОВО</div><div class="countdown-ds" style="color:var(--success)">Проект завершён</div>`:`<div class="countdown-time"></div><div class="countdown-ds"></div>`;
        
        let stagesHtml='<div class="gantt-stage"><div class="gantt-stage-icon">📝</div><div>Создание</div></div>';
        if(reqs.some(r=>r.files&&r.files.length>0))stagesHtml+='<div class="gantt-stage"><div class="gantt-stage-icon">📎</div><div>Файлы</div></div>';
        stagesHtml+='<div class="gantt-stage"><div class="gantt-stage-icon">⚙️</div><div>Обработка</div></div>';
        if(reqs.some(r=>['Оплачено','Договорённость'].includes(r.status)))stagesHtml+='<div class="gantt-stage"><div class="gantt-stage-icon">💰</div><div>Оплата</div></div>';
        stagesHtml+='<div class="gantt-stage"><div class="gantt-stage-icon">✅</div><div>Финал</div></div>';
        
        const badgeHtml=hasOverdue?'<span class="overdue-badge">Просрочен</span>':(allDone?'<span style="font-size:12px;padding:4px 10px;background:rgba(19,115,51,.12);color:#137333;border-radius:12px;font-weight:500">✓ Завершён</span>':'<span style="font-size:12px;padding:4px 10px;background:rgba(26,115,232,.12);color:#1a73e8;border-radius:12px;font-weight:500">Активен</span>');
        
        const startDateStr=startDate.toLocaleDateString('ru-RU',{day:'numeric',month:'long',year:'numeric'});
        const endDateStr=endDate.toLocaleDateString('ru-RU',{day:'numeric',month:'long',year:'numeric'});
        
        // 🔥 ИСПРАВЛЕНО: Добавлен data-project и cursor:pointer
        cardsHtml+=`
        <div class="gantt-card ${cardClass}" data-project="${escapeHtml(proj.name)}" style="cursor:pointer">
            <div class="gantt-header">
                <div class="gantt-info">
                    <div class="gantt-title">📁 ${escapeHtml(proj.name)} ${badgeHtml}</div>
                    <div class="gantt-meta">
                        <span>📅 Старт: ${startDateStr}</span>
                        <span>⏰ Дедлайн: ${endDateStr}</span>
                        <span>📋 ${reqs.length} заяв.</span>
                    </div>
                </div>
                <div class="countdown ${countdownClass}" data-deadline="${dlMs}">
                    <div class="countdown-label">${countdownLabel}</div>
                    ${countdownContent}
                </div>
            </div>
            <div class="gantt-track">
                <div class="gantt-bar ${cardClass}" style="left:0;width:${barWidth}%">${stagesHtml}</div>
                ${markersHtml}
                ${!allDone?`<div class="gantt-now" style="left:${nowPct}%"></div>`:''}
            </div>
            <div class="gantt-events">${chipsHtml||'<span style="color:var(--text-secondary);font-size:12px">Нет событий</span>'}</div>
        </div>`;
    });
    
    let legendHtml=`<div class="gantt-legend">
        <div class="gantt-legend-item"><div class="gantt-legend-bar" style="background:linear-gradient(90deg,#bbdefb,#64b5f6,#1a73e8)"></div><span>Активный</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-bar" style="background:linear-gradient(90deg,#c8e6c9,#66bb6a,#137333)"></div><span>Завершён</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-bar" style="background:linear-gradient(90deg,#ffcdd2,#ef5350,#d93025)"></div><span>Просрочен</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#d93025"></div><span>Сейчас</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#137333"></div><span>Создание</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#e37400"></div><span>Статус</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#1a73e8"></div><span>Правка</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#9334e6"></div><span>Файл</span></div>
    </div>`;
    
    container.innerHTML=`<div class="back" id="backBtn">← Назад к проектам</div><h2>📊 Дашборд проектов</h2>${statsHtml}${cardsHtml}${legendHtml}`;
    document.getElementById('backBtn').addEventListener('click',()=>{CP=null;render()});
    
    // 🔥 ДОБАВЛЕНО: Обработчик клика на карточки проектов
    document.querySelectorAll('.gantt-card').forEach(card=>{
        card.addEventListener('click',(e)=>{
            // Не срабатывать, если клик был на внутренних интерактивных элементах
            if(e.target.closest('.gantt-marker')||e.target.closest('.gantt-chip'))return;
            const projectName=card.dataset.project;
            if(projectName)openProject(projectName);
        });
    });
    
    startGanttTimers();
}

async function render(){
    const container=document.getElementById('cnt');
    if(!container||!U)return;
    try{
        container.innerHTML='';selectedMassIds=[];stopGanttTimers();
        if(!CP){
            let html='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px"><h3 style="margin:0">📂 Проекты</h3><div style="display:flex;gap:8px;flex-wrap:wrap">';
            // 🔥 ИСПРАВЛЕНО: Кнопка дашборда
            if(U.role==='director'){html+=`<button class="btn btn-pri btn-sm" id="dashboardBtn">📊 Дашборд проектов</button><button class="btn btn-suc btn-sm" id="exportBtn">📥 Excel</button><button class="btn btn-gra btn-sm" id="deleteUserBtn">🗑️ Пользователи</button>`}
            html+='</div></div>';
            html+=`<div class="search-bar"><input type="text" id="globalSearch" placeholder="🔍 Поиск по заявкам..."><select id="filterStatus"><option value="">Все статусы</option>`;
            (D.system.statuses||[]).forEach(s=>{if(s&&s.name)html+=`<option value="${escapeHtml(s.name)}">${escapeHtml(s.name)}</option>`});
            html+=`</select><button class="btn btn-pri btn-sm" id="searchBtn">Найти</button></div><div id="searchResults"></div>`;
            html+='<div class="grid">';
            if(U.role==='procurement')html+=`<div class="card" id="allRequestsCard" style="border-left:4px solid var(--warning)"><div style="display:flex;justify-content:space-between;align-items:center"><h3 style="margin:0">📋 Все новые заявки</h3><span id="unreadBadgeCard" class="hidden unread-badge">0</span></div></div>`;
            if(U.role==='director'){const pc=(D.requests||[]).filter(r=>r.status==='Выполняется').length;html+=`<div class="card" id="pendingRequestsCard" style="border-left:4px solid var(--warning)"><h3>⏳ Ожидают решения (${pc})</h3></div>`}
            if(U.role!=='client')html+='<div class="card card-new" id="newProjectBtn"><h3>+ Новый проект</h3></div>';
            (D.projects.active||[]).forEach(p=>{const name=p.name||p;const pr=(D.requests||[]).filter(r=>r.project===name);const oc=pr.filter(isOverdue).length;html+=`<div class="card"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px"><div><h3 class="project-name" data-project="${escapeHtml(name)}" style="cursor:pointer;margin:0 0 4px 0">📁 ${escapeHtml(name)}</h3><small style="color:var(--text-secondary)">Заявок: ${pr.length}${oc?` <span class="overdue-badge">⚠ ${oc} просрочено</span>`:''}</small></div><div style="display:flex;gap:6px">`;if(U.role==='director')html+=`<button class="btn btn-gra btn-sm edit-deadline-btn" data-project="${escapeHtml(name)}">✏️</button><button class="btn btn-gra btn-sm arch-btn" data-archive="${escapeHtml(name)}">📦</button>`;html+=`</div></div></div>`});
            if((D.projects.active||[]).length===0&&U.role!=='procurement')html+='<p>Нет активных проектов</p>';
            html+='</div><h3 style="margin:20px 0 12px">🗄️ Архив</h3><div class="grid">';
            (D.projects.archive||[]).forEach(p=>{let name='[Неизвестный проект]';if(p&&p.name&&typeof p.name==='string')name=p.name;else if(typeof p==='string')name=p;else if(p&&p.name&&typeof p.name==='object')name=p.name.name||'[Исправлено]';html+=`<div class="card arch-c"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px"><h3 class="project-name" data-project="${escapeHtml(name)}" style="cursor:pointer;margin:0">📁 ${escapeHtml(name)}</h3>`;if(U.role==='director')html+=`<button class="btn btn-dan btn-sm delete-archive-btn" data-project="${escapeHtml(name)}">🗑️</button>`;html+=`</div></div>`});
            if((D.projects.archive||[]).length===0)html+='<p>Архив пуст</p>';
            container.innerHTML=html+'</div>';
            document.querySelectorAll('.project-name').forEach(el=>el.addEventListener('click',()=>openProject(el.getAttribute('data-project'))));
            document.querySelectorAll('.arch-btn').forEach(b=>b.addEventListener('click',e=>{e.stopPropagation();archiveProject(b.getAttribute('data-archive'))}));
            document.querySelectorAll('.edit-deadline-btn').forEach(b=>b.addEventListener('click',e=>{e.stopPropagation();showEditDeadlineModal(b.getAttribute('data-project'))}));
            document.querySelectorAll('.delete-archive-btn').forEach(b=>b.addEventListener('click',e=>{e.stopPropagation();deleteFromArchive(b.getAttribute('data-project'))}));
            document.getElementById('newProjectBtn')?.addEventListener('click',showProjectModal);
            document.getElementById('allRequestsCard')?.addEventListener('click',()=>openProject('ALL_REQUESTS'));
            document.getElementById('pendingRequestsCard')?.addEventListener('click',()=>openProject('PENDING_REQUESTS'));
            document.getElementById('dashboardBtn')?.addEventListener('click',renderGanttDashboard);
            document.getElementById('exportBtn')?.addEventListener('click',exportExcel);
            document.getElementById('deleteUserBtn')?.addEventListener('click',showDeleteUserModal);
            document.getElementById('searchBtn')?.addEventListener('click',doSearch);
            document.getElementById('globalSearch')?.addEventListener('keypress',e=>{if(e.key==='Enter')doSearch()});
            document.getElementById('filterStatus')?.addEventListener('change',doSearch);
            await updateUnreadCount();return;
        }
        if(CP==='SEARCH'){const q=(window._searchQuery||'').toLowerCase();const fs=window._searchFilter||'';let results=D.requests.filter(r=>{const mq=!q||(r.id||'').toLowerCase().includes(q)||(r.body||'').toLowerCase().includes(q)||(r.author||'').toLowerCase().includes(q)||(r.project||'').toLowerCase().includes(q);const ms=!fs||r.status===fs;return mq&&ms});let html=`<div class="back" id="backBtn">← Назад</div><h2>🔍 Результаты поиска (${results.length})</h2>`;html+=renderRequestsTable(results,true);container.innerHTML=html;document.getElementById('backBtn').addEventListener('click',()=>{CP=null;render()});attachTableListeners();return}
        if(CP==='ALL_REQUESTS'){const an=D.requests.filter(r=>r.status==='Отправлено на рассмотрение');let html=`<div class="back" id="backBtn">← Назад</div><h2>📋 Все новые заявки (${an.length})</h2>`;html+=renderRequestsTable(an,true);container.innerHTML=html;document.getElementById('backBtn').addEventListener('click',()=>{CP=null;render()});attachTableListeners();await markRequestsAsRead(an.map(r=>r.id));return}
        if(CP==='PENDING_REQUESTS'){const pn=D.requests.filter(r=>r.status==='Выполняется');let html=`<div class="back" id="backBtn">← Назад</div><h2>⏳ Ожидают решения (${pn.length})</h2>`;html+=renderRequestsTable(pn,true);container.innerHTML=html;document.getElementById('backBtn').addEventListener('click',()=>{CP=null;render()});attachTableListeners();return}
        const reqs=D.requests.filter(r=>r&&r.project===CP);const ab=IsArch?' 🔒 Архив':'';const ca=(U.role==='client'||U.role==='procurement')&&!IsArch;let html=`<div class="back" id="backBtn">← Назад</div><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:16px"><h2 style="margin:0">${escapeHtml(CP)}${ab}</h2>`;if(ca)html+=`<button class="btn btn-suc" id="addRequestBtn">+ Новая заявка</button>`;html+='</div>';
        html+=`<div class="search-bar"><input type="text" id="projectSearch" placeholder="🔍 Фильтр..."><select id="projectFilter"><option value="">Все статусы</option>`;(D.system.statuses||[]).forEach(s=>{if(s&&s.name)html+=`<option value="${escapeHtml(s.name)}">${escapeHtml(s.name)}</option>`});html+=`</select></div>`;
        html+=renderRequestsTable(reqs,false);container.innerHTML=html;document.getElementById('backBtn').addEventListener('click',()=>{CP=null;render()});document.getElementById('addRequestBtn')?.addEventListener('click',()=>openModal());document.getElementById('projectSearch')?.addEventListener('input',filterProjectTable);document.getElementById('projectFilter')?.addEventListener('change',filterProjectTable);attachTableListeners();
    }catch(e){showNotification("Ошибка: "+e.message,"Ошибка")}
}
function filterProjectTable(){const q=(document.getElementById('projectSearch')?.value||'').toLowerCase();const fs=document.getElementById('projectFilter')?.value||'';document.querySelectorAll('#reqTable tbody tr').forEach(row=>{const t=row.textContent.toLowerCase();const s=row.getAttribute('data-status')||'';row.style.display=(!q||t.includes(q))&&(!fs||s===fs)?'':'none'})}
function renderRequestsTable(reqs,showProject){
    if(reqs.length===0)return'<p style="padding:20px;text-align:center;color:var(--text-secondary)">Нет заявок</p>';
    const cm=U.role==='procurement'||U.role==='director';
    let html='<div class="table-wrapper"><table id="reqTable"><thead><tr>';if(cm)html+='<th style="width:40px"><input type="checkbox" class="mass-checkbox" id="selectAll"></th>';html+='<th>ID</th>';if(showProject)html+='<th>Проект</th>';html+='<th>Срок</th><th>Приоритет</th><th>Бюджет</th><th>Статус</th><th>Описание</th><th>Автор</th><th>Файлы</th><th>Действия</th></tr></thead><tbody>';
    reqs.forEach(r=>{const ds=getDisplayStatus(r);const sc=getStatusClass(ds);const ov=isOverdue(r);const ob=ov?'<span class="overdue-badge">⚠ Просрочено</span>':'';const wh=(U.role==='client'&&r.deadline&&r.status==='Отправлено на рассмотрение'&&getWorkingDaysDiff(new Date(),new Date(r.deadline))<5)?'<span class="warning">⚠ <5 дней</span>':'';const ad=getUserDisplayName(r.author);const vf=(r.files||[]).filter(f=>f&&typeof f==='object'&&f.path&&f.name&&typeof f.path==='string'&&typeof f.name==='string');const fh=vf.length>0?vf.map(f=>`<a class="file-item" href="/download/${encodeURI(f.path)}" target="_blank" onclick="event.stopPropagation()">📄 ${escapeHtml(f.name)}</a>`).join(''):'<span style="color:var(--text-secondary);font-size:12px">—</span>';
    let btns='';const ce=(r.author===U.name&&r.status==='Отправлено на рассмотрение')||(U.role==='procurement')||(U.role==='director');if(ce&&!IsArch)btns+=`<button class="btn btn-gra btn-sm edit-btn" data-id="${r.id}">✏️</button>`;
    if(!IsArch){if(U.role==='procurement'&&r.status==='Отправлено на рассмотрение'){btns+=`<button class="btn btn-wrn btn-sm action-btn" data-id="${r.id}" data-status="Выполняется">✓ Принять</button><button class="btn btn-dan btn-sm action-btn" data-id="${r.id}" data-status="Отклонено">✕ Отклонить</button>`}if(U.role==='director'&&r.status==='Выполняется'){btns+=`<button class="btn btn-suc btn-sm action-btn" data-id="${r.id}" data-status="Оплачено">💰 Оплачено</button><button class="btn btn-gra btn-sm action-btn" data-id="${r.id}" data-status="Договорённость">🤝 Договор.</button><button class="btn btn-dan btn-sm action-btn" data-id="${r.id}" data-status="Отклонено">✕</button>`}}
    btns+=`<button class="btn btn-gra btn-sm audit-btn" data-id="${r.id}" title="История">📋</button>`;
    html+=`<tr data-id="${r.id}" data-status="${r.status}" class="${ov?'overdue':''}">`;if(cm)html+=`<td><input type="checkbox" class="mass-checkbox row-checkbox" data-id="${r.id}"></td>`;html+=`<td><b>${escapeHtml(r.id)}</b>${ob}${wh}</td>`;if(showProject)html+=`<td><span class="project-link" data-project="${escapeHtml(r.project)}">${escapeHtml(r.project)}</span></td>`;html+=`<td>${escapeHtml(r.deadline||'—')}</td><td>${getPriorityIcon(r.priority)}</td><td>${r.budget?formatMoney(r.budget):'<span style="color:var(--text-secondary)">—</span>'}</td><td><span class="st ${sc}">${escapeHtml(ds)}</span></td><td style="max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escapeHtml(r.body||'')}">${escapeHtml(r.body||'')}</td><td>${escapeHtml(ad)}</td><td>${fh}</td><td style="white-space:nowrap">${btns}</td></tr>`});
    html+='</tbody></table></div>';if(cm)html+=`<div id="massActionsBar" class="mass-actions-bar hidden"><span>Выбрано: <b id="massCount">0</b></span><div><button class="btn btn-sm" id="massAccept">✓ Принять все</button><button class="btn btn-sm" id="massReject">✕ Отклонить все</button></div></div>`;return html}
function attachTableListeners(){document.querySelectorAll('.action-btn').forEach(b=>b.addEventListener('click',()=>showCommentModal(b.getAttribute('data-id'),b.getAttribute('data-status'))));document.querySelectorAll('.audit-btn').forEach(b=>b.addEventListener('click',()=>showAudit(b.getAttribute('data-id'))));document.querySelectorAll('.edit-btn').forEach(b=>b.addEventListener('click',()=>openModal(b.getAttribute('data-id'))));document.querySelectorAll('.project-link').forEach(el=>el.addEventListener('click',()=>openProject(el.getAttribute('data-project'))));const sa=document.getElementById('selectAll');if(sa){sa.addEventListener('change',()=>{document.querySelectorAll('.row-checkbox').forEach(c=>{if(c.closest('tr').style.display!=='none')c.checked=sa.checked});updateMassCount()})}document.querySelectorAll('.row-checkbox').forEach(c=>c.addEventListener('change',updateMassCount));document.getElementById('massAccept')?.addEventListener('click',()=>doMassAction('Выполняется'));document.getElementById('massReject')?.addEventListener('click',()=>doMassAction('Отклонено'))}
function updateMassCount(){selectedMassIds=[];document.querySelectorAll('.row-checkbox:checked').forEach(c=>selectedMassIds.push(c.getAttribute('data-id')));const bar=document.getElementById('massActionsBar');const cnt=document.getElementById('massCount');if(bar&&cnt){if(selectedMassIds.length>0){bar.classList.remove('hidden');cnt.textContent=selectedMassIds.length}else bar.classList.add('hidden')}}
async function doMassAction(status){if(selectedMassIds.length===0)return;if(!confirm(`Применить "${status}" к ${selectedMassIds.length} заявкам?`))return;const res=await api('/bulk-process',{ids:selectedMassIds,status,comment:'Массовое действие'});if(res.error)showNotification(res.error);else{showNotification(`Обработано: ${selectedMassIds.length}`,"Успешно");playSound('success');await loadData()}}
function doSearch(){window._searchQuery=document.getElementById('globalSearch')?.value||'';window._searchFilter=document.getElementById('filterStatus')?.value||'';CP='SEARCH';render()}
let currentActionId=null,currentActionStatus=null;
function showCommentModal(id,status){currentActionId=id;currentActionStatus=status;document.getElementById('commentText').value='';document.getElementById('commentTitle').innerText={'Выполняется':'Принять в работу','Оплачено':'Подтверждение оплаты','Договорённость':'Договорённость','Отклонено':'Причина отклонения'}[status]||'Комментарий';document.getElementById('commentModal').style.display='flex'}
async function submitComment(){const c=document.getElementById('commentText').value.trim();let res;if(currentActionStatus==='Отклонено')res=await api('/rjt',{id:currentActionId,reason:c});else res=await api('/sts',{id:currentActionId,status:currentActionStatus,comment:c});if(res.error)showNotification(res.error);else{closeCommentModal();playSound('success');await loadData();await updateUnreadCount()}}
function closeCommentModal(){document.getElementById('commentModal').style.display='none'}
function showProjectModal(){document.getElementById('projectName').value='';document.getElementById('projectDeadline').value='';document.getElementById('projectModal').style.display='flex'}
async function createProjectHandler(){const n=document.getElementById('projectName').value.trim();const d=document.getElementById('projectDeadline').value;if(!n||!d)return showNotification("Заполните поля");const res=await api('/prj',{name:n,deadline:d});if(res.error)showNotification(res.error);else{closeProjectModal();await loadData()}}
function closeProjectModal(){document.getElementById('projectModal').style.display='none'}
function showEditDeadlineModal(name){const p=D.projects.active.find(x=>x.name===name);if(!p)return;document.getElementById('editProjectName').value=name;document.getElementById('editProjectDeadline').value=p.deadline||'';document.getElementById('editProjectDeadlineModal').style.display='flex'}
async function saveProjectDeadline(){const n=document.getElementById('editProjectName').value;const d=document.getElementById('editProjectDeadline').value;if(!d)return showNotification("Укажите срок");const res=await api('/update-project-deadline',{name:n,deadline:d});if(res.error)showNotification(res.error);else{closeEditDeadlineModal();await loadData()}}
function closeEditDeadlineModal(){document.getElementById('editProjectDeadlineModal').style.display='none'}
async function openProject(p){CP=p;IsArch=(p!=='ALL_REQUESTS'&&p!=='PENDING_REQUESTS'&&D.projects.archive&&D.projects.archive.some(x=>{const name=(x&&typeof x==='object'&&x.name)?x.name:(typeof x==='string'?x:'');return name===p}));await render()}
async function archiveProject(n){if(U.role!=='director')return showNotification("Нет прав");if(confirm('В архив "'+n+'"?')){const res=await api('/arc',{name:n});if(res.error)showNotification(res.error);else await loadData()}}
function openModal(id=null){document.getElementById('mdl').style.display='flex';const fdd=document.getElementById('fdd');if(fdd)fdd.min=new Date().toISOString().split('T')[0];selectedFilesForUpload=[];document.getElementById('selectedFiles').innerHTML='';setupFileUpload();if(id){const r=D.requests.find(x=>x.id===id);if(!r)return;document.getElementById('mtl').innerText='Редактирование '+id;document.getElementById('fid').value=r.id;document.getElementById('fpr').value=r.project;fdd.value=r.deadline||'';document.getElementById('fpriority').value=r.priority||'medium';document.getElementById('fbudget').value=r.budget||'';document.getElementById('fbd').value=r.body||''}else{document.getElementById('mtl').innerText='Новая заявка';document.getElementById('fid').value='';document.getElementById('fpr').value=CP;fdd.value='';document.getElementById('fpriority').value='medium';document.getElementById('fbudget').value='';document.getElementById('fbd').value=''}const proj=D.projects.active.find(p=>p.name===CP);if(proj&&proj.deadline)fdd.max=proj.deadline;else fdd.removeAttribute('max')}
function closeModal(){document.getElementById('mdl').style.display='none'}
function setupFileUpload(){const area=document.getElementById('fileUploadArea');const input=document.getElementById('fileInput');if(!area||!input)return;area.onclick=()=>input.click();area.ondragover=(e)=>{e.preventDefault();area.classList.add('dragover')};area.ondragleave=()=>area.classList.remove('dragover');area.ondrop=(e)=>{e.preventDefault();area.classList.remove('dragover');handleFiles(e.dataTransfer.files)};input.onchange=()=>handleFiles(input.files)}
function handleFiles(files){selectedFilesForUpload=Array.from(files);const el=document.getElementById('selectedFiles');if(!el)return;el.innerHTML=selectedFilesForUpload.map((f,i)=>`<div class="file-item">📄 ${escapeHtml(f.name)} (${(f.size/1024).toFixed(1)}KB) <span onclick="event.stopPropagation();removeFile(${i})" style="cursor:pointer;color:var(--danger)">✕</span></div>`).join('')}
function removeFile(i){selectedFilesForUpload.splice(i,1);handleFiles(selectedFilesForUpload)}
async function uploadFilesForRequest(requestId){for(const file of selectedFilesForUpload){const reader=new FileReader();const base64=await new Promise((resolve)=>{reader.onload=()=>resolve(reader.result.split(',')[1]);reader.readAsDataURL(file)});await api('/upload-file',{requestId,fileName:file.name,fileData:base64},'POST')}selectedFilesForUpload=[]}
async function exportExcel(){showNotification("⏳ Генерация отчёта...","Подождите");const res=await api('/export-excel',{},'POST');if(res.ok)showNotification("✅ Отчёт сохранён","Успешно");else showNotification(res.error||"Ошибка","Ошибка")}

window.onload=()=>{
    document.querySelectorAll('#atabs .tab').forEach(t=>t.addEventListener('click',()=>switchTab(t.getAttribute('data-tab'))));
    document.getElementById('loginBtn')?.addEventListener('click',login);
    document.getElementById('regBtn')?.addEventListener('click',reg);
    document.getElementById('vrfBtn')?.addEventListener('click',vrf);
    document.getElementById('rsndBtn')?.addEventListener('click',rsnd);
    document.getElementById('recoverBtn')?.addEventListener('click',recoverRequest);
    document.getElementById('recoverResetBtn')?.addEventListener('click',recoverReset);
    document.getElementById('logoutBtn')?.addEventListener('click',logout);
    document.getElementById('modalCancel')?.addEventListener('click',closeModal);
    document.getElementById('cancelCommentBtn')?.addEventListener('click',closeCommentModal);
    document.getElementById('submitCommentBtn')?.addEventListener('click',submitComment);
    document.getElementById('createProjectBtn')?.addEventListener('click',createProjectHandler);
    document.getElementById('cancelProjectBtn')?.addEventListener('click',closeProjectModal);
    document.getElementById('saveProjectDeadlineBtn')?.addEventListener('click',saveProjectDeadline);
    document.getElementById('cancelEditDeadlineBtn')?.addEventListener('click',closeEditDeadlineModal);
    document.getElementById('notificationOkBtn')?.addEventListener('click',()=>document.getElementById('notificationModal').style.display='none');
    document.getElementById('closeAuditBtn')?.addEventListener('click',()=>document.getElementById('auditModal').style.display='none');
    document.getElementById('confirmDeleteUserBtn')?.addEventListener('click',confirmDeleteUser);
    document.getElementById('cancelDeleteUserBtn')?.addEventListener('click',()=>document.getElementById('deleteUserModal').style.display='none');
    document.getElementById('notificationBell')?.addEventListener('click',(e)=>toggleNotificationDropdown(e));
    document.addEventListener('click',function(e){const dropdown=document.getElementById('notificationDropdown');const bell=document.getElementById('notificationBell');if(!dropdown||!bell)return;if(!dropdown.contains(e.target)&&!bell.contains(e.target)&&dropdown.classList.contains('show'))dropdown.classList.remove('show')});
    document.getElementById('frm')?.addEventListener('submit',async e=>{e.preventDefault();const fid=document.getElementById('fid').value;const fpr=document.getElementById('fpr').value;const fdd=document.getElementById('fdd').value;const fp=document.getElementById('fpriority').value;const fb=document.getElementById('fbudget').value;const fbd=document.getElementById('fbd').value;if(!fpr||!fdd||!fbd)return showNotification("Заполните поля");const proj=D.projects.active.find(p=>p.name===fpr);if(proj&&proj.deadline&&new Date(fdd)>new Date(proj.deadline))return showNotification("Срок не может превышать срок проекта");let res;if(fid)res=await api('/edt',{id:fid,body:fbd,deadline:fdd,priority:fp,budget:fb||null});else res=await api('/req',{project:fpr,deadline:fdd,body:fbd,priority:fp,budget:fb||null,author:U.name});if(res.error)showNotification(res.error);else{if(res.id&&selectedFilesForUpload.length>0)await uploadFilesForRequest(res.id);closeModal();playSound('success');await loadData();await updateUnreadCount()}});
    try{const saved=localStorage.getItem('z1u');const token=localStorage.getItem('z1t');if(saved&&token){U=JSON.parse(saved);TOKEN=token;showApp();loadData();updateUnreadCount();requestNotificationPermission();startNotificationChecker()}}catch(e){}
};
</script>
</body>
</html>
'@

# ==========================================================
# 🌐 БЛОК 10: СЕРВЕР
# ==========================================================
function Add-UrlAcl { param([string]$url); try { $existing = netsh http show urlacl | Select-String $url; if (-not $existing) { netsh http add urlacl url=$url user=Everyone listen=yes delegate=yes | Out-Null } } catch { Write-Host "⚠️ URL ACL: $_" -ForegroundColor Yellow } }

Repair-Database
Migrate-Database

$localIps = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "Loopback" -and $_.IPAddress -notlike "169.254*" }).IPAddress
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$port/")
$listener.Prefixes.Add("http://localhost:$port/")
foreach ($ip in $localIps) { $listener.Prefixes.Add("http://$ip`:$port/") }
Add-UrlAcl -url "http://+:$port/"
foreach ($ip in $localIps) { Add-UrlAcl -url "http://$ip`:$port/" }

try {
    $listener.Start()
    Write-Host "`n🟢 ФОРМА ЗАЯВКИ 3.7.9 ЗАПУЩЕНА!`n" -ForegroundColor Green
    Write-Host "📍 http://localhost:$port" -ForegroundColor Cyan
    foreach ($ip in $localIps) { Write-Host "📍 http://$ip`:$port" -ForegroundColor Yellow }
    Write-Host "`n💡 Firewall: New-NetFirewallRule -DisplayName 'Zayavka379' -Direction Inbound -Protocol TCP -LocalPort $port -Action Allow`n" -ForegroundColor White
} catch {
    Write-Host "`n❌ Ошибка: $_" -ForegroundColor Red
    Write-Host "💡 Запустите от Администратора" -ForegroundColor Yellow
    exit 1
}

function Send-Json($obj, $code=200) { $ctx.Response.StatusCode = $code; $json = if ($obj) { $obj | ConvertTo-Json -Depth 15 } else { '{"error":"null"}' }; $bytes = [System.Text.Encoding]::UTF8.GetBytes($json); $ctx.Response.ContentType = "application/json; charset=utf-8"; $ctx.Response.ContentLength64 = $bytes.Length; $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) }
function Read-Body { $reader = New-Object System.IO.StreamReader($ctx.Request.InputStream, [System.Text.Encoding]::UTF8); $body = $reader.ReadToEnd(); $reader.Close(); return $body | ConvertFrom-Json }
function Get-AuthUser { $token = $ctx.Request.Headers["X-Token"]; if ($token) { return Get-UserByToken $token }; $username = $ctx.Request.Headers["X-Username"]; if ($username) { return Get-User $username }; return $null }

while ($true) {
    Export-Excel
    try {
        $ctx = $listener.GetContext()
        $m = $ctx.Request.HttpMethod; $p = $ctx.Request.Url.AbsolutePath
        Write-Host "[HTTP] $m $p - $($ctx.Request.RemoteEndPoint.Address)" -ForegroundColor DarkGray
        if ($p -eq "/favicon.ico" -and $m -eq "GET") { $ctx.Response.StatusCode = 204; continue }
        if ($m -eq "GET" -and $p -match "^/$") { $bytes = [System.Text.Encoding]::UTF8.GetBytes($ui); $ctx.Response.ContentType = "text/html; charset=utf-8"; $ctx.Response.ContentLength64 = $bytes.Length; $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) }
        elseif ($p -match "^/download/(.+)$" -and $m -eq "GET") { $fn = [Uri]::UnescapeDataString($matches[1]); $fp = Join-Path $uploadsDir $fn; if (Test-Path $fp) { $bytes = [System.IO.File]::ReadAllBytes($fp); $ctx.Response.ContentType = "application/octet-stream"; $ctx.Response.AddHeader("Content-Disposition", "attachment; filename*=UTF-8''$([Uri]::EscapeDataString($fn))"); $ctx.Response.ContentLength64 = $bytes.Length; $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) } else { Send-Json @{error="Файл не найден"} 404 } }
        elseif ($p -eq "/reg" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Register $b.username $b.password $b.role $b.name) }
        elseif ($p -eq "/vrf" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Verify $b.username $b.code) }
        elseif ($p -eq "/log" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-Login $b.username $b.password) }
        elseif ($p -eq "/recover-request" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-RecoverRequest $b.username) }
        elseif ($p -eq "/recover-reset" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-RecoverReset $b.username $b.code $b.newPassword) }
        elseif ($p -eq "/rsnd" -and $m -eq "POST") { $b = Read-Body; Send-Json (Auth-RecoverRequest $b.username) }
        elseif ($p -eq "/dat" -and $m -eq "GET") { $u = Get-AuthUser; if($u){Send-Json (Get-Db)}else{Send-Json @{error="Auth"} 401} }
        elseif ($p -eq "/updates" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if ($u) { Send-Json (Get-UserUpdates $u.username $b.lastCheck) } else { Send-Json @{error="Auth"} 401 } }
        elseif ($p -eq "/dashboard" -and $m -eq "GET") { $u = Get-AuthUser; if ($u) { Send-Json (Get-DashboardData $u.role) } else { Send-Json @{error="Auth"} 401 } }
        elseif ($p -eq "/export-excel" -and $m -eq "POST") { $u = Get-AuthUser; if ($u -and $u.role -eq "director") { Export-Excel $true; Send-Json @{ok=$true} } else { Send-Json @{error="Нет прав"} 403 } }
        elseif ($p -eq "/prj" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; try { Send-Json (Action-CreateProject $b.name $b.deadline $u.role) } catch { Send-Json @{error="$($_.Exception.Message)"} 500 } }
        elseif ($p -eq "/update-project-deadline" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json (Action-UpdateProjectDeadline $b.name $b.deadline $u.role) }
        elseif ($p -eq "/arc" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; Send-Json (Action-Archive $b.name $u.role) }
        elseif ($p -eq "/delete-archive" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; Send-Json (Action-DeleteFromArchive $b.name $u.role) }
        elseif ($p -eq "/delete-user" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json (Action-DeleteUser $b.username $u.role) }
        elseif ($p -eq "/req" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; if ($u.role -in @('client','procurement')) { Send-Json (Action-CreateRequest $b.project $b.deadline $b.body $u.name $b.priority $b.budget $u.role) } else { Send-Json @{error="Access Denied"} 403 } }
        elseif ($p -eq "/sts" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; $comment = if ($b.comment -ne $null) { $b.comment.ToString() } else { " " }; Send-Json (Action-Process $b.id $b.status $u.name $u.role $comment) }
        elseif ($p -eq "/bulk-process" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json (Action-BulkProcess $b.ids $b.status $u.name $u.role $b.comment) }
        elseif ($p -eq "/rjt" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; $reason = if ($b.reason -ne $null) { $b.reason.ToString() } else { " " }; Send-Json (Action-Reject $b.id $reason $u.name $u.role) }
        elseif ($p -eq "/edt" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json (Action-EditRequest $b.id $b.body $b.deadline $b.priority $u.name $b.budget) }
        elseif ($p -eq "/upload-file" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; $fileBytes = [Convert]::FromBase64String($b.fileData); Send-Json (Action-UploadFile $b.requestId $b.fileName $fileBytes $u.name) }
        elseif ($p -eq "/mark-read" -and $m -eq "POST") { $b = Read-Body; $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json @{ok = MarkRequestAsRead $u.username $b.requestId} }
        elseif ($p -eq "/unread-count" -and $m -eq "GET") { $u = Get-AuthUser; if (-not $u) { Send-Json @{error="Auth"} 401; continue }; Send-Json @{count = GetUnreadRequestsCount $u.username} }
        else { Send-Json @{error="Not Found"} 404 }
    } catch { Write-Host "❌ $_" -ForegroundColor Red; try { Send-Json @{error="Server: $_"} 500 } catch {} }
    finally { try { $ctx.Response.Close() } catch {} }
    Start-Sleep -Milliseconds 20
}
