# ============================================================
#  GestorCertificados.ps1
#  Herramienta para activar/desactivar certificados digitales
#  Mueve certificados entre almacén Personal (visible en navegador)
#  y Otras Personas (oculto en navegador)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Constantes de almacenes ──────────────────────────────────
$STORE_PERSONAL    = "My"
$STORE_INACTIVO    = "AddressBook"   # "Otras personas" en certmgr.msc

# ── Paleta de colores ────────────────────────────────────────
$COLOR_BG          = [System.Drawing.Color]::FromArgb(245, 247, 250)
$COLOR_PANEL       = [System.Drawing.Color]::White
$COLOR_ACTIVO      = [System.Drawing.Color]::FromArgb(220, 242, 231)   # verde suave
$COLOR_ACTIVO_TXT  = [System.Drawing.Color]::FromArgb(22, 120, 60)
$COLOR_INACTIVO    = [System.Drawing.Color]::FromArgb(240, 240, 240)
$COLOR_INACTIVO_TXT= [System.Drawing.Color]::FromArgb(110, 110, 110)
$COLOR_ACCENT      = [System.Drawing.Color]::FromArgb(37, 99, 235)     # azul
$COLOR_DANGER      = [System.Drawing.Color]::FromArgb(220, 38, 38)     # rojo
$COLOR_HEADER      = [System.Drawing.Color]::FromArgb(30, 41, 59)      # azul oscuro
$COLOR_CADUCADO    = [System.Drawing.Color]::FromArgb(254, 226, 226)

# ════════════════════════════════════════════════════════════
#  FUNCIONES DE CERTIFICADOS
# ════════════════════════════════════════════════════════════

function Get-AllCertificates {
    $list = [System.Collections.Generic.List[PSObject]]::new()

    $stores = @(
        [PSCustomObject]@{ Name = $STORE_PERSONAL; Label = "Activo" },
        [PSCustomObject]@{ Name = $STORE_INACTIVO; Label = "Inactivo" }
    )

    foreach ($s in $stores) {
        try {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s.Name, "CurrentUser")
            $store.Open("ReadOnly")
            foreach ($cert in $store.Certificates) {
                # Ignorar certificados sin clave privada si están en Personal (raíces, etc.)
                $nombre = if ($cert.FriendlyName -and $cert.FriendlyName.Trim() -ne "") {
                    $cert.FriendlyName.Trim()
                } else {
                    # Extraer CN del Subject como fallback
                    if ($cert.Subject -match "CN=([^,]+)") { $Matches[1].Trim() } else { $cert.Subject }
                }

                $list.Add([PSCustomObject]@{
                    NombreDescriptivo = $nombre
                    Estado            = $s.Label
                    StoreName         = $s.Name
                    Vencimiento       = $cert.NotAfter
                    VencimientoStr    = $cert.NotAfter.ToString("dd/MM/yyyy")
                    Caducado          = ($cert.NotAfter -lt (Get-Date))
                    Thumbprint        = $cert.Thumbprint
                    Subject           = $cert.Subject
                    CertObject        = $cert
                })
            }
            $store.Close()
        } catch {
            # Almacén no accesible, se ignora
        }
    }

    return $list
}

function Move-Certificate {
    param(
        [PSObject]$CertInfo,
        [string]$TargetStoreName
    )
    try {
        # Quitar del almacén origen
        $src = New-Object System.Security.Cryptography.X509Certificates.X509Store($CertInfo.StoreName, "CurrentUser")
        $src.Open("ReadWrite")
        $src.Remove($CertInfo.CertObject)
        $src.Close()

        # Añadir al almacén destino
        $dst = New-Object System.Security.Cryptography.X509Certificates.X509Store($TargetStoreName, "CurrentUser")
        $dst.Open("ReadWrite")
        $dst.Add($CertInfo.CertObject)
        $dst.Close()

        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error al mover el certificado:`n$($_.Exception.Message)",
            "Error", "OK", "Error") | Out-Null
        return $false
    }
}

# ════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DEL FORMULARIO
# ════════════════════════════════════════════════════════════

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Gestor de Certificados Digitales"
$form.Size            = New-Object System.Drawing.Size(860, 620)
$form.MinimumSize     = New-Object System.Drawing.Size(700, 500)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $COLOR_BG
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ── TableLayoutPanel: 3 filas sin solapamiento ───────────────
$tbl = New-Object System.Windows.Forms.TableLayoutPanel
$tbl.Dock        = "Fill"
$tbl.ColumnCount = 1
$tbl.RowCount    = 3
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))  # búsqueda
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # grid
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))  # botones
$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tbl.Padding     = New-Object System.Windows.Forms.Padding(0)
$tbl.Margin      = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($tbl)

# ── Panel de búsqueda ────────────────────────────────────────
$pnlSearch = New-Object System.Windows.Forms.Panel
$pnlSearch.Dock      = "Fill"
$pnlSearch.BackColor = $COLOR_PANEL
$pnlSearch.Padding   = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text     = "🔍  Buscar:"
$lblSearch.Location = New-Object System.Drawing.Point(12, 16)
$lblSearch.AutoSize = $true
$lblSearch.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlSearch.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location  = New-Object System.Drawing.Point(80, 12)
$txtSearch.Width     = 340
$txtSearch.Height    = 28
$txtSearch.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSearch.BackColor = $COLOR_BG
$pnlSearch.Controls.Add($txtSearch)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text      = "Escribe nombre de empresa o CIF"
$lblHint.Location  = New-Object System.Drawing.Point(430, 16)
$lblHint.AutoSize  = $true
$lblHint.ForeColor = [System.Drawing.Color]::Gray
$lblHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$pnlSearch.Controls.Add($lblHint)

$chkSoloActivos = New-Object System.Windows.Forms.CheckBox
$chkSoloActivos.Text     = "Solo activos"
$chkSoloActivos.Location = New-Object System.Drawing.Point(640, 14)
$chkSoloActivos.AutoSize = $true
$pnlSearch.Controls.Add($chkSoloActivos)

# ── DataGridView ─────────────────────────────────────────────
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock                    = "Fill"
$grid.BackgroundColor         = $COLOR_BG
$grid.BorderStyle             = "None"
$grid.RowHeadersVisible       = $false
$grid.AllowUserToAddRows      = $false
$grid.AllowUserToDeleteRows   = $false
$grid.ReadOnly                = $true
$grid.SelectionMode           = "FullRowSelect"
$grid.MultiSelect             = $false
$grid.AutoSizeRowsMode        = "AllCells"
$grid.ColumnHeadersHeightSizeMode = "AutoSize"
$grid.DefaultCellStyle.Font   = New-Object System.Drawing.Font("Segoe UI", 9.5)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$grid.ColumnHeadersDefaultCellStyle.Font        = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.BackColor   = $COLOR_HEADER
$grid.ColumnHeadersDefaultCellStyle.ForeColor   = [System.Drawing.Color]::White
$grid.EnableHeadersVisualStyles = $false
$grid.GridColor = [System.Drawing.Color]::FromArgb(226, 232, 240)

# ── Panel de botones de acción ───────────────────────────────
$pnlActions = New-Object System.Windows.Forms.Panel
$pnlActions.Dock      = "Fill"
$pnlActions.BackColor = $COLOR_PANEL

$btnActivar = New-Object System.Windows.Forms.Button
$btnActivar.Text      = "✅  ACTIVAR  (mover a Personal)"
$btnActivar.Location  = New-Object System.Drawing.Point(12, 12)
$btnActivar.Size      = New-Object System.Drawing.Size(240, 34)
$btnActivar.BackColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
$btnActivar.ForeColor = [System.Drawing.Color]::White
$btnActivar.FlatStyle = "Flat"
$btnActivar.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
$btnActivar.FlatAppearance.BorderSize = 0
$btnActivar.Enabled   = $false
$pnlActions.Controls.Add($btnActivar)

$btnDesactivar = New-Object System.Windows.Forms.Button
$btnDesactivar.Text      = "⏸  DESACTIVAR  (mover a Otras personas)"
$btnDesactivar.Location  = New-Object System.Drawing.Point(264, 12)
$btnDesactivar.Size      = New-Object System.Drawing.Size(270, 34)
$btnDesactivar.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$btnDesactivar.ForeColor = [System.Drawing.Color]::White
$btnDesactivar.FlatStyle = "Flat"
$btnDesactivar.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
$btnDesactivar.FlatAppearance.BorderSize = 0
$btnDesactivar.Enabled   = $false
$pnlActions.Controls.Add($btnDesactivar)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text      = "🔄  Actualizar lista"
$btnRefresh.Location  = New-Object System.Drawing.Point(548, 12)
$btnRefresh.Size      = New-Object System.Drawing.Size(160, 34)
$btnRefresh.BackColor = $COLOR_ACCENT
$btnRefresh.ForeColor = [System.Drawing.Color]::White
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$btnRefresh.FlatAppearance.BorderSize = 0
$pnlActions.Controls.Add($btnRefresh)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(720, 18)
$lblStatus.AutoSize  = $true
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$pnlActions.Controls.Add($lblStatus)

# ── Añadir al TableLayoutPanel en orden de filas ─────────────
$tbl.Controls.Add($pnlSearch,  0, 0)
$tbl.Controls.Add($grid,       0, 1)
$tbl.Controls.Add($pnlActions, 0, 2)

# Columnas
$colNombre = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNombre.Name            = "NombreDescriptivo"
$colNombre.HeaderText      = "Nombre descriptivo  (empresa / CIF)"
$colNombre.AutoSizeMode    = "Fill"
$colNombre.MinimumWidth    = 300
$colNombre.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$colEstado = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEstado.Name            = "Estado"
$colEstado.HeaderText      = "Estado"
$colEstado.Width           = 110
$colEstado.DefaultCellStyle.Alignment = "MiddleCenter"

$colVence = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVence.Name             = "VencimientoStr"
$colVence.HeaderText       = "Vence el"
$colVence.Width            = 110
$colVence.DefaultCellStyle.Alignment = "MiddleCenter"

$grid.Columns.AddRange($colNombre, $colEstado, $colVence)

# ════════════════════════════════════════════════════════════
#  LÓGICA DE DATOS
# ════════════════════════════════════════════════════════════

$script:CertCache = @()

function Refresh-CertList {
    $lblStatus.Text = "Cargando..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:CertCache = Get-AllCertificates
    Apply-Filter
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $total   = $script:CertCache.Count
    $activos = ($script:CertCache | Where-Object { $_.Estado -eq "Activo" }).Count
    $lblStatus.Text = "$total certificados  |  $activos activos"
}

function Normalize-Text {
    param([string]$text)
    if (-not $text) { return "" }
    $n = $text.ToLower()
    $n = $n -replace '[áàäâã]','a'
    $n = $n -replace '[éèëê]','e'
    $n = $n -replace '[íìïî]','i'
    $n = $n -replace '[óòöôõ]','o'
    $n = $n -replace '[úùüû]','u'
    $n = $n -replace '[ñ]','n'
    return $n
}

function Apply-Filter {
    $filtro  = Normalize-Text ($txtSearch.Text.Trim())
    $soloAct = $chkSoloActivos.Checked

    # Construir lista filtrada con foreach para evitar problemas de scope
    $filtered = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($c in $script:CertCache) {
        $campos = (Normalize-Text $c.NombreDescriptivo) + " " + (Normalize-Text $c.Subject)

        $matchFiltro = ($filtro -eq "") -or ($campos.Contains($filtro))
        $matchEstado = (-not $soloAct) -or ($c.Estado -eq "Activo")

        if ($matchFiltro -and $matchEstado) {
            $filtered.Add($c)
        }
    }

    $grid.Rows.Clear()
    foreach ($c in $filtered) {
        $idx = $grid.Rows.Add($c.NombreDescriptivo, $c.Estado, $c.VencimientoStr)
        $row = $grid.Rows[$idx]
        $row.Tag = $c   # guardar el objeto completo en Tag para recuperarlo después

        if ($c.Caducado) {
            $row.DefaultCellStyle.BackColor = $COLOR_CADUCADO
            $row.DefaultCellStyle.ForeColor = $COLOR_DANGER
        } elseif ($c.Estado -eq "Activo") {
            $row.Cells["Estado"].Style.BackColor = $COLOR_ACTIVO
            $row.Cells["Estado"].Style.ForeColor = $COLOR_ACTIVO_TXT
            $row.Cells["Estado"].Style.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
        } else {
            $row.Cells["Estado"].Style.BackColor = $COLOR_INACTIVO
            $row.Cells["Estado"].Style.ForeColor = $COLOR_INACTIVO_TXT
        }
    }

    # Forzar scroll al inicio
    $grid.ClearSelection()
    if ($grid.Rows.Count -gt 0) {
        $grid.FirstDisplayedScrollingRowIndex = 0
    }

    # Deshabilitar botones al cambiar filtro (pierde selección)
    $btnActivar.Enabled    = $false
    $btnDesactivar.Enabled = $false
}

# ════════════════════════════════════════════════════════════
#  EVENTOS
# ════════════════════════════════════════════════════════════

$txtSearch.Add_TextChanged({ Apply-Filter })
$chkSoloActivos.Add_CheckedChanged({ Apply-Filter })

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -gt 0) {
        $cert = $grid.SelectedRows[0].Tag
        if ($cert) {
            $btnActivar.Enabled    = ($cert.Estado -eq "Inactivo")
            $btnDesactivar.Enabled = ($cert.Estado -eq "Activo")
        }
    } else {
        $btnActivar.Enabled    = $false
        $btnDesactivar.Enabled = $false
    }
})

$btnActivar.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $cert = $grid.SelectedRows[0].Tag
    if (-not $cert) { return }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "¿Activar el certificado?`n`n$($cert.NombreDescriptivo)`n`nSe moverá al almacén Personal y será visible en el navegador.",
        "Confirmar activación", "YesNo", "Question")

    if ($res -eq "Yes") {
        if (Move-Certificate -CertInfo $cert -TargetStoreName $STORE_PERSONAL) {
            $lblStatus.Text = "✅ Activado: $($cert.NombreDescriptivo)"
            Refresh-CertList
        }
    }
})

$btnDesactivar.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $cert = $grid.SelectedRows[0].Tag
    if (-not $cert) { return }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "¿Desactivar el certificado?`n`n$($cert.NombreDescriptivo)`n`nSe moverá a 'Otras personas' y quedará oculto en el navegador.",
        "Confirmar desactivación", "YesNo", "Question")

    if ($res -eq "Yes") {
        if (Move-Certificate -CertInfo $cert -TargetStoreName $STORE_INACTIVO) {
            $lblStatus.Text = "⏸ Desactivado: $($cert.NombreDescriptivo)"
            Refresh-CertList
        }
    }
})

$btnRefresh.Add_Click({ Refresh-CertList })

$form.Add_Shown({
    Refresh-CertList
    if ($grid.Rows.Count -gt 0) {
        $grid.FirstDisplayedScrollingRowIndex = 0
        $grid.ClearSelection()
    }
})

# ════════════════════════════════════════════════════════════
#  ARRANCAR
# ════════════════════════════════════════════════════════════
[System.Windows.Forms.Application]::Run($form)
