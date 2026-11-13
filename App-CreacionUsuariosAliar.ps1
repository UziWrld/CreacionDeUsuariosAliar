<#
========================================================================================
 APLICACIÓN DE ESCRITORIO CORPORATIVA: GESTIÓN ACTIVE DIRECTORY
 Empresa: Aliar S.A. (aliar.com.co)
 Módulo: Aprovisionamiento y Gestión de Permisos en Dominio
========================================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = $PSScriptRoot
if (-not $scriptDir -and $MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptDir -and $MyInvocation.MyCommand.Definition -and (Test-Path -Path $MyInvocation.MyCommand.Definition -ErrorAction SilentlyContinue)) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
if (-not $scriptDir) {
    $scriptDir = (Get-Location).Path
}
$jsonPath = Join-Path $scriptDir "aliar_ad_data.json"

# Captura de Excepciones Globales de WinForms y .NET
[System.AppDomain]::CurrentDomain.add_UnhandledException({
    param($s, $e)
    try {
        [System.Windows.Forms.MessageBox]::Show("Ocurrió un error inesperado en la aplicación:

$($e.ExceptionObject)", "Error Crítico", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {}
})

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    try {
        [System.Windows.Forms.MessageBox]::Show("Ocurrió un error en la interfaz:

$($e.Exception.Message)", "Error de Interfaz", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {}
})
# ======================================================================================
# 0. FUNCIONES DE RENDIMIENTO Y OPTIMIZACIÓN DE DIBUJADO (DOUBLE BUFFERING)
# ======================================================================================
function Enable-DoubleBuffer($ctrl) {
    try {
        $prop = [System.Windows.Forms.Control].GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"NonPublic, Instance")
        if ($prop) { $prop.SetValue($ctrl, $true, $null) }
    } catch {}
}

# ======================================================================================
# 1. CARGA DE DATOS REALES DE ACTIVE DIRECTORY (ALIAR S.A.) - ALTA VELOCIDAD CON DICCIONARIOS
# ======================================================================================
$global:RealGroups = New-Object System.Collections.Generic.List[PSCustomObject]
$global:RealUsers = New-Object System.Collections.Generic.List[PSCustomObject]
$global:AreaFolderGroups = New-Object System.Collections.Generic.List[PSCustomObject]
$global:OtherGroups = New-Object System.Collections.Generic.List[PSCustomObject]

# Diccionarios O(1) para búsquedas instantáneas sin congelamiento de UI
$global:UserLookupBySam = New-Object 'System.Collections.Generic.Dictionary[string, PSCustomObject]'
$global:UserLookupByName = New-Object 'System.Collections.Generic.Dictionary[string, PSCustomObject]'

function Clean-Text($str) {
    if (-not $str) { return "" }
    $norm = $str.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $norm.Length; $i++) {
        $c = $norm[$i]
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().ToLower() -replace "[^a-z0-9]", ""
}

if (Test-Path $jsonPath) {
    try {
        $rawJson = [System.IO.File]::ReadAllText($jsonPath, [System.Text.Encoding]::UTF8)
        $data = $rawJson | ConvertFrom-Json
        
        if ($data.groups) {
            foreach ($g in $data.groups) {
                $global:RealGroups.Add($g)
                if ($g.name -like "GR_DC_*") {
                    $global:AreaFolderGroups.Add($g)
                } else {
                    $global:OtherGroups.Add($g)
                }
            }
        }
        
        if ($data.users) {
            foreach ($u in $data.users) {
                $sede = ""
                if ($u.n -match '-\s*([A-Za-z0-9_]+)\s*$') {
                    $sede = $matches[1]
                }
                $uGroups = if ($u.g) { $u.g } else { @() }
                $userObj = [PSCustomObject]@{
                    name = $u.n
                    user = $u.u
                    groups = $uGroups
                    department = $u.d
                    title = $u.t
                    sede = $sede
                    searchStr = ("$($u.n) $($u.u) $sede $($u.d)").ToLower()
                }
                $global:RealUsers.Add($userObj)
                
                if ($u.u) {
                    $samKey = $u.u.Trim().ToLower()
                    if (-not $global:UserLookupBySam.ContainsKey($samKey)) {
                        $global:UserLookupBySam.Add($samKey, $userObj)
                    }
                }
                if ($u.n) {
                    $nameKey = Clean-Text $u.n
                    if ($nameKey -and (-not $global:UserLookupByName.ContainsKey($nameKey))) {
                        $global:UserLookupByName.Add($nameKey, $userObj)
                    }
                }
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error cargando base de datos local aliar_ad_data.json:`n$_", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
} else {
}

# Unidades Organizativas (OU) y Sedes Oficiales de Active Directory (Ruta: USUARIOS en aliar.com.co)
$global:SedesConfig = @(
    # CUNDINAMARCA / BOGOTA
    [PSCustomObject]@{ OU = "AL_CUN_PL_ENGATIVA";         Code = "CUN_PLENG";  Label = "AL_CUN_PL_ENGATIVA - Planta Engativa (Bogota)" }
    [PSCustomObject]@{ OU = "AL_CUN_PL_GUADALUPE";        Code = "CUN_PLGUA";  Label = "AL_CUN_PL_GUADALUPE - Planta Guadalupe (Cundinamarca)" }
    [PSCustomObject]@{ OU = "AL_CUN_DIST_BOGOTA";         Code = "CUN_DISBOG"; Label = "AL_CUN_DIST_BOGOTA - Distribucion Bogota" }
    [PSCustomObject]@{ OU = "AL_CUN_DIST_PALOQUEMAO";     Code = "CUN_DISPLQ"; Label = "AL_CUN_DIST_PALOQUEMAO - Distribucion Paloquemao" }
    [PSCustomObject]@{ OU = "AL_CUN_ADMON_BOGOTA";        Code = "CUN_ADM";    Label = "AL_CUN_ADMON_BOGOTA - Administracion Bogota" }
    [PSCustomObject]@{ OU = "AL_CUN_PV_CUNDINAMARCA";     Code = "CUN_PV";     Label = "AL_CUN_PV_CUNDINAMARCA - Puntos de Venta Cundinamarca" }
    
    # META / PUERTO GAITAN / VILLAVICENCIO
    [PSCustomObject]@{ OU = "AL_MET_FRIGORIFICO";         Code = "MET_FRIG";   Label = "AL_MET_FRIGORIFICO - Planta Frigorifico (Puerto Gaitan)" }
    [PSCustomObject]@{ OU = "AL_MET_PL_AGROINDUSTRIALES"; Code = "MET_PLAGR";  Label = "AL_MET_PL_AGROINDUSTRIALES - Planta Alimentos Concentrados" }
    [PSCustomObject]@{ OU = "AL_MET_PORCICULTURA";        Code = "MET_PORCI";  Label = "AL_MET_PORCICULTURA - Granjas Porcicultura (La Fazenda)" }
    [PSCustomObject]@{ OU = "AL_MET_AGRICULTURA";         Code = "MET_AGRI";   Label = "AL_MET_AGRICULTURA - Agricultura y Granos (Puerto Gaitan)" }
    [PSCustomObject]@{ OU = "AL_MET_GANADERIA";           Code = "MET_GAN";    Label = "AL_MET_GANADERIA - Ganaderia (Meta)" }
    [PSCustomObject]@{ OU = "AL_MET_TALLERES_AGR";        Code = "MET_TALLER"; Label = "AL_MET_TALLERES_AGR - Talleres Agricolas y Maquinaria" }
    [PSCustomObject]@{ OU = "AL_MET_MANT_PLANTAS";        Code = "MET_MANT";   Label = "AL_MET_MANT_PLANTAS - Mantenimiento Plantas (Meta)" }
    [PSCustomObject]@{ OU = "AL_MET_OBRAS_CIVILES";       Code = "MET_OBRAS";  Label = "AL_MET_OBRAS_CIVILES - Obras Civiles y Construccion (Meta)" }
    [PSCustomObject]@{ OU = "AL_MET_ADMON_FINCAS";        Code = "MET_ADM";    Label = "AL_MET_ADMON_FINCAS - Administracion Fincas (Gaitan/Vcio)" }
    [PSCustomObject]@{ OU = "AL_MET_DIST_VILLAVICENCIO";  Code = "MET_DISVIL"; Label = "AL_MET_DIST_VILLAVICENCIO - Distribucion Villavicencio" }
    [PSCustomObject]@{ OU = "AL_MET_PV_META";             Code = "MET_PV";     Label = "AL_MET_PV_META - Puntos de Venta Meta y Llanos" }
    
    # SANTANDER / BUCARAMANGA / PIEDECUESTA
    [PSCustomObject]@{ OU = "AL_SAN_OFICINA_PPAL";        Code = "SAN_OFPP";   Label = "AL_SAN_OFICINA_PPAL - Oficina Principal Piedecuesta" }
    [PSCustomObject]@{ OU = "AL_SAN_DIST_BMANGA";         Code = "SAN_DISBGA"; Label = "AL_SAN_DIST_BMANGA - Distribucion Bucaramanga" }
    [PSCustomObject]@{ OU = "AL_SAN_PV_SANTANDER";        Code = "SAN_PV";     Label = "AL_SAN_PV_SANTANDER - Puntos de Venta Santander" }
    
    # ANTIOQUIA / MEDELLIN
    [PSCustomObject]@{ OU = "AL_ANT_DIST_MEDELLIN";       Code = "ANT_DISMED"; Label = "AL_ANT_DIST_MEDELLIN - Distribucion Medellin" }
    [PSCustomObject]@{ OU = "AL_ANT_PV_ANTIOQUIA";        Code = "ANT_PV";     Label = "AL_ANT_PV_ANTIOQUIA - Puntos de Venta Antioquia" }
    
    # ATLANTICO / BARRANQUILLA
    [PSCustomObject]@{ OU = "AL_ATL_DIST_BARRANQUILLA";   Code = "ATL_DISBQA"; Label = "AL_ATL_DIST_BARRANQUILLA - Distribucion Barranquilla" }
    [PSCustomObject]@{ OU = "AL_ATL_PV_ATLANTICO";        Code = "ATL_PV";     Label = "AL_ATL_PV_ATLANTICO - Puntos de Venta Atlantico" }
    
    # OTROS PUNTOS DE VENTA Y REGIONALES
    [PSCustomObject]@{ OU = "AL_CES_PV_CESAR";            Code = "CES_PV";     Label = "AL_CES_PV_CESAR - Puntos de Venta Cesar (Valledupar)" }
    [PSCustomObject]@{ OU = "AL_CAS_PV_CASANARE";         Code = "CAS_PV";     Label = "AL_CAS_PV_CASANARE - Puntos de Venta Casanare (Yopal/Villanueva)" }
    [PSCustomObject]@{ OU = "AL_BOL_PV_BOLIVAR";          Code = "BOL_PV";     Label = "AL_BOL_PV_BOLIVAR - Puntos de Venta Bolivar (Cartagena)" }
    [PSCustomObject]@{ OU = "AL_VIC_ADMON_FINCAS";        Code = "VIC_ADMON";  Label = "AL_VIC_ADMON_FINCAS - Administracion Fincas (Vichada)" }
    
    # ESPECIALES / EXTERNOS / PRACTICANTES
    [PSCustomObject]@{ OU = "AL_PRACTICANTES";            Code = "PRACTICANTE"; Label = "AL_PRACTICANTES - Practicantes y Aprendices SENA" }
    [PSCustomObject]@{ OU = "AL_ESPECIALES";              Code = "ESPECIALES";  Label = "AL_ESPECIALES - Cuentas Especiales / Genericas" }
    [PSCustomObject]@{ OU = "AL_EXTERNOS";                Code = "EXTERNOS";    Label = "AL_EXTERNOS - Personal Externo / Contratistas" }
    [PSCustomObject]@{ OU = "AL_NEXO";                    Code = "NEXO";        Label = "AL_NEXO - Nexo" }
    [PSCustomObject]@{ OU = "AL_INACTIVOS";               Code = "INACTIVOS";   Label = "AL_INACTIVOS - Inactivos / Retirados" }
)

$global:SedesAliar = @($global:SedesConfig | ForEach-Object { $_.Label })
$global:SedesCodigos = @($global:SedesConfig | ForEach-Object { $_.Code })
$global:SedesOUs = @($global:SedesConfig | ForEach-Object { $_.OU })

$global:AD_OU_Map = @{}
$global:UsuariosBaseOU = "OU=USUARIOS,OU=E_ALIAR,DC=aliar,DC=com,DC=co"

function Get-TargetOUPath($ouOrCode) {
    $ouName = "AL_CUN_PL_ENGATIVA"
    if ($ouOrCode) {
        $str = $ouOrCode.ToString().Trim()
        $found = $global:SedesConfig | Where-Object { $_.OU -eq $str -or $_.Code -eq $str -or $_.Label -eq $str -or $_.Label.StartsWith($str) } | Select-Object -First 1
        if ($found) {
            $ouName = $found.OU
        } elseif ($str.StartsWith("AL_")) {
            $ouName = $str
        } elseif ($str -match '^([A-Za-z0-9_]+)') {
            $raw = $matches[1]
            $found2 = $global:SedesConfig | Where-Object { $_.OU -like "*$raw*" -or $_.Code -eq $raw } | Select-Object -First 1
            if ($found2) { $ouName = $found2.OU }
        }
    }
    
    # 1. Si ya se descubrio el DistinguishedName exacto en vivo en el DA
    if ($global:AD_OU_Map -and $global:AD_OU_Map.ContainsKey($ouName)) {
        return $global:AD_OU_Map[$ouName]
    }
    
    # 2. Consultar en vivo en Active Directory si esta disponible
    try {
        if (Get-Command Get-ADOrganizationalUnit -ErrorAction SilentlyContinue) {
            $ouObj = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction Stop | Select-Object -First 1
            if ($ouObj) {
                if (-not $global:AD_OU_Map) { $global:AD_OU_Map = @{} }
                $global:AD_OU_Map[$ouName] = $ouObj.DistinguishedName
                return $ouObj.DistinguishedName
            }
        }
    } catch {}
    
    # 3. Fallback con base USUARIOS oficial
    return "OU=$ouName,OU=USUARIOS,OU=E_ALIAR,DC=aliar,DC=com,DC=co"
}

function Get-SelectedSedeInfo {
    $text = if ($cmbSede.Text) { $cmbSede.Text.Trim() } else { "" }
    foreach ($item in $global:SedesConfig) {
        if ($text -eq $item.Label -or $text -eq $item.OU -or $text -eq $item.Code -or $text.StartsWith($item.OU)) {
            return $item
        }
    }
    if ($text -match '^([A-Za-z0-9_]+)') {
        $raw = $matches[1].ToUpper()
        $ouVal = if ($raw.StartsWith("AL_")) { $raw } else { "AL_$raw" }
        return [PSCustomObject]@{ OU = $ouVal; Code = $raw; Label = $text }
    }
    return $global:SedesConfig[0]
}

# Departamentos corporativos en Aliar S.A.
$global:DepartamentosAliar = @(
    "Planta Engativá",
    "Planta Frigorífico",
    "Agronomía",
    "Tecnología de Información",
    "Gestión Humana",
    "Contabilidad y Finanzas",
    "Planta Alimentos",
    "Logística y Cadena de Suministro",
    "Calidad e Inocuidad",
    "Mantenimiento Industrial",
    "Comercial y Ventas",
    "Comercio Exterior",
    "Compras y Abastecimiento",
    "Seguridad y Salud en el Trabajo",
    "Auditoría y Control Interno",
    "Gerencia General"
)

$global:PerfilesWeb = @(
    "GR_Perfil_Muy_Restrictivo",
    "GR_Perfil_Moderadamente_Restrictivo",
    "GR_Perfil_Equilibrado",
    "GR_Perfil_Relativamente_Libre"
)

# ======================================================================================
# 2. FUNCIONES DE LIMPIEZA, PARSER Y AUTO-DETECCION POR CARGO
# ======================================================================================
function Clean-Text($str) {
    if (-not $str) { return "" }
    $norm = $str.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $norm.Length; $i++) {
        $c = $norm[$i]
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().ToLower() -replace "[^a-z0-9]", ""
}

function To-TitleCase($str) {
    if (-not $str) { return "" }
    $clean = $str.ToString().Trim().ToLower()
    $textInfo = (Get-Culture).TextInfo
    return $textInfo.ToTitleCase($clean)
}

function Generate-RandomPassword {
    return "Aliar2026*"
}







function Parse-AliarWorkerString($str) {
    if (-not $str) { return $null }
    $str = $str.Trim()
    if ($str -match '^(\d{5,12})[\s;,|\t]+(.+)$') {
        $cedula = $matches[1]
        $resto = $matches[2].Trim()
        
        $cargoKeywords = 'OPERARIO|AUXILIAR|ANALISTA|ASISTENTE|SUPERVISOR|COORDINADOR|JEFE|DIRECTOR|GERENTE|LIDER|TECNICO|TECNOLOGO|INGENIERO|CONDUCTOR|VETERINARIO|MEDICO|APRENDIZ|PRACTICANTE|OFICIAL|MECANICO|ELECTRICO|ADMINISTRADOR|ASESOR|AUDITOR|MERCADERISTA|VENDEDOR|CAJERO|DESPOSTADOR|CHOFER'
        if ($resto -match '^(.*?)\s+((?:' + $cargoKeywords + ')\b.*)$') {
            $nombresCompletos = $matches[1].Trim()
            $cargo = $matches[2].Trim()
        } else {
            $parts = $resto -split '\s+'
            if ($parts.Count -ge 5) {
                $nombresCompletos = ($parts[0..3] -join ' ')
                $cargo = ($parts[4..($parts.Count - 1)] -join ' ')
            } else {
                $nombresCompletos = $resto
                $cargo = "OPERARIO PLANTA"
            }
        }
        
        $nParts = $nombresCompletos -split '\s+'
        $first = ''; $middle = ''; $last = ''; $secondLast = ''
        if ($nParts.Count -ge 4) {
            $first = $nParts[0]; $middle = $nParts[1]; $last = $nParts[2]; $secondLast = ($nParts[3..($nParts.Count - 1)] -join ' ')
        } elseif ($nParts.Count -eq 3) {
            $first = $nParts[0]; $last = $nParts[1]; $secondLast = $nParts[2]
        } elseif ($nParts.Count -eq 2) {
            $first = $nParts[0]; $last = $nParts[1]
        } else {
            $first = $nombresCompletos
        }
        
        return [PSCustomObject]@{
            Cedula = $cedula
            PrimerNombre = To-TitleCase $first
            SegundoNombre = To-TitleCase $middle
            PrimerApellido = To-TitleCase $last
            SegundoApellido = To-TitleCase $secondLast
            Cargo = To-TitleCase $cargo
        }
    }
    return $null
}

# Auto-deteccion inteligente de Sede, Departamento y Perfil segun el Cargo
function Get-RoleInferredSettings($cargo) {
    $c = if ($cargo) { $cargo.ToUpper() } else { "" }
    
    $sede = "CUN_PLENG"
    $dept = "Planta Engativá"
    $profile = "GR_Perfil_Muy_Restrictivo"
    
    # 1. Detectar Departamento y Perfil por especialidad
    if ($c -match '\bGESTION HUMANA\b|\bTALENTO\b|\bGH\b|\bSELECCION\b|\bBIENESTAR\b|\bNOMINA\b') {
        $dept = "Gestión Humana"
        $sede = "SAN_OFPP"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bCONTAB|\bFINANZ|\bTESORER|\bCARTERA|\bCOSTOS|\bIMPUESTOS|\bPAGADOR') {
        $dept = "Contabilidad y Finanzas"
        $sede = "SAN_OFPP"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bBODEGA\b|\bLOGIST|\bALMACEN\b|\bDESPACHO|\bCONDUCTOR\b|\bCHOFER\b|\bENTREGA\b|\bDISTRIB') {
        $dept = "Logística y Cadena de Suministro"
        $sede = "CUN_DISBOG"
        $profile = "GR_Perfil_Muy_Restrictivo"
    } elseif ($c -match '\bCALIDAD\b|\bINOCUIDAD\b|\bLABORATORIO\b|\bMICROBIOL') {
        $dept = "Calidad e Inocuidad"
        $sede = "CUN_PLENG"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bMANTENIMIENTO\b|\bELECTRICO\b|\bMECANICO\b|\bSOLDADOR\b') {
        $dept = "Mantenimiento Industrial"
        $sede = "CUN_PLENG"
        $profile = "GR_Perfil_Muy_Restrictivo"
    } elseif ($c -match '\bVETERINAR|\bZOOTECN|\bGRANJA\b|\bPORCIC|\bCERDOS\b|\bGALPON') {
        $dept = "Agronomía"
        $sede = "MET_PORCI"
        $profile = "GR_Perfil_Muy_Restrictivo"
    } elseif ($c -match '\bAGRONOM|\bCAMPO\b|\bCULTIVO\b|\bPALMA\b|\bCOSECHA\b') {
        $dept = "Agronomía"
        $sede = "MET_AGRI"
        $profile = "GR_Perfil_Muy_Restrictivo"
    } elseif ($c -match '\bALIMENTOS\b|\bCONCENTRADOS\b|\bMOLINO\b') {
        $dept = "Planta Alimentos"
        $sede = "MET_PLAGR"
        $profile = "GR_Perfil_Muy_Restrictivo"
    } elseif ($c -match '\bTI\b|\bSISTEMAS\b|\bTECNOLOG|\bDESARROLLADOR\b|\bSOPORTE\b|\bREDES\b') {
        $dept = "Tecnología de Información"
        $sede = "SAN_OFPP"
        $profile = "GR_Perfil_Relativamente_Libre"
    } elseif ($c -match '\bVENTAS\b|\bCOMERCIAL\b|\bASESOR\b|\bPUNTO DE VENTA\b|\bTIENDA\b') {
        $dept = "Comercial y Ventas"
        $sede = "CUN_PV"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bCOMPRAS\b|\bABASTECIM|\bPROVEEDOR') {
        $dept = "Compras y Abastecimiento"
        $sede = "SAN_OFPP"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bSST\b|\bSEGURIDAD Y SALUD\b|\bAMBIENTAL\b|\bHSEQ\b') {
        $dept = "Seguridad y Salud en el Trabajo"
        $sede = "CUN_PLENG"
        $profile = "GR_Perfil_Moderadamente_Restrictivo"
    } elseif ($c -match '\bOPERARIO\b|\bDESPOSTE\b|\bFAENADO\b|\bEMPAQUE\b|\bPLANTA\b|\bCORTE\b') {
        $dept = "Planta Engativá"
        $sede = "CUN_PLENG"
        $profile = "GR_Perfil_Muy_Restrictivo"
    }
    
    # 2. Refinar Sede por pistas geográficas
    if ($c -match '\bFRIGO|\bPUERTO GAITAN\b|\bMETA\b') {
        $sede = "MET_FRIG"
        if ($dept -eq "Planta Engativá") { $dept = "Planta Frigorífico" }
    } elseif ($c -match '\bENGATIVA\b|\bBOGOTA\b') {
        if ($sede -ne "CUN_PV" -and $sede -ne "CUN_DISBOG") { $sede = "CUN_PLENG" }
    } elseif ($c -match '\bBUCARAMANGA\b|\bPIEDECUESTA\b|\bSANTANDER\b') {
        $sede = "SAN_OFPP"
    } elseif ($c -match '\bMEDELLIN\b|\bANTIOQUIA\b') {
        $sede = "ANT_DISMED"
    }
    
    $vpn = $false
    if ($c -match '\bDIRECTOR|\bGERENTE|\bLIDER|\bJEFE|\bCOORDINADOR|\bTI\b|\bSISTEMAS\b|\bDESARROLLADOR|\bINGENIERO|\bADMINISTRATIVO|\bANALISTA|\bESPECIALISTA') {
        $vpn = $true
    }
    
    return [PSCustomObject]@{
        Sede = $sede
        Departamento = $dept
        PerfilWeb = $profile
        Vpn = $vpn
    }
}

# Configurador de ComboBox con altura homogénea de 28px
function Setup-StyledComboBox($cmb, $items, $itemHeight = 22) {
    $cmb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmb.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $cmb.ItemHeight = $itemHeight
    $cmb.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cmb.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $items | ForEach-Object { [void]$cmb.Items.Add($_) }
    if ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
    
    $cmb.add_DrawItem({
        param($s, $e)
        if ($e.Index -lt 0) { return }
        $e.DrawBackground()
        $text = $s.Items[$e.Index].ToString()
        $brush = if (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected) {
            [System.Drawing.SystemBrushes]::HighlightText
        } else {
            [System.Drawing.SystemBrushes]::WindowText
        }
        $sf = New-Object System.Drawing.StringFormat
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $e.Graphics.DrawString($text, $s.Font, $brush, [float]($e.Bounds.X + 3), [float]($e.Bounds.Y + ($e.Bounds.Height / 2)), $sf)
        $e.DrawFocusRectangle()
    })
}

# ======================================================================================
# 3. CONSTRUCCION DE LA VENTANA PRINCIPAL
# ======================================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Aliar S.A. | Active Directory Provisioning Hub [v8.2 Corporativo Actualizado]"
$form.ClientSize = New-Object System.Drawing.Size(1360, 890)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 720)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249) # Slate 100
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --------------------------------------------------------------------------------------
# HEADER SUPERIOR CORPORATIVO
# --------------------------------------------------------------------------------------
# HEADER SUPERIOR CORPORATIVO
# --------------------------------------------------------------------------------------
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 52
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42) # Slate 900
$form.Controls.Add($pnlHeader)

$logoPath = Join-Path $scriptDir "Logo-web-1.png"
if (Test-Path $logoPath) {
    try {
        $imgLogo = [System.Drawing.Image]::FromFile($logoPath)
        $picLogo = New-Object System.Windows.Forms.PictureBox
        $picLogo.Image = $imgLogo
        $picLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $picLogo.Size = New-Object System.Drawing.Size(115, 38)
        $picLogo.Location = New-Object System.Drawing.Point(16, 7)
        $picLogo.BackColor = [System.Drawing.Color]::Transparent
        $pnlHeader.Controls.Add($picLogo)
        
        try {
            $hIcon = ([System.Drawing.Bitmap]$imgLogo).GetHicon()
            $form.Icon = [System.Drawing.Icon]::FromHandle($hIcon)
        } catch {}
    } catch {}
}

$lblTitle = New-Object System.Windows.Forms.Label
if (Test-Path $logoPath) {
    $lblTitle.Text = "|  Gestión de Usuarios Active Directory"
    $lblTitle.Location = New-Object System.Drawing.Point(138, 14)
} else {
    $lblTitle.Text = "Aliar S.A.  |  Gestión de Usuarios Active Directory"
    $lblTitle.Location = New-Object System.Drawing.Point(18, 14)
}
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12.0, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$pnlHeader.Controls.Add($lblTitle)

$pnlHeaderRight = New-Object System.Windows.Forms.Panel
$pnlHeaderRight.Dock = [System.Windows.Forms.DockStyle]::Right
$pnlHeaderRight.Width = 540
$pnlHeader.Controls.Add($pnlHeaderRight)

$lblStats = New-Object System.Windows.Forms.Label
$lblStats.Text = "Base Local: $($global:RealGroups.Count) Grupos | $($global:RealUsers.Count) Colaboradores"
$lblStats.ForeColor = [System.Drawing.Color]::FromArgb(52, 211, 153) # Emerald 400
$lblStats.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblStats.Location = New-Object System.Drawing.Point(10, 18)
$lblStats.AutoSize = $true
$pnlHeaderRight.Controls.Add($lblStats)

$btnRefreshAD = New-Object System.Windows.Forms.Button
$btnRefreshAD.Text = "Refrescar DA en Vivo"
$btnRefreshAD.Size = New-Object System.Drawing.Size(155, 28)
$btnRefreshAD.Location = New-Object System.Drawing.Point(370, 12)
$btnRefreshAD.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$btnRefreshAD.ForeColor = [System.Drawing.Color]::White
$btnRefreshAD.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefreshAD.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRefreshAD.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlHeaderRight.Controls.Add($btnRefreshAD)

# --------------------------------------------------------------------------------------
# CONTENEDOR SPLIT PRINCIPAL (IZQUIERDA: PESTANAS / DERECHA: ACCION Y FICHA)
# --------------------------------------------------------------------------------------
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitMain.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel2
$splitMain.Size = New-Object System.Drawing.Size(1360, 825)
$splitMain.Panel1MinSize = 540
$splitMain.Panel2MinSize = 420
$splitMain.SplitterDistance = 910
$splitMain.BackColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
$form.Controls.Add($splitMain)
$splitMain.BringToFront()

# ======================================================================================
# PANEL IZQUIERDO: PESTANAS DE GESTION
# ======================================================================================
$tabControlLeft = New-Object System.Windows.Forms.TabControl
$tabControlLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControlLeft.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$splitMain.Panel1.Controls.Add($tabControlLeft)

$tab1 = New-Object System.Windows.Forms.TabPage
$tab1.Text = " 1. Crear Usuario Individual "
$tab1.BackColor = [System.Drawing.Color]::White
$tabControlLeft.TabPages.Add($tab1)

$tab2 = New-Object System.Windows.Forms.TabPage
$tab2.Text = " 2. Gestionar Permisos (Existente) "
$tab2.BackColor = [System.Drawing.Color]::White
$tabControlLeft.TabPages.Add($tab2)

$tab3 = New-Object System.Windows.Forms.TabPage
$tab3.Text = " 3. Carga Masiva (Lotes Multi-Cargo / GH) "
$tab3.BackColor = [System.Drawing.Color]::White
$tabControlLeft.TabPages.Add($tab3)

# --------------------------------------------------------------------------------------
# TAB 1: CREAR USUARIO INDIVIDUAL
# --------------------------------------------------------------------------------------
$pnlTab1Scroll = New-Object System.Windows.Forms.Panel
$pnlTab1Scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab1Scroll.AutoScroll = $true
$pnlTab1Scroll.Padding = New-Object System.Windows.Forms.Padding(10)
$tab1.Controls.Add($pnlTab1Scroll)

$tblTab1 = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab1.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab1.MinimumSize = New-Object System.Drawing.Size(650, 760)
$tblTab1.ColumnCount = 1
$tblTab1.RowCount = 4
[void]$tblTab1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 85)))  # Fila 0: Auto-completar
[void]$tblTab1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 235))) # Fila 1: Datos Basicos
[void]$tblTab1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180))) # Fila 2: Usuario Modelo
[void]$tblTab1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # Fila 3: Perfiles y Carpetas
$pnlTab1Scroll.Controls.Add($tblTab1)

# Tab 1 - Fila 0: Barra de Auto-Completar
$grpSmartPaste = New-Object System.Windows.Forms.GroupBox
$grpSmartPaste.Text = " Pegar Datos Rápidos de Gestión Humana (Autocompletar en 1 Clic) "
$grpSmartPaste.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpSmartPaste.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199) # Sky 600
$grpSmartPaste.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpSmartPaste.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 8)
$tblTab1.Controls.Add($grpSmartPaste, 0, 0)

$tblSmartInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblSmartInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblSmartInner.ColumnCount = 3
$tblSmartInner.RowCount = 2
[void]$tblSmartInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblSmartInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 195)))
[void]$tblSmartInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95)))
[void]$tblSmartInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
[void]$tblSmartInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$grpSmartPaste.Controls.Add($tblSmartInner)

$lblSmartHelp = New-Object System.Windows.Forms.Label
$lblSmartHelp.Text = "Pega aquí el texto de Gestión Humana y presiona Enter (ej: 12345678 PEPITO PEREZ RODRIGUEZ OPERARIO PLANTA):"
$lblSmartHelp.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblSmartHelp.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblSmartHelp.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblSmartInner.Controls.Add($lblSmartHelp, 0, 0)
$tblSmartInner.SetColumnSpan($lblSmartHelp, 3)

$txtSmartPaste = New-Object System.Windows.Forms.TextBox
$txtSmartPaste.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSmartPaste.Font = New-Object System.Drawing.Font("Segoe UI", 11.5)
$txtSmartPaste.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tblSmartInner.Controls.Add($txtSmartPaste, 0, 1)

$btnSmartParse = New-Object System.Windows.Forms.Button
$btnSmartParse.Text = "Auto-Rellenar Formulario"
$btnSmartParse.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSmartParse.Height = 28
$btnSmartParse.MaximumSize = New-Object System.Drawing.Size(0, 28)
$btnSmartParse.BackColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$btnSmartParse.ForeColor = [System.Drawing.Color]::White
$btnSmartParse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSmartParse.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSmartParse.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$tblSmartInner.Controls.Add($btnSmartParse, 1, 1)

$btnSmartClear = New-Object System.Windows.Forms.Button
$btnSmartClear.Text = "Limpiar"
$btnSmartClear.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSmartClear.Height = 28
$btnSmartClear.MaximumSize = New-Object System.Drawing.Size(0, 28)
$btnSmartClear.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$btnSmartClear.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$btnSmartClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSmartClear.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSmartClear.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$tblSmartInner.Controls.Add($btnSmartClear, 2, 1)

# Tab 1 - Fila 1: Datos Básicos del Colaborador
$grpDatos = New-Object System.Windows.Forms.GroupBox
$grpDatos.Text = " Información del Colaborador (Aliar S.A.) "
$grpDatos.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpDatos.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$grpDatos.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpDatos.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 8)
$tblTab1.Controls.Add($grpDatos, 0, 1)

$tblDatos = New-Object System.Windows.Forms.TableLayoutPanel
$tblDatos.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblDatos.ColumnCount = 4
$tblDatos.RowCount = 3
[void]$tblDatos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$tblDatos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$tblDatos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$tblDatos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$tblDatos.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54)))
[void]$tblDatos.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54)))
[void]$tblDatos.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$grpDatos.Controls.Add($tblDatos)

function Make-InputCell($labelText, [ref]$txtRef) {
    $pnl = New-Object System.Windows.Forms.TableLayoutPanel
    $pnl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnl.Margin = New-Object System.Windows.Forms.Padding(4, 1, 4, 1)
    $pnl.ColumnCount = 1
    $pnl.RowCount = 2
    [void]$pnl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
    [void]$pnl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $pnl.Controls.Add($lbl, 0, 0)

    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock = [System.Windows.Forms.DockStyle]::Fill
    $t.Font = New-Object System.Drawing.Font("Segoe UI", 11.5)
    $t.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnl.Controls.Add($t, 0, 1)

    $txtRef.Value = $t
    return $pnl
}

$txtFirstName = $null
$txtMiddleName = $null
$txtLastName = $null
$txtSecondLastName = $null

$cellFN = Make-InputCell "Primer Nombre *" ([ref]$txtFirstName)
$cellMN = Make-InputCell "Segundo Nombre" ([ref]$txtMiddleName)
$cellLN = Make-InputCell "Primer Apellido *" ([ref]$txtLastName)
$cellSN = Make-InputCell "Segundo Apellido" ([ref]$txtSecondLastName)

$tblDatos.Controls.Add($cellFN, 0, 0)
$tblDatos.Controls.Add($cellMN, 1, 0)
$tblDatos.Controls.Add($cellLN, 2, 0)
$tblDatos.Controls.Add($cellSN, 3, 0)

# Fila 1: Cédula, Cargo, Sede, Departamento
$txtEmployeeId = $null
$txtTitle = $null
$cellId = Make-InputCell "Cédula / Documento *" ([ref]$txtEmployeeId)
$cellTitle = Make-InputCell "Cargo en Aliar *" ([ref]$txtTitle)

# Sede Dropdown Tab 1
$pnlSede = New-Object System.Windows.Forms.TableLayoutPanel
$pnlSede.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlSede.Margin = New-Object System.Windows.Forms.Padding(4, 1, 4, 1)
$pnlSede.ColumnCount = 1
$pnlSede.RowCount = 2
[void]$pnlSede.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
[void]$pnlSede.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$lblSede = New-Object System.Windows.Forms.Label
$lblSede.Text = "Sede / Planta de Operación *"
$lblSede.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblSede.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblSede.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$pnlSede.Controls.Add($lblSede, 0, 0)
$cmbSede = New-Object System.Windows.Forms.ComboBox
Setup-StyledComboBox $cmbSede $global:SedesAliar 22
$cmbSede.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$cmbSede.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$cmbSede.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$cmbSede.add_TextChanged({ Update-AllOutputs })
$pnlSede.Controls.Add($cmbSede, 0, 1)

# Depto Dropdown Tab 1
$pnlDept = New-Object System.Windows.Forms.TableLayoutPanel
$pnlDept.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlDept.Margin = New-Object System.Windows.Forms.Padding(4, 1, 4, 1)
$pnlDept.ColumnCount = 1
$pnlDept.RowCount = 2
[void]$pnlDept.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
[void]$pnlDept.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$lblDept = New-Object System.Windows.Forms.Label
$lblDept.Text = "Departamento *"
$lblDept.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblDept.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblDept.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$pnlDept.Controls.Add($lblDept, 0, 0)
$cmbDepartment = New-Object System.Windows.Forms.ComboBox
Setup-StyledComboBox $cmbDepartment $global:DepartamentosAliar 22
$pnlDept.Controls.Add($cmbDepartment, 0, 1)

$tblDatos.Controls.Add($cellId, 0, 1)
$tblDatos.Controls.Add($cellTitle, 1, 1)
$tblDatos.Controls.Add($pnlSede, 2, 1)
$tblDatos.Controls.Add($pnlDept, 3, 1)

# Fila 2: Usuario sAMAccountName y Contraseña
$pnlSam = New-Object System.Windows.Forms.TableLayoutPanel
$pnlSam.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlSam.Margin = New-Object System.Windows.Forms.Padding(4, 1, 4, 1)
$pnlSam.ColumnCount = 1
$pnlSam.RowCount = 2
[void]$pnlSam.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
[void]$pnlSam.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$lblSam = New-Object System.Windows.Forms.Label
$lblSam.Text = "Usuario de Dominio (nombre.apellido) *"
$lblSam.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblSam.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblSam.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$pnlSam.Controls.Add($lblSam, 0, 0)
$txtSam = New-Object System.Windows.Forms.TextBox
$txtSam.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSam.Font = New-Object System.Drawing.Font("Segoe UI", 11.5, [System.Drawing.FontStyle]::Bold)
$txtSam.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtSam.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$pnlSam.Controls.Add($txtSam, 0, 1)

$pnlPass = New-Object System.Windows.Forms.TableLayoutPanel
$pnlPass.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlPass.Margin = New-Object System.Windows.Forms.Padding(4, 1, 4, 1)
$pnlPass.ColumnCount = 2
$pnlPass.RowCount = 2
[void]$pnlPass.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$pnlPass.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
[void]$pnlPass.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
[void]$pnlPass.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Contraseña Temporal *"
$lblPass.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblPass.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblPass.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$pnlPass.Controls.Add($lblPass, 0, 0)
$pnlPass.SetColumnSpan($lblPass, 2)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtPass.Font = New-Object System.Drawing.Font("Segoe UI", 11.5)
$txtPass.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtPass.Text = "Aliar2026*"
$pnlPass.Controls.Add($txtPass, 0, 1)

$btnGenPass = New-Object System.Windows.Forms.Button
$btnGenPass.Text = "Generar Clave"
$btnGenPass.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnGenPass.Height = 28
$btnGenPass.MaximumSize = New-Object System.Drawing.Size(0, 28)
$btnGenPass.MinimumSize = New-Object System.Drawing.Size(0, 28)
$btnGenPass.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenPass.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnGenPass.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPass.Controls.Add($btnGenPass, 1, 1)

$tblDatos.Controls.Add($pnlSam, 0, 2)
$tblDatos.SetColumnSpan($pnlSam, 2)
$tblDatos.Controls.Add($pnlPass, 2, 2)
$tblDatos.SetColumnSpan($pnlPass, 2)

# Fila 3 eliminada (OU Destino no requerida para el usuario)

# Tab 1 - Fila 2: Clonar Usuario Modelo
$grpModelo = New-Object System.Windows.Forms.GroupBox
$grpModelo.Text = " Clonar Permisos y Sede de un Usuario Modelo (Previsualizacion Segura) "
$grpModelo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpModelo.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$grpModelo.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpModelo.Padding = New-Object System.Windows.Forms.Padding(10)
$tblTab1.Controls.Add($grpModelo, 0, 2)

$tblModelo = New-Object System.Windows.Forms.TableLayoutPanel
$tblModelo.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblModelo.ColumnCount = 2
$tblModelo.RowCount = 1
[void]$tblModelo.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$tblModelo.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$grpModelo.Controls.Add($tblModelo)

$pnlModLeft = New-Object System.Windows.Forms.Panel
$pnlModLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlModLeft.Padding = New-Object System.Windows.Forms.Padding(4)
$tblModelo.Controls.Add($pnlModLeft, 0, 0)

$lblSearchModelo = New-Object System.Windows.Forms.Label
$lblSearchModelo.Text = "Buscar colaborador activo en el DA (escribe nombre, apellido o usuario):"
$lblSearchModelo.Dock = [System.Windows.Forms.DockStyle]::Top
$lblSearchModelo.Height = 20
$lblSearchModelo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlModLeft.Controls.Add($lblSearchModelo)

$txtSearchModelo = New-Object System.Windows.Forms.TextBox
$txtSearchModelo.Dock = [System.Windows.Forms.DockStyle]::Top
$txtSearchModelo.Font = New-Object System.Drawing.Font("Segoe UI", 11.5)
$txtSearchModelo.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlModLeft.Controls.Add($txtSearchModelo)

$lstSearchResults = New-Object System.Windows.Forms.ListBox
$lstSearchResults.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstSearchResults.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlModLeft.Controls.Add($lstSearchResults)
$lstSearchResults.BringToFront()

$pnlModRight = New-Object System.Windows.Forms.Panel
$pnlModRight.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlModRight.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$pnlModRight.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlModRight.Padding = New-Object System.Windows.Forms.Padding(8)
$tblModelo.Controls.Add($pnlModRight, 1, 0)

$lblModeloNombre = New-Object System.Windows.Forms.Label
$lblModeloNombre.Text = "Selecciona un usuario modelo de la lista izquierda..."
$lblModeloNombre.Dock = [System.Windows.Forms.DockStyle]::Top
$lblModeloNombre.Height = 22
$lblModeloNombre.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlModRight.Controls.Add($lblModeloNombre)

$lblModeloCount = New-Object System.Windows.Forms.Label
$lblModeloCount.Text = "Grupos en DA: 0"
$lblModeloCount.Dock = [System.Windows.Forms.DockStyle]::Top
$lblModeloCount.Height = 18
$lblModeloCount.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
$lblModeloCount.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlModRight.Controls.Add($lblModeloCount)

$btnClonarModelo = New-Object System.Windows.Forms.Button
$btnClonarModelo.Text = "Cargar Permisos y Sede al Formulario"
$btnClonarModelo.Dock = [System.Windows.Forms.DockStyle]::Bottom
$btnClonarModelo.Height = 28
$btnClonarModelo.BackColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
$btnClonarModelo.ForeColor = [System.Drawing.Color]::White
$btnClonarModelo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClonarModelo.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClonarModelo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$btnClonarModelo.Enabled = $false
$pnlModRight.Controls.Add($btnClonarModelo)

$lstModeloGrupos = New-Object System.Windows.Forms.ListBox
$lstModeloGrupos.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstModeloGrupos.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$pnlModRight.Controls.Add($lstModeloGrupos)
$lstModeloGrupos.BringToFront()

# Tab 1 - Fila 3: Perfiles y Carpetas Departamentales (Expandible Verticalmente)
$grpPermisos = New-Object System.Windows.Forms.GroupBox
$grpPermisos.Text = " Asignación de Perfil de Navegación y Carpetas Departamentales "
$grpPermisos.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpPermisos.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$grpPermisos.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpPermisos.Padding = New-Object System.Windows.Forms.Padding(10)
$tblTab1.Controls.Add($grpPermisos, 0, 3)

$tblPermisosInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblPermisosInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblPermisosInner.ColumnCount = 1
$tblPermisosInner.RowCount = 3
[void]$tblPermisosInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$tblPermisosInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100))) # Radios perfil y VPN
    [void]$tblPermisosInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))  # Filtro carpetas
[void]$tblPermisosInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # CheckedListBox
$grpPermisos.Controls.Add($tblPermisosInner)

# Perfil Web Top Tab 1
$pnlPerfilesTop = New-Object System.Windows.Forms.Panel
$pnlPerfilesTop.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblPermisosInner.Controls.Add($pnlPerfilesTop, 0, 0)

$lblPerfiles = New-Object System.Windows.Forms.Label
$lblPerfiles.Text = "Perfil de Navegación Web (Excluyente):"
$lblPerfiles.Location = New-Object System.Drawing.Point(4, 4)
$lblPerfiles.AutoSize = $true
$lblPerfiles.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPerfilesTop.Controls.Add($lblPerfiles)

$rbPerfil1 = New-Object System.Windows.Forms.RadioButton
$rbPerfil1.Text = "GR_Perfil_Moderadamente_Restrictivo (Estándar Aliar)"
$rbPerfil1.Location = New-Object System.Drawing.Point(4, 24)
$rbPerfil1.AutoSize = $true
$rbPerfil1.Checked = $true
$rbPerfil1.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPerfilesTop.Controls.Add($rbPerfil1)

$rbPerfil2 = New-Object System.Windows.Forms.RadioButton
$rbPerfil2.Text = "GR_Perfil_Muy_Restrictivo (Planta/Operativo)"
$rbPerfil2.Location = New-Object System.Drawing.Point(4, 44)
$rbPerfil2.AutoSize = $true
$rbPerfil2.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPerfilesTop.Controls.Add($rbPerfil2)

$rbPerfil3 = New-Object System.Windows.Forms.RadioButton
$rbPerfil3.Text = "GR_Perfil_Equilibrado"
$rbPerfil3.Location = New-Object System.Drawing.Point(380, 24)
$rbPerfil3.AutoSize = $true
$rbPerfil3.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPerfilesTop.Controls.Add($rbPerfil3)

$rbPerfil4 = New-Object System.Windows.Forms.RadioButton
$rbPerfil4.Text = "GR_Perfil_Relativamente_Libre (Líderes/TI)"
$rbPerfil4.Location = New-Object System.Drawing.Point(380, 44)
$rbPerfil4.AutoSize = $true
$rbPerfil4.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlPerfilesTop.Controls.Add($rbPerfil4)

$chkVpn = New-Object System.Windows.Forms.CheckBox
$chkVpn.Text = "Acceso Remoto Seguro VPN (GR_VPN_SSL)"
    $chkVpn.Location = New-Object System.Drawing.Point(4, 70)
$chkVpn.AutoSize = $true
$chkVpn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$chkVpn.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$pnlPerfilesTop.Controls.Add($chkVpn)

# Filtro Carpetas Tab 1
$pnlCarpetasHeader = New-Object System.Windows.Forms.TableLayoutPanel
$pnlCarpetasHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlCarpetasHeader.ColumnCount = 2
$pnlCarpetasHeader.RowCount = 2
[void]$pnlCarpetasHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$pnlCarpetasHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
    [void]$pnlCarpetasHeader.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
    [void]$pnlCarpetasHeader.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tblPermisosInner.Controls.Add($pnlCarpetasHeader, 0, 1)

$lblCarpetas = New-Object System.Windows.Forms.Label
$lblCarpetas.Text = "Carpetas Departamentales Compartidas ($($global:AreaFolderGroups.Count) grupos GR_DC_* disponibles):"
$lblCarpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblCarpetas.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlCarpetasHeader.Controls.Add($lblCarpetas, 0, 0)
$pnlCarpetasHeader.SetColumnSpan($lblCarpetas, 2)

$txtFilterCarpetas = New-Object System.Windows.Forms.TextBox
$txtFilterCarpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtFilterCarpetas.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtFilterCarpetas.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlCarpetasHeader.Controls.Add($txtFilterCarpetas, 0, 1)

$btnLimpiarFiltro = New-Object System.Windows.Forms.Button
$btnLimpiarFiltro.Text = "Limpiar Filtro"
$btnLimpiarFiltro.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnLimpiarFiltro.Height = 28
$btnLimpiarFiltro.MaximumSize = New-Object System.Drawing.Size(0, 28)
$btnLimpiarFiltro.MinimumSize = New-Object System.Drawing.Size(0, 28)
$btnLimpiarFiltro.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLimpiarFiltro.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLimpiarFiltro.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlCarpetasHeader.Controls.Add($btnLimpiarFiltro, 1, 1)

$chkListCarpetas = New-Object System.Windows.Forms.CheckedListBox
$chkListCarpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
$chkListCarpetas.CheckOnClick = $true
$chkListCarpetas.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tblPermisosInner.Controls.Add($chkListCarpetas, 0, 2)

$global:SelectedGroupNames = New-Object System.Collections.Generic.HashSet[string]

function Refresh-CarpetasList($filter = "") {
    $chkListCarpetas.Items.Clear()
    $filtered = $global:AreaFolderGroups
    if ($filter) {
        $cleanFilter = $filter.ToLower()
        $filtered = $global:AreaFolderGroups | Where-Object { $_.name.ToLower().Contains($cleanFilter) -or $_.desc.ToLower().Contains($cleanFilter) }
    }
    foreach ($g in $filtered) {
        $idx = $chkListCarpetas.Items.Add($g.name)
        if ($global:SelectedGroupNames.Contains($g.name)) {
            $chkListCarpetas.SetItemChecked($idx, $true)
        }
    }
}
Refresh-CarpetasList

# ======================================================================================
# TAB 2: GESTIONAR PERMISOS (COLABORADOR EXISTENTE)
# ======================================================================================
$pnlTab2Scroll = New-Object System.Windows.Forms.Panel
$pnlTab2Scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2Scroll.AutoScroll = $true
$pnlTab2Scroll.AutoScrollMinSize = New-Object System.Drawing.Size(650, 800)
$pnlTab2Scroll.Padding = New-Object System.Windows.Forms.Padding(10)
$tab2.Controls.Add($pnlTab2Scroll)

$tblTab2 = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2.MinimumSize = New-Object System.Drawing.Size(650, 780)
$tblTab2.ColumnCount = 1
$tblTab2.RowCount = 4
[void]$tblTab2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150))) # Fila 0: Búsqueda del colaborador a gestionar (altura para resultados)
[void]$tblTab2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 145))) # Fila 1: Ficha del colaborador y lista de membresías actuales en DA
[void]$tblTab2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 140))) # Fila 2: Clonar permisos de usuario modelo (altura completa)
[void]$tblTab2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # Fila 3: Modificar membresías y carpetas
$pnlTab2Scroll.Controls.Add($tblTab2)

# Tab 2 - Fila 0: Búsqueda de Colaborador a Gestionar
# Tab 2 - Fila 0: Búsqueda de Colaborador a Gestionar
$grpTab2Search = New-Object System.Windows.Forms.GroupBox
$grpTab2Search.Text = " 1. Buscar Colaborador a Gestionar en Active Directory "
$grpTab2Search.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTab2Search.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$grpTab2Search.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpTab2Search.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 8)
$tblTab2.Controls.Add($grpTab2Search, 0, 0)

$tblTab2SearchInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2SearchInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2SearchInner.ColumnCount = 1
$tblTab2SearchInner.RowCount = 3
[void]$tblTab2SearchInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2SearchInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) # Label
[void]$tblTab2SearchInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) # TextBox
[void]$tblTab2SearchInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # ListBox
$grpTab2Search.Controls.Add($tblTab2SearchInner)

$lblTab2Search = New-Object System.Windows.Forms.Label
$lblTab2Search.Text = "Escribe el nombre, apellido o usuario de dominio a consultar:"
$lblTab2Search.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblTab2Search.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblTab2Search.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$tblTab2SearchInner.Controls.Add($lblTab2Search, 0, 0)

$txtTab2Search = New-Object System.Windows.Forms.TextBox
$txtTab2Search.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtTab2Search.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtTab2Search.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tblTab2SearchInner.Controls.Add($txtTab2Search, 0, 1)

$lstTab2SearchResults = New-Object System.Windows.Forms.ListBox
$lstTab2SearchResults.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstTab2SearchResults.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lstTab2SearchResults.IntegralHeight = $false
$tblTab2SearchInner.Controls.Add($lstTab2SearchResults, 0, 2)
$grpTab2Current = New-Object System.Windows.Forms.GroupBox
$grpTab2Current.Text = " Datos del Colaborador y Permisos Actuales en el Dominio "
$grpTab2Current.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTab2Current.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$grpTab2Current.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpTab2Current.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 6)
$tblTab2.Controls.Add($grpTab2Current, 0, 1)

$tblTab2CurrentCols = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2CurrentCols.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentCols.ColumnCount = 2
$tblTab2CurrentCols.RowCount = 1
[void]$tblTab2CurrentCols.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 54)))
[void]$tblTab2CurrentCols.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 46)))
$grpTab2Current.Controls.Add($tblTab2CurrentCols)

# Columna 0: Ficha del Colaborador
$pnlTab2CurrentInfo = New-Object System.Windows.Forms.Panel
$pnlTab2CurrentInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2CurrentInfo.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$pnlTab2CurrentInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlTab2CurrentInfo.Padding = New-Object System.Windows.Forms.Padding(6)
$tblTab2CurrentCols.Controls.Add($pnlTab2CurrentInfo, 0, 0)

$tblTab2CurrentInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2CurrentInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentInner.ColumnCount = 1
$tblTab2CurrentInner.RowCount = 4
[void]$tblTab2CurrentInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2CurrentInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Nombre
[void]$tblTab2CurrentInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) # Cuenta
[void]$tblTab2CurrentInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) # Sede/Dept/Cargo
[void]$tblTab2CurrentInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Perfil Web y VPN
$pnlTab2CurrentInfo.Controls.Add($tblTab2CurrentInner)

$lblTab2ColabName = New-Object System.Windows.Forms.Label
$lblTab2ColabName.Text = "Ningún colaborador seleccionado. Busca y selecciona uno arriba."
$lblTab2ColabName.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblTab2ColabName.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$lblTab2ColabName.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentInner.Controls.Add($lblTab2ColabName, 0, 0)

$lblTab2ColabAccount = New-Object System.Windows.Forms.Label
$lblTab2ColabAccount.Text = "Usuario de Dominio: -"
$lblTab2ColabAccount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblTab2ColabAccount.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$lblTab2ColabAccount.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentInner.Controls.Add($lblTab2ColabAccount, 0, 1)

$lblTab2ColabDetails = New-Object System.Windows.Forms.Label
$lblTab2ColabDetails.Text = "Sede: -  |  Departamento: -  |  Cargo: -"
$lblTab2ColabDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblTab2ColabDetails.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblTab2ColabDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentInner.Controls.Add($lblTab2ColabDetails, 0, 2)

$lblTab2ColabAccess = New-Object System.Windows.Forms.Label
$lblTab2ColabAccess.Text = "Perfil Web: -  |  Acceso VPN: -"
$lblTab2ColabAccess.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblTab2ColabAccess.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
$lblTab2ColabAccess.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2CurrentInner.Controls.Add($lblTab2ColabAccess, 0, 3)

# Columna 1: Permisos y Membresías Actuales en el Dominio
$pnlTab2CurrentPerms = New-Object System.Windows.Forms.Panel
$pnlTab2CurrentPerms.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2CurrentPerms.Padding = New-Object System.Windows.Forms.Padding(0)
$tblTab2CurrentCols.Controls.Add($pnlTab2CurrentPerms, 1, 0)

$tblTab2PermsInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2PermsInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2PermsInner.ColumnCount = 1
$tblTab2PermsInner.RowCount = 2
[void]$tblTab2PermsInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2PermsInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
[void]$tblTab2PermsInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$pnlTab2CurrentPerms.Controls.Add($tblTab2PermsInner)

$lblTab2CurrentPermsTitle = New-Object System.Windows.Forms.Label
$lblTab2CurrentPermsTitle.Text = "Membresías y Carpetas Actuales en DA (0):"
$lblTab2CurrentPermsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblTab2CurrentPermsTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$lblTab2CurrentPermsTitle.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2PermsInner.Controls.Add($lblTab2CurrentPermsTitle, 0, 0)

$lstTab2CurrentPerms = New-Object System.Windows.Forms.ListBox
$lstTab2CurrentPerms.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstTab2CurrentPerms.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lstTab2CurrentPerms.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$lstTab2CurrentPerms.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lstTab2CurrentPerms.Items.Add("(Selecciona un colaborador para listar sus membresías)")
$tblTab2PermsInner.Controls.Add($lstTab2CurrentPerms, 0, 1)
# Tab 2 - Fila 2: Clonar Permisos de un Usuario Modelo
$grpTab2Modelo = New-Object System.Windows.Forms.GroupBox
$grpTab2Modelo.Text = " 2. Clonar Permisos desde un Colaborador Modelo (Opcional) "
$grpTab2Modelo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTab2Modelo.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$grpTab2Modelo.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpTab2Modelo.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 8)
$tblTab2.Controls.Add($grpTab2Modelo, 0, 2)

$tblTab2ModeloInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2ModeloInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2ModeloInner.ColumnCount = 2
$tblTab2ModeloInner.RowCount = 1
[void]$tblTab2ModeloInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 54)))
[void]$tblTab2ModeloInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 46)))
$grpTab2Modelo.Controls.Add($tblTab2ModeloInner)

$pnlTab2ModLeft = New-Object System.Windows.Forms.Panel
$pnlTab2ModLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2ModLeft.Padding = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
$tblTab2ModeloInner.Controls.Add($pnlTab2ModLeft, 0, 0)

$lblTab2ModSearch = New-Object System.Windows.Forms.Label
$lblTab2ModSearch.Text = "Buscar usuario modelo activo en el DA:"
$lblTab2ModSearch.Dock = [System.Windows.Forms.DockStyle]::Top
$lblTab2ModSearch.Height = 18
$lblTab2ModSearch.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblTab2ModSearch.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$pnlTab2ModLeft.Controls.Add($lblTab2ModSearch)

$txtTab2ModSearch = New-Object System.Windows.Forms.TextBox
$txtTab2ModSearch.Dock = [System.Windows.Forms.DockStyle]::Top
$txtTab2ModSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtTab2ModSearch.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlTab2ModLeft.Controls.Add($txtTab2ModSearch)

$lstTab2ModResults = New-Object System.Windows.Forms.ListBox
$lstTab2ModResults.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstTab2ModResults.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2ModLeft.Controls.Add($lstTab2ModResults)
$lstTab2ModResults.BringToFront()

$pnlTab2ModRight = New-Object System.Windows.Forms.Panel
$pnlTab2ModRight.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2ModRight.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$pnlTab2ModRight.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlTab2ModRight.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$tblTab2ModeloInner.Controls.Add($pnlTab2ModRight, 1, 0)

$tblTab2ModRightInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2ModRightInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2ModRightInner.ColumnCount = 1
$tblTab2ModRightInner.RowCount = 3
[void]$tblTab2ModRightInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2ModRightInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Info
[void]$tblTab2ModRightInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) # Count
[void]$tblTab2ModRightInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # Button
$pnlTab2ModRight.Controls.Add($tblTab2ModRightInner)

$lblTab2ModInfo = New-Object System.Windows.Forms.Label
$lblTab2ModInfo.Text = "Selecciona un modelo a la izquierda..."
$lblTab2ModInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblTab2ModInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$tblTab2ModRightInner.Controls.Add($lblTab2ModInfo, 0, 0)

$lblTab2ModCount = New-Object System.Windows.Forms.Label
$lblTab2ModCount.Text = "Grupos en DA: 0"
$lblTab2ModCount.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblTab2ModCount.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
$lblTab2ModCount.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$tblTab2ModRightInner.Controls.Add($lblTab2ModCount, 0, 1)

$btnTab2Clone = New-Object System.Windows.Forms.Button
$btnTab2Clone.Text = "Copiar Permisos del Modelo"
$btnTab2Clone.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnTab2Clone.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$btnTab2Clone.BackColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
$btnTab2Clone.ForeColor = [System.Drawing.Color]::White
$btnTab2Clone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTab2Clone.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnTab2Clone.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$btnTab2Clone.Enabled = $false
$tblTab2ModRightInner.Controls.Add($btnTab2Clone, 0, 2)
# Tab 2 - Fila 3: Modificación de Permisos
$grpTab2Perms = New-Object System.Windows.Forms.GroupBox
$grpTab2Perms.Text = " 3. Modificar Membresías y Carpetas Departamentales "
$grpTab2Perms.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTab2Perms.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$grpTab2Perms.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpTab2Perms.Padding = New-Object System.Windows.Forms.Padding(10)
$tblTab2.Controls.Add($grpTab2Perms, 0, 3)

$tblTab2PermsInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab2PermsInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2PermsInner.ColumnCount = 1
$tblTab2PermsInner.RowCount = 3
[void]$tblTab2PermsInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab2PermsInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100))) # Radios perfil y VPN
[void]$tblTab2PermsInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))  # Filtro carpetas
[void]$tblTab2PermsInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # CheckedListBox
$grpTab2Perms.Controls.Add($tblTab2PermsInner)

# Perfil Web Tab 2
$pnlTab2PerfilesTop = New-Object System.Windows.Forms.Panel
$pnlTab2PerfilesTop.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab2PermsInner.Controls.Add($pnlTab2PerfilesTop, 0, 0)

$lblTab2Perfiles = New-Object System.Windows.Forms.Label
$lblTab2Perfiles.Text = "Perfil de Navegación Web:"
$lblTab2Perfiles.Location = New-Object System.Drawing.Point(4, 4)
$lblTab2Perfiles.AutoSize = $true
$lblTab2Perfiles.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2PerfilesTop.Controls.Add($lblTab2Perfiles)

$rbTab2Perfil1 = New-Object System.Windows.Forms.RadioButton
$rbTab2Perfil1.Text = "GR_Perfil_Moderadamente_Restrictivo"
$rbTab2Perfil1.Location = New-Object System.Drawing.Point(4, 25)
$rbTab2Perfil1.AutoSize = $true
$rbTab2Perfil1.Checked = $true
$rbTab2Perfil1.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2PerfilesTop.Controls.Add($rbTab2Perfil1)

$rbTab2Perfil2 = New-Object System.Windows.Forms.RadioButton
$rbTab2Perfil2.Text = "GR_Perfil_Muy_Restrictivo"
$rbTab2Perfil2.Location = New-Object System.Drawing.Point(4, 47)
$rbTab2Perfil2.AutoSize = $true
$rbTab2Perfil2.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2PerfilesTop.Controls.Add($rbTab2Perfil2)

$rbTab2Perfil3 = New-Object System.Windows.Forms.RadioButton
$rbTab2Perfil3.Text = "GR_Perfil_Equilibrado"
$rbTab2Perfil3.Location = New-Object System.Drawing.Point(340, 25)
$rbTab2Perfil3.AutoSize = $true
$rbTab2Perfil3.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2PerfilesTop.Controls.Add($rbTab2Perfil3)

$rbTab2Perfil4 = New-Object System.Windows.Forms.RadioButton
$rbTab2Perfil4.Text = "GR_Perfil_Relativamente_Libre"
$rbTab2Perfil4.Location = New-Object System.Drawing.Point(340, 47)
$rbTab2Perfil4.AutoSize = $true
$rbTab2Perfil4.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2PerfilesTop.Controls.Add($rbTab2Perfil4)

$chkTab2Vpn = New-Object System.Windows.Forms.CheckBox
$chkTab2Vpn.Text = "Acceso Remoto Seguro VPN (GR_VPN_SSL)"
$chkTab2Vpn.Location = New-Object System.Drawing.Point(4, 71)
$chkTab2Vpn.AutoSize = $true
$chkTab2Vpn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$chkTab2Vpn.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$pnlTab2PerfilesTop.Controls.Add($chkTab2Vpn)

# Filtro Carpetas Tab 2
$pnlTab2CarpetasHeader = New-Object System.Windows.Forms.TableLayoutPanel
$pnlTab2CarpetasHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab2CarpetasHeader.ColumnCount = 2
$pnlTab2CarpetasHeader.RowCount = 2
[void]$pnlTab2CarpetasHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$pnlTab2CarpetasHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 115)))
[void]$pnlTab2CarpetasHeader.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$pnlTab2CarpetasHeader.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tblTab2PermsInner.Controls.Add($pnlTab2CarpetasHeader, 0, 1)

$lblTab2Carpetas = New-Object System.Windows.Forms.Label
$lblTab2Carpetas.Text = "Selecciona o desmarca carpetas departamentales para actualizar permisos:"
$lblTab2Carpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblTab2Carpetas.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2CarpetasHeader.Controls.Add($lblTab2Carpetas, 0, 0)
$pnlTab2CarpetasHeader.SetColumnSpan($lblTab2Carpetas, 2)

$txtTab2FilterCarpetas = New-Object System.Windows.Forms.TextBox
$txtTab2FilterCarpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtTab2FilterCarpetas.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtTab2FilterCarpetas.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtTab2FilterCarpetas.Margin = New-Object System.Windows.Forms.Padding(0, 1, 6, 1)
$pnlTab2CarpetasHeader.Controls.Add($txtTab2FilterCarpetas, 0, 1)

$btnTab2LimpiarFiltro = New-Object System.Windows.Forms.Button
$btnTab2LimpiarFiltro.Text = "Limpiar Filtro"
$btnTab2LimpiarFiltro.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnTab2LimpiarFiltro.Height = 28
$btnTab2LimpiarFiltro.Margin = New-Object System.Windows.Forms.Padding(0, 1, 0, 1)
$btnTab2LimpiarFiltro.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$btnTab2LimpiarFiltro.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTab2LimpiarFiltro.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnTab2LimpiarFiltro.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTab2CarpetasHeader.Controls.Add($btnTab2LimpiarFiltro, 1, 1)

$chkListTab2Carpetas = New-Object System.Windows.Forms.CheckedListBox
$chkListTab2Carpetas.Dock = [System.Windows.Forms.DockStyle]::Fill
$chkListTab2Carpetas.CheckOnClick = $true
$chkListTab2Carpetas.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tblTab2PermsInner.Controls.Add($chkListTab2Carpetas, 0, 2)

$global:Tab2SelectedUser = $null
$global:Tab2OriginalGroups = New-Object System.Collections.Generic.HashSet[string]
$global:Tab2CurrentSelectedGroups = New-Object System.Collections.Generic.HashSet[string]

$global:Tab2Refreshing = $false

function Refresh-Tab2CarpetasList($filter = "") {
    $wasRefreshing = $global:Tab2Refreshing
    $global:Tab2Refreshing = $true
    $chkListTab2Carpetas.BeginUpdate()
    try {
        $chkListTab2Carpetas.Items.Clear()
        $filtered = $global:AreaFolderGroups
        if ($filter) {
            $cleanFilter = $filter.ToLower()
            $filtered = $global:AreaFolderGroups | Where-Object { $_.name.ToLower().Contains($cleanFilter) -or $_.desc.ToLower().Contains($cleanFilter) }
        }
        foreach ($g in $filtered) {
            $idx = $chkListTab2Carpetas.Items.Add($g.name)
            if ($global:Tab2CurrentSelectedGroups.Contains($g.name)) {
                $chkListTab2Carpetas.SetItemChecked($idx, $true)
            }
        }
    } finally {
        $chkListTab2Carpetas.EndUpdate()
        $global:Tab2Refreshing = $wasRefreshing
    }
}
Refresh-Tab2CarpetasList
# TAB 3: CARGA MASIVA MULTI-CARGO / MULTI-SEDE / MULTI-ÁREA
# ======================================================================================
$pnlTab3Scroll = New-Object System.Windows.Forms.Panel
$pnlTab3Scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab3Scroll.Padding = New-Object System.Windows.Forms.Padding(12)
$tab3.Controls.Add($pnlTab3Scroll)

$tblTab3 = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab3.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab3.ColumnCount = 1
$tblTab3.RowCount = 4
[void]$tblTab3.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 165))) # Entrada (ampliado para multilínea y botones completos)
[void]$tblTab3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))  # Barra de herramientas y clasificación
[void]$tblTab3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # DataGridView Interactivo
[void]$tblTab3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # Resumen y exportacion
$pnlTab3Scroll.Controls.Add($tblTab3)

# Tab 3 - Fila 0: Entrada de Datos
# Tab 3 - Fila 0: Entrada de Datos
$grpTab3Input = New-Object System.Windows.Forms.GroupBox
$grpTab3Input.Text = " Paso 1: Pegar Lista de Colaboradores Enviada por Gestión Humana (Soporta Multi-Cargo y Multi-Área) "
$grpTab3Input.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTab3Input.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$grpTab3Input.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpTab3Input.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 8)
$tblTab3.Controls.Add($grpTab3Input, 0, 0)

$tblTab3InputInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab3InputInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab3InputInner.ColumnCount = 2
$tblTab3InputInner.RowCount = 1
[void]$tblTab3InputInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab3InputInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 185)))
$grpTab3Input.Controls.Add($tblTab3InputInner)

$txtBatchInput = New-Object System.Windows.Forms.TextBox
$txtBatchInput.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtBatchInput.Multiline = $true
$txtBatchInput.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtBatchInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$tblTab3InputInner.Controls.Add($txtBatchInput, 0, 0)

$tblTab3SideBtns = New-Object System.Windows.Forms.TableLayoutPanel
$tblTab3SideBtns.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblTab3SideBtns.ColumnCount = 1
$tblTab3SideBtns.RowCount = 2
[void]$tblTab3SideBtns.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblTab3SideBtns.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$tblTab3SideBtns.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tblTab3InputInner.Controls.Add($tblTab3SideBtns, 1, 0)

$btnBatchLoadFile = New-Object System.Windows.Forms.Button
$btnBatchLoadFile.Text = "Cargar CSV / TXT"
$btnBatchLoadFile.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnBatchLoadFile.Margin = New-Object System.Windows.Forms.Padding(4, 2, 0, 4)
$btnBatchLoadFile.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$btnBatchLoadFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBatchLoadFile.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnBatchLoadFile.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$tblTab3SideBtns.Controls.Add($btnBatchLoadFile, 0, 0)

$btnBatchClear = New-Object System.Windows.Forms.Button
$btnBatchClear.Text = "Limpiar Lista"
$btnBatchClear.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnBatchClear.Margin = New-Object System.Windows.Forms.Padding(4, 4, 0, 2)
$btnBatchClear.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$btnBatchClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBatchClear.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnBatchClear.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$tblTab3SideBtns.Controls.Add($btnBatchClear, 0, 1)
$pnlBatchTools = New-Object System.Windows.Forms.Panel
$pnlBatchTools.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlBatchTools.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 6)
$tblTab3.Controls.Add($pnlBatchTools, 0, 1)

$btnProcessBatch = New-Object System.Windows.Forms.Button
$btnProcessBatch.Text = "Procesar y Clasificar Lista de Colaboradores"
$btnProcessBatch.Dock = [System.Windows.Forms.DockStyle]::Left
$btnProcessBatch.Width = 340
$btnProcessBatch.Height = 32
$btnProcessBatch.BackColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$btnProcessBatch.ForeColor = [System.Drawing.Color]::White
$btnProcessBatch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnProcessBatch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnProcessBatch.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlBatchTools.Controls.Add($btnProcessBatch)
# Tab 3 - Fila 2: DataGridView Interactivo y Editable con Dropdowns por Fila
$gridBatch = New-Object System.Windows.Forms.DataGridView
$gridBatch.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridBatch.AllowUserToAddRows = $false
$gridBatch.AllowUserToDeleteRows = $true
$gridBatch.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$gridBatch.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::RowHeaderSelect
$gridBatch.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
$gridBatch.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$gridBatch.BackgroundColor = [System.Drawing.Color]::White
$gridBatch.RowHeadersWidth = 28
Enable-DoubleBuffer $gridBatch
$gridBatch.add_DataError({
    param($s, $e)
    $e.ThrowException = $false
})
$tblTab3.Controls.Add($gridBatch, 0, 2)

# Columnas del DataGridView
$colCed = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCed.Name = "Cedula"; $colCed.HeaderText = "Cédula"; $colCed.Width = 90; [void]$gridBatch.Columns.Add($colCed)
$colNom = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colNom.Name = "Nombre"; $colNom.HeaderText = "Nombre Completo"; $colNom.Width = 200; [void]$gridBatch.Columns.Add($colNom)
$colUsr = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colUsr.Name = "Usuario"; $colUsr.HeaderText = "Usuario de Dominio"; $colUsr.Width = 140; [void]$gridBatch.Columns.Add($colUsr)
$colCrg = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCrg.Name = "Cargo"; $colCrg.HeaderText = "Cargo"; $colCrg.Width = 150; [void]$gridBatch.Columns.Add($colCrg)

# Sede Dropdown Column
$colSede = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colSede.Name = "Sede"
$colSede.HeaderText = "Sede / Ubicación (OU)"
$colSede.Width = 160
$colSede.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::ComboBox
$colSede.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:SedesCodigos | ForEach-Object { [void]$colSede.Items.Add($_) }
$global:SedesOUs | ForEach-Object { if (-not $colSede.Items.Contains($_)) { [void]$colSede.Items.Add($_) } }
[void]$gridBatch.Columns.Add($colSede)

# Departamento Dropdown Column
$colDpt = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colDpt.Name = "Departamento"
$colDpt.HeaderText = "Departamento"
$colDpt.Width = 170
$colDpt.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::ComboBox
$colDpt.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:DepartamentosAliar | ForEach-Object { [void]$colDpt.Items.Add($_) }
[void]$gridBatch.Columns.Add($colDpt)

# Perfil Web Dropdown Column
$colPrf = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colPrf.Name = "PerfilWeb"
$colPrf.HeaderText = "Perfil de Navegacion Web"
$colPrf.Width = 180
$colPrf.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::ComboBox
$colPrf.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:PerfilesWeb | ForEach-Object { [void]$colPrf.Items.Add($_) }
[void]$gridBatch.Columns.Add($colPrf)

# VPN CheckBox Column (Por defecto activado)
$colVpn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colVpn.Name = "Vpn"
$colVpn.HeaderText = "Acceso VPN"
$colVpn.Width = 85
$colVpn.ThreeState = $false
$colVpn.TrueValue = $true
$colVpn.FalseValue = $false
$colVpn.DefaultCellStyle.NullValue = $true
$colVpn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
[void]$gridBatch.Columns.Add($colVpn)

$colClv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClv.Name = "Clave"; $colClv.HeaderText = "Clave Temporal"; $colClv.Width = 100; [void]$gridBatch.Columns.Add($colClv)
$colEst = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colEst.Name = "EstadoDA"; $colEst.HeaderText = "Estado en Active Directory"; $colEst.Width = 170; [void]$gridBatch.Columns.Add($colEst)

# Tab 3 - Fila 3: Barra inferior de Resumen y Exportacion
$pnlTab3Footer = New-Object System.Windows.Forms.Panel
$pnlTab3Footer.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlTab3Footer.Padding = New-Object System.Windows.Forms.Padding(4)
$tblTab3.Controls.Add($pnlTab3Footer, 0, 3)

$lblBatchStatus = New-Object System.Windows.Forms.Label
$lblBatchStatus.Text = "Total en Lote: 0 colaboradores | Disponibles: 0 | Ya registrados en DA: 0"
$lblBatchStatus.Dock = [System.Windows.Forms.DockStyle]::Left
$lblBatchStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblBatchStatus.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$lblBatchStatus.AutoSize = $true
$pnlTab3Footer.Controls.Add($lblBatchStatus)

$btnExportCsvReport = New-Object System.Windows.Forms.Button
$btnExportCsvReport.Text = "Exportar Reporte a GH (CSV / Excel)"
$btnExportCsvReport.Dock = [System.Windows.Forms.DockStyle]::Right
$btnExportCsvReport.Width = 220
$btnExportCsvReport.BackColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$btnExportCsvReport.ForeColor = [System.Drawing.Color]::White
$btnExportCsvReport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportCsvReport.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$btnExportCsvReport.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlTab3Footer.Controls.Add($btnExportCsvReport)

$btnCopyBatchCreds = New-Object System.Windows.Forms.Button
$btnCopyBatchCreds.Text = "Copiar Credenciales del Lote"
$btnCopyBatchCreds.Dock = [System.Windows.Forms.DockStyle]::Right
$btnCopyBatchCreds.Width = 200
$btnCopyBatchCreds.BackColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$btnCopyBatchCreds.ForeColor = [System.Drawing.Color]::White
$btnCopyBatchCreds.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCopyBatchCreds.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$btnCopyBatchCreds.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlTab3Footer.Controls.Add($btnCopyBatchCreds)

# ======================================================================================
# PANEL DERECHO: FICHA Y CENTRO DE ACCIÓN EN ACTIVE DIRECTORY
# ======================================================================================
$pnlRight = New-Object System.Windows.Forms.Panel
$pnlRight.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlRight.BackColor = [System.Drawing.Color]::White
$pnlRight.Padding = New-Object System.Windows.Forms.Padding(12)
$splitMain.Panel2.Controls.Add($pnlRight)

$tblRight = New-Object System.Windows.Forms.TableLayoutPanel
$tblRight.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblRight.ColumnCount = 1
$tblRight.RowCount = 4
[void]$tblRight.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))  # Header Title
[void]$tblRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # Ficha Card (Toma todo el alto disponible)
[void]$tblRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))  # Botón Copiar Credenciales
[void]$tblRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))  # Boton Accion Principal
$pnlRight.Controls.Add($tblRight)

# Fila 0: Titulo Header
$lblRightTitle = New-Object System.Windows.Forms.Label
$lblRightTitle.Text = "Ficha de Creación en Active Directory"
$lblRightTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblRightTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$lblRightTitle.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblRight.Controls.Add($lblRightTitle, 0, 0)

# Fila 1: Tarjeta de Resumen Visual del Colaborador
$pnlCardSummary = New-Object System.Windows.Forms.Panel
$pnlCardSummary.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlCardSummary.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$pnlCardSummary.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlCardSummary.Padding = New-Object System.Windows.Forms.Padding(12)
$tblRight.Controls.Add($pnlCardSummary, 0, 1)

$tblCardInner = New-Object System.Windows.Forms.TableLayoutPanel
$tblCardInner.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblCardInner.ColumnCount = 1
$tblCardInner.RowCount = 6
[void]$tblCardInner.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) # Nombre
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Cuenta
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Detalles
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) # Badge (26px para evitar cualquier corte)
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) # Titulo grupos
[void]$tblCardInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # ListBox
$pnlCardSummary.Controls.Add($tblCardInner)

$lblSumDisplayName = New-Object System.Windows.Forms.Label
$lblSumDisplayName.Text = "Nombre del Colaborador - SEDE"
$lblSumDisplayName.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$lblSumDisplayName.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$lblSumDisplayName.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblCardInner.Controls.Add($lblSumDisplayName, 0, 0)

$lblSumUser = New-Object System.Windows.Forms.Label
$lblSumUser.Text = "Usuario de Dominio: -"
$lblSumUser.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblSumUser.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
$lblSumUser.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblCardInner.Controls.Add($lblSumUser, 0, 1)

$lblSumDetails = New-Object System.Windows.Forms.Label
$lblSumDetails.Text = "Cédula: -  |  Cargo: -  |  Dept: -"
$lblSumDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblSumDetails.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblSumDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblCardInner.Controls.Add($lblSumDetails, 0, 2)

$lblSumStatusBadge = New-Object System.Windows.Forms.Label
$lblSumStatusBadge.Text = "Estado: Pendiente de Información"
$lblSumStatusBadge.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(217, 119, 6)
$lblSumStatusBadge.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblSumStatusBadge.AutoEllipsis = $true
$tblCardInner.Controls.Add($lblSumStatusBadge, 0, 3)

$lblSumGroupsTitle = New-Object System.Windows.Forms.Label
$lblSumGroupsTitle.Text = "Membresías y Carpetas a Asignar en el DA:"
$lblSumGroupsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblSumGroupsTitle.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblCardInner.Controls.Add($lblSumGroupsTitle, 0, 4)

$lstSumGroups = New-Object System.Windows.Forms.ListBox
$lstSumGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstSumGroups.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$tblCardInner.Controls.Add($lstSumGroups, 0, 5)

# Fila 2: Botón Copiar Credenciales
$btnCopyCreds = New-Object System.Windows.Forms.Button
$btnCopyCreds.Text = "Copiar Datos de Acceso (Usuario y Clave)"
$btnCopyCreds.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnCopyCreds.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)
$btnCopyCreds.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$btnCopyCreds.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$btnCopyCreds.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCopyCreds.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopyCreds.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tblRight.Controls.Add($btnCopyCreds, 0, 2)

# Fila 3: Boton Gigante de Accion Directa
$btnExecuteDirect = New-Object System.Windows.Forms.Button
$btnExecuteDirect.Text = "CREAR USUARIO EN ACTIVE DIRECTORY"
$btnExecuteDirect.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnExecuteDirect.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$btnExecuteDirect.BackColor = [System.Drawing.Color]::FromArgb(5, 150, 105) # Emerald 600
$btnExecuteDirect.ForeColor = [System.Drawing.Color]::White
$btnExecuteDirect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExecuteDirect.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnExecuteDirect.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$tblRight.Controls.Add($btnExecuteDirect, 0, 3)
# ======================================================================================
# 4. LOGICA DE NEGOCIO Y EVENTOS
# ======================================================================================

function Generate-SamAccount {
    $first = Clean-Text $txtFirstName.Text
    $last = Clean-Text $txtLastName.Text
    if (-not $txtPass.Text) {
        $txtPass.Text = "Aliar2026*"
    }
    if (-not $first -or -not $last) { return "" }
    return "$first.$last"
}

function Get-SelectedProfile {
    if ($rbPerfil1.Checked) { return "GR_Perfil_Moderadamente_Restrictivo" }
    if ($rbPerfil2.Checked) { return "GR_Perfil_Muy_Restrictivo" }
    if ($rbPerfil3.Checked) { return "GR_Perfil_Equilibrado" }
    if ($rbPerfil4.Checked) { return "GR_Perfil_Relativamente_Libre" }
    return "GR_Perfil_Moderadamente_Restrictivo"
}

function Get-SelectedSedeCode {
    $info = Get-SelectedSedeInfo
    return $info.Code
}

function Get-SelectedSedeOU {
    $info = Get-SelectedSedeInfo
    return $info.OU
}

function Get-AllTargetGroups {
    $groups = New-Object System.Collections.Generic.List[string]
    $profile = Get-SelectedProfile
    if ($profile) { $groups.Add($profile) }
    if ($chkVpn.Checked) { $groups.Add("GR_VPN_SSL") }
    foreach ($grp in $global:SelectedGroupNames) {
        if (-not $groups.Contains($grp)) {
            $groups.Add($grp)
        }
    }
    return $groups
}

function Update-AllOutputs {
    # ------------------ CASO PESTAÑA 1: CREAR INDIVIDUAL ------------------
    if ($tabControlLeft.SelectedIndex -eq 0) {
        $sam = $txtSam.Text.Trim()
        $sedeCode = Get-SelectedSedeCode
        
        $firstParts = @(To-TitleCase $txtFirstName.Text.Trim(), To-TitleCase $txtMiddleName.Text.Trim()) | Where-Object { $_ }
        $lastParts = @(To-TitleCase $txtLastName.Text.Trim(), To-TitleCase $txtSecondLastName.Text.Trim()) | Where-Object { $_ }
        $fullName = ("$($firstParts -join ' ') $($lastParts -join ' ')").Trim()
        
        if ($fullName) {
            $lblSumDisplayName.Text = "$fullName - $sedeCode"
        } else {
            $lblSumDisplayName.Text = "Nombre del Colaborador - SEDE"
        }
        
        if ($sam) {
            $lblSumUser.Text = "Usuario de Dominio: $sam"
            
            # Búsqueda instantánea O(1) en diccionario sin lag
            $existing = $null
            $samKey = $sam.ToLower()
            if ($global:UserLookupBySam.ContainsKey($samKey)) {
                $existing = $global:UserLookupBySam[$samKey]
            } elseif ($fullName) {
                $nameKey = Clean-Text $fullName
                if ($nameKey -and $global:UserLookupByName.ContainsKey($nameKey)) {
                    $existing = $global:UserLookupByName[$nameKey]
                }
            }
            
            if ($existing) {
                $lblSumStatusBadge.Text = "Estado: [YA REGISTRADO EN EL DA] Sede: $($existing.sede)"
                $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(220, 38, 38)
            } else {
                $lblSumStatusBadge.Text = "Estado: [DISPONIBLE PARA CREAR EN EL DA]"
                $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
            }
        } else {
            $lblSumUser.Text = "Usuario de Dominio: -"
            $lblSumStatusBadge.Text = "Estado: Pendiente de Información"
            $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(217, 119, 6)
        }
        
        $ced = $txtEmployeeId.Text.Trim()
        $crg = To-TitleCase $txtTitle.Text.Trim()
        $dpt = if ($cmbDepartment.SelectedItem) { $cmbDepartment.SelectedItem.ToString() } else { "-" }
        $passDisplay = if ($txtPass.Text.Trim()) { $txtPass.Text.Trim() } else { "Aliar2026*" }
        $lblSumDetails.Text = "Cédula: $ced  |  Cargo: $crg  |  Clave: $passDisplay"
        
        $lstSumGroups.BeginUpdate()
        $lstSumGroups.Items.Clear()
        $targetGroups = Get-AllTargetGroups
        $lblSumGroupsTitle.Text = "Membresías y Carpetas a Asignar en el DA ($($targetGroups.Count)):"
        foreach ($g in $targetGroups) {
            [void]$lstSumGroups.Items.Add($g)
        }
        $lstSumGroups.EndUpdate()
        
        $btnCopyCreds.Visible = $true
        $btnCopyCreds.Text = "Copiar Datos de Acceso (Usuario y Clave)"
        $tblRight.RowStyles[2].Height = 34
        $btnExecuteDirect.Text = "CREAR USUARIO EN ACTIVE DIRECTORY"
        $btnExecuteDirect.BackColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
    }
    # ------------------ CASO PESTANA 2: GESTIONAR PERMISOS ------------------
    elseif ($tabControlLeft.SelectedIndex -eq 1) {
        if ($global:Tab2SelectedUser) {
            $u = $global:Tab2SelectedUser
            $lblSumDisplayName.Text = "$($u.name)"
            $lblSumUser.Text = "Usuario de Dominio: $($u.user)"
            $lblSumDetails.Text = "Sede: $($u.sede)  |  Dept: $($u.department)  |  Cargo: $($u.title)"
            
            # Calcular diferencias
            $toAdd = @()
            $toRemove = @()
            
            $currentWebProfile = ""
            if ($rbTab2Perfil1.Checked) { $currentWebProfile = "GR_Perfil_Moderadamente_Restrictivo" }
            elseif ($rbTab2Perfil2.Checked) { $currentWebProfile = "GR_Perfil_Muy_Restrictivo" }
            elseif ($rbTab2Perfil3.Checked) { $currentWebProfile = "GR_Perfil_Equilibrado" }
            elseif ($rbTab2Perfil4.Checked) { $currentWebProfile = "GR_Perfil_Relativamente_Libre" }
            
            if ($currentWebProfile -and (-not $global:Tab2OriginalGroups.Contains($currentWebProfile))) {
                $toAdd += $currentWebProfile
            }
            
            $hasOrigVpn = ($global:Tab2OriginalGroups.Contains("GR_VPN_SSL") -or $global:Tab2OriginalGroups.Contains("GR_VPN_Acceso"))
            if ($chkTab2Vpn.Checked -and (-not $hasOrigVpn)) {
                $toAdd += "GR_VPN_SSL"
            } elseif ((-not $chkTab2Vpn.Checked) -and $hasOrigVpn) {
                if ($global:Tab2OriginalGroups.Contains("GR_VPN_SSL")) { $toRemove += "GR_VPN_SSL" }
                if ($global:Tab2OriginalGroups.Contains("GR_VPN_Acceso")) { $toRemove += "GR_VPN_Acceso" }
            }
            
            foreach ($grp in $global:Tab2CurrentSelectedGroups) {
                if (-not $global:Tab2OriginalGroups.Contains($grp)) {
                    $toAdd += $grp
                }
            }
            foreach ($grp in $global:Tab2OriginalGroups) {
                if ($grp -like "GR_DC_*" -and (-not $global:Tab2CurrentSelectedGroups.Contains($grp))) {
                    $toRemove += $grp
                }
            }
            
            $lstSumGroups.BeginUpdate()
            $lstSumGroups.Items.Clear()
            if ($toAdd.Count -gt 0 -or $toRemove.Count -gt 0) {
                $lblSumStatusBadge.Text = "Estado: Modificación Pendiente ($($toAdd.Count) agregar / $($toRemove.Count) remover)"
                $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
                $lblSumGroupsTitle.Text = "Diferencial de Permisos para $($u.user):"
                foreach ($g in $toAdd) {
                    [void]$lstSumGroups.Items.Add("[+] AGREGAR: $g")
                }
                foreach ($g in $toRemove) {
                    [void]$lstSumGroups.Items.Add("[-] REMOVER: $g")
                }
            } else {
                $lblSumStatusBadge.Text = "Estado: Sin Cambios Pendientes ($($u.groups.Count) membresías activas)"
                $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
                $lblSumGroupsTitle.Text = "Membresías Actuales de $($u.user) en DA ($($u.groups.Count)):"
                if ($u.groups.Count -eq 0) {
                    [void]$lstSumGroups.Items.Add("(El colaborador no tiene membresías registradas en DA)")
                } else {
                    foreach ($g in ($u.groups | Sort-Object)) {
                        [void]$lstSumGroups.Items.Add("• $g")
                    }
                }
            }
            $lstSumGroups.EndUpdate()
            
            $btnCopyCreds.Visible = $false
            $tblRight.RowStyles[2].Height = 0
            $btnExecuteDirect.Text = "APLICAR CAMBIOS DE PERMISOS EN EL DA"
            $btnExecuteDirect.BackColor = [System.Drawing.Color]::FromArgb(2, 132, 199)
        } else {
            $lblSumDisplayName.Text = "Selecciona un Colaborador en Tab 2"
            $lblSumUser.Text = "Usuario de Dominio: -"
            $lblSumDetails.Text = "Sede: -  |  Dept: -  |  Cargo: -"
            $lblSumStatusBadge.Text = "Estado: Sin Colaborador Seleccionado"
            $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
            $lstSumGroups.Items.Clear()
            $lblSumGroupsTitle.Text = "Membresias a Modificar:"
            $btnCopyCreds.Visible = $false
            $tblRight.RowStyles[2].Height = 0
            $btnExecuteDirect.Text = "SELECCIONA UN COLABORADOR PRIMERO"
            $btnExecuteDirect.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
        }
    }
    # ------------------ CASO PESTANA 3: CARGA MASIVA ------------------
    elseif ($tabControlLeft.SelectedIndex -eq 2) {
        $count = $gridBatch.Rows.Count
        $lblSumDisplayName.Text = "Aprovisionamiento Masivo (Lotes)"
        $lblSumUser.Text = "Total Colaboradores en Lote: $count"
        
        $availableCount = 0
        $alreadyCount = 0
        $sedesPresent = New-Object System.Collections.Generic.HashSet[string]
        $deptsPresent = New-Object System.Collections.Generic.HashSet[string]
        
        $vpnCount = 0
        foreach ($r in $gridBatch.Rows) {
            if ($r.Cells["EstadoDA"].Value -like "*Disponible*") { $availableCount++ } else { $alreadyCount++ }
            if ($r.Cells["Sede"].Value) { [void]$sedesPresent.Add($r.Cells["Sede"].Value.ToString()) }
            if ($r.Cells["Departamento"].Value) { [void]$deptsPresent.Add($r.Cells["Departamento"].Value.ToString()) }
            if ([bool]$r.Cells["Vpn"].Value) { $vpnCount++ }
        }
        
        $lblSumDetails.Text = "Listos para crear: $availableCount  |  Ya registrados: $alreadyCount"
        $lblSumStatusBadge.Text = "Estado: $availableCount cuentas listas para crear en el DA"
        $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
        
        $lstSumGroups.Items.Clear()
        $lblSumGroupsTitle.Text = "Resumen de Sedes, Areas y VPN del Lote:"
        [void]$lstSumGroups.Items.Add("- Sedes en este lote ($($sedesPresent.Count)): " + ($sedesPresent -join ", "))
        [void]$lstSumGroups.Items.Add("- Departamentos ($($deptsPresent.Count)): " + ($deptsPresent -join ", "))
        [void]$lstSumGroups.Items.Add("- Con Acceso VPN Seleccionado: $vpnCount colaboradores")
        [void]$lstSumGroups.Items.Add("- Contenedor Base: USUARIOS (aliar.com.co)")
        [void]$lstSumGroups.Items.Add("- Regla: Cada colaborador se crea con su Sede, Perfil y VPN especifico")
        
        $btnCopyCreds.Visible = $true
        $btnCopyCreds.Text = "Copiar Credenciales del Lote (Para GH)"
        $tblRight.RowStyles[2].Height = 34
        $btnExecuteDirect.Text = "CREAR TODOS LOS USUARIOS DEL LOTE EN EL DA"
        $btnExecuteDirect.BackColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
    }
}

# Auto-completar desde Gestión Humana (Tab 1)
function Process-SmartPaste {
    $raw = $txtSmartPaste.Text.Trim()
    if (-not $raw) {
        [System.Windows.Forms.MessageBox]::Show("Por favor ingresa o pega el texto recibido de Gestión Humana.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $parsed = Parse-AliarWorkerString $raw
    if ($parsed) {
        $txtEmployeeId.Text = $parsed.Cedula
        $txtFirstName.Text = $parsed.PrimerNombre
        $txtMiddleName.Text = $parsed.SegundoNombre
        $txtLastName.Text = $parsed.PrimerApellido
        $txtSecondLastName.Text = $parsed.SegundoApellido
        $txtTitle.Text = $parsed.Cargo
        
        $firstClean = Clean-Text $parsed.PrimerNombre
        $lastClean = Clean-Text $parsed.PrimerApellido
        $newSam = "$firstClean.$lastClean"
        $txtSam.Text = $newSam
        
        if (-not $txtPass.Text) {
            $txtPass.Text = Generate-RandomPassword
        }
        
        # Auto-detección inteligente de sede y departamento
        $inferred = Get-RoleInferredSettings $parsed.Cargo
        
        # Buscar si ya está registrado en AD para heredar su sede exacta
        $cleanSearch = "$firstClean $lastClean"
        $found = $null
        if ($global:UserLookupBySam.ContainsKey($newSam.ToLower())) {
            $found = $global:UserLookupBySam[$newSam.ToLower()]
        } elseif ($global:UserLookupByName.ContainsKey($cleanSearch)) {
            $found = $global:UserLookupByName[$cleanSearch]
        }
        
        $targetSedeCode = if ($found -and $found.sede) { $found.sede } else { $inferred.Sede }
        if ($targetSedeCode) {
            $cfg = $global:SedesConfig | Where-Object { $_.Code -eq $targetSedeCode -or $_.OU -eq $targetSedeCode -or $_.Label -like "*$targetSedeCode*" } | Select-Object -First 1
            if ($cfg) {
                for ($i = 0; $i -lt $cmbSede.Items.Count; $i++) {
                    if ($cmbSede.Items[$i].ToString() -eq $cfg.Label -or $cmbSede.Items[$i].ToString().StartsWith($cfg.OU)) {
                        $cmbSede.SelectedIndex = $i
                        break
                    }
                }
            }
        }
        
        for ($i = 0; $i -lt $cmbDepartment.Items.Count; $i++) {
            if ($cmbDepartment.Items[$i].ToString() -eq $inferred.Departamento) {
                $cmbDepartment.SelectedIndex = $i
                break
            }
        }
        
        # Perfil web auto-detectado
        if ($inferred.PerfilWeb -eq "GR_Perfil_Muy_Restrictivo") { $rbPerfil2.Checked = $true }
        elseif ($inferred.PerfilWeb -eq "GR_Perfil_Equilibrado") { $rbPerfil3.Checked = $true }
        elseif ($inferred.PerfilWeb -eq "GR_Perfil_Relativamente_Libre") { $rbPerfil4.Checked = $true }
        else { $rbPerfil1.Checked = $true }
        
        $chkVpn.Checked = [bool]$inferred.Vpn
        
        Update-AllOutputs
    } else {
        [System.Windows.Forms.MessageBox]::Show("No se pudo interpretar el formato del texto.`nEjemplo válido: 12345678 PEPITO PEREZ RODRIGUEZ OPERARIO PLANTA", "Formato No Reconocido", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

$btnSmartParse.add_Click({ Process-SmartPaste })

$txtSmartPaste.add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Process-SmartPaste
    }
})

$btnSmartClear.add_Click({
    $txtSmartPaste.Text = ""
    $txtFirstName.Text = ""
    $txtMiddleName.Text = ""
    $txtLastName.Text = ""
    $txtSecondLastName.Text = ""
    $txtEmployeeId.Text = ""
    $txtTitle.Text = ""
    $txtSam.Text = ""
    $txtPass.Text = "Aliar2026*"
    $cmbSede.SelectedIndex = 0
    $cmbDepartment.SelectedIndex = 0
    $global:SelectedGroupNames.Clear()
    Refresh-CarpetasList
    Update-AllOutputs
})

# Generar clave aleatoria
$btnGenPass.add_Click({
    $txtPass.Text = Generate-RandomPassword
})

# Sincronizar SAM al escribir nombres o apellidos (sin duplicar llamadas a Update-AllOutputs)
$txtFirstName.add_TextChanged({ $txtSam.Text = Generate-SamAccount })
$txtLastName.add_TextChanged({ $txtSam.Text = Generate-SamAccount })
$txtMiddleName.add_TextChanged({ Update-AllOutputs })
$txtSecondLastName.add_TextChanged({ Update-AllOutputs })
$txtEmployeeId.add_TextChanged({ Update-AllOutputs })
$txtTitle.add_TextChanged({ Update-AllOutputs })

$txtFirstName.add_LostFocus({ if ($txtFirstName.Text) { $txtFirstName.Text = To-TitleCase $txtFirstName.Text } })
$txtMiddleName.add_LostFocus({ if ($txtMiddleName.Text) { $txtMiddleName.Text = To-TitleCase $txtMiddleName.Text } })
$txtLastName.add_LostFocus({ if ($txtLastName.Text) { $txtLastName.Text = To-TitleCase $txtLastName.Text } })
$txtSecondLastName.add_LostFocus({ if ($txtSecondLastName.Text) { $txtSecondLastName.Text = To-TitleCase $txtSecondLastName.Text } })
$txtTitle.add_LostFocus({ if ($txtTitle.Text) { $txtTitle.Text = To-TitleCase $txtTitle.Text } })
$txtSam.add_TextChanged({ Update-AllOutputs })
$txtPass.add_TextChanged({ Update-AllOutputs })
$cmbSede.add_SelectedIndexChanged({ Update-AllOutputs })
$cmbDepartment.add_SelectedIndexChanged({ Update-AllOutputs })

$rbPerfil1.add_CheckedChanged({ Update-AllOutputs })
$rbPerfil2.add_CheckedChanged({ Update-AllOutputs })
$rbPerfil3.add_CheckedChanged({ Update-AllOutputs })
$rbPerfil4.add_CheckedChanged({ Update-AllOutputs })
$chkVpn.add_CheckedChanged({ Update-AllOutputs })

# Búsqueda ultra-rápida de Usuario Modelo en DA (Tab 1)
$txtSearchModelo.add_TextChanged({
    $query = $txtSearchModelo.Text.Trim().ToLower()
    $lstSearchResults.BeginUpdate()
    $lstSearchResults.Items.Clear()
    if ($query.Length -ge 2) {
        $foundCount = 0
        $totalUsers = $global:RealUsers.Count
        for ($i = 0; $i -lt $totalUsers; $i++) {
            $u = $global:RealUsers[$i]
            if ($u.searchStr.IndexOf($query) -ge 0) {
                [void]$lstSearchResults.Items.Add("$($u.name) ($($u.user))")
                $foundCount++
                if ($foundCount -ge 30) { break }
            }
        }
    }
    $lstSearchResults.EndUpdate()
})

$lstSearchResults.add_SelectedIndexChanged({
    if ($lstSearchResults.SelectedIndex -ge 0) {
        $selectedText = $lstSearchResults.SelectedItem.ToString()
        if ($selectedText -match '\((.*?)\)$') {
            $userSam = $matches[1].Trim().ToLower()
            $uObj = $null
            if ($global:UserLookupBySam.ContainsKey($userSam)) {
                $uObj = $global:UserLookupBySam[$userSam]
            }
            if (-not $uObj) {
                $uObj = [PSCustomObject]@{
                    name = ($selectedText -replace '\s*\(.*?\)$', '')
                    user = $userSam
                    groups = @()
                    department = ""
                    title = ""
                    sede = ""
                }
            }
            
            # Consultar en vivo en Active Directory para obtener todas sus membresías reales
            if (Get-Command Get-ADUser -ErrorAction SilentlyContinue) {
                try {
                    $liveModel = Get-ADUser -Identity $userSam -Properties MemberOf, Department, Title, DistinguishedName -ErrorAction SilentlyContinue
                    if ($liveModel) {
                        if ($liveModel.MemberOf) {
                            $liveGrps = @($liveModel.MemberOf | ForEach-Object {
                                if ($_ -match "^CN=([^,]+)") { $matches[1] }
                            })
                            if ($liveGrps.Count -gt 0) { $uObj.groups = $liveGrps }
                        }
                        if ($liveModel.Department) { $uObj.department = $liveModel.Department }
                        if ($liveModel.Title) { $uObj.title = $liveModel.Title }
                        if ($liveModel.DistinguishedName -match 'OU=(AL_[^,]+)') {
                            $cfg = $global:SedesConfig | Where-Object { $_.OU -eq $matches[1] } | Select-Object -First 1
                            if ($cfg) { $uObj.sede = $cfg.Code }
                        }
                    }
                } catch {}
            }
            
            if ($uObj) {
                $lblModeloNombre.Text = "$($uObj.name)"
                $lblModeloCount.Text = "Grupos en DA: $($uObj.groups.Count)"
                $lstModeloGrupos.BeginUpdate()
                $lstModeloGrupos.Items.Clear()
                foreach ($grp in $uObj.groups) {
                    [void]$lstModeloGrupos.Items.Add($grp)
                }
                $lstModeloGrupos.EndUpdate()
                $btnClonarModelo.Enabled = $true
                $btnClonarModelo.Tag = $uObj
            }
        }
    }
})

$btnClonarModelo.add_Click({
    $uObj = $btnClonarModelo.Tag
    if ($uObj) {
        # 1. Sede
        $targetSede = $uObj.sede
        if ($targetSede) {
            $cfg = $global:SedesConfig | Where-Object { $_.Code -eq $targetSede -or $_.OU -eq $targetSede -or $_.Label -like "*$targetSede*" } | Select-Object -First 1
            if ($cfg) {
                for ($i = 0; $i -lt $cmbSede.Items.Count; $i++) {
                    if ($cmbSede.Items[$i].ToString() -eq $cfg.Label -or $cmbSede.Items[$i].ToString().StartsWith($cfg.OU)) {
                        $cmbSede.SelectedIndex = $i
                        break
                    }
                }
            }
        }
        
        # 2. Departamento
        if ($uObj.department) {
            for ($i = 0; $i -lt $cmbDepartment.Items.Count; $i++) {
                if ($cmbDepartment.Items[$i].ToString().ToLower() -eq $uObj.department.ToLower()) {
                    $cmbDepartment.SelectedIndex = $i
                    break
                }
            }
        }
        
        # 3. Perfil Web
        if ($uObj.groups -contains "GR_Perfil_Muy_Restrictivo") { $rbPerfil2.Checked = $true }
        elseif ($uObj.groups -contains "GR_Perfil_Equilibrado") { $rbPerfil3.Checked = $true }
        elseif ($uObj.groups -contains "GR_Perfil_Relativamente_Libre") { $rbPerfil4.Checked = $true }
        elseif ($uObj.groups -contains "GR_Perfil_Moderadamente_Restrictivo") { $rbPerfil1.Checked = $true }
        else { $rbPerfil2.Checked = $true }
        
        # 4. VPN
        $chkVpn.Checked = ($uObj.groups -contains "GR_VPN_SSL" -or $uObj.groups -contains "GR_VPN_Acceso")
        
        # 5. Carpetas y Membresías Departamentales
        $global:SelectedGroupNames.Clear()
        $clonedCount = 0
        foreach ($grp in $uObj.groups) {
            if ($grp -like "GR_Perfil_*" -or $grp -eq "GR_VPN_SSL" -or $grp -eq "GR_VPN_Acceso" -or $grp -eq "Usuarios del dominio" -or $grp -eq "Domain Users") {
                continue
            }
            [void]$global:SelectedGroupNames.Add($grp)
            $clonedCount++
            
            # Asegurar que esté presente en la lista de carpetas para que aparezca tildado
            $already = $global:AreaFolderGroups | Where-Object { $_.name -eq $grp }
            if (-not $already) {
                $newG = [PSCustomObject]@{ name = $grp; desc = "Membresía heredada del modelo" }
                $global:AreaFolderGroups.Add($newG)
            }
        }
        
        Refresh-CarpetasList $txtFilterCarpetas.Text
        Update-AllOutputs
        
        [System.Windows.Forms.MessageBox]::Show("Se clonaron exitosamente los permisos del colaborador modelo '$($uObj.name)':`n`n- Sede: $($uObj.sede)`n- Departamento: $($uObj.department)`n- Membresías departamentales asignadas: $clonedCount grupos`n- VPN: $(if ($chkVpn.Checked) { 'Sí' } else { 'No' })", "Permisos Clonados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$txtFilterCarpetas.add_TextChanged({
    Refresh-CarpetasList $txtFilterCarpetas.Text
})

$btnLimpiarFiltro.add_Click({
    $txtFilterCarpetas.Text = ""
    Refresh-CarpetasList ""
})

$chkListCarpetas.add_ItemCheck({
    param($s, $e)
    $item = $chkListCarpetas.Items[$e.Index]
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        [void]$global:SelectedGroupNames.Add($item)
    } else {
        [void]$global:SelectedGroupNames.Remove($item)
    }
    Update-AllOutputs
})

# ======================================================================================
# EVENTOS DE LA PESTAÑA 2: GESTIONAR PERMISOS (EXISTENTE)
# ======================================================================================
# Búsqueda de Usuario Modelo para clonar en Tab 2
$txtTab2ModSearch.add_TextChanged({
    $query = $txtTab2ModSearch.Text.Trim().ToLower()
    $lstTab2ModResults.BeginUpdate()
    $lstTab2ModResults.Items.Clear()
    if ($query.Length -ge 2) {
        $foundCount = 0
        $totalUsers = $global:RealUsers.Count
        for ($i = 0; $i -lt $totalUsers; $i++) {
            $u = $global:RealUsers[$i]
            if ($u.searchStr.IndexOf($query) -ge 0) {
                [void]$lstTab2ModResults.Items.Add("$($u.name) ($($u.user))")
                $foundCount++
                if ($foundCount -ge 20) { break }
            }
        }
    }
    $lstTab2ModResults.EndUpdate()
})

$lstTab2ModResults.add_SelectedIndexChanged({
    if ($lstTab2ModResults.SelectedIndex -ge 0) {
        $selectedText = $lstTab2ModResults.SelectedItem.ToString()
        if ($selectedText -match '\((.*?)\)$') {
            $modelSam = $matches[1].Trim().ToLower()
            $mObj = $null
            if ($global:UserLookupBySam.ContainsKey($modelSam)) {
                $mObj = $global:UserLookupBySam[$modelSam]
            }
            if (-not $mObj) {
                $mObj = [PSCustomObject]@{
                    name = ($selectedText -replace '\s*\(.*?\)$', '')
                    user = $modelSam
                    groups = @()
                    department = ""
                    title = ""
                    sede = ""
                }
            }
            
            # Consultar en vivo en Active Directory para obtener membresías reales
            if (Get-Command Get-ADUser -ErrorAction SilentlyContinue) {
                try {
                    $liveModel = Get-ADUser -Identity $modelSam -Properties MemberOf, Department, Title, DistinguishedName -ErrorAction SilentlyContinue
                    if ($liveModel -and $liveModel.MemberOf) {
                        $liveGrps = @($liveModel.MemberOf | ForEach-Object {
                            if ($_ -match "^CN=([^,]+)") { $matches[1] }
                        })
                        if ($liveGrps.Count -gt 0) { $mObj.groups = $liveGrps }
                    }
                } catch {}
            }
            
            if ($mObj) {
                $lblTab2ModInfo.Text = "$($mObj.name)"
                $lblTab2ModCount.Text = "Membresías del modelo: $($mObj.groups.Count) grupos"
                $btnTab2Clone.Enabled = $true
            }
        }
    }
})

$btnTab2Clone.add_Click({
    if (-not $global:Tab2SelectedUser) {
        [System.Windows.Forms.MessageBox]::Show("Por favor primero busca y selecciona un colaborador a gestionar en la sección 1 arriba.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ($lstTab2ModResults.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Por favor selecciona un colaborador modelo en la lista.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selectedText = $lstTab2ModResults.SelectedItem.ToString()
    if ($selectedText -match '\((.*?)\)$') {
        $modelSam = $matches[1].Trim().ToLower()
        $mObj = $null
        if ($global:UserLookupBySam.ContainsKey($modelSam)) {
            $mObj = $global:UserLookupBySam[$modelSam]
        }
        if ($mObj) {
            $global:Tab2Refreshing = $true
            try {
                # Aplicar perfil web del modelo
                if ($mObj.groups -contains "GR_Perfil_Muy_Restrictivo") { $rbTab2Perfil2.Checked = $true }
                elseif ($mObj.groups -contains "GR_Perfil_Equilibrado") { $rbTab2Perfil3.Checked = $true }
                elseif ($mObj.groups -contains "GR_Perfil_Relativamente_Libre") { $rbTab2Perfil4.Checked = $true }
                elseif ($mObj.groups -contains "GR_Perfil_Moderadamente_Restrictivo") { $rbTab2Perfil1.Checked = $true }
                else { $rbTab2Perfil2.Checked = $true }
                
                # Aplicar VPN del modelo
                $chkTab2Vpn.Checked = ($mObj.groups -contains "GR_VPN_Acceso" -or $mObj.groups -contains "GR_VPN_SSL")
                
                # Aplicar carpetas y membresías departamentales del modelo
                $global:Tab2CurrentSelectedGroups.Clear()
                $clonedCount = 0
                foreach ($grp in $mObj.groups) {
                    if ($grp -like "GR_Perfil_*" -or $grp -eq "GR_VPN_SSL" -or $grp -eq "GR_VPN_Acceso" -or $grp -eq "Usuarios del dominio" -or $grp -eq "Domain Users") {
                        continue
                    }
                    [void]$global:Tab2CurrentSelectedGroups.Add($grp)
                    $clonedCount++
                    
                    # Asegurar presencia en checklist Tab 2
                    $already = $global:AreaFolderGroups | Where-Object { $_.name -eq $grp }
                    if (-not $already) {
                        $newG = [PSCustomObject]@{ name = $grp; desc = "Membresía heredada del modelo" }
                        $global:AreaFolderGroups.Add($newG)
                    }
                }
                
                Refresh-Tab2CarpetasList $txtTab2FilterCarpetas.Text
            } finally {
                $global:Tab2Refreshing = $false
            }
            Update-AllOutputs
            [System.Windows.Forms.MessageBox]::Show("Se clonaron los permisos del modelo '$($mObj.name)' al formulario:`n`n- Membresías departamentales asignadas: $clonedCount grupos`n- VPN: $(if ($chkTab2Vpn.Checked) { 'Sí' } else { 'No' })`n`nRevisa el diferencial de permisos a aplicar en el panel derecho.", "Permisos Clonados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# ======================================================================================
$txtTab2Search.add_TextChanged({
    $query = $txtTab2Search.Text.Trim().ToLower()
    $lstTab2SearchResults.BeginUpdate()
    $lstTab2SearchResults.Items.Clear()
    if ($query.Length -ge 2) {
        $foundCount = 0
        $totalUsers = $global:RealUsers.Count
        for ($i = 0; $i -lt $totalUsers; $i++) {
            $u = $global:RealUsers[$i]
            if ($u.searchStr.IndexOf($query) -ge 0) {
                [void]$lstTab2SearchResults.Items.Add("$($u.name) ($($u.user))")
                $foundCount++
                if ($foundCount -ge 30) { break }
            }
        }
    }
    $lstTab2SearchResults.EndUpdate()
})

$lstTab2SearchResults.add_SelectedIndexChanged({
    if ($lstTab2SearchResults.SelectedIndex -ge 0) {
        $selectedText = $lstTab2SearchResults.SelectedItem.ToString()
        if ($selectedText -match '\((.*?)\)$') {
            $userSam = $matches[1].Trim().ToLower()
            $uObj = $null
            if ($global:UserLookupBySam.ContainsKey($userSam)) {
                $uObj = $global:UserLookupBySam[$userSam]
            }
            if ($uObj) {
                $global:Tab2Refreshing = $true
                try {
                    $global:Tab2SelectedUser = $uObj
                    $global:Tab2OriginalGroups.Clear()
                    $global:Tab2CurrentSelectedGroups.Clear()
                    
                    # Consultar en vivo en Active Directory para obtener los grupos actualizados
                    if (Get-Command Get-ADUser -ErrorAction SilentlyContinue) {
                        try {
                            $adU = Get-ADUser -Identity $userSam -Properties MemberOf, Department, Title -ErrorAction SilentlyContinue
                            if ($adU -and $adU.MemberOf) {
                                $liveMbrs = @()
                                foreach ($dn in $adU.MemberOf) {
                                    if ($dn -match "^CN=([^,]+)") { $liveMbrs += $matches[1] }
                                }
                                if ($liveMbrs.Count -gt 0) { $uObj.groups = $liveMbrs }
                            }
                            if ($adU.Department) { $uObj.department = $adU.Department }
                            if ($adU.Title) { $uObj.title = $adU.Title }
                        } catch {}
                    }
                    
                    foreach ($grp in $uObj.groups) {
                        [void]$global:Tab2OriginalGroups.Add($grp)
                        if ($grp -like "GR_DC_*") {
                            [void]$global:Tab2CurrentSelectedGroups.Add($grp)
                        }
                    }
                    
                    $lblTab2ColabName.Text = "$($uObj.name)"
                    $lblTab2ColabAccount.Text = "Usuario de Dominio: $($uObj.user)"
                    $lblTab2ColabDetails.Text = "Sede: $($uObj.sede)  |  Dept: $($uObj.department)  |  Cargo: $($uObj.title)"
                    
                    $pWeb = "Moderadamente Restrictivo"
                    if ($uObj.groups -contains "GR_Perfil_Muy_Restrictivo") { $pWeb = "Muy Restrictivo"; $rbTab2Perfil2.Checked = $true }
                    elseif ($uObj.groups -contains "GR_Perfil_Equilibrado") { $pWeb = "Equilibrado"; $rbTab2Perfil3.Checked = $true }
                    elseif ($uObj.groups -contains "GR_Perfil_Relativamente_Libre") { $pWeb = "Relativamente Libre"; $rbTab2Perfil4.Checked = $true }
                    else { $rbTab2Perfil1.Checked = $true }
                    
                    $hasVpn = ($uObj.groups -contains "GR_VPN_Acceso" -or $uObj.groups -contains "GR_VPN_SSL")
                    $chkTab2Vpn.Checked = $hasVpn
                    $vpnText = if ($hasVpn) { "Habilitado" } else { "No asignado" }
                    
                    $lblTab2ColabAccess.Text = "Perfil Web: $pWeb  |  VPN: $vpnText"
                    $lblTab2CurrentPermsTitle.Text = "Membresías y Carpetas en DA ($($uObj.groups.Count) grupos):"
                    
                    $lstTab2CurrentPerms.BeginUpdate()
                    $lstTab2CurrentPerms.Items.Clear()
                    if ($uObj.groups.Count -eq 0) {
                        [void]$lstTab2CurrentPerms.Items.Add("(Sin membresías o carpetas asignadas en DA)")
                    } else {
                        foreach ($g in ($uObj.groups | Sort-Object)) {
                            [void]$lstTab2CurrentPerms.Items.Add($g)
                        }
                    }
                    $lstTab2CurrentPerms.EndUpdate()
                    
                    Refresh-Tab2CarpetasList $txtTab2FilterCarpetas.Text
                } finally {
                    $global:Tab2Refreshing = $false
                }
                Update-AllOutputs
            }
        }
    }
})

$txtTab2FilterCarpetas.add_TextChanged({
    Refresh-Tab2CarpetasList $txtTab2FilterCarpetas.Text
})

$btnTab2LimpiarFiltro.add_Click({
    $txtTab2FilterCarpetas.Text = ""
    Refresh-Tab2CarpetasList ""
})

$chkListTab2Carpetas.add_ItemCheck({
    param($s, $e)
    if (-not $global:Tab2Refreshing) {
        $item = $chkListTab2Carpetas.Items[$e.Index]
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            [void]$global:Tab2CurrentSelectedGroups.Add($item)
        } else {
            [void]$global:Tab2CurrentSelectedGroups.Remove($item)
        }
        Update-AllOutputs
    }
})

$rbTab2Perfil1.add_CheckedChanged({ if (-not $global:Tab2Refreshing) { Update-AllOutputs } })
$rbTab2Perfil2.add_CheckedChanged({ if (-not $global:Tab2Refreshing) { Update-AllOutputs } })
$rbTab2Perfil3.add_CheckedChanged({ if (-not $global:Tab2Refreshing) { Update-AllOutputs } })
$rbTab2Perfil4.add_CheckedChanged({ if (-not $global:Tab2Refreshing) { Update-AllOutputs } })
$chkTab2Vpn.add_CheckedChanged({ if (-not $global:Tab2Refreshing) { Update-AllOutputs } })

# ======================================================================================
# EVENTOS DE LA PESTANA 3: CARGA MASIVA MULTI-CARGO / MULTI-SEDE
# ======================================================================================
$btnBatchLoadFile.add_Click({
    $openDlg = New-Object System.Windows.Forms.OpenFileDialog
    $openDlg.Filter = "Archivos de Texto o CSV (*.csv;*.txt)|*.csv;*.txt|Todos los archivos (*.*)|*.*"
    if ($openDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $content = Get-Content $openDlg.FileName -Raw -Encoding utf8
        $txtBatchInput.Text = $content
    }
})

$btnBatchClear.add_Click({
    $txtBatchInput.Text = ""
    $gridBatch.Rows.Clear()
    $lblBatchStatus.Text = "Total en Lote: 0 colaboradores | Disponibles: 0 | Ya registrados en DA: 0"
    Update-AllOutputs
})

# Procesar Lote con Auto-Detección Inteligente Individual por Fila (Optimizado sin congelamiento)
$gridBatch.add_CellClick({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
        $col = $gridBatch.Columns[$e.ColumnIndex]
        if ($col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
            $gridBatch.BeginEdit($true)
            if ($gridBatch.EditingControl -is [System.Windows.Forms.ComboBox]) {
                ($gridBatch.EditingControl -as [System.Windows.Forms.ComboBox]).DroppedDown = $true
            }
        }
    }
})

$gridBatch.add_CellEnter({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
        $col = $gridBatch.Columns[$e.ColumnIndex]
        if ($col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
            $gridBatch.BeginEdit($true)
            if ($gridBatch.EditingControl -is [System.Windows.Forms.ComboBox]) {
                ($gridBatch.EditingControl -as [System.Windows.Forms.ComboBox]).DroppedDown = $true
            }
        }
    }
})

$gridBatch.add_CellContentClick({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
        if ($gridBatch.Columns[$e.ColumnIndex] -is [System.Windows.Forms.DataGridViewCheckBoxColumn]) {
            $gridBatch.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    }
})

$global:BatchLoading = $false
$btnProcessBatch.add_Click({
    $lines = $txtBatchInput.Lines | Where-Object { $_.Trim() }
    if ($lines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Por favor ingresa o carga la lista de colaboradores en la caja de texto superior.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $gridBatch.SuspendLayout()
    $global:BatchLoading = $true
    $gridBatch.Rows.Clear()
    
    $availableCount = 0
    $alreadyCount = 0
    
    foreach ($line in $lines) {
        $parsed = Parse-AliarWorkerString $line
        if ($parsed) {
            $first = Clean-Text $parsed.PrimerNombre
            $last = Clean-Text $parsed.PrimerApellido
            $userSam = "$first.$last"
            $pass = Generate-RandomPassword
            
            # Auto-detección inteligente específica para este colaborador
            $inferred = Get-RoleInferredSettings $parsed.Cargo
            $rowSede = $inferred.Sede
            $rowDept = $inferred.Departamento
            $rowProfile = $inferred.PerfilWeb
            
            $firsts = @($parsed.PrimerNombre, $parsed.SegundoNombre) | Where-Object { $_ }
            $lasts = @($parsed.PrimerApellido, $parsed.SegundoApellido) | Where-Object { $_ }
            $fullDisplay = ("$($firsts -join ' ') $($lasts -join ' ')").Trim()
            
            # Validación ultra-rápida en memoria O(1) contra Active Directory
            $exists = $null
            $samKey = $userSam.ToLower()
            if ($global:UserLookupBySam.ContainsKey($samKey)) {
                $exists = $global:UserLookupBySam[$samKey]
            } elseif ($fullDisplay) {
                $nameKey = Clean-Text $fullDisplay
                if ($nameKey -and $global:UserLookupByName.ContainsKey($nameKey)) {
                    $exists = $global:UserLookupByName[$nameKey]
                }
            }
            
            $estadoDA = "Disponible para Crear"
            if ($exists) {
                $estadoDA = "Ya Registrado en DA ($($exists.sede))"
                if ($exists.sede -and ($global:SedesCodigos -contains $exists.sede)) {
                    $rowSede = $exists.sede
                }
                $alreadyCount++
            } else {
                $availableCount++
            }
            
            # Sanitizar valores contra las listas válidas de los ComboBoxColumn para evitar DataError
                        if ($rowSede -and (-not $colSede.Items.Contains($rowSede))) {
                [void]$colSede.Items.Add($rowSede)
            }
            if (-not $rowSede) { $rowSede = "CUN_PLENG" }
            if (-not ($global:DepartamentosAliar -contains $rowDept)) { $rowDept = "Planta Engativá" }
            if (-not ($global:PerfilesWeb -contains $rowProfile)) { $rowProfile = "GR_Perfil_Muy_Restrictivo" }
            
            
            $hasVpn = $true # Por defecto activado para todo el lote
            
            $rowIdx = $gridBatch.Rows.Add(
                $parsed.Cedula,
                $fullDisplay,
                $userSam,
                $parsed.Cargo,
                $rowSede,
                $rowDept,
                $rowProfile,
                $hasVpn,
                $pass,
                $estadoDA
            )
            
            # Colorear estado
            if ($exists) {
                $gridBatch.Rows[$rowIdx].Cells["EstadoDA"].Style.ForeColor = [System.Drawing.Color]::FromArgb(220, 38, 38)
            } else {
                $gridBatch.Rows[$rowIdx].Cells["EstadoDA"].Style.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
            }
        } else {
        }
    }
    
    $gridBatch.ResumeLayout()
    $global:BatchLoading = $false
    
    $total = $gridBatch.Rows.Count
    $lblBatchStatus.Text = "Total en Lote: $total colaboradores | Listos para crear: $availableCount | Ya registrados en DA: $alreadyCount"
    Update-AllOutputs
    [System.Windows.Forms.MessageBox]::Show("Se procesaron $total colaboradores individualmente:`n`n- Cada fila ha sido auto-clasificada con su propia Sede, Departamento, Perfil Web y Acceso VPN.`n- Cuentas listas para creación: $availableCount`n- Cuentas ya registradas: $alreadyCount`n`nPuedes modificar la Sede, Departamento o marcar/desmarcar la casilla VPN de cualquier colaborador en la tabla.", "Lote Procesado con Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Actualizar resumen al cambiar cualquier celda de la tabla
$gridBatch.add_CellValueChanged({
    param($s, $e)
    if (-not $global:BatchLoading) {
        Update-AllOutputs
    }
})

function Copy-SingleUserCredentialsToClipboard($notify = $true) {
    $sam = $txtSam.Text.Trim()
    $pass = if ($txtPass.Text.Trim()) { $txtPass.Text.Trim() } else { "Aliar2026*" }
    
    if (-not $sam) {
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("Por favor ingresa primero la información del colaborador para generar su usuario de dominio.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        return $false
    }
    
    $clip = "Usuario de dominio: $sam`r`nContraseña: $pass"
    
    try {
        [System.Windows.Forms.Clipboard]::SetText($clip)
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("Datos de acceso copiados al portapapeles exitosamente:`n`n$clip", "Copiado al Portapapeles", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        return $true
    } catch {
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("Error al copiar al portapapeles: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        return $false
    }
}

function Copy-BatchCredentialsToClipboard($notify = $true) {
    if ($gridBatch.Rows.Count -eq 0) {
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("No hay colaboradores en la tabla para copiar credenciales.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        return $false
    }
    
    $sb = New-Object System.Text.StringBuilder
    $count = 0
    foreach ($row in $gridBatch.Rows) {
        $uSam = $row.Cells["Usuario"].Value
        if (-not $uSam) { continue }
        $count++
        $uPass = $row.Cells["Clave"].Value
        if (-not $uPass) { $uPass = "Aliar2026*" }
        
        [void]$sb.AppendLine("Usuario de dominio: $uSam")
        [void]$sb.AppendLine("Contraseña: $uPass")
        [void]$sb.AppendLine("")
    }
    
    $clip = $sb.ToString().TrimEnd()
    
    try {
        [System.Windows.Forms.Clipboard]::SetText($clip)
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("¡Credenciales de $count colaboradores copiadas al portapapeles exitosamente!`n`nFormato copiado:`nUsuario de dominio: nombre.apellido`nContraseña: clave", "Copiado al Portapapeles", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        return $true
    } catch {
        if ($notify) {
            [System.Windows.Forms.MessageBox]::Show("Error al copiar al portapapeles: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        return $false
    }
}

# Evento botón Copiar Credenciales en Panel Derecho
$btnCopyCreds.add_Click({
    if ($tabControlLeft.SelectedIndex -eq 0) {
        Copy-SingleUserCredentialsToClipboard $true
    } elseif ($tabControlLeft.SelectedIndex -eq 2) {
        Copy-BatchCredentialsToClipboard $true
    }
})

# Evento botón Copiar Credenciales en Footer Tab 3
$btnCopyBatchCreds.add_Click({
    Copy-BatchCredentialsToClipboard $true
})

# Menú contextual de clic derecho en la tabla de Carga Masiva (Tab 3)
$ctxGridBatch = New-Object System.Windows.Forms.ContextMenuStrip
$miCopyRow = New-Object System.Windows.Forms.ToolStripMenuItem("Copiar Usuario y Clave de este Colaborador")
$miCopyAll = New-Object System.Windows.Forms.ToolStripMenuItem("Copiar Todas las Credenciales del Lote (Para GH)")
[void]$ctxGridBatch.Items.Add($miCopyRow)
[void]$ctxGridBatch.Items.Add($miCopyAll)
$gridBatch.ContextMenuStrip = $ctxGridBatch

$miCopyRow.add_Click({
    if ($gridBatch.CurrentRow) {
        $r = $gridBatch.CurrentRow
        $sam = $r.Cells["Usuario"].Value
        $pass = $r.Cells["Clave"].Value
        $nom = $r.Cells["Nombre"].Value
        if ($sam) {
            if (-not $pass) { $pass = "Aliar2026*" }
            $txt = "Usuario de dominio: $sam`r`nContraseña: $pass"
            try {
                [System.Windows.Forms.Clipboard]::SetText($txt)
                [System.Windows.Forms.MessageBox]::Show("Datos de acceso copiados al portapapeles:`n`n$txt", "Copiado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {}
        }
    }
})

$miCopyAll.add_Click({
    Copy-BatchCredentialsToClipboard $true
})

# Exportar Reporte Masivo a CSV para Gestion Humana
$btnExportCsvReport.add_Click({
    if ($gridBatch.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hay datos en la tabla para exportar. Procesa un lote primero.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Archivo CSV (*.csv)|*.csv"
    $saveDialog.FileName = "Reporte_Creacion_Usuarios_Aliar_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvRows = @()
        foreach ($row in $gridBatch.Rows) {
            $sedeCode = $row.Cells["Sede"].Value
            $nomComp = "$($row.Cells['Nombre'].Value) - $sedeCode"
            $uSam = $row.Cells["Usuario"].Value
            
            $vpnVal = if ([bool]$row.Cells["Vpn"].Value) { "SI" } else { "NO" }
            $csvRows += [PSCustomObject]@{
                "Cedula" = $row.Cells["Cedula"].Value
                "Nombre Completo" = $nomComp
                "Usuario de Dominio" = $uSam
                "Correo Corporativo" = "$uSam@aliar.com.co"
                "Cargo" = $row.Cells["Cargo"].Value
                "Sede" = $sedeCode
                "Departamento" = $row.Cells["Departamento"].Value
                "Perfil Web" = $row.Cells["PerfilWeb"].Value
                "Acceso VPN" = $vpnVal
                "Contrasena Temporal" = $row.Cells["Clave"].Value
                "Estado en Active Directory" = $row.Cells["EstadoDA"].Value
                "Fecha Creacion" = (Get-Date -Format "yyyy-MM-dd HH:mm")
            }
        }
        $csvRows | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding utf8
        [System.Windows.Forms.MessageBox]::Show("Reporte para Gestion Humana exportado con exito en:`n$($saveDialog.FileName)", "Exportacion Completa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Cambio de Pestana Principal
$tabControlLeft.add_SelectedIndexChanged({
    Update-AllOutputs
})

# ======================================================================================
# SINCRONIZACIÓN Y CONSULTA ACTUALIZADA DEL DOMINIO (EN VIVO O LOCAL)
# ======================================================================================
function Sync-ADLive($showSuccessDialog = $false) {
    $btnRefreshAD.Enabled = $false
    $btnRefreshAD.Text = "Consultando DA..."
    $form.Refresh()
    
    try {
        if (Get-Command Get-ADGroup -ErrorAction SilentlyContinue) {
            $liveGroups = Get-ADGroup -Filter * -Properties Description -ErrorAction Stop | Select-Object Name, Description
            if ($liveGroups -and $liveGroups.Count -gt 0) {
                $global:RealGroups = New-Object System.Collections.Generic.List[PSCustomObject]
                $global:AreaFolderGroups = New-Object System.Collections.Generic.List[PSCustomObject]
                $global:OtherGroups = New-Object System.Collections.Generic.List[PSCustomObject]
                foreach ($g in $liveGroups) {
                    $obj = [PSCustomObject]@{ name = $g.Name; desc = if ($g.Description) { $g.Description } else { "" } }
                    $global:RealGroups.Add($obj)
                    if ($obj.name -like "GR_DC_*") {
                        $global:AreaFolderGroups.Add($obj)
                    } else {
                        $global:OtherGroups.Add($obj)
                    }
                }
                $lblStats.Text = "(*) Dominio Conectado: $($global:RealGroups.Count) Grupos de Dominio"
                $lblStats.ForeColor = [System.Drawing.Color]::FromArgb(52, 211, 153)
                Refresh-CarpetasList $txtFilterCarpetas.Text
                Refresh-Tab2CarpetasList $txtTab2FilterCarpetas.Text
                
                # Descubrir en vivo la ruta base USUARIOS y cada OU AL_*
                try {
                    $liveOUs = Get-ADOrganizationalUnit -Filter "Name -like 'AL_*' -or Name -eq 'USUARIOS'" -ErrorAction SilentlyContinue
                    if ($liveOUs) {
                        foreach ($ou in $liveOUs) {
                            $global:AD_OU_Map[$ou.Name] = $ou.DistinguishedName
                            if ($ou.Name -eq "USUARIOS") {
                                $global:UsuariosBaseOU = $ou.DistinguishedName
                            } elseif ($ou.Name.StartsWith("AL_")) {
                                $knownOU = $global:SedesConfig | Where-Object { $_.OU -eq $ou.Name }
                                if (-not $knownOU) {
                                    $newObj = [PSCustomObject]@{
                                        OU = $ou.Name
                                        Code = ($ou.Name -replace '^AL_', '')
                                        Label = "$($ou.Name) (OU Detectada en DA)"
                                    }
                                    $global:SedesConfig += $newObj
                                    [void]$cmbSede.Items.Add($newObj.Label)
                                    if (-not $colSede.Items.Contains($newObj.OU)) { [void]$colSede.Items.Add($newObj.OU) }
                                    if (-not $colSede.Items.Contains($newObj.Code)) { [void]$colSede.Items.Add($newObj.Code) }
                                }
                            }
                        }
                    }
                } catch {}
                
                if ($showSuccessDialog) {
                    [System.Windows.Forms.MessageBox]::Show("Se sincronizó exitosamente la lista de grupos y unidades organizativas (OU) en vivo desde el Dominio aliar.com.co.", "Dominio Sincronizado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
        } else {
            $lblStats.Text = "Base Local: $($global:RealGroups.Count) Grupos | $($global:RealUsers.Count) Colaboradores"
            if ($showSuccessDialog) {
                [System.Windows.Forms.MessageBox]::Show("El módulo de ActiveDirectory no está disponible o este equipo no está conectado al controlador de dominio.`nSe mantienen los $($global:RealGroups.Count) grupos corporativos de la base local.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    } catch {
        $lblStats.Text = "Base Local: $($global:RealGroups.Count) Grupos | $($global:RealUsers.Count) Colaboradores"
        if ($showSuccessDialog) {
            [System.Windows.Forms.MessageBox]::Show("No se pudo conectar con el Controlador de Dominio:`n`n$_`n`nSe continuará con la base local.", "Aviso de Conexión", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } finally {
        $btnRefreshAD.Text = "Refrescar DA en Vivo"
        $btnRefreshAD.Enabled = $true
    }
}

$btnRefreshAD.add_Click({
    Sync-ADLive $true
})

# ======================================================================================
# 5. EJECUCION DIRECTA EN ACTIVE DIRECTORY (BOTON ACCION PRINCIPAL)
# ======================================================================================
$btnExecuteDirect.add_Click({
    # ------------------ CASO PESTANA 1: CREAR INDIVIDUAL ------------------
    if ($tabControlLeft.SelectedIndex -eq 0) {
        $sam = $txtSam.Text.Trim()
        $first = To-TitleCase $txtFirstName.Text.Trim()
        $last = To-TitleCase $txtLastName.Text.Trim()
        $middle = To-TitleCase $txtMiddleName.Text.Trim()
        $secondLast = To-TitleCase $txtSecondLastName.Text.Trim()
        $ced = $txtEmployeeId.Text.Trim()
        $title = To-TitleCase $txtTitle.Text.Trim()
        $dept = if ($cmbDepartment.SelectedItem) { $cmbDepartment.SelectedItem.ToString() } else { "Planta Engativa" }
        $sedeCode = Get-SelectedSedeCode
        $pass = $txtPass.Text.Trim()
        if (-not $pass) {
            $pass = "Aliar2026*"
            $txtPass.Text = $pass
        }
        $targetGroups = Get-AllTargetGroups
        
        if (-not $sam -or -not $first -or -not $last -or -not $ced) {
            [System.Windows.Forms.MessageBox]::Show("Por favor completa los campos obligatorios (*) antes de aprovisionar en el DA.", "Campos Requeridos", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $firstParts = @($first, $middle) | Where-Object { $_ }
        $lastParts = @($last, $secondLast) | Where-Object { $_ }
        $displayName = ("$($firstParts -join ' ') $($lastParts -join ' ') - $sedeCode").Trim()
        
        $sedeInfo = Get-SelectedSedeInfo
        $targetOU = Get-TargetOUPath $sedeInfo.OU
        $msgConfirm = "¿Confirma la creación del usuario en Active Directory?`n`n" +
                      "Colaborador: $displayName`n" +
                      "Usuario de Dominio: $sam`n" +
                      "Cédula: $ced`n" +
                      "Cargo: $title`n" +
                      "Sede / Ubicación: $($sedeInfo.Label)`n" +
                      "Ruta en Active Directory: $targetOU`n" +
                      "Departamento: $dept`n" +
                      "Contraseña Temporal: $pass`n" +
                      "Membresías a asignar: $($targetGroups.Count) grupos"
        
                # Validación de duplicidad: Verificar si ya existe en Active Directory
        $alreadyRegistered = $null
        $samKey = $sam.ToLower()
        if ($global:UserLookupBySam.ContainsKey($samKey)) {
            $alreadyRegistered = $global:UserLookupBySam[$samKey]
        } elseif ($displayName) {
            $firstParts = @($first, $middle) | Where-Object { $_ }
            $lastParts = @($last, $secondLast) | Where-Object { $_ }
            $fullClean = Clean-Text ("$($firstParts -join ' ') $($lastParts -join ' ')")
            if ($fullClean -and $global:UserLookupByName.ContainsKey($fullClean)) {
                $alreadyRegistered = $global:UserLookupByName[$fullClean]
            }
        }
        
        # Validación en vivo si Get-ADUser está disponible
        if (-not $alreadyRegistered -and (Get-Command Get-ADUser -ErrorAction SilentlyContinue)) {
            try {
                $liveUser = Get-ADUser -Identity $sam -ErrorAction SilentlyContinue
                if ($liveUser) {
                    $alreadyRegistered = [PSCustomObject]@{ name = $liveUser.Name; user = $liveUser.SamAccountName; sede = "" }
                }
            } catch {}
        }
        
        if ($alreadyRegistered) {
            [System.Windows.Forms.MessageBox]::Show("No se puede crear el usuario porque ya se encuentra registrado en Active Directory:`n`nUsuario de Dominio: $($alreadyRegistered.user)`nNombre en DA: $($alreadyRegistered.name)`n`nNo se permite crear cuentas duplicadas en el dominio. Para modificar sus permisos o sede, ve a la pestaña '2. Gestionar Permisos (Existente)'.", "Colaborador Ya Registrado en DA", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $res = [System.Windows.Forms.MessageBox]::Show($msgConfirm, "Confirmación de Aprovisionamiento", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        
        try {
            if (-not (Get-Command New-ADUser -ErrorAction SilentlyContinue)) {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            
            $sedeInfo = Get-SelectedSedeInfo
            $targetOU = Get-TargetOUPath $sedeInfo.OU
            
            $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
            New-ADUser -Name $displayName `
                       -GivenName ("$first $middle").Trim() `
                       -Surname ("$last $secondLast").Trim() `
                       -DisplayName $displayName `
                       -SamAccountName $sam `
                       -UserPrincipalName "$sam@aliar.com.co" `
                       -Department $dept `
                       -Title $title `
                       -EmployeeID $ced `
                       -Path $targetOU `
                       -AccountPassword $secPass `
                       -Enabled $true `
                       -ChangePasswordAtLogon $true `
                       -PasswordNeverExpires $false
            
            $assignedCount = 0
            $failedGroups = @()
            foreach ($grp in $targetGroups) {
                $added = $false
                for ($retry = 0; $retry -lt 3 -and (-not $added); $retry++) {
                    try {
                        Add-ADGroupMember -Identity $grp -Members $sam -ErrorAction Stop
                        $assignedCount++
                        $added = $true
                    } catch {
                        if ($grp -eq "GR_VPN_SSL") {
                            try {
                                Add-ADGroupMember -Identity "GR_VPN_Acceso" -Members $sam -ErrorAction Stop
                                $assignedCount++
                                $added = $true
                                break
                            } catch {}
                        }
                        if ($retry -lt 2) {
                            Start-Sleep -Milliseconds 250
                        }
                    }
                }
                if (-not $added) {
                    $failedGroups += $grp
                }
            }
            
            $lblSumStatusBadge.Text = "Estado: [CREADO EXITOSAMENTE EN EL DA]"
            $lblSumStatusBadge.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
            
            Copy-SingleUserCredentialsToClipboard $false
            
            if ($failedGroups.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Colaborador aprovisionado en Active Directory:`n`nUsuario de Dominio: $sam`nContraseña Temporal: $pass`nUbicación OU: $targetOU`nMembresías asignadas: $assignedCount de $($targetGroups.Count)`n`n(Los datos de acceso han sido copiados automáticamente al portapapeles para notificar a GH).`n`nAtención: Los siguientes grupos no se pudieron asignar automáticamente:`n" + ($failedGroups -join ", "), "Aprovisionado con Observaciones", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            } else {
                [System.Windows.Forms.MessageBox]::Show("¡Colaborador aprovisionado con éxito en Active Directory!`n`nUsuario de Dominio: $sam`nContraseña Temporal: $pass`nUbicación OU: $targetOU`nMembresías asignadas: $assignedCount de $($targetGroups.Count)`n`n(Los datos de acceso han sido copiados automáticamente al portapapeles listos para pegar en correo o chat).", "Operación Exitosa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Ocurrió un error al ejecutar la creación en Active Directory:`n$_", "Error en Active Directory", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    # ------------------ CASO PESTAÑA 2: GESTIONAR PERMISOS EXISTENTE ------------------
    elseif ($tabControlLeft.SelectedIndex -eq 1) {
        if (-not $global:Tab2SelectedUser) {
            [System.Windows.Forms.MessageBox]::Show("Por favor busca y selecciona un colaborador registrado en la lista de arriba.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $u = $global:Tab2SelectedUser
        $targetUser = $u.user
        
        $toAdd = @()
        $toRemove = @()
        
        $currentWebProfile = ""
        if ($rbTab2Perfil1.Checked) { $currentWebProfile = "GR_Perfil_Moderadamente_Restrictivo" }
        elseif ($rbTab2Perfil2.Checked) { $currentWebProfile = "GR_Perfil_Muy_Restrictivo" }
        elseif ($rbTab2Perfil3.Checked) { $currentWebProfile = "GR_Perfil_Equilibrado" }
        elseif ($rbTab2Perfil4.Checked) { $currentWebProfile = "GR_Perfil_Relativamente_Libre" }
        
        if ($currentWebProfile -and (-not $global:Tab2OriginalGroups.Contains($currentWebProfile))) {
            $toAdd += $currentWebProfile
        }
        
        $hasOrigVpn = ($global:Tab2OriginalGroups.Contains("GR_VPN_SSL") -or $global:Tab2OriginalGroups.Contains("GR_VPN_Acceso"))
        if ($chkTab2Vpn.Checked -and (-not $hasOrigVpn)) {
            $toAdd += "GR_VPN_SSL"
        } elseif ((-not $chkTab2Vpn.Checked) -and $hasOrigVpn) {
            if ($global:Tab2OriginalGroups.Contains("GR_VPN_SSL")) { $toRemove += "GR_VPN_SSL" }
            if ($global:Tab2OriginalGroups.Contains("GR_VPN_Acceso")) { $toRemove += "GR_VPN_Acceso" }
        }
        
        foreach ($grp in $global:Tab2CurrentSelectedGroups) {
            if (-not $global:Tab2OriginalGroups.Contains($grp)) {
                $toAdd += $grp
            }
        }
        foreach ($grp in $global:Tab2OriginalGroups) {
            if ($grp -like "GR_DC_*" -and (-not $global:Tab2CurrentSelectedGroups.Contains($grp))) {
                $toRemove += $grp
            }
        }
        
        if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No hay cambios pendientes de permisos para el colaborador '$targetUser'.", "Sin Cambios", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $msgConfirm = "¿Desea aplicar los siguientes cambios de membresías para '$targetUser'?`n`n" +
                      "Grupos a agregar: $($toAdd.Count)`n" +
                      "Grupos a remover: $($toRemove.Count)"

        
        if ([System.Windows.Forms.MessageBox]::Show($msgConfirm, "Confirmación de Permisos", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        




        
        try {
            if (-not (Get-Command Add-ADGroupMember -ErrorAction SilentlyContinue)) {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            $addedOk = 0
            $removedOk = 0
            
            foreach ($grp in $toAdd) {
                try {
                    Add-ADGroupMember -Identity $grp -Members $targetUser -ErrorAction SilentlyContinue
                    $addedOk++
                } catch {}
            }
            foreach ($grp in $toRemove) {
                try {
                    Remove-ADGroupMember -Identity $grp -Members $targetUser -Confirm:$false -ErrorAction SilentlyContinue
                    $removedOk++
                } catch {}
            }
            
            [System.Windows.Forms.MessageBox]::Show("Permisos actualizados con éxito para '$targetUser':`n`n- Agregados: $addedOk grupos`n- Removidos: $removedOk grupos", "Permisos Aplicados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al actualizar permisos en Active Directory: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    # ------------------ CASO PESTAÑA 3: CARGA MASIVA MULTI-CARGO ------------------
    elseif ($tabControlLeft.SelectedIndex -eq 2) {
        if ($gridBatch.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No hay colaboradores en la tabla masiva para crear. Carga o procesa una lista primero.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $availableRows = @()
        foreach ($r in $gridBatch.Rows) {
            if ($r.Cells["EstadoDA"].Value -like "*Disponible*") {
                $availableRows += $r
            }
        }
        
        if ($availableRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Todos los colaboradores del lote ya están registrados en Active Directory. No hay cuentas nuevas por crear.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $count = $availableRows.Count
        
        $msgConfirm = "¿Desea iniciar la creación masiva de $count usuarios disponibles en Active Directory?`n`n" +
                      "Nota: Cada cuenta se creará en su respectiva Unidad Organizativa (OU) según la sede de la tabla.`n" +
                      "Ruta Contenedora: USUARIOS (aliar.com.co)"

        
        if ([System.Windows.Forms.MessageBox]::Show($msgConfirm, "Confirmar Creación Masiva Individualizada", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        




        
        try {
            if (-not (Get-Command New-ADUser -ErrorAction SilentlyContinue)) {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            $createdOk = 0
            foreach ($row in $availableRows) {
                $ced = $row.Cells["Cedula"].Value
                $uBaseName = $row.Cells["Nombre"].Value
                $uSam = $row.Cells["Usuario"].Value
                $uCargo = $row.Cells["Cargo"].Value
                $uSede = $row.Cells["Sede"].Value
                $uDept = $row.Cells["Departamento"].Value
                $uProfile = $row.Cells["PerfilWeb"].Value
                $isVpn = [bool]$row.Cells["Vpn"].Value
                $uPass = $row.Cells["Clave"].Value
                
                $uBaseName = To-TitleCase $uBaseName
                $uCargo = To-TitleCase $uCargo
                $rowSedeCode = $uSede
                $foundCfg = $global:SedesConfig | Where-Object { $_.OU -eq $uSede -or $_.Code -eq $uSede -or $_.Label -eq $uSede } | Select-Object -First 1
                if ($foundCfg) { $rowSedeCode = $foundCfg.Code }
                $dNameWithSede = "$uBaseName - $rowSedeCode"
                $targetOU = Get-TargetOUPath $uSede
                
                try {
                    $secPass = ConvertTo-SecureString $uPass -AsPlainText -Force
                    New-ADUser -Name $dNameWithSede `
                               -DisplayName $dNameWithSede `
                               -SamAccountName $uSam `
                               -UserPrincipalName "$uSam@aliar.com.co" `
                               -Department $uDept `
                               -Title $uCargo `
                               -EmployeeID $ced `
                               -Path $targetOU `
                               -AccountPassword $secPass `
                               -Enabled $true `
                               -ChangePasswordAtLogon $true `
                               -PasswordNeverExpires $false
                    
                    if ($uProfile) {
                        for ($retry = 0; $retry -lt 3; $retry++) {
                            try {
                                Add-ADGroupMember -Identity $uProfile -Members $uSam -ErrorAction Stop
                                break
                            } catch { Start-Sleep -Milliseconds 250 }
                        }
                    }
                    if ($isVpn) {
                        $vpnAdded = $false
                        for ($retry = 0; $retry -lt 3 -and (-not $vpnAdded); $retry++) {
                            try {
                                Add-ADGroupMember -Identity "GR_VPN_SSL" -Members $uSam -ErrorAction Stop
                                $vpnAdded = $true
                            } catch {
                                try {
                                    Add-ADGroupMember -Identity "GR_VPN_Acceso" -Members $uSam -ErrorAction Stop
                                    $vpnAdded = $true
                                    break
                                } catch {}
                                Start-Sleep -Milliseconds 250
                            }
                        }
                    }
                    
                    $createdOk++
                    $row.Cells["EstadoDA"].Value = "Creado en DA Exitosamente"
                    $row.Cells["EstadoDA"].Style.ForeColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
                } catch {
                    $row.Cells["EstadoDA"].Value = "Error al Crear"
                    $row.Cells["EstadoDA"].Style.ForeColor = [System.Drawing.Color]::FromArgb(220, 38, 38)
                }
            }
            Copy-BatchCredentialsToClipboard $false
            [System.Windows.Forms.MessageBox]::Show("Creación masiva completada exitosamente:`n`n$createdOk de $count colaboradores creados individualmente en Active Directory con sus respectivas Sedes y Departamentos.`n`n(Las credenciales del lote han sido copiadas automáticamente al portapapeles para entrega a GH).", "Lote Creado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error general en creación masiva: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# ======================================================================================
# 6. RESPONSIVIDAD DINÁMICA, OPTIMIZACIÓN VISUAL Y EVENTOS DE CARGA
# ======================================================================================
$form.add_Resize({
    if ($form.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized) {
        $rightPanelWidth = 430
        $newDistance = $form.ClientSize.Width - $rightPanelWidth
        if ($newDistance -ge 540) {
            $splitMain.SplitterDistance = $newDistance
        }
    }
})

$form.add_Load({
    Enable-DoubleBuffer $form
    Enable-DoubleBuffer $splitMain
    Enable-DoubleBuffer $tblTab1
    Enable-DoubleBuffer $tblTab2
    Enable-DoubleBuffer $tblTab3
    Enable-DoubleBuffer $gridBatch
    
    $rightPanelWidth = 430
    $newDistance = $form.ClientSize.Width - $rightPanelWidth
    if ($newDistance -ge 540) {
        $splitMain.SplitterDistance = $newDistance
    }
    $txtPass.Text = Generate-RandomPassword
    Update-AllOutputs
})

$form.add_Shown({
    # Consulta automática del Dominio en vivo al iniciar
    Sync-ADLive $false
})

# Lanzamiento de la aplicación
[System.Windows.Forms.Application]::Run($form)