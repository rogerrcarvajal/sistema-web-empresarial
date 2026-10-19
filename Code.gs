/**
 * ============================================================================
 * Proyecto: Sistema Web Empresarial — Módulo de Autenticación, Usuarios y RBAC
 * Autor: Rogerr Carvajal <rogerrcarvajal@gmail.com>
 * Rol: Desarrollador de Soluciones con Modelos & Agentes Inteligentes
 * Repositorio: https://github.com/rogerrcarvajal/sistema-web-empresarial
 * Estándar: Clean Architecture / DDD / Zero-Trust Security / Google Apps Script
 * ============================================================================
 */

'use strict';

/**
 * 1. CONFIGURACIÓN CENTRALIZADA DEL SISTEMA
 */
const APP_CONFIG = {
  APP_NAME: 'Sistema Web Empresarial — Core Auth & RBAC',
  APP_VERSION: '1.0.0',
  AUTHOR: 'Rogerr Carvajal (@rogerrcarvajal)',
  SHEETS: {
    USERS: 'Usuarios',
    ROLES: 'Roles',
    PERMISSIONS: 'Permisos',
    ROLE_PERMISSIONS: 'Rol_Permisos',
    SESSIONS: 'Sesiones',
    RESET_TOKENS: 'Recuperacion_Tokens',
    AUDIT_LOGS: 'Auditoria_Seguridad',
    SETTINGS: 'Configuracion'
  },
  SECURITY: {
    SESSION_DURATION_MINUTES: 120,
    RESET_TOKEN_DURATION_MINUTES: 15,
    MAX_LOGIN_ATTEMPTS: 5,
    LOCKOUT_MINUTES: 15,
    PBKDF2_ITERATIONS: 1500,
    CACHE_TTL_SECONDS: 300
  },
  HEADERS: {
    Usuarios: [
      'id_usuario', 'nombre_completo', 'email', 'rol_id', 'password_hash',
      'password_salt', 'estado', 'intentos_fallidos', 'bloqueado_hasta',
      'ultimo_acceso', 'creado_en', 'actualizado_en'
    ],
    Roles: ['id_rol', 'nombre_rol', 'descripcion', 'nivel_jerarquia', 'estado'],
    Permisos: ['id_permiso', 'codigo_permiso', 'modulo', 'descripcion'],
    Rol_Permisos: ['id_asignacion', 'id_rol', 'codigo_permiso', 'asignado_en'],
    Sesiones: [
      'token_hash', 'id_usuario', 'email', 'rol_id', 'ip_context',
      'creado_en', 'expira_en', 'activo'
    ],
    Recuperacion_Tokens: [
      'token_hash', 'id_usuario', 'email', 'creado_en', 'expira_en', 'usado'
    ],
    Auditoria_Seguridad: [
      'id_evento', 'timestamp', 'id_usuario', 'email', 'evento',
      'modulo', 'resultado', 'detalle'
    ],
    Configuracion: ['clave', 'valor', 'descripcion', 'actualizado_en']
  }
};

/**
 * 2. PUNTO DE ENTRADA WEB Y PLANTILLAS HTML
 */
function doGet(e) {
  try {
    ensureDatabaseInitialized_();
  } catch (err) {
    console.error('Aviso de inicialización automática:', err);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Incluye archivos HTML/CSS/JS parciales en la plantilla principal.
 * @param {string} filename Nombre del archivo a incluir.
 * @return {string} Contenido HTML evaluado.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 3. INICIALIZACIÓN Y APROVISIONAMIENTO AUTOMÁTICO DE GOOGLE SHEETS
 * Crea las hojas, cabeceras con formato corporativo, roles, permisos y usuario admin por defecto.
 */
function setupDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Este script debe estar vinculado a un Google Spreadsheet activo o configurar SpreadsheetApp.openById.');
    }

    const sheetNames = Object.values(APP_CONFIG.SHEETS);
    sheetNames.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      const headers = APP_CONFIG.HEADERS[sheetName];
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      if (sheet.getLastRow() === 0 && headers && headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#0b0f19');
        headerRange.setFontColor('#38bdf8');
        headerRange.setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    });

    seedInitialData_(ss);
    clearSystemCache_();

    return {
      success: true,
      message: 'Base de datos en Google Sheets aprovisionada correctamente con esquema RBAC y usuario Administrador inicial.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Verifica que las hojas existan antes de procesar peticiones.
 */
function ensureDatabaseInitialized_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('DB_INITIALIZED_V1') === 'true') {
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const usersSheet = ss.getSheetByName(APP_CONFIG.SHEETS.USERS);
  if (!usersSheet || usersSheet.getLastRow() === 0) {
    setupDatabase();
  }
  cache.put('DB_INITIALIZED_V1', 'true', 21600);
}

/**
 * Inserta datos semilla por lotes (Batch Operations) si las hojas están vacías.
 */
function seedInitialData_(ss) {
  const now = new Date().toISOString();

  // 1. Semilla de Roles
  const rolesSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ROLES);
  if (rolesSheet.getLastRow() <= 1) {
    const defaultRoles = [
      ['ROL_ADMIN', 'Administrador Global', 'Control total del sistema, usuarios, roles y configuración', 100, 'ACTIVO'],
      ['ROL_SUPERVISOR', 'Supervisor de Operaciones', 'Gestión de inventario, ventas, logística, SSOMA y reportes', 50, 'ACTIVO'],
      ['ROL_OPERADOR', 'Operador Especialista', 'Operaciones estándar de módulos asignados y perfil personal', 20, 'ACTIVO'],
      ['ROL_AUDITOR', 'Auditor / Consulta', 'Visualización de tableros, reportes y trazabilidad en modo lectura', 10, 'ACTIVO']
    ];
    rolesSheet.getRange(2, 1, defaultRoles.length, defaultRoles[0].length).setValues(defaultRoles);
  }

  // 2. Semilla de Permisos Empresariales (Multi-Módulo: Core, Inventario, Ventas, Logística, SSOMA)
  const permsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PERMISSIONS);
  if (permsSheet.getLastRow() <= 1) {
    const defaultPerms = [
      ['PERM_01', 'DASHBOARD_VIEW', 'CORE', 'Visualizar panel principal y métricas operativas'],
      ['PERM_02', 'USERS_VIEW', 'USUARIOS', 'Listar y buscar usuarios del sistema'],
      ['PERM_03', 'USERS_MANAGE', 'USUARIOS', 'Crear, editar, bloquear y cambiar roles de usuarios'],
      ['PERM_04', 'ROLES_MANAGE', 'SEGURIDAD', 'Configurar matriz de roles, permisos y políticas del sistema'],
      ['PERM_05', 'AUDIT_VIEW', 'SEGURIDAD', 'Consultar bitácora de auditoría Zero-Trust'],
      ['PERM_06', 'INVENTARIO_ACCESS', 'INVENTARIO', 'Acceso al módulo de control de stock y almacenes'],
      ['PERM_07', 'VENTAS_ACCESS', 'VENTAS', 'Acceso al módulo comercial, cotizaciones y facturación'],
      ['PERM_08', 'LOGISTICA_ACCESS', 'LOGISTICA', 'Acceso al módulo de despacho, flota y cadena de suministro'],
      ['PERM_09', 'SSOMA_ACCESS', 'SSOMA', 'Acceso a Seguridad, Salud Ocupacional y Medio Ambiente']
    ];
    permsSheet.getRange(2, 1, defaultPerms.length, defaultPerms[0].length).setValues(defaultPerms);
  }

  // 3. Semilla de Matriz Rol_Permisos
  const rpSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ROLE_PERMISSIONS);
  if (rpSheet.getLastRow() <= 1) {
    const rolePerms = [
      // Admin tiene todos los permisos
      ['RP_01', 'ROL_ADMIN', 'DASHBOARD_VIEW', now],
      ['RP_02', 'ROL_ADMIN', 'USERS_VIEW', now],
      ['RP_03', 'ROL_ADMIN', 'USERS_MANAGE', now],
      ['RP_04', 'ROL_ADMIN', 'ROLES_MANAGE', now],
      ['RP_05', 'ROL_ADMIN', 'AUDIT_VIEW', now],
      ['RP_06', 'ROL_ADMIN', 'INVENTARIO_ACCESS', now],
      ['RP_07', 'ROL_ADMIN', 'VENTAS_ACCESS', now],
      ['RP_08', 'ROL_ADMIN', 'LOGISTICA_ACCESS', now],
      ['RP_09', 'ROL_ADMIN', 'SSOMA_ACCESS', now],
      // Supervisor
      ['RP_10', 'ROL_SUPERVISOR', 'DASHBOARD_VIEW', now],
      ['RP_11', 'ROL_SUPERVISOR', 'USERS_VIEW', now],
      ['RP_12', 'ROL_SUPERVISOR', 'INVENTARIO_ACCESS', now],
      ['RP_13', 'ROL_SUPERVISOR', 'VENTAS_ACCESS', now],
      ['RP_14', 'ROL_SUPERVISOR', 'LOGISTICA_ACCESS', now],
      ['RP_15', 'ROL_SUPERVISOR', 'SSOMA_ACCESS', now],
      // Operador
      ['RP_16', 'ROL_OPERADOR', 'DASHBOARD_VIEW', now],
      ['RP_17', 'ROL_OPERADOR', 'INVENTARIO_ACCESS', now],
      ['RP_18', 'ROL_OPERADOR', 'VENTAS_ACCESS', now],
      ['RP_19', 'ROL_OPERADOR', 'LOGISTICA_ACCESS', now],
      // Auditor
      ['RP_20', 'ROL_AUDITOR', 'DASHBOARD_VIEW', now],
      ['RP_21', 'ROL_AUDITOR', 'AUDIT_VIEW', now],
      ['RP_22', 'ROL_AUDITOR', 'SSOMA_ACCESS', now]
    ];
    rpSheet.getRange(2, 1, rolePerms.length, rolePerms[0].length).setValues(rolePerms);
  }

  // 4. Semilla de Configuración Global
  const settingsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    const defaultSettings = [
      ['ALLOW_PUBLIC_REGISTRATION', 'true', 'Permite que nuevos usuarios se registren desde la pantalla de login', now],
      ['DEFAULT_REGISTRATION_ROLE', 'ROL_OPERADOR', 'Rol asignado por defecto al auto-registrarse', now],
      ['SESSION_TTL_MINUTES', '120', 'Tiempo de expiración de sesión activa en minutos', now],
      ['MAX_LOGIN_ATTEMPTS', '5', 'Intentos fallidos antes de bloqueo temporal de cuenta', now],
      ['COMPANY_MODULE_TITLE', 'Enterprise Core Suite — Inventario, Ventas, Logística & SSOMA', 'Subtítulo corporativo del sistema', now]
    ];
    settingsSheet.getRange(2, 1, defaultSettings.length, defaultSettings[0].length).setValues(defaultSettings);
  }

  // 5. Semilla de Usuarios Iniciales (Contraseña no en texto plano: Hash + Salt único)
  const usersSheet = ss.getSheetByName(APP_CONFIG.SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    const adminSalt = generateSalt_();
    const adminHash = hashPasswordWithSalt_('Admin2026!', adminSalt);

    const supSalt = generateSalt_();
    const supHash = hashPasswordWithSalt_('Supervisor2026!', supSalt);

    const defaultUsers = [
      [
        'USR_0001',
        'Rogerr Carvajal (Admin)',
        'admin@empresa.com',
        'ROL_ADMIN',
        adminHash,
        adminSalt,
        'ACTIVO',
        0,
        '',
        now,
        now,
        now
      ],
      [
        'USR_0002',
        'Supervisor Operaciones SSOMA',
        'supervisor@empresa.com',
        'ROL_SUPERVISOR',
        supHash,
        supSalt,
        'ACTIVO',
        0,
        '',
        now,
        now,
        now
      ]
    ];
    usersSheet.getRange(2, 1, defaultUsers.length, defaultUsers[0].length).setValues(defaultUsers);
  }
}

/**
 * 4. MOTOR CRIPTOGRÁFICO ZERO-TRUST (SHA-256 ITERADO + SALT ALEATORIO + TIEMPO CONSTANTE)
 */
function generateSalt_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function sha256Hex_(input) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );
  let hex = '';
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    const byteStr = byteVal.toString(16);
    hex += (byteStr.length === 1 ? '0' : '') + byteStr;
  }
  return hex;
}

/**
 * Deriva un hash robusto aplicando múltiples rondas de SHA-256 con salt por usuario y pepper de script.
 */
function hashPasswordWithSalt_(plainPassword, salt) {
  const pepper = 'RC_ENTERPRISE_ZERO_TRUST_PEPPER_2026';
  let digest = sha256Hex_(salt + ':' + String(plainPassword) + ':' + pepper);
  const rounds = APP_CONFIG.SECURITY.PBKDF2_ITERATIONS;
  for (let i = 0; i < rounds; i++) {
    digest = sha256Hex_(digest + ':' + salt + ':' + i);
  }
  return digest;
}

/**
 * Comparación en tiempo constante para mitigar ataques de temporización (Timing Attacks).
 */
function constantTimeEquals_(a, b) {
  const strA = String(a || '');
  const strB = String(b || '');
  if (strA.length !== strB.length) return false;
  let diff = 0;
  for (let i = 0; i < strA.length; i++) {
    diff |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
  }
  return diff === 0;
}

/**
 * 5. CAPA DE ACCESO A DATOS POR LOTES (BATCH I/O) Y CACHÉ MULTICAPA
 */
function getSheetDataAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sheet: null, headers: [], rows: [], rawValues: [] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheet: sheet, headers: [], rows: [], rawValues: [] };
  }

  // Lectura masiva en una sola llamada (Evita lectura celda por celda)
  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allValues[0].map(h => String(h).trim());
  const rows = [];

  for (let r = 1; r < allValues.length; r++) {
    const rowData = allValues[r];
    // Ignorar filas completamente vacías
    if (rowData.every(cell => cell === '' || cell === null)) continue;
    const obj = { _rowNumber: r + 1 };
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]] = rowData[c];
    }
    rows.push(obj);
  }

  return { sheet: sheet, headers: headers, rows: rows, rawValues: allValues };
}

/**
 * Actualiza una fila completa en una sola operación por lotes.
 */
function updateRowByObject_(sheet, headers, rowNumber, updatedObj) {
  const rowArray = headers.map(h => (updatedObj[h] !== undefined ? updatedObj[h] : ''));
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowArray]);
}

/**
 * Inserta una fila completa al final de la hoja en una sola operación.
 */
function appendRowByObject_(sheet, headers, newObj) {
  const rowArray = headers.map(h => (newObj[h] !== undefined ? newObj[h] : ''));
  sheet.appendRow(rowArray);
}

/**
 * Obtiene configuraciones, roles y permisos desde CacheService para minimizar lecturas a Sheets.
 */
function getCachedCatalogs_() {
  const cache = CacheService.getScriptCache();
  const cachedJson = cache.get('CATALOGS_CACHE_V1');
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (e) {
      // Continuar a recarga desde Sheets
    }
  }

  const rolesData = getSheetDataAsObjects_(APP_CONFIG.SHEETS.ROLES).rows;
  const permsData = getSheetDataAsObjects_(APP_CONFIG.SHEETS.PERMISSIONS).rows;
  const rpData = getSheetDataAsObjects_(APP_CONFIG.SHEETS.ROLE_PERMISSIONS).rows;
  const settingsRows = getSheetDataAsObjects_(APP_CONFIG.SHEETS.SETTINGS).rows;

  const settings = {};
  settingsRows.forEach(item => {
    settings[String(item.clave).trim()] = String(item.valor).trim();
  });

  const rolePermissionsMap = {};
  rolesData.forEach(r => {
    rolePermissionsMap[r.id_rol] = [];
  });
  rpData.forEach(rp => {
    if (!rolePermissionsMap[rp.id_rol]) {
      rolePermissionsMap[rp.id_rol] = [];
    }
    if (rolePermissionsMap[rp.id_rol].indexOf(rp.codigo_permiso) === -1) {
      rolePermissionsMap[rp.id_rol].push(rp.codigo_permiso);
    }
  });

  const cleanRoles = rolesData.map(r => ({
    id_rol: String(r.id_rol),
    nombre_rol: String(r.nombre_rol),
    descripcion: String(r.descripcion),
    nivel_jerarquia: Number(r.nivel_jerarquia) || 10,
    estado: String(r.estado),
    permisos: rolePermissionsMap[r.id_rol] || []
  }));

  const cleanPerms = permsData.map(p => ({
    id_permiso: String(p.id_permiso),
    codigo_permiso: String(p.codigo_permiso),
    modulo: String(p.modulo),
    descripcion: String(p.descripcion)
  }));

  const payload = {
    roles: cleanRoles,
    permissions: cleanPerms,
    rolePermissionsMap: rolePermissionsMap,
    settings: settings
  };

  cache.put('CATALOGS_CACHE_V1', JSON.stringify(payload), APP_CONFIG.SECURITY.CACHE_TTL_SECONDS);
  return payload;
}

function clearSystemCache_() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['CATALOGS_CACHE_V1', 'DB_INITIALIZED_V1']);
}

/**
 * Registra eventos de seguridad en la hoja de Auditoría.
 */
function logSecurityEvent_(idUsuario, email, evento, modulo, resultado, detalle) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(APP_CONFIG.SHEETS.AUDIT_LOGS);
    if (!auditSheet) return;
    const idEvento = 'EVT_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
    auditSheet.appendRow([
      idEvento,
      new Date().toISOString(),
      idUsuario || 'ANONIMO',
      email || 'N/A',
      evento,
      modulo,
      resultado,
      detalle
    ]);
  } catch (e) {
    console.error('Error al escribir log de auditoría:', e);
  }
}

/**
 * Sanitiza un objeto de usuario para NUNCA exponer hashes, salts ni metadatos sensibles al frontend.
 */
function sanitizeUserForClient_(userRow, catalogs) {
  const roleId = String(userRow.rol_id);
  const roleObj = (catalogs.roles || []).find(r => r.id_rol === roleId);
  const permissions = (catalogs.rolePermissionsMap && catalogs.rolePermissionsMap[roleId]) || [];

  return {
    id_usuario: String(userRow.id_usuario),
    nombre_completo: String(userRow.nombre_completo),
    email: String(userRow.email).toLowerCase(),
    rol_id: roleId,
    nombre_rol: roleObj ? roleObj.nombre_rol : roleId,
    estado: String(userRow.estado),
    ultimo_acceso: String(userRow.ultimo_acceso || ''),
    creado_en: String(userRow.creado_en || ''),
    permisos: permissions
  };
}

/**
 * 6. VALIDACIÓN DE ENTRADAS Y LIMITACIÓN DE TASA (RATE LIMITING)
 */
function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validatePasswordStrength_(password) {
  const pwd = String(password || '');
  if (pwd.length < 8) {
    return { valid: false, message: 'La contraseña debe tener al menos 8 caracteres.' };
  }
  if (!/[A-Z]/.test(pwd) || !/[a-z]/.test(pwd) || !/[0-9]/.test(pwd)) {
    return { valid: false, message: 'La contraseña debe incluir mayúsculas, minúsculas y números.' };
  }
  return { valid: true };
}

function checkRateLimit_(actionKey, maxRequests, windowSeconds) {
  const cache = CacheService.getScriptCache();
  const key = 'RL_' + sha256Hex_(actionKey).substring(0, 32);
  const current = Number(cache.get(key) || '0');
  if (current >= maxRequests) {
    return false;
  }
  cache.put(key, String(current + 1), windowSeconds);
  return true;
}

/**
 * 7. GESTIÓN DE SESIONES Y AUTORIZACIÓN EN SERVIDOR
 */
function createSessionForUser_(userRow) {
  const rawToken = Utilities.getUuid() + '-' + Utilities.getUuid() + '-' + new Date().getTime();
  const tokenHash = sha256Hex_(rawToken);
  const catalogs = getCachedCatalogs_();
  const ttlMinutes = Number(catalogs.settings.SESSION_TTL_MINUTES) || APP_CONFIG.SECURITY.SESSION_DURATION_MINUTES;

  const now = new Date();
  const expiresAt = new Date(now.getTime() + ttlMinutes * 60000);

  const sessionObj = {
    token_hash: tokenHash,
    id_usuario: String(userRow.id_usuario),
    email: String(userRow.email).toLowerCase(),
    rol_id: String(userRow.rol_id),
    ip_context: 'GAS_WEBAPP_CLIENT',
    creado_en: now.toISOString(),
    expira_en: expiresAt.toISOString(),
    activo: 'TRUE'
  };

  const sessionsTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.SESSIONS);
  appendRowByObject_(sessionsTable.sheet, sessionsTable.headers, sessionObj);

  // Guardar sesión rápida en CacheService por 15 min para acelerar validaciones subsecuentes
  const cache = CacheService.getScriptCache();
  cache.put('SESS_' + tokenHash, JSON.stringify(sessionObj), Math.min(ttlMinutes * 60, 900));

  return {
    sessionToken: rawToken,
    expiresAt: expiresAt.toISOString()
  };
}

/**
 * Verifica que el token de sesión sea válido, no haya expirado y que el usuario siga ACTIVO.
 * Si se especifica requiredPermission, verifica autorización RBAC en el servidor.
 */
function requireAuthAndPermission_(rawToken, requiredPermission) {
  if (!rawToken || typeof rawToken !== 'string') {
    throw new Error('NO_AUTH: Sesión no proporcionada o inválida.');
  }

  const tokenHash = sha256Hex_(rawToken.trim());
  const nowIso = new Date().toISOString();
  const sessionsTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.SESSIONS);

  const sessionRow = sessionsTable.rows.find(
    s => constantTimeEquals_(String(s.token_hash), tokenHash) && String(s.activo).toUpperCase() === 'TRUE'
  );

  if (!sessionRow) {
    throw new Error('NO_AUTH: La sesión no existe o fue cerrada.');
  }

  if (String(sessionRow.expira_en) < nowIso) {
    sessionRow.activo = 'FALSE';
    updateRowByObject_(sessionsTable.sheet, sessionsTable.headers, sessionRow._rowNumber, sessionRow);
    CacheService.getScriptCache().remove('SESS_' + tokenHash);
    throw new Error('SESSION_EXPIRED: Tu sesión ha expirado por seguridad. Inicia sesión nuevamente.');
  }

  const usersTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.USERS);
  const userRow = usersTable.rows.find(u => String(u.id_usuario) === String(sessionRow.id_usuario));

  if (!userRow || String(userRow.estado).toUpperCase() !== 'ACTIVO') {
    throw new Error('FORBIDDEN: La cuenta de usuario se encuentra inactiva o suspendida.');
  }

  const catalogs = getCachedCatalogs_();
  const userPermissions = catalogs.rolePermissionsMap[String(userRow.rol_id)] || [];

  if (requiredPermission && userPermissions.indexOf(requiredPermission) === -1) {
    logSecurityEvent_(
      userRow.id_usuario,
      userRow.email,
      'ACCESO_DENEGADO_RBAC',
      requiredPermission,
      'BLOQUEADO',
      'Intento de ejecutar operación sin el permiso requerido: ' + requiredPermission
    );
    throw new Error('FORBIDDEN: No tienes permisos suficientes (' + requiredPermission + ') para realizar esta operación.');
  }

  return {
    userRow: userRow,
    usersTable: usersTable,
    sessionRow: sessionRow,
    sessionsTable: sessionsTable,
    catalogs: catalogs,
    userClean: sanitizeUserForClient_(userRow, catalogs)
  };
}

/**
 * 8. ENDPOINTS PÚBLICOS Y PROTEGIDOS (API RPC PARA FRONTEND)
 */

/**
 * Obtiene el estado inicial de la aplicación y valida una sesión existente si se envía un token.
 */
function apiGetBootstrapState(sessionToken) {
  try {
    ensureDatabaseInitialized_();
    const catalogs = getCachedCatalogs_();
    const publicConfig = {
      appName: APP_CONFIG.APP_NAME,
      appVersion: APP_CONFIG.APP_VERSION,
      companyTitle: catalogs.settings.COMPANY_MODULE_TITLE || 'Enterprise Suite',
      allowPublicRegistration: String(catalogs.settings.ALLOW_PUBLIC_REGISTRATION).toLowerCase() === 'true'
    };

    if (!sessionToken) {
      return {
        success: true,
        authenticated: false,
        config: publicConfig
      };
    }

    const authContext = requireAuthAndPermission_(sessionToken, null);
    return {
      success: true,
      authenticated: true,
      config: publicConfig,
      user: authContext.userClean,
      catalogs: {
        roles: catalogs.roles,
        permissions: catalogs.permissions
      }
    };
  } catch (err) {
    const catalogs = getCachedCatalogs_();
    return {
      success: true,
      authenticated: false,
      authError: err.message,
      config: {
        appName: APP_CONFIG.APP_NAME,
        appVersion: APP_CONFIG.APP_VERSION,
        companyTitle: catalogs.settings.COMPANY_MODULE_TITLE || 'Enterprise Suite',
        allowPublicRegistration: String(catalogs.settings.ALLOW_PUBLIC_REGISTRATION).toLowerCase() === 'true'
      }
    };
  }
}

/**
 * Autenticación de usuario con protección contra fuerza bruta y bloqueo temporal.
 */
function apiLogin(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureDatabaseInitialized_();
    const email = normalizeEmail_(payload && payload.email);
    const password = String((payload && payload.password) || '');

    if (!email || !password || !isValidEmail_(email)) {
      return { success: false, message: 'Credenciales inválidas. Verifica tu correo y contraseña.' };
    }

    if (!checkRateLimit_('LOGIN_' + email, 10, 60)) {
      return {
        success: false,
        message: 'Demasiadas solicitudes seguidas. Espera un minuto antes de volver a intentar.'
      };
    }

    const usersTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.USERS);
    const userRow = usersTable.rows.find(u => normalizeEmail_(u.email) === email);
    const catalogs = getCachedCatalogs_();

    // Mensaje genérico para no revelar si un correo existe o no
    const genericErrorMsg = 'Credenciales inválidas o cuenta no autorizada.';

    if (!userRow) {
      logSecurityEvent_('DESCONOCIDO', email, 'LOGIN_FALLIDO', 'AUTH', 'ERROR', ' Intento con email no registrado');
      return { success: false, message: genericErrorMsg };
    }

    if (String(userRow.estado).toUpperCase() !== 'ACTIVO') {
      logSecurityEvent_(userRow.id_usuario, email, 'LOGIN_CUENTA_INACTIVA', 'AUTH', 'BLOQUEADO', 'Cuenta en estado ' + userRow.estado);
      return { success: false, message: genericErrorMsg };
    }

    const now = new Date();
    const nowIso = now.toISOString();

    if (userRow.bloqueado_hasta && String(userRow.bloqueado_hasta) > nowIso) {
      logSecurityEvent_(userRow.id_usuario, email, 'LOGIN_CUENTA_BLOQUEADA', 'AUTH', 'BLOQUEADO', 'Cuenta bloqueada temporalmente hasta ' + userRow.bloqueado_hasta);
      return {
        success: false,
        message: 'Cuenta bloqueada temporalmente por múltiples intentos fallidos. Intenta más tarde o recupera tu contraseña.'
      };
    }

    const computedHash = hashPasswordWithSalt_(password, String(userRow.password_salt));
    const passwordMatches = constantTimeEquals_(computedHash, String(userRow.password_hash));

    if (!passwordMatches) {
      const maxAttempts = Number(catalogs.settings.MAX_LOGIN_ATTEMPTS) || APP_CONFIG.SECURITY.MAX_LOGIN_ATTEMPTS;
      const attempts = (Number(userRow.intentos_fallidos) || 0) + 1;
      userRow.intentos_fallidos = attempts;
      userRow.actualizado_en = nowIso;

      if (attempts >= maxAttempts) {
        const lockUntil = new Date(now.getTime() + APP_CONFIG.SECURITY.LOCKOUT_MINUTES * 60000);
        userRow.bloqueado_hasta = lockUntil.toISOString();
      }

      updateRowByObject_(usersTable.sheet, usersTable.headers, userRow._rowNumber, userRow);
      logSecurityEvent_(userRow.id_usuario, email, 'LOGIN_FALLIDO', 'AUTH', 'ERROR', 'Contraseña incorrecta. Intento #' + attempts);
      return { success: false, message: genericErrorMsg };
    }

    // Éxito: reiniciar contador de intentos fallidos y registrar último acceso
    userRow.intentos_fallidos = 0;
    userRow.bloqueado_hasta = '';
    userRow.ultimo_acceso = nowIso;
    userRow.actualizado_en = nowIso;
    updateRowByObject_(usersTable.sheet, usersTable.headers, userRow._rowNumber, userRow);

    const sessionInfo = createSessionForUser_(userRow);
    logSecurityEvent_(userRow.id_usuario, email, 'LOGIN_EXITOSO', 'AUTH', 'OK', 'Inicio de sesión correcto');

    return {
      success: true,
      message: 'Bienvenido, ' + userRow.nombre_completo,
      sessionToken: sessionInfo.sessionToken,
      expiresAt: sessionInfo.expiresAt,
      user: sanitizeUserForClient_(userRow, catalogs),
      catalogs: {
        roles: catalogs.roles,
        permissions: catalogs.permissions
      }
    };
  } catch (err) {
    return { success: false, message: 'Error procesando el inicio de sesión: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cierre de sesión seguro invalidando el token en servidor y caché.
 */
function apiLogout(sessionToken) {
  try {
    if (!sessionToken) return { success: true };
    const tokenHash = sha256Hex_(String(sessionToken).trim());
    const sessionsTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.SESSIONS);
    const sessionRow = sessionsTable.rows.find(s => constantTimeEquals_(String(s.token_hash), tokenHash));

    if (sessionRow) {
      sessionRow.activo = 'FALSE';
      updateRowByObject_(sessionsTable.sheet, sessionsTable.headers, sessionRow._rowNumber, sessionRow);
      logSecurityEvent_(sessionRow.id_usuario, sessionRow.email, 'LOGOUT', 'AUTH', 'OK', 'Sesión cerrada por el usuario');
    }

    CacheService.getScriptCache().remove('SESS_' + tokenHash);
    return { success: true, message: 'Sesión cerrada correctamente.' };
  } catch (err) {
    return { success: false, message: 'Error al cerrar sesión.' };
  }
}

/**
 * Registro público de usuarios (solo si el Administrador lo habilitó en Configuración).
 * No revela si un correo ya está registrado para proteger la privacidad de las cuentas.
 */
function apiRegisterPublicUser(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureDatabaseInitialized_();
    const catalogs = getCachedCatalogs_();
    const allowRegistration = String(catalogs.settings.ALLOW_PUBLIC_REGISTRATION).toLowerCase() === 'true';

    if (!allowRegistration) {
      return {
        success: false,
        message: 'El auto-registro público está deshabilitado por el Administrador del sistema.'
      };
    }

    const nombre = String((payload && payload.nombre_completo) || '').trim();
    const email = normalizeEmail_(payload && payload.email);
    const password = String((payload && payload.password) || '');

    if (nombre.length < 3) {
      return { success: false, message: 'Ingresa tu nombre completo válido (mínimo 3 caracteres).' };
    }
    if (!isValidEmail_(email)) {
      return { success: false, message: 'Ingresa un correo electrónico corporativo válido.' };
    }

    const pwdValidation = validatePasswordStrength_(password);
    if (!pwdValidation.valid) {
      return { success: false, message: pwdValidation.message };
    }

    if (!checkRateLimit_('REGISTER_' + email, 3, 300)) {
      return { success: false, message: 'Demasiadas solicitudes de registro. Intenta más tarde.' };
    }

    const usersTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.USERS);
    const existing = usersTable.rows.find(u => normalizeEmail_(u.email) === email);

    // Respuesta uniforme que no revela si el correo ya existe en la base de datos
    const neutralSuccessMsg = 'Solicitud procesada. Si el correo es elegible, tu cuenta ha sido activada y ya puedes iniciar sesión.';

    if (existing) {
      logSecurityEvent_(existing.id_usuario, email, 'REGISTRO_DUPLICADO_OCULTO', 'AUTH', 'AVISO', 'Intento de registro sobre correo ya existente');
      return { success: true, message: neutralSuccessMsg };
    }

    const salt = generateSalt_();
    const passwordHash = hashPasswordWithSalt_(password, salt);
    const nowIso = new Date().toISOString();
    const defaultRole = catalogs.settings.DEFAULT_REGISTRATION_ROLE || 'ROL_OPERADOR';
    const newId = 'USR_' + String(usersTable.rows.length + 1).padStart(4, '0') + '_' + Math.floor(Math.random() * 90 + 10);

    const newUser = {
      id_usuario: newId,
      nombre_completo: nombre,
      email: email,
      rol_id: defaultRole,
      password_hash: passwordHash,
      password_salt: salt,
      estado: 'ACTIVO',
      intentos_fallidos: 0,
      bloqueado_hasta: '',
      ultimo_acceso: '',
      creado_en: nowIso,
      actualizado_en: nowIso
    };

    appendRowByObject_(usersTable.sheet, usersTable.headers, newUser);
    logSecurityEvent_(newId, email, 'REGISTRO_PUBLICO_EXITOSO', 'AUTH', 'OK', 'Usuario registrado con rol ' + defaultRole);

    return { success: true, message: neutralSuccessMsg };
  } catch (err) {
    return { success: false, message: 'No fue posible procesar el registro en este momento.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Solicitud de recuperación de contraseña por correo.
 * Genera un código temporal de un solo uso (hasheado en BD) y NUNCA revela si el correo existe.
 */
function apiRequestPasswordReset(emailInput) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureDatabaseInitialized_();
    const email = normalizeEmail_(emailInput);
    const neutralMsg = 'Si el correo está asociado a una cuenta activa, recibirás un código de seguridad con vigencia de 15 minutos.';

    if (!isValidEmail_(email)) {
      return { success: false, message: 'Formato de correo electrónico inválido.' };
    }

    if (!checkRateLimit_('RESET_REQ_' + email, 3, 300)) {
      return { success: true, message: neutralMsg };
    }

    const usersTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.USERS);
    const userRow = usersTable.rows.find(
      u => normalizeEmail_(u.email) === email && String(u.estado).toUpperCase() === 'ACTIVO'
    );

    if (!userRow) {
      logSecurityEvent_('DESCONOCIDO', email, 'RECUPERACION_CORREO_NO_EXISTE', 'AUTH', 'AVISO', 'Solicitud de recuperación para correo inexistente o inactivo');
      return { success: true, message: neutralMsg };
    }

    // Generar código OTP criptográfico de 8 caracteres alfanuméricos
    const rawCode = Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
    const tokenHash = sha256Hex_(email + ':' + rawCode);
    const now = new Date();
    const expiresAt = new Date(now.getTime() + APP_CONFIG.SECURITY.RESET_TOKEN_DURATION_MINUTES * 60000);

    const tokensTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.RESET_TOKENS);
    appendRowByObject_(tokensTable.sheet, tokensTable.headers, {
      token_hash: tokenHash,
      id_usuario: userRow.id_usuario,
      email: email,
      creado_en: now.toISOString(),
      expira_en: expiresAt.toISOString(),
      usado: 'FALSE'
    });

    // Envío de correo mediante MailApp de Google Apps Script
    try {
      MailApp.sendEmail({
        to: email,
        subject: '[' + APP_CONFIG.APP_NAME + '] Código de Recuperación de Contraseña',
        htmlBody:
          '<div style="font-family: Inter, sans-serif; background:#0b0f19; color:#f8fafc; padding:24px; border-radius:12px;">' +
          '<h2 style="color:#38bdf8; margin-top:0;">Recuperación de Acceso Empresarial</h2>' +
          '<p>Hola <strong>' + userRow.nombre_completo + '</strong>,</p>' +
          '<p>Se ha solicitado restablecer tu contraseña. Utiliza el siguiente código de verificación de un solo uso:</p>' +
          '<div style="background:#1e293b; border:1px solid #38bdf8; padding:16px; font-size:22px; font-weight:bold; letter-spacing:4px; text-align:center; border-radius:8px; color:#38bdf8;">' +
          rawCode +
          '</div>' +
          '<p style="color:#94a3b8; font-size:13px; margin-top:16px;">Este código expira en 15 minutos. Si no solicitaste este cambio, ignora este mensaje.</p>' +
          '</div>'
      });
    } catch (mailErr) {
      console.warn('MailApp no pudo enviar correo en este entorno:', mailErr);
    }

    logSecurityEvent_(userRow.id_usuario, email, 'TOKEN_RECUPERACION_EMITIDO', 'AUTH', 'OK', 'Token temporal generado y enviado por correo');
    return { success: true, message: neutralMsg };
  } catch (err) {
    return { success: false, message: 'No fue posible procesar la solicitud.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Confirma el restablecimiento de contraseña validando el hash del token temporal.
 */
function apiConfirmPasswordReset(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const email = normalizeEmail_(payload && payload.email);
    const code = String((payload && payload.code) || '').trim().toUpperCase();
    const newPassword = String((payload && payload.newPassword) || '');

    if (!isValidEmail_(email) || !code || code.length < 6) {
      return { success: false, message: 'Datos de verificación inválidos.' };
    }

    const pwdCheck = validatePasswordStrength_(newPassword);
    if (!pwdCheck.valid) {
      return { success: false, message: pwdCheck.message };
    }

    if (!checkRateLimit_('RESET_CONFIRM_' + email, 5, 300)) {
      return { success: false, message: 'Demasiados intentos de verificación. Espera unos minutos.' };
    }

    const expectedHash = sha256Hex_(email + ':' + code);
    const nowIso = new Date().toISOString();
    const tokensTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.RESET_TOKENS);

    const tokenRow = tokensTable.rows.find(
      t =>
        normalizeEmail_(t.email) === email &&
        constantTimeEquals_(String(t.token_hash), expectedHash) &&
        String(t.usado).toUpperCase() === 'FALSE' &&
        String(t.expira_en) > nowIso
    );

    if (!tokenRow) {
      logSecurityEvent_('DESCONOCIDO', email, 'RECUPERACION_CODIGO_INVALIDO', 'AUTH', 'ERROR', 'Código inválido o expirado');
      return { success: false, message: 'El código de verificación es inválido o ha expirado.' };
    }

    const usersTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.USERS);
    const userRow = usersTable.rows.find(u => String(u.id_usuario) === String(tokenRow.id_usuario));
    if (!userRow) {
      return { success: false, message: 'No se pudo completar el restablecimiento.' };
    }

    // Marcar token como usado
    tokenRow.usado = 'TRUE';
    updateRowByObject_(tokensTable.sheet, tokensTable.headers, tokenRow._rowNumber, tokenRow);

    // Actualizar contraseña con nuevo Salt y desbloquear cuenta
    const newSalt = generateSalt_();
    userRow.password_salt = newSalt;
    userRow.password_hash = hashPasswordWithSalt_(newPassword, newSalt);
    userRow.intentos_fallidos = 0;
    userRow.bloqueado_hasta = '';
    userRow.actualizado_en = nowIso;
    updateRowByObject_(usersTable.sheet, usersTable.headers, userRow._rowNumber, userRow);

    logSecurityEvent_(userRow.id_usuario, email, 'PASSWORD_RESTABLECIDO', 'AUTH', 'OK', 'Contraseña actualizada vía flujo de recuperación');
    return {
      success: true,
      message: 'Tu contraseña ha sido actualizada exitosamente. Ya puedes iniciar sesión.'
    };
  } catch (err) {
    return { success: false, message: 'Error al restablecer la contraseña.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualización de perfil personal y cambio de contraseña autenticado.
 */
function apiUpdateOwnProfile(sessionToken, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const auth = requireAuthAndPermission_(sessionToken, null);
    const userRow = auth.userRow;
    const nowIso = new Date().toISOString();

    if (payload.nombre_completo) {
      const nuevoNombre = String(payload.nombre_completo).trim();
      if (nuevoNombre.length < 3) {
        return { success: false, message: 'El nombre completo debe tener al menos 3 caracteres.' };
      }
      userRow.nombre_completo = nuevoNombre;
    }

    if (payload.currentPassword && payload.newPassword) {
      const currentHash = hashPasswordWithSalt_(String(payload.currentPassword), String(userRow.password_salt));
      if (!constantTimeEquals_(currentHash, String(userRow.password_hash))) {
        logSecurityEvent_(userRow.id_usuario, userRow.email, 'CAMBIO_PASSWORD_FALLIDO', 'PERFIL', 'ERROR', 'Contraseña actual incorrecta');
        return { success: false, message: 'La contraseña actual ingresada no es correcta.' };
      }
      const pwdCheck = validatePasswordStrength_(payload.newPassword);
      if (!pwdCheck.valid) {
        return { success: false, message: pwdCheck.message };
      }
      const newSalt = generateSalt_();
      userRow.password_salt = newSalt;
      userRow.password_hash = hashPasswordWithSalt_(String(payload.newPassword), newSalt);
      logSecurityEvent_(userRow.id_usuario, userRow.email, 'CAMBIO_PASSWORD_EXITOSO', 'PERFIL', 'OK', 'Usuario actualizó su contraseña');
    }

    userRow.actualizado_en = nowIso;
    updateRowByObject_(auth.usersTable.sheet, auth.usersTable.headers, userRow._rowNumber, userRow);

    return {
      success: true,
      message: 'Perfil actualizado correctamente.',
      user: sanitizeUserForClient_(userRow, auth.catalogs)
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 9. ENDPOINTS ADMINISTRATIVOS PROTEGIDOS POR RBAC (USUARIOS, ROLES, AUDITORÍA Y MÓDULOS)
 */

/**
 * Obtiene el listado de usuarios sanitizados, roles, permisos y métricas operativas.
 */
function apiGetAdminDashboardData(sessionToken) {
  try {
    const auth = requireAuthAndPermission_(sessionToken, 'DASHBOARD_VIEW');
    const canViewUsers = auth.userClean.permisos.indexOf('USERS_VIEW') !== -1;
    const canViewAudit = auth.userClean.permisos.indexOf('AUDIT_VIEW') !== -1;

    const sanitizedUsers = canViewUsers
      ? auth.usersTable.rows.map(u => sanitizeUserForClient_(u, auth.catalogs))
      : [];

    let recentLogs = [];
    if (canViewAudit) {
      const auditData = getSheetDataAsObjects_(APP_CONFIG.SHEETS.AUDIT_LOGS).rows;
      recentLogs = auditData
        .slice(-30)
        .reverse()
        .map(l => ({
          id_evento: String(l.id_evento),
          timestamp: String(l.timestamp),
          id_usuario: String(l.id_usuario),
          email: String(l.email),
          evento: String(l.evento),
          modulo: String(l.modulo),
          resultado: String(l.resultado),
          detalle: String(l.detalle)
        }));
    }

    const totalUsers = auth.usersTable.rows.length;
    const activeUsers = auth.usersTable.rows.filter(u => String(u.estado).toUpperCase() === 'ACTIVO').length;
    const activeSessions = auth.sessionsTable.rows.filter(s => String(s.activo).toUpperCase() === 'TRUE').length;

    return {
      success: true,
      metrics: {
        totalUsers: totalUsers,
        activeUsers: activeUsers,
        activeSessions: activeSessions,
        totalRoles: auth.catalogs.roles.length,
        totalPermissions: auth.catalogs.permissions.length
      },
      users: sanitizedUsers,
      roles: auth.catalogs.roles,
      permissions: auth.catalogs.permissions,
      settings: auth.catalogs.settings,
      auditLogs: recentLogs
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Crea o actualiza un usuario desde el panel de Administración (Requiere USERS_MANAGE).
 */
function apiAdminSaveUser(sessionToken, userPayload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const auth = requireAuthAndPermission_(sessionToken, 'USERS_MANAGE');
    const nowIso = new Date().toISOString();
    const idUsuario = String((userPayload && userPayload.id_usuario) || '').trim();
    const nombre = String((userPayload && userPayload.nombre_completo) || '').trim();
    const email = normalizeEmail_(userPayload && userPayload.email);
    const rolId = String((userPayload && userPayload.rol_id) || 'ROL_OPERADOR').trim();
    const estado = String((userPayload && userPayload.estado) || 'ACTIVO').trim().toUpperCase();
    const plainPassword = String((userPayload && userPayload.password) || '').trim();

    if (nombre.length < 3 || !isValidEmail_(email)) {
      return { success: false, message: 'Nombre completo y correo electrónico válido son obligatorios.' };
    }

    const roleExists = auth.catalogs.roles.some(r => r.id_rol === rolId);
    if (!roleExists) {
      return { success: false, message: 'El rol seleccionado no existe en el esquema RBAC.' };
    }

    if (!idUsuario) {
      // Crear nuevo usuario
      const duplicate = auth.usersTable.rows.find(u => normalizeEmail_(u.email) === email);
      if (duplicate) {
        return { success: false, message: 'Ya existe un usuario registrado con ese correo electrónico.' };
      }
      const pwdCheck = validatePasswordStrength_(plainPassword);
      if (!pwdCheck.valid) {
        return { success: false, message: pwdCheck.message };
      }

      const salt = generateSalt_();
      const hash = hashPasswordWithSalt_(plainPassword, salt);
      const newId = 'USR_' + String(auth.usersTable.rows.length + 1).padStart(4, '0') + '_' + Math.floor(Math.random() * 90 + 10);

      const newUserObj = {
        id_usuario: newId,
        nombre_completo: nombre,
        email: email,
        rol_id: rolId,
        password_hash: hash,
        password_salt: salt,
        estado: estado,
        intentos_fallidos: 0,
        bloqueado_hasta: '',
        ultimo_acceso: '',
        creado_en: nowIso,
        actualizado_en: nowIso
      };

      appendRowByObject_(auth.usersTable.sheet, auth.usersTable.headers, newUserObj);
      logSecurityEvent_(auth.userRow.id_usuario, auth.userRow.email, 'ADMIN_CREAR_USUARIO', 'USUARIOS', 'OK', 'Creado usuario ' + email + ' con rol ' + rolId);

      return {
        success: true,
        message: 'Usuario creado exitosamente.',
        user: sanitizeUserForClient_(newUserObj, auth.catalogs)
      };
    } else {
      // Editar usuario existente
      const targetUser = auth.usersTable.rows.find(u => String(u.id_usuario) === idUsuario);
      if (!targetUser) {
        return { success: false, message: 'Usuario no encontrado.' };
      }

      const emailConflict = auth.usersTable.rows.find(
        u => normalizeEmail_(u.email) === email && String(u.id_usuario) !== idUsuario
      );
      if (emailConflict) {
        return { success: false, message: 'Otro usuario ya utiliza ese correo electrónico.' };
      }

      targetUser.nombre_completo = nombre;
      targetUser.email = email;
      targetUser.rol_id = rolId;
      targetUser.estado = estado;
      targetUser.actualizado_en = nowIso;

      if (estado === 'ACTIVO') {
        targetUser.intentos_fallidos = 0;
        targetUser.bloqueado_hasta = '';
      }

      if (plainPassword) {
        const pwdCheck = validatePasswordStrength_(plainPassword);
        if (!pwdCheck.valid) {
          return { success: false, message: pwdCheck.message };
        }
        const newSalt = generateSalt_();
        targetUser.password_salt = newSalt;
        targetUser.password_hash = hashPasswordWithSalt_(plainPassword, newSalt);
      }

      updateRowByObject_(auth.usersTable.sheet, auth.usersTable.headers, targetUser._rowNumber, targetUser);
      logSecurityEvent_(auth.userRow.id_usuario, auth.userRow.email, 'ADMIN_EDITAR_USUARIO', 'USUARIOS', 'OK', 'Actualizado usuario ' + email);

      return {
        success: true,
        message: 'Usuario actualizado exitosamente.',
        user: sanitizeUserForClient_(targetUser, auth.catalogs)
      };
    }
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza los permisos asignados a un Rol en lote y refresca la caché (Requiere ROLES_MANAGE).
 */
function apiAdminUpdateRolePermissions(sessionToken, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const auth = requireAuthAndPermission_(sessionToken, 'ROLES_MANAGE');
    const targetRoleId = String((payload && payload.id_rol) || '').trim();
    const permissionCodes = Array.isArray(payload && payload.permisos) ? payload.permisos : [];

    if (!targetRoleId) {
      return { success: false, message: 'ID de rol inválido.' };
    }

    const rpTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.ROLE_PERMISSIONS);
    const remainingRows = rpTable.rows.filter(r => String(r.id_rol) !== targetRoleId);
    const nowIso = new Date().toISOString();

    const newMatrixRows = [];
    let counter = 1;

    remainingRows.forEach(r => {
      newMatrixRows.push([
        'RP_' + String(counter++).padStart(2, '0'),
        String(r.id_rol),
        String(r.codigo_permiso),
        String(r.asignado_en || nowIso)
      ]);
    });

    permissionCodes.forEach(code => {
      newMatrixRows.push([
        'RP_' + String(counter++).padStart(2, '0'),
        targetRoleId,
        String(code),
        nowIso
      ]);
    });

    // Reescritura por lotes en la hoja Rol_Permisos
    const sheet = rpTable.sheet;
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, rpTable.headers.length).clearContent();
    }
    if (newMatrixRows.length > 0) {
      sheet.getRange(2, 1, newMatrixRows.length, rpTable.headers.length).setValues(newMatrixRows);
    }

    clearSystemCache_();
    logSecurityEvent_(
      auth.userRow.id_usuario,
      auth.userRow.email,
      'RBAC_ACTUALIZAR_PERMISOS',
      'SEGURIDAD',
      'OK',
      'Rol ' + targetRoleId + ' actualizado con [' + permissionCodes.join(', ') + ']'
    );

    const updatedCatalogs = getCachedCatalogs_();
    return {
      success: true,
      message: 'Matriz de permisos actualizada en lote para el rol ' + targetRoleId,
      roles: updatedCatalogs.roles
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza parámetros globales del sistema (Auto-registro, rol por defecto, expiración de sesión).
 */
function apiAdminUpdateSettings(sessionToken, settingsPayload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const auth = requireAuthAndPermission_(sessionToken, 'ROLES_MANAGE');
    const settingsTable = getSheetDataAsObjects_(APP_CONFIG.SHEETS.SETTINGS);
    const nowIso = new Date().toISOString();

    Object.keys(settingsPayload || {}).forEach(key => {
      const row = settingsTable.rows.find(r => String(r.clave) === key);
      if (row) {
        row.valor = String(settingsPayload[key]);
        row.actualizado_en = nowIso;
        updateRowByObject_(settingsTable.sheet, settingsTable.headers, row._rowNumber, row);
      }
    });

    clearSystemCache_();
    logSecurityEvent_(auth.userRow.id_usuario, auth.userRow.email, 'CONFIG_ACTUALIZADA', 'SEGURIDAD', 'OK', JSON.stringify(settingsPayload));

    const updatedCatalogs = getCachedCatalogs_();
    return {
      success: true,
      message: 'Configuración de seguridad actualizada.',
      settings: updatedCatalogs.settings
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Simulador de ejecución de operación protegida para módulos empresariales (Inventario, Ventas, Logística, SSOMA).
 * Demuestra cómo cualquier módulo futuro valida permisos en el servidor antes de procesar datos.
 */
function apiExecuteProtectedModuleAction(sessionToken, moduleKey) {
  try {
    const permissionMap = {
      INVENTARIO: 'INVENTARIO_ACCESS',
      VENTAS: 'VENTAS_ACCESS',
      LOGISTICA: 'LOGISTICA_ACCESS',
      SSOMA: 'SSOMA_ACCESS'
    };
    const requiredPerm = permissionMap[String(moduleKey).toUpperCase()];
    if (!requiredPerm) {
      return { success: false, message: 'Módulo empresarial no reconocido.' };
    }

    const auth = requireAuthAndPermission_(sessionToken, requiredPerm);
    logSecurityEvent_(
      auth.userRow.id_usuario,
      auth.userRow.email,
      'ACCESO_MODULO_' + String(moduleKey).toUpperCase(),
      String(moduleKey).toUpperCase(),
      'OK',
      'Validación Zero-Trust autorizada para ' + requiredPerm
    );

    return {
      success: true,
      module: moduleKey,
      requiredPermission: requiredPerm,
      authorizedUser: auth.userClean.nombre_completo,
      role: auth.userClean.nombre_rol,
      timestamp: new Date().toISOString(),
      message: 'Operación autorizada en servidor mediante permiso [' + requiredPerm + ']. Listo para conectar lógica de ' + moduleKey + '.'
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
