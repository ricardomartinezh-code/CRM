function doGet(e){
  const action = String(e?.parameter?.action || '').trim();
  if(!action){
    return jsonResponse({ ok: true, message: 'Servicio disponible' }, e);
  }
  if(action === 'login'){
    return jsonResponse({ error: 'Utiliza POST para iniciar sesión.' }, e);
  }
  if(action === 'me'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return jsonResponse({ user: auth.user }, e);
  }
  if(action === 'listUsers'){
    const auth = requireAuth_(e, { requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleListUsers_(e, auth.user);
  }
  if(action === 'listTeams'){
    const auth = requireAuth_(e, { requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleListTeams_(e, auth.user);
  }
  if(action === 'getLeads'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return handleGetLeads_(e, auth.user);
  }
  if(action === 'updateLead'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return handleUpdateLead_(e, auth.user);
  }
  if(action === 'listSheets'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return handleListSheets_(e, auth.user);
  }
  if(action === 'listAsesores'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return handleListAsesores_(e, auth.user);
  }
  if(action === 'diagnostics'){
    const auth = requireAuth_(e, { requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleDiagnostics_(e, auth.user);
  }
  if(action === 'resolveDuplicates'){
    const auth = requireAuth_(e, { requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleResolveDuplicates_(e, auth.user);
  }
  if(action === 'report'){
    const auth = requireAuth_(e, { requireActive: true });
    if(auth.error) return auth.response;
    return handleReport_(e, auth.user);
  }
  return jsonResponse({error:'accion no soportada'}, e);
}

function doPost(e){
  const body = parseJsonBody_(e);
  const action = String(body.action || '').trim();
  if(action === 'login'){
    return handleLogin_(e, body);
  }
  if(!action){
    return jsonResponse({ error: 'Acción requerida' }, e);
  }
  if(action === 'createUser'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleCreateUser_(e, body, auth.user);
  }
  if(action === 'updateUser'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleUpdateUser_(e, body, auth.user);
  }
  if(action === 'deleteUser'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleDeleteUser_(e, body, auth.user);
  }
  if(action === 'createTeam'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleCreateTeam_(e, body, auth.user);
  }
  if(action === 'updateTeam'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleUpdateTeam_(e, body, auth.user);
  }
  if(action === 'deleteTeam'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleDeleteTeam_(e, body, auth.user);
  }
  if(action === 'importLeads'){
    const auth = requireAuth_(e, { body, requireActive: true });
    if(auth.error) return auth.response;
    return handleImportLeads_(e, body, auth.user);
  }
  if(action === 'requestPasswordReset'){
    return handlePasswordResetRequest_(e, body);
  }
  if(action === 'syncActivos'){
    const auth = requireAuth_(e, { body, requireActive: true });
    if(auth.error) return auth.response;
    return handleSyncActivos_(e, body, auth.user);
  }
  if(action === 'repairSheetStructure'){
    const auth = requireAuth_(e, { body, requireActive: true, role: 'admin' });
    if(auth.error) return auth.response;
    return handleRepairSheetStructure_(e, body, auth.user);
  }
  return jsonResponse({error:'accion no soportada'}, e);
}

function jsonResponse(obj, e){
  const cb = e?.parameter?.callback;
  const out = cb ? `${cb}(${JSON.stringify(obj)})` : JSON.stringify(obj);
  return ContentService.createTextOutput(out)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

const USERS_SHEET_NAME = 'Usuarios';
const USER_HEADERS = ['Email','UserID','Nombre','PasswordHash','Salt','Rol','Planteles','Bases','Activo','UltimoIngreso','CreadoEl','ActualizadoEl','TokenVersion'];
const TEAMS_SHEET_NAME = 'Equipos';
const TEAM_HEADERS = ['TeamID','Nombre','Descripcion','Planteles','Bases','Miembros','Activo','CreadoEl','ActualizadoEl'];
const TOKEN_SECRET_PROPERTY = 'AUTH_TOKEN_SECRET';
const TOKEN_EXP_MINUTES = 12 * 60;
const ADMIN_ROLES = new Set(['admin']);
const SPREADSHEET_ID_PROPERTY = 'APP_SPREADSHEET_ID';
const SPREADSHEET_ERROR_MESSAGE = 'No se encontró la hoja de cálculo principal. Configura el ID del libro en la propiedad de script "APP_SPREADSHEET_ID".';
const DIAGNOSTICS_PRIORITY_BASE_PROPERTY = 'DIAGNOSTICS_PRIORITY_BASE';
const DIAGNOSTICS_ALERT_RECIPIENTS_PROPERTY = 'DIAGNOSTICS_ALERT_RECIPIENTS';
const DIAGNOSTICS_ALERT_WEBHOOK_PROPERTY = 'DIAGNOSTICS_ALERT_WEBHOOK';
const DIAGNOSTICS_CONTROL_SHEET_NAME = 'Control Diagnósticos';
const DIAGNOSTICS_CONTROL_HEADERS = [
  'Fecha',
  'Base solicitada',
  'Base analizada',
  'Errores detectados',
  'Advertencias',
  'Filas inválidas',
  'Estado del sistema',
  'Resumen'
];
let ACTIVE_USER = null;

const COLUMNS = ['ID','Nombre','Matricula','Correo','Teléfono','Plantel','Modalidad','Programa','Etapa','Estado','Asesor','Comentario','Asignación','Toque 1','Toque 2','Toque 3','Toque 4','CRM','Resolución','Metadatos'];
const REQUIRED_HEADERS = ['id','nombre','telefono','etapa'];
const REQUIRED_HEADER_LABELS = {
  id: 'ID',
  nombre: 'Nombre',
  telefono: 'Teléfono',
  etapa: 'Etapa'
};
const NEGATIVE_STATES = ['no contesta whatsapp','cuelga','buzon de voz'];
const OPTIONAL_HEADER_NAMES = [
  'Telefono','Telefonos','Telefono 1','Telefono 2','Telefono fijo','Telefono celular','Celular','Movil','Whatsapp','Tel',
  'Telefono normalizado','Telefono aircall','Telefono whatsapp','Whatsapp link','Link whatsapp','Resolucion','Resolución',
  'Fecha','Fecha alta','Fecha de alta','Registrado','Fuente','Origen','Base','Campana','Campaña','Tipificacion','Tipificación',
  'Notas','Observaciones','ID CRM','ID lead','Lead id','Fecha asignacion','Fecha asignación'
];
const KNOWN_HEADER_NAMES = Array.from(new Set([...COLUMNS, ...OPTIONAL_HEADER_NAMES]));
const KNOWN_HEADER_SET = new Set(KNOWN_HEADER_NAMES.map(normalizeHeader_));
const VALID_ETAPAS = ['nuevo','contactado','no contactado','inscrito','descartado'];
const ETAPA_ESTADO_MAP = {
  'nuevo': ['nuevo'],
  'contactado': ['interesado','no interesado','en labor de venta'],
  'no contactado': ['no contesta whatsapp','mensaje no contestado','buzon de voz','cuelga'],
  'inscrito': ['inscrito'],
  'descartado': ['spam','no desea informacion','inscrito en otra escuela','limite intentos de contacto']
};
const VALID_ESTADOS_SET = new Set([].concat(...Object.values(ETAPA_ESTADO_MAP)));
const CANONICAL_ETAPA_LABELS = {
  'nuevo': 'Nuevo',
  'contactado': 'Contactado',
  'no contactado': 'No contactado',
  'inscrito': 'Inscrito',
  'descartado': 'Descartado'
};
const CANONICAL_ESTADO_LABELS = (() => {
  const map = Object.create(null);
  Object.keys(ETAPA_ESTADO_MAP).forEach(key => {
    (ETAPA_ESTADO_MAP[key] || []).forEach(value => {
      const normalized = String(value || '').trim().toLowerCase();
      if(!normalized || map[normalized]) return;
      map[normalized] = formatTitleCase_(normalized);
    });
  });
  return map;
})();

const ROUND_ROBIN_PROPERTY_PREFIX = 'ROUND_ROBIN_INDEX_';

function getSpreadsheet_(){
  let ss = null;
  try{
    ss = SpreadsheetApp.getActive();
  }catch(err){
    ss = null;
  }
  if(ss) return ss;
  let spreadsheetId = '';
  try{
    const props = PropertiesService.getScriptProperties();
    spreadsheetId = props ? props.getProperty(SPREADSHEET_ID_PROPERTY) : '';
  }catch(err){
    spreadsheetId = '';
  }
  if(spreadsheetId){
    try{
      ss = SpreadsheetApp.openById(spreadsheetId);
    }catch(err){
      Logger.log('No se pudo abrir la hoja de cálculo configurada: ' + err);
      ss = null;
    }
  }
  return ss;
}

function sanitizeSheetName_(name){
  return String(name || '').replace(/^\s*\*+\s*/, '').trim();
}

function normalizeState_(value){
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/whasapp/g, 'whatsapp')
    .replace(/\s+/g, ' ')
    .trim();
}

function safeDate_(value){
  if(value instanceof Date){
    return isNaN(value.getTime()) ? new Date() : value;
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? new Date() : d;
}

function formatDateTime_(value){
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const dateObj = safeDate_(value);
  return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd HH:mm');
}

function buildToqueValue_(fechaValue, estado){
  const base = formatDateTime_(fechaValue);
  const stateText = String(estado || '').trim();
  return stateText ? `${base} · ${stateText}` : base;
}

function extractToqueState_(value){
  const str = String(value || '').trim();
  if(!str) return '';
  const seps = [' · ', ' | ', ' - '];
  for(let i=0;i<seps.length;i++){
    const sep = seps[i];
    const idx = str.lastIndexOf(sep);
    if(idx >= 0) return str.substring(idx + sep.length).trim();
  }
  return '';
}

function parsePhoneValue_(value){
  const raw = String(value || '').trim();
  if(!raw) return { raw: '', digits: '', local: '', intl: '', wa: '' };
  const digitsOriginal = raw.replace(/\D/g, '');
  if(!digitsOriginal) return { raw, digits: '', local: '', intl: '', wa: '' };
  let digits = digitsOriginal;
  if(digits.startsWith('0052')) digits = digits.slice(4);
  if(digits.startsWith('521')) digits = digits.slice(3);
  else if(digits.startsWith('52')) digits = digits.slice(2);
  if(digits.length > 10 && (digits.startsWith('044') || digits.startsWith('045'))){
    digits = digits.slice(3);
  }
  if(digits.length > 10 && digits.startsWith('01')){
    digits = digits.slice(2);
  }
  if(digits.length === 11 && digits.startsWith('1')){
    digits = digits.slice(1);
  }
  if(digits.length > 10 && digits.startsWith('0')){
    digits = digits.replace(/^0+/, '');
  }
  let local = digits;
  if(local.length > 10){
    local = local.slice(0, 10);
  }
  if(local.length < 7){
    return { raw, digits: '', local: '', intl: '', wa: '' };
  }
  const intl = local.length === 10 ? `52${local}` : digitsOriginal;
  const wa = local.length === 10 ? `521${local}` : digitsOriginal;
  return { raw, digits: digitsOriginal, local, intl, wa };
}

function buildContactLinks_(telefono, correo, nombre){
  const phone = parsePhoneValue_(telefono);
  const callDigits = phone.intl || phone.digits;
  const waDigits = phone.wa || callDigits;
  const email = String(correo || '').trim();
  const leadName = String(nombre || '').trim();
  const subject = 'Seguimiento UNIDEP';
  const greeting = leadName ? `Hola ${leadName}, ` : 'Hola, ';
  return {
    TelefonoNormalizado: phone.local || '',
    TelefonoAircall: callDigits && callDigits.length >= 7 ? `https://dashboard.aircall.io/call/${callDigits}` : '',
    TelefonoWhatsapp: waDigits && waDigits.length >= 7 ? `https://wa.me/${waDigits}` : '',
    CorreoGmail: email ? `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(email)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(greeting)}` : ''
  };
}

function applyContactLinksToRow_(row, map, telefono, correo, nombre){
  const contact = buildContactLinks_(telefono, correo, nombre);
  setRowValue_(row, map, 'Telefono normalizado', contact.TelefonoNormalizado);
  setRowValue_(row, map, 'TelefonoNormalizado', contact.TelefonoNormalizado);
  setRowValue_(row, map, 'Telefono aircall', contact.TelefonoAircall);
  setRowValue_(row, map, 'TelefonoAircall', contact.TelefonoAircall);
  setRowValue_(row, map, 'Telefono whatsapp', contact.TelefonoWhatsapp);
  setRowValue_(row, map, 'TelefonoWhatsapp', contact.TelefonoWhatsapp);
  setRowValue_(row, map, 'Correo gmail', contact.CorreoGmail);
  setRowValue_(row, map, 'CorreoGmail', contact.CorreoGmail);
}

function normalizeIdentifierKey_(value){
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

function normalizeEmailKey_(value){
  return String(value || '')
    .trim()
    .toLowerCase();
}

function normalizePhoneKey_(value){
  const parsed = parsePhoneValue_(value);
  if(parsed.local) return parsed.local;
  if(parsed.digits) return parsed.digits;
  return String(value || '').replace(/\D+/g, '');
}

function generateLeadId_(){
  return 'IMP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function computeResolutionForLead_(sheetName, etapa, estado){
  const etapaNorm = String(etapa || '').toLowerCase().trim();
  if(!etapaNorm) return '';
  if(etapaNorm === 'descartado') return 'Descartado';
  if(etapaNorm === 'inscrito') return 'Inscrito';
  if(etapaNorm === 'nuevo' || etapaNorm === 'no contactado') return 'Nuevo';
  if(etapaNorm === 'contactado'){
    const base = sanitizeSheetName_(sheetName).toLowerCase();
    if(base.indexOf('plantel') >= 0) return 'Seguimiento';
    if(base.indexOf('regreso') >= 0) return 'Seguimiento';
    if(base.indexOf('recicl') >= 0) return 'Recuperado';
    if(base.indexOf('recuper') >= 0) return 'Rescatado';
    return 'Seguimiento';
  }
  return '';
}

function coerceMetadataValue_(value){
  if(value === undefined || value === null) return '';
  if(typeof value === 'string') return value.trim();
  if(Array.isArray(value)){
    return value
      .map(item => {
        if(item === undefined || item === null) return '';
        if(typeof item === 'string') return item.trim();
        try{
          return JSON.stringify(item);
        }catch(err){
          return String(item);
        }
      })
      .filter(Boolean)
      .join(' | ');
  }
  if(typeof value === 'object'){
    try{
      return JSON.stringify(value);
    }catch(err){
      return String(value);
    }
  }
  return String(value || '').trim();
}

function normalizeHeader_(value){
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

const HEADER_SYNONYMS = {
  id: ['lead id', 'id lead', 'id crm', 'crm id', 'id prospecto', 'folio', 'folio id', 'id alumno', 'id registro'],
  nombre: ['nombre completo', 'nombre del lead', 'nombre y apellidos', 'nombre alumno', 'lead', 'prospecto'],
  telefono: [
    'telefono 1',
    'telefono principal',
    'telefono celular',
    'telefono movil',
    'telefono móvil',
    'telefono contacto',
    'telefono de contacto',
    'numero',
    'numero de contacto',
    'número de contacto',
    'celular',
    'whatsapp',
    'whats app',
    'telefono whatsapp',
    'telefono whatsapp principal'
  ],
  etapa: ['fase', 'fase actual', 'status', 'estatus', 'estatus general', 'pipeline', 'stage'],
  estado: ['subestado', 'sub estado', 'estatus detalle', 'status detalle', 'detalle estado', 'resultado', 'subestatus'],
  asesor: ['ejecutivo', 'ejecutiva', 'consultor', 'consultora', 'agente', 'asesor asignado'],
  metadatos: ['metadato', 'metadata', 'datos extra', 'informacion extra', 'info extra']
};

const NORMALIZED_HEADER_ALIASES = Object.freeze(Object.keys(HEADER_SYNONYMS).reduce((acc, key) => {
  const normalizedKey = normalizeHeader_(key);
  const items = HEADER_SYNONYMS[key] || [];
  const normalizedItems = items
    .map(item => normalizeHeader_(item))
    .filter(Boolean);
  acc[normalizedKey] = Array.from(new Set([normalizedKey, ...normalizedItems]));
  return acc;
}, {}));

function normalizeValue_(value){
  return normalizeHeader_(value);
}

function formatTitleCase_(value){
  return String(value || '')
    .toLowerCase()
    .split(/\s+/)
    .map(part => {
      if(!part) return '';
      if(part === 'crm') return 'CRM';
      if(part === 'id') return 'ID';
      return part.charAt(0).toUpperCase() + part.slice(1);
    })
    .join(' ')
    .trim();
}

function formatEtapaLabel_(normalized){
  const key = String(normalized || '').trim().toLowerCase();
  if(!key) return '';
  return CANONICAL_ETAPA_LABELS[key] || formatTitleCase_(key);
}

function formatEstadoLabel_(normalized){
  const key = String(normalized || '').trim().toLowerCase();
  if(!key) return '';
  return CANONICAL_ESTADO_LABELS[key] || formatTitleCase_(key);
}

function getColumnMap_(sheet){
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    const raw = String(header || '');
    const trimmed = raw.trim();
    const normalized = normalizeHeader_(raw);
    [raw, trimmed, normalized].forEach(key => {
      if(key && map[key] === undefined) map[key] = index;
    });
  });
  return {headers, map};
}

function getColumnIndex_(map, headerName){
  if(!map) return undefined;
  const direct = map[headerName];
  if(direct !== undefined) return direct;
  const trimmed = String(headerName || '').trim();
  if(trimmed && map[trimmed] !== undefined) return map[trimmed];
  const normalized = normalizeHeader_(headerName);
  if(normalized){
    if(map[normalized] !== undefined) return map[normalized];
    const aliases = NORMALIZED_HEADER_ALIASES[normalized] || [];
    for(let i = 0; i < aliases.length; i++){
      const alias = aliases[i];
      if(alias && map[alias] !== undefined) return map[alias];
    }
  }
  return undefined;
}

function handleGetLeads_(e, user){
  e = e || { parameter: {} };
  const sheetName = e.parameter.sheet || 'Lead Recuperados';
  if(user && !userCanAccessSheet_(user, sheetName)){
    return jsonResponse({ error: 'No tienes permiso para consultar esta base.' }, e);
  }
  const asesor = String(e.parameter.asesor || '').trim().toLowerCase();
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return jsonResponse([], e);
  const {headers, map} = getColumnMap_(sheet);
  const values = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),headers.length).getValues();
  const idxAsesor = getColumnIndex_(map, 'Asesor');
  const rows = values
    .filter(r => {
      const a = idxAsesor !== undefined ? String(r[idxAsesor] || '').toLowerCase().trim() : '';
      return !asesor || a === asesor;
    })
    .map(r => {
      const obj = {};
      COLUMNS.forEach(h => {
        const idx = getColumnIndex_(map, h);
        obj[h] = idx !== undefined ? r[idx] : '';
      });
      const contact = buildContactLinks_(obj['Teléfono'] || obj['Telefono'] || '', obj['Correo'] || '', obj['Nombre'] || '');
      Object.assign(obj, contact);
      return obj;
    });
  return jsonResponse(rows, e);
}

function handleListSheets_(e, user){
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheets = ss.getSheets();
  const baseOrder = ['planteles', 'regresos', 'reciclados', 'rescate'];
  const baseLookup = new Map();
  const names = [];
  for(let i = 0; i < sheets.length; i++){
    const sheet = sheets[i];
    if(!sheet) continue;
    const name = sheet.getName();
    const sanitized = sanitizeSheetName_(name);
    const trimmed = String(sanitized || '').trim();
    if(!trimmed) continue;
    const baseKey = canonicalBaseKey_(trimmed);
    if(baseOrder.includes(baseKey) && !baseLookup.has(baseKey)){
      baseLookup.set(baseKey, name);
    }
    const normalized = normalizeScopeKey_(trimmed);
    if(normalized === 'catalogo') continue;
    if(isLeadSheet_(sheet)) names.push(name);
  }
  const orderedBaseNames = baseOrder
    .map(key => baseLookup.get(key))
    .filter(Boolean);
  const ordered = Array.from(new Set([...orderedBaseNames, ...names]));
  const filtered = filterSheetsForUser_(ordered, user);
  return jsonResponse({ sheets: filtered }, e);
}

function isLeadSheet_(sheet){
  if(!sheet || sheet.isSheetHidden()) return false;
  const lastColumn = sheet.getLastColumn();
  if(lastColumn <= 0) return false;
  const { map } = getColumnMap_(sheet);
  return REQUIRED_HEADERS.every(h => getColumnIndex_(map, h) !== undefined);
}

function handleListAsesores_(e, user){
  e = e || { parameter: {} };
  const sheetName = e.parameter.sheet || 'Lead Recuperados';
  if(user && !userCanAccessSheet_(user, sheetName)){
    return jsonResponse({asesores:[]}, e);
  }
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return jsonResponse({asesores:[]}, e);
  const {headers, map} = getColumnMap_(sheet);
  const values = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),headers.length).getValues();
  const idx = getColumnIndex_(map, 'Asesor');
  const set = new Set();
  if(idx !== undefined) values.forEach(r=>{
    const a = r[idx];
    if(a){
      const name = String(a).trim();
      if(name.toLowerCase() !== 'sistema') set.add(name);
    }
  });
  return jsonResponse({asesores:[...set]}, e);
}

function handleReport_(e, user){
  e = e || { parameter: {} };
  const type = e.parameter.type || 'estado';
  const asesor = String(e.parameter.asesor || '').trim().toLowerCase();
  const etapaFilter = String(e.parameter.etapa || '').trim().toLowerCase();
  const fi = e.parameter.fi || '';
  const ff = e.parameter.ff || '';
  const sheetName = e.parameter.base || e.parameter.sheet || 'Lead Recuperados';
  if(user && !userCanAccessSheet_(user, sheetName)){
    return jsonResponse({data:{}}, e);
  }
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return jsonResponse({data:{}}, e);
  const {headers, map} = getColumnMap_(sheet);
  const values = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),headers.length).getValues();
  const idx = {
    asesor: getColumnIndex_(map, 'Asesor'),
    etapa: getColumnIndex_(map, 'Etapa'),
    estado: getColumnIndex_(map, 'Estado'),
    asignacion: getColumnIndex_(map, 'Asignación')
  };
  const filtered = values.filter(r=>{
    const a = idx.asesor !== undefined ? String(r[idx.asesor]||'').toLowerCase().trim() : '';
    const et = idx.etapa !== undefined ? String(r[idx.etapa]||'').toLowerCase().trim() : '';
    const f = idx.asignacion !== undefined ? r[idx.asignacion] : '';
    return (!asesor || a===asesor) &&
           (!etapaFilter || et===etapaFilter) &&
           (!fi || f>=fi) &&
           (!ff || f<=ff);
  });
  let key;
  if(type === 'detalle') key = idx.asesor;
  else if(type === 'sistema') key = idx.etapa;
  else key = idx.estado;
  const out = {};
  if(key !== undefined){
    filtered.forEach(r=>{ const k = String(r[key]||''); if(k) out[k]=(out[k]||0)+1; });
  }
  return jsonResponse({data:out}, e);
}

function handleUpdateLead_(e, user){
  e = e || { parameter: {} };
  const sheetName = e.parameter.sheet || 'Lead Recuperados';
  if(user && !userCanAccessSheet_(user, sheetName)){
    return jsonResponse({error:'No tienes permiso para actualizar esta base.'}, e);
  }
  const id = e.parameter.id;
  const etapa = e.parameter.etapa || '';
  const estado = e.parameter.estado || '';
  const comentario = e.parameter.comentario || '';
  const fecha = e.parameter.fecha || new Date().toISOString();
  const registrarToque = e.parameter.registrarToque !== 'false';
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return jsonResponse({error:'hoja no encontrada'}, e);
  const {headers, map} = getColumnMap_(sheet);
  const values = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),headers.length).getValues();
  const idx = {
    ID: getColumnIndex_(map, 'ID'),
    Etapa: getColumnIndex_(map, 'Etapa'),
    Estado: getColumnIndex_(map, 'Estado'),
    Comentario: getColumnIndex_(map, 'Comentario'),
    Asignacion: getColumnIndex_(map, 'Asignación'),
    Toques: [
      getColumnIndex_(map, 'Toque 1'),
      getColumnIndex_(map, 'Toque 2'),
      getColumnIndex_(map, 'Toque 3'),
      getColumnIndex_(map, 'Toque 4')
    ],
    Resol: getColumnIndex_(map, 'Resolución')
  };
  const rowIndex = values.findIndex(r => idx.ID !== undefined && String(r[idx.ID]) === String(id));
  if(rowIndex < 0) return jsonResponse({error:'id no encontrado'}, e);
  const row = values[rowIndex];
  if(idx.Etapa !== undefined) row[idx.Etapa] = etapa;
  if(idx.Estado !== undefined) row[idx.Estado] = estado;
  if(idx.Comentario !== undefined) row[idx.Comentario] = comentario;
  if(registrarToque){
    if(idx.Asignacion !== undefined && !row[idx.Asignacion]) row[idx.Asignacion] = formatDateTime_(fecha);
    const toqueValue = buildToqueValue_(fecha, estado || etapa);
    for(const c of idx.Toques){
      if(c !== undefined && !row[c]){
        row[c] = toqueValue;
        break;
      }
    }
  }
  const toqueStates = idx.Toques
    .map(c => c !== undefined ? normalizeState_(extractToqueState_(row[c])) : '')
    .filter(Boolean);
  if(registrarToque){
    const negatives = toqueStates.filter(s => NEGATIVE_STATES.indexOf(s) >= 0).length;
    if(negatives >= 3){
      if(idx.Etapa !== undefined) row[idx.Etapa] = 'Descartado';
      if(idx.Estado !== undefined) row[idx.Estado] = 'Límite intentos de contacto';
    }
  }
  const finalEtapa = idx.Etapa !== undefined ? row[idx.Etapa] : etapa;
  const finalEstado = idx.Estado !== undefined ? row[idx.Estado] : estado;
  if(idx.Resol !== undefined){
    row[idx.Resol] = computeResolutionForLead_(sheetName, finalEtapa, finalEstado);
  }
  sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
  const responseLead = {
    etapa: finalEtapa,
    estado: finalEstado,
    comentario: idx.Comentario !== undefined ? row[idx.Comentario] : comentario,
    asignacion: idx.Asignacion !== undefined ? row[idx.Asignacion] : '',
    resolucion: idx.Resol !== undefined ? row[idx.Resol] : '',
    toques: idx.Toques.map(c => c !== undefined ? row[c] || '' : '')
  };
  return jsonResponse({ok:true, lead: responseLead}, e);
}

function computeDiagnosticsReport_(requestedSheet){
  const requested = String(requestedSheet || '').trim();
  const timezone = Session.getScriptTimeZone() || '';
  const result = {
    timestamp: new Date().toISOString(),
    sheet: requested,
    connection: false,
    read: false,
    write: false,
    duplicates: [],
    errors: [],
    warnings: [],
    checks: {
      connection: { ok: false, message: '' },
      read: { ok: false, message: '' },
      write: { ok: false, message: '' }
    },
    sheets: [],
    aggregate: { duplicateIds: [], duplicatePhones: [] },
    login: getAuthDiagnostics_(),
    status: { spreadsheet: '', totalSheets: 0, requestedSheet: requested },
    summary: { sheetsWithErrors: 0, sheetsWithWarnings: 0, invalidRows: 0 },
    system: { status: 'En espera', timezone }
  };
  const ss = getSpreadsheet_();
  if(!ss){
    result.errors.push(SPREADSHEET_ERROR_MESSAGE);
    result.checks.connection.message = SPREADSHEET_ERROR_MESSAGE;
    result.system.status = 'Con incidencias';
    return result;
  }
  try{
    result.status.spreadsheet = ss.getName();
  }catch(err){
    result.status.spreadsheet = '';
  }
  const leadSheets = ss.getSheets().filter(isLeadSheet_);
  result.status.totalSheets = leadSheets.length;
  result.status.sheetNames = leadSheets.map(s => s.getName());
  let targetSheet = requested ? ss.getSheetByName(requested) : null;
  if(requested && !targetSheet){
    result.warnings.push('La hoja solicitada "' + requested + '" no existe. Se utilizará la primera hoja disponible.');
  }
  if(targetSheet && !isLeadSheet_(targetSheet)){
    result.warnings.push('La hoja "' + targetSheet.getName() + '" no coincide con el formato esperado de leads.');
  }
  if(!targetSheet){
    targetSheet = leadSheets.length ? leadSheets[0] : null;
  }
  if(!targetSheet){
    result.checks.connection.message = 'Sin hojas válidas para auditar.';
    result.errors.push('No se encontró ninguna hoja válida para analizar.');
    result.system.status = 'Con incidencias';
    return result;
  }
  result.sheet = targetSheet.getName();
  result.connection = true;
  result.checks.connection.ok = true;
  result.checks.connection.message = 'Hoja "' + result.sheet + '" disponible para análisis.';

  try{
    const rows = Math.max(1, Math.min(targetSheet.getLastRow(), 5));
    const cols = Math.max(1, Math.min(targetSheet.getLastColumn(), 5));
    targetSheet.getRange(1, 1, rows, cols).getValues();
    result.read = true;
    result.checks.read.ok = true;
    result.checks.read.message = 'Lectura realizada correctamente.';
  }catch(err){
    result.errors.push('Error de lectura en ' + result.sheet + ': ' + err.message);
    result.checks.read.message = err.message;
  }

  try{
    const cell = targetSheet.getRange(1,1);
    cell.setValue(cell.getValue());
    result.write = true;
    result.checks.write.ok = true;
    result.checks.write.message = 'Escritura verificada correctamente.';
  }catch(err){
    result.errors.push('Error de escritura en ' + result.sheet + ': ' + err.message);
    result.checks.write.message = err.message;
  }

  const trackers = {
    ids: Object.create(null),
    phones: Object.create(null)
  };
  leadSheets.forEach(sheet => {
    const summary = auditSheetForDiagnostics_(sheet, trackers);
    result.sheets.push(summary);
  });

  result.aggregate.duplicateIds = buildAggregateList_(trackers.ids, 10);
  result.aggregate.duplicatePhones = buildAggregateList_(trackers.phones, 10);

  const selectedSummary = result.sheets.find(s => s.name === result.sheet);
  if(selectedSummary){
    result.duplicates = selectedSummary.duplicateIds.map(d => d.value);
  }

  result.summary.sheetsWithErrors = result.sheets.filter(s => !s.ok).length;
  result.summary.sheetsWithWarnings = result.sheets.filter(s => s.hasWarnings).length;
  result.summary.invalidRows = result.sheets.reduce((acc, s) => acc + s.invalidCount, 0);
  result.system.status = result.summary.sheetsWithErrors ? 'Con incidencias' : 'Operativo';

  return result;
}

function handleDiagnostics_(e, user){
  e = e || { parameter: {} };
  const requestedSheet = e.parameter.sheet || '';
  const report = computeDiagnosticsReport_(requestedSheet);
  return jsonResponse(report, e);
}

function runDailyDiagnosticsAudit(){
  const requestedSheet = getPriorityDiagnosticsBase_();
  const report = computeDiagnosticsReport_(requestedSheet);
  const processed = processDiagnosticsReport_(report, requestedSheet);
  return { requestedSheet, report, processed };
}

function ensureDailyDiagnosticsTrigger(){
  const handlerName = 'runDailyDiagnosticsAudit';
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === handlerName);
  if(exists) return false;
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  return true;
}

function processDiagnosticsReport_(report, requestedSheet){
  const summary = { logged: false, notified: false };
  try{
    summary.logged = logDiagnosticsAudit_(report, requestedSheet);
  }catch(err){
    try{
      Logger.log('[Diagnostics] Error al registrar el resultado: ' + err.message);
    }catch(loggingError){
      // ignored on purpose
    }
  }
  try{
    const sheetsWithErrors = Number(report?.summary?.sheetsWithErrors || 0);
    if(sheetsWithErrors > 0){
      summary.notified = notifyDiagnosticsIncident_(report);
    }
  }catch(err){
    try{
      Logger.log('[Diagnostics] Error al enviar la notificación: ' + err.message);
    }catch(loggingError){
      // ignored on purpose
    }
  }
  return summary;
}

function logDiagnosticsAudit_(report, requestedSheet){
  const sheet = getDiagnosticsControlSheet_();
  if(!sheet) return false;
  const requested = String(requestedSheet || '').trim() || String(report?.status?.requestedSheet || '').trim();
  const analyzed = String(report?.sheet || '').trim();
  const errors = Number(report?.summary?.sheetsWithErrors || 0);
  const warnings = Number(report?.summary?.sheetsWithWarnings || 0);
  const invalidRows = Number(report?.summary?.invalidRows || 0);
  const status = String(report?.system?.status || '').trim();
  const errorMessages = Array.isArray(report?.errors) ? report.errors.filter(Boolean) : [];
  const warningMessages = Array.isArray(report?.warnings) ? report.warnings.filter(Boolean) : [];
  const messageParts = [];
  if(errorMessages.length){
    messageParts.push('Errores: ' + errorMessages.join(' | '));
  }
  if(warningMessages.length){
    messageParts.push('Advertencias: ' + warningMessages.join(' | '));
  }
  if(!messageParts.length){
    if(status){
      messageParts.push('Estado: ' + status);
    }else{
      messageParts.push('Sin incidencias registradas.');
    }
  }
  const summaryText = messageParts.join(' · ');
  const timestamp = formatDateTime_(new Date());
  sheet.appendRow([
    timestamp,
    requested,
    analyzed,
    errors,
    warnings,
    invalidRows,
    status,
    summaryText
  ]);
  return true;
}

function notifyDiagnosticsIncident_(report){
  const recipients = getDiagnosticsAlertRecipients_();
  const webhookUrl = getDiagnosticsWebhookUrl_();
  const sheet = String(report?.sheet || '').trim() || 'bases';
  const errors = Number(report?.summary?.sheetsWithErrors || 0);
  const warnings = Number(report?.summary?.sheetsWithWarnings || 0);
  const invalidRows = Number(report?.summary?.invalidRows || 0);
  const status = String(report?.system?.status || '').trim();
  const errorMessages = Array.isArray(report?.errors) ? report.errors.filter(Boolean) : [];
  const warningMessages = Array.isArray(report?.warnings) ? report.warnings.filter(Boolean) : [];
  const header = 'Diagnóstico diario · Incidencias detectadas en ' + sheet;
  const bodyLines = [
    header,
    '',
    'Estado: ' + (status || 'Desconocido'),
    'Hojas con errores: ' + errors,
    'Hojas con advertencias: ' + warnings,
    'Filas inválidas: ' + invalidRows,
    ''
  ];
  if(errorMessages.length){
    bodyLines.push('Errores:');
    errorMessages.forEach(message => bodyLines.push('• ' + message));
    bodyLines.push('');
  }
  if(warningMessages.length){
    bodyLines.push('Advertencias:');
    warningMessages.forEach(message => bodyLines.push('• ' + message));
    bodyLines.push('');
  }
  const body = bodyLines.join('\n');
  let notified = false;
  if(recipients.length){
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: header,
      body
    });
    notified = true;
  }
  if(webhookUrl){
    const payload = {
      text: header,
      status,
      summary: report?.summary || {},
      errors: errorMessages,
      warnings: warningMessages,
      timestamp: report?.timestamp || new Date().toISOString()
    };
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    notified = true;
  }
  return notified;
}

function getPriorityDiagnosticsBase_(){
  try{
    const props = PropertiesService.getScriptProperties();
    if(props){
      const configured = String(props.getProperty(DIAGNOSTICS_PRIORITY_BASE_PROPERTY) || '').trim();
      if(configured) return configured;
    }
  }catch(err){
    try{
      Logger.log('[Diagnostics] Error al leer la base prioritaria: ' + err.message);
    }catch(loggingError){
      // ignored on purpose
    }
  }
  try{
    const ss = getSpreadsheet_();
    if(!ss) return '';
    const leadSheets = ss.getSheets().filter(isLeadSheet_);
    if(leadSheets.length){
      return leadSheets[0].getName();
    }
  }catch(err){
    try{
      Logger.log('[Diagnostics] No fue posible determinar la base prioritaria: ' + err.message);
    }catch(loggingError){
      // ignored on purpose
    }
  }
  return '';
}

function getDiagnosticsControlSheet_(){
  const ss = getSpreadsheet_();
  if(!ss){
    throw new Error(SPREADSHEET_ERROR_MESSAGE);
  }
  let sheet = ss.getSheetByName(DIAGNOSTICS_CONTROL_SHEET_NAME);
  if(sheet) return sheet;
  sheet = ss.insertSheet(DIAGNOSTICS_CONTROL_SHEET_NAME);
  sheet.getRange(1, 1, 1, DIAGNOSTICS_CONTROL_HEADERS.length).setValues([DIAGNOSTICS_CONTROL_HEADERS]);
  return sheet;
}

function getDiagnosticsAlertRecipients_(){
  try{
    const props = PropertiesService.getScriptProperties();
    if(!props) return [];
    const raw = props.getProperty(DIAGNOSTICS_ALERT_RECIPIENTS_PROPERTY);
    if(!raw) return [];
    return String(raw)
      .split(/[,;\n]/)
      .map(item => String(item || '').trim())
      .filter(Boolean);
  }catch(err){
    return [];
  }
}

function getDiagnosticsWebhookUrl_(){
  try{
    const props = PropertiesService.getScriptProperties();
    if(!props) return '';
    return String(props.getProperty(DIAGNOSTICS_ALERT_WEBHOOK_PROPERTY) || '').trim();
  }catch(err){
    return '';
  }
}

function handleResolveDuplicates_(e, user){
  e = e || { parameter: {} };
  const sheetName = e.parameter.sheet || '';
  if(user && !userCanAccessSheet_(user, sheetName) && !isAdminUser_(user)){
    return jsonResponse({ ok:false, error:'No tienes permiso para depurar esta base.' }, e);
  }
  if(!sheetName){
    return jsonResponse({ ok:false, error:'Falta el parámetro "sheet".' }, e);
  }
  const ss = getSpreadsheet_();
  if(!ss){
    return jsonResponse({ ok:false, error: SPREADSHEET_ERROR_MESSAGE }, e);
  }
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet){
    return jsonResponse({ ok:false, error:'La hoja "' + sheetName + '" no existe.' }, e);
  }
  const duplicates = scanSheetDuplicates_(sheet);
  const idGroups = duplicates.idGroups || [];
  const phoneGroups = duplicates.phoneGroups || [];
  if(!idGroups.length && !phoneGroups.length){
    return jsonResponse({
      ok: true,
      sheet: sheet.getName(),
      displayName: sanitizeSheetName_(sheet.getName()),
      removedRows: 0,
      highlightedRows: 0,
      groupsProcessed: 0,
      message: 'No se encontraron duplicados para eliminar.'
    }, e);
  }
  if(!duplicates.hasIdColumn && !duplicates.hasPhoneColumn){
    return jsonResponse({ ok:false, error:'La hoja no tiene columnas de ID o Teléfono para detectar duplicados.' }, e);
  }

  const keepers = new Set();
  const toDelete = new Set();
  let groupsProcessed = 0;
  const registerGroup = rows => {
    if(!Array.isArray(rows) || rows.length <= 1) return;
    const sorted = rows.slice().sort((a,b)=>a-b);
    let keep = sorted.find(r => keepers.has(r));
    if(keep === undefined){
      keep = sorted.find(r => !toDelete.has(r));
    }
    if(keep === undefined){
      keep = sorted[0];
    }
    keepers.add(keep);
    toDelete.delete(keep);
    sorted.forEach(row => {
      if(row === keep) return;
      if(keepers.has(row)) return;
      toDelete.add(row);
    });
    groupsProcessed++;
  };

  idGroups.forEach(group => registerGroup(group.rows));
  phoneGroups.forEach(group => registerGroup(group.rows));

  const lastRow = duplicates.lastRow;
  const lastColumn = duplicates.lastColumn;
  const highlightRows = Array.from(keepers)
    .filter(row => row >= 2 && row <= lastRow)
    .sort((a,b)=>a-b);
  if(lastColumn > 0 && highlightRows.length){
    const color = '#f59e0b';
    highlightRows.forEach(row => {
      sheet.getRange(row, 1, 1, lastColumn).setBackground(color);
    });
  }
  SpreadsheetApp.flush();
  const deleteRows = Array.from(toDelete)
    .filter(row => row >= 2 && row <= lastRow && highlightRows.indexOf(row) === -1)
    .sort((a,b)=>b-a);
  deleteRows.forEach(row => sheet.deleteRow(row));

  const displayName = sanitizeSheetName_(sheet.getName());
  const response = {
    ok: true,
    sheet: sheet.getName(),
    displayName,
    removedRows: deleteRows.length,
    highlightedRows: highlightRows.length,
    keptRows: highlightRows,
    deletedRows: deleteRows,
    groupsProcessed,
    duplicateGroups: { ids: idGroups.length, phones: phoneGroups.length },
    message: deleteRows.length
      ? 'Se eliminaron ' + deleteRows.length + ' filas duplicadas en "' + (displayName || sheet.getName()) + '".'
      : 'No se encontraron duplicados para eliminar.'
  };
  return jsonResponse(response, e);
}

function handleRepairSheetStructure_(e, body, user){
  body = body || {};
  const sheetName = String(body.sheet || '').trim();
  if(!sheetName){
    return jsonResponse({ ok:false, error:'Falta el parámetro "sheet".' }, e);
  }
  if(user && !userCanAccessSheet_(user, sheetName) && !isAdminUser_(user)){
    return jsonResponse({ ok:false, error:'No tienes permiso para modificar esta base.' }, e);
  }
  const ss = getSpreadsheet_();
  if(!ss){
    return jsonResponse({ ok:false, error: SPREADSHEET_ERROR_MESSAGE }, e);
  }
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet){
    return jsonResponse({ ok:false, error:'La hoja "' + sheetName + '" no existe.' }, e);
  }
  const options = {
    ensureColumns: Array.isArray(body.ensureColumns) ? body.ensureColumns.map(col => String(col || '').trim()).filter(Boolean) : [],
    ensureContactColumns: body.ensureContactColumns !== false,
    fixInvalidRows: body.fixInvalidRows !== false,
    rebuildContactLinks: body.rebuildContactLinks !== false
  };
  const result = repairSheetStructure_(sheet, options);
  return jsonResponse(result, e);
}

function repairSheetStructure_(sheet, options){
  const opts = options || {};
  const ensureColumns = new Set();
  (opts.ensureColumns || []).forEach(name => {
    const normalized = normalizeHeader_(name);
    if(!normalized) return;
    ensureColumns.add(name);
  });
  if(opts.ensureContactColumns){
    ['Telefono normalizado','TelefonoNormalizado','Telefono aircall','TelefonoAircall','Telefono whatsapp','TelefonoWhatsapp','Correo gmail','CorreoGmail']
      .forEach(name => ensureColumns.add(name));
  }

  const initial = getColumnMap_(sheet);
  const headers = initial.headers.slice();
  const headerSet = new Set(headers.map(normalizeHeader_));
  const addedColumns = [];
  ensureColumns.forEach(name => {
    const normalized = normalizeHeader_(name);
    if(!normalized || headerSet.has(normalized)) return;
    headerSet.add(normalized);
    headers.push(name);
    addedColumns.push(name);
  });

  if(addedColumns.length){
    const targetColumns = headers.length;
    const maxColumns = sheet.getMaxColumns();
    if(targetColumns > maxColumns){
      sheet.insertColumnsAfter(Math.max(1, maxColumns), targetColumns - maxColumns);
    }
    sheet.getRange(1, 1, 1, targetColumns).setValues([headers]);
  }

  const { headers: updatedHeaders, map } = getColumnMap_(sheet);
  const response = {
    ok: true,
    sheet: sheet.getName(),
    displayName: sanitizeSheetName_(sheet.getName()),
    addedColumns,
    updatedRows: 0,
    generatedIds: 0,
    normalizedPhones: 0,
    normalizedEtapas: 0,
    normalizedEstados: 0,
    rebuiltContacts: 0,
    actions: []
  };
  if(addedColumns.length){
    response.actions.push('Se agregaron columnas faltantes: ' + addedColumns.join(', '));
  }

  const dataRows = Math.max(0, sheet.getLastRow() - 1);
  if(dataRows <= 0){
    if(!response.actions.length){
      response.actions.push('No hay filas de datos para reparar.');
    }
    return response;
  }

  const shouldFixRows = opts.fixInvalidRows !== false;
  const shouldRebuildContacts = opts.rebuildContactLinks !== false;
  if(!shouldFixRows && !shouldRebuildContacts){
    if(!response.actions.length){
      response.actions.push('No se solicitaron reparaciones de filas.');
    }
    return response;
  }

  const range = sheet.getRange(2, 1, dataRows, updatedHeaders.length);
  const values = range.getValues();
  const idx = {
    ID: getColumnIndex_(map, 'ID'),
    Nombre: getColumnIndex_(map, 'Nombre'),
    Telefono: getColumnIndex_(map, 'Teléfono'),
    TelefonoAlt: getColumnIndex_(map, 'Telefono'),
    Correo: getColumnIndex_(map, 'Correo'),
    Etapa: getColumnIndex_(map, 'Etapa'),
    Estado: getColumnIndex_(map, 'Estado'),
    Resolucion: getColumnIndex_(map, 'Resolución'),
    ResolucionAlt: getColumnIndex_(map, 'Resolucion')
  };
  const contactColumns = [
    { header: 'Telefono normalizado', key: 'TelefonoNormalizado' },
    { header: 'TelefonoNormalizado', key: 'TelefonoNormalizado' },
    { header: 'Telefono aircall', key: 'TelefonoAircall' },
    { header: 'TelefonoAircall', key: 'TelefonoAircall' },
    { header: 'Telefono whatsapp', key: 'TelefonoWhatsapp' },
    { header: 'TelefonoWhatsapp', key: 'TelefonoWhatsapp' },
    { header: 'Correo gmail', key: 'CorreoGmail' },
    { header: 'CorreoGmail', key: 'CorreoGmail' }
  ];

  for(let i = 0; i < values.length; i++){
    const row = values[i];
    let rowChanged = false;
    let phoneForContacts = '';
    if(shouldFixRows && idx.ID !== undefined){
      const currentId = String(row[idx.ID] || '').trim();
      if(!currentId){
        const newId = generateLeadId_();
        row[idx.ID] = newId;
        response.generatedIds++;
        rowChanged = true;
      }
    }

    if(shouldFixRows){
      const phoneIndex = idx.Telefono !== undefined ? idx.Telefono : idx.TelefonoAlt;
      if(phoneIndex !== undefined){
        const rawPhone = row[phoneIndex];
        let parsed = parsePhoneValue_(rawPhone);
        if(!parsed.local && rawPhone){
          const digitsOnly = String(rawPhone).replace(/\D+/g, '');
          if(digitsOnly.length >= 10){
            parsed = parsePhoneValue_(digitsOnly.slice(-10));
          }
        }
        if(parsed.local){
          phoneForContacts = parsed.local;
          let phoneChanged = false;
          if(idx.Telefono !== undefined && row[idx.Telefono] !== parsed.local){
            row[idx.Telefono] = parsed.local;
            phoneChanged = true;
          }
          if(idx.TelefonoAlt !== undefined && row[idx.TelefonoAlt] !== parsed.local){
            row[idx.TelefonoAlt] = parsed.local;
            phoneChanged = true;
          }
          if(phoneChanged){
            response.normalizedPhones++;
            rowChanged = true;
          }
        }
      }
    }

    let finalEtapaNorm = '';
    let finalEtapaLabel = '';
    if(shouldFixRows && idx.Etapa !== undefined){
      const rawEtapa = row[idx.Etapa];
      let etapaNorm = normalizeValue_(rawEtapa);
      if(!etapaNorm || VALID_ETAPAS.indexOf(etapaNorm) === -1){
        etapaNorm = 'nuevo';
      }
      finalEtapaNorm = etapaNorm;
      finalEtapaLabel = formatEtapaLabel_(etapaNorm);
      if(row[idx.Etapa] !== finalEtapaLabel){
        row[idx.Etapa] = finalEtapaLabel;
        rowChanged = true;
        response.normalizedEtapas++;
      }
    }else if(idx.Etapa !== undefined){
      finalEtapaLabel = row[idx.Etapa];
      finalEtapaNorm = normalizeValue_(finalEtapaLabel);
    }

    let finalEstadoLabel = '';
    if(shouldFixRows && idx.Estado !== undefined){
      const rawEstado = row[idx.Estado];
      let estadoNorm = normalizeValue_(rawEstado);
      const validStates = finalEtapaNorm ? (ETAPA_ESTADO_MAP[finalEtapaNorm] || []) : [];
      if(!estadoNorm || !VALID_ESTADOS_SET.has(estadoNorm)){
        estadoNorm = validStates.length ? validStates[0] : estadoNorm;
      }else if(validStates.length && validStates.indexOf(estadoNorm) === -1){
        estadoNorm = validStates[0] || estadoNorm;
      }
      if(estadoNorm){
        finalEstadoLabel = formatEstadoLabel_(estadoNorm);
        if(row[idx.Estado] !== finalEstadoLabel){
          row[idx.Estado] = finalEstadoLabel;
          rowChanged = true;
          response.normalizedEstados++;
        }
      }
    }else if(idx.Estado !== undefined){
      finalEstadoLabel = row[idx.Estado];
    }

    if(shouldFixRows){
      const resolution = computeResolutionForLead_(sheet.getName(), finalEtapaLabel || row[idx.Etapa] || '', finalEstadoLabel || row[idx.Estado] || '');
      if(idx.Resolucion !== undefined && row[idx.Resolucion] !== resolution){
        row[idx.Resolucion] = resolution;
        rowChanged = true;
      }
      if(idx.ResolucionAlt !== undefined && row[idx.ResolucionAlt] !== resolution){
        row[idx.ResolucionAlt] = resolution;
        rowChanged = true;
      }
    }

    if(shouldRebuildContacts){
      const telefono = phoneForContacts || (idx.Telefono !== undefined ? row[idx.Telefono] : (idx.TelefonoAlt !== undefined ? row[idx.TelefonoAlt] : ''));
      const correo = idx.Correo !== undefined ? String(row[idx.Correo] || '').trim() : '';
      const nombre = idx.Nombre !== undefined ? String(row[idx.Nombre] || '').trim() : '';
      const contact = buildContactLinks_(telefono, correo, nombre);
      let contactChanged = false;
      contactColumns.forEach(item => {
        const columnIndex = getColumnIndex_(map, item.header);
        if(columnIndex === undefined) return;
        const value = contact[item.key] || '';
        if(row[columnIndex] !== value){
          row[columnIndex] = value;
          rowChanged = true;
          contactChanged = true;
        }
      });
      if(contactChanged){
        response.rebuiltContacts++;
      }
    }

    if(rowChanged){
      response.updatedRows++;
    }
  }

  if(response.updatedRows > 0){
    range.setValues(values);
    response.actions.push('Se actualizaron ' + response.updatedRows + ' filas.');
  }
  if(response.generatedIds > 0){
    response.actions.push('Se asignaron ' + response.generatedIds + ' IDs faltantes.');
  }
  if(response.normalizedPhones > 0){
    response.actions.push('Se normalizaron ' + response.normalizedPhones + ' teléfonos.');
  }
  if(response.normalizedEtapas > 0){
    response.actions.push('Se corrigieron ' + response.normalizedEtapas + ' etapas.');
  }
  if(response.normalizedEstados > 0){
    response.actions.push('Se ajustaron ' + response.normalizedEstados + ' estados.');
  }
  if(response.rebuiltContacts > 0){
    response.actions.push('Se regeneraron enlaces de contacto en ' + response.rebuiltContacts + ' filas.');
  }
  if(!response.actions.length){
    response.actions.push('No se requirieron cambios en la hoja.');
  }
  return response;
}

function scanSheetDuplicates_(sheet){
  if(!sheet){
    return { idGroups: [], phoneGroups: [], lastRow: 0, lastColumn: 0, hasIdColumn: false, hasPhoneColumn: false };
  }
  const {headers, map} = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const dataRowCount = Math.max(0, lastRow - 1);
  const idGroups = [];
  const phoneGroups = [];
  const idIndex = getColumnIndex_(map, 'ID');
  let phoneIndex = getColumnIndex_(map, 'Teléfono');
  if(phoneIndex === undefined) phoneIndex = getColumnIndex_(map, 'Telefono');
  const idOccurrences = Object.create(null);
  const phoneOccurrences = Object.create(null);

  if(dataRowCount > 0){
    const values = sheet.getRange(2, 1, dataRowCount, headers.length).getValues();
    values.forEach((row, index) => {
      const rowNumber = index + 2;
      if(idIndex !== undefined){
        const idValue = String(row[idIndex] || '').trim();
        if(idValue){
          if(!idOccurrences[idValue]) idOccurrences[idValue] = [];
          idOccurrences[idValue].push(rowNumber);
        }
      }
      if(phoneIndex !== undefined){
        const rawPhone = row[phoneIndex];
        if(rawPhone){
          const parsed = parsePhoneValue_(rawPhone);
          if(parsed.local){
            const key = parsed.local;
            if(!phoneOccurrences[key]) phoneOccurrences[key] = [];
            phoneOccurrences[key].push(rowNumber);
          }
        }
      }
    });
  }

  Object.keys(idOccurrences).forEach(key => {
    const rows = idOccurrences[key];
    if(rows.length > 1){
      idGroups.push({ key, rows: rows.slice().sort((a,b)=>a-b) });
    }
  });
  Object.keys(phoneOccurrences).forEach(key => {
    const rows = phoneOccurrences[key];
    if(rows.length > 1){
      phoneGroups.push({ key, rows: rows.slice().sort((a,b)=>a-b) });
    }
  });

  return {
    idGroups,
    phoneGroups,
    lastRow,
    lastColumn,
    hasIdColumn: idIndex !== undefined,
    hasPhoneColumn: phoneIndex !== undefined
  };
}

function auditSheetForDiagnostics_(sheet, globalTrackers){
  const {headers, map} = getColumnMap_(sheet);
  const missingColumns = REQUIRED_HEADERS
    .filter(h => getColumnIndex_(map, h) === undefined)
    .map(key => REQUIRED_HEADER_LABELS[key] || key);
  const unknownColumns = headers
    .map((header, idx) => ({ header, idx, normalized: normalizeHeader_(header) }))
    .filter(item => item.header && !KNOWN_HEADER_SET.has(item.normalized))
    .map(item => ({ name: item.header, column: columnToLetter_(item.idx + 1) }));
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
  const values = dataRowCount > 0 ? sheet.getRange(2, 1, dataRowCount, headers.length).getValues() : [];
  const idx = {
    ID: getColumnIndex_(map, 'ID'),
    Telefono: getColumnIndex_(map, 'Teléfono'),
    Correo: getColumnIndex_(map, 'Correo'),
    Etapa: getColumnIndex_(map, 'Etapa'),
    Estado: getColumnIndex_(map, 'Estado')
  };
  if(idx.Telefono === undefined) idx.Telefono = getColumnIndex_(map, 'Telefono');

  const invalidRows = [];
  const idOccurrences = Object.create(null);
  const phoneOccurrences = Object.create(null);

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const issues = [];
    if(idx.ID !== undefined){
      const idValue = String(row[idx.ID] || '').trim();
      if(!idValue){
        issues.push('ID vacío');
      }else{
        if(!idOccurrences[idValue]) idOccurrences[idValue] = [];
        idOccurrences[idValue].push(rowNumber);
        trackOccurrence_(globalTrackers.ids, idValue, sheet.getName(), rowNumber);
      }
    }

    if(idx.Telefono !== undefined){
      const rawPhone = row[idx.Telefono];
      if(!rawPhone){
        issues.push('Teléfono vacío');
      }else{
        const parsed = parsePhoneValue_(rawPhone);
        if(!parsed.local){
          issues.push('Teléfono inválido (' + rawPhone + ')');
        }else{
          const key = parsed.local;
          if(!phoneOccurrences[key]) phoneOccurrences[key] = { rows: [], samples: [] };
          const entry = phoneOccurrences[key];
          entry.rows.push(rowNumber);
          const sample = String(rawPhone);
          if(entry.samples.length < 3 && entry.samples.indexOf(sample) === -1) entry.samples.push(sample);
          trackOccurrence_(globalTrackers.phones, key, sheet.getName(), rowNumber);
        }
      }
    }

    if(idx.Etapa !== undefined){
      const etapaRaw = row[idx.Etapa];
      const etapaNorm = normalizeValue_(etapaRaw);
      if(!etapaNorm){
        issues.push('Etapa vacía');
      }else if(VALID_ETAPAS.indexOf(etapaNorm) === -1){
        issues.push('Etapa no reconocida (' + etapaRaw + ')');
      }
      if(idx.Estado !== undefined){
        const estadoRaw = row[idx.Estado];
        const estadoNorm = normalizeValue_(estadoRaw);
        if(estadoNorm){
          if(!VALID_ESTADOS_SET.has(estadoNorm)){
            issues.push('Estado no reconocido (' + estadoRaw + ')');
          }else if(VALID_ETAPAS.indexOf(etapaNorm) >= 0){
            const validStates = ETAPA_ESTADO_MAP[etapaNorm] || [];
            if(validStates.length && validStates.indexOf(estadoNorm) === -1){
              issues.push('Estado "' + estadoRaw + '" no corresponde a etapa "' + etapaRaw + '"');
            }
          }
        }
      }
    }

    if(idx.Correo !== undefined){
      const email = String(row[idx.Correo] || '').trim();
      if(email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){
        issues.push('Correo no válido (' + email + ')');
      }
    }

    if(issues.length){
      invalidRows.push({ row: rowNumber, issues });
    }
  });

  const duplicateIdEntries = Object.keys(idOccurrences)
    .filter(key => idOccurrences[key].length > 1)
    .map(key => ({ value: key, rows: idOccurrences[key] }));
  const duplicatePhoneEntries = Object.keys(phoneOccurrences)
    .filter(key => phoneOccurrences[key].rows.length > 1)
    .map(key => ({ value: key, rows: phoneOccurrences[key].rows, samples: phoneOccurrences[key].samples }));

  const duplicateSampleLimit = 50;
  const invalidSampleLimit = 50;

  const summary = {
    name: sheet.getName(),
    totalRows: values.length,
    missingColumns,
    unknownColumns,
    duplicateIds: duplicateIdEntries.slice(0, duplicateSampleLimit),
    duplicateIdTotal: duplicateIdEntries.length,
    duplicatePhones: duplicatePhoneEntries.slice(0, duplicateSampleLimit),
    duplicatePhoneTotal: duplicatePhoneEntries.length,
    invalidRows: invalidRows.slice(0, invalidSampleLimit),
    invalidCount: invalidRows.length,
    truncatedInvalid: invalidRows.length > invalidSampleLimit
  };
  const hasWarnings = unknownColumns.length > 0;
  const hasErrors = missingColumns.length > 0 || duplicateIdEntries.length > 0 || duplicatePhoneEntries.length > 0 || invalidRows.length > 0;
  summary.ok = !hasErrors;
  summary.hasWarnings = hasWarnings;
  return summary;
}

function buildAggregateList_(store, sampleLimit){
  const limit = sampleLimit || 10;
  return Object.keys(store || {})
    .map(key => {
      const occ = store[key];
      if(!Array.isArray(occ) || occ.length <= 1) return null;
      return {
        value: key,
        count: occ.length,
        occurrences: occ.slice(0, limit),
        truncated: occ.length > limit
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.count - a.count);
}

function trackOccurrence_(store, key, sheetName, rowNumber){
  if(!key) return;
  if(!store[key]) store[key] = [];
  store[key].push({ sheet: sheetName, row: rowNumber });
}

function columnToLetter_(column){
  let col = Number(column) || 0;
  let letter = '';
  while(col > 0){
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
function parseJsonBody_(e){
  if(!e || !e.postData || !e.postData.contents) return {};
  try{
    return JSON.parse(e.postData.contents);
  }catch(err){
    return {};
  }
}

function extractToken_(e, body){
  if(body && body.token){
    return String(body.token || '').trim();
  }
  const paramToken = e?.parameter?.token;
  if(paramToken){
    return String(paramToken || '').trim();
  }
  const headers = e?.headers;
  if(headers){
    const authHeader = headers['Authorization'] || headers['authorization'] || headers['AUTHORIZATION'];
    if(authHeader && typeof authHeader === 'string'){
      const trimmed = authHeader.trim();
      if(trimmed.toLowerCase().startsWith('bearer ')){
        return trimmed.slice(7).trim();
      }
      return trimmed;
    }
  }
  return '';
}

function userHasRole_(user, roles){
  if(!roles) return true;
  const list = Array.isArray(roles) ? roles : [roles];
  const current = String(user?.role || '').trim().toLowerCase();
  return list.some(role => {
    const normalized = String(role || '').trim().toLowerCase();
    if(!normalized) return false;
    if(normalized === 'admin') return ADMIN_ROLES.has(current);
    return normalized === current;
  });
}

function isAdminUser_(user){
  return userHasRole_(user, 'admin');
}

function unauthorizedResponse_(message, e){
  return jsonResponse({ error: message || 'No autorizado', code: 'UNAUTHORIZED' }, e);
}

function forbiddenResponse_(message, e){
  return jsonResponse({ error: message || 'Operación no permitida', code: 'FORBIDDEN' }, e);
}

function requireAuth_(e, options){
  const opts = options || {};
  const token = extractToken_(e, opts.body);
  if(!token){
    return { error: true, response: unauthorizedResponse_('Falta token de autenticación.', e) };
  }
  const verification = verifyToken_(token);
  if(!verification.ok){
    return { error: true, response: unauthorizedResponse_(verification.message || 'Token inválido.', e) };
  }
  const user = verification.user;
  if(opts.requireActive && user && user.active === false){
    return { error: true, response: forbiddenResponse_('Tu usuario está inactivo.', e) };
  }
  if(opts.role && !userHasRole_(user, opts.role)){
    return { error: true, response: forbiddenResponse_('No tienes permisos suficientes.', e) };
  }
  ACTIVE_USER = user;
  return { user, token };
}

function generateSecret_(){
  const uuid = Utilities.getUuid().replace(/-/g, '');
  const uuid2 = Utilities.getUuid().replace(/-/g, '');
  return (uuid + uuid2).slice(0, 64);
}

function generateSalt_(){
  return Utilities.getUuid().replace(/-/g, '');
}

function generateTokenVersion_(){
  return Utilities.getUuid().replace(/-/g, '');
}

function generateTemporaryPassword_(){
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789@#$%*';
  let out = '';
  for(let i = 0; i < 10; i++){
    const index = Math.floor(Math.random() * chars.length);
    out += chars.charAt(index);
  }
  return out;
}

function getTokenSecret_(){
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(TOKEN_SECRET_PROPERTY);
  if(!secret){
    secret = generateSecret_();
    props.setProperty(TOKEN_SECRET_PROPERTY, secret);
  }
  return secret;
}

function hashPassword_(password, salt){
  const combo = `${String(salt || '')}::${String(password || '')}`;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combo);
  return Utilities.base64Encode(bytes);
}

function verifyPassword_(password, hash, salt){
  if(!hash) return false;
  const computed = hashPassword_(password, salt);
  return computed === hash;
}

function normalizeBoolean_(value){
  if(value === true || value === false) return Boolean(value);
  if(typeof value === 'number') return value !== 0;
  const str = String(value || '').trim().toLowerCase();
  if(!str) return false;
  return ['1','true','si','sí','active','activo','activa','yes','y'].indexOf(str) >= 0;
}

function normalizeRole_(role){
  const str = String(role || '').trim().toLowerCase();
  const valid = ['admin','coordinador','asesor','viewer'];
  if(!str) return 'asesor';
  if(ADMIN_ROLES.has(str)) return 'admin';
  if(valid.indexOf(str) >= 0) return str;
  return 'asesor';
}

function parseList_(value){
  if(Array.isArray(value)){
    return value.map(v => String(v || '').trim()).filter(Boolean);
  }
  return String(value || '')
    .split(/[;,\n]/)
    .map(v => v.trim())
    .filter(Boolean);
}

function joinList_(value){
  if(!value) return '';
  if(Array.isArray(value)){
    return value.map(v => String(v || '').trim()).filter(Boolean).join(', ');
  }
  return String(value || '').trim();
}

function ensureUserSheet_(){
  const ss = getSpreadsheet_();
  if(!ss) throw new Error(SPREADSHEET_ERROR_MESSAGE);
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  let created = false;
  if(!sheet){
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
    sheet.hideSheet();
    created = true;
  }
  const data = getColumnMap_(sheet);
  let { headers, map } = data;
  let updated = created;
  USER_HEADERS.forEach(header => {
    if(getColumnIndex_(map, header) === undefined){
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1, 1, 1).setValue(header);
      updated = true;
    }
  });
  if(updated){
    const refreshed = getColumnMap_(sheet);
    headers = refreshed.headers;
    map = refreshed.map;
  }
  return { sheet, headers, map };
}

function sanitizeTeamId_(value){
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toUpperCase();
}

function ensureTeamSheet_(){
  const ss = getSpreadsheet_();
  if(!ss) throw new Error(SPREADSHEET_ERROR_MESSAGE);
  let sheet = ss.getSheetByName(TEAMS_SHEET_NAME);
  let created = false;
  if(!sheet){
    sheet = ss.insertSheet(TEAMS_SHEET_NAME);
    sheet.getRange(1, 1, 1, TEAM_HEADERS.length).setValues([TEAM_HEADERS]);
    created = true;
  }
  const data = getColumnMap_(sheet);
  let { headers, map } = data;
  let updated = created;
  TEAM_HEADERS.forEach(header => {
    if(getColumnIndex_(map, header) === undefined){
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1, 1, 1).setValue(header);
      updated = true;
    }
  });
  if(updated){
    const refreshed = getColumnMap_(sheet);
    headers = refreshed.headers;
    map = refreshed.map;
  }
  return { sheet, headers, map };
}

function getTeamSheetData_(){
  return ensureTeamSheet_();
}

function getUserSheetData_(){
  return ensureUserSheet_();
}

function readUsersWithMeta_(){
  const { sheet, map } = getUserSheetData_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
  const users = values.map((row, idx) => {
    const user = mapUserRow_(row, map);
    user.rowIndex = idx + 2;
    return user;
  });
  return { sheet, map, users };
}

function mapUserRow_(row, map){
  const get = name => {
    const idx = getColumnIndex_(map, name);
    if(idx === undefined) return '';
    return row[idx];
  };
  const email = String(get('Email') || '').trim().toLowerCase();
  const userId = String(get('UserID') || '').trim();
  const name = String(get('Nombre') || '').trim();
  const role = normalizeRole_(get('Rol'));
  const planteles = parseList_(get('Planteles'));
  const bases = parseList_(get('Bases'));
  const activoRaw = get('Activo');
  const active = activoRaw === '' ? true : normalizeBoolean_(activoRaw);
  return {
    email,
    userId,
    name,
    role,
    planteles,
    bases,
    active,
    passwordHash: String(get('PasswordHash') || ''),
    salt: String(get('Salt') || ''),
    lastLogin: get('UltimoIngreso') || '',
    createdAt: get('CreadoEl') || '',
    updatedAt: get('ActualizadoEl') || '',
    tokenVersion: String(get('TokenVersion') || '')
  };
}

function sanitizeUserForClient_(user){
  if(!user) return null;
  return {
    email: user.email,
    userId: user.userId,
    name: user.name,
    role: user.role,
    planteles: Array.isArray(user.planteles) ? user.planteles : [],
    bases: Array.isArray(user.bases) ? user.bases : [],
    active: user.active !== false,
    lastLogin: user.lastLogin || '',
    createdAt: user.createdAt || '',
    updatedAt: user.updatedAt || ''
  };
}

function mapTeamRow_(row, map){
  const get = name => {
    const idx = getColumnIndex_(map, name);
    if(idx === undefined) return '';
    return row[idx];
  };
  const teamId = sanitizeTeamId_(get('TeamID'));
  const name = String(get('Nombre') || '').trim();
  const description = String(get('Descripcion') || '').trim();
  const planteles = parseList_(get('Planteles'));
  const bases = parseList_(get('Bases'));
  const members = parseList_(get('Miembros'));
  const activoRaw = get('Activo');
  const active = activoRaw === '' ? true : normalizeBoolean_(activoRaw);
  return {
    teamId,
    name,
    description,
    planteles,
    bases,
    members,
    active,
    createdAt: get('CreadoEl') || '',
    updatedAt: get('ActualizadoEl') || ''
  };
}

function readTeamsWithMeta_(){
  const { sheet, map } = getTeamSheetData_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
  const teams = values.map((row, idx) => {
    const team = mapTeamRow_(row, map);
    team.rowIndex = idx + 2;
    return team;
  });
  return { sheet, map, teams };
}

function sanitizeTeamForClient_(team){
  if(!team) return null;
  return {
    teamId: team.teamId,
    name: team.name,
    description: team.description || '',
    planteles: Array.isArray(team.planteles) ? team.planteles : [],
    bases: Array.isArray(team.bases) ? team.bases : [],
    members: Array.isArray(team.members) ? team.members : [],
    active: team.active !== false,
    createdAt: team.createdAt || '',
    updatedAt: team.updatedAt || ''
  };
}

function normalizeScopeKey_(value){
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

const GLOBAL_BASE_SCOPE_TOKENS = new Set([
  'todas',
  'todas las bases',
  'todas bases',
  'acceso global',
  'global',
  'global bases'
].map(normalizeScopeKey_));

const GLOBAL_PLANTEL_SCOPE_TOKENS = new Set([
  'todos',
  'todos los planteles',
  'todos los campus',
  'todos campus',
  'acceso global',
  'global',
  'global planteles'
].map(normalizeScopeKey_));

const BASE_NAME_ALIAS_MAP = (() => {
  const entries = [
    ['plantel', 'planteles'],
    ['planteles', 'planteles'],
    ['regreso', 'regresos'],
    ['regresos', 'regresos'],
    ['reciclado', 'reciclados'],
    ['reciclados', 'reciclados'],
    ['reciclada', 'reciclados'],
    ['recicladas', 'reciclados'],
    ['rescate', 'rescate'],
    ['rescatado', 'rescate'],
    ['rescatados', 'rescate'],
    ['rescatada', 'rescate'],
    ['rescatadas', 'rescate']
  ];
  const map = Object.create(null);
  entries.forEach(([alias, canonical]) => {
    const key = normalizeScopeKey_(alias);
    if(!key || map[key]) return;
    map[key] = canonical;
  });
  return map;
})();

function normalizeSheetKey_(name){
  const normalized = normalizeScopeKey_(sanitizeSheetName_(name));
  if(!normalized) return '';
  const withoutPrefix = normalized.replace(/^\d+\s*/, '').trim();
  return withoutPrefix || normalized;
}

function canonicalBaseKey_(value){
  const normalized = normalizeSheetKey_(value);
  if(!normalized) return '';
  return BASE_NAME_ALIAS_MAP[normalized] || normalized;
}

function getRoundRobinPropertyKey_(sheetName){
  const baseKey = canonicalBaseKey_(sheetName) || normalizeScopeKey_(sheetName) || 'default';
  const sanitized = String(baseKey || 'default')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'DEFAULT';
  return `${ROUND_ROBIN_PROPERTY_PREFIX}${sanitized}`;
}

function readRoundRobinIndex_(sheetName){
  try{
    const props = PropertiesService.getScriptProperties();
    if(!props) return 0;
    const key = getRoundRobinPropertyKey_(sheetName);
    const raw = props.getProperty(key);
    const value = Number(raw);
    if(!isFinite(value) || value < 0) return 0;
    return Math.floor(value);
  }catch(err){
    return 0;
  }
}

function writeRoundRobinIndex_(sheetName, index){
  try{
    const props = PropertiesService.getScriptProperties();
    if(!props) return;
    const key = getRoundRobinPropertyKey_(sheetName);
    const safeIndex = Math.max(0, Math.floor(Number(index) || 0));
    props.setProperty(key, String(safeIndex));
  }catch(err){
    // ignored on purpose
  }
}

function getActiveAsesoresForSheet_(sheetName){
  try{
    const meta = readUsersWithMeta_();
    const asesores = meta.users
      .filter(user => {
        if(!user || user.active === false) return false;
        const role = String(user.role || '').trim().toLowerCase();
        if(role !== 'asesor') return false;
        return userCanAccessSheet_(user, sheetName);
      })
      .map(user => ({
        name: String(user.name || '').trim() || user.userId || user.email,
        lead: user
      }));
    asesores.sort((a, b) => a.name.localeCompare(b.name, 'es', { sensitivity: 'base' }));
    return asesores.map(entry => entry.name);
  }catch(err){
    return [];
  }
}

function assignAsesoresRoundRobin_(rows, sheetName){
  if(!Array.isArray(rows) || !rows.length) return;
  const asesores = getActiveAsesoresForSheet_(sheetName);
  if(!asesores.length){
    try{
      Logger.log(`[RoundRobin] Base ${sheetName} · Sin asesores activos disponibles para asignación.`);
    }catch(err){
      // logging failures are ignored
    }
    return;
  }
  let pointer = readRoundRobinIndex_(sheetName);
  if(pointer >= asesores.length){
    pointer = pointer % asesores.length;
  }
  const startPointer = pointer;
  rows.forEach(row => {
    const current = asesores[pointer % asesores.length];
    row.asesorValue = current;
    pointer = (pointer + 1) % asesores.length;
  });
  writeRoundRobinIndex_(sheetName, pointer);
  try{
    Logger.log(`[RoundRobin] Base ${sheetName} · Inicio:${startPointer} · Fin:${pointer} · Asignados:${rows.length}`);
    rows.forEach((row, idx) => {
      Logger.log(`[RoundRobin] ${sheetName} · #${idx + 1} Lead:${row.leadId || ''} -> ${row.asesorValue}`);
    });
  }catch(err){
    // logging failures are ignored
  }
}

function isGlobalBaseValue_(value){
  if(value === undefined || value === null) return false;
  return GLOBAL_BASE_SCOPE_TOKENS.has(normalizeScopeKey_(value));
}

function hasGlobalBaseAccess_(values){
  if(!values) return false;
  const list = Array.isArray(values) ? values : parseList_(values);
  return list.some(isGlobalBaseValue_);
}

function filterSheetsForUser_(names, user){
  if(!user) return names;
  const allowed = Array.isArray(user.bases) ? user.bases : [];
  if(!allowed.length) return names;
  if(hasGlobalBaseAccess_(allowed)) return names;
  const allowedSet = new Set(allowed.map(canonicalBaseKey_).filter(Boolean));
  if(!allowedSet.size) return names;
  return names.filter(name => allowedSet.has(canonicalBaseKey_(name)));
}

function userCanAccessSheet_(user, sheetName){
  if(!user) return false;
  const allowed = Array.isArray(user.bases) ? user.bases : [];
  if(!allowed.length) return true;
  if(hasGlobalBaseAccess_(allowed)) return true;
  const target = canonicalBaseKey_(sheetName);
  if(!target) return false;
  return allowed.some(base => canonicalBaseKey_(base) === target);
}

function findUserByEmail_(email){
  const target = String(email || '').trim().toLowerCase();
  const meta = readUsersWithMeta_();
  const match = target ? meta.users.find(u => u.email === target) || null : null;
  return { sheet: meta.sheet, map: meta.map, user: match, users: meta.users };
}

function findUserById_(userId){
  const target = String(userId || '').trim();
  const meta = readUsersWithMeta_();
  const match = target ? meta.users.find(u => u.userId === target) || null : null;
  return { sheet: meta.sheet, map: meta.map, user: match, users: meta.users };
}

function getRowValue_(row, map, header){
  const idx = getColumnIndex_(map, header);
  if(idx === undefined) return '';
  return row[idx];
}

function setRowValue_(row, map, header, value){
  const idx = getColumnIndex_(map, header);
  if(idx === undefined) return;
  row[idx] = value;
}

function createToken_(user){
  const now = Date.now();
  const payload = {
    uid: user.userId,
    email: user.email,
    name: user.name,
    role: user.role,
    planteles: user.planteles || [],
    bases: user.bases || [],
    iat: now,
    exp: now + TOKEN_EXP_MINUTES * 60 * 1000,
    ver: user.tokenVersion || '',
    active: user.active !== false
  };
  const payloadJson = JSON.stringify(payload);
  const payloadBytes = Utilities.newBlob(payloadJson, 'application/json').getBytes();
  const payloadB64 = Utilities.base64EncodeWebSafe(payloadBytes);
  const signatureBytes = Utilities.computeHmacSha256Signature(payloadB64, getTokenSecret_());
  const signature = Utilities.base64EncodeWebSafe(signatureBytes);
  return `${payloadB64}.${signature}`;
}

function verifyToken_(token){
  if(!token) return { ok:false, message:'Token faltante' };
  const parts = token.split('.');
  if(parts.length !== 2) return { ok:false, message:'Token inválido' };
  const payloadBytes = Utilities.base64DecodeWebSafe(parts[0]);
  let payloadJson = '';
  try{
    payloadJson = Utilities.newBlob(payloadBytes).getDataAsString();
  }catch(err){
    return { ok:false, message:'Token no válido' };
  }
  let payload;
  try{
    payload = JSON.parse(payloadJson);
  }catch(err){
    return { ok:false, message:'Token dañado' };
  }
  const expectedBytes = Utilities.computeHmacSha256Signature(parts[0], getTokenSecret_());
  const expectedSig = Utilities.base64EncodeWebSafe(expectedBytes);
  if(expectedSig !== parts[1]){
    return { ok:false, message:'Token no válido' };
  }
  if(payload.exp && Date.now() > Number(payload.exp || 0)){
    return { ok:false, message:'Token expirado' };
  }
  let lookup;
  try{
    lookup = findUserById_(payload.uid);
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return { ok:false, message };
  }
  const user = lookup.user;
  if(!user){
    return { ok:false, message:'Usuario no encontrado' };
  }
  if(user.active === false){
    return { ok:false, message:'Usuario inactivo' };
  }
  if((user.tokenVersion || '') && payload.ver && user.tokenVersion !== payload.ver){
    return { ok:false, message:'Token desactualizado' };
  }
  const sanitized = sanitizeUserForClient_(user);
  return { ok:true, user: sanitized };
}

function handleLogin_(e, body){
  const email = String(body?.email || '').trim().toLowerCase();
  const password = String(body?.password || '');
  if(!email || !password){
    return jsonResponse({ error:'Correo y contraseña obligatorios.' }, e);
  }
  let lookup;
  try{
    lookup = findUserByEmail_(email);
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const user = lookup.user;
  if(!user || !user.passwordHash){
    return jsonResponse({ error:'Credenciales inválidas.' }, e);
  }
  if(user.active === false){
    return jsonResponse({ error:'Tu usuario está inactivo.' }, e);
  }
  if(!verifyPassword_(password, user.passwordHash, user.salt)){
    return jsonResponse({ error:'Credenciales inválidas.' }, e);
  }
  const sheet = lookup.sheet;
  const map = lookup.map;
  const rowIndex = user.rowIndex;
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  const now = new Date();
  const timestamp = formatDateTime_(now);
  setRowValue_(row, map, 'UltimoIngreso', timestamp);
  setRowValue_(row, map, 'ActualizadoEl', timestamp);
  if(!user.tokenVersion){
    user.tokenVersion = generateTokenVersion_();
    setRowValue_(row, map, 'TokenVersion', user.tokenVersion);
  }
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  user.lastLogin = timestamp;
  user.updatedAt = timestamp;
  const token = createToken_(user);
  const sanitized = sanitizeUserForClient_(user);
  sanitized.lastLogin = timestamp;
  sanitized.updatedAt = timestamp;
  return jsonResponse({ ok:true, token, user: sanitized }, e);
}

function handlePasswordResetRequest_(e, body){
  const email = String(body?.email || '').trim().toLowerCase();
  const userId = String(body?.userId || '').trim();
  if(!email || !userId){
    return jsonResponse({ error:'Correo e ID de usuario obligatorios.' }, e);
  }
  let lookup;
  try{
    lookup = findUserByEmail_(email);
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const user = lookup.user;
  if(!user || String(user.userId || '') !== userId){
    return jsonResponse({ error:'No encontramos un usuario activo con esos datos.' }, e);
  }
  if(user.active === false){
    return jsonResponse({ error:'Tu usuario está inactivo. Contacta al administrador.' }, e);
  }
  const sheet = lookup.sheet;
  const map = lookup.map;
  const rowIndex = user.rowIndex;
  if(!sheet || !map || !rowIndex){
    return jsonResponse({ error:'No se pudo restablecer la contraseña.' }, e);
  }
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  const tempPassword = generateTemporaryPassword_();
  user.salt = generateSalt_();
  user.passwordHash = hashPassword_(tempPassword, user.salt);
  user.tokenVersion = generateTokenVersion_();
  const timestamp = formatDateTime_(new Date());
  setRowValue_(row, map, 'PasswordHash', user.passwordHash);
  setRowValue_(row, map, 'Salt', user.salt);
  setRowValue_(row, map, 'TokenVersion', user.tokenVersion);
  setRowValue_(row, map, 'ActualizadoEl', timestamp);
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  user.updatedAt = timestamp;
  const sanitized = sanitizeUserForClient_(user);
  sanitized.updatedAt = timestamp;
  return jsonResponse({ ok:true, tempPassword, message:'Contraseña temporal generada.', user: sanitized }, e);
}

function getAuthDiagnostics_(){
  try{
    const meta = readUsersWithMeta_();
    const count = meta.users.length;
    return {
      enabled: true,
      ok: count > 0,
      users: count,
      message: count ? 'Activo' : 'Sin usuarios registrados'
    };
  }catch(err){
    return { enabled: false, ok: false, message: err.message || 'No disponible' };
  }
}

function handleListUsers_(e, user){
  try{
    const meta = readUsersWithMeta_();
    const list = meta.users.map(u => sanitizeUserForClient_(u));
    return jsonResponse({ users: list }, e);
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
}

function handleListTeams_(e, user){
  try{
    const meta = readTeamsWithMeta_();
    const list = meta.teams.map(team => sanitizeTeamForClient_(team));
    return jsonResponse({ teams: list }, e);
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
}

function handleCreateUser_(e, body, requester){
  const email = String(body?.email || '').trim().toLowerCase();
  const userId = String(body?.userId || '').trim();
  const name = String(body?.name || '').trim();
  const role = normalizeRole_(body?.role);
  const planteles = parseList_(body?.planteles);
  const bases = parseList_(body?.bases);
  const password = String(body?.password || '').trim();
  const activeField = body?.active;
  const active = activeField === undefined ? true : normalizeBoolean_(activeField);
  if(!email || !userId || !name){
    return jsonResponse({ error:'Correo, ID y nombre son obligatorios.' }, e);
  }
  if(!password){
    return jsonResponse({ error:'La contraseña es obligatoria.' }, e);
  }
  let meta;
  try{
    meta = readUsersWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  if(meta.users.some(u => u.email === email)){
    return jsonResponse({ error:'El correo ya está registrado.' }, e);
  }
  if(meta.users.some(u => u.userId === userId)){
    return jsonResponse({ error:'El ID ya está registrado.' }, e);
  }
  const salt = generateSalt_();
  const hash = hashPassword_(password, salt);
  const now = new Date();
  const createdAt = formatDateTime_(now);
  const sheet = meta.sheet;
  const map = meta.map;
  const totalCols = sheet.getLastColumn();
  const rowValues = new Array(totalCols).fill('');
  setRowValue_(rowValues, map, 'Email', email);
  setRowValue_(rowValues, map, 'UserID', userId);
  setRowValue_(rowValues, map, 'Nombre', name);
  setRowValue_(rowValues, map, 'Rol', role);
  setRowValue_(rowValues, map, 'Planteles', joinList_(planteles));
  setRowValue_(rowValues, map, 'Bases', joinList_(bases));
  setRowValue_(rowValues, map, 'Activo', active ? 'TRUE' : 'FALSE');
  setRowValue_(rowValues, map, 'PasswordHash', hash);
  setRowValue_(rowValues, map, 'Salt', salt);
  setRowValue_(rowValues, map, 'CreadoEl', createdAt);
  setRowValue_(rowValues, map, 'ActualizadoEl', createdAt);
  const tokenVersion = generateTokenVersion_();
  setRowValue_(rowValues, map, 'TokenVersion', tokenVersion);
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, totalCols).setValues([rowValues]);
  const savedRow = sheet.getRange(nextRow, 1, 1, totalCols).getValues()[0];
  const savedUser = mapUserRow_(savedRow, map);
  savedUser.rowIndex = nextRow;
  savedUser.tokenVersion = tokenVersion;
  const sanitized = sanitizeUserForClient_(savedUser);
  return jsonResponse({ ok:true, user: sanitized }, e);
}

function handleUpdateUser_(e, body, requester){
  const targetId = String(body?.userId || '').trim();
  if(!targetId){
    return jsonResponse({ error:'Falta el ID del usuario.' }, e);
  }
  let meta;
  try{
    meta = readUsersWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const current = meta.users.find(u => u.userId === targetId);
  if(!current){
    return jsonResponse({ error:'Usuario no encontrado.' }, e);
  }
  const sheet = meta.sheet;
  const map = meta.map;
  const rowIndex = current.rowIndex;
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  if(body.email !== undefined){
    const email = String(body.email || '').trim().toLowerCase();
    if(!email) return jsonResponse({ error:'El correo no puede quedar vacío.' }, e);
    if(email !== current.email && meta.users.some(u => u.email === email)){
      return jsonResponse({ error:'El correo ya está registrado.' }, e);
    }
    current.email = email;
  }
  if(body.name !== undefined){
    const name = String(body.name || '').trim();
    if(!name) return jsonResponse({ error:'El nombre no puede quedar vacío.' }, e);
    current.name = name;
  }
  if(body.role !== undefined){
    current.role = normalizeRole_(body.role);
  }
  if(body.planteles !== undefined){
    current.planteles = parseList_(body.planteles);
  }
  if(body.bases !== undefined){
    current.bases = parseList_(body.bases);
  }
  if(body.active !== undefined){
    current.active = normalizeBoolean_(body.active);
  }
  let passwordChanged = false;
  if(body.password){
    const password = String(body.password || '').trim();
    if(!password){
      return jsonResponse({ error:'La contraseña no puede quedar vacía.' }, e);
    }
    current.salt = generateSalt_();
    current.passwordHash = hashPassword_(password, current.salt);
    current.tokenVersion = generateTokenVersion_();
    passwordChanged = true;
  }
  setRowValue_(row, map, 'Email', current.email);
  setRowValue_(row, map, 'Nombre', current.name);
  setRowValue_(row, map, 'Rol', current.role);
  setRowValue_(row, map, 'Planteles', joinList_(current.planteles));
  setRowValue_(row, map, 'Bases', joinList_(current.bases));
  setRowValue_(row, map, 'Activo', current.active ? 'TRUE' : 'FALSE');
  if(passwordChanged){
    setRowValue_(row, map, 'PasswordHash', current.passwordHash);
    setRowValue_(row, map, 'Salt', current.salt);
    setRowValue_(row, map, 'TokenVersion', current.tokenVersion);
  }else if(!current.tokenVersion){
    current.tokenVersion = generateTokenVersion_();
    setRowValue_(row, map, 'TokenVersion', current.tokenVersion);
  }
  const updatedAt = formatDateTime_(new Date());
  setRowValue_(row, map, 'ActualizadoEl', updatedAt);
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  const refreshed = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  const updatedUser = mapUserRow_(refreshed, map);
  updatedUser.rowIndex = rowIndex;
  const sanitized = sanitizeUserForClient_(updatedUser);
  return jsonResponse({ ok:true, user: sanitized }, e);
}

function handleDeleteUser_(e, body, requester){
  const targetId = String(body?.userId || '').trim();
  if(!targetId){
    return jsonResponse({ error:'Falta el ID del usuario.' }, e);
  }
  let meta;
  try{
    meta = readUsersWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const current = meta.users.find(u => u.userId === targetId);
  if(!current){
    return jsonResponse({ error:'Usuario no encontrado.' }, e);
  }
  const sheet = meta.sheet;
  const map = meta.map;
  const rowIndex = current.rowIndex;
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  setRowValue_(row, map, 'Activo', 'FALSE');
  const tokenVersion = generateTokenVersion_();
  setRowValue_(row, map, 'TokenVersion', tokenVersion);
  setRowValue_(row, map, 'ActualizadoEl', formatDateTime_(new Date()));
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  const updatedUser = mapUserRow_(row, map);
  updatedUser.rowIndex = rowIndex;
  const sanitized = sanitizeUserForClient_(updatedUser);
  return jsonResponse({ ok:true, user: sanitized }, e);
}

function handleCreateTeam_(e, body, requester){
  const rawId = String(body?.teamId || '').trim();
  const name = String(body?.name || '').trim();
  const description = String(body?.description || '').trim();
  const planteles = parseList_(body?.planteles);
  const bases = parseList_(body?.bases);
  const members = parseList_(body?.members);
  const activeField = body?.active;
  const active = activeField === undefined ? true : normalizeBoolean_(activeField);
  const teamId = sanitizeTeamId_(rawId);
  if(!teamId){
    return jsonResponse({ error:'El ID del equipo es obligatorio.' }, e);
  }
  if(!name){
    return jsonResponse({ error:'El nombre del equipo es obligatorio.' }, e);
  }
  let meta;
  try{
    meta = readTeamsWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  if(meta.teams.some(team => team.teamId === teamId)){
    return jsonResponse({ error:'El ID del equipo ya está registrado.' }, e);
  }
  const sheet = meta.sheet;
  const map = meta.map;
  const totalCols = sheet.getLastColumn();
  const rowValues = new Array(totalCols).fill('');
  setRowValue_(rowValues, map, 'TeamID', teamId);
  setRowValue_(rowValues, map, 'Nombre', name);
  setRowValue_(rowValues, map, 'Descripcion', description);
  setRowValue_(rowValues, map, 'Planteles', joinList_(planteles));
  setRowValue_(rowValues, map, 'Bases', joinList_(bases));
  setRowValue_(rowValues, map, 'Miembros', joinList_(members));
  setRowValue_(rowValues, map, 'Activo', active ? 'TRUE' : 'FALSE');
  const now = formatDateTime_(new Date());
  setRowValue_(rowValues, map, 'CreadoEl', now);
  setRowValue_(rowValues, map, 'ActualizadoEl', now);
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, totalCols).setValues([rowValues]);
  const savedRow = sheet.getRange(nextRow, 1, 1, totalCols).getValues()[0];
  const savedTeam = mapTeamRow_(savedRow, map);
  savedTeam.rowIndex = nextRow;
  const sanitized = sanitizeTeamForClient_(savedTeam);
  return jsonResponse({ ok:true, team: sanitized }, e);
}

function handleUpdateTeam_(e, body, requester){
  const targetId = sanitizeTeamId_(body?.teamId);
  if(!targetId){
    return jsonResponse({ error:'Falta el ID del equipo.' }, e);
  }
  let meta;
  try{
    meta = readTeamsWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const current = meta.teams.find(team => team.teamId === targetId);
  if(!current){
    return jsonResponse({ error:'Equipo no encontrado.' }, e);
  }
  const sheet = meta.sheet;
  const map = meta.map;
  const rowIndex = current.rowIndex;
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  if(body.name !== undefined){
    const name = String(body.name || '').trim();
    if(!name) return jsonResponse({ error:'El nombre no puede quedar vacío.' }, e);
    current.name = name;
  }
  if(body.description !== undefined){
    current.description = String(body.description || '').trim();
  }
  if(body.planteles !== undefined){
    current.planteles = parseList_(body.planteles);
  }
  if(body.bases !== undefined){
    current.bases = parseList_(body.bases);
  }
  if(body.members !== undefined){
    current.members = parseList_(body.members);
  }
  if(body.active !== undefined){
    current.active = normalizeBoolean_(body.active);
  }
  setRowValue_(row, map, 'Nombre', current.name);
  setRowValue_(row, map, 'Descripcion', current.description);
  setRowValue_(row, map, 'Planteles', joinList_(current.planteles));
  setRowValue_(row, map, 'Bases', joinList_(current.bases));
  setRowValue_(row, map, 'Miembros', joinList_(current.members));
  setRowValue_(row, map, 'Activo', current.active ? 'TRUE' : 'FALSE');
  const updatedAt = formatDateTime_(new Date());
  setRowValue_(row, map, 'ActualizadoEl', updatedAt);
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  const refreshed = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  const updatedTeam = mapTeamRow_(refreshed, map);
  updatedTeam.rowIndex = rowIndex;
  const sanitized = sanitizeTeamForClient_(updatedTeam);
  return jsonResponse({ ok:true, team: sanitized }, e);
}

function handleDeleteTeam_(e, body, requester){
  const targetId = sanitizeTeamId_(body?.teamId);
  if(!targetId){
    return jsonResponse({ error:'Falta el ID del equipo.' }, e);
  }
  let meta;
  try{
    meta = readTeamsWithMeta_();
  }catch(err){
    const message = err && err.message ? err.message : SPREADSHEET_ERROR_MESSAGE;
    return jsonResponse({ error: message }, e);
  }
  const current = meta.teams.find(team => team.teamId === targetId);
  if(!current){
    return jsonResponse({ error:'Equipo no encontrado.' }, e);
  }
  const sheet = meta.sheet;
  const map = meta.map;
  const rowIndex = current.rowIndex;
  const totalCols = sheet.getLastColumn();
  const row = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  current.active = false;
  setRowValue_(row, map, 'Activo', 'FALSE');
  const updatedAt = formatDateTime_(new Date());
  setRowValue_(row, map, 'ActualizadoEl', updatedAt);
  sheet.getRange(rowIndex, 1, 1, totalCols).setValues([row]);
  const refreshed = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
  const updatedTeam = mapTeamRow_(refreshed, map);
  updatedTeam.rowIndex = rowIndex;
  const sanitized = sanitizeTeamForClient_(updatedTeam);
  return jsonResponse({ ok:true, team: sanitized }, e);
}

function addLeadIndexEntry_(store, key, ref){
  if(!key) return;
  if(!store.has(key)) store.set(key, []);
  const list = store.get(key);
  if(list.indexOf(ref) === -1) list.push(ref);
}

function removeLeadIndexEntry_(store, key, ref){
  if(!key || !store.has(key)) return;
  const list = store.get(key);
  const filtered = list.filter(item => item !== ref);
  if(filtered.length){
    store.set(key, filtered);
  }else{
    store.delete(key);
  }
}

function collectLeadKeysForRow_(row, map){
  const idKey = normalizeIdentifierKey_(getRowValue_(row, map, 'ID'));
  const matriculaKey = normalizeIdentifierKey_(getRowValue_(row, map, 'Matricula'));
  const emailKey = normalizeEmailKey_(getRowValue_(row, map, 'Correo'));
  const phoneCandidates = [
    getRowValue_(row, map, 'Teléfono'),
    getRowValue_(row, map, 'Telefono'),
    getRowValue_(row, map, 'Telefono normalizado'),
    getRowValue_(row, map, 'TelefonoNormalizado')
  ];
  let phoneKey = '';
  for(let i=0;i<phoneCandidates.length;i++){
    const key = normalizePhoneKey_(phoneCandidates[i]);
    if(key){
      phoneKey = key;
      break;
    }
  }
  return { id: idKey, matricula: matriculaKey, email: emailKey, phone: phoneKey };
}

function buildSheetLeadIndex_(values, map){
  const indexes = {
    byId: new Map(),
    byMatricula: new Map(),
    byEmail: new Map(),
    byPhone: new Map()
  };
  values.forEach((row, index) => {
    if(!row) return;
    const ref = { row, rowNumber: index + 2, index };
    const keys = collectLeadKeysForRow_(row, map);
    if(keys.id) addLeadIndexEntry_(indexes.byId, keys.id, ref);
    if(keys.matricula) addLeadIndexEntry_(indexes.byMatricula, keys.matricula, ref);
    if(keys.email) addLeadIndexEntry_(indexes.byEmail, keys.email, ref);
    if(keys.phone) addLeadIndexEntry_(indexes.byPhone, keys.phone, ref);
  });
  return indexes;
}

function syncIndexForKey_(store, oldKey, newKey, ref){
  if(oldKey && (!newKey || oldKey !== newKey)){
    removeLeadIndexEntry_(store, oldKey, ref);
  }
  if(newKey){
    addLeadIndexEntry_(store, newKey, ref);
  }
}

function updateLeadIndexesForRef_(indexes, oldKeys, newKeys, ref){
  syncIndexForKey_(indexes.byId, oldKeys.id, newKeys.id, ref);
  syncIndexForKey_(indexes.byMatricula, oldKeys.matricula, newKeys.matricula, ref);
  syncIndexForKey_(indexes.byEmail, oldKeys.email, newKeys.email, ref);
  syncIndexForKey_(indexes.byPhone, oldKeys.phone, newKeys.phone, ref);
}

function findLeadReferenceForSync_(entry, indexes){
  if(!entry || typeof entry !== 'object'){
    return { error:true, reason:'Registro inválido.' };
  }
  const attempts = [];
  const idKey = entry.idKey || normalizeIdentifierKey_(entry.id || '');
  if(idKey) attempts.push({ key: idKey, map: indexes.byId, label: 'ID' });
  const matriculaKey = entry.matriculaKey || normalizeIdentifierKey_(entry.matricula || (entry.updates && entry.updates.matricula) || '');
  if(matriculaKey) attempts.push({ key: matriculaKey, map: indexes.byMatricula, label: 'Matrícula' });
  const phoneKey = entry.phoneKey || normalizePhoneKey_((entry.updates && entry.updates.telefono) || '');
  if(phoneKey) attempts.push({ key: phoneKey, map: indexes.byPhone, label: 'Teléfono' });
  const seen = new Set();
  for(let i=0;i<attempts.length;i++){
    const attempt = attempts[i];
    if(!attempt.key || seen.has(attempt.key)) continue;
    seen.add(attempt.key);
    const list = attempt.map.get(attempt.key) || [];
    if(!list.length) continue;
    if(list.length > 1){
      return { error:true, reason: attempt.label + ' coincide con múltiples registros.' };
    }
    return { ref: list[0] };
  }
  return { error:true, reason:'No se encontró coincidencia en la base.' };
}

function handleSyncActivos_(e, body, user){
  const entries = Array.isArray(body?.rows) ? body.rows.filter(row => row && typeof row === 'object') : [];
  if(!entries.length){
    return jsonResponse({ ok:false, error:'No se recibieron coincidencias para actualizar.' }, e);
  }
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ ok:false, error: SPREADSHEET_ERROR_MESSAGE }, e);
  const result = { ok:true, updated:0, unmatched:[] };
  const grouped = new Map();
  entries.forEach(entry => {
    const sheetName = String(entry?.sheet || '').trim();
    if(!sheetName){
      result.unmatched.push({ sheet:'', id: String(entry?.id || ''), matricula: String(entry?.updates?.matricula || ''), reason:'Hoja no especificada.' });
      return;
    }
    if(user && !userCanAccessSheet_(user, sheetName)){
      result.unmatched.push({ sheet: sheetName, id: String(entry?.id || ''), matricula: String(entry?.updates?.matricula || ''), reason:'No tienes permiso para actualizar esta base.' });
      return;
    }
    if(!grouped.has(sheetName)) grouped.set(sheetName, []);
    grouped.get(sheetName).push(entry);
  });
  grouped.forEach((items, sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if(!sheet){
      items.forEach(entry => {
        result.unmatched.push({ sheet: sheetName, id: String(entry?.id || ''), matricula: String(entry?.updates?.matricula || ''), reason:'La hoja no existe.' });
      });
      return;
    }
    const { headers, map } = getColumnMap_(sheet);
    const totalCols = sheet.getLastColumn();
    const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
    const values = dataRowCount > 0 ? sheet.getRange(2, 1, dataRowCount, headers.length).getValues() : [];
    const indexes = buildSheetLeadIndex_(values, map);
    const pending = new Map();
    items.forEach(entry => {
      const updates = entry?.updates || {};
      const meaningfulUpdates = Object.keys(updates).filter(key => String(updates[key] || '').trim());
      if(!meaningfulUpdates.length){
        result.unmatched.push({ sheet: sheetName, id: String(entry?.id || ''), matricula: String(entry?.updates?.matricula || ''), reason:'Sin campos para actualizar.' });
        return;
      }
      const match = findLeadReferenceForSync_(entry, indexes);
      if(match.error || !match.ref){
        result.unmatched.push({ sheet: sheetName, id: String(entry?.id || ''), matricula: String(entry?.updates?.matricula || ''), reason: match.reason || 'No se encontró coincidencia en la base.' });
        return;
      }
      const ref = match.ref;
      const row = ref.row;
      const oldKeys = collectLeadKeysForRow_(row, map);
      if(updates.nombre){
        setRowValue_(row, map, 'Nombre', String(updates.nombre).trim());
      }
      if(updates.matricula){
        const matricula = String(updates.matricula).trim();
        setRowValue_(row, map, 'Matricula', matricula);
      }
      if(updates.telefono){
        const telefono = String(updates.telefono).trim();
        setRowValue_(row, map, 'Teléfono', telefono);
        setRowValue_(row, map, 'Telefono', telefono);
      }
      if(updates.correo){
        setRowValue_(row, map, 'Correo', String(updates.correo).trim());
      }
      if(updates.campus){
        setRowValue_(row, map, 'Plantel', String(updates.campus).trim());
      }
      if(updates.modalidad){
        setRowValue_(row, map, 'Modalidad', String(updates.modalidad).trim());
      }
      if(updates.programa){
        setRowValue_(row, map, 'Programa', String(updates.programa).trim());
      }
      setRowValue_(row, map, 'Etapa', 'Inscrito');
      setRowValue_(row, map, 'Estado', 'Inscrito');
      const resolution = computeResolutionForLead_(sheetName, 'Inscrito', 'Inscrito');
      if(resolution){
        setRowValue_(row, map, 'Resolución', resolution);
        setRowValue_(row, map, 'Resolucion', resolution);
      }
      const finalNombre = String(getRowValue_(row, map, 'Nombre') || '').trim();
      const finalTelefono = String(getRowValue_(row, map, 'Teléfono') || getRowValue_(row, map, 'Telefono') || '').trim();
      const finalCorreo = String(getRowValue_(row, map, 'Correo') || '').trim();
      applyContactLinksToRow_(row, map, finalTelefono, finalCorreo, finalNombre);
      const timestamp = formatDateTime_(new Date());
      setRowValue_(row, map, 'ActualizadoEl', timestamp);
      const newKeys = collectLeadKeysForRow_(row, map);
      updateLeadIndexesForRef_(indexes, oldKeys, newKeys, ref);
      pending.set(ref.rowNumber, row);
      result.updated++;
    });
    pending.forEach((rowValues, rowNumber) => {
      sheet.getRange(rowNumber, 1, 1, totalCols).setValues([rowValues]);
    });
  });
  if(result.updated && grouped.size){
    result.message = `Registros actualizados: ${result.updated}${result.unmatched.length ? ` · No encontrados: ${result.unmatched.length}` : ''}`;
  }else if(!result.updated && !result.unmatched.length){
    result.message = 'No se aplicaron cambios.';
  }
  return jsonResponse(result, e);
}

function handleImportLeads_(e, body, user){
  const sheetName = String(body?.sheet || '').trim();
  if(!sheetName){
    return jsonResponse({ ok:false, error:'Falta la base de destino.' }, e);
  }
  if(user && !userCanAccessSheet_(user, sheetName)){
    return jsonResponse({ ok:false, error:'No tienes permiso para importar en esta base.' }, e);
  }
  const rows = Array.isArray(body?.rows) ? body.rows.filter(row => row && typeof row === 'object') : [];
  if(!rows.length){
    return jsonResponse({ ok:false, error:'No se recibieron registros para importar.' }, e);
  }
  const ss = getSpreadsheet_();
  if(!ss) return jsonResponse({ ok:false, error: SPREADSHEET_ERROR_MESSAGE }, e);
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet){
    return jsonResponse({ ok:false, error:'La hoja "' + sheetName + '" no existe.' }, e);
  }
  const { headers, map } = getColumnMap_(sheet);
  const totalCols = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const dataRowCount = Math.max(0, lastRow - 1);
  const existingValues = dataRowCount > 0 ? sheet.getRange(2, 1, dataRowCount, headers.length).getValues() : [];
  const indexes = buildSheetLeadIndex_(existingValues, map);
  const seen = { id:new Set(), matricula:new Set(), email:new Set(), phone:new Set() };
  const preparedRows = [];
  const needsAssignment = [];
  const skipped = [];
  rows.forEach((rawRow, idx) => {
    const data = rawRow && typeof rawRow === 'object' ? rawRow : {};
    const idValue = String(data.id || '').trim();
    const nombre = String(data.nombre || '').trim();
    const matriculaValue = String(data.matricula || '').trim();
    const telefonoValue = String(data.telefono || '').trim();
    const correoValue = String(data.correo || '').trim();
    const campus = String(data.campus || '').trim();
    const modalidad = String(data.modalidad || '').trim();
    const programa = String(data.programa || '').trim();
    const etapaValue = String(data.etapa || '').trim() || 'Nuevo';
    const estadoValue = String(data.estado || '').trim();
    const asesorValue = String(data.asesor || '').trim();
    const comentarioValue = String(data.comentario || '').trim();
    let asignacionValue = String(data.asignacion || '').trim();
    const metadataValue = coerceMetadataValue_(data.metadatos);
    const idKey = normalizeIdentifierKey_(idValue);
    const matriculaKey = normalizeIdentifierKey_(matriculaValue);
    const emailKey = normalizeEmailKey_(correoValue);
    const phoneKey = normalizePhoneKey_(telefonoValue);
    const reasons = [];
    if(idKey){
      if(indexes.byId.has(idKey)) reasons.push('ID existente en la base');
      if(seen.id.has(idKey)) reasons.push('ID duplicado en la importación');
    }
    if(matriculaKey){
      if(indexes.byMatricula.has(matriculaKey)) reasons.push('Matrícula existente en la base');
      if(seen.matricula.has(matriculaKey)) reasons.push('Matrícula duplicada en la importación');
    }
    if(emailKey){
      if(indexes.byEmail.has(emailKey)) reasons.push('Correo existente en la base');
      if(seen.email.has(emailKey)) reasons.push('Correo duplicado en la importación');
    }
    if(phoneKey){
      if(indexes.byPhone.has(phoneKey)) reasons.push('Teléfono existente en la base');
      if(seen.phone.has(phoneKey)) reasons.push('Teléfono duplicado en la importación');
    }
    if(reasons.length){
      skipped.push({
        row: idx + 1,
        id: idValue,
        matricula: matriculaValue,
        telefono: telefonoValue,
        correo: correoValue,
        motivo: reasons.join(' · ')
      });
      return;
    }
    if(idKey) seen.id.add(idKey);
    if(matriculaKey) seen.matricula.add(matriculaKey);
    if(emailKey) seen.email.add(emailKey);
    if(phoneKey) seen.phone.add(phoneKey);
    const leadId = idValue || generateLeadId_();
    if(!asignacionValue){
      asignacionValue = formatDateTime_(new Date());
    }
    const timestamp = formatDateTime_(new Date());
    const resolution = computeResolutionForLead_(sheet.getName(), etapaValue, estadoValue || etapaValue);
    const prepared = {
      leadId,
      nombre,
      matriculaValue,
      correoValue,
      telefonoValue,
      campus,
      modalidad,
      programa,
      etapaValue,
      estadoValue,
      asesorValue,
      comentarioValue,
      metadataValue,
      asignacionValue,
      resolution,
      createdAt: timestamp,
      updatedAt: timestamp
    };
    if(!asesorValue){
      needsAssignment.push(prepared);
    }
    preparedRows.push(prepared);
  });
  if(!preparedRows.length){
    const message = skipped.length ? 'No se importó ningún registro por incidencias detectadas.' : 'No hay registros válidos para importar.';
    return jsonResponse({ ok:false, error: message, skipped }, e);
  }
  if(needsAssignment.length){
    assignAsesoresRoundRobin_(needsAssignment, sheetName);
  }
  const toInsert = preparedRows.map(row => {
    const rowValues = new Array(totalCols).fill('');
    setRowValue_(rowValues, map, 'ID', row.leadId);
    setRowValue_(rowValues, map, 'Nombre', row.nombre);
    setRowValue_(rowValues, map, 'Matricula', row.matriculaValue);
    setRowValue_(rowValues, map, 'Correo', row.correoValue);
    setRowValue_(rowValues, map, 'Teléfono', row.telefonoValue);
    setRowValue_(rowValues, map, 'Telefono', row.telefonoValue);
    setRowValue_(rowValues, map, 'Plantel', row.campus);
    setRowValue_(rowValues, map, 'Modalidad', row.modalidad);
    setRowValue_(rowValues, map, 'Programa', row.programa);
    setRowValue_(rowValues, map, 'Etapa', row.etapaValue);
    if(row.estadoValue){
      setRowValue_(rowValues, map, 'Estado', row.estadoValue);
    }
    if(row.asesorValue){
      setRowValue_(rowValues, map, 'Asesor', row.asesorValue);
    }
    if(row.comentarioValue){
      setRowValue_(rowValues, map, 'Comentario', row.comentarioValue);
    }
    if(row.metadataValue){
      setRowValue_(rowValues, map, 'Metadatos', row.metadataValue);
    }
    setRowValue_(rowValues, map, 'Asignación', row.asignacionValue);
    setRowValue_(rowValues, map, 'Asignacion', row.asignacionValue);
    if(row.resolution){
      setRowValue_(rowValues, map, 'Resolución', row.resolution);
      setRowValue_(rowValues, map, 'Resolucion', row.resolution);
    }
    applyContactLinksToRow_(rowValues, map, row.telefonoValue, row.correoValue, row.nombre);
    setRowValue_(rowValues, map, 'CreadoEl', row.createdAt);
    setRowValue_(rowValues, map, 'ActualizadoEl', row.updatedAt);
    return rowValues;
  });
  const startRow = lastRow + 1;
  sheet.getRange(startRow, 1, toInsert.length, totalCols).setValues(toInsert);
  const messageParts = [`Registros importados: ${toInsert.length}`];
  if(skipped.length){
    messageParts.push(`Omitidos: ${skipped.length}`);
  }
  const response = { ok:true, inserted: toInsert.length, skipped, message: messageParts.join(' · ') };
  return jsonResponse(response, e);
}
