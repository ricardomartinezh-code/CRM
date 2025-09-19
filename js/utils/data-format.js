(function (global, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    global.ReLeadDataFormat = factory();
  }
})(typeof self !== 'undefined' ? self : globalThis, function () {
  const cleanValue = (val) => {
    if (val === undefined || val === null) return '';
    if (typeof val === 'string') return val.trim();
    return String(val).trim();
  };

  function parseNumericValue(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
    if (Array.isArray(value)) return value.length;
    if (typeof value === 'object') {
      if ('total' in value && value.total !== value) {
        const total = parseNumericValue(value.total);
        if (Number.isFinite(total)) return total;
      }
      if ('count' in value && value.count !== value) {
        const count = parseNumericValue(value.count);
        if (Number.isFinite(count)) return count;
      }
      if ('value' in value && value.value !== value) {
        const nested = parseNumericValue(value.value);
        if (Number.isFinite(nested)) return nested;
      }
      if ('valor' in value && value.valor !== value) {
        const nested = parseNumericValue(value.valor);
        if (Number.isFinite(nested)) return nested;
      }
      if (Array.isArray(value.registros)) return value.registros.length;
    }
    const text = cleanValue(value);
    if (!text) return NaN;
    const match = text.match(/-?\d+(?:[.,]\d+)?/);
    if (!match) return NaN;
    const normalized = match[0].replace(',', '.');
    const num = Number(normalized);
    return Number.isFinite(num) ? num : NaN;
  }

  function formatLosgDateTime(value) {
    if (value === null || value === undefined || value === '') return '';
    if (Array.isArray(value)) {
      const first = value.find((item) => item !== undefined && item !== null);
      return first !== undefined ? formatLosgDateTime(first) : '';
    }
    if (value instanceof Date) {
      if (Number.isNaN(value.valueOf())) return '';
      return value.toISOString().replace('T', ' ').slice(0, 16);
    }
    if (typeof value === 'object') {
      const candidate = [
        'fecha',
        'date',
        'datetime',
        'timestamp',
        'inicio',
        'start',
        'hora',
        'time',
        'first',
        'last',
      ]
        .map((key) => value[key])
        .find((val) => val !== undefined && val !== null && val !== '');
      if (candidate !== undefined) {
        return formatLosgDateTime(candidate);
      }
    }
    const text = cleanValue(value);
    if (!text) return '';
    const parsed = new Date(text);
    if (Number.isNaN(parsed.valueOf())) return text;
    return parsed.toISOString().replace('T', ' ').slice(0, 16);
  }

  function formatLosgList(value) {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value)) {
      return value
        .map((item) => formatLosgList(item))
        .filter(Boolean)
        .join(' | ');
    }
    if (value instanceof Date) {
      return formatLosgDateTime(value);
    }
    if (typeof value === 'object') {
      const name = cleanValue(
        value.name ||
          value.nombre ||
          value.displayName ||
          value.usuario ||
          value.user ||
          value.title
      );
      const email = cleanValue(value.email || value.correo || value.mail);
      const role = cleanValue(value.role || value.rol || value.perfil);
      const descriptor = cleanValue(
        value.descripcion ||
          value.description ||
          value.detalle ||
          value.detail ||
          value.etiqueta ||
          value.label ||
          value.tipo ||
          value.action ||
          value.accion ||
          value.resumen
      );
      const target = cleanValue(
        value.destino ||
          value.target ||
          value.asignado ||
          value.asignadoA ||
          value.para ||
          value.to ||
          value.recipient
      );
      const parts = [];
      if (descriptor && target) {
        parts.push(`${descriptor} → ${target}`);
      } else {
        if (descriptor) parts.push(descriptor);
        if (target) parts.push(target);
      }
      if (name) parts.push(name);
      if (role) parts.push(role);
      if (email) parts.push(email);
      if (!parts.length) {
        const values = Object.values(value).map(cleanValue).filter(Boolean);
        if (values.length) {
          return values.join(' / ');
        }
      }
      return parts.join(' · ');
    }
    return cleanValue(value);
  }

  function formatLosgAssignments(value) {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value)) {
      return value
        .map((item) => formatLosgAssignments(item))
        .filter(Boolean)
        .join(' | ');
    }
    if (typeof value === 'object') {
      const permiso = cleanValue(
        value.permiso ||
          value.permission ||
          value.tipo ||
          value.scope ||
          value.descripcion ||
          value.description
      );
      const destinatario = cleanValue(
        value.usuario ||
          value.user ||
          value.destino ||
          value.target ||
          value.asignado ||
          value.asignadoA ||
          value.para ||
          value.to ||
          value.nombre
      );
      const detalle = cleanValue(
        value.detalle || value.detail || value.resumen || value.notes || value.resultado
      );
      const parts = [];
      if (permiso) parts.push(permiso);
      if (destinatario) parts.push(`→ ${destinatario}`);
      if (detalle) parts.push(`(${detalle})`);
      const text = parts.filter(Boolean).join(' ');
      if (text) return text.trim();
      return formatLosgList(value);
    }
    return cleanValue(value);
  }

  function formatLosgNumber(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return formatLosgDateTime(value);
    if (typeof value === 'number' && Number.isFinite(value)) return String(value);
    const numeric = parseNumericValue(value);
    if (Number.isFinite(numeric)) return String(numeric);
    if (typeof value === 'object') return formatLosgList(value);
    const text = cleanValue(value);
    if (!text) return '';
    const num = Number(text);
    return Number.isFinite(num) ? String(num) : text;
  }

  function normalizeLosgEntry(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const lowerMap = new Map();
    Object.keys(raw).forEach((key) => {
      lowerMap.set(String(key || '').toLowerCase(), raw[key]);
    });
    const pick = (...keys) => {
      for (const key of keys) {
        if (!key) continue;
        const normalized = String(key).toLowerCase();
        if (lowerMap.has(normalized)) {
          return lowerMap.get(normalized);
        }
      }
      return undefined;
    };
    const rawTotal = pick(
      'accesos',
      'conexiones',
      'logins',
      'login_count',
      'entradas',
      'ingresos',
      'visitas',
      'sesiones'
    );
    const rawMobile = pick(
      'moviles',
      'mobile',
      'mobilelogins',
      'dispositivosmoviles',
      'accesosmoviles'
    );
    const rawDesktop = pick(
      'escritorio',
      'desktop',
      'desktoplogins',
      'dispositivosescritorio',
      'accesosescritorio'
    );
    const totalNumeric = parseNumericValue(rawTotal);
    const mobileNumeric = parseNumericValue(rawMobile);
    const desktopNumeric = parseNumericValue(rawDesktop);
    const horariosRaw = pick('horarios', 'horas', 'times', 'session_times', 'sesiones', 'schedule');
    let deviceSummary = formatLosgList(
      pick(
        'dispositivos',
        'devices',
        'device_types',
        'tipodispositivo',
        'tiposdispositivo',
        'origenes'
      )
    );
    if (!deviceSummary) {
      const pieces = [];
      if (Number.isFinite(mobileNumeric) && mobileNumeric > 0) {
        pieces.push(`Móvil${mobileNumeric > 1 ? ` (${mobileNumeric})` : ''}`);
      }
      if (Number.isFinite(desktopNumeric) && desktopNumeric > 0) {
        pieces.push(`Escritorio${desktopNumeric > 1 ? ` (${desktopNumeric})` : ''}`);
      }
      if (!pieces.length && Number.isFinite(totalNumeric) && totalNumeric > 0) {
        pieces.push(totalNumeric === 1 ? 'Único acceso' : 'Múltiples dispositivos');
      }
      deviceSummary = pieces.join(' · ');
    }
    const normalized = {
      fecha: formatLosgDateTime(pick('fecha', 'date', 'dia', 'day')),
      usuario: formatLosgList(pick('usuario', 'user', 'nombre', 'name', 'displayname')),
      correo: formatLosgList(pick('correo', 'email', 'mail', 'correo_institucional')),
      rol: formatLosgList(pick('rol', 'role', 'perfil')),
      total_accesos: formatLosgNumber(rawTotal),
      primer_acceso: formatLosgDateTime(
        pick(
          'primeracceso',
          'firstlogin',
          'primeravez',
          'first_seen',
          'inicio',
          'primer_ingreso',
          'primeraconexion'
        )
      ),
      ultimo_acceso: formatLosgDateTime(
        pick(
          'ultimoacceso',
          'lastlogin',
          'ultimavez',
          'last_seen',
          'final',
          'ultimo_ingreso',
          'ultimaconexion'
        )
      ),
      accesos_moviles: formatLosgNumber(rawMobile),
      accesos_escritorio: formatLosgNumber(rawDesktop),
      dispositivos: deviceSummary,
      horarios: formatLosgList(horariosRaw),
      importaciones: formatLosgNumber(
        pick('importaciones', 'imports', 'import_count', 'accionesimportacion', 'importactivity')
      ),
      exportaciones: formatLosgNumber(
        pick('exportaciones', 'exports', 'export_count', 'accionesexportacion', 'exportactivity')
      ),
      cambios_masivos: formatLosgNumber(
        pick(
          'cambiosmasivos',
          'massupdates',
          'bulkupdates',
          'bulk',
          'bulk_count',
          'movimientosmasivos',
          'actualizacionesmasivas'
        )
      ),
      diagnosticos: formatLosgNumber(
        pick(
          'diagnosticos',
          'diagnostics',
          'diagnostics_count',
          'ejecucionesdiagnostico',
          'diagnosticosrealizados'
        )
      ),
      resultado_diagnostico: formatLosgList(
        pick(
          'resultadodiagnostico',
          'diagnostics_result',
          'diagnosticosresultado',
          'diagnosticos_resumen',
          'diagnosticosdetalle',
          'diagnosticresult'
        )
      ),
      permisos_asignados: formatLosgAssignments(
        pick(
          'permisosasignados',
          'permisos',
          'permissions',
          'permission_changes',
          'permission_assignments'
        )
      ),
      planteles_asignados: formatLosgAssignments(
        pick('plantelasignado', 'planteles', 'campus', 'planteles_asignados', 'campus_assignments')
      ),
    };
    const hasData = Object.values(normalized).some((value) => cleanValue(value));
    return hasData ? normalized : null;
  }

  return {
    cleanValue,
    parseNumericValue,
    formatLosgDateTime,
    formatLosgList,
    formatLosgAssignments,
    formatLosgNumber,
    normalizeLosgEntry,
  };
});
