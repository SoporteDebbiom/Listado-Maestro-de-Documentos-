
var SYNC = {
  initialized:  false,
  connected:    false,
  onlineCount:  0,
  onlineUsers:  [],
  sessionId:    Date.now().toString(36) + Math.random().toString(36).substr(2, 6),
  _version:     0,
  _isSaving:    false,
  _saveTimer:   null,
  _saveQueue:   false,
  _client:      null,
  _channel:     null,
  _presence:    null,
  _lastSaveAt:  0,
  _initialLoadDone: false,
  _heartbeatTimer: null,
  _vigenciaTimer:  null,
  _backupTimer:    null,
  _lastBackupAt:   0,
  _lastSyncTime:   0
};

var SAVE_DEBOUNCE = 1500;
var HEARTBEAT_INTERVAL = 30000;
var VIGENCIA_CHECK_INTERVAL = 60000;
var AUTO_BACKUP_INTERVAL = 3600000;
async function initSync() {
  var client = typeof getDB === 'function' ? getDB() : null;
  if (!client) {
    return false;
  }

  try {
    SYNC._client = client;

    var result = await SYNC._client
      .from('app_state')
      .select('version')
      .eq('id', 1)
      .single();

    if (result.error) {
      return false;
    }

    SYNC._version    = result.data.version || 0;
    SYNC.initialized = true;
    SYNC.connected   = true;

    _startHeartbeat();
    _startVigenciaCheck();
    _startAutoBackup();

    updateSyncIndicator();
    return true;

  } catch (e) {
    return false;
  }
}
function syncSaveState() {
  localSaveState();
  if (!SYNC.initialized || !SYNC._client) return;
  if (!SYNC._initialLoadDone) {
    return;
  }

  clearTimeout(SYNC._saveTimer);
  SYNC._saveTimer = setTimeout(_doSave, SAVE_DEBOUNCE);
}

async function _doSave() {
  if (SYNC._isSaving) {
    SYNC._saveQueue = true;
    return;
  }
  if (typeof getCsrfToken === 'function' && !getCsrfToken()) {
    return;
  }

  SYNC._isSaving = true;
  _showSyncAnimation(true);

  try {
    var newVersion = Date.now();
    var stateData = {
      records:   STATE.records || [],
      obsoletos: STATE.obsoletos || [],
      papelera:  STATE.papelera || [],
      salidas:   STATE.salidas || [],
      logs:      STATE.logs || [],
      elaboros:  STATE.elaboros || [],
      nextId:    STATE.nextId || 1000
    };
    if (typeof _dataHash === 'function') stateData._hash = _dataHash(stateData);

    var modifiedBy = AUTH.currentUser ? AUTH.currentUser.username : 'Sistema';

    var result = await SYNC._client
      .from('app_state')
      .update({
        state:       stateData,
        version:     newVersion,
        modified_by: modifiedBy,
        session_id:  SYNC.sessionId,
        updated_at:  new Date().toISOString()
      })
      .eq('id', 1);

    if (result.error) {
      SYNC.connected = false;
    } else {
      SYNC._version    = newVersion;
      SYNC._lastSaveAt = Date.now();
      SYNC._lastSyncTime = Date.now();
      SYNC.connected   = true;
      if (typeof _unsavedChanges !== 'undefined') _unsavedChanges = false;
      _updateLastSyncUI();
    }
  } catch (e) {
    SYNC.connected = false;
  }

  updateSyncIndicator();
  _showSyncAnimation(false);
  SYNC._isSaving = false;

  if (SYNC._saveQueue) {
    SYNC._saveQueue = false;
    setTimeout(_doSave, 300);
  }
}

async function forceSync() {
  if (!SYNC.initialized) {
    showToast('Sincronización no disponible (modo local)', 'error');
    return;
  }
  _showSyncAnimation(true);

  await _doSave();
  await _pullRemote(true);

  _showSyncAnimation(false);
  showToast('✅ Sincronizado', 'success');
}
function listenStateChanges() {
  if (!SYNC.initialized || !SYNC._client) return;

  if (SYNC._channel) {
    SYNC._client.removeChannel(SYNC._channel);
    SYNC._channel = null;
  }

  SYNC._channel = SYNC._client
    .channel('db-changes')
    .on('postgres_changes',
      { event: 'UPDATE', schema: 'public', table: 'app_state' },
      function(payload) {
        _handleRealtimeUpdate(payload.new, false);
      }
    )
    .subscribe(function(status) {
      if (status === 'SUBSCRIBED') {
        SYNC.connected = true;
        updateSyncIndicator();
      } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
        SYNC.connected = false;
        updateSyncIndicator();
        setTimeout(function() {
          if (SYNC.initialized && !SYNC.connected) {
            listenStateChanges();
          }
        }, 5000);
      }
    });
}

function _mergeById(local, remote) {
  if (!local || local.length === 0) return remote || [];
  if (!remote || remote.length === 0) return local || [];
  var merged = {};
  remote.forEach(function(r) { if (r && r.id) merged[r.id] = r; });
  local.forEach(function(r) {
    if (r && r.id && !merged[r.id]) merged[r.id] = r;
  });
  return Object.values(merged);
}

function _mergeArrays(local, remote) {
  var set = {};
  (local || []).forEach(function(v) { set[v] = true; });
  (remote || []).forEach(function(v) { set[v] = true; });
  return Object.keys(set);
}

function _handleRealtimeUpdate(row, forceApply) {
  if (!row) return;

  var remoteVersion = row.version || 0;

  if (!forceApply) {
    var remoteSessionId = row.session_id || '';
    if (remoteSessionId && remoteSessionId === SYNC.sessionId) {
      return;
    }
    if (!remoteSessionId) {
      var myUser = AUTH.currentUser ? AUTH.currentUser.username : '';
      var remoteUser = row.modified_by || '';
      if (myUser && remoteUser === myUser && (Date.now() - SYNC._lastSaveAt < 5000)) {
        return;
      }
    }
  }

  if (!forceApply && remoteVersion <= SYNC._version) return;

  var by = row.modified_by || '?';

  if (row.state) {
    STATE.records   = _mergeById(STATE.records,   row.state.records   || []);
    STATE.obsoletos = _mergeById(STATE.obsoletos, row.state.obsoletos || []);
    STATE.papelera  = _mergeById(STATE.papelera,  row.state.papelera  || []);
    STATE.salidas   = _mergeById(STATE.salidas,   row.state.salidas   || []);
    STATE.logs      = (row.state.logs || []).length >= (STATE.logs || []).length ? row.state.logs : STATE.logs;
    STATE.elaboros  = _mergeArrays(STATE.elaboros, row.state.elaboros || []);
    STATE.nextId    = Math.max(STATE.nextId || 1000, row.state.nextId || 1000);
  }

  SYNC._version = remoteVersion;
  localSaveState();

  if (AUTH.currentUser) {
    render();
    if (by && by !== (AUTH.currentUser ? AUTH.currentUser.username : '')) {
      showToast('📥 ' + by + ' actualizó los datos', 'info');
      if (typeof _playNotificationSound === 'function') _playNotificationSound();
    }
    if (typeof _unsavedChanges !== 'undefined') _unsavedChanges = false;
    SYNC._lastSyncTime = Date.now();
    _updateLastSyncUI();
  }
}

async function _pullRemote(forceApply) {
  if (!SYNC._client) return;

  try {
    var result = await SYNC._client
      .from('app_state')
      .select('*')
      .eq('id', 1)
      .single();

    if (result.error || !result.data) return;

    var shouldApply = forceApply || (result.data.version > SYNC._version);
    if (shouldApply && result.data.state) {
      _handleRealtimeUpdate(result.data, !!forceApply);
    }
  } catch (e) {
  }
}
function _startHeartbeat() {
  clearInterval(SYNC._heartbeatTimer);
  SYNC._heartbeatTimer = setInterval(async function() {
    if (!SYNC.initialized || !SYNC._client) return;

    try {
      var result = await SYNC._client
        .from('app_state')
        .select('version')
        .eq('id', 1)
        .single();

      if (result.error) {
        if (SYNC.connected) {
          SYNC.connected = false;
          updateSyncIndicator();
        }
        return;
      }

      if (!SYNC.connected) {
        SYNC.connected = true;
        updateSyncIndicator();
        listenStateChanges();
        if (AUTH.currentUser) registerPresence();
        showToast('🔄 Reconectado a la nube', 'success');
      }

      var remoteVersion = result.data.version || 0;
      if (remoteVersion > SYNC._version) {
        await _pullRemote(false);
      }

    } catch (e) {
      if (SYNC.connected) {
        SYNC.connected = false;
        updateSyncIndicator();
      }
    }
  }, HEARTBEAT_INTERVAL);
}
function _startVigenciaCheck() {
  clearInterval(SYNC._vigenciaTimer);

  var prevVencidosCount = -1;

  SYNC._vigenciaTimer = setInterval(function() {
    if (!AUTH.currentUser) return;
    if (typeof calcEstado !== 'function') return;

    var currentVencidos = 0;
    (STATE.records || []).forEach(function(r) {
      if (calcEstado(r.fechaEmision) === 'Vencido') currentVencidos++;
    });

    if (prevVencidosCount >= 0 && currentVencidos !== prevVencidosCount) {
      render();
      if (currentVencidos > prevVencidosCount) {
        var nuevos = currentVencidos - prevVencidosCount;
        showToast('⚠️ ' + nuevos + ' documento(s) acaba(n) de vencer', 'error');
      }
    }
    prevVencidosCount = currentVencidos;
  }, VIGENCIA_CHECK_INTERVAL);
}
function generateBackup() {
  if (!STATE.records || STATE.records.length === 0) {
    showToast('No hay datos para respaldar', 'error');
    return;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0')
    + String(now.getSeconds()).padStart(2,'0');

  var backupData = {
    _meta: {
      tipo:      'DEBBIOM_BACKUP',
      version:   '2.0',
      fecha:     now.toISOString(),
      timestamp: timestamp,
      usuario:   AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
      syncVersion: SYNC._version,
      totalRegistros: (STATE.records || []).length,
      totalObsoletos: (STATE.obsoletos || []).length,
      totalPapelera:  (STATE.papelera || []).length,
      totalSalidas:   (STATE.salidas || []).length,
      totalLogs:      (STATE.logs || []).length
    },
    records:   STATE.records   || [],
    obsoletos: STATE.obsoletos || [],
    papelera:  STATE.papelera  || [],
    salidas:   STATE.salidas   || [],
    logs:      STATE.logs || [],
    elaboros:  STATE.elaboros  || [],
    nextId:    STATE.nextId    || 1000,
    users:     (AUTH.users || []).map(function(u) {
      return { username: u.username, role: u.role, full_name: u.full_name || u.fullName || '' };
    })
  };

  var json = JSON.stringify(backupData, null, 2);
  var blob = new Blob([json], { type: 'application/json' });
  var url  = URL.createObjectURL(blob);

  var a = document.createElement('a');
  a.href = url;
  a.download = 'DEBBIOM_BACKUP_' + timestamp + '.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  SYNC._lastBackupAt = Date.now();
  addLog('backup', 'Backup JSON generado: DEBBIOM_BACKUP_' + timestamp + '.json');
  showToast('✅ Backup descargado + guardado en la nube', 'success');

  if (SYNC.initialized && SYNC._client) {
    _saveCloudBackup('json', 'Backup JSON descargado en ' + (AUTH.currentUser ? AUTH.currentUser.username : 'desconocido'));
  }
}

function generateBackupExcel() {
  if (typeof XLSX === 'undefined') {
    showToast('Error: Librería XLSX no disponible', 'error');
    return;
  }
  if (!STATE.records || STATE.records.length === 0) {
    showToast('No hay datos para respaldar', 'error');
    return;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0');

  var wb = XLSX.utils.book_new();

  var recHeaders = ['ID','No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Vigencia','Estado','Días Vigencia','Tipo Resguardo','Ubicación','Copias','Tipo Resguardo Copia','Usuarios','Observaciones','URLs Archivos','URLs Archivos Copia'];
  var recRows = (STATE.records || []).map(function(r) {
    var urls = (r.archivoURLs || []).join(' | ');
    var urlsCopia = (r.archivoURLsCopia || []).join(' | ');
    if (r.archivos && r.archivos.length > 0) {
      var nombres = r.archivos.map(function(f) { return f.name || ''; }).join(' | ');
      urls = urls ? urls + ' | [Adjuntos: ' + nombres + ']' : '[Adjuntos: ' + nombres + ']';
    }
    if (r.archivosCopia && r.archivosCopia.length > 0) {
      var nombresCopia = r.archivosCopia.map(function(f) { return f.name || ''; }).join(' | ');
      urlsCopia = urlsCopia ? urlsCopia + ' | [Adjuntos: ' + nombresCopia + ']' : '[Adjuntos: ' + nombresCopia + ']';
    }
    return [r.id, r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro,
      typeof renderVigenciaDate === 'function' ? renderVigenciaDate(r.fechaEmision) : '',
      typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '',
      typeof calcDiasVigencia === 'function' ? calcDiasVigencia(r.fechaEmision) : '',
      r.tipoResguardo, r.ubicacion, r.copias, r.tipoResguardoCopia, r.usuarios, r.observaciones,
      urls, urlsCopia
    ];
  });
  var ws1 = XLSX.utils.aoa_to_sheet([recHeaders].concat(recRows));
  XLSX.utils.book_append_sheet(wb, ws1, 'Registros');

  if ((STATE.obsoletos || []).length > 0) {
    var obsHeaders = ['ID','Fecha Obsoleto','No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Motivo','Observaciones','URLs Archivos'];
    var obsRows = STATE.obsoletos.map(function(r) {
      var urls = (r.archivoURLs || []).join(' | ');
      return [r.id, r.fechaObsoleto || '', r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro, r.motivo, r.observaciones, urls];
    });
    var ws2 = XLSX.utils.aoa_to_sheet([obsHeaders].concat(obsRows));
    XLSX.utils.book_append_sheet(wb, ws2, 'Obsoletos');
  }

  if ((STATE.salidas || []).length > 0) {
    var salHeaders = ['Nombre Documento','Versión','Fecha Emisión','Área','Motivo','Destinatario','Correo','Fecha Salida'];
    var salRows = STATE.salidas.map(function(r) {
      return [r.nombreDoc, r.version, r.fechaEmision, r.area, r.motivo, r.destinatario, r.correo, r.fechaSalida];
    });
    var ws3 = XLSX.utils.aoa_to_sheet([salHeaders].concat(salRows));
    XLSX.utils.book_append_sheet(wb, ws3, 'Salidas');
  }

  if ((STATE.papelera || []).length > 0) {
    var papHeaders = ['ID','No.','Nombre Documento','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Versión','Fecha Emisión','Elaboró','Observaciones','Fecha Eliminación'];
    var papRows = STATE.papelera.map(function(r) {
      return [r.id, r.no, r.nombreDoc, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.version, r.fechaEmision, r.elaboro, r.observaciones, r.deletedAt || ''];
    });
    var ws4 = XLSX.utils.aoa_to_sheet([papHeaders].concat(papRows));
    XLSX.utils.book_append_sheet(wb, ws4, 'Papelera');
  }

  if ((STATE.logs || []).length > 0) {
    var logHeaders = ['Fecha','Hora','Usuario','Tipo','Detalle','Documento'];
    var logRows = STATE.logs.map(function(l) {
      var fecha = '', hora = '';
      if (l.ts) {
        try {
          var d = new Date(l.ts);
          fecha = d.toLocaleDateString('es-MX');
          hora = d.toLocaleTimeString('es-MX');
        } catch(e) { fecha = l.ts; }
      }
      return [fecha, hora, l.user || '', l.type || '', l.details || '', l.doc || ''];
    });
    var ws5 = XLSX.utils.aoa_to_sheet([logHeaders].concat(logRows));
    XLSX.utils.book_append_sheet(wb, ws5, 'Logs');
  }

  if (AUTH.users && AUTH.users.length > 0) {
    var userHeaders = ['Usuario','Nombre Completo','Rol'];
    var userRows = AUTH.users.map(function(u) {
      return [u.username, u.full_name || u.fullName || '', u.role || 'user'];
    });
    var ws6 = XLSX.utils.aoa_to_sheet([userHeaders].concat(userRows));
    XLSX.utils.book_append_sheet(wb, ws6, 'Usuarios');
  }

  if (STATE.elaboros && STATE.elaboros.length > 0) {
    var elaHeaders = ['Nombre'];
    var elaRows = STATE.elaboros.map(function(e) { return [e]; });
    var ws7 = XLSX.utils.aoa_to_sheet([elaHeaders].concat(elaRows));
    XLSX.utils.book_append_sheet(wb, ws7, 'Elaboradores');
  }

  var resumen = [
    ['DEBBIOM - Backup COMPLETO'],
    [''],
    ['Fecha de Backup', now.toLocaleString('es-MX')],
    ['Generado por', AUTH.currentUser ? AUTH.currentUser.username : 'Sistema'],
    ['Versión de Sync', String(SYNC._version)],
    ['Next ID', STATE.nextId || 1000],
    [''],
    ['Sección', 'Total'],
    ['Registros activos', (STATE.records || []).length],
    ['Documentos obsoletos', (STATE.obsoletos || []).length],
    ['Registro de salidas', (STATE.salidas || []).length],
    ['Papelera', (STATE.papelera || []).length],
    ['Logs de actividad', (STATE.logs || []).length],
    ['Usuarios', (AUTH.users || []).length],
    ['Elaboradores', (STATE.elaboros || []).length]
  ];
  if (typeof getAlertasResumen === 'function') {
    var alertas = getAlertasResumen();
    resumen.push(['']);
    resumen.push(['Estado de Vigencia', 'Total']);
    resumen.push(['Vencidos', alertas.vencidos]);
    resumen.push(['Críticos (≤15 días)', alertas.criticos]);
    resumen.push(['Aviso (≤30 días)', alertas.avisos]);
    resumen.push(['Próximos (≤60 días)', alertas.proximos]);
  }
  var ws8 = XLSX.utils.aoa_to_sheet(resumen);
  XLSX.utils.book_append_sheet(wb, ws8, 'Resumen');

  var filename = 'DEBBIOM_BACKUP_COMPLETO_' + timestamp + '.xlsx';
  XLSX.writeFile(wb, filename);

  SYNC._lastBackupAt = Date.now();
  addLog('backup', 'Backup Excel COMPLETO generado: ' + filename);
  showToast('✅ Backup Excel descargado + guardado en la nube (' + (STATE.logs||[]).length + ' logs)', 'success');

  if (SYNC.initialized && SYNC._client) {
    _saveCloudBackup('excel', 'Backup Excel descargado en ' + (AUTH.currentUser ? AUTH.currentUser.username : 'desconocido'));
  }
}

function restoreBackup(file) {
  if (!file) return;

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = JSON.parse(e.target.result);

      if (!data._meta || data._meta.tipo !== 'DEBBIOM_BACKUP') {
        showToast('❌ Archivo no es un backup válido de DEBBIOM', 'error');
        return;
      }

      if (!data.records || !Array.isArray(data.records)) {
        showToast('❌ Backup corrupto: no contiene registros', 'error');
        return;
      }

      var msg = '¿Restaurar backup del ' + new Date(data._meta.fecha).toLocaleString('es-MX') + '?\n\n'
        + 'Contenido:\n'
        + '• ' + data.records.length + ' registros\n'
        + '• ' + (data.obsoletos || []).length + ' obsoletos\n'
        + '• ' + (data.salidas || []).length + ' salidas\n\n'
        + '⚠️ Esto REEMPLAZARÁ todos los datos actuales.';

      if (!confirm(msg)) return;

      STATE.records   = data.records;
      STATE.obsoletos = data.obsoletos || [];
      STATE.papelera  = data.papelera  || [];
      STATE.salidas   = data.salidas   || [];
      STATE.logs      = data.logs      || [];
      STATE.elaboros  = data.elaboros  || [];
      STATE.nextId    = data.nextId    || 1000;
      saveState();
      addLog('restore', 'Backup restaurado: ' + data._meta.fecha + ' (' + data.records.length + ' registros)');
      render();
      showToast('✅ Backup restaurado: ' + data.records.length + ' registros', 'success');

    } catch (err) {
      showToast('❌ Error al leer el archivo de backup', 'error');
    }
  };
  reader.readAsText(file);
}

function _startAutoBackup() {
  clearInterval(SYNC._backupTimer);
  SYNC._backupTimer = setInterval(function() {
    if (SYNC.initialized && SYNC._client && (STATE.records||[]).length > 0) {
      _saveCloudBackup('auto', 'Auto-backup programado');
    }
  }, AUTO_BACKUP_INTERVAL);
}
async function _saveCloudBackup(tipo, nota) {
  if (!SYNC.initialized || !SYNC._client) {
    return null;
  }
  if (!STATE.records || STATE.records.length === 0) {
    return null;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0')
    + String(now.getSeconds()).padStart(2,'0');

  var nombre = 'BACKUP_' + tipo.toUpperCase() + '_' + timestamp;

  var stateData = {
    records:   STATE.records   || [],
    obsoletos: STATE.obsoletos || [],
    papelera:  STATE.papelera  || [],
    salidas:   STATE.salidas   || [],
    logs:      STATE.logs || [],
    elaboros:  STATE.elaboros  || [],
    nextId:    STATE.nextId    || 1000
  };

  var usersData = (AUTH.users || []).map(function(u) {
    return { username: u.username, role: u.role, full_name: u.full_name || u.fullName || '' };
  });

  try {
    var result = await SYNC._client
      .from('backups')
      .insert({
        nombre:          nombre,
        tipo:            tipo,
        created_by:      AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
        created_at:      now.toISOString(),
        sync_version:    SYNC._version || 0,
        total_registros: (STATE.records || []).length,
        total_obsoletos: (STATE.obsoletos || []).length,
        total_salidas:   (STATE.salidas || []).length,
        total_papelera:  (STATE.papelera || []).length,
        total_logs:      (STATE.logs || []).length,
        state:           stateData,
        users:           usersData,
        nota:            nota || ''
      })
      .select('id, nombre')
      .single();

    if (result.error) {
      return null;
    }

    return result.data;
  } catch (e) {
    return null;
  }
}

async function listCloudBackups() {
  if (!SYNC.initialized || !SYNC._client) return [];

  try {
    var result = await SYNC._client
      .from('backups')
      .select('id, nombre, tipo, created_by, created_at, sync_version, total_registros, total_obsoletos, total_salidas, total_papelera, total_logs, nota')
      .order('created_at', { ascending: false })
      .limit(50);

    if (result.error) {
      return [];
    }

    return result.data || [];
  } catch (e) {
    return [];
  }
}

async function restoreCloudBackup(backupId) {
  if (!SYNC.initialized || !SYNC._client) {
    showToast('No hay conexión a Supabase', 'error');
    return false;
  }

  try {
    var result = await SYNC._client
      .from('backups')
      .select('*')
      .eq('id', backupId)
      .single();

    if (result.error || !result.data) {
      showToast('Backup no encontrado', 'error');
      return false;
    }

    var bk = result.data;

    if (!bk.state || !bk.state.records) {
      showToast('Backup corrupto: sin registros', 'error');
      return false;
    }

    var msg = '¿Restaurar backup "' + bk.nombre + '"?\n\n'
      + 'Fecha: ' + new Date(bk.created_at).toLocaleString('es-MX') + '\n'
      + 'Creado por: ' + bk.created_by + '\n'
      + 'Registros: ' + bk.total_registros + '\n'
      + 'Obsoletos: ' + bk.total_obsoletos + '\n'
      + 'Salidas: ' + bk.total_salidas + '\n\n'
      + '⚠️ Esto REEMPLAZARÁ todos los datos actuales en TODOS los dispositivos.';

    if (!confirm(msg)) return false;

    STATE.records   = bk.state.records   || [];
    STATE.obsoletos = bk.state.obsoletos || [];
    STATE.papelera  = bk.state.papelera  || [];
    STATE.salidas   = bk.state.salidas   || [];
    STATE.logs      = bk.state.logs      || [];
    STATE.elaboros  = bk.state.elaboros  || [];
    STATE.nextId    = bk.state.nextId    || 1000;
    saveState();
    addLog('restore-cloud', 'Backup restaurado desde nube: ' + bk.nombre + ' (' + bk.total_registros + ' registros)');
    render();
    showToast('✅ Backup restaurado: ' + bk.total_registros + ' registros (sincronizando a todos los dispositivos...)', 'success');
    return true;

  } catch (e) {
    showToast('Error al restaurar backup', 'error');
    return false;
  }
}

async function deleteCloudBackup() {
  showToast('Los backups no se pueden eliminar', 'error');
  return false;
}

async function saveManualCloudBackup(nota) {
  var n = nota || 'Backup manual';
  showToast('☁️ Guardando backup en la nube...', 'info');
  var result = await _saveCloudBackup('manual', n);
  if (result) {
    addLog('backup-cloud', 'Backup guardado en la nube: ' + result.nombre);
    showToast('✅ Backup guardado en la nube', 'success');
  } else {
    showToast('❌ Error al guardar backup en la nube', 'error');
  }
  return result;
}

async function downloadCloudBackup(backupId) {
  if (!SYNC.initialized || !SYNC._client) return;

  try {
    var result = await SYNC._client
      .from('backups')
      .select('*')
      .eq('id', backupId)
      .single();

    if (result.error || !result.data) {
      showToast('Backup no encontrado', 'error');
      return;
    }

    var bk = result.data;
    var backupData = {
      _meta: {
        tipo:           'DEBBIOM_BACKUP',
        version:        '2.0',
        fecha:          bk.created_at,
        usuario:        bk.created_by,
        syncVersion:    bk.sync_version,
        totalRegistros: bk.total_registros,
        totalObsoletos: bk.total_obsoletos,
        totalSalidas:   bk.total_salidas,
        totalPapelera:  bk.total_papelera,
        totalLogs:      bk.total_logs,
        origen:         'supabase_cloud',
        nota:           bk.nota
      },
      records:   bk.state.records   || [],
      obsoletos: bk.state.obsoletos || [],
      papelera:  bk.state.papelera  || [],
      salidas:   bk.state.salidas   || [],
      logs:      bk.state.logs      || [],
      elaboros:  bk.state.elaboros  || [],
      nextId:    bk.state.nextId    || 1000,
      users:     bk.users           || []
    };

    var json = JSON.stringify(backupData, null, 2);
    var blob = new Blob([json], { type: 'application/json' });
    var url  = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = bk.nombre + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('✅ Backup descargado: ' + bk.nombre + '.json', 'success');
  } catch (e) {
    showToast('Error descargando backup', 'error');
  }
}

async function _cleanOldCloudBackups() { return; }
function registerPresence() {
  if (!SYNC.initialized || !SYNC._client || !AUTH.currentUser) return;

  if (SYNC._presence) {
    try {
      SYNC._presence.untrack();
      SYNC._client.removeChannel(SYNC._presence);
    } catch (e) {}
    SYNC._presence = null;
  }

  SYNC._presence = SYNC._client.channel('online-users', {
    config: { presence: { key: SYNC.sessionId } }
  });

  SYNC._presence.on('presence', { event: 'sync' }, function() {
    var state = SYNC._presence.presenceState();
    var users = [];
    for (var key in state) {
      var entries = state[key];
      if (entries && entries.length > 0) {
        users.push(entries[0].username);
      }
    }
    SYNC.onlineUsers = users.filter(function(v, i, a) { return a.indexOf(v) === i; });
    SYNC.onlineCount = SYNC.onlineUsers.length;
    updateOnlineCountUI();
  });

  SYNC._presence.subscribe(async function(status) {
    if (status === 'SUBSCRIBED') {
      await SYNC._presence.track({
        username:  AUTH.currentUser.username,
        fullName:  AUTH.currentUser.fullName || AUTH.currentUser.username,
        online_at: new Date().toISOString()
      });
    }
  });
}

function unregisterPresence() {
  if (SYNC._presence) {
    try {
      SYNC._presence.untrack();
      SYNC._client.removeChannel(SYNC._presence);
    } catch (e) {}
    SYNC._presence = null;
  }
}

function updateOnlineCountUI() {
  var el1 = document.querySelector('.online-count');
  if (el1) el1.textContent = String(SYNC.onlineCount);
  var el2 = document.getElementById('sidebar-online-count');
  if (el2) el2.textContent = String(SYNC.onlineCount);
}
function _updateLastSyncUI() {
  var el = document.getElementById('last-sync-time');
  if (!el) return;
  if (!SYNC._lastSyncTime) { el.textContent = '—'; return; }
  var diff = Math.floor((Date.now() - SYNC._lastSyncTime) / 1000);
  if (diff < 5) el.textContent = 'Ahora';
  else if (diff < 60) el.textContent = 'Hace ' + diff + 's';
  else if (diff < 3600) el.textContent = 'Hace ' + Math.floor(diff/60) + ' min';
  else el.textContent = 'Hace ' + Math.floor(diff/3600) + 'h';
}
setInterval(_updateLastSyncUI, 10000);

function updateSyncIndicator() {
  var el = document.getElementById('sync-indicator');
  if (!el) return;
  if (!SYNC.initialized) {
    el.className = 'sync-indicator sync-offline';
    el.innerHTML = '<i class="fas fa-database"></i> Local';
  } else if (SYNC.connected) {
    el.className = 'sync-indicator sync-online';
    el.innerHTML = '<i class="fas fa-cloud"></i> En línea';
  } else {
    el.className = 'sync-indicator sync-reconnecting';
    el.innerHTML = '<i class="fas fa-cloud-arrow-up"></i> Reconectando...';
  }
}

function _showSyncAnimation(active) {
  var btn = document.getElementById('btn-sync');
  if (!btn) return;
  if (active) {
    btn.classList.add('syncing');
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sincronizando...';
  } else {
    btn.classList.remove('syncing');
    btn.innerHTML = '<i class="fas fa-arrows-rotate"></i> Sincronizar';
  }
}
async function syncLoadState() {
  var status = await _syncLoadWithStatus();
  return status === 'loaded';
}

async function _syncLoadWithStatus() {
  if (!SYNC.initialized || !SYNC._client) return 'error';

  try {
    var result = await SYNC._client
      .from('app_state')
      .select('*')
      .eq('id', 1)
      .single();

    if (result.error) return 'error';
    if (!result.data || !result.data.state) return 'empty';

    if (result.data.state.records && result.data.state.records.length > 0) {
      STATE.records   = result.data.state.records;
      STATE.obsoletos = result.data.state.obsoletos || [];
      STATE.papelera  = result.data.state.papelera  || [];
      STATE.salidas   = result.data.state.salidas   || [];
      STATE.logs      = result.data.state.logs      || [];
      STATE.elaboros  = result.data.state.elaboros  || [];
      STATE.nextId    = result.data.state.nextId    || 1000;
      SYNC._version   = result.data.version || 0;
      localSaveState();
      return 'loaded';
    }

    return 'empty';
  } catch (e) {
    return 'error';
  }
}
function syncSaveUsers() { /* Usuarios gestionados por Supabase Auth */ }
async function syncLoadUsers() { return null; }
function listenUserChanges() { /* Usuarios gestionados por Supabase Auth */ }
function localSaveState() {}

function localLoadState() { return false; }
function cleanupSync() {
  clearTimeout(SYNC._saveTimer);
  clearInterval(SYNC._heartbeatTimer);
  clearInterval(SYNC._vigenciaTimer);
  clearInterval(SYNC._backupTimer);
  if (SYNC._channel) {
    SYNC._client.removeChannel(SYNC._channel);
  }
  unregisterPresence();
}

window.addEventListener('beforeunload', cleanupSync);
