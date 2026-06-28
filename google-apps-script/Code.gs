const SHEET_ID = 'REPLACE_WITH_GOOGLE_SHEET_ID';
const SHEET_NAME = 'Postulaciones';
const MIN_SUBMIT_DELAY_MS = 4000;

const HEADERS = [
  'created_at',
  'role',
  'full_name',
  'phone',
  'email',
  'evidence_links',
  'ai_brief',
  'ai_projects',
  'superpower',
  'role_question',
  'consent',
  'page_url',
  'user_agent',
  'submission_key',
  'status',
];

function doGet() {
  return jsonResponse_({
    status: 'ok',
    message: 'WeSpeak application endpoint online.',
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = normalizePayload_(e && e.parameter ? e.parameter : {});
    const validation = validatePayload_(payload);

    if (!validation.ok) {
      return jsonResponse_({
        status: validation.status,
        message: validation.message,
      });
    }

    const sheet = getSheet_();

    if (isDuplicate_(sheet, payload.submissionKey)) {
      return jsonResponse_({
        status: 'duplicate',
        message:
          'Ya existe una postulación registrada para este correo y esta vacante. Si necesitas actualizarla, contacta al equipo.',
      });
    }

    sheet.appendRow([
      new Date().toISOString(),
      safeCell_(payload.role),
      safeCell_(payload.fullName),
      safeCell_(payload.phone),
      safeCell_(payload.email),
      safeCell_(payload.evidenceLinks),
      safeCell_(payload.aiBrief),
      safeCell_(payload.aiProjects),
      safeCell_(payload.superpower),
      safeCell_(payload.roleQuestion),
      payload.consent ? 'true' : 'false',
      safeCell_(payload.pageUrl),
      safeCell_(payload.userAgent),
      safeCell_(payload.submissionKey),
      'new',
    ]);

    return jsonResponse_({
      status: 'success',
      message:
        '¡Listo! Tu postulación quedó guardada correctamente y ya está visible para el equipo.',
    });
  } catch (error) {
    return jsonResponse_({
      status: 'error',
      message: 'No pudimos guardar la postulación en Google Sheets.',
      detail: String(error && error.message ? error.message : error),
    });
  } finally {
    lock.releaseLock();
  }
}

function normalizePayload_(raw) {
  const role = raw.role === 'assistant' ? 'assistant' : 'content';
  const email = normalizeEmail_(raw.email);

  return {
    role,
    fullName: normalizeName_(raw.fullName),
    phone: normalizePhone_(raw.phone),
    email,
    evidenceLinks: normalizeText_(raw.evidenceLinks),
    aiBrief: normalizeText_(raw.aiBrief),
    aiProjects: normalizeText_(raw.aiProjects),
    superpower: normalizeText_(raw.superpower),
    roleQuestion: normalizeText_(raw.roleQuestion),
    consent: String(raw.consent).toLowerCase() === 'true',
    pageUrl: normalizeText_(raw.pageUrl),
    userAgent: normalizeText_(raw.userAgent),
    company: normalizeText_(raw.company),
    startedAt: Number(raw.startedAt || 0),
    submissionKey: [role, email].join('::'),
  };
}

function validatePayload_(payload) {
  if (payload.company) {
    return {
      ok: false,
      status: 'spam',
      message: 'No pudimos procesar esta postulación por una validación anti-spam.',
    };
  }

  if (!payload.startedAt || Date.now() - payload.startedAt < MIN_SUBMIT_DELAY_MS) {
    return {
      ok: false,
      status: 'spam',
      message: 'La postulación se envió demasiado rápido. Inténtalo de nuevo.',
    };
  }

  if (payload.fullName.length < 6) {
    return invalid_('El nombre completo es obligatorio.');
  }

  if (!isValidName_(payload.fullName)) {
    return invalid_('El nombre no puede llevar números ni caracteres raros.');
  }

  if (!isValidPhone_(payload.phone)) {
    return invalid_('El teléfono debe tener entre 10 y 15 dígitos y no puede llevar letras.');
  }

  if (!isValidEmail_(payload.email)) {
    return invalid_('El correo no es válido.');
  }

  if (payload.evidenceLinks.length < 10) {
    return invalid_('La evidencia o links del candidato son obligatorios.');
  }

  if (startsWithNumber_(payload.evidenceLinks)) {
    return invalid_('La evidencia no puede empezar con un número.');
  }

  if (payload.aiBrief && startsWithNumber_(payload.aiBrief)) {
    return invalid_('El brief IA no puede empezar con un número.');
  }

  if (payload.aiBrief && payload.aiBrief.length < 20) {
    return invalid_('El brief IA necesita un poco más de contexto o debe ir vacío.');
  }

  if (payload.aiProjects && startsWithNumber_(payload.aiProjects)) {
    return invalid_('Los proyectos IA no pueden empezar con un número.');
  }

  if (payload.aiProjects && payload.aiProjects.length < 10) {
    return invalid_('Los proyectos IA necesitan una descripción útil o deben ir vacíos.');
  }

  if (payload.superpower.length < 20) {
    return invalid_('El superpoder necesita una explicación más sustanciosa.');
  }

  if (payload.roleQuestion.length < 40) {
    return invalid_('La respuesta del caso del rol es obligatoria y debe tener contexto.');
  }

  if (!payload.consent) {
    return invalid_('El consentimiento de datos es obligatorio.');
  }

  return { ok: true };
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}

function isDuplicate_(sheet, submissionKey) {
  if (sheet.getLastRow() <= 1) return false;

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length)
    .getValues();
  const keyIndex = HEADERS.indexOf('submission_key');

  return values.some((row) => String(row[keyIndex]).trim() === submissionKey);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function invalid_(message) {
  return {
    ok: false,
    status: 'invalid',
    message,
  };
}

function normalizeText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeName_(value) {
  return String(value || '')
    .replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’. -]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeEmail_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function normalizePhone_(value) {
  return String(value || '')
    .replace(/[^\d+()\-\s]/g, '')
    .trim();
}

function isValidName_(value) {
  return (
    value.length >= 6 &&
    /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’. -]+$/.test(value) &&
    /[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/.test(value)
  );
}

function isValidPhone_(value) {
  var digits = String(value || '').replace(/\D/g, '');
  return /^[-+()\d\s]{10,20}$/.test(value) && digits.length >= 10 && digits.length <= 15;
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function startsWithNumber_(value) {
  return /^\d/.test(String(value || '').trim());
}

function safeCell_(value) {
  const normalized = normalizeText_(value);
  return /^[=+\-@]/.test(normalized) ? "'" + normalized : normalized;
}
