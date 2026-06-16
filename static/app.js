/* ── Sesión ──────────────────────────────────────────────────────────────── */
const SESSION_KEY     = 'tachyon_session'
const SESSION_MINUTES = 15
let sessionTimer      = null
let countdownInterval = null

function getSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY)
    if (!raw) return null
    const s = JSON.parse(raw)
    if (Date.now() > s.expires) { sessionStorage.removeItem(SESSION_KEY); return null }
    return s
  } catch { return null }
}

function setSession(key) {
  const s = { key, expires: Date.now() + SESSION_MINUTES * 60 * 1000 }
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(s))
}

function resetSessionTimer() {
  const s = getSession()
  if (!s) return
  setSession(s.key)           // renueva expiración
  startCountdown()
}

function startCountdown() {
  clearInterval(countdownInterval)
  countdownInterval = setInterval(() => {
    const s = getSession()
    if (!s) { logout(); return }
    const mins  = Math.floor((s.expires - Date.now()) / 60000)
    const secs  = Math.floor(((s.expires - Date.now()) % 60000) / 1000)
    const el    = document.getElementById('sessionTimer')
    if (el) el.textContent = `Sesión: ${mins}:${secs.toString().padStart(2, '0')}`
    if (Date.now() >= s.expires) logout()
  }, 1000)
}

// Renovar sesión en cada interacción del usuario
;['click', 'keydown', 'mousemove', 'touchstart'].forEach(ev =>
  document.addEventListener(ev, () => { if (getSession()) resetSessionTimer() }, { passive: true })
)

/* ── Login ───────────────────────────────────────────────────────────────── */
async function doLogin() {
  const keyInput = document.getElementById('accessKey')
  const key      = keyInput.value.trim()
  if (!key) return

  const btn     = document.getElementById('btnLogin')
  const spinner = document.getElementById('loginSpinner')
  const btnText = document.getElementById('loginBtnText')
  const errEl   = document.getElementById('loginError')

  btn.disabled      = true
  spinner.style.display = 'block'
  btnText.textContent   = 'Verificando…'
  errEl.classList.remove('show')

  try {
    // Verificamos la clave enviando un FormData vacío excepto la key
    const fd = new FormData()
    fd.append('key',  key)
    fd.append('check', '1')     // señal para que el backend solo valide

    const res  = await fetch('/verificar-clave', { method: 'POST', body: fd })
    const data = await res.json()

    if (!res.ok || !data.ok) {
      errEl.classList.add('show')
      keyInput.value = ''
      keyInput.focus()
      return
    }

    // Clave correcta → guardar sesión y mostrar app
    setSession(key)
    document.getElementById('loginOverlay').classList.add('hidden')
    document.getElementById('sessionBadge').style.display = 'flex'
    startCountdown()

  } catch {
    errEl.textContent = 'No se pudo conectar con el servidor'
    errEl.classList.add('show')
  } finally {
    btn.disabled          = false
    spinner.style.display = 'none'
    btnText.textContent   = 'Entrar'
  }
}

function logout() {
  clearInterval(countdownInterval)
  sessionStorage.removeItem(SESSION_KEY)
  document.getElementById('loginOverlay').classList.remove('hidden')
  document.getElementById('sessionBadge').style.display = 'none'
  document.getElementById('accessKey').value = ''
  // Limpiar resultado si lo había
  nuevo(true)
}

function toggleKey() {
  const inp = document.getElementById('accessKey')
  inp.type  = inp.type === 'password' ? 'text' : 'password'
}

/* ── Init ────────────────────────────────────────────────────────────────── */
;(function init() {
  const s = getSession()
  if (s) {
    document.getElementById('loginOverlay').classList.add('hidden')
    document.getElementById('sessionBadge').style.display = 'flex'
    startCountdown()
  }
})()

/* ── Drag & drop ─────────────────────────────────────────────────────────── */
let archivoActual = null

const dz = document.getElementById('dropzone')
dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag') })
dz.addEventListener('dragleave', ()  => dz.classList.remove('drag'))
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag')
  const file = e.dataTransfer.files[0]
  if (file) setArchivo(file)
})

function onFileChange(e) {
  const file = e.target.files[0]
  if (file) setArchivo(file)
}

function setArchivo(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    mostrarAlerta('Solo se aceptan archivos .xlsx o .xls', 'error'); return
  }
  archivoActual = file
  document.getElementById('fileName').textContent = file.name
  document.getElementById('fileSelected').classList.add('show')
  ocultarAlerta()
  actualizarBoton()
}

function quitarArchivo() {
  archivoActual = null
  document.getElementById('fileSelected').classList.remove('show')
  document.getElementById('fileInput').value = ''
  actualizarBoton()
}

function actualizarBoton() {
  const btn = document.getElementById('btnProcesar')
  const txt = document.getElementById('btnText')
  btn.disabled    = !archivoActual
  txt.textContent = archivoActual ? 'Analizar archivo' : 'Selecciona un archivo para continuar'
}

/* ── Alertas ─────────────────────────────────────────────────────────────── */
function mostrarAlerta(msg, tipo = 'error') {
  const el = document.getElementById('alerta')
  el.className = `alert show ${tipo}`
  document.getElementById('alertaMsg').textContent = msg
}
function ocultarAlerta() {
  document.getElementById('alerta').classList.remove('show')
}

/* ── Procesar ────────────────────────────────────────────────────────────── */
async function procesar() {
  const s = getSession()
  if (!s) { logout(); return }
  if (!archivoActual) { mostrarAlerta('Selecciona un archivo Excel'); return }

  const btn     = document.getElementById('btnProcesar')
  const spinner = document.getElementById('spinner')
  const btnText = document.getElementById('btnText')

  btn.disabled          = true
  spinner.style.display = 'block'
  btnText.textContent   = 'Procesando…'
  ocultarAlerta()

  const form = new FormData()
  form.append('key',  s.key)
  form.append('file', archivoActual)

  try {
    const res  = await fetch('/procesar', { method: 'POST', body: form })
    const data = await res.json()

    if (!res.ok) {
      if (res.status === 401) { logout(); return }
      mostrarAlerta(data.error || 'Error al procesar el archivo', 'error')
      return
    }

    resetSessionTimer()
    mostrarResultado(data.resumen, archivoActual.name)

  } catch {
    mostrarAlerta('No se pudo conectar con el servidor. Verifica tu conexión.', 'error')
  } finally {
    btn.disabled          = false
    spinner.style.display = 'none'
    btnText.textContent   = archivoActual ? 'Analizar archivo' : 'Selecciona un archivo para continuar'
  }
}

/* ── Resultado ───────────────────────────────────────────────────────────── */
function mostrarResultado(r, nombre) {
  document.getElementById('resNombre').textContent      = nombre
  document.getElementById('mRegistros').textContent     = r.total_registros?.toLocaleString('es-MX') ?? '—'
  document.getElementById('mPassed').textContent        = r.total_passed?.toLocaleString('es-MX')    ?? '—'
  document.getElementById('mFailed').textContent        = r.total_failed?.toLocaleString('es-MX')    ?? '—'
  document.getElementById('mHorasTotal').textContent    = (r.horas_totales?.toFixed(2)  ?? '—') + 'h'
  document.getElementById('mHorasMuertas').textContent  = (r.horas_muertas?.toFixed(2)  ?? '—') + 'h'
  document.getElementById('mHorasReales').textContent   = (r.horas_reales?.toFixed(2)   ?? '—') + 'h'

  const tags = document.getElementById('resTags')
  tags.innerHTML = ''
  const addTag = (txt, azul) => {
    const t = document.createElement('span')
    t.className   = 'tag' + (azul ? ' blue' : '')
    t.textContent = txt
    tags.appendChild(t)
  }
  addTag(`${r.modelos} modelos`)
  r.fechas?.forEach(f => addTag(f, true))

  document.getElementById('resultado').classList.add('show')
  document.getElementById('resultado').scrollIntoView({ behavior: 'smooth', block: 'start' })
}

/* ── Descargar ───────────────────────────────────────────────────────────── */
function descargar() {
  const s = getSession()
  if (!s) { logout(); return }
  const a  = document.createElement('a')
  a.href   = '/descargar'
  a.click()
  resetSessionTimer()
}

/* ── Nuevo ───────────────────────────────────────────────────────────────── */
function nuevo(silencioso = false) {
  archivoActual = null
  document.getElementById('fileSelected').classList.remove('show')
  document.getElementById('fileInput').value = ''
  document.getElementById('resultado').classList.remove('show')
  document.getElementById('resTags').innerHTML = ''
  ocultarAlerta()
  actualizarBoton()
  if (!silencioso) window.scrollTo({ top: 0, behavior: 'smooth' })
}
