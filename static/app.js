function mostrarResultado(r, nombre) {

document.getElementById('resNombre')
.textContent = nombre

// Métricas nuevas

const registros =
document.getElementById('mRegistros')

const passed =
document.getElementById('mPassed')

const failed =
document.getElementById('mFailed')

const horasTotal =
document.getElementById('mHorasTotal')

const horasMuertas =
document.getElementById('mHorasMuertas')

const horasReales =
document.getElementById('mHorasReales')

if (registros)
registros.textContent =
r.modelos?.toLocaleString('es-MX') ?? '0'

if (passed)
passed.textContent =
r.piezas_totales?.toLocaleString('es-MX') ?? '0'

if (failed)
failed.textContent =
r.minutos_totales?.toLocaleString('es-MX') ?? '0'

if (horasTotal)
horasTotal.textContent =
"${r.horas_totales ?? 0} h"

if (horasMuertas)
horasMuertas.textContent = '—'

if (horasReales)
horasReales.textContent = '—'

const tags =
document.getElementById('resTags')

tags.innerHTML = ''

const tag =
document.createElement('span')

tag.className = 'tag'

tag.textContent =
"${r.modelos} modelos procesados"

tags.appendChild(tag)

document.getElementById('resultado')
.classList.add('show')

document.getElementById('resultado')
.scrollIntoView({
behavior: 'smooth',
block: 'start'
})
}