// Hub Evaluación 360 · Healthy Food Marcalman Ecuador
// Una sola página sin framework. Datos: localStorage (demo) o una lista de SharePoint (una fila por registro, columna Datos = JSON).
(() => {
  'use strict';
  const CFG = window.HUB_CONFIG || {};
  // Cargos y competencias del Manual de Funciones. No van en el código publicado: se guardan en SharePoint (fila de configuración).
  // En la copia local se cargan de datos-manual.js / datos-personas.js (no se publican) y se suben al guardar la Configuración.
  let MAN = { cargos: [], competencias: {} };
  const NO_EVALUADOS = ['presidente', 'gerente-general']; // el Gerente General solo evalúa: nadie lo evalúa
  const DEMO = CFG.modo !== 'sharepoint' || !CFG.clientId || /[?&]demo\b/.test(location.search) || location.hash === '#demo';
  const app = document.getElementById('app');
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clon = o => JSON.parse(JSON.stringify(o));
  const ahora = () => new Date().toISOString();
  const fechaLocal = s => (/^\d{4}-\d{2}-\d{2}$/.test(s || '') ? new Date(s + 'T12:00:00') : new Date(s));
  const fmtFecha = s => s ? fechaLocal(s).toLocaleDateString('es-EC', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
  const fmtFechaHora = s => s ? new Date(s).toLocaleString('es-EC', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
  const sumarDias = (s, d) => { const x = fechaLocal(s); x.setDate(x.getDate() + d); return x.toISOString(); };
  const diasHasta = s => Math.ceil((fechaLocal(s) - new Date()) / 864e5);
  const vacio = v => v == null || (typeof v === 'string' && !v.trim()) || (Array.isArray(v) && !v.some(x => String(x || '').trim()))
    || (typeof v === 'object' && !Array.isArray(v) && !Object.values(v).some(x => String(x || '').trim()));
  const sinTildes = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

  // ---------- Catálogos fijos ----------
  const ROLES = {
    auto: { nombre: 'Autoevaluación', corto: 'Auto', de: 'Autoevaluación', obj: 'la autoevaluación' },
    jefe: { nombre: 'Jefe inmediato', corto: 'Jefe', de: 'Evaluación del jefe', obj: 'la evaluación como jefe' },
    par: { nombre: 'Par', corto: 'Par', de: 'Mirada de un par', obj: 'su mirada como par' },
    colab: { nombre: 'Colaborador', corto: 'Colab.', de: 'Mirada de un colaborador', obj: 'su mirada como colaborador' }
  };
  const COLORES = { lima: '#8cc63f', teal: '#2a9d8f', naranja: '#f39200', verde: '#1f7a3d', azul: '#3d6fb6', morado: '#7a4fb0' };
  const ICONOS = { lima: '+', teal: '↗', naranja: '!', verde: '◎', azul: '★', morado: '◆' };
  const ZONAS = { tarjeta: 'Tarjeta superior', indicador: 'Recuadro del indicador', plan: 'Plan de mejora', extra: 'Sección adicional' };
  const TIPOS = { largo: 'Texto abierto', corto: 'Frase corta', acciones: 'Lista de acciones', competencias: 'Por competencia del cargo' };
  const TODOS = ['auto', 'jefe', 'par', 'colab'];

  const CAMPOS_BASE = [
    { id: 'fortaleza', etiqueta: 'Fortaleza reconocida', tipo: 'largo', zona: 'tarjeta', color: 'lima', activo: true, obligatorio: true, base: true,
      fuentes: TODOS, sugerirCompetencias: true,
      pregunta: { auto: '¿Cuál es la fortaleza con la que más aportas desde tu cargo?', jefe: '¿Qué fortaleza reconoces en {nombre}?',
        par: '¿Qué fortaleza valoras de {nombre} cuando trabajan juntos?', colab: '¿Qué hace especialmente bien {nombre} como líder?' },
      ayuda: 'Una sola fortaleza, la más clara. Si puedes, piensa en un hecho de este año que la demuestre.',
      ejemplo: 'Innovación e implementación tecnológica.' },
    { id: 'conducta', etiqueta: 'Conducta de desarrollo', tipo: 'largo', zona: 'tarjeta', color: 'teal', activo: true, obligatorio: true, base: true,
      fuentes: TODOS, sugerirCompetencias: true,
      pregunta: { auto: '¿Qué comportamiento concreto te propones desarrollar?', jefe: '¿Qué comportamiento concreto debería desarrollar {nombre}?',
        par: '¿Qué podría hacer distinto {nombre} para que el trabajo entre áreas fluya mejor?', colab: '¿Qué te gustaría que {nombre} haga más, o de otra forma?' },
      ayuda: 'Que se pueda ver: una acción o un hábito, no un rasgo de personalidad. "Documentar acuerdos" en vez de "ser más ordenado".',
      ejemplo: 'Acordar entregables con cada área y documentar su validación antes de cerrar una iniciativa.' },
    { id: 'resultado', etiqueta: 'Resultado esperado', tipo: 'largo', zona: 'tarjeta', color: 'naranja', activo: true, obligatorio: true, base: true,
      fuentes: ['auto', 'jefe'],
      pregunta: { auto: 'Si desarrollas esa conducta, ¿qué resultado va a notar la empresa?', jefe: '¿Qué resultado esperas ver de {nombre} en este período?' },
      ayuda: 'El impacto en el área o en la empresa: qué cambia, para quién.',
      ejemplo: 'Mejorar la estructuración de la información empresarial y los sistemas de gestión, registrados.' },
    { id: 'indicador', etiqueta: 'Indicador observable y medible', tipo: 'corto', zona: 'indicador', color: 'verde', activo: true, obligatorio: true, base: true,
      fuentes: ['auto', 'jefe'], validar: 'medible',
      pregunta: { auto: '¿Con qué indicador vamos a saber que lo lograste?', jefe: '¿Con qué indicador medirás el avance de {nombre}?' },
      ayuda: 'Que tenga unidad (%, número, días…) y que otra persona lo pueda verificar sin preguntarte.',
      ejemplo: '% de iniciativas con impacto cuantificado, evidencia y cierre aprobado.' },
    { id: 'meta', etiqueta: 'Meta sugerida', tipo: 'corto', zona: 'indicador', color: 'verde', activo: true, obligatorio: true, base: true,
      fuentes: ['auto', 'jefe'], validar: 'meta',
      pregunta: { auto: '¿Cuál es la meta y cada cuánto se mide?', jefe: '¿Qué meta le propones a {nombre}, y cada cuánto se mide?' },
      ayuda: 'Un número y una frecuencia.', ejemplo: '≥ 90% mensual.' },
    { id: 'plan', etiqueta: 'Plan de mejora', tipo: 'acciones', zona: 'plan', color: 'verde', activo: true, obligatorio: true, base: true, max: 5,
      fuentes: ['auto', 'jefe'],
      pregunta: { auto: '¿Qué acciones concretas vas a hacer?', jefe: '¿Qué acciones concretas le propones a {nombre}?' },
      ayuda: 'De 1 a 5 acciones. Empieza cada una con un verbo. Se revisan a los 30, 60 y 90 días.',
      ejemplo: 'Definir plan de digitalización con hitos y responsables.\nCuantificar ahorro, tiempo o reducción de errores.\nLevantar el plan de mantenimiento para el 2027.' },
    { id: 'competencias', etiqueta: 'Evidencias por competencia', tipo: 'competencias', zona: 'extra', color: 'azul', activo: false, obligatorio: false, base: true,
      fuentes: ['auto', 'jefe'],
      pregunta: { auto: 'Cuéntanos un hecho que muestre cada competencia de tu cargo.', jefe: '¿Qué hecho muestra cada competencia de {nombre}?' },
      ayuda: 'Son las competencias del Manual de Funciones para este cargo. Deja en blanco las que no apliquen.', ejemplo: '' },
    { id: 'comentario', etiqueta: 'Comentario adicional', tipo: 'largo', zona: 'extra', color: 'morado', activo: false, obligatorio: false, base: true,
      fuentes: TODOS,
      pregunta: { auto: '¿Algo más que quieras que se tome en cuenta?', jefe: '¿Algo más sobre {nombre} que quieras dejar registrado?',
        par: '¿Algo más que quieras compartir?', colab: '¿Algo más que quieras compartir?' },
      ayuda: 'Opcional.', ejemplo: '' }
  ];

  const confBase = () => ({
    ciclo: { id: '2026', nombre: 'Evaluación 360 · 2026', inicio: '2026-11-16', cierre: '2026-12-11',
      pares: false, maxPares: 2, colaboradores: false, anonimo: true, jefeVeAuto: true, seguimiento: [30, 60, 90] },
    cargos: {},
    personas: clon(window.HUB_PERSONAS || CFG.personas || []),
    campos: clon(CAMPOS_BASE),
    administradores: clon(CFG.administradores || [])
  });

  // ---------- Estado ----------
  const filas = new Map(); // id → registro
  let CONF = confBase();
  let YO = null;           // persona que usa la app
  let CORREO = '';         // cuenta Microsoft (modo SharePoint)

  const cache = {
    leer(k) { try { return JSON.parse(localStorage.getItem('e360.' + k)); } catch (_) { return null; } },
    guardar(k, v) { try { localStorage.setItem('e360.' + k, JSON.stringify(v)); } catch (_) {} },
    quitar(k) { try { localStorage.removeItem('e360.' + k); } catch (_) {} }
  };

  // ---------- Personas, cargos y fuentes ----------
  const persona = id => CONF.personas.find(p => p.id === id) || null;
  const cargo = id => MAN.cargos.find(c => c.id === id) || { id, nombre: id || 'Sin cargo', area: '', competencias: [], nivel: 'operativo', mision: '' };
  const cargoDe = p => cargo(p && p.cargo);
  const participaCargo = id => {
    const v = CONF.cargos[id];
    if (typeof v === 'boolean') return v;
    const c = cargo(id);
    return c.nivel === 'estrategico' && !NO_EVALUADOS.includes(c.id);
  };
  const evaluados = () => CONF.personas.filter(p => participaCargo(p.cargo));
  const esEvaluado = p => !!p && participaCargo(p.cargo);
  const primerNombre = p => (p ? p.nombre.split(/\s+/)[0] : '');
  const iniciales = n => { const w = String(n || '?').split(/\s+/).filter(Boolean); return ((w[0] || '?')[0] + ((w.length >= 4 ? w[2] : w[1]) || ' ')[0]).toUpperCase().trim(); };
  const colorDe = id => { let h = 0; for (const c of String(id)) h = (h * 31 + c.charCodeAt(0)) % 360; return `hsl(${h} 45% 38%)`; };
  const avatar = (p, cls = '') => `<span class="avatar ${cls}" style="--c:${p ? colorDe(p.id) : '#9aa59c'}">${p ? esc(iniciales(p.nombre)) : '?'}</span>`;
  const esAdmin = () => !!((YO && YO.tthh) || (CORREO && (CONF.administradores || []).map(x => x.toLowerCase()).includes(CORREO)) || (CORREO && (CFG.administradores || []).map(x => x.toLowerCase()).includes(CORREO)));

  function fuentesDe(ev) {
    const f = [{ rol: 'auto', evaluador: ev.id }];
    if (ev.jefe && persona(ev.jefe)) f.push({ rol: 'jefe', evaluador: ev.jefe });
    if (CONF.ciclo.pares && ev.jefe) {
      // Rotación: cada persona recibe (y responde) como máximo maxPares miradas de pares con el mismo jefe.
      const her = evaluados().filter(p => p.jefe === ev.jefe).sort((a, b) => a.id.localeCompare(b.id));
      const i = her.findIndex(p => p.id === ev.id), k = Math.min(Number(CONF.ciclo.maxPares) || 0, her.length - 1);
      for (let d = 1; d <= k; d++) f.push({ rol: 'par', evaluador: her[(i + d) % her.length].id });
    }
    if (CONF.ciclo.colaboradores) CONF.personas.filter(p => p.jefe === ev.id).forEach(p => f.push({ rol: 'colab', evaluador: p.id }));
    return f;
  }
  const camposDe = rol => CONF.campos.filter(c => c.activo && (c.fuentes || []).includes(rol));
  const camposAcuerdo = () => CONF.campos.filter(c => c.activo && (c.fuentes || []).some(f => f === 'auto' || f === 'jefe'));
  const preguntaDe = (c, rol, ev) => ((c.pregunta && (c.pregunta[rol] || c.pregunta.jefe)) || c.etiqueta).split('{nombre}').join(primerNombre(ev));

  // ---------- Registros ----------
  const idResp = (ev, rol, evaluador, ciclo = CONF.ciclo.id) => `${ciclo}|${ev}|${rol}|${evaluador}`;
  const idAcuerdo = (ev, ciclo = CONF.ciclo.id) => `${ciclo}|${ev}|acuerdo|`;
  const resp = (ev, rol, evaluador) => filas.get(idResp(ev, rol, evaluador)) || null;
  const acuerdoDe = ev => filas.get(idAcuerdo(ev)) || null;
  const enviado = r => !!r && r.estado === 'enviado';
  function nuevoRegistro(tipo, ev, rol, evaluador) {
    return { id: tipo === 'acuerdo' ? idAcuerdo(ev) : idResp(ev, rol, evaluador), tipo, ciclo: CONF.ciclo.id, evaluado: ev, rol, evaluador: evaluador || '',
      estado: 'borrador', campos: {}, historial: [], creado: ahora() };
  }
  function hist(reg, accion, detalle) {
    reg.historial = reg.historial || [];
    reg.historial.push({ t: ahora(), quien: YO ? YO.nombre : (CORREO || 'Sistema'), accion, detalle: detalle || '' });
  }
  function progreso(reg, campos) {
    const total = campos.length || 1;
    const hechos = campos.filter(c => reg && !vacio(reg.campos[c.id])).length;
    return { hechos, total, pct: Math.round((hechos / total) * 100) };
  }

  function estadoDe(ev) {
    const auto = resp(ev.id, 'auto', ev.id);
    const jefe = ev.jefe ? resp(ev.id, 'jefe', ev.jefe) : null;
    const otras = fuentesDe(ev).filter(f => f.rol === 'par' || f.rol === 'colab');
    const otrasEnv = otras.filter(f => enviado(resp(ev.id, f.rol, f.evaluador))).length;
    const ac = acuerdoDe(ev.id);
    const hitos = CONF.ciclo.seguimiento.map(d => !!(ac && ac.seguimiento && ac.seguimiento[d]));
    const listoAcuerdo = enviado(auto) && (enviado(jefe) || !ev.jefe);
    let etapa, clase = '';
    if (ac && ac.estado === 'confirmado') { etapa = hitos.every(Boolean) ? 'Completado' : 'En seguimiento'; clase = 'ok'; }
    else if (ac && ac.estado === 'cerrado') { etapa = 'Por confirmar'; clase = 'info'; }
    else if (listoAcuerdo) { etapa = ac ? 'En conversación' : 'Listo para acuerdo'; clase = 'info'; }
    else if (enviado(auto)) { etapa = 'Falta el jefe'; clase = 'borr'; }
    else if (enviado(jefe)) { etapa = 'Falta autoevaluación'; clase = 'borr'; }
    else if (auto || jefe) { etapa = 'En curso'; clase = 'borr'; }
    else etapa = 'Por iniciar';
    return { auto, jefe, otras, otrasEnv, ac, hitos, listoAcuerdo, etapa, clase };
  }
  // El jefe trabaja sobre lo que escribió el evaluado: ve su respuesta en cada pregunta y la particulariza.
  const jefeVeAuto = () => CONF.ciclo.jefeVeAuto !== false;
  const autoParaJefe = () => (R && R.rol === 'jefe' && jefeVeAuto()) ? resp(R.ev.id, 'auto', R.ev.id) : null;
  const pillEstado = r => !r ? '<span class="pill">Sin iniciar</span>' : enviado(r) ? '<span class="pill ok">Enviada</span>' : '<span class="pill borr">Borrador</span>';

  // ---------- Almacenamiento ----------
  const Store = {
    async cargar() {
      filas.clear();
      if (DEMO) {
        const o = cache.leer('demo.filas');
        if (!o) semillaDemo();
        else Object.values(o).forEach(r => filas.set(r.id, r));
      } else await spCargar();
      aplicarConfig();
    },
    async guardar(reg) {
      reg.modificado = ahora();
      reg.modificadoPor = YO ? YO.nombre : CORREO;
      filas.set(reg.id, reg);
      if (DEMO) { cache.guardar('demo.filas', Object.fromEntries(filas)); return; }
      await spGuardar(reg);
    }
  };
  function aplicarConfig() {
    const c = filas.get('config');
    MAN = (c && c.datos && c.datos.manual) || window.HUB_MANUAL || { cargos: [], competencias: {} };
    const b = confBase();
    CONF = c && c.datos ? { ...b, ...clon(c.datos), ciclo: { ...b.ciclo, ...(c.datos.ciclo || {}) } } : b;
    // Campos nuevos de esta versión que aún no estén en la configuración guardada.
    CAMPOS_BASE.forEach(cb => { if (!CONF.campos.some(x => x.id === cb.id)) CONF.campos.push({ ...clon(cb), activo: false }); });
    const n = $('#ciclo-nombre'); if (n) n.textContent = CONF.ciclo.nombre;
  }

  // ---------- SharePoint ----------
  const SPC = CFG.sharepoint || {};
  const API = `https://${String(SPC.sitio || '').split(':')[0]}/_api/v2.0`;
  const URL_APP = location.origin + location.pathname.replace(/index\.html$/, '');
  let pca = null, cuenta = null, spIds = null;
  const itemDe = new Map();
  async function iniciarMsal() {
    pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenantId, redirectUri: URL_APP, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: 'localStorage' }
    });
    await pca.initialize();
    const r = await pca.handleRedirectPromise();
    const propias = pca.getAllAccounts().filter(a => a.tenantId === CFG.tenantId);
    cuenta = (r && r.account) || pca.getActiveAccount() || propias[0] || null;
    if (cuenta) { pca.setActiveAccount(cuenta); CORREO = String(cuenta.username || '').toLowerCase(); }
  }
  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes: CFG.permisos, account: cuenta })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        let ultimo = 0;
        try { ultimo = Number(sessionStorage.getItem('e360.redir') || 0); sessionStorage.setItem('e360.redir', String(Date.now())); } catch (_) {}
        if (Date.now() - ultimo < 60000) throw new Error('Microsoft no entregó el permiso. ¿Ya lo aprobó el administrador?');
        await pca.acquireTokenRedirect({ scopes: CFG.permisos, account: cuenta, loginHint: cuenta && cuenta.username });
        return new Promise(() => {});
      }
      throw e;
    }
  }
  async function sp(ruta, op = {}) {
    const headers = { Authorization: 'Bearer ' + await token(), Accept: 'application/json' };
    if (op.json) headers['Content-Type'] = 'application/json';
    const r = await fetch(ruta.startsWith('http') ? ruta : API + ruta, { method: op.method || 'GET', headers, body: op.json ? JSON.stringify(op.json) : undefined });
    if (!r.ok) { const e = new Error('SharePoint respondió ' + r.status); e.status = r.status; throw e; }
    return r.status === 204 ? null : r.json();
  }
  async function spResolver() {
    if (spIds) return spIds;
    const clave = SPC.sitio + '|' + SPC.lista;
    const g = cache.leer('sp.ids');
    if (g && g.clave === clave) return (spIds = g);
    const sitio = (await sp(`/sites/${SPC.sitio}?$select=id`)).id;
    const lista = (await sp(`/sites/${sitio}/lists?$select=id,name`)).value.find(x => x.name === SPC.lista);
    if (!lista) throw new Error(`No se encontró la lista ${SPC.lista} en SharePoint (o tu cuenta no tiene acceso).`);
    spIds = { clave, sitio, lista: lista.id };
    cache.guardar('sp.ids', spIds);
    return spIds;
  }
  async function spCargar() {
    const x = await spResolver();
    let ruta = `/sites/${x.sitio}/lists/${x.lista}/items?$expand=fields($select=Datos)&$select=id&$top=500`;
    itemDe.clear();
    while (ruta) {
      const r = await sp(ruta);
      for (const it of r.value) {
        let d = null;
        try { d = JSON.parse((it.fields && it.fields.Datos) || 'null'); } catch (_) {}
        if (!d || !d.id) continue;
        itemDe.set(d.id, it.id);
        filas.set(d.id, d);
      }
      ruta = r['@odata.nextLink'] || null;
    }
  }
  function tituloFila(r) {
    if (r.tipo === 'config') return 'Configuración';
    const ev = persona(r.evaluado), q = persona(r.evaluador);
    return [r.ciclo, ev ? ev.nombre : r.evaluado, r.tipo === 'acuerdo' ? 'Acuerdo' : ROLES[r.rol].nombre, q && r.rol !== 'auto' ? q.nombre : '', r.estado].filter(Boolean).join(' · ');
  }
  async function spGuardar(reg) {
    const x = await spResolver();
    const base = `/sites/${x.sitio}/lists/${x.lista}/items`;
    const fields = { Title: tituloFila(reg).slice(0, 250), Datos: JSON.stringify(reg) };
    const item = itemDe.get(reg.id);
    if (item) await sp(`${base}/${item}/fields`, { method: 'PATCH', json: fields });
    else itemDe.set(reg.id, (await sp(base, { method: 'POST', json: { fields } })).id);
  }

  // ---------- Guardado automático ----------
  let tGuardar = null, porGuardar = null, estadoGuardado = '';
  function marcarGuardado(txt) { estadoGuardado = txt; const g = $('#guardado'); if (g) g.textContent = txt; }
  function programarGuardado(reg) {
    porGuardar = reg;
    marcarGuardado('Guardando…');
    clearTimeout(tGuardar);
    tGuardar = setTimeout(vaciarGuardado, DEMO ? 350 : 1200);
  }
  async function vaciarGuardado() {
    clearTimeout(tGuardar);
    const reg = porGuardar; porGuardar = null;
    if (!reg) return;
    try { await Store.guardar(reg); marcarGuardado('Guardado ' + new Date().toLocaleTimeString('es-EC', { hour: '2-digit', minute: '2-digit' })); }
    catch (e) { porGuardar = reg; marcarGuardado('Sin guardar: revisa la conexión'); aviso('No se pudo guardar: ' + e.message); }
  }
  window.addEventListener('beforeunload', e => { if (porGuardar) { vaciarGuardado(); e.preventDefault(); } });

  // ---------- UI común ----------
  function aviso(txt) {
    const a = $('#aviso'); a.textContent = txt; a.hidden = false;
    clearTimeout(aviso.t); aviso.t = setTimeout(() => { a.hidden = true; }, 3200);
  }
  function modal(html, alAbrir) {
    const m = $('#modal');
    m.innerHTML = `<div class="modal-caja">${html}</div>`;
    m.hidden = false;
    m.onclick = e => { if (e.target === m || e.target.closest('[data-cerrar]')) cerrarModal(); };
    if (alAbrir) alAbrir(m);
  }
  function cerrarModal() { const m = $('#modal'); m.hidden = true; m.innerHTML = ''; }
  function confirmar(titulo, texto, si = 'Confirmar', peligro = false) {
    return new Promise(res => {
      modal(`<h3>${titulo}</h3><p class="sub">${texto}</p><div class="fila"><button class="btn sec" data-no>Cancelar</button><button class="btn ${peligro ? 'peligro' : ''}" data-si>${si}</button></div>`, m => {
        $('[data-si]', m).onclick = () => { cerrarModal(); res(true); };
        $('[data-no]', m).onclick = () => { cerrarModal(); res(false); };
      });
    });
  }
  function pedirTexto(titulo, texto, etiqueta, obligatorio = true) {
    return new Promise(res => {
      modal(`<h3>${titulo}</h3><p class="sub">${texto}</p><label class="campo-form">${etiqueta}<textarea id="m-txt"></textarea></label>
        <div class="fila"><button class="btn sec" data-no>Cancelar</button><button class="btn" data-si>Aceptar</button></div>`, m => {
        const t = $('#m-txt', m); t.focus();
        $('[data-si]', m).onclick = () => { if (obligatorio && !t.value.trim()) { t.focus(); return; } cerrarModal(); res(t.value.trim()); };
        $('[data-no]', m).onclick = () => { cerrarModal(); res(null); };
      });
    });
  }
  function confeti() {
    const c = document.createElement('div'); c.className = 'confeti';
    const col = Object.values(COLORES);
    for (let i = 0; i < 90; i++) {
      const p = document.createElement('i');
      p.style.left = Math.random() * 100 + 'vw'; p.style.background = col[i % col.length];
      p.style.animationDuration = 1.8 + Math.random() * 1.8 + 's'; p.style.animationDelay = Math.random() * .5 + 's';
      c.appendChild(p);
    }
    document.body.appendChild(c); setTimeout(() => c.remove(), 4500);
  }
  function anillo(pct, tam = 110, grosor = 11, color = '#fff', fondo = 'rgba(255,255,255,.22)', texto = true) {
    const r = (tam - grosor) / 2, L = 2 * Math.PI * r;
    return `<svg class="anillo-grande" width="${tam}" height="${tam}" viewBox="0 0 ${tam} ${tam}" aria-label="${pct}%">
      <circle cx="${tam / 2}" cy="${tam / 2}" r="${r}" fill="none" stroke="${fondo}" stroke-width="${grosor}"/>
      <circle cx="${tam / 2}" cy="${tam / 2}" r="${r}" fill="none" stroke="${color}" stroke-width="${grosor}" stroke-linecap="round"
        stroke-dasharray="${(L * pct) / 100} ${L}" transform="rotate(-90 ${tam / 2} ${tam / 2})"/>
      ${texto ? `<text x="50%" y="54%" text-anchor="middle" dominant-baseline="middle" font-size="${tam / 4.4}" font-weight="800" fill="${color}">${pct}%</text>` : ''}</svg>`;
  }

  // ---------- Ficha (formato de Talento Humano) ----------
  function valorHTML(c, v) {
    if (vacio(v)) return '<span class="vacio-campo">Por completar</span>';
    if (c.tipo === 'acciones') return (Array.isArray(v) ? v : [v]).filter(x => String(x || '').trim()).map(esc).join('<br>');
    if (c.tipo === 'competencias') return Object.entries(v).filter(([, t]) => String(t || '').trim()).map(([k, t]) => `<b>${esc(k)}:</b> ${esc(t)}`).join('<br>');
    return esc(v);
  }
  function fichaHTML(ev, valores, op = {}) {
    const campos = op.campos || camposAcuerdo();
    const zona = z => campos.filter(c => (c.zona || 'extra') === z);
    const ic = c => `<span class="icono-campo" style="background:${COLORES[c.color] || COLORES.verde}">${ICONOS[c.color] || '•'}</span>`;
    const tarjetas = zona('tarjeta').map(c => `<div class="f-card ${op.vivo === c.id ? 'vivo' : ''}"><h4 style="color:${COLORES[c.color] || COLORES.verde}">${ic(c)}${esc(c.etiqueta)}</h4><p>${valorHTML(c, valores[c.id])}</p></div>`).join('');
    const ind = zona('indicador');
    const indHTML = ind.length ? `<div class="f-card f-ind ${ind.some(c => op.vivo === c.id) ? 'vivo' : ''}">${ind.map((c, k) => k === 0
      ? `<h4 style="color:var(--verde)">${esc(c.etiqueta)}</h4><div class="grande">${valorHTML(c, valores[c.id])}</div>`
      : `<div class="grande">${esc(c.etiqueta)}: ${valorHTML(c, valores[c.id])}</div>`).join('')}</div>` : '';
    const plan = zona('plan');
    const planHTML = plan.length ? `<div class="f-card f-plan ${plan.some(c => op.vivo === c.id) ? 'vivo' : ''}">${plan.map(c => {
      const v = valores[c.id];
      const items = (Array.isArray(v) ? v : (v ? [v] : [])).filter(x => String(x || '').trim());
      return `<h4 style="color:var(--verde)">${esc(c.etiqueta)}</h4>${items.length ? `<ol>${items.map((x, k) => `<li><span class="num">${k + 1}</span><span>${esc(x)}</span></li>`).join('')}</ol>` : '<p><span class="vacio-campo">Por completar</span></p>'}`;
    }).join('')}<div class="f-seg">Seguimiento sugerido: ${CONF.ciclo.seguimiento.join(' · ')} días</div></div>` : '';
    const extra = zona('extra').filter(c => !op.final || !vacio(valores[c.id]));
    const extraHTML = extra.length ? `<div class="f-extra"><h4>Información adicional</h4><dl>${extra.map(c => `<dt>${esc(c.etiqueta)}</dt><dd>${valorHTML(c, valores[c.id])}</dd>`).join('')}</dl></div>` : '';
    const c = cargoDe(ev);
    return `<article class="ficha ${op.mini ? 'mini' : ''}">
      <header class="ficha-cab"><h2>${esc(op.titulo || (ev ? ev.nombre : ''))}</h2><span>${esc(c.nombre)}</span></header>
      <div class="ficha-cuerpo">
        ${tarjetas ? `<div class="ficha-tres">${tarjetas}</div>` : ''}
        ${indHTML || planHTML ? `<div class="ficha-dos">${indHTML}${planHTML}</div>` : ''}
        ${extraHTML}
      </div></article>`;
  }

  // ---------- Tareas de la persona que entra ----------
  function misTareas() {
    const t = { auto: null, equipo: [], otras: [] };
    if (!YO) return t;
    for (const ev of evaluados()) {
      for (const f of fuentesDe(ev)) {
        if (f.evaluador !== YO.id) continue;
        const r = resp(ev.id, f.rol, YO.id);
        const item = { ev, rol: f.rol, r, prog: progreso(r, camposDe(f.rol)) };
        if (f.rol === 'auto') t.auto = item;
        else if (f.rol === 'jefe') t.equipo.push(item);
        else t.otras.push(item);
      }
    }
    return t;
  }

  // ---------- Barra superior ----------
  function pintarBarra(vista) {
    $('#barra').hidden = false;
    const links = [['inicio', 'Mi espacio']];
    if (esAdmin()) links.push(['panel', 'Panel Talento Humano'], ['config', 'Configuración']);
    $('#nav').innerHTML = links.map(([v, t]) => `<a href="#/${v}" class="${vista === v ? 'activo' : ''}">${t}</a>`).join('');
    if (DEMO) {
      $('#yo').innerHTML = `<span class="demo-tag" title="Los datos quedan solo en este navegador">DEMO</span>
        <select id="ver-como" title="Ver la app como…">${CONF.personas.map(p => `<option value="${p.id}" ${YO && YO.id === p.id ? 'selected' : ''}>${esc(p.nombre)} · ${esc(cargoDe(p).nombre)}</option>`).join('')}</select>
        <button class="btn fant chico" id="reiniciar" title="Borrar los datos de prueba">↺</button>`;
      $('#ver-como').onchange = e => { YO = persona(e.target.value); cache.guardar('demo.yo', YO.id); location.hash = '#/inicio'; pintar(); };
      $('#reiniciar').onclick = async () => {
        if (!await confirmar('¿Reiniciar la demo?', 'Se borran los datos de prueba de este navegador y se vuelven a cargar los ejemplos.', 'Reiniciar', true)) return;
        cache.quitar('demo.filas'); B = null; await Store.cargar(); YO = persona(YO && YO.id) || persona('dc'); pintar(); aviso('Demo reiniciada');
      };
    } else {
      $('#yo').innerHTML = `${avatar(YO, 'p')}<span>${esc(YO ? primerNombre(YO) : CORREO)}</span><button class="btn fant chico" id="salir">Salir</button>`;
      $('#salir').onclick = () => pca.logoutRedirect({ account: cuenta, postLogoutRedirectUri: URL_APP });
    }
  }

  // ---------- Router ----------
  function ruta() { return location.hash.replace(/^#\/?/, '').split('/').filter(Boolean).map(decodeURIComponent); }
  let vistaActual = '';
  async function pintar() {
    if (porGuardar) await vaciarGuardado();
    cerrarModal();
    const [vista = 'inicio', ...a] = ruta();
    vistaActual = vista;
    pintarBarra(vista);
    window.scrollTo(0, 0);
    try {
      if (vista === 'r') return vResponder(a[0], a[1], a[2]);
      if (vista === 'acuerdo') return vAcuerdo(a[0]);
      if (vista === 'ficha') return vFicha(a[0]);
      if (vista === 'panel') return vPanel();
      if (vista === 'persona') return vPersona(a[0]);
      if (vista === 'config') return vConfig(a[0]);
      return vInicio();
    } catch (e) {
      console.error(e);
      app.innerHTML = `<div class="tarjeta vacio"><div class="grande">⚠️</div><h2>Algo salió mal</h2><p>${esc(e.message)}</p><a class="btn" href="#/inicio">Volver</a></div>`;
    }
  }
  window.addEventListener('hashchange', pintar);
  const sinPermiso = () => { app.innerHTML = `<div class="tarjeta vacio"><div class="grande">🔒</div><h2>Esta sección no está disponible para tu cuenta</h2><p>Si crees que es un error, escribe a Talento Humano.</p><a class="btn" href="#/inicio">Volver a mi espacio</a></div>`; };

  // ---------- Vista: Mi espacio ----------
  function tarjetaTarea(it) {
    const { ev, rol, r, prog } = it;
    const env = enviado(r);
    const titulo = rol === 'auto' ? 'Tu autoevaluación' : rol === 'jefe' ? 'Evaluación como jefe' : ROLES[rol].de;
    const cta = env ? 'Ver lo que enviaste' : r ? 'Continuar' : 'Comenzar';
    const anon = (rol === 'par' || rol === 'colab') && CONF.ciclo.anonimo ? '<span class="chico-txt">· anónima</span>' : '';
    return `<a class="tarjeta tarea ${env ? 'hecha' : ''} ${!env && rol === 'auto' ? 'destacada' : ''}" href="#/r/${ev.id}/${rol}">
      <div class="tarea-cab">${avatar(ev)}<div><span class="tipo">${titulo}</span> ${anon}<b>${esc(ev.nombre)}</b><span class="sub">${esc(cargoDe(ev).nombre)}</span></div></div>
      <div class="barra-prog"><i style="width:${env ? 100 : prog.pct}%"></i></div>
      <div class="tarea-pie">${pillEstado(r)}<span class="btn chico ${env ? 'sec' : ''}">${cta} →</span></div></a>`;
  }
  function tarjetaEquipo(ev) {
    const e = estadoDe(ev);
    const r = e.jefe;
    let cta, href;
    if (e.ac && (e.ac.estado === 'cerrado' || e.ac.estado === 'confirmado')) { cta = 'Ver ficha y seguimiento'; href = `#/ficha/${ev.id}`; }
    else if (e.listoAcuerdo) { cta = e.ac ? 'Continuar la conversación' : 'Abrir sala de acuerdo'; href = `#/acuerdo/${ev.id}`; }
    else if (!enviado(r)) { cta = r ? 'Continuar tu evaluación' : jefeVeAuto() && !enviado(e.auto) ? 'Evaluar (aún sin su autoevaluación)' : 'Evaluar sobre su autoevaluación'; href = `#/r/${ev.id}/jefe`; }
    else { cta = 'Ver tu evaluación'; href = `#/r/${ev.id}/jefe`; }
    const paso = (ok, b, t, txt) => `<i class="${ok ? 'ok' : b ? 'b' : ''}" title="${t}">${ok ? '✓' : txt}</i>`;
    return `<a class="tarjeta tarea ${e.listoAcuerdo && !e.ac ? 'destacada' : ''}" href="${href}">
      <div class="tarea-cab">${avatar(ev)}<div><span class="tipo">Tu equipo</span><b>${esc(ev.nombre)}</b><span class="sub">${esc(cargoDe(ev).nombre)}</span></div></div>
      <div class="tubo">${paso(enviado(e.auto), e.auto, 'Autoevaluación', 'A')}<span class="sep"></span>${paso(enviado(r), r, 'Jefe', 'J')}<span class="sep"></span>${paso(e.ac && e.ac.estado !== 'borrador', e.ac, 'Acuerdo', 'Ac')}<span class="sep"></span>${paso(e.ac && e.ac.estado === 'confirmado', false, 'Confirmado', 'C')}
        ${e.otras.length ? `<span class="chico-txt" style="margin-left:8px">360: ${e.otrasEnv}/${e.otras.length}</span>` : ''}</div>
      <div class="tarea-pie"><span class="pill ${e.clase}">${e.etapa}</span><span class="btn chico">${cta} →</span></div></a>`;
  }
  function vInicio() {
    if (!YO) {
      if (esAdmin()) { location.hash = '#/panel'; return; }
      app.innerHTML = `<div class="tarjeta ingreso"><span class="logo">360</span><h2>Tu cuenta no participa en este ciclo</h2>
        <p class="sub">Ingresaste como <b>${esc(CORREO)}</b>. Si deberías participar, pide a Talento Humano que registre tu correo en la configuración.</p></div>`;
      return;
    }
    const t = misTareas();
    const lista = [t.auto, ...t.equipo, ...t.otras].filter(Boolean);
    const hechas = lista.filter(x => enviado(x.r)).length;
    const pct = lista.length ? Math.round((hechas / lista.length) * 100) : 100;
    const dias = diasHasta(CONF.ciclo.cierre);
    const ac = acuerdoDe(YO.id);
    const equipo = evaluados().filter(p => p.jefe === YO.id);
    const miFicha = ac && (ac.estado === 'cerrado' || ac.estado === 'confirmado');
    app.innerHTML = `
      <section class="hero">
        <h1>Hola, ${esc(primerNombre(YO))} 👋</h1>
        <p>${esEvaluado(YO)
          ? 'Esta evaluación es una conversación, no un examen: tú y tu jefe miran el mismo año desde dos lados y acuerdan juntos una fortaleza, una conducta a desarrollar y un plan con seguimiento.'
          : 'Aquí están las evaluaciones que te tocan como jefe o como parte de la mirada 360 de tus compañeros.'}</p>
        <div class="hero-dato">
          <div><b>${lista.length - hechas}</b><span>pendientes</span></div>
          <div><b>${dias >= 0 ? dias : 0}</b><span>${dias >= 0 ? 'días para el cierre' : 'ciclo cerrado'} (${fmtFecha(CONF.ciclo.cierre)})</span></div>
          <div><b>${esc(CONF.ciclo.nombre)}</b><span>ciclo activo</span></div>
        </div>
        ${lista.length ? anillo(pct) : ''}
      </section>
      ${miFicha ? `<div class="titulo-sec"><h2>Tu ficha acordada</h2><a href="#/ficha/${YO.id}">Ver completa y seguimiento →</a></div>
        <a href="#/ficha/${YO.id}" style="text-decoration:none;color:inherit">${ac.estado === 'cerrado' ? '<p class="pill info sin" style="margin:0 0 10px">Tu jefe cerró el acuerdo: falta tu confirmación</p>' : ''}${fichaHTML(YO, ac.campos, { final: true })}</a>` : ''}
      ${t.auto && !miFicha ? `<div class="titulo-sec"><h2>Tu evaluación</h2><span class="sub">${YO.jefe && jefeVeAuto() ? 'Cuando la envíes, tu jefe la verá para hacer la suya sobre tus respuestas' : 'Solo tú la ves hasta que la envíes'}</span></div><div class="tareas">${tarjetaTarea(t.auto)}</div>` : ''}
      ${equipo.length ? `<div class="titulo-sec"><h2>Tu equipo</h2><span class="sub">${equipo.length} persona${equipo.length > 1 ? 's' : ''} a cargo</span></div><div class="tareas">${equipo.map(tarjetaEquipo).join('')}</div>` : ''}
      ${t.otras.length ? `<div class="titulo-sec"><h2>Miradas 360 que te pidieron</h2><span class="sub">${CONF.ciclo.anonimo ? 'Se muestran sin tu nombre' : ''}</span></div><div class="tareas">${t.otras.map(tarjetaTarea).join('')}</div>` : ''}
      ${!lista.length && !equipo.length && !miFicha ? '<div class="tarjeta vacio"><div class="grande">🌱</div><p>No tienes evaluaciones pendientes en este ciclo.</p></div>' : ''}`;
  }

  // ---------- Vista: responder (pregunta por pregunta) ----------
  let R = null;
  function vResponder(evId, rol, evalId) {
    const ev = persona(evId);
    if (!ev || !ROLES[rol]) return sinPermiso();
    const evaluadorId = evalId || (YO && YO.id);
    const propio = YO && evaluadorId === YO.id && fuentesDe(ev).some(f => f.rol === rol && f.evaluador === YO.id);
    const jefeLee = YO && ev.jefe === YO.id && rol === 'auto' && jefeVeAuto() && enviado(resp(ev.id, 'auto', ev.id));
    if (!propio && !esAdmin() && !jefeLee) return sinPermiso();
    const existente = resp(ev.id, rol, evaluadorId);
    if (!propio || enviado(existente)) return vRespuestaLectura(ev, rol, evaluadorId, existente);
    const reg = existente ? clon(existente) : nuevoRegistro('respuesta', ev.id, rol, evaluadorId);
    const campos = camposDe(rol);
    const pasos = ['intro', ...campos.map(c => c.id), 'revision'];
    const guardadoPaso = Number(sessionStorage.getItem('e360.paso.' + reg.id) || 0);
    R = { reg, ev, rol, campos, pasos, i: Math.min(guardadoPaso, pasos.length - 1), vivo: null };
    estadoGuardado = existente ? 'Borrador guardado' : '';
    pintarPaso();
  }
  function irPaso(i) {
    R.i = Math.max(0, Math.min(R.pasos.length - 1, i));
    sessionStorage.setItem('e360.paso.' + R.reg.id, R.i);
    pintarPaso();
  }
  function cambiarValor(c, v) {
    const r = R.reg;
    if (!r.historial.length) hist(r, 'Comenzó ' + ROLES[R.rol].obj);
    r.campos[c.id] = v;
    R.vivo = c.id;
    const viv = $('#vivo'); if (viv) viv.innerHTML = fichaHTML(R.ev, r.campos, { mini: true, campos: R.campos, vivo: c.id, titulo: tituloFicha() });
    const p = $('#pistas'); if (p) p.innerHTML = pistasHTML(c, v);
    programarGuardado(r);
  }
  const tituloFicha = () => R.rol === 'auto' || R.rol === 'jefe' ? R.ev.nombre : 'Tu mirada sobre ' + primerNombre(R.ev);
  function pistasHTML(c, v) {
    if (c.tipo === 'competencias') { const n = Object.values(v || {}).filter(x => String(x || '').trim()).length; return `<span class="${n ? 'ok' : 'no'}">${n} de ${cargoDe(R.ev).competencias.length} competencias con evidencia</span>`; }
    if (c.tipo === 'acciones') {
      const xs = (v || []).filter(x => String(x || '').trim());
      const verbo = xs.length && xs.every(x => /^\s*[a-záéíóúñ]+(ar|er|ir)\b/i.test(x));
      return `<span class="${xs.length ? 'ok' : 'no'}">${xs.length} acción${xs.length === 1 ? '' : 'es'}</span><span class="${verbo ? 'ok' : 'no'}">Cada una empieza con un verbo</span><span class="${xs.length && xs.length <= (c.max || 5) ? 'ok' : 'no'}">Máximo ${c.max || 5}</span>`;
    }
    const t = String(v || ''), pal = t.trim() ? t.trim().split(/\s+/).length : 0;
    if (c.validar === 'medible') return `<span class="${/%|\d|porcentaje|n[uú]mero|cantidad|d[ií]as|horas|tasa|[ií]ndice|promedio/i.test(t) ? 'ok' : 'no'}">Tiene unidad (%, número, días…)</span><span class="${pal >= 4 ? 'ok' : 'no'}">Se entiende sin explicarlo</span>`;
    if (c.validar === 'meta') return `<span class="${/\d/.test(t) ? 'ok' : 'no'}">Tiene un número</span><span class="${/diari|seman|quincen|mensual|mes|bimest|trimest|semest|anual|año/i.test(t) ? 'ok' : 'no'}">Tiene frecuencia</span>`;
    if (c.tipo === 'corto') return `<span class="${pal >= 2 ? 'ok' : 'no'}">Respuesta clara</span>`;
    const comps = cargoDe(R.ev).competencias.filter(k => sinTildes(t).includes(sinTildes(k)));
    return `<span class="${pal >= 6 ? 'ok' : 'no'}">${pal < 6 ? 'Un poco más de detalle ayuda' : 'Buen nivel de detalle'} (${pal} palabras)</span>
      <span class="${/\d|por ejemplo|cuando|en el|en la|proyecto|caso/i.test(t) ? 'ok' : 'no'}">Incluye un ejemplo o hecho concreto</span>
      ${c.sugerirCompetencias ? `<span class="${comps.length ? 'ok' : 'no'}">Conectada con una competencia del cargo</span>` : ''}`;
  }
  function campoPasoHTML(c) {
    const v = R.reg.campos[c.id];
    const color = COLORES[c.color] || COLORES.verde;
    let entrada;
    if (c.tipo === 'acciones') {
      const xs = Array.isArray(v) ? v.slice() : (v ? [v] : []);
      while (xs.length < Math.min(c.max || 5, Math.max(1, xs.filter(x => String(x || '').trim()).length + 1))) xs.push('');
      entrada = `<div class="acciones-lista">${xs.map((x, k) => `<div class="accion-item"><span class="num ${k === 0 ? 'lima' : ''}">${k + 1}</span>
        <input class="entrada" data-acc="${k}" value="${esc(x)}" placeholder="${k === 0 ? 'Ej.: Definir plan de digitalización con hitos y responsables' : 'Otra acción (opcional)'}"></div>`).join('')}</div>`;
    } else if (c.tipo === 'competencias') {
      const o = v || {};
      const sa = autoParaJefe() && enviado(autoParaJefe()) ? (autoParaJefe().campos[c.id] || {}) : null;
      entrada = cargoDe(R.ev).competencias.map(k => `<div class="comp-item"><b>${esc(k)}</b><p>${esc((MAN.competencias[k] || {}).descripcion || '')}</p>
        ${sa && String(sa[k] || '').trim() ? `<p style="color:var(--teal)"><b>${esc(primerNombre(R.ev))}:</b> ${esc(sa[k])}</p>` : ''}
        <textarea data-comp="${esc(k)}" placeholder="Un hecho que lo muestre…">${esc(o[k] || '')}</textarea></div>`).join('') || '<p class="sub">Este cargo no tiene competencias registradas en el manual.</p>';
    } else if (c.tipo === 'corto') {
      entrada = `<input class="entrada corta" id="inp" value="${esc(v || '')}" placeholder="Escribe aquí…" autocomplete="off">`;
    } else {
      entrada = `<textarea class="entrada" id="inp" placeholder="Escribe aquí… (se guarda solo)">${esc(v || '')}</textarea>`;
    }
    const comps = c.sugerirCompetencias && c.tipo === 'largo' ? cargoDe(R.ev).competencias : [];
    return `<span class="etiqueta-campo" style="color:${color}"><span class="icono-campo" style="background:${color}">${ICONOS[c.color] || '•'}</span>${esc(c.etiqueta)}${c.obligatorio ? '' : ' · opcional'}</span>
      <h2 class="pregunta">${esc(preguntaDe(c, R.rol, R.ev))}</h2>
      <p class="ayuda">${esc(c.ayuda || '')}</p>
      ${suRespuestaHTML(c)}
      ${entrada}
      ${comps.length ? `<div class="chips" title="Competencias del cargo según el Manual de Funciones">${comps.map(k => `<button class="chip ${sinTildes(v).includes(sinTildes(k)) ? 'on' : ''}" data-chip="${esc(k)}" title="${esc((MAN.competencias[k] || {}).descripcion || '')}">${esc(k)}</button>`).join('')}</div>` : ''}
      <div class="pistas" id="pistas">${pistasHTML(c, v)}</div>
      ${c.ejemplo ? `<details class="ejemplo"><summary>Ver un ejemplo</summary><div>${esc(c.ejemplo).replace(/\n/g, '<br>')}</div></details>` : ''}`;
  }
  function suRespuestaHTML(c) {
    const a = autoParaJefe();
    if (!a || !enviado(a) || !(camposDe('auto').some(x => x.id === c.id))) return '';
    const va = a.campos[c.id];
    const pn = esc(primerNombre(R.ev));
    return `<div class="su-resp"><small>Lo que escribió ${pn}</small><div>${valorHTML(c, va)}</div>
      ${!vacio(va) && c.tipo !== 'competencias' ? `<button class="btn sec chico" data-partir="${c.id}">↓ Partir de su respuesta para particularizarla</button>` : ''}</div>`;
  }
  function introHTML() {
    const { ev, rol } = R;
    const c = cargoDe(ev);
    const quien = rol === 'auto' ? 'tu cargo' : `el cargo de ${primerNombre(ev)}`;
    const reglas = rol === 'auto'
      ? [['Tu jefe la verá', jefeVeAuto() ? 'Cuando la envíes, tu jefe verá cada respuesta tuya y escribirá la suya sobre ella.' : 'Tu jefe no la verá hasta la sala de acuerdo.'],
        ['Se guarda solo', 'Puedes salir y volver cuando quieras; seguirás donde quedaste.'],
        ['Luego conversan', 'Con las dos miradas, tu jefe abre la sala de acuerdo y cierran juntos tu ficha.']]
      : rol === 'jefe'
        ? [['Sobre su respuesta', !jefeVeAuto() ? 'Responde sin ver la autoevaluación.' : enviado(resp(ev.id, 'auto', ev.id)) ? `En cada pregunta verás lo que escribió ${primerNombre(ev)}: puedes partir de su texto y particularizarlo.` : `${primerNombre(ev)} aún no envía su autoevaluación. Te conviene esperar: cuando la envíe, la verás en cada pregunta.`],
          ['Se guarda solo', 'Puedes salir y volver cuando quieras.'],
          ['Después: acuerdo', 'Abrirás la sala de acuerdo para conversar y cerrar la ficha final.']]
        : [['Anónima', CONF.ciclo.anonimo ? `${primerNombre(ev)} y su jefe verán tu respuesta sin tu nombre.` : 'Tu nombre será visible.'],
          ['Breve', `Son ${R.campos.length} preguntas cortas.`],
          ['Constructiva', 'Piensa en hechos concretos que hayas visto.']];
    return `<span class="tipo">${ROLES[rol].de}</span>
      <h2 class="pregunta">${rol === 'auto' ? `${esc(primerNombre(ev))}, miremos tu año en ${R.campos.length} pasos` : `Tu mirada sobre ${esc(ev.nombre)}`}</h2>
      <p class="ayuda">Son preguntas abiertas: escribe como lo dirías en una conversación. Mientras respondes, a la derecha se va armando la ficha.</p>
      <div class="intro-cargo"><b>${esc(c.nombre)}</b>${c.area ? ` · ${esc(c.area)}` : ''}<br>${esc(c.mision || '')}
        ${c.competencias.length ? `<div class="comps" style="margin-top:10px">${c.competencias.map(k => `<span>${esc(k)}</span>`).join('')}</div>` : ''}
        <div class="chico-txt" style="margin-top:8px">Misión y competencias de ${quien} según el Manual de Funciones.</div></div>
      <div class="reglas">${reglas.map(([a, b]) => `<div><b>${a}</b>${b}</div>`).join('')}</div>
      ${autoParaJefe() && enviado(autoParaJefe()) ? `<h3 style="margin:22px 0 10px;font-size:15px">Así llenó ${esc(primerNombre(ev))} su autoevaluación</h3>${fichaHTML(ev, autoParaJefe().campos, { campos: camposDe('auto'), mini: true, titulo: 'Autoevaluación · ' + ev.nombre })}` : ''}`;
  }
  function revisionHTML() {
    const faltan = R.campos.filter(c => c.obligatorio && vacio(R.reg.campos[c.id]));
    return `<span class="tipo">Revisión final</span>
      <h2 class="pregunta">${faltan.length ? 'Casi listo' : 'Todo listo para enviar'}</h2>
      <p class="ayuda">${faltan.length ? `Falta: ${faltan.map(c => `<a href="#" data-ir-campo="${c.id}">${esc(c.etiqueta)}</a>`).join(', ')}.` : 'Revisa la ficha. Al enviar ya no se puede editar (Talento Humano puede reabrirla si hace falta).'}</p>
      ${fichaHTML(R.ev, R.reg.campos, { campos: R.campos, titulo: tituloFicha() })}`;
  }
  function pintarPaso() {
    const { pasos, i, reg } = R;
    const p = pasos[i];
    const campo = R.campos.find(c => c.id === p);
    const cuerpo = p === 'intro' ? introHTML() : p === 'revision' ? revisionHTML() : campoPasoHTML(campo);
    const faltan = R.campos.filter(c => c.obligatorio && vacio(reg.campos[c.id])).length;
    const boton = p === 'revision'
      ? `<button class="btn" data-act="enviar" ${faltan ? 'disabled' : ''}>Enviar ${R.rol === 'auto' ? 'mi autoevaluación' : 'evaluación'} ✓</button>`
      : `<span class="atajo"><kbd>Ctrl</kbd> + <kbd>Enter</kbd></span><button class="btn" data-act="sig">${p === 'intro' ? 'Empezar →' : 'Siguiente →'}</button>`;
    app.innerHTML = `<div class="flujo">
      <section class="tarjeta paso">
        <div class="paso-top"><a class="btn fant chico" href="#/inicio">← Mi espacio</a>
          <div class="puntos">${pasos.map((x, k) => { const c = R.campos.find(z => z.id === x); const ok = c ? !vacio(reg.campos[c.id]) : k < i;
            return `<i data-ir="${k}" class="${ok ? 'listo' : ''} ${k === i ? 'aqui' : ''}" title="${c ? esc(c.etiqueta) : x === 'intro' ? 'Inicio' : 'Revisión'}"></i>`; }).join('')}</div>
          <span class="chico-txt">${i}/${pasos.length - 1}</span></div>
        ${cuerpo}
        <div class="paso-pie">${i > 0 ? '<button class="btn sec" data-act="atras">← Atrás</button>' : ''}<span class="guardado" id="guardado">${esc(estadoGuardado)}</span><span class="espacio"></span>${boton}</div>
      </section>
      ${p === 'revision' ? '' : `<aside class="lateral"><h3>Así va quedando</h3><div id="vivo">${fichaHTML(R.ev, reg.campos, { mini: true, campos: R.campos, titulo: tituloFicha() })}</div></aside>`}
    </div>`;
    $$('[data-ir]').forEach(el => { el.onclick = () => irPaso(Number(el.dataset.ir)); });
    $$('[data-ir-campo]').forEach(el => { el.onclick = e => { e.preventDefault(); irPaso(pasos.indexOf(el.dataset.irCampo)); }; });
    const atras = $('[data-act="atras"]'); if (atras) atras.onclick = () => irPaso(i - 1);
    const sig = $('[data-act="sig"]'); if (sig) sig.onclick = siguiente;
    const env = $('[data-act="enviar"]'); if (env) env.onclick = enviarRespuesta;
    if (campo) enlazarCampo(campo);
  }
  function siguiente() {
    const c = R.campos.find(x => x.id === R.pasos[R.i]);
    if (c && c.obligatorio && vacio(R.reg.campos[c.id])) {
      const p = $('.paso'); p.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-8px)' }, { transform: 'translateX(8px)' }, { transform: 'translateX(0)' }], { duration: 260 });
      aviso('Esta respuesta es necesaria para seguir');
      return;
    }
    irPaso(R.i + 1);
  }
  function enlazarCampo(c) {
    const autoAlto = t => { t.style.height = 'auto'; t.style.height = Math.max(120, t.scrollHeight) + 'px'; };
    const inp = $('#inp');
    if (inp) {
      inp.oninput = () => { if (inp.tagName === 'TEXTAREA') autoAlto(inp); cambiarValor(c, inp.value); actualizarChips(inp.value); };
      if (inp.tagName === 'TEXTAREA') autoAlto(inp);
      inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length);
    }
    $$('[data-chip]').forEach(b => {
      b.onclick = () => {
        const k = b.dataset.chip;
        if (sinTildes(inp.value).includes(sinTildes(k))) return;
        inp.value = inp.value.trim() ? `${inp.value.trim()} (${k})` : `${k}: `;
        cambiarValor(c, inp.value); actualizarChips(inp.value); autoAlto(inp); inp.focus();
      };
    });
    const accs = $$('[data-acc]');
    const leerAcc = () => accs.map(x => x.value);
    accs.forEach((x, k) => {
      x.oninput = () => {
        cambiarValor(c, leerAcc());
        // Al escribir en la última acción aparece una nueva casilla (hasta el máximo).
        if (k === accs.length - 1 && x.value.trim() && accs.length < (c.max || 5)) { const pos = x.selectionStart; pintarPaso(); const n = $$('[data-acc]')[k]; n.focus(); n.setSelectionRange(pos, pos); }
      };
      x.onkeydown = e => { if (e.key === 'Enter' && !e.ctrlKey) { e.preventDefault(); const n = $$('[data-acc]')[k + 1]; if (n) n.focus(); } };
    });
    if (accs.length && !inp) accs[Math.max(0, accs.findIndex(x => !x.value.trim()))].focus();
    $$('[data-partir]').forEach(b => {
      b.onclick = async () => {
        const va = autoParaJefe().campos[c.id];
        if (!vacio(R.reg.campos[c.id]) && !await confirmar('¿Reemplazar tu texto?', 'Se pondrá la respuesta de ' + esc(primerNombre(R.ev)) + ' para que la ajustes.', 'Reemplazar')) return;
        cambiarValor(c, clon(va)); pintarPaso();
      };
    });
    $$('[data-comp]').forEach(t => {
      t.oninput = () => { const o = { ...(R.reg.campos[c.id] || {}) }; o[t.dataset.comp] = t.value; cambiarValor(c, o); };
    });
  }
  function actualizarChips(v) { $$('[data-chip]').forEach(b => b.classList.toggle('on', sinTildes(v).includes(sinTildes(b.dataset.chip)))); }
  document.addEventListener('keydown', e => {
    if (vistaActual !== 'r' || !R || !$('.paso')) return;
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); const s = $('[data-act="sig"]'); if (s) siguiente(); }
  });
  async function enviarRespuesta() {
    const { reg, ev, rol } = R;
    if (!await confirmar('¿Enviar?', rol === 'auto' ? 'Tu jefe la verá cuando los dos hayan enviado. Ya no podrás editarla.' : `Ya no podrás editarla. ${rol === 'jefe' ? `Cuando ${primerNombre(ev)} también envíe, se abre la sala de acuerdo.` : ''}`, 'Enviar')) return;
    // Acciones vacías fuera.
    R.campos.filter(c => c.tipo === 'acciones').forEach(c => { if (Array.isArray(reg.campos[c.id])) reg.campos[c.id] = reg.campos[c.id].filter(x => String(x || '').trim()); });
    reg.estado = 'enviado'; reg.enviado = ahora();
    hist(reg, 'Envió ' + ROLES[rol].obj);
    porGuardar = reg;
    await vaciarGuardado();
    if (porGuardar) return;
    sessionStorage.removeItem('e360.paso.' + reg.id);
    const e = estadoDe(ev);
    const siguienteTxt = rol === 'auto' ? (e.listoAcuerdo ? 'Tu jefe ya envió la suya: pronto te invitará a conversar.' : 'Cuando tu jefe envíe la suya, se abre la sala de acuerdo.')
      : rol === 'jefe' ? (e.listoAcuerdo ? `${primerNombre(ev)} ya envió su autoevaluación: puedes abrir la sala de acuerdo.` : `Falta la autoevaluación de ${primerNombre(ev)}.`)
        : 'Gracias por tu mirada: ayuda a que la evaluación sea más justa.';
    app.innerHTML = `<div class="tarjeta exito"><div class="check-anim">✓</div><h2>¡Enviado!</h2><p class="sub">${esc(siguienteTxt)}</p>
      <div class="fila" style="justify-content:center;margin-top:20px"><a class="btn sec" href="#/inicio">Volver a mi espacio</a>
      ${rol === 'jefe' && e.listoAcuerdo ? `<a class="btn" href="#/acuerdo/${ev.id}">Abrir sala de acuerdo →</a>` : ''}</div></div>`;
    confeti();
  }
  function vRespuestaLectura(ev, rol, evaluadorId, r) {
    const q = persona(evaluadorId);
    const anon = (rol === 'par' || rol === 'colab') && CONF.ciclo.anonimo && !esAdmin();
    const campos = camposDe(rol);
    app.innerHTML = `<div class="fila" style="margin-bottom:14px"><a class="btn fant chico" href="${esAdmin() && (!YO || evaluadorId !== YO.id) ? `#/persona/${ev.id}` : '#/inicio'}">← Volver</a></div>
      <div class="tarjeta esp-cab">${avatar(ev, 'g')}<div><span class="tipo">${ROLES[rol].de}</span><h2>${esc(ev.nombre)}</h2>
        <span class="sub">${anon ? 'Mirada anónima' : `Por ${esc(q ? q.nombre : evaluadorId)}`} · ${r ? (enviado(r) ? 'Enviada el ' + fmtFechaHora(r.enviado) : 'Borrador') : 'Sin iniciar'}</span></div>
        <span class="espacio"></span>${esAdmin() && enviado(r) ? '<button class="btn peligro chico" id="reabrir">Reabrir</button>' : ''}</div>
      <div style="margin-top:14px">${r ? fichaHTML(ev, r.campos, { campos, titulo: rol === 'auto' || rol === 'jefe' ? ev.nombre : `Mirada sobre ${ev.nombre}` }) : '<div class="tarjeta vacio">Aún no hay respuestas.</div>'}</div>`;
    const b = $('#reabrir');
    if (b) b.onclick = async () => {
      const motivo = await pedirTexto('Reabrir respuesta', 'La persona podrá editarla otra vez. Queda registrado en la trazabilidad.', 'Motivo');
      if (!motivo) return;
      const x = clon(r); x.estado = 'borrador'; delete x.enviado; hist(x, 'Talento Humano reabrió la respuesta', motivo);
      await Store.guardar(x); aviso('Respuesta reabierta'); pintar();
    };
  }

  // ---------- Vista: sala de acuerdo ----------
  let A = null;
  function similitud(a, b) {
    const txt = t => sinTildes(Array.isArray(t) ? t.join(' ') : typeof t === 'object' && t ? Object.values(t).join(' ') : t);
    const pal = t => new Set(txt(t).split(/[^a-z0-9%]+/).filter(w => w.length > 3 && !VACIAS.has(w)).map(w => w.slice(0, 6)));
    const nums = t => (txt(t).match(/\d+([.,]\d+)?/g) || []).sort().join('|');
    const x = pal(a), y = pal(b);
    if (!x.size || !y.size) return null;
    let com = 0; x.forEach(w => { if (y.has(w)) com++; });
    let s = com / Math.min(x.size, y.size);
    if (nums(a) && nums(b) && nums(a) !== nums(b)) s = Math.min(s, 0.3); // mismas palabras, otra cifra: no es "lo mismo"
    return s;
  }
  const VACIAS = new Set('para como esta este esto estos estas sobre entre cada pero porque cuando donde desde hasta tambien mucho muy mas menos todo todos toda todas tiene tienen hacer hace sus las los del con que una uno unos unas ser son fue han hay ante bajo segun sin tras durante mediante otro otra otros otras mismo misma debe deben puede pueden areas area definir revisar realizar hacer mejorar cerrar cada semana mensual semanal'.split(' '));
  function chipCoincidencia(s) {
    if (s == null) return '';
    if (s >= 0.45) return '<span class="pill ok coincide">Ven lo mismo</span>';
    if (s >= 0.15) return '<span class="pill info coincide">Coinciden en parte</span>';
    return '<span class="pill alerta coincide" title="Miradas distintas: buen punto para conversar">Miradas distintas · conversar</span>';
  }
  function vAcuerdo(evId) {
    const ev = persona(evId);
    if (!ev) return sinPermiso();
    const esJefe = YO && ev.jefe === YO.id;
    if (!esJefe && !esAdmin()) return sinPermiso();
    const e = estadoDe(ev);
    if (e.ac && e.ac.estado !== 'borrador') { location.hash = `#/ficha/${ev.id}`; return; }
    if (!e.listoAcuerdo) {
      app.innerHTML = `<div class="tarjeta exito"><div class="grande" style="font-size:44px">⏳</div><h2>La sala de acuerdo se abre con las dos miradas</h2>
        <div class="fila" style="justify-content:center;margin-top:14px"><span class="pill ${enviado(e.auto) ? 'ok' : 'borr'}">Autoevaluación de ${esc(primerNombre(ev))}: ${enviado(e.auto) ? 'enviada' : 'pendiente'}</span>
        ${ev.jefe ? `<span class="pill ${enviado(e.jefe) ? 'ok' : 'borr'}">Evaluación del jefe: ${enviado(e.jefe) ? 'enviada' : 'pendiente'}</span>` : ''}</div>
        <div class="fila" style="justify-content:center;margin-top:20px"><a class="btn sec" href="#/inicio">Volver</a>${esJefe && !enviado(e.jefe) ? `<a class="btn" href="#/r/${ev.id}/jefe">Completar mi evaluación</a>` : ''}</div></div>`;
      return;
    }
    const campos = camposAcuerdo();
    let reg = e.ac ? clon(e.ac) : null;
    if (!reg) {
      reg = nuevoRegistro('acuerdo', ev.id, 'acuerdo', '');
      const fuente = e.jefe || e.auto;
      campos.forEach(c => { if (fuente.campos[c.id] != null) reg.campos[c.id] = clon(fuente.campos[c.id]); });
      reg.fechaConversacion = new Date().toISOString().slice(0, 10);
    }
    A = { reg, ev, e, campos };
    estadoGuardado = e.ac ? 'Borrador guardado' : '';
    pintarAcuerdo();
  }
  function editorAcuerdo(c, v) {
    if (c.tipo === 'acciones') {
      const xs = (Array.isArray(v) ? v : v ? [v] : []).slice();
      while (xs.length < (c.max || 5)) xs.push('');
      return xs.map((x, k) => `<div class="accion-item" style="margin:4px 0"><span class="num ${k === 0 ? 'lima' : ''}">${k + 1}</span><input data-ac="${c.id}" data-k="${k}" value="${esc(x)}" placeholder="${k ? 'Opcional' : 'Acción acordada'}"></div>`).join('');
    }
    if (c.tipo === 'competencias') return cargoDe(A.ev).competencias.map(k => `<div style="margin:6px 0"><b style="font-size:13px">${esc(k)}</b><textarea data-ac="${c.id}" data-comp="${esc(k)}" rows="2">${esc((v || {})[k] || '')}</textarea></div>`).join('');
    if (c.tipo === 'corto') return `<input data-ac="${c.id}" value="${esc(v || '')}">`;
    return `<textarea data-ac="${c.id}" rows="3">${esc(v || '')}</textarea>`;
  }
  function pintarAcuerdo() {
    const { reg, ev, e, campos } = A;
    const otras = e.otras.map(f => resp(ev.id, f.rol, f.evaluador)).filter(enviado);
    const bloque = c => {
      const va = e.auto ? e.auto.campos[c.id] : null, vj = e.jefe ? e.jefe.campos[c.id] : null;
      const color = COLORES[c.color] || COLORES.verde;
      const otrasC = otras.filter(r => !vacio(r.campos[c.id]));
      return `<section class="tarjeta esp-campo">
        <div class="esp-campo-cab"><span class="icono-campo" style="background:${color}">${ICONOS[c.color] || '•'}</span><h3 style="color:${color};font-size:15px">${esc(c.etiqueta)}</h3>${c.tipo !== 'competencias' ? chipCoincidencia(similitud(va, vj)) : ''}</div>
        <div class="miradas">
          <div class="mirada auto"><small>Cómo se ve ${esc(primerNombre(ev))}</small>${valorHTML(c, va)}${!vacio(va) ? `<div class="usar"><button class="btn fant chico" data-usar="auto" data-c="${c.id}">Usar esta ↓</button></div>` : ''}</div>
          <div class="mirada jefe"><small>Cómo lo ve ${ev.jefe ? esc(primerNombre(persona(ev.jefe))) + ' (jefe)' : 'el jefe'}</small>${e.jefe ? valorHTML(c, vj) : '<span class="vacio-campo">Sin jefe asignado</span>'}${!vacio(vj) ? `<div class="usar"><button class="btn fant chico" data-usar="jefe" data-c="${c.id}">Usar esta ↓</button></div>` : ''}</div>
        </div>
        ${otrasC.length ? `<details class="otras"><summary>Otras miradas 360 (${otrasC.length})${CONF.ciclo.anonimo ? ' · anónimas' : ''}</summary><ul>${otrasC.map(r => `<li>${CONF.ciclo.anonimo ? '' : `<b>${esc((persona(r.evaluador) || {}).nombre || '')}:</b> `}${valorHTML(c, r.campos[c.id])}</li>`).join('')}</ul></details>` : ''}
        <div class="acuerdo-caja"><label>Lo que acordamos</label>${editorAcuerdo(c, reg.campos[c.id])}</div>
      </section>`;
    };
    app.innerHTML = `<div class="fila" style="margin-bottom:14px"><a class="btn fant chico" href="#/inicio">← Mi espacio</a></div>
      <div class="tarjeta esp-cab">${avatar(ev, 'g')}<div><span class="tipo">Sala de acuerdo</span><h2>${esc(ev.nombre)}</h2><span class="sub">${esc(cargoDe(ev).nombre)}</span></div>
        <span class="espacio"></span><div class="esp-fuentes"><span class="pill ok">Autoevaluación</span>${e.jefe ? '<span class="pill ok">Jefe</span>' : ''}${e.otras.length ? `<span class="pill info">360: ${otras.length}/${e.otras.length}</span>` : ''}</div></div>
      <div class="tarjeta" style="padding:16px 22px;margin-top:14px;line-height:1.5">
        <b>Cómo usar esta sala:</b> siéntense juntos (o en llamada). Para cada punto lean las dos miradas, conversen y escriban abajo lo que acuerdan.
        Donde dice <span class="pill alerta sin">Miradas distintas</span> está la conversación más valiosa.
        <label class="campo-form" style="max-width:240px">Fecha de la conversación<input type="date" id="fecha-conv" value="${esc(reg.fechaConversacion || '')}"></label>
      </div>
      ${campos.map(bloque).join('')}
      <div class="esp-pie"><span class="guardado" id="guardado">${esc(estadoGuardado)}</span><span class="espacio"></span>
        <button class="btn sec" id="previa">Vista previa de la ficha</button><button class="btn" id="cerrar-ac">Cerrar acuerdo ✓</button></div>`;
    const leer = () => {
      campos.forEach(c => {
        const els = $$(`[data-ac="${c.id}"]`);
        if (c.tipo === 'acciones') reg.campos[c.id] = els.map(x => x.value);
        else if (c.tipo === 'competencias') { const o = {}; els.forEach(x => { o[x.dataset.comp] = x.value; }); reg.campos[c.id] = o; }
        else if (els[0]) reg.campos[c.id] = els[0].value;
      });
      reg.fechaConversacion = $('#fecha-conv').value;
      if (!reg.historial.length) hist(reg, 'Abrió la sala de acuerdo');
      programarGuardado(reg);
    };
    $$('[data-ac], #fecha-conv').forEach(x => { x.oninput = leer; });
    $$('[data-usar]').forEach(b => {
      b.onclick = () => {
        const c = campos.find(x => x.id === b.dataset.c);
        const fuente = b.dataset.usar === 'auto' ? e.auto : e.jefe;
        reg.campos[c.id] = clon(fuente.campos[c.id]);
        if (!reg.historial.length) hist(reg, 'Abrió la sala de acuerdo');
        programarGuardado(reg);
        const y = window.scrollY; pintarAcuerdo(); window.scrollTo(0, y);
        aviso(`${c.etiqueta}: se copió la mirada ${b.dataset.usar === 'auto' ? 'de ' + primerNombre(ev) : 'del jefe'}`);
      };
    });
    $('#previa').onclick = () => modal(`${fichaHTML(ev, reg.campos)}<div class="fila"><button class="btn" data-cerrar>Cerrar</button></div>`, m => { m.firstElementChild.style.maxWidth = '900px'; });
    $('#cerrar-ac').onclick = async () => {
      const faltan = campos.filter(c => c.obligatorio && vacio(reg.campos[c.id]));
      if (faltan.length) { aviso('Falta acordar: ' + faltan.map(c => c.etiqueta).join(', ')); return; }
      if (!await confirmar('¿Cerrar el acuerdo?', `${esc(primerNombre(ev))} recibirá la ficha para confirmarla. Desde ahí empieza el seguimiento a ${CONF.ciclo.seguimiento.join(', ')} días.`, 'Cerrar acuerdo')) return;
      campos.filter(c => c.tipo === 'acciones').forEach(c => { reg.campos[c.id] = (reg.campos[c.id] || []).filter(x => String(x || '').trim()); });
      reg.estado = 'cerrado'; reg.cerrado = ahora(); reg.cerradoPor = YO ? YO.nombre : CORREO;
      hist(reg, 'Cerró el acuerdo', reg.fechaConversacion ? 'Conversación del ' + fmtFecha(reg.fechaConversacion) : '');
      porGuardar = reg; await vaciarGuardado();
      if (porGuardar) return;
      confeti(); location.hash = `#/ficha/${ev.id}`;
    };
  }

  // ---------- Vista: ficha final + seguimiento ----------
  function vFicha(evId) {
    const ev = persona(evId);
    if (!ev) return sinPermiso();
    const yoMismo = YO && YO.id === ev.id, esJefe = YO && ev.jefe === YO.id;
    if (!yoMismo && !esJefe && !esAdmin()) return sinPermiso();
    const ac = acuerdoDe(ev.id);
    if (!ac || ac.estado === 'borrador') {
      app.innerHTML = `<div class="tarjeta vacio"><div class="grande">📝</div><h2>La ficha aún no está acordada</h2><p>Se arma en la sala de acuerdo, cuando la persona y su jefe hayan enviado sus miradas.</p>
        <a class="btn" href="${esJefe || esAdmin() ? `#/acuerdo/${ev.id}` : '#/inicio'}">${esJefe || esAdmin() ? 'Ir a la sala de acuerdo' : 'Volver'}</a></div>`;
      return;
    }
    const base = ac.confirmado || ac.cerrado;
    const plan = camposAcuerdo().find(c => c.tipo === 'acciones');
    const acciones = plan ? (ac.campos[plan.id] || []) : [];
    const hitos = CONF.ciclo.seguimiento.map(d => {
      const s = (ac.seguimiento || {})[d];
      const fecha = sumarDias(base, d);
      const vencido = !s && new Date(fecha) < new Date();
      return `<div class="tarjeta hito ${s ? 'hecho' : ''} ${vencido ? 'vencido' : ''}">
        <div class="dia">${d}<small>días · ${fmtFecha(fecha)}</small></div>
        ${s ? `<div class="estado-acc">${acciones.map((x, k) => { const st = (s.acciones || [])[k] || 'sin'; return `<span class="pill ${st === 'cumplida' ? 'ok' : st === 'curso' ? 'info' : ''}" title="${esc(x)}">${k + 1}. ${st === 'cumplida' ? 'Cumplida' : st === 'curso' ? 'En curso' : 'Sin iniciar'}</span>`; }).join('')}</div>
          ${s.nota ? `<p style="margin:6px 0;font-size:14px">${esc(s.nota)}</p>` : ''}<span class="chico-txt">${esc(s.quien)} · ${fmtFechaHora(s.fecha)}</span>`
          : `<p class="sub" style="margin:8px 0">${vencido ? 'Revisión pendiente' : 'Próxima revisión'}</p>`}
        ${ac.estado === 'confirmado' && (esJefe || yoMismo || esAdmin()) ? `<div style="margin-top:10px"><button class="btn chico ${s ? 'sec' : ''}" data-seg="${d}">${s ? 'Actualizar' : 'Registrar avance'}</button></div>` : ''}
      </div>`;
    }).join('');
    app.innerHTML = `<div class="fila no-imprimir" style="margin-bottom:14px"><a class="btn fant chico" href="${esAdmin() && !yoMismo && !esJefe ? '#/panel' : '#/inicio'}">← Volver</a><span class="espacio"></span>
        ${esJefe || esAdmin() ? `<a class="btn sec chico" href="#/persona/${ev.id}">Trazabilidad</a>` : ''}<button class="btn sec chico" onclick="window.print()">Imprimir / PDF</button></div>
      ${ac.estado === 'cerrado' && yoMismo ? `<div class="tarjeta no-imprimir" style="padding:20px 22px;margin-bottom:14px;box-shadow:inset 0 0 0 2px var(--lima)">
          <h3>Tu jefe cerró el acuerdo. ¿Lo confirmas?</h3><p class="sub">Confirma que conversaron y que esta ficha refleja lo acordado. Si tienes una observación, déjala escrita: queda en el registro.</p>
          <label class="campo-form">Observación (opcional)<textarea id="obs"></textarea></label>
          <div class="fila" style="margin-top:12px"><button class="btn" id="confirmar-ac">Confirmo el acuerdo ✓</button></div></div>` : ''}
      ${ac.estado === 'cerrado' && !yoMismo ? `<p class="pill info sin no-imprimir" style="margin:0 0 12px">Esperando la confirmación de ${esc(primerNombre(ev))}</p>` : ''}
      ${fichaHTML(ev, ac.campos, { final: true })}
      <p class="chico-txt" style="margin:10px 4px">${esc(CONF.ciclo.nombre)} · Conversación: ${fmtFecha(ac.fechaConversacion) || '—'} · Cerrado por ${esc(ac.cerradoPor || '')} el ${fmtFechaHora(ac.cerrado)}
        ${ac.confirmado ? ` · Confirmado por ${esc(ev.nombre)} el ${fmtFechaHora(ac.confirmado)}` : ''}${ac.observacion ? ` · Observación: “${esc(ac.observacion)}”` : ''}</p>
      <div class="titulo-sec"><h2>Seguimiento</h2><span class="sub">${ac.estado === 'confirmado' ? 'Cuenta desde la confirmación' : 'Empieza cuando se confirme el acuerdo'}</span></div>
      <div class="linea-tiempo">${hitos}</div>`;
    const b = $('#confirmar-ac');
    if (b) b.onclick = async () => {
      const x = clon(ac); x.estado = 'confirmado'; x.confirmado = ahora(); x.observacion = $('#obs').value.trim();
      hist(x, 'Confirmó el acuerdo', x.observacion);
      await Store.guardar(x); confeti(); aviso('Acuerdo confirmado'); pintar();
    };
    $$('[data-seg]').forEach(btn => { btn.onclick = () => registrarAvance(ev, ac, Number(btn.dataset.seg), acciones); });
  }
  function registrarAvance(ev, ac, d, acciones) {
    const s = (ac.seguimiento || {})[d] || {};
    modal(`<h3>Revisión a los ${d} días</h3><p class="sub">${esc(ev.nombre)} · ¿cómo van las acciones del plan?</p>
      ${acciones.map((x, k) => `<div style="margin:12px 0"><div style="font-size:14px;margin-bottom:6px"><span class="num" style="display:inline-grid;width:22px;height:22px;font-size:11px;margin-right:6px">${k + 1}</span>${esc(x)}</div>
        <div class="tabs" style="margin:0" data-acc-est="${k}">${[['sin', 'Sin iniciar'], ['curso', 'En curso'], ['cumplida', 'Cumplida']].map(([v, t]) => `<button type="button" data-v="${v}" class="${((s.acciones || [])[k] || 'sin') === v ? 'on' : ''}">${t}</button>`).join('')}</div></div>`).join('')}
      <label class="campo-form">Nota de la revisión<textarea id="seg-nota" placeholder="Evidencias, ajustes, compromisos…">${esc(s.nota || '')}</textarea></label>
      <div class="fila"><button class="btn sec" data-cerrar>Cancelar</button><button class="btn" id="seg-ok">Guardar avance</button></div>`, m => {
      $$('[data-acc-est]', m).forEach(g => $$('button', g).forEach(bt => { bt.onclick = () => { $$('button', g).forEach(z => z.classList.remove('on')); bt.classList.add('on'); }; }));
      $('#seg-ok', m).onclick = async () => {
        const x = clon(ac); x.seguimiento = x.seguimiento || {};
        x.seguimiento[d] = { fecha: ahora(), quien: YO ? YO.nombre : CORREO, nota: $('#seg-nota', m).value.trim(),
          acciones: $$('[data-acc-est]', m).map(g => ($('button.on', g) || {}).dataset.v || 'sin') };
        const cumplidas = x.seguimiento[d].acciones.filter(v => v === 'cumplida').length;
        hist(x, `Registró la revisión de ${d} días`, `${cumplidas}/${acciones.length} acciones cumplidas`);
        await Store.guardar(x); cerrarModal(); aviso('Avance registrado'); pintar();
      };
    });
  }

  // ---------- Vista: panel de Talento Humano ----------
  let filtroPanel = { q: '', area: '', etapa: '' };
  function vPanel() {
    if (!esAdmin()) return sinPermiso();
    const evs = evaluados();
    const est = new Map(evs.map(p => [p.id, estadoDe(p)]));
    const n = evs.length || 1;
    const contar = f => evs.filter(p => f(est.get(p.id), p)).length;
    const kAuto = contar(e => enviado(e.auto)), kJefe = contar((e, p) => enviado(e.jefe) || !p.jefe);
    const kAc = contar(e => e.ac && e.ac.estado !== 'borrador'), kConf = contar(e => e.ac && e.ac.estado === 'confirmado');
    const kSeg = contar(e => e.hitos.some(Boolean));
    const etapas = [...new Set([...est.values()].map(e => e.etapa))];
    const areas = [...new Set(evs.map(p => cargoDe(p).area).filter(Boolean))].sort();
    const visibles = evs.filter(p => {
      const e = est.get(p.id), q = sinTildes(filtroPanel.q);
      return (!q || sinTildes(p.nombre + ' ' + cargoDe(p).nombre).includes(q)) && (!filtroPanel.area || cargoDe(p).area === filtroPanel.area) && (!filtroPanel.etapa || e.etapa === filtroPanel.etapa);
    });
    const tubo = (e, p) => {
      const i = (ok, b, t, txt) => `<i class="${ok ? 'ok' : b ? 'b' : ''}" title="${t}">${ok ? '✓' : txt}</i>`;
      return `<div class="tubo">${i(enviado(e.auto), e.auto, 'Autoevaluación', 'A')}${i(enviado(e.jefe) || !p.jefe, e.jefe, 'Jefe', 'J')}
        ${e.otras.length ? `<i class="${e.otrasEnv === e.otras.length ? 'ok' : e.otrasEnv ? 'b' : ''}" title="Miradas 360" style="width:auto;padding:0 6px">${e.otrasEnv}/${e.otras.length}</i>` : ''}
        <span class="sep"></span>${i(e.ac && e.ac.estado !== 'borrador', e.ac, 'Acuerdo', 'Ac')}${i(e.ac && e.ac.estado === 'confirmado', false, 'Confirmado', 'C')}
        <span class="sep"></span>${CONF.ciclo.seguimiento.map((d, k) => `<i class="${e.hitos[k] ? 'lima' : ''}" title="Seguimiento ${d} días">${d}</i>`).join('')}</div>`;
    };
    // Temas de desarrollo: competencias mencionadas en las conductas a desarrollar (insumo para el plan de capacitación).
    const textosConducta = evs.map(p => { const e = est.get(p.id); const r = (e.ac && e.ac.estado !== 'borrador') ? e.ac : e.jefe; return r ? [r.campos.conducta, r.campos.plan].flat().join(' ') : ''; }).filter(Boolean);
    const conteoComp = Object.keys(MAN.competencias).map(k => [k, textosConducta.filter(t => sinTildes(t).includes(sinTildes(k))).length]).filter(x => x[1]).sort((a, b) => b[1] - a[1]);
    const conteoPal = {};
    textosConducta.forEach(t => new Set(sinTildes(t).split(/[^a-z]+/).filter(w => w.length > 5 && !VACIAS.has(w))).forEach(w => { conteoPal[w] = (conteoPal[w] || 0) + 1; }));
    const temas = Object.entries(conteoPal).sort((a, b) => b[1] - a[1]).slice(0, 18);
    const embudo = [['Autoevaluaciones', kAuto], ['Evaluaciones de jefe', kJefe], ['Acuerdos cerrados', kAc], ['Confirmados', kConf], ['Con seguimiento', kSeg]];
    app.innerHTML = `
      <div class="titulo-sec" style="margin-top:0"><div><h1 style="font-size:26px">Panel de Talento Humano</h1><span class="sub">${esc(CONF.ciclo.nombre)} · ${fmtFecha(CONF.ciclo.inicio)} al ${fmtFecha(CONF.ciclo.cierre)}</span></div>
        <div class="fila"><button class="btn sec" id="csv">Exportar a Excel (CSV)</button><a class="btn sec" href="#/config">Configurar</a></div></div>
      <div class="kpis">
        <div class="tarjeta kpi">${anillo(100, 46, 6, '#1f7a3d', '#e7ece6', false)}<div><b>${evs.length}</b><span>participantes</span></div></div>
        ${[['Autoevaluaciones', kAuto], ['Jefes', kJefe], ['Acuerdos', kAc], ['Confirmados', kConf]].map(([t, v]) => `<div class="tarjeta kpi">${anillo(Math.round((v / n) * 100), 46, 6, '#1f7a3d', '#e7ece6', false)}<div><b>${v}<small style="font-size:14px;color:var(--gris-2)">/${evs.length}</small></b><span>${t}</span></div></div>`).join('')}
      </div>
      <div class="dos-col" style="margin-top:14px">
        <div class="tarjeta"><div class="filtros"><b>Avance del ciclo</b></div><div class="embudo">${embudo.map(([t, v]) => `<div class="embudo-fila"><span>${t}</span><div class="barra-prog"><i style="width:${(v / n) * 100}%"></i></div><b>${Math.round((v / n) * 100)}%</b></div>`).join('')}</div></div>
        <div class="tarjeta"><div class="filtros"><b>Temas de desarrollo</b><span class="chico-txt">de las conductas y planes acordados · insumo para el plan de capacitación</span></div>
          ${conteoComp.length ? `<div class="embudo">${conteoComp.slice(0, 5).map(([k, v]) => `<div class="embudo-fila"><span>${esc(k)}</span><div class="barra-prog"><i style="width:${(v / Math.max(1, textosConducta.length)) * 100}%"></i></div><b>${v}</b></div>`).join('')}</div>` : ''}
          <div class="temas">${temas.length ? temas.map(([w, v]) => `<span class="tema" style="font-size:${12 + Math.min(8, v * 2)}px">${esc(w)}</span>`).join('') : '<span class="sub">Aparecerán cuando haya evaluaciones de jefe o acuerdos.</span>'}</div></div>
      </div>
      <div class="tarjeta" style="margin-top:14px">
        <div class="filtros"><input id="f-q" placeholder="Buscar persona o cargo…" value="${esc(filtroPanel.q)}">
          <select id="f-area"><option value="">Todas las áreas</option>${areas.map(a => `<option ${filtroPanel.area === a ? 'selected' : ''}>${esc(a)}</option>`).join('')}</select>
          <select id="f-etapa"><option value="">Todas las etapas</option>${etapas.map(a => `<option ${filtroPanel.etapa === a ? 'selected' : ''}>${esc(a)}</option>`).join('')}</select>
          <span class="espacio"></span><span class="chico-txt">${visibles.length} de ${evs.length}</span></div>
        <div class="tabla-wrap"><table class="tabla"><thead><tr><th>Persona</th><th>Jefe evaluador</th><th>Avance</th><th>Etapa</th><th>Última actividad</th><th></th></tr></thead><tbody>
        ${visibles.map(p => {
          const e = est.get(p.id), j = persona(p.jefe);
          const regs = [e.auto, e.jefe, e.ac].filter(Boolean);
          const ult = regs.map(r => r.modificado || r.creado).sort().pop();
          const pend = [];
          if (!enviado(e.auto)) pend.push(p);
          if (j && !enviado(e.jefe)) pend.push(j);
          const correos = pend.map(x => x.correo).filter(Boolean);
          return `<tr class="clic" data-p="${p.id}"><td><div class="persona-cel">${avatar(p, 'p')}<div><b>${esc(p.nombre)}</b><span class="chico-txt">${esc(cargoDe(p).nombre)}</span></div></div></td>
            <td>${j ? esc(j.nombre) : '<span class="pill alerta sin">Sin jefe asignado</span>'}</td><td>${tubo(e, p)}</td><td><span class="pill ${e.clase}">${e.etapa}</span></td>
            <td class="chico-txt">${ult ? fmtFechaHora(ult) : '—'}</td>
            <td>${correos.length ? `<a class="btn fant chico" data-stop href="mailto:${correos.join(',')}?subject=${encodeURIComponent('Evaluación 360 pendiente')}&body=${encodeURIComponent(`Hola, recuerda completar la ${CONF.ciclo.nombre} de ${p.nombre} antes del ${fmtFecha(CONF.ciclo.cierre)}: ${URL_APP}`)}" title="Recordar por correo">✉ Recordar</a>` : ''}</td></tr>`;
        }).join('') || '<tr><td colspan="6" class="vacio">Sin resultados</td></tr>'}
        </tbody></table></div>
      </div>`;
    const refiltrar = () => { filtroPanel = { q: $('#f-q').value, area: $('#f-area').value, etapa: $('#f-etapa').value }; const pos = $('#f-q').selectionStart; vPanel(); const q = $('#f-q'); q.focus(); q.setSelectionRange(pos, pos); };
    $('#f-q').oninput = refiltrar; $('#f-area').onchange = refiltrar; $('#f-etapa').onchange = refiltrar;
    $$('tr[data-p]').forEach(tr => { tr.onclick = e => { if (!e.target.closest('[data-stop]')) location.hash = '#/persona/' + tr.dataset.p; }; });
    $('#csv').onclick = exportarCSV;
  }
  function exportarCSV() {
    const campos = CONF.campos.filter(c => c.activo);
    const val = (c, v) => vacio(v) ? '' : c.tipo === 'acciones' ? (Array.isArray(v) ? v : [v]).filter(Boolean).map((x, k) => `${k + 1}. ${x}`).join(' | ')
      : c.tipo === 'competencias' ? Object.entries(v).filter(([, t]) => t).map(([k, t]) => `${k}: ${t}`).join(' | ') : String(v);
    const cab = ['Ciclo', 'Evaluado', 'Cargo', 'Área', 'Tipo', 'Evaluador', 'Estado', 'Enviado/Cerrado', 'Confirmado', ...campos.map(c => c.etiqueta), ...CONF.ciclo.seguimiento.map(d => `Seguimiento ${d} días`)];
    const filasCSV = [cab];
    [...filas.values()].filter(r => r.ciclo === CONF.ciclo.id && r.tipo !== 'config').sort((a, b) => a.id.localeCompare(b.id)).forEach(r => {
      const ev = persona(r.evaluado) || { nombre: r.evaluado }, q = persona(r.evaluador);
      const plan = campos.find(c => c.tipo === 'acciones');
      filasCSV.push([r.ciclo, ev.nombre, cargoDe(ev).nombre, cargoDe(ev).area, r.tipo === 'acuerdo' ? 'Acuerdo final' : ROLES[r.rol].nombre, q ? q.nombre : '', r.estado,
        fmtFechaHora(r.enviado || r.cerrado), fmtFechaHora(r.confirmado), ...campos.map(c => val(c, r.campos[c.id])),
        ...CONF.ciclo.seguimiento.map(d => { const s = (r.seguimiento || {})[d]; if (!s) return ''; const n = (s.acciones || []).filter(v => v === 'cumplida').length; return `${fmtFecha(s.fecha)} · ${n}/${plan ? (r.campos[plan.id] || []).length : 0} cumplidas${s.nota ? ' · ' + s.nota : ''}`; })]);
    });
    const csv = '﻿' + filasCSV.map(f => f.map(x => `"${String(x == null ? '' : x).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = `Evaluacion360_${CONF.ciclo.id}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  // ---------- Vista: persona (detalle y trazabilidad) ----------
  // Ficha del evaluado y ficha del jefe una junto a otra (y la acordada, si existe).
  function fichasComparadas(ev, e, esJefe) {
    const ver = r => r && (esAdmin() || enviado(r));
    const col = (r, rol, titulo) => `<div><div class="fila" style="margin-bottom:8px"><b>${titulo}</b>${pillEstado(r)}${r && r.enviado ? `<span class="chico-txt">${fmtFechaHora(r.enviado)}</span>` : ''}</div>
      ${ver(r) ? fichaHTML(ev, r.campos, { campos: camposDe(rol), titulo: ev.nombre }) : '<div class="tarjeta vacio">Aún no enviada.</div>'}</div>`;
    const autoVisible = esAdmin() || (esJefe && jefeVeAuto());
    return `<div class="titulo-sec"><h2>Fichas: evaluado y jefe</h2><button class="btn sec chico no-imprimir" onclick="window.print()">Imprimir / PDF</button></div>
      <div class="fichas-par">${autoVisible ? col(e.auto, 'auto', 'Llenada por ' + esc(ev.nombre)) : ''}${ev.jefe ? col(e.jefe, 'jefe', 'Llenada por el jefe · ' + esc((persona(ev.jefe) || {}).nombre || '')) : ''}</div>
      ${e.ac && e.ac.estado !== 'borrador' ? `<div class="titulo-sec"><h2>Ficha acordada</h2><a href="#/ficha/${ev.id}">Seguimiento →</a></div>${fichaHTML(ev, e.ac.campos, { final: true })}` : ''}`;
  }
  function vPersona(id) {
    const ev = persona(id);
    if (!ev) return sinPermiso();
    const esJefe = YO && ev.jefe === YO.id;
    if (!esAdmin() && !esJefe) return sinPermiso();
    const e = estadoDe(ev);
    const fuentes = fuentesDe(ev);
    const anon = !esAdmin() && CONF.ciclo.anonimo;
    const regs = [...fuentes.map(f => resp(ev.id, f.rol, f.evaluador)), e.ac].filter(Boolean);
    const eventos = regs.flatMap(r => (r.historial || []).map(h => ({ ...h, tipo: r.tipo === 'acuerdo' ? 'Acuerdo' : ROLES[r.rol].nombre, anonimo: anon && (r.rol === 'par' || r.rol === 'colab') })))
      .sort((a, b) => b.t.localeCompare(a.t));
    app.innerHTML = `<div class="fila" style="margin-bottom:14px"><a class="btn fant chico" href="${esAdmin() ? '#/panel' : '#/inicio'}">← Volver</a></div>
      <div class="tarjeta esp-cab">${avatar(ev, 'g')}<div><span class="tipo">${esc(CONF.ciclo.nombre)}</span><h2>${esc(ev.nombre)}</h2><span class="sub">${esc(cargoDe(ev).nombre)} · Jefe: ${esc((persona(ev.jefe) || {}).nombre || 'sin asignar')}</span></div>
        <span class="espacio"></span><span class="pill ${e.clase}">${e.etapa}</span></div>
      <div class="dos-col" style="margin-top:14px">
        <div class="tarjeta"><div class="filtros"><b>Fuentes de la evaluación</b></div>
          <table class="tabla"><tbody>${fuentes.map(f => { const r = resp(ev.id, f.rol, f.evaluador), q = persona(f.evaluador);
            const ver = esAdmin() || (esJefe && (f.rol === 'jefe' || (f.rol === 'auto' && jefeVeAuto() && enviado(r))));
            return `<tr><td><b>${ROLES[f.rol].nombre}</b><br><span class="chico-txt">${anon && f.rol !== 'auto' && f.rol !== 'jefe' ? 'Anónimo' : esc(q ? q.nombre : '')}</span></td><td>${pillEstado(r)}</td>
              <td class="chico-txt">${r && r.enviado ? fmtFechaHora(r.enviado) : ''}</td><td>${ver && r ? `<a class="btn fant chico" href="#/r/${ev.id}/${f.rol}/${f.evaluador}">Ver</a>` : ''}</td></tr>`; }).join('')}
            <tr><td><b>Acuerdo final</b></td><td>${e.ac ? `<span class="pill ${e.ac.estado === 'borrador' ? 'borr' : 'ok'}">${e.ac.estado === 'borrador' ? 'En conversación' : e.ac.estado === 'cerrado' ? 'Cerrado' : 'Confirmado'}</span>` : '<span class="pill">Pendiente</span>'}</td>
              <td class="chico-txt">${e.ac && e.ac.cerrado ? fmtFechaHora(e.ac.cerrado) : ''}</td><td>${e.ac && e.ac.estado !== 'borrador' ? `<a class="btn fant chico" href="#/ficha/${ev.id}">Ficha</a>` : e.listoAcuerdo ? `<a class="btn fant chico" href="#/acuerdo/${ev.id}">Sala</a>` : ''}</td></tr>
          </tbody></table>
          ${esAdmin() && e.ac && e.ac.estado !== 'borrador' ? '<div style="padding:12px 16px"><button class="btn peligro chico" id="reabrir-ac">Reabrir acuerdo</button></div>' : ''}</div>
        <div class="tarjeta"><div class="filtros"><b>Trazabilidad</b><span class="chico-txt">quién hizo qué y cuándo</span></div>
          <div style="padding:16px 20px">${eventos.length ? `<ul class="traza">${eventos.map(h => `<li class="${/Envi|Cerr|Confirm/.test(h.accion) ? 'fuerte' : ''}"><b>${h.anonimo ? 'Anónimo' : esc(h.quien)}</b> · ${esc(h.accion)} <span class="pill sin">${esc(h.tipo)}</span>${h.detalle ? `<br><span class="sub">${esc(h.detalle)}</span>` : ''}<time>${fmtFechaHora(h.t)}</time></li>`).join('')}</ul>` : '<p class="sub">Sin actividad todavía.</p>'}</div></div>
      </div>
      ${fichasComparadas(ev, e, esJefe)}`;
    const b = $('#reabrir-ac');
    if (b) b.onclick = async () => {
      const motivo = await pedirTexto('Reabrir acuerdo', 'El jefe podrá volver a editarlo en la sala de acuerdo. La confirmación y el seguimiento se conservan en el historial.', 'Motivo');
      if (!motivo) return;
      const x = clon(e.ac); x.estado = 'borrador'; delete x.confirmado; hist(x, 'Talento Humano reabrió el acuerdo', motivo);
      await Store.guardar(x); aviso('Acuerdo reabierto'); pintar();
    };
  }

  // ---------- Vista: configuración ----------
  let B = null; // borrador de configuración
  let campoAbierto = null;
  function vConfig(tab = 'ciclo') {
    if (!esAdmin()) return sinPermiso();
    if (!B) B = clon(CONF);
    const tabs = [['ciclo', 'Ciclo'], ['cargos', 'Cargos que participan'], ['personas', 'Personas y jefes'], ['campos', 'Campos de la ficha'], ['accesos', 'Accesos']];
    const fc = filas.get('config');
    const faltaBase = !DEMO && !(fc && fc.datos && fc.datos.manual && fc.datos.personas && fc.datos.personas.length);
    const cambios = faltaBase || JSON.stringify(B) !== JSON.stringify(CONF);
    const cont = { ciclo: confCiclo, cargos: confCargos, personas: confPersonas, campos: confCampos, accesos: confAccesos }[tab] || confCiclo;
    app.innerHTML = `<div class="titulo-sec" style="margin-top:0"><div><h1 style="font-size:26px">Configuración</h1><span class="sub">Lo que guardes aquí aplica a toda la empresa.</span></div></div>
      <div class="tabs">${tabs.map(([k, t]) => `<button data-tab="${k}" class="${tab === k ? 'on' : ''}">${t}</button>`).join('')}</div>
      <div id="conf">${cont()}</div>
      <div class="pegajoso"><div class="tarjeta">${faltaBase ? '<span class="pill alerta">Falta subir el manual y las personas a SharePoint</span>' : cambios ? '<span class="pill borr">Cambios sin guardar</span>' : '<span class="pill ok">Todo guardado</span>'}
        <button class="btn sec" id="descartar" ${cambios ? '' : 'disabled'}>Descartar</button><button class="btn" id="guardar-conf" ${cambios ? '' : 'disabled'}>Guardar cambios</button></div></div>`;
    $$('[data-tab]').forEach(b => { b.onclick = () => { location.hash = '#/config/' + b.dataset.tab; }; });
    $('#descartar').onclick = () => { B = null; vConfig(tab); };
    $('#guardar-conf').onclick = async () => {
      const ids = B.personas.map(p => p.id);
      if (new Set(ids).size !== ids.length) { aviso('Hay dos personas con el mismo identificador'); return; }
      const reg = filas.get('config') ? clon(filas.get('config')) : { id: 'config', tipo: 'config', historial: [], creado: ahora() };
      if (!MAN.cargos.length) { aviso('No hay datos del Manual de Funciones para guardar'); return; }
      reg.datos = { ...clon(B), manual: clon(MAN) };
      hist(reg, 'Guardó la configuración');
      try { await Store.guardar(reg); } catch (e) { aviso('No se pudo guardar: ' + e.message); return; }
      aplicarConfig(); B = null; aviso('Configuración guardada'); pintar();
    };
    enlazarConfig(tab);
  }
  const refrescarConfig = () => { const [, tab] = ruta(); const y = window.scrollY; vConfig(tab); window.scrollTo(0, y); };
  const interruptor = (attrs, on) => `<label class="interruptor"><input type="checkbox" ${attrs} ${on ? 'checked' : ''}><span></span></label>`;
  function confCiclo() {
    const c = B.ciclo;
    return `<div class="tarjeta conf-sec"><h3>Ciclo de evaluación</h3><p class="sub">Cambiar el identificador inicia un ciclo nuevo; lo anterior se conserva.</p>
      <div class="rejilla">
        <label class="campo-form">Nombre<input data-ciclo="nombre" value="${esc(c.nombre)}"></label>
        <label class="campo-form">Identificador<input data-ciclo="id" value="${esc(c.id)}"></label>
        <label class="campo-form">Inicio<input type="date" data-ciclo="inicio" value="${esc(c.inicio)}"></label>
        <label class="campo-form">Cierre<input type="date" data-ciclo="cierre" value="${esc(c.cierre)}"></label>
        <label class="campo-form">Seguimiento (días, separados por coma)<input data-ciclo="seguimiento" value="${esc(c.seguimiento.join(', '))}"></label>
        <label class="campo-form">Pares por persona (si se activan)<input type="number" min="1" max="10" data-ciclo="maxPares" value="${esc(c.maxPares)}"></label>
      </div></div>
      <div class="tarjeta conf-sec"><h3>Fuentes de la evaluación 360</h3><p class="sub">La autoevaluación y la del jefe inmediato van siempre. Agrega más miradas si el ciclo lo amerita.</p>
      ${[['pares', 'Pares', 'Compañeros con el mismo jefe que también participan (p. ej., las jefaturas que reportan a Gerencia General).'],
        ['colaboradores', 'Colaboradores', 'Las personas a cargo del evaluado (según el organigrama) responden una mirada corta.'],
        ['anonimo', 'Miradas 360 anónimas', 'El evaluado y su jefe ven las respuestas de pares y colaboradores sin nombre. Talento Humano sí ve quién respondió.'],
        ['jefeVeAuto', 'El jefe evalúa sobre la autoevaluación', 'En cada pregunta el jefe ve lo que escribió el evaluado y puede partir de su texto para particularizarlo. Si se apaga, el jefe responde sin verla.']]
        .map(([k, t, d]) => `<div class="cargo-fila">${interruptor(`data-ciclo-b="${k}"`, c[k])}<div class="info"><b>${t}</b><span class="sub">${d}</span></div></div>`).join('')}
      ${carga()}</div>`;
  }
  // Cuántas evaluaciones le tocan a cada persona con la configuración en edición (para no sobrecargar a nadie).
  function carga() {
    const real = CONF; CONF = B;
    const n = new Map();
    try { evaluados().forEach(ev => fuentesDe(ev).forEach(f => n.set(f.evaluador, (n.get(f.evaluador) || 0) + 1))); } finally { CONF = real; }
    const top = [...n.entries()].sort((a, b) => b[1] - a[1]);
    if (!top.length) return '';
    const total = top.reduce((s, x) => s + x[1], 0);
    const p = id => B.personas.find(x => x.id === id);
    return `<div class="intro-cargo" style="margin-bottom:0"><b>Carga con esta configuración:</b> ${total} evaluaciones en total, ${top.length} personas responden.
      Más cargadas: ${top.slice(0, 4).map(([id, k]) => `${esc((p(id) || {}).nombre || id)} (${k})`).join(', ')}.</div>`;
  }
  function confCargos() {
    const grupo = (nivel, titulo, nota) => {
      const cs = MAN.cargos.filter(c => c.nivel === nivel);
      const pordef = c => c.nivel === 'estrategico' && !NO_EVALUADOS.includes(c.id);
      return `<div class="tarjeta conf-sec"><div class="fila"><div><h3>${titulo}</h3><span class="sub">${nota}</span></div><span class="espacio"></span>
        <button class="btn fant chico" data-todos="${nivel}" data-v="1">Todos</button><button class="btn fant chico" data-todos="${nivel}" data-v="0">Ninguno</button></div>
        ${cs.map(c => { const on = B.cargos[c.id] !== undefined ? B.cargos[c.id] : pordef(c);
          const gente = B.personas.filter(p => p.cargo === c.id);
          return `<div class="cargo-fila ${on ? '' : 'off'}">${interruptor(`data-cargo="${c.id}"`, on)}<div class="info"><b>${esc(c.nombre)} <span class="chico-txt">${esc(c.codigo)} · ${esc(c.area)}</span></b>
            <span class="sub">Jefe inmediato (manual): ${esc(c.jefeInmediato || '—')} · ${gente.length ? gente.map(p => esc(p.nombre)).join(', ') : '<i>sin personas asignadas</i>'}</span>
            <div class="comps">${c.competencias.map(k => `<span>${esc(k)}</span>`).join('')}</div></div></div>`; }).join('')}</div>`;
    };
    return grupo('estrategico', 'Nivel estratégico', 'Según el Manual de Funciones v07: entran por defecto, excepto Presidente y Gerente General (el Gerente General solo evalúa a su equipo; nadie lo evalúa).')
      + grupo('operativo', 'Nivel operativo', 'No entran por defecto. Actívalos si un ciclo los incluye.');
  }
  function confPersonas() {
    const opCargo = sel => MAN.cargos.map(c => `<option value="${c.id}" ${sel === c.id ? 'selected' : ''}>${esc(c.nombre)}</option>`).join('');
    const opJefe = (p) => `<option value="">— Sin jefe —</option>` + B.personas.filter(x => x.id !== p.id).map(x => `<option value="${x.id}" ${p.jefe === x.id ? 'selected' : ''}>${esc(x.nombre)}</option>`).join('');
    const arbol = padre => {
      const hijos = B.personas.filter(p => (p.jefe || '') === padre && (padre || !p.jefe || !B.personas.some(x => x.id === p.jefe)));
      return hijos.length ? `<ul class="${padre ? '' : 'arbol'}">${hijos.map(p => {
        const on = B.cargos[p.cargo] !== undefined ? B.cargos[p.cargo] : (cargo(p.cargo).nivel === 'estrategico' && !NO_EVALUADOS.includes(p.cargo));
        return `<li><span class="nodo ${on ? 'si' : ''}">${avatar(p, 'p')}<span>${esc(p.nombre)}<small>${esc(cargo(p.cargo).nombre)}${on ? ' · participa' : ''}</small></span></span>${arbol(p.id)}</li>`; }).join('')}</ul>` : '';
    };
    return `<div class="tarjeta conf-sec"><div class="fila"><div><h3>Personas y jefe evaluador</h3><span class="sub">Tomado del organigrama del Manual de Funciones. El correo es el de Microsoft 365: con él la app reconoce a cada persona.</span></div>
        <span class="espacio"></span><button class="btn chico" id="add-persona">+ Persona</button></div>
      <div class="tabla-wrap"><table class="tabla"><thead><tr><th>Nombre</th><th>Correo</th><th>Cargo</th><th>Jefe evaluador</th><th title="Acceso al panel y la configuración">TTHH</th><th></th></tr></thead><tbody>
      ${B.personas.map((p, k) => `<tr><td><input class="inp" data-per="${k}" data-k="nombre" value="${esc(p.nombre)}" style="min-width:190px"></td>
        <td><input class="inp" data-per="${k}" data-k="correo" value="${esc(p.correo || '')}" placeholder="nombre@healthyfood.com.ec" style="min-width:220px"></td>
        <td><select class="inp" data-per="${k}" data-k="cargo" style="max-width:240px">${opCargo(p.cargo)}</select></td>
        <td><select class="inp" data-per="${k}" data-k="jefe" style="max-width:200px">${opJefe(p)}</select></td>
        <td><input type="checkbox" data-per="${k}" data-k="tthh" ${p.tthh ? 'checked' : ''}></td>
        <td><button class="btn fant chico" data-del-per="${k}" title="Quitar">✕</button></td></tr>`).join('')}
      </tbody></table></div></div>
      <div class="tarjeta conf-sec"><h3>Organigrama resultante</h3><p class="sub">Borde verde = participa en el ciclo.</p><div class="organi">${arbol('')}</div></div>`;
  }
  function confCampos() {
    return `<div class="tarjeta conf-sec"><div class="fila"><div><h3>Campos de la ficha</h3><span class="sub">Cada campo es una pregunta abierta. Elige quién la responde y dónde aparece en la ficha final.</span></div>
        <span class="espacio"></span><button class="btn sec chico" id="restaurar">Restaurar sugeridos</button><button class="btn chico" id="add-campo">+ Campo</button></div>
      <div style="margin-top:14px">${B.campos.map((c, k) => {
        const abierto = campoAbierto === c.id;
        return `<div class="campo-conf ${c.activo ? '' : 'off'}"><header data-abrir="${c.id}">${interruptor(`data-cam="${k}" data-k="activo" data-stop`, c.activo)}
          <span class="icono-campo" style="background:${COLORES[c.color] || COLORES.verde}">${ICONOS[c.color] || '•'}</span><b>${esc(c.etiqueta)}</b>
          <span class="pill sin">${TIPOS[c.tipo] || c.tipo}</span><span class="chico-txt">${(c.fuentes || []).map(f => ROLES[f].corto).join(' · ')}</span>
          <button class="btn fant chico" data-mover="${k}" data-d="-1" data-stop ${k === 0 ? 'disabled' : ''}>↑</button><button class="btn fant chico" data-mover="${k}" data-d="1" data-stop ${k === B.campos.length - 1 ? 'disabled' : ''}>↓</button></header>
          ${abierto ? `<div class="cuerpo"><div class="rejilla">
            <label class="campo-form">Nombre en la ficha<input data-cam="${k}" data-k="etiqueta" value="${esc(c.etiqueta)}"></label>
            <label class="campo-form">Tipo de respuesta<select data-cam="${k}" data-k="tipo">${Object.entries(TIPOS).map(([v, t]) => `<option value="${v}" ${c.tipo === v ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
            <label class="campo-form">Dónde aparece<select data-cam="${k}" data-k="zona">${Object.entries(ZONAS).map(([v, t]) => `<option value="${v}" ${c.zona === v ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
            <label class="campo-form">Color<select data-cam="${k}" data-k="color">${Object.keys(COLORES).map(v => `<option ${c.color === v ? 'selected' : ''}>${v}</option>`).join('')}</select></label></div>
            <div class="campo-form">Quién responde<div class="marcas">${TODOS.map(f => `<label><input type="checkbox" data-cam="${k}" data-fuente="${f}" ${(c.fuentes || []).includes(f) ? 'checked' : ''}>${ROLES[f].nombre}</label>`).join('')}
              <label style="margin-left:auto"><input type="checkbox" data-cam="${k}" data-k="obligatorio" ${c.obligatorio ? 'checked' : ''}>Obligatorio</label>
              ${c.tipo === 'largo' ? `<label><input type="checkbox" data-cam="${k}" data-k="sugerirCompetencias" ${c.sugerirCompetencias ? 'checked' : ''}>Sugerir competencias del cargo</label>` : ''}</div></div>
            <div class="rejilla">${TODOS.filter(f => (c.fuentes || []).includes(f)).map(f => `<label class="campo-form">Pregunta · ${ROLES[f].nombre}<textarea data-cam="${k}" data-preg="${f}" rows="2">${esc((c.pregunta || {})[f] || '')}</textarea></label>`).join('')}</div>
            <p class="chico-txt">Usa {nombre} para el nombre de la persona evaluada.</p>
            <div class="rejilla"><label class="campo-form">Ayuda (debajo de la pregunta)<textarea data-cam="${k}" data-k="ayuda" rows="2">${esc(c.ayuda || '')}</textarea></label>
              <label class="campo-form">Ejemplo${c.tipo === 'acciones' ? ' (una acción por línea)' : ''}<textarea data-cam="${k}" data-k="ejemplo" rows="2">${esc(c.ejemplo || '')}</textarea></label></div>
            ${!c.base ? `<div style="margin-top:12px"><button class="btn peligro chico" data-del-cam="${k}">Eliminar campo</button></div>` : ''}
          </div>` : ''}</div>`;
      }).join('')}</div></div>
      <div class="tarjeta conf-sec"><h3>Vista previa de la ficha</h3><p class="sub">Así se verá el acuerdo final (con el ejemplo de cada campo).</p>
        ${fichaHTML(B.personas[0] || null, Object.fromEntries(B.campos.map(c => [c.id, c.tipo === 'acciones' ? String(c.ejemplo || '').split('\n').filter(Boolean) : c.tipo === 'competencias' ? {} : c.ejemplo])), { campos: B.campos.filter(c => c.activo && (c.fuentes || []).some(f => f === 'auto' || f === 'jefe')), titulo: 'Nombre de la persona' })}</div>`;
  }
  function confAccesos() {
    return `<div class="tarjeta conf-sec"><h3>Administradores</h3><p class="sub">Pueden ver el panel, todas las respuestas, reabrir y configurar. Además, toda persona marcada como TTHH en "Personas".</p>
      <label class="campo-form">Correos (uno por línea)<textarea id="admins" rows="4">${esc((B.administradores || []).join('\n'))}</textarea></label></div>
      <div class="tarjeta conf-sec"><h3>Dónde se guarda</h3><p class="sub">${DEMO ? 'Modo demo: los datos quedan solo en este navegador. Para usarla en serio, ver README (lista de SharePoint).' : `Lista <b>${esc(SPC.lista)}</b> en <b>${esc(SPC.sitio)}</b>. Una fila por respuesta, con historial de versiones de SharePoint.`}</p></div>`;
  }
  function enlazarConfig(tab) {
    const cont = $('#conf');
    $$('[data-stop], .campo-conf header .interruptor', cont).forEach(x => x.addEventListener('click', e => e.stopPropagation()));
    $$('[data-ciclo]', cont).forEach(x => { x.onchange = () => { const k = x.dataset.ciclo; B.ciclo[k] = k === 'seguimiento' ? x.value.split(/[,;\s]+/).map(Number).filter(n => n > 0).sort((a, b) => a - b) : k === 'maxPares' ? Math.max(1, Number(x.value) || 1) : x.value.trim(); refrescarConfig(); }; });
    $$('[data-ciclo-b]', cont).forEach(x => { x.onchange = () => { B.ciclo[x.dataset.cicloB] = x.checked; refrescarConfig(); }; });
    $$('[data-cargo]', cont).forEach(x => { x.onchange = () => { B.cargos[x.dataset.cargo] = x.checked; refrescarConfig(); }; });
    $$('[data-todos]', cont).forEach(x => { x.onclick = () => { MAN.cargos.filter(c => c.nivel === x.dataset.todos).forEach(c => { B.cargos[c.id] = x.dataset.v === '1'; }); refrescarConfig(); }; });
    $$('[data-per]', cont).forEach(x => {
      x.onchange = () => { const p = B.personas[Number(x.dataset.per)]; p[x.dataset.k] = x.type === 'checkbox' ? x.checked : x.dataset.k === 'correo' ? x.value.trim().toLowerCase() : x.value; refrescarConfig(); };
    });
    $$('[data-del-per]', cont).forEach(x => { x.onclick = async () => { const p = B.personas[Number(x.dataset.delPer)]; if (!await confirmar('¿Quitar a ' + esc(p.nombre) + '?', 'Sus respuestas ya guardadas no se borran.', 'Quitar', true)) return; B.personas.splice(Number(x.dataset.delPer), 1); B.personas.forEach(q => { if (q.jefe === p.id) q.jefe = ''; }); refrescarConfig(); }; });
    const add = $('#add-persona', cont);
    if (add) add.onclick = () => { B.personas.push({ id: 'p' + Date.now().toString(36), nombre: 'Nueva persona', cargo: MAN.cargos[0].id, jefe: '', correo: '' }); refrescarConfig(); };
    $$('[data-abrir]', cont).forEach(h => { h.onclick = () => { campoAbierto = campoAbierto === h.dataset.abrir ? null : h.dataset.abrir; refrescarConfig(); }; });
    $$('[data-cam]', cont).forEach(x => {
      x.onchange = () => {
        const c = B.campos[Number(x.dataset.cam)];
        if (x.dataset.fuente) { const f = new Set(c.fuentes || []); x.checked ? f.add(x.dataset.fuente) : f.delete(x.dataset.fuente); c.fuentes = TODOS.filter(r => f.has(r)); }
        else if (x.dataset.preg) { c.pregunta = c.pregunta || {}; c.pregunta[x.dataset.preg] = x.value; }
        else c[x.dataset.k] = x.type === 'checkbox' ? x.checked : x.value;
        refrescarConfig();
      };
    });
    $$('[data-mover]', cont).forEach(x => { x.onclick = () => { const k = Number(x.dataset.mover), d = Number(x.dataset.d); const [c] = B.campos.splice(k, 1); B.campos.splice(k + d, 0, c); refrescarConfig(); }; });
    $$('[data-del-cam]', cont).forEach(x => { x.onclick = () => { B.campos.splice(Number(x.dataset.delCam), 1); refrescarConfig(); }; });
    const ac = $('#add-campo', cont);
    if (ac) ac.onclick = () => { const id = 'c' + Date.now().toString(36); B.campos.push({ id, etiqueta: 'Nuevo campo', tipo: 'largo', zona: 'extra', color: 'azul', activo: true, obligatorio: false, fuentes: ['auto', 'jefe'], pregunta: { auto: '', jefe: '' }, ayuda: '', ejemplo: '' }); campoAbierto = id; refrescarConfig(); };
    const rs = $('#restaurar', cont);
    if (rs) rs.onclick = async () => { if (!await confirmar('¿Restaurar los campos sugeridos?', 'Se reemplazan los campos y preguntas por los originales. Las respuestas ya guardadas no se pierden.', 'Restaurar')) return; B.campos = clon(CAMPOS_BASE); refrescarConfig(); };
    const ad = $('#admins', cont);
    if (ad) ad.onchange = () => { B.administradores = ad.value.split(/[\s,;]+/).map(x => x.trim().toLowerCase()).filter(Boolean); refrescarConfig(); };
  }

  // ---------- Datos de ejemplo (solo demo) ----------
  function semillaDemo() {
    filas.clear();
    CONF = confBase();
    const yoAntes = YO;
    const t = (d, h) => `2026-11-${String(d).padStart(2, '0')}T${String(h).padStart(2, '0')}:15:00-05:00`;
    const r = (ev, rol, evaluador, estado, campos, eventos) => {
      const reg = nuevoRegistro('respuesta', ev, rol, evaluador);
      Object.assign(reg, { estado, campos, creado: eventos[0][0] });
      reg.historial = eventos.map(([tt, accion]) => ({ t: tt, quien: persona(evaluador).nombre, accion, detalle: '' }));
      if (estado === 'enviado') reg.enviado = eventos[eventos.length - 1][0];
      reg.modificado = eventos[eventos.length - 1][0];
      filas.set(reg.id, reg);
    };
    r('ei', 'auto', 'ei', 'enviado', {
      fortaleza: 'Innovación e implementación tecnológica: este año puse en producción las apps de mantenimiento, asistencia y recepción (Orientación a resultados).',
      conducta: 'Cerrar formalmente cada proyecto con el área usuaria, con un acta de entrega antes de pasar al siguiente.',
      resultado: 'Que las áreas usen las herramientas sin depender de mí y que la información quede ordenada en los sistemas.',
      indicador: 'Número de proyectos con acta de cierre firmada por el área.',
      meta: '100% de proyectos nuevos, revisión mensual.',
      plan: ['Crear una plantilla de acta de cierre.', 'Capacitar a cada área usuaria antes de cerrar.', 'Levantar el plan de mantenimiento para el 2027.']
    }, [[t(5, 9), 'Comenzó la autoevaluación'], [t(6, 16), 'Envió la autoevaluación']]);
    r('ei', 'jefe', 'jmg', 'enviado', {
      fortaleza: 'Innovación e implementación tecnológica.',
      conducta: 'Acordar entregables con cada área y documentar su validación antes de cerrar una iniciativa.',
      resultado: 'Mejorar la estructuración de la información empresarial y los sistemas de gestión, registrados.',
      indicador: '% de iniciativas con impacto cuantificado, evidencia y cierre aprobado.',
      meta: '≥ 90% mensual.',
      plan: ['Definir plan de digitalización con hitos y responsables.', 'Cuantificar ahorro, tiempo o reducción de errores.', 'Levantar el plan de mantenimiento para el 2027.']
    }, [[t(6, 10), 'Comenzó la evaluación del jefe'], [t(7, 11), 'Envió la evaluación del jefe']]);
    const ac = nuevoRegistro('acuerdo', 'ei', 'acuerdo', '');
    Object.assign(ac, {
      estado: 'confirmado', fechaConversacion: '2026-11-10', cerrado: t(10, 12), cerradoPor: 'José Miguel González', confirmado: t(10, 17), observacion: '',
      campos: clon(filas.get(idResp('ei', 'jefe', 'jmg')).campos), creado: t(10, 11), modificado: t(10, 17),
      historial: [{ t: t(10, 11), quien: 'José Miguel González', accion: 'Abrió la sala de acuerdo', detalle: '' },
        { t: t(10, 12), quien: 'José Miguel González', accion: 'Cerró el acuerdo', detalle: 'Conversación del 10 nov 2026' },
        { t: t(10, 17), quien: 'Erick Javier Iñaguazo Pindo', accion: 'Confirmó el acuerdo', detalle: '' }]
    });
    filas.set(ac.id, ac);
    r('pg', 'auto', 'pg', 'enviado', {
      fortaleza: 'Conocimiento del sistema FSSC 22000 y liderazgo en las auditorías: cero no conformidades mayores este año.',
      conducta: 'Delegar más el seguimiento diario a los supervisores de calidad y enfocarme en la planificación.',
      resultado: 'Más tiempo para proyectos de mejora y un equipo de calidad más autónomo.',
      indicador: '% de verificaciones diarias cerradas por los supervisores sin mi intervención.',
      meta: '≥ 85% mensual.',
      plan: ['Definir matriz de responsabilidades del equipo de calidad.', 'Revisar indicadores con los supervisores cada semana.']
    }, [[t(9, 8), 'Comenzó la autoevaluación'], [t(11, 15), 'Envió la autoevaluación']]);
    r('bu', 'auto', 'bu', 'borrador', {
      fortaleza: 'Relación con los clientes del exterior y respuesta rápida a sus requerimientos.',
      conducta: ''
    }, [[t(11, 18), 'Comenzó la autoevaluación']]);
    r('jh', 'jefe', 'dn', 'borrador', {
      fortaleza: 'Compromiso con el cumplimiento del plan de producción, incluso en semanas de alta demanda.'
    }, [[t(12, 8), 'Comenzó la evaluación del jefe']]);
    cache.guardar('demo.filas', Object.fromEntries(filas));
    YO = yoAntes;
  }

  // ---------- Arranque ----------
  // Datos locales opcionales (no publicados): si no existen, se sigue sin ellos.
  const cargarScript = src => new Promise(res => { const s = document.createElement('script'); s.src = src; s.onload = s.onerror = () => res(); document.head.appendChild(s); });
  async function iniciar() {
    try {
      await Promise.all([cargarScript('datos-manual.js?v=8'), cargarScript('datos-personas.js?v=8')]);
      MAN = window.HUB_MANUAL || MAN;
      if (DEMO) {
        if (!(window.HUB_PERSONAS || []).length) {
          app.innerHTML = '<div class="tarjeta ingreso"><span class="logo">360</span><h2>La demo solo funciona en la copia local</h2><p class="sub">Abre la app sin <b>?demo</b> para ingresar con tu cuenta.</p></div>';
          return;
        }
        await Store.cargar();
        YO = persona(cache.leer('demo.yo')) || persona('dc') || CONF.personas[0];
      } else {
        if (!window.msal) throw new Error('No se pudo cargar el inicio de sesión de Microsoft. Revisa la conexión.');
        await iniciarMsal();
        if (!cuenta) {
          app.innerHTML = `<div class="tarjeta ingreso"><span class="logo">360</span><h2>Evaluación 360</h2><p class="sub">Ingresa con tu cuenta de Healthy Food (Microsoft 365).</p>
            <button class="btn" id="entrar" style="margin-top:14px">Ingresar con Microsoft</button></div>`;
          $('#entrar').onclick = () => pca.loginRedirect({ scopes: CFG.permisos, prompt: 'select_account' });
          return;
        }
        await Store.cargar();
        YO = CONF.personas.find(p => (p.correo || '').toLowerCase() === CORREO) || null;
        // Sin tiempo real: se vuelve a leer cada minuto mientras no se esté escribiendo.
        setInterval(async () => {
          if (document.hidden || porGuardar || ['r', 'acuerdo', 'config'].includes(vistaActual) || !$('#modal').hidden) return;
          try { await Store.cargar(); YO = CONF.personas.find(p => (p.correo || '').toLowerCase() === CORREO) || null; pintar(); } catch (_) {}
        }, 60000);
      }
      if (location.hash === '#demo') history.replaceState(null, '', '#/inicio');
      pintar();
    } catch (e) {
      console.error(e);
      app.innerHTML = `<div class="tarjeta ingreso"><div class="grande" style="font-size:40px">⚠️</div><h2>No se pudo abrir la evaluación</h2><p class="sub">${esc(e.message)}</p>
        <button class="btn" onclick="location.reload()">Reintentar</button></div>`;
    }
  }
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar); else iniciar();
})();
