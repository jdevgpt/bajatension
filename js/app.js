/* Interfaz de KMZ + ANEXO III Red BT */
(function () {
  'use strict';
  const APP_VERSION = '1.14.0';
  const C = window.Core;
  const $ = id => document.getElementById(id);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const estado = { fotos: [], kmzGenerado: null, kmzModificado: null, modelo: null };

  // ───────── persistencia de las tablas (solo texto) ─────────
  const LS = { main: 'bajatension.main', def: 'bajatension.def', suf: 'bajatension.suf1', comp: 'bajatension.comp', orig: 'bajatension.orig', mapa: 'bajatension.mapa', vk: 'bajatension.mapaVK', vkf: 'bajatension.mapaFotos', rev: 'bajatension.revAppsheet' };
  function lsJson(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  const lsGet = k => { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* sin almacenamiento */ } };
  $('txtMain').value = lsGet(LS.main);
  $('txtDef').value = lsGet(LS.def);
  $('chkSuf1').checked = lsGet(LS.suf) === '1';
  $('txtComp').value = lsGet(LS.comp);
  $('txtComp').addEventListener('input', () => { lsSet(LS.comp, $('txtComp').value); clearTimeout(estado.tComp); estado.tComp = setTimeout(() => pintarSecciones(), 600); });
  const infoTabla = (ta, info) => { const n = C.splitTable(ta.value).length; $(info).textContent = n ? n + ' filas' : ''; };
  [['txtMain', 'infoMain', LS.main], ['txtDef', 'infoDef', LS.def]].forEach(([ta, info, k]) => {
    $(ta).addEventListener('input', () => { lsSet(k, $(ta).value); comprobarLuego(); clearTimeout(estado.tRen); estado.tRen = setTimeout(() => { if (estado.fotos.length && ctTabla() !== estado.ctFotos) { renombrarCampo(); pintarFotos(); } }, 700); infoTabla($(ta), info); actualizarBtnExcel(); clearTimeout(estado.tNum); estado.tNum = setTimeout(() => { revisarTablas(); pintarNumeracion(); pintarSecciones(); avisarTabla(); marcarFotosSiToca(); }, 500); });
    infoTabla($(ta), info);
  });
  $('chkSuf1').addEventListener('change', () => lsSet(LS.suf, $('chkSuf1').checked ? '1' : ''));

  // ───────── carga de archivos de tabla (.txt/.csv/.xlsx) ─────────
  document.querySelectorAll('[data-file]').forEach(b => b.addEventListener('click', () => $(b.dataset.file).click()));
  async function leerTabla(file) {
    if (/\.xlsx$/i.test(file.name)) return xlsxATexto(await file.arrayBuffer());
    const buf = await file.arrayBuffer();
    let t = new TextDecoder('utf-8').decode(buf);
    if (t.includes('�')) t = new TextDecoder('windows-1252').decode(buf);
    return t;
  }
  async function xlsxATexto(buf) {
    const z = await JSZip.loadAsync(buf);
    const P = s => new DOMParser().parseFromString(s, 'application/xml');
    const wb = P(await z.file('xl/workbook.xml').async('string'));
    const rid = wb.getElementsByTagName('sheet')[0].getAttribute('r:id');
    const rels = P(await z.file('xl/_rels/workbook.xml.rels').async('string'));
    let target = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === rid).getAttribute('Target');
    target = target.replace(/^\//, '').replace(/^(?!xl\/)/, 'xl/');
    const ss = [];
    if (z.file('xl/sharedStrings.xml')) {
      const d = P(await z.file('xl/sharedStrings.xml').async('string'));
      [...d.getElementsByTagName('si')].forEach(si => ss.push([...si.getElementsByTagName('t')].map(t => t.textContent).join('')));
    }
    const sh = P(await z.file(target).async('string'));
    const colN = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
    const filas = [];
    [...sh.getElementsByTagName('row')].forEach(row => {
      const f = [];
      [...row.getElementsByTagName('c')].forEach(c => {
        const t = c.getAttribute('t'), v = c.getElementsByTagName('v')[0];
        let val = '';
        if (t === 's' && v) val = ss[+v.textContent] || '';
        else if (t === 'inlineStr') val = [...c.getElementsByTagName('t')].map(x => x.textContent).join('');
        else if (v) val = v.textContent;
        f[colN(c.getAttribute('r'))] = String(val).replace(/[\t\n\r]+/g, ' ');
      });
      filas.push(Array.from(f, x => x == null ? '' : x).join('\t'));
    });
    return filas.join('\n');
  }
  [['fMain', 'txtMain', 'infoMain', LS.main], ['fDef', 'txtDef', 'infoDef', LS.def]].forEach(([fi, ta, info, k]) => {
    $(fi).addEventListener('change', async e => {
      const f = e.target.files[0]; if (!f) return;
      try { $(ta).value = await leerTabla(f); lsSet(k, $(ta).value); infoTabla($(ta), info); $(info).textContent += ' · ' + f.name; pintarNumeracion(); actualizarBtnExcel(); avisarTabla(); }
      catch (err) { alert('No se pudo leer ' + f.name + ': ' + err.message); }
      e.target.value = '';
    });
  });

  // ───────── fotos ─────────
  function zonaDrop(zona, input, onFiles) {
    zona.addEventListener('click', () => input.click());
    zona.addEventListener('dragover', e => { e.preventDefault(); zona.classList.add('on'); });
    zona.addEventListener('dragleave', () => zona.classList.remove('on'));
    zona.addEventListener('drop', e => { e.preventDefault(); zona.classList.remove('on'); onFiles([...e.dataTransfer.files]); });
    input.addEventListener('change', e => { onFiles([...e.target.files]); e.target.value = ''; });
  }
  // Fotos: se renombran como RenombrarImagenes.cmd (con el CT de la tabla de posiciones) y se redimensionan a 1366×768.
  //   f.crudo: nombre tal cual viene del ZIP · f.orig: nombre tras RenombrarImagenes · f.nombre: + renumeración de apoyos
  //   f.leerOriginal(): bytes originales · f.leer(): bytes redimensionados (los que van al KMZ y a la carpeta FOTOS)
  const ES_IMG = /\.(jpe?g|png|gif|webp)$/i;
  function nuevaFoto(nombre, leerOriginal, origen) {
    const crudo = nombre.split(/[\\/]/).pop();
    const f = { crudo, orig: crudo, nombre: renombrar(crudo), leerOriginal, red: null, origen };
    f.leer = () => f.red || (f.red = redimensionar(f));
    return f;
  }
  async function anadirFotos(files, origen) {
    origen = origen || 'pos';
    $('infoFotos').textContent = 'Leyendo…';
    // misma carpeta: una foto con el mismo nombre que otra ya cargada la sustituye (como al copiar en Windows)
    const meter = nf => { const i = estado.fotos.findIndex(x => x.crudo.toLowerCase() === nf.crudo.toLowerCase()); if (i >= 0) estado.fotos[i] = nf; else estado.fotos.push(nf); };
    for (const f of files) {
      if (/\.zip$/i.test(f.name)) {
        try {
          const z = await JSZip.loadAsync(f);
          z.forEach((path, entry) => {
            if (entry.dir || /(^|\/)(__MACOSX|\.)/.test(path)) return;
            if (ES_IMG.test(path)) meter(nuevaFoto(path, () => entry.async('uint8array'), origen));
          });
        } catch (e) { alert('No se pudo abrir ' + f.name + ': ' + e.message); }
      } else if (ES_IMG.test(f.name)) {
        meter(nuevaFoto(f.name, async () => new Uint8Array(await f.arrayBuffer()), origen));
      }
    }
    renombrarCampo(); pintarFotos(); marcarFotosSiToca();
  }
  // CT de la tabla de posiciones (el más repetido en la columna CT)
  function ctTabla() {
    try { return String(modeloDesdeTablas().ct || ''); } catch (e) {}
    const cnt = {}; C.splitTable($('txtMain').value).forEach(r => { const v = String(r[1] || '').trim(); if (/^\d+$/.test(v)) cnt[v] = (cnt[v] || 0) + 1; });
    return Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0] || '';
  }
  function renombrarCampo() {
    if (!estado.fotos.length) return;
    // fotos de filas que la revisión de AppSheet ha quitado por ser de otro centro: no se usan
    const otros = new Set(((estado.revAplicada && estado.revAplicada.quitadas) || []).map(q => q.vk).filter(Boolean));
    if (otros.size) {
      const vkFoto = f => { const m = f.crudo.match(/^(?:PT\s*)?(\d{7,}|\d{5,}-\d+)(?=[^\d-]|$)/i); return m ? m[1] : ''; };
      estado.fotos.forEach(f => { if (otros.has(vkFoto(f))) { f.descartadaPor = 'otro centro'; f.otroCentro = vkFoto(f); estado.descartadas.push(f); } });
      estado.fotos = estado.fotos.filter(f => !f.otroCentro);
    }
    const ct = ctTabla(); estado.ctFotos = ct;
    if (!ct) { estado.fotos.forEach(f => { f.orig = f.crudo; f.nombre = renombrar(f.orig); f.campo = null; }); return; }
    const r = C.renombrarFotosCampo(estado.fotos.map(f => (f.crudoVK = C.corregirNombreFoto(f.crudo, estado.mapaVK || {}, estado.mapaFotos || {}))), ct);
    estado.fotos.forEach((f, i) => { f.campo = r[i]; f.orig = r[i].nuevo; f.nombre = renombrar(f.orig); });
  }
  // Redimensiona para que quepa en 1366×768 (una vertical queda 576×768), sin agrandar; JPEG calidad 0,9
  async function redimensionar(f) {
    const bytes = await f.leerOriginal();
    if (!/\.(jpe?g|png|webp)$/i.test(f.crudo)) return bytes;
    try {
      const bmp = await createImageBitmap(new Blob([bytes]), { imageOrientation: 'from-image' });
      const w = bmp.width, h = bmp.height;
      const k = Math.min(1, 1366 / w, 768 / h);
      if (k >= 1) { bmp.close && bmp.close(); return bytes; }
      const W = Math.round(w * k), H = Math.round(h * k);
      const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
      const g = cv.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(bmp, 0, 0, W, H); bmp.close && bmp.close();
      const tipo = /\.png$/i.test(f.crudo) ? 'image/png' : /\.webp$/i.test(f.crudo) ? 'image/webp' : 'image/jpeg';
      const blob = await new Promise(res => cv.toBlob(res, tipo, 0.9));
      return blob ? new Uint8Array(await blob.arrayBuffer()) : bytes;
    } catch (e) { console.warn('No se pudo redimensionar', f.crudo, e); return bytes; }
  }
  // ── Fotos repetidas: el técnico cambió la foto y AppSheet guardó las dos → mismo nombre final ──
  // Se elige cuál se queda; las descartadas salen de estado.fotos (no van a FOTOS, ni a FOTOS ORIGINALES, ni al KMZ).
  estado.descartadas = [];
  const horaDe = f => { const m = f.crudo.match(/\.(\d{6})\.[^.]+$/); return m ? `${m[1].slice(0, 2)}:${m[1].slice(2, 4)}:${m[1].slice(4, 6)}` : ''; };
  const urlMini = f => f.mini || (f.mini = f.leerOriginal().then(b => URL.createObjectURL(new Blob([b], { type: 'image/jpeg' }))));
  function gruposRepetidas() {
    const g = new Map();
    estado.fotos.forEach(f => { if (f.campo && f.campo.estado === 'omitido' && f.campo.destino) { const k = f.campo.destino.toLowerCase(); if (!g.has(k)) g.set(k, { destino: f.campo.destino, fotos: [] }); g.get(k).fotos.push(f); } });
    estado.descartadas.forEach(f => { const k = (f.destinoDesc || '').toLowerCase(); if (!k || f.otroCentro) return; if (!g.has(k)) g.set(k, { destino: f.destinoDesc, fotos: [] }); g.get(k).fotos.push(f); });
    g.forEach((v, k) => { const gan = estado.fotos.find(f => f.campo && f.campo.estado !== 'omitido' && (f.campo.destino || '').toLowerCase() === k); if (gan) v.fotos.unshift(gan); });
    return [...g.values()].filter(v => v.fotos.length > 1 || v.fotos.some(f => f.descartadaPor));
  }
  function pintarRepetidas() {
    const zona = $('repetidas'); if (!zona) return;
    const gs = gruposRepetidas();
    if (!gs.length) { zona.innerHTML = ''; return; }
    const pend = gs.filter(g => !g.fotos.some(f => f.descartadaPor)).length;
    zona.innerHTML = `<div class="num-box"><b>Fotos repetidas</b> (${gs.length}): el técnico cambió la foto y se guardaron las dos. Elige cuál se queda con el nombre; las otras se eliminan (no salen en FOTOS ORIGINALES, ni en el KMZ, ni en la carpeta final).${pend ? ` <b>${pend === 1 ? 'Falta 1' : 'Faltan ' + pend} por decidir.</b>` : ''}` +
      gs.map((g, gi) => `<div class="rep-g"><div class="rep-t">${esc(g.destino)}</div><div class="rep-fs">` + g.fotos.map((f, fi) => {
        const desc = !!f.descartadaPor, elegida = !desc && g.fotos.some(x => x.descartadaPor);
        return `<div class="rep-f${desc ? ' desc' : ''}${elegida ? ' ok' : ''}"><img data-g="${gi}" data-f="${fi}" alt=""><div class="info">${esc(f.crudo)}${horaDe(f) ? `<br>hora ${horaDe(f)}` : ''}${f.origen === 'def' ? ' · defectos' : ''}</div>` +
          (desc ? `<span class="t-aviso">Descartada</span> <button class="sec peq rep-rec" data-g="${gi}" data-f="${fi}">Recuperar</button>` : elegida ? '<b>✔ Se queda</b>' : `<button class="peq rep-usar" data-g="${gi}" data-f="${fi}">Usar esta</button>`) + '</div>';
      }).join('') + '</div></div>').join('') + '</div>';
    estado.gruposRep = gs;
    zona.querySelectorAll('img[data-g]').forEach(img => urlMini(gs[+img.dataset.g].fotos[+img.dataset.f]).then(u => { img.src = u; }));
  }
  function aplicarRepetidas() { renombrarCampo(); pintarFotos(); estado.fotosMarcadas = false; marcarFotosSiToca(); }
  document.addEventListener('click', e => {
    const b = e.target.closest('.rep-usar,.rep-rec'); if (!b || !estado.gruposRep) return;
    const g = estado.gruposRep[+b.dataset.g], f = g.fotos[+b.dataset.f];
    if (b.classList.contains('rep-usar')) {
      // las demás del grupo se descartan
      g.fotos.forEach(x => { if (x === f || x.descartadaPor) return; x.descartadaPor = f.crudo; x.destinoDesc = g.destino; estado.descartadas.push(x); });
      estado.fotos = estado.fotos.filter(x => !x.descartadaPor);
    } else {
      // recuperar: vuelve la descartada y se deshace la elección del grupo
      g.fotos.filter(x => x.descartadaPor).forEach(x => { delete x.descartadaPor; delete x.destinoDesc; estado.fotos.push(x); });
      estado.descartadas = estado.descartadas.filter(x => !g.fotos.includes(x));
    }
    aplicarRepetidas();
  });
  function pintarFotos() {
    if (estado.mapaListo) pintarNumeracion();
    const n = estado.fotos.length;
    $('btnQuitarFotos').hidden = !n; $('btnCarpeta').hidden = !n;
    if (!n) { $('infoFotos').textContent = 'Sin fotos.'; $('detFotos').innerHTML = ''; pintarRepetidas(); comprobarLuego(); return; }
    if (!estado.ctFotos) { $('infoFotos').textContent = `${n} fotos cargadas (${estado.fotos.filter(f => f.origen !== 'def').length} de posiciones, ${estado.fotos.filter(f => f.origen === 'def').length} de defectos). Pega la tabla de posiciones para renombrarlas con su CT.`; $('detFotos').innerHTML = ''; return; }
    const ren = estado.fotos.filter(f => f.campo && f.campo.estado === 'renombrado').length;
    const mal = estado.fotos.filter(f => f.campo && (f.campo.estado === 'ignorado' || f.campo.estado === 'omitido'));
    const nP = estado.fotos.filter(f => f.origen !== 'def').length, nD = n - nP;
    comprobarLuego(); pintarRepetidas();
    const nRep = estado.descartadas.filter(f => !f.otroCentro).length, nOtro = estado.descartadas.length - nRep;
    $('infoFotos').textContent = `${n} fotos (${nP} de posiciones${nD ? `, ${nD} de defectos` : ', sin ZIP de defectos'}${nRep ? `; ${nRep} repetida${nRep === 1 ? '' : 's'} descartada${nRep === 1 ? '' : 's'}` : ''}${nOtro ? `; ${nOtro} de otro centro, fuera` : ''}) · ${ren} renombradas con CT ${estado.ctFotos}${n - ren - mal.length ? ` · ${n - ren - mal.length} ya tenían su nombre` : ''}${mal.length ? ` · ${mal.length} sin renombrar` : ''}. Se redimensionan a 1366×768.`;
    const filas = estado.fotos.map(f => `<tr${f.campo && (f.campo.estado === 'ignorado' || f.campo.estado === 'omitido') ? ' class="t-aviso"' : ''}><td>${f.origen === 'def' ? 'Defectos' : 'Posiciones'}</td><td>${esc(f.crudo)}${f.crudoVK && f.crudoVK !== f.crudo ? `<br><span class="info">VK corregido: ${esc(f.crudoVK)}</span>` : ''}</td><td>${esc(f.nombre)}</td><td>${f.campo ? (f.campo.estado === 'renombrado' ? 'Renombrada' : f.campo.estado === 'igual' ? 'Igual' : (f.campo.estado === 'omitido' ? 'Omitida' : 'Ignorada') + ' (' + esc(f.campo.motivo) + ')') : ''}</td></tr>`).join('');
    $('detFotos').innerHTML = `<details${mal.length ? ' open' : ''}><summary>Nombres de las fotos${mal.length ? ` (${mal.length} sin renombrar)` : ''}</summary><table><tr><th>ZIP</th><th>Nombre en el ZIP</th><th>Nombre nuevo</th><th></th></tr>${filas}</table></details>`;
  }
  zonaDrop($('dropFotos'), $('fFotos'), fs => anadirFotos(fs, 'pos'));
  zonaDrop($('dropFotosDef'), $('fFotosDef'), fs => anadirFotos(fs, 'def'));
  $('btnQuitarFotos').addEventListener('click', () => { estado.fotos = []; estado.descartadas = []; estado.fotosMarcadas = false; estado.carpetaHecha = false; pintarFotos(); });


  // ───────── ANEXO II: se guarda en este equipo (IndexedDB) ─────────
  const idb = () => new Promise((res, rej) => { const r = indexedDB.open('bajatension', 1); r.onupgradeneeded = () => r.result.createObjectStore('archivos'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  async function idbGet(k) { const db = await idb(); return new Promise(res => { const q = db.transaction('archivos').objectStore('archivos').get(k); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); }); }
  async function idbSet(k, v) { const db = await idb(); return new Promise((res, rej) => { const t = db.transaction('archivos', 'readwrite'); v == null ? t.objectStore('archivos').delete(k) : t.objectStore('archivos').put(v, k); t.oncomplete = res; t.onerror = () => rej(t.error); }); }
  const ANEXO2_NOMBRE = 'ANEXO II. PROCEDIMIENTO PARA LA REVISION EXPRÉS DE LAS INSTALACIONES BT DE ENDESA. BEL011..docx';
  async function pintarAnexo2() {
    let a = null; try { a = await idbGet('anexo2'); } catch (e) {}
    $('btnQuitarAnexo2').hidden = !a;
    if (!a) {
      try { const r = await fetch('plantillas/anexo2.docx'); if (r.ok) a = { nombre: ANEXO2_NOMBRE, bytes: new Uint8Array(await r.arrayBuffer()), deLaApp: true }; } catch (e) {}
    }
    estado.anexo2 = a;
    $('infoAnexo2').textContent = !a ? 'Sin archivo del ANEXO II.' : a.deLaApp ? 'ANEXO II: el de la app (BEL011).' : `ANEXO II: ${a.nombre} (guardado en este equipo; «Quitar» vuelve al de la app).`;
  }
  zonaDrop($('dropAnexo2'), $('fAnexo2'), async files => {
    const f = files[0]; if (!f) return;
    try { await idbSet('anexo2', { nombre: f.name, bytes: new Uint8Array(await f.arrayBuffer()) }); } catch (e) { alert('No se pudo guardar el archivo: ' + e.message); }
    pintarAnexo2();
  });
  $('btnQuitarAnexo2').addEventListener('click', async () => { if (!confirm('¿Quitar el archivo del ANEXO II de este equipo?')) return; try { await idbSet('anexo2', null); } catch (e) {} pintarAnexo2(); });
  pintarAnexo2();

  // ───────── FOTOS ORIGINALES/«CT»: las fotos renombradas sin redimensionar (se guardan aparte en el ordenador) ─────────
  async function descargarOriginales(ct) {
    const z = new JSZip(), orig = z.folder('FOTOS ORIGINALES').folder(ct);
    const usados = new Set();
    for (const f of estado.fotos) {
      let n = f.nombre, i = 2; while (usados.has(n.toLowerCase())) n = f.nombre.replace(/(\.\w+)$/, `_${i++}$1`); usados.add(n.toLowerCase());
      orig.file(n, await f.leerOriginal(), { compression: 'STORE' });
    }
    descargar(await z.generateAsync({ type: 'blob' }), `FOTOS ORIGINALES - ${ct}.zip`);
    await new Promise(r => setTimeout(r, 500));
    return true;
  }
  $('btnCarpeta').addEventListener('click', async () => {
    const ct = ctTabla(); if (!ct) { alert('Pega primero la tabla de posiciones: el nombre de la carpeta es el número del centro.'); return; }
    const b = $('btnCarpeta'); b.disabled = true; const t = b.textContent; b.textContent = 'Preparando…';
    try { await descargarOriginales(ct); } catch (e) { alert(e.message); } finally { b.disabled = false; b.textContent = t; }
  });

  $('btnNuevo').addEventListener('click', () => {
    if (!confirm('¿Empezar un proyecto nuevo? Se borra todo lo del proyecto actual: tablas, correcciones, fotos, KMZ, captura del ANEXO I, ANEXO III, datos del informe/certificado y el envío pendiente al informe.')) return;
    // pantalla
    $('txtMain').value = ''; $('txtDef').value = ''; $('txtComp').value = ''; $('chkSuf1').checked = false;
    $('res1').innerHTML = ''; $('res2').innerHTML = ''; $('infoMain').textContent = ''; $('infoDef').textContent = '';
    $('infoKmz').textContent = 'Ningún KMZ cargado.'; $('btnUsarGenerado').hidden = true;
    $('fotosPick').innerHTML = '<div class="info">Carga el KMZ modificado para elegir las fotos.</div>';
    ['fMain', 'fDef', 'fFotos', 'fFotosDef', 'fKmz', 'fCaptura'].forEach(id => { if ($(id)) $(id).value = ''; });
    ['revAppsheet', 'comprobacion', 'repetidas', 'detFotos', 'res3', 'p3Defs', 'p3DirInfo'].forEach(id => { if ($(id)) $(id).innerHTML = ''; });
    // captura del ANEXO I y paso 3
    if ($('p3CapPrev').src) { try { URL.revokeObjectURL($('p3CapPrev').src); } catch (e) {} }
    $('p3CapPrev').removeAttribute('src'); $('p3CapPrev').hidden = true; $('btnQuitarCaptura').hidden = true;
    ['p3Tec', 'p3Fecha', 'p3Cd', 'p3Dir', 'p3Pob'].forEach(id => { $(id).value = ''; });
    $('p3Form').hidden = true; $('p3Vacio').hidden = false;
    // memoria de la página
    if (estado.miniaturas) estado.miniaturas.forEach(u => { if (u) URL.revokeObjectURL(u); });
    estado.miniaturas = new Map(); estado.fotoElegida = {}; estado.avisoRegenerar = false; estado.fotosMarcadas = false; estado.tablaAvisada = null;
    (estado.fotos || []).concat(estado.descartadas || []).forEach(f => { if (f.mini) f.mini.then(u => URL.revokeObjectURL(u)).catch(() => {}); });
    estado.captura = null; estado.anexo3 = null; estado.p3Defs = []; estado.gruposRep = null; estado.revDeshacer = null; estado.carpetaHecha = false; estado.ctFotos = '';
    estado.fotos = []; estado.descartadas = []; estado.mapaVK = {}; estado.mapaFotos = {}; estado.revAplicada = null; estado.revPendiente = null; estado.kmzGenerado = null; estado.kmzModificado = null; estado.kmzLeido = null; estado.modelo = null; estado.informe = null;
    clearTimeout(estado.tGrav);
    // memoria del navegador (el orden de gravedad es un ajuste y se conserva)
    [LS.main, LS.def, LS.suf, LS.comp, LS.orig, LS.mapa, LS.vk, LS.vkf, LS.rev, 'jdev_kmz_informe'].forEach(k => { try { localStorage.removeItem(k); } catch (e) {} });
    pintarFotos(); actualizarBtnExcel(); pintarNumeracion(); $('detSecc').open = false; $('secciones').innerHTML = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // ───────── leyenda (única imagen propia permitida) ─────────
  function leyendaPNG(filas) {
    const fila = 30, pad = 6;
    const cv = document.createElement('canvas'); const ctx = cv.getContext('2d');
    ctx.font = '13px Arial, Helvetica, sans-serif';
    const w = Math.max(190, Math.ceil(Math.max(...filas.map(f => ctx.measureText(f.label).width))) + 58);
    const h = filas.length * fila + pad * 2;
    cv.width = w; cv.height = h;
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h);
    ctx.strokeStyle = '#888'; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, w - 1, h - 1);
    const pin = (y, col) => { ctx.strokeStyle = '#555'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(20, y - 2); ctx.lineTo(17, y + 11); ctx.stroke(); ctx.beginPath(); ctx.arc(21, y - 5, 7, 0, 7); ctx.fillStyle = col; ctx.fill(); ctx.strokeStyle = 'rgba(0,0,0,.45)'; ctx.lineWidth = 1; ctx.stroke(); };
    filas.forEach((f, i) => {
      const y = pad + i * fila + fila / 2;
      switch (f.k) {
        case 'APOYO': pin(y, '#f7d117'); break;
        case 'APOYO_CGP': pin(y, '#1ea51e'); break;
        case 'POST': pin(y, '#7fc8f8'); break;
        case 'CGP': ctx.beginPath(); ctx.arc(20, y, 6, 0, 7); ctx.fillStyle = '#006400'; ctx.fill(); break;
        case 'GPAL': ctx.fillStyle = '#0000ff'; ctx.fillRect(14, y - 6, 12, 12); break;
        case 'DEF': ctx.strokeStyle = '#333'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(15, y + 9); ctx.lineTo(15, y - 9); ctx.stroke(); ctx.beginPath(); ctx.moveTo(15, y - 9); ctx.lineTo(27, y - 4); ctx.lineTo(15, y + 1); ctx.closePath(); ctx.fillStyle = '#ff0000'; ctx.fill(); break;
        case 'ZOBX': case 'Q91X': ctx.beginPath(); ctx.moveTo(20, y - 9); ctx.lineTo(28, y); ctx.lineTo(20, y + 9); ctx.lineTo(12, y); ctx.closePath(); ctx.fillStyle = f.k === 'ZOBX' ? '#ff0000' : '#2060ff'; ctx.fill(); break;
        case 'SAL': ctx.strokeStyle = f.color; ctx.lineWidth = 3; ctx.beginPath(); ctx.moveTo(8, y); ctx.lineTo(32, y); ctx.stroke(); break;
      }
      ctx.fillStyle = '#000'; ctx.textBaseline = 'middle'; ctx.fillText(f.label, 46, y + 1);
    });
    return new Promise(res => cv.toBlob(b => res({ blob: b, w, h }), 'image/png'));
  }

  // ───────── vista previa esquemática (SVG) ─────────
  function vistaPrevia(m) {
    const pts = m.nodes.filter(n => n.coords); if (!pts.length) return '';
    const lat0 = pts.reduce((s, n) => s + n.coords.lat, 0) / pts.length, k = Math.cos(lat0 * Math.PI / 180);
    const X = n => n.coords.lon * k, Y = n => -n.coords.lat;
    const xs = pts.map(X), ys = pts.map(Y), minx = Math.min(...xs), maxx = Math.max(...xs), miny = Math.min(...ys), maxy = Math.max(...ys);
    const W = 900, H = 420, m0 = 24, s = Math.min((W - 2 * m0) / ((maxx - minx) || 1e-6), (H - 2 * m0) / ((maxy - miny) || 1e-6));
    const px = n => (m0 + (X(n) - minx) * s).toFixed(1), py = n => (m0 + (Y(n) - miny) * s).toFixed(1);
    let g = '';
    (m.adicionales || []).forEach(x => { if (x.a.coords && x.b.coords) g += `<line x1="${px(x.a)}" y1="${py(x.a)}" x2="${px(x.b)}" y2="${py(x.b)}" stroke="${C.salHex(x.salida)}" stroke-width="2.5"><title>${C.salName(x.salida)}: ${x.a.name} → ${x.b.name} (recorrido) · ${Math.round(x.dist)} m</title></line>`; });
    (m.compartidos || []).forEach(c => { const t = c.tramo; g += `<line x1="${px(t.a)}" y1="${py(t.a)}" x2="${px(t.b)}" y2="${py(t.b)}" stroke="${C.salHex(c.salida)}" stroke-width="${2.5 + 3 * c.orden}" stroke-opacity=".55" stroke-linecap="round"><title>${C.salName(c.salida)} comparte ${t.a.name} → ${t.b.name}</title></line>`; });
    m.tramos.forEach(t => { if (t.dist != null) g += `<line x1="${px(t.a)}" y1="${py(t.a)}" x2="${px(t.b)}" y2="${py(t.b)}" stroke="${C.salHex(t.salida)}" stroke-width="2.5"><title>${t.a.name} → ${t.b.name} · ${Math.round(t.dist)} m</title></line>`; });
    pts.forEach(n => {
      const col = n.esPT ? '#e00' : n.clase === 'POST' ? '#7fc8f8' : n.clase === 'GPAL' ? '#00f' : n.cgp ? '#1ea51e' : '#f7d117';
      const def = n.defectos.length ? `<circle cx="${px(n)}" cy="${py(n)}" r="9" fill="none" stroke="#e00" stroke-dasharray="2 2"/>` : '';
      g += `${def}<circle cx="${px(n)}" cy="${py(n)}" r="${n.esPT ? 7 : 5}" fill="${col}" stroke="#333"><title>${esc(n.name)}${n.defectos.length ? ' · ' + n.defectos.map(d => d.codigo).join(', ') : ''}</title></circle>`;
      if (!n.esPT) g += `<text x="${+px(n) + 7}" y="${+py(n) - 6}" font-size="10" fill="#555">${n.apoyo}</text>`;
    });
    const hReal = Math.min(H, 2 * m0 + (maxy - miny) * s + 10);
    return `<svg class="prev" viewBox="0 0 ${W} ${hReal}" width="100%">${g}</svg><div class="info">Vista esquemática (no es un mapa). Círculo rojo discontinuo = apoyo con defecto. Pasa el ratón por encima para ver nombres y distancias.</div>`;
  }

  const listaAvisos = (arr, cls) => arr.length ? `<ul class="av">${arr.map(a => `<li class="${cls}">${esc(a)}</li>`).join('')}</ul>` : '';
  const descargar = (blob, nombre) => { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); };

  // Avisa a GestCD (misma web) de que se ha hecho un paso de este centro: 'fotos', 'kmz' o 'excel'
  function marcarPaso(paso, ct) {
    try {
      if (!ct) ct = modeloDesdeTablas().ct;
      const k = String(ct).trim().toUpperCase().replace(/^0+(?=\d)/, ''); if (!k) return false;
      const m = JSON.parse(localStorage.getItem('jdev_pasos') || '{}') || {};
      m[k] = Object.assign({}, m[k], { [paso]: Date.now() });
      localStorage.setItem('jdev_pasos', JSON.stringify(m));
      return true;
    } catch (e) { return false; }
  }
  // Al meter la tabla de posiciones se avisa a GestCD para que cree el centro si no lo tiene
  function avisarTabla() {
    try {
      const m = modeloDesdeTablas(); const k = String(m.ct).replace(/^0+(?=\d)/, '');
      if (estado.tablaAvisada === k) return;
      const reg = JSON.parse(localStorage.getItem('jdev_pasos') || '{}') || {};
      const tipo = m.pt && /^CT/i.test(String(m.pt.row.tipo).trim()) ? 'CT' : 'PT';
      reg[k] = Object.assign({}, reg[k], { tabla: Date.now(), tipo });
      localStorage.setItem('jdev_pasos', JSON.stringify(reg));
      estado.tablaAvisada = k;
    } catch (e) { /* tabla incompleta */ }
  }
  // Marca las fotos en cuanto se conozca el CT (pueden subirse antes que la tabla)
  function marcarFotosSiToca() {
    if (!estado.fotos.length || estado.fotosMarcadas) return;
    if (marcarPaso('fotos')) estado.fotosMarcadas = true;
  }
  function resaltar(id) {
    const el = $(id); el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    el.classList.add('on'); setTimeout(() => el.classList.remove('on'), 2500);
  }

  // Provincia del centro: la de GestCD (misma web), o la de la app de informes, o la última elegida
  const PROVS = ['Badajoz', 'Cáceres', 'Sevilla', 'Málaga', 'Córdoba', 'Huelva', 'Cádiz', 'Granada', 'Jaén', 'Almería'];
  const sinAcento = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
  const esBoletin = prov => sinAcento(prov) === 'MALAGA';
  function provinciaCentro(ct) {
    const k = String(ct).replace(/^0+(?=\d)/, '');
    const norm = p => PROVS.find(x => sinAcento(x) === sinAcento(p)) || p;
    try {
      const g = JSON.parse(localStorage.getItem('gestcd_v1') || 'null');
      const cs = ((g && g.centros) || []).filter(c => String(c.numero).replace(/^0+(?=\d)/, '') === k && c.provincia);
      const c = cs.find(x => !(x.pasos && x.pasos.sharepoint)) || cs[cs.length - 1];
      if (c) { const p = norm(c.provincia); if (!PROVS.includes(p)) PROVS.push(p); return { nombre: p, origen: 'GestCD', tecnico: c.operario && c.operario !== 'Sin asignar' ? c.operario : '' }; }
    } catch (e) {}
    try {
      const ib = JSON.parse(localStorage.getItem('ibt_estado') || 'null');
      for (const l of (ib && ib.lineas) || []) for (const c of l.centros || []) if (String(c.centro).replace(/^0+(?=\d)/, '') === k && c.provincia) { const p = norm(c.provincia); if (!PROVS.includes(p)) PROVS.push(p); return { nombre: p, origen: 'la app de informes' }; }
    } catch (e) {}
    return { nombre: lsGet('bajatension.prov') || 'Badajoz', origen: '' };
  }

  function modeloDesdeTablas() {
    const main = C.splitTable($('txtMain').value);
    if (!main.length) throw new Error('Falta la tabla de posiciones.');
    return C.buildModel(main, C.splitTable($('txtDef').value), { compartidos: $('txtComp').value });
  }

  // ───────── PASO 1: KMZ ─────────
  $('btnKmz').addEventListener('click', async () => {
    if (!estado.fotos.length && !confirm('No has cargado el ZIP de fotos de posiciones.\n\n¿Generar el KMZ sin fotos?\n\nAceptar: generar sin fotos · Cancelar: volver para cargarlas')) { resaltar('dropFotos'); return; }
    if (estado.fotos.length && C.splitTable($('txtDef').value).length && !estado.fotos.some(f => f.origen === 'def') && !confirm('Hay tabla de defectos pero no has cargado el ZIP de fotos de defectos.\n\n¿Seguir sin las fotos de los defectos?\n\nAceptar: seguir · Cancelar: volver para cargarlo')) { resaltar('dropFotosDef'); return; }
    const nAv = comprobar();
    if (nAv && !confirm('La comprobación ha encontrado cosas que revisar (las tienes encima del botón: apoyos sin INICIO, FIN o SECCIÓN, numeración, observaciones de apoyos compartidos o fotos).\n\n¿Generar el KMZ igualmente?\n\nAceptar: generar · Cancelar: revisarlo antes')) { resaltar('comprobacion'); return; }
    const btn = $('btnKmz'); btn.disabled = true; btn.textContent = 'Generando…';
    try {
      const m = modeloDesdeTablas();
      if (estado.fotos.length) { renombrarCampo(); pintarFotos(); btn.textContent = 'Preparando fotos…'; await descargarOriginales(m.ct); for (const f of estado.fotos) await f.leer(); btn.textContent = 'Generando KMZ…'; }
      const ph = C.asociarFotos(m, estado.fotos.map(f => f.nombre), { sufijo1EsCGP: $('chkSuf1').checked });
      // nombres únicos dentro de files/
      const usados = new Set(['leyenda.png']); const nombreDe = new Map();
      const fotoPor = new Map(estado.fotos.map(f => [f.nombre, f]));
      const todas = [];
      m.nodes.forEach(n => todas.push(...n.fotos, ...n.fotosCGP)); m.defectos.forEach(d => todas.push(...d.fotos));
      todas.forEach(f => { if (nombreDe.has(f.full)) return; let b = f.base, i = 2; while (usados.has(b)) b = f.base.replace(/(\.\w+)$/, `_${i++}$1`); usados.add(b); nombreDe.set(f.full, b); });
      const filasLey = C.leyendaFilas(m);
      const ley = await leyendaPNG(filasLey);
      const kml = C.buildKml(m, f => nombreDe.get(f.full), { w: ley.w, h: ley.h });
      const z = new JSZip();
      z.file('doc.kml', kml, { compression: 'DEFLATE' });
      z.file('files/leyenda.png', ley.blob, { compression: 'STORE' });
      for (const [full, base] of nombreDe) z.file('files/' + base, await fotoPor.get(full).leer(), { compression: 'STORE' });
      const blob = await z.generateAsync({ type: 'blob', mimeType: 'application/vnd.google-earth.kmz' });
      const nombre = `ANEXO I - ${m.ct}.kmz`;
      estado.kmzGenerado = { blob, nombre }; estado.modelo = m; if (estado.avisoRegenerar) { estado.avisoRegenerar = false; pintarNumeracion(); }
      descargar(blob, nombre);
      marcarPaso('kmz', m.ct); marcarFotosSiToca();
      $('btnUsarGenerado').hidden = false; actualizarBtnExcel();
      pintarRes1(m, ph, nombre);
    } catch (e) {
      console.error(e); $('res1').innerHTML = `<div class="res"><div class="err-box">${esc(e.message)}</div></div>`;
    } finally { btn.disabled = false; btn.textContent = 'Generar KMZ'; }
  });

  function pintarRes1(m, ph, nombre) {
    const nApoyos = m.nodes.filter(n => !n.esPT).length, nCgp = m.nodes.filter(n => n.cgp).length;
    const tramos = m.tramos.filter(t => t.dist != null).length;
    const d = C.distancias(m);
    const nZ = m.defectos.filter(x => x.codigo === 'ZOBX').length, nQ = m.defectos.filter(x => x.codigo === 'ZOBX_Q91X').length, nD = m.defectos.length - nZ - nQ;
    let h = `<div class="res"><div class="ok-box">✔ <b>${esc(nombre)}</b> generado y descargado.</div>
      <div class="chips"><span class="chip">PT/CT <b>${esc(m.pt ? m.pt.name : m.ct)}</b></span><span class="chip">Apoyos <b>${nApoyos}</b></span><span class="chip">Tramos <b>${tramos}</b></span><span class="chip">CGP/CPM <b>${nCgp}</b></span>
      <span class="chip">ZOBX <b>${nZ}</b></span>${nQ ? `<span class="chip">ZOBX_Q91X <b>${nQ}</b></span>` : ''}<span class="chip">Otros defectos <b>${nD}</b></span>
      <span class="chip">Fotos <b>${ph.asociadas}/${ph.total}</b></span><span class="chip">Red total <b>${C.fmtMiles(d.total)} m</b></span></div>`;
    h += m.avisos.length ? `<b class="t-aviso">Avisos (${m.avisos.length})</b>${listaAvisos(m.avisos, 't-aviso')}` : '<div class="t-ok">Sin avisos.</div>';
    if (m.notas.length) h += `<details><summary>Notas (${m.notas.length})</summary>${listaAvisos(m.notas, '')}</details>`;
    h += `<details><summary>Distancias por salida</summary><table><tr><th>Salida</th><th>Metros (suma de tramos)</th></tr>${Object.entries(d.porSalida).map(([s, v]) => `<tr><td>${s}</td><td>${C.fmtMiles(v)}</td></tr>`).join('')}<tr><th>Total red${(m.compartidos || []).length ? ' (los tramos compartidos cuentan una vez por cada línea)' : ''}</th><th>${C.fmtMiles(d.total)}</th></tr></table></details>`;
    if (m.defectos.length) h += `<details><summary>Defectos asignados (${m.defectos.length})</summary><div class="scroll"><table><tr><th>Código</th><th>Descripción</th><th>Apoyo</th><th>Cómo se asoció</th><th>Distancia</th><th>Fotos</th></tr>${m.defectos.map(x => `<tr><td>${esc(x.codigo)}</td><td>${esc(x.desc)}</td><td>${esc(x.node.esPT ? x.node.name : x.node.code)}</td><td>${esc(x.metodo)}</td><td class="${x.lejos ? 't-aviso' : ''}">${x.dist == null ? '' : C.fmtMiles(x.dist) + ' m'}</td><td>${x.fotos.length}</td></tr>`).join('')}</table></div></details>`;
    if (ph.huerfanas.length) h += `<details open><summary class="t-aviso">Fotos sin asociar (${ph.huerfanas.length})</summary>${listaAvisos(ph.huerfanas.map(x => x.nombre + ' — ' + x.motivo), 't-aviso')}</details>`;
    h += vistaPrevia(m) + '</div>';
    $('res1').innerHTML = h;
  }

  // ───────── PASO 2: ANEXO III ─────────
  function actualizarBtnExcel() { $('btnExcel').disabled = !$('txtMain').value.trim(); }
  zonaDrop($('dropKmz'), $('fKmz'), files => {
    const f = files.find(x => /\.(kmz|kml)$/i.test(x.name)); if (!f) { alert('Elige un archivo .kmz o .kml'); return; }
    estado.kmzModificado = { blob: f, nombre: f.name };
    $('infoKmz').textContent = 'KMZ cargado: ' + f.name; actualizarBtnExcel(); prepararFotos(true);
  });
  $('btnUsarGenerado').addEventListener('click', () => {
    estado.kmzModificado = estado.kmzGenerado;
    $('infoKmz').textContent = 'Usando el KMZ recién generado (sin modificar): ' + estado.kmzGenerado.nombre; actualizarBtnExcel(); prepararFotos(true);
  });

  // ───────── elección de la foto de cada fila (columna FOTO) ─────────
  estado.fotoElegida = {}; estado.miniaturas = new Map();
  async function prepararFotos(nuevoKmz) {
    const zona = $('fotosPick');
    if (nuevoKmz) { estado.fotoElegida = {}; estado.miniaturas.forEach(u => URL.revokeObjectURL(u)); estado.miniaturas = new Map(); }
    const kmz = estado.kmzModificado || estado.kmzGenerado;
    if (!kmz) { zona.innerHTML = ''; return; }
    try {
      const m = modeloDesdeTablas();
      const k = await leerKmz(kmz); estado.kmzLeido = k;
      const an = C.buildAnexo(m, C.parseKml(k.kml), { gravedad: $('txtGrav').value.split(/[\s,;]+/), mapaCodigos: Object.assign({}, estado.mapaVK, mapaTotal()) });
      const filas = an.filas.filter(f => f.candidatas.length);
      if (!filas.length) { zona.innerHTML = '<div class="info">No hay fotos que elegir (ni foto _1 del PT ni fotos de defectos en el KMZ).</div>'; return; }
      const leer = lectorImagenes(k.zip);
      for (const f of filas) for (const c of f.candidatas) if (!estado.miniaturas.has(c.src)) { const d = await leer(c.src); estado.miniaturas.set(c.src, d ? URL.createObjectURL(new Blob([d], { type: d[0] === 0x89 ? 'image/png' : 'image/jpeg' })) : ''); }
      zona.innerHTML = `<p class="info" style="margin:4px 0 8px">Pulsa la foto que quieres en cada fila. Viene marcada la que elegiría la app (foto _1 en el PT; 1ª foto del defecto más grave en los apoyos).</p>` + filas.map(f => {
        const sel = Object.prototype.hasOwnProperty.call(estado.fotoElegida, f.clave) ? estado.fotoElegida[f.clave] : (f.fotoSrc || '');
        const titulo = f.apoyo === '0' || +f.apoyo === 0 ? `Apoyo 0 · PT/CT ${esc(f.ct)}` : `Apoyo ${esc(f.apoyo)} · ${esc([f.defec, f.q91x].filter(Boolean).join(' / ') || 'defecto')}`;
        return `<div class="fp-fila" data-k="${esc(f.clave)}"><div class="fp-tit">${titulo}</div><div class="fp-fotos">` +
          f.candidatas.map(c => `<button type="button" class="fp-foto${sel === c.src ? ' on' : ''}" data-src="${esc(c.src)}" title="${esc(c.src.split('/').pop())}"><img loading="lazy" src="${estado.miniaturas.get(c.src) || ''}" alt=""><span>${esc((suf => c.etiqueta === 'General' ? ('General ' + suf).trim() : (suf || c.etiqueta))(c.src.split('/').pop().replace(/\.\w+$/, '').replace(/^\d+_?/, '')))}</span></button>`).join('') +
          `<button type="button" class="fp-foto fp-no${sel === '' ? ' on' : ''}" data-src=""><span>Sin foto</span></button></div></div>`;
      }).join('');
    } catch (e) { zona.innerHTML = `<div class="err-box">${esc(e.message)}</div>`; }
  }
  $('fotosPick').addEventListener('click', e => {
    const b = e.target.closest('.fp-foto'); if (!b) return;
    const fila = b.closest('.fp-fila');
    estado.fotoElegida[fila.dataset.k] = b.dataset.src;
    fila.querySelectorAll('.fp-foto').forEach(x => x.classList.toggle('on', x === b));
  });
  $('fotosPick').addEventListener('dblclick', e => { const b = e.target.closest('.fp-foto'); if (b && b.dataset.src) window.open(estado.miniaturas.get(b.dataset.src), '_blank'); });


  // ── Comprobador de metros (+ ANEXO III solo desde el KMZ, con los metros reales de cada línea) ──
  estado.metros = [];
  function fechaGestcd(ct) {
    try {
      const k = String(ct).replace(/^0+(?=\d)/, '');
      const g = JSON.parse(localStorage.getItem('gestcd_v1') || 'null');
      const cs = ((g && g.centros) || []).filter(c => String(c.numero).replace(/^0+(?=\d)/, '') === k && c.fecha);
      return cs.length ? cs[cs.length - 1].fecha : '';
    } catch (e) { return ''; }
  }
  const isoADmy = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? `${m[3]}/${m[2]}/${m[1]}` : ''; };
  zonaDrop($('dropMetros'), $('fMetros'), async files => {
    files = files.filter(x => /\.(kmz|kml)$/i.test(x.name));
    if (!files.length) { alert('Elige archivos .kmz o .kml'); return; }
    const out = $('resMetros'); out.innerHTML = '<div class="info">Midiendo…</div>';
    estado.metros = [];
    let h = '', suma = 0;
    const m = v => C.fmtMiles(Math.round(v));
    for (const f of files) {
      try {
        const k = await leerKmz({ nombre: f.name, blob: f });
        const r = C.medirKml(k.kml); suma += r.total;
        let ct = ''; try { ct = C.tablaDesdeKmz(C.parseKml(k.kml), '').ct; } catch (e) {}
        const i = estado.metros.push({ nombre: f.name, k, ct }) - 1;
        h += `<div class="res"><div class="ok-box"><b>${esc(f.name)}</b>: <b>${m(r.total)} m</b> de red en ${r.lineas} línea${r.lineas === 1 ? '' : 's'}${r.compartidas ? ` (${r.compartidas} de tramos compartidos)` : ''}</div>`;
        if (!r.lineas) h += '<div class="t-aviso">Este KMZ no tiene líneas.</div>';
        else h += `<table><tr><th>Carpeta / salida</th><th>Líneas</th><th>Metros</th></tr>${r.grupos.map(g => `<tr><td>${esc(g.carpeta)}</td><td>${g.lineas}${g.compartidas ? ` (${g.compartidas} compartidas)` : ''}</td><td>${m(g.metros)}</td></tr>`).join('')}<tr><th>Total</th><th>${r.lineas}</th><th>${m(r.total)}</th></tr></table>`;
        if (ct) h += `<div class="fila m-anexo" data-i="${i}"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" class="m-sel"> ANEXO III de ${esc(ct)}</label>
          <label style="display:flex;gap:6px;align-items:center">Fecha <input type="date" class="m-fecha" value="${esc(fechaGestcd(ct))}"></label></div>`;
        h += '</div>';
      } catch (e) { h += `<div class="res t-aviso">${esc(f.name)}: no lo puedo leer (${esc(e.message)})</div>`; }
    }
    if (files.length > 1) h += `<div class="ok-box">Suma de los ${files.length} KMZ: <b>${C.fmtMiles(Math.round(suma))} m</b></div>`;
    if (estado.metros.some(x => x.ct)) h += `<div class="acciones"><button id="btnAnexoMetros">Generar ANEXO III de los marcados</button></div>
      <p class="info">El ANEXO III se monta solo con el KMZ: apoyos, secciones, observaciones y defectos de sus popups, INICIO/FIN de las líneas y DIST. con los metros reales de cada línea dibujada. En los tramos por los que van dos o más salidas (circuito doble) la DIST. lleva los metros de todas sus líneas, así que el total cuadra con el KMZ. Si una CGP/CPM está separada de su apoyo (cable grapado por la pared), su distancia se suma a la DIST. del apoyo: los metros de la línea si la dibujas del apoyo a la CGP, o en línea recta si no. La foto de cada fila es la que elegiría la app (foto _1 del PT y 1ª foto del defecto más grave).</p><div id="resAnexoMetros"></div>`;
    out.innerHTML = h; $('btnLimpiarMetros').hidden = false;
  });
  $('btnLimpiarMetros').addEventListener('click', () => { estado.metros = []; $('resMetros').innerHTML = ''; $('fMetros').value = ''; $('btnLimpiarMetros').hidden = true; });
  $('resMetros').addEventListener('click', async e => {
    if (e.target.id !== 'btnAnexoMetros') return;
    const sel = [...document.querySelectorAll('.m-anexo')].filter(d => d.querySelector('.m-sel').checked);
    if (!sel.length) { alert('Marca los KMZ de los que quieres el ANEXO III.'); return; }
    const sinFecha = sel.filter(d => !d.querySelector('.m-fecha').value);
    if (sinFecha.length) { alert('Pon la fecha de la revisión de ' + sinFecha.map(d => estado.metros[+d.dataset.i].ct).join(', ') + ' (no va dentro del KMZ).'); sinFecha[0].querySelector('.m-fecha').focus(); return; }
    const btn = e.target; btn.disabled = true; btn.textContent = 'Generando…';
    let h = '';
    for (const d of sel) {
      const it = estado.metros[+d.dataset.i];
      try {
        const kd = C.parseKml(it.k.kml);
        const t = C.tablaDesdeKmz(kd, isoADmy(d.querySelector('.m-fecha').value));
        const mod = C.buildModel(t.rows, [], {});
        const an = C.buildAnexo(mod, kd, { gravedad: $('txtGrav').value.split(/[\s,;]+/), metrosLinea: true, circuitos: t.circuitos });
        an.extraTotal = t.compartidos;
        const tpl = await JSZip.loadAsync(Uint8Array.from(atob(window.ANEXO_TEMPLATE_B64), c => c.charCodeAt(0)));
        const gx = await C.generarAnexoXlsx(tpl, an, mod.ct, lectorImagenes(it.k.zip));
        const blob = await tpl.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const nombre = `ANEXO III - ${mod.ct}.xlsx`;
        descargar(blob, nombre); await new Promise(r => setTimeout(r, 400));
        const avisos = t.avisos.concat(mod.avisos, an.avisos, gx.avisos);
        const tot = an.totalDist + Math.round(t.compartidos);
        const dobles = Object.keys(t.circuitos).length, medido = C.medirKml(it.k.kml).total;
        h += `<div class="ok-box">✔ <b>${esc(nombre)}</b>: ${an.filas.length} filas · total <b>${C.fmtMiles(tot)} m</b> (KMZ: ${C.fmtMiles(Math.round(medido))} m de líneas${an.cgpRecta >= 0.5 ? ` + ${C.fmtMiles(Math.round(an.cgpRecta))} m de apoyos a CGP/CPM sin línea dibujada` : ''})${dobles ? ` · ${dobles} tramo${dobles === 1 ? '' : 's'} con circuito doble o más, con sus metros multiplicados` : ''}</div>` +
          (avisos.length ? listaAvisos(avisos, 't-aviso') : '');
      } catch (er) { h += `<div class="err-box">${esc(it.nombre)}: ${esc(er.message)}</div>`; }
    }
    $('resAnexoMetros').innerHTML = h;
    btn.disabled = false; btn.textContent = 'Generar ANEXO III de los marcados';
  });

  async function leerKmz(k) {
    if (/\.kml$/i.test(k.nombre)) return { kml: await k.blob.text(), zip: null };
    const z = await JSZip.loadAsync(k.blob);
    const f = z.file('doc.kml') || z.file(/\.kml$/i)[0];
    if (!f) throw new Error('El KMZ no contiene ningún .kml');
    return { kml: await f.async('string'), zip: z };
  }
  function lectorImagenes(zip) {
    return async src => {
      if (!zip) return null;
      let p = src; try { p = decodeURIComponent(src); } catch (e) {}
      p = p.replace(/^\.?\//, '');
      let f = zip.file(p);
      if (!f) { const base = p.split(/[\\/]/).pop().toLowerCase(); f = zip.file(new RegExp('(^|/)' + base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$', 'i'))[0]; }
      return f ? await f.async('uint8array') : null;
    };
  }

  // Orden de gravedad de defectos (para elegir la foto)
  const LS_GRAV = 'bajatension.gravedad';
  $('txtGrav').value = lsGet(LS_GRAV) || C.GRAVEDAD_DEFECTO.join(', ');
  $('txtGrav').addEventListener('input', () => { lsSet(LS_GRAV, $('txtGrav').value); clearTimeout(estado.tGrav); estado.tGrav = setTimeout(() => prepararFotos(false), 600); });
  $('btnGravDef').addEventListener('click', () => { $('txtGrav').value = C.GRAVEDAD_DEFECTO.join(', '); lsSet(LS_GRAV, ''); });

  $('btnExcel').addEventListener('click', async () => {
    let kmz = estado.kmzModificado;
    if (!kmz) {
      if (!estado.kmzGenerado) { alert('No has subido el KMZ (ANEXO I).\n\nSúbelo para generar el Excel.'); resaltar('dropKmz'); return; }
      if (!confirm('No has subido el KMZ modificado (ANEXO I).\n\n¿Generar el Excel con el KMZ recién generado, sin los cambios que hayas hecho en Google Earth?\n\nAceptar: generar con el recién generado · Cancelar: volver para subir el modificado')) { resaltar('dropKmz'); return; }
      kmz = estado.kmzGenerado;
    }
    const btn = $('btnExcel'); btn.disabled = true; btn.textContent = 'Generando…';
    try {
      const m = modeloDesdeTablas(); // siempre desde las tablas guardadas
      const k = await leerKmz(kmz);
      const an = C.buildAnexo(m, C.parseKml(k.kml), { gravedad: $('txtGrav').value.split(/[\s,;]+/), fotoElegida: estado.fotoElegida, mapaCodigos: Object.assign({}, estado.mapaVK, mapaTotal()) });
      const tpl = await JSZip.loadAsync(Uint8Array.from(atob(window.ANEXO_TEMPLATE_B64), c => c.charCodeAt(0)));
      const gx = await C.generarAnexoXlsx(tpl, an, m.ct, lectorImagenes(k.zip));
      an.avisos.push(...gx.avisos);
      if (!k.zip) an.avisos.push('Has subido un .kml: no lleva fotos, así que la columna FOTO queda vacía.');
      const blob = await tpl.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const nombre = `ANEXO III - ${m.ct}.xlsx`;
      marcarPaso('excel', m.ct);
      estado.anexo3 = { blob, nombre, kmz, kzip: k.zip, ct: m.ct };
      const conDef = an.filas.filter(f => f.defec || f.q91x);
      let h = `<div class="res"><div class="ok-box">✔ <b>${esc(nombre)}</b> generado desde <i>${esc(kmz.nombre)}</i>. Se descarga en la carpeta final (paso 3). <button class="sec peq" id="btnSoloAnexo3">Descargar solo el ANEXO III</button></div>
        <div class="chips"><span class="chip">Filas <b>${an.filas.length}</b></span><span class="chip">Con defecto <b>${conDef.length}</b></span><span class="chip">Suma DIST. <b>${C.fmtMiles(an.totalDist)} m</b></span><span class="chip">Fotos en FOTO <b>${gx.nFotos}</b></span><span class="chip">Defectos leídos de <b>${an.fuenteDef}</b></span></div>`;
      const inf = an.informe;
      const [fy, fm, fd] = (inf.fecha || '--').split('-');
      estado.informe = inf; prepararPaso3();
      h += an.avisos.length ? `<b class="t-aviso">Avisos (${an.avisos.length})</b>${listaAvisos(an.avisos, 't-aviso')}` : '<div class="t-ok">Topología del KMZ coincide con la tabla. Sin avisos.</div>';
      if (an.resumen.length) h += `<details open><summary>Supuestos aplicados</summary>${listaAvisos(an.resumen, '')}</details>`;
      h += `<details><summary>Ver filas del ANEXO III</summary><div class="scroll"><table><tr><th>Nº</th><th>TIPO</th><th>INICIO</th><th>FIN</th><th>SECCIÓN</th><th>DIST.</th><th>DEFEC.</th><th>Q91X</th><th>OBSERVACIONES</th><th>FOTO</th></tr>${an.filas.map(f => `<tr><td>${esc(f.apoyo)}</td><td>${esc(f.tipo)}</td><td>${esc(f.inicio)}</td><td>${esc(f.fin)}</td><td>${esc(f.seccion)}</td><td>${f.dist}</td><td>${esc(f.defec)}</td><td>${esc(f.q91x)}</td><td>${esc(f.obs)}</td><td>${f.fotoVm ? esc(f.fotoSrc.split('/').pop()) + '<br><span class="info">' + esc(f.fotoMotivo) + '</span>' : ''}</td></tr>`).join('')}</table></div></details></div>`;
      $('res2').innerHTML = h;
      $('btnSoloAnexo3').addEventListener('click', () => descargar(blob, nombre));
    } catch (e) {
      console.error(e); $('res2').innerHTML = `<div class="res"><div class="err-box">${esc(e.message)}</div></div>`;
    } finally { btn.textContent = 'Generar ANEXO III'; actualizarBtnExcel(); }
  });






  // ───────── Revisión de las tablas de AppSheet (CT / VK / fotos cruzados) ─────────
  estado.mapaVK = lsJson(LS.vk, {}); estado.mapaFotos = lsJson(LS.vkf, {}); estado.revAplicada = lsJson(LS.rev, null);
  const letraCol = k => { let s = ''; k++; while (k > 0) { const r = (k - 1) % 26; s = String.fromCharCode(65 + r) + s; k = Math.floor((k - 1) / 26); } return s; };
  function revisarTablas() {
    const zona = $('revAppsheet'); if (!zona) return;
    let r = null;
    try { r = C.revisarAppsheet($('txtMain').value, $('txtDef').value); } catch (e) { console.error(e); }
    estado.revPendiente = r && (r.cambios.length || r.quitadas.length) ? r : null;
    const ap = estado.revAplicada;
    let h = '';
    if (estado.revPendiente) {
      const nP = r.cambios.filter(c => c.tabla === 'POSICIONES').length, nD = r.cambios.length - nP;
      h += `<div class="num-box"><b>⚠ Datos cruzados de AppSheet</b> en el centro ${esc(r.ct)}: ${r.cambios.length} celda${r.cambios.length === 1 ? '' : 's'} que corregir (${nP} en posiciones, ${nD} en defectos)${r.quitadas.length ? ` y ${r.quitadas.length} fila${r.quitadas.length === 1 ? '' : 's'} de otro centro` : ''}.
        <div class="info">Todavía no se ha cambiado nada: revisa la lista, corrige cualquier valor propuesto que no sea correcto (se pueden editar) y pulsa «Aplicar correcciones». El VK y el nombre de las fotos mandan sobre la columna CT; el VK se rehace con CT + cuadro + salida + apoyo; los negativos pasan a positivos; las rutas de las fotos que no coinciden con el VK, los defectos y las fotos del ZIP pasan al VK de su fila.</div>
        ${tablaCambios(r, true)}
        <div class="acciones" style="margin-top:8px"><button class="peq" id="btnAplicarRev">Aplicar correcciones</button></div></div>`;
    } else if (ap) {
      h += `<div class="num-box ok">✔ Correcciones de AppSheet aplicadas: ${ap.cambios.length} celda${ap.cambios.length === 1 ? '' : 's'}${ap.quitadas.length ? `, ${ap.quitadas.length} fila${ap.quitadas.length === 1 ? '' : 's'} de otro centro quitada${ap.quitadas.length === 1 ? '' : 's'}` : ''}${Object.keys(estado.mapaVK).length ? ` y ${Object.keys(estado.mapaVK).length} VK cambiado${Object.keys(estado.mapaVK).length === 1 ? '' : 's'} (las fotos del ZIP se renombran con el VK nuevo)` : ''}.
        <div class="acciones" style="margin-top:6px"><button class="sec peq" id="btnListadoRev">Descargar listado de cambios</button><button class="sec peq" id="btnCopiarMain">Copiar posiciones corregidas</button>${$('txtDef').value.trim() ? '<button class="sec peq" id="btnCopiarDef">Copiar defectos corregidos</button>' : ''}<button class="sec peq" id="btnDeshacerRev">Deshacer</button></div>
        <details><summary>Ver los cambios</summary>${tablaCambios(ap)}</details></div>`;
    }
    if (r && r.revisar.length) h += `<div style="margin-top:6px">${listaAvisos(r.revisar, 't-aviso')}</div>`;
    zona.innerHTML = h;
  }
  function tablaCambios(r, editable) {
    const filas = r.cambios.map((c, i) => `<tr${c.conflicto ? ' class="t-aviso" style="background:#fdecea"' : ''}><td>${c.tabla}</td><td>${letraCol(c.col)}${c.fila}</td><td>${esc(c.antes)}</td><td>${editable
        ? `<input class="rv-in campo" style="margin:0;padding:4px 6px;font-size:12px" data-i="${i}"${c.dep ? ` data-dep="${esc(c.dep)}"` : ''}${c.esVK ? ` data-vk="${esc(c.dep)}" data-prev="${esc(c.despues)}"` : ''} value="${esc(c.despues)}">`
        : esc(c.despues)}</td><td>${esc(c.motivo)}</td></tr>`).join('')
      + r.quitadas.map((q, i) => `<tr class="t-aviso"><td>${q.tabla || 'POSICIONES'}</td><td>fila ${q.fila}</td><td>${esc(q.vk)}</td><td>${editable ? `<label><input type="checkbox" class="rv-q" data-i="${i}" checked> quitar la fila</label>` : '(fila quitada)'}</td><td>${esc(q.motivo)}</td></tr>`).join('');
    return `<details${editable || r.cambios.length + r.quitadas.length <= 12 ? ' open' : ''}><summary>Cambios (${r.cambios.length + r.quitadas.length})</summary><div class="scroll" style="max-height:320px"><table class="revt"><colgroup><col style="width:92px"><col style="width:58px"><col><col><col style="width:22%"></colgroup><tr><th>Tabla</th><th>Celda</th><th>Antes</th><th>Después</th><th>Motivo</th></tr>${filas}</table></div></details>`;
  }
  function aplicarRevision() {
    const r = estado.revPendiente; if (!r) return;
    // valores elegidos (los propuestos o los que se hayan editado)
    const val = i => { const el = $('revAppsheet').querySelector(`.rv-in[data-i="${i}"]`); return el ? el.value.trim() : r.cambios[i].despues; };
    const cambios = r.cambios.map((c, i) => Object.assign({}, c, { despues: val(i) }));
    const quitadas = r.quitadas.filter((q, i) => { const el = $('revAppsheet').querySelector(`.rv-q[data-i="${i}"]`); return !el || el.checked; });
    const vacios = cambios.filter(c => !c.despues && c.antes);
    if (vacios.length && !confirm(`Has dejado vacío${vacios.length === 1 ? '' : 's'} ${vacios.length} valor${vacios.length === 1 ? '' : 'es'} (${vacios.map(c => letraCol(c.col) + c.fila).join(', ')}). ¿Aplicar igualmente?`)) return;
    const editar = (texto, tabla) => {
      const sep = /\t/.test(texto) ? '\t' : ';';
      const lin = String(texto || '').replace(/\r\n?/g, '\n').split('\n');
      cambios.filter(c => c.tabla === tabla).forEach(c => { const cs = (lin[c.fila - 1] || '').split(sep); while (cs.length <= c.col) cs.push(''); cs[c.col] = c.despues; lin[c.fila - 1] = cs.join(sep); });
      const fuera = new Set(quitadas.filter(q => (q.tabla || 'POSICIONES') === tabla).map(q => q.fila - 1));
      return lin.filter((l, i) => !fuera.has(i)).join('\n');
    };
    const mainNuevo = editar($('txtMain').value, 'POSICIONES'), defNuevo = $('txtDef').value.trim() ? editar($('txtDef').value, 'DEFECTOS') : $('txtDef').value;
    estado.revDeshacer = { main: $('txtMain').value, def: $('txtDef').value, mapa: Object.assign({}, estado.mapaVK), mapaF: Object.assign({}, estado.mapaFotos), ap: estado.revAplicada };
    // fotos del ZIP: VK cambiados y fotos concretas de las rutas corregidas
    const base = v => String(v || '').split('/').pop().trim();
    const nuevoVK = {}, nuevasF = {};
    cambios.forEach(c => {
      if (c.esVK && c.antes && c.despues && c.antes !== c.despues) nuevoVK[c.antes] = c.despues;
      if (c.ruta && base(c.antes) && base(c.despues) && base(c.antes) !== base(c.despues)) nuevasF[base(c.antes).toLowerCase()] = base(c.despues);
    });
    const mapa = Object.assign({}, estado.mapaVK);
    Object.keys(mapa).forEach(k => { if (nuevoVK[mapa[k]]) mapa[k] = nuevoVK[mapa[k]]; });
    Object.keys(nuevoVK).forEach(k => { if (!(k in mapa)) mapa[k] = nuevoVK[k]; });
    estado.mapaVK = mapa; lsSet(LS.vk, JSON.stringify(mapa));
    estado.mapaFotos = Object.assign({}, estado.mapaFotos, nuevasF); lsSet(LS.vkf, JSON.stringify(estado.mapaFotos));
    const prev = estado.revAplicada || { cambios: [], quitadas: [] };
    estado.revAplicada = { ct: r.ct, cambios: prev.cambios.concat(cambios), quitadas: prev.quitadas.concat(quitadas) };
    lsSet(LS.rev, JSON.stringify(estado.revAplicada));
    $('txtMain').value = mainNuevo; lsSet(LS.main, mainNuevo);
    $('txtDef').value = defNuevo; lsSet(LS.def, defNuevo);
    ['txtMain', 'txtDef'].forEach(id => $(id).dispatchEvent(new Event('input')));
    if (estado.fotos.length) { renombrarCampo(); pintarFotos(); }
  }
  $('revAppsheet').addEventListener('input', e => {
    const inp = e.target; if (!inp.classList.contains('rv-in')) return;
    inp.dataset.editado = '1';
    if (inp.dataset.vk == null) return;
    const prev = inp.dataset.prev, nuevo = inp.value.trim(); if (!prev || !nuevo || prev === nuevo) return;
    $('revAppsheet').querySelectorAll(`.rv-in[data-dep="${CSS.escape(inp.dataset.vk)}"]`).forEach(o => { if (o !== inp && !o.dataset.editado) o.value = o.value.split(prev).join(nuevo); });
    inp.dataset.prev = nuevo;
  });
  $('revAppsheet').addEventListener('click', async e => {
    if (e.target.id === 'btnAplicarRev') aplicarRevision();
    else if (e.target.id === 'btnDeshacerRev') {
      const d = estado.revDeshacer;
      if (!d) { if (!confirm('Se borra el registro de correcciones y el cambio de nombres de las fotos. Las tablas se quedan como están (pega de nuevo las originales si quieres volver atrás). ¿Seguir?')) return; estado.mapaVK = {}; estado.mapaFotos = {}; estado.revAplicada = null; }
      else { estado.mapaVK = d.mapa; estado.mapaFotos = d.mapaF || {}; estado.revAplicada = d.ap; $('txtMain').value = d.main; $('txtDef').value = d.def; lsSet(LS.main, d.main); lsSet(LS.def, d.def); estado.revDeshacer = null; }
      lsSet(LS.vk, JSON.stringify(estado.mapaVK)); lsSet(LS.vkf, JSON.stringify(estado.mapaFotos)); lsSet(LS.rev, JSON.stringify(estado.revAplicada));
      ['txtMain', 'txtDef'].forEach(id => $(id).dispatchEvent(new Event('input')));
      if (estado.fotos.length) { renombrarCampo(); pintarFotos(); }
    }
    else if (e.target.id === 'btnCopiarMain' || e.target.id === 'btnCopiarDef') {
      const t = $(e.target.id === 'btnCopiarMain' ? 'txtMain' : 'txtDef').value;
      try { await navigator.clipboard.writeText(t); e.target.textContent = '✔ Copiada'; } catch (er) { alert('No se pudo copiar: selecciona el texto de la tabla y cópialo a mano.'); }
    }
    else if (e.target.id === 'btnListadoRev') {
      const ap = estado.revAplicada; if (!ap) return;
      const q = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
      const filas = [['Tabla', 'Fila', 'Celda', 'Antes', 'Después', 'Motivo']].concat(
        ap.cambios.map(c => [c.tabla, c.fila, letraCol(c.col) + c.fila, c.antes, c.despues, c.motivo]),
        ap.quitadas.map(x => [x.tabla || 'POSICIONES', x.fila, '', x.vk, 'fila quitada', x.motivo]),
        [[]], [['Fotos del ZIP', '', '', 'Nombre / VK antiguo', 'Nombre / VK nuevo', '']], Object.keys(estado.mapaFotos).map(k => ['', '', '', k, estado.mapaFotos[k], 'foto renombrada (ruta corregida)']), Object.keys(estado.mapaVK).map(k => ['', '', '', k, estado.mapaVK[k], 'el resto de fotos que empiezan por el VK antiguo']));
      descargar(new Blob(['﻿' + filas.map(f => f.map(q).join(';')).join('\r\n')], { type: 'text/csv;charset=utf-8' }), `Cambios AppSheet - ${ap.ct}.csv`);
    }
  });

  // ───────── Comprobación previa: qué falta o no encaja antes de generar ─────────
  function comprobar() {
    const zona = $('comprobacion'); if (!zona) return;
    let m; try { m = modeloDesdeTablas(); } catch (e) { zona.innerHTML = ''; return; }
    const nom = n => n.esPT ? 'PT ' + n.ct : `apoyo ${n.apoyo} (${C.salName(+n.salida)})`;
    const nomF = r => `apoyo ${r.apoyo} (${C.salName(+r.salida)}, fila ${r.linea})`;
    const vacio = v => String(v == null ? '' : v).trim() === '';
    const filas = m.filas.filter(r => !r.esPT);
    // ── tabla de posiciones ──
    let nNum = 0; try { const p = C.analizarNumeracion(m); nNum = (p.huecos || p.problemas || p || []).length || 0; } catch (e) {}
    const sinIni = filas.filter(r => vacio(r.inicio)).map(nomF);
    const sinFin = filas.filter(r => vacio(r.fin)).map(nomF);
    const sinSec = filas.filter(r => vacio(r.seccion)).map(nomF);
    const RE_COMP = /compart|com[uú]n|doble\s*circ|circuito\s*doble|triple|(\d+|dos|tres|cuatro)\s*circuitos|\bsal(ida)?s?\s*\.?\s*\d|\bS\s?0?\d\b|otra\s+salida|varias\s+salidas|misma\s+l[ií]nea/i;
    const obsComp = filas.filter(r => RE_COMP.test(r.obs || '')).map(r => `${nomF(r)}: «${r.obs}»`);
    const grupos = [
      ['Faltan números de apoyo (arriba puedes renumerar)', nNum ? [`Hay ${nNum} problema${nNum === 1 ? '' : 's'} de numeración: revísalo en el recuadro de numeración de arriba.`] : []],
      ['Falta INICIO', sinIni],
      ['Falta FIN', sinFin],
      ['Falta SECCIÓN (puedes ponerla en «Secciones de los apoyos»)', sinSec],
      ['Observaciones que hablan de apoyos compartidos con otras salidas: indícalo en «Recorrido de salidas que pasan por apoyos de otra»', obsComp]
    ];
    // ── fotos ──
    let ph = null;
    if (estado.fotos.length && estado.ctFotos) {
      ph = C.asociarFotos(m, estado.fotos.map(f => f.nombre), { sufijo1EsCGP: $('chkSuf1').checked });
      grupos.push(
        ['Fotos que no se han podido renombrar', estado.fotos.filter(f => f.campo && (f.campo.estado === 'ignorado' || f.campo.estado === 'omitido')).map(f => `${f.crudo} — ${f.campo.motivo}`)],
        ['Fotos que no encajan con ningún apoyo o defecto', ph.huerfanas.map(x => x.nombre + ' — ' + x.motivo)],
        ['Apoyos sin ninguna foto', m.nodes.filter(n => !n.fotos.length && !(n.fotosCGP || []).length).map(nom)],
        ['Defectos sin foto', m.defectos.filter(d => !d.fotos.length).map(d => `${d.codigo} en ${nom(d.node)}`)],
        ['Apoyos con CGP/CPM sin foto de la CGP/CPM', m.nodes.filter(n => n.cgp && !(n.fotosCGP || []).length).map(nom)]);
    }
    const mal = grupos.filter(g => g[1].length);
    const fotosTxt = ph ? ` (${ph.asociadas} de ${ph.total} fotos asociadas)` : '';
    zona.innerHTML = mal.length
      ? `<div class="t-aviso"><b>Comprobación antes de generar el KMZ</b>${fotosTxt}:</div>` + mal.map(([t, l]) => `<details${l.length <= 8 ? ' open' : ''}><summary class="t-aviso">${t} (${l.length})</summary>${listaAvisos(l, 't-aviso')}</details>`).join('')
      : `<div class="ok-box">✔ Comprobación: numeración, INICIO, FIN y SECCIÓN completos, sin observaciones de apoyos compartidos${ph ? `; las ${ph.total} fotos encajan y todos los apoyos, defectos y CGP/CPM tienen foto` : ''}.</div>` + (estado.fotos.length ? '' : '<div class="info">Carga los ZIP de fotos para comprobarlas también.</div>');
    return mal.length;
  }
  function comprobarLuego() { clearTimeout(estado.tComp2); estado.tComp2 = setTimeout(comprobar, 400); }

  // ───────── PASO 3: Informe BT / Certificado de Revisión BT y carpeta final ─────────
  const DEF_MAP = Object.fromEntries((window.DEFECTOS_BT || []).map(d => [d.c, d]));
  const UP = x => String(x || '').toLocaleUpperCase('es-ES').trim();
  const oracion = x => { x = String(x || '').replace(/\s+/g, ' ').trim().toLocaleLowerCase('es-ES'); return x ? x.charAt(0).toLocaleUpperCase('es-ES') + x.slice(1) : ''; };
  let DIRS = null;
  const cargarDirs = () => DIRS || (DIRS = fetch('plantillas/direcciones.json').then(r => r.ok ? r.json() : {}).catch(() => ({})));
  const LS_POB = 'bajatension.poblaciones', LS_TEC = 'bajatension.tecnicos';
  function prepararPaso3() {
    const inf = estado.informe; if (!inf) return;
    $('p3Vacio').hidden = true; $('p3Form').hidden = false;
    const prov = provinciaCentro(inf.centro);
    $('p3Prov').innerHTML = PROVS.map(p => `<option${p === prov.nombre ? ' selected' : ''}>${esc(p)}</option>`).join('');
    $('p3Cd').value = inf.centro;
    $('p3Fecha').value = inf.fecha || fechaGestcd(inf.centro) || '';
    $('p3Tec').value = prov.tecnico || $('p3Tec').value || '';
    $('p3TecList').innerHTML = lsJson(LS_TEC, []).map(t => `<option value="${esc(t)}">`).join('');
    const pobs = lsJson(LS_POB, {}); $('p3Pob').value = pobs[inf.centro] || '';
    $('p3Dir').value = ''; $('p3DirInfo').textContent = 'Buscando la dirección…';
    Promise.resolve(cargarDirs()).then(d => {
      if (estado.informe !== inf) return;
      const k = UP(inf.centro), r = d[k] != null ? d[k] : d[k.replace(/^0+/, '')];
      if (r) { $('p3Dir').value = oracion(r); $('p3DirInfo').textContent = 'Dirección de la base de datos (puedes corregirla).'; }
      else $('p3DirInfo').textContent = 'El centro no está en la base de datos de direcciones: escríbela.';
    });
    estado.p3Defs = inf.defectos.map(d => { const cod = UP(d.cod); return { apoyo: UP(d.apoyo), cod, des: UP(DEF_MAP[cod] ? DEF_MAP[cod].d : d.des), linea: UP(d.linea || inf.centro), obs: UP(d.obs) }; });
    $('p3Defs').innerHTML = estado.p3Defs.length
      ? `<details><summary>Tabla de defectos del documento (${estado.p3Defs.length}, sin ZOBX)</summary><table><tr><th>Apoyo</th><th>Línea</th><th>Código</th><th>Descripción</th><th>Observaciones</th></tr>${estado.p3Defs.map(d => `<tr><td>${esc(d.apoyo)}</td><td>${esc(d.linea)}</td><td>${esc(d.cod)}</td><td>${esc(d.des)}</td><td>${esc(d.obs)}</td></tr>`).join('')}</table></details>`
      : '<div class="info">Sin defectos (solo ZOBX): el documento llevará «sin defectos».</div>';
    $('res3').innerHTML = '';
    pintarPaso3();
  }
  const esCert = () => esBoletin($('p3Prov').value);
  function pintarPaso3() {
    if (!estado.informe) return;
    const cert = esCert(), ct = estado.informe.centro, doc = cert ? 'CERTIFICADO DE REVISION BT' : 'INFORME BT';
    $('tituloPaso3').textContent = cert ? 'Certificado de Revisión BT' : 'Informe BT';
    $('p3TecWrap').hidden = !cert;
    let firma = null; try { firma = JSON.parse(localStorage.getItem('ibt_firma') || 'null'); } catch (e) {}
    $('p3Firma').textContent = cert ? 'El certificado va sin firma (lo firma el ingeniero a mano).' : firma ? 'El informe lleva la firma guardada en este equipo (la de InformesConBT).' : 'No hay firma guardada en este equipo: el informe saldrá sin firmar (la firma se guarda en InformesConBT).';
    $('btnEnviarInf').textContent = cert ? 'Enviar al Certificado BT' : 'Enviar al Informe BT';
    $('p3Contenido').innerHTML = `La carpeta <b>${esc(ct)}</b> lleva: ANEXO I - ${esc(ct)}.kmz (el modificado), ANEXO I - ${esc(ct)}.jpg (la captura), ANEXO II, ANEXO III - ${esc(ct)}.xlsx, ${doc} - ${esc(ct)}.docx y FOTOS. El PDF lo sacas tú desde Word.`;
  }
  $('btnEnviarInf').addEventListener('click', () => {
    const provincia = $('p3Prov').value;
    const datos = Object.assign({}, estado.informe, { provincia, documento: esBoletin(provincia) ? 'boletin' : 'informe', tecnico: $('p3Tec').value.trim() || provinciaCentro(estado.informe.centro).tecnico || '' });
    try { localStorage.setItem('jdev_kmz_informe', JSON.stringify(datos)); } catch (e) { alert('No se pudo guardar en este navegador'); return; }
    window.open('../InformesConBT/?kmz=1', 'informesconbt');
  });
  $('p3Prov').addEventListener('change', () => { lsSet('bajatension.prov', $('p3Prov').value); pintarPaso3(); });
  $('p3Pob').addEventListener('input', () => { const m = lsJson(LS_POB, {}); m[$('p3Cd').value] = $('p3Pob').value; lsSet(LS_POB, JSON.stringify(m)); });
  // captura del ANEXO I
  zonaDrop($('dropCaptura'), $('fCaptura'), async files => {
    const f = files.find(x => /^image\//.test(x.type) || /\.(jpe?g|png|webp|bmp)$/i.test(x.name)); if (!f) return;
    try {
      const bmp = await createImageBitmap(f);
      const k = Math.min(1, 2000 / Math.max(bmp.width, bmp.height));
      const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * k); cv.height = Math.round(bmp.height * k);
      const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(bmp, 0, 0, cv.width, cv.height);
      const paraWord = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.88));
      let original = f;
      if (!/\.jpe?g$/i.test(f.name)) { const c2 = document.createElement('canvas'); c2.width = bmp.width; c2.height = bmp.height; const g2 = c2.getContext('2d'); g2.fillStyle = '#fff'; g2.fillRect(0, 0, c2.width, c2.height); g2.drawImage(bmp, 0, 0); original = await new Promise(r => c2.toBlob(r, 'image/jpeg', 0.92)); }
      estado.captura = { word: { bytes: new Uint8Array(await paraWord.arrayBuffer()), type: 'image/jpeg', w: cv.width, h: cv.height }, jpg: original };
      $('p3CapPrev').src = URL.createObjectURL(paraWord); $('p3CapPrev').hidden = false; $('btnQuitarCaptura').hidden = false;
      let ctc = estado.informe && estado.informe.centro; if (!ctc) try { ctc = modeloDesdeTablas().ct; } catch (e) {}
      if (ctc) marcarPaso('captura', ctc);
    } catch (e) { alert('No se pudo leer esa imagen.'); }
  });
  $('btnQuitarCaptura').addEventListener('click', () => { estado.captura = null; $('p3CapPrev').hidden = true; $('btnQuitarCaptura').hidden = true; });
  const dataURLBytes = u => Uint8Array.from(atob(u.split(',')[1]), c => c.charCodeAt(0));
  async function generarWord() {
    const inf = estado.informe, cert = esCert();
    const falta = [];
    if (!$('p3Fecha').value) falta.push('la fecha');
    if (!$('p3Dir').value.trim()) falta.push('la dirección');
    if (!$('p3Pob').value.trim()) falta.push('la población');
    if (cert && !$('p3Tec').value.trim()) falta.push('el técnico de campo');
    if (falta.length) { alert('Falta ' + falta.join(', ') + '.'); return null; }
    if (!estado.captura && !confirm('No has subido la captura del ANEXO I: el documento irá sin ella.\n\n¿Seguir?')) { resaltar('dropCaptura'); return null; }
    if (cert) { const t = lsJson(LS_TEC, []), v = $('p3Tec').value.trim(); if (!t.includes(v)) { t.push(v); lsSet(LS_TEC, JSON.stringify(t.slice(-30))); } }
    const r = await fetch(cert ? 'plantillas/plantilla_certificado.docx' : 'plantillas/plantilla.docx');
    if (!r.ok) throw new Error('No encuentro la plantilla del documento.');
    let firma = null; if (!cert) try { firma = JSON.parse(localStorage.getItem('ibt_firma') || 'null'); } catch (e) {}
    let nombreTec = ''; try { nombreTec = JSON.parse(localStorage.getItem('ibt_nombre') || '""'); } catch (e) {}
    const [y, m, d] = $('p3Fecha').value.split('-');
    const data = {
      nombre: nombreTec, tecnico: cert ? $('p3Tec').value.trim() : '', fecha: `${d}/${m}/${y}`, centro: UP(inf.centro), direccion: $('p3Dir').value.trim(), poblacion: $('p3Pob').value.trim(), provincia: $('p3Prov').value,
      sinDefectos: !estado.p3Defs.length, defectos: estado.p3Defs, textoSinDefectos: cert ? 'SIN DEFECTOS DETECTADOS' : '',
      firma: firma ? { bytes: dataURLBytes(firma.data), type: 'image/png', w: firma.w, h: firma.h } : null,
      foto: estado.captura ? estado.captura.word : null
    };
    const docx = await MotorBT.generarDocx(new Uint8Array(await r.arrayBuffer()), data);
    try { const mk = JSON.parse(localStorage.getItem('jdev_informes') || '{}'); mk[UP(inf.centro).replace(/^0+(?=\d)/, '')] = Date.now(); localStorage.setItem('jdev_informes', JSON.stringify(mk)); } catch (e) {}
    return { blob: new Blob([docx], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), nombre: `${cert ? 'CERTIFICADO DE REVISION BT' : 'INFORME BT'} - ${inf.centro}.docx` };
  }
  async function conBoton(id, txt, fn) { const b = $(id), t = b.textContent; b.disabled = true; b.textContent = txt; try { await fn(); } catch (e) { console.error(e); $('res3').innerHTML = `<div class="err-box">${esc(e.message)}</div>`; } finally { b.disabled = false; b.textContent = t; } }
  $('btnWord').addEventListener('click', () => conBoton('btnWord', 'Generando…', async () => {
    const w = await generarWord(); if (!w) return;
    descargar(w.blob, w.nombre);
    $('res3').innerHTML = `<div class="ok-box">✔ <b>${esc(w.nombre)}</b> descargado.</div>`;
  }));
  $('btnCarpetaFinal').addEventListener('click', () => conBoton('btnCarpetaFinal', 'Preparando la carpeta…', async () => {
    const a3 = estado.anexo3, ct = estado.informe.centro;
    const w = await generarWord(); if (!w) return;
    const z = new JSZip(), raiz = z.folder(ct), avisos = [];
    raiz.file(`ANEXO I - ${ct}.kmz`, a3.kmz.blob);
    if (a3.kmz === estado.kmzGenerado) avisos.push('El KMZ de la carpeta es el recién generado, no uno modificado en Google Earth.');
    if (estado.captura) raiz.file(`ANEXO I - ${ct}.jpg`, estado.captura.jpg); else avisos.push('Sin captura del ANEXO I.');
    if (estado.anexo2) raiz.file(estado.anexo2.nombre, estado.anexo2.bytes); else avisos.push('Sin ANEXO II.');
    raiz.file(a3.nombre, a3.blob);
    raiz.file(w.nombre, w.blob);
    const fotos = raiz.folder('FOTOS');
    let nF = 0;
    if (estado.fotos.length) {
      const usados = new Set();
      for (const f of estado.fotos) { let n = f.nombre, i = 2; while (usados.has(n.toLowerCase())) n = f.nombre.replace(/(\.\w+)$/, `_${i++}$1`); usados.add(n.toLowerCase()); fotos.file(n, await f.leer(), { compression: 'STORE' }); nF++; }
    } else if (a3.kzip) {
      // sin fotos cargadas (p. ej. tras recargar la página): las del KMZ, que ya están renombradas y redimensionadas
      const fs = Object.keys(a3.kzip.files).filter(n => /^files\/[^/]+\.(jpe?g|png)$/i.test(n) && !/leyenda\.png$/i.test(n));
      for (const n of fs) { fotos.file(n.slice(6), await a3.kzip.file(n).async('uint8array'), { compression: 'STORE' }); nF++; }
      if (nF) avisos.push('FOTOS sale de las fotos que van dentro del KMZ (no tenías el ZIP de fotos cargado).');
    }
    if (!nF) avisos.push('La carpeta FOTOS va vacía.');
    descargar(await z.generateAsync({ type: 'blob' }), `${ct}.zip`);
    $('res3').innerHTML = `<div class="ok-box">✔ Carpeta <b>${esc(ct)}</b> descargada (${esc(ct)}.zip) con ${nF} fotos y ${esc(w.nombre)}.</div>${avisos.length ? listaAvisos(avisos, 't-aviso') : ''}`;
  }));

  // ───────── secciones de los apoyos ─────────
  function selectorSeccion(valor, attrs) {
    const v = C.normSeccion(valor || '');
    const ops = ['<option value="">—</option>'].concat(C.SECCIONES.map(x => `<option${x === v ? ' selected' : ''}>${x}</option>`));
    if (v && !C.SECCIONES.includes(v)) ops.push(`<option selected value="${esc(valor)}">${esc(valor)}</option>`);
    ops.push('<option value="__otra">Otra…</option>');
    return `<select ${attrs}>${ops.join('')}</select>`;
  }
  // Selectores de una fila: la 1ª es la de su salida; luego las de tramos compartidos y las que se añadan con «+ sección»
  function celdaSecciones(f, n) {
    const partes = trim(f.r.seccion).split(/\s*\/\s*/);
    const total = Math.max(n || 0, partes.filter(Boolean).length, 1 + f.compartidas.length, 1);
    let h = '';
    for (let i = 0; i < total; i++) {
      const etiqueta = i === 0 ? '' : (f.compartidas[i - 1] ? `Compartido con ${C.salName(f.compartidas[i - 1].salida)}` : `Sección ${i + 1}`);
      const quitar = i > 0 && !f.compartidas[i - 1] ? '<button type="button" class="sec peq sx" title="Quitar esta sección">×</button>'
        : i > 0 ? `<button type="button" class="sec peq sxsal" data-sal="${f.compartidas[i - 1].salida}" title="Quitar esta salida del tramo">×</button>` : '';
      h += i === 0 ? `<div class="comp" style="margin-top:0">${selectorSeccion(partes[0] || '', 'data-p="0"')}<button type="button" class="sec peq smas">+ sección</button><button type="button" class="sec peq ssal" title="Otra salida va también por el tramo que llega a este apoyo (línea paralela)">+ salida</button><button type="button" class="sec peq scop" title="Copiar las secciones y salidas de este apoyo para pegarlas en otros">Copiar</button></div>`
        : `<div class="comp">${etiqueta}: ${selectorSeccion(partes[i] || '', `data-p="${i}"`)}${quitar}</div>`;
    }
    return h;
  }
  function pintarSecciones() {
    comprobarLuego();
    const zona = $('secciones');
    if (!$('detSecc').open) return;
    let m = null;
    try { if ($('txtMain').value.trim()) m = modeloDesdeTablas(); } catch (e) { m = null; }
    if (!m) { zona.innerHTML = '<div class="info">Pega primero la tabla de posiciones.</div>'; return; }
    const filas = C.filasSecciones(m);
    estado.secc = { m, filas };
    const salidas = [...new Set(filas.map(f => f.salida))].sort((a, b) => a - b);
    const varios = new Set(filas.map(f => f.cuadro)).size > 1;
    let h = `<div class="secc-bulk"><b style="width:100%;font-size:13px">Poner la misma sección a varios apoyos seguidos</b>
      <label>Salida<select id="sbSal">${salidas.map(s => `<option value="${s}">${C.salName(s)}</option>`).join('')}</select></label>
      <label>Desde apoyo<input id="sbDe" type="number" min="1"></label><label>Hasta apoyo<input id="sbA" type="number" min="1"></label>
      <label>Sección${selectorSeccion('', 'id="sbSec"')}</label><button class="peq" id="btnSecBulk">Aplicar</button>
      <label>o va también la salida<input id="sbComp" type="number" min="1" style="width:70px" title="Puede ser una salida que no tenga apoyos propios en la tabla"></label><button class="peq sec" id="btnCompBulk">Añadir salida</button></div>
      <div class="info">La salida que va «también» puede ser una que no tenga ningún apoyo propio (por ejemplo, la SAL02 que va junto a la SAL01 en 5 apoyos y termina con ella): se dibuja su línea paralela y cuentan sus metros.</div>`;
    const cl = estado.clipSecc;
    if (cl) h += `<div class="num-box" id="barraPegar" style="margin:8px 0">📋 Copiado el <b>apoyo ${esc(cl.de)}</b>: ${esc(cl.resumen)}. Marca los apoyos donde pegarlo
      <div class="acciones" style="margin-top:6px"><button class="peq" id="btnPegar">Pegar en los marcados</button><button class="sec peq" id="btnMarcarSal">Marcar todos los de ${esc(C.salName(cl.salida))}</button><button class="sec peq" id="btnCancelarCopia">Cancelar</button></div></div>`;
    h += `<div class="scroll"><table class="tsec"><tr>${cl ? '<th></th>' : ''}<th>Apoyo</th><th>Salida</th><th>Sección</th></tr>` + filas.map((f, i) => `<tr data-i="${i}">${cl ? `<td><input type="checkbox" class="spegar"${f.cuadro === cl.f.cuadro && f.salida === cl.f.salida && f.apoyo === cl.f.apoyo ? ' disabled' : ''}></td>` : ''}<td><b>${esc(f.apoyo)}</b> <span class="info">${esc(f.r.node.name)}</span></td><td>${varios ? 'C' + f.cuadro + ' ' : ''}${C.salName(f.salida)}</td><td>${celdaSecciones(f)}</td></tr>`).join('') + '</table></div>';
    h += '<div class="info">Los cambios se escriben en la tabla de posiciones (columna SECCIÓN) y salen en el KMZ y en el ANEXO III. Con varias secciones en un apoyo la celda queda como «RZ95 / RZ50» (la de su salida primero).<br><b>+ salida</b>: otra salida va también por el tramo que llega a ese apoyo (circuito doble). Se dibuja su línea paralela en el KMZ, sus metros cuentan en el KMZ y en el ANEXO III, y aquí puedes ponerle su sección. Se anota en el recuadro «Recorrido de salidas…» de abajo.</div>';
    zona.innerHTML = h;
  }
  // Repinta la celda con una sección más, conservando lo elegido
  function celdaSeccionesDesde(f, tr, n) {
    const actual = [...tr.querySelectorAll('select')].map(s => s.value === '__otra' ? '' : s.value);
    const falso = { r: { seccion: actual.join(' / ') }, compartidas: f.compartidas };
    const html = celdaSecciones(falso, n);
    // los huecos vacíos se respetan: se rellenan por posición
    const tmp = document.createElement('td'); tmp.innerHTML = html;
    [...tmp.querySelectorAll('select')].forEach((s, i) => { if (actual[i] !== undefined && [...s.options].some(o => o.value === actual[i])) s.value = actual[i]; });
    return tmp.innerHTML;
  }

  // ── «+ salida»: otra salida pasa por el tramo que llega a este apoyo (circuito doble) ──
  // Se escribe en el recuadro de recorridos como «SAL01 3-4 SAL02» (del 3 al 4 de la SAL01 va también la SAL02):
  // así se dibuja la paralela, se cuentan sus metros en el KMZ y en el ANEXO III y la sección sale en el editor.
  function tramoQueLlega(f) { return estado.secc.m.tramos.find(t => t.b === f.r.node && t.dist != null) || null; }
  const lineaComp = (f, t, d) => `${C.salName(f.salida)} ${t.a.esPT ? 0 : t.a.apoyo}-${f.apoyo} ${C.salName(d)}`;
  const normLinea = l => l.toUpperCase().replace(/SALIDA/g, 'SAL').replace(/SAL\s*0*(\d+)/g, 'SAL$1').replace(/\s+/g, ' ').replace(/\s*-\s*/g, '-').trim();
  function escribirComp(lineas) { $('txtComp').value = lineas.join('\n'); lsSet(LS.comp, $('txtComp').value); pintarSecciones(); }
  function anadirSalidaTramo(f, d) {
    const t = tramoQueLlega(f);
    if (!t) { alert(`El apoyo ${f.apoyo} no tiene tramo que llegue a él (revisa INICIO/FIN).`); return; }
    if (d === f.salida) return;
    const ls = $('txtComp').value.split(/\r?\n/).filter(x => x.trim());
    const nueva = lineaComp(f, t, d);
    if (!ls.some(x => normLinea(x) === normLinea(nueva))) ls.push(nueva);
    escribirComp(ls);
  }
  function quitarSalidaTramo(f, d) {
    const t = tramoQueLlega(f); if (!t) return;
    const quitar = normLinea(lineaComp(f, t, d));
    const ls = $('txtComp').value.split(/\r?\n/).filter(x => x.trim());
    const quedan = ls.filter(x => normLinea(x) !== quitar);
    if (quedan.length === ls.length) { alert(`Este tramo compartido con ${C.salName(d)} viene de una línea del recuadro «Recorrido de salidas…» que abarca más apoyos. Quítalo o cámbialo ahí.`); return; }
    escribirComp(quedan);
  }
  $('secciones').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-i]'); if (!tr) return;
    const f = estado.secc.filas[+tr.dataset.i];
    if (e.target.classList.contains('ssal')) {
      // salidas con apoyos en la tabla + las que solo existen como compartidas (sin apoyos propios)
      const enComp = estado.secc.filas.flatMap(x => x.compartidas.map(c => c.salida));
      const todas = [...new Set(estado.secc.filas.map(x => x.salida).concat(enComp))];
      const sals = todas.filter(x => x !== f.salida && !f.compartidas.some(c => c.salida === x)).sort((a, b) => a - b);
      const nueva = Math.max(0, ...todas) + 1;
      const conApoyos = new Set(estado.secc.filas.map(x => x.salida));
      const sel = document.createElement('select'); sel.className = 'ssel';
      sel.innerHTML = '<option value="">¿Qué salida?</option>' + sals.map(x => `<option value="${x}">${C.salName(x)}${conApoyos.has(x) ? '' : ' (sin apoyos propios)'}</option>`).join('')
        + `<option value="${nueva}">${C.salName(nueva)} — salida nueva sin apoyos propios</option><option value="__n">Otra (escribir número)…</option>`;
      e.target.replaceWith(sel); sel.focus();
      sel.addEventListener('change', () => {
        let d = sel.value; if (!d) return;
        if (d === '__n') { d = prompt('Número de la salida (por ejemplo 4):', ''); if (!d || !/^\d+$/.test(d.trim())) { pintarSecciones(); return; } }
        anadirSalidaTramo(f, +d);
      });
      return;
    }
    if (e.target.classList.contains('sxsal')) { quitarSalidaTramo(f, +e.target.dataset.sal); return; }
  });


  // ── Copiar / pegar secciones y salidas compartidas entre apoyos ──
  function copiarApoyo(f, tr) {
    const sels = [...tr.querySelectorAll('select:not(.ssel)')].map(x => x.value === '__otra' ? '' : x.value);
    const comp = {}; f.compartidas.forEach((c, i) => { comp[c.salida] = sels[i + 1] || ''; });
    const extras = sels.slice(1 + f.compartidas.length).filter(Boolean);
    const partes = [sels[0] ? `${C.salName(f.salida)} ${sels[0]}` : C.salName(f.salida)].concat(f.compartidas.map(c => `${C.salName(c.salida)}${comp[c.salida] ? ' ' + comp[c.salida] : ''}`), extras);
    estado.clipSecc = { f, de: f.apoyo + ' (' + C.salName(f.salida) + ')', salida: f.salida, propia: sels[0] || '', comp, extras, resumen: partes.join(' · ') };
    pintarSecciones();
  }
  function pegarEn(objetivos) {
    const cl = estado.clipSecc; if (!cl || !objetivos.length) return;
    // 1) salidas compartidas: una línea «SALxx a-b SALyy» por tramo
    const ls = $('txtComp').value.split(/\r?\n/).filter(x => x.trim());
    const sinTramo = [];
    objetivos.forEach(f => {
      const t = tramoQueLlega(f);
      Object.keys(cl.comp).map(Number).forEach(d => {
        if (d === f.salida) return;
        if (!t) { sinTramo.push(f.apoyo); return; }
        const l = lineaComp(f, t, d); if (!ls.some(x => normLinea(x) === normLinea(l))) ls.push(l);
      });
    });
    $('txtComp').value = ls.join('\n'); lsSet(LS.comp, $('txtComp').value);
    // 2) secciones, en el orden de salidas que tenga cada apoyo después de compartir
    const m2 = modeloDesdeTablas(), f2 = C.filasSecciones(m2);
    const clave = f => `${f.cuadro}|${f.salida}|${f.apoyo}`, objs = new Set(objetivos.map(clave));
    const cambios = f2.filter(f => objs.has(clave(f))).map(f => {
      const partes = [cl.propia].concat(f.compartidas.map(c => cl.comp[c.salida] != null ? cl.comp[c.salida] : (f.partes[f.compartidas.indexOf(c) + 1] || '')), cl.extras);
      while (partes.length > 1 && !partes[partes.length - 1]) partes.pop();
      return { r: f.r, seccion: partes.join(' / ') };
    });
    const txt = C.editarTabla($('txtMain').value, m2, cambios);
    $('txtMain').value = txt; lsSet(LS.main, txt); infoTabla($('txtMain'), 'infoMain');
    estado.clipSecc = null;
    pintarSecciones(); pintarNumeracion();
    $('secciones').querySelectorAll('tr[data-i]').forEach(tr => { const f = estado.secc.filas[+tr.dataset.i]; if (objs.has(clave(f))) tr.classList.add('mod'); });
    if (sinTramo.length) alert(`No se ha podido compartir la salida en ${sinTramo.length === 1 ? 'el apoyo' : 'los apoyos'} ${[...new Set(sinTramo)].join(', ')}: no tienen tramo que llegue a ellos (revisa INICIO/FIN). Las secciones sí se han pegado.`);
  }
  $('secciones').addEventListener('click', e => {
    if (e.target.classList.contains('scop')) { const tr = e.target.closest('tr[data-i]'); copiarApoyo(estado.secc.filas[+tr.dataset.i], tr); return; }
    if (e.target.id === 'btnCancelarCopia') { estado.clipSecc = null; pintarSecciones(); return; }
    if (e.target.id === 'btnMarcarSal') { const sal = estado.clipSecc.salida; $('secciones').querySelectorAll('tr[data-i]').forEach(tr => { const c = tr.querySelector('.spegar'); if (c && !c.disabled && estado.secc.filas[+tr.dataset.i].salida === sal) c.checked = true; }); return; }
    if (e.target.id === 'btnPegar') {
      const objs = [...$('secciones').querySelectorAll('tr[data-i]')].filter(tr => { const c = tr.querySelector('.spegar'); return c && c.checked; }).map(tr => estado.secc.filas[+tr.dataset.i]);
      if (!objs.length) return alert('Marca primero los apoyos donde quieres pegar.');
      pegarEn(objs);
    }
  });

  function guardarSecciones(cambios) {
    const m = estado.secc.m;
    const txt = C.editarTabla($('txtMain').value, m, cambios);
    $('txtMain').value = txt; lsSet(LS.main, txt);
  }
  function valorFila(tr) {
    return [...tr.querySelectorAll('select:not(.ssel)')].map(s => s.value === '__otra' ? '' : s.value.trim()).filter(Boolean).join(' / ');
  }
  $('secciones').addEventListener('change', e => {
    const sel = e.target.closest('select'); if (!sel || sel.classList.contains('ssel')) return;
    if (sel.value === '__otra') {
      const v = prompt('Escribe la sección:', '');
      if (v && v.trim()) { const o = document.createElement('option'); o.value = o.textContent = C.normSeccion(v); sel.insertBefore(o, sel.lastElementChild); sel.value = o.value; }
      else { sel.value = ''; }
    }
    if (sel.id === 'sbSec') return;
    const tr = sel.closest('tr[data-i]'); if (!tr) return;
    const f = estado.secc.filas[+tr.dataset.i];
    guardarSecciones([{ r: f.r, seccion: valorFila(tr) }]);
    tr.classList.add('mod');
  });
  $('secciones').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-i]');
    if (tr && e.target.classList.contains('smas')) {
      const f = estado.secc.filas[+tr.dataset.i];
      const n = tr.querySelectorAll('select').length + 1;
      tr.cells[2].innerHTML = celdaSeccionesDesde(f, tr, n);
      return;
    }
    if (tr && e.target.classList.contains('sx')) {
      e.target.closest('.comp').remove();
      const f = estado.secc.filas[+tr.dataset.i];
      guardarSecciones([{ r: f.r, seccion: valorFila(tr) }]); tr.classList.add('mod');
      tr.cells[2].innerHTML = celdaSeccionesDesde(f, tr, tr.querySelectorAll('select').length); // renumerar etiquetas
      return;
    }
    if (e.target.id === 'btnCompBulk') {
      const sal = +$('sbSal').value, de = +$('sbDe').value, a = +$('sbA').value, d = +$('sbComp').value;
      if (!de || !a || !d) return alert('Pon desde qué apoyo, hasta cuál y qué salida va también por esos tramos.');
      if (d === sal) return alert('Esa es la misma salida.');
      const [lo, hi] = de <= a ? [de, a] : [a, de];
      const fs = estado.secc.filas.filter(f => f.salida === sal && +f.apoyo >= lo && +f.apoyo <= hi);
      if (!fs.length) return alert('No hay apoyos en ese rango.');
      const ls = $('txtComp').value.split(/\r?\n/).filter(x => x.trim());
      let n = 0; fs.forEach(f => { const t = tramoQueLlega(f); if (!t) return; const l = lineaComp(f, t, d); if (!ls.some(x => normLinea(x) === normLinea(l))) { ls.push(l); n++; } });
      escribirComp(ls);
      return;
    }
    if (e.target.id !== 'btnSecBulk') return;
    const sal = +$('sbSal').value, de = +$('sbDe').value, a = +$('sbA').value, sec = $('sbSec').value === '__otra' ? '' : $('sbSec').value;
    if (!de || !a || !sec) return alert('Pon desde qué apoyo, hasta cuál y la sección.');
    const [lo, hi] = de <= a ? [de, a] : [a, de];
    const cambios = estado.secc.filas.filter(f => f.salida === sal && +f.apoyo >= lo && +f.apoyo <= hi).map(f => {
      const resto = trim(f.r.seccion).split(/\s*\/\s*/).slice(1);
      return { r: f.r, seccion: [sec].concat(resto).filter((x, i) => i === 0 || x).join(' / ') };
    });
    if (!cambios.length) return alert('No hay apoyos en ese rango.');
    guardarSecciones(cambios); pintarSecciones();
    $('secciones').querySelectorAll('tr[data-i]').forEach(tr => { const f = estado.secc.filas[+tr.dataset.i]; if (f.salida === sal && +f.apoyo >= lo && +f.apoyo <= hi) tr.classList.add('mod'); });
  });
  const trim = s => String(s == null ? '' : s).trim();
  $('detSecc').addEventListener('toggle', pintarSecciones);

  // ───────── numeración de apoyos: huecos, repetidos y renumeración ─────────
  function mapaTotal() { return lsJson(LS.mapa, {}); }
  function renombrar(nombre) {
    const m = mapaTotal(), ks = Object.keys(m).sort((a, b) => b.length - a.length);
    if (!ks.length) return nombre;
    const re = new RegExp('(?<!\\d)(' + ks.join('|') + ')(?!\\d)', 'g');
    return String(nombre).replace(re, x => m[x]);
  }
  const trimTxt = v => String(v == null ? '' : v).trim();
  function pintarNumeracion() {
    const zona = $('numeracion');
    let m = null;
    try { if ($('txtMain').value.trim()) m = modeloDesdeTablas(); } catch (e) { m = null; }
    const mapa = mapaTotal(), nCamb = Object.keys(mapa).filter(k => mapa[k] !== k).length;
    let h = '';
    if (nCamb) {
      h += `<div class="num-box ok" style="margin-bottom:8px">✔ Renumeración aplicada: ${nCamb} apoyo${nCamb === 1 ? '' : 's'} con número nuevo (tabla, defectos y fotos).
        <button class="sec peq" id="btnDeshacerNum">Deshacer</button> ${estado.fotos.length ? '<button class="sec peq" id="btnZipRen">Descargar fotos renombradas</button>' : ''}
        ${estado.avisoRegenerar ? '<br><b>Vuelve a generar el KMZ</b>: el que tenías es de antes del cambio.' : ''}</div>`;
    }
    if (!m) { zona.innerHTML = h; enganchar(); return; }
    const problemas = C.analizarNumeracion(m);
    const salidas = [...new Set(m.nodes.filter(n => !n.esPT).map(n => n.cuadro + '|' + n.salida))].map(k => k.split('|').map(Number));
    if (!problemas.length) h += `<div class="num-box ok">✔ Numeración de apoyos sin huecos ni repetidos.</div>`;
    else {
      h += '<div class="num-box"><b>Revisa la numeración de los apoyos</b>';
      problemas.forEach(g => {
        const nom = g.seguida ? `Cuadro ${g.cuadro} (numeración seguida entre salidas)` : (salidas.some(x => x[0] !== g.cuadro) ? `Cuadro ${g.cuadro} · ` : '') + C.salName(g.salida);
        const txt = [];
        if (g.faltan.length) txt.push(`falta${g.faltan.length > 1 ? 'n' : ''} el${g.faltan.length > 1 ? ' los' : ''} ${g.faltan.join(', ')} (${g.total} apoyos numerados hasta el ${g.max})`);
        if (g.repetidos.length) txt.push(`número${g.repetidos.length > 1 ? 's' : ''} repetido${g.repetidos.length > 1 ? 's' : ''}: ${g.repetidos.join(', ')} (corrígelo en la tabla)`);
        const prop = g.faltan.length && !g.repetidos.length ? C.propuestaCompactar(m, g.cuadro, g.salida) : [];
        h += `<div class="num-g">• <b>${nom}</b>: ${txt.join('; ')}.` + (prop.length ? ` <button class="peq" data-compactar="${g.cuadro}|${g.salida == null ? '' : g.salida}">Quitar huecos (${prop[0].node.apoyo}–${prop[prop.length - 1].node.apoyo} → ${prop[0].nuevo}–${prop[prop.length - 1].nuevo})</button>` : '') + '</div>';
      });
      h += '</div>';
    }
    h += `<details style="margin-top:6px"><summary>Cambiar el número de un apoyo a mano</summary><div class="num-man">
      <label>Salida<select id="nmSal">${salidas.map(([c, sa]) => `<option value="${c}|${sa}">${salidas.some(x => x[0] !== c) ? 'C' + c + ' ' : ''}${C.salName(sa)}</option>`).join('')}</select></label>
      <label>Apoyo actual<input id="nmDe" type="number" min="1"></label>
      <label>Número nuevo<input id="nmA" type="number" min="1"></label>
      <label style="flex-direction:row;align-items:center;gap:6px;width:auto"><input id="nmSig" type="checkbox" style="width:auto"> y los siguientes también</label>
      <button class="peq" id="btnNumMan">Aplicar</button></div></details>
    <details style="margin-top:6px"${estado.ifAbierto ? ' open' : ''} id="detIniFin"><summary>Cambiar el INICIO y el FIN de un tramo</summary><div class="num-man">
      <label>Salida<select id="ifSal">${salidas.map(([c, sa]) => `<option value="${c}|${sa}">${salidas.some(x => x[0] !== c) ? 'C' + c + ' ' : ''}${C.salName(sa)}</option>`).join('')}</select></label>
      <label>Apoyo<input id="ifAp" type="number" min="1"></label>
      <label>INICIO<input id="ifIni" type="text" inputmode="numeric" style="width:80px"></label>
      <label>FIN<input id="ifFin" type="text" inputmode="numeric" style="width:80px"></label>
      <button class="peq" id="btnIniFin">Aplicar</button></div>
      <div class="info" id="ifInfo">Pon la salida y el apoyo: se rellenan su INICIO y su FIN actuales para que los cambies. El tramo va del apoyo INICIO al apoyo FIN (0 es el PT).</div></details>`;
    zona.innerHTML = h; enganchar();
    function enganchar() {
      zona.querySelectorAll('[data-compactar]').forEach(b => b.onclick = () => { const [c, sa] = b.dataset.compactar.split('|'); const mm = modeloDesdeTablas(); aplicarRenumeracion(mm, C.propuestaCompactar(mm, +c, sa === '' ? null : +sa)); });
      if ($('btnNumMan')) $('btnNumMan').onclick = () => {
        const [c, sa] = $('nmSal').value.split('|').map(Number), de = +$('nmDe').value, a = +$('nmA').value;
        if (!de || !a) return alert('Pon el número actual y el nuevo.');
        const mm = modeloDesdeTablas();
        const prop = C.propuestaCambio(mm, c, sa, de, a, $('nmSig').checked);
        if (!prop.length) return alert(`No hay ningún apoyo ${de} en esa salida.`);
        aplicarRenumeracion(mm, prop);
      };
      if ($('btnDeshacerNum')) $('btnDeshacerNum').onclick = deshacerRenumeracion;
      // INICIO / FIN de un tramo
      const filaIF = () => {
        const [c, sa] = $('ifSal').value.split('|').map(Number), ap = +$('ifAp').value; if (!ap) return null;
        const mm = modeloDesdeTablas(); const n = mm.nodes.find(x => !x.esPT && x.cuadro === c && x.salida === sa && +x.apoyo === ap);
        return n ? { mm, n } : null;
      };
      if ($('ifAp')) {
        const rellenar = () => { const x = filaIF(); if (!x) { $('ifInfo').textContent = $('ifAp').value ? 'No hay ese apoyo en esa salida.' : 'Pon la salida y el apoyo.'; return; } $('ifIni').value = trimTxt(x.n.row.inicio); $('ifFin').value = trimTxt(x.n.row.fin); $('ifInfo').textContent = `Apoyo ${x.n.apoyo} (${C.salName(x.n.salida)}): ahora INICIO ${trimTxt(x.n.row.inicio) || '—'} y FIN ${trimTxt(x.n.row.fin) || '—'}.`; };
        $('ifAp').oninput = rellenar; $('ifSal').onchange = rellenar;
        $('detIniFin').ontoggle = () => { estado.ifAbierto = $('detIniFin').open; };
        $('btnIniFin').onclick = () => {
          const x = filaIF(); if (!x) return alert('Elige una salida y un apoyo que exista.');
          const ini = $('ifIni').value.trim(), fin = $('ifFin').value.trim();
          if (!/^\d*$/.test(ini) || !/^\d*$/.test(fin)) return alert('INICIO y FIN tienen que ser números de apoyo (0 para el PT).');
          const txt = C.editarTabla($('txtMain').value, x.mm, [{ r: x.n.row, inicio: ini, fin }]);
          $('txtMain').value = txt; lsSet(LS.main, txt); infoTabla($('txtMain'), 'infoMain');
          if (estado.kmzGenerado) { estado.avisoRegenerar = true; if (estado.kmzModificado === estado.kmzGenerado) estado.kmzModificado = null; estado.kmzGenerado = null; $('btnUsarGenerado').hidden = true; $('res1').innerHTML = ''; actualizarBtnExcel(); }
          const ap = $('ifAp').value, sal = $('ifSal').value;
          revisarTablas(); pintarNumeracion(); pintarSecciones(); comprobarLuego();
          if ($('ifAp')) { $('ifSal').value = sal; $('ifAp').value = ap; $('ifInfo').textContent = `✔ Apoyo ${ap}: INICIO ${ini || '—'} y FIN ${fin || '—'} guardados en la tabla.`; $('ifIni').value = ini; $('ifFin').value = fin; }
        };
      }
      if ($('btnZipRen')) $('btnZipRen').onclick = zipRenombradas;
    }
  }
  function aplicarRenumeracion(m, prop) {
    const r = C.renumerar(m, $('txtMain').value, $('txtDef').value, prop);
    if (r.errores.length) return alert('No se puede aplicar:\n' + r.errores.join('\n'));
    const lista = prop.slice(0, 12).map(x => `${x.node.apoyo}→${x.nuevo}`).join(', ') + (prop.length > 12 ? '…' : '');
    if (!confirm(`Se van a cambiar ${prop.length} apoyo${prop.length === 1 ? '' : 's'} (${lista}).\n\nSe actualizan los números y los INICIO/FIN de la tabla, los códigos de la tabla de defectos y el nombre de las fotos con el código del apoyo nuevo.\n\n¿Seguir?`)) return;
    if (!lsGet(LS.orig)) lsSet(LS.orig, JSON.stringify({ main: $('txtMain').value, def: $('txtDef').value }));
    // componer con lo ya aplicado: original → actual
    const tot = mapaTotal(), vals = new Set(Object.values(tot));
    Object.keys(tot).forEach(k => { if (r.mapa[tot[k]]) tot[k] = r.mapa[tot[k]]; });
    Object.keys(r.mapa).forEach(k => { if (!vals.has(k) && !(k in tot)) tot[k] = r.mapa[k]; });
    lsSet(LS.mapa, JSON.stringify(tot));
    $('txtMain').value = r.mainText; $('txtDef').value = r.defText; lsSet(LS.main, r.mainText); lsSet(LS.def, r.defText);
    infoTabla($('txtMain'), 'infoMain'); infoTabla($('txtDef'), 'infoDef');
    estado.fotos.forEach(f => { f.nombre = renombrar(f.orig); });
    estado.fotoElegida = {}; pintarSecciones();
    if (estado.kmzGenerado) { estado.avisoRegenerar = true; if (estado.kmzModificado === estado.kmzGenerado) estado.kmzModificado = null; estado.kmzGenerado = null; $('btnUsarGenerado').hidden = true; $('res1').innerHTML = ''; actualizarBtnExcel(); }
    pintarNumeracion();
  }
  function deshacerRenumeracion() {
    const o = lsJson(LS.orig, null); if (!o) return;
    if (!confirm('¿Volver a la numeración original? Se restauran las tablas y los nombres de las fotos.')) return;
    $('txtMain').value = o.main; $('txtDef').value = o.def; lsSet(LS.main, o.main); lsSet(LS.def, o.def);
    try { localStorage.removeItem(LS.orig); localStorage.removeItem(LS.mapa); } catch (e) {}
    infoTabla($('txtMain'), 'infoMain'); infoTabla($('txtDef'), 'infoDef');
    estado.fotos.forEach(f => { f.nombre = f.orig; }); estado.fotoElegida = {}; pintarFotos();
    pintarNumeracion(); pintarSecciones();
  }
  async function zipRenombradas() {
    const z = new JSZip();
    for (const f of estado.fotos) z.file(f.nombre, await f.leer(), { compression: 'STORE' });
    let ct = ''; try { ct = modeloDesdeTablas().ct; } catch (e) {}
    descargar(await z.generateAsync({ type: 'blob' }), `Fotos renombradas${ct ? ' - ' + ct : ''}.zip`);
  }
  revisarTablas(); pintarNumeracion(); estado.mapaListo = true; actualizarBtnExcel(); comprobarLuego();

  // ───────── PWA y aviso de versión nueva ─────────
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
  async function comprobarVersion() {
    try {
      const r = await fetch('version.json?t=' + Date.now(), { cache: 'no-store' });
      const v = (await r.json()).version;
      if (v && v !== APP_VERSION) $('banner').classList.add('on');
    } catch (e) { /* sin conexión */ }
  }
  $('btnActualizar').addEventListener('click', async () => {
    const b = $('btnActualizar'); b.disabled = true; b.textContent = 'Actualizando…';
    try {
      // 1) descargar de nuevo los archivos de la app saltándose la caché del navegador
      const archivos = ['./', 'index.html', 'js/app.js', 'js/core.js', 'js/template.js', 'vendor/jszip.min.js', 'sw.js', 'manifest.webmanifest'];
      await Promise.all(archivos.map(u => fetch(u, { cache: 'reload' }).catch(() => {})));
      // 2) borrar la copia sin conexión y actualizar el service worker
      const ks = await caches.keys(); await Promise.all(ks.map(k => caches.delete(k)));
      if ('serviceWorker' in navigator) {
        const regs = await navigator.serviceWorker.getRegistrations();
        await Promise.all(regs.map(r => r.update().catch(() => {})));
      }
    } catch (e) {}
    // 3) recargar pidiendo la página nueva
    location.replace(location.pathname + '?v=' + Date.now());
  });
  if (/[?&]v=\d+/.test(location.search)) history.replaceState(null, '', location.pathname);
  comprobarVersion(); setInterval(comprobarVersion, 30 * 60 * 1000);
  window.__APP_VERSION = APP_VERSION;
})();
