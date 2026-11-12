# Hub Evaluación 360 · Healthy Food

App web (PWA, sin framework) para la evaluación 360 de los cargos del **nivel estratégico** del Manual de Funciones (GTH-MFC-01 v07).
Reemplaza el formulario por una conversación guiada: cada persona y su jefe responden las mismas preguntas abiertas,
luego se sientan en la **sala de acuerdo**, cierran juntos la **ficha** (el formato de Talento Humano) y le dan **seguimiento a 30 · 60 · 90 días**.

Publicada en **https://sisgeshealthy.github.io/hub-evaluacion360/** (ingreso con la cuenta Microsoft 365 del cargo).
Prueba local: servir esta carpeta en `http://localhost:5500/` (URI registrada en Azure) o abrir `…/?demo` para la demo con datos de ejemplo
(la demo necesita los archivos locales `datos-manual.js` y `datos-personas.js`, que no se publican).

## Cómo funciona

| Paso | Quién | Qué ve |
|---|---|---|
| 1. Autoevaluación | La persona | Una pregunta por pantalla, con ayuda, ejemplo, competencias de su cargo para tocar y pistas de calidad ("¿tiene unidad?", "¿empieza con verbo?"). La ficha se arma en vivo a la derecha. Se guarda sola. |
| 2. Evaluación del jefe | Jefe inmediato (organigrama) | Las mismas preguntas, redactadas hacia la otra persona. **A ciegas**: no ve la autoevaluación hasta que los dos envían. |
| 3. Miradas 360 (opcional) | Pares y colaboradores | Preguntas cortas (fortaleza y conducta). **Anónimas** para el evaluado y el jefe. Se activan en Configuración → Ciclo. Los pares rotan: cada persona da y recibe como máximo *N* miradas. |
| 4. Sala de acuerdo | Jefe (con la persona) | Las dos miradas lado a lado, con un indicador de coincidencia ("Ven lo mismo" / "Coinciden en parte" / "Miradas distintas · conversar"). Se escribe lo acordado y se cierra. |
| 5. Confirmación | La persona | Confirma la ficha (con observación opcional). Desde ahí corre el seguimiento. |
| 6. Seguimiento | Jefe o persona | A los 30, 60 y 90 días se marca cada acción del plan (sin iniciar / en curso / cumplida) con una nota. |

**Panel de Talento Humano:** avance del ciclo, embudo, estado de cada persona, recordatorio por correo, temas de desarrollo
(palabras y competencias que más aparecen en las conductas y planes: insumo para el plan de capacitación), exportación a Excel (CSV)
y **trazabilidad** por persona (quién hizo qué y cuándo; reabrir respuestas o acuerdos queda registrado con motivo).

**Configuración (solo Talento Humano y administradores):**
- **Ciclo:** nombre, fechas, días de seguimiento, fuentes 360 (pares, colaboradores), anonimato, evaluación a ciegas. Muestra la carga de evaluaciones por persona.
- **Cargos que participan:** los 22 cargos del manual, agrupados por nivel. Por defecto: nivel estratégico menos Presidente y Gerente General (el Gerente General solo evalúa; nadie lo evalúa).
- **Personas y jefes:** del organigrama del manual; aquí se cargan los correos de Microsoft 365. Muestra el organigrama resultante.
- **Campos de la ficha:** activar/desactivar, renombrar, cambiar preguntas por rol, ayuda, ejemplo, tipo (texto abierto, frase corta, lista de acciones, por competencia), dónde aparece en la ficha, orden. Vista previa de la ficha.
- **Accesos:** correos administradores.

## Archivos

- `index.html`, `styles.css`, `app.js`: la app. `config.js`: modo, SharePoint, Azure y administradores.
- **Datos internos (no se publican, `.gitignore`):** `datos-manual.js` (cargos, misión, jefe, competencias del Manual de Funciones)
  y `datos-personas.js` (organigrama, jefes y cuentas). La app publicada los lee de SharePoint (fila "Configuración" de la lista).
  - Si cambia el manual: `py -I herramientas/manual-a-datos.py "ruta/MANUAL DE FUNCIONES.docx"`, abrir la app en `localhost:5500`,
    Configuración → cambiar cualquier cosa → Guardar: el manual se sube a SharePoint.
  - Personas, jefes y correos se editan directamente en Configuración → Personas.
- `sw.js`, `manifest.webmanifest`, `icons/`: instalable y abre sin señal.
- Al publicar cambios, subir el `?v=` de `index.html` y `VERSION` en `sw.js`.

## Dónde se guarda

- Sitio SharePoint `/sites/Evaluacion360`, lista "Evaluación 360 - Registros" (dirección `Lists/Evaluacion360`): una fila por respuesta,
  acuerdo y la configuración. Columna `Datos` (texto sin formato, JSON completo). Historial de versiones activado (500).
- Ingreso: registro de Azure "Hub Comex" (permiso delegado `AllSites.Write`). URIs SPA: la de GitHub Pages y `http://localhost:5500/`.
- Las cuentas que evalúan necesitan permiso **Colaborar** en la lista.

## Confidencialidad (leer antes de producción)

Las reglas de la app (a ciegas, anonimato de las miradas 360, quién ve qué) se aplican **en la app**. Quien tenga permiso de
lectura en la lista de SharePoint podría abrirla directamente y ver todas las filas (la columna `Datos` tiene el JSON completo).
Para un piloto con el nivel estratégico se recomienda:
- lista en un sitio propio, no en el Espacio Colaborativo;
- vista por defecto de la lista sin la columna `Datos` y sin enlaces publicados;
- auditoría con el historial de versiones.

Si se necesita que el servidor haga cumplir las reglas (p. ej. para incluir al nivel operativo o pares anónimos con garantía),
el siguiente paso es un intermediario con permisos de aplicación (como el Worker de hub-averias) que entregue a cada cuenta
solo lo que le corresponde.
