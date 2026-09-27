# 📋 Cambios de la plataforma E&G

Qué se le fue cambiando a la plataforma, lo más nuevo arriba. Cada línea es un
cambio que ya está en vivo: para verlo hay que recargar con **Ctrl + F5**.

> Este archivo lo genera `scripts/generar_cambios.py` desde el historial del
> repositorio. No se escribe a mano: si algo falta aquí, es que no se desplegó.


## 27 de septiembre de 2026

- v295 — Paso 0: la cola ya no re-sube ni pisa, no se pierden cambios con el navegador lleno, aviso al cerrar (SW v295)  
  <sub>`310a4e18c`</sub>

## 26 de septiembre de 2026

- v294 — Centro de costos y Utilidad leen la venta del servidor, no de la copia del navegador (SW v294)  
  <sub>`efe4ab981`</sub>

## 25 de septiembre de 2026

- Cartera: número de factura único en toda la cartera + aviso de factura repetida (SW v293)  
  <sub>`ca64a5ce4`</sub>
- Cotizaciones: botones Seleccionar todo / No seleccionar nada en el modal de adjudicación (SW v292)  
  <sub>`931249852`</sub>
- Remisiones: sello v2 (diseño propio, Montserrat, azul corporativo) y 3 días (SW v291)  
  <sub>`312925b11`</sub>
- Remisiones: sello E&G «No se aceptan devoluciones después de 5 días» al pie del PDF (SW v290)  
  <sub>`7341b109f`</sub>

## 24 de septiembre de 2026

- v289 — la parcial desde la 2ª entrega muestra todo lo entregado hasta hoy  
  <sub>`ac8b0cc4c`</sub>

## 22 de septiembre de 2026

- v288 — una remisión parcial no cuenta las entregas posteriores  
  <sub>`d424e5a16`</sub>
- v287 — Fase 0: agrupar las subidas pesadas a GitHub  
  <sub>`7341ae699`</sub>
- v286 — botón «Traer proveedor y costo de la cotización»  
  <sub>`3ea5a9494`</sub>
- v285 — el plan de compras ya no desaparece sin decir nada  
  <sub>`54aee1160`</sub>
- v284 — carril rápido de guardado (proxy Supabase aparte) + reintentos que no duplican  
  <sub>`7b49c7ba1`</sub>
- v283 — el ping no puede ser más estricto que la escritura que protege  
  <sub>`7377cd668`</sub>
- v282 — el «Supabase 401» era el respaldo, no Supabase  
  <sub>`7800bbbaf`</sub>

## 19 de septiembre de 2026

- OP: no pasa a despachada mientras queden lineas de COMPRA sin despachar (caso OP-2026-0085)  
  <sub>`17bb55d31`</sub>

## 17 de septiembre de 2026

- OP: el boton quitar aparece SIEMPRE, tambien en lineas ya despachadas  
  <sub>`31b6e5525`</sub>
- OP: quitar linea pregunta con modal propio, no con confirm()/prompt()  
  <sub>`b99a63057`</sub>
- OP: poder quitar una linea de la orden  
  <sub>`794282210`</sub>

## 16 de septiembre de 2026

- Cotizaciones: el PDF y el Excel guardan antes de salir para el cliente  
  <sub>`3a23b0a08`</sub>
- Cotizaciones: normalizar el numero escrito a mano y frenar el duplicado  
  <sub>`ae0c10e16`</sub>
- Fix: la OC perdia el autoretenedor del proveedor (MUNDIAL con retefuente 2,5%)  
  <sub>`5fe881f93`</sub>

## 15 de septiembre de 2026

- 🚢 v273 — «Analiza el mes»: cabe en la pantalla y dice qué IMPORTAR  
  <sub>`257b227a5`</sub>
- 📈 v272 — Botón «Analiza el mes»: qué comprar y dónde se pierde plata  
  <sub>`2f09e47a8`</sub>
- 🏷️ v271 — caché energy-v272 para que el módulo Proveedores llegue a los equipos  
  <sub>`26e73c3e9`</sub>
- 🏷️ v271 — Módulo Proveedores: a cómo nos venden y qué les compramos  
  <sub>`8f45df3f2`</sub>
- Cotizaciones: filtros tipo Excel abiertos a todo el equipo  
  <sub>`582dcec75`</sub>
- Cotizaciones: filtros tipo Excel en la tabla de ítems (piloto solo Andrea)  
  <sub>`751c0111b`</sub>
- 🔗 v270 — Ligar no deja pasar un ítem de BODEGA sin producto  
  <sub>`d7ac05a65`</sub>
- Cotizaciones: no pisar GitHub si no se pudo leer + no perder ítems (LM2164, 15-sep)  
  <sub>`4e4896c5c`</sub>

## 14 de septiembre de 2026

- Cotizaciones: alarma anti-duplicado + portero de versión (SB_WRITE_V 130)  
  <sub>`5d0ffaf16`</sub>
- Cotizaciones: evitar ítems duplicados y copias renumeradas (incidente LM2137 14-sep)  
  <sub>`4637ecd29`</sub>
- 🔒 Mario sin acceso a Remisiones ni Órdenes de Pedido (SW v266)  
  <sub>`ff8bfe246`</sub>

## 11 de septiembre de 2026

- OP: lo que se compra no sale de bodega, y el cruce distingue la pieza exacta (v265)  
  <sub>`6942d3810`</sub>
- Definitiva: lo ya entregado solo en el resumen y líneas gemelas unidas (v264)  
  <sub>`03dfd917a`</sub>
- Cuadro de cierre: un renglón que cubre varias líneas gemelas se reparte (v263)  
  <sub>`dfa705b6d`</sub>
- ﻿Remisión definitiva de una sola entrega: sin cuadro repetido, cabe en una hoja (v262)  
  <sub>`626b9346d`</sub>

## 10 de septiembre de 2026

- 🔗 v261 — Guardar una remisión ligada le avisa a su OP lo que cambió  
  <sub>`78dde1132`</sub>

## 9 de septiembre de 2026

- 🔗 v260 — La línea de la OP guarda SIEMPRE a qué producto entró el material  
  <sub>`ab7de119d`</sub>
- 🗑️ v259 — Anular un movimiento sabiendo si tocó el saldo, y la remisión no repite el ítem 1  
  <sub>`78f52878a`</sub>
- 📦 v258 — El kardex no puede sobrevivir al movimiento que cuenta  
  <sub>`de3d2f8da`</sub>
- 🛒 v257 — El plan de compras dice qué líneas deja por fuera  
  <sub>`77f8f8d57`</sub>
- 📋 CAMBIOS.md — la bitácora de la plataforma, generada del historial  
  <sub>`75fdf50d6`</sub>
- 🔄 Botar el caché viejo: energy-v249 → v256  
  <sub>`9ee753d69`</sub>
- 📑 OP: unir varias cotizaciones en una sola orden  
  <sub>`65d892060`</sub>

## 8 de septiembre de 2026

- v255 — Al ligar, una línea en CERO se adueñaba del ítem que había salido en otra remisión  
  <sub>`e369d156c`</sub>
- v254 — Insertar una fila en la mitad de una cotización, sin dejar nada apuntando mal  
  <sub>`0a5802e75`</sub>
- v253 — El cuadro de saldo inventaba un pendiente cuando la orden repite una descripción  
  <sub>`b1b2889fc`</sub>
- v252 — Las entregas parciales que nacen de una OP ya no necesitan marcarse a mano  
  <sub>`b7ec6272e`</sub>
- v251 — Al ligar, un 0 en la remisión tomaba la cantidad de la OP y movía inventario de más  
  <sub>`c510e7358`</sub>
- v250 — El consecutivo de remisión salía de un número quemado que ya existía  
  <sub>`8f432b02c`</sub>

## 7 de septiembre de 2026

- v249 — Entregas parciales: el amarre con la OP dejaba de existir al guardar  
  <sub>`2aca49ee8`</sub>

## 4 de septiembre de 2026

- v249 — Carril "Lista para cerrar": las OP terminadas ya no desaparecen  
  <sub>`090ac4fc3`</sub>
- v248 — Ligar emparejaba corrido cuando la OP nace de adjudicación parcial  
  <sub>`d1178b4ea`</sub>
- v247 — "Sin espacio" con internet inestable: la cola guardaba una copia entera  
  <sub>`35282ca55`</sub>
- v246 — Ligar una remisión ya hecha ahora INGRESA el material a bodega  
  <sub>`320944498`</sub>

## 3 de septiembre de 2026

- v245 — Tablero: carril de Certificados, y la factura se lee de la cotización  
  <sub>`071db538d`</sub>
- v244 — El logo vuelve a salir: se trae del repo, no de WordPress  
  <sub>`c5cb6dab2`</sub>
- v243 — Tablero de OP: la misma lista, en columnas por etapa  
  <sub>`d23d506d5`</sub>
- v242 — Cerrar una OP exige TODOS los certificados, no uno  
  <sub>`1ffa3b07a`</sub>

## 2 de septiembre de 2026

- v241 — En una entrega parcial, el papel imprime solo lo que sale  
  <sub>`1e78e93d6`</sub>

## 1 de septiembre de 2026

- v240 — Entregas parciales y remisión DEFINITIVA  
  <sub>`db5c91c0c`</sub>
- v239 — Historial de OP + filtros por facturar / en camino / atrasadas  
  <sub>`95e86c0d1`</sub>
- v238 — La OP ya no puede nacer vacía (OP-2026-0041 y 0037)  
  <sub>`965039dcb`</sub>
- v237 — El cliente puede volver a pedir sobre la misma cotización: la OP repetida es válida  
  <sub>`54263e7cb`</sub>
- v236 — Crear OP avisa CUÁL OP ya tiene esa cotización, y distingue duplicado de adicional  
  <sub>`8b6f8ccd2`</sub>
- v235 — Ligar remisión ya hecha ahora SÍ descuenta el inventario  
  <sub>`3c83ccdeb`</sub>
- Remisiones: el logo va incrustado, ya no se trae de la página web (SW v234)  
  <sub>`3b2882cc0`</sub>

## 31 de agosto de 2026

- Informe semanal comercial: solo gerencia general (fuera Mario y Sheila)  
  <sub>`27d8f5b20`</sub>
- Cotizaciones: buscador de precios ya cotizados, por cliente (SW v233)  
  <sub>`d03b080c7`</sub>

## 29 de agosto de 2026

- OP: la observacion se escribe DESDE LA LISTA (SW v232)  
  <sub>`ccf410c78`</sub>
- OP certificados: el archivo vive AFUERA, y sin colada no es un pendiente (SW v231)  
  <sub>`bb0d4d588`</sub>
- OP: la colada NO es el MTR — corregido el aviso de la v229 (SW v230)  
  <sub>`937bc60c1`</sub>
- OP: si pasa por nosotros, ingresa a bodega — y no dos remisiones por la misma entrega (SW v229)  
  <sub>`a6c3c3a30`</sub>
- Aviso a gerencia: la memoria pasa del repo a Supabase, y vuelve el respaldo  
  <sub>`47fa40e09`</sub>
- OP: el cruce con inventario elegia la copia VACIA de la pieza (SW v228)  
  <sub>`186758f2b`</sub>

## 28 de agosto de 2026

- Aviso de cotizacion a gerencia: solo al enviarla, no cada 15 minutos  
  <sub>`1cadea5b9`</sub>
- OP: observaciones para decir POR QUE esta detenida (SW v227)  
  <sub>`67d57e33b`</sub>
- OP: "Marcar como enviada" disponible siempre, no solo con todo despachado (SW v226)  
  <sub>`52a22f737`</sub>

## 27 de agosto de 2026

- OP: el envio como paso propio + el plan se entera si algo deja de ser de bodega (SW v225)  
  <sub>`0692eb5fb`</sub>
- Desviacion tecnica: viaja de la cotizacion hasta el cliente (SW v224)  
  <sub>`3bfa8a1fd`</sub>
- OP: Andrea aprueba junto a Alberto (SW v223)  
  <sub>`a792176e4`</sub>

## 26 de agosto de 2026

- OP: color propio para las cerradas (SW v222)  
  <sub>`eaf145588`</sub>

## 25 de agosto de 2026

- OP: Alexandra llega hasta el final, el resto mira; certificados y aviso a gerencia  
  <sub>`f349a22de`</sub>

## 24 de agosto de 2026

- OP y Plan de compras: que los dos digan lo mismo  
  <sub>`63a4d94a1`</sub>
- Cotizaciones: corregir adjudicaciones parciales ya registradas  
  <sub>`135540a22`</sub>
- Cotizaciones: la adjudicacion se guarda sola + Cartera reconoce al cliente  
  <sub>`7808a5588`</sub>
- OP: recibir y despachar en un solo acto (cross-docking)  
  <sub>`ad48b09dd`</sub>
- OP: ligar una remision que ya se habia hecho aparte (caso MONTITEC)  
  <sub>`c250f9eae`</sub>
- OP y Plan de compras: 5 correcciones reportadas por Andrea  
  <sub>`b5885f843`</sub>

## 23 de agosto de 2026

- feat(OP): recibir compras en bodega + ORIGEN del movimiento de inventario (SW v214)  
  <sub>`ba7eefbf1`</sub>
- feat(OP): boton para abrir la remision desde la OP (SW v213)  
  <sub>`6c2bf3e8e`</sub>
- feat(OP): alistar y despachar en UN SOLO ACTO — remision + kardex (SW v212)  
  <sub>`ed752c2e4`</sub>
- feat: modulo Centro de Costos — solo Alberto y Andrea (SW v211)  
  <sub>`299093d96`</sub>
- feat(plan): boton "Centro de costos" — fletes y otros gastos del pedido (SW v210)  
  <sub>`728c8e82e`</sub>
- feat(OP): el plan de compras incluye lo que sale de BODEGA, con costo de inventario (SW v209)  
  <sub>`a3cbd1074`</sub>
- fix(OP): Andrea tambien recibe el aviso de ALISTAR  
  <sub>`f1c667d70`</sub>
- feat(OP): compras ANTES que gerencia + avisos por etapa + el plan avisa a la OP (SW v208)  
  <sub>`4e19f2293`</sub>
- feat(OP): aviso instantaneo, horario ampliado y checklist de certificados (SW v207)  
  <sub>`5f59ad0f2`</sub>
- fix(OP): falsos positivos del cruce con inventario (SW v206)  
  <sub>`faf22ea45`</sub>
- feat(OP): aviso de aprobacion por correo sin depender de Apps Script  
  <sub>`c686751a3`</sub>
- feat(OP): cruce automatico con inventario — deja de comprarse lo que ya hay (SW v205)  
  <sub>`53e1559f1`</sub>
- feat(OP): el aviso de aprobacion sale desde el correo institucional  
  <sub>`08b5cc470`</sub>
- feat(OP): la OP genera y enlaza su Plan de Compras (SW v204)  
  <sub>`5d1cc0d5c`</sub>
- feat(OP): el aviso de Telegram admite varios destinatarios  
  <sub>`e90439a2c`</sub>
- feat(OP): aviso por Telegram al jefe reusando el bot de egresos  
  <sub>`30e3e4a7a`</sub>
- fix(OP): el aviso de aprobacion no le llegaba a nadie + lenguaje llano (SW v203)  
  <sub>`05e3e7d09`</sub>
- fix(OP): 409 por numero duplicado — culpa de reiniciar el consecutivo en pruebas (SW v202)  
  <sub>`fba00c8a6`</sub>
- feat(OP): la via — estado grafico de la orden (SW v201)  
  <sub>`254d42ee1`</sub>
- feat(OP): editar origen/proveedor/costo por linea + descartar y anular (SW v200)  
  <sub>`069890e4c`</sub>
- fix(OP): identidad visual propia + 401 al crear OP + Procesos O.C. sale del menu (SW v199)  
  <sub>`78dc95ebb`</sub>
- feat(OP): modulo Ordenes de Pedido — esqueleto con compuerta de aprobacion (SW v198)  
  <sub>`b2c2a39dc`</sub>

## 22 de agosto de 2026

- fix(OP): la tabla de zonas arranca vacia, no se siembra  
  <sub>`b48c91dcf`</sub>
- feat(OP): SQL del esqueleto de Orden de Pedido (para correr en Supabase)  
  <sub>`c2986225b`</sub>
- chore(correos): retirar a Lina Cifuentes del recordatorio de sabados (ya no labora)  
  <sub>`060ab6d6f`</sub>
- feat(informes): informe comercial SEMANAL a gerencia con copia a vendedores y Andrea  
  <sub>`ec4fbaf9b`</sub>
- feat(cotizaciones): bandeja "Sin cerrar", fecha de rechazo y Excel de analisis completo (SW v197)  
  <sub>`137d292ee`</sub>

## 21 de agosto de 2026

- Fix: la ruta de los cMaps debe ser absoluta  
  <sub>`976f56bf8`</sub>
- Arreglo: los certificados de fabricantes asiaticos ya no se abren en blanco  
  <sub>`72ec99711`</sub>

## 20 de agosto de 2026

- PDF Studio: mostrar la version en el pie para poder verificar la cache  
  <sub>`c0ce507a8`</sub>
- PDF Studio v194: lista de proveedores como chips, con los frecuentes ya puestos  
  <sub>`6e1d24cce`</sub>
- PDF Studio v193: quita la imagen de fondo del proveedor sin tocar el texto  
  <sub>`c471b169b`</sub>
- PDF Studio v192: limpieza automatica de certificados (proveedor + sello + fondo)  
  <sub>`f79e5a201`</sub>
- SW: caché energy-v191 para que el equipo tome la colada en las entradas  
  <sub>`acb5ec13f`</sub>
- Abastecimiento v191: la ENTRADA también registra la COLADA  
  <sub>`56d13c23a`</sub>

## 19 de agosto de 2026

- Remisiones v190: PDF con membrete ASSET + planilla de viaje a nombre de Asset  
  <sub>`a2a2e8ab4`</sub>
- Rotación v189: los productos archivados vuelven al análisis (y el kardex ya no se corta  
  <sub>`18834e042`</sub>
- Inventario: archivados 6 códigos repetidos que estaban en cero  
  <sub>`22659332c`</sub>
- Inventario: corregidos 48 errores de escritura en nombres y marcas  
  <sub>`aec73d89f`</sub>
- Abastecimiento v188: aviso al crear un producto que ya existe  
  <sub>`b28824c3b`</sub>
- Plan de Compras v187: marcar COMPRADO por proveedor (para los que no llevan OC)  
  <sub>`0e8632427`</sub>
- Plan de Compras v186: check de compra por línea + planes con varias cotizaciones  
  <sub>`e7b0548ed`</sub>

## 17 de agosto de 2026

- SW: caché energy-v185 para que los equipos tomen la planilla de viaje  
  <sub>`12f26ab93`</sub>
- Planilla de transporte de mercancía por viaje (v185)  
  <sub>`e071a7484`</sub>

## 14 de agosto de 2026

- Adjudicación v184: el modal "¿Qué le adjudicaron?" muestra el N° de ítem  
  <sub>`07837ac99`</sub>
- Remisiones v183: el N° de ítem de la cotización viaja a la remisión  
  <sub>`f3e31dce2`</sub>

## 13 de agosto de 2026

- informe comercial: el valor adjudicado nunca llegaba a los KPIs  
  <sub>`e0f87be0e`</sub>

## 12 de agosto de 2026

- v182: arregla el modal de adjudicacion + desviacion tecnica al plan y a la OC  
  <sub>`65fc92e61`</sub>

## 10 de agosto de 2026

- Retiro de Lina Cifuentes (ya no labora en la empresa)  
  <sub>`7a8d47acc`</sub>
- Cotizaciones: casilla Obs. proveedor sale en blanco  
  <sub>`c085b8eaf`</sub>
- Respaldo: agregar cartera, compras y proveedores (7 tablas sin cubrir)  
  <sub>`c5c36b470`</sub>

## 8 de agosto de 2026

- Suspender informes automaticos: solo queda el Resumen de Pagos 6PM  
  <sub>`537a6d618`</sub>

## 6 de agosto de 2026

- 🔒 Tareas: cada quien ve las suyas; solo Andrea y Alberto ven todas  
  <sub>`6fafd3ceb`</sub>
- 📋 Remisiones: Lista de Verificación de Despacho (SIG-CAL-FR-009)  
  <sub>`d6e8744ef`</sub>

## 4 de agosto de 2026

- 🗑️ Abastecimiento: borrar producto (o archivarlo si ya tiene historial)  
  <sub>`31c92180d`</sub>
- 📦 Abastecimiento: el origen (IMP o PZ) es obligatorio, se quita "nacional"  
  <sub>`0c97126be`</sub>
- 📦 Abastecimiento: editar producto + consecutivo automático por familia  
  <sub>`8cbde4b05`</sub>
- 📷 Fotos a IndexedDB: el localStorage deja de llenarse  
  <sub>`09f14e5ee`</sub>
- 🛑 Cuota llena: ya no se pierden cambios + auto-liberación de cachés  
  <sub>`c15de6267`</sub>

## 28 de julio de 2026

- Confirmacion visual antes de registrar una SALIDA del carrito (v173)  
  <sub>`e22f890e1`</sub>
- Ajuste de inventario deja rastro en Kardex (v172)  
  <sub>`9cf2d6ea1`</sub>
- Cerrar Mes ya no borra ordenes: las archiva (v172)  
  <sub>`5c6b562f8`</sub>
- Fix O.C.: el merge ya no pierde avance (v172)  
  <sub>`b5b14a31a`</sub>

## 27 de julio de 2026

- Sync: timeout adaptativo al subir a GitHub por proxy (v171)  
  <sub>`3f844a261`</sub>

## 26 de julio de 2026

- O.C.: busqueda GLOBAL (todas incl. archivadas) + estado en tabla (v170)  
  <sub>`b75a88f52`</sub>
- O.C.: buscador por N/cliente/descripcion/observacion (v169)  
  <sub>`0d04f23bc`</sub>
- O.C.: filtro Sin facturar muestra TODAS las pendientes de facturacion (v168)  
  <sub>`06a4a1dee`</sub>
- O.C.: clic en la orden muestra sus items para identificarla (v167)  
  <sub>`e601aeab8`</sub>
- O.C.: estado de facturacion + observaciones + errores visibles (v166)  
  <sub>`d22abacc6`</sub>
- Abastecimiento: comprobante tambien para ENTRADAS (v165)  
  <sub>`861773d40`</sub>

## 24 de julio de 2026

- Abastecimiento: hoja de alistamiento imprimible para bodega (v164)  
  <sub>`215543f25`</sub>
- PDF Studio — Sesion 3: formularios (rellenar + crear campos) + fix nitidez (v163)  
  <sub>`568aec8db`</sub>

## 23 de julio de 2026

- PDF Studio — Sesion 4: tapar datos (redaccion real) y quitar marcas de agua (v162)  
  <sub>`90a2b8fd6`</sub>
- Cartera: base/IVA reales de la factura + RETEFUENTE (saldo neto) (v161)  
  <sub>`128f1d8d7`</sub>
- Procesos O.C.: completado automatico, archivado y foco en lo pendiente (v160)  
  <sub>`5c7a78eef`</sub>
- Remisiones: orden y filtro por N de remision (v159)  
  <sub>`5eedac9de`</sub>
- Remisiones: blindaje anti-perdida tras incidente 26215 INDUMEC (v158)  
  <sub>`98f505208`</sub>
- PDF Studio: vista EN VIVO — lo que ves es lo que se descarga (v157)  
  <sub>`5a81348a9`</sub>
- PDF Studio E&G — Sesion 2: editar (texto, firma, sellos, marca, membrete) (v156)  
  <sub>`b5741c201`</sub>
- PDF Studio E&G — Sesion 1: organizar paginas + version PC (v155)  
  <sub>`cf6ef673d`</sub>
- Cartera: Nelsy solo consulta + Excel foto al corte de fecha (v154)  
  <sub>`c5cfc6d93`</sub>
- BD cotizaciones: busqueda "cliente + item" (v153)  
  <sub>`a0e97e2c5`</sub>
- BD cotizaciones: cadena documental completa al Facturar (v152)  
  <sub>`93039b07d`</sub>
- Cotizaciones: alternativas por item (3 -> 3.1) que NO suman al total (v151)  
  <sub>`61a693ce6`</sub>

## 22 de julio de 2026

- Cartera: modal de historial 📚 a pantalla casi completa (96vw) para ver toda la tabla  
  <sub>`bc54eefaa`</sub>
- Cartera: editar/anular/borrar facturas históricas (solo andrea) con nota automática; ANULADAS fuera de todos los totales  
  <sub>`6e3a84396`</sub>

## 21 de julio de 2026

- Cartera: pestaña Pagos con filtro por año/mes, buscador, total recibido y export Excel  
  <sub>`3730029e3`</sub>
- Cartera F4: historial completo por cliente, correos de cobro con Claude (copiar), cuenta de cobro con consecutivo compartido, e informe 7PM leyendo Supabase  
  <sub>`f6721ed58`</sub>
- Cartera: pago en lote con búsqueda de combinación exacta y ReteIVA, columnas con/sin IVA, y auto-completar OC al facturar desde cotización  
  <sub>`ae2d2057e`</sub>
- Cartera en la Nube (F3): nueva factura con autollenado, pagos/abonos, BD clientes, notas y amarre cotización→OC→factura  
  <sub>`c0afb3e30`</sub>
- Cartera en la Nube (F2): módulo lee Supabase — semáforo de mora, comisiones y meses sin depender del PC  
  <sub>`7d583f914`</sub>

## 18 de julio de 2026

- Kanban personal: cada quien sus tareas, dirección ve todo con colores (v150)  
  <sub>`178cbdb25`</sub>
- Oferta Técnica: varios adjuntos por ítem — PDF + fotos juntos (v149)  
  <sub>`3700d75f1`</sub>
- Oferta Técnica: fichas como imagen (JPG/PNG) + arrastrar y soltar (v148)  
  <sub>`fc22d6d34`</sub>
- Oferta Técnica: solo información técnica en specs IA (v147)  
  <sub>`a95f786b7`</sub>
- Oferta Técnica: colapsar renglones en blanco de las specs (v146)  
  <sub>`9f240aa22`</sub>
- Oferta Técnica: diseño denso + specs IA exhaustivas (v145)  
  <sub>`0381ce437`</sub>
- Oferta Técnica en cotizaciones: PDF membrete + fichas anexas + specs IA (v144)  
  <sub>`deb091255`</sub>
- Cotización PDF: columna IMAGEN centrada + descripción en mayúsculas (v143)  
  <sub>`c3ec1edbc`</sub>
- Tarjeta sábado dashboard lee Programación Equipo + sync rotación (v142)  
  <sub>`5c11f431e`</sub>

## 17 de julio de 2026

- 🏷️ Remisiones: rótulos rediseñados — media carta, 2 por hoja, ítems con cantidad (diseño skill rotulos)  
  <sub>`3ac4d27e4`</sub>

## 16 de julio de 2026

- Crear remision desde la BD, como Crear OC (v141)  
  <sub>`48f27832f`</sub>
- Cotizaciones: puente Adjudicada-Remision + vencimiento automatico (v140)  
  <sub>`f76911d64`</sub>
- Busqueda inteligente: medidas EXACTAS (v139)  
  <sub>`d3da37026`</sub>
- Abastecimiento: busqueda inteligente en los 3 buscadores (v138)  
  <sub>`c5e7944c3`</sub>
- Abastecimiento: aviso anti-doble remision + cruce remision-bodega (v137)  
  <sub>`cfd9f0f60`</sub>

## 15 de julio de 2026

- Cartera: vista Facturado/Pagado por mes en plataforma (v136)  
  <sub>`3414819f7`</sub>
- Candado anti doble-clic al guardar cotizacion (v135)  
  <sub>`bac3e81dc`</sub>

## 14 de julio de 2026

- Consecutivo nace al GUARDAR (v134): numeracion unica, correlativa, sin huecos  
  <sub>`ac4611b0a`</sub>

## 13 de julio de 2026

- Abastecimiento: botón '📋 Descarga para conteo' — Excel marca EYG para inventario físico ciego, agrupado por familia, respeta filtros (SW v133)  
  <sub>`c940ee7b2`</sub>
- Presencia se suelta al salir del editor: vigilante 2s + visibilitychange (SW v132)  
  <sub>`2394fb313`</sub>
- Presencia mejorada: burbuja grande + fila resaltada en dorado con nombre de quien está en cada ítem (SW v131)  
  <sub>`0d2129fdd`</sub>
- Fix arranque Realtime: currentUser es let global, window.currentUser nunca existía y _rtBoot esperaba eternamente (SW v130)  
  <sub>`b91a9d6d0`</sub>
- Migración cotizaciones→Supabase: escrituras vía proxy (RLS), Supabase manda al abrir, colaboración por ítem + Realtime/presencia (SW v129)  
  <sub>`338dbce1b`</sub>

## 12 de julio de 2026

- Cartera: boton 'Descargar libro completo' (Excel tal cual) protegido con clave, para Andrea/Alberto/Nelsy (SW v128)  
  <sub>`40da8354f`</sub>

## 10 de julio de 2026

- Prevencion cotizaciones: choque de numero = SOLO AVISAR (no renumerar solo) + banner 'nueva version, recargar' (no auto-recarga) (SW v127)  
  <sub>`4f3fd3029`</sub>
- Blindaje cotizaciones: no reclamar/renumerar al EDITAR una guardada (fix rebote LM1790-1 -> -4) (SW v126)  
  <sub>`98a557e1c`</sub>
- Cotizaciones: foto grande por item (pantalla+PDF). Remisiones: contenido por caja en rotulos (CAJA X de N) guardado en base companera (SW v125)  
  <sub>`3bdff90ee`</sub>

## 9 de julio de 2026

- Visitas: nueva base Actas de visita (bitacora) ligada al semanario por consecutivo VIS-2026-NNNN (SW v124)  
  <sub>`61646278e`</sub>
- Plan de Compras: base retefuente compras 27->10 UVT ($523.740), vigente 1-jul-2026 (Decreto 0572/2025) (SW v123)  
  <sub>`699e653a3`</sub>
- Plan de Compras: boton Utilidad (venta cotiz vs costo plan) solo gerencia+Andrea, con transporte (SW v122)  
  <sub>`fb88a4394`</sub>
- Cartera: dar acceso de visualización a Nelsy (SW v121)  
  <sub>`cf1666164`</sub>

## 8 de julio de 2026

- Informe de pagos: cron a las 23:17 UTC (minuto no en punto) para que GitHub no lo salte por congestión de la hora exacta  
  <sub>`89f233868`</sub>
- Cotizaciones: pestaña Clientes visible SOLO para gerencia (Alberto) y Andrea — oculta el tab + candado en cotizShowTab para el resto + SW v120  
  <sub>`7b2204077`</sub>
- Programación Equipo: agregar a Sheila (híbrido) al roster + SW v119  
  <sub>`35a4d28fe`</sub>
- Visitas: una visita puede tener VARIOS comerciales (casillas en vez de selector único). Aparece en la fila de cada comercial en el semanario con marca 👥 compartida; editar/borrar afecta a todos. Backward-compatible (visitas viejas usan comercial único) + SW v118  
  <sub>`055581d6d`</sub>
- FIX visitas se borran solas: guardar/borrar ahora une remoto + LOCAL + el cambio (antes una visita recién creada se perdía al guardar/borrar otra porque el remoto baja con retraso y no la incluía) + SW v117  
  <sub>`a853e32df`</sub>
- Semanario visitas: quitado el campo 'Cliente' del formulario Programar visita (queda 'Empresa/Cliente' como dato principal); JS blindado para no romper con el campo ausente; visitas viejas conservan su dato + SW v116  
  <sub>`5bd4396c1`</sub>
- Visitas: borrado con lápida (deleted:true) para que NO revivan al sincronizar + filtrar borradas en semanario y export + nombres completos en el semanario (Mario Rodríguez/Sandra Sánchez, solo etiqueta, sin tocar usuarios/equipo) + SW v115  
  <sub>`9381b9ca7`</sub>
- Remisiones anti-borrado: el aviso de 'ya existe el número' ahora consulta el SERVIDOR (no la memoria local, que estaba vieja y por eso SERSOLCA pisó a QCIEN sin avisar) + SW v114  
  <sub>`43f6765d2`</sub>
- Plan de Compras anti-borrado: avisa (consultando el servidor) antes de reemplazar un plan que YA existe con ese CC — antes se sobrescribía en silencio si dos personas hacían plan de la misma cotización. Igual que Remisiones + SW v113  
  <sub>`284c6232d`</sub>
- Semanario visitas: cuando hay 2+ visitas el mismo día se ven CLARAMENTE separadas (borde completo, más margen, numeradas 1/2·2/2) con la hora en badge visible, ordenadas por hora — antes se veían pegadas y parecían una sola + SW v112  
  <sub>`31b1c3df7`</sub>
- Programación Equipo: nuevo estado 🚗 Visita de campo (ciclo Presencial→Remoto→Visita de campo→Vacaciones→Descanso), en resumen mensual, Excel y leyenda + SW v111  
  <sub>`6ec3dd339`</sub>
- Blindaje anti-duplicados en Remisiones y Plan de Compras: al leer, si un corte de red dejó líneas repetidas (mismo item, distinto id) se muestra solo la más reciente (_dedupLineas). Aditivo, no toca el guardado + SW v110  
  <sub>`be637774b`</sub>
- Indicador de versión visible en la barra lateral (lee la versión real instalada; toca para actualizar) — para detectar equipos con código viejo + SW v109  
  <sub>`f5447ce6f`</sub>
- FIX CRÍTICO pérdida de precios: _cotizReconciliar prefería SIEMPRE la versión renumerada (renumeradaAt de v105) sobre cualquier edición más reciente → al poner precios a una cotización renumerada, la versión vieja en $0 ganaba y los borraba. Ahora gana updatedAt más reciente; renumeradaAt solo desempata + SW v108  
  <sub>`69547cb53`</sub>
- Fix Base de Datos vacía para Sandra/Mario: todo el equipo ve TODAS las cotizaciones (el filtro .includes('comercial') atrapaba a 'Asistente Comercial' y 'Coordinador Comercial' y les mostraba la BD vacía) + SW v107  
  <sub>`f223573bc`</sub>

## 7 de julio de 2026

- Sostenibilidad sesión 1: guardado sin pérdida en Remisiones y Plan de Compras (insertar→borrar viejas, adiós DELETE+POST), aviso global de cuota localStorage llena, e informe 🩺 Salud del Sistema (mensual día 1 + chequeo lunes solo-si-alertas: backups verificados, tamaños con tendencia, Supabase, robots fallidos) + SW v106  
  <sub>`f7380e0c9`</sub>
- Paso A (fase intermedia): CONSECUTIVO DEL SERVIDOR — reclamo atómico del número en Supabase al guardar cotización nueva (INSERT sobre PK; 409 → siguiente número + form actualizado antes del PDF; offline → flujo local) + backfill completo verificado + SW v105  
  <sub>`e39160e5c`</sub>
- Programación Equipo: botón Excel solo para dirección + SW v104  
  <sub>`088e6e8f2`</sub>
- Programación Equipo: todo el equipo VE (nav abierta), solo dirección EDITA (peIsManager en celdas, semana estándar y roster; tab Equipo oculto para no-dirección) + SW v103  
  <sub>`111b29f57`</sub>
- FIX layout: eliminado </div> sobrante que cerraba .content tras Configuración — Inventario/Visitas/Programación/Mensajería/Abastecimiento/Remisiones/Compras vuelven al contenedor (desaparece el espacio en blanco gigante del semanario y recuperan el padding) + SW v102  
  <sub>`c318d3f80`</sub>
- Cola de sync: reintento automático cada 60s (falla puntual del proxy ya no requiere acción del usuario) + SW v101  
  <sub>`598832ff1`</sub>
- Nuevo módulo Programación Equipo (solo gerencia/Andrea/Mario): semana Lun-Sáb con estados por persona (presencial/remoto/vacaciones), roster editable presenciales/híbridos, semana estándar 1-clic, resumen mensual + Excel E&G, y pestaña Sábados movida aquí desde Visitas + SW v100  
  <sub>`645d35788`</sub>
- Chat: reintento automático si el proxy falla (degrada a Haiku) + mensaje de error claro + SW v99  
  <sub>`922b488e5`</sub>
- Chat ENERGY asesor técnico: mecánico/eléctrico/instrumentación con modo experto automático (Sonnet solo en preguntas técnicas), sugerencia de inventario E&G y aviso de seguridad en temas críticos + SW v98  
  <sub>`5232310b9`</sub>
- Arreglos rotos del análisis: actividad reciente REAL + tarjeta Alertas activas (adiós al 7 fijo), calendario reuniones navega meses, alerta cotiz sin seguimiento revivida (clave correcta, resumen), Teams chat real en Equipo (sin presencia falsa), WhatsApp equipo copia al grupo, nombres chat ENERGY corregidos, fuera schedule-reports.yml duplicado + SW v97  
  <sub>`5b9f32374`</sub>
- Clientes: fusionados ARROW y METAL (confirmado por usuario) + emparejador umbral 5 letras + SW v96  
  <sub>`7675e824b`</sub>
- Clientes: botón 📗 Excel con marca E&G, emparejador normalizado anti-duplicados (_cliNorm/_cliMismo) y limpieza de 6 fusiones (59→51) + SW v95  
  <sub>`8d7d1b073`</sub>
- Informe de pagos: remitente fijo info@eygenergygroup.com (no desde el correo de Andrea)  
  <sub>`6c7792cb0`</sub>
- Informe de Pagos 6PM migrado a la nube: GitHub Actions + Graph (OneDrive Andrea), reporta desde el último informe enviado, envía siempre (sin pagos = aviso)  
  <sub>`4a810a15c`</sub>
- Bloque A cotizaciones: BD clientes auto-alimentada (upsert al guardar + seed 59 clientes), fechaEnvio revive el semáforo, timeouts de red 10-20s (fix Sincronizando eterno), auto-marcar Vencidas + SW v94  
  <sub>`1a9acb975`</sub>
- Botón 🔄 Actualizar global: sube cola + fuerza descarga de todos los datos + 'hace Xs' + detector de versión nueva con oferta de recarga + SW v93  
  <sub>`20f4a97df`</sub>
- Cotizaciones: motivo de rechazo obligatorio (desplegable 5 motivos al marcar Rechazada, visible en BD) + ítems en $0 salen como NO COTIZADO en PDF y Excel + SW v92  
  <sub>`d2bc262c7`</sub>
- Abrir cotización sin colgarse: verificación con tope 2,5s (internet flojo abre con copia local + aviso), toast 'Abriendo…' y mensaje claro si no se encuentra + SW v91  
  <sub>`e94bba0a7`</sub>
- SW v90  
  <sub>`e09afc5b3`</sub>
- Consecutivo vivo: el formulario detecta cuando otra persona usa el mismo número (vigilante en poll 8s) y actualiza campo+modo edición al renumerar + SW v90  
  <sub>`ebebfed96`</sub>
- Extraer con IA ya no borra ítems existentes: pregunta agregar/reemplazar + botón deshacer (↩️ Recuperar ítems anteriores) + SW v89  
  <sub>`85c4f69ed`</sub>
- Fix pisadas al reconectar: la cola offline re-mezcla con el remoto antes de subir (cotizaciones/solicitudes/semanario/sábados) + poll inmediato al volver la red + SW v88  
  <sub>`111172cad`</sub>
- Sábados fase 2: tarjeta Dashboard 'Este sábado trabaja' + recordatorio viernes 4PM por correo + SW v87  
  <sub>`83102b1ed`</sub>
- Sábados: cronograma rotativo de asistentes (Alexandra/Lina/Sandra) en Visitas Comerciales + SW v86  
  <sub>`0540d7577`</sub>

## 6 de julio de 2026

- Semanario: mover modal fuera de page-visitas (overlay fijo) + SW v85  
  <sub>`af95fbc75`</sub>
- Semanario: modal con pie fijo (Guardar siempre visible) + SW v84  
  <sub>`04e1095d1`</sub>
- Semanario: exportar visitas a calendario (.ics) para Outlook/Teams + SW v83  
  <sub>`1b648fc25`</sub>
- Visitas: nuevo Semanario (programacion semanal de visitas) + bump SW v82  
  <sub>`c1b17732e`</sub>
- Remisiones: permitir editar N° para sufijos (ej. 1741-1) + bump SW v81  
  <sub>`7ee6af696`</sub>

## 5 de julio de 2026

- Migracion automatica del umbral de flete viejo (5.000.000 -> 500.000)  
  <sub>`d29eda68f`</sub>
- Observaciones por defecto completas + umbral flete a $500.000  
  <sub>`7646c32d3`</sub>
- Flete interno del pedido: reparte el flete en los precios (reversible)  
  <sub>`d7fb0d5b7`</sub>
- Calculadora: marca en el buscador + stock disponible + aviso de cantidad  
  <sub>`488e4893d`</sub>
- Calculadora: muestra y arrastra la MARCA del item (ademas del proveedor)  
  <sub>`8a9ccf61e`</sub>
- Calculadora: agrega proveedor (auto del inventario) + cantidad -> total  
  <sub>`fc59740c9`</sub>
- Calculadora utilidad+flete en Cotizaciones + Inventario lee de Abastecimiento  
  <sub>`87b4577b1`</sub>

## 4 de julio de 2026

- Sync: refresco más rápido (polling en paralelo + al abrir módulo + 10s→8s)  
  <sub>`155b090a3`</sub>
- Cotizaciones: generar PDF sin ventana emergente (arregla "no pasa nada" del jefe)  
  <sub>`bf98e0423`</sub>
- Cotizaciones: anti-colisión de consecutivo + propagación más rápida + fix toggle autoretenedor  
  <sub>`99ee44f07`</sub>

## 3 de julio de 2026

- Plan de Compras: campo Observaciones en OC + panel diagnóstico de sincronización  
  <sub>`7cce4db16`</sub>

## 2 de julio de 2026

- Usuario Mario: corregir correo a director.comercial@eygenergygroup.com  
  <sub>`47dc30e35`</sub>
- Usuarios: crear Mario (Coord. Comercial) y Sandra (Asist. Comercial) + claves  
  <sub>`91145e9d4`</sub>
- Procesos O.C.: ligar/cambiar a mano la cotizacion de una OC existente  
  <sub>`f7dff37e7`</sub>
- Cotizacion Facturada: auto-guardar al marcar (no depende de "Guardar")  
  <sub>`6dc1b15d7`</sub>
- Cotizacion Facturada: preguntar N° factura al cambiar estado y escribir en la OC al instante  
  <sub>`835a77a16`</sub>
- Procesos O.C.: mostrar N° de factura en la lista y en el informe imprimible  
  <sub>`cad42cfbc`</sub>
- Plan de Compras: Excel identico al PDF + cotizacion Facturada propaga a la OC  
  <sub>`63b4ced0a`</sub>
- Cotizaciones: boton Analisis por cliente (solo Andrea) - adjudicadas, cotizado vs adjudicado, items caidos. SW v61  
  <sub>`d1d412a3d`</sub>
- Cotizaciones: boton Comparar (cotizado vs adjudicado/Plan de Compras) con % adjudicado e items caidos. SW v60  
  <sub>`e4d965b26`</sub>

## 1 de julio de 2026

- Cotizaciones: mostrar quien edito (updatedBy) en la BD + aviso anti-pisada si otro la edito hace <10 min. SW v59  
  <sub>`65a39cc22`</sub>
- Cotizaciones: observacion extensa (condiciones comerciales completas) como default en cotizacion nueva. SW v58  
  <sub>`4a1d3b53d`</sub>
- Cotizaciones: ocultar tambien MARGEN a no-gerencia (solo Alberto/Andrea ven FACTOR y MARGEN). SW v57  
  <sub>`37773e97c`</sub>
- Cotizaciones: boton "Quitar imagen" y que no se arrastre a otras  
  <sub>`581fcf4d4`</sub>
- email-to-cotiz.gs: max_tokens 8000->16000 (solicitud 79 items, la salida topaba en ~66)  
  <sub>`a43ad4f6d`</sub>
- email-to-cotiz.gs: subir max_tokens 2000->8000 y limite Excel 5000->40000 (solicitud de 79 items)  
  <sub>`587fa3ff5`</sub>
- Solicitudes: evitar cotizaciones duplicadas (doble candado)  
  <sub>`0e65bdef4`</sub>

## 30 de junio de 2026

- Trazabilidad cotizacion <-> remision (columna cotizacion_id en Supabase)  
  <sub>`1a41a379f`</sub>
- Trazabilidad: desde la OC, ver la cotizacion ligada (local o LIBRO). SW v53  
  <sub>`d08c0ccba`</sub>
- Cotizaciones: boton Nueva pregunta Guardar/Descartar/Cancelar si hay cotizacion sin guardar (SW v52)  
  <sub>`ea03cbdbf`</sub>
- Informe: adjudicadas transparentes (cuenta todas las ganadas por adjudicadaAt o fecha de cotizacion)  
  <sub>`03ef66b31`</sub>
- Informe comercial: bloque Cotizado/Adjudicado/Facturado (hoy y mes)  
  <sub>`0670d6b72`</sub>
- Cotizaciones <-> Procesos OC: enlace de vuelta  
  <sub>`481ce8836`</sub>
- Cotizaciones: crear Orden de Compra (Procesos OC) desde una adjudicada  
  <sub>`d495042a8`</sub>
- Plan de Compras: editar si un proveedor retiene o es autoretenedor  
  <sub>`319678b4e`</sub>
- Extractor LIBRO: inferir fecha de cotizaciones sin fecha de generacion  
  <sub>`140f90848`</sub>

## 29 de junio de 2026

- Cotizaciones: columna FACTOR (gerencia) que calcula V. Unitario desde el costo  
  <sub>`be32f6634`</sub>
- Etapa 2: informe comercial lee cotizaciones de Supabase (fuente principal)  
  <sub>`db3a09f57`</sub>
- Backup diario: incluir cotizaciones y cotizacion_items desde Supabase  
  <sub>`4cf8d403c`</sub>
- Cotizaciones: doble escritura a Supabase (Etapa 1 de "todo en Supabase")  
  <sub>`40680c450`</sub>
- Remisiones "Traer de cotizacion": mostrar TODAS (incluye LIBRO/Supabase)  
  <sub>`02b8ceb9f`</sub>
- Cotizaciones BD: ordenar por columna + filtro "Realizo"  
  <sub>`11bb4bfe9`</sub>
- Cotizaciones del LIBRO: editables (adoptar al editar)  
  <sub>`ccfc81ccb`</sub>
- Cotizaciones BD: mostrar historico del LIBRO (Supabase) junto a las locales  
  <sub>`d46b331f1`</sub>

## 22 de junio de 2026

- Fix: cotizaciones largas se cortaban a la mitad al extraer con IA  
  <sub>`5eaebabc8`</sub>
- Agrega botones 'BD completa' en Cotizaciones y Remisiones (solo Andrea/Alberto)  
  <sub>`50d29a91f`</sub>

## 19 de junio de 2026

- fix(cartera): abrir pestaña antes del await para no bloquear el popup  
  <sub>`787985d9e`</sub>
- feat(OC): datos del proveedor (NIT/dir/ciudad/tel/contacto) en la OC, se guardan y autocompletan  
  <sub>`48b6c27e2`</sub>

## 18 de junio de 2026

- feat(abastecimiento): boton X para cerrar el aviso de stock critico  
  <sub>`68f87ca33`</sub>
- fix(robots): OC dispara solo con 'ORDEN DE COMPRA' (tolera typos como ENEEGY); cotiz IGNORA 'ORDEN DE COMPRA'  
  <sub>`cf10062b8`</sub>

## 17 de junio de 2026

- fix(email-to-oc): dedup por mensaje (no por hilo) para órdenes con asunto repetido  
  <sub>`7a6f4dd49`</sub>
- fix(email-to-oc): re-chequeo de asunto en una sola línea (evita ReferenceError subjU)  
  <sub>`dddac2c58`</sub>
- fix(email-to-oc): aceptar asuntos con cliente en medio (ORDEN DE COMPRA CIAM ENERGY GROUP)  
  <sub>`54ed989b7`</sub>

## 16 de junio de 2026

- Solicitudes: boton "Actualizar" para traer solicitudes nuevas del correo  
  <sub>`7461cb4a3`</sub>
- Abastecimiento: botones crear producto y familia  
  <sub>`f80576464`</sub>

## 12 de junio de 2026

- fix: en Solicitudes, ocultar Cotizar si ya tiene cotizacion y mostrar Ver cotizacion  
  <sub>`dca13f635`</sub>
- feat: Excel en Solicitudes (filtra por estado) + estado Para revision gerencia + extraccion literal de correos  
  <sub>`2f68a9c8f`</sub>

## 11 de junio de 2026

- Cotizaciones: botón Exportar a Excel para análisis (rango de fechas)  
  <sub>`c848d1c28`</sub>
- Bot WhatsApp: recibe fotos/PDF y crea solicitudes de cotización  
  <sub>`b9b70bb8b`</sub>
- Cotizaciones: desviación técnica como textarea que crece (ver texto completo)  
  <sub>`540918d53`</sub>
- Fix ordenes.json atascado: caer al proxy cuando el PAT falla (401/403)  
  <sub>`a2d490cac`</sub>
- Fix Email-to-OC: la llamada interna del trigger Gmail pasa el token  
  <sub>`77d632ed3`</sub>
- Plan logístico: sede de recogida editable por proveedor  
  <sub>`e51d93efb`</sub>
- Cotizaciones: BD muestra 'Realizó' y estado inicial Borrador  
  <sub>`7113ecf34`</sub>
- email-to-cotiz: sincronizar con versión real (ENERGY Cotiz) + fix no-perder-solicitudes  
  <sub>`b2cee4226`</sub>
- Cotizaciones: limpiar el formulario al guardar una cotización nueva  
  <sub>`8afda034d`</sub>

## 10 de junio de 2026

- Proteger endpoint Email-to-OC con token (evita OCs falsas)  
  <sub>`2264ed3de`</sub>
- Seguridad bot WhatsApp + no perder solicitudes de cotización  
  <sub>`616706d06`</sub>
- Proteger proxy: token + rate-limit + tope de modelo  
  <sub>`a8a983734`</sub>
- Crear workflow de backup automático de Supabase (faltaba)  
  <sub>`3692e99bc`</sub>
- Informe comercial: conversión = ganadas/emitidas (excluye borradores)  
  <sub>`6d533dcf3`</sub>
- email-to-cotiz: anti-duplicados — no crear solicitud si ya hay una del mismo cliente+descripcion en <24h  
  <sub>`f89e53abb`</sub>
- Solicitudes/sesion: #1 merges leen fresco de raw + borrado pegajoso (no resucitar); #2 policia (no cotizar 2 veces) + asigna cotizacionId; sesion del dia no re-pide clave al recargar  
  <sub>`819a11f1a`</sub>
- Service worker: pedir HTML/JS propios siempre frescos (cache:no-store) + bump v37 — los despliegues ahora aparecen al recargar sin quedar pegados en cache  
  <sub>`fe042376a`</sub>
- Solicitudes: ocultar las canceladas de todas las vistas (se conservan como tombstone para que el sync no las resucite)  
  <sub>`301918757`</sub>
- Cotizacion: descripcion en textarea que crece (renglon completo) + foto por item (camara, comprimida, guardada en archivo aparte, casilla para incluir en PDF cliente)  
  <sub>`05700f156`</sub>
- Sync mas rapido y confiable: refresco 30s->10s + deteccion de cambios por hash completo (evita que se pisen cambios por updates no detectados)  
  <sub>`58e2d37e4`</sub>
- Fix borrar cotizaciones: borrado logico (deleted flag) para que la union con el servidor no las resucite; ocultarlas en BD, agente, remision y buscador  
  <sub>`dcee8f310`</sub>
- Fix token vencido: si el PAT falla, caer automaticamente al proxy compartido (no bloquear al equipo)  
  <sub>`49d5909b7`</sub>
- Informe comercial: adelantar a 5:13 PM Colombia (minuto impar) para reducir retraso de GitHub Actions  
  <sub>`ae8b66e26`</sub>

## 9 de junio de 2026

- Cotizaciones: PDF imprime colores y margenes correctos (print-color-adjust + @page margin 0)  
  <sub>`116ac7a06`</sub>

## 7 de junio de 2026

- feat: boton Ayuda (manual integrado segun usuario) + clave una sola vez al dia por equipo  
  <sub>`1e30e0243`</sub>
- fix: blindaje anti-pisadas para cotizaciones.json (merge por id con remoto antes de guardar, como solicitudes)  
  <sub>`c9318a234`</sub>
- feat: informe lee tambien cotizaciones archivadas por Cerrar Mes (data/historico)  
  <sub>`a9abadcfc`</sub>
- feat: informe gerencial — semanal desde mayo (cotizado/facturado/recaudado del libro de cartera), cartera vencida, top deudores, vendedores, comparativa mensual + gerenciageneral como destinatario  
  <sub>`c3dbcc646`</sub>
- fix: PDF cotizacion sin VENDEDOR/ELABORO/APROBO de cara al cliente (siguen guardados internamente)  
  <sub>`ebec0fd50`</sub>
- feat: informe comercial diario 7PM a Andrea + Datos_EYG.xlsx (historico 445 cotizaciones, conversion, top clientes, margen)  
  <sub>`8e6e10696`</sub>
- feat: plan de compras — marcar proveedor como No requiere OC (columna sin_oc)  
  <sub>`37685e7c2`</sub>
- feat: modulo Plan de Compras (Supabase) — desde cotizacion, dividir proveedores, OC E6-FC-01 y ruta logistica identicas al Excel, retefuente 27 UVT 2026  
  <sub>`d2270495f`</sub>

## 6 de junio de 2026

- feat: remisiones — traer items desde una cotizacion (checkboxes, cantidades editables, cliente automatico)  
  <sub>`f3175c9dc`</sub>
- fix: unir solicitudes con el remoto antes de guardar (evita perdidas) + SW v31  
  <sub>`f7f293757`</sub>

## 5 de junio de 2026

- Remisiones: fix consecutivo (sufijos -1), importar Excel con comparacion vs BD, consulta rapida de items en busqueda  
  <sub>`03982daa3`</sub>

## 3 de junio de 2026

- 🐛 Abastecimiento: refrescar la pestaña Kardex al registrar un movimiento  
  <sub>`e49f24b3f`</sub>
- ✨ Seguridad: cierre de sesión automático por inactividad (8 h)  
  <sub>`d706d25cb`</sub>

## 2 de junio de 2026

- ✨ Abastecimiento: botón "Sincronizar con Excel" (concilia stock vs Supabase)  
  <sub>`b6ffb2ae8`</sub>
- ✨ Abastecimiento: mostrar marca y proveedor al dar salida  
  <sub>`eeeafc9fe`</sub>
- 🐛 Fix sync: escritura atascada bloqueaba la lectura (solicitudes no llegaban a todos)  
  <sub>`21a629c34`</sub>

## 1 de junio de 2026

- chore: bump SW v28->v29  
  <sub>`17d776a38`</sub>
- feat: PDF remision premium + rotulo hasta 3 por hoja  
  <sub>`4fdbe53b3`</sub>
- chore: bump SW v27->v28  
  <sub>`d0e072d57`</sub>
- fix: quitar franja superior del PDF + respaldo incluye OCs/cotizaciones  
  <sub>`48c9c5985`</sub>
- feat: script respaldo diario Supabase  
  <sub>`872e0b026`</sub>
- chore: bump SW v26->v27  
  <sub>`a9b32173d`</sub>
- feat: PDF remision industrial + boton respaldo Supabase  
  <sub>`ca70d81fb`</sub>
- chore: bump SW cache v25->v26 (forzar actualizacion remisiones/PDF)  
  <sub>`437d69ef9`</sub>
- feat: PDF y rotulo de remisiones con logo y diseno corporativo  
  <sub>`0b292bfb9`</sub>
- feat: modulo Remisiones (Supabase) + Configuracion al final del menu  
  <sub>`296ef9e92`</sub>
- feat: boton editar (lapiz) en Kardex para corregir colada/remision/lote/notas  
  <sub>`85b304c39`</sub>
- feat: ajuste rapido de stock por producto (conteo fisico) - solo Andrea/Alberto  
  <sub>`aeff4c091`</sub>
- fix: absFetch toleraba mal respuestas 201 vacias (POST kardex) - movimientos no guardaban stock  
  <sub>`0c2043c3a`</sub>
- feat: Movimientos tipo carrito (Excel) en Abastecimiento - entradas/salidas multiples  
  <sub>`d4ae7fa03`</sub>

## 31 de mayo de 2026

- fix: quitar stock critico del reporte 5PM  
  <sub>`ee707695d`</sub>
- fix: workflow alertas lee productos directamente  
  <sub>`bcf45e967`</sub>
- fix: alertas stock sin tabla extra  
  <sub>`7f2c2eb20`</sub>
- feat: workflow alertas stock critico cada 15min  
  <sub>`f3f74d4bb`</sub>
- feat: seccion stock critico en reporte diario  
  <sub>`7d0cd5d81`</sub>
- feat: alertas stock critico en plataforma  
  <sub>`342fe3bb4`</sub>
- feat: workflow recordatorio backup viernes 4:45PM  
  <sub>`c9f413e71`</sub>
- feat: banner + correo recordatorio backup viernes  
  <sub>`62f3e5f35`</sub>
- feat: boton descargar Excel backup inventario  
  <sub>`d410fbc0c`</sub>
- fix: instalar openpyxl en workflow para Excel adjunto  
  <sub>`d208ff855`</sub>
- feat: reporte con Excel adjunto + kardex Supabase + agotados separados  
  <sub>`924774275`</sub>
- fix: kardex del dia al inicio del reporte  
  <sub>`2fffa99f2`</sub>
- feat: inventory report con kardex del dia y agotados separados  
  <sub>`dc177a892`</sub>
- fix: scroll reset en cambio de pestaña Abastecimiento  
  <sub>`2e0f87540`</sub>
- fix: reset scroll al navegar a Abastecimiento  
  <sub>`d748df697`</sub>
- feat: busqueda por nombre en movimientos - flujo paso a paso  
  <sub>`16e4bb5e4`</sub>
- fix: mover page-abastecimiento dentro del content div  
  <sub>`f6f2b4f3f`</sub>

## 29 de mayo de 2026

- fix: header abastecimiento + auto-load al abrir  
  <sub>`7fcd149fa`</sub>
- feat: modulo Abastecimiento + usuarios Yesid y Lina  
  <sub>`871e4ffda`</sub>

## 25 de mayo de 2026

- fix: abrir cartera sin bloqueo popup (window.open antes del await)  
  <sub>`aaabf2c59`</sub>
- fix: precio proveedor no pierde foco al escribir  
  <sub>`9013a3319`</sub>
- fix: miembros de equipo no se borran al recargar  
  <sub>`f8551fc26`</sub>
- fix: sincronizar contraseñas desde GitHub antes del login  
  <sub>`58174efa0`</sub>
- feat: base de datos clientes con autocomplete + pre-llenado cotizacion desde solicitud  
  <sub>`d162636af`</sub>
- fix: cotizGuardarBD usa numero del campo si difiere de cotizEditId  
  <sub>`c305b4285`</sub>
- fix: evitar que cola obsoleta sobreescriba backup reciente del mismo archivo  
  <sub>`07cc380df`</sub>
- fix: max_tokens 2000 para listas de productos largas  
  <sub>`66897ef35`</sub>
- debug: log respuesta de Claude para diagnostico  
  <sub>`34aae1990`</sub>
- fix: includeInlineImages para emails con imagen adjunta o pegada  
  <sub>`fe33895f8`</sub>
- feat: solicitudes muestran lista de productos, forma de pago y observaciones  
  <sub>`7ab7d9300`</sub>
- feat: extraccion de productos, forma de pago y observaciones con Claude  
  <sub>`01de9c6a3`</sub>
- config: URL Apps Script cotizaciones actualizada  
  <sub>`2d81db33b`</sub>
- feat: notif cotiz creada/enviada, solicitud→enviada al enviar cotiz, filtro default pendientes  
  <sub>`5c57f43f3`</sub>
- feat: notificaciones a Andrea y Gerencia en todos los eventos  
  <sub>`2bcd12b8e`</sub>
- feat: notificacion por correo al crear solicitud o enviar cotizacion  
  <sub>`b39980e4a`</sub>
- feat: email-to-cotiz agrega doPost para notificaciones de solicitudes y cotizaciones  
  <sub>`eb75c01f3`</sub>
- feat: solicitudes manuales - boton Nueva solicitud + formulario + editar  
  <sub>`493aff75d`</sub>

## 24 de mayo de 2026

- feat: email-to-cotiz usando Gmail polling (reemplaza Power Automate)  
  <sub>`70c54a16f`</sub>

## 23 de mayo de 2026

- feat: email-to-cotiz.gs — solicitudes desde info@ + WhatsApp a 3 destinatarios  
  <sub>`2ad18589b`</sub>
- feat: solicitudes cotiz desde email — semaforo 12h (Index.html)  
  <sub>`f29619632`</sub>
- fix(sync): proteger items locales pendientes de subida contra sobreescritura del polling  
  <sub>`4bd2f4a74`</sub>
- feat(cotizaciones): email del encabezado PDF se ajusta al vendedor seleccionado  
  <sub>`1deec8ef2`</sub>
- fix(cotizaciones): restringir vista BD solo a usuarios con rol Comercial en equipo  
  <sub>`809e94cc6`</sub>
- feat(cotizaciones): colaboradores solo ven sus propias cotizaciones en BD  
  <sub>`c243fa6cc`</sub>
- feat(cotizaciones): agregar campos vendedor, realizada por y aprobada por + filtro BD  
  <sub>`ef8a7cacc`</sub>

## 22 de mayo de 2026

- fix: sync sin token usa proxy GH; no muestra error si proxy configurado  
  <sub>`db514b4d6`</sub>
- feat: cotizaciones - proveedores BD, precio proveedor/margen, busqueda BD, consecutivo fix, alertas WA  
  <sub>`726d4492e`</sub>
- fix: cartera - quitar test de conexion engañoso por mixed content  
  <sub>`0f43fba8d`</sub>
- fix: auto-completar exige las 4 etapas con fecha + HE done si requerida  
  <sub>`3f8c4ed2b`</sub>
- feat: reporte con fechaIngreso, dias, KPIs + historico completo  
  <sub>`ed777a5f8`</sub>
- feat: fechaIngreso OP + consecutivo cotiz editable + sync claves  
  <sub>`ef1cfca45`</sub>
- fix: convertir num a str al unir lista en resumen etapas  
  <sub>`ce6668832`</sub>
- feat: informe diario con valor + HE + seccion pendientes HE  
  <sub>`ad41a6d94`</sub>
- fix: actualizar GS_SCRIPT_URL a deployment correcto  
  <sub>`bcf4eda0b`</sub>
- fix: preservar valor y hojaEntrada en Google Sheets  
  <sub>`ce3945094`</sub>
- fix: reescribir timeline OC con helper seguro + HE condicional  
  <sub>`214f75fb0`</sub>
- feat: etapa Hoja de Entrada en timeline cuando requerida  
  <sub>`079e901d2`</sub>
- fix extractor cotizaciones proxy y duplicados OC  
  <sub>`88360cb2d`</sub>

## 12 de mayo de 2026

- 🐛 Fix: informe diario por correo incluía OCs borradas  
  <sub>`5df5c8987`</sub>
- 🐛 Fix: OCs borradas reaparecían tras recargar (sync GH+GS)  
  <sub>`0125a996a`</sub>
- 🩹 Botón "Verificar estados" en Procesos OC  
  <sub>`56f314571`</sub>
- 🐛 Fix: OC se marcaba como completada sin facturar  
  <sub>`be9ee68ee`</sub>
- ⚡ email-to-oc: cambiar modelo a Haiku 4.5 (menos sobrecargado)  
  <sub>`9575c7d1f`</sub>
- 🔧 email-to-oc: detección de duplicados por núcleo numérico  
  <sub>`461fbbeba`</sub>
- 🔧 email-to-oc: retry 529 + label-based query + num normalizado  
  <sub>`707d38e8c`</sub>
- 🎯 KPIs 100% flexibles por miembro + foto individual  
  <sub>`ea36a9f7d`</sub>

## 11 de mayo de 2026

- 📷 Importar metas KPI desde imagen con Claude Vision  
  <sub>`d30470b10`</sub>
- 📊 Dashboard Gerencial v1 con KPIs medibles 0-100  
  <sub>`f32ccad26`</sub>
- 📊 Equipo: botones KPIs/Evaluación + 4 tools nuevas en Energy IA  
  <sub>`f3b39e6af`</sub>
- 🔔 teams-notify: migrar a Workflows app (Adaptive Card)  
  <sub>`b0d77c557`</sub>
- 📅 email-to-oc: separar fechaCreacion (emisión OC) de fechaCompra  
  <sub>`17078f594`</sub>

## 7 de mayo de 2026

- 🔧 syncNow muestra el estado real (no falso éxito) + diagnóstico atascos  
  <sub>`4fb865591`</sub>

## 5 de mayo de 2026

- 🗑️ Papelera + Vaciar Cronograma para Proyectos  
  <sub>`08de9ab13`</sub>
- 🗑️ Papelera con Deshacer para OPs  
  <sub>`8e70a0439`</sub>
- 📊 Kardex del día en informe de inventario  
  <sub>`d0548ee7d`</sub>
- 🐛 Fix CORS bloqueando polling + eliminar OP duplicada 2026005442  
  <sub>`0d3858e3f`</sub>
- 🔍 Búsqueda global Ctrl+K — un solo cuadro busca en 8 módulos  
  <sub>`5c426c90d`</sub>
- 🔁 Sync robusto: reintento auto + indicador prominente + encolar fallos OPs  
  <sub>`ffee135e6`</sub>
- ⚠️ op_report.py: banner si datos llevan +4h sin actualizarse  
  <sub>`581de5232`</sub>

## 4 de mayo de 2026

- 🐛 email-to-oc.gs: no marcar leído si Claude falla + log detallado  
  <sub>`62640e06c`</sub>
- ✨ email-to-oc.gs: agregar polling Gmail + trigger 1 min  
  <sub>`251cc54cb`</sub>
- ✨ Auto-creación OC + fix bug savePOCData + nuevo Apps Script email→OC  
  <sub>`3750e062b`</sub>
- 🐛 Fix bug cierre de mes: auto-complete solo si las 4 etapas tienen fecha  
  <sub>`fef61d77c`</sub>
- Bot WhatsApp ENERGY — Apps Script con Twilio + Claude tools  
  <sub>`e61600377`</sub>
- OP cierre de mes: histórico también compartido a GitHub  
  <sub>`39391f9ac`</sub>
- Cotizaciones: 7 estados (Borrador/Enviada/Adjudicada/Facturada/Rechazada/Vencida/Pendiente)  
  <sub>`8ec576a0e`</sub>
- Cotizaciones: PDF imprimible matching formato E&G + buscar inventario  
  <sub>`5d106cc09`</sub>
- Function calling en chat ENERGY: el asistente ya ejecuta acciones  
  <sub>`b556e46f3`</sub>

## 1 de mayo de 2026

- Nuevo módulo Históricos (Solo Andrea/Alberto)  
  <sub>`1dafbdf6e`</sub>
- Zoom de fotos en visitas, mensajería e inventario  
  <sub>`c69202b2b`</sub>
- Cierre de mes manual + histórico para 6 módulos (Tareas, Reuniones, Visitas, Mensajería, Cotizaciones, Minutas)  
  <sub>`108bcb6e1`</sub>
- Permisos por usuario y módulo (paso 6 del roadmap)  
  <sub>`006075e59`</sub>
- Auto-refresh del historial de Visitas y Mensajería al sincronizar  
  <sub>`9ef1a4fc3`</sub>
- Proxy unificado: Claude IA + GitHub Writes en un solo Apps Script  
  <sub>`e48cb1bd5`</sub>
- Bloqueo colaborativo de O.C. (paso 4 del roadmap)  
  <sub>`b3d5199cc`</sub>
- Fix: _saveCompanyConfig usa proxy GitHub (no solo PAT) + botón eliminar más visible  
  <sub>`4a081f051`</sub>
- Auto-recarga al detectar nueva versión + SW v14  
  <sub>`0de1856fe`</sub>
- Proxy GitHub para escritura compartida + eliminar visitas/mensajería (admin)  
  <sub>`d56de12d4`</sub>
- Fotos de Inventario integradas en módulo Inventario + sync compartido  
  <sub>`a901b829f`</sub>
- Eliminar Andrés Barrera del equipo (ya no trabaja)  
  <sub>`a75642f23`</sub>
- Fix items demo que reaparecen + botón 'Resetear datos locales'  
  <sub>`4928023a9`</sub>
- Sync en tiempo real con polling + ETag (paso 1 del roadmap)  
  <sub>`7f238b1c7`</sub>
- Indicador de sync en topbar + cola offline (paso 1+5 del roadmap)  
  <sub>`830eef0b2`</sub>
- Eliminar datos demo hardcoded (Ecopetrol, Mansarovar, Hocol, Perenco, etc.)  
  <sub>`5f572bfc5`</sub>
- Lectura compartida sin token: cualquier navegador ve los mismos datos  
  <sub>`97f6d1aa2`</sub>
- Visitas y Mensajería: avisos claros del sync a GitHub + scroll a historial  
  <sub>`7bd3c21c6`</sub>
- Notificar a Gerencia (Teams) login y logout de usuarios  
  <sub>`179ac2c67`</sub>
- Visitas y Mensajería: fotos opcionales + base de datos compartida  
  <sub>`38876cad6`</sub>
- Fix bug: campo updatedAt contaminado con etiquetas '✅ Facturado'/'❌ Pendiente'  
  <sub>`3a3f4bb73`</sub>

## 30 de abril de 2026

- Fix móvil: el FAB ⚡ tapaba el botón de enviar del chat  
  <sub>`ef624d711`</sub>
- Reporte OP de las 5pm: clasificar por urgencia real, no por etapa pendiente  
  <sub>`21e33e1eb`</sub>
- Fix: login móvil bloqueado por crypto.subtle/cache stale  
  <sub>`0e3410772`</sub>
- Fix chat móvil iOS: header y input siempre visibles + acceso directo a Configuración  
  <sub>`04cac6879`</sub>
- Hamburger más grande y atajo a Configuración desde el chat móvil  
  <sub>`80e890f7f`</sub>
- Sincronizar URL del proxy entre dispositivos via data/config.json  
  <sub>`a56cca707`</sub>
- PWA instalable en celular + chat con memoria y contexto rico  
  <sub>`fbd495c56`</sub>
- Fase 1 (4/4): Hash SHA-256 de contraseñas y código de acceso  
  <sub>`9432923e7`</sub>
- Mostrar audit log en cards de O.C. (Procesos Orden de Compra)  
  <sub>`bb0f220a1`</sub>
- Fase 1 (3/4): Audit log — quién creó/modificó cada cosa  
  <sub>`3846875b8`</sub>
- Fase 1 (2/4): Backup automático de TODOS los datos a GitHub  
  <sub>`7b117bedd`</sub>
- Cambiar proxy de Claude a Google Apps Script (gratis)  
  <sub>`f986d3ecd`</sub>
- Fase 1 (1/4): Proxy seguro de Claude API  
  <sub>`71988872c`</sub>
- Dashboard: mostrar etapa actual y progresión de cada O.C.  
  <sub>`076061ad7`</sub>
- Dashboard: notificaciones reales + sección O.C. en proceso  
  <sub>`347eca050`</sub>
- fix  
  <sub>`6bcaef71b`</sub>

## 29 de abril de 2026

- Create apps-script.gs  
  <sub>`c2211fd90`</sub>
- Invnetario  
  <sub>`f2926491c`</sub>
- Excel: agregar autofiltros en encabezados del informe OP  
  <sub>`7ad36992f`</sub>
- Fix 422: traer SHA antes de guardar y reintentar si falla  
  <sub>`976fd783e`</sub>

## 23 de abril de 2026

- chore: sincronizar horario reporte OC con inventario — 4:00 PM Colombia  
  <sub>`60f152778`</sub>

## 22 de abril de 2026

- chore: cambiar horario informe inventario a 4:00 PM Colombia (19:10 UTC)  
  <sub>`38e264699`</sub>
- TES  
  <sub>`b58d795c6`</sub>
- mo  
  <sub>`cb87c5a43`</sub>
- Update inventory-report.yml  
  <sub>`bfd02c932`</sub>

## 21 de abril de 2026

- te  
  <sub>`4e0e5bf4d`</sub>
- mod  
  <sub>`e0ebb1711`</sub>
- Fix: ajustar hora de reportes a 5 PM Colombia (21:30 UTC con margen para retrasos de GitHub)  
  <sub>`748ebeaf9`</sub>
- chN  
  <sub>`02a6071e0`</sub>
- desplegable  
  <sub>`f4479259d`</sub>
- SCREE  
  <sub>`0ab25c40f`</sub>
- TEST  
  <sub>`ba038e695`</sub>
- REV  
  <sub>`17580def2`</sub>
- PUSG  
  <sub>`08903b8c0`</sub>

## 20 de abril de 2026

- actualizar  
  <sub>`8e82d3466`</sub>
- modulos  
  <sub>`b5f94a476`</sub>

## 17 de abril de 2026

- act_Inv  
  <sub>`7f5d90fa0`</sub>
- Horario  
  <sub>`830dfc3f4`</sub>
- Sincronizar teams  
  <sub>`c087555b8`</sub>
- Actualizar Inventario  
  <sub>`0d5e310b5`</sub>
- Ajuste horario workflows  
  <sub>`860c3930b`</sub>
- SEGURIDAD - ELIMINAR CREDENCIALES  
  <sub>`185e2fd17`</sub>
- CERRAR MES SOLO JEFE  
  <sub>`7cfd04442`</sub>
- fix auto-completar OP  
  <sub>`d7e4db42a`</sub>

## 14 de abril de 2026

- Update op-report.yml  
  <sub>`e6033017b`</sub>

## 10 de abril de 2026

- Update schedule-reports.yml  
  <sub>`cc8dd99c9`</sub>
- Rename main.yml to schedule-reports.yml  
  <sub>`681e9f1a5`</sub>
- Create main.yml  
  <sub>`76756524a`</sub>

## 8 de abril de 2026

- feat: renombrar OC → OP en módulo de órdenes (etiquetas UI)  
  <sub>`ef48ac3ba`</sub>
- chore: eliminar oc-report.yml — renombrado a op-report.yml  
  <sub>`b3632d981`</sub>
- chore: eliminar oc_report.py — renombrado a op_report.py  
  <sub>`152a67fb4`</sub>
- feat: renombrar OC → OP — workflow op-report.yml  
  <sub>`4348de801`</sub>
- feat: renombrar OC → OP (Orden de Pedido) — script op_report.py  
  <sub>`80c6f9463`</sub>
- fix: corregir encoding UTF-8 caracteres especiales y emojis  
  <sub>`c954689ac`</sub>
- fix: restaurar app script funcional + preservar inventario actualizado  
  <sub>`92d7b7a48`</sub>
- fix: load_config tolerante a archivo faltante en GitHub Actions  
  <sub>`3e1bf7d08`</sub>
- feat: workflow reporte diario OC  
  <sub>`4bf1617c1`</sub>
- feat: agregar script reporte diario OC  
  <sub>`29ea0cfe1`</sub>
- Agregar keywords mecanico/electrico: threadolet, sockolet, flanches, filtros, conduit, cajas, tomas  
  <sub>`117d4ab5e`</sub>
- Agregar auto-clasificacion por familia/categoria en inventory_report.py  
  <sub>`cbd3a95c3`</sub>
- Diagnostico: imprimir encabezados de TODAS las hojas para detectar FAMILIA/CATEGORIA  
  <sub>`f36f70ab3`</sub>
- Fix: buscar hoja correcta en Excel (no KARDEX) - lee todas las hojas y elige la de mayor score  
  <sub>`f291bd28b`</sub>
- Mejorar logging de columnas en sync_inventory.py  
  <sub>`d23a11109`</sub>

## 7 de abril de 2026

- Fix Gantt timeline + inventory field mapping  
  <sub>`769898b12`</sub>
- Update sync-inventory.yml  
  <sub>`904af9bb2`</sub>

## 6 de abril de 2026

- Update teams-notify.yml  
  <sub>`3c5f80854`</sub>
- Create sync-inventory.yml  
  <sub>`d76a6ac2e`</sub>
- Create sync_inventory.py  
  <sub>`5cd6e2cc8`</sub>

## 5 de abril de 2026

- Create inventory_report.py  
  <sub>`ff360528c`</sub>
- Create inventory-report.yml  
  <sub>`a038f2971`</sub>

## 27 de marzo de 2026

- Rediseno visual Cotizaciones IA - E&G corporativo colores  
  <sub>`da4f076e0`</sub>
- Cotizaciones IA: formato E&G + base de datos BD  
  <sub>`72d9576cb`</sub>
- Agregar Cotizaciones IA con generador de Excel  
  <sub>`71e5cb728`</sub>
- Fix encoding + agregar campo token GitHub en Configuracion  
  <sub>`c20743d4d`</sub>
- Notificaciones Teams via GitHub Actions (token en localStorage)  
  <sub>`b7d7575d2`</sub>
- Create teams-notify.yml  
  <sub>`70444bc93`</sub>
