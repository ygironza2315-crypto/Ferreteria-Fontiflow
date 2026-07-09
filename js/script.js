const MI_TELEFONO = "573171318781";
let productosBase = []; // Aquí se guarda todo el Excel

// --- 1. CARGA INICIAL OPTIMIZADA PARA MATRIZ DE PRECIOS ---
async function cargarInventario() {
    const contenedor = document.getElementById('contenedor-productos');
    if (!contenedor) return;

    try {
        const response = await fetch('inventario.xlsx?v=' + Date.now());
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];

        // MODIFICACIÓN TÉCNICA CRUCIAL:
        productosBase = XLSX.utils.sheet_to_json(hoja, { raw: false, defval: "0" });

        // === COPIA Y PEGA ESTE BLOQUE DE CONTROL AQUÍ ===
        const pruebaArticulo = productosBase.find(p => p.nombre && p.nombre.includes("SEGUETA"));
        console.log("[Fontiflow Super Debug] Fila completa encontrada en el JSON:", pruebaArticulo);
        // ================================================

        // Renderizado inicial: debeHacerScroll = false para que cargue arriba
        renderizarProductos(productosBase, false);
        configurarFiltrosCategorias();

    } catch (error) {
        console.error("Error:", error);
    }
}

// ==========================================================================
// 1. RENDERIZADO INTELIGENTE (CON DOBLE BOTÓN COMPACTO: COTIZAR Y AÑADIR)
// ==========================================================================
function renderizarProductos(lista, debeHacerScroll = true) {
    const contenedor = document.getElementById('contenedor-productos');
    if (!contenedor) return; 
    
    contenedor.innerHTML = ""; 

    if (lista.length === 0) {
        const msgNoEncontrado = document.getElementById('mensaje-no-encontrado');
        if (msgNoEncontrado) msgNoEncontrado.style.display = 'block';
        return;
    }

    lista.forEach((prod) => {
        const nombre = prod.nombre || "Producto";
        
        // Control de imágenes múltiples
        const imagenes = prod.img ? prod.img.toString().split(',') : ["logo-fontiflow.png"];
        const imgInicial = imagenes[0].trim();
        const listaImgsAtributo = imagenes.map(i => i.trim()).join(',');

        // Filtro inteligente para marcas vacías o en cero
        let cadenaMarca = prod.marca ? prod.marca.toString().trim() : "Genérica";
        if (cadenaMarca === "0" || cadenaMarca === "") {
            cadenaMarca = "Genérica";
        }
        const marcas = cadenaMarca.split(',');

        // Guardamos las matrices crudas en los atributos del contenedor
        const matrizPrecios = prod.precio ? prod.precio.toString().trim() : "0";
        const matrizMedidas = prod.medidas ? prod.medidas.toString().trim() : "Única";

        const cambiaPorMedida = (marcas.length <= 1 && imagenes.length > 1) ? "true" : "false";
        
        const card = document.createElement('div');
        card.className = `col-12 col-md-6 mb-4 item-producto`; 
        
        card.setAttribute('data-precios', matrizPrecios);
        card.setAttribute('data-medidas', matrizMedidas);
        
        card.innerHTML = `
            <div class="card card-producto-horizontal shadow-sm border-0 h-100">
                <div class="row g-0 align-items-center h-100">
                    <div class="col-md-4 contenedor-img-horizontal fondo-producto-azul d-flex align-items-center justify-content-center" style="background-color: #ffffff !important;">
                        <img src="img/${imgInicial}" 
                             data-imagenes="${listaImgsAtributo}" 
                             data-cambia-por-medida="${cambiaPorMedida}"
                             class="img-fluid rounded-start" 
                             onerror="this.src='img/logo-fontiflow.png'" 
                             style="--escala-medida: 0.85;">
                    </div>
                    <div class="col-md-8">
                        <div class="card-body">
                            <h4 class="fw-bold mb-1" style="color: #002d5a;">${nombre}</h4>
                            <p class="small text-muted mb-2 text-uppercase font-weight-bold" style="font-size: 0.75rem;">${prod.cat || ''}</p>
                            
                            <div class="mb-3 depto-marca">
                                ${marcas.length === 1 ? `
                                    <span class="text-muted small">Marca: <strong style="color: #002d5a;">${marcas[0].trim()}</strong></span>
                                ` : marcas.length > 1 ? `
                                    <span class="text-muted d-block small mb-1">Selecciona la marca:</span>
                                    <div class="d-flex flex-wrap gap-1">
                                        ${marcas.map((m, idx) => `
                                            <button type="button" class="btn btn-sm btn-outline-primary py-0 px-2 btn-marca-variante ${idx === 0 ? 'active' : ''}" 
                                                    style="font-size: 0.75rem;" onclick="cambiarMarcaFoto(this, ${idx})">
                                                ${m.trim()}
                                            </button>
                                        `).join('')}
                                    </div>
                                ` : ''}
                            </div>

                            <div class="d-flex flex-wrap gap-2 mb-3 contenedor-botones-medida">
                            </div>
                            
                            <div class="d-flex justify-content-between align-items-center border-top pt-3">
                                
                                <div class="seccion-precio-real justify-content-between align-items-center w-100">
                                    <div>
                                        <small class="text-muted d-block">Precio unitario</small>
                                        <h3 class="fw-bold mb-0 precio-display" style="color: #ff6600; font-size: 1.4rem;">$0</h3>
                                    </div>
                                    <div class="d-flex gap-1">
                                        <a href="#" target="_blank" class="btn btn-sm btn-success fw-bold px-2 btn-cotizar" style="font-size: 0.75rem;">💬 Cotizar</a>
                                        <button type="button" class="btn btn-sm btn-warning fw-bold px-2 shadow-sm btn-anadir" style="font-size: 0.75rem;" 
                                                onclick="agregarAlCarrito(this)">🛒 Añadir</button>
                                    </div>
                                </div>

                                <div class="seccion-precio-bloqueado justify-content-between align-items-center w-100">
                                    <div>
                                        <small class="text-muted d-block">Precio unitario</small>
                                        <h3 class="fw-bold mb-0 text-secondary" style="font-size: 1.1rem; padding-top: 4px;">🔒 Privado</h3>
                                    </div>
                                    <div>
                                        <button type="button" class="btn btn-sm btn-primary fw-bold px-3 shadow-sm btn-abrir-auth" style="font-size: 0.75rem;">🔑 Ver Precio</button>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>
                </div>
            </div>`;
        contenedor.appendChild(card);
        
        window.sincronizarPrecioYVisual(card);
    });

    if (debeHacerScroll) {
        const yOffset = -120; 
        const y = contenedor.getBoundingClientRect().top + window.pageYOffset + yOffset;
        window.scrollTo({top: y, behavior: 'smooth'});
    }
}
// --- 3. FILTROS (Aquí sí queremos scroll) ---
function configurarFiltrosCategorias() {
    const botones = document.querySelectorAll('.btn-categoria-card');
    botones.forEach(btn => {
        btn.addEventListener('click', function () {
            const buscadorInput = document.getElementById('buscador-blog');
            if (buscadorInput) buscadorInput.value = "";

            botones.forEach(b => b.classList.remove('active'));
            this.classList.add('active');

            const filtro = this.getAttribute('data-filter').toLowerCase();
            const filtrados = filtro === "todos"
                ? productosBase
                : productosBase.filter(p => p.cat?.toLowerCase().trim() === filtro);

            // Pasamos 'true' para que al filtrar sí suba la pantalla
            renderizarProductos(filtrados, true);
        });
    });
}
// --- 4. LÓGICA DEL BUSCADOR PROTEGIDA ---
const buscador = document.getElementById('buscador-blog');
if (buscador) {
    buscador.addEventListener('input', function (e) {
        const termino = e.target.value.toLowerCase().trim();
        const productos = document.querySelectorAll('.item-producto');
        const seccionCat = document.getElementById('seccion-categorias');
        const msgNoEncontrado = document.getElementById('mensaje-no-encontrado'); // Lo guardamos en variable
        let encontrados = 0;

        if (termino.length > 0) {
            if (seccionCat) seccionCat.style.display = 'none'; // Ocultar líneas al buscar

            productos.forEach(prod => {
                const contenido = prod.innerText.toLowerCase();
                if (contenido.includes(termino)) {
                    prod.style.display = 'block';
                    encontrados++;
                } else {
                    prod.style.display = 'none';
                }
            });

            // PROTECCIÓN CRUCIAL: Solo cambia el estilo si el elemento de verdad existe en el HTML
            if (msgNoEncontrado) {
                msgNoEncontrado.style.display = (encontrados === 0) ? 'block' : 'none';
            }
        } else {
            // Si borra el buscador, regresamos a mostrar "TODOS"
            if (seccionCat) seccionCat.style.display = 'block';
            if (msgNoEncontrado) msgNoEncontrado.style.display = 'none'; // PROTECCIÓN

            // Re-renderizamos pasándole false para que no mande la pantalla abajo bruscamente
            renderizarProductos(productosBase, false);
        }
    });
}

// Iniciar al cargar el documento
document.addEventListener('DOMContentLoaded', cargarInventario);

// --- MANEJO DEL FORMULARIO DE COTIZACIÓN ---
document.addEventListener('DOMContentLoaded', function () {
    const formulario = document.getElementById('form-cotizacion');

    // Solo ejecutamos si el formulario existe en la página actual
    if (formulario) {
        formulario.addEventListener('submit', function (e) {
            e.preventDefault();

            // Captura de datos
            const nombre = document.getElementById('nombre').value;
            const whatsapp = document.getElementById('whatsapp').value;
            const proyecto = document.getElementById('proyecto').value;
            const lista = document.getElementById('lista').value;

            // Construcción del mensaje para Fontiflow
            const mensaje = `Hola Fontiflow! 👋\n\n` +
                `Mi nombre es: *${nombre}*\n` +
                `Tipo de proyecto: *${proyecto}*\n` +
                `Mi contacto: ${whatsapp}\n\n` +
                `*Lista de materiales:*\n${lista}`;

            // Número de WhatsApp de la ferretería
            const telefono = "573171318781";
            const url = `https://wa.me/${telefono}?text=${encodeURIComponent(mensaje)}`;

            // Abrir en nueva pestaña
            window.open(url, '_blank');
        });
    }

// ==========================================================================
// 2. MOTOR CENTRAL (SINCRO DE PRECIOS, MARCAS, IMÁGENES Y DATA DE BOTONES)
// ==========================================================================
window.sincronizarPrecioYVisual = function(contenedorPrincipal) {
    // A. Detectar qué botón de Marca está seleccionado actualmente
    const btnsMarca = Array.from(contenedorPrincipal.querySelectorAll('.btn-marca-variante'));
    const idxMarca = btnsMarca.findIndex(b => b.classList.contains('active'));
    const safeIdxMarca = idxMarca !== -1 ? idxMarca : 0;

    // === ADAPTACIÓN INTELIGENTE PARA GRUPOS DE LA MISMA MARCA ===
    let cadenaMedidas = contenedorPrincipal.getAttribute('data-medidas') || "Única";
    let cadenaMatriz = contenedorPrincipal.getAttribute('data-precios') || "0";

    if (btnsMarca.length <= 1) {
        cadenaMedidas = cadenaMedidas.split('|').join(',');
        cadenaMatriz = cadenaMatriz.split('|').join(',');
    }
    // ============================================================

    // B. GENERACIÓN DINÁMICA DE BOTONES DE MEDIDA SEGÚN LA MARCA
    const bloquesMedidas = cadenaMedidas.split('|');
    const bloqueMedidaSeleccionado = bloquesMedidas[safeIdxMarca] || bloquesMedidas[0] || "Única";
    const listaMedidas = bloqueMedidaSeleccionado.split(',');

    // Detectamos cuál botón de medida estaba activo antes de redibujar
    const btnsMedidaAntes = Array.from(contenedorPrincipal.querySelectorAll('.btn-medida'));
    let idxMedidaActiva = btnsMedidaAntes.findIndex(b => b.classList.contains('active'));
    if (idxMedidaActiva === -1) idxMedidaActiva = 0;

    // Inyectamos los nuevos botones con los textos correctos de la marca
    const contenedorBotonesMedida = contenedorPrincipal.querySelector('.contenedor-botones-medida');
    if (contenedorBotonesMedida) {
        contenedorBotonesMedida.innerHTML = listaMedidas.map((m, i) => `
            <button type="button" class="btn btn-outline-secondary btn-sm btn-medida ${i === idxMedidaActiva ? 'active' : ''}" 
                    onclick="actualizarPrecio(this, '', ${i})">
                ${m.trim()}
            </button>
        `).join('');
    }

    // C. Extraer los bloques de precios correspondientes
    const bloquesMarcas = cadenaMatriz.split('|'); 
    const bloqueSeleccionado = bloquesMarcas[safeIdxMarca] || bloquesMarcas[0] || "0";
    const listaPrecios = bloqueSeleccionado.split(','); 

    // D. CONTROL DE EXISTENCIAS: Validar disponibilidad de las nuevas medidas creadas
    const btnsMedida = Array.from(contenedorPrincipal.querySelectorAll('.btn-medida'));
    let primerIndiceDisponible = -1;

    btnsMedida.forEach((btn, i) => {
        const precioAsignado = listaPrecios[i]?.trim();

        if (!precioAsignado || precioAsignado === "0" || precioAsignado === "") {
            btn.disabled = true;
            btn.classList.remove('active');
            btn.style.cursor = 'not-allowed';
        } else {
            btn.disabled = false;
            btn.style.cursor = 'pointer';
            if (primerIndiceDisponible === -1) {
                primerIndiceDisponible = i;
            }
        }
    });

    // E. CONTROL AUTOMÁTICO DE SELECCIÓN RESIDUAL
    let idxMedida = btnsMedida.findIndex(b => b.classList.contains('active'));
    if (idxMedida === -1 && primerIndiceDisponible !== -1) {
        btnsMedida[primerIndiceDisponible].classList.add('active');
        idxMedida = primerIndiceDisponible;
    }

    // F. Renderizar el precio final calculado
    const safeIdxMedida = idxMedida !== -1 ? idxMedida : 0;
    const precioFinal = listaPrecios[safeIdxMedida]?.trim() || "0";

    const displayPrecio = contenedorPrincipal.querySelector('.precio-display');
    if (displayPrecio) {
        displayPrecio.innerText = `$${precioFinal}`;
    }

    // G. Aplicar el escalado dinámico por tamaño e intercambio de imagen
    const img = contenedorPrincipal.querySelector('.fondo-producto-azul img');
    if (img) {
        const nuevaEscala = 0.85 + (safeIdxMedida * 0.12);
        img.style.setProperty('--escala-medida', nuevaEscala);

        const cambiaPorMedida = img.getAttribute('data-cambia-por-medida');
        const atributoImgs = img.getAttribute('data-imagenes');
        
        if (atributoImgs) {
            const listaImgs = atributoImgs.split(',');
            if (cambiaPorMedida === "true") {
                const totalMedidasOpciones = btnsMedida.length;
                let indiceFoto = safeIdxMedida;

                if (listaImgs.length < totalMedidasOpciones && listaImgs.length > 0) {
                    const medidasPorFoto = totalMedidasOpciones / listaImgs.length;
                    indiceFoto = Math.floor(safeIdxMedida / medidasPorFoto);
                }

                if (listaImgs[indiceFoto]) {
                    img.src = `img/${listaImgs[indiceFoto].trim()}`;
                }
            } else {
                if (listaImgs[safeIdxMarca]) {
                    img.src = `img/${listaImgs[safeIdxMarca].trim()}`;
                }
            }
        }
    }

    // === CONFIGURACIÓN SIMÉTRICA DE AMBOS BOTONES EN TIEMPO REAL ===
    const nombreProducto = contenedorPrincipal.querySelector('h4')?.innerText || "Producto";
    
    const btnMarcaActivo = contenedorPrincipal.querySelector('.btn-marca-variante.active');
    const marcaTexto = btnMarcaActivo ? btnMarcaActivo.innerText : (contenedorPrincipal.querySelector('.depto-marca strong')?.innerText || "Genérica");
    
    const btnMedidaActivo = contenedorPrincipal.querySelector('.btn-medida.active');
    const medidaTexto = btnMedidaActivo ? btnMedidaActivo.innerText : "Única";

    // 1. Datos dinámicos para el Carrito
    const btnAnadir = contenedorPrincipal.querySelector('.btn-anadir');
    if (btnAnadir) {
        btnAnadir.setAttribute('data-nombre', nombreProducto);
        btnAnadir.setAttribute('data-marca', marcaTexto);
        btnAnadir.setAttribute('data-medida', medidaTexto);
        btnAnadir.setAttribute('data-precio', precioFinal);
    }

    // 2. Enlace directo para la Cotización Express
    const btnWhatsApp = contenedorPrincipal.querySelector('.btn-cotizar');
    if (btnWhatsApp) {
        const mensajeOriginal = `Hola Fontiflow, me interesa cotizar:\n\n*Producto:* ${nombreProducto}\n*Marca:* ${marcaTexto}\n*Medida:* ${medidaTexto}\n*Precio Catálogo:* $${precioFinal}`;
        const mensajeCodificado = encodeURIComponent(mensajeOriginal);
        btnWhatsApp.href = `https://wa.me/${MI_TELEFONO}?text=${mensajeCodificado}`;
    }
};

// ==========================================================================
// 3. ACCIONADORES DE INTERFAZ (CALIBRADO PARA REPARTO DE IMÁGENES MUCHOS-A-UNO)
// ==========================================================================
window.actualizarPrecio = function (boton, precioOmitido, indice) {
    const cardBody = boton.closest('.card-body');
    if (!cardBody) return;

    cardBody.querySelectorAll('.btn-medida').forEach(b => b.classList.remove('active'));
    boton.classList.add('active');

    // Buscamos el contenedor que sí tiene guardada la matriz de precios
    const contenedorPrincipal = boton.closest('.item-producto');
    if (contenedorPrincipal) {
        window.sincronizarPrecioYVisual(contenedorPrincipal);

        // Soporte para cambio de foto por medida (posiciones/grupos)
        const img = contenedorPrincipal.querySelector('.fondo-producto-azul img');
        if (img) {
            const cambiaPorMedida = img.getAttribute('data-cambia-por-medida');
            if (cambiaPorMedida === "true" && indice !== undefined) {
                const atributoImgs = img.getAttribute('data-imagenes');
                if (atributoImgs) {
                    const listaImgs = atributoImgs.split(',');

                    // === AJUSTE MATEMÁTICO DE DISTRIBUCIÓN DE IMÁGENES ===
                    const totalMedidas = contenedorPrincipal.querySelectorAll('.btn-medida').length;
                    let indiceFoto = indice;

                    // Si el usuario puso menos imágenes que medidas en el Excel, las agrupamos proporcionalmente
                    if (listaImgs.length < totalMedidas && listaImgs.length > 0) {
                        const medidasPorFoto = totalMedidas / listaImgs.length;
                        indiceFoto = Math.floor(indice / medidasPorFoto);
                    }

                    // Ejecutamos el cambio de imagen de forma segura
                    if (listaImgs[indiceFoto]) {
                        img.src = `img/${listaImgs[indiceFoto].trim()}`;
                    }
                    // =====================================================
                }
            }
        }
    }
};

window.cambiarMarcaFoto = function (boton, indiceMarca) {
    // SOLUCIÓN: Buscamos el contenedor que sí tiene guardada la matriz de precios
    const contenedorPrincipal = boton.closest('.item-producto');
    if (!contenedorPrincipal) return;

    const contenedorMarcas = boton.closest('.depto-marca');
    contenedorMarcas.querySelectorAll('.btn-marca-variante').forEach(b => b.classList.remove('active'));
    boton.classList.add('active');

    const img = contenedorPrincipal.querySelector('.fondo-producto-azul img');
    if (img) {
        const cambiaPorMedida = img.getAttribute('data-cambia-por-medida');
        if (cambiaPorMedida !== "true" && indiceMarca !== undefined) {
            const atributoImgs = img.getAttribute('data-imagenes');
            if (atributoImgs) {
                const listaImgs = atributoImgs.split(',');
                if (listaImgs[indiceMarca]) {
                    img.src = `img/${listaImgs[indiceMarca].trim()}`;
                }
            }
        }
    }

    // Al mutar la marca, recalculamos la intersección de la matriz de precios
    window.sincronizarPrecioYVisual(contenedorPrincipal);
};
// ==========================================================================
// LÓGICA ESTRUCTURADA PARA EL EFECTO LUPA EN EL CATÁLOGO
// ==========================================================================
const catálogoProductos = document.getElementById('contenedor-productos');

if (catálogoProductos) {
    // Detecta el movimiento del mouse sobre la foto para mover la lupa
    catálogoProductos.addEventListener('mousemove', function (e) {
        const cajaFoto = e.target.closest('.fondo-producto-azul');
        if (!cajaFoto) return;

        const imagen = cajaFoto.querySelector('img');
        if (!imagen) return;

        // Calculamos la posición del cursor relativa a la caja blanca (0% a 100%)
        const dimensiones = cajaFoto.getBoundingClientRect();
        const coordenadaX = ((e.clientX - dimensiones.left) / dimensiones.width) * 100;
        const coordenadaY = ((e.clientY - dimensiones.top) / dimensiones.height) * 100;

        // Desplazamos el centro del zoom hacia la posición exacta del mouse
        imagen.style.transformOrigin = `${coordenadaX}% ${coordenadaY}%`;
    });

    // Cuando el mouse sale del producto, restauramos el centro original suavemente
    catálogoProductos.addEventListener('mouseout', function (e) {
        const cajaFoto = e.target.closest('.fondo-producto-azul');
        if (!cajaFoto) return;

        const imagen = cajaFoto.querySelector('img');
        if (imagen) {
            imagen.style.transformOrigin = 'center center';
        }
    });
}
// ==========================================================================
// 5. SISTEMA CENTRAL E-COMMERCE HÍBRIDO (CARRITO + CHECKOUT DE ENVÍO)
// ==========================================================================
window.carrito = [];

function inicializarInterfazCarrito() {
    if (document.getElementById('btn-carrito-flotante')) return;

    // 1. Botón Flotante Redondo
    const botonFlotanteHTML = `
        <div id="btn-carrito-flotante" class="position-fixed bottom-0 end-0 m-4 shadow" style="z-index: 1050; display: none;">
            <button class="btn btn-primary btn-lg rounded-pill fw-bold d-flex align-items-center gap-2 px-4 py-3" 
                    style="background-color: #002d5a; border-color: #002d5a;" data-bs-toggle="modal" data-bs-target="#modalCarrito" onclick="mostrarPaso(1)">
                🛒 Mi Pedido 
                <span id="badge-contador-carrito" class="badge bg-warning text-dark rounded-pill">0</span>
            </button>
        </div>
    `;

    // 2. Modal de Dos Pasos (Carrito y Formulario de Pago)
    const modalCarritoHTML = `
        <div class="modal fade" id="modalCarrito" tabindex="-1" aria-labelledby="modalCarritoLabel" aria-hidden="true" style="z-index: 1060;">
            <div class="modal-dialog modal-dialog-centered modal-md">
                <div class="modal-content border-0 shadow">
                    
                    <!-- CABECERA -->
                    <div class="modal-header text-white border-0" style="background-color: #002d5a;">
                        <h5 class="modal-title fw-bold" id="modalCarritoLabel">🛒 Mi Carrito Fontiflow</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>

                    <!-- CUERPO DE LA VENTANA (VISTAS INTERCAMBIABLES) -->
                    <div class="modal-body" style="max-height: 450px; overflow-y: auto;">
                        
                        <!-- PASO 1: LISTA DE PRODUCTOS -->
                        <div id="vista-paso-1">
                            <div id="lista-items-carrito"></div>
                            <div class="d-flex justify-content-between align-items-center border-top pt-3 mt-3 bg-light p-2 rounded">
                                <span class="text-muted fw-bold">Total Productos:</span>
                                <h4 class="fw-bold mb-0 text-dark" id="total-carrito-p1">$0</h4>
                            </div>
                            <div class="text-end mt-3">
                                <button type="button" class="btn btn-primary fw-bold px-4" style="background-color: #ff6600; border-color: #ff6600;" onclick="mostrarPaso(2)">
                                    Siguiente: Datos de Envío ➔
                                </button>
                            </div>
                        </div>

                        <!-- PASO 2: FORMULARIO DE DESPACHO Y PAGO -->
                        <div id="vista-paso-2" style="display: none;">
                            <h6 class="fw-bold mb-3 text-muted border-bottom pb-2">📋 Datos para el Despacho</h6>
                            <form id="form-checkout" onsubmit="procesarOrdenFinal(event)">
                                <div class="mb-2">
                                    <label class="small fw-bold text-muted">Nombre Completo *</label>
                                    <input type="text" id="chk-nombre" class="form-control form-control-sm" required placeholder="Ej: Juan Pérez">
                                </div>
                                <div class="mb-2">
                                    <label class="small fw-bold text-muted">Teléfono de Contacto *</label>
                                    <input type="tel" id="chk-telefono" class="form-control form-control-sm" required placeholder="Ej: 315XXXXXXX">
                                </div>
                                <div class="mb-3">
                                    <label class="small fw-bold text-muted">Dirección de Entrega *</label>
                                    <input type="text" id="chk-direccion" class="form-control form-control-sm" required placeholder="Calle, Carrera, Número, Barrio">
                                </div>

                                <h6 class="fw-bold mb-3 text-muted border-bottom pb-2 mt-4">💳 Selecciona el Método de Pago</h6>
                                
                                <!-- Opción A: WhatsApp Híbrido -->
                                <div class="form-check p-3 border rounded mb-2 bg-white shadow-sm list-group-item-action" style="cursor:pointer;">
                                    <input class="form-check-input ms-1" type="radio" name="metodoPago" id="pago-whatsapp" value="whatsapp" checked>
                                    <label class="form-check-label fw-bold d-block text-dark ms-4" for="pago-whatsapp">
                                        🚚 Contra Entrega / Transferencia (Manual)
                                        <small class="text-muted d-block fw-normal font-size-xs">Pagas en efectivo al recibir o transfieres directo a nuestro Nequi/Daviplata.</small>
                                    </label>
                                </div>

                                <!-- Opción B: Pasarela de Pago Seguro -->
                                <div class="form-check p-3 border rounded mb-3 bg-white shadow-sm list-group-item-action" style="cursor:pointer;">
                                    <input class="form-check-input ms-1" type="radio" name="metodoPago" id="pago-online" value="online">
                                    <label class="form-check-label fw-bold d-block text-dark ms-4" for="pago-online">
                                        💳 Pago Online Automático (PSE, Tarjetas, Nequi)
                                        <small class="text-muted d-block fw-normal font-size-xs">Pagas de forma segura e inmediata a través de nuestra pasarela virtual integrada.</small>
                                    </label>
                                </div>

                                <div class="d-flex justify-content-between align-items-center border-top pt-3 bg-light p-2 rounded">
                                    <div>
                                        <small class="text-muted d-block">Total a Pagar</small>
                                        <h4 class="fw-bold mb-0 text-success" id="total-carrito-p2">$0</h4>
                                    </div>
                                    <div class="d-flex gap-2">
                                        <button type="button" class="btn btn-outline-secondary btn-sm fw-bold" onclick="mostrarPaso(1)">⬅ Volver</button>
                                        <button type="submit" class="btn btn-success btn-sm fw-bold px-3 shadow-sm">🚀 Confirmar Compra</button>
                                    </div>
                                </div>
                            </form>
                        </div>

                    </div>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', botonFlotanteHTML);
    document.body.insertAdjacentHTML('beforeend', modalCarritoHTML);
}

// CONTROLADOR DE CAMBIO DE PASOS INTERNOS
window.mostrarPaso = function(paso) {
    const p1 = document.getElementById('vista-paso-1');
    const p2 = document.getElementById('vista-paso-2');
    
    if (paso === 1) {
        if(p1) p1.style.display = 'block';
        if(p2) p2.style.display = 'none';
        renderizarCarrito();
    } else {
        if (window.carrito.length === 0) return;
        if(p1) p1.style.display = 'none';
        if(p2) p2.style.display = 'block';
        
        // Sincronizamos el total en el paso 2
        const totalAcumulado = window.carrito.reduce((suma, item) => suma + (item.precio * item.cantidad), 0);
        const t2 = document.getElementById('total-carrito-p2');
        if (t2) t2.innerText = `$${totalAcumulado}`;
    }
};

// CONTROL DE INYECCIÓN INTEGRADO
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", inicializarInterfazCarrito);
} else {
    inicializarInterfazCarrito();
}

// AÑADIR PRODUCTO DETECTANDO VARIANTES ACTÚALES
window.agregarAlCarrito = function(boton) {
    const nombre = boton.getAttribute('data-nombre');
    const marca = boton.getAttribute('data-marca');
    const medida = boton.getAttribute('data-medida');
    const precio = parseFloat(boton.getAttribute('data-precio')) || 0;

    if (precio === 0) return; 

    const idUnico = `${nombre}-${marca}-${medida}`;
    const productoExistente = window.carrito.find(item => item.id === idUnico);

    if (productoExistente) {
        productoExistente.cantidad += 1;
    } else {
        window.carrito.push({
            id: idUnico, nombre: nombre, marca: marca, medida: medida, precio: precio, cantidad: 1
        });
    }

    actualizarContadoresUX();

    const textoOriginal = boton.innerHTML;
    boton.innerText = "¡Añadido! ✓";
    boton.classList.replace('btn-warning', 'btn-success');
    setTimeout(() => {
        boton.innerHTML = textoOriginal;
        boton.classList.replace('btn-success', 'btn-warning');
    }, 800);
};

function actualizarContadoresUX() {
    const totalItems = window.carrito.reduce((suma, item) => suma + item.cantidad, 0);
    const wrapperFlotante = document.getElementById('btn-carrito-flotante');
    const badgeContador = document.getElementById('badge-contador-carrito');

    if (badgeContador) badgeContador.innerText = totalItems;
    if (wrapperFlotante) wrapperFlotante.style.display = totalItems > 0 ? 'block' : 'none';
}

window.renderizarCarrito = function() {
    const contenedorLista = document.getElementById('lista-items-carrito');
    const contenedorTotalP1 = document.getElementById('total-carrito-p1');
    if (!contenedorLista) return;

    if (window.carrito.length === 0) {
        contenedorLista.innerHTML = `<p class="text-center text-muted my-4">El carrito está vacío. Añade productos del catálogo.</p>`;
        if (contenedorTotalP1) contenedorTotalP1.innerText = "$0";
        return;
    }

    let acumuladoTotal = 0;

    contenedorLista.innerHTML = window.carrito.map((item, idx) => {
        const subtotalItem = item.precio * item.cantidad;
        acumuladoTotal += subtotalItem;

        return `
            <div class="d-flex align-items-center justify-content-between border-bottom py-2">
                <div style="max-width: 65%;">
                    <h6 class="fw-bold mb-0 text-dark" style="font-size: 0.9rem;">${item.nombre}</h6>
                    <small class="text-muted d-block" style="font-size: 0.75rem;">Marca: ${item.marca} | Med: ${item.medida}</small>
                    <span class="fw-bold text-muted small" style="font-size: 0.8rem;">$${item.precio} c/u</span>
                </div>
                <div class="d-flex align-items-center gap-2">
                    <div class="input-group input-group-sm" style="width: 80px;">
                        <button class="btn btn-outline-secondary py-0" type="button" onclick="modificarCantidad(${idx}, -1)">-</button>
                        <span class="form-control text-center py-0 fw-bold bg-white" style="font-size: 0.8rem; padding: 2px;">${item.cantidad}</span>
                        <button class="btn btn-outline-secondary py-0" type="button" onclick="modificarCantidad(${idx}, 1)">+</button>
                    </div>
                    <button class="btn btn-sm btn-link text-danger p-0 ms-1" onclick="removerDelCarrito(${idx})">🗑️</button>
                </div>
            </div>
        `;
    }).join('');

    if (contenedorTotalP1) contenedorTotalP1.innerText = `$${acumuladoTotal}`;
};

window.modificarCantidad = function(indice, cambio) {
    if (!window.carrito[indice]) return;
    window.carrito[indice].cantidad += cambio;
    if (window.carrito[indice].cantidad <= 0) window.carrito.splice(indice, 1);
    actualizarContadoresUX();
    renderizarCarrito();
};

window.removerDelCarrito = function(indice) {
    window.carrito.splice(indice, 1);
    actualizarContadoresUX();
    renderizarCarrito();
};

// ==========================================================================
// MÓDULO CENTRAL DE PROCESAMIENTO Y DESPACHO DE VENTAS
// ==========================================================================
window.procesarOrdenFinal = function(event) {
    event.preventDefault(); // Evitamos que la página se recargue

    const nombre = document.getElementById('chk-nombre').value.trim();
    const telefono = document.getElementById('chk-telefono').value.trim();
    const direccion = document.getElementById('chk-direccion').value.trim();
    const tipoPago = document.querySelector('input[name="metodoPago"]:checked').value;

    let totalGeneral = window.carrito.reduce((suma, item) => suma + (item.precio * item.cantidad), 0);

    if (tipoPago === "whatsapp") {
        // --- METODO 1: ORDEN FORMAL PARA PAGO CONTRA ENTREGA/MANUAL ---
        
        // 1. AUTOMATIZACIÓN: Forzamos el cierre de la ventana del carrito de inmediato
        const btnCerrarModal = document.querySelector('#modalCarrito .btn-close');
        if (btnCerrarModal) btnCerrarModal.click();

        // Construcción de la plantilla para WhatsApp
        let plantilla = `🛒 *NUEVA ORDEN DE COMPRA - FONTIFLOW*\n`;
        plantilla += `¡Hola! Acabo de armar un pedido desde la tienda online:\n\n`;
        plantilla += `👤 *Cliente:* ${nombre}\n`;
        plantilla += `📞 *Teléfono:* ${telefono}\n`;
        plantilla += `📍 *Dirección:* ${direccion}\n`;
        plantilla += `💳 *Método de Pago:* Contra Entrega / Transferencia\n\n`;
        plantilla += `📦 *DETALLE DE ARTÍCULOS:*\n`;

        window.carrito.forEach((item, i) => {
            plantilla += `${i + 1}. *${item.nombre}* [x${item.cantidad}]\n`;
            plantilla += `   _M:_ ${item.marca} | _Med:_ ${item.medida} ($${item.precio} c/u)\n`;
        });

        plantilla += `\n💰 *TOTAL NETO A PAGAR:* $${totalGeneral}\n\n`;
        plantilla += `Por favor, confírmenme el despacho de la mercancía. ¡Muchas gracias!`;

        const stringCodificado = encodeURIComponent(plantilla);
        
        // Despachamos el pedido abriendo el chat de WhatsApp
        window.open(`https://wa.me/${MI_TELEFONO}?text=${stringCodificado}`, '_blank');

        // 2. AUTOMATIZACIÓN: Lanzamos la ventana flotante con los números de cuenta
        // Le ponemos un mini retraso de medio segundo para que la animación se vea perfecta
        setTimeout(() => {
            mostrarVentanaDatosTransferencia();
        }, 500);

        // Vaciamos el carrito local para que la orden quede cerrada y limpia
        window.carrito = [];
        if (typeof actualizarContadoresUX === "function") actualizarContadoresUX();

    } else {
        // --- METODO 2: PASARELA DE PAGO VIRTUAL DIRECTA ---
        
        // Cerramos el modal del carrito de inmediato
        const btnCerrarModal = document.querySelector('#modalCarrito .btn-close');
        if (btnCerrarModal) btnCerrarModal.click();

        // Llamamos directo a la pasarela nativa que ya está cargada en el HTML
        ejecutarPasarelaTransaccional(totalGeneral, nombre, telefono, direccion);
    }
};

// ==========================================================================
// MOTOR TRANSACCIONAL AUTOSUFICIENTE (CON VALIDACIÓN DE MONTO MÍNIMO)
// ==========================================================================
function ejecutarPasarelaTransaccional(total, nombre, telefono, direccion) {
    
    // 🚀 1. CANDADO DE SEGURIDAD COMERCIAL (Colocar al principio de todo)
    if (total < 1500) {
        alert("Para pagos online con Tarjeta o PSE, el monto mínimo de compra debe ser de $1.500 pesos.\n\nPara montos menores, por favor selecciona el método de pago 'Contra Entrega' o paga en efectivo al recibir tu pedido en Ciudad del Valle.");
        return; // 🛑 Detiene la ejecución aquí mismo. No cierra el modal ni abre Wompi.
    }

    // 2. Si el total es de $1.500 o más, el código continúa su curso normal hacia abajo:
    const montoEnCentavos = Math.round(total * 100);
    const referenciaFactura = `FF-${Date.now()}`; 

    const btnCerrarModal = document.querySelector('#modalCarrito .btn-close');
    if (btnCerrarModal) btnCerrarModal.click();

    const abrirCheckoutWompi = async () => {
        try {
            const llavePublica = 'pub_prod_iOswYsKdmhiVZAzFOVa2pZMBHXFYebMj'; 
            const secretoIntegridad = 'prod_integrity_AIjDiDLgxcB2qoXB2GSoTzvHYLz3LYee'; 

            // Cifrado digital SHA-256
            const cadenaParaCifrar = referenciaFactura + montoEnCentavos + 'COP' + secretoIntegridad;
            const encoder = new TextEncoder();
            const data = encoder.encode(cadenaParaCifrar);
            const hashBuffer = await crypto.subtle.digest('SHA-256', data);
            const hashArray = Array.from(new Uint8Array(hashBuffer));
            const firmaHexadecimal = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');

            const checkout = new WidgetCheckout({
                currency: 'COP',
                amountInCents: montoEnCentavos,
                reference: referenciaFactura,
                publicKey: llavePublica,
                signature: firmaHexadecimal, 
                customerData: {
                    fullName: nombre,
                    phoneNumber: telefono,
                    phoneNumberPrefix: '+57' 
                }
            });

            checkout.open(function ( resultado ) {
                const transaccion = resultado.transaction;
                
                if (transaccion.status === 'APPROVED') {
                    alert(`¡Excelente! Pago aprobado con éxito.\nID Transacción: ${transaccion.id}`);

                    let msgExito = `✅ *PAGO CONFIRMADO ONLINE - FONTIFLOW*\n`;
                    msgExito += `El cliente ha pagado con éxito a través de PSE/Tarjeta.\n\n`;
                    msgExito += `👤 *Comprador:* ${nombre}\n`;
                    msgExito += `📞 *Teléfono:* ${telefono}\n`;
                    msgExito += `📍 *Dirección de Despacho:* ${direccion}\n\n`;
                    msgExito += `💰 *Monto Recaudado:* $${total}\n`;
                    msgExito += `🆔 *ID Transacción:* ${transaccion.id}\n`;
                    msgExito += `🧾 *Referencia:* ${referenciaFactura}\n\n`;
                    msgExito += `Por favor procedan con el alistamiento de los materiales.`;

                    window.open(`https://wa.me/${MI_TELEFONO}?text=${encodeURIComponent(msgExito)}`, '_blank');

                    window.carrito = [];
                    if (typeof actualizarContadoresUX === "function") actualizarContadoresUX();

                } else if (transaccion.status === 'DECLINED') {
                    alert('La transacción fue declinada por el banco. Intenta con otro medio de pago o usa la opción Contra Entrega.');
                }
            });

        } catch (error) {
            console.error("Error al inicializar el objeto Wompi o generar firma:", error);
            alert("Ocurrió un problema de configuración en las credenciales de la pasarela. Revisa la consola (F12).");
        }
    };

    if (typeof WidgetCheckout !== 'undefined') {
        abrirCheckoutWompi();
    } else {
        console.log("Inyectando script de Wompi en caliente desde JavaScript...");
        
        const scriptInyectado = document.createElement('script');
        scriptInyectado.src = 'https://checkout.wompi.co/widget.js';
        scriptInyectado.type = 'text/javascript';
        
        scriptInyectado.onload = () => {
            abrirCheckoutWompi();
        };

        scriptInyectado.onerror = () => {
            alert("Error de conexión: No se pudo descargar el sistema de pagos seguros de Wompi.");
        };

        document.head.appendChild(scriptInyectado);
    }
}
// ==========================================================================
// VENTANA FLOTANTE EXCLUSIVA: DATOS DE TRANSFERENCIA DIRECTA (NEQUI/DAVIPLATA)
// ==========================================================================
function mostrarVentanaDatosTransferencia() {
    // Seguridad: si por alguna razón ya existía una ventana abierta, la removemos
    const existente = document.getElementById('alerta-transferencia-fontiflow');
    if (existente) existente.remove();
    const fondoExistente = document.getElementById('fondo-alerta-transferencia');
    if (fondoExistente) fondoExistente.remove();

    // Estructura HTML y CSS integrada de la ventana flotante
    const htmlAlerta = `
        <div id="fondo-alerta-transferencia" class="position-fixed top-0 start-0 w-100 h-100 bg-dark opacity-50" style="z-index: 1999;"></div>

        <div id="alerta-transferencia-fontiflow" class="position-fixed top-50 start-50 translate-middle p-4 rounded shadow-lg bg-white border border-light text-center animate__animated animate__zoomIn" style="z-index: 2000; max-width: 90%; width: 420px; font-family: sans-serif;">
            <div class="mb-2" style="font-size: 2.5rem;">🎉</div>
            <h4 class="fw-bold mb-1" style="color: #002d5a;">¡Pedido Recibido!</h4>
            <p class="text-muted small mb-4">Tu lista ya fue enviada a nuestro WhatsApp. Si elegiste pagar por transferencia, aquí tienes nuestras cuentas oficiales:</p>
            
            <div class="p-3 mb-2 rounded border border-light text-start shadow-sm" style="background-color: #fcf4f9; border-left: 5px solid #d10074 !important;">
                <span class="fw-bold d-block" style="color: #d10074; font-size: 0.8rem; letter-spacing: 1px;">MOVIMIENTO ELECTRÓNICO</span>
                <strong class="fs-5 d-block text-dark mt-1">📱 NEQUI: 3164235294</strong>
                <small class="text-muted d-block mt-1">Titular: Ferretería Fontiflow</small>
            </div>

            <div class="p-3 mb-4 rounded border border-light text-start shadow-sm" style="background-color: #fff5f5; border-left: 5px solid #e60000 !important;">
                <span class="fw-bold d-block" style="color: #e60000; font-size: 0.8rem; letter-spacing: 1px;">DEPOSITOS EN LÍNEA</span>
                <strong class="fs-5 d-block text-dark mt-1">❤️ DAVIPLATA: 3164235294</strong>
                <small class="text-muted d-block mt-1">Titular: Ferretería Fontiflow</small>
            </div>

            <div class="p-2 bg-light rounded mb-3">
                <p class="text-danger fw-bold m-0" style="font-size: 0.78rem;">
                    ⚠️ IMPORTANTE: Adjunta el comprobante de pago en el chat de WhatsApp que se abrió para validar tu despacho.
                </p>
            </div>

            <button type="button" class="btn btn-primary fw-bold w-100 py-2 shadow-sm" 
                    style="background-color: #002d5a; border-color: #002d5a;" 
                    onclick="cerrarVentanaTransferencia()">
                Entendido, Finalizar
            </button>
        </div>
    `;

    // Inyectamos el componente vivo al body de la página web
    document.body.insertAdjacentHTML('beforeend', htmlAlerta);
}

// Función encargada de limpiar la pantalla al dar clic en cerrar
window.cerrarVentanaTransferencia = function() {
    document.getElementById('alerta-transferencia-fontiflow')?.remove();
    document.getElementById('fondo-alerta-transferencia')?.remove();
};

});

