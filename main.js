// ============================================
// CONFIGURACIÓN
// ============================================

// URL del Web App desplegado de Google Apps Script
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzcV7o_86qlxy0mtjg93knlv44VHy1MpKPAasc5Z35LaVochZutsVQUT7JheiIg5Fs0/exec"; // Tomar como referencia al archivo Código.js en el repositorio

// Modo de desarrollo (usa datos de prueba cuando el backend no está configurado)
const DEV_MODE = false;

// ============================================
// VARIABLES GLOBALES
// ============================================

let dataTable = null;
let quillRequerimiento = null;
let quillComentarios = null;
let proyectosData = [];
let currentProyectoId = null;
let adjuntosData = [];        // Array de adjuntos en el formulario
let adjuntoIdCounter = 0;     // Contador para IDs únicos de adjuntos

// ============================================
// DATOS DE PRUEBA (para desarrollo sin backend)
// ============================================

const datosDemo = [
    {
        id: '735709281',
        fechaCreacion: '15/3/2026',
        marca: 'IEmpresa',
        responsable: 'Erick',
        area: 'OSE',
        participantes: 'Roy',
        requerimiento: '<p>Crear un formulario que carece thoth etc etc conectado con bitrix</p>',
        comentarios: '<p>opcional</p>',
        nombreProyecto: '✅ Calendario Metoring',
        accesos: [
            { titulo: 'Admin', valor: 'admin | Password: 123' },
            { titulo: 'Cliente', valor: 'cliente | Password: 123' },
            { titulo: 'OSE', valor: 'ose | Password: 123' }
        ],
        plataformas: [
            { titulo: 'Data', valor: 'https://baltic.bitrix24.es/crm/deal/category/272/' },
            { titulo: 'Github', valor: 'https://baltic.bitrix24.es/crm/deal/category/272/' },
            { titulo: 'Loocker', valor: 'https://baltic.bitrix24.es/crm/deal/category/272/' }
        ],
        imagenes: [
            { titulo: 'Diagrama de secuencia', valor: 'https://colegionarval.org/wp-content/uploads/tipos-de-inteligencia.jpg' }
        ],
        estado: 'En proceso',
        fechaActualizacion: '15/3/2026',
        comentariosEstado: 'opcional'
    },
    {
        id: '735709282',
        fechaCreacion: '14/3/2026',
        marca: 'JVN',
        responsable: 'Scott',
        area: 'DAD',
        participantes: 'Mayra, Luis',
        requerimiento: '<p>Desarrollar módulo de reportes para el sistema de inventarios</p>',
        comentarios: '<p>Incluir exportación a Excel y PDF</p>',
        nombreProyecto: 'Módulo Reportes Inventario',
        accesos: [
            { titulo: 'Admin', valor: 'admin_jvn | Password: secure123' }
        ],
        plataformas: [
            { titulo: 'Repositorio', valor: 'https://github.com/ejemplo/reportes' }
        ],
        imagenes: [],
        estado: 'Terminado',
        fechaActualizacion: '16/3/2026',
        comentariosEstado: 'Entregado al cliente'
    },
    {
        id: '735709283',
        fechaCreacion: '13/3/2026',
        marca: 'Blackwell',
        responsable: 'Luigui',
        area: '.ID.',
        participantes: '',
        requerimiento: '<p>Integración con API de pagos</p>',
        comentarios: '',
        nombreProyecto: 'Gateway de Pagos',
        accesos: [],
        plataformas: [
            { titulo: 'Documentación API', valor: 'https://docs.stripe.com/' }
        ],
        imagenes: [],
        estado: 'Rechazado',
        fechaActualizacion: '14/3/2026',
        comentariosEstado: 'Cliente canceló el proyecto'
    }
];

// ============================================
// INICIALIZACIÓN
// ============================================

$(document).ready(function() {
    initQuillEditors();
    initDataTable();
    initEventListeners();
    loadProyectos();
});

/**
 * Inicializa los editores Quill
 */
function initQuillEditors() {
    const toolbarOptions = [
        ['bold', 'italic', 'underline', 'strike'],
        ['blockquote', 'code-block'],
        [{ 'list': 'ordered'}, { 'list': 'bullet' }],
        [{ 'header': [1, 2, 3, false] }],
        ['link'],
        ['clean']
    ];

    quillRequerimiento = new Quill('#editorRequerimiento', {
        theme: 'snow',
        modules: { toolbar: toolbarOptions },
        placeholder: 'Describe el requerimiento del proyecto...'
    });

    quillComentarios = new Quill('#editorComentarios', {
        theme: 'snow',
        modules: { toolbar: toolbarOptions },
        placeholder: 'Agrega comentarios adicionales...'
    });
}

/**
 * Inicializa DataTable
 */
function initDataTable() {
    dataTable = $('#tablaProyectos').DataTable({
        responsive: true,
        language: {
            url: 'https://cdn.datatables.net/plug-ins/1.13.7/i18n/es-ES.json'
        },
        order: [[1, 'desc']],
        pageLength: 10,
        columns: [
            { data: 'id', width: '80px' },
            {
                data: 'fechaCreacion',
                width: '130px',
                render: function(data) {
                    return displayDate(data);
                }
            },
            { 
                data: 'nombreProyecto',
                render: function(data, type, row) {
                    return `<a href="#" class="project-name-link" data-id="${row.id}">${data}</a>`;
                }
            },
            { 
                data: 'marca',
                render: function(data) {
                    const badgeClass = `badge-${data.toLowerCase().replace('.', '')}`;
                    return `<span class="badge badge-marca ${badgeClass}">${data}</span>`;
                }
            },
            { data: 'responsable' },
            { data: 'area' },
            { 
                data: 'estado',
                render: function(data) {
                    let badgeClass = 'badge-en-proceso';
                    if (data === 'Terminado') badgeClass = 'badge-terminado';
                    if (data === 'Rechazado') badgeClass = 'badge-rechazado';
                    if (data === 'Observaciones') badgeClass = 'badge-observaciones';
                    return `<span class="badge badge-estado ${badgeClass}">${data}</span>`;
                }
            },
            {
                data: null,
                orderable: false,
                render: function(data, type, row) {
                    return `
                        <div class="d-flex">
                            <button class="btn btn-info btn-action" onclick="verProyecto('${row.id}')" title="Ver detalle">
                                <i class="fas fa-eye text-white"></i>
                            </button>
                            <button class="btn btn-warning btn-action" onclick="editarProyecto('${row.id}')" title="Editar">
                                <i class="fas fa-edit text-white"></i>
                            </button>
                            <button class="btn btn-danger btn-action" onclick="eliminarProyecto('${row.id}')" title="Eliminar">
                                <i class="fas fa-trash text-white"></i>
                            </button>
                        </div>
                    `;
                }
            }
        ]
    });
}

/**
 * Inicializa event listeners
 */
function initEventListeners() {
    // Nuevo proyecto
    $('#btnNuevoProyecto').on('click', nuevoProyecto);
    
    // Guardar proyecto
    $('#btnGuardarProyecto').on('click', guardarProyecto);
    
    // Agregar campos dinámicos
    $('#btnAddAcceso').on('click', () => agregarCampoDinamico('accesos'));
    $('#btnAddPlataforma').on('click', () => agregarCampoDinamico('plataformas'));
    
    // Filtros
    $('#filterMarca, #filterEstado, #filterResponsable').on('change', aplicarFiltros);
    
    // Click en nombre de proyecto
    $(document).on('click', '.project-name-link', function(e) {
        e.preventDefault();
        const id = $(this).data('id');
        verProyecto(id);
    });
    
    // Editar desde modal detalle
    $('#btnEditarDesdeDetalle').on('click', function() {
        $('#modalDetalle').modal('hide');
        setTimeout(() => {
            editarProyecto(currentProyectoId);
        }, 300);
    });
    
    // Limpiar formulario al cerrar modal
    $('#modalProyecto').on('hidden.bs.modal', limpiarFormulario);

    // Marca y Área con opción "Otros" libre
    $('#marca').on('change', function() { toggleOtrosInput('marca', $(this).val()); });
    $('#area').on('change',  function() { toggleOtrosInput('area',  $(this).val()); });

    // Inicializar módulo de adjuntos (drag & drop, paste, selección)
    initAdjuntos();
}

// ============================================
// CRUD OPERATIONS
// ============================================

/**
 * Carga los proyectos desde el backend
 */
function loadProyectos() {
    showLoading(true);
    
    if (DEV_MODE || APPS_SCRIPT_URL === 'TU_URL_DE_APPS_SCRIPT_AQUI') {
        // Usar datos de demostración
        setTimeout(() => {
            proyectosData = datosDemo;
            renderProyectos();
            updateStats();
            populateFilters();
            showLoading(false);
        }, 800);
        return;
    }
    
    // Llamada real al backend
    $.ajax({
        url: APPS_SCRIPT_URL,
        method: 'GET',
        data: { action: 'getProyectos' },
        success: function(response) {
            if (response.success) {
                proyectosData = response.data;
                renderProyectos();
                updateStats();
                populateFilters();
            } else {
                showError('Error al cargar proyectos: ' + response.message);
            }
        },
        error: function(xhr, status, error) {
            console.error("Detalle del error:", xhr, status, error);
            showError('Error de conexión: ' + error);
        },
        complete: function() {
            showLoading(false);
        }
    });
}

/**
 * Renderiza los proyectos en la tabla
 */
function renderProyectos() {
    dataTable.clear();
    dataTable.rows.add(proyectosData);
    dataTable.draw();
}

/**
 * Abre modal para nuevo proyecto
 */
function nuevoProyecto() {
    currentProyectoId = null;
    $('#modalProyectoTitle').text('Nuevo Proyecto');
    $('#proyectoId').val('');
    limpiarFormulario();
    $('#modalProyecto').modal('show');
}

/**
 * Carga un proyecto para edición
 */
function editarProyecto(id) {
    const proyecto = proyectosData.find(p => p.id === id);
    if (!proyecto) {
        showError('Proyecto no encontrado');
        return;
    }
    
    currentProyectoId = id;
    $('#modalProyectoTitle').text('Editar Proyecto');
    $('#proyectoId').val(id);
    
    // Llenar campos básicos
    $('#nombreProyecto').val(proyecto.nombreProyecto);
    $('#responsable').val(proyecto.responsable);
    $('#participantes').val(proyecto.participantes);
    $('#estado').val(proyecto.estado);
    $('#comentariosEstado').val(proyecto.comentariosEstado);

    // Marca – soporte de opción libre
    const marcasPredefinidas = ['IEmpresa','JVN','Blackwell','ITAE','Eurocoach','Iberoteca','Thoth','Otros',''];
    if (marcasPredefinidas.includes(proyecto.marca)) {
        $('#marca').val(proyecto.marca);
        $('#marcaOtrosDiv').hide().find('input').val('');
    } else {
        $('#marca').val('Otros');
        $('#marcaOtros').val(proyecto.marca);
        $('#marcaOtrosDiv').show();
    }

    // Área – soporte de opción libre
    const areasPredefinidas = ['OSE','DAD','.ID.','Otros',''];
    if (areasPredefinidas.includes(proyecto.area)) {
        $('#area').val(proyecto.area);
        $('#areaOtrosDiv').hide().find('input').val('');
    } else {
        $('#area').val('Otros');
        $('#areaOtros').val(proyecto.area);
        $('#areaOtrosDiv').show();
    }
    
    // Llenar editores Quill
    quillRequerimiento.root.innerHTML = proyecto.requerimiento || '';
    quillComentarios.root.innerHTML = proyecto.comentarios || '';
    
    // Llenar campos dinámicos (accesos y plataformas)
    $('#accesosContainer').empty();
    $('#plataformasContainer').empty();
    
    if (proyecto.accesos && proyecto.accesos.length > 0) {
        proyecto.accesos.forEach(acceso => agregarCampoDinamico('accesos', acceso.titulo, acceso.valor));
    }
    if (proyecto.plataformas && proyecto.plataformas.length > 0) {
        proyecto.plataformas.forEach(plat => agregarCampoDinamico('plataformas', plat.titulo, plat.valor));
    }

    // Cargar adjuntos existentes en el nuevo sistema
    adjuntosData = [];
    if (proyecto.imagenes && proyecto.imagenes.length > 0) {
        proyecto.imagenes.forEach(img => {
            const fileId     = getDriveFileId(img.valor);
            const isImg      = isDriveImageUrl(img.valor);
            const thumbUrl   = fileId
                ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`
                : (isImg ? img.valor : null);
            const viewUrl    = fileId ? `https://drive.google.com/file/d/${fileId}/view`    : img.valor;
            const previewUrl = fileId ? `https://drive.google.com/file/d/${fileId}/preview` : img.valor;
            adjuntosData.push({
                id          : 'adj_' + (++adjuntoIdCounter),
                titulo      : img.titulo,
                valor       : thumbUrl || img.valor,  // guardar thumbnailUrl para imágenes
                thumbnail   : thumbUrl,
                thumbnailUrl: thumbUrl,
                viewUrl,
                previewUrl,
                fileId,
                isImage     : isImg,
                type        : isImg ? 'image' : 'other',
                uploading   : false,
                error       : false
            });
        });
    }
    updateAdjuntosGrid();
    
    actualizarVisibilidadNoItems();
    $('#modalProyecto').modal('show');
}

/**
 * Guarda el proyecto (crear o actualizar)
 */
async function guardarProyecto() {
    if (!$('#formProyecto')[0].checkValidity()) {
        $('#formProyecto')[0].reportValidity();
        return;
    }

    // Validar campos "Otros" si están activos
    if ($('#marca').val() === 'Otros' && !$('#marcaOtros').val().trim()) {
        showError('Por favor escribe el nombre de la marca.');
        $('#marcaOtros').focus();
        return;
    }
    if ($('#area').val() === 'Otros' && !$('#areaOtros').val().trim()) {
        showError('Por favor escribe el nombre del área.');
        $('#areaOtros').focus();
        return;
    }
    
    const proyecto = {
        id: $('#proyectoId').val() || generarId(),
        fechaCreacion: $('#proyectoId').val() ? 
            proyectosData.find(p => p.id === $('#proyectoId').val())?.fechaCreacion : 
            formatDate(new Date()),
        marca: getMarcaValue(),
        responsable: $('#responsable').val(),
        area: getAreaValue(),
        participantes: $('#participantes').val(),
        requerimiento: quillRequerimiento.root.innerHTML,
        comentarios: quillComentarios.root.innerHTML,
        nombreProyecto: $('#nombreProyecto').val(),
        accesos: obtenerCamposDinamicos('accesos'),
        plataformas: obtenerCamposDinamicos('plataformas'),
        imagenes: adjuntosData
            .filter(a => !a.uploading && !a.error)
            .map(a => ({ titulo: a.titulo, valor: a.valor })),
        estado: $('#estado').val(),
        fechaActualizacion: formatDate(new Date()),
        comentariosEstado: $('#comentariosEstado').val()
    };
    
    showLoading(true);
    
    try {
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            redirect: 'follow', // IMPORTANTE para evitar el 405
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                action: currentProyectoId ? 'updateProyecto' : 'createProyecto',
                proyecto: proyecto
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            Swal.fire({ icon: 'success', title: currentProyectoId ? 'Actualizado' : 'Creado', text: result.message, timer: 2000 });
            $('#modalProyecto').modal('hide');
            loadProyectos();
        } else {
            showError('Error al guardar: ' + result.message);
        }
    } catch (error) {
        showError('Error de conexión: ' + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * Elimina un proyecto usando Fetch API
 */
async function eliminarProyecto(id) {
    const proyecto = proyectosData.find(p => p.id === id);
    const result = await Swal.fire({ title: '¿Eliminar?', icon: 'warning', showCancelButton: true, confirmButtonText: 'Sí' });
    
    if (result.isConfirmed) {
        showLoading(true);
        try {
            const response = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                redirect: 'follow',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'deleteProyecto', id: id })
            });
            const res = await response.json();
            if (res.success) {
                Swal.fire('Eliminado', '', 'success');
                loadProyectos();
            } else {
                showError(res.message);
            }
        } catch (error) {
            showError('Error: ' + error.message);
        } finally {
            showLoading(false);
        }
    }
}

/**
 * Muestra el detalle de un proyecto
 */
function verProyecto(id) {
    const proyecto = proyectosData.find(p => p.id === id);
    if (!proyecto) {
        showError('Proyecto no encontrado');
        return;
    }
    
    currentProyectoId = id;
    $('#modalDetalleTitle').text(proyecto.nombreProyecto);
    
    let estadoBadge = 'badge-en-proceso';
    if (proyecto.estado === 'Terminado') estadoBadge = 'badge-terminado';
    if (proyecto.estado === 'Rechazado') estadoBadge = 'badge-rechazado';
    if (proyecto.estado === 'Observaciones') estadoBadge = 'badge-observaciones';
    
    let accesosHtml = '';
    if (proyecto.accesos && proyecto.accesos.length > 0) {
        accesosHtml = `<div class="acceso-grid">${proyecto.accesos.map(a => `
            <div class="acceso-card">
                <span class="acceso-label">${a.titulo}</span>
                <div class="acceso-value">${a.valor}</div>
            </div>
        `).join('')}</div>`;
    } else {
        accesosHtml = '<p class="text-muted small mb-0"><i class="fas fa-info-circle me-1"></i>No hay accesos registrados</p>';
    }
    
    let plataformasHtml = '';
    if (proyecto.plataformas && proyecto.plataformas.length > 0) {
        plataformasHtml = `<div class="plataforma-grid">${proyecto.plataformas.map(p => `
            <div class="plataforma-card">
                <span class="plataforma-label">${p.titulo}</span>
                <a class="plataforma-link" href="${p.valor}" target="_blank"><i class="fas fa-external-link-alt me-1"></i>${p.valor}</a>
            </div>
        `).join('')}</div>`;
    } else {
        plataformasHtml = '<p class="text-muted small mb-0"><i class="fas fa-info-circle me-1"></i>No hay plataformas registradas</p>';
    }
    
    const imagenesHtml = renderAdjuntosDetalle(proyecto.imagenes);
    
    const marcaClass = proyecto.marca
        ? proyecto.marca.toLowerCase().replace(/\./g, '').replace(/ /g, '')
        : 'otros';

    const html = `
        <div class="detalle-proyecto">

            <!-- Info general del proyecto -->
            <div class="detalle-info-grid mb-4">
                ${proyecto.marca ? `
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-tag me-1"></i>Marca</span>
                    <span class="detalle-info-value"><span class="badge badge-marca badge-${marcaClass} fs-6">${proyecto.marca}</span></span>
                </div>` : ''}
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-user me-1"></i>Responsable</span>
                    <span class="detalle-info-value fw-semibold">${proyecto.responsable}</span>
                </div>
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-building me-1"></i>Área</span>
                    <span class="detalle-info-value">${proyecto.area}</span>
                </div>
                ${proyecto.participantes ? `
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-users me-1"></i>Participantes</span>
                    <span class="detalle-info-value">${proyecto.participantes}</span>
                </div>` : ''}
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-calendar-plus me-1"></i>Creado</span>
                    <span class="detalle-info-value">${proyecto.fechaCreacion}</span>
                </div>
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-hashtag me-1"></i>ID</span>
                    <span class="detalle-info-value text-muted">${proyecto.id}</span>
                </div>
            </div>

            <div class="row g-4">
                <div class="col-lg-8">

                    <!-- Plataformas / URLs -->
                    <div class="detail-section">
                        <h6><i class="fas fa-globe me-2"></i>Plataformas / URLs</h6>
                        ${plataformasHtml}
                    </div>

                    <!-- Accesos -->
                    <div class="detail-section">
                        <h6><i class="fas fa-key me-2"></i>Accesos</h6>
                        ${accesosHtml}
                    </div>

                    <!-- Requerimiento -->
                    <div class="detail-section">
                        <h6><i class="fas fa-file-alt me-2"></i>Requerimiento</h6>
                        <div class="detalle-rich-content">
                            ${proyecto.requerimiento || '<em class="text-muted">Sin requerimiento especificado</em>'}
                        </div>
                    </div>

                    <!-- Imágenes / Documentos -->
                    <div class="detail-section">
                        <h6><i class="fas fa-images me-2"></i>Adjuntos</h6>
                        ${imagenesHtml}
                    </div>

                    ${proyecto.comentarios ? `
                    <div class="detail-section">
                        <h6><i class="fas fa-comments me-2"></i>Comentarios</h6>
                        <div class="detalle-rich-content">
                            ${proyecto.comentarios}
                        </div>
                    </div>` : ''}

                </div>

                <div class="col-lg-4">
                    <!-- Tarjeta de estado -->
                    <div class="card border-0 shadow-sm detalle-estado-card">
                        <div class="card-body">
                            <p class="detalle-estado-titulo">Estado del Proyecto</p>
                            <div class="text-center mb-3">
                                <span class="badge badge-estado ${estadoBadge} detalle-estado-badge">${proyecto.estado}</span>
                            </div>
                            <hr>
                            <div class="detalle-meta-row">
                                <span class="detalle-meta-label">Última actualización</span>
                                <span class="detalle-meta-value">${proyecto.fechaActualizacion || '-'}</span>
                            </div>
                            ${proyecto.comentariosEstado ? `
                            <hr>
                            <div class="detalle-meta-row">
                                <span class="detalle-meta-label">Nota de estado</span>
                                <span class="detalle-meta-value">${proyecto.comentariosEstado}</span>
                            </div>` : ''}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    $('#detalleContent').html(html);
    $('#modalDetalle').modal('show');
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Agrega un campo dinámico (acceso, plataforma o imagen)
 */
function agregarCampoDinamico(tipo, titulo = '', valor = '') {
    const container = $(`#${tipo}Container`);
    const index = container.children('.dynamic-field').length;
    
    let placeholderTitulo = 'Título';
    let placeholderValor = 'URL o valor';
    let iconClass = 'fa-link';
    
    if (tipo === 'accesos') {
        placeholderTitulo = 'Ej: Admin, Cliente, Usuario';
        placeholderValor = 'Ej: usuario | Password: 123';
        iconClass = 'fa-key';
    } else if (tipo === 'plataformas') {
        placeholderTitulo = 'Ej: Github, Bitrix, Loocker';
        placeholderValor = 'https://...';
        iconClass = 'fa-globe';
    } else if (tipo === 'imagenes') {
        placeholderTitulo = 'Ej: Diagrama, Wireframe, Mockup';
        placeholderValor = 'URL de la imagen';
        iconClass = 'fa-image';
    }
    
    const html = `
        <div class="dynamic-field" data-index="${index}">
            <button type="button" class="btn btn-outline-danger btn-sm btn-remove" onclick="removerCampoDinamico(this, '${tipo}')">
                <i class="fas fa-times"></i>
            </button>
            <div class="row g-2">
                <div class="col-md-4">
                    <div class="input-group input-group-sm">
                        <span class="input-group-text"><i class="fas ${iconClass}"></i></span>
                        <input type="text" class="form-control campo-titulo" placeholder="${placeholderTitulo}" value="${titulo}">
                    </div>
                </div>
                <div class="col-md-8">
                    <input type="text" class="form-control form-control-sm campo-valor" placeholder="${placeholderValor}" value="${valor}">
                </div>
            </div>
        </div>
    `;
    
    container.append(html);
    actualizarVisibilidadNoItems();
}

/**
 * Remueve un campo dinámico
 */
function removerCampoDinamico(btn, tipo) {
    $(btn).closest('.dynamic-field').remove();
    actualizarVisibilidadNoItems();
}

/**
 * Obtiene los valores de campos dinámicos
 */
function obtenerCamposDinamicos(tipo) {
    const items = [];
    $(`#${tipo}Container .dynamic-field`).each(function() {
        const titulo = $(this).find('.campo-titulo').val().trim();
        const valor = $(this).find('.campo-valor').val().trim();
        if (titulo || valor) {
            items.push({ titulo, valor });
        }
    });
    return items;
}

/**
 * Actualiza la visibilidad de los mensajes "No hay items"
 */
function actualizarVisibilidadNoItems() {
    ['accesos', 'plataformas'].forEach(tipo => {
        const hasItems = $(`#${tipo}Container .dynamic-field`).length > 0;
        $(`#no${tipo.charAt(0).toUpperCase() + tipo.slice(1)}`).toggle(!hasItems);
    });
}

/**
 * Limpia el formulario
 */
function limpiarFormulario() {
    $('#formProyecto')[0].reset();
    quillRequerimiento.root.innerHTML = '';
    quillComentarios.root.innerHTML = '';
    $('#accesosContainer').empty();
    $('#plataformasContainer').empty();
    adjuntosData = [];
    updateAdjuntosGrid();
    actualizarVisibilidadNoItems();
    currentProyectoId = null;
    $('#marcaOtrosDiv, #areaOtrosDiv').hide().find('input').val('');
}

/**
 * Actualiza las estadísticas
 */
function updateStats() {
    const total = proyectosData.length;
    const enProceso = proyectosData.filter(p => p.estado === 'En proceso').length;
    const terminados = proyectosData.filter(p => p.estado === 'Terminado').length;
    const observaciones = proyectosData.filter(p => p.estado === 'Observaciones').length;
    const rechazados = proyectosData.filter(p => p.estado === 'Rechazado').length;
    
    $('#statTotal').text(total);
    $('#statEnProceso').text(enProceso);
    $('#statTerminados').text(terminados);
    $('#statObservaciones').text(observaciones);
    $('#statRechazados').text(rechazados);
}

/**
 * Popula los filtros con valores únicos
 */
function populateFilters() {
    // Marcas
    const marcas = [...new Set(proyectosData.map(p => p.marca))];
    const filterMarca = $('#filterMarca');
    filterMarca.find('option:not(:first)').remove();
    marcas.forEach(marca => {
        filterMarca.append(`<option value="${marca}">${marca}</option>`);
    });
    
    // Responsables
    const responsables = [...new Set(proyectosData.map(p => p.responsable))];
    const filterResponsable = $('#filterResponsable');
    filterResponsable.find('option:not(:first)').remove();
    responsables.forEach(resp => {
        filterResponsable.append(`<option value="${resp}">${resp}</option>`);
    });
}

/**
 * Aplica los filtros a la tabla
 */
function aplicarFiltros() {
    const marca = $('#filterMarca').val();
    const estado = $('#filterEstado').val();
    const responsable = $('#filterResponsable').val();
    
    let filteredData = proyectosData;
    
    if (marca) {
        filteredData = filteredData.filter(p => p.marca === marca);
    }
    if (estado) {
        filteredData = filteredData.filter(p => p.estado === estado);
    }
    if (responsable) {
        filteredData = filteredData.filter(p => p.responsable === responsable);
    }
    
    dataTable.clear();
    dataTable.rows.add(filteredData);
    dataTable.draw();
}

/**
 * Muestra/oculta el overlay de carga
 */
function showLoading(show) {
    if (show) {
        $('#loadingOverlay').addClass('show');
    } else {
        $('#loadingOverlay').removeClass('show');
    }
}

/**
 * Muestra un mensaje de error
 */
function showError(message) {
    Swal.fire({
        icon: 'error',
        title: 'Error',
        text: message,
        confirmButtonColor: '#2563eb'
    });
}

// ============================================
// ADJUNTOS — Drag & Drop / Paste / Upload
// ============================================

// —— Helpers para URLs de Google Drive ——

/** Extrae el fileId de cualquier URL de Google Drive */
function getDriveFileId(url) {
    if (!url) return null;
    // thumbnail?id=FILE_ID  /  uc?export=view&id=FILE_ID
    let m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) return m[1];
    // /d/FILE_ID/  (file/d/, document/d/, etc.)
    m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) return m[1];
    return null;
}

/** Devuelve la URL de miniatura funcional para un archivo Drive */
function getDriveThumbnailUrl(url, sz = 'w800') {
    const id = getDriveFileId(url);
    return id ? `https://drive.google.com/thumbnail?id=${id}&sz=${sz}` : null;
}

/** Devuelve la URL de preview embeddable */
function getDrivePreviewUrl(url) {
    const id = getDriveFileId(url);
    return id ? `https://drive.google.com/file/d/${id}/preview` : url;
}

/** Devuelve la URL de vista/descarga */
function getDriveViewUrl(url) {
    const id = getDriveFileId(url);
    return id ? `https://drive.google.com/file/d/${id}/view` : url;
}

/** Detecta si una URL apunta a una imagen */
function isDriveImageUrl(url) {
    if (!url) return false;
    if (url.includes('thumbnail?id=') || url.includes('/thumbnail?')) return true;
    if (url.includes('uc?export=view') || url.includes('export=view')) return true;
    if (/\.(jpg|jpeg|png|gif|webp|bmp|svg)(\?|$)/i.test(url)) return true;
    return false;
}

// Temp store para adjuntos del modal de detalle
let _detalleAdjuntos = [];

/**
 * Renderiza el grid de adjuntos en la vista de detalle.
 * Miniaturas clickeables; docs muestran preview embebido en SweetAlert.
 */
function renderAdjuntosDetalle(imagenes) {
    if (!imagenes || imagenes.length === 0) {
        return '<p class="text-muted small mb-0"><i class="fas fa-info-circle me-1"></i>No hay adjuntos registrados</p>';
    }
    _detalleAdjuntos = imagenes.map(img => {
        const fileId     = getDriveFileId(img.valor);
        const isImg      = isDriveImageUrl(img.valor);
        const thumbUrl   = fileId
            ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w600`
            : (isImg ? img.valor : null);
        return {
            titulo      : img.titulo,
            valor       : img.valor,
            fileId,
            isImage     : isImg,
            thumbnailUrl: thumbUrl,
            thumbnail   : thumbUrl,
            viewUrl     : fileId ? `https://drive.google.com/file/d/${fileId}/view`    : img.valor,
            previewUrl  : fileId ? `https://drive.google.com/file/d/${fileId}/preview` : img.valor,
            type        : isImg ? 'image' : 'other'
        };
    });

    const cards = _detalleAdjuntos.map((a, i) => {
        const thumb = a.thumbnailUrl
            ? `<img src="${a.thumbnailUrl}" alt="${escHtml(a.titulo)}"
                   onerror="this.style.display='none';this.nextElementSibling.style.display='flex'">
               <div style="display:none;width:100%;height:100%;align-items:center;justify-content:center">
                   <i class="fas ${getFileIcon(a.type)} fa-2x"></i>
               </div>`
            : `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center">
                   <i class="fas ${getFileIcon(a.type)} fa-2x"></i>
               </div>`;
        return `
            <div class="adjunto-detalle-card" onclick="expandDetalleAdjunto(${i})" title="${escHtml(a.titulo)}">
                <div class="adjunto-detalle-thumb">${thumb}</div>
                <div class="adjunto-detalle-label">${escHtml(a.titulo)}</div>
            </div>`;
    });
    return `<div class="adjuntos-detalle-grid">${cards.join('')}</div>`;
}

/** Expande un adjunto de la vista de detalle por índice */
function expandDetalleAdjunto(index) {
    const a = _detalleAdjuntos[index];
    if (a) expandAdjunto(a);
}

/**
 * Inicializa el módulo de adjuntos: drop zone, input file, paste global
 */
function initAdjuntos() {
    const dropZone = document.getElementById('adjuntosDropZone');
    const fileInput = document.getElementById('adjuntosFileInput');
    if (!dropZone || !fileInput) return;

    // Botón "Seleccionar archivos"
    document.getElementById('btnSelectFiles')?.addEventListener('click', () => fileInput.click());
    // Clic directo en la zona
    dropZone.addEventListener('click', (e) => {
        if (e.target.id !== 'btnSelectFiles' && !e.target.closest('#btnSelectFiles')) fileInput.click();
    });

    // Cambio en el input file
    fileInput.addEventListener('change', (e) => {
        handleFiles(Array.from(e.target.files));
        fileInput.value = '';
    });

    // Drag & Drop
    dropZone.addEventListener('dragover',  (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', (e) => { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files);
        if (files.length) handleFiles(files);
    });

    // Pegar desde portapapeles (Ctrl+V) — solo cuando el modal está abierto
    document.addEventListener('paste', (e) => {
        if (!$('#modalProyecto').hasClass('show')) return;
        const files = Array.from(e.clipboardData.items)
            .filter(i => i.kind === 'file')
            .map(i => i.getAsFile())
            .filter(Boolean);
        if (files.length) { e.preventDefault(); handleFiles(files); }
    });
}

/** Devuelve el tipo de archivo a partir del MIME type */
function getFileType(file) {
    const m = file.type;
    if (m.startsWith('image/')) return 'image';
    if (m === 'application/pdf') return 'pdf';
    if (m.includes('word') || m.includes('document')) return 'doc';
    if (m.includes('excel') || m.includes('spreadsheet')) return 'xls';
    return 'other';
}

/** Devuelve la clase FontAwesome + color para el tipo de archivo */
function getFileIcon(type) {
    const icons  = { image:'fa-image', pdf:'fa-file-pdf', doc:'fa-file-word', xls:'fa-file-excel', other:'fa-file' };
    const colors = { image:'text-info', pdf:'text-danger', doc:'text-primary', xls:'text-success', other:'text-secondary' };
    return `${icons[type]||icons.other} ${colors[type]||colors.other}`;
}

/**
 * Procesa un array de File objects: crea entrada en adjuntosData y sube a Drive
 */
async function handleFiles(files) {
    for (const file of files) {
        const id      = 'adj_' + (++adjuntoIdCounter);
        const type    = getFileType(file);
        const tempUrl = URL.createObjectURL(file);

        adjuntosData.push({
            id, type, uploading: true, error: false,
            titulo      : file.name.replace(/\.[^/.]+$/, ''),
            valor       : tempUrl,
            thumbnail   : type === 'image' ? tempUrl : null,  // blob URL para preview instantáneo
            thumbnailUrl: null,
            viewUrl     : null,
            previewUrl  : null,
            fileId      : null,
            isImage     : type === 'image',
            _tempUrl    : tempUrl
        });
        updateAdjuntosGrid();

        try {
            const result = await uploadFileToGoogleDrive(file);
            const idx = adjuntosData.findIndex(a => a.id === id);
            if (idx !== -1) {
                // Revocar blob URL solo si ya tenemos una URL de Drive real
                if (result.fileId && adjuntosData[idx]._tempUrl) {
                    URL.revokeObjectURL(adjuntosData[idx]._tempUrl);
                }
                Object.assign(adjuntosData[idx], {
                    valor       : result.valor,
                    thumbnailUrl: result.thumbnailUrl,
                    viewUrl     : result.viewUrl,
                    previewUrl  : result.previewUrl,
                    fileId      : result.fileId,
                    isImage     : result.isImage,
                    // Usar thumbnail Drive URL para imagen (reemplaza blob URL)
                    thumbnail   : result.isImage ? (result.thumbnailUrl || result.valor) : null,
                    uploading   : false,
                    _tempUrl    : null
                });
                updateAdjuntosGrid();
            }
        } catch (err) {
            const idx = adjuntosData.findIndex(a => a.id === id);
            if (idx !== -1) { adjuntosData[idx].uploading = false; adjuntosData[idx].error = true; updateAdjuntosGrid(); }
            console.error('Error al subir adjunto:', err);
        }
    }
}

/**
 * Sube un archivo a Google Drive vía el Apps Script endpoint.
 * Devuelve un objeto con thumbnailUrl, viewUrl, previewUrl, fileId, isImage.
 */
async function uploadFileToGoogleDrive(file) {
    // En modo dev no hay backend — usar blob URL temporal
    if (DEV_MODE || !APPS_SCRIPT_URL || APPS_SCRIPT_URL === 'TU_URL_DE_APPS_SCRIPT_AQUI') {
        const tempUrl = URL.createObjectURL(file);
        const isImage = file.type.startsWith('image/');
        return { valor: tempUrl, thumbnailUrl: isImage ? tempUrl : null,
                 viewUrl: tempUrl, previewUrl: tempUrl, fileId: null, isImage };
    }
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const base64Data = e.target.result.split(',')[1];
                const resp = await fetch(APPS_SCRIPT_URL, {
                    method : 'POST',
                    redirect: 'follow',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body   : JSON.stringify({ action:'uploadImage', fileName: file.name, base64Data, mimeType: file.type })
                });
                const result = await resp.json();
                if (result.success) {
                    resolve({
                        valor       : result.data.url,           // thumbnailUrl para imágenes, viewUrl para docs
                        thumbnailUrl: result.data.thumbnailUrl,
                        viewUrl     : result.data.viewUrl,
                        previewUrl  : result.data.previewUrl,
                        fileId      : result.data.fileId,
                        isImage     : result.data.isImage
                    });
                } else {
                    reject(new Error(result.message || 'Error al subir'));
                }
            } catch (err) { reject(err); }
        };
        reader.onerror = () => reject(new Error('Error al leer archivo'));
        reader.readAsDataURL(file);
    });
}

/** Actualiza el grid de previsualización de adjuntos */
function updateAdjuntosGrid() {
    const grid = document.getElementById('adjuntosPreviewGrid');
    if (!grid) return;
    grid.innerHTML = adjuntosData.length === 0 ? '' : adjuntosData.map(renderAdjuntoCard).join('');
}

/** Renderiza una tarjeta de adjunto */
function renderAdjuntoCard(a) {
    // Durante carga: usar blob URL (a.thumbnail). Tras subida: usar thumbnailUrl de Drive
    const imgSrc = a.uploading ? a.thumbnail : (a.thumbnailUrl || a.thumbnail);
    const thumb = imgSrc
        ? `<img src="${imgSrc}" alt="${escHtml(a.titulo)}" onclick="expandAdjuntoById('${a.id}')" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'">
           <div style="display:none;width:100%;height:100%;align-items:center;justify-content:center;cursor:pointer" onclick="expandAdjuntoById('${a.id}')"><i class="fas ${getFileIcon(a.type)} fa-2x"></i></div>`
        : `<i class="fas ${getFileIcon(a.type)} fa-2x" onclick="expandAdjuntoById('${a.id}')" style="cursor:pointer"></i>`;
    const spinner = a.uploading
        ? `<div class="adjunto-uploading"><div class="spinner-border spinner-border-sm text-primary" role="status"></div><span class="ms-1 small text-muted">Subiendo...</span></div>`
        : '';
    const errBadge = a.error
        ? `<span class="adjunto-error-badge" title="Error al subir"><i class="fas fa-exclamation-triangle text-warning"></i></span>`
        : '';
    return `
        <div class="adjunto-card" data-id="${a.id}">
            ${errBadge}
            <div class="adjunto-thumb">${thumb}${spinner}</div>
            <div class="adjunto-info">
                <input type="text" class="form-control form-control-sm adjunto-titulo-input"
                       value="${escHtml(a.titulo)}" placeholder="Título..."
                       oninput="updateAdjuntoTitulo('${a.id}',this.value)">
            </div>
            <button type="button" class="adjunto-remove" onclick="removeAdjunto('${a.id}')" title="Eliminar">
                <i class="fas fa-times"></i>
            </button>
        </div>`;
}

/** Escapa HTML para uso seguro en atributos */
function escHtml(str) {
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/** Actualiza el título de un adjunto al escribir en su input */
function updateAdjuntoTitulo(id, value) {
    const a = adjuntosData.find(x => x.id === id);
    if (a) a.titulo = value;
}

/** Elimina un adjunto del grid y de memoria */
function removeAdjunto(id) {
    const a = adjuntosData.find(x => x.id === id);
    if (a && a._tempUrl) URL.revokeObjectURL(a._tempUrl);
    adjuntosData = adjuntosData.filter(x => x.id !== id);
    updateAdjuntosGrid();
}

/** Expande un adjunto por su ID (lookup en adjuntosData) */
function expandAdjuntoById(id) {
    const a = adjuntosData.find(x => x.id === id);
    if (a && !a.uploading) expandAdjunto(a);
}

/**
 * Muestra vista expandida de un adjunto.
 * Acepta un objeto adjunto con: { titulo, valor, thumbnailUrl, viewUrl, previewUrl, isImage, type }
 */
function expandAdjunto(a) {
    const isImg = a.isImage || a.type === 'image';
    if (isImg) {
        const imgUrl = a.thumbnailUrl || a.thumbnail || a.valor;
        const viewUrl = a.viewUrl || getDriveViewUrl(a.valor);
        Swal.fire({
            title           : a.titulo,
            imageUrl        : imgUrl,
            imageAlt        : a.titulo,
            showConfirmButton: false,
            showCloseButton : true,
            imageWidth      : '100%',
            footer          : viewUrl
                ? `<a href="${viewUrl}" target="_blank" class="text-primary small"><i class="fas fa-external-link-alt me-1"></i>Abrir en Drive</a>`
                : '',
            customClass     : { popup: 'swal-image-popup' }
        });
    } else {
        const previewUrl = a.previewUrl || getDrivePreviewUrl(a.valor);
        const viewUrl    = a.viewUrl    || getDriveViewUrl(a.valor);
        Swal.fire({
            title           : a.titulo,
            html            : `<div class="doc-preview-wrap"><iframe src="${previewUrl}" frameborder="0" allowfullscreen></iframe></div>`,
            showConfirmButton: false,
            showCloseButton : true,
            width           : '90vw',
            padding         : '0.75rem',
            footer          : `<a href="${viewUrl}" target="_blank" class="text-primary small"><i class="fas fa-external-link-alt me-1"></i>Abrir en Drive</a>`,
            customClass     : { popup: 'swal-doc-popup' }
        });
    }
}

// ============================================
// OTROS — Marca / Área libre
// ============================================

/** Muestra u oculta el input libre cuando se elige "Otros" */
function toggleOtrosInput(field, value) {
    const $div = $(`#${field}OtrosDiv`);
    if (value === 'Otros') $div.show().find('input').focus();
    else                   $div.hide().find('input').val('');
}

/** Devuelve el valor real de Marca (libre si se eligió Otros) */
function getMarcaValue() {
    return $('#marca').val() === 'Otros'
        ? ($('#marcaOtros').val().trim() || 'Otros')
        : $('#marca').val();
}

/** Devuelve el valor real de Área (libre si se eligió Otros) */
function getAreaValue() {
    return $('#area').val() === 'Otros'
        ? ($('#areaOtros').val().trim() || 'Otros')
        : $('#area').val();
}

/**
 * Genera un ID único
 */
function generarId() {
    return Math.floor(100000000 + Math.random() * 900000000).toString();
}

/**
 * Formatea una fecha con hora en zona horaria de Perú (America/Lima, UTC-5)
 * Formato: d/m/yyyy, HH:MM
 */
function formatDate(date) {
    // Convertir a zona horaria de Perú usando toLocaleString
    const peruStr = date.toLocaleString('en-US', { timeZone: 'America/Lima' });
    const peruDate = new Date(peruStr);
    const day = peruDate.getDate();
    const month = peruDate.getMonth() + 1;
    const year = peruDate.getFullYear();
    const hours = String(peruDate.getHours()).padStart(2, '0');
    const minutes = String(peruDate.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year}, ${hours}:${minutes}`;
}

/**
 * Normaliza y muestra cualquier fecha (ISO, string del backend, Date obj)
 * en el formato correcto con zona horaria Perú.
 */
function displayDate(value) {
    if (!value || value === '-') return '-';
    // Si ya tiene el formato correcto (d/m/yyyy, HH:MM) devolverlo tal cual
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(value.toString())) return value;
    // Intentar parsear como fecha ISO u otro formato reconocible
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return formatDate(parsed);
    return value;
}
