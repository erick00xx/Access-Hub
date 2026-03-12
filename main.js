// ============================================
// CONFIGURACIÓN
// ============================================

// URL del Web App desplegado de Google Apps Script
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx52j94Jn2fDzFjNgI6a2IQXaPyyPxMzQqf6VxMnjSxXATVez5VAOn10VZcfGkY5tel/exec"; // Tomar como referencia al archivo Código.js en el repositorio

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

// Sesión activa
let currentUser = null;    // Nombre del perfil activo (null = sin sesión)
let isAdmin = false;    // true = modo administrador sin restricciones
let usuariosData = [];     // Objetos cargados {nombre, pass} desde la hoja
let responsablesData = []; // Nombres cargados desde la hoja "Responsables"

// ============================================
// DATOS DE PRUEBA (para desarrollo sin backend)
// ============================================

const datosDemo = [
    {
        id: '735709281',
        fechaCreacion: '15/3/2026',
        marca: 'IEmpresa',
        responsable: 'Erick | Scott',
        area: 'OSE',
        participantes: 'Roy',
        requerimiento: '<p>Crear un formulario que carece thoth etc etc conectado con bitrix</p>',
        comentarios: '<p>opcional</p>',
        nombreProyecto: '✅ Calendario Metoring',
        accesos: [
            { titulo: 'Admin', valor: 'admin | 123' },
            { titulo: 'Cliente', valor: 'cliente | 123' },
            { titulo: 'OSE', valor: 'ose | 123' }
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

$(document).ready(function () {
    initQuillEditors();
    initDataTable();
    initEventListeners();
    loadResponsablesAndShowLogin();
});

/**
 * Inicializa los editores Quill
 */
function initQuillEditors() {
    const toolbarOptions = [
        ['bold', 'italic', 'underline', 'strike'],
        ['blockquote', 'code-block'],
        [{ 'list': 'ordered' }, { 'list': 'bullet' }],
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
        order: [[0, 'desc']],
        pageLength: 10,
        columns: [
            {
                data: 'nombreProyecto',
                render: function (data, type, row) {
                    const formattedDate = displayDate(row.fechaCreacion);
                    return `<div class="project-cell">
                                <a href="#" class="project-name-link fw-bold text-dark fs-6" style="text-decoration: none;" data-id="${row.id}">${data}</a>
                                <div class="project-date text-muted mt-1" style="font-size: 0.75rem;">Creado el: ${formattedDate}</div>
                            </div>`;
                }
            },
            {
                data: 'marca',
                render: function (data) {
                    return getMarcaBadgeHtml(data);
                }
            },
            {
                data: 'responsable',
                render: function (data) {
                    return getResponsablesBadgesHtml(data);
                }
            },
            { data: 'area' },
            {
                data: 'estado',
                render: function (data, type, row) {
                    let badgeClass = 'badge-en-proceso';
                    if (data === 'Terminado') badgeClass = 'badge-terminado';
                    if (data === 'Rechazado') badgeClass = 'badge-rechazado';
                    if (data === 'Observaciones') badgeClass = 'badge-observaciones';
                    const fechaEst = row.fechaActualizacion ? `<br><span class="estado-fecha">${row.fechaActualizacion}</span>` : '';
                    return `<span class="badge badge-estado ${badgeClass}">${data}</span>${fechaEst}`;
                }
            },
            {
                data: null,
                orderable: false,
                render: function (data, type, row) {
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
    // Función para el template de Select2 con Avatar
    function formatResponsableTemplate(state) {
        if (!state.id) return state.text;
        const name = state.text;
        const initials = name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
        const gradient = getAvatarGradient(name);
        return $(`<div class="d-flex align-items-center gap-2 px-1">
            <div class="s2-user-avatar" style="background:${gradient}">${initials}</div>
            <span class="fw-medium">${escHtml(name)}</span>
        </div>`);
    }

    function formatResponsableSelection(state) {
        if (!state.id) return state.text;
        const name = state.text;
        const initials = name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
        const gradient = getAvatarGradient(name);
        return $(`<div class="d-flex align-items-center gap-1">
            <div class="s2-user-avatar-sm" style="background:${gradient}">${initials}</div>
            <span class="small fw-medium">${escHtml(name)}</span>
        </div>`);
    }

    // Inicializar Select2 para responsables
    $('#responsable').select2({
        theme: 'bootstrap-5',
        placeholder: 'Selecciona al menos un responsable',
        allowClear: true,
        width: '100%',
        templateResult: formatResponsableTemplate,
        templateSelection: formatResponsableSelection
    });

    // Los listeners de unselecting y clearing fueron removidos porque currentUser ya no reside en este input

    // Nuevo proyecto
    $('#btnNuevoProyecto').on('click', nuevoProyecto);

    // Nuevo proyecto
    $('#btnNuevoProyecto').on('click', nuevoProyecto);

    // Guardar proyecto o avanzar al siguiente tab
    $('#btnGuardarProyecto').on('click', guardarOContinuarProyecto);

    // Escuchar el cambio de tabs en el modal para actualizar el botón Siguiente/Guardar
    $('#modalProyectoTabs button[data-bs-toggle="tab"]').on('shown.bs.tab', function (e) {
        updateModalSaveButtonState();
    });

    // Agregar campos dinámicos
    $('#btnAddAcceso').on('click', () => agregarCampoDinamico('accesos'));
    $('#btnAddPlataforma').on('click', () => agregarCampoDinamico('plataformas'));

    // Filtros
    $('#filterMarca, #filterEstado, #filterResponsable').on('change', aplicarFiltros);

    // Click en nombre de proyecto
    $(document).on('click', '.project-name-link', function (e) {
        e.preventDefault();
        const id = $(this).data('id');
        verProyecto(id);
    });

    // Editar desde modal detalle
    $('#btnEditarDesdeDetalle').on('click', function () {
        $('#modalDetalle').modal('hide');
        setTimeout(() => {
            editarProyecto(currentProyectoId);
        }, 300);
    });

    // Limpiar formulario al cerrar modal
    $('#modalProyecto').on('hidden.bs.modal', limpiarFormulario);

    // Toggle comentarios iniciales
    $('#btnToggleComentarios').on('click', function () {
        const $collapse = $('#comentariosCollapse');
        const $icon = $('#iconToggleComentarios');
        const isVisible = $collapse.is(':visible');
        $collapse.toggle(!isVisible);
        $icon.toggleClass('fa-chevron-right', isVisible).toggleClass('fa-chevron-down', !isVisible);
    });

    // Marca y Área con opción "Otros" libre
    $('#marca').on('change', function () { toggleOtrosInput('marca', $(this).val()); });
    $('#area').on('change', function () { toggleOtrosInput('area', $(this).val()); });

    // Sincronizar radio buttons de Estado con el select oculto
    $(document).on('change', 'input[name="estadoRadio"]', function () {
        $('#estado').val($(this).val());
        // Actualizar estilo visual de la opción activa
        $('.estado-radio-option').removeClass('active');
        $(this).closest('.estado-radio-option').addClass('active');
    });

    // Inicializar adjuntos (drag & drop, paste, selección)
    initAdjuntos();
}

// ============================================
// LOGIN / SESIÓN
// ============================================

// Contraseña admin pre-cargada en paralelo
let _cachedAdminPass = null;
// Indica si los proyectos ya están cargados en caché
let _proyectosCargados = false;

/**
 * Carga responsables, proyectos y password de admin en paralelo, luego muestra login.
 * Los proyectos se guardan en caché para no recargarlos al cambiar de perfil.
 */
async function loadResponsablesAndShowLogin() {
    if (DEV_MODE || APPS_SCRIPT_URL === 'TU_URL_DE_APPS_SCRIPT_AQUI') {
        usuariosData = [
            { nombre: 'Erick', pass: '123' },
            { nombre: 'Scott', pass: '123' },
            { nombre: 'Luigui', pass: '123' },
            { nombre: 'Mayra', pass: '123' }
        ];
        responsablesData = usuariosData.map(u => u.nombre);
        _cachedAdminPass = 'admin123';
        // Simular datos de demo
        proyectosData = datosDemo;
        _proyectosCargados = true;
        showLoginScreen();
        return;
    }

    showLoading(true);
    try {
        // Cargar todo en paralelo para minimizar tiempo de espera
        const [respResp, respPass, respProyectos] = await Promise.allSettled([
            fetch(`${APPS_SCRIPT_URL}?action=getResponsables`).then(r => r.json()),
            fetch(`${APPS_SCRIPT_URL}?action=getPassword`).then(r => r.json()),
            fetch(`${APPS_SCRIPT_URL}?action=getProyectos`).then(r => r.json())
        ]);

        // Responsables
        if (respResp.status === 'fulfilled' && respResp.value.success) {
            usuariosData = respResp.value.data || [];
            responsablesData = usuariosData.map(u => u.nombre);
        } else {
            usuariosData = [];
            responsablesData = [];
        }

        // Password admin
        if (respPass.status === 'fulfilled' && respPass.value.success) {
            _cachedAdminPass = respPass.value.data;
        }

        // Proyectos en caché
        if (respProyectos.status === 'fulfilled' && respProyectos.value.success) {
            proyectosData = respProyectos.value.data || [];
            _proyectosCargados = true;
        }
    } catch (e) {
        console.error('Error en carga inicial:', e);
    } finally {
        showLoading(false);
    }
    showLoginScreen();
}

/**
 * Devuelve un gradiente consistente para el avatar según el nombre.
 */
function getAvatarGradient(name) {
    const gradients = [
        ['#2563eb', '#0ea5e9'],
        ['#7c3aed', '#a78bfa'],
        ['#059669', '#34d399'],
        ['#d97706', '#fbbf24'],
        ['#dc2626', '#f87171'],
        ['#0891b2', '#38bdf8'],
        ['#db2777', '#f9a8d4'],
    ];
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash += name.charCodeAt(i);
    const g = gradients[hash % gradients.length];
    return `linear-gradient(135deg, ${g[0]}, ${g[1]})`;
}

/**
 * Genera el HTML de badges para múltiples responsables separados por "|".
 */
function getResponsablesBadgesHtml(responsablesStr) {
    if (!responsablesStr) return '<span class="text-muted small">Sin asignar</span>';

    const responsables = responsablesStr.split('|').map(r => r.trim()).filter(Boolean);

    return responsables.map(name => {
        const initials = name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
        const gradient = getAvatarGradient(name);
        return `
            <div class="d-inline-flex align-items-center gap-1 bg-light border rounded px-2 py-1 me-1 mb-1" style="font-size:0.75rem" title="${escHtml(name)}">
                <div class="s2-user-avatar-sm" style="background:${gradient}">${initials}</div>
                <span class="fw-medium text-dark">${escHtml(name)}</span>
            </div>`;
    }).join('');
}

/**
 * Genera el HTML de un badge para la marca, asegurando que las no predefinidas tengan un color bonito y legible.
 */
function getMarcaBadgeHtml(marca) {
    if (!marca) return '';
    const predefinedBrands = ['IEmpresa', 'JVN', 'Blackwell', 'ITAE', 'Eurocoach', 'Iberoteca', 'Thoth'];

    const upperMarca = marca.toUpperCase();
    const predefinedMatched = predefinedBrands.find(b => b.toUpperCase() === upperMarca);

    if (predefinedMatched) {
        const badgeClass = `badge-${predefinedMatched.toLowerCase().replace(/\./g, '').replace(/ /g, '')}`;
        return `<span class="badge badge-marca ${badgeClass}">${marca}</span>`;
    } else {
        // Generar color consistente basado en el nombre string hash
        let hash = 0;
        for (let i = 0; i < marca.length; i++) {
            hash = marca.charCodeAt(i) + ((hash << 5) - hash);
        }

        // HSL colores pastel lindos
        const h = Math.abs(hash) % 360;
        const s = 70 + (Math.abs(hash) % 20); // 70-90% saturación
        const bgL = 94; // Fondo claro (Lightness)
        const textL = 30; // Texto oscurecido
        const borderL = 86; // Borde intermedio

        return `<span class="badge badge-marca" style="background-color: hsl(${h}, ${s}%, ${bgL}%); color: hsl(${h}, ${s}%, ${textL}%); border: 1px solid hsl(${h}, ${s}%, ${borderL}%);">${marca}</span>`;
    }
}

/**
 * Muestra el overlay de selección de perfil.
 */
function showLoginScreen() {
    const cards = responsablesData.map(name => {
        const initials = name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
        const gradient = getAvatarGradient(name);
        return `
            <div class="login-perfil-card" onclick="selectProfile('${escHtml(name)}')">
                <div class="login-perfil-avatar" style="background:${gradient}">${initials}</div>
                <div class="login-perfil-name">${escHtml(name)}</div>
            </div>`;
    }).join('');
    $('#loginPerfiles').html(cards);
    $('#loginOverlay').addClass('show');
}

/**
 * Inicia el proceso de login genérico.
 */
async function promptLogin(isAdminRole, expectedPass, userName = null) {
    const title = isAdminRole ? 'Acceso Restringido' : `Acceso: ${userName}`;
    const placeholder = isAdminRole ? 'Contraseña de administrador' : 'Ingresa tu contraseña';

    const { value: inputPass } = await Swal.fire({
        title: title,
        input: 'password',
        inputPlaceholder: placeholder,
        showCancelButton: true,
        confirmButtonText: 'Ingresar',
        cancelButtonText: 'Cancelar',
        confirmButtonColor: '#1d4ed8',
        cancelButtonColor: '#64748b',
        width: '340px',
        buttonsStyling: true,
        showClass: { popup: '' },
        hideClass: { popup: '' },
        customClass: { container: 'swal-on-top', popup: 'login-pass-popup', title: 'login-pass-title', input: 'login-pass-input' },
        inputAttributes: { autocomplete: 'new-password', spellcheck: 'false' }
    });

    if (inputPass === undefined) return; // cancelado

    if (inputPass === expectedPass || (expectedPass === '' && inputPass === '')) {
        currentUser = isAdminRole ? null : userName;
        isAdmin = isAdminRole;
        finishLogin();
    } else {
        Swal.fire({
            icon: 'error',
            title: 'Contraseña incorrecta',
            showConfirmButton: false,
            timer: 1600,
            width: '300px',
            showClass: { popup: '' },
            hideClass: { popup: '' }
        });
    }
}

/**
 * Selecciona un perfil de usuario normal y arranca la app.
 */
function selectProfile(name) {
    const user = usuariosData.find(u => u.nombre === name);
    const expectedPass = user ? (user.pass || '').toString().trim() : '';
    promptLogin(false, expectedPass, name);
}

/**
 * Acceso administrador: usa la contraseña pre-cargada al inicio.
 * Si aún no está disponible, hace el fetch en ese momento.
 */
async function loginAsAdmin() {
    // Usar contraseña ya pre-cargada; si falta, obtenerla ahora
    let adminPass = _cachedAdminPass || 'admin123';
    if (!_cachedAdminPass && !DEV_MODE && APPS_SCRIPT_URL !== 'TU_URL_DE_APPS_SCRIPT_AQUI') {
        try {
            const resp = await fetch(`${APPS_SCRIPT_URL}?action=getPassword`);
            const result = await resp.json();
            if (result.success) {
                adminPass = result.data.toString().trim();
                _cachedAdminPass = adminPass;
            }
        } catch (e) {
            console.error('Error obteniendo contraseña admin:', e);
        }
    }

    promptLogin(true, adminPass);
}

/**
 * Finaliza el login: configura UI y muestra la app.
 * Si los proyectos ya están en caché no los vuelve a pedir al servidor.
 */
function finishLogin() {
    populateFormResponsable();
    updateNavUser();
    $('#loginOverlay').removeClass('show');
    if (_proyectosCargados) {
        // Datos ya disponibles: solo actualizar stats/filtros/tabla
        updateStats();
        populateFilters();
        aplicarFiltros();
    } else {
        loadProyectos();
    }
}

/**
 * Rellena dinámicamente el <select id="responsable"> con los nombres de la hoja.
 */
function populateFormResponsable() {
    const $sel = $('#responsable');
    const $fijo = $('#responsableFijo');
    const $asterisk = $('#responsableAsterisk');
    $sel.empty();

    if (!isAdmin && currentUser) {
        const initials = currentUser.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
        const gradient = getAvatarGradient(currentUser);
        $fijo.html(`
            <div class="d-inline-flex align-items-center gap-1 bg-light border rounded px-2 py-1" style="font-size:0.75rem">
                <div class="s2-user-avatar-sm" style="background:${gradient}">${initials}</div>
                <span class="fw-semibold text-primary">${escHtml(currentUser)}</span>
            </div>
        `).removeClass('d-none');
        $sel.removeAttr('required');
        $asterisk.addClass('d-none');
    } else {
        $fijo.addClass('d-none').empty();
        $sel.attr('required', 'required');
        $asterisk.removeClass('d-none');
    }

    responsablesData.forEach(name => {
        if (!isAdmin && currentUser && name === currentUser) return;
        $sel.append(`<option value="${escHtml(name)}">${escHtml(name)}</option>`);
    });
}

/**
 * Actualiza el indicador de usuario en la navbar.
 */
function updateNavUser() {
    if (isAdmin) {
        $('#navUserName').text('Administrador');
        $('#navAdminBadge').removeClass('d-none');
    } else {
        $('#navUserName').text(currentUser);
        $('#navAdminBadge').addClass('d-none');
    }
    $('#navUserInfo').removeClass('d-none');
}

/**
 * Vuelve a la pantalla de selección de perfil.
 * NO recarga proyectos: los datos ya están en caché.
 */
function cambiarPerfil() {
    currentUser = null;
    isAdmin = false;
    $('#navUserInfo').addClass('d-none');
    showLoginScreen();
    // Nota: al seleccionar perfil, finishLogin() usará _proyectosCargados=true
    // y solo re-aplicará filtros, sin petición al servidor.
}

/**
 * Aplica las restricciones de UI según el perfil activo (bloquea filtro de responsable).
 */
function applyUserRestrictions() {
    if (!isAdmin && currentUser) {
        $('#filterResponsable').val(currentUser).prop('disabled', true);
        $('#filterResponsableLabel').html(
            `Mi perfil <i class="fas fa-lock ms-1 text-muted" style="font-size:0.65rem" title="Filtro bloqueado a tu perfil"></i>`
        );
    } else {
        $('#filterResponsable').prop('disabled', false).val('');
        $('#filterResponsableLabel').text('Filtrar por Responsable');
    }
}

// ============================================
// WIZARD TABS LOGIC
// ============================================

/**
 * Lógica para el botón Siguiente/Guardar.
 * Si no estamos en el último tab ("Estado"), valida y avanza.
 * Si estamos en el último, guarda el proyecto.
 */
function guardarOContinuarProyecto() {
    const tabsList = ['#tab-general', '#tab-accesos', '#tab-estado'];
    const activeTabButton = document.querySelector('#modalProyectoTabs button.active');
    const targetId = activeTabButton ? activeTabButton.getAttribute('data-bs-target') : '#tab-general';

    // Obtener el input/select con error visible usando Form API native functionality
    if (!$('#formProyecto')[0].checkValidity()) {
        const invalidFields = $('#formProyecto :invalid');
        if (invalidFields.length > 0) {
            // Find which tab the first invalid element belongs to
            const firstInvalid = invalidFields.first();
            const parentTab = firstInvalid.closest('.tab-pane');
            if (parentTab && parentTab.attr('id') !== targetId.substring(1)) {
                $(`#modalProyectoTabs button[data-bs-target="#${parentTab.attr('id')}"]`).tab('show');
                // Allow Bootstrap to animate the tab transition
                setTimeout(() => $('#formProyecto')[0].reportValidity(), 200);
            } else {
                $('#formProyecto')[0].reportValidity();
            }
        }
        return;
    }

    const currentIndex = tabsList.indexOf(targetId);

    // Si NO estamos en el último tab (índice 2, "Estado")
    if (currentIndex < tabsList.length - 1) {
        // En "Datos Generales", validamos si el "Responsable" múltiple está lleno (para Admins)
        if (currentIndex === 0 && isAdmin) {
            const responsablesSeleccionados = Array.from($('#responsable')[0].selectedOptions).map(o => o.value);
            if (responsablesSeleccionados.length === 0) {
                showError('Por favor selecciona al menos un responsable.');
                return;
            }
        }

        // Validar campos "Otros" si están activos en pestaña 1
        if (currentIndex === 0) {
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
        }

        // Avanzar al siguiente tab
        const nextTarget = tabsList[currentIndex + 1];
        $(`#modalProyectoTabs button[data-bs-target="${nextTarget}"]`).tab('show');
    } else {
        // Ejecutar guardado final si estamos en Estado
        guardarProyecto();
    }
}

/**
 * Actualiza el texto y apariencia del botón para leer "Siguiente" o "Guardar Proyecto".
 */
function updateModalSaveButtonState() {
    const $btn = $('#btnGuardarProyecto');
    const activeTab = document.querySelector('#modalProyectoTabs button.active');
    const target = activeTab ? activeTab.getAttribute('data-bs-target') : '#tab-general';

    if (target === '#tab-estado') {
        $btn.html('<i class="fas fa-save me-1"></i>Guardar Proyecto');
        $btn.removeClass('btn-info text-white').addClass('btn-primary');
    } else {
        // Mantener bootstrap consistent colors pero cambiar contenido
        $btn.html('Siguiente <i class="fas fa-chevron-right ms-1"></i>');
        $btn.removeClass('btn-primary').addClass('btn-info text-white');
    }
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
            updateStats();
            populateFilters();
            aplicarFiltros();
            showLoading(false);
        }, 800);
        return;
    }

    // Llamada real al backend
    $.ajax({
        url: APPS_SCRIPT_URL,
        method: 'GET',
        data: { action: 'getProyectos' },
        success: function (response) {
            if (response.success) {
                proyectosData = response.data;
                _proyectosCargados = true;
                updateStats();
                populateFilters();
                aplicarFiltros();
            } else {
                showError('Error al cargar proyectos: ' + response.message);
            }
        },
        error: function (xhr, status, error) {
            console.error("Detalle del error:", xhr, status, error);
            showError('Error de conexión: ' + error);
        },
        complete: function () {
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
    // perfil activo ya se maneja fuera del select cuando no es admin
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
    // Responsable: soporte multi (separado por " | ")
    let responsables = (proyecto.responsable || '').split(' | ').map(r => r.trim()).filter(Boolean);

    // Si no es admin y hay usuario, se excluye del select2 porque ya aparece fijo en el span
    if (!isAdmin && currentUser) {
        responsables = responsables.filter(r => r !== currentUser);
    }
    $('#responsable').val(responsables).trigger('change');

    $('#participantes').val(proyecto.participantes);
    $('#estado').val(proyecto.estado);
    // Sincronizar radio buttons con el estado cargado
    $('input[name="estadoRadio"]').prop('checked', false);
    $(`input[name="estadoRadio"][value="${proyecto.estado}"]`).prop('checked', true);
    $('.estado-radio-option').removeClass('active');
    $(`input[name="estadoRadio"][value="${proyecto.estado}"]`).closest('.estado-radio-option').addClass('active');
    $('#comentariosEstado').val(proyecto.comentariosEstado);

    // Marca – soporte de opción libre
    const marcasPredefinidas = ['IEmpresa', 'JVN', 'Blackwell', 'ITAE', 'Eurocoach', 'Iberoteca', 'Thoth', 'Otros', ''];
    if (marcasPredefinidas.includes(proyecto.marca)) {
        $('#marca').val(proyecto.marca);
        $('#marcaOtrosDiv').hide().find('input').val('');
    } else {
        $('#marca').val('Otros');
        $('#marcaOtros').val(proyecto.marca);
        $('#marcaOtrosDiv').show();
    }

    // Área – soporte de opción libre
    const areasPredefinidas = ['OSE', 'DAD', '.ID.', 'Otros', ''];
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
    // Mostrar comentarios si tienen contenido
    if (proyecto.comentarios && proyecto.comentarios.replace(/<[^>]*>/g, '').trim()) {
        $('#comentariosCollapse').show();
        $('#iconToggleComentarios').removeClass('fa-chevron-right').addClass('fa-chevron-down');
    } else {
        $('#comentariosCollapse').hide();
        $('#iconToggleComentarios').removeClass('fa-chevron-down').addClass('fa-chevron-right');
    }

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
            const fileId = getDriveFileId(img.valor);
            const isImg = isDriveImageUrl(img.valor);
            const thumbUrl = fileId
                ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`
                : (isImg ? img.valor : null);
            const viewUrl = fileId ? `https://drive.google.com/file/d/${fileId}/view` : img.valor;
            const previewUrl = fileId ? `https://drive.google.com/file/d/${fileId}/preview` : img.valor;
            adjuntosData.push({
                id: 'adj_' + (++adjuntoIdCounter),
                titulo: img.titulo,
                valor: thumbUrl || img.valor,  // guardar thumbnailUrl para imágenes
                thumbnail: thumbUrl,
                thumbnailUrl: thumbUrl,
                viewUrl,
                previewUrl,
                fileId,
                isImage: isImg,
                type: isImg ? 'image' : 'other',
                uploading: false,
                error: false
            });
        });
    }
    updateAdjuntosGrid();

    actualizarVisibilidadNoItems();

    // Activa la primera pestaña al abrir modal
    $('#tab-general-btn').tab('show');
    updateModalSaveButtonState();

    $('#modalProyecto').modal('show');
}

/**
 * Guarda el proyecto (crear o actualizar)
 * Note: CheckValidity y saltos manuales se manejan en guardarOContinuarProyecto.
 */
async function guardarProyecto() {
    const proyecto = {
        id: $('#proyectoId').val() || generarId(),
        fechaCreacion: $('#proyectoId').val() ?
            proyectosData.find(p => p.id === $('#proyectoId').val())?.fechaCreacion :
            formatDate(new Date()),
        marca: getMarcaValue(),
        responsable: (!isAdmin && currentUser)
            ? [currentUser, ...Array.from($('#responsable')[0].selectedOptions).map(o => o.value)].join(' | ')
            : Array.from($('#responsable')[0].selectedOptions).map(o => o.value).join(' | '),
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
        accesosHtml = `<div class="acceso-grid">${proyecto.accesos.map((a, idx) => {
            const partes = (a.valor || '').split(' | ');
            const usuario = partes[0] ? partes[0].trim() : '';
            const pass = partes.slice(1).join(' | ').trim();
            const uid = `accp_${proyecto.id}_${idx}`;
            const uEsc = usuario.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
            const pEsc = pass.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
            return `
            <div class="acceso-card">
                <div class="acceso-title-bar">
                    <i class="fas fa-key me-2"></i>${a.titulo}
                </div>
                <div class="acceso-fields">
                    <div class="acceso-field-row">
                        <span class="acceso-field-label"><i class="fas fa-user me-1"></i>Usuario</span>
                        <div class="acceso-field-value-wrap">
                            <code class="acceso-field-value">${usuario || '-'}</code>
                            ${usuario ? `<button type="button" class="btn-acceso-action" data-v="${uEsc}" onclick="copiarAlPortapapeles(this.dataset.v, this)" title="Copiar usuario"><i class="fas fa-copy"></i></button>` : ''}
                        </div>
                    </div>
                    <div class="acceso-field-row">
                        <span class="acceso-field-label"><i class="fas fa-lock me-1"></i>Contraseña</span>
                        <div class="acceso-field-value-wrap">
                            <code class="acceso-field-value pass-hidden" id="${uid}" data-pass="${pEsc}">${pass ? '••••••••' : '-'}</code>
                            ${pass ? `
                            <button type="button" class="btn-acceso-action btn-toggle-pass" data-target="${uid}" onclick="toggleAccesoPassword(this)" title="Ver/ocultar contraseña"><i class="fas fa-eye"></i></button>
                            <button type="button" class="btn-acceso-action" data-target="${uid}" onclick="copiarPassSpan(this)" title="Copiar contraseña"><i class="fas fa-copy"></i></button>` : ''}
                        </div>
                    </div>
                </div>
            </div>`;
        }).join('')}</div>`;
    } else {
        accesosHtml = '<p class="text-muted small mb-0"><i class="fas fa-info-circle me-1"></i>No hay accesos registrados</p>';
    }

    let plataformasHtml = '';
    if (proyecto.plataformas && proyecto.plataformas.length > 0) {
        plataformasHtml = `<div class="plataforma-grid">${proyecto.plataformas.map((p, pi) => {
            const urlEsc = p.valor.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
            return `
            <div class="plataforma-card">
                <span class="plataforma-label">${p.titulo}</span>
                <div class="d-flex align-items-center gap-2">
                    <a class="plataforma-link" href="${p.valor}" target="_blank"><i class="fas fa-external-link-alt me-1"></i>${p.valor}</a>
                    <button type="button" class="btn-acceso-action" data-v="${urlEsc}" onclick="copiarAlPortapapeles(this.dataset.v, this)" title="Copiar URL"><i class="fas fa-copy"></i></button>
                </div>
            </div>`;
        }).join('')}</div>`;
    } else {
        plataformasHtml = '<p class="text-muted small mb-0"><i class="fas fa-info-circle me-1"></i>No hay plataformas registradas</p>';
    }

    const imagenesHtml = renderAdjuntosDetalle(proyecto.imagenes);

    const html = `
        <div class="detalle-proyecto">

            <!-- Info general del proyecto -->
            <div class="detalle-info-grid mb-4">
                ${proyecto.marca ? `
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-tag me-1"></i>Marca</span>
                    <span class="detalle-info-value" style="font-size: 1.1em;">${getMarcaBadgeHtml(proyecto.marca)}</span>
                </div>` : ''}
                <div class="detalle-info-item">
                    <span class="detalle-info-label"><i class="fas fa-user me-1"></i>Responsable</span>
                    <div class="mt-1">${getResponsablesBadgesHtml(proyecto.responsable)}</div>
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
                                <span class="detalle-meta-label"><i class="fas fa-calendar-check me-1 text-warning"></i>Fecha de cambio de estado</span>
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
    let html = '';

    if (tipo === 'accesos') {
        // Separar usuario y contraseña del campo valor (formato: "usuario | pass")
        const partes = valor.split(' | ');
        const usuarioVal = partes[0] ? partes[0].trim() : '';
        const passVal = partes.slice(1).join(' | ').trim();

        html = `
        <div class="dynamic-field" data-index="${index}">
            <button type="button" class="btn btn-outline-danger btn-sm btn-remove" onclick="removerCampoDinamico(this, '${tipo}')">
                <i class="fas fa-times"></i>
            </button>
            <div class="row g-2 align-items-center">
                <div class="col-sm-3">
                    <div class="input-group input-group-sm">
                        <span class="input-group-text"><i class="fas fa-key"></i></span>
                        <input type="text" class="form-control campo-titulo" placeholder="Tipo (Admin, Cliente...)" value="${titulo}">
                    </div>
                </div>
                <div class="col-sm-4">
                    <div class="input-group input-group-sm">
                        <span class="input-group-text"><i class="fas fa-user"></i></span>
                        <input type="text" class="form-control campo-usuario" placeholder="Usuario" value="${usuarioVal}" autocomplete="off">
                    </div>
                </div>
                <div class="col-sm-5">
                    <div class="input-group input-group-sm">
                        <span class="input-group-text"><i class="fas fa-lock"></i></span>
                        <input type="password" class="form-control campo-password" placeholder="Contraseña" value="${passVal}" autocomplete="new-password">
                        <button type="button" class="btn btn-outline-secondary" onclick="togglePasswordField(this)" tabindex="-1">
                            <i class="fas fa-eye"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>`;
    } else {
        let placeholderTitulo = 'Título';
        let placeholderValor = 'URL o valor';
        let iconClass = 'fa-link';

        if (tipo === 'plataformas') {
            placeholderTitulo = 'Ej: Github, Bitrix, Loocker';
            placeholderValor = 'https://';
            iconClass = 'fa-globe';
        } else if (tipo === 'imagenes') {
            placeholderTitulo = 'Ej: Diagrama, Wireframe, Mockup';
            placeholderValor = 'URL de la imagen';
            iconClass = 'fa-image';
        }

        html = `
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
        </div>`;
    }

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
    $(`#${tipo}Container .dynamic-field`).each(function () {
        const titulo = $(this).find('.campo-titulo').val().trim();
        if (tipo === 'accesos') {
            const usuario = $(this).find('.campo-usuario').val().trim();
            const password = $(this).find('.campo-password').val().trim();
            if (titulo || usuario || password) {
                items.push({ titulo, valor: `${usuario} | ${password}` });
            }
        } else {
            const valor = $(this).find('.campo-valor').val().trim();
            if (titulo || valor) {
                items.push({ titulo, valor });
            }
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
    // Resetear radio buttons de estado
    $('input[name="estadoRadio"]').prop('checked', false);
    $('input[name="estadoRadio"][value="En proceso"]').prop('checked', true);
    $('.estado-radio-option').removeClass('active');
    $('input[name="estadoRadio"][value="En proceso"]').closest('.estado-radio-option').addClass('active');
    $('#estado').val('En proceso');
    quillRequerimiento.root.innerHTML = '';
    quillComentarios.root.innerHTML = '';
    $('#accesosContainer').empty();
    $('#plataformasContainer').empty();
    adjuntosData = [];
    updateAdjuntosGrid();
    actualizarVisibilidadNoItems();
    currentProyectoId = null;
    $('#marcaOtrosDiv, #areaOtrosDiv').hide().find('input').val('');
    // Ocultar comentarios y resetear
    $('#comentariosCollapse').hide();
    $('#iconToggleComentarios').removeClass('fa-chevron-down').addClass('fa-chevron-right');
    // Resetear responsable multi-select y Select2
    $('#responsable').val(null).trigger('change');
    // Volver a la primera pestaña
    $('#tab-general-btn').tab('show');
}

/**
 * Actualiza las estadísticas
 */
function updateStats() {
    // Para no-admin, las estadísticas reflejan solo sus proyectos
    const statsData = (!isAdmin && currentUser)
        ? proyectosData.filter(p =>
            (p.responsable || '').split(' | ').map(r => r.trim()).includes(currentUser))
        : proyectosData;

    $('#statTotal').text(statsData.length);
    $('#statEnProceso').text(statsData.filter(p => p.estado === 'En proceso').length);
    $('#statTerminados').text(statsData.filter(p => p.estado === 'Terminado').length);
    $('#statObservaciones').text(statsData.filter(p => p.estado === 'Observaciones').length);
    $('#statRechazados').text(statsData.filter(p => p.estado === 'Rechazado').length);
}

/**
 * Popula los filtros con valores únicos
 */
function populateFilters() {
    // Marcas (desde los datos reales)
    const marcas = [...new Set(proyectosData.map(p => p.marca).filter(Boolean))];
    const filterMarca = $('#filterMarca');
    filterMarca.find('option:not(:first)').remove();
    marcas.forEach(marca => {
        filterMarca.append(`<option value="${marca}">${marca}</option>`);
    });

    // Responsables desde la hoja "Responsables" (no desde los datos de proyectos)
    const filterResponsable = $('#filterResponsable');
    filterResponsable.find('option:not(:first)').remove();
    responsablesData.forEach(name => {
        filterResponsable.append(`<option value="${name}">${name}</option>`);
    });

    // Bloquear / desbloquear según perfil
    applyUserRestrictions();
}

/**
 * Aplica los filtros a la tabla
 */
function aplicarFiltros() {
    const marca = $('#filterMarca').val();
    const estado = $('#filterEstado').val();
    // Para no-admin, siempre usar su propio perfil como filtro de responsable
    const responsable = (!isAdmin && currentUser) ? currentUser : $('#filterResponsable').val();

    let filteredData = proyectosData;

    if (marca) {
        filteredData = filteredData.filter(p => p.marca === marca);
    }
    if (estado) {
        filteredData = filteredData.filter(p => p.estado === estado);
    }
    if (responsable) {
        filteredData = filteredData.filter(p =>
            (p.responsable || '').split(' | ').map(r => r.trim()).includes(responsable)
        );
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
        const fileId = getDriveFileId(img.valor);
        const isImg = isDriveImageUrl(img.valor);
        const thumbUrl = fileId
            ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w600`
            : (isImg ? img.valor : null);
        return {
            titulo: img.titulo,
            valor: img.valor,
            fileId,
            isImage: isImg,
            thumbnailUrl: thumbUrl,
            thumbnail: thumbUrl,
            viewUrl: fileId ? `https://drive.google.com/file/d/${fileId}/view` : img.valor,
            previewUrl: fileId ? `https://drive.google.com/file/d/${fileId}/preview` : img.valor,
            type: isImg ? 'image' : 'other'
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
                <div class="adjunto-detalle-label adjunto-hover-label">${escHtml(a.titulo)}</div>
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
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
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
    const icons = { image: 'fa-image', pdf: 'fa-file-pdf', doc: 'fa-file-word', xls: 'fa-file-excel', other: 'fa-file' };
    const colors = { image: 'text-info', pdf: 'text-danger', doc: 'text-primary', xls: 'text-success', other: 'text-secondary' };
    return `${icons[type] || icons.other} ${colors[type] || colors.other}`;
}

/**
 * Procesa un array de File objects: crea entrada en adjuntosData y sube a Drive
 */
async function handleFiles(files) {
    for (const file of files) {
        const id = 'adj_' + (++adjuntoIdCounter);
        const type = getFileType(file);
        const tempUrl = URL.createObjectURL(file);

        adjuntosData.push({
            id, type, uploading: true, error: false,
            titulo: file.name.replace(/\.[^/.]+$/, ''),
            valor: tempUrl,
            thumbnail: type === 'image' ? tempUrl : null,  // blob URL para preview instantáneo
            thumbnailUrl: null,
            viewUrl: null,
            previewUrl: null,
            fileId: null,
            isImage: type === 'image',
            _tempUrl: tempUrl
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
                    valor: result.valor,
                    thumbnailUrl: result.thumbnailUrl,
                    viewUrl: result.viewUrl,
                    previewUrl: result.previewUrl,
                    fileId: result.fileId,
                    isImage: result.isImage,
                    // Usar thumbnail Drive URL para imagen (reemplaza blob URL)
                    thumbnail: result.isImage ? (result.thumbnailUrl || result.valor) : null,
                    uploading: false,
                    _tempUrl: null
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
        return {
            valor: tempUrl, thumbnailUrl: isImage ? tempUrl : null,
            viewUrl: tempUrl, previewUrl: tempUrl, fileId: null, isImage
        };
    }
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const base64Data = e.target.result.split(',')[1];
                const resp = await fetch(APPS_SCRIPT_URL, {
                    method: 'POST',
                    redirect: 'follow',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ action: 'uploadImage', fileName: file.name, base64Data, mimeType: file.type })
                });
                const result = await resp.json();
                if (result.success) {
                    resolve({
                        valor: result.data.url,           // thumbnailUrl para imágenes, viewUrl para docs
                        thumbnailUrl: result.data.thumbnailUrl,
                        viewUrl: result.data.viewUrl,
                        previewUrl: result.data.previewUrl,
                        fileId: result.data.fileId,
                        isImage: result.data.isImage
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
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
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
            title: '',
            imageUrl: imgUrl,
            imageAlt: a.titulo,
            showConfirmButton: false,
            showCloseButton: true,
            imageWidth: '100%',
            footer: viewUrl
                ? `<a href="${viewUrl}" target="_blank" class="text-primary small"><i class="fas fa-external-link-alt me-1"></i>Abrir en Drive</a>`
                : '',
            customClass: { popup: 'swal-image-popup' }
        });
    } else {
        const previewUrl = a.previewUrl || getDrivePreviewUrl(a.valor);
        const viewUrl = a.viewUrl || getDriveViewUrl(a.valor);
        Swal.fire({
            title: a.titulo,
            html: `<div class="doc-preview-wrap"><iframe src="${previewUrl}" frameborder="0" allowfullscreen></iframe></div>`,
            showConfirmButton: false,
            showCloseButton: true,
            width: '90vw',
            padding: '0.75rem',
            footer: `<a href="${viewUrl}" target="_blank" class="text-primary small"><i class="fas fa-external-link-alt me-1"></i>Abrir en Drive</a>`,
            customClass: { popup: 'swal-doc-popup' }
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
    else $div.hide().find('input').val('');
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

// ============================================
// UTILIDADES DE ACCESOS / CREDENCIALES
// ============================================

/**
 * Alterna visibilidad de contraseña en tarjeta de detalle
 */
function toggleAccesoPassword(btn) {
    const span = document.getElementById(btn.dataset.target);
    const hidden = span.classList.contains('pass-hidden');
    span.textContent = hidden ? span.dataset.pass : '••••••••';
    span.classList.toggle('pass-hidden', !hidden);
    btn.querySelector('i').className = hidden ? 'fas fa-eye-slash' : 'fas fa-eye';
}

/**
 * Copia la contraseña del span referenciado
 */
function copiarPassSpan(btn) {
    const span = document.getElementById(btn.dataset.target);
    copiarAlPortapapeles(span.dataset.pass, btn);
}

/**
 * Copia texto al portapapeles con feedback visual
 */
function copiarAlPortapapeles(text, btn) {
    const doFeedback = () => {
        const icon = btn.querySelector('i');
        const orig = icon.className;
        icon.className = 'fas fa-check';
        btn.classList.add('btn-acceso-copied');
        setTimeout(() => {
            icon.className = orig;
            btn.classList.remove('btn-acceso-copied');
        }, 1600);
    };
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(doFeedback);
    } else {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        doFeedback();
    }
}

/**
 * Alterna visibilidad de contraseña en el formulario
 */
function togglePasswordField(btn) {
    const input = btn.closest('.input-group').querySelector('input');
    const icon = btn.querySelector('i');
    if (input.type === 'password') {
        input.type = 'text';
        icon.className = 'fas fa-eye-slash';
    } else {
        input.type = 'password';
        icon.className = 'fas fa-eye';
    }
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
