/**
 * Configuración de Azure AD (MSAL) y SharePoint.
 */
//#region "Configuración de Azure AD (MSAL) y SharePoint."
const msalConfig = {
    auth: {
        // ID de la aplicación registrada en Entra ID (Azure AD)
        clientId: "a3428d41-17e3-431d-b1fb-212838d61686", // "TU_CLIENT_ID_AQUI"
        // URL del tenant (ej: https://login.microsoftonline.com/TU_TENANT_ID)
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.href,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

const loginRequest = {
    scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/User.Read"]
};

// Datos de SharePoint (Graph API)
const spConfig = {
    siteUrl: "gbg48832240.sharepoint.com",
    sitePath: "/sites/Reservademaquinas",
    listName: "JLG 45 Mantenimiento"
};


//#endregion   

// ==========================================
// CONFIGURACIÓN DE PRESELECCIÓN POR DISPOSITIVO (PERSONALIZACIÓN)
// ==========================================
// Puedes asociar un ID de dispositivo (que verás en la barra superior de la web o en el log)
// con un departamento y una máquina por defecto. Esto es útil para preseleccionar
// opciones en equipos específicos (por ejemplo, PCs de departamentos concretos).
//
// Los valores de 'departamento' y 'maquina' deben coincidir EXACTAMENTE con los atributos 'value' en el HTML.
// Ejemplo de configuración:
// const deviceDefaults = {
//     "mgb-user-1234": { departamento: "Mantenimiento", maquina: "JLG-45 Mantenimiento" },
//     "mgb-user-5678": { departamento: "Limpieza", maquina: "Boom Mantenimiento" }
// };
const deviceDefaults = {
    // Agrega aquí los identificadores de tus PCs y sus opciones por defecto:
    "mgb-user-3277": { departamento: "Audiovisuales Exposiciones" }, // mio
    "mgb-user-6718": { departamento: "Audiovisuales Exposiciones" }  // tambien mio
};

// Genera o recupera un ID único y persistente para este navegador/PC
function getOrCreateDeviceId() {
    let deviceId = localStorage.getItem('mgb_device_id');
    if (!deviceId) {
        const randomNum = Math.floor(1000 + Math.random() * 9000);
        deviceId = `mgb-user-${randomNum}`;
        localStorage.setItem('mgb_device_id', deviceId);
    }
    return deviceId;
}

// Aplica la preselección configurada para este dispositivo
function applyDeviceDefaults() {
    const deviceId = getOrCreateDeviceId();
    const defaults = deviceDefaults[deviceId];
    if (defaults) {
        if (defaults.departamento) {
            const selectDept = document.getElementById('departamentos');
            if (selectDept) selectDept.value = defaults.departamento;
        }
        /*

        if (defaults.maquina) {
            const selectMaq = document.getElementById('maquinas');
            if (selectMaq) selectMaq.value = defaults.maquina;
        }
        */
    }
}

// Registra la sesión de usuario en la lista de SharePoint 'RegistroAccesos' (para funcionar 100% en GitHub Pages)
async function logUserSession(username) {
    if (!graphAccessToken) {
        console.log("No se puede registrar acceso: Token de acceso no disponible.");
        return;
    }
    try {
        const deviceId = getOrCreateDeviceId();
        const timestamp = new Date().toLocaleString('es-ES');

        // Obtener el ID del sitio de SharePoint
        const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${spConfig.siteUrl}:${spConfig.sitePath}`, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });
        if (!siteResponse.ok) {
            console.error("Error al obtener ID del sitio para registro de acceso:", await siteResponse.text());
            return;
        }
        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        // Formamos la línea del log para la columna Title (que viene por defecto)
        const logContent = `Dispositivo: ${deviceId} | Cuenta: ${username} | Fecha: ${timestamp}`;

        const payload = {
            fields: {
                Title: logContent
            }
        };

        const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/RegistroAccesos/items`;

        const response = await fetch(listEndpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${graphAccessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            console.log("Acceso registrado con éxito en SharePoint.");
        } else {
            console.warn("No se pudo registrar el acceso en SharePoint (puede que la lista 'RegistroAccesos' no exista todavía). Código:", response.status);
        }
    } catch (error) {
        console.error("Error al registrar acceso en SharePoint:", error);
    }
}

let myMSALObj;
let graphAccessToken = null;
let currentDate = new Date();
let reservations = [];
let selectedReservationDate = "";

// DOM Elements
// -------------------------------------------------------------------
// Ig: 
// -------------------------------------------------------------------
let selectedMaquina = document.getElementById('maquinas');
selectedMaquina.addEventListener('change', (e) => {
    currentList = e.target.value;
    loadReservations();
});
let selectedDepartamento = document.getElementById('departamentos')
// -------------------------------------------------------------------

const btnLogin = document.getElementById('btn-login');
const btnLogout = document.getElementById('btn-logout');
const userInfo = document.getElementById('user-info');
const datePicker = document.getElementById('date-picker');
const btnPrevDay = document.getElementById('btn-prev-day');
const btnNextDay = document.getElementById('btn-next-day');
const btnRefresh = document.getElementById('btn-refresh');
const timeline = document.getElementById('timeline');
const currentDateDisplay = document.getElementById('current-date-display');

// Modal Elements
const modal = document.getElementById('reservation-modal');
const btnCloseModal = document.getElementById('btn-close-modal');
const btnCancel = document.getElementById('btn-cancel');
const btnSave = document.getElementById('btn-save');
const inputDept = document.getElementById('input-dept');
const inputStart = document.getElementById('input-start');
const inputEnd = document.getElementById('input-end');
const formError = document.getElementById('form-error');

// Initialization: intenta iniciar sesión lo primero
window.onload = () => {
    myMSALObj = new msal.PublicClientApplication(msalConfig);
    setupEventListeners();
    applyDeviceDefaults(); // Aplicar preselección configurada para este dispositivo

    // Las versiones anteriores usaban un pop-up para iniciar sesión y por tanto el usuario debía pulsar el botón.
    // ahora utiliza una redirección y así se puede ejecutar automáticamente lo primero de todo.
    myMSALObj.handleRedirectPromise()
        .then(response => {
            if (response !== null) {
                // Acabamos de volver de un inicio de sesión por redirección exitoso
                showLoggedInView(response.account.username);
                graphAccessToken = response.accessToken;
                initDatePicker();
                loadReservations();
                logUserSession(response.account.username); // Registrar inicio de sesión en SharePoint

                // Iniciar autorefresco
                if (!window.refreshInterval) {
                    window.refreshInterval = setInterval(loadReservations, 60000);
                }
            } else {
                // Carga normal de la página. Comprobar si ya hay una cuenta en caché.
                const currentAccounts = myMSALObj.getAllAccounts();
                if (currentAccounts.length > 0) {
                    showLoggedInView(currentAccounts[0].username);
                    getToken(); // getToken se encargará de llamar a logUserSession una vez obtenido el token silenciosamente
                } else {
                    // Si el usuario no ha cerrado sesión explícitamente, iniciar sesión automáticamente
                    if (sessionStorage.getItem('user_logged_out') !== 'true') {
                        signIn();
                    } else {
                        showLoggedOutView();
                    }
                }
            }
        })
        .catch(error => {
            console.error("Error al procesar la redirección de MSAL:", error);
            showLoggedOutView();
        });
};

// inicializo aquí mis cosas auxiliares
function setupIgn() {
    // mapDptoColor.set("Iluminacion Mantenimiento", "#ff0000");
}

function setupEventListeners() {
    btnLogin.addEventListener('click', signIn);
    btnLogout.addEventListener('click', signOut);

    btnPrevDay.addEventListener('click', () => changeDate(-7));
    btnNextDay.addEventListener('click', () => changeDate(7));
    datePicker.addEventListener('change', (e) => {
        currentDate = new Date(e.target.value);
        updateDateDisplay();
        loadReservations();
    });

    btnRefresh.addEventListener('click', loadReservations);

    btnCloseModal.addEventListener('click', closeModal);
    btnCancel.addEventListener('click', closeModal);
    btnSave.addEventListener('click', saveReservation);
}

// ==========================================
//#region AUTHENTICATION (MSAL)
// ==========================================
function signIn() {
    // Limpiar indicador de logout manual para permitir el flujo
    sessionStorage.removeItem('user_logged_out');

    myMSALObj.loginRedirect(loginRequest)
        .catch(error => {
            console.error("Error al iniciar redirección de login:", error);
            alert("Error de autenticación: " + error.message);
        });
}

function signOut() {
    // Registrar que el usuario ha querido cerrar sesión para evitar bucles de auto-login
    sessionStorage.setItem('user_logged_out', 'true');

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(userInfo.textContent)
    };
    myMSALObj.logoutRedirect(logoutRequest);
}

function getToken() {
    myMSALObj.acquireTokenSilent(loginRequest)
        .then(response => {
            graphAccessToken = response.accessToken;
            initDatePicker();
            loadReservations();
            logUserSession(response.account.username); // Registrar inicio de sesión en SharePoint

            // Loop de refresco automático cada 60 segundos
            if (!window.refreshInterval) {
                window.refreshInterval = setInterval(loadReservations, 60000);
            }
        })
        .catch(error => {
            console.error("Silent token failed:", error);
            if (error instanceof msal.InteractionRequiredAuthError || error.name === "InteractionRequiredAuthError" || error.errorCode === "consent_required") {
                // Si falla el token silencioso y requiere interacción, redirigimos
                myMSALObj.acquireTokenRedirect(loginRequest);
            }
        });
}

function showLoggedInView(username) {
    const deviceId = getOrCreateDeviceId();
    userInfo.innerHTML = `<i class="fa-solid fa-user"></i> ${username} <span style="font-size: 0.85em; opacity: 0.75; margin-left: 8px;">(ID: ${deviceId})</span>`;
    userInfo.classList.remove('hidden');
    btnLogout.classList.remove('hidden');
    btnLogin.classList.add('hidden'); // Ocultar botón de inicio de sesión ya que está correctamente autenticado
}

function showLoggedOutView() {
    const deviceId = getOrCreateDeviceId();
    // Mostramos el ID del dispositivo de forma visible para facilitar la configuración manual al administrador
    userInfo.innerHTML = `<span style="font-size: 0.9em; opacity: 0.75; margin-right: 12px;">ID Dispositivo: <b>${deviceId}</b></span>`;
    userInfo.classList.remove('hidden');
    btnLogout.classList.add('hidden');
    btnLogin.classList.remove('hidden'); // Mostrar el botón para permitir inicio de sesión manual
    timeline.innerHTML = '<div class="loading-spinner">Inicia sesión para ver las reservas.</div>';
}
//#endregion
// ==========================================
//#region UI LOGIC & CALENDAR
// ==========================================
// Helper to get Monday of the week for a given date
function getStartOfWeek(d) {
    const date = new Date(d);
    const day = date.getDay(); // 0 is Sunday, 1 is Monday...
    const diff = date.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(date.setDate(diff));
}

// Helper to get all 7 dates of the week
function getWeekDates(d) {
    const monday = getStartOfWeek(d);
    const dates = [];
    for (let i = 0; i < 7; i++) {
        const day = new Date(monday);
        day.setDate(monday.getDate() + i);
        dates.push(day);
    }
    return dates;
}

// Busca reservas solapadas en la misma semana/día
function findConflicts(resList) {
    const conflicts = [];
    for (let i = 0; i < resList.length; i++) {
        for (let j = i + 1; j < resList.length; j++) {
            const r1 = resList[i];
            const r2 = resList[j];

            const start1 = new Date(r1.start);
            const end1 = new Date(r1.end);
            const start2 = new Date(r2.start);
            const end2 = new Date(r2.end);

            // Si se solapan en tiempo (comenzando o terminando en el mismo rango)
            if (start1 < end2 && start2 < end1) {
                conflicts.push({ r1, r2 });
            }
        }
    }
    return conflicts;
}

// Elimina una reserva de SharePoint
async function deleteReservation(itemId) {
    try {
        if (!graphAccessToken) return false;

        const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${spConfig.siteUrl}:${spConfig.sitePath}`, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });
        if (!siteResponse.ok) return false;
        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        const deleteEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${selectedMaquina.value}/items/${itemId}`;
        const response = await fetch(deleteEndpoint, {
            method: 'DELETE',
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });

        return response.ok;
    } catch (error) {
        console.error("Error al eliminar reserva en conflicto:", error);
        return false;
    }
}

// Muestra la notificación en la parte superior
function showNotification(message) {
    let banner = document.getElementById('notification-banner');
    if (!banner) {
        banner = document.createElement('div');
        banner.id = 'notification-banner';
        banner.className = 'notification-banner';
        const header = document.querySelector('.glass-header');
        if (header) {
            header.after(banner);
        } else {
            document.body.prepend(banner);
        }
    }

    banner.innerHTML = `
        <span><i class="fa-solid fa-triangle-exclamation"></i> ${message}</span>
        <button onclick="this.parentElement.remove()" class="icon-btn" style="color: white; margin-left: 1rem; cursor: pointer; background: transparent; border: none;"><i class="fa-solid fa-xmark"></i></button>
    `;
    banner.classList.remove('hidden');

    // Auto-eliminar después de 10 segundos
    setTimeout(() => {
        if (banner && banner.parentElement) banner.remove();
    }, 10000);
}

// Maneja y resuelve los conflictos eliminando la reserva más nueva
async function handleConflicts(conflicts) {
    const { r1, r2 } = conflicts[0];
    const id1 = parseInt(r1.id) || 0;
    const id2 = parseInt(r2.id) || 0;

    const older = id1 < id2 ? r1 : r2;
    const newer = id1 < id2 ? r2 : r1;

    const startDateTime = new Date(newer.start);
    const startStr = startDateTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const dateStr = startDateTime.toLocaleDateString('es-ES', { day: 'numeric', month: 'short' });

    showNotification(`¡Conflicto detectado! La reserva de "${newer.department}" del ${dateStr} a las ${startStr} se solapaba con una reserva anterior de "${older.department}". Ha sido auto-cancelada automáticamente.`);

    // Eliminar la reserva más reciente
    const success = await deleteReservation(newer.id);
    if (success) {
        // Recargar las reservas
        await loadReservations();
    } else {
        // Si no se puede borrar, renderizar el timeline actual para no buclar
        renderTimeline();
    }
}

function initDatePicker() {
    const today = new Date();
    datePicker.value = today.toISOString().split('T')[0];
    updateDateDisplay();
}

function changeDate(days) {
    currentDate.setDate(currentDate.getDate() + days);
    datePicker.value = currentDate.toISOString().split('T')[0];
    updateDateDisplay();
    loadReservations();
}

function updateDateDisplay() {
    const monday = getStartOfWeek(currentDate);
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    let displayStr = "";
    if (monday.getMonth() === sunday.getMonth()) {
        displayStr = `Semana del ${monday.getDate()} al ${sunday.getDate()} de ${monday.toLocaleDateString('es-ES', { month: 'long' })} de ${monday.getFullYear()}`;
    } else if (monday.getFullYear() === sunday.getFullYear()) {
        displayStr = `Semana del ${monday.getDate()} de ${monday.toLocaleDateString('es-ES', { month: 'long' })} al ${sunday.getDate()} de ${sunday.toLocaleDateString('es-ES', { month: 'long' })} de ${monday.getFullYear()}`;
    } else {
        displayStr = `Semana del ${monday.getDate()} de ${monday.toLocaleDateString('es-ES', { month: 'long' })} de ${monday.getFullYear()} al ${sunday.getDate()} de ${sunday.toLocaleDateString('es-ES', { month: 'long' })} de ${sunday.getFullYear()}`;
    }

    currentDateDisplay.textContent = displayStr;
}

// Asigna un color al dpto indicado de forma directa
function getDeptColorClass(deptName) {
    /*
    const hash = deptName.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    const colorIndex = (hash % 5) + 1;
    return `dept-color-${colorIndex}`;
    */
    // que coñ... es una funcion arrow? desenrrollándolo aquí abajo:
    // let hash = 0;
    // for (let i = 0; i < deptName.length; i++) {
    //     hash += deptName.charCodeAt(i);
    // }
    //const dpto = selectedDepartamento.value; // 0 1 2 3 4
    let dpto;
    // guarro pero funciona
    switch (deptName.toLowerCase()) {
        case "iluminacion mantenimiento":
            dpto = 0;
            break;
        case "iluminacion exposiciones":
            dpto = 1;
            break;
        case "mantenimiento":
            dpto = 2;
            break;
        case "seguridad":
            dpto = 3;
            break;
        case "limpieza":
            dpto = 4;
            break;
        case "montaje exposiciones":
            dpto = 5;
            break;
        case "audiovisuales exposiciones":
            dpto = 6;
            break;
        default:
            dpto = 99;
            break;
    }

    return `dept-color-${dpto}`;


}

// Construir la vista de la línea de tiempo (por horas, vista semanal)
function renderTimeline() {
    timeline.innerHTML = '';

    const weekDates = getWeekDates(currentDate);
    const startHour = 6;
    const endHour = 20;

    const todayStr = new Date().toISOString().split('T')[0];
    const dayNames = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];

    weekDates.forEach((date, index) => {
        const dateStr = date.toISOString().split('T')[0];
        const isToday = (dateStr === todayStr);

        // Crear columna de día
        const dayColumn = document.createElement('div');
        dayColumn.className = `day-column ${isToday ? 'today' : ''}`;

        // Cabecera de la columna
        const dayHeader = document.createElement('div');
        dayHeader.className = 'day-header';

        const dayNameSpan = document.createElement('span');
        dayNameSpan.className = 'day-name';
        dayNameSpan.textContent = dayNames[index];

        const dayDateSpan = document.createElement('span');
        dayDateSpan.className = 'day-date';
        // Formato: 21 may
        const dayDateOptions = { day: 'numeric', month: 'short' };
        dayDateSpan.textContent = date.toLocaleDateString('es-ES', dayDateOptions);

        dayHeader.appendChild(dayNameSpan);
        dayHeader.appendChild(dayDateSpan);
        dayColumn.appendChild(dayHeader);

        // Slots de tiempo para este día
        const slotsContainer = document.createElement('div');
        slotsContainer.className = 'day-slots';

        for (let h = startHour; h <= endHour; h++) {
            const hourStr = h.toString().padStart(2, '0') + ':00';

            // Determinar si esta hora está ocupada
            const overlappingReservation = reservations.find(r => {
                const rStart = new Date(r.start);
                const rEnd = new Date(r.end);

                const rDateStr = rStart.toISOString().split('T')[0];
                if (rDateStr !== dateStr) return false;

                const slotTime = new Date(`${dateStr}T${hourStr}:00`);
                return slotTime >= rStart && slotTime < rEnd;
            });

            const slot = document.createElement('div');
            slot.className = `time-slot ${overlappingReservation ? 'occupied' : 'free'}`;

            const label = document.createElement('span');
            label.className = 'slot-hour';
            label.textContent = hourStr;
            slot.appendChild(label);

            if (overlappingReservation) {
                const deptClass = getDeptColorClass(overlappingReservation.department);
                slot.classList.add(deptClass);

                const deptSpan = document.createElement('span');
                deptSpan.className = 'slot-dept';
                deptSpan.textContent = overlappingReservation.department;
                slot.appendChild(deptSpan);

                // Tooltip con detalles completos
                const rStartStr = new Date(overlappingReservation.start).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                const rEndStr = new Date(overlappingReservation.end).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                slot.title = `${overlappingReservation.department} (${rStartStr} - ${rEndStr})`;
            } else {
                const statusSpan = document.createElement('span');
                statusSpan.className = 'slot-status';
                statusSpan.textContent = 'Libre';
                slot.appendChild(statusSpan);

                // Permitir clic para reservar si está libre
                slot.addEventListener('click', () => openModal(dateStr, hourStr));
            }

            slotsContainer.appendChild(slot);
        }

        dayColumn.appendChild(slotsContainer);
        timeline.appendChild(dayColumn);
    });
}

//#endregion
// ==========================================
//#region DATA FETCHING & SAVING (GRAPH API)
// ==========================================

// Obtener reservas desde SharePoint
async function loadReservations() {
    timeline.innerHTML = '<div class="loading-spinner"><i class="fa-solid fa-circle-notch fa-spin"></i> Cargando reservas...</div>';

    /*
    if (MOCK_MODE && !graphAccessToken) {
        // Datos falsos para pruebas
        setTimeout(() => {
            const today = datePicker.value;
            reservations = [
                { id: '1', department: 'Mantenimiento Mecánico', start: `${today}T08:00:00`, end: `${today}T10:00:00` },
                { id: '2', department: 'Electricidad', start: `${today}T13:00:00`, end: `${today}T15:00:00` }
            ];
            renderTimeline();
        }, 800);
        return;
    }
    */

    try {
        if (!graphAccessToken) {
            throw new Error("No se ha obtenido el Token de Acceso todavía. Por favor, cierra sesión y vuelve a entrar.");
        }

        // OBTENER el Site ID
        const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${spConfig.siteUrl}:${spConfig.sitePath}`, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });

        if (!siteResponse.ok) {
            const errText = await siteResponse.text();
            throw new Error(`Falló al obtener el Sitio (${siteResponse.status}): ${errText}`);
        }

        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        // OBTENER la lista
        const selectedDate = datePicker.value;
        const nextDate = new Date(currentDate);
        nextDate.setDate(nextDate.getDate() + 1);
        const nextDateStr = nextDate.toISOString().split('T')[0];

        //const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${spConfig.listName}/items?expand=fields&$top=999`;
        const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${selectedMaquina.value}/items?expand=fields&$top=999`;

        const response = await fetch(listEndpoint, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });

        if (!response.ok) {
            if (response.status === 400) {
                // Diagnosticar nombres de columnas
                //const colsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${spConfig.listName}/columns`, { headers: { 'Authorization': `Bearer ${graphAccessToken}` } });
                const colsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${selectedMaquina.value}/columns`, { headers: { 'Authorization': `Bearer ${graphAccessToken}` } });
                if (colsRes.ok) {
                    const colsData = await colsRes.json();
                    const availableFields = colsData.value.map(c => `• ${c.displayName} -> <b>${c.name}</b>`).join('<br>');
                    throw new Error(`El nombre de alguna columna es diferente en SharePoint.<br><br><b>Columnas encontradas en tu lista:</b><br>${availableFields}<br><br>Copia este mensaje y pásamelo para que ajuste el código.`);
                }
            }
            const errText = await response.text();
            throw new Error(`Falló al obtener la Lista (${response.status}): ${errText}`);
        }

        const data = await response.json();

        // Mapear los datos de SharePoint a nuestro modelo interno y filtrar localmente
        const weekDates = getWeekDates(currentDate);
        const weekDateStrings = weekDates.map(d => d.toISOString().split('T')[0]);

        reservations = data.value.map(item => {
            // Buscamos las keys de forma case-insensitive por si Graph API cambia las mayúsculas <- LOL
            const f = item.fields;
            const keys = Object.keys(f);
            const titleKey = keys.find(k => k.toLowerCase() === 'title') || 'Title';
            const iniKey = keys.find(k => k.toLowerCase() === 'fechaini') || 'Fechaini';
            const finKey = keys.find(k => k.toLowerCase() === 'fechafin') || 'Fechafin';

            return {
                id: item.id,
                department: f[titleKey] || 'Sin Departamento',
                start: f[iniKey],
                end: f[finKey]
            };
        }).filter(r => {
            // Filtramos las reservas que correspondan a la semana actual
            if (!r.start) return false;
            const rDate = new Date(r.start).toISOString().split('T')[0];
            return weekDateStrings.includes(rDate);
        });

        // Buscar y resolver conflictos de solapamiento
        const conflicts = findConflicts(reservations);
        if (conflicts.length > 0) {
            await handleConflicts(conflicts);
            return;
        }

        renderTimeline();
    } catch (error) {
        console.error("Error al cargar reservas:", error);
        timeline.innerHTML = `<div class="error-message" style="word-break:break-word; max-width:100%; white-space:pre-wrap;"><b>Error detallado de Microsoft intentando acceder a ${selectedMaquina.value}:</b><br/><br/>${error.message}</div>`;
    }
}
//#endregion
// ==========================================
//#region MODAL LOGIC
// ==========================================
function openModal(dateStr, startHourStr) {
    selectedReservationDate = dateStr;
    inputStart.value = startHourStr;

    // Default end time: start time + 1 hour
    const endH = parseInt(startHourStr.split(':')[0]) + 1;
    inputEnd.value = `${endH.toString().padStart(2, '0')}:00`;

    //inputDept.value = '';
    inputDept.value = selectedDepartamento.value;
    formError.classList.add('hidden');

    modal.classList.remove('hidden');
}

function closeModal() {
    modal.classList.add('hidden');
}

function checkOverlap(newStart, newEnd) {
    return reservations.some(r => {
        const rStart = new Date(r.start);
        const rEnd = new Date(r.end);

        // Verificar si hay solapamiento de intervalos de tiempo
        return (newStart < rEnd && newEnd > rStart);
    });
}

async function saveReservation(e) {
    e.preventDefault();

    const selectedDate = selectedReservationDate;
    const startStr = inputStart.value;
    const endStr = inputEnd.value;
    const dept = inputDept.value.trim();

    if (!dept || !startStr || !endStr) return;

    const startDateTime = new Date(`${selectedDate}T${startStr}:00`);
    const endDateTime = new Date(`${selectedDate}T${endStr}:00`);

    if (endDateTime <= startDateTime) {
        formError.textContent = "La hora de fin debe ser posterior a la de inicio.";
        formError.classList.remove('hidden');
        return;
    }

    if (checkOverlap(startDateTime, endDateTime)) {
        formError.textContent = "El horario seleccionado se solapa con una reserva existente.";
        formError.classList.remove('hidden');
        return;
    }

    btnSave.disabled = true;
    btnSave.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Guardando...';

    try {
        // Obtener el ID del sitio nuevamente
        const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${spConfig.siteUrl}:${spConfig.sitePath}`, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });
        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        // Datos para SharePoint
        const payload = {
            fields: {
                Title: dept,
                Fechaini: startDateTime.toISOString(),
                Fechafin: endDateTime.toISOString()
            }
        };

        //const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${spConfig.listName}/items`;
        const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${selectedMaquina.value}/items`;

        const response = await fetch(listEndpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${graphAccessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) throw new Error("Error al guardar en SharePoint");

        closeModal();
        await loadReservations();
    } catch (error) {
        console.error("Error al guardar reserva:", error);
        formError.textContent = "Hubo un error al comunicar con SharePoint.";
        formError.classList.remove('hidden');
    } finally {
        btnSave.disabled = false;
        btnSave.innerHTML = 'Confirmar Reserva';
    }
}
//#endregion