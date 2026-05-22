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
// No hay MODO MOCK (Simulación sin conexión a API)
// Cambia a 'false' cuando configures los IDs reales.
//const MOCK_MODE = true;
// const MOCK_MODE = false;
// ==========================================

let myMSALObj;
let graphAccessToken = null;
let currentDate = new Date();
let reservations = [];
let selectedReservationDate = "";

// DOM Elements
// Ig: 
let selectedMaquina = document.getElementById('maquinas');
selectedMaquina.addEventListener('change', (e) => {
    currentList = e.target.value;
    loadReservations();
});

let selectedDepartamento = document.getElementById('departamentos')

/*
mas

// de aqui cogería el value
<select id="ddlViewBy">
  <option value="1">test1</option>
  <option value="2" selected="selected">test2</option>
  <option value="3">test3</option>
</select>
Running this code:

var e = document.getElementById("ddlViewBy");
var value = e.value;
var text = e.options[e.selectedIndex].text;

*/

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

// Ign
const mapDptoColor = new Map();

// Initialization
window.onload = () => {
    myMSALObj = new msal.PublicClientApplication(msalConfig);
    checkAuthStatus();

    setupEventListeners();

    //alert(selectedMaquina.text);
};

// inicializo aquí mis cosas auxiliares
function setupIgn() {
    mapDptoColor.set("Iluminacion", "#ff0000");
    mapDptoColor.set("Mantenimiento", "#00ff00");
    mapDptoColor.set("Audiovisuales", "#0000ff");
    mapDptoColor.set("Seguridad", "#ff00ff");
    mapDptoColor.set("Informatica", "#00ffff");
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
function checkAuthStatus() {
    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts.length > 0) {
        showLoggedInView(currentAccounts[0].username);
        getToken();
    }
}

function signIn() {
    /*
    if (MOCK_MODE && msalConfig.auth.clientId === "TU_CLIENT_ID_AQUI") {
        showLoggedInView("Usuario de Prueba (Modo Mock)");
        initDatePicker();
        loadReservations();
        return;
    }
    */
    myMSALObj.loginPopup(loginRequest)
        .then(response => {
            showLoggedInView(response.account.username);
            graphAccessToken = response.accessToken;
            initDatePicker();
            loadReservations();

            // Iniciar autorefresco
            if (!window.refreshInterval) {
                window.refreshInterval = setInterval(loadReservations, 60000);
            }
        })
        .catch(error => {
            console.error(error);
            alert("Error de autenticación: " + error.message);
        });
}

function signOut() {
    /*
    if (MOCK_MODE && msalConfig.auth.clientId === "TU_CLIENT_ID_AQUI") {
        userInfo.classList.add('hidden');
        btnLogout.classList.add('hidden');
        btnLogin.classList.remove('hidden');
        timeline.innerHTML = '<div class="loading-spinner">Inicia sesión para ver las reservas.</div>';
        return;
    }
    */
    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(userInfo.textContent)
    };
    myMSALObj.logoutPopup(logoutRequest);
}

function getToken() {
    myMSALObj.acquireTokenSilent(loginRequest)
        .then(response => {
            graphAccessToken = response.accessToken;
            initDatePicker();
            loadReservations();

            // Loop de refresco automático cada 60 segundos
            if (!window.refreshInterval) {
                window.refreshInterval = setInterval(loadReservations, 60000);
            }
        })
        .catch(error => {
            console.error("Silent token failed:", error);
            if (error instanceof msal.InteractionRequiredAuthError || error.name === "InteractionRequiredAuthError" || error.errorCode === "consent_required") {
                myMSALObj.acquireTokenPopup(loginRequest)
                    .then(response => {
                        graphAccessToken = response.accessToken;
                        initDatePicker();
                        loadReservations();
                    }).catch(err => console.error("Popup token failed:", err));
            }
        });
}

function showLoggedInView(username) {
    userInfo.textContent = username;
    userInfo.classList.remove('hidden');
    btnLogout.classList.remove('hidden');
    //btnLogin.classList.add('hidden');
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

    switch (deptName.toLowerCase()) {
        case "iluminacion":
            dpto = 0;
            break;
        case "mantenimiento":
            dpto = 1;
            break;
        case "audiovisuales":
            dpto = 2;
            break;
        case "seguridad":
            dpto = 3;
            break;
        case "informatica":
            dpto = 4;
            break;
        default:
            dpto = 5;
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