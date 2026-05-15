/**
 * Configuración de Azure AD (MSAL) y SharePoint.
 * IMPORTANTE: DEBES REEMPLAZAR ESTOS VALORES CON LOS DE TU ENTORNO.
 */
const msalConfig = {
    auth: {
        // ID de la aplicación registrada en Entra ID (Azure AD)
        //clientId: "a3428d41-17e3-431d-b1fb-212838d61686",
        clientId: "TU_CLIENT_ID_AQUI",
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
    scopes: ["Sites.ReadWrite.All", "User.Read"]
};

// Datos de SharePoint (Graph API)
const spConfig = {
    siteUrl: "gbg48832240.sharepoint.com",
    sitePath: "/sites/Reservademaquinas",
    listName: "JLG 45 Mantenimiento"
};

// ==========================================
// MODO MOCK (Simulación sin conexión a API)
// Cambia a 'false' cuando configures los IDs reales.
const MOCK_MODE = true;
// ==========================================

let myMSALObj;
let graphAccessToken = null;
let currentDate = new Date();
let reservations = [];

// DOM Elements
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

// Initialization
window.onload = () => {
    // Inicializar MSAL solo si no estamos en Mock Mode estricto o si hay un ClientID
    if (msalConfig.auth.clientId !== "TU_CLIENT_ID_AQUI") {
        myMSALObj = new msal.PublicClientApplication(msalConfig);
        checkAuthStatus();
    } else if (MOCK_MODE) {
        // Simular login para pruebas de UI
        showLoggedInView("Usuario de Prueba (Modo Mock)");
        initDatePicker();
        loadReservations();
    } else {
        alert("Por favor, configura el clientId en app.js para usar la autenticación real.");
    }

    setupEventListeners();
};

function setupEventListeners() {
    btnLogin.addEventListener('click', signIn);
    btnLogout.addEventListener('click', signOut);

    btnPrevDay.addEventListener('click', () => changeDate(-1));
    btnNextDay.addEventListener('click', () => changeDate(1));
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
// AUTHENTICATION (MSAL)
// ==========================================
function checkAuthStatus() {
    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts.length > 0) {
        showLoggedInView(currentAccounts[0].username);
        getToken();
    }
}

function signIn() {
    if (MOCK_MODE && msalConfig.auth.clientId === "TU_CLIENT_ID_AQUI") {
        showLoggedInView("Usuario de Prueba (Modo Mock)");
        initDatePicker();
        loadReservations();
        return;
    }

    myMSALObj.loginPopup(loginRequest)
        .then(response => {
            showLoggedInView(response.account.username);
            getToken();
        })
        .catch(error => {
            console.error(error);
            alert("Error de autenticación: " + error.message);
        });
}

function signOut() {
    if (MOCK_MODE && msalConfig.auth.clientId === "TU_CLIENT_ID_AQUI") {
        userInfo.classList.add('hidden');
        btnLogout.classList.add('hidden');
        btnLogin.classList.remove('hidden');
        timeline.innerHTML = '<div class="loading-spinner">Inicia sesión para ver las reservas.</div>';
        return;
    }

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
            setInterval(loadReservations, 60000);
        })
        .catch(error => {
            if (error instanceof msal.InteractionRequiredAuthError) {
                myMSALObj.acquireTokenPopup(loginRequest)
                    .then(response => {
                        graphAccessToken = response.accessToken;
                        initDatePicker();
                        loadReservations();
                    });
            }
        });
}

function showLoggedInView(username) {
    userInfo.textContent = username;
    userInfo.classList.remove('hidden');
    btnLogout.classList.remove('hidden');
    btnLogin.classList.add('hidden');
}

// ==========================================
// UI LOGIC & CALENDAR
// ==========================================
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
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    currentDateDisplay.textContent = currentDate.toLocaleDateString('es-ES', options);
}

// Genera un color consistente basado en el nombre del departamento
function getDeptColorClass(deptName) {
    const hash = deptName.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    const colorIndex = (hash % 5) + 1;
    return `dept-color-${colorIndex}`;
}

// Construir la vista de la línea de tiempo (por horas)
function renderTimeline() {
    timeline.innerHTML = '';

    // Configurar horas de trabajo: 06:00 a 22:00
    const startHour = 8;
    const endHour = 20;

    const selectedDateStr = datePicker.value;

    for (let h = startHour; h <= endHour; h++) {
        const hourStr = h.toString().padStart(2, '0') + ':00';

        // Determinar si esta hora está ocupada
        // Una reserva ocupa la hora si su inicio <= h y su fin > h
        const overlappingReservation = reservations.find(r => {
            const rStart = new Date(r.start);
            const rEnd = new Date(r.end);

            // Comprobar si la reserva es el mismo día
            const rDateStr = rStart.toISOString().split('T')[0];
            if (rDateStr !== selectedDateStr) return false;

            // La lógica simplificada: si la hora evaluada cae dentro de la reserva
            const slotTime = new Date(`${selectedDateStr}T${hourStr}:00`);
            return slotTime >= rStart && slotTime < rEnd;
        });

        const slot = document.createElement('div');
        slot.className = `time-slot ${overlappingReservation ? 'occupied' : 'free'}`;

        const label = document.createElement('div');
        label.className = 'time-label';
        label.textContent = hourStr;

        const content = document.createElement('div');
        content.className = 'slot-content';

        if (overlappingReservation) {
            const deptClass = getDeptColorClass(overlappingReservation.department);
            content.classList.add(deptClass);

            const rStartStr = new Date(overlappingReservation.start).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            const rEndStr = new Date(overlappingReservation.end).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

            content.innerHTML = `
                <div class="reservation-details">
                    <span class="dept-name">${overlappingReservation.department}</span>
                    <span class="time-range">${rStartStr} - ${rEndStr}</span>
                </div>
                <i class="fa-solid fa-lock"></i>
            `;
        } else {
            // Permitir clic para reservar si está libre
            slot.addEventListener('click', () => openModal(hourStr));
        }

        slot.appendChild(label);
        slot.appendChild(content);
        timeline.appendChild(slot);
    }
}

// ==========================================
// DATA FETCHING & SAVING (GRAPH API)
// ==========================================

// Obtener reservas desde SharePoint
async function loadReservations() {
    timeline.innerHTML = '<div class="loading-spinner"><i class="fa-solid fa-circle-notch fa-spin"></i> Cargando reservas...</div>';

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

    try {
        // OBTENER el Site ID
        const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${spConfig.siteUrl}:${spConfig.sitePath}`, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });
        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        // OBTENER la lista
        // (Nota: Nombres de columnas internos. Fecha_x0020_ini, Fecha_x0020_fin, Departamento)
        const selectedDate = datePicker.value;
        const nextDate = new Date(currentDate);
        nextDate.setDate(nextDate.getDate() + 1);
        const nextDateStr = nextDate.toISOString().split('T')[0];

        // Filtramos por la fecha seleccionada usando OData en Graph
        const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${spConfig.listName}/items?expand=fields(select=Departamento,Fecha_x0020_ini,Fecha_x0020_fin)&$filter=fields/Fecha_x0020_ini ge '${selectedDate}T00:00:00Z' and fields/Fecha_x0020_ini lt '${nextDateStr}T00:00:00Z'`;

        const response = await fetch(listEndpoint, {
            headers: { 'Authorization': `Bearer ${graphAccessToken}` }
        });

        if (!response.ok) throw new Error("Error obteniendo datos");

        const data = await response.json();

        // Mapear los datos de SharePoint a nuestro modelo interno
        reservations = data.value.map(item => ({
            id: item.id,
            department: item.fields.Departamento,
            start: item.fields.Fecha_x0020_ini,
            end: item.fields.Fecha_x0020_fin
        }));

        renderTimeline();
    } catch (error) {
        console.error("Error al cargar reservas:", error);
        timeline.innerHTML = `<div class="error-message">Error cargando datos. Asegúrate de tener permisos o revisa la consola.</div>`;
    }
}

// ==========================================
// MODAL LOGIC
// ==========================================
function openModal(startHourStr) {
    inputStart.value = startHourStr;

    // Default end time: start time + 1 hour
    const endH = parseInt(startHourStr.split(':')[0]) + 1;
    inputEnd.value = `${endH.toString().padStart(2, '0')}:00`;

    inputDept.value = '';
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

    const selectedDate = datePicker.value;
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

    if (MOCK_MODE && !graphAccessToken) {
        // Simular guardado
        setTimeout(() => {
            reservations.push({
                id: Math.random().toString(),
                department: dept,
                start: startDateTime.toISOString(),
                end: endDateTime.toISOString()
            });
            closeModal();
            renderTimeline();
            btnSave.disabled = false;
            btnSave.innerHTML = 'Confirmar Reserva';
        }, 1000);
        return;
    }

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
                Departamento: dept,
                Fecha_x0020_ini: startDateTime.toISOString(),
                Fecha_x0020_fin: endDateTime.toISOString()
            }
        };

        const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${spConfig.listName}/items`;

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
