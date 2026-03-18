// Global variables
let chartInstances = {};
let currentUser = null;
let balanceViewMode = 'monthly';
let latestProcessedRows = [];
let activeDashboardMenu = 'laporan';
let lastMetricSnapshot = {};
let googleSheetReportCache = {
    url: '',
    rows: [],
    isLoading: false
};
let monthlyDistributionState = {
    months: [],
    allChannelData: {},
    allRevenueData: {},
    channelByMonth: new Map(),
    revenueByMonth: new Map(),
    channelIndex: -1,
    revenueIndex: -1,
    colors: []
};
let pendingUserAction = {
    type: '',
    username: ''
};
const USERS_STORAGE_KEY = 'appUsers';
const THEME_STORAGE_KEY = 'dashboardTheme';
const DATE_PARSING_MODE_STORAGE_KEY = 'dateParsingMode';
const MANUAL_INPUT_DATA_KEY = 'manualInputDataset';
const MANUAL_INPUT_HISTORY_KEY = 'manualInputHistory';
const ACTIVITY_LOG_STORAGE_KEY = 'activityLogs';
const MAX_ACTIVITY_LOG_ITEMS = 500;
let dateParsingMode = 'auto';
let activityLogGlobalHandlersBound = false;
const MANUAL_TRANSACTION_CHANNELS = [
    { key: 'qris', label: 'Qris' },
    { key: 'mobileBanking', label: 'Mobile Banking' },
    { key: 'edc', label: 'EDC' },
    { key: 'atm', label: 'ATM' },
    { key: 'teller', label: 'Teller' }
];
const MANUAL_REPORT_FIELDS = [
    {
        key: 'laporanRKUD',
        transaksiKey: 'transaksiRKUD',
        label: 'Laporan RKUD',
        inputId: 'manualLaporanRKUD',
        editInputId: 'manualEditLaporanRKUD',
        transaksiInputId: 'manualTransaksiRKUD',
        editTransaksiInputId: 'manualEditTransaksiRKUD',
        transaksiChannelInputIds: {
            qris: 'manualQrisRKUD',
            mobileBanking: 'manualMobileBankingRKUD',
            edc: 'manualEDCRKUD',
            atm: 'manualATMRKUD',
            teller: 'manualTellerRKUD'
        },
        editTransaksiChannelInputIds: {
            qris: 'manualEditQrisRKUD',
            mobileBanking: 'manualEditMobileBankingRKUD',
            edc: 'manualEditEDCRKUD',
            atm: 'manualEditATMRKUD',
            teller: 'manualEditTellerRKUD'
        }
    },
    {
        key: 'laporanPDRD',
        transaksiKey: 'transaksiPDRD',
        label: 'Laporan PDRD',
        inputId: 'manualLaporanPDRD',
        editInputId: 'manualEditLaporanPDRD',
        transaksiInputId: 'manualTransaksiPDRD',
        editTransaksiInputId: 'manualEditTransaksiPDRD',
        transaksiChannelInputIds: {
            qris: 'manualQrisPDRD',
            mobileBanking: 'manualMobileBankingPDRD',
            edc: 'manualEDCPDRD',
            atm: 'manualATMPDRD',
            teller: 'manualTellerPDRD'
        },
        editTransaksiChannelInputIds: {
            qris: 'manualEditQrisPDRD',
            mobileBanking: 'manualEditMobileBankingPDRD',
            edc: 'manualEditEDCPDRD',
            atm: 'manualEditATMPDRD',
            teller: 'manualEditTellerPDRD'
        }
    },
    {
        key: 'laporanBRIBPHTB',
        transaksiKey: 'transaksiBRIBPHTB',
        label: 'Laporan BRI-BPHTB',
        inputId: 'manualLaporanBRIBPHTB',
        editInputId: 'manualEditLaporanBRIBPHTB',
        transaksiInputId: 'manualTransaksiBRIBPHTB',
        editTransaksiInputId: 'manualEditTransaksiBRIBPHTB',
        transaksiChannelInputIds: {
            qris: 'manualQrisBRIBPHTB',
            mobileBanking: 'manualMobileBankingBRIBPHTB',
            edc: 'manualEDCBRIBPHTB',
            atm: 'manualATMBRIBPHTB',
            teller: 'manualTellerBRIBPHTB'
        },
        editTransaksiChannelInputIds: {
            qris: 'manualEditQrisBRIBPHTB',
            mobileBanking: 'manualEditMobileBankingBRIBPHTB',
            edc: 'manualEditEDCBRIBPHTB',
            atm: 'manualEditATMBRIBPHTB',
            teller: 'manualEditTellerBRIBPHTB'
        }
    },
    {
        key: 'laporanBRIPBB',
        transaksiKey: 'transaksiBRIPBB',
        label: 'Laporan BRI-PBB',
        inputId: 'manualLaporanBRIPBB',
        editInputId: 'manualEditLaporanBRIPBB',
        transaksiInputId: 'manualTransaksiBRIPBB',
        editTransaksiInputId: 'manualEditTransaksiBRIPBB',
        transaksiChannelInputIds: {
            qris: 'manualQrisBRIPBB',
            mobileBanking: 'manualMobileBankingBRIPBB',
            edc: 'manualEDCBRIPBB',
            atm: 'manualATMBRIPBB',
            teller: 'manualTellerBRIPBB'
        },
        editTransaksiChannelInputIds: {
            qris: 'manualEditQrisBRIPBB',
            mobileBanking: 'manualEditMobileBankingBRIPBB',
            edc: 'manualEditEDCBRIPBB',
            atm: 'manualEditATMBRIPBB',
            teller: 'manualEditTellerBRIPBB'
        }
    },
    {
        key: 'laporanMandiri',
        transaksiKey: 'transaksiMandiri',
        label: 'Laporan Mandiri',
        inputId: 'manualLaporanMandiri',
        editInputId: 'manualEditLaporanMandiri',
        transaksiInputId: 'manualTransaksiMandiri',
        editTransaksiInputId: 'manualEditTransaksiMandiri',
        transaksiChannelInputIds: {
            qris: 'manualQrisMandiri',
            mobileBanking: 'manualMobileBankingMandiri',
            edc: 'manualEDCMandiri',
            atm: 'manualATMMandiri',
            teller: 'manualTellerMandiri'
        },
        editTransaksiChannelInputIds: {
            qris: 'manualEditQrisMandiri',
            mobileBanking: 'manualEditMobileBankingMandiri',
            edc: 'manualEditEDCMandiri',
            atm: 'manualEditATMMandiri',
            teller: 'manualEditTellerMandiri'
        }
    }
];
const DEFAULT_USERS = [
    {
        username: 'admin',
        password: 'admin123',
        role: 'admin',
        fullName: 'Administrator',
        photoUrl: '',
        permissions: ['view', 'upload', 'download', 'manage']
    },
    {
        username: 'user',
        password: 'user123',
        role: 'user',
        fullName: 'Regular User',
        photoUrl: '',
        permissions: ['view', 'upload']
    },
    {
        username: 'guest',
        password: 'guest123',
        role: 'guest',
        fullName: 'Guest User',
        photoUrl: '',
        permissions: ['view']
    }
];

function getStoredUsers() {
    try {
        const rawUsers = localStorage.getItem(USERS_STORAGE_KEY);
        if (!rawUsers) {
            localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(DEFAULT_USERS));
            return [...DEFAULT_USERS];
        }

        const parsedUsers = JSON.parse(rawUsers);
        if (!Array.isArray(parsedUsers) || parsedUsers.length === 0) {
            localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(DEFAULT_USERS));
            return [...DEFAULT_USERS];
        }

        return parsedUsers;
    } catch (error) {
        localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(DEFAULT_USERS));
        return [...DEFAULT_USERS];
    }
}

function saveStoredUsers(users) {
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users || []));
}

function getPermissionsByRole(role) {
    if (role === 'admin') return ['view', 'upload', 'download', 'manage'];
    if (role === 'user') return ['view', 'upload'];
    return ['view'];
}

function isAdminUser() {
    return Boolean(currentUser && currentUser.role === 'admin');
}

function getAllowedMenusByRole(role) {
    if (role === 'admin') return ['sorotan', 'grafik', 'laporan', 'manual', 'user', 'activity-log'];
    if (role === 'user') return ['sorotan', 'grafik', 'laporan', 'password'];
    return ['laporan', 'password'];
}

function canAccessMenu(menuKey) {
    if (!currentUser) return false;
    return getAllowedMenusByRole(currentUser.role).includes(menuKey);
}

function getFallbackMenuForCurrentRole() {
    const allowedMenus = getAllowedMenusByRole(currentUser?.role);
    return allowedMenus[0] || 'laporan';
}

function getMenuLabel(menuKey) {
    const labelMap = {
        sorotan: 'Sorotan Utama',
        grafik: 'Grafik Operasional',
        laporan: 'Laporan Transaksi',
        manual: 'Input Data Manual',
        user: 'Pengelolaan User',
        password: 'Ubah Password',
        'activity-log': 'Log Aktivitas'
    };

    return labelMap[menuKey] || 'Menu';
}

function normalizeActivityLogLevel(level) {
    const safeLevel = String(level || '').trim().toLowerCase();
    if (['info', 'warning', 'error', 'success'].includes(safeLevel)) {
        return safeLevel;
    }
    return 'info';
}

function getActivityLogs() {
    try {
        const raw = localStorage.getItem(ACTIVITY_LOG_STORAGE_KEY);
        if (!raw) return [];

        const parsed = JSON.parse(raw);
        if (!Array.isArray(parsed)) return [];

        return parsed
            .filter((log) => log && typeof log === 'object')
            .map((log) => ({
                id: String(log.id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
                timestamp: String(log.timestamp || new Date().toISOString()),
                level: normalizeActivityLogLevel(log.level),
                category: String(log.category || 'sistem').trim() || 'sistem',
                event: String(log.event || 'Peristiwa').trim() || 'Peristiwa',
                message: String(log.message || '').trim(),
                user: String(log.user || '-').trim() || '-',
                source: String(log.source || 'Dashboard').trim() || 'Dashboard'
            }));
    } catch (error) {
        return [];
    }
}

function saveActivityLogs(logs) {
    try {
        const safeLogs = Array.isArray(logs) ? logs.slice(0, 10) : [];
        localStorage.setItem(ACTIVITY_LOG_STORAGE_KEY, JSON.stringify(safeLogs));
    } catch (error) {
        console.error('Gagal menyimpan log aktivitas:', error);
    }
}

function addActivityLog(entry) {
    const logs = getActivityLogs();
    logs.unshift(entry);
    // Batasi hanya 10 log terbaru
    saveActivityLogs(logs.slice(0, 10));
}

function logActivityEvent(eventName, message, options = {}) {
    const nextEntry = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        timestamp: new Date().toISOString(),
        level: normalizeActivityLogLevel(options.level || 'info'),
        category: String(options.category || 'aktivitas').trim() || 'aktivitas',
        event: String(eventName || 'Peristiwa').trim() || 'Peristiwa',
        message: String(message || '').trim(),
        user: currentUser?.username || '-',
        source: String(options.source || 'Dashboard').trim() || 'Dashboard'
    };

    addActivityLog(nextEntry);
}

function escapeLogHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function formatActivityLogTimestamp(timestamp) {
    const parsedDate = new Date(String(timestamp || ''));
    if (!Number.isFinite(parsedDate.getTime())) return '-';
    return parsedDate.toLocaleString('id-ID');
}

function showActivityLogStatus(message, type) {
    const statusElement = document.getElementById('activityLogStatus');
    if (!statusElement) return;

    statusElement.textContent = String(message || '');
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = 'block';

    setTimeout(() => {
        statusElement.style.display = 'none';
    }, 4000);
}

function renderActivityLogs() {
    const tableBody = document.getElementById('activityLogTableBody');
    const emptyState = document.getElementById('activityLogEmpty');
    if (!tableBody || !emptyState) return;

    let logs = getActivityLogs().sort((left, right) => {
        return new Date(right.timestamp).getTime() - new Date(left.timestamp).getTime();
    });
    // Tampilkan hanya 10 log terbaru
    logs = logs.slice(0, 10);

    if (!logs.length) {
        tableBody.innerHTML = '';
        emptyState.style.display = 'block';
        return;
    }

    const rowsHtml = logs.map((log, index) => {
        const level = normalizeActivityLogLevel(log.level);
        const levelLabelMap = {
            info: 'INFO',
            warning: 'WARNING',
            error: 'ERROR',
            success: 'SUCCESS'
        };

        return `
            <tr>
                <td>${index + 1}</td>
                <td>${escapeLogHtml(formatActivityLogTimestamp(log.timestamp))}</td>
                <td><span class="activity-log-badge level-${level}">${levelLabelMap[level] || 'INFO'}</span></td>
                <td>${escapeLogHtml(log.category || '-')}</td>
                <td>${escapeLogHtml(log.event || '-')}</td>
                <td class="activity-log-message">${escapeLogHtml(log.message || '-')}</td>
                <td>${escapeLogHtml(log.user || '-')}</td>
                <td>${escapeLogHtml(log.source || '-')}</td>
            </tr>
        `;
    }).join('');

    tableBody.innerHTML = rowsHtml;
    emptyState.style.display = 'none';
}

function clearActivityLogs() {
    if (!isAdminUser()) {
        showActivityLogStatus('❌ Hanya admin yang dapat menghapus log aktivitas.', 'error');
        return;
    }

    if (!confirm('Hapus semua log aktivitas yang tersimpan?')) {
        return;
    }

    saveActivityLogs([]);
    renderActivityLogs();
    showActivityLogStatus('✅ Log aktivitas berhasil dihapus.', 'success');
}

function exportActivityLogs(format = 'json') {
    if (!isAdminUser()) {
        showActivityLogStatus('❌ Hanya admin yang dapat mengekspor log aktivitas.', 'error');
        return;
    }

    const logs = getActivityLogs();
    if (!logs.length) {
        showActivityLogStatus('⚠️ Belum ada data log untuk diekspor.', 'warning');
        return;
    }

    if (String(format || '').toLowerCase() !== 'json') {
        showActivityLogStatus('❌ Format export belum didukung.', 'error');
        return;
    }

    try {
        const payload = JSON.stringify(logs, null, 2);
        const blob = new Blob([payload], { type: 'application/json' });
        const link = document.createElement('a');
        const timestamp = new Date().toISOString().replace(/[.:]/g, '-');
        const objectUrl = URL.createObjectURL(blob);

        link.href = objectUrl;
        link.download = `activity-log-${timestamp}.json`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(objectUrl);

        showActivityLogStatus('✅ Export log aktivitas berhasil.', 'success');
    } catch (error) {
        showActivityLogStatus('❌ Gagal export log aktivitas.', 'error');
    }
}

function initGlobalErrorLogging() {
    if (activityLogGlobalHandlersBound) return;
    activityLogGlobalHandlersBound = true;

    window.addEventListener('error', (event) => {
        const source = String(event.filename || 'window').split('/').pop();
        const location = event.lineno ? `:${event.lineno}${event.colno ? `:${event.colno}` : ''}` : '';
        const message = String(event.message || 'Unknown error');

        logActivityEvent('JavaScript Error', `${message} (${source}${location})`, {
            level: 'error',
            category: 'error',
            source: source || 'window'
        });
    });

    window.addEventListener('unhandledrejection', (event) => {
        const reason = typeof event.reason === 'string'
            ? event.reason
            : (event.reason?.message || 'Unhandled promise rejection');

        logActivityEvent('Promise Rejection', String(reason || 'Unhandled promise rejection'), {
            level: 'error',
            category: 'error',
            source: 'promise'
        });
    });
}

function showUserManageStatus(message, type) {
    const statusElement = document.getElementById('userManageStatus');
    if (!statusElement) return;
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = 'block';
}

function showPasswordChangeStatus(message, type) {
    const statusElement = document.getElementById('passwordChangeStatus');
    if (!statusElement) return;
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = 'block';
}

function resetPasswordChangeForm() {
    const currentPasswordInput = document.getElementById('currentPasswordInput');
    const newPasswordInput = document.getElementById('newPasswordInput');
    const confirmNewPasswordInput = document.getElementById('confirmNewPasswordInput');
    if (currentPasswordInput) currentPasswordInput.value = '';
    if (newPasswordInput) newPasswordInput.value = '';
    if (confirmNewPasswordInput) confirmNewPasswordInput.value = '';
}

function changeOwnPassword() {
    if (!currentUser) {
        showPasswordChangeStatus('❌ Sesi login tidak ditemukan.', 'error');
        return;
    }

    if (!['user', 'guest'].includes(String(currentUser.role || ''))) {
        showPasswordChangeStatus('❌ Menu ini hanya untuk role user dan guest.', 'error');
        return;
    }

    const currentPasswordInput = document.getElementById('currentPasswordInput');
    const newPasswordInput = document.getElementById('newPasswordInput');
    const confirmNewPasswordInput = document.getElementById('confirmNewPasswordInput');
    if (!currentPasswordInput || !newPasswordInput || !confirmNewPasswordInput) return;

    const currentPassword = String(currentPasswordInput.value || '').trim();
    const newPassword = String(newPasswordInput.value || '').trim();
    const confirmNewPassword = String(confirmNewPasswordInput.value || '').trim();

    if (!currentPassword || !newPassword || !confirmNewPassword) {
        showPasswordChangeStatus('❌ Semua field password wajib diisi.', 'error');
        return;
    }

    if (newPassword.length < 6 || newPassword.length > 64) {
        showPasswordChangeStatus('❌ Password baru harus 6-64 karakter.', 'error');
        return;
    }

    if (newPassword !== confirmNewPassword) {
        showPasswordChangeStatus('❌ Konfirmasi password baru tidak sama.', 'error');
        return;
    }

    const users = getStoredUsers();
    const userIndex = users.findIndex(user => user.username === currentUser.username);
    if (userIndex < 0) {
        showPasswordChangeStatus('❌ Data akun tidak ditemukan.', 'error');
        return;
    }

    if (String(users[userIndex].password || '') !== currentPassword) {
        showPasswordChangeStatus('❌ Password lama tidak sesuai.', 'error');
        return;
    }

    if (currentPassword === newPassword) {
        showPasswordChangeStatus('❌ Password baru harus berbeda dari password lama.', 'error');
        return;
    }

    users[userIndex].password = newPassword;
    saveStoredUsers(users);
    resetPasswordChangeForm();
    showPasswordChangeStatus('✅ Password login berhasil diperbarui.', 'success');
}

function onManagedUserChange() {
    const userSelect = document.getElementById('manageUserSelect');
    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const roleSelect = document.getElementById('manageRole');
    if (!userSelect || !usernameInput || !passwordInput || !fullNameInput || !roleSelect) return;

    const users = getStoredUsers();
    const selectedUser = users.find(user => user.username === userSelect.value);
    if (!selectedUser) return;

    usernameInput.value = selectedUser.username || '';
    passwordInput.value = selectedUser.password || '';
    fullNameInput.value = selectedUser.fullName || '';
    roleSelect.value = selectedUser.role || 'guest';
}

function renderUserManagement() {
    const userSelect = document.getElementById('manageUserSelect');
    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const roleSelect = document.getElementById('manageRole');

    const users = getStoredUsers();
    if (userSelect) {
        const fixedSelectableUsers = ['admin', 'user', 'guest'];
        const managedUsers = fixedSelectableUsers
            .map((username) => users.find((user) => user.username === username))
            .filter(Boolean);

        const previousSelectedUsername = userSelect.value;
        userSelect.innerHTML = managedUsers
            .map(user => `<option value="${user.username}">${user.username} (${user.role})</option>`)
            .join('');

        if (managedUsers.length > 0) {
            const selectedExists = managedUsers.some(user => user.username === previousSelectedUsername);
            userSelect.value = selectedExists ? previousSelectedUsername : managedUsers[0].username;
            onManagedUserChange();
        } else if (usernameInput && passwordInput && fullNameInput && roleSelect) {
            usernameInput.value = '';
            passwordInput.value = '';
            fullNameInput.value = '';
            roleSelect.value = 'guest';
        }
    } else if (usernameInput && passwordInput && fullNameInput && roleSelect) {
        usernameInput.value = '';
        passwordInput.value = '';
        fullNameInput.value = '';
        roleSelect.value = 'guest';
    }

    renderRegisteredUsersTable(users);
}

function renderRegisteredUsersTable(users) {
    const tableBody = document.getElementById('registeredUsersTableBody');
    const emptyState = document.getElementById('registeredUsersEmpty');
    if (!tableBody || !emptyState) return;

    const safeUsers = Array.isArray(users) ? users : [];
    const escapeHtml = (value) => String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    const escapeJsString = (value) => String(value || '')
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/\r/g, '\\r')
        .replace(/\n/g, '\\n');

    if (safeUsers.length === 0) {
        tableBody.innerHTML = '';
        emptyState.style.display = 'block';
        return;
    }

    const rowsHtml = safeUsers.map((user, index) => {
        const permissions = Array.isArray(user.permissions) && user.permissions.length > 0
            ? user.permissions.join(', ')
            : '-';

        return `
            <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(user.username)}</td>
                <td>${escapeHtml(user.fullName || user.username)}</td>
                <td>${escapeHtml(user.role || 'guest')}</td>
                <td>${escapeHtml(permissions)}</td>
                <td>
                    <div style="display:flex; gap:6px; flex-wrap:wrap;">
                        <button type="button" onclick="requestManagedUserEdit('${escapeJsString(user.username)}')" style="padding:6px 10px; font-size:0.78em;">✏️ Edit</button>
                        <button type="button" onclick="deleteManagedUser('${escapeJsString(user.username)}')" style="padding:6px 10px; font-size:0.78em; background:#dc2626;">🗑️ Hapus</button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');

    tableBody.innerHTML = rowsHtml;
    emptyState.style.display = 'none';
}

function closeManagedUserEditPopup(event) {
    if (event && event.target && event.currentTarget && event.target !== event.currentTarget) {
        return;
    }

    const popup = document.getElementById('userEditPopup');
    const originalInput = document.getElementById('editPopupOriginalUsername');
    if (popup) {
        popup.style.display = 'none';
    }
    if (originalInput) {
        originalInput.value = '';
    }
}

function openManagedUserEditPopup(targetUsername) {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat mengubah data user.', 'error');
        return;
    }

    const userSelect = document.getElementById('manageUserSelect');
    const selectedUsername = String(targetUsername || userSelect?.value || '').trim();
    if (!selectedUsername) {
        showUserManageStatus('❌ Pilih user yang akan diubah.', 'error');
        return;
    }

    const users = getStoredUsers();
    const selectedUser = users.find(user => user.username === selectedUsername);
    if (!selectedUser) {
        showUserManageStatus('❌ Data user tidak ditemukan.', 'error');
        return;
    }

    const popup = document.getElementById('userEditPopup');
    const originalInput = document.getElementById('editPopupOriginalUsername');
    const usernameInput = document.getElementById('editPopupUsername');
    const passwordInput = document.getElementById('editPopupPassword');
    const fullNameInput = document.getElementById('editPopupFullName');
    const roleInput = document.getElementById('editPopupRole');

    if (!popup || !originalInput || !usernameInput || !passwordInput || !fullNameInput || !roleInput) {
        return;
    }

    originalInput.value = selectedUser.username || '';
    usernameInput.value = selectedUser.username || '';
    passwordInput.value = selectedUser.password || '';
    fullNameInput.value = selectedUser.fullName || '';
    roleInput.value = selectedUser.role || 'guest';

    popup.style.display = 'flex';
}

function saveManagedUserData(selectedUsername, newUsername, newPassword, newFullName, newRole) {
    const targetUsername = String(selectedUsername || '').trim();
    const updatedUsername = String(newUsername || '').trim();
    const updatedPassword = String(newPassword || '').trim();
    const updatedFullName = String(newFullName || '').trim();
    const updatedRole = String(newRole || 'guest').trim();

    if (!targetUsername) {
        showUserManageStatus('❌ Pilih user yang akan diubah.', 'error');
        return false;
    }

    if (!updatedUsername || !updatedPassword) {
        showUserManageStatus('❌ Username dan password baru wajib diisi.', 'error');
        return false;
    }

    const users = getStoredUsers();
    const userIndex = users.findIndex(user => user.username === targetUsername);
    if (userIndex < 0) {
        showUserManageStatus('❌ Data user tidak ditemukan.', 'error');
        return false;
    }

    const duplicateUser = users.find((user, index) => index !== userIndex && user.username === updatedUsername);
    if (duplicateUser) {
        showUserManageStatus('❌ Username sudah dipakai user lain.', 'error');
        return false;
    }

    const oldUsername = users[userIndex].username;
    users[userIndex].username = updatedUsername;
    users[userIndex].password = updatedPassword;
    users[userIndex].fullName = updatedFullName || users[userIndex].fullName;
    users[userIndex].role = ['admin', 'user', 'guest'].includes(updatedRole) ? updatedRole : 'guest';
    users[userIndex].permissions = getPermissionsByRole(users[userIndex].role);

    saveStoredUsers(users);

    if (currentUser && currentUser.username === oldUsername) {
        currentUser.username = users[userIndex].username;
        currentUser.fullName = users[userIndex].fullName;
        currentUser.photoUrl = users[userIndex].photoUrl;
        currentUser.role = users[userIndex].role;
        currentUser.permissions = users[userIndex].permissions;
        localStorage.setItem('userSession', JSON.stringify({
            ...currentUser,
            username: users[userIndex].username,
            fullName: users[userIndex].fullName,
            photoUrl: users[userIndex].photoUrl,
            role: users[userIndex].role,
            permissions: users[userIndex].permissions
        }));
        updateUserDisplay();
        applyPermissions();
    }

    renderUserManagement();
    const refreshedSelect = document.getElementById('manageUserSelect');
    if (refreshedSelect) {
        refreshedSelect.value = updatedUsername;
        onManagedUserChange();
    }
    showUserManageStatus('✅ Data user berhasil diperbarui.', 'success');
    return true;
}

function saveManagedUserEditPopup() {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat mengubah data user.', 'error');
        return;
    }

    const originalInput = document.getElementById('editPopupOriginalUsername');
    const usernameInput = document.getElementById('editPopupUsername');
    const passwordInput = document.getElementById('editPopupPassword');
    const fullNameInput = document.getElementById('editPopupFullName');
    const roleInput = document.getElementById('editPopupRole');
    if (!originalInput || !usernameInput || !passwordInput || !fullNameInput || !roleInput) return;

    const isSaved = saveManagedUserData(
        originalInput.value,
        usernameInput.value,
        passwordInput.value,
        fullNameInput.value,
        roleInput.value
    );

    if (isSaved) {
        closeManagedUserEditPopup();
    }
}

function selectManagedUserForEdit(username) {
    openManagedUserEditPopup(username);
}

function requestManagedUserEdit(targetUsername) {
    openManagedUserEditPopup(targetUsername);
}

function updateManagedUserCredentials() {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat mengubah data user.', 'error');
        return;
    }

    const userSelect = document.getElementById('manageUserSelect');
    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const roleSelect = document.getElementById('manageRole');
    if (!userSelect || !usernameInput || !passwordInput || !fullNameInput || !roleSelect) return;

    saveManagedUserData(
        userSelect.value,
        usernameInput.value,
        passwordInput.value,
        fullNameInput.value,
        roleSelect.value
    );
}

function closeUserActionPopup(event) {
    if (event && event.target && event.currentTarget && event.target !== event.currentTarget) {
        return;
    }

    const popup = document.getElementById('userActionPopup');
    if (popup) {
        popup.style.display = 'none';
    }

    pendingUserAction = {
        type: '',
        username: ''
    };
}

function openUserActionPopup(actionType, targetUsername) {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat mengelola data user.', 'error');
        return;
    }

    const userSelect = document.getElementById('manageUserSelect');
    const selectedUsername = String(targetUsername || userSelect?.value || '').trim();
    if (!selectedUsername) {
        showUserManageStatus('❌ Pilih user terlebih dahulu.', 'error');
        return;
    }

    const users = getStoredUsers();
    const selectedUser = users.find(user => user.username === selectedUsername);
    if (!selectedUser) {
        showUserManageStatus('❌ Data user tidak ditemukan.', 'error');
        return;
    }

    const popup = document.getElementById('userActionPopup');
    const titleElement = document.getElementById('userActionPopupTitle');
    const messageElement = document.getElementById('userActionPopupMessage');
    const confirmButton = document.getElementById('userActionPopupConfirmBtn');
    if (!popup || !titleElement || !messageElement || !confirmButton) {
        return;
    }

    if (actionType !== 'delete') {
        return;
    }

    pendingUserAction = {
        type: actionType,
        username: selectedUsername
    };

    titleElement.textContent = '🗑️ Konfirmasi Hapus User';
    messageElement.textContent = `Apakah Anda yakin ingin menghapus user ${selectedUsername}?`;
    confirmButton.textContent = '🗑️ Ya, Hapus';
    confirmButton.style.background = '#dc2626';

    popup.style.display = 'flex';
}

function executeUserActionPopup() {
    const { type, username } = pendingUserAction;
    closeUserActionPopup();

    if (!type || !username) {
        return;
    }

    if (type === 'delete') {
        deleteManagedUserDirect(username);
    }
}

function addManagedUser() {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat menambah user.', 'error');
        return;
    }

    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const roleSelect = document.getElementById('manageRole');
    if (!usernameInput || !passwordInput || !fullNameInput || !roleSelect) return;

    const newUsername = String(usernameInput.value || '').trim();
    const newPassword = String(passwordInput.value || '').trim();
    const newFullName = String(fullNameInput.value || '').trim();
    const newRole = String(roleSelect.value || 'guest').trim();

    if (!newUsername || !newPassword) {
        showUserManageStatus('❌ Username dan password wajib diisi untuk menambah user.', 'error');
        return;
    }

    const users = getStoredUsers();
    if (users.some(user => user.username === newUsername)) {
        showUserManageStatus('❌ Username sudah terdaftar.', 'error');
        return;
    }

    const role = ['admin', 'user', 'guest'].includes(newRole) ? newRole : 'guest';
    users.push({
        username: newUsername,
        password: newPassword,
        role,
        fullName: newFullName || newUsername,
        photoUrl: '',
        permissions: getPermissionsByRole(role)
    });

    saveStoredUsers(users);
    renderUserManagement();

    const userSelect = document.getElementById('manageUserSelect');
    if (userSelect) {
        userSelect.value = newUsername;
        onManagedUserChange();
    }

    showUserManageStatus('✅ User baru berhasil ditambahkan.', 'success');
}

function deleteManagedUser(targetUsername) {
    openUserActionPopup('delete', targetUsername);
}

function deleteManagedUserDirect(targetUsername) {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat menghapus user.', 'error');
        return;
    }

    const userSelect = document.getElementById('manageUserSelect');
    const resolvedTargetUsername = String(targetUsername || '').trim();
    if (!userSelect && !resolvedTargetUsername) return;

    const selectedUsername = resolvedTargetUsername || String(userSelect?.value || '').trim();
    if (!selectedUsername) {
        showUserManageStatus('❌ Pilih user yang akan dihapus.', 'error');
        return;
    }

    if (userSelect) {
        const optionExists = Array.from(userSelect.options).some(option => option.value === selectedUsername);
        if (optionExists) {
            userSelect.value = selectedUsername;
            onManagedUserChange();
        }
    }

    if (selectedUsername === currentUser.username) {
        showUserManageStatus('❌ User yang sedang login tidak dapat dihapus.', 'error');
        return;
    }

    const users = getStoredUsers();
    const selectedUser = users.find(user => user.username === selectedUsername);
    if (!selectedUser) {
        showUserManageStatus('❌ User tidak ditemukan.', 'error');
        return;
    }

    const remainingAdmins = users.filter(user => user.role === 'admin' && user.username !== selectedUsername);
    if (selectedUser.role === 'admin' && remainingAdmins.length === 0) {
        showUserManageStatus('❌ Tidak dapat menghapus admin terakhir.', 'error');
        return;
    }

    const filteredUsers = users.filter(user => user.username !== selectedUsername);
    saveStoredUsers(filteredUsers);
    renderUserManagement();
    showUserManageStatus('✅ User berhasil dihapus.', 'success');
}

function switchDashboardMenu(menuKey) {
    const validMenus = ['sorotan', 'grafik', 'laporan', 'manual', 'user', 'password', 'activity-log'];
    if (!validMenus.includes(menuKey)) return;

    if (!canAccessMenu(menuKey)) {
        showStatus('❌ Anda tidak memiliki akses ke menu tersebut.', 'error');
        return;
    }

    const previousMenu = activeDashboardMenu;
    activeDashboardMenu = menuKey;

    if (previousMenu !== menuKey) {
        logActivityEvent('Perpindahan Menu', `Berpindah ke menu ${getMenuLabel(menuKey)}.`, {
            level: 'info',
            category: 'menu',
            source: 'dashboard'
        });
    }

    const buttonMap = {
        sorotan: document.getElementById('menuBtnSorotan'),
        grafik: document.getElementById('menuBtnGrafik'),
        laporan: document.getElementById('menuBtnLaporan'),
        manual: document.getElementById('menuBtnManual'),
        user: document.getElementById('menuBtnUser'),
        password: document.getElementById('menuBtnPassword'),
        'activity-log': document.getElementById('menuBtnActivityLog')
    };

    const sectionMap = {
        sorotan: document.getElementById('menuSectionSorotan'),
        grafik: document.getElementById('menuSectionGrafik'),
        laporan: document.getElementById('menuSectionLaporan'),
        manual: document.getElementById('menuSectionManual'),
        user: document.getElementById('menuSectionUser'),
        password: document.getElementById('menuSectionPassword'),
        'activity-log': document.getElementById('menuSectionActivityLog')
    };

    Object.entries(buttonMap).forEach(([key, element]) => {
        if (!element) return;
        element.classList.toggle('active', key === menuKey);
    });

    Object.entries(sectionMap).forEach(([key, element]) => {
        if (!element) return;
        element.classList.toggle('active', key === menuKey);
    });

    const datasetHero = document.querySelector('.dataset-hero');
    if (datasetHero) {
        datasetHero.style.display = menuKey === 'laporan' ? 'none' : '';
    }

    if (menuKey === 'user') {
        renderUserManagement();
    }

    if (menuKey === 'laporan') {
        loadGoogleSheetReport();
        requestAnimationFrame(() => {
            replayGoogleSheetChartAnimation();
        });
    } else {
        stopGoogleSheetTrendArrowAnimation(chartInstances.googleSheetReport);
    }

    if (menuKey === 'manual') {
        ensureManualInputDefaultDate();
        updateManualTotalPemasukanPreview();
        renderManualInputHistoryTable();
    }

    if (menuKey === 'password') {
        const statusElement = document.getElementById('passwordChangeStatus');
        if (statusElement) {
            statusElement.style.display = 'none';
        }
        resetPasswordChangeForm();
    }

    if (menuKey === 'activity-log') {
        renderActivityLogs();
    }

    updateActiveMenuChip(menuKey);

    const searchInput = document.getElementById('dashboardSearch');
    if (searchInput) {
        applyDashboardFilter(String(searchInput.value || '').trim());
    }

    setTimeout(() => {
        Object.values(chartInstances).forEach(chart => {
            if (chart && typeof chart.resize === 'function') {
                chart.resize();
            }
        });
    }, 80);
}

function updateActiveMenuChip(menuKey) {
    const activeMenuChip = document.getElementById('activeMenuChip');
    if (!activeMenuChip) return;

    const label = getMenuLabel(menuKey);
    activeMenuChip.textContent = `Mode: ${label}`;
}

function updateDashboardClock() {
    const clockElement = document.getElementById('dashboardClock');
    if (!clockElement) return;

    const now = new Date();
    const dateText = now.toLocaleDateString('id-ID', {
        weekday: 'short',
        day: '2-digit',
        month: 'short',
        year: 'numeric'
    });
    const timeText = now.toLocaleTimeString('id-ID', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });

    clockElement.textContent = `${dateText} • ${timeText}`;
}

function getMenuBySearchKeyword(query) {
    const keyword = String(query || '').toLowerCase();
    if (!keyword) return null;

    const keywordMenuMap = [
        { menu: 'sorotan', terms: ['sorotan', 'ringkasan', 'utama', 'rekomendasi', 'kesiapan', 'qris'] },
        { menu: 'grafik', terms: ['grafik', 'chart', 'kanal', 'penerimaan', 'saldo'] },
        { menu: 'laporan', terms: ['laporan', 'transaksi', 'google', 'sheet', 'online'] },
        { menu: 'manual', terms: ['manual', 'input', 'isi data', 'entri'] },
        { menu: 'user', terms: ['user', 'admin', 'akses', 'role', 'akun'] },
        { menu: 'password', terms: ['password', 'kata sandi', 'ganti password', 'ubah password'] },
        { menu: 'activity-log', terms: ['log', 'aktivitas', 'riwayat', 'error', 'peristiwa'] }
    ];

    for (const item of keywordMenuMap) {
        if (!canAccessMenu(item.menu)) continue;
        if (item.terms.some(term => keyword.includes(term))) {
            return item.menu;
        }
    }

    return null;
}

function applyDashboardFilter(query) {
    const normalizedQuery = String(query || '').trim().toLowerCase();
    const activeSection = document.querySelector('.menu-section.active');
    if (!activeSection) return;

    const cards = activeSection.querySelectorAll('.stat-card, .summary-card, .chart-container, .summary-decision-item');
    if (!cards.length) return;

    if (!normalizedQuery) {
        cards.forEach(card => card.classList.remove('search-hidden'));
        return;
    }

    cards.forEach(card => {
        const textContent = String(card.textContent || '').toLowerCase();
        const isVisible = textContent.includes(normalizedQuery);
        card.classList.toggle('search-hidden', !isVisible);
    });
}

function handleDashboardSearch(event) {
    const keyword = String(event?.target?.value || '').trim();
    const mappedMenu = getMenuBySearchKeyword(keyword);

    if (mappedMenu && mappedMenu !== activeDashboardMenu) {
        switchDashboardMenu(mappedMenu);
    } else {
        applyDashboardFilter(keyword);
    }
}

function clearDashboardSearch() {
    const searchInput = document.getElementById('dashboardSearch');
    if (!searchInput) return;
    searchInput.value = '';
    applyDashboardFilter('');
}

function getTodayDateInputValue() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function ensureManualInputDefaultDate() {
    const dateInput = document.getElementById('manualTanggalLaporan');
    if (!dateInput) return;
    if (!String(dateInput.value || '').trim()) {
        dateInput.value = getTodayDateInputValue();
    }
}

function formatManualAmountInput(amount) {
    const numericAmount = Number(amount);
    if (!Number.isFinite(numericAmount) || numericAmount <= 0) return '';
    return numericAmount.toLocaleString('id-ID', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function formatManualTransactionInput(amount) {
    const numericAmount = Number(amount);
    if (!Number.isFinite(numericAmount) || numericAmount <= 0) return '';
    return String(Math.round(numericAmount));
}

function getManualReportFieldByKey(reportKey) {
    return MANUAL_REPORT_FIELDS.find((field) => field.key === reportKey) || null;
}

function closeManualReportInputModal(event) {
    if (event && event.target && event.currentTarget && event.target !== event.currentTarget) {
        return;
    }

    const modal = document.getElementById('manualReportInputModal');
    const fieldKeyInput = document.getElementById('manualReportInputFieldKey');
    if (modal) {
        modal.style.display = 'none';
    }
    if (fieldKeyInput) {
        fieldKeyInput.value = '';
    }
}

function updateManualReportInputModalTotal() {
    const channelInputMap = {
        qris: 'manualReportInputQris',
        mobileBanking: 'manualReportInputMobileBanking',
        edc: 'manualReportInputEDC',
        atm: 'manualReportInputATM',
        teller: 'manualReportInputTeller'
    };

    const total = MANUAL_TRANSACTION_CHANNELS.reduce((sum, channel) => {
        const rawValue = String(document.getElementById(channelInputMap[channel.key])?.value || '').trim();
        const parsedValue = Number.parseInt(rawValue, 10);
        return sum + (Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0);
    }, 0);

    const totalInput = document.getElementById('manualReportInputTotal');
    if (totalInput) {
        totalInput.value = formatManualTransactionInput(total);
    }
}

function openManualReportInputModal(reportKey) {
    const field = getManualReportFieldByKey(reportKey);
    if (!field) {
        showStatus('❌ Form laporan tidak ditemukan.', 'error');
        return;
    }

    const modal = document.getElementById('manualReportInputModal');
    const fieldKeyInput = document.getElementById('manualReportInputFieldKey');
    const titleElement = document.getElementById('manualReportInputTitle');
    const amountLabelElement = document.getElementById('manualReportInputAmountLabel');
    const amountInput = document.getElementById('manualReportInputAmount');
    const qrisInput = document.getElementById('manualReportInputQris');
    const mobileInput = document.getElementById('manualReportInputMobileBanking');
    const edcInput = document.getElementById('manualReportInputEDC');
    const atmInput = document.getElementById('manualReportInputATM');
    const tellerInput = document.getElementById('manualReportInputTeller');

    if (!modal || !fieldKeyInput || !amountInput || !qrisInput || !mobileInput || !edcInput || !atmInput || !tellerInput) {
        showStatus('❌ Komponen popup input belum siap.', 'error');
        return;
    }

    const channelIds = field.transaksiChannelInputIds || {};
    const sourceAmount = document.getElementById(field.inputId);

    fieldKeyInput.value = field.key;
    if (titleElement) {
        titleElement.textContent = `🧾 Form Input ${field.label.replace('Laporan ', '')}`;
    }
    if (amountLabelElement) {
        amountLabelElement.textContent = field.label;
    }

    amountInput.value = String(sourceAmount?.value || '');
    qrisInput.value = String(document.getElementById(channelIds.qris)?.value || '');
    mobileInput.value = String(document.getElementById(channelIds.mobileBanking)?.value || '');
    edcInput.value = String(document.getElementById(channelIds.edc)?.value || '');
    atmInput.value = String(document.getElementById(channelIds.atm)?.value || '');
    tellerInput.value = String(document.getElementById(channelIds.teller)?.value || '');

    updateManualReportInputModalTotal();
    modal.style.display = 'flex';
}

function saveManualReportInputModal() {
    const fieldKey = String(document.getElementById('manualReportInputFieldKey')?.value || '').trim();
    const field = getManualReportFieldByKey(fieldKey);
    if (!field) {
        showStatus('❌ Form laporan tidak valid.', 'error');
        return;
    }

    const amountInput = document.getElementById('manualReportInputAmount');
    const qrisInput = document.getElementById('manualReportInputQris');
    const mobileInput = document.getElementById('manualReportInputMobileBanking');
    const edcInput = document.getElementById('manualReportInputEDC');
    const atmInput = document.getElementById('manualReportInputATM');
    const tellerInput = document.getElementById('manualReportInputTeller');
    if (!amountInput || !qrisInput || !mobileInput || !edcInput || !atmInput || !tellerInput) {
        showStatus('❌ Input popup tidak lengkap.', 'error');
        return;
    }

    const amountValueRaw = String(amountInput.value || '').trim();
    const parsedAmount = parseAmount(amountValueRaw);
    const normalizedAmount = Number.isFinite(parsedAmount) && parsedAmount > 0
        ? formatManualAmountInput(parsedAmount)
        : '';

    const normalizeTransaction = (value) => {
        const parsedValue = Number.parseInt(String(value || '').trim(), 10);
        return Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0;
    };

    const qris = normalizeTransaction(qrisInput.value);
    const mobileBanking = normalizeTransaction(mobileInput.value);
    const edc = normalizeTransaction(edcInput.value);
    const atm = normalizeTransaction(atmInput.value);
    const teller = normalizeTransaction(tellerInput.value);
    const total = qris + mobileBanking + edc + atm + teller;

    const sourceAmountInput = document.getElementById(field.inputId);
    const sourceTotalInput = document.getElementById(field.transaksiInputId);
    const channelIds = field.transaksiChannelInputIds || {};
    const sourceQrisInput = document.getElementById(channelIds.qris);
    const sourceMobileInput = document.getElementById(channelIds.mobileBanking);
    const sourceEdcInput = document.getElementById(channelIds.edc);
    const sourceAtmInput = document.getElementById(channelIds.atm);
    const sourceTellerInput = document.getElementById(channelIds.teller);

    if (!sourceAmountInput || !sourceTotalInput || !sourceQrisInput || !sourceMobileInput || !sourceEdcInput || !sourceAtmInput || !sourceTellerInput) {
        showStatus('❌ Target input laporan tidak ditemukan.', 'error');
        return;
    }

    sourceAmountInput.value = normalizedAmount;
    sourceQrisInput.value = formatManualTransactionInput(qris);
    sourceMobileInput.value = formatManualTransactionInput(mobileBanking);
    sourceEdcInput.value = formatManualTransactionInput(edc);
    sourceAtmInput.value = formatManualTransactionInput(atm);
    sourceTellerInput.value = formatManualTransactionInput(teller);
    sourceTotalInput.value = formatManualTransactionInput(total);

    updateManualTotalPemasukanPreview();
    closeManualReportInputModal();
}

function toManualTransactionChannelKey(transaksiKey, channelKey) {
    const safeChannel = String(channelKey || '').trim();
    if (!safeChannel) return transaksiKey;
    return `${transaksiKey}${safeChannel.charAt(0).toUpperCase()}${safeChannel.slice(1)}`;
}

function calculateManualReportTotal(reportAmounts) {
    return MANUAL_REPORT_FIELDS.reduce((sum, field) => {
        const value = Number(reportAmounts?.[field.key] || 0);
        return sum + (Number.isFinite(value) && value > 0 ? value : 0);
    }, 0);
}

function calculateManualReportTransactionTotal(reportTransactions) {
    return MANUAL_REPORT_FIELDS.reduce((sum, field) => {
        const value = Number(reportTransactions?.[field.transaksiKey] || 0);
        return sum + (Number.isFinite(value) && value > 0 ? value : 0);
    }, 0);
}

function normalizeManualReportAmounts(record) {
    const normalized = {};
    MANUAL_REPORT_FIELDS.forEach((field) => {
        const value = Number(record?.[field.key]);
        normalized[field.key] = Number.isFinite(value) && value > 0 ? value : 0;
    });

    const calculatedTotal = calculateManualReportTotal(normalized);
    const fallbackTotal = Number(record?.totalPemasukan || 0);
    if (calculatedTotal <= 0 && Number.isFinite(fallbackTotal) && fallbackTotal > 0) {
        normalized.laporanRKUD = fallbackTotal;
    }

    return normalized;
}

function normalizeManualReportTransactions(record) {
    const normalized = {};
    const normalizedChannels = normalizeManualReportTransactionChannels(record);

    MANUAL_REPORT_FIELDS.forEach((field) => {
        const channelTotal = MANUAL_TRANSACTION_CHANNELS.reduce((sum, channel) => {
            const channelKey = toManualTransactionChannelKey(field.transaksiKey, channel.key);
            const value = Number(normalizedChannels[channelKey] || 0);
            return sum + (Number.isFinite(value) && value > 0 ? value : 0);
        }, 0);
        const fallbackValue = Number(record?.[field.transaksiKey]);
        const roundedFallbackValue = Number.isFinite(fallbackValue) && fallbackValue > 0 ? Math.round(fallbackValue) : 0;
        const finalValue = channelTotal > 0 ? channelTotal : roundedFallbackValue;

        if (channelTotal <= 0 && finalValue > 0) {
            const tellerKey = toManualTransactionChannelKey(field.transaksiKey, 'teller');
            normalizedChannels[tellerKey] = finalValue;
        }

        normalized[field.transaksiKey] = finalValue;
    });

    const calculatedTotal = calculateManualReportTransactionTotal(normalized);
    const fallbackTotal = Number(record?.totalTransaksi || 0);
    if (calculatedTotal <= 0 && Number.isFinite(fallbackTotal) && fallbackTotal > 0) {
        normalized.transaksiRKUD = Math.round(fallbackTotal);
    }

    return normalized;
}

function normalizeManualReportTransactionChannels(record) {
    const normalized = {};

    MANUAL_REPORT_FIELDS.forEach((field) => {
        MANUAL_TRANSACTION_CHANNELS.forEach((channel) => {
            const key = toManualTransactionChannelKey(field.transaksiKey, channel.key);
            const value = Number(record?.[key]);
            normalized[key] = Number.isFinite(value) && value > 0 ? Math.round(value) : 0;
        });
    });

    return normalized;
}

function getManualReportAmountsFromForm(isEdit = false) {
    const amounts = {};
    MANUAL_REPORT_FIELDS.forEach((field) => {
        const targetId = isEdit ? field.editInputId : field.inputId;
        const rawValue = String(document.getElementById(targetId)?.value || '').trim();
        const parsedValue = parseAmount(rawValue);
        amounts[field.key] = Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0;
    });
    return amounts;
}

function getManualReportTransactionsFromForm(isEdit = false) {
    const transactions = {};
    MANUAL_REPORT_FIELDS.forEach((field) => {
        const channelIds = isEdit ? field.editTransaksiChannelInputIds : field.transaksiChannelInputIds;
        const totalByChannels = MANUAL_TRANSACTION_CHANNELS.reduce((sum, channel) => {
            const rawValue = String(document.getElementById(channelIds?.[channel.key])?.value || '').trim();
            const parsedValue = Number.parseInt(rawValue, 10);
            return sum + (Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0);
        }, 0);

        const totalInput = document.getElementById(isEdit ? field.editTransaksiInputId : field.transaksiInputId);
        if (totalInput) {
            totalInput.value = formatManualTransactionInput(totalByChannels);
        }

        transactions[field.transaksiKey] = totalByChannels;
    });
    return transactions;
}

function getManualReportTransactionChannelsFromForm(isEdit = false) {
    const channels = {};

    MANUAL_REPORT_FIELDS.forEach((field) => {
        const channelIds = isEdit ? field.editTransaksiChannelInputIds : field.transaksiChannelInputIds;
        MANUAL_TRANSACTION_CHANNELS.forEach((channel) => {
            const key = toManualTransactionChannelKey(field.transaksiKey, channel.key);
            const rawValue = String(document.getElementById(channelIds?.[channel.key])?.value || '').trim();
            const parsedValue = Number.parseInt(rawValue, 10);
            channels[key] = Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0;
        });
    });

    return channels;
}

function setManualReportAmountsToForm(reportAmounts, isEdit = false) {
    MANUAL_REPORT_FIELDS.forEach((field) => {
        const targetId = isEdit ? field.editInputId : field.inputId;
        const inputElement = document.getElementById(targetId);
        if (!inputElement) return;
        inputElement.value = formatManualAmountInput(reportAmounts?.[field.key] || 0);
    });
}

function setManualReportTransactionsToForm(reportTransactions, isEdit = false, reportTransactionChannels = {}) {
    MANUAL_REPORT_FIELDS.forEach((field) => {
        const channelIds = isEdit ? field.editTransaksiChannelInputIds : field.transaksiChannelInputIds;
        let channelTotal = 0;

        MANUAL_TRANSACTION_CHANNELS.forEach((channel) => {
            const key = toManualTransactionChannelKey(field.transaksiKey, channel.key);
            const inputElement = document.getElementById(channelIds?.[channel.key]);
            if (!inputElement) return;

            const channelValue = Number(reportTransactionChannels?.[key] || 0);
            const safeChannelValue = Number.isFinite(channelValue) && channelValue > 0 ? Math.round(channelValue) : 0;
            inputElement.value = formatManualTransactionInput(safeChannelValue);
            channelTotal += safeChannelValue;
        });

        const targetId = isEdit ? field.editTransaksiInputId : field.transaksiInputId;
        const totalInputElement = document.getElementById(targetId);
        if (!totalInputElement) return;

        const fallbackTotal = Number(reportTransactions?.[field.transaksiKey] || 0);
        const finalTotal = channelTotal > 0 ? channelTotal : (Number.isFinite(fallbackTotal) && fallbackTotal > 0 ? Math.round(fallbackTotal) : 0);
        totalInputElement.value = formatManualTransactionInput(finalTotal);
    });
}

function setManualAutoTotalDisplay(totalAmount, isEdit = false) {
    const targetId = isEdit ? 'manualEditTotalPemasukan' : 'manualTotalPemasukan';
    const inputElement = document.getElementById(targetId);
    if (!inputElement) return;

    const safeTotal = Number(totalAmount);
    if (!Number.isFinite(safeTotal) || safeTotal <= 0) {
        inputElement.value = '';
        return;
    }

    inputElement.value = formatCurrency(safeTotal);
}

function setManualAutoTransactionDisplay(totalTransaction, isEdit = false) {
    const targetId = isEdit ? 'manualEditTotalTransaksi' : 'manualTotalTransaksi';
    const inputElement = document.getElementById(targetId);
    if (!inputElement) return;

    const safeTotal = Number(totalTransaction);
    if (!Number.isFinite(safeTotal) || safeTotal <= 0) {
        inputElement.value = '';
        return;
    }

    inputElement.value = String(Math.round(safeTotal));
}

function renderManualChannelMonthlySummary(history = null) {
    const tableBody = document.getElementById('manualChannelMonthlyTableBody');
    const emptyState = document.getElementById('manualChannelMonthlyEmpty');
    if (!tableBody || !emptyState) return;

    const safeHistory = Array.isArray(history) ? history : getManualInputHistory();
    const monthlyMap = new Map();

    safeHistory.forEach((record) => {
        const reportDate = new Date(`${String(record?.tanggalLaporan || '').trim()}T00:00:00`);
        if (!Number.isFinite(reportDate.getTime())) return;

        const monthKey = `${reportDate.getFullYear()}-${String(reportDate.getMonth() + 1).padStart(2, '0')}`;
        if (!monthlyMap.has(monthKey)) {
            monthlyMap.set(monthKey, {
                monthKey,
                monthLabel: reportDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }),
                qris: 0,
                mobileBanking: 0,
                edc: 0,
                atm: 0,
                teller: 0
            });
        }

        const bucket = monthlyMap.get(monthKey);
        const reportChannels = normalizeManualReportTransactionChannels(record);

        MANUAL_REPORT_FIELDS.forEach((field) => {
            MANUAL_TRANSACTION_CHANNELS.forEach((channel) => {
                const channelKey = toManualTransactionChannelKey(field.transaksiKey, channel.key);
                const channelValue = Number(reportChannels?.[channelKey] || 0);
                if (!Number.isFinite(channelValue) || channelValue <= 0) return;
                bucket[channel.key] += Math.round(channelValue);
            });
        });
    });

    const monthlyRows = Array.from(monthlyMap.values()).sort((a, b) => String(a.monthKey).localeCompare(String(b.monthKey)));

    if (!monthlyRows.length) {
        tableBody.innerHTML = '';
        emptyState.style.display = 'block';
        return;
    }

    const formatCount = (value) => Number(value || 0).toLocaleString('id-ID');
    const rowsHtml = monthlyRows.map((row) => {
        const totalKanal = Number(row.qris || 0)
            + Number(row.mobileBanking || 0)
            + Number(row.edc || 0)
            + Number(row.atm || 0)
            + Number(row.teller || 0);

        return `
            <tr>
                <td>${row.monthLabel}</td>
                <td>${formatCount(row.qris)}</td>
                <td>${formatCount(row.mobileBanking)}</td>
                <td>${formatCount(row.edc)}</td>
                <td>${formatCount(row.atm)}</td>
                <td>${formatCount(row.teller)}</td>
                <td><strong>${formatCount(totalKanal)}</strong></td>
            </tr>
        `;
    }).join('');

    tableBody.innerHTML = rowsHtml;
    emptyState.style.display = 'none';
}

function getManualChannelTotalsForYear(year, history = null) {
    const targetYear = Number.parseInt(String(year || ''), 10);
    const safeHistory = Array.isArray(history) ? history : getManualInputHistory();
    const totals = {
        qris: 0,
        mobileBanking: 0,
        edc: 0,
        atm: 0,
        teller: 0,
        hasData: false
    };

    if (!Number.isInteger(targetYear) || targetYear <= 0) {
        return totals;
    }

    safeHistory.forEach((record) => {
        const reportDate = new Date(`${String(record?.tanggalLaporan || '').trim()}T00:00:00`);
        if (!Number.isFinite(reportDate.getTime()) || reportDate.getFullYear() !== targetYear) return;

        const reportChannels = normalizeManualReportTransactionChannels(record);
        MANUAL_REPORT_FIELDS.forEach((field) => {
            MANUAL_TRANSACTION_CHANNELS.forEach((channel) => {
                const channelKey = toManualTransactionChannelKey(field.transaksiKey, channel.key);
                const channelValue = Number(reportChannels?.[channelKey] || 0);
                if (!Number.isFinite(channelValue) || channelValue <= 0) return;
                totals[channel.key] += Math.round(channelValue);
                totals.hasData = true;
            });
        });
    });

    return totals;
}

function getManualReportAmountTotalsForYear(year, history = null) {
    const targetYear = Number.parseInt(String(year || ''), 10);
    const safeHistory = Array.isArray(history) ? history : getManualInputHistory();
    const totals = {
        byReport: {},
        total: 0,
        hasData: false
    };

    MANUAL_REPORT_FIELDS.forEach((field) => {
        totals.byReport[field.key] = {
            label: field.label,
            amount: 0
        };
    });

    if (!Number.isInteger(targetYear) || targetYear <= 0) {
        return totals;
    }

    safeHistory.forEach((record) => {
        const reportDate = new Date(`${String(record?.tanggalLaporan || '').trim()}T00:00:00`);
        if (!Number.isFinite(reportDate.getTime()) || reportDate.getFullYear() !== targetYear) return;

        const reportAmounts = normalizeManualReportAmounts(record);
        MANUAL_REPORT_FIELDS.forEach((field) => {
            const amount = Number(reportAmounts?.[field.key] || 0);
            if (!Number.isFinite(amount) || amount <= 0) return;
            totals.byReport[field.key].amount += amount;
            totals.total += amount;
            totals.hasData = true;
        });
    });

    return totals;
}

function renderGoogleSheetCreditBreakdown(creditTotals, activeYear) {
    const container = document.getElementById('gsTotalCreditBreakdown');
    if (!container) return;

    if (!creditTotals?.hasData) {
        container.innerHTML = `<div style="font-size:0.8em; color:#64748b; margin-top:8px;">Detail pemasukan manual tahun ${activeYear} belum tersedia.</div>`;
        return;
    }

    const rowsHtml = MANUAL_REPORT_FIELDS.map((field) => {
        const item = creditTotals.byReport[field.key] || { label: field.label, amount: 0 };
        return `<div style="display:flex; justify-content:space-between; gap:8px;"><span>${item.label}</span><strong>${formatCurrency(Number(item.amount || 0))}</strong></div>`;
    }).join('');

    container.innerHTML = `
        <div style="margin-top:10px; font-size:0.8em; color:#475569; line-height:1.45; display:grid; gap:4px;">
            ${rowsHtml}
        </div>
    `;
}

function validateManualReportPairing(reportAmounts, reportTransactions) {
    for (const field of MANUAL_REPORT_FIELDS) {
        const amount = Number(reportAmounts?.[field.key] || 0);
        const transaction = Number(reportTransactions?.[field.transaksiKey] || 0);
        if (amount > 0 && transaction <= 0) {
            return `❌ Total transaksi untuk ${field.label} harus lebih dari 0.`;
        }
        if (transaction > 0 && amount <= 0) {
            return `❌ Nominal untuk ${field.label} harus lebih dari 0 jika transaksinya diisi.`;
        }
    }
    return '';
}

function updateManualTotalPemasukanPreview() {
    const reportAmounts = getManualReportAmountsFromForm(false);
    const reportTransactions = getManualReportTransactionsFromForm(false);
    const totalAmount = calculateManualReportTotal(reportAmounts);
    const totalTransaction = calculateManualReportTransactionTotal(reportTransactions);
    setManualAutoTotalDisplay(totalAmount, false);
    setManualAutoTransactionDisplay(totalTransaction, false);
}

function updateManualEditTotalPemasukanPreview() {
    const reportAmounts = getManualReportAmountsFromForm(true);
    const reportTransactions = getManualReportTransactionsFromForm(true);
    const totalAmount = calculateManualReportTotal(reportAmounts);
    const totalTransaction = calculateManualReportTransactionTotal(reportTransactions);
    setManualAutoTotalDisplay(totalAmount, true);
    setManualAutoTransactionDisplay(totalTransaction, true);
}

function closeManualEditModal(event) {
    if (event && event.target && event.currentTarget && event.target !== event.currentTarget) {
        return;
    }

    const modal = document.getElementById('manualEditModal');
    const rowIndexInput = document.getElementById('manualEditRowIndex');
    if (modal) {
        modal.style.display = 'none';
    }
    if (rowIndexInput) {
        rowIndexInput.value = '';
    }
}

function openManualEditModal(index) {
    const history = getManualInputHistory();
    const targetIndex = Number.parseInt(index, 10);
    if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= history.length) {
        showStatus('❌ Data riwayat tidak ditemukan untuk diedit.', 'error');
        return;
    }

    const target = history[targetIndex];
    const modal = document.getElementById('manualEditModal');
    const indexInput = document.getElementById('manualEditRowIndex');
    const tanggalInput = document.getElementById('manualEditTanggalLaporan');
    const reportAmounts = normalizeManualReportAmounts(target);
    const reportTransactionChannels = normalizeManualReportTransactionChannels(target);
    const reportTransactions = normalizeManualReportTransactions(target);
    const totalPemasukan = calculateManualReportTotal(reportAmounts);
    const totalTransaksi = calculateManualReportTransactionTotal(reportTransactions);

    if (indexInput) indexInput.value = String(targetIndex);
    if (tanggalInput) tanggalInput.value = String(target.tanggalLaporan || '');
    setManualReportAmountsToForm(reportAmounts, true);
    setManualReportTransactionsToForm(reportTransactions, true, reportTransactionChannels);
    setManualAutoTotalDisplay(totalPemasukan, true);
    setManualAutoTransactionDisplay(totalTransaksi, true);
    if (modal) modal.style.display = 'flex';
}

function saveManualHistoryEdit() {
    const indexInput = document.getElementById('manualEditRowIndex');
    const tanggalRaw = String(document.getElementById('manualEditTanggalLaporan')?.value || '').trim();
    const rowIndex = Number.parseInt(String(indexInput?.value || '-1'), 10);
    const reportAmounts = getManualReportAmountsFromForm(true);
    const reportTransactionChannels = getManualReportTransactionChannelsFromForm(true);
    const reportTransactions = getManualReportTransactionsFromForm(true);
    const transaksi = calculateManualReportTransactionTotal(reportTransactions);
    const pemasukan = calculateManualReportTotal(reportAmounts);

    if (!Number.isInteger(rowIndex) || rowIndex < 0) {
        showStatus('❌ Baris edit tidak valid.', 'error');
        return;
    }

    if (!tanggalRaw) {
        showStatus('❌ Tanggal Laporan wajib diisi.', 'error');
        return;
    }

    const selectedDate = new Date(`${tanggalRaw}T00:00:00`);
    if (!Number.isFinite(selectedDate.getTime())) {
        showStatus('❌ Format Tanggal Laporan tidak valid.', 'error');
        return;
    }

    if (!Number.isFinite(transaksi) || transaksi <= 0) {
        showStatus('❌ Total Transaksi otomatis harus lebih dari 0 (isi minimal satu laporan transaksi).', 'error');
        return;
    }

    if (!Number.isFinite(pemasukan) || pemasukan <= 0) {
        showStatus('❌ Total Pemasukan otomatis harus lebih dari 0 (isi minimal satu laporan).', 'error');
        return;
    }

    const pairingError = validateManualReportPairing(reportAmounts, reportTransactions);
    if (pairingError) {
        showStatus(pairingError, 'error');
        return;
    }

    const history = getManualInputHistory();
    if (rowIndex >= history.length) {
        showStatus('❌ Data riwayat tidak ditemukan.', 'error');
        return;
    }

    history[rowIndex] = {
        ...history[rowIndex],
        tanggalLaporan: tanggalRaw,
        totalTransaksi: transaksi,
        ...reportAmounts,
        ...reportTransactionChannels,
        ...reportTransactions,
        totalPemasukan: pemasukan
    };

    saveManualInputHistory(history);
    renderManualInputHistoryTable();

    const dataset = buildDatasetFromManualHistory(history);
    const rows = normalizeRowsForProcessing(dataset.rows);
    if (rows.length > 0) {
        saveManualInputDataset(rows, {
            updatedFrom: 'manual-edit',
            totalRecords: history.length
        });
        processData(rows);
    }

    closeManualEditModal();
    showStatus('✅ Data manual berhasil diperbarui.', 'success');
}

function deleteManualHistoryRow(index) {
    const rowIndex = Number.parseInt(index, 10);
    if (!Number.isInteger(rowIndex) || rowIndex < 0) {
        showStatus('❌ Baris data yang akan dihapus tidak valid.', 'error');
        return;
    }

    const history = getManualInputHistory();
    if (rowIndex >= history.length) {
        showStatus('❌ Data riwayat tidak ditemukan.', 'error');
        return;
    }

    const target = history[rowIndex];
    const reportDate = new Date(`${target.tanggalLaporan}T00:00:00`);
    const reportDateLabel = Number.isFinite(reportDate.getTime())
        ? reportDate.toLocaleDateString('id-ID')
        : String(target.tanggalLaporan || '-');

    if (!confirm(`Hapus data manual tanggal ${reportDateLabel} dengan total ${formatCurrency(Number(target.totalPemasukan || 0))}?`)) {
        return;
    }

    const updatedHistory = history.filter((_, currentIndex) => currentIndex !== rowIndex);
    saveManualInputHistory(updatedHistory);
    renderManualInputHistoryTable();

    if (updatedHistory.length === 0) {
        try {
            localStorage.removeItem(MANUAL_INPUT_DATA_KEY);
        } catch (error) {
            console.error('Gagal menghapus data input manual tersimpan:', error);
        }

        resetGoogleSheetSummary();
        updateGoogleSheetMeta('Sumber laporan belum tersedia. Gunakan Input Data Manual atau muat CSV.', 'error');
        showStatus('✅ Data manual berhasil dihapus. Riwayat sudah kosong.', 'success');
        return;
    }

    const dataset = buildDatasetFromManualHistory(updatedHistory);
    const rows = normalizeRowsForProcessing(dataset.rows);
    if (rows.length > 0) {
        saveManualInputDataset(rows, {
            totalRecords: updatedHistory.length,
            source: 'manual-history'
        });
        processData(rows);
        showStatus('✅ Data manual berhasil dihapus dan laporan diperbarui.', 'success');
        return;
    }

    showStatus('⚠️ Data berhasil dihapus, namun dataset laporan tidak dapat diproses.', 'warning');
}

function getManualInputHistory() {
    try {
        const raw = localStorage.getItem(MANUAL_INPUT_HISTORY_KEY);
        if (!raw) return [];

        const parsed = JSON.parse(raw);
        if (!Array.isArray(parsed)) return [];

        return parsed
            .filter((item) => item && typeof item === 'object')
            .map((item) => {
                const reportDate = String(item.tanggalLaporan || '').trim();
                const reportAmounts = normalizeManualReportAmounts(item);
                const reportTransactionChannels = normalizeManualReportTransactionChannels(item);
                const reportTransactions = normalizeManualReportTransactions(item);
                const calculatedTotal = calculateManualReportTotal(reportAmounts);
                const calculatedTransactions = calculateManualReportTransactionTotal(reportTransactions);
                const fallbackTotal = Number(item.totalPemasukan);
                const totalPemasukan = calculatedTotal > 0 ? calculatedTotal : fallbackTotal;
                const fallbackTransactions = Number(item.totalTransaksi);
                const totalTransaksi = calculatedTransactions > 0 ? calculatedTransactions : fallbackTransactions;
                const isValid = reportDate && Number.isFinite(totalPemasukan) && totalPemasukan > 0 && Number.isFinite(totalTransaksi) && totalTransaksi > 0;
                if (!isValid) return null;
                return {
                    ...item,
                    ...reportAmounts,
                    ...reportTransactionChannels,
                    ...reportTransactions,
                    totalTransaksi,
                    totalPemasukan
                };
            })
            .filter(Boolean);
    } catch (error) {
        return [];
    }
}

function saveManualInputHistory(history) {
    try {
        localStorage.setItem(MANUAL_INPUT_HISTORY_KEY, JSON.stringify(Array.isArray(history) ? history : []));
    } catch (error) {
        console.error('Gagal menyimpan riwayat input manual:', error);
    }
}

function buildSyntheticRowsFromManualRecord(record, recordIndex = 0) {
    const reportAmounts = normalizeManualReportAmounts(record);
    const reportTransactionChannels = normalizeManualReportTransactionChannels(record);
    const reportTransactions = normalizeManualReportTransactions(record);
    const calculatedTotal = calculateManualReportTotal(reportAmounts);
    const calculatedTotalTransaksi = calculateManualReportTransactionTotal(reportTransactions);
    const fallbackTotal = Number(record.totalPemasukan);
    const fallbackTransaksi = Math.max(1, Number.parseInt(record.totalTransaksi, 10));
    const totalPemasukan = calculatedTotal > 0 ? calculatedTotal : fallbackTotal;
    const totalTransaksi = calculatedTotalTransaksi > 0 ? calculatedTotalTransaksi : fallbackTransaksi;
    const selectedDate = new Date(`${record.tanggalLaporan}T00:00:00`);
    if (!Number.isFinite(selectedDate.getTime()) || !Number.isFinite(totalPemasukan) || totalPemasukan <= 0) return [];

    const selectedYear = selectedDate.getFullYear();
    const selectedMonth = selectedDate.getMonth();
    const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate();
    const monthLabel = selectedDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
    const hasDetailedSplit = MANUAL_REPORT_FIELDS.some((field) => {
        const amount = Number(reportAmounts[field.key] || 0);
        const transaction = Number(reportTransactions[field.transaksiKey] || 0);
        return amount > 0 && transaction > 0;
    });

    const buildSyntheticRow = (index, amountPerTransaction, jenisPenerimaan, channelName = 'Teller') => {
        const syntheticDate = new Date(
            selectedYear,
            selectedMonth,
            (index % daysInMonth) + 1,
            8 + ((index + recordIndex) % 10),
            ((index * 11) + (recordIndex * 3)) % 60,
            0
        );

        const day = String(syntheticDate.getDate()).padStart(2, '0');
        const month = String(syntheticDate.getMonth() + 1).padStart(2, '0');
        const year = syntheticDate.getFullYear();
        const hour = String(syntheticDate.getHours()).padStart(2, '0');
        const minute = String(syntheticDate.getMinutes()).padStart(2, '0');
        const second = String(syntheticDate.getSeconds()).padStart(2, '0');
        const tanggal = `${day}/${month}/${year}`;
        const jam = `${hour}:${minute}:${second}`;

        return {
            AccountNo: 'MANUAL-INPUT',
            PostDate: `${tanggal} ${jam}`,
            Tanggal: tanggal,
            Jam: jam,
            Bulan: monthLabel,
            Source: 'Input Manual',
            __sourceFile: 'Input Manual',
            'Jenis Penerimaan': jenisPenerimaan,
            'Jenis Kanal': channelName,
            'Credit Amount': Number(amountPerTransaction.toFixed(2))
        };
    };

    if (hasDetailedSplit) {
        let runningIndex = 0;
        const rows = [];

        MANUAL_REPORT_FIELDS.forEach((field) => {
            const reportAmount = Number(reportAmounts[field.key] || 0);
            const reportTransaction = Number(reportTransactions[field.transaksiKey] || 0);
            if (!(reportAmount > 0) || !(reportTransaction > 0)) return;

            const channelBreakdown = MANUAL_TRANSACTION_CHANNELS.map((channel) => {
                const key = toManualTransactionChannelKey(field.transaksiKey, channel.key);
                const value = Number(reportTransactionChannels[key] || 0);
                return {
                    key: channel.key,
                    label: channel.label,
                    count: Number.isFinite(value) && value > 0 ? Math.round(value) : 0
                };
            });

            const channelTransactionTotal = channelBreakdown.reduce((sum, channel) => sum + channel.count, 0);
            if (channelTransactionTotal > 0) {
                const amountPerTransaction = reportAmount / channelTransactionTotal;
                channelBreakdown.forEach((channel) => {
                    for (let i = 0; i < channel.count; i += 1) {
                        rows.push(buildSyntheticRow(runningIndex, amountPerTransaction, field.label, channel.label));
                        runningIndex += 1;
                    }
                });
                return;
            }

            const amountPerTransaction = reportAmount / reportTransaction;
            for (let i = 0; i < reportTransaction; i += 1) {
                rows.push(buildSyntheticRow(runningIndex, amountPerTransaction, field.label, 'Teller'));
                runningIndex += 1;
            }
        });

        if (rows.length > 0) {
            return rows;
        }
    }

    const amountPerTransaction = totalPemasukan / totalTransaksi;
    return Array.from({ length: totalTransaksi }, (_, index) => buildSyntheticRow(index, amountPerTransaction, 'Input Manual'));
}

function buildDatasetFromManualHistory(history) {
    const safeHistory = Array.isArray(history) ? history : [];
    const rows = safeHistory.flatMap((record, index) => buildSyntheticRowsFromManualRecord(record, index));
    return {
        createdAt: safeHistory.length > 0 ? safeHistory[safeHistory.length - 1].savedAt : new Date().toISOString(),
        rows,
        metadata: {
            totalRecords: safeHistory.length
        }
    };
}

function renderManualInputHistoryTable() {
    const tableBody = document.getElementById('manualHistoryTableBody');
    const emptyState = document.getElementById('manualHistoryEmpty');
    if (!tableBody || !emptyState) return;

    const sourceHistory = getManualInputHistory();
    renderManualChannelMonthlySummary(sourceHistory);

    const history = sourceHistory
        .map((item, originalIndex) => ({ item, originalIndex }))
        .sort((a, b) => String(a.item.tanggalLaporan || '').localeCompare(String(b.item.tanggalLaporan || '')));

    if (history.length === 0) {
        tableBody.innerHTML = '';
        emptyState.style.display = 'block';
        return;
    }

    const escapeHtml = (value) => String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const rowsHtml = history.map((entry, index) => {
        const { item, originalIndex } = entry;
        const reportDate = new Date(`${item.tanggalLaporan}T00:00:00`);
        const reportDateLabel = Number.isFinite(reportDate.getTime())
            ? reportDate.toLocaleDateString('id-ID')
            : '-';
        const periodLabel = Number.isFinite(reportDate.getTime())
            ? reportDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' })
            : '-';
        const savedAtDate = new Date(String(item.savedAt || ''));
        const savedAtLabel = Number.isFinite(savedAtDate.getTime())
            ? savedAtDate.toLocaleString('id-ID')
            : '-';
        const actionButton = `
            <div style="display:flex; gap:6px; justify-content:center; flex-wrap:wrap;">
                <button type="button" onclick="openManualEditModal(${originalIndex})" style="padding:6px 10px; border-radius:8px; background:#0ea5e9; color:#fff; border:none; font-weight:600; cursor:pointer;">✏️ Edit</button>
                <button type="button" onclick="deleteManualHistoryRow(${originalIndex})" style="padding:6px 10px; border-radius:8px; background:#ef4444; color:#fff; border:none; font-weight:600; cursor:pointer;">🗑️ Hapus</button>
            </div>
        `;

        return `
            <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(reportDateLabel)}</td>
                <td>${escapeHtml(periodLabel)}</td>
                <td>${Number(item.totalTransaksi || 0).toLocaleString('id-ID')}</td>
                <td>${escapeHtml(formatCurrency(Number(item.totalPemasukan || 0)))}</td>
                <td>${escapeHtml(savedAtLabel)}</td>
                <td>${actionButton}</td>
            </tr>
        `;
    }).join('');

    tableBody.innerHTML = rowsHtml;
    emptyState.style.display = 'none';
}

function saveManualInputDataset(rows, metadata = {}) {
    try {
        const payload = {
            createdAt: new Date().toISOString(),
            rows: Array.isArray(rows) ? rows : [],
            metadata: metadata || {}
        };
        localStorage.setItem(MANUAL_INPUT_DATA_KEY, JSON.stringify(payload));
    } catch (error) {
        console.error('Gagal menyimpan data input manual:', error);
    }
}

function getManualInputDataset() {
    const history = getManualInputHistory();
    if (history.length > 0) {
        return buildDatasetFromManualHistory(history);
    }

    try {
        const raw = localStorage.getItem(MANUAL_INPUT_DATA_KEY);
        if (!raw) return null;

        const parsed = JSON.parse(raw);
        if (!parsed || !Array.isArray(parsed.rows) || parsed.rows.length === 0) {
            return null;
        }

        return parsed;
    } catch (error) {
        return null;
    }
}

function normalizeRowsForProcessing(rows) {
    return (rows || [])
        .map(normalizeCsvRow)
        .filter(isValidTransactionRow)
        .filter(isAllowedDisplayChannel)
        .sort((a, b) => getRowTimestamp(a) - getRowTimestamp(b));
}

function applyManualDatasetToReport(dataset) {
    const rows = normalizeRowsForProcessing(dataset?.rows || []);
    if (rows.length === 0) return false;

    googleSheetReportCache = {
        url: 'manual-input://local',
        rows,
        isLoading: false
    };

    renderGoogleSheetChart(rows);

    const createdAt = String(dataset?.createdAt || '').trim();
    const createdAtLabel = createdAt
        ? new Date(createdAt).toLocaleString('id-ID')
        : new Date().toLocaleString('id-ID');
    updateGoogleSheetMeta(`Sumber Laporan: Input Manual tersimpan (${createdAtLabel})`);

    return true;
}

function loadManualDatasetAsPrimary() {
    const manualDataset = getManualInputDataset();
    if (!manualDataset) return false;

    const rows = normalizeRowsForProcessing(manualDataset.rows);
    if (rows.length === 0) return false;

    processData(rows);
    applyManualDatasetToReport(manualDataset);
    return true;
}

function resetManualInputData() {
    const ids = [
        'manualTanggalLaporan',
        'manualLaporanRKUD',
        'manualQrisRKUD',
        'manualMobileBankingRKUD',
        'manualEDCRKUD',
        'manualATMRKUD',
        'manualTellerRKUD',
        'manualTransaksiRKUD',
        'manualLaporanPDRD',
        'manualQrisPDRD',
        'manualMobileBankingPDRD',
        'manualEDCPDRD',
        'manualATMPDRD',
        'manualTellerPDRD',
        'manualTransaksiPDRD',
        'manualLaporanBRIBPHTB',
        'manualQrisBRIBPHTB',
        'manualMobileBankingBRIBPHTB',
        'manualEDCBRIBPHTB',
        'manualATMBRIBPHTB',
        'manualTellerBRIBPHTB',
        'manualTransaksiBRIBPHTB',
        'manualLaporanBRIPBB',
        'manualQrisBRIPBB',
        'manualMobileBankingBRIPBB',
        'manualEDCBRIPBB',
        'manualATMBRIPBB',
        'manualTellerBRIPBB',
        'manualTransaksiBRIPBB',
        'manualLaporanMandiri',
        'manualQrisMandiri',
        'manualMobileBankingMandiri',
        'manualEDCMandiri',
        'manualATMMandiri',
        'manualTellerMandiri',
        'manualTransaksiMandiri',
        'manualTotalPemasukan',
        'manualTotalTransaksi'
    ];

    ids.forEach((id) => {
        const element = document.getElementById(id);
        if (element) {
            element.value = '';
        }
    });

    ensureManualInputDefaultDate();
    updateManualTotalPemasukanPreview();
}

function submitManualInputData() {
    const tanggalLaporanRaw = String(document.getElementById('manualTanggalLaporan')?.value || '').trim();
    const reportAmounts = getManualReportAmountsFromForm(false);
    const reportTransactionChannels = getManualReportTransactionChannelsFromForm(false);
    const reportTransactions = getManualReportTransactionsFromForm(false);
    const totalTransaksi = calculateManualReportTransactionTotal(reportTransactions);
    const totalPemasukan = calculateManualReportTotal(reportAmounts);

    if (!tanggalLaporanRaw) {
        showStatus('❌ Tanggal Laporan wajib diisi.', 'error');
        return;
    }

    const selectedDate = new Date(`${tanggalLaporanRaw}T00:00:00`);
    if (!Number.isFinite(selectedDate.getTime())) {
        showStatus('❌ Format Tanggal Laporan tidak valid.', 'error');
        return;
    }

    if (!Number.isFinite(totalTransaksi) || totalTransaksi <= 0) {
        showStatus('❌ Total Transaksi otomatis harus lebih dari 0 (isi minimal satu laporan transaksi).', 'error');
        return;
    }

    if (!Number.isFinite(totalPemasukan) || totalPemasukan <= 0) {
        showStatus('❌ Total Pemasukan otomatis harus lebih dari 0 (isi minimal satu laporan).', 'error');
        return;
    }

    const pairingError = validateManualReportPairing(reportAmounts, reportTransactions);
    if (pairingError) {
        showStatus(pairingError, 'error');
        return;
    }

    const transactionCount = totalTransaksi;

    const history = getManualInputHistory();
    const nextEntry = {
        tanggalLaporan: tanggalLaporanRaw,
        totalTransaksi: transactionCount,
        ...reportAmounts,
        ...reportTransactionChannels,
        ...reportTransactions,
        totalPemasukan,
        savedAt: new Date().toISOString()
    };

    history.push(nextEntry);

    saveManualInputHistory(history);
    renderManualInputHistoryTable();

    const dataset = buildDatasetFromManualHistory(history);
    const aggregateRows = normalizeRowsForProcessing(dataset.rows);
    if (aggregateRows.length > 0) {
        saveManualInputDataset(aggregateRows, {
            totalRecords: history.length,
            source: 'manual-history'
        });
    }

    document.getElementById('loading').style.display = 'block';
    document.getElementById('dashboard').style.display = 'none';

    setTimeout(() => {
        const syntheticRows = buildSyntheticRowsFromManualRecord(nextEntry, history.length);
        processData(aggregateRows.length > 0 ? aggregateRows : syntheticRows);
        switchDashboardMenu('laporan');
        showStatus(`✅ Data manual diproses (${transactionCount} transaksi estimasi) dengan total ${formatCurrency(totalPemasukan)} dan ditampilkan di Laporan Transaksi.`, 'success');
    }, 80);
}

function triggerElementAnimation(element, className) {
    if (!element) return;
    element.classList.remove(className);
    void element.offsetWidth;
    element.classList.add(className);
}

function updateMetricWithAnimation(metricKey, elementId, nextTextValue) {
    const element = document.getElementById(elementId);
    if (!element) return;

    const nextValue = String(nextTextValue ?? '');
    const previousValue = lastMetricSnapshot[metricKey];
    const hasChanged = previousValue !== undefined && previousValue !== nextValue;

    element.textContent = nextValue;
    if (hasChanged) {
        triggerElementAnimation(element, 'metric-updated');
        const card = element.closest('.stat-card, .summary-card, .chart-container, .summary-decision-item');
        if (card) {
            triggerElementAnimation(card, 'card-updated');
        }
    }

    lastMetricSnapshot[metricKey] = nextValue;
}

function initializeMetricSnapshot() {
    if (Object.keys(lastMetricSnapshot).length > 0) return;
    const trackedIds = [
        'totalCredit',
        'totalTransactions',
        'currentDatasetLabel',
        'currentDatasetSub',
        'summaryTopChannel',
        'summaryTopRevenue',
        'summaryTopSource',
        'summaryAvgCredit',
        'summaryAction',
        'summaryTargetWarning',
        'summaryTrendInsight',
        'summarySegmentationInsight',
        'summaryQrisInsight',
        'summarySimulationInsight',
        'summaryHealthScore'
    ];

    trackedIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            lastMetricSnapshot[id] = String(element.textContent || '');
        }
    });
}

function initDashboardEnhancements() {
    const searchInput = document.getElementById('dashboardSearch');
    if (searchInput) {
        searchInput.addEventListener('input', handleDashboardSearch);
    }

    initializeMetricSnapshot();
    updateActiveMenuChip(activeDashboardMenu);
    ensureManualInputDefaultDate();
    updateManualTotalPemasukanPreview();
    renderManualInputHistoryTable();
    renderActivityLogs();
    updateDashboardClock();
    setInterval(updateDashboardClock, 1000);
}

function updateThemeToggleUI(theme) {
    const toggleIcon = document.getElementById('themeToggleIcon');
    const toggleText = document.getElementById('themeToggleText');
    if (toggleIcon) {
        toggleIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
    }
    if (toggleText) {
        toggleText.textContent = theme === 'dark' ? 'Light Mode' : 'Dark Mode';
    }
}

function applyTheme(theme) {
    const isDark = theme === 'dark';
    document.body.classList.toggle('dark-mode', isDark);
    updateThemeToggleUI(isDark ? 'dark' : 'light');
}

function toggleTheme() {
    const nextTheme = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
    applyTheme(nextTheme);
    localStorage.setItem(THEME_STORAGE_KEY, nextTheme);
}

function initTheme() {
    const storedTheme = String(localStorage.getItem(THEME_STORAGE_KEY) || 'light').toLowerCase();
    const theme = storedTheme === 'dark' ? 'dark' : 'light';
    applyTheme(theme);
}

// Check authentication
function checkAuth() {
    const session = localStorage.getItem('userSession');
    
    if (!session) {
        // Belum login, redirect ke login page
        window.location.href = 'login.html';
        return false;
    }
    
    try {
        currentUser = JSON.parse(session);
        updateUserDisplay();
        applyPermissions();
        logActivityEvent('Login Berhasil', `User ${currentUser?.username || '-'} berhasil masuk ke dashboard.`, {
            level: 'success',
            category: 'autentikasi',
            source: 'auth'
        });
        return true;
    } catch (error) {
        console.error('Error parsing session:', error);
        logActivityEvent('Session Tidak Valid', 'Data sesi login tidak dapat diproses.', {
            level: 'error',
            category: 'autentikasi',
            source: 'auth'
        });
        localStorage.removeItem('userSession');
        window.location.href = 'login.html';
        return false;
    }
}

// Update user display
function updateUserDisplay() {
    if (!currentUser) return;

    const userNameElement = document.getElementById('userName');
    if (userNameElement) {
        userNameElement.textContent = currentUser.username || currentUser.fullName || '-';
    }

    const sidebarBrandTitle = document.getElementById('sidebarBrandTitle');
    if (sidebarBrandTitle) {
        const displayName = currentUser.fullName || currentUser.username || 'User';
        sidebarBrandTitle.textContent = `📁 ${displayName}`;
    }
}

// Apply permissions
function applyPermissions() {
    if (!currentUser) return;
    
    const role = currentUser.role || 'guest';
    const permissions = currentUser.permissions || [];
    const allowedMenus = getAllowedMenusByRole(role);

    const menuButtonMap = {
        sorotan: document.getElementById('menuBtnSorotan'),
        grafik: document.getElementById('menuBtnGrafik'),
        laporan: document.getElementById('menuBtnLaporan'),
        manual: document.getElementById('menuBtnManual'),
        user: document.getElementById('menuBtnUser'),
        password: document.getElementById('menuBtnPassword'),
        'activity-log': document.getElementById('menuBtnActivityLog')
    };

    const menuSectionMap = {
        sorotan: document.getElementById('menuSectionSorotan'),
        grafik: document.getElementById('menuSectionGrafik'),
        laporan: document.getElementById('menuSectionLaporan'),
        manual: document.getElementById('menuSectionManual'),
        user: document.getElementById('menuSectionUser'),
        password: document.getElementById('menuSectionPassword'),
        'activity-log': document.getElementById('menuSectionActivityLog')
    };

    Object.entries(menuButtonMap).forEach(([menuKey, element]) => {
        if (!element) return;
        element.style.display = allowedMenus.includes(menuKey) ? '' : 'none';
    });

    Object.entries(menuSectionMap).forEach(([menuKey, element]) => {
        if (!element) return;
        element.style.display = allowedMenus.includes(menuKey) ? '' : 'none';
    });

    const uploadSection = document.getElementById('uploadSection');
    if (uploadSection) {
        uploadSection.style.display = role === 'guest' ? 'none' : '';
    }

    const exportPdfButton = document.getElementById('menuBtnExportPdf');
    if (exportPdfButton) {
        exportPdfButton.style.display = role === 'guest' ? 'none' : '';
    }
    
    // Check upload permission
    if (!permissions.includes('upload')) {
        const btnLoad = document.getElementById('btnLoadCSV');
        const fileInput = document.getElementById('csvFile');

        if (btnLoad) btnLoad.disabled = true;
        if (fileInput) fileInput.disabled = true;
        
        if (role !== 'guest') {
            showStatus('⚠️ Anda hanya memiliki akses view. Upload CSV dinonaktifkan.', 'warning');
        }
    } else {
        const btnLoad = document.getElementById('btnLoadCSV');
        const fileInput = document.getElementById('csvFile');
        if (btnLoad) btnLoad.disabled = false;
        if (fileInput) fileInput.disabled = false;
    }

    if (!canAccessMenu(activeDashboardMenu)) {
        switchDashboardMenu(getFallbackMenuForCurrentRole());
    }
}

// Handle logout
function handleLogout() {
    if (confirm('Apakah Anda yakin ingin logout?')) {
        logActivityEvent('Logout', `User ${currentUser?.username || '-'} keluar dari dashboard.`, {
            level: 'info',
            category: 'autentikasi',
            source: 'auth'
        });
        localStorage.removeItem('userSession');
        window.location.href = 'login.html';
    }
}

// Parse number from Indonesian format
function parseAmount(str) {
    if (str === null || str === undefined || str === '') return 0;
    if (typeof str === 'number') return Number.isFinite(str) ? str : 0;

    const normalized = String(str).trim().replace(/\s+/g, '');
    if (!normalized) return 0;
    return parseFloat(normalized.replace(/\./g, '').replace(',', '.')) || 0;
}

// Format currency
function formatCurrency(amount) {
    return 'Rp ' + amount.toLocaleString('id-ID', { 
        minimumFractionDigits: 2, 
        maximumFractionDigits: 2 
    });
}

function formatCurrencyTick(value) {
    return formatCurrency(Number(value));
}

function formatDistributionTooltipLabel(context) {
    const total = context.dataset.data.reduce((sum, current) => sum + current, 0);
    const percentage = total > 0 ? ((context.parsed / total) * 100).toFixed(1) : '0.0';
    return context.label + ': ' + formatCurrency(context.parsed) + ' (' + percentage + '%)';
}

function getTopEntry(dataObject) {
    const entries = Object.entries(dataObject || {});
    if (entries.length === 0) {
        return { label: '-', value: 0 };
    }

    const [label, value] = entries.reduce((maxEntry, currentEntry) => {
        return currentEntry[1] > maxEntry[1] ? currentEntry : maxEntry;
    });

    return { label, value };
}

function formatHourRange(hourIndex) {
    if (hourIndex < 0 || hourIndex > 23) return '-';
    const nextHour = (hourIndex + 1) % 24;
    return `${String(hourIndex).padStart(2, '0')}:00-${String(nextHour).padStart(2, '0')}:00`;
}

function formatDateLabel(timestamp) {
    const date = new Date(timestamp);
    if (isNaN(date.getTime())) return '-';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function normalizeDateParsingMode(mode) {
    const safeMode = String(mode || '').trim().toLowerCase();
    if (safeMode === 'ddmm' || safeMode === 'mmdd' || safeMode === 'auto') {
        return safeMode;
    }
    return 'auto';
}

function loadDateParsingMode() {
    try {
        const storedMode = localStorage.getItem(DATE_PARSING_MODE_STORAGE_KEY);
        dateParsingMode = normalizeDateParsingMode(storedMode);
    } catch (error) {
        dateParsingMode = 'auto';
    }
    return dateParsingMode;
}

function setDateParsingMode(mode) {
    const normalizedMode = normalizeDateParsingMode(mode);
    dateParsingMode = normalizedMode;

    try {
        localStorage.setItem(DATE_PARSING_MODE_STORAGE_KEY, normalizedMode);
    } catch (error) {
        console.warn('Gagal menyimpan mode parser tanggal:', error);
    }

    if (Array.isArray(latestProcessedRows) && latestProcessedRows.length > 0) {
        processData(latestProcessedRows);
    }
}

function getMonthNumberFromText(value) {
    const text = String(value || '').toLowerCase();
    if (!text) return null;

    const monthPattern = /(januari|january|februari|february|maret|march|april|mei|may|juni|june|juli|july|agustus|august|september|oktober|october|november|desember|december)/;
    const match = text.match(monthPattern);
    if (!match) return null;

    const monthMap = {
        januari: 1, january: 1,
        februari: 2, february: 2,
        maret: 3, march: 3,
        april: 4,
        mei: 5, may: 5,
        juni: 6, june: 6,
        juli: 7, july: 7,
        agustus: 8, august: 8,
        september: 9,
        oktober: 10, october: 10,
        november: 11,
        desember: 12, december: 12
    };

    return monthMap[match[1]] || null;
}

function parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, order) {
    const first = Number.parseInt(String(firstPart || ''), 10);
    const second = Number.parseInt(String(secondPart || ''), 10);
    const yearRaw = String(yearPart || '').trim();
    if (!Number.isInteger(first) || !Number.isInteger(second) || !yearRaw) return null;

    const day = order === 'ddmm' ? first : second;
    const month = order === 'ddmm' ? second : first;
    if (day < 1 || day > 31 || month < 1 || month > 12) return null;

    const fullYear = yearRaw.length === 2 ? `20${yearRaw}` : yearRaw;
    const isoDate = `${String(fullYear).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${normalizedTime}`;
    const parsedDate = Date.parse(isoDate);
    return Number.isNaN(parsedDate) ? null : parsedDate;
}

function parseAmbiguousSlashDate(firstPart, secondPart, yearPart, normalizedTime, row) {
    const first = Number.parseInt(String(firstPart || ''), 10);
    const second = Number.parseInt(String(secondPart || ''), 10);
    if (!Number.isInteger(first) || !Number.isInteger(second)) return null;

    if (first > 12 && second <= 12) {
        return parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'ddmm');
    }

    if (second > 12 && first <= 12) {
        return parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'mmdd');
    }

    const monthFromLabel = getMonthNumberFromText(row?.Bulan);
    if (Number.isInteger(monthFromLabel)) {
        const ddmmCandidate = parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'ddmm');
        const mmddCandidate = parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'mmdd');
        if (ddmmCandidate !== null && mmddCandidate !== null) {
            const ddmmMonth = new Date(ddmmCandidate).getMonth() + 1;
            const mmddMonth = new Date(mmddCandidate).getMonth() + 1;
            if (ddmmMonth === monthFromLabel && mmddMonth !== monthFromLabel) return ddmmCandidate;
            if (mmddMonth === monthFromLabel && ddmmMonth !== monthFromLabel) return mmddCandidate;
        }
    }

    return parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'ddmm');
}

function parseRowTimestamp(row) {
    const postDateText = String(row?.PostDate || '').trim();

    const rawDateText = String(row?.Tanggal || '').trim();
    const rawTimeText = String(row?.Jam || '').trim();
    if (!rawDateText && !postDateText) return null;

    let dateText = rawDateText || postDateText;
    let timeText = rawTimeText || '00:00:00';

    const splitDateTime = dateText.split(/\s+/);
    if (!rawTimeText && splitDateTime.length >= 2) {
        const possibleTime = splitDateTime[splitDateTime.length - 1];
        if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(possibleTime)) {
            timeText = possibleTime;
            dateText = splitDateTime.slice(0, -1).join(' ');
        }
    }

    const normalizedTime = /^\d{1,2}:\d{2}(:\d{2})?$/.test(timeText) ? (timeText.length === 5 ? `${timeText}:00` : timeText) : '00:00:00';

    const slashMatch = dateText.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (slashMatch) {
        const firstPart = slashMatch[1];
        const secondPart = slashMatch[2];
        const yearPart = slashMatch[3];

        const mode = normalizeDateParsingMode(dateParsingMode);
        let parsedSlashDate = null;

        if (mode === 'ddmm') {
            parsedSlashDate = parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'ddmm');
        } else if (mode === 'mmdd') {
            parsedSlashDate = parseSlashDateByOrder(firstPart, secondPart, yearPart, normalizedTime, 'mmdd');
        } else {
            parsedSlashDate = parseAmbiguousSlashDate(firstPart, secondPart, yearPart, normalizedTime, row);
        }

        if (parsedSlashDate !== null) return parsedSlashDate;
    }

    const normalizedDateText = dateText.replace(/\s+/g, ' ').toLowerCase();
    const monthMap = {
        januari: '01', january: '01',
        februari: '02', february: '02',
        maret: '03', march: '03',
        april: '04',
        mei: '05', may: '05',
        juni: '06', june: '06',
        juli: '07', july: '07',
        agustus: '08', august: '08',
        september: '09',
        oktober: '10', october: '10',
        november: '11',
        desember: '12', december: '12'
    };

    const monthNameMatch = normalizedDateText.match(/^(\d{1,2})\s+([a-z]+)\s+(\d{2,4})$/);
    if (monthNameMatch) {
        const day = monthNameMatch[1];
        const monthName = monthNameMatch[2];
        const month = monthMap[monthName];
        const yearRaw = monthNameMatch[3];
        const fullYear = yearRaw.length === 2 ? `20${yearRaw}` : yearRaw;

        if (month) {
            const isoDate = `${String(fullYear).padStart(4, '0')}-${month}-${String(day).padStart(2, '0')}T${normalizedTime}`;
            const parsedMonthNameDate = Date.parse(isoDate);
            if (!isNaN(parsedMonthNameDate)) return parsedMonthNameDate;
        }
    }

    const directDate = Date.parse(`${dateText} ${normalizedTime}`);
    if (!isNaN(directDate)) return directDate;

    return null;
}

function isDigitalChannel(channelName) {
    return /qris|mobile|internet|virtual|\bva\b|edc|atm|transfer|digital/i.test(String(channelName || ''));
}

function isTellerChannel(channelName) {
    return /teller|teler/i.test(String(channelName || ''));
}

function isQrisChannel(channelName) {
    return /qris/i.test(String(channelName || ''));
}

function isMobileChannel(channelName) {
    return /mobile/i.test(String(channelName || ''));
}

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function formatPercent(value) {
    return `${(Number(value || 0) * 100).toFixed(1)}%`;
}

function predictNextValue(series, options = {}) {
    const values = (series || []).map(value => Number(value)).filter(value => Number.isFinite(value));
    if (values.length === 0) return 0;
    if (values.length === 1) return values[0];

    const deltas = [];
    for (let index = 1; index < values.length; index += 1) {
        deltas.push(values[index] - values[index - 1]);
    }

    const weighted = deltas.reduce((accumulator, delta, index) => {
        const weight = index + 1;
        accumulator.sum += delta * weight;
        accumulator.weight += weight;
        return accumulator;
    }, { sum: 0, weight: 0 });

    const trendDelta = weighted.weight > 0 ? (weighted.sum / weighted.weight) : 0;
    const rawPrediction = values[values.length - 1] + trendDelta;

    const minValue = Number.isFinite(options.min) ? options.min : null;
    const maxValue = Number.isFinite(options.max) ? options.max : null;

    if (minValue !== null && maxValue !== null) return clamp(rawPrediction, minValue, maxValue);
    if (minValue !== null) return Math.max(minValue, rawPrediction);
    if (maxValue !== null) return Math.min(maxValue, rawPrediction);
    return rawPrediction;
}

function predictAverageValue(series) {
    const values = (series || []).map(value => Number(value)).filter(value => Number.isFinite(value) && value >= 0);
    if (values.length === 0) return 0;
    if (values.length === 1) return values[0];

    if (values.length === 2) {
        return (values[1] * 0.7) + (values[0] * 0.3);
    }

    const trendPrediction = predictNextValue(values, { min: 0 });
    const recentAverage = (values[values.length - 1] + values[values.length - 2] + values[values.length - 3]) / 3;
    const blendedPrediction = (trendPrediction * 0.6) + (recentAverage * 0.4);
    const conservativeFloor = values[values.length - 1] * 0.65;

    return Math.max(conservativeFloor, blendedPrediction);
}

function getNextMonthLabel(lastMonthKey) {
    const [yearText, monthText] = String(lastMonthKey || '').split('-');
    const year = Number.parseInt(yearText, 10);
    const monthIndex = Number.parseInt(monthText, 10) - 1;
    if (!Number.isFinite(year) || !Number.isFinite(monthIndex)) return 'Prediksi Bulan Depan';

    const nextMonthDate = new Date(year, monthIndex + 1, 1);
    return `Prediksi ${nextMonthDate.toLocaleDateString('id-ID', { month: 'short', year: 'numeric' })}`;
}

function calculateDiversificationScore(channelTxMap) {
    const values = Object.values(channelTxMap || {}).filter(value => value > 0);
    if (values.length <= 1) return values.length === 1 ? 20 : 0;

    const total = values.reduce((sum, value) => sum + value, 0);
    if (total <= 0) return 0;

    const hhi = values.reduce((sum, value) => {
        const share = value / total;
        return sum + (share * share);
    }, 0);

    const normalized = (1 - hhi) / (1 - (1 / values.length));
    return clamp(normalized * 100, 0, 100);
}

function buildStrategicInsights(data) {
    const monthlyMap = new Map();

    const findLargestDeltaEntry = (currentMap, previousMap, direction = 'up') => {
        const mergedKeys = new Set([
            ...Object.keys(currentMap || {}),
            ...Object.keys(previousMap || {})
        ]);

        let best = null;
        mergedKeys.forEach((key) => {
            const currentValue = Number(currentMap?.[key] || 0);
            const previousValue = Number(previousMap?.[key] || 0);
            const diff = currentValue - previousValue;

            if (direction === 'up' && diff <= 0) return;
            if (direction === 'down' && diff >= 0) return;

            if (!best || Math.abs(diff) > Math.abs(best.diff)) {
                best = { key, diff, currentValue, previousValue };
            }
        });

        return best;
    };

    data.forEach(row => {
        const timestamp = getRowTimestamp(row);
        if (!timestamp) return;

        const date = new Date(timestamp);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const monthKey = `${year}-${month}`;
        const monthLabel = date.toLocaleDateString('id-ID', { month: 'short', year: 'numeric' });

        if (!monthlyMap.has(monthKey)) {
            monthlyMap.set(monthKey, {
                key: monthKey,
                label: monthLabel,
                totalTx: 0,
                totalCredit: 0,
                digitalTx: 0,
                qrisTx: 0,
                tellerTx: 0,
                mobileTx: 0,
                hourly: Array(24).fill(0),
                channelTx: {},
                revenueCredit: {},
                sourceCredit: {}
            });
        }

        const monthData = monthlyMap.get(monthKey);
        const credit = parseAmount(row['Credit Amount']);
        const channelName = resolveChannelName(row);
        const revenueName = row['Jenis Penerimaan'] || 'Lainnya';
        const sourceName = resolveSourceName(row);
        const hourText = String(row.Jam || '').trim();

        monthData.totalTx += 1;
        monthData.totalCredit += credit;
        monthData.channelTx[channelName] = (monthData.channelTx[channelName] || 0) + 1;

        if (isDigitalChannel(channelName)) monthData.digitalTx += 1;
        if (isQrisChannel(channelName)) monthData.qrisTx += 1;
        if (isTellerChannel(channelName)) monthData.tellerTx += 1;
        if (isMobileChannel(channelName)) monthData.mobileTx += 1;

        if (revenueName && credit > 0) {
            monthData.revenueCredit[revenueName] = (monthData.revenueCredit[revenueName] || 0) + credit;
        }

        if (sourceName && credit > 0) {
            monthData.sourceCredit[sourceName] = (monthData.sourceCredit[sourceName] || 0) + credit;
        }

        if (/^\d{1,2}:\d{2}/.test(hourText)) {
            const hour = Number.parseInt(hourText.split(':')[0], 10);
            if (Number.isFinite(hour) && hour >= 0 && hour <= 23) {
                monthData.hourly[hour] += 1;
            }
        }
    });

    const now = new Date();
    const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const sortedMonths = Array.from(monthlyMap.values())
        .sort((a, b) => a.key.localeCompare(b.key));

    const nonFutureMonths = sortedMonths.filter(item => String(item.key) <= currentMonthKey);
    const currentYearMonths = nonFutureMonths.filter(item => String(item.key).startsWith(`${now.getFullYear()}-`));
    const trendSourceMonths = currentYearMonths.length > 0
        ? currentYearMonths
        : (nonFutureMonths.length > 0 ? nonFutureMonths : sortedMonths);

    const orderedMonths = trendSourceMonths
        .slice(-6)
        .map(monthData => {
            const totalTx = monthData.totalTx || 0;
            const totalCredit = monthData.totalCredit || 0;
            const topRevenue = getTopEntry(monthData.revenueCredit);
            const topSource = getTopEntry(monthData.sourceCredit);
            const peakHourCount = Math.max(...monthData.hourly);
            const peakHourIndex = peakHourCount > 0 ? monthData.hourly.indexOf(peakHourCount) : -1;

            return {
                ...monthData,
                digitalShare: totalTx > 0 ? monthData.digitalTx / totalTx : 0,
                qrisShare: totalTx > 0 ? monthData.qrisTx / totalTx : 0,
                nonCashShare: totalTx > 0 ? (totalTx - monthData.tellerTx) / totalTx : 0,
                mobileShare: totalTx > 0 ? monthData.mobileTx / totalTx : 0,
                avgValue: totalTx > 0 ? totalCredit / totalTx : 0,
                topRevenueLabel: topRevenue.label,
                topRevenueShare: totalCredit > 0 ? topRevenue.value / totalCredit : 0,
                topSourceLabel: topSource.label,
                topSourceShare: totalCredit > 0 ? topSource.value / totalCredit : 0,
                peakHourIndex
            };
        });

    if (orderedMonths.length === 0) {
        return {
            available: false,
            targetWarning: 'Data tanggal belum cukup untuk membaca target bulanan dan peringatan dini.',
            trendInsight: 'Unggah data minimal 2 bulan agar tren mudah dibandingkan.',
            segmentationInsight: 'Sumber penerimaan belum bisa dipetakan.',
            qrisInsight: 'Data kanal digital belum cukup untuk dihitung.',
            peakInsight: 'Pola jam transaksi belum terbaca.',
            simulationInsight: 'Simulasi kebijakan belum tersedia karena data belum cukup.',
            healthScoreText: 'Belum bisa dinilai',
            healthStatus: 'red',
            monthlyTrend: null
        };
    }

    const latest = orderedMonths[orderedMonths.length - 1];
    const previous = orderedMonths.length >= 2 ? orderedMonths[orderedMonths.length - 2] : null;

    const digitalDrop = previous ? (previous.digitalShare - latest.digitalShare) : 0;
    const avgDrop = previous && previous.avgValue > 0 ? ((previous.avgValue - latest.avgValue) / previous.avgValue) : 0;
    const earlyWarnings = [];
    if (digitalDrop > 0.05) earlyWarnings.push(`porsi digital turun ${formatPercent(digitalDrop)} vs bulan lalu`);
    if (avgDrop > 0.1) earlyWarnings.push(`rata-rata nilai transaksi turun ${formatPercent(avgDrop)}`);

    const digitalTarget = 0.8;
    const targetGap = digitalTarget - latest.digitalShare;
    const nonDigitalShare = Math.max(0, 1 - latest.digitalShare);
    const dominanceText = latest.digitalShare >= 0.5
        ? 'Komposisi transaksi saat ini didominasi Kanal Digital.'
        : 'Komposisi transaksi saat ini masih didominasi Kanal Non Digital (Teller).';
    const targetText = targetGap <= 0
        ? `Target Kanal Digital sudah tercapai. Capaian Kanal Digital saat ini ${formatPercent(latest.digitalShare)} (target ${formatPercent(digitalTarget)}), sedangkan Kanal Non Digital (Teller) ${formatPercent(nonDigitalShare)}. ${dominanceText}`
        : `Target Kanal Digital belum tercapai. Capaian Kanal Digital saat ini ${formatPercent(latest.digitalShare)}, masih kurang ${formatPercent(targetGap)} dari target ${formatPercent(digitalTarget)}. Kanal Non Digital (Teller) saat ini ${formatPercent(nonDigitalShare)}. ${dominanceText}`;
    const warningText = earlyWarnings.length > 0
        ? `Perlu perhatian: ${earlyWarnings.join('; ')}.`
        : 'Kondisi aman: belum ada penurunan penting pada bulan ini.';


    const first = orderedMonths[0];
    const digitalDelta = latest.digitalShare - first.digitalShare;
    const avgDelta = first.avgValue > 0 ? ((latest.avgValue - first.avgValue) / first.avgValue) : 0;

    const historicalDigital = orderedMonths.map(item => Number((item.digitalShare * 100).toFixed(2)));
    const historicalQris = orderedMonths.map(item => Number((item.qrisShare * 100).toFixed(2)));
    const historicalRevenueDominance = orderedMonths.map(item => Number((item.topRevenueShare * 100).toFixed(2)));
    const historicalAvgValue = orderedMonths.map(item => item.avgValue);

    const forecastDigital = Number(predictNextValue(historicalDigital, { min: 0, max: 100 }).toFixed(2));
    const forecastQris = Number(predictNextValue(historicalQris, { min: 0, max: 100 }).toFixed(2));
    const forecastRevenueDominance = Number(predictNextValue(historicalRevenueDominance, { min: 0, max: 100 }).toFixed(2));
    const forecastAvgValue = Number(predictAverageValue(historicalAvgValue).toFixed(2));

    const trendInsight = `Dalam ${orderedMonths.length} bulan terakhir, porsi digital ${digitalDelta >= 0 ? 'naik' : 'turun'} ${formatPercent(Math.abs(digitalDelta))}, nilai rata-rata transaksi ${avgDelta >= 0 ? 'naik' : 'turun'} ${formatPercent(Math.abs(avgDelta))}. Prediksi bulan depan: digital sekitar ${forecastDigital.toFixed(1)}% dan rata-rata nilai transaksi sekitar ${formatCurrency(forecastAvgValue)}.`;

    const hasMonthActivity = (monthData) => {
        return Number(monthData?.totalTx || 0) > 0 || Number(monthData?.totalCredit || 0) > 0;
    };

    const movementCandidates = orderedMonths.filter(hasMonthActivity);
    let movementLatest = movementCandidates[movementCandidates.length - 1] || latest;
    let movementPrevious = movementCandidates.length >= 2 ? movementCandidates[movementCandidates.length - 2] : null;

    if (movementLatest && movementPrevious && movementLatest.key === currentMonthKey) {
        const latestCredit = Number(movementLatest.totalCredit || 0);
        const latestTx = Number(movementLatest.totalTx || 0);
        const previousCredit = Number(movementPrevious.totalCredit || 0);
        const previousTx = Number(movementPrevious.totalTx || 0);

        const creditLikelyPartial = previousCredit > 0 && latestCredit < (previousCredit * 0.6);
        const txLikelyPartial = previousTx > 0 && latestTx < (previousTx * 0.6);

        if ((creditLikelyPartial || txLikelyPartial) && movementCandidates.length >= 3) {
            movementLatest = movementCandidates[movementCandidates.length - 2];
            movementPrevious = movementCandidates[movementCandidates.length - 3];
        }
    }

    let monthMovementInsight = `Belum ada pembanding bulan sebelumnya untuk ${latest.label}.`;
    if (movementPrevious && movementLatest) {
        const latestLabel = String(movementLatest.label || 'Bulan ini');
        const previousLabel = String(movementPrevious.label || 'bulan sebelumnya');
        const creditDelta = Number(movementLatest.totalCredit || 0) - Number(movementPrevious.totalCredit || 0);
        const txDelta = Number(movementLatest.totalTx || 0) - Number(movementPrevious.totalTx || 0);
        const avgValueDelta = Number(movementLatest.avgValue || 0) - Number(movementPrevious.avgValue || 0);
        const creditDeltaRatio = movementPrevious.totalCredit > 0 ? (creditDelta / movementPrevious.totalCredit) : 0;
        const txDeltaRatio = movementPrevious.totalTx > 0 ? (txDelta / movementPrevious.totalTx) : 0;
        const avgValueDeltaRatio = movementPrevious.avgValue > 0 ? (avgValueDelta / movementPrevious.avgValue) : 0;

        const direction = creditDelta >= 0 ? 'naik' : 'turun';
        const reasonParts = [];

        if (creditDelta >= 0) {
            if (txDelta > 0) {
                reasonParts.push(`jumlah transaksi naik ${Math.abs(txDelta).toLocaleString('id-ID')} (${formatPercent(Math.abs(txDeltaRatio))})`);
            }
            if (avgValueDelta > 0) {
                reasonParts.push(`rata-rata nilai transaksi naik ${formatPercent(Math.abs(avgValueDeltaRatio))}`);
            }

            const topRevenueUp = findLargestDeltaEntry(movementLatest.revenueCredit, movementPrevious.revenueCredit, 'up');
            if (topRevenueUp && topRevenueUp.diff > 0) {
                reasonParts.push(`kenaikan terbesar dari jenis penerimaan ${topRevenueUp.key} sebesar ${formatCurrency(topRevenueUp.diff)}`);
            }

            const topChannelUp = findLargestDeltaEntry(movementLatest.channelTx, movementPrevious.channelTx, 'up');
            if (topChannelUp && topChannelUp.diff > 0) {
                reasonParts.push(`aktivitas kanal ${topChannelUp.key} naik ${Math.abs(topChannelUp.diff).toLocaleString('id-ID')} transaksi`);
            }
        } else {
            if (txDelta < 0) {
                reasonParts.push(`jumlah transaksi turun ${Math.abs(txDelta).toLocaleString('id-ID')} (${formatPercent(Math.abs(txDeltaRatio))})`);
            }
            if (avgValueDelta < 0) {
                reasonParts.push(`rata-rata nilai transaksi turun ${formatPercent(Math.abs(avgValueDeltaRatio))}`);
            }

            const topRevenueDown = findLargestDeltaEntry(movementLatest.revenueCredit, movementPrevious.revenueCredit, 'down');
            if (topRevenueDown && topRevenueDown.diff < 0) {
                reasonParts.push(`penurunan terbesar dari jenis penerimaan ${topRevenueDown.key} sebesar ${formatCurrency(Math.abs(topRevenueDown.diff))}`);
            }

            const topChannelDown = findLargestDeltaEntry(movementLatest.channelTx, movementPrevious.channelTx, 'down');
            if (topChannelDown && topChannelDown.diff < 0) {
                reasonParts.push(`aktivitas kanal ${topChannelDown.key} turun ${Math.abs(topChannelDown.diff).toLocaleString('id-ID')} transaksi`);
            }
        }

        const headline = `Grafik ${latestLabel} ${direction} ${formatPercent(Math.abs(creditDeltaRatio))} dibanding ${previousLabel}.`;
        monthMovementInsight = reasonParts.length > 0
            ? `${headline} Penyebab utama: ${reasonParts.join('; ')}.`
            : `${headline} Perubahan dipengaruhi kombinasi volume transaksi dan komposisi kanal/jenis penerimaan.`;
    }

    const segmentationInsight = latest.topSourceShare >= 0.6
        ? `Penerimaan masih sangat bergantung pada ${latest.topSourceLabel} (${formatPercent(latest.topSourceShare)}). Disarankan tambah kolaborasi bank persepsi/BUMD agar risiko lebih terjaga.`
        : `Sumber penerimaan cukup berimbang. Kontributor terbesar saat ini ${latest.topSourceLabel} (${formatPercent(latest.topSourceShare)}).`;

    const digitalShare = Number(latest.digitalShare || 0);
    const digitalStrategyNonDigitalShare = Math.max(0, 1 - digitalShare);
    const topDigitalChannel = Object.entries(latest.channelTx || {})
        .filter(([channelName]) => isDigitalChannel(channelName))
        .sort((a, b) => Number(b[1] || 0) - Number(a[1] || 0))[0] || ['-'];
    const topDigitalChannelLabel = String(topDigitalChannel[0] || '-');
    const qrisInsight = `Porsi Kanal Digital (QRIS, Mobile Banking, EDC, ATM) saat ini ${formatPercent(digitalShare)}, sedangkan Kanal Non Digital (Teller) ${formatPercent(digitalStrategyNonDigitalShare)}. Strategi untuk pimpinan: fokuskan akselerasi kanal digital pada kanal dominan (${topDigitalChannelLabel}), lakukan migrasi bertahap transaksi Teller ke kanal digital, serta tetapkan target bulanan per OPD/unit dengan monitoring persentase digital vs non digital.`;

    const peakInsight = latest.peakHourIndex >= 0
        ? `Jam tersibuk ada di ${formatHourRange(latest.peakHourIndex)}. Fokuskan kapasitas server/helpdesk di jam ini, sambil dorong transaksi di jam lain agar beban lebih merata.`
        : 'Jam puncak belum teridentifikasi.';

    const additionalDigitalTx = latest.totalTx * 0.05;
    const simulationPad = additionalDigitalTx * latest.avgValue;
    const simulationInsight = `Perkiraan cepat: jika transaksi digital naik 5%, potensi kenaikan PAD sekitar ${formatCurrency(simulationPad)} dengan asumsi nilai rata-rata transaksi tetap.`;

    const digitalScore = clamp((latest.digitalShare / 0.8) * 100, 0, 100);
    const growthRate = previous && previous.totalTx > 0 ? ((latest.totalTx - previous.totalTx) / previous.totalTx) : 0;
    const growthScore = clamp(50 + (growthRate * 250), 0, 100);
    const nonCashScore = clamp((latest.nonCashShare / 0.8) * 100, 0, 100);
    const diversificationScore = calculateDiversificationScore(latest.channelTx);

    const healthIndex = (
        (digitalScore * 0.35) +
        (growthScore * 0.25) +
        (nonCashScore * 0.2) +
        (diversificationScore * 0.2)
    );

    const healthStatus = healthIndex >= 75 ? 'green' : (healthIndex >= 55 ? 'yellow' : 'red');
    const healthReadableText = healthStatus === 'green'
        ? 'Kesiapan Digital: Baik'
        : (healthStatus === 'yellow' ? 'Kesiapan Digital: Cukup (perlu peningkatan)' : 'Kesiapan Digital: Perlu perhatian');

    return {
        available: true,
        targetWarning: `${targetText} ${warningText}`,
        trendInsight: monthMovementInsight,
        segmentationInsight,
        qrisInsight,
        peakInsight,
        simulationInsight,
        healthScoreText: `${healthReadableText} — nilai ${healthIndex.toFixed(1)}/100`,
        healthStatus,
        monthlyTrend: {
            labels: orderedMonths.map(item => item.label),
            labelsWithForecast: [...orderedMonths.map(item => item.label), getNextMonthLabel(latest.key)],
            digitalShare: [...historicalDigital, forecastDigital],
            qrisShare: [...historicalQris, forecastQris],
            dominantRevenueShare: [...historicalRevenueDominance, forecastRevenueDominance],
            avgValue: [...historicalAvgValue, forecastAvgValue],
            historicalLength: orderedMonths.length
        }
    };
}

function updateExecutiveSummary(totalCredit, transactionCount, channelData, channelTransactionData, revenueTypeData, hourlyData, fileSourceData, strategicInsights) {
    const topChannel = getTopEntry(channelData);
    const topChannelByCount = getTopEntry(channelTransactionData);
    const topRevenue = getTopEntry(revenueTypeData);
    const topSource = getTopEntry(fileSourceData);

    const channelEntries = Object.entries(channelTransactionData || {});
    const tellerCount = Object.entries(channelTransactionData || {}).reduce((total, [channelName, count]) => {
        return /teller|teler/i.test(channelName) ? total + count : total;
    }, 0);
    const nonTellerCount = Math.max(0, transactionCount - tellerCount);
    const digitalCount = channelEntries.reduce((total, [channelName, count]) => {
        return /qris|mobile|internet|virtual|\bva\b|edc|atm|transfer|digital/i.test(channelName)
            ? total + count
            : total;
    }, 0);
    const qrisCount = channelEntries.reduce((total, [channelName, count]) => {
        return /qris/i.test(channelName) ? total + count : total;
    }, 0);

    const tellerRatio = transactionCount > 0 ? tellerCount / transactionCount : 0;
    const digitalRatio = transactionCount > 0 ? digitalCount / transactionCount : 0;
    const qrisRatio = transactionCount > 0 ? qrisCount / transactionCount : 0;
    const topSourceRatio = totalCredit > 0 ? topSource.value / totalCredit : 0;

    const avgCredit = transactionCount > 0 ? totalCredit / transactionCount : 0;

    updateMetricWithAnimation('summaryTopChannel', 'summaryTopChannel', `${topChannel.label} (${formatCurrency(topChannel.value)})`);
    updateMetricWithAnimation('summaryTopRevenue', 'summaryTopRevenue', `${topRevenue.label} (${formatCurrency(topRevenue.value)})`);
    updateMetricWithAnimation('summaryTopSource', 'summaryTopSource', `${topSource.label} (${formatCurrency(topSource.value)})`);
    updateMetricWithAnimation('summaryAvgCredit', 'summaryAvgCredit', formatCurrency(avgCredit));

    const actionElement = document.getElementById('summaryAction');
    const recommendationsElement = document.getElementById('summaryRecommendations');
    const statusIconElement = document.getElementById('summaryStatusIcon');
    actionElement.className = 'summary-action';
    if (statusIconElement) {
        statusIconElement.className = 'summary-status-icon';
    }

    const tellerDominanceRatio = transactionCount > 0 ? (topChannelByCount.value / transactionCount) : 0;
    const isTellerDominant = /teller|teler/i.test(topChannelByCount.label) && tellerDominanceRatio >= 0.5;
    const isNonTellerAboveTeller = nonTellerCount > tellerCount;

    let actionText = 'Aksi: Data belum cukup untuk menyusun rekomendasi.';
    if (isTellerDominant) {
        actionElement.classList.add('warning');
        if (statusIconElement) {
            statusIconElement.classList.add('active', 'warning');
            statusIconElement.textContent = '⚠️ PERINGATAN';
        }
        actionText = `⚠️ Warning: Transaksi Non Digital mendominasi ${(tellerDominanceRatio * 100).toFixed(1)}%. Tingkatkan prioritas Transaksi Kanal Digital untuk diversifikasi transaksinya.`;
    } else if (isNonTellerAboveTeller) {
        actionElement.classList.add('success');
        if (statusIconElement) {
            statusIconElement.classList.add('active', 'success');
            statusIconElement.textContent = '✅ BAGUS';
        }
        actionText = `✅ Bagus: Transaksi Digital (${nonTellerCount}) sudah di atas Transaksi Non Digital (${tellerCount}). Lanjutkan peningkatan Kanal Pembayaran Digital untuk mendorong pertumbuhan yang lebih tinggi.`;
    } else if (topChannel.value > 0) {
        actionText = `Aksi: Prioritaskan optimalisasi kanal ${topChannel.label} dan dorong replikasi ke kanal lain.`;
    }

    updateMetricWithAnimation('summaryAction', 'summaryAction', actionText);

    const recommendations = [];
    if (tellerRatio >= 0.5) {
        recommendations.push(`Tetapkan target TP2DD untuk menurunkan porsi Transaksi Non Digital dari ${(tellerRatio * 100).toFixed(1)}% menjadi <50% dalam 1-2 triwulan melalui migrasi transaksi ke Kanal Digital.`);
    } else {
        recommendations.push(`Pertahankan momentum ETPD dengan menjaga porsi Kanal Digital ${Math.max(digitalRatio * 100, 100 - tellerRatio * 100).toFixed(1)}% dan tetapkan KPI kenaikan bulanan transaksi Non Tunai.`);
    }

    if (qrisRatio < 0.2) {
        recommendations.push(`Percepat adopsi Kanal Digital (saat ini ${(qrisRatio * 100).toFixed(1)}%) lewat perluasan merchant/OPD, sosialisasi wajib scan, dan insentif kanal Non Tunai.`);
    } else {
        recommendations.push(`Optimalkan performa Kanal Digital (saat ini ${(qrisRatio * 100).toFixed(1)}%) dengan monitoring harian dan perluasan use case pajak/retribusi prioritas.`);
    }

    if (topSourceRatio >= 0.6) {
        recommendations.push(`Kurangi konsentrasi penerimaan pada satu sumber (${(topSourceRatio * 100).toFixed(1)}%) dengan strategi pemerataan kanal per jenis pajak/retribusi.`);
    }

    if (recommendationsElement) {
        const finalRecommendations = recommendations.slice(0, 3);
        recommendationsElement.innerHTML = '';
        finalRecommendations.forEach(text => {
            const item = document.createElement('li');
            item.textContent = text;
            recommendationsElement.appendChild(item);
        });
    }

    const targetWarningElement = document.getElementById('summaryTargetWarning');
    const trendInsightElement = document.getElementById('summaryTrendInsight');
    const segmentationElement = document.getElementById('summarySegmentationInsight');
    const qrisElement = document.getElementById('summaryQrisInsight');
    const simulationElement = document.getElementById('summarySimulationInsight');
    const healthElement = document.getElementById('summaryHealthScore');
    const healthPillElement = document.getElementById('summaryHealthPill');

    if (targetWarningElement) updateMetricWithAnimation('summaryTargetWarning', 'summaryTargetWarning', strategicInsights?.targetWarning || '-');
    if (trendInsightElement) updateMetricWithAnimation('summaryTrendInsight', 'summaryTrendInsight', strategicInsights?.trendInsight || '-');
    if (segmentationElement) updateMetricWithAnimation('summarySegmentationInsight', 'summarySegmentationInsight', strategicInsights?.segmentationInsight || '-');
    if (qrisElement) updateMetricWithAnimation('summaryQrisInsight', 'summaryQrisInsight', strategicInsights?.qrisInsight || '-');
    if (simulationElement) updateMetricWithAnimation('summarySimulationInsight', 'summarySimulationInsight', strategicInsights?.simulationInsight || '-');
    if (healthElement) updateMetricWithAnimation('summaryHealthScore', 'summaryHealthScore', strategicInsights?.healthScoreText || '-');
    if (healthPillElement) {
        const status = strategicInsights?.healthStatus || 'red';
        const healthText = status === 'green' ? 'BAIK' : (status === 'yellow' ? 'CUKUP' : 'WASPADA');
        healthPillElement.className = `summary-health-pill ${status}`;
        healthPillElement.textContent = healthText;
    }

    if (isTellerDominant) return 'warning';
    if (isNonTellerAboveTeller) return 'success';
    return 'neutral';
}

function updateTransactionBreakdown(channelTransactionData) {
    const container = document.getElementById('transactionBreakdown');
    if (!container) return;

    const escapeHtml = (value) => String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const entries = Object.entries(channelTransactionData || {})
        .filter(([, count]) => count > 0)
        .sort((a, b) => b[1] - a[1]);

    if (entries.length === 0) {
        container.innerHTML = '<span class="breakdown-item">Tidak ada data kanal</span>';
        return;
    }

    container.innerHTML = entries
        .map(([channelName, count]) => {
            const rawText = `${channelName}: ${count}`;
            return `<span class="breakdown-item" title="${escapeHtml(rawText)}">${escapeHtml(rawText)}</span>`;
        })
        .join('');
}

function updateCreditBreakdown(channelData) {
    const container = document.getElementById('creditBreakdown');
    if (!container) return;

    const escapeHtml = (value) => String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const entries = Object.entries(channelData || {})
        .filter(([, amount]) => amount > 0)
        .sort((a, b) => b[1] - a[1]);

    if (entries.length === 0) {
        container.innerHTML = '<span class="breakdown-item">Tidak ada data pemasukan kanal</span>';
        return;
    }

    container.innerHTML = entries
        .map(([channelName, amount]) => {
            const rawText = `${channelName}: ${formatCurrency(amount)}`;
            return `<span class="breakdown-item" title="${escapeHtml(rawText)}">${escapeHtml(rawText)}</span>`;
        })
        .join('');
}

async function exportDashboardPDF() {
    if (!currentUser || currentUser.role === 'guest') {
        showStatus('❌ Role guest tidak memiliki akses export laporan.', 'error');
        return;
    }

    const dashboard = document.getElementById('dashboard');
    if (!dashboard || dashboard.style.display === 'none') {
        showStatus('⚠️ Muat data terlebih dahulu sebelum export PDF.', 'warning');
        return;
    }

    if (!window.html2canvas || !window.jspdf) {
        showStatus('❌ Library export PDF belum siap. Coba refresh halaman.', 'error');
        return;
    }

    const exportTarget = document.querySelector('.admin-main');
    if (!exportTarget) {
        showStatus('❌ Area dashboard tidak ditemukan untuk diexport.', 'error');
        return;
    }

    const previousMenu = activeDashboardMenu;
    try {
        showStatus('⏳ Sedang membuat PDF, mohon tunggu...', 'warning');
        window.scrollTo({ top: 0, behavior: 'auto' });
        document.body.classList.add('pdf-export-mode');

        await new Promise(resolve => setTimeout(resolve, 180));
        Object.values(chartInstances).forEach(chart => {
            if (chart && typeof chart.resize === 'function') {
                chart.resize();
            }
        });
        await new Promise(resolve => setTimeout(resolve, 120));

        const datasetHero = document.querySelector('.dataset-hero');
        const sorotanStats = document.querySelector('#menuSectionSorotan .stats-grid');
        const sorotanSummary = document.querySelector('#menuSectionSorotan .summary-card');
        const isRenderableElement = (element) => {
            if (!element) return false;
            const rect = element.getBoundingClientRect();
            const style = window.getComputedStyle(element);
            return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
        };
        const grafikCards = Array.from(document.querySelectorAll('#menuSectionGrafik .chart-container'))
            .filter(isRenderableElement);

        const chapters = [
            {
                blocks: [
                    isRenderableElement(datasetHero) ? { element: datasetHero } : null,
                    isRenderableElement(sorotanStats) ? { element: sorotanStats } : null,
                    isRenderableElement(sorotanSummary) ? { element: sorotanSummary } : null
                ].filter(Boolean)
            },
            {
                blocks: grafikCards.map((card) => ({
                    element: card
                }))
            }
        ].filter(chapter => chapter.blocks.length > 0);

        if (chapters.length === 0) {
            throw new Error('Tidak ada bagian laporan yang dapat diexport.');
        }

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 10;
        const contentWidth = pageWidth - (margin * 2);
        const footerReserve = 12;
        let cursorY = 20;

        const periodText = (document.getElementById('period')?.textContent || '-').trim();
        const accountText = (document.getElementById('accountNo')?.textContent || '-').trim();
        const generatedAt = new Date().toLocaleString('id-ID');
        const displayedUserName = (document.getElementById('userName')?.textContent || '').trim();
        const currentFullName = String(currentUser?.fullName || '').trim();
        const printedByName = currentFullName || displayedUserName || '-';

        const sanitizePdfText = (value) => {
            return String(value || '')
                .normalize('NFKD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/[^\x20-\x7E]/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
        };

        const resolveChartForExportBlock = (blockElement) => {
            if (!blockElement) return null;
            if (blockElement.querySelector('#balanceChart')) return { key: 'balance', chart: chartInstances.balance };
            if (blockElement.querySelector('#channelChart')) return { key: 'channel', chart: chartInstances.channel };
            if (blockElement.querySelector('#revenueTypeChart')) return { key: 'revenue', chart: chartInstances.revenue };
            if (blockElement.querySelector('#laporanTransaksiChart')) return { key: 'googleSheetReport', chart: chartInstances.googleSheetReport };
            return null;
        };

        const prepareChartForExportCapture = (chartMeta) => {
            const chart = chartMeta?.chart;
            const chartKey = chartMeta?.key || '';
            if (!chart) return () => {};

            const previousSuppressOverlay = Boolean(chart.$suppressExportOverlay);
            const previousTooltipEnabled = chart?.options?.plugins?.tooltip?.enabled;
            const previousPlugins = Array.isArray(chart?.config?.plugins) ? [...chart.config.plugins] : null;

            const shouldRestartBalanceArrow = chartKey === 'balance' && Number(chart.$balanceTrendArrowAnimationFrameId || 0) > 0;
            const shouldRestartGoogleArrow = chartKey === 'googleSheetReport' && Number(chart.$trendArrowAnimationFrameId || 0) > 0;

            if (chartKey === 'balance') {
                stopBalanceTrendArrowAnimation(chart);
            }

            if (chartKey === 'googleSheetReport') {
                stopGoogleSheetTrendArrowAnimation(chart);
            }

            chart.$suppressExportOverlay = true;
            if (chart?.options?.plugins?.tooltip) {
                chart.options.plugins.tooltip.enabled = false;
            }
            if (Array.isArray(chart?.config?.plugins) && (chartKey === 'balance' || chartKey === 'googleSheetReport')) {
                chart.config.plugins = chart.config.plugins.filter((plugin) => {
                    const pluginId = String(plugin?.id || '');
                    return pluginId !== 'balanceTrendArrow' && pluginId !== 'googleSheetTrendArrow';
                });
            }

            chart.$balanceTrendArrowInfo = null;
            chart.$trendArrowInfo = null;
            chart.$balanceTrendUserHoverActive = false;
            chart.$balanceTrendTooltipActiveIndex = -1;
            chart.$trendTooltipActiveIndex = -1;

            if (typeof chart.setActiveElements === 'function') {
                chart.setActiveElements([]);
            }
            if (chart.tooltip && typeof chart.tooltip.setActiveElements === 'function') {
                chart.tooltip.setActiveElements([], { x: 0, y: 0 });
            }
            chart.update('none');

            return () => {
                chart.$balanceTrendArrowInfo = null;
                chart.$trendArrowInfo = null;
                chart.$suppressExportOverlay = previousSuppressOverlay;
                if (chart?.options?.plugins?.tooltip && typeof previousTooltipEnabled !== 'undefined') {
                    chart.options.plugins.tooltip.enabled = previousTooltipEnabled;
                }
                if (previousPlugins && Array.isArray(chart?.config?.plugins)) {
                    chart.config.plugins = previousPlugins;
                }
                if (typeof chart.setActiveElements === 'function') {
                    chart.setActiveElements([]);
                }
                if (chart.tooltip && typeof chart.tooltip.setActiveElements === 'function') {
                    chart.tooltip.setActiveElements([], { x: 0, y: 0 });
                }

                if (shouldRestartBalanceArrow && balanceViewMode === 'monthly') {
                    startBalanceTrendArrowAnimation(chart);
                }
                if (shouldRestartGoogleArrow) {
                    startGoogleSheetTrendArrowAnimation(chart);
                }
                chart.update('none');
            };
        };

        const appendChartDetailForExport = (blockElement) => {
            if (!blockElement) return null;

            let chart = null;
            let detailTitle = '';
            let totalTitle = '';
            let buildRows = null;

            if (blockElement.querySelector('#channelChart')) {
                chart = chartInstances.channel;
                detailTitle = 'Detail Transaksi Berdasarkan Kanal';
                totalTitle = 'Total Kanal';
                buildRows = () => {
                    const labels = chart?.data?.labels || [];
                    const values = chart?.data?.datasets?.[0]?.data || [];
                    if (!Array.isArray(labels) || !Array.isArray(values) || labels.length === 0 || values.length === 0) {
                        return [];
                    }
                    return labels.map((label, index) => ({
                        label: String(label || '-'),
                        value: Number(values[index] || 0)
                    })).filter(item => Number.isFinite(item.value) && item.value > 0);
                };
            } else if (blockElement.querySelector('#revenueTypeChart')) {
                chart = chartInstances.revenue;
                detailTitle = 'Detail Jenis Penerimaan';
                totalTitle = 'Total Jenis Penerimaan';
                buildRows = () => {
                    const labels = chart?.data?.labels || [];
                    const values = chart?.data?.datasets?.[0]?.data || [];
                    if (!Array.isArray(labels) || !Array.isArray(values) || labels.length === 0 || values.length === 0) {
                        return [];
                    }
                    return labels.map((label, index) => ({
                        label: String(label || '-'),
                        value: Number(values[index] || 0)
                    })).filter(item => Number.isFinite(item.value) && item.value > 0);
                };
            } else if (blockElement.querySelector('#balanceChart')) {
                chart = chartInstances.balance;
                detailTitle = 'Detail Pendapatan Bulanan';
                totalTitle = 'Total Pendapatan Bulanan';
                buildRows = () => {
                    const labels = chart?.data?.labels || [];
                    const barDataset = (chart?.data?.datasets || []).find(dataset => dataset?.label === 'Pendapatan Transaksi Bulanan' && dataset?.type === 'bar');
                    const lineDataset = (chart?.data?.datasets || []).find(dataset => dataset?.label === 'Tren Pendapatan Bulanan' && dataset?.type === 'line');
                    const barValues = Array.isArray(barDataset?.data) ? barDataset.data : [];
                    const lineValues = Array.isArray(lineDataset?.data) ? lineDataset.data : [];
                    if (!Array.isArray(labels) || labels.length === 0 || barValues.length === 0) {
                        return [];
                    }

                    return labels.map((label, index) => ({
                        label: String(label || '-'),
                        value: Number(barValues[index] || 0),
                        trendValue: Number(lineValues[index] || 0)
                    })).filter(item => Number.isFinite(item.value) && item.value > 0);
                };
            } else {
                return null;
            }

            if (!chart || typeof buildRows !== 'function') {
                return null;
            }

            const rows = buildRows();

            if (rows.length === 0) return null;

            const total = rows.reduce((sum, item) => sum + item.value, 0);
            if (!Number.isFinite(total) || total <= 0) return null;

            const wrapper = document.createElement('div');
            wrapper.setAttribute('data-export-channel-detail', 'true');
            wrapper.style.marginTop = '10px';
            wrapper.style.padding = '10px 12px';
            wrapper.style.border = '1px dashed #cbd5e1';
            wrapper.style.borderRadius = '10px';
            wrapper.style.background = '#f8fafc';
            wrapper.style.color = '#1f2937';
            wrapper.style.fontSize = '12px';
            wrapper.style.lineHeight = '1.45';

            const title = document.createElement('div');
            title.textContent = detailTitle;
            title.style.fontWeight = '700';
            title.style.marginBottom = '6px';
            wrapper.appendChild(title);

            const list = document.createElement('ul');
            list.style.margin = '0';
            list.style.paddingLeft = '18px';

            rows.forEach(item => {
                const percent = ((item.value / total) * 100).toFixed(1);
                const listItem = document.createElement('li');
                if (Number.isFinite(item.trendValue)) {
                    listItem.textContent = `${item.label}: ${percent}% (${formatCurrency(item.value)}) • Tren: ${formatCurrency(item.trendValue)}`;
                } else {
                    listItem.textContent = `${item.label}: ${percent}% (${formatCurrency(item.value)})`;
                }
                list.appendChild(listItem);
            });

            wrapper.appendChild(list);

            const totalText = document.createElement('div');
            totalText.style.marginTop = '6px';
            totalText.style.fontWeight = '700';
            totalText.textContent = `${totalTitle}: ${formatCurrency(total)}`;
            wrapper.appendChild(totalText);

            blockElement.appendChild(wrapper);
            return wrapper;
        };

        for (let chapterIndex = 0; chapterIndex < chapters.length; chapterIndex++) {
            const chapter = chapters[chapterIndex];
            if (chapterIndex > 0) {
                pdf.addPage();
            }

            cursorY = 20;

            for (let index = 0; index < chapter.blocks.length; index++) {
                const block = chapter.blocks[index];
                if (!isRenderableElement(block.element)) {
                    continue;
                }

                const chartMeta = resolveChartForExportBlock(block.element);
                const restoreChartStateAfterCapture = prepareChartForExportCapture(chartMeta);
                const temporaryDetailNode = appendChartDetailForExport(block.element);
                await new Promise(resolve => requestAnimationFrame(() => resolve()));

                const canvas = await window.html2canvas(block.element, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: '#ffffff',
                    scrollX: 0,
                    scrollY: 0,
                    windowWidth: Math.max(block.element.scrollWidth, block.element.clientWidth)
                });

                if (temporaryDetailNode && temporaryDetailNode.parentNode) {
                    temporaryDetailNode.parentNode.removeChild(temporaryDetailNode);
                }
                restoreChartStateAfterCapture();

                if (!canvas || !Number.isFinite(canvas.width) || !Number.isFinite(canvas.height) || canvas.width <= 0 || canvas.height <= 0) {
                    continue;
                }

                const imageData = canvas.toDataURL('image/png', 1.0);
                const imageHeight = (canvas.height * contentWidth) / canvas.width;
                if (!Number.isFinite(imageHeight) || imageHeight <= 0) {
                    continue;
                }

                if (cursorY + 4 > pageHeight - footerReserve) {
                    pdf.addPage();
                    cursorY = margin;
                }

                const maxDrawableHeight = pageHeight - footerReserve - margin;
                let scaleFactor = 1;
                if (imageHeight > maxDrawableHeight) {
                    scaleFactor = maxDrawableHeight / imageHeight;
                }

                const drawWidth = contentWidth * scaleFactor;
                const drawHeight = imageHeight * scaleFactor;
                const drawX = margin + ((contentWidth - drawWidth) / 2);

                if (!Number.isFinite(drawX) || !Number.isFinite(cursorY) || !Number.isFinite(drawWidth) || !Number.isFinite(drawHeight) || drawWidth <= 0 || drawHeight <= 0) {
                    continue;
                }

                const availableHeight = pageHeight - footerReserve - cursorY;
                if (drawHeight > availableHeight) {
                    pdf.addPage();
                    cursorY = margin;
                }

                pdf.addImage(imageData, 'PNG', drawX, cursorY, drawWidth, drawHeight);
                cursorY += drawHeight + 5;

                if (index !== chapter.blocks.length - 1 && cursorY > pageHeight - 36) {
                    pdf.addPage();
                    cursorY = margin;
                }
            }
        }

        const loadImageAsDataUrl = (src) => {
            return new Promise((resolve) => {
                const image = new Image();
                image.onload = () => {
                    try {
                        const canvas = document.createElement('canvas');
                        canvas.width = image.naturalWidth || image.width;
                        canvas.height = image.naturalHeight || image.height;
                        const context = canvas.getContext('2d');
                        if (!context) {
                            resolve(null);
                            return;
                        }
                        context.drawImage(image, 0, 0);
                        resolve(canvas.toDataURL('image/png'));
                    } catch (error) {
                        resolve(null);
                    }
                };
                image.onerror = () => resolve(null);
                image.src = src;
            });
        };

        const logoDataUrl = await loadImageAsDataUrl('data/pipakatan.png');
        const logoSize = 4.6;
        const logoY = 3.4;
        const headerTextX = margin + (logoDataUrl ? (logoSize + 1.6) : 0);

        const totalPages = pdf.getNumberOfPages();
        for (let page = 1; page <= totalPages; page++) {
            pdf.setPage(page);
            pdf.setDrawColor(226, 232, 240);
            pdf.setLineWidth(0.25);
            pdf.line(margin, 8, pageWidth - margin, 8);

            if (logoDataUrl) {
                pdf.addImage(logoDataUrl, 'PNG', margin, logoY, logoSize, logoSize);
            }

            pdf.setFont('helvetica', 'bold');
            pdf.setFontSize(8.5);
            pdf.text(sanitizePdfText('DINUNG PIPAKATAN'), headerTextX, 6.5);

            pdf.setFont('helvetica', 'normal');
            pdf.setFontSize(8.5);
            pdf.text(sanitizePdfText(`Dicetak: ${generatedAt}`), margin, pageHeight - 6);
            pdf.text(sanitizePdfText(`Halaman ${page} dari ${totalPages}`), pageWidth - margin, pageHeight - 6, { align: 'right' });
            pdf.setFontSize(7.8);
            pdf.text(sanitizePdfText(`Dicetak oleh: ${printedByName}`), margin, pageHeight - 2.2);
        }

        const safePeriod = periodText.replace(/[^a-zA-Z0-9-_]/g, '_') || 'periode';
        const dateStr = new Date().toISOString().slice(0, 10);
        pdf.save(`laporan-gabungan-pipakatan-${safePeriod}-${dateStr}.pdf`);

        showStatus('✅ Export PDF berhasil!', 'success');
    } catch (error) {
        console.error('Error exporting PDF:', error);
        showStatus('❌ Gagal export PDF: ' + error.message, 'error');
    } finally {
        document.body.classList.remove('pdf-export-mode');
        switchDashboardMenu(previousMenu);
    }
}

function cleanHeaderKey(key) {
    return String(key || '').replace(/^\uFEFF/, '').trim();
}

function normalizeHeaderLookupKey(key) {
    return cleanHeaderKey(key).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function pickFirstNonEmptyValue(values, fallback = '') {
    for (const value of values || []) {
        const text = String(value || '').trim();
        if (text) return text;
    }
    return fallback;
}

function resolveChannelName(row) {
    const rawChannel = pickFirstNonEmptyValue([
        row?.['Jenis Kanal'],
        row?.Kanal,
        row?.Channel,
        row?.['KETERANGAN'],
        row?.Keterangan,
        row?.Source
    ], '');

    if (!rawChannel) return 'Tidak Diklasifikasikan';

    if (/qris/i.test(rawChannel)) return 'QRIS';
    if (/teller|teler/i.test(rawChannel)) return 'Teller';
    if (/mobile|m-banking|mbanking|livin|brimo|bca\s*mobile/i.test(rawChannel)) return 'Mobile Banking';
    if (/internet|i-banking|ibanking/i.test(rawChannel)) return 'Internet Banking';
    if (/virtual\s*account|\bva\b/i.test(rawChannel)) return 'Virtual Account';
    if (/atm/i.test(rawChannel)) return 'ATM';
    if (/edc/i.test(rawChannel)) return 'EDC';
    if (/transfer/i.test(rawChannel)) return 'Transfer';

    return rawChannel;
}

function resolveSourceName(row) {
    return pickFirstNonEmptyValue([
        row?.Source,
        row?.['KETERANGAN'],
        row?.Keterangan,
        row?.['Jenis Penerimaan'],
        row?.__sourceFile
    ], 'Tanpa Sumber');
}

const ALLOWED_DISPLAY_CHANNELS = new Set(['teller', 'atm', 'mobile banking', 'qris', 'edc']);

function isSupportedDisplayChannel(channelName) {
    return ALLOWED_DISPLAY_CHANNELS.has(String(channelName || '').trim().toLowerCase());
}

function isAllowedDisplayChannel(row) {
    const resolvedChannel = resolveChannelName(row);
    return isSupportedDisplayChannel(resolvedChannel);
}

function normalizeCsvRow(row) {
    const normalized = {};

    const setNormalizedValue = (rawKey, rawValue) => {
        const normalizedKey = cleanHeaderKey(rawKey);
        const nextValue = typeof rawValue === 'string' ? rawValue.trim() : rawValue;
        const nextText = String(nextValue || '').trim();

        if (!Object.prototype.hasOwnProperty.call(normalized, normalizedKey)) {
            normalized[normalizedKey] = nextValue;
            return;
        }

        const currentValue = normalized[normalizedKey];
        const currentText = String(currentValue || '').trim();
        if (!currentText && nextText) {
            normalized[normalizedKey] = nextValue;
        }
    };

    Object.entries(row || {}).forEach(([key, value]) => {
        setNormalizedValue(key, value);
    });

    const hasDirectTaxSchema = [
        'NOP',
        'Total',
        'Tanggal',
        'Jenis Penerimaan',
        'Jenis Kanal'
    ].every((key) => Object.prototype.hasOwnProperty.call(normalized, key));

    const hasBankMutationSchema = [
        'NO',
        'TANGGAL - WAKTU',
        'NOMOR ARSIP',
        'KETERANGAN',
        'DEBET',
        'KREDIT',
        'SALDO AKHIR'
    ].every((key) => Object.prototype.hasOwnProperty.call(normalized, key));

    const findByAliases = (aliases, fallback = '') => {
        for (const alias of aliases) {
            const targetKey = normalizeHeaderLookupKey(alias);
            const foundKey = Object.keys(normalized).find(key => normalizeHeaderLookupKey(key) === targetKey);
            if (foundKey && normalized[foundKey] !== undefined && normalized[foundKey] !== '') {
                return normalized[foundKey];
            }
        }
        return fallback;
    };

    const combinedDateTime = findByAliases([
        'Tanggal Waktu',
        'TANGGAL - WAKTU',
        'TANGGAL- WAKTU',
        'TANGGAL-WAKTU',
        'Tanggal-Waktu',
        'Datetime',
        'DateTime'
    ]);
    if (combinedDateTime) {
        const [datePart, timePart] = String(combinedDateTime).split(/\s+/);
        if (!normalized.Tanggal && datePart) normalized.Tanggal = datePart;
        if (!normalized.Jam && timePart) normalized.Jam = timePart;
        if (!normalized.PostDate) normalized.PostDate = String(combinedDateTime);
    }

    if (!normalized.AccountNo) normalized.AccountNo = findByAliases(['AccountNo', 'Account No', 'NoRekening', 'Rekening', 'Nomor Arsip', 'NOMOR ARSIP', 'Kode Billing', 'NOP', 'No', 'NO', 'Nomor'], '');
    if (!normalized.PostDate) normalized.PostDate = findByAliases(['PostDate', 'PostingDate', 'TransactionDate', 'Tanggal Waktu', 'TANGGAL - WAKTU', 'DateTime'], '');
    if (!normalized.Tanggal) normalized.Tanggal = findByAliases(['Tanggal', 'Date', 'Tgl'], '');
    if (!normalized.Jam) normalized.Jam = findByAliases(['Jam', 'Time', 'Waktu'], '');
    if (!normalized['Credit Amount']) normalized['Credit Amount'] = findByAliases(['Credit Amount', 'Credit', 'Jumlah Kredit', 'Kredit', 'KREDIT', 'Total'], '0');
    if (!normalized['Debit Amount']) normalized['Debit Amount'] = findByAliases(['Debit Amount', 'Debit', 'Jumlah Debit', 'Debet', 'DEBET'], '0');
    const explicitCloseBalanceValue = findByAliases(['Close Balance', 'Balance', 'Saldo', 'Saldo Akhir', 'SALDO AKHIR'], '');
    if (!normalized['Close Balance']) normalized['Close Balance'] = explicitCloseBalanceValue || '0';
    if (!normalized.Source) normalized.Source = findByAliases(['Source', 'Sumber', 'Keterangan', 'KETERANGAN'], '');
    if (!normalized.Bulan) normalized.Bulan = findByAliases(['Bulan', 'Month', 'Tahun', 'Tanggal_Bulan', 'Tanggal Bulan', 'Periode'], '');
    const explicitChannelValue = findByAliases(['Jenis Kanal', 'JenisKanal', 'Kanal', 'Channel', 'KETERANGAN', 'Keterangan'], '');
    if (!normalized['Jenis Kanal']) normalized['Jenis Kanal'] = explicitChannelValue;
    if (!normalized['Jenis Penerimaan']) normalized['Jenis Penerimaan'] = findByAliases(['Jenis Penerimaan', 'JenisPenerimaan', 'Penerimaan', 'RevenueType', 'Jenis Pajak'], '');
    if (!normalized.Nama) normalized.Nama = findByAliases(['Nama', 'Nama Wajib Pajak', 'Wajib Pajak', 'Customer Name'], '');

    const hasBillingCsvSignature = Boolean(String(findByAliases(['Kode Billing', 'No', 'Nomor'], '') || '').trim());
    const hasBillingTotal = Boolean(String(findByAliases(['Total', 'Credit Amount', 'Jumlah Kredit', 'Kredit'], '') || '').trim());
    const hasBillingDate = Boolean(String(findByAliases(['Tanggal', 'Date', 'Tgl'], '') || '').trim());
    const hasBillingRevenue = Boolean(String(findByAliases(['Jenis Pajak', 'Jenis Penerimaan', 'Penerimaan'], '') || '').trim());
    const isBillingFormat = hasBillingCsvSignature && hasBillingTotal && hasBillingDate && hasBillingRevenue;

    if (!String(normalized['Jenis Penerimaan'] || '').trim() && isBillingFormat) {
        normalized['Jenis Penerimaan'] = findByAliases(['Jenis Pajak'], 'Lainnya');
    }

    if (!String(normalized['Jenis Kanal'] || '').trim() && isBillingFormat) {
        normalized['Jenis Kanal'] = 'Teller';
    }

    if (hasDirectTaxSchema) {
        normalized.AccountNo = String(normalized.NOP || normalized.AccountNo || '').trim();
        normalized['Credit Amount'] = String(normalized.Total || normalized['Credit Amount'] || '0').trim();
        normalized['Jenis Penerimaan'] = String(normalized['Jenis Penerimaan'] || normalized['Jenis Pajak'] || 'Lainnya').trim();
        normalized['Jenis Kanal'] = String(normalized['Jenis Kanal'] || 'Teller').trim();
        if (!String(normalized.PostDate || '').trim()) {
            normalized.PostDate = String(normalized.Tanggal || '').trim();
        }
    }

    normalized.__hasExplicitChannel = Boolean(String(explicitChannelValue || '').trim());
    normalized.__hasJenisPenerimaan = Boolean(String(normalized['Jenis Penerimaan'] || '').trim());
    normalized.__hasExplicitCloseBalance = Boolean(String(explicitCloseBalanceValue || '').trim());
    normalized.__isBillingFormat = isBillingFormat || hasDirectTaxSchema;

    if (hasBankMutationSchema) {
        if (!String(normalized.PostDate || '').trim()) {
            normalized.PostDate = String(findByAliases(['TANGGAL - WAKTU', 'Tanggal Waktu', 'DateTime'], '') || '').trim();
        }
        if (!String(normalized['Credit Amount'] || '').trim()) {
            normalized['Credit Amount'] = String(findByAliases(['KREDIT', 'Kredit', 'Credit Amount'], '0') || '0').trim();
        }
        if (!String(normalized['Debit Amount'] || '').trim()) {
            normalized['Debit Amount'] = String(findByAliases(['DEBET', 'Debet', 'Debit Amount'], '0') || '0').trim();
        }
        if (!String(normalized['Close Balance'] || '').trim()) {
            normalized['Close Balance'] = String(findByAliases(['SALDO AKHIR', 'Saldo Akhir', 'Close Balance'], '0') || '0').trim();
        }
    }

    const pokokAmount = parseAmount(findByAliases(['Pokok'], '0'));
    const dendaAmount = parseAmount(findByAliases(['Denda'], '0'));
    if (parseAmount(normalized['Credit Amount']) === 0 && (pokokAmount > 0 || dendaAmount > 0)) {
        normalized['Credit Amount'] = `${(pokokAmount + dendaAmount).toFixed(2).replace('.', ',')}`;
    }

    return normalized;
}

function isValidTransactionRow(row) {
    const hasDate = Boolean(String(row.PostDate || row.Tanggal || '').trim());
    const hasAccount = Boolean(String(row.AccountNo || row['NOMOR ARSIP'] || row.NOP || '').trim());
    const hasDescription = Boolean(String(row.Remarks || row['KETERANGAN'] || row.Keterangan || row.Source || '').trim());
    const hasAmountMovement = parseAmount(row['Credit Amount']) > 0 || parseAmount(row['Debit Amount']) > 0;

    const looksLikeBankMutationSchema = [
        'NO',
        'TANGGAL - WAKTU',
        'NOMOR ARSIP',
        'KETERANGAN',
        'DEBET',
        'KREDIT'
    ].some((key) => Object.prototype.hasOwnProperty.call(row || {}, key));

    if (!hasDate || !hasAmountMovement) {
        return false;
    }

    if (looksLikeBankMutationSchema) {
        return hasAccount && hasDescription;
    }

    return hasAccount || hasDescription;
}

function getRowTimestamp(row) {
    return parseRowTimestamp(row) ?? 0;
}

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
    statusDiv.style.display = 'block';

    if (type === 'warning' || type === 'error') {
        logActivityEvent('Status Dashboard', String(message || ''), {
            level: type === 'error' ? 'error' : 'warning',
            category: 'status',
            source: 'dashboard'
        });
    }
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 5000);
}

// Destroy existing charts
function destroyCharts() {
    Object.values(chartInstances).forEach(chart => {
        if (chart) chart.destroy();
    });
    chartInstances = {};
}

function getDistributionColorSet(count, baseColors) {
    const palette = Array.isArray(baseColors) && baseColors.length > 0
        ? baseColors
        : ['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b', '#fa709a', '#fee140', '#30cfd0', '#a8edea', '#fed6e3', '#c471f5', '#fa8bff'];
    return Array.from({ length: Math.max(0, count) }, (_, index) => palette[index % palette.length]);
}

function convertDistributionMapToSortedEntries(dataMap) {
    return Object.entries(dataMap || {})
        .filter(([, value]) => Number(value) > 0)
        .sort((a, b) => Number(b[1]) - Number(a[1]));
}

function getCurrentMonthKey() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function aggregateDistributionMaps(monthEntries, mapKey) {
    return (monthEntries || []).reduce((aggregate, monthEntry) => {
        const source = monthEntry?.[mapKey] || {};
        Object.entries(source).forEach(([label, value]) => {
            const numericValue = Number(value || 0);
            if (!Number.isFinite(numericValue) || numericValue <= 0) return;
            aggregate[label] = (aggregate[label] || 0) + numericValue;
        });
        return aggregate;
    }, {});
}

function buildFullMonthDistributionEntries(monthlyDistribution) {
    const sourceEntries = Array.isArray(monthlyDistribution) ? monthlyDistribution : [];
    if (sourceEntries.length === 0) return [];

    const keyedEntries = new Map();
    let minYear = Number.POSITIVE_INFINITY;
    let maxYear = Number.NEGATIVE_INFINITY;
    const currentMonthKey = getCurrentMonthKey();

    sourceEntries.forEach((entry) => {
        const key = String(entry?.key || '').trim();
        const match = key.match(/^(\d{4})-(\d{2})$/);
        if (!match) return;
        if (key > currentMonthKey) return;

        const year = Number.parseInt(match[1], 10);
        const month = Number.parseInt(match[2], 10);
        if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) return;

        minYear = Math.min(minYear, year);
        maxYear = Math.max(maxYear, year);
        keyedEntries.set(key, {
            key,
            label: String(entry?.label || '').trim() || new Date(year, month - 1, 1).toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }),
            channelData: { ...(entry?.channelData || {}) },
            revenueData: { ...(entry?.revenueData || {}) }
        });
    });

    if (!Number.isFinite(minYear) || !Number.isFinite(maxYear)) {
        const currentYear = Number.parseInt(currentMonthKey.slice(0, 4), 10);
        const currentMonth = Number.parseInt(currentMonthKey.slice(5, 7), 10);
        if (!Number.isInteger(currentYear) || !Number.isInteger(currentMonth)) {
            return [];
        }

        return Array.from({ length: currentMonth }, (_, index) => {
            const month = index + 1;
            const key = `${currentYear}-${String(month).padStart(2, '0')}`;
            return {
                key,
                label: new Date(currentYear, month - 1, 1).toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }),
                channelData: {},
                revenueData: {}
            };
        });
    }

    const completedEntries = [];
    for (let year = minYear; year <= maxYear; year += 1) {
        for (let month = 1; month <= 12; month += 1) {
            const key = `${year}-${String(month).padStart(2, '0')}`;
            if (key > currentMonthKey) continue;
            const existing = keyedEntries.get(key);
            if (existing) {
                completedEntries.push(existing);
            } else {
                completedEntries.push({
                    key,
                    label: new Date(year, month - 1, 1).toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }),
                    channelData: {},
                    revenueData: {}
                });
            }
        }
    }

    return completedEntries;
}

function calculateNextMonthlyFilterIndex(currentIndex, step, totalMonths) {
    const cycleLength = Math.max(1, Number(totalMonths) + 1);
    const normalizedStep = step < 0 ? -1 : 1;
    const currentCycleIndex = Number(currentIndex) + 1;
    const nextCycleIndex = (currentCycleIndex + normalizedStep + cycleLength) % cycleLength;
    return nextCycleIndex - 1;
}

function getMonthlyFilterLabel(months, selectedIndex) {
    if (!Array.isArray(months) || months.length === 0 || selectedIndex < 0) {
        return 'Semua Bulan';
    }
    return String(months[selectedIndex]?.label || 'Semua Bulan');
}

function updateDistributionMonthlyFilterControl(type) {
    const isChannel = type === 'channel';
    const prevButton = document.getElementById(isChannel ? 'channelMonthPrev' : 'revenueMonthPrev');
    const nextButton = document.getElementById(isChannel ? 'channelMonthNext' : 'revenueMonthNext');
    const labelElement = document.getElementById(isChannel ? 'channelMonthLabel' : 'revenueMonthLabel');
    const months = monthlyDistributionState.months || [];
    const selectedIndex = isChannel ? monthlyDistributionState.channelIndex : monthlyDistributionState.revenueIndex;
    const hasMonthlyData = months.length > 0;

    if (labelElement) {
        labelElement.textContent = getMonthlyFilterLabel(months, selectedIndex);
    }

    if (prevButton) prevButton.disabled = !hasMonthlyData;
    if (nextButton) nextButton.disabled = !hasMonthlyData;
}

function applyChannelChartMonthlyFilter() {
    const chart = chartInstances.channel;
    if (!chart) return;

    const months = monthlyDistributionState.months || [];
    const selectedMonthKey = monthlyDistributionState.channelIndex >= 0
        ? months[monthlyDistributionState.channelIndex]?.key
        : null;
    const sourceData = selectedMonthKey
        ? (monthlyDistributionState.channelByMonth.get(selectedMonthKey) || {})
        : (monthlyDistributionState.allChannelData || {});

    const entries = convertDistributionMapToSortedEntries(sourceData);
    const isMonthlySelection = Boolean(selectedMonthKey);
    const labels = entries.length > 0
        ? entries.map(([label]) => label)
        : (isMonthlySelection ? ['Teller', 'ATM', 'Mobile Banking', 'QRIS', 'EDC'] : []);
    const values = entries.length > 0
        ? entries.map(([, value]) => Number(value))
        : (isMonthlySelection ? [0, 0, 0, 0, 0] : []);
    const colors = getDistributionColorSet(labels.length, monthlyDistributionState.colors);

    chart.data.labels = labels;
    chart.data.datasets[0].data = values;
    chart.data.datasets[0].backgroundColor = colors;
    chart.data.datasets[0].borderColor = Array.from({ length: labels.length }, () => '#ffffff');
    chart.update();

    updateDistributionMonthlyFilterControl('channel');
}

function applyRevenueChartMonthlyFilter() {
    const chart = chartInstances.revenue;
    if (!chart) return;

    const months = monthlyDistributionState.months || [];
    const selectedMonthKey = monthlyDistributionState.revenueIndex >= 0
        ? months[monthlyDistributionState.revenueIndex]?.key
        : null;
    const sourceData = selectedMonthKey
        ? (monthlyDistributionState.revenueByMonth.get(selectedMonthKey) || {})
        : (monthlyDistributionState.allRevenueData || {});

    const entries = convertDistributionMapToSortedEntries(sourceData);
    const isMonthlySelection = Boolean(selectedMonthKey);
    const labels = entries.length > 0
        ? entries.map(([label]) => label)
        : (isMonthlySelection ? ['Tidak ada transaksi'] : []);
    const values = entries.length > 0
        ? entries.map(([, value]) => Number(value))
        : (isMonthlySelection ? [0] : []);
    const colors = getDistributionColorSet(labels.length, monthlyDistributionState.colors);

    chart.data.labels = labels;
    chart.data.datasets[0].data = values;
    chart.data.datasets[0].backgroundColor = colors;
    chart.data.datasets[0].borderColor = Array.from({ length: labels.length }, () => '#ffffff');
    chart.update();

    updateDistributionMonthlyFilterControl('revenue');
}

function changeChannelChartMonth(step) {
    const totalMonths = (monthlyDistributionState.months || []).length;
    if (totalMonths <= 0) return;
    monthlyDistributionState.channelIndex = calculateNextMonthlyFilterIndex(monthlyDistributionState.channelIndex, step, totalMonths);
    applyChannelChartMonthlyFilter();
}

function changeRevenueChartMonth(step) {
    const totalMonths = (monthlyDistributionState.months || []).length;
    if (totalMonths <= 0) return;
    monthlyDistributionState.revenueIndex = calculateNextMonthlyFilterIndex(monthlyDistributionState.revenueIndex, step, totalMonths);
    applyRevenueChartMonthlyFilter();
}

function updateBalanceToggleControls(activeMode, hasDaily, hasMonthly) {
    const dailyButton = document.getElementById('balanceToggleDaily');
    const monthlyButton = document.getElementById('balanceToggleMonthly');
    const modeBadge = document.getElementById('balanceModeBadge');
    if (!dailyButton || !monthlyButton) return;

    dailyButton.disabled = !hasDaily;
    monthlyButton.disabled = !hasMonthly;

    dailyButton.classList.toggle('active', activeMode === 'daily');
    monthlyButton.classList.toggle('active', activeMode === 'monthly');

    if (modeBadge) {
        modeBadge.textContent = `Mode: ${activeMode === 'monthly' ? 'Bulanan' : 'Harian'}`;
    }
}

function setBalanceViewMode(mode) {
    if (!['daily', 'monthly'].includes(mode)) return;
    balanceViewMode = mode;

    if (Array.isArray(latestProcessedRows) && latestProcessedRows.length > 0) {
        processData(latestProcessedRows);
    }
}

function detectCsvDelimiter(csvText) {
    const text = String(csvText || '');
    const lines = text
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean)
        .slice(0, 20);

    if (lines.length === 0) return ',';

    const splitCsvLine = (line, delimiter) => {
        const parts = [];
        let token = '';
        let inQuotes = false;

        for (let index = 0; index < line.length; index += 1) {
            const char = line[index];
            const nextChar = line[index + 1];

            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    token += '"';
                    index += 1;
                } else {
                    inQuotes = !inQuotes;
                }
                continue;
            }

            if (!inQuotes && char === delimiter) {
                parts.push(token.trim());
                token = '';
                continue;
            }

            token += char;
        }

        parts.push(token.trim());
        return parts;
    };

    const expectedHeaderKeys = new Set([
        'no',
        'tanggalwaktu',
        'nomorarsip',
        'keterangan',
        'debet',
        'kredit',
        'saldoakhir',
        'jenispenerimaan',
        'jeniskanal',
        'tanggal',
        'bulan'
    ]);

    const headerLine = lines[0] || '';
    const delimiterCandidates = [';', ',', '\t', '|'];
    const headerRank = delimiterCandidates
        .map((delimiter) => {
            const fields = splitCsvLine(headerLine, delimiter);
            const score = fields.reduce((sum, field) => {
                const key = normalizeHeaderLookupKey(field);
                return sum + (expectedHeaderKeys.has(key) ? 1 : 0);
            }, 0);
            return { delimiter, score };
        })
        .sort((a, b) => b.score - a.score);

    if (headerRank[0] && headerRank[0].score >= 3) {
        return headerRank[0].delimiter;
    }

    const countDelimiter = (line, delimiter) => {
        let inQuotes = false;
        let count = 0;
        for (let i = 0; i < line.length; i += 1) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (!inQuotes && char === delimiter) {
                count += 1;
            }
        }
        return count;
    };

    const delimiterScores = {
        ';': 0,
        ',': 0,
        '\t': 0,
        '|': 0
    };

    lines.forEach(line => {
        delimiterScores[';'] += countDelimiter(line, ';');
        delimiterScores[','] += countDelimiter(line, ',');
        delimiterScores['\t'] += countDelimiter(line, '\t');
        delimiterScores['|'] += countDelimiter(line, '|');
    });

    const ranked = Object.entries(delimiterScores).sort((a, b) => b[1] - a[1]);
    return ranked[0][1] > 0 ? ranked[0][0] : ',';
}

function getExpectedCsvHeaderKeys() {
    return new Set([
        'no',
        'tanggalwaktu',
        'nomorarsip',
        'keterangan',
        'debet',
        'kredit',
        'saldoakhir',
        'jenispenerimaan',
        'jeniskanal',
        'tanggal',
        'bulan'
    ]);
}

function normalizeCsvTextBeforeParsing(csvText) {
    const text = String(csvText || '').replace(/^\uFEFF/, '');
    const lines = text.split(/\r?\n/);
    const nonEmptyIndexes = lines
        .map((line, index) => ({ line: String(line || '').trim(), index }))
        .filter(item => item.line.length > 0)
        .map(item => item.index);

    if (nonEmptyIndexes.length < 2) return text;

    const headerIndex = nonEmptyIndexes[0];
    const sampleIndex = nonEmptyIndexes[1];
    const headerLine = lines[headerIndex] || '';
    const sampleLine = lines[sampleIndex] || '';

    const countOutsideQuotes = (line, char) => {
        let inQuotes = false;
        let count = 0;
        for (let i = 0; i < line.length; i += 1) {
            const current = line[i];
            if (current === '"') {
                inQuotes = !inQuotes;
            } else if (!inQuotes && current === char) {
                count += 1;
            }
        }
        return count;
    };

    const headerComma = countOutsideQuotes(headerLine, ',');
    const headerSemicolon = countOutsideQuotes(headerLine, ';');
    const sampleComma = countOutsideQuotes(sampleLine, ',');
    const sampleSemicolon = countOutsideQuotes(sampleLine, ';');

    const isMixedHeaderDelimiter = headerComma >= 5 && headerSemicolon === 0 && sampleSemicolon >= 5 && sampleSemicolon > sampleComma;
    if (!isMixedHeaderDelimiter) return text;

    lines[headerIndex] = headerLine.replace(/,/g, ';');
    return lines.join('\n');
}

function looksLikeBrokenCsvRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return false;
    const sampleRow = rows.find(item => item && typeof item === 'object') || rows[0];
    const keys = Object.keys(sampleRow || {});
    if (keys.length === 0) return false;
    if (keys.length === 1 && /;|\||\t/.test(keys[0])) return true;
    if (keys.some(key => /;|\||\t/.test(String(key)))) return true;
    return false;
}

function evaluateParsedRowsQuality(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return -1;

    let usableRows = 0;
    let datedRows = 0;
    let amountRows = 0;
    let channelRows = 0;

    rows.forEach((rawRow) => {
        const normalizedRow = normalizeCsvRow(rawRow || {});
        const hasDate = Boolean(String(normalizedRow.PostDate || normalizedRow.Tanggal || '').trim());
        const hasAmount = parseAmount(normalizedRow['Credit Amount']) > 0 || parseAmount(normalizedRow['Debit Amount']) > 0;
        const hasChannel = Boolean(String(resolveChannelName(normalizedRow) || '').trim());

        if (isValidTransactionRow(normalizedRow)) usableRows += 1;
        if (hasDate) datedRows += 1;
        if (hasAmount) amountRows += 1;
        if (hasChannel) channelRows += 1;
    });

    return (usableRows * 5) + (datedRows * 2) + (amountRows * 2) + channelRows;
}

function parseCsvTextWithFallback(csvText) {
    const normalizedCsvText = normalizeCsvTextBeforeParsing(csvText);
    const primaryDelimiter = detectCsvDelimiter(normalizedCsvText);
    const candidateDelimiters = Array.from(new Set([
        primaryDelimiter,
        primaryDelimiter === ';' ? ',' : ';',
        '\t',
        '|'
    ]));

    const parseWithDelimiter = (delimiter) => Papa.parse(normalizedCsvText, {
        header: true,
        skipEmptyLines: 'greedy',
        delimiter,
        quoteChar: '"',
        escapeChar: '"',
        transformHeader: function(header) {
            return cleanHeaderKey(header);
        }
    });

    let bestResult = null;
    let bestScore = -1;
    const expectedHeaderKeys = getExpectedCsvHeaderKeys();

    candidateDelimiters.forEach((delimiter) => {
        const result = parseWithDelimiter(delimiter);
        const rows = result.data || [];
        if (looksLikeBrokenCsvRows(rows)) return;

        const matchedHeaderCount = (result.meta?.fields || []).reduce((sum, field) => {
            const key = normalizeHeaderLookupKey(field);
            return sum + (expectedHeaderKeys.has(key) ? 1 : 0);
        }, 0);

        const score = evaluateParsedRowsQuality(rows) + (matchedHeaderCount * 20);
        if (score > bestScore) {
            bestScore = score;
            bestResult = result;
        }
    });

    if (bestResult) {
        return bestResult;
    }

    return parseWithDelimiter(primaryDelimiter);
}

function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ''));
        reader.onerror = () => reject(new Error('Gagal membaca file CSV.'));
        reader.readAsText(file);
    });
}

function updateGoogleSheetMeta(message, type = 'info') {
    const metaElement = document.getElementById('gsReportMeta');
    if (!metaElement) return;

    metaElement.textContent = String(message || '-');
    if (type === 'error') {
        metaElement.style.color = '#dc2626';
        return;
    }

    if (type === 'loading') {
        metaElement.style.color = '#b45309';
        return;
    }

    metaElement.style.color = '';
}

function resetGoogleSheetSummary() {
    const totalCreditElement = document.getElementById('gsTotalCredit');
    const totalTransactionsElement = document.getElementById('gsTotalTransactions');
    const transaksiQrisElement = document.getElementById('gsTransactionsQris');
    const transaksiMobileBankingElement = document.getElementById('gsTransactionsMobileBanking');
    const transaksiEDCElement = document.getElementById('gsTransactionsEDC');
    const transaksiATMElement = document.getElementById('gsTransactionsATM');
    const transaksiTellerElement = document.getElementById('gsTransactionsTeller');
    const totalCreditBreakdownElement = document.getElementById('gsTotalCreditBreakdown');

    if (chartInstances.googleSheetReport) {
        chartInstances.googleSheetReport.destroy();
        chartInstances.googleSheetReport = null;
    }

    if (totalCreditElement) {
        totalCreditElement.textContent = formatCurrency(0);
    }

    if (totalTransactionsElement) {
        totalTransactionsElement.textContent = '0';
    }

    if (transaksiQrisElement) transaksiQrisElement.textContent = '0';
    if (transaksiMobileBankingElement) transaksiMobileBankingElement.textContent = '0';
    if (transaksiEDCElement) transaksiEDCElement.textContent = '0';
    if (transaksiATMElement) transaksiATMElement.textContent = '0';
    if (transaksiTellerElement) transaksiTellerElement.textContent = '0';
    if (totalCreditBreakdownElement) {
        totalCreditBreakdownElement.innerHTML = '<div style="font-size:0.8em; color:#64748b; margin-top:8px;">Detail pemasukan manual belum tersedia.</div>';
    }
}

function stopGoogleSheetTrendArrowAnimation(chart) {
    if (!chart || !chart.$trendArrowAnimationFrameId) return;

    cancelAnimationFrame(chart.$trendArrowAnimationFrameId);
    chart.$trendArrowAnimationFrameId = 0;
    chart.$trendArrowInfo = null;
}

function getGoogleSheetTrendPathMetrics(chart) {
    if (!chart || !chart.data) return null;

    const lineDatasetIndex = chart.data.datasets.findIndex(dataset => dataset.type === 'line');
    if (lineDatasetIndex < 0) return null;

    const lineMeta = chart.getDatasetMeta(lineDatasetIndex);
    const points = (lineMeta?.data || []).filter(point => Number.isFinite(point?.x) && Number.isFinite(point?.y));
    if (points.length < 2) return null;

    const segmentLengths = [];
    const cumulativeDistances = [0];
    let totalLength = 0;

    for (let index = 0; index < points.length - 1; index += 1) {
        const from = points[index];
        const to = points[index + 1];
        const length = Math.hypot(to.x - from.x, to.y - from.y);
        segmentLengths.push(length);
        totalLength += length;
        cumulativeDistances.push(totalLength);
    }

    if (totalLength <= 0) return null;

    return {
        lineDatasetIndex,
        points,
        segmentLengths,
        cumulativeDistances,
        totalLength
    };
}

function getGoogleSheetTrendArrowState(chart) {
    const pathMetrics = getGoogleSheetTrendPathMetrics(chart);
    if (!pathMetrics) return null;

    const { lineDatasetIndex, points, segmentLengths, totalLength } = pathMetrics;

    const progress = Number(chart.$trendArrowProgress || 0);
    let targetDistance = progress * totalLength;
    let segmentIndex = 0;

    while (segmentIndex < segmentLengths.length - 1 && targetDistance > segmentLengths[segmentIndex]) {
        targetDistance -= segmentLengths[segmentIndex];
        segmentIndex += 1;
    }

    const fromPoint = points[segmentIndex];
    const toPoint = points[segmentIndex + 1];
    const currentSegmentLength = segmentLengths[segmentIndex] || 1;
    const ratio = Math.min(Math.max(targetDistance / currentSegmentLength, 0), 1);
    const arrowX = fromPoint.x + (toPoint.x - fromPoint.x) * ratio;
    const arrowY = fromPoint.y + (toPoint.y - fromPoint.y) * ratio;
    const angle = Math.atan2(toPoint.y - fromPoint.y, toPoint.x - fromPoint.x);

    const values = chart.data.datasets[lineDatasetIndex]?.data || [];
    const fromValue = Number(values[segmentIndex] || 0);
    const toValue = Number(values[segmentIndex + 1] || 0);
    let arrowColor = '#64748b';
    if (toValue > fromValue) arrowColor = '#16a34a';
    if (toValue < fromValue) arrowColor = '#dc2626';

    let nearestPointIndex = -1;
    let nearestPointDistance = Number.POSITIVE_INFINITY;
    points.forEach((point, index) => {
        const distance = Math.hypot(point.x - arrowX, point.y - arrowY);
        if (distance < nearestPointDistance) {
            nearestPointDistance = distance;
            nearestPointIndex = index;
        }
    });

    return {
        lineDatasetIndex,
        segmentIndex,
        arrowX,
        arrowY,
        angle,
        arrowColor,
        nearestPointIndex,
        nearestPointDistance
    };
}

function clearGoogleSheetTrendTooltip(chart) {
    if (!chart || !chart.tooltip) return;

    chart.$trendTooltipActiveIndex = -1;
    chart.setActiveElements([]);
    chart.tooltip.setActiveElements([], { x: 0, y: 0 });
}

function updateGoogleSheetTrendTooltip(chart, arrowState) {
    if (!chart || !arrowState) return;

    const pointIndex = arrowState.nearestPointIndex;
    if (pointIndex < 0) {
        chart.$trendArrowInfo = null;
        return;
    }

    const labels = chart.data?.labels || [];
    const barDatasetIndex = 0;
    const barDataset = chart.data?.datasets?.[barDatasetIndex] || {};
    const lineDataset = chart.data?.datasets?.[arrowState.lineDatasetIndex] || {};
    const barValues = barDataset.data || [];
    const lineValues = lineDataset.data || [];
    const currentValue = Number(lineValues[pointIndex] || 0);
    const prevValue = pointIndex > 0 ? Number(lineValues[pointIndex - 1] || 0) : currentValue;
    let trendLabel = 'Stabil';
    if (currentValue > prevValue) trendLabel = 'Naik';
    if (currentValue < prevValue) trendLabel = 'Turun';

    const totalValue = Number(barValues[pointIndex] || 0);
    const infoLines = [
        `${String(barDataset.label || 'Total Pemasukan')}: ${formatCurrency(totalValue)}`,
        `${String(lineDataset.label || 'Tren Naik/Turun')}: ${trendLabel}`
    ];

    chart.$trendArrowInfo = {
        monthLabel: String(labels[pointIndex] || '-'),
        lines: infoLines,
        x: Number(arrowState.arrowX || 0),
        y: Number(arrowState.arrowY || 0)
    };
}

function startGoogleSheetTrendArrowAnimation(chart) {
    if (!chart) return;

    stopGoogleSheetTrendArrowAnimation(chart);
    chart.$trendTooltipActiveIndex = -1;
    chart.$trendArrowProgress = 0;
    chart.$trendArrowPhase = 'move';
    chart.$trendArrowTargetPointIndex = 1;
    chart.$trendArrowPhaseStart = performance.now();

    const animateArrow = (timestamp) => {
        if (!chart || !chart.ctx || !chart.canvas) return;

        const pathMetrics = getGoogleSheetTrendPathMetrics(chart);
        if (!pathMetrics) {
            chart.$trendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
            return;
        }

        const travelDuration = 7600;
        const pointPauseDuration = 2400;
        const maxPointIndex = pathMetrics.points.length - 1;
        const currentTarget = Math.min(Math.max(Number(chart.$trendArrowTargetPointIndex || 1), 1), maxPointIndex);
        const phaseStart = Number(chart.$trendArrowPhaseStart || timestamp);
        const elapsedInPhase = Math.max(0, timestamp - phaseStart);

        if (chart.$trendArrowPhase === 'hold') {
            const holdPointIndex = Math.min(Math.max(Number(chart.$trendArrowHoldPointIndex || currentTarget), 1), maxPointIndex);
            chart.$trendArrowProgress = pathMetrics.cumulativeDistances[holdPointIndex] / pathMetrics.totalLength;

            if (elapsedInPhase >= pointPauseDuration) {
                if (holdPointIndex >= maxPointIndex) {
                    chart.$trendArrowTargetPointIndex = 1;
                    chart.$trendArrowProgress = 0;
                } else {
                    chart.$trendArrowTargetPointIndex = holdPointIndex + 1;
                }
                chart.$trendArrowPhase = 'move';
                chart.$trendArrowPhaseStart = timestamp;
            }
        } else {
            const fromPointIndex = Math.max(currentTarget - 1, 0);
            const segmentLength = Number(pathMetrics.segmentLengths[fromPointIndex] || 0);
            const segmentDuration = Math.max((segmentLength / pathMetrics.totalLength) * travelDuration, 1);
            const progressInSegment = Math.min(elapsedInPhase / segmentDuration, 1);
            const distanceBeforeSegment = Number(pathMetrics.cumulativeDistances[fromPointIndex] || 0);
            const traveledDistance = distanceBeforeSegment + (segmentLength * progressInSegment);
            chart.$trendArrowProgress = traveledDistance / pathMetrics.totalLength;

            if (progressInSegment >= 1) {
                chart.$trendArrowPhase = 'hold';
                chart.$trendArrowHoldPointIndex = currentTarget;
                chart.$trendArrowPhaseStart = timestamp;
            }
        }

        chart.$trendArrowRenderState = getGoogleSheetTrendArrowState(chart);
        updateGoogleSheetTrendTooltip(chart, chart.$trendArrowRenderState);
        chart.draw();
        chart.$trendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
    };

    chart.$trendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
}

const googleSheetTrendArrowPlugin = {
    id: 'googleSheetTrendArrow',
    afterDatasetsDraw(chart) {
        if (chart?.$suppressExportOverlay) return;

        const arrowState = chart.$trendArrowRenderState || getGoogleSheetTrendArrowState(chart);
        if (!arrowState) return;

        const context = chart.ctx;
        context.save();
        context.translate(arrowState.arrowX, arrowState.arrowY);
        context.rotate(arrowState.angle);
        context.beginPath();
        context.moveTo(9, 0);
        context.lineTo(-6, 5);
        context.lineTo(-6, -5);
        context.closePath();
        context.fillStyle = arrowState.arrowColor;
        context.fill();
        context.restore();

        const arrowInfo = chart.$trendArrowInfo;
        if (!arrowInfo) return;

        const titleText = String(arrowInfo.monthLabel || '-');
        const infoLines = Array.isArray(arrowInfo.lines) ? arrowInfo.lines : [];
        if (infoLines.length === 0) return;

        context.save();
        context.font = '12px Roboto';
        const titleWidth = context.measureText(titleText).width;
        const linesWidth = infoLines.reduce((width, lineText) => Math.max(width, context.measureText(String(lineText || '')).width), 0);
        const textWidth = Math.max(titleWidth, linesWidth);
        const paddingX = 8;
        const lineHeight = 15;
        const boxWidth = textWidth + (paddingX * 2);
        const boxHeight = 8 + lineHeight + 4 + (infoLines.length * lineHeight) + 8;
        const area = chart.chartArea || { left: 0, right: chart.width, top: 0, bottom: chart.height };
        let boxX = arrowInfo.x + 14;
        let boxY = arrowInfo.y - boxHeight - 10;

        if (boxX + boxWidth > area.right - 4) {
            boxX = arrowInfo.x - boxWidth - 14;
        }
        if (boxX < area.left + 4) {
            boxX = area.left + 4;
        }
        if (boxY < area.top + 4) {
            boxY = arrowInfo.y + 10;
        }

        context.fillStyle = 'rgba(15, 23, 42, 0.9)';
        context.beginPath();
        context.roundRect(boxX, boxY, boxWidth, boxHeight, 6);
        context.fill();

        context.fillStyle = '#ffffff';
        context.textBaseline = 'top';
        context.fillText(titleText, boxX + paddingX, boxY + 8);

        let currentTextY = boxY + 8 + lineHeight + 4;
        infoLines.forEach((lineText) => {
            context.fillText(String(lineText || ''), boxX + paddingX, currentTextY);
            currentTextY += lineHeight;
        });
        context.restore();
    },
    afterDestroy(chart) {
        stopGoogleSheetTrendArrowAnimation(chart);
    }
};

function renderGoogleSheetChart(rows) {
    const canvas = document.getElementById('laporanTransaksiChart');
    if (!canvas) return;

    if (chartInstances.googleSheetReport) {
        chartInstances.googleSheetReport.destroy();
        chartInstances.googleSheetReport = null;
    }

    const activeYear = new Date().getFullYear();
    const monthLabels = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const monthlyCredit = Array(12).fill(0);
    let totalCredit = 0;
    let totalTransactions = 0;
    const manualHistory = getManualInputHistory();
    const transactionByChannel = {
        qris: 0,
        mobileBanking: 0,
        edc: 0,
        atm: 0,
        teller: 0
    };

    rows.forEach((row) => {
        const timestamp = getRowTimestamp(row);
        if (!Number.isFinite(timestamp) || timestamp <= 0) return;

        const rowDate = new Date(timestamp);
        if (rowDate.getFullYear() !== activeYear) return;

        const credit = parseAmount(row['Credit Amount']);
        const pokok = parseAmount(row.Pokok);
        const denda = parseAmount(row.Denda);
        const effectiveCredit = credit > 0 ? credit : (pokok + denda);

        if (effectiveCredit > 0) {
            monthlyCredit[rowDate.getMonth()] += effectiveCredit;
            totalCredit += effectiveCredit;
        }

        const channelName = String(resolveChannelName(row) || '').trim().toLowerCase();
        if (channelName === 'qris') transactionByChannel.qris += 1;
        if (channelName === 'mobile banking') transactionByChannel.mobileBanking += 1;
        if (channelName === 'edc') transactionByChannel.edc += 1;
        if (channelName === 'atm') transactionByChannel.atm += 1;
        if (channelName === 'teller') transactionByChannel.teller += 1;

        totalTransactions += 1;
    });

    const manualChannelTotals = getManualChannelTotalsForYear(activeYear, manualHistory);
    const manualCreditTotals = getManualReportAmountTotalsForYear(activeYear, manualHistory);
    if (manualChannelTotals.hasData) {
        transactionByChannel.qris = manualChannelTotals.qris;
        transactionByChannel.mobileBanking = manualChannelTotals.mobileBanking;
        transactionByChannel.edc = manualChannelTotals.edc;
        transactionByChannel.atm = manualChannelTotals.atm;
        transactionByChannel.teller = manualChannelTotals.teller;
    }

    if (manualCreditTotals.hasData) {
        totalCredit = manualCreditTotals.total;
    }

    const totalCreditElement = document.getElementById('gsTotalCredit');
    const totalTransactionsElement = document.getElementById('gsTotalTransactions');
    const transaksiQrisElement = document.getElementById('gsTransactionsQris');
    const transaksiMobileBankingElement = document.getElementById('gsTransactionsMobileBanking');
    const transaksiEDCElement = document.getElementById('gsTransactionsEDC');
    const transaksiATMElement = document.getElementById('gsTransactionsATM');
    const transaksiTellerElement = document.getElementById('gsTransactionsTeller');
    if (totalCreditElement) {
        totalCreditElement.textContent = formatCurrency(totalCredit);
    }
    if (totalTransactionsElement) {
        totalTransactionsElement.textContent = totalTransactions.toLocaleString('id-ID');
    }
    if (transaksiQrisElement) transaksiQrisElement.textContent = transactionByChannel.qris.toLocaleString('id-ID');
    if (transaksiMobileBankingElement) transaksiMobileBankingElement.textContent = transactionByChannel.mobileBanking.toLocaleString('id-ID');
    if (transaksiEDCElement) transaksiEDCElement.textContent = transactionByChannel.edc.toLocaleString('id-ID');
    if (transaksiATMElement) transaksiATMElement.textContent = transactionByChannel.atm.toLocaleString('id-ID');
    if (transaksiTellerElement) transaksiTellerElement.textContent = transactionByChannel.teller.toLocaleString('id-ID');
    renderGoogleSheetCreditBreakdown(manualCreditTotals, activeYear);

    chartInstances.googleSheetReport = new Chart(canvas, {
        plugins: [googleSheetTrendArrowPlugin],
        type: 'bar',
        data: {
            labels: monthLabels,
            datasets: [
                {
                    label: `Total Pemasukan Bulanan (${activeYear})`,
                    data: monthlyCredit,
                    backgroundColor: 'rgba(14, 165, 233, 0.28)',
                    borderColor: '#0284c7',
                    borderWidth: 1,
                    borderRadius: 6
                },
                {
                    type: 'line',
                    label: 'Tren Naik/Turun',
                    data: monthlyCredit,
                    borderColor: '#16a34a',
                    borderWidth: 2,
                    tension: 0.28,
                    fill: false,
                    pointRadius: 3,
                    pointHoverRadius: 5,
                    pointBackgroundColor(context) {
                        const dataIndex = context.dataIndex;
                        const values = context.dataset.data || [];
                        const currentValue = Number(values[dataIndex] || 0);
                        const prevValue = dataIndex > 0 ? Number(values[dataIndex - 1] || 0) : currentValue;

                        if (currentValue > prevValue) return '#16a34a';
                        if (currentValue < prevValue) return '#dc2626';
                        return '#64748b';
                    },
                    segment: {
                        borderColor(segmentContext) {
                            const from = Number(segmentContext.p0?.parsed?.y || 0);
                            const to = Number(segmentContext.p1?.parsed?.y || 0);
                            if (to > from) return '#16a34a';
                            if (to < from) return '#dc2626';
                            return '#64748b';
                        }
                    }
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            animation: {
                duration: 900,
                easing: 'easeOutQuart',
                delay(context) {
                    if (context.type !== 'data' || context.mode !== 'default') return 0;
                    return context.dataIndex * 45;
                }
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            if (context.dataset.type === 'line') {
                                const values = context.dataset.data || [];
                                const index = context.dataIndex;
                                const currentValue = Number(values[index] || 0);
                                const prevValue = index > 0 ? Number(values[index - 1] || 0) : currentValue;
                                let trendLabel = 'Stabil';
                                if (currentValue > prevValue) trendLabel = 'Naik';
                                if (currentValue < prevValue) trendLabel = 'Turun';
                                return `${context.dataset.label}: ${trendLabel}`;
                            }
                            return `${context.dataset.label}: ${formatCurrency(Number(context.parsed.y || 0))}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return formatCurrencyTick(value);
                        }
                    }
                }
            }
        }
    });

    startGoogleSheetTrendArrowAnimation(chartInstances.googleSheetReport);
}

function replayGoogleSheetChartAnimation() {
    const reportChart = chartInstances.googleSheetReport;
    if (!reportChart) return;

    reportChart.reset();
    reportChart.update();
    startGoogleSheetTrendArrowAnimation(reportChart);
}

function loadGoogleSheetReport(forceReload = false) {
    const manualDataset = getManualInputDataset();
    if (manualDataset && applyManualDatasetToReport(manualDataset)) {
        return;
    }

    const rows = normalizeRowsForProcessing(latestProcessedRows || []);
    if (rows.length === 0) {
        resetGoogleSheetSummary();
        updateGoogleSheetMeta('Sumber laporan belum tersedia. Gunakan Input Data Manual atau muat CSV.', 'error');
        return;
    }

    googleSheetReportCache = {
        url: 'local-data://processed',
        rows,
        isLoading: false
    };

    renderGoogleSheetChart(rows);
    const latestSource = String(rows[rows.length - 1]?.__sourceFile || rows[rows.length - 1]?.Source || 'Data Lokal').trim();
    updateGoogleSheetMeta(`Sumber Laporan: ${latestSource}`);
}

function refreshGoogleSheetReport() {
    loadGoogleSheetReport(true);
    showStatus('⏳ Refresh data laporan diproses...', 'warning');
}

// Load CSV from file input
function loadCSV() {
    // Check permission
    if (!currentUser || !currentUser.permissions.includes('upload')) {
        showStatus('❌ Anda tidak memiliki izin untuk upload file!', 'error');
        return;
    }
    
    const fileInput = document.getElementById('csvFile');
    const files = Array.from(fileInput.files || []);
    
    if (files.length === 0) {
        showStatus('Silakan pilih minimal 1 file CSV terlebih dahulu!', 'error');
        return;
    }
    
    const invalidFile = files.find(file => !file.name.toLowerCase().endsWith('.csv'));
    if (invalidFile) {
        showStatus('Semua file harus berformat CSV! File tidak valid: ' + invalidFile.name, 'error');
        return;
    }
    
    document.getElementById('loading').style.display = 'block';
    document.getElementById('dashboard').style.display = 'none';

    const parsePromises = files.map(file => readFileAsText(file)
        .then(csvText => {
            const parseResults = parseCsvTextWithFallback(csvText);
            const normalizedRows = (parseResults.data || []).map(normalizeCsvRow);
            return { fileName: file.name, data: normalizedRows };
        })
        .catch(error => {
            throw new Error(file.name + ' - ' + error.message);
        })
    );

    Promise.all(parsePromises)
        .then(parsedResults => {
            const mergedData = parsedResults.flatMap(result =>
                result.data.map(row => ({
                    ...row,
                    __sourceFile: result.fileName
                }))
            );
            activeDashboardMenu = 'sorotan';
            processData(mergedData);

            const fileNames = files.map(file => file.name).join(', ');
            showStatus('✅ Data berhasil dimuat dari ' + files.length + ' file: ' + fileNames, 'success');
        })
        .catch(error => {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('dashboard').style.display = 'block';
            showStatus('❌ Error membaca file: ' + error.message, 'error');
        });
}

// Load default CSV from data folder
function loadDefaultCSV() {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('dashboard').style.display = 'none';
    
    fetch(`data/rekening-bank.csv?v=${Date.now()}`, { cache: 'no-store' })
        .then(response => {
            if (!response.ok) {
                throw new Error('File CSV tidak ditemukan');
            }
            return response.text();
        })
        .then(csvText => {
            const results = parseCsvTextWithFallback(csvText);
            const rowsWithSource = (results.data || []).map(row => ({
                ...normalizeCsvRow(row),
                __sourceFile: 'data/rekening-bank.csv'
            }));

            const validRows = rowsWithSource
                .filter(isValidTransactionRow)
                .filter(isAllowedDisplayChannel);

            if (validRows.length === 0) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('dashboard').style.display = 'block';
                showStatus('❌ Data default kosong atau format tidak sesuai!', 'error');
                return;
            }

            const latestRow = validRows.reduce((latest, current) => {
                return getRowTimestamp(current) > getRowTimestamp(latest) ? current : latest;
            }, validRows[0]);

            const latestSource = (latestRow.Source || '').trim();
            const latestMonth = (latestRow.Bulan || '').trim();

            let latestData = validRows;
            if (latestSource) {
                latestData = validRows.filter(row => (row.Source || '').trim() === latestSource);
            } else if (latestMonth) {
                latestData = validRows.filter(row => (row.Bulan || '').trim() === latestMonth);
            }

            processData(latestData);
            showStatus('✅ Data default terbaru berhasil dimuat!', 'success');
        })
        .catch(error => {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('dashboard').style.display = 'block';
            showStatus('❌ Error: ' + error.message + '. Pastikan file data/rekening-bank.csv ada!', 'error');
        });
}

// Process and display data
function processData(data) {
    // Filter empty rows
    data = data
        .map(normalizeCsvRow)
        .filter(isValidTransactionRow)
        .filter(isAllowedDisplayChannel)
        .sort((a, b) => getRowTimestamp(a) - getRowTimestamp(b));

    latestProcessedRows = data;
    
    if (data.length === 0) {
        document.getElementById('loading').style.display = 'none';
        document.getElementById('dashboard').style.display = 'block';
        showStatus('❌ Data tidak tampil: periksa format kolom CSV dan pastikan kanal termasuk Teller, Mobile Banking, QRIS, EDC, atau ATM.', 'error');
        return;
    }
    
    // Destroy existing charts
    destroyCharts();
    
    // Calculate statistics
    let totalCredit = 0;
    let transactionCount = data.length;
    let accountNo = '';
    
    const channelData = {};
    const channelTransactionData = {};
    const revenueTypeData = {};
    const hourlyData = Array(24).fill(0);
    const balanceOverTime = [];
    const monthlyBalanceMap = new Map();
    const fileSourceData = {};
    const monthlyDistributionMap = new Map();
    let runningBalance = 0;
    
    data.forEach(row => {
        const credit = parseAmount(row['Credit Amount']);
        const pokok = parseAmount(row.Pokok);
        const denda = parseAmount(row.Denda);
        const effectiveCredit = credit > 0 ? credit : (pokok + denda);
        const debit = parseAmount(row['Debit Amount']);
        const balance = parseAmount(row['Close Balance']);
        const hasExplicitCloseBalance = Boolean(row.__hasExplicitCloseBalance);
        const date = row['Tanggal'] || row['PostDate'];
        const rowTimestamp = getRowTimestamp(row);
        const channel = resolveChannelName(row);
        const revenueType = row['Jenis Penerimaan'] || 'Lainnya';
        const includeMonetary = effectiveCredit > 0;
        const time = row['Jam'];
        const sourceFile = resolveSourceName(row);
        
        if (!accountNo && row['AccountNo']) {
            accountNo = row['AccountNo'];
        }
        
        if (includeMonetary) {
            totalCredit += effectiveCredit;
        }
        
        // Channel data
        if (includeMonetary && channel) {
            channelData[channel] = (channelData[channel] || 0) + effectiveCredit;
        }

        if (channel) {
            channelTransactionData[channel] = (channelTransactionData[channel] || 0) + 1;
        }
        
        // Revenue type data
        if (includeMonetary && revenueType) {
            revenueTypeData[revenueType] = (revenueTypeData[revenueType] || 0) + effectiveCredit;
        }

        if (rowTimestamp > 0 && includeMonetary) {
            const rowDate = new Date(rowTimestamp);
            const monthKey = `${rowDate.getFullYear()}-${String(rowDate.getMonth() + 1).padStart(2, '0')}`;
            const monthLabel = rowDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });

            if (!monthlyDistributionMap.has(monthKey)) {
                monthlyDistributionMap.set(monthKey, {
                    key: monthKey,
                    label: monthLabel,
                    channelData: {},
                    revenueData: {}
                });
            }

            const monthBucket = monthlyDistributionMap.get(monthKey);
            if (channel) {
                monthBucket.channelData[channel] = (monthBucket.channelData[channel] || 0) + effectiveCredit;
            }
            if (revenueType) {
                monthBucket.revenueData[revenueType] = (monthBucket.revenueData[revenueType] || 0) + effectiveCredit;
            }
        }
        
        // Hourly data
        if (time) {
            const hour = parseInt(time.split(':')[0]);
            if (!isNaN(hour) && hour >= 0 && hour < 24) {
                hourlyData[hour]++;
            }
        }

        // Source file data
        if (includeMonetary) {
            fileSourceData[sourceFile] = (fileSourceData[sourceFile] || 0) + effectiveCredit;
        }

        if (balance > 0) {
            runningBalance = balance;
        } else {
            runningBalance += (credit - debit);
        }
        
        // Balance over time
        const effectiveBalance = balance > 0 ? balance : runningBalance;
        if (date && rowTimestamp > 0 && Number.isFinite(effectiveBalance) && effectiveBalance !== 0) {
            const rowDate = new Date(rowTimestamp);
            const dateKey = `${rowDate.getFullYear()}-${String(rowDate.getMonth() + 1).padStart(2, '0')}-${String(rowDate.getDate()).padStart(2, '0')}`;
            balanceOverTime.push({
                date: formatDateLabel(rowTimestamp),
                balance: effectiveBalance,
                timestamp: rowTimestamp,
                dateKey: dateKey
            });
        }

        if (rowTimestamp > 0) {
            const rowDate = new Date(rowTimestamp);
            const monthKey = `${rowDate.getFullYear()}-${String(rowDate.getMonth() + 1).padStart(2, '0')}`;
            const monthLabel = rowDate.toLocaleDateString('id-ID', { month: 'short', year: 'numeric' });

            if (!monthlyBalanceMap.has(monthKey)) {
                monthlyBalanceMap.set(monthKey, {
                    key: monthKey,
                    label: monthLabel,
                    totalCredit: 0,
                    latestBalance: effectiveBalance,
                    latestTimestamp: rowTimestamp
                });
            }

            const monthData = monthlyBalanceMap.get(monthKey);
            if (includeMonetary) {
                monthData.totalCredit += effectiveCredit;
            }

            if (rowTimestamp >= monthData.latestTimestamp) {
                monthData.latestTimestamp = rowTimestamp;
                monthData.latestBalance = effectiveBalance;
            }
        }
    });

    balanceOverTime.sort((a, b) => a.timestamp - b.timestamp);

    const dailyBalanceOverTime = Array.from(
        balanceOverTime.reduce((map, item) => {
            map.set(item.dateKey, item);
            return map;
        }, new Map()).values()
    ).sort((a, b) => a.timestamp - b.timestamp);

    const balanceSeries = dailyBalanceOverTime.length >= 3 ? dailyBalanceOverTime : balanceOverTime;
    const rawMonthlySeries = Array.from(monthlyBalanceMap.values()).sort((a, b) => a.key.localeCompare(b.key));
    const now = new Date();
    const fixedYear = String(now.getFullYear());
    const fixedMonthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const monthlyMapByKey = new Map(rawMonthlySeries.map(item => [item.key, item]));
    const currentYearMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const monthlySeries = fixedMonthLabels.map((monthLabel, monthIndex) => {
        const monthKey = `${fixedYear}-${String(monthIndex + 1).padStart(2, '0')}`;
        const source = monthlyMapByKey.get(monthKey);
        const isFutureMonth = monthKey > currentYearMonthKey;
        return {
            key: monthKey,
            label: `${monthLabel} ${fixedYear}`,
            latestBalance: (!isFutureMonth && source) ? source.latestBalance : 0,
            totalCredit: (!isFutureMonth && source) ? source.totalCredit : 0
        };
    });

    const availableMonthlyRowsCount = rawMonthlySeries.reduce((total, item) => {
        return total + ((item.totalCredit > 0 || item.latestBalance > 0) ? 1 : 0);
    }, 0);
    const balanceChartData = {
        daily: {
            labels: balanceSeries.map(item => item.date),
            balances: balanceSeries.map(item => item.balance),
            credits: balanceSeries.map(() => null)
        },
        monthly: {
            labels: monthlySeries.map(item => item.label),
            balances: monthlySeries.map(item => item.latestBalance),
            credits: monthlySeries.map(item => item.totalCredit)
        },
        hasDaily: balanceSeries.length > 0,
        hasMonthly: availableMonthlyRowsCount > 0
    };

    const monthlyDistribution = Array.from(monthlyDistributionMap.values())
        .sort((a, b) => String(a.key).localeCompare(String(b.key)));
    
    // Update header info
    document.getElementById('accountNo').textContent = accountNo || '-';
    const latestRow = data[data.length - 1] || {};
    const latestPeriod = latestRow['Bulan'] || '-';
    const latestSource = latestRow.__sourceFile || latestRow.Source || 'CSV Upload';
    document.getElementById('period').textContent = latestPeriod;

    updateMetricWithAnimation('currentDatasetLabel', 'currentDatasetLabel', `File CSV Aktif: ${latestSource} • Rekening: ${accountNo || '-'}`);
    updateMetricWithAnimation('currentDatasetSub', 'currentDatasetSub', `Sumber: ${latestSource} • Periode: ${latestPeriod}`);
    
    // Update statistics
    updateMetricWithAnimation('totalCredit', 'totalCredit', formatCurrency(totalCredit));
    updateMetricWithAnimation('totalTransactions', 'totalTransactions', transactionCount);
    updateCreditBreakdown(channelData);
    updateTransactionBreakdown(channelTransactionData);

    const strategicInsights = buildStrategicInsights(data);

    // Update executive summary
    const dashboardStatus = updateExecutiveSummary(totalCredit, transactionCount, channelData, channelTransactionData, revenueTypeData, hourlyData, fileSourceData, strategicInsights);
    
    // Chart colors
    const colors = [
        '#667eea', '#764ba2', '#f093fb', '#4facfe',
        '#43e97b', '#fa709a', '#fee140', '#30cfd0',
        '#a8edea', '#fed6e3', '#c471f5', '#fa8bff'
    ];
    
    // Create charts
    createCharts(balanceChartData, channelData, revenueTypeData, fileSourceData, colors, dashboardStatus, strategicInsights, monthlyDistribution);
    
    // Show dashboard
    document.getElementById('loading').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';
    switchDashboardMenu(activeDashboardMenu);

    if (activeDashboardMenu === 'laporan') {
        loadGoogleSheetReport(true);
    }
}

function stopBalanceTrendArrowAnimation(chart) {
    if (!chart) return;

    if (chart.$balanceTrendArrowAnimationFrameId) {
        cancelAnimationFrame(chart.$balanceTrendArrowAnimationFrameId);
        chart.$balanceTrendArrowAnimationFrameId = 0;
    }

    clearBalanceTrendArrowTooltip(chart);
}

function clearBalanceTrendArrowTooltip(chart) {
    if (!chart || !chart.tooltip) return;

    chart.$balanceTrendTooltipActiveIndex = -1;
    chart.setActiveElements([]);
    chart.tooltip.setActiveElements([], { x: 0, y: 0 });
    chart.$balanceTrendArrowInfo = null;
}

function updateBalanceTrendArrowInfo(chart, pointIndex, anchorPoint = null, lineDatasetIndex = -1) {
    if (!chart || !Number.isInteger(pointIndex) || pointIndex < 0) {
        if (chart) chart.$balanceTrendArrowInfo = null;
        return;
    }

    const labels = chart.data?.labels || [];
    const barDatasetIndex = chart.data?.datasets?.findIndex(dataset => dataset?.label === 'Pendapatan Transaksi Bulanan' && dataset?.type === 'bar' && !dataset?.hidden);
    const safeBarDatasetIndex = barDatasetIndex >= 0 ? barDatasetIndex : 1;
    const safeLineDatasetIndex = lineDatasetIndex >= 0 ? lineDatasetIndex : getBalanceTrendArrowDatasetIndex(chart);

    const barDataset = chart.data?.datasets?.[safeBarDatasetIndex] || {};
    const lineDataset = chart.data?.datasets?.[safeLineDatasetIndex] || {};
    const barValues = barDataset.data || [];
    const lineValues = lineDataset.data || [];

    const currentValue = Number(lineValues[pointIndex] || 0);
    const prevValue = pointIndex > 0 ? Number(lineValues[pointIndex - 1] || 0) : currentValue;
    let trendLabel = 'Stabil';
    if (currentValue > prevValue) trendLabel = 'Naik';
    if (currentValue < prevValue) trendLabel = 'Turun';

    const totalValue = Number(barValues[pointIndex] || 0);
    const infoLines = [
        `${String(barDataset.label || 'Pendapatan Transaksi Bulanan')}: ${formatCurrency(totalValue)}`,
        `${String(lineDataset.label || 'Tren Pendapatan Bulanan')}: ${formatCurrency(currentValue)} (${trendLabel})`
    ];

    chart.$balanceTrendArrowInfo = {
        monthLabel: String(labels[pointIndex] || '-'),
        lines: infoLines,
        x: Number(anchorPoint?.x || 0),
        y: Number(anchorPoint?.y || 0)
    };
}

function showBalanceTrendTooltipAtIndex(chart, lineDatasetIndex, pointIndex, anchorPoint = null) {
    if (!chart || !chart.tooltip || !Number.isInteger(pointIndex) || pointIndex < 0) return;

    const activeElements = (chart.data?.datasets || [])
        .map((dataset, datasetIndex) => ({ dataset, datasetIndex }))
        .filter(({ dataset }) => !dataset?.hidden)
        .filter(({ dataset, datasetIndex }) => {
            const value = dataset?.data?.[pointIndex];
            if (!Number.isFinite(Number(value))) return false;
            const metaPoint = chart.getDatasetMeta(datasetIndex)?.data?.[pointIndex];
            return Boolean(metaPoint) && Number.isFinite(Number(metaPoint?.x)) && Number.isFinite(Number(metaPoint?.y));
        })
        .map(({ datasetIndex }) => ({ datasetIndex, index: pointIndex }));

    if (!activeElements.length) {
        clearBalanceTrendArrowTooltip(chart);
        return;
    }

    const linePoint = chart.getDatasetMeta(lineDatasetIndex)?.data?.[pointIndex];
    const anchorX = Number(anchorPoint?.x ?? linePoint?.x ?? 0);
    const anchorY = Number(anchorPoint?.y ?? linePoint?.y ?? 0);

    chart.setActiveElements(activeElements);
    chart.tooltip.setActiveElements(activeElements, { x: anchorX, y: anchorY });
    chart.$balanceTrendTooltipActiveIndex = pointIndex;
}

function updateBalanceTrendArrowTooltip(chart, arrowState) {
    if (!chart || !chart.tooltip || !arrowState) return;

    const hitThreshold = 10;
    const pointIndex = Number(arrowState.nearestPointDataIndex);
    const pointDistance = Number(arrowState.nearestPointDistance);

    if (!Number.isInteger(pointIndex) || pointIndex < 0 || !Number.isFinite(pointDistance) || pointDistance > hitThreshold) {
        if (chart.$balanceTrendTooltipActiveIndex !== -1) {
            clearBalanceTrendArrowTooltip(chart);
        }
        return;
    }

    showBalanceTrendTooltipAtIndex(chart, arrowState.lineDatasetIndex, pointIndex, {
        x: arrowState.arrowX,
        y: arrowState.arrowY
    });
}

function getBalanceTrendArrowDatasetIndex(chart) {
    if (!chart?.data?.datasets) return -1;
    return chart.data.datasets.findIndex(dataset => dataset?.label === 'Tren Pendapatan Bulanan' && dataset?.type === 'line' && !dataset?.hidden);
}

function getBalanceTrendArrowPathMetrics(chart) {
    const lineDatasetIndex = getBalanceTrendArrowDatasetIndex(chart);
    if (lineDatasetIndex < 0) return null;

    const lineMeta = chart.getDatasetMeta(lineDatasetIndex);
    const points = (lineMeta?.data || [])
        .map((point, dataIndex) => ({
            point,
            dataIndex,
            x: Number(point?.x),
            y: Number(point?.y)
        }))
        .filter(item => Number.isFinite(item.x) && Number.isFinite(item.y));
    if (points.length < 2) return null;

    const segmentLengths = [];
    const cumulativeDistances = [0];
    let totalLength = 0;

    for (let index = 0; index < points.length - 1; index += 1) {
        const from = points[index];
        const to = points[index + 1];
        const length = Math.hypot(to.x - from.x, to.y - from.y);
        segmentLengths.push(length);
        totalLength += length;
        cumulativeDistances.push(totalLength);
    }

    if (totalLength <= 0) return null;

    return {
        lineDatasetIndex,
        points,
        segmentLengths,
        cumulativeDistances,
        totalLength
    };
}

function getBalanceTrendArrowState(chart) {
    const pathMetrics = getBalanceTrendArrowPathMetrics(chart);
    if (!pathMetrics) return null;

    const { lineDatasetIndex, points, segmentLengths, totalLength } = pathMetrics;
    const progress = Number(chart.$balanceTrendArrowProgress || 0);
    let targetDistance = progress * totalLength;
    let segmentIndex = 0;

    while (segmentIndex < segmentLengths.length - 1 && targetDistance > segmentLengths[segmentIndex]) {
        targetDistance -= segmentLengths[segmentIndex];
        segmentIndex += 1;
    }

    const fromPoint = points[segmentIndex];
    const toPoint = points[segmentIndex + 1];
    const currentSegmentLength = segmentLengths[segmentIndex] || 1;
    const ratio = Math.min(Math.max(targetDistance / currentSegmentLength, 0), 1);
    const arrowX = fromPoint.x + (toPoint.x - fromPoint.x) * ratio;
    const arrowY = fromPoint.y + (toPoint.y - fromPoint.y) * ratio;
    const angle = Math.atan2(toPoint.y - fromPoint.y, toPoint.x - fromPoint.x);

    const values = chart.data.datasets[lineDatasetIndex]?.data || [];
    const fromValue = Number(values[fromPoint.dataIndex] || 0);
    const toValue = Number(values[toPoint.dataIndex] || 0);
    let arrowColor = '#64748b';
    if (toValue > fromValue) arrowColor = '#16a34a';
    if (toValue < fromValue) arrowColor = '#dc2626';

    let nearestPointDataIndex = -1;
    let nearestPointDistance = Number.POSITIVE_INFINITY;
    points.forEach((point) => {
        const distance = Math.hypot(point.x - arrowX, point.y - arrowY);
        if (distance < nearestPointDistance) {
            nearestPointDistance = distance;
            nearestPointDataIndex = point.dataIndex;
        }
    });

    return {
        lineDatasetIndex,
        arrowX,
        arrowY,
        angle,
        arrowColor,
        nearestPointDataIndex,
        nearestPointDistance
    };
}

function startBalanceTrendArrowAnimation(chart) {
    if (!chart) return;

    stopBalanceTrendArrowAnimation(chart);
    chart.$balanceTrendTooltipActiveIndex = -1;
    chart.$balanceTrendArrowProgress = 0;
    chart.$balanceTrendArrowPhase = 'move';
    chart.$balanceTrendArrowTargetPointIndex = 1;
    chart.$balanceTrendArrowPhaseStart = performance.now();

    const animateArrow = (timestamp) => {
        if (!chart || !chart.ctx || !chart.canvas) return;

        const pathMetrics = getBalanceTrendArrowPathMetrics(chart);
        if (!pathMetrics) {
            chart.$balanceTrendArrowRenderState = null;
            clearBalanceTrendArrowTooltip(chart);
            chart.draw();
            chart.$balanceTrendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
            return;
        }

        const travelDuration = 6500;
        const pointPauseDuration = 2000;
        const maxPointIndex = pathMetrics.points.length - 1;
        const currentTarget = Math.min(Math.max(Number(chart.$balanceTrendArrowTargetPointIndex || 1), 1), maxPointIndex);
        const phaseStart = Number(chart.$balanceTrendArrowPhaseStart || timestamp);
        const elapsedInPhase = Math.max(0, timestamp - phaseStart);

        if (chart.$balanceTrendArrowPhase === 'hold') {
            const holdPointIndex = Math.min(Math.max(Number(chart.$balanceTrendArrowHoldPointIndex || currentTarget), 1), maxPointIndex);
            chart.$balanceTrendArrowProgress = pathMetrics.cumulativeDistances[holdPointIndex] / pathMetrics.totalLength;

            const holdPoint = pathMetrics.points[holdPointIndex];
            const holdDataIndex = Number(holdPoint?.dataIndex);
            if (!chart.$balanceTrendUserHoverActive && Number.isInteger(holdDataIndex) && holdDataIndex >= 0) {
                showBalanceTrendTooltipAtIndex(chart, pathMetrics.lineDatasetIndex, holdDataIndex, {
                    x: Number(holdPoint?.x || 0),
                    y: Number(holdPoint?.y || 0)
                });

                updateBalanceTrendArrowInfo(chart, holdDataIndex, {
                    x: Number(holdPoint?.x || 0),
                    y: Number(holdPoint?.y || 0)
                }, pathMetrics.lineDatasetIndex);
            }

            if (elapsedInPhase >= pointPauseDuration) {
                if (holdPointIndex >= maxPointIndex) {
                    chart.$balanceTrendArrowTargetPointIndex = 1;
                    chart.$balanceTrendArrowProgress = 0;
                } else {
                    chart.$balanceTrendArrowTargetPointIndex = holdPointIndex + 1;
                }
                chart.$balanceTrendArrowPhase = 'move';
                chart.$balanceTrendArrowPhaseStart = timestamp;
            }
        } else {
            const fromPointIndex = Math.max(currentTarget - 1, 0);
            const segmentLength = Number(pathMetrics.segmentLengths[fromPointIndex] || 0);
            const segmentDuration = Math.max((segmentLength / pathMetrics.totalLength) * travelDuration, 1);
            const linearProgress = Math.min(elapsedInPhase / segmentDuration, 1);
            const easedProgress = 0.5 - (Math.cos(Math.PI * linearProgress) / 2);
            const distanceBeforeSegment = Number(pathMetrics.cumulativeDistances[fromPointIndex] || 0);
            const traveledDistance = distanceBeforeSegment + (segmentLength * easedProgress);
            chart.$balanceTrendArrowProgress = traveledDistance / pathMetrics.totalLength;

            if (!chart.$balanceTrendUserHoverActive && chart.$balanceTrendTooltipActiveIndex !== -1) {
                clearBalanceTrendArrowTooltip(chart);
            }

            if (!chart.$balanceTrendUserHoverActive) {
                chart.$balanceTrendArrowInfo = null;
            }

            if (linearProgress >= 1) {
                chart.$balanceTrendArrowPhase = 'hold';
                chart.$balanceTrendArrowHoldPointIndex = currentTarget;
                chart.$balanceTrendArrowPhaseStart = timestamp;
            }
        }

        chart.$balanceTrendArrowRenderState = getBalanceTrendArrowState(chart);
        chart.draw();
        chart.$balanceTrendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
    };

    chart.$balanceTrendArrowAnimationFrameId = requestAnimationFrame(animateArrow);
}

const balanceTrendArrowPlugin = {
    id: 'balanceTrendArrow',
    afterEvent(chart, args) {
        if (chart?.$suppressExportOverlay) return;

        const eventType = String(args?.event?.type || '');
        if (!eventType) return;

        if (eventType === 'mousemove') {
            chart.$balanceTrendUserHoverActive = Boolean(args?.inChartArea);
            if (chart.$balanceTrendUserHoverActive) {
                chart.$balanceTrendTooltipActiveIndex = -1;
            }
            return;
        }

        if (eventType === 'mouseout' || eventType === 'mouseleave') {
            chart.$balanceTrendUserHoverActive = false;
        }
    },
    afterDatasetsDraw(chart) {
        if (chart?.$suppressExportOverlay) return;

        const arrowState = chart.$balanceTrendArrowRenderState || getBalanceTrendArrowState(chart);
        if (!arrowState) return;

        const context = chart.ctx;
        context.save();
        context.translate(arrowState.arrowX, arrowState.arrowY);
        context.rotate(arrowState.angle);
        context.beginPath();
        context.moveTo(9, 0);
        context.lineTo(-6, 5);
        context.lineTo(-6, -5);
        context.closePath();
        context.fillStyle = arrowState.arrowColor;
        context.fill();
        context.restore();

        const arrowInfo = chart.$balanceTrendArrowInfo;
        if (!arrowInfo) return;

        const titleText = String(arrowInfo.monthLabel || '-');
        const infoLines = Array.isArray(arrowInfo.lines) ? arrowInfo.lines : [];
        if (infoLines.length === 0) return;

        context.save();
        context.font = '12px Roboto';
        const titleWidth = context.measureText(titleText).width;
        const linesWidth = infoLines.reduce((width, lineText) => Math.max(width, context.measureText(String(lineText || '')).width), 0);
        const textWidth = Math.max(titleWidth, linesWidth);
        const paddingX = 8;
        const lineHeight = 15;
        const boxWidth = textWidth + (paddingX * 2);
        const boxHeight = 8 + lineHeight + 4 + (infoLines.length * lineHeight) + 8;
        const area = chart.chartArea || { left: 0, right: chart.width, top: 0, bottom: chart.height };
        let boxX = Number(arrowInfo.x || 0) + 14;
        let boxY = Number(arrowInfo.y || 0) - boxHeight - 10;

        if (boxX + boxWidth > area.right - 4) {
            boxX = Number(arrowInfo.x || 0) - boxWidth - 14;
        }
        if (boxX < area.left + 4) {
            boxX = area.left + 4;
        }
        if (boxY < area.top + 4) {
            boxY = Number(arrowInfo.y || 0) + 10;
        }

        context.fillStyle = 'rgba(15, 23, 42, 0.9)';
        context.beginPath();
        context.roundRect(boxX, boxY, boxWidth, boxHeight, 6);
        context.fill();

        context.fillStyle = '#ffffff';
        context.textBaseline = 'top';
        context.fillText(titleText, boxX + paddingX, boxY + 8);

        let currentTextY = boxY + 8 + lineHeight + 4;
        infoLines.forEach((lineText) => {
            context.fillText(String(lineText || ''), boxX + paddingX, currentTextY);
            currentTextY += lineHeight;
        });
        context.restore();
    },
    afterDestroy(chart) {
        stopBalanceTrendArrowAnimation(chart);
    }
};

// Create all charts
function createCharts(balanceChartData, channelData, revenueTypeData, fileSourceData, colors, dashboardStatus = 'neutral', strategicInsights = null, monthlyDistribution = []) {
    const textColor = '#1f2937';
    const mutedTextColor = '#6b7280';
    const statusPalette = {
        warning: {
            line: '#d97706',
            areaTop: 'rgba(217, 119, 6, 0.35)',
            areaBottom: 'rgba(217, 119, 6, 0.03)',
            barTop: 'rgba(245, 158, 11, 0.95)',
            barBottom: 'rgba(245, 158, 11, 0.42)',
            barBorder: '#b45309',
            grid: 'rgba(217, 119, 6, 0.18)'
        },
        success: {
            line: '#16a34a',
            areaTop: 'rgba(22, 163, 74, 0.34)',
            areaBottom: 'rgba(22, 163, 74, 0.03)',
            barTop: 'rgba(34, 197, 94, 0.95)',
            barBottom: 'rgba(34, 197, 94, 0.42)',
            barBorder: '#15803d',
            grid: 'rgba(34, 197, 94, 0.18)'
        },
        neutral: {
            line: '#6366f1',
            areaTop: 'rgba(99, 102, 241, 0.35)',
            areaBottom: 'rgba(99, 102, 241, 0.02)',
            barTop: 'rgba(79, 172, 254, 0.95)',
            barBottom: 'rgba(79, 172, 254, 0.42)',
            barBorder: '#3b82f6',
            grid: 'rgba(148, 163, 184, 0.22)'
        }
    };
    const palette = statusPalette[dashboardStatus] || statusPalette.neutral;
    const gridColor = palette.grid;
    const tickFont = {
        family: "'Roboto', sans-serif",
        size: 12,
        weight: '600'
    };

    const balanceCanvas = document.getElementById('balanceChart');
    const balanceContext = balanceCanvas.getContext('2d');
    const balanceGradient = balanceContext.createLinearGradient(0, 0, 0, balanceCanvas.height || 360);
    balanceGradient.addColorStop(0, palette.areaTop);
    balanceGradient.addColorStop(1, palette.areaBottom);

    const barCanvas = document.getElementById('fileSourceChart');
    const barContext = barCanvas.getContext('2d');
    const barGradient = barContext.createLinearGradient(0, 0, 0, barCanvas.height || 360);
    barGradient.addColorStop(0, palette.barTop);
    barGradient.addColorStop(1, palette.barBottom);

    const distributionBorderColors = Object.values(channelData).map(() => '#ffffff');
    const revenueBorderColors = Object.values(revenueTypeData).map(() => '#ffffff');

    monthlyDistributionState = {
        months: buildFullMonthDistributionEntries(monthlyDistribution),
        allChannelData: {},
        allRevenueData: {},
        channelByMonth: new Map(),
        revenueByMonth: new Map(),
        channelIndex: -1,
        revenueIndex: -1,
        colors: Array.isArray(colors) ? colors : []
    };
    monthlyDistributionState.months.forEach((monthEntry) => {
        monthlyDistributionState.channelByMonth.set(monthEntry.key, { ...(monthEntry.channelData || {}) });
        monthlyDistributionState.revenueByMonth.set(monthEntry.key, { ...(monthEntry.revenueData || {}) });
    });
    monthlyDistributionState.allChannelData = aggregateDistributionMaps(monthlyDistributionState.months, 'channelData');
    monthlyDistributionState.allRevenueData = aggregateDistributionMaps(monthlyDistributionState.months, 'revenueData');

    const commonLegend = {
        position: 'bottom',
        labels: {
            color: textColor,
            usePointStyle: true,
            pointStyle: 'circle',
            boxWidth: 10,
            boxHeight: 10,
            padding: 14,
            font: {
                family: "'Roboto', sans-serif",
                size: 11,
                weight: '600'
            }
        }
    };

    const commonTooltip = {
        backgroundColor: 'rgba(15, 23, 42, 0.92)',
        titleColor: '#f8fafc',
        bodyColor: '#e2e8f0',
        borderColor: 'rgba(148, 163, 184, 0.35)',
        borderWidth: 1,
        cornerRadius: 10,
        padding: 12,
        displayColors: false
    };

    let activeMode = balanceViewMode;
    if (activeMode === 'monthly' && !balanceChartData.hasMonthly) {
        activeMode = 'daily';
    }
    if (activeMode === 'daily' && !balanceChartData.hasDaily) {
        activeMode = 'monthly';
    }
    balanceViewMode = activeMode;
    updateBalanceToggleControls(activeMode, balanceChartData.hasDaily, balanceChartData.hasMonthly);

    const selectedBalanceData = activeMode === 'monthly' ? balanceChartData.monthly : balanceChartData.daily;

    // Balance chart
    chartInstances.balance = new Chart(balanceCanvas, {
        plugins: [balanceTrendArrowPlugin],
        type: 'line',
        data: {
            labels: selectedBalanceData.labels,
            datasets: [
                {
                    label: activeMode === 'monthly' ? 'Saldo Akhir Bulanan' : 'Saldo Harian',
                    data: selectedBalanceData.balances,
                    borderColor: palette.line,
                    backgroundColor: balanceGradient,
                    borderWidth: 3.2,
                    fill: true,
                    tension: 0.35,
                    pointRadius: 2.4,
                    pointHoverRadius: 5,
                    pointHitRadius: 12,
                    pointBackgroundColor: palette.line,
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 1.5,
                    yAxisID: 'y',
                    hidden: activeMode === 'monthly'
                },
                {
                    label: 'Pendapatan Transaksi Bulanan',
                    data: selectedBalanceData.credits,
                    type: 'bar',
                    yAxisID: 'y1',
                    backgroundColor: 'rgba(14, 165, 233, 0.28)',
                    borderColor: '#0284c7',
                    borderWidth: 1,
                    borderRadius: 7,
                    borderSkipped: false,
                    maxBarThickness: 40,
                    hidden: activeMode !== 'monthly'
                },
                {
                    label: 'Tren Pendapatan Bulanan',
                    data: selectedBalanceData.credits,
                    type: 'line',
                    yAxisID: 'y1',
                    borderColor: '#0369a1',
                    backgroundColor: 'rgba(14, 165, 233, 0.12)',
                    borderWidth: 2.4,
                    fill: false,
                    tension: 0.32,
                    pointRadius: 3,
                    pointHoverRadius: 5,
                    pointBackgroundColor: '#0284c7',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 1.3,
                    spanGaps: true,
                    hidden: activeMode !== 'monthly'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    ...commonLegend,
                    labels: {
                        ...commonLegend.labels,
                        filter: function(item) {
                            if (item.text === 'Saldo Akhir Bulanan') {
                                return false;
                            }
                            if (item.text === 'Pendapatan Transaksi Bulanan') {
                                return activeMode === 'monthly';
                            }
                            if (item.text === 'Tren Pendapatan Bulanan') {
                                return activeMode === 'monthly';
                            }
                            return true;
                        }
                    }
                },
                tooltip: {
                    ...commonTooltip,
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: ${formatCurrency(context.parsed.y)}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    display: activeMode !== 'monthly',
                    beginAtZero: false,
                    grid: {
                        color: gridColor,
                        drawBorder: false
                    },
                    ticks: {
                        color: mutedTextColor,
                        font: tickFont,
                        padding: 8,
                        callback: function(value) {
                            return formatCurrencyTick(value);
                        }
                    }
                },
                y1: {
                    display: activeMode === 'monthly',
                    position: 'right',
                    beginAtZero: true,
                    grid: {
                        drawOnChartArea: false,
                        drawBorder: false
                    },
                    ticks: {
                        color: mutedTextColor,
                        font: tickFont,
                        callback: function(value) {
                            return formatCurrencyTick(value);
                        }
                    }
                },
                x: {
                    grid: {
                        display: false,
                        drawBorder: false
                    },
                    ticks: {
                        color: mutedTextColor,
                        font: tickFont,
                        padding: 6,
                        autoSkip: true,
                        maxTicksLimit: 12,
                        maxRotation: 35,
                        minRotation: 35
                    }
                }
            }
        }
    });

    if (activeMode === 'monthly') {
        startBalanceTrendArrowAnimation(chartInstances.balance);
    }

    const trendCanvas = document.getElementById('trendInsightChart');
    if (trendCanvas && strategicInsights?.monthlyTrend?.labels?.length) {
        const trendData = strategicInsights.monthlyTrend;
        const forecastIndex = Math.max(0, (trendData.historicalLength || trendData.labels.length) - 1);

        const forecastLineStyle = {
            borderDash: function(context) {
                return context.p0DataIndex === forecastIndex ? [6, 6] : undefined;
            }
        };

        const forecastPointRadius = function(context) {
            return context.dataIndex === (trendData.labelsWithForecast.length - 1) ? 4 : 2;
        };

        chartInstances.trendInsight = new Chart(trendCanvas, {
            type: 'line',
            data: {
                labels: trendData.labelsWithForecast,
                datasets: [
                    {
                        label: '% Kontribusi Digital',
                        data: trendData.digitalShare,
                        borderColor: '#0ea5e9',
                        backgroundColor: 'rgba(14, 165, 233, 0.12)',
                        borderWidth: 2.5,
                        tension: 0.32,
                        pointRadius: forecastPointRadius,
                        segment: forecastLineStyle,
                        yAxisID: 'y'
                    },
                    {
                        label: '% QRIS',
                        data: trendData.qrisShare,
                        borderColor: '#22c55e',
                        backgroundColor: 'rgba(34, 197, 94, 0.12)',
                        borderWidth: 2.5,
                        tension: 0.32,
                        pointRadius: forecastPointRadius,
                        segment: forecastLineStyle,
                        yAxisID: 'y'
                    },
                    {
                        label: '% Dominasi Jenis Penerimaan',
                        data: trendData.dominantRevenueShare,
                        borderColor: '#f59e0b',
                        backgroundColor: 'rgba(245, 158, 11, 0.12)',
                        borderWidth: 2.5,
                        tension: 0.32,
                        pointRadius: forecastPointRadius,
                        segment: forecastLineStyle,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Rata-rata Nilai Transaksi',
                        data: trendData.avgValue,
                        borderColor: palette.line,
                        backgroundColor: 'rgba(99, 102, 241, 0.1)',
                        borderWidth: 2.8,
                        tension: 0.35,
                        pointRadius: forecastPointRadius,
                        segment: forecastLineStyle,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: commonLegend,
                    tooltip: {
                        ...commonTooltip,
                        displayColors: true,
                        callbacks: {
                            label: function(context) {
                                const isForecast = context.dataIndex === (trendData.labelsWithForecast.length - 1);
                                if (context.dataset.yAxisID === 'y1') {
                                    return `${context.dataset.label}${isForecast ? ' (Prediksi)' : ''}: ${formatCurrency(context.parsed.y)}`;
                                }
                                return `${context.dataset.label}${isForecast ? ' (Prediksi)' : ''}: ${context.parsed.y.toFixed(1)}%`;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        position: 'left',
                        min: 0,
                        max: 100,
                        grid: {
                            color: gridColor,
                            drawBorder: false
                        },
                        ticks: {
                            color: mutedTextColor,
                            font: tickFont,
                            callback: function(value) {
                                return `${value}%`;
                            }
                        }
                    },
                    y1: {
                        position: 'right',
                        grid: {
                            drawOnChartArea: false,
                            drawBorder: false
                        },
                        ticks: {
                            color: mutedTextColor,
                            font: tickFont,
                            callback: function(value) {
                                return formatCurrencyTick(value);
                            }
                        }
                    },
                    x: {
                        grid: {
                            display: false,
                            drawBorder: false
                        },
                        ticks: {
                            color: mutedTextColor,
                            font: tickFont
                        }
                    }
                }
            }
        });
    }
    
    // Channel distribution chart
    chartInstances.channel = new Chart(document.getElementById('channelChart'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(channelData),
            datasets: [{
                data: Object.values(channelData),
                backgroundColor: colors,
                borderColor: distributionBorderColors,
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: commonLegend,
                tooltip: {
                    ...commonTooltip,
                    callbacks: {
                        label: function(context) {
                            return formatDistributionTooltipLabel(context);
                        }
                    }
                }
            },
            cutout: '62%'
        }
    });
    
    // Revenue type chart
    chartInstances.revenue = new Chart(document.getElementById('revenueTypeChart'), {
        type: 'pie',
        data: {
            labels: Object.keys(revenueTypeData),
            datasets: [{
                data: Object.values(revenueTypeData),
                backgroundColor: colors,
                borderColor: revenueBorderColors,
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: commonLegend,
                tooltip: {
                    ...commonTooltip,
                    callbacks: {
                        label: function(context) {
                            return formatDistributionTooltipLabel(context);
                        }
                    }
                }
            }
        }
    });

    updateDistributionMonthlyFilterControl('channel');
    updateDistributionMonthlyFilterControl('revenue');
    applyChannelChartMonthlyFilter();
    applyRevenueChartMonthlyFilter();
    
    // File source summary chart
    chartInstances.fileSource = new Chart(barCanvas, {
        type: 'bar',
        data: {
            labels: Object.keys(fileSourceData),
            datasets: [{
                label: 'Total Pemasukan',
                data: Object.values(fileSourceData),
                backgroundColor: barGradient,
                borderColor: palette.barBorder,
                borderWidth: 1,
                borderRadius: 10,
                borderSkipped: false,
                maxBarThickness: 48
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    ...commonTooltip,
                    callbacks: {
                        label: function(context) {
                            return 'Total: ' + formatCurrency(context.parsed.y);
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: gridColor,
                        drawBorder: false
                    },
                    ticks: {
                        color: mutedTextColor,
                        font: tickFont,
                        callback: function(value) {
                            return formatCurrencyTick(value);
                        }
                    }
                },
                x: {
                    grid: {
                        display: false,
                        drawBorder: false
                    },
                    ticks: {
                        color: mutedTextColor,
                        font: tickFont,
                        maxRotation: 45,
                        minRotation: 20
                    }
                }
            }
        }
    });
}

// Initialize on page load
window.addEventListener('DOMContentLoaded', function() {
    loadDateParsingMode();
    initTheme();
    initDashboardEnhancements();
    initGlobalErrorLogging();

    if (checkAuth()) {
        const dashboardElement = document.getElementById('dashboard');
        if (dashboardElement) {
            dashboardElement.style.display = 'block';
        }

        const hasManualPrimaryData = loadManualDatasetAsPrimary();
        if (hasManualPrimaryData) {
            activeDashboardMenu = 'laporan';
            switchDashboardMenu('laporan');
            showStatus('ℹ️ Data input manual tersimpan dimuat sebagai laporan utama.', 'warning');
            return;
        }

        switchDashboardMenu(activeDashboardMenu);
        loadDefaultCSV();
    }
});
