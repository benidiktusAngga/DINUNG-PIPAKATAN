// Global variables
let chartInstances = {};
let currentUser = null;
let balanceViewMode = 'monthly';
let latestProcessedRows = [];
let activeDashboardMenu = 'sorotan';
const USERS_STORAGE_KEY = 'appUsers';
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

function showUserManageStatus(message, type) {
    const statusElement = document.getElementById('userManageStatus');
    if (!statusElement) return;
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = 'block';
}

function onManagedUserChange() {
    const userSelect = document.getElementById('manageUserSelect');
    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const photoUrlInput = document.getElementById('managePhotoUrl');
    const roleSelect = document.getElementById('manageRole');
    if (!userSelect || !usernameInput || !passwordInput || !fullNameInput || !photoUrlInput || !roleSelect) return;

    const users = getStoredUsers();
    const selectedUser = users.find(user => user.username === userSelect.value);
    if (!selectedUser) return;

    usernameInput.value = selectedUser.username || '';
    passwordInput.value = selectedUser.password || '';
    fullNameInput.value = selectedUser.fullName || '';
    photoUrlInput.value = selectedUser.photoUrl || '';
    roleSelect.value = selectedUser.role || 'guest';
}

function renderUserManagement() {
    const userSelect = document.getElementById('manageUserSelect');
    if (!userSelect) return;

    const users = getStoredUsers();
    const previousSelectedUsername = userSelect.value;
    userSelect.innerHTML = users
        .map(user => `<option value="${user.username}">${user.username} (${user.role})</option>`)
        .join('');

    if (users.length > 0) {
        const selectedExists = users.some(user => user.username === previousSelectedUsername);
        userSelect.value = selectedExists ? previousSelectedUsername : users[0].username;
        onManagedUserChange();
    }
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
    const photoUrlInput = document.getElementById('managePhotoUrl');
    const roleSelect = document.getElementById('manageRole');
    if (!userSelect || !usernameInput || !passwordInput || !fullNameInput || !photoUrlInput || !roleSelect) return;

    const selectedUsername = String(userSelect.value || '').trim();
    const newUsername = String(usernameInput.value || '').trim();
    const newPassword = String(passwordInput.value || '').trim();
    const newFullName = String(fullNameInput.value || '').trim();
    const newPhotoUrl = String(photoUrlInput.value || '').trim();
    const newRole = String(roleSelect.value || 'guest').trim();

    if (!selectedUsername) {
        showUserManageStatus('❌ Pilih user yang akan diubah.', 'error');
        return;
    }

    if (!newUsername || !newPassword) {
        showUserManageStatus('❌ Username dan password baru wajib diisi.', 'error');
        return;
    }

    const users = getStoredUsers();
    const userIndex = users.findIndex(user => user.username === selectedUsername);
    if (userIndex < 0) {
        showUserManageStatus('❌ Data user tidak ditemukan.', 'error');
        return;
    }

    const duplicateUser = users.find((user, index) => index !== userIndex && user.username === newUsername);
    if (duplicateUser) {
        showUserManageStatus('❌ Username sudah dipakai user lain.', 'error');
        return;
    }

    const oldUsername = users[userIndex].username;
    users[userIndex].username = newUsername;
    users[userIndex].password = newPassword;
    users[userIndex].fullName = newFullName || users[userIndex].fullName;
    users[userIndex].photoUrl = newPhotoUrl;
    users[userIndex].role = ['admin', 'user', 'guest'].includes(newRole) ? newRole : 'guest';
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
        refreshedSelect.value = newUsername;
        onManagedUserChange();
    }
    showUserManageStatus('✅ Data user berhasil diperbarui.', 'success');
}

function addManagedUser() {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat menambah user.', 'error');
        return;
    }

    const usernameInput = document.getElementById('manageUsername');
    const passwordInput = document.getElementById('managePassword');
    const fullNameInput = document.getElementById('manageFullName');
    const photoUrlInput = document.getElementById('managePhotoUrl');
    const roleSelect = document.getElementById('manageRole');
    if (!usernameInput || !passwordInput || !fullNameInput || !photoUrlInput || !roleSelect) return;

    const newUsername = String(usernameInput.value || '').trim();
    const newPassword = String(passwordInput.value || '').trim();
    const newFullName = String(fullNameInput.value || '').trim();
    const newPhotoUrl = String(photoUrlInput.value || '').trim();
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
        photoUrl: newPhotoUrl,
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

function deleteManagedUser() {
    if (!isAdminUser()) {
        showUserManageStatus('❌ Hanya admin yang dapat menghapus user.', 'error');
        return;
    }

    const userSelect = document.getElementById('manageUserSelect');
    if (!userSelect) return;

    const selectedUsername = String(userSelect.value || '').trim();
    if (!selectedUsername) {
        showUserManageStatus('❌ Pilih user yang akan dihapus.', 'error');
        return;
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

    if (!confirm(`Hapus user ${selectedUsername}?`)) {
        return;
    }

    const filteredUsers = users.filter(user => user.username !== selectedUsername);
    saveStoredUsers(filteredUsers);
    renderUserManagement();
    showUserManageStatus('✅ User berhasil dihapus.', 'success');
}

function switchDashboardMenu(menuKey) {
    const validMenus = ['sorotan', 'grafik', 'tren', 'user'];
    if (!validMenus.includes(menuKey)) return;

    if (menuKey === 'user' && !isAdminUser()) {
        showStatus('❌ Anda tidak memiliki akses ke menu Pengelolaan User.', 'error');
        return;
    }

    activeDashboardMenu = menuKey;

    const buttonMap = {
        sorotan: document.getElementById('menuBtnSorotan'),
        grafik: document.getElementById('menuBtnGrafik'),
        tren: document.getElementById('menuBtnTren'),
        user: document.getElementById('menuBtnUser')
    };

    const sectionMap = {
        sorotan: document.getElementById('menuSectionSorotan'),
        grafik: document.getElementById('menuSectionGrafik'),
        tren: document.getElementById('menuSectionTren'),
        user: document.getElementById('menuSectionUser')
    };

    Object.entries(buttonMap).forEach(([key, element]) => {
        if (!element) return;
        element.classList.toggle('active', key === menuKey);
    });

    Object.entries(sectionMap).forEach(([key, element]) => {
        if (!element) return;
        element.classList.toggle('active', key === menuKey);
    });

    if (menuKey === 'user') {
        renderUserManagement();
    }

    setTimeout(() => {
        Object.values(chartInstances).forEach(chart => {
            if (chart && typeof chart.resize === 'function') {
                chart.resize();
            }
        });
    }, 80);
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
        return true;
    } catch (error) {
        console.error('Error parsing session:', error);
        localStorage.removeItem('userSession');
        window.location.href = 'login.html';
        return false;
    }
}

// Update user display
function updateUserDisplay() {
    if (!currentUser) return;
    
    // Set avatar based on role
    const avatars = {
        'admin': '👑',
        'user': '👤',
        'guest': '👥'
    };
    
    const avatarImage = document.getElementById('userAvatarImage');
    const avatarFallback = document.getElementById('userAvatarFallback');
    const photoUrl = String(currentUser.photoUrl || '').trim();

    if (avatarFallback) {
        avatarFallback.textContent = avatars[currentUser.role] || '👤';
    }

    if (avatarImage) {
        if (photoUrl) {
            avatarImage.src = photoUrl;
            avatarImage.style.display = 'block';
            if (avatarFallback) avatarFallback.style.display = 'none';
        } else {
            avatarImage.removeAttribute('src');
            avatarImage.style.display = 'none';
            if (avatarFallback) avatarFallback.style.display = 'inline';
        }

        avatarImage.onerror = function() {
            avatarImage.style.display = 'none';
            if (avatarFallback) avatarFallback.style.display = 'inline';
        };
    }
    document.getElementById('userName').textContent = currentUser.fullName;
    document.getElementById('userRole').textContent = `Role: ${currentUser.role}`;
    
    // Set role badge
    const badge = document.getElementById('roleBadge');
    badge.textContent = currentUser.role;
    badge.className = `role-badge ${currentUser.role}`;

    const sidebarBrandTitle = document.getElementById('sidebarBrandTitle');
    if (sidebarBrandTitle) {
        const displayName = currentUser.fullName || currentUser.username || 'User';
        sidebarBrandTitle.textContent = `📁 ${displayName}`;
    }
}

// Apply permissions
function applyPermissions() {
    if (!currentUser) return;
    
    const permissions = currentUser.permissions || [];
    
    // Check upload permission
    if (!permissions.includes('upload')) {
        const btnLoad = document.getElementById('btnLoadCSV');
        const fileInput = document.getElementById('csvFile');
        
        btnLoad.disabled = true;
        fileInput.disabled = true;
        
        showStatus('⚠️ Anda hanya memiliki akses view. Upload CSV dinonaktifkan.', 'warning');
    }

    const menuUserButton = document.getElementById('menuBtnUser');
    const menuUserSection = document.getElementById('menuSectionUser');
    if (menuUserButton) {
        menuUserButton.style.display = isAdminUser() ? 'block' : 'none';
    }
    if (menuUserSection) {
        menuUserSection.style.display = isAdminUser() ? '' : 'none';
    }
    if (!isAdminUser() && activeDashboardMenu === 'user') {
        switchDashboardMenu('sorotan');
    }
}

// Handle logout
function handleLogout() {
    if (confirm('Apakah Anda yakin ingin logout?')) {
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

function parseRowTimestamp(row) {
    const postDateText = String(row?.PostDate || '').trim();
    const parsedPostDate = postDateText ? Date.parse(postDateText) : NaN;
    if (!isNaN(parsedPostDate)) return parsedPostDate;

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

    const directDate = Date.parse(`${dateText} ${normalizedTime}`);
    if (!isNaN(directDate)) return directDate;

    const slashMatch = dateText.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (slashMatch) {
        const day = slashMatch[1];
        const month = slashMatch[2];
        const yearRaw = slashMatch[3];
        const fullYear = yearRaw.length === 2 ? `20${yearRaw}` : yearRaw;
        const isoDate = `${String(fullYear).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${normalizedTime}`;
        const parsedSlashDate = Date.parse(isoDate);
        if (!isNaN(parsedSlashDate)) return parsedSlashDate;
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
        const channelName = row['Jenis Kanal'] || 'Unknown';
        const revenueName = row['Jenis Penerimaan'] || 'Lainnya';
        const sourceName = row.__sourceFile || row.Source || 'Tanpa Sumber';
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

    const orderedMonths = Array.from(monthlyMap.values())
        .sort((a, b) => a.key.localeCompare(b.key))
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
            qrisInsight: 'Data QRIS belum cukup untuk dihitung.',
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
    const peakShift = previous && latest.peakHourIndex >= 0 && previous.peakHourIndex >= 0
        ? Math.abs(latest.peakHourIndex - previous.peakHourIndex)
        : 0;

    const earlyWarnings = [];
    if (digitalDrop > 0.05) earlyWarnings.push(`porsi digital turun ${formatPercent(digitalDrop)} vs bulan lalu`);
    if (avgDrop > 0.1) earlyWarnings.push(`rata-rata nilai transaksi turun ${formatPercent(avgDrop)}`);
    if (peakShift >= 2) earlyWarnings.push(`jam puncak bergeser ${peakShift} jam`);

    const mobileTarget = 0.82;
    const targetGap = mobileTarget - latest.mobileShare;
    const targetText = targetGap <= 0
        ? `Target Mobile Banking sudah tercapai. Capaian saat ini ${formatPercent(latest.mobileShare)}.`
        : `Target Mobile Banking belum tercapai. Capaian saat ini ${formatPercent(latest.mobileShare)}, masih kurang ${formatPercent(targetGap)} dari target ${formatPercent(mobileTarget)}.`;
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

    const segmentationInsight = latest.topSourceShare >= 0.6
        ? `Penerimaan masih sangat bergantung pada ${latest.topSourceLabel} (${formatPercent(latest.topSourceShare)}). Disarankan tambah kolaborasi bank persepsi/BUMD agar risiko lebih terjaga.`
        : `Sumber penerimaan cukup berimbang. Kontributor terbesar saat ini ${latest.topSourceLabel} (${formatPercent(latest.topSourceShare)}).`;

    const qrisPotential15 = Math.max(0, 0.15 - latest.qrisShare) * latest.totalCredit;
    const qrisPotential20 = Math.max(0, 0.20 - latest.qrisShare) * latest.totalCredit;
    const qrisInsight = `Porsi QRIS saat ini ${formatPercent(latest.qrisShare)}. Potensi tambahan jika naik ke 15% sekitar ${formatCurrency(qrisPotential15)}, dan jika ke 20% sekitar ${formatCurrency(qrisPotential20)}. Daftar OPD nonaktif QRIS bisa ditampilkan jika data memuat kolom OPD/Kecamatan.`;

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
        trendInsight,
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

    document.getElementById('summaryTopChannel').textContent = `${topChannel.label} (${formatCurrency(topChannel.value)})`;
    document.getElementById('summaryTopRevenue').textContent = `${topRevenue.label} (${formatCurrency(topRevenue.value)})`;
    document.getElementById('summaryTopSource').textContent = `${topSource.label} (${formatCurrency(topSource.value)})`;
    document.getElementById('summaryAvgCredit').textContent = formatCurrency(avgCredit);

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
        actionText = `⚠️ Warning: Transaksi kanal ${topChannelByCount.label} mendominasi ${(tellerDominanceRatio * 100).toFixed(1)}%. Tingkatkan prioritas kanal non-teller untuk diversifikasi transaksi.`;
    } else if (isNonTellerAboveTeller) {
        actionElement.classList.add('success');
        if (statusIconElement) {
            statusIconElement.classList.add('active', 'success');
            statusIconElement.textContent = '✅ BAGUS';
        }
        actionText = `✅ Bagus: Transaksi non-teller (${nonTellerCount}) sudah di atas teller (${tellerCount}). Lanjutkan peningkatan adopsi kanal digital untuk mendorong pertumbuhan yang lebih tinggi.`;
    } else if (topChannel.value > 0) {
        actionText = `Aksi: Prioritaskan optimalisasi kanal ${topChannel.label} dan dorong replikasi ke kanal lain.`;
    }

    actionElement.textContent = actionText;

    const recommendations = [];
    if (tellerRatio >= 0.5) {
        recommendations.push(`Tetapkan target TP2DD untuk menurunkan porsi teller dari ${(tellerRatio * 100).toFixed(1)}% menjadi <50% dalam 1-2 triwulan melalui migrasi transaksi ke kanal digital.`);
    } else {
        recommendations.push(`Pertahankan momentum ETPD dengan menjaga porsi kanal digital ${Math.max(digitalRatio * 100, 100 - tellerRatio * 100).toFixed(1)}% dan tetapkan KPI kenaikan bulanan transaksi nontunai.`);
    }

    if (qrisRatio < 0.2) {
        recommendations.push(`Percepat adopsi QRIS (saat ini ${(qrisRatio * 100).toFixed(1)}%) lewat perluasan merchant/OPD, sosialisasi wajib scan, dan insentif kanal nontunai.`);
    } else {
        recommendations.push(`Optimalkan performa QRIS (saat ini ${(qrisRatio * 100).toFixed(1)}%) dengan monitoring harian dan perluasan use case pajak/retribusi prioritas.`);
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

    if (targetWarningElement) targetWarningElement.textContent = strategicInsights?.targetWarning || '-';
    if (trendInsightElement) trendInsightElement.textContent = strategicInsights?.trendInsight || '-';
    if (segmentationElement) segmentationElement.textContent = strategicInsights?.segmentationInsight || '-';
    if (qrisElement) qrisElement.textContent = strategicInsights?.qrisInsight || '-';
    if (simulationElement) simulationElement.textContent = strategicInsights?.simulationInsight || '-';
    if (healthElement) healthElement.textContent = strategicInsights?.healthScoreText || '-';
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

    const entries = Object.entries(channelTransactionData || {})
        .filter(([, count]) => count > 0)
        .sort((a, b) => b[1] - a[1]);

    if (entries.length === 0) {
        container.innerHTML = '<span class="breakdown-item">Tidak ada data kanal</span>';
        return;
    }

    container.innerHTML = entries
        .map(([channelName, count]) => `<span class="breakdown-item">${channelName}: ${count}</span>`)
        .join('');
}

function updateCreditBreakdown(channelData) {
    const container = document.getElementById('creditBreakdown');
    if (!container) return;

    const entries = Object.entries(channelData || {})
        .filter(([, amount]) => amount > 0)
        .sort((a, b) => b[1] - a[1]);

    if (entries.length === 0) {
        container.innerHTML = '<span class="breakdown-item">Tidak ada data pemasukan kanal</span>';
        return;
    }

    container.innerHTML = entries
        .map(([channelName, amount]) => `<span class="breakdown-item">${channelName}: ${formatCurrency(amount)}</span>`)
        .join('');
}

async function exportDashboardPDF() {
    const dashboard = document.getElementById('dashboard');
    if (!dashboard || dashboard.style.display === 'none') {
        showStatus('⚠️ Muat data terlebih dahulu sebelum export PDF.', 'warning');
        return;
    }

    if (!window.html2canvas || !window.jspdf) {
        showStatus('❌ Library export PDF belum siap. Coba refresh halaman.', 'error');
        return;
    }

    const exportTarget = document.querySelector('.container');
    if (!exportTarget) {
        showStatus('❌ Area dashboard tidak ditemukan untuk diexport.', 'error');
        return;
    }

    try {
        showStatus('⏳ Sedang membuat PDF, mohon tunggu...', 'warning');
        window.scrollTo({ top: 0, behavior: 'auto' });

        const canvas = await window.html2canvas(exportTarget, {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: -window.scrollY,
            windowWidth: document.documentElement.scrollWidth,
            windowHeight: document.documentElement.scrollHeight
        });

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const imageWidth = pageWidth;
        const imageHeight = (canvas.height * imageWidth) / canvas.width;

        const imageData = canvas.toDataURL('image/jpeg', 0.95);
        let remainingHeight = imageHeight;
        let positionY = 0;

        pdf.addImage(imageData, 'JPEG', 0, positionY, imageWidth, imageHeight);
        remainingHeight -= pageHeight;

        while (remainingHeight > 0) {
            positionY = remainingHeight - imageHeight;
            pdf.addPage();
            pdf.addImage(imageData, 'JPEG', 0, positionY, imageWidth, imageHeight);
            remainingHeight -= pageHeight;
        }

        const periodText = (document.getElementById('period')?.textContent || 'periode').trim();
        const safePeriod = periodText.replace(/[^a-zA-Z0-9-_]/g, '_') || 'periode';
        const dateStr = new Date().toISOString().slice(0, 10);
        pdf.save(`laporan-pipakatan-${safePeriod}-${dateStr}.pdf`);

        showStatus('✅ Export PDF berhasil!', 'success');
    } catch (error) {
        console.error('Error exporting PDF:', error);
        showStatus('❌ Gagal export PDF: ' + error.message, 'error');
    }
}

function cleanHeaderKey(key) {
    return String(key || '').replace(/^\uFEFF/, '').trim();
}

function normalizeHeaderLookupKey(key) {
    return cleanHeaderKey(key).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeCsvRow(row) {
    const normalized = {};

    Object.entries(row || {}).forEach(([key, value]) => {
        const normalizedKey = cleanHeaderKey(key);
        normalized[normalizedKey] = typeof value === 'string' ? value.trim() : value;
    });

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

    const combinedDateTime = findByAliases(['Tanggal Waktu', 'TANGGAL- WAKTU', 'Tanggal-Waktu', 'Datetime', 'DateTime']);
    if (combinedDateTime) {
        const [datePart, timePart] = String(combinedDateTime).split(/\s+/);
        if (!normalized.Tanggal && datePart) normalized.Tanggal = datePart;
        if (!normalized.Jam && timePart) normalized.Jam = timePart;
        if (!normalized.PostDate) normalized.PostDate = String(combinedDateTime);
    }

    if (!normalized.AccountNo) normalized.AccountNo = findByAliases(['AccountNo', 'Account No', 'NoRekening', 'Rekening', 'Nomor Arsip', 'Kode Billing', 'NOP'], '');
    if (!normalized.PostDate) normalized.PostDate = findByAliases(['PostDate', 'PostingDate', 'TransactionDate', 'Tanggal Waktu', 'DateTime'], '');
    if (!normalized.Tanggal) normalized.Tanggal = findByAliases(['Tanggal', 'Date', 'Tgl'], '');
    if (!normalized.Jam) normalized.Jam = findByAliases(['Jam', 'Time', 'Waktu'], '');
    if (!normalized['Credit Amount']) normalized['Credit Amount'] = findByAliases(['Credit Amount', 'Credit', 'Jumlah Kredit', 'Kredit', 'Total'], '0');
    if (!normalized['Debit Amount']) normalized['Debit Amount'] = findByAliases(['Debit Amount', 'Debit', 'Jumlah Debit', 'Debet'], '0');
    if (!normalized['Close Balance']) normalized['Close Balance'] = findByAliases(['Close Balance', 'Balance', 'Saldo', 'Saldo Akhir'], '0');
    if (!normalized.Source) normalized.Source = findByAliases(['Source', 'Sumber'], '');
    if (!normalized.Bulan) normalized.Bulan = findByAliases(['Bulan', 'Month', 'Tahun'], '');
    if (!normalized['Jenis Kanal']) normalized['Jenis Kanal'] = findByAliases(['Jenis Kanal', 'Kanal', 'Channel', 'Nama'], '');
    if (!normalized['Jenis Penerimaan']) normalized['Jenis Penerimaan'] = findByAliases(['Jenis Penerimaan', 'Penerimaan', 'RevenueType', 'Jenis Pajak'], '');

    const pokokAmount = parseAmount(findByAliases(['Pokok'], '0'));
    const dendaAmount = parseAmount(findByAliases(['Denda'], '0'));
    if (parseAmount(normalized['Credit Amount']) === 0 && (pokokAmount > 0 || dendaAmount > 0)) {
        normalized['Credit Amount'] = `${(pokokAmount + dendaAmount).toFixed(2).replace('.', ',')}`;
    }

    return normalized;
}

function isValidTransactionRow(row) {
    const hasDate = Boolean(String(row.PostDate || row.Tanggal || '').trim());
    const hasAccount = Boolean(String(row.AccountNo || '').trim());
    const hasAmount = parseAmount(row['Credit Amount']) > 0 || parseAmount(row['Debit Amount']) > 0 || parseAmount(row['Close Balance']) > 0;

    return hasDate && (hasAccount || hasAmount);
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

    const parsePromises = files.map(file => new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: 'greedy',
            transformHeader: function(header) {
                return cleanHeaderKey(header);
            },
            complete: function(results) {
                const normalizedRows = (results.data || []).map(normalizeCsvRow);
                resolve({ fileName: file.name, data: normalizedRows });
            },
            error: function(error) {
                reject(new Error(file.name + ' - ' + error.message));
            }
        });
    }));

    Promise.all(parsePromises)
        .then(parsedResults => {
            const mergedData = parsedResults.flatMap(result =>
                result.data.map(row => ({
                    ...row,
                    __sourceFile: result.fileName
                }))
            );
            processData(mergedData);

            const fileNames = files.map(file => file.name).join(', ');
            showStatus('✅ Data berhasil dimuat dari ' + files.length + ' file: ' + fileNames, 'success');
        })
        .catch(error => {
            document.getElementById('loading').style.display = 'none';
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
            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: 'greedy',
                transformHeader: function(header) {
                    return cleanHeaderKey(header);
                },
                complete: function(results) {
                    const rowsWithSource = (results.data || []).map(row => ({
                        ...normalizeCsvRow(row),
                        __sourceFile: 'data/rekening-bank.csv'
                    }));

                    const validRows = rowsWithSource.filter(isValidTransactionRow);

                    if (validRows.length === 0) {
                        document.getElementById('loading').style.display = 'none';
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
                },
                error: function(error) {
                    document.getElementById('loading').style.display = 'none';
                    showStatus('❌ Error parsing CSV: ' + error.message, 'error');
                }
            });
        })
        .catch(error => {
            document.getElementById('loading').style.display = 'none';
            showStatus('❌ Error: ' + error.message + '. Pastikan file data/rekening-bank.csv ada!', 'error');
        });
}

// Process and display data
function processData(data) {
    // Filter empty rows
    data = data
        .map(normalizeCsvRow)
        .filter(isValidTransactionRow)
        .sort((a, b) => getRowTimestamp(a) - getRowTimestamp(b));

    latestProcessedRows = data;
    
    if (data.length === 0) {
        document.getElementById('loading').style.display = 'none';
        showStatus('❌ File CSV kosong atau format tidak sesuai!', 'error');
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
    let runningBalance = 0;
    
    data.forEach(row => {
        const credit = parseAmount(row['Credit Amount']);
        const debit = parseAmount(row['Debit Amount']);
        const balance = parseAmount(row['Close Balance']);
        const date = row['Tanggal'] || row['PostDate'];
        const rowTimestamp = getRowTimestamp(row);
        const channel = row['Jenis Kanal'] || 'Unknown';
        const revenueType = row['Jenis Penerimaan'] || 'Lainnya';
        const time = row['Jam'];
        const sourceFile = row.__sourceFile || row.Source || 'Tanpa Sumber';
        
        if (!accountNo && row['AccountNo']) {
            accountNo = row['AccountNo'];
        }
        
        totalCredit += credit;
        
        // Channel data
        if (channel && credit > 0) {
            channelData[channel] = (channelData[channel] || 0) + credit;
        }

        if (channel) {
            channelTransactionData[channel] = (channelTransactionData[channel] || 0) + 1;
        }
        
        // Revenue type data
        if (revenueType && credit > 0) {
            revenueTypeData[revenueType] = (revenueTypeData[revenueType] || 0) + credit;
        }
        
        // Hourly data
        if (time) {
            const hour = parseInt(time.split(':')[0]);
            if (!isNaN(hour) && hour >= 0 && hour < 24) {
                hourlyData[hour]++;
            }
        }

        // Source file data
        if (credit > 0) {
            fileSourceData[sourceFile] = (fileSourceData[sourceFile] || 0) + credit;
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
            monthData.totalCredit += credit;

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
    const monthlySeries = Array.from(monthlyBalanceMap.values()).sort((a, b) => a.key.localeCompare(b.key));
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
        hasMonthly: monthlySeries.length >= 2
    };
    
    // Update header info
    document.getElementById('accountNo').textContent = accountNo || '-';
    const latestPeriod = data[data.length - 1]?.['Bulan'] || '-';
    document.getElementById('period').textContent = latestPeriod;
    
    // Update statistics
    document.getElementById('totalCredit').textContent = formatCurrency(totalCredit);
    document.getElementById('totalTransactions').textContent = transactionCount;
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
    createCharts(balanceChartData, channelData, revenueTypeData, fileSourceData, colors, dashboardStatus, strategicInsights);
    
    // Show dashboard
    document.getElementById('loading').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';
    switchDashboardMenu(activeDashboardMenu);
}

// Create all charts
function createCharts(balanceChartData, channelData, revenueTypeData, fileSourceData, colors, dashboardStatus = 'neutral', strategicInsights = null) {
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
        family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
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
                family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
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
                    yAxisID: 'y'
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
                            if (item.text === 'Pendapatan Transaksi Bulanan') {
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
    switchDashboardMenu(activeDashboardMenu);
    if (checkAuth()) {
        loadDefaultCSV();
    }
});