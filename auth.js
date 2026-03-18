const USERS_STORAGE_KEY = 'appUsers';
const LOGIN_GUARD_STORAGE_KEY = 'loginGuardState';
const AUTH_SECURITY = {
    USERNAME_REGEX: /^[a-zA-Z0-9._-]{3,32}$/,
    MIN_PASSWORD_LENGTH: 6,
    MAX_PASSWORD_LENGTH: 64,
    MAX_FAILED_ATTEMPTS: 5,
    LOCKOUT_DURATION_MS: 5 * 60 * 1000
};

// Database users default (dalam aplikasi production, ini harus di backend/database)
const defaultUsers = [
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

function getUsersFromStorage() {
    try {
        const rawUsers = localStorage.getItem(USERS_STORAGE_KEY);
        if (!rawUsers) {
            localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(defaultUsers));
            return [...defaultUsers];
        }

        const parsedUsers = JSON.parse(rawUsers);
        const sanitizedUsers = sanitizeUsers(parsedUsers);

        if (sanitizedUsers.length === 0) {
            localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(defaultUsers));
            return [...defaultUsers];
        }

        return sanitizedUsers;
    } catch (error) {
        localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(defaultUsers));
        return [...defaultUsers];
    }
}

function normalizeInput(value) {
    return String(value || '').normalize('NFKC').trim();
}

function normalizePasswordInput(value) {
    return String(value || '').normalize('NFKC');
}

function hasControlChars(value) {
    return /[\u0000-\u001F\u007F]/.test(value);
}

function sanitizeUsers(rawUsers) {
    if (!Array.isArray(rawUsers)) return [];

    return rawUsers
        .filter((user) => user && typeof user === 'object')
        .map((user) => {
            const username = normalizeInput(user.username);
            const password = typeof user.password === 'string' ? user.password : '';
            const role = ['admin', 'user', 'guest'].includes(user.role) ? user.role : 'guest';
            const fullName = normalizeInput(user.fullName || username);
            const photoUrl = normalizeInput(user.photoUrl || '');
            const permissions = Array.isArray(user.permissions)
                ? user.permissions.filter((permission) => ['view', 'upload', 'download', 'manage'].includes(permission))
                : ['view'];

            if (!AUTH_SECURITY.USERNAME_REGEX.test(username)) return null;
            if (password.length < AUTH_SECURITY.MIN_PASSWORD_LENGTH || password.length > AUTH_SECURITY.MAX_PASSWORD_LENGTH) return null;
            if (hasControlChars(password)) return null;

            return {
                username,
                password,
                role,
                fullName: fullName || username,
                photoUrl,
                permissions: permissions.length > 0 ? permissions : ['view']
            };
        })
        .filter(Boolean);
}

function timingSafeEqual(left, right) {
    if (typeof left !== 'string' || typeof right !== 'string') return false;

    const normalizedLeft = String(left);
    const normalizedRight = String(right);
    const maxLength = Math.max(normalizedLeft.length, normalizedRight.length);
    let diff = normalizedLeft.length ^ normalizedRight.length;

    for (let i = 0; i < maxLength; i++) {
        const leftCode = i < normalizedLeft.length ? normalizedLeft.charCodeAt(i) : 0;
        const rightCode = i < normalizedRight.length ? normalizedRight.charCodeAt(i) : 0;
        diff |= leftCode ^ rightCode;
    }

    return diff === 0;
}

function getLoginGuardState() {
    try {
        const raw = localStorage.getItem(LOGIN_GUARD_STORAGE_KEY);
        if (!raw) return { failedAttempts: 0, lockoutUntil: 0 };

        const parsed = JSON.parse(raw);
        return {
            failedAttempts: Number.isFinite(parsed.failedAttempts) ? Math.max(0, parsed.failedAttempts) : 0,
            lockoutUntil: Number.isFinite(parsed.lockoutUntil) ? Math.max(0, parsed.lockoutUntil) : 0
        };
    } catch (error) {
        return { failedAttempts: 0, lockoutUntil: 0 };
    }
}

function saveLoginGuardState(state) {
    localStorage.setItem(LOGIN_GUARD_STORAGE_KEY, JSON.stringify({
        failedAttempts: state.failedAttempts,
        lockoutUntil: state.lockoutUntil
    }));
}

function getLockoutRemainingMs() {
    const state = getLoginGuardState();
    const remaining = state.lockoutUntil - Date.now();
    return remaining > 0 ? remaining : 0;
}

function registerFailedAttempt() {
    const state = getLoginGuardState();
    const nextAttempts = state.failedAttempts + 1;
    const shouldLock = nextAttempts >= AUTH_SECURITY.MAX_FAILED_ATTEMPTS;

    saveLoginGuardState({
        failedAttempts: shouldLock ? 0 : nextAttempts,
        lockoutUntil: shouldLock ? Date.now() + AUTH_SECURITY.LOCKOUT_DURATION_MS : 0
    });
}

function resetLoginGuard() {
    saveLoginGuardState({ failedAttempts: 0, lockoutUntil: 0 });
}

// Toggle password visibility
function togglePassword() {
    const passwordInput = document.getElementById('password');
    const toggleIcon = document.querySelector('.toggle-icon');
    
    if (passwordInput.type === 'password') {
        passwordInput.type = 'text';
        toggleIcon.textContent = '🙈';
    } else {
        passwordInput.type = 'password';
        toggleIcon.textContent = '👁️';
    }
}

// Fill login form (untuk demo)
function fillLogin(username, password) {
    document.getElementById('username').value = username;
    document.getElementById('password').value = password;
    document.getElementById('username').focus();
}

// Show error message
function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    
    setTimeout(() => {
        errorDiv.style.display = 'none';
    }, 5000);
}

function showSuccessPopup() {
    const popup = document.getElementById('successPopup');
    if (!popup) return;
    popup.classList.add('show');
    popup.setAttribute('aria-hidden', 'false');
}

function hidePopupById(popupId) {
    const popup = document.getElementById(popupId);
    if (!popup) return;
    popup.classList.remove('show');
    popup.setAttribute('aria-hidden', 'true');
}

function showFailedPopup() {
    const popup = document.getElementById('failedPopup');
    if (!popup) return;

    popup.classList.add('show');
    popup.setAttribute('aria-hidden', 'false');

    setTimeout(() => {
        hidePopupById('failedPopup');
    }, 1600);
}

function setupPopupOverlayDismiss() {
    const popupIds = ['successPopup', 'failedPopup'];
    popupIds.forEach((popupId) => {
        const popup = document.getElementById(popupId);
        if (!popup) return;

        popup.addEventListener('click', function(event) {
            if (event.target === popup) {
                hidePopupById(popupId);
            }
        });
    });
}

// Handle login
function handleLogin(event) {
    event.preventDefault();

    const username = normalizeInput(document.getElementById('username').value);
    const password = normalizePasswordInput(document.getElementById('password').value);

    const lockoutRemainingMs = getLockoutRemainingMs();
    if (lockoutRemainingMs > 0) {
        const lockoutRemainingMinutes = Math.ceil(lockoutRemainingMs / 60000);
        showError(`❌ Terlalu banyak percobaan login. Coba lagi dalam ${lockoutRemainingMinutes} menit.`);
        showFailedPopup();
        return false;
    }
    
    // Validasi input
    if (!username || !password) {
        showError('❌ Username dan password harus diisi!');
        return false;
    }

    if (!AUTH_SECURITY.USERNAME_REGEX.test(username)) {
        showError('❌ Format username tidak valid. Gunakan 3-32 karakter (huruf, angka, titik, underscore, strip).');
        showFailedPopup();
        return false;
    }

    if (password.length < AUTH_SECURITY.MIN_PASSWORD_LENGTH || password.length > AUTH_SECURITY.MAX_PASSWORD_LENGTH || hasControlChars(password)) {
        showError('❌ Format password tidak valid. Gunakan 6-64 karakter tanpa karakter kontrol.');
        showFailedPopup();
        return false;
    }
    
    const users = getUsersFromStorage();

    // Cari user di database
    const user = users.find((storedUser) => timingSafeEqual(storedUser.username, username));
    const passwordValid = user ? timingSafeEqual(user.password, password) : false;
    
    if (user && passwordValid) {
        // Login berhasil
        const sessionData = {
            username: user.username,
            role: user.role,
            fullName: user.fullName,
            photoUrl: user.photoUrl || '',
            permissions: user.permissions,
            loginTime: new Date().toISOString()
        };
        
        // Simpan session ke localStorage
        localStorage.setItem('userSession', JSON.stringify(sessionData));

        resetLoginGuard();

        showSuccessPopup();
        
        // Redirect ke dashboard
        setTimeout(() => {
            window.location.href = 'index.html';
        }, 1200);
    } else {
        // Login gagal
        registerFailedAttempt();
        showFailedPopup();
        showError('❌ Username atau password tidak valid!');
        
        // Shake animation
        const form = document.getElementById('loginForm');
        form.style.animation = 'none';
        setTimeout(() => {
            form.style.animation = 'shake 0.5s ease';
        }, 10);
    }
    
    return false;
}

// Check if already logged in
function checkExistingSession() {
    const session = localStorage.getItem('userSession');
    if (session) {
        // Sudah login, redirect ke dashboard
        window.location.href = 'index.html';
    }
}

// Jalankan saat halaman dimuat
window.addEventListener('DOMContentLoaded', function() {
    checkExistingSession();
    setupPopupOverlayDismiss();
});

// Handle enter key
document.addEventListener('keypress', function(event) {
    if (event.key === 'Enter') {
        const form = document.getElementById('loginForm');
        if (form) {
            handleLogin(event);
        }
    }
});
