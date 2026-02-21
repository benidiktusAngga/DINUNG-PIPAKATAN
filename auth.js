const USERS_STORAGE_KEY = 'appUsers';

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
        if (!Array.isArray(parsedUsers) || parsedUsers.length === 0) {
            localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(defaultUsers));
            return [...defaultUsers];
        }

        return parsedUsers;
    } catch (error) {
        localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(defaultUsers));
        return [...defaultUsers];
    }
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

// Handle login
function handleLogin(event) {
    event.preventDefault();
    
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value;
    
    // Validasi input
    if (!username || !password) {
        showError('❌ Username dan password harus diisi!');
        return false;
    }
    
    const users = getUsersFromStorage();

    // Cari user di database
    const user = users.find(u => u.username === username && u.password === password);
    
    if (user) {
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
        
        // Redirect ke dashboard
        window.location.href = 'index.html';
    } else {
        // Login gagal
        showError('❌ Username atau password salah!');
        
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
window.addEventListener('DOMContentLoaded', checkExistingSession);

// Handle enter key
document.addEventListener('keypress', function(event) {
    if (event.key === 'Enter') {
        const form = document.getElementById('loginForm');
        if (form) {
            handleLogin(event);
        }
    }
});