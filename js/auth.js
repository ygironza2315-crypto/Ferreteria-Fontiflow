import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
// 🚀 Importamos updateProfile para guardar el nombre real del cliente
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut, updateProfile } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";

// 🚨 REEMPLAZA ESTO CON TUS LLAVES REALES DE LA CONSOLA DE FIREBASE
const firebaseConfig = {
    apiKey: "AIzaSyDzNhMwca_skQbzXswbLPq2A0cETfxdi90",
    authDomain: "ferreteria-fontiflow-auth.firebaseapp.com",
    projectId: "ferreteria-fontiflow-auth",
    storageBucket: "ferreteria-fontiflow-auth.firebasestorage.app",
    messagingSenderId: "405773965005",
    appId: "1:405773965005:web:1335205a830d6cbd493b84"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// Controladores globales de apertura y cierre por clicks
document.addEventListener('click', (e) => {
    if (e.target.matches('.btn-abrir-auth') || e.target.closest('.btn-abrir-auth')) {
        const modal = document.getElementById('authModalCustom');
        if (modal) modal.style.display = 'flex';
    }
    if (e.target.matches('.btn-cerrar-auth') || e.target.closest('.btn-cerrar-auth')) {
        const modal = document.getElementById('authModalCustom');
        if (modal) modal.style.display = 'none';
    }
});

// Escuchador del Estado de la Sesión en Tiempo Real (Actualizado a ASYNC para soportar updateProfile)
onAuthStateChanged(auth, async (user) => {
    if (user) {
        document.body.classList.add('user-logged-in');
        
        const txtNombre = document.getElementById('txtNombreUsuario');
        
        // ==========================================================================
        // 🧠 CONTROL DE USUARIOS ANTIGUOS: Si no tienen nombre, se los pedimos una vez
        // ==========================================================================
        if (!user.displayName) {
            // Verificamos si ya le preguntamos en esta sesión de navegación para no ser molestos
            if (!sessionStorage.getItem('intentoNombre')) {
                sessionStorage.setItem('intentoNombre', 'true'); // Evita bucles infinitos si cancela
                
                const nombreAsignado = prompt("¡Bienvenido de nuevo! Para personalizar tu cuenta, por favor ingresa tu Nombre Completo o el de tu Empresa:");
                
                if (nombreAsignado && nombreAsignado.trim() !== "") {
                    try {
                        // Guardamos el nombre real en la nube de Firebase
                        await updateProfile(user, { displayName: nombreAsignado.trim() });
                        window.location.reload(); // Refrescamos para aplicar el cambio visual
                        return;
                    } catch (err) {
                        console.error("Error al actualizar usuario antiguo:", err);
                    }
                }
            }
            
            // Si el cliente cancela el cartel, mostramos el correo provisionalmente en esta visita
            if (txtNombre) {
                txtNombre.innerText = user.email.split('@')[0].toUpperCase();
            }
        } else {
            // Si el usuario ya tiene su nombre configurado (usuarios nuevos o ya actualizados)
            if (txtNombre) {
                txtNombre.innerText = user.displayName.toUpperCase();
            }
        }
        // ==========================================================================
        
        const modal = document.getElementById('authModalCustom');
        if (modal) modal.style.display = 'none';
    } else {
        document.body.classList.remove('user-logged-in');
    }
});

// Procesamiento de formularios
document.addEventListener('DOMContentLoaded', () => {
    const seccionLogin = document.getElementById('modalSeccionLogin');
    const seccionRegistro = document.getElementById('modalSeccionRegistro');
    const errorDiv = document.getElementById('modalErrorMessage');

    document.getElementById('btnModalIrARegistro')?.addEventListener('click', () => {
        seccionLogin?.classList.add('d-none');
        seccionRegistro?.classList.remove('d-none');
        errorDiv?.classList.add('d-none');
    });

    document.getElementById('btnModalIrALogin')?.addEventListener('click', () => {
        seccionRegistro?.classList.add('d-none');
        seccionLogin?.classList.remove('d-none');
        errorDiv?.classList.add('d-none');
    });

    // Evento 1: Iniciar Sesión
    document.getElementById('modalLoginForm')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (errorDiv) errorDiv.classList.add('d-none');
        const email = document.getElementById('modalEmail').value;
        const password = document.getElementById('modalPassword').value;

        try {
            await signInWithEmailAndPassword(auth, email, password);
        } catch (error) {
            if (errorDiv) errorDiv.innerText = "Correo o contraseña incorrectos.";
            errorDiv?.classList.remove('d-none');
        }
    });

    // Evento 2: Registro con Nombre Personalizado
    document.getElementById('modalRegisterForm')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (errorDiv) errorDiv.classList.add('d-none');
        
        // Captura de datos del formulario
        const name = document.getElementById('modalRegName').value.trim();
        const email = document.getElementById('modalRegEmail').value.trim();
        const password = document.getElementById('modalRegPassword').value;
        const passwordConfirm = document.getElementById('modalRegPasswordConfirm').value;

        if (password !== passwordConfirm) {
            if (errorDiv) errorDiv.innerText = "Las contraseñas no coinciden.";
            errorDiv?.classList.remove('d-none');
            return;
        }
        if (password.length < 6) {
            if (errorDiv) errorDiv.innerText = "La contraseña debe tener mínimo 6 caracteres.";
            errorDiv?.classList.remove('d-none');
            return;
        }

        try {
            // 1. Creamos la cuenta con correo y contraseña en los servidores de Google
            await createUserWithEmailAndPassword(auth, email, password);
            
            // 2. Inyectamos de inmediato el nombre del cliente en su perfil recién creado
            await updateProfile(auth.currentUser, { displayName: name });
            
            // 3. Forzamos un refresco rápido de la página para que el menú asimile el nombre nuevo
            window.location.reload();
            
        } catch (error) {
            if (errorDiv) {
                if (error.code === 'auth/email-already-in-use') {
                    errorDiv.innerText = "Este correo ya está registrado.";
                } else {
                    errorDiv.innerText = "Error al crear la cuenta. Inténtalo de nuevo.";
                }
            }
            errorDiv?.classList.remove('d-none');
        }
    });

    // Evento 3: Cerrar Sesión
    document.getElementById('btnCerrarSesion')?.addEventListener('click', () => {
        signOut(auth).then(() => { window.location.reload(); });
    });
});