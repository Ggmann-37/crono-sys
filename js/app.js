/**
 * Crono-SYS - Orquestador Principal de la Aplicación
 * Conecta ExcelParser (SheetJS), SmartScheduler, CalendarView y SupabaseService
 * Incluye Fase Previa (Landing estilo Outlook) y Panel Login/Signup dividido (estilo GitHub)
 */

const App = {
  state: {
    scheduleId: null,
    fileName: null,
    baseBlocks: [],    // Bloques del Excel (Horario Base)
    taskBlocks: [],    // Bloques asignados a tareas/estudio
    breakBlocks: [],   // Bloques de descanso calculados
    tasks: [],         // Tareas registradas
    selectedBlock: null,
    currentView: 'landing', // 'landing', 'auth', 'app'
    authMode: 'signup',     // 'signup' o 'login'
    isPreviewMode: false,   // Si es true, no guarda en la nube
    pendingConfirmationEmail: ''
  },

  turnstile: {
    siteKey: '0x4AAAAAAGKBPrSRb8JwpJkV',
    widgetId: null,
    token: null,
    readyPromise: null
  },

  /**
   * Inicialización de la aplicación
   */
  async init() {
    console.log('Iniciando Crono-SYS...');

    // 1. Inicializar Supabase
    SupabaseService.init();
    this.initTurnstile();

    // 2. Inicializar Calendario estilo Outlook/Teams
    const calendarContainer = document.getElementById('calendarMainContainer');
    CalendarView.init(calendarContainer);

    // 3. Vincular Eventos de la Interfaz
    this.bindEvents();

    // 4. Configurar Detector Responsivo en Segundo Plano (Móvil / Ordenador)
    this.initResponsiveDetector();

    // 5. Configurar Estado de Autenticación
    this.setupAuth();

    // 6. Comprobar sesión existente
    const session = await SupabaseService.getSession();
    const confirmedEmail = session?.user?.email_confirmed_at || session?.user?.confirmed_at;
    if (session && session.user && confirmedEmail) {
      this.clearPendingEmailConfirmation();
      // Si el usuario ya está autenticado, va directamente a la app
      await this.loadInitialData();
      this.showView('app', { preview: false });
    } else if (this.restorePendingEmailConfirmation()) {
      this.showView('auth', { mode: 'signup' });
    } else {
      // Si no, muestra la fase previa (Landing Hero)
      this.showView('landing');
    }

    // 7. Actualizar estadísticas iniciales
    this.updateStats();
  },

  /**
   * Detector responsivo en segundo plano que identifica si es móvil u ordenador
   * y adapta los estilos, calendario y controles dinámicamente
   */
  initResponsiveDetector() {
    const updateDeviceClass = () => {
      const isMobile = window.matchMedia('(max-width: 768px)').matches || 
                       ('ontouchstart' in window && window.innerWidth <= 1024);
      if (isMobile) {
        document.body.classList.add('is-mobile-device');
        document.body.classList.remove('is-desktop-device');
      } else {
        document.body.classList.add('is-desktop-device');
        document.body.classList.remove('is-mobile-device');
      }
    };

    updateDeviceClass();
    window.addEventListener('resize', updateDeviceClass, { passive: true });
    window.addEventListener('orientationchange', updateDeviceClass, { passive: true });
  },

  // ==========================================
  // GESTIÓN DE VISTAS (LANDING, AUTH, APP)
  // ==========================================

  /**
   * Cambia de pantalla entre Landing (fase previa), Auth (GitHub style) y App
   */
  showView(viewName, options = {}) {
    if (this.state.pendingConfirmationEmail && viewName !== 'auth') {
      this.showView('auth', { mode: 'signup' });
      return;
    }

    const viewLanding = document.getElementById('viewLanding');
    const viewAuth = document.getElementById('viewAuth');
    const viewApp = document.getElementById('viewApp');

    if (viewLanding) viewLanding.classList.add('view-hidden');
    if (viewAuth) viewAuth.classList.add('view-hidden');
    if (viewApp) viewApp.classList.add('view-hidden');

    this.state.currentView = viewName;

    if (viewName === 'landing') {
      if (viewLanding) viewLanding.classList.remove('view-hidden');
      window.scrollTo(0, 0);
    } else if (viewName === 'auth') {
      if (viewAuth) viewAuth.classList.remove('view-hidden');
      this.setAuthMode(options.mode || 'signup');
      window.scrollTo(0, 0);
      const emailInput = document.getElementById('authEmailInput');
      if (emailInput) setTimeout(() => emailInput.focus(), 100);
    } else if (viewName === 'app') {
      if (viewApp) viewApp.classList.remove('view-hidden');
      this.state.isPreviewMode = !!options.preview;

      const previewBanner = document.getElementById('previewModeBanner');
      const statSecurity = document.getElementById('statSecurityText');
      const authText = document.getElementById('authStatusText');

      if (this.state.isPreviewMode) {
        if (previewBanner) previewBanner.style.display = 'flex';
        if (statSecurity) statSecurity.textContent = 'Modo Vista Previa';
        if (authText) authText.textContent = 'Vista Previa';
      } else {
        if (previewBanner) previewBanner.style.display = 'none';
        if (statSecurity) {
          statSecurity.textContent = SupabaseService.currentUser ? 'Ha iniciado sesión' : 'No ha iniciado sesión';
        }
        if (authText) {
          authText.textContent = SupabaseService.currentUser ? SupabaseService.currentUser.email.split('@')[0] : 'Modo Local';
        }
      }

      // Re-dibujar el calendario para recalcular dimensiones correctamente
      setTimeout(() => {
        this.updateCalendar();
        CalendarView.setupAutoScroll();
      }, 50);

      window.scrollTo(0, 0);
    }
  },

  /**
   * Alterna entre modo 'signup' y 'login' en la pantalla dividida (Captura 2)
   */
  setAuthMode(mode) {
    this.state.authMode = mode;

    const leftHeadline = document.getElementById('authLeftHeadline');
    const cardHeadline = document.getElementById('authCardHeadline');
    const groupConfirm = document.getElementById('groupConfirmPassword');
    const inputConfirm = document.getElementById('authConfirmPasswordInput');
    const passwordHelper = document.getElementById('authPasswordHelper');
    const submitBtn = document.getElementById('btnAuthFormSubmit');
    const googleButtonText = document.getElementById('authGoogleButtonText');
    const githubButtonText = document.getElementById('authGitHubButtonText');
    const promptText = document.getElementById('authSwitchPromptText');
    const actionLink = document.getElementById('authSwitchActionLink');
    const mismatchError = document.getElementById('passwordMismatchError');
    const confirmationPanel = document.getElementById('authEmailConfirmation');
    const authForm = document.getElementById('authMainForm');
    const authBackLink = document.getElementById('btnAuthBackToLanding');

    if (this.state.pendingConfirmationEmail) {
      this.showEmailConfirmation(this.state.pendingConfirmationEmail);
      return;
    }

    if (mismatchError) mismatchError.style.display = 'none';
    if (confirmationPanel) confirmationPanel.classList.add('view-hidden');
    if (authForm) authForm.classList.remove('view-hidden');
    const switchBar = document.querySelector('.auth-top-switch-bar');
    if (switchBar) switchBar.classList.remove('view-hidden');
    if (authBackLink) authBackLink.classList.remove('view-hidden');
    this.resetTurnstile();

    if (mode === 'signup') {
      if (leftHeadline) leftHeadline.textContent = 'Crea tu cuenta gratuita';
      if (cardHeadline) cardHeadline.textContent = 'Crear cuenta en Crono-SYS';
      if (groupConfirm) groupConfirm.style.display = 'block';
      if (inputConfirm) inputConfirm.setAttribute('required', 'required');
      if (passwordHelper) passwordHelper.textContent = 'La contraseña debe tener al menos 6 caracteres.';
      if (submitBtn) submitBtn.textContent = 'Crear Cuenta';
      if (googleButtonText) googleButtonText.textContent = 'Registrarse con Google';
      if (githubButtonText) githubButtonText.textContent = 'Registrarse con GitHub';
      if (promptText) promptText.textContent = '¿Ya tienes una cuenta?';
      if (actionLink) actionLink.innerHTML = 'Iniciar sesión &rarr;';
    } else {
      if (leftHeadline) leftHeadline.textContent = 'Bienvenido de nuevo a Crono-SYS';
      if (cardHeadline) cardHeadline.textContent = 'Iniciar sesión en Crono-SYS';
      if (groupConfirm) groupConfirm.style.display = 'none';
      if (inputConfirm) inputConfirm.removeAttribute('required');
      if (passwordHelper) passwordHelper.textContent = '';
      if (submitBtn) submitBtn.textContent = 'Iniciar Sesión';
      if (googleButtonText) googleButtonText.textContent = 'Iniciar sesión con Google';
      if (githubButtonText) githubButtonText.textContent = 'Iniciar sesión con GitHub';
      if (promptText) promptText.textContent = '¿No tienes una cuenta?';
      if (actionLink) actionLink.innerHTML = 'Regístrate gratis &rarr;';
    }
  },

  showEmailConfirmation(email) {
    this.state.pendingConfirmationEmail = email;
    sessionStorage.setItem('cronosys_pending_confirmation_email', email);

    const authForm = document.getElementById('authMainForm');
    const confirmationPanel = document.getElementById('authEmailConfirmation');
    const confirmationEmail = document.getElementById('authConfirmationEmail');
    const switchBar = document.querySelector('.auth-top-switch-bar');
    const cardHeadline = document.getElementById('authCardHeadline');
    const authBackLink = document.getElementById('btnAuthBackToLanding');

    if (authForm) authForm.classList.add('view-hidden');
    if (confirmationPanel) confirmationPanel.classList.remove('view-hidden');
    if (confirmationEmail) confirmationEmail.textContent = email;
    if (switchBar) switchBar.classList.add('view-hidden');
    if (authBackLink) authBackLink.classList.add('view-hidden');
    if (cardHeadline) cardHeadline.textContent = 'Verificación pendiente';

    this.resetTurnstile();
  },

  restorePendingEmailConfirmation() {
    const email = sessionStorage.getItem('cronosys_pending_confirmation_email');
    if (!email) return false;
    this.state.pendingConfirmationEmail = email;
    return true;
  },

  clearPendingEmailConfirmation() {
    this.state.pendingConfirmationEmail = '';
    sessionStorage.removeItem('cronosys_pending_confirmation_email');
  },

  // ==========================================
  // VINCULACIÓN DE EVENTOS
  // ==========================================

  bindEvents() {
    // --- NAVEGACIÓN DESDE LA LANDING (CAPTURA 1) ---
    const btnHeroLogin = document.getElementById('btnHeroLogin');
    const btnNavLogin = document.getElementById('btnNavLogin');
    if (btnHeroLogin) btnHeroLogin.addEventListener('click', () => this.showView('auth', { mode: 'login' }));
    if (btnNavLogin) btnNavLogin.addEventListener('click', () => this.showView('auth', { mode: 'login' }));

    const btnHeroSignup = document.getElementById('btnHeroSignup');
    const btnNavSignup = document.getElementById('btnNavSignup');
    if (btnHeroSignup) btnHeroSignup.addEventListener('click', () => this.showView('auth', { mode: 'signup' }));
    if (btnNavSignup) btnNavSignup.addEventListener('click', () => this.showView('auth', { mode: 'signup' }));

    // "Probar gratis" -> Preview interactiva que NO guarda
    const btnHeroPreview = document.getElementById('btnHeroPreview');
    const btnNavPreview = document.getElementById('btnNavPreview');
    if (btnHeroPreview) btnHeroPreview.addEventListener('click', () => this.showView('app', { preview: true }));
    if (btnNavPreview) btnNavPreview.addEventListener('click', () => this.showView('app', { preview: true }));

    const btnLandingTemplate = document.getElementById('btnLandingDownloadTemplate');
    if (btnLandingTemplate) {
      btnLandingTemplate.addEventListener('click', (e) => {
        e.preventDefault();
        ExcelParser.downloadTemplate();
      });
    }

    const landingLogo = document.getElementById('landingLogoLink');
    if (landingLogo) {
      landingLogo.addEventListener('click', (e) => {
        e.preventDefault();
        this.showView('landing');
      });
    }

    // --- NAVEGACIÓN DESDE EL PANEL DE LOGIN / SIGNUP (CAPTURA 2) ---
    const switchLink = document.getElementById('authSwitchActionLink');
    if (switchLink) {
      switchLink.addEventListener('click', (e) => {
        e.preventDefault();
        this.setAuthMode(this.state.authMode === 'signup' ? 'login' : 'signup');
      });
    }

    const btnAuthBack = document.getElementById('btnAuthBackToLanding');
    if (btnAuthBack) {
      btnAuthBack.addEventListener('click', (e) => {
        e.preventDefault();
        this.showView('landing');
      });
    }

    const authBrandHome = document.getElementById('authBrandHomeLink');
    if (authBrandHome) {
      authBrandHome.addEventListener('click', (e) => {
        e.preventDefault();
        this.showView('landing');
      });
    }

    // Envío del formulario principal de Auth (Login / Signup)
    const authForm = document.getElementById('authMainForm');
    if (authForm) {
      authForm.addEventListener('submit', (e) => this.handleMainAuthSubmit(e));
    }

    const googleButton = document.getElementById('btnAuthGoogle');
    if (googleButton) {
      googleButton.addEventListener('click', async () => {
        if (this.state.pendingConfirmationEmail) {
          this.showToast('Verifica tu correo antes de iniciar sesión', 'error');
          return;
        }

        googleButton.disabled = true;
        try {
          await SupabaseService.signInWithGoogle();
        } catch (err) {
          googleButton.disabled = false;
          this.showToast(err.message, 'error');
        }
      });
    }

    const githubButton = document.getElementById('btnAuthGitHub');
    if (githubButton) {
      githubButton.addEventListener('click', async () => {
        if (this.state.pendingConfirmationEmail) {
          this.showToast('Verifica tu correo antes de iniciar sesión', 'error');
          return;
        }

        githubButton.disabled = true;
        try {
          await SupabaseService.signInWithGitHub();
        } catch (err) {
          githubButton.disabled = false;
          this.showToast(err.message, 'error');
        }
      });
    }

    const resendConfirmation = document.getElementById('btnResendConfirmation');
    if (resendConfirmation) {
      resendConfirmation.addEventListener('click', async () => {
        const email = this.state.pendingConfirmationEmail;
        if (!email) return;

        resendConfirmation.disabled = true;
        resendConfirmation.textContent = 'Enviando...';
        try {
          await SupabaseService.resendConfirmation(email);
          this.showToast('Correo de verificación reenviado', 'success');
        } catch (err) {
          this.showToast(err.message, 'error');
        } finally {
          resendConfirmation.disabled = false;
          resendConfirmation.textContent = 'Reenviar correo';
        }
      });
    }

    // --- NAVEGACIÓN DESDE EL BANNER DE VISTA PREVIA ---
    const btnExitPreview = document.getElementById('btnExitPreviewMode');
    if (btnExitPreview) {
      btnExitPreview.addEventListener('click', () => this.showView('landing'));
    }

    const btnSavePreview = document.getElementById('btnSaveFromPreview');
    if (btnSavePreview) {
      btnSavePreview.addEventListener('click', () => this.showView('auth', { mode: 'signup' }));
    }

    // Retorno a la Portada desde la App
    const btnAppLogoHome = document.getElementById('btnAppHomeLogo');
    if (btnAppLogoHome) {
      btnAppLogoHome.addEventListener('click', (e) => {
        e.preventDefault();
        this.showView('landing');
      });
    }

    // --- ACCIONES DENTRO DEL HORARIO Y TAREAS ---
    const btnDownload = document.getElementById('btnDownloadTemplate');
    if (btnDownload) {
      btnDownload.addEventListener('click', () => ExcelParser.downloadTemplate());
    }

    const btnModalDownload = document.getElementById('btnModalDownloadTemplate');
    if (btnModalDownload) {
      btnModalDownload.addEventListener('click', () => ExcelParser.downloadTemplate());
    }

    const btnOpenUpload = document.getElementById('btnOpenUploadModal');
    if (btnOpenUpload) {
      btnOpenUpload.addEventListener('click', () => this.openModal('uploadModal'));
    }

    const btnOpenTask = document.getElementById('btnOpenTaskModal');
    if (btnOpenTask) {
      btnOpenTask.addEventListener('click', () => this.openModal('taskModal'));
    }

    const btnQuickAdd = document.getElementById('btnQuickAddTask');
    if (btnQuickAdd) {
      btnQuickAdd.addEventListener('click', () => this.openModal('taskModal'));
    }

    const btnEmptyAdd = document.getElementById('btnEmptyStateAdd');
    if (btnEmptyAdd) {
      btnEmptyAdd.addEventListener('click', () => this.openModal('taskModal'));
    }

    // --- PANEL DE USUARIO ESTILO OUTLOOK (CAPTURA 3) ---
    const btnProfileAvatar = document.getElementById('btnMsProfileAvatar');
    const msFlyout = document.getElementById('msAccountFlyout');

    if (btnProfileAvatar && msFlyout) {
      btnProfileAvatar.addEventListener('click', (e) => {
        e.stopPropagation();
        msFlyout.classList.toggle('active');
      });

      // Cerrar al hacer clic fuera
      document.addEventListener('click', (e) => {
        if (!msFlyout.contains(e.target) && e.target !== btnProfileAvatar) {
          msFlyout.classList.remove('active');
        }
      });
    }

    // Cerrar sesión desde el panel Outlook
    const btnFlyoutSignOut = document.getElementById('btnFlyoutSignOut');
    if (btnFlyoutSignOut) {
      btnFlyoutSignOut.addEventListener('click', async () => {
        try {
          await SupabaseService.signOut();
          if (msFlyout) msFlyout.classList.remove('active');
          this.showToast('Sesión cerrada', 'info');
          this.showView('landing');
        } catch (err) {
          this.showToast(err.message, 'error');
        }
      });
    }

    // Iniciar sesión con otra cuenta
    const btnSwitchAccount = document.getElementById('btnFlyoutSwitchAccount');
    if (btnSwitchAccount) {
      btnSwitchAccount.addEventListener('click', async () => {
        try {
          await SupabaseService.signOut();
          if (msFlyout) msFlyout.classList.remove('active');
          this.showView('auth', { mode: 'login' });
        } catch (err) {
          this.showToast(err.message, 'error');
        }
      });
    }

    // Abrir modal de Cambiar Contraseña
    const btnOpenPassword = document.getElementById('btnOpenChangePassword');
    if (btnOpenPassword) {
      btnOpenPassword.addEventListener('click', (e) => {
        e.preventDefault();
        if (msFlyout) msFlyout.classList.remove('active');
        this.openModal('modalChangePassword');
      });
    }

    // Abrir modal de Cambiar Email
    const btnOpenEmail = document.getElementById('btnOpenChangeEmail');
    if (btnOpenEmail) {
      btnOpenEmail.addEventListener('click', (e) => {
        e.preventDefault();
        if (msFlyout) msFlyout.classList.remove('active');
        this.openModal('modalChangeEmail');
      });
    }

    // Abrir modal de Eliminar Cuenta
    const btnOpenDeleteAcc = document.getElementById('btnOpenDeleteAccount');
    if (btnOpenDeleteAcc) {
      btnOpenDeleteAcc.addEventListener('click', (e) => {
        e.preventDefault();
        if (msFlyout) msFlyout.classList.remove('active');
        this.openModal('modalDeleteAccount');
      });
    }

    // Formulario de Cambiar Contraseña
    const formPassword = document.getElementById('formChangePassword');
    if (formPassword) {
      formPassword.addEventListener('submit', async (e) => {
        e.preventDefault();
        const p1 = document.getElementById('inputNewPassword').value;
        const p2 = document.getElementById('inputConfirmNewPassword').value;
        const err = document.getElementById('changePasswordError');

        if (p1 !== p2) {
          if (err) err.style.display = 'block';
          this.showToast('Las contraseñas no coinciden', 'error');
          return;
        }

        if (err) err.style.display = 'none';

        try {
          this.showToast('Actualizando contraseña...', 'info');
          await SupabaseService.updatePassword(p1);
          this.closeModal('modalChangePassword');
          formPassword.reset();
          this.showToast('Contraseña actualizada con éxito', 'success');
        } catch (error) {
          this.showToast(error.message, 'error');
        }
      });
    }

    // Formulario de Cambiar Email
    const formEmail = document.getElementById('formChangeEmail');
    if (formEmail) {
      formEmail.addEventListener('submit', async (e) => {
        e.preventDefault();
        const newEmail = document.getElementById('inputNewEmail').value.trim();
        if (!newEmail) return;

        try {
          this.showToast('Enviando solicitud...', 'info');
          await SupabaseService.updateEmail(newEmail);
          this.closeModal('modalChangeEmail');
          formEmail.reset();
          this.showToast('Correo actualizado. Revisa tu bandeja de entrada si requiere confirmación.', 'success');
        } catch (error) {
          this.showToast(error.message, 'error');
        }
      });
    }

    // Confirmación de Eliminar Cuenta
    const btnConfirmDelete = document.getElementById('btnConfirmDeleteAccount');
    if (btnConfirmDelete) {
      btnConfirmDelete.addEventListener('click', async () => {
        try {
          this.showToast('Eliminando cuenta y datos...', 'info');
          await SupabaseService.deleteAccount();
          this.state.baseBlocks = [];
          this.state.taskBlocks = [];
          this.state.breakBlocks = [];
          this.state.tasks = [];
          this.closeModal('modalDeleteAccount');
          this.showToast('Tu cuenta y datos han sido eliminados', 'info');
          this.showView('landing');
        } catch (error) {
          this.showToast(error.message, 'error');
        }
      });
    }

    document.querySelectorAll('[data-close]').forEach(btn => {
      btn.addEventListener('click', () => {
        const modalId = btn.getAttribute('data-close');
        this.closeModal(modalId);
      });
    });

    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) {
          overlay.classList.remove('active');
        }
      });
    });

    this.setupDropzone();

    const taskForm = document.getElementById('taskForm');
    if (taskForm) {
      taskForm.addEventListener('submit', (e) => this.handleTaskFormSubmit(e));
    }

    const btnAllDays = document.getElementById('btnSelectAllDays');
    if (btnAllDays) {
      btnAllDays.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('input[name="prefDays"]').forEach(cb => cb.checked = true);
      });
    }

    const btnWorkdays = document.getElementById('btnSelectWorkdays');
    if (btnWorkdays) {
      btnWorkdays.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('input[name="prefDays"]').forEach(cb => {
          cb.checked = (cb.value !== 'sabado' && cb.value !== 'domingo');
        });
      });
    }

    const btnViewAll = document.getElementById('btnViewAllWeek');
    const btnViewWork = document.getElementById('btnViewWorkdays');

    if (btnViewAll && btnViewWork) {
      btnViewAll.addEventListener('click', () => {
        btnViewAll.classList.add('active');
        btnViewWork.classList.remove('active');
        CalendarView.toggleWorkdaysOnly(false);
      });

      btnViewWork.addEventListener('click', () => {
        btnViewWork.classList.add('active');
        btnViewAll.classList.remove('active');
        CalendarView.toggleWorkdaysOnly(true);
      });
    }

    const btnOptimize = document.getElementById('btnAutoOptimizeAll');
    if (btnOptimize) {
      btnOptimize.addEventListener('click', () => this.reorganizeAllTasks());
    }

    const btnClear = document.getElementById('btnClearCalendar');
    if (btnClear) {
      btnClear.addEventListener('click', () => this.clearScheduledTasks());
    }

    const btnDeleteBlock = document.getElementById('btnDeleteBlock');
    if (btnDeleteBlock) {
      btnDeleteBlock.addEventListener('click', () => this.deleteSelectedBlock());
    }

    // Botón de Cerrar Sesión
    const btnSignOut = document.getElementById('btnSignOut');
    if (btnSignOut) {
      btnSignOut.addEventListener('click', async () => {
        try {
          await SupabaseService.signOut();
          this.closeModal('authModal');
          this.showToast('Sesión cerrada correctamente', 'info');
          this.showView('landing');
        } catch (err) {
          this.showToast(err.message, 'error');
        }
      });
    }
  },

  // ==========================================
  // MANEJO DEL FORMULARIO PRINCIPAL DE AUTH
  // ==========================================

  initTurnstile() {
    this.turnstile.readyPromise = new Promise((resolve) => {
      const waitForTurnstile = () => {
        if (window.turnstile && typeof window.turnstile.render === 'function') {
          resolve();
          return;
        }
        window.setTimeout(waitForTurnstile, 100);
      };
      waitForTurnstile();
    });
  },

  async renderTurnstile() {
    const container = document.getElementById('turnstileWidget');
    if (!container) return;

    await this.turnstile.readyPromise;
    if (this.turnstile.widgetId !== null) return;

    this.turnstile.widgetId = window.turnstile.render(container, {
      sitekey: this.turnstile.siteKey,
      action: 'authenticate',
      execution: 'render',
      appearance: 'always',
      theme: 'light',
      size: 'normal',
      callback: (token) => {
        this.turnstile.token = token;
        this.setTurnstileError('');
        const submitBtn = document.getElementById('btnAuthFormSubmit');
        if (submitBtn) submitBtn.disabled = false;
      },
      'expired-callback': () => {
        this.turnstile.token = null;
        this.setTurnstileError('La verificación ha caducado. Completa el desafío de nuevo.');
        const submitBtn = document.getElementById('btnAuthFormSubmit');
        if (submitBtn) submitBtn.disabled = true;
      },
      'error-callback': () => {
        this.turnstile.token = null;
        this.setTurnstileError('No se pudo cargar la verificación. Recarga la página e inténtalo de nuevo.');
        const submitBtn = document.getElementById('btnAuthFormSubmit');
        if (submitBtn) submitBtn.disabled = true;
      }
    });
  },

  resetTurnstile() {
    this.turnstile.token = null;
    this.setTurnstileError('');
    const submitBtn = document.getElementById('btnAuthFormSubmit');
    if (submitBtn) submitBtn.disabled = true;
    if (this.turnstile.widgetId !== null && window.turnstile) {
      window.turnstile.reset(this.turnstile.widgetId);
    } else {
      this.renderTurnstile().catch((error) => {
        console.error('No se pudo inicializar Turnstile:', error);
      });
    }
  },

  setTurnstileError(message) {
    const errorElement = document.getElementById('turnstileError');
    if (errorElement) errorElement.textContent = message;
  },

  async handleMainAuthSubmit(e) {
    e.preventDefault();

    if (this.state.pendingConfirmationEmail) {
      this.showEmailConfirmation(this.state.pendingConfirmationEmail);
      this.showToast('Verifica tu correo antes de iniciar sesión', 'error');
      return;
    }

    const emailInput = document.getElementById('authEmailInput');
    const passwordInput = document.getElementById('authPasswordInput');
    const confirmInput = document.getElementById('authConfirmPasswordInput');
    const mismatchError = document.getElementById('passwordMismatchError');

    const email = emailInput ? emailInput.value.trim() : '';
    const password = passwordInput ? passwordInput.value : '';

    if (!email || !password) {
      this.showToast('Por favor completa todos los campos requeridos', 'error');
      return;
    }

    if (!this.turnstile.token) {
      this.setTurnstileError('Completa la verificación de seguridad para continuar.');
      this.showToast('Completa la verificación de seguridad', 'error');
      return;
    }

    // Modo SIGNUP: Validar que las contraseñas coincidan obligatoriamente
    if (this.state.authMode === 'signup') {
      const confirmPassword = confirmInput ? confirmInput.value : '';

      if (password !== confirmPassword) {
        if (mismatchError) mismatchError.style.display = 'block';
        if (confirmInput) confirmInput.focus();
        this.showToast('Las contraseñas no coinciden', 'error');
        return;
      }

      if (mismatchError) mismatchError.style.display = 'none';

      try {
        this.showToast('Creando tu cuenta...', 'info');
        const result = await SupabaseService.signUp(email, password, '', this.turnstile.token);
        const user = result?.user || null;
        const isConfirmed = !!(
          result?.session &&
          user &&
          (user.email_confirmed_at || user.confirmed_at)
        );

        if (!isConfirmed) {
          this.showToast('Revisa tu correo para confirmar la cuenta', 'info');
          this.showEmailConfirmation(email);
          return;
        }

        this.showToast('Cuenta creada con éxito', 'success');
        await this.loadInitialData();
        this.showView('app', { preview: false });
      } catch (err) {
        this.showToast(err.message, 'error');
        this.resetTurnstile();
      }
    } else {
      // Modo LOGIN
      try {
        this.showToast('Iniciando sesión...', 'info');
        await SupabaseService.signIn(email, password, this.turnstile.token);
        this.showToast('Sesión iniciada correctamente', 'success');

        await this.loadInitialData();
        this.showView('app', { preview: false });
      } catch (err) {
        this.showToast(err.message, 'error');
        this.resetTurnstile();
      }
    }
  },

  /**
   * Configura la zona de arrastrar y soltar archivos Excel
   */
  setupDropzone() {
    const dropzone = document.getElementById('dropzoneExcel');
    const fileInput = document.getElementById('excelFileInput');
    const btnSelect = document.getElementById('btnSelectExcelFile');

    if (!dropzone || !fileInput) return;

    if (btnSelect) {
      btnSelect.addEventListener('click', () => fileInput.click());
    }

    dropzone.addEventListener('click', (e) => {
      if (e.target !== btnSelect) fileInput.click();
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files.length > 0) {
        this.processExcelFile(files[0]);
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        this.processExcelFile(fileInput.files[0]);
        fileInput.value = ''; // Reset
      }
    });
  },

  /**
   * Procesa el archivo Excel cargado y valida estrictamente el formato
   */
  async processExcelFile(file) {
    this.showToast('Analizando archivo Excel con SheetJS...', 'info');

    const result = await ExcelParser.parseExcelFile(file);

    // Requisito estricto: si no cumple el formato, dar mensaje "formato no valido"
    if (!result.valid) {
      this.showToast(result.error || 'formato no valido', 'error');
      return;
    }

    // Formato válido: guardar bloques base en memoria
    this.state.fileName = result.fileName;
    this.state.baseBlocks = result.blocks;

    // Solo guardar de forma permanente si NO estamos en vista previa
    if (!this.state.isPreviewMode) {
      const saveRes = await SupabaseService.saveSchedule(
        'Horario Base Semanal',
        result.fileName,
        result.blocks
      );

      if (saveRes.schedule && saveRes.schedule.id) {
        this.state.scheduleId = saveRes.schedule.id;
      }
    }

    // Actualizar nombre de archivo en la interfaz
    const badge = document.getElementById('scheduleFileBadge');
    if (badge) {
      badge.textContent = result.fileName;
      badge.style.background = '#dcfce7';
      badge.style.color = '#15803d';
    }

    // Reorganizar las tareas existentes si las hubiera con el nuevo horario base
    if (this.state.tasks.length > 0) {
      this.reorganizeAllTasks(false);
    } else {
      this.updateCalendar();
    }

    this.closeModal('uploadModal');
    this.showToast(`Horario importado con éxito (${result.blocks.length} bloques).`, 'success');
  },

  /**
   * Maneja el formulario para añadir una nueva tarea / estudio
   */
  async handleTaskFormSubmit(e) {
    e.preventDefault();

    const title = document.getElementById('taskTitle').value.trim();
    if (!title) {
      this.showToast('Por favor introduce un título', 'error');
      return;
    }

    const category = document.getElementById('taskCategory').value;
    const priority = document.getElementById('taskPriority').value;
    const hours = parseInt(document.getElementById('taskDurationHours').value, 10) || 0;
    const minutes = parseInt(document.getElementById('taskDurationMinutes').value, 10) || 0;
    const totalMinutes = hours * 60 + minutes;

    if (totalMinutes <= 0) {
      this.showToast('El tiempo total debe ser mayor a 0 minutos', 'error');
      return;
    }

    const maxSessionMinutes = parseInt(document.getElementById('taskMaxSession').value, 10) || 60;
    const restBetweenMinutes = parseInt(document.getElementById('taskRestBetween').value, 10) || 0;
    const timePreference = document.getElementById('taskTimePreference').value;

    const checkedDays = Array.from(document.querySelectorAll('input[name="prefDays"]:checked'))
      .map(cb => cb.value);

    if (checkedDays.length === 0) {
      this.showToast('Selecciona al menos un día permitido', 'error');
      return;
    }

    const newTask = {
      id: 'task-' + Date.now(),
      scheduleId: this.state.scheduleId,
      title: title,
      category: category,
      priority: priority,
      durationMinutes: totalMinutes,
      maxSessionMinutes: maxSessionMinutes,
      restBetweenMinutes: restBetweenMinutes,
      timePreference: timePreference,
      preferredDays: checkedDays,
      status: 'scheduled',
      createdAt: new Date().toISOString()
    };

    const currentOccupied = [
      ...this.state.baseBlocks,
      ...this.state.taskBlocks,
      ...this.state.breakBlocks
    ];

    const scheduleRes = SmartScheduler.scheduleTask(newTask, currentOccupied);

    if (scheduleRes.placedBlocks.length === 0) {
      this.showToast(scheduleRes.message, 'error');
      return;
    }

    this.state.tasks.push(newTask);
    this.state.taskBlocks.push(...scheduleRes.placedBlocks);
    this.state.breakBlocks.push(...scheduleRes.breakBlocks);

    // Solo guardar en Supabase/Local si NO es preview
    if (!this.state.isPreviewMode) {
      await SupabaseService.saveTask(newTask);
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();

    document.getElementById('taskTitle').value = '';
    this.closeModal('taskModal');

    if (scheduleRes.success) {
      this.showToast(scheduleRes.message, 'success');
    } else {
      this.showToast(scheduleRes.message, 'info');
    }
  },

  openNewTaskModalWithSlot(day, hour) {
    document.querySelectorAll('input[name="prefDays"]').forEach(cb => {
      cb.checked = (cb.value === day);
    });

    const prefSelect = document.getElementById('taskTimePreference');
    if (prefSelect) {
      if (hour >= 8 && hour < 14) prefSelect.value = 'morning';
      else if (hour >= 14 && hour < 20) prefSelect.value = 'afternoon';
      else if (hour >= 20) prefSelect.value = 'evening';
      else prefSelect.value = 'any';
    }

    this.openModal('taskModal');
  },

  reorganizeAllTasks(showToastFeedback = true) {
    if (this.state.tasks.length === 0) {
      if (showToastFeedback) this.showToast('No hay tareas registradas para organizar', 'info');
      return;
    }

    const result = SmartScheduler.reorganizeAll(this.state.tasks, this.state.baseBlocks);
    this.state.taskBlocks = result.taskBlocks;
    this.state.breakBlocks = result.breakBlocks;

    this.updateCalendar();
    this.updateStats();

    if (showToastFeedback) {
      this.showToast('¡Horario y tareas re-organizados con éxito!', 'success');
    }
  },

  clearScheduledTasks() {
    if (this.state.taskBlocks.length === 0 && this.state.tasks.length === 0) {
      this.showToast('El calendario ya está limpio', 'info');
      return;
    }

    if (confirm('¿Deseas vaciar las tareas y descansos programados? El horario base de Excel se mantendrá.')) {
      this.state.taskBlocks = [];
      this.state.breakBlocks = [];
      this.state.tasks = [];
      if (!this.state.isPreviewMode) {
        localStorage.removeItem('cronosys_tasks');
      }

      this.updateCalendar();
      this.renderTaskCards();
      this.updateStats();
      this.showToast('Tareas y descansos eliminados', 'info');
    }
  },

  showBlockDetailModal(block) {
    this.state.selectedBlock = block;

    const titleEl = document.getElementById('detailBlockTitle');
    const dayEl = document.getElementById('detailBlockDay');
    const timeEl = document.getElementById('detailBlockTime');
    const typeEl = document.getElementById('detailBlockType');
    const notesEl = document.getElementById('detailBlockNotes');
    const notesWrap = document.getElementById('detailBlockNotesWrap');

    if (titleEl) titleEl.textContent = block.title;
    if (dayEl) dayEl.textContent = block.day;
    if (timeEl) timeEl.textContent = `${block.startTime} - ${block.endTime}`;

    const typeLabels = {
      base: '🏢 Horario Base (Excel)',
      study: '📚 Tiempo de Estudio',
      task: '💼 Tarea Programada',
      break: '☕ Descanso / Pausa'
    };

    if (typeEl) typeEl.textContent = typeLabels[block.type] || block.type;

    if (notesEl && notesWrap) {
      if (block.notes) {
        notesWrap.style.display = 'block';
        notesEl.textContent = block.notes;
      } else {
        notesWrap.style.display = 'none';
      }
    }

    this.openModal('blockDetailModal');
  },

  deleteSelectedBlock() {
    const block = this.state.selectedBlock;
    if (!block) return;

    if (block.type === 'base') {
      this.state.baseBlocks = this.state.baseBlocks.filter(b => b.id !== block.id);
    } else if (block.type === 'study' || block.type === 'task') {
      this.state.taskBlocks = this.state.taskBlocks.filter(b => b.id !== block.id);
    } else if (block.type === 'break') {
      this.state.breakBlocks = this.state.breakBlocks.filter(b => b.id !== block.id);
    }

    this.closeModal('blockDetailModal');
    this.updateCalendar();
    this.updateStats();
    this.showToast(`Bloque "${block.title}" eliminado`, 'info');
  },

  updateCalendar() {
    const all = [
      ...this.state.baseBlocks,
      ...this.state.taskBlocks,
      ...this.state.breakBlocks
    ];
    CalendarView.setBlocks(all);
  },

  renderTaskCards() {
    const container = document.getElementById('taskCardsGrid');
    if (!container) return;

    if (this.state.tasks.length === 0) {
      container.innerHTML = `
        <div class="empty-state-card">
          <div class="empty-state-icon">
            <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="12" cy="12" r="10"></circle>
              <polyline points="12 6 12 12 16 14"></polyline>
            </svg>
          </div>
          <div class="empty-state-title">No hay tareas o estudios añadidos aún</div>
          <p class="empty-state-desc">
            Sube tu horario en Excel y añade sesiones de estudio o tareas. Crono-SYS las organizará de manera inteligente en los huecos disponibles con tus descansos preferidos.
          </p>
          <button class="btn btn-primary" id="btnEmptyStateAddCard">
            Crear Primera Tarea
          </button>
        </div>
      `;

      const btn = document.getElementById('btnEmptyStateAddCard');
      if (btn) btn.addEventListener('click', () => this.openModal('taskModal'));
      return;
    }

    let html = '';
    this.state.tasks.forEach(task => {
      const hours = Math.floor(task.durationMinutes / 60);
      const mins = task.durationMinutes % 60;
      const durationStr = `${hours > 0 ? hours + 'h ' : ''}${mins > 0 ? mins + 'm' : ''}`.trim();

      const daysStr = (task.preferredDays || []).map(d => d.substring(0, 3).toUpperCase()).join(', ');

      html += `
        <div class="task-item-card" data-id="${task.id}">
          <div class="task-item-header">
            <span class="task-item-title">${task.title}</span>
            <span class="task-category-pill ${task.category || 'estudio'}">${task.category || 'Estudio'}</span>
          </div>

          <div class="task-item-details">
            <div class="task-meta-row">
              <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <polyline points="12 6 12 12 16 14"></polyline>
              </svg>
              <span>Tiempo total: <strong>${durationStr}</strong> (en bloques de ${task.maxSessionMinutes || 60}m)</span>
            </div>

            <div class="task-meta-row">
              <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="16" y1="2" x2="16" y2="6"></line>
                <line x1="8" y1="2" x2="8" y2="6"></line>
              </svg>
              <span>Días: <strong>${daysStr}</strong></span>
            </div>

            <div class="task-meta-row">
              <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 8h1a4 4 0 0 1 0 8h-1"></path>
                <path d="M2 8h16v9a4 4 0 0 1-4 4H6a4 4 0 0 1-4-4V8z"></path>
              </svg>
              <span>Descanso: <strong>${task.restBetweenMinutes || 0} min</strong></span>
            </div>
          </div>

          <div class="task-item-actions">
            <span style="font-size: 0.75rem; color: #16a34a; font-weight: 700;">● Programada</span>
            <button class="btn btn-secondary btn-delete-task" data-task-id="${task.id}" style="padding: 0.3rem 0.6rem; font-size: 0.75rem; color: #dc2626;">
              Eliminar
            </button>
          </div>
        </div>
      `;
    });

    container.innerHTML = html;

    container.querySelectorAll('.btn-delete-task').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const taskId = btn.getAttribute('data-task-id');
        this.deleteTask(taskId);
      });
    });
  },

  async deleteTask(taskId) {
    this.state.tasks = this.state.tasks.filter(t => t.id !== taskId);
    this.state.taskBlocks = this.state.taskBlocks.filter(b => b.taskId !== taskId);
    this.state.breakBlocks = this.state.breakBlocks.filter(b => b.taskId !== taskId);

    if (!this.state.isPreviewMode) {
      await SupabaseService.deleteTask(taskId);
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();
    this.showToast('Tarea eliminada del planificador', 'info');
  },

  updateStats() {
    const statBase = document.getElementById('statBaseCount');
    if (statBase) {
      statBase.textContent = `${this.state.baseBlocks.length} bloques`;
    }

    let totalStudyMin = 0;
    this.state.taskBlocks.forEach(b => {
      const s = CalendarView.timeToMinutes(b.startTime);
      let e = CalendarView.timeToMinutes(b.endTime);
      if (e <= s) e = 1440;
      totalStudyMin += (e - s);
    });

    const statStudy = document.getElementById('statStudyHours');
    if (statStudy) {
      const h = Math.floor(totalStudyMin / 60);
      const m = totalStudyMin % 60;
      statStudy.textContent = `${h}h ${m.toString().padStart(2, '0')}m`;
    }

    const statBreaks = document.getElementById('statBreaksCount');
    if (statBreaks) {
      statBreaks.textContent = `${this.state.breakBlocks.length} pausas`;
    }
  },

  async loadInitialData() {
    const schedRes = await SupabaseService.loadActiveSchedule();
    if (schedRes.success && schedRes.data) {
      this.state.scheduleId = schedRes.data.id;
      this.state.fileName = schedRes.data.fileName || 'Horario Activo';
      this.state.baseBlocks = (schedRes.data.blocks || []).filter(b => b.type === 'base');

      const badge = document.getElementById('scheduleFileBadge');
      if (badge && this.state.fileName) {
        badge.textContent = this.state.fileName;
        badge.style.background = '#dcfce7';
        badge.style.color = '#15803d';
      }
    }

    const taskRes = await SupabaseService.loadTasks();
    if (taskRes.success && taskRes.tasks) {
      this.state.tasks = taskRes.tasks;
      if (this.state.tasks.length > 0) {
        this.reorganizeAllTasks(false);
      }
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();
  },

  // ==========================================
  // AUTENTICACIÓN
  // ==========================================

  setupAuth() {
    SupabaseService.onAuthStateChange((event, session) => {
      this.updateAuthUI(session ? session.user : null);

      if (event === 'SIGNED_IN' && session?.user) {
        const confirmedEmail = session.user.email_confirmed_at || session.user.confirmed_at;
        if (confirmedEmail) {
          this.clearPendingEmailConfirmation();
          this.loadInitialData()
            .then(() => this.showView('app', { preview: false }));
        }
      }
    });

    SupabaseService.getSession().then(session => {
      this.updateAuthUI(session ? session.user : null);
    });
  },

  updateAuthUI(user) {
    const statSecurity = document.getElementById('statSecurityText');
    const headerInitials = document.getElementById('headerAvatarInitials');
    const flyoutInitials = document.getElementById('flyoutAvatarInitials');
    const flyoutName = document.getElementById('flyoutUserName');
    const flyoutEmail = document.getElementById('flyoutUserEmail');

    if (user) {
      // Extraer iniciales inteligentes (ej. Javier García -> JG o JM)
      let initials = 'JM';
      const emailName = user.email ? user.email.split('@')[0] : 'U';
      const parts = emailName.replace(/[._-]/g, ' ').trim().split(/\s+/);
      if (parts.length >= 2) {
        initials = (parts[0][0] + parts[1][0]).toUpperCase();
      } else if (parts[0].length >= 2) {
        initials = parts[0].substring(0, 2).toUpperCase();
      }

      const displayName = user.user_metadata?.full_name || emailName.replace(/[._-]/g, ' ').toUpperCase();

      if (headerInitials) headerInitials.textContent = initials;
      if (flyoutInitials) flyoutInitials.textContent = initials;
      if (flyoutName) flyoutName.textContent = displayName;
      if (flyoutEmail) flyoutEmail.textContent = user.email;

      if (statSecurity && !this.state.isPreviewMode) statSecurity.textContent = 'Ha iniciado sesión';
    } else {
      if (headerInitials) headerInitials.textContent = 'IN';
      if (flyoutInitials) flyoutInitials.textContent = 'IN';
      if (flyoutName) flyoutName.textContent = 'MODO INVITADO';
      if (flyoutEmail) flyoutEmail.textContent = 'Sin sesión sincronizada';
      if (statSecurity && !this.state.isPreviewMode) statSecurity.textContent = 'No ha iniciado sesión';
    }
  },

  // ==========================================
  // UTILIDADES DE INTERFAZ (MODALES Y TOASTS)
  // ==========================================

  openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.add('active');
  },

  closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.remove('active');
  },

  showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let iconSvg = '';
    if (type === 'error') {
      iconSvg = `
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"></circle>
          <line x1="15" y1="9" x2="9" y2="15"></line>
          <line x1="9" y1="9" x2="15" y2="15"></line>
        </svg>
      `;
    } else if (type === 'success') {
      iconSvg = `
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
          <polyline points="22 4 12 14.01 9 11.01"></polyline>
        </svg>
      `;
    } else {
      iconSvg = `
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"></circle>
          <line x1="12" y1="16" x2="12" y2="12"></line>
          <line x1="12" y1="8" x2="12.01" y2="8"></line>
        </svg>
      `;
    }

    toast.innerHTML = `${iconSvg}<span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(100%)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }
};

// Arrancar cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});

window.App = App;
