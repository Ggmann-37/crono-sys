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
    pendingConfirmationEmail: '',
    pendingImport: null,
    editingTaskId: null,
    previewInitialized: false,
    activeModalId: null,
    modalTrigger: null
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
      CalendarView.setMobileDayView(window.matchMedia('(max-width: 640px)').matches);
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
        this.setStorageStatus('preview');
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

  startPreview() {
    if (this.state.previewInitialized) return;
    this.state.previewInitialized = true;
    this.state.fileName = 'Horario de ejemplo';
    this.state.baseBlocks = [
      { id: 'demo-base-mon', day: 'lunes', startTime: '09:00', endTime: '12:00', title: 'Clases', type: 'base', color: '#2563eb' },
      { id: 'demo-base-tue', day: 'martes', startTime: '10:00', endTime: '13:00', title: 'Trabajo', type: 'base', color: '#2563eb' },
      { id: 'demo-base-wed', day: 'miercoles', startTime: '09:00', endTime: '12:00', title: 'Clases', type: 'base', color: '#2563eb' },
      { id: 'demo-base-thu', day: 'jueves', startTime: '10:00', endTime: '13:00', title: 'Trabajo', type: 'base', color: '#2563eb' },
      { id: 'demo-base-fri', day: 'viernes', startTime: '09:00', endTime: '11:00', title: 'Clases', type: 'base', color: '#2563eb' }
    ];
    this.state.tasks = [
      {
        id: 'demo-task-math',
        title: 'Repasar matemáticas',
        category: 'estudio',
        priority: 'high',
        durationMinutes: 120,
        maxSessionMinutes: 60,
        restBetweenMinutes: 15,
        timePreference: 'afternoon',
        preferredDays: ['lunes', 'miercoles', 'viernes'],
        status: 'pending'
      },
      {
        id: 'demo-task-project',
        title: 'Avanzar el proyecto',
        category: 'proyecto',
        priority: 'medium',
        durationMinutes: 90,
        maxSessionMinutes: 45,
        restBetweenMinutes: 10,
        timePreference: 'any',
        preferredDays: ['martes', 'jueves'],
        status: 'pending'
      }
    ];

    const scheduleBadge = document.getElementById('scheduleFileBadge');
    if (scheduleBadge) {
      scheduleBadge.textContent = this.state.fileName;
      scheduleBadge.style.background = '#dcfce7';
      scheduleBadge.style.color = '#15803d';
    }

    const result = SmartScheduler.reorganizeAll(this.state.tasks, this.state.baseBlocks);
    this.state.taskBlocks = result.taskBlocks;
    this.state.breakBlocks = result.breakBlocks;
    const taskStatuses = new Map(result.report.map(item => [item.taskId, item.success ? 'scheduled' : 'pending']));
    this.state.tasks = this.state.tasks.map(task => ({ ...task, status: taskStatuses.get(task.id) || 'pending' }));
    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();
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
    if (btnHeroPreview) btnHeroPreview.addEventListener('click', () => {
      this.startPreview();
      this.showView('app', { preview: true });
    });
    if (btnNavPreview) btnNavPreview.addEventListener('click', () => {
      this.startPreview();
      this.showView('app', { preview: true });
    });

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
      btnOpenTask.addEventListener('click', () => this.openTaskModal());
    }

    const btnQuickAdd = document.getElementById('btnQuickAddTask');
    if (btnQuickAdd) {
      btnQuickAdd.addEventListener('click', () => this.openTaskModal());
    }

    const btnEmptyAdd = document.getElementById('btnEmptyStateAdd');
    if (btnEmptyAdd) {
      btnEmptyAdd.addEventListener('click', () => this.openTaskModal());
    }

    const taskFilter = document.getElementById('taskStatusFilter');
    if (taskFilter) taskFilter.addEventListener('change', () => this.renderTaskCards());

    const taskCards = document.getElementById('taskCardsGrid');
    if (taskCards) {
      taskCards.addEventListener('click', (event) => {
        const button = event.target.closest('button[data-task-action]');
        if (!button) return;
        const { taskAction, taskId } = button.dataset;
        if (taskAction === 'complete') this.setTaskCompleted(taskId, true);
        if (taskAction === 'reopen') this.setTaskCompleted(taskId, false);
        if (taskAction === 'edit') this.openTaskModal(this.state.tasks.find(task => task.id === taskId));
        if (taskAction === 'delete') this.deleteTask(taskId);
      });
    }

    const confirmImport = document.getElementById('btnConfirmExcelImport');
    if (confirmImport) confirmImport.addEventListener('click', () => this.confirmExcelImport());
    const cancelImport = document.getElementById('btnCancelExcelImport');
    if (cancelImport) cancelImport.addEventListener('click', () => this.cancelExcelImport());

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
          this.closeModal(overlay.id);
        }
      });
    });

    document.addEventListener('keydown', (event) => {
      const modal = this.state.activeModalId && document.getElementById(this.state.activeModalId);
      if (!modal) return;

      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeModal(this.state.activeModalId);
        return;
      }

      if (event.key !== 'Tab') return;
      const focusable = [...modal.querySelectorAll(
        'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])'
      )].filter(element => element.getClientRects().length > 0);
      if (focusable.length === 0) {
        event.preventDefault();
        modal.querySelector('.modal-container')?.focus();
        return;
      }

      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
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
    if (this.state.pendingImport) this.cancelExcelImport();
    this.showToast('Analizando archivo Excel con SheetJS...', 'info');

    const result = await ExcelParser.parseExcelFile(file);

    if (!result.valid) {
      this.showToast(result.error || 'No se pudo validar el archivo.', 'error');
      return;
    }

    this.state.pendingImport = result;
    const preview = document.getElementById('excelImportPreview');
    const confirmButton = document.getElementById('btnConfirmExcelImport');
    const cancelButton = document.getElementById('btnCancelExcelImport');
    const warnings = result.warnings || [];
    if (preview) {
      preview.innerHTML = `
        <h4>Revisa el horario detectado</h4>
        <p><strong>${this.escapeHtml(result.fileName)}</strong>: ${result.blocks.length} bloques.</p>
        ${warnings.length
          ? `<p class="import-warning">Se detectaron posibles solapamientos:</p><ul>${warnings.map(warning => `<li>${this.escapeHtml(warning)}</li>`).join('')}</ul>`
          : '<p>No se detectaron solapamientos.</p>'}
      `;
      preview.classList.remove('view-hidden');
    }
    if (confirmButton) confirmButton.classList.remove('view-hidden');
    if (cancelButton) cancelButton.classList.remove('view-hidden');
    this.showToast('Archivo revisado. Confirma la importación cuando estés listo.', 'success');
  },

  async confirmExcelImport() {
    const result = this.state.pendingImport;
    if (!result) return;

    this.state.pendingImport = null;
    this.state.fileName = result.fileName;
    this.state.baseBlocks = result.blocks;
    let saveRes = { isLocal: false };

    if (!this.state.isPreviewMode) {
      saveRes = await SupabaseService.saveSchedule(
        'Horario Base Semanal',
        result.fileName,
        result.blocks
      );
      if (saveRes.schedule?.id) this.state.scheduleId = saveRes.schedule.id;
      this.setStorageStatus(saveRes.isLocal ? 'local' : 'cloud');
    }

    const badge = document.getElementById('scheduleFileBadge');
    if (badge) {
      badge.textContent = result.fileName;
      badge.style.background = '#dcfce7';
      badge.style.color = '#15803d';
    }

    if (this.state.tasks.length > 0) await this.reorganizeAllTasks(false);
    else this.updateCalendar();

    this.resetExcelImportPreview();
    this.closeModal('uploadModal');
    this.showToast(
      this.state.isPreviewMode
        ? `Horario de prueba importado (${result.blocks.length} bloques). No se guardará.`
        : saveRes?.isLocal
        ? saveRes.error
          ? 'Horario guardado solo en este dispositivo; no se pudo sincronizar.'
          : 'Horario guardado en este dispositivo.'
          : `Horario importado y sincronizado (${result.blocks.length} bloques).`,
      saveRes?.isLocal && saveRes.error && !this.state.isPreviewMode ? 'error' : 'success'
    );
  },

  cancelExcelImport() {
    this.state.pendingImport = null;
    this.resetExcelImportPreview();
  },

  resetExcelImportPreview() {
    const preview = document.getElementById('excelImportPreview');
    const confirmButton = document.getElementById('btnConfirmExcelImport');
    const cancelButton = document.getElementById('btnCancelExcelImport');
    if (preview) {
      preview.innerHTML = '';
      preview.classList.add('view-hidden');
    }
    if (confirmButton) confirmButton.classList.add('view-hidden');
    if (cancelButton) cancelButton.classList.add('view-hidden');
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

    const editingTask = this.state.tasks.find(task => task.id === this.state.editingTaskId);
    const task = {
      ...(editingTask || {}),
      id: editingTask?.id || `task-${Date.now()}`,
      scheduleId: this.state.scheduleId || editingTask?.scheduleId || null,
      title: title,
      category: category,
      priority: priority,
      durationMinutes: totalMinutes,
      maxSessionMinutes: maxSessionMinutes,
      restBetweenMinutes: restBetweenMinutes,
      timePreference: timePreference,
      preferredDays: checkedDays,
      status: 'pending',
      createdAt: editingTask?.createdAt || new Date().toISOString()
    };

    const currentOccupied = [
      ...this.state.baseBlocks,
      ...this.state.taskBlocks.filter(block => block.taskId !== task.id),
      ...this.state.breakBlocks.filter(block => block.taskId !== task.id)
    ];

    const scheduleRes = SmartScheduler.scheduleTask(task, currentOccupied);
    task.status = scheduleRes.success ? 'scheduled' : 'pending';
    this.state.tasks = editingTask
      ? this.state.tasks.map(existing => existing.id === task.id ? task : existing)
      : [...this.state.tasks, task];
    this.state.taskBlocks = currentOccupied.filter(block => block.type === 'study' || block.type === 'task')
      .concat(scheduleRes.placedBlocks);
    this.state.breakBlocks = currentOccupied.filter(block => block.type === 'break')
      .concat(scheduleRes.breakBlocks);

    // Solo guardar en Supabase/Local si NO es preview
    let saveRes = { isLocal: false, task };
    if (!this.state.isPreviewMode) {
      saveRes = await SupabaseService.saveTask(task);
      if (saveRes.task?.id && saveRes.task.id !== task.id) {
        const temporaryId = task.id;
        task.id = saveRes.task.id;
        [...this.state.taskBlocks, ...this.state.breakBlocks]
          .filter(block => block.taskId === temporaryId)
          .forEach(block => { block.taskId = task.id; });
      }
      this.setStorageStatus(saveRes.isLocal ? 'local' : 'cloud');
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();

    document.getElementById('taskForm').reset();
    this.state.editingTaskId = null;
    document.getElementById('taskModalTitle').textContent = '⚡ Organizar Tarea / Tiempo de Estudio';
    document.getElementById('btnSubmitTask').textContent = '⚡ Organizar Automáticamente';
    this.closeModal('taskModal');

    if (this.state.isPreviewMode && scheduleRes.success) {
      this.showToast(scheduleRes.message, 'success');
    } else if (this.state.isPreviewMode) {
      this.showToast(scheduleRes.message, 'info');
    } else if (saveRes.isLocal) {
      this.showToast(
        saveRes.error
          ? 'Tarea guardada solo en este dispositivo; no se pudo sincronizar.'
          : 'Tarea guardada en este dispositivo.',
        saveRes.error ? 'error' : 'info'
      );
    } else if (scheduleRes.success) {
      this.showToast(scheduleRes.message, 'success');
    } else {
      this.showToast(scheduleRes.message, 'info');
    }
  },

  openTaskModal(task = null) {
    const form = document.getElementById('taskForm');
    if (form) form.reset();

    this.state.editingTaskId = task?.id || null;
    const title = document.getElementById('taskModalTitle');
    const submit = document.getElementById('btnSubmitTask');
    if (task) {
      document.getElementById('taskTitle').value = task.title;
      document.getElementById('taskCategory').value = task.category || 'estudio';
      document.getElementById('taskPriority').value = task.priority || 'medium';
      document.getElementById('taskDurationHours').value = Math.floor(task.durationMinutes / 60);
      document.getElementById('taskDurationMinutes').value = task.durationMinutes % 60;
      document.getElementById('taskMaxSession').value = String(task.maxSessionMinutes || 60);
      document.getElementById('taskRestBetween').value = String(task.restBetweenMinutes || 0);
      document.getElementById('taskTimePreference').value = task.timePreference || 'any';
      document.querySelectorAll('input[name="prefDays"]').forEach(input => {
        input.checked = (task.preferredDays || []).includes(input.value);
      });
      if (title) title.textContent = 'Editar tarea';
      if (submit) submit.textContent = 'Guardar cambios y reorganizar';
    } else {
      document.querySelectorAll('input[name="prefDays"]').forEach(input => {
        input.checked = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes'].includes(input.value);
      });
      if (title) title.textContent = '⚡ Organizar Tarea / Tiempo de Estudio';
      if (submit) submit.textContent = '⚡ Organizar Automáticamente';
    }

    this.openModal('taskModal');
  },

  openNewTaskModalWithSlot(day, hour) {
    this.openTaskModal();
    document.querySelectorAll('input[name="prefDays"]').forEach(cb => {
      cb.checked = cb.value === day;
    });

    const prefSelect = document.getElementById('taskTimePreference');
    if (prefSelect) {
      if (hour >= 8 && hour < 14) prefSelect.value = 'morning';
      else if (hour >= 14 && hour < 20) prefSelect.value = 'afternoon';
      else if (hour >= 20) prefSelect.value = 'evening';
      else prefSelect.value = 'any';
    }

  },

  async reorganizeAllTasks(showToastFeedback = true) {
    const openTasks = this.state.tasks.filter(task => task.status !== 'completed');
    if (openTasks.length === 0) {
      if (showToastFeedback) this.showToast('No hay tareas registradas para organizar', 'info');
      return;
    }

    const result = SmartScheduler.reorganizeAll(openTasks, this.state.baseBlocks);
    this.state.taskBlocks = result.taskBlocks;
    this.state.breakBlocks = result.breakBlocks;
    const statuses = new Map(result.report.map(item => [item.taskId, item.success ? 'scheduled' : 'pending']));
    this.state.tasks = this.state.tasks.map(task => ({
      ...task,
      status: task.status === 'completed' ? task.status : (statuses.get(task.id) || 'pending')
    }));

    if (!this.state.isPreviewMode) {
      const tasksToSave = this.state.tasks.filter(task => task.status !== 'completed');
      const saveResults = await Promise.all(tasksToSave.map(task => SupabaseService.saveTask(task)));
      saveResults.forEach((save, index) => {
        const task = tasksToSave[index];
        if (!save.task?.id || save.task.id === task.id) return;
        const oldId = task.id;
        task.id = save.task.id;
        [...this.state.taskBlocks, ...this.state.breakBlocks]
          .filter(block => block.taskId === oldId)
          .forEach(block => { block.taskId = task.id; });
      });
      if (saveResults.some(save => save.isLocal)) this.setStorageStatus('local');
      else if (saveResults.length) this.setStorageStatus('cloud');
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();

    if (showToastFeedback) {
      const incomplete = result.report.filter(item => !item.success);
      if (incomplete.length) {
        const remaining = incomplete.reduce((total, item) => total + item.unplacedMinutes, 0);
        this.showToast(`Reorganización parcial: ${remaining} min pendientes en ${incomplete.length} tarea(s). Revisa “Mis tareas”.`, 'info');
      } else {
        this.showToast('Horario y tareas reorganizados con éxito.', 'success');
      }
    }
  },

  async clearScheduledTasks() {
    if (this.state.taskBlocks.length === 0 && this.state.tasks.length === 0) {
      this.showToast('El calendario ya está limpio', 'info');
      return;
    }

    if (confirm('¿Deseas vaciar las tareas y descansos programados? El horario base de Excel se mantendrá.')) {
      let failedIds = [];
      if (!this.state.isPreviewMode) {
        const results = await Promise.all(this.state.tasks.map(task => SupabaseService.deleteTask(task.id)));
        failedIds = this.state.tasks.filter((task, index) => !results[index].success).map(task => task.id);
      }
      this.state.tasks = this.state.tasks.filter(task => failedIds.includes(task.id));
      this.state.taskBlocks = this.state.taskBlocks.filter(block => failedIds.includes(block.taskId));
      this.state.breakBlocks = this.state.breakBlocks.filter(block => failedIds.includes(block.taskId));

      this.updateCalendar();
      this.renderTaskCards();
      this.updateStats();
      if (failedIds.length) {
        this.showToast(`No se pudieron eliminar ${failedIds.length} tarea(s) de la nube. Siguen en la lista.`, 'error');
      } else {
        this.showToast('Tareas y descansos eliminados', 'info');
      }
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

    const filter = document.getElementById('taskStatusFilter')?.value || 'all';
    const visibleTasks = this.state.tasks.filter(task =>
      filter === 'all' ||
      (filter === 'completed' ? task.status === 'completed' : task.status !== 'completed')
    );

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
      if (btn) btn.addEventListener('click', () => this.openTaskModal());
      return;
    }

    if (visibleTasks.length === 0) {
      container.innerHTML = '<p class="task-filter-empty">No hay tareas en este filtro.</p>';
      return;
    }

    let html = '';
    visibleTasks.forEach(task => {
      const hours = Math.floor(task.durationMinutes / 60);
      const mins = task.durationMinutes % 60;
      const durationStr = `${hours > 0 ? hours + 'h ' : ''}${mins > 0 ? mins + 'm' : ''}`.trim();
      const daysStr = (task.preferredDays || []).map(d => d.substring(0, 3).toUpperCase()).join(', ');
      const placedMinutes = this.state.taskBlocks
        .filter(block => block.taskId === task.id)
        .reduce((total, block) => {
          const start = CalendarView.timeToMinutes(block.startTime);
          let end = CalendarView.timeToMinutes(block.endTime);
          if (end <= start) end = 1440;
          return total + end - start;
        }, 0);
      const remainingMinutes = Math.max(0, task.durationMinutes - placedMinutes);
      const statusText = task.status === 'completed'
        ? 'Completada'
        : remainingMinutes > 0
          ? `Pendiente de encajar ${remainingMinutes} min`
          : 'Programada';
      const categoryLabels = {
        estudio: 'Estudio',
        tarea: 'Tarea',
        proyecto: 'Proyecto',
        personal: 'Personal'
      };
      const category = ['estudio', 'tarea', 'proyecto', 'personal'].includes(task.category)
        ? task.category
        : 'estudio';
      html += `
        <article class="task-item-card ${task.status === 'completed' ? 'is-completed' : ''}" data-id="${this.escapeHtml(task.id)}">
          <div class="task-item-header">
            <h3 class="task-item-title">${this.escapeHtml(task.title)}</h3>
            <span class="task-category-pill ${category}">${categoryLabels[category]}</span>
          </div>

          <div class="task-item-details">
            <div class="task-meta-row">
              <span>Tiempo total: <strong>${this.escapeHtml(durationStr)}</strong> (sesiones de ${Number(task.maxSessionMinutes) || 60} min)</span>
            </div>
            <div class="task-meta-row">
              <span>Días preferidos: <strong>${this.escapeHtml(daysStr || 'Sin días')}</strong></span>
            </div>
            <div class="task-meta-row">
              <span>Descanso: <strong>${Number(task.restBetweenMinutes) || 0} min</strong></span>
            </div>
          </div>

          <div class="task-item-actions">
            <span class="task-status ${task.status === 'completed' ? 'completed' : remainingMinutes > 0 ? 'pending' : 'scheduled'}">${statusText}</span>
            <div class="task-action-buttons">
              <button class="btn btn-secondary" type="button" aria-label="Editar ${this.escapeHtml(task.title)}" data-task-action="edit" data-task-id="${this.escapeHtml(task.id)}">Editar</button>
              <button class="btn btn-secondary" type="button" aria-label="${task.status === 'completed' ? 'Reabrir' : 'Completar'} ${this.escapeHtml(task.title)}" data-task-action="${task.status === 'completed' ? 'reopen' : 'complete'}" data-task-id="${this.escapeHtml(task.id)}">
                ${task.status === 'completed' ? 'Reabrir' : 'Completar'}
              </button>
              <button class="btn btn-secondary btn-delete-task" type="button" data-task-action="delete" data-task-id="${this.escapeHtml(task.id)}">Eliminar</button>
            </div>
          </div>
        </article>
      `;
    });

    container.innerHTML = html;
  },

  async setTaskCompleted(taskId, completed) {
    const task = this.state.tasks.find(item => item.id === taskId);
    if (!task) return;

    task.status = completed ? 'completed' : 'pending';
    let localFallback = false;
    let syncFailed = false;
    if (completed) {
      this.state.taskBlocks = this.state.taskBlocks.filter(block => block.taskId !== taskId);
      this.state.breakBlocks = this.state.breakBlocks.filter(block => block.taskId !== taskId);
      if (!this.state.isPreviewMode) {
        const save = await SupabaseService.saveTask(task);
        if (save.task?.id) task.id = save.task.id;
        localFallback = save.isLocal;
        syncFailed = !!save.error;
        this.setStorageStatus(save.isLocal ? 'local' : 'cloud');
      }
    } else {
      await this.reorganizeAllTasks(false);
    }

    this.updateCalendar();
    this.renderTaskCards();
    this.updateStats();
    this.showToast(
      localFallback
        ? syncFailed
          ? 'Tarea completada. El cambio quedó solo en este dispositivo porque falló la sincronización.'
          : 'Tarea completada en este dispositivo.'
        : completed ? 'Tarea marcada como completada.' : 'Tarea reabierta.',
      syncFailed ? 'error' : localFallback ? 'info' : 'success'
    );
  },

  async deleteTask(taskId) {
    if (!this.state.isPreviewMode) {
      const result = await SupabaseService.deleteTask(taskId);
      if (!result.success) {
        this.showToast('No se pudo borrar la tarea de la nube; sigue en tu lista.', 'error');
        return;
      }
    }

    this.state.tasks = this.state.tasks.filter(task => task.id !== taskId);
    this.state.taskBlocks = this.state.taskBlocks.filter(block => block.taskId !== taskId);
    this.state.breakBlocks = this.state.breakBlocks.filter(block => block.taskId !== taskId);
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
    this.state.scheduleId = null;
    this.state.fileName = null;
    this.state.baseBlocks = [];
    this.state.taskBlocks = [];
    this.state.breakBlocks = [];
    this.state.tasks = [];
    const badge = document.getElementById('scheduleFileBadge');
    if (badge) {
      badge.textContent = 'Sin archivo cargado';
      badge.removeAttribute('style');
    }

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
        await this.reorganizeAllTasks(false);
      }
    }

    this.setStorageStatus(schedRes.isLocal || taskRes?.isLocal ? 'local' : 'cloud');
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
    if (!modal) return;
    this.state.modalTrigger = document.activeElement;
    this.state.activeModalId = modalId;
    const dialog = modal.querySelector('.modal-container');
    const title = modal.querySelector('.modal-title');
    if (dialog) {
      dialog.setAttribute('role', 'dialog');
      dialog.setAttribute('aria-modal', 'true');
      dialog.setAttribute('aria-label', title?.textContent?.trim() || 'Diálogo');
      dialog.setAttribute('tabindex', '-1');
    }
    modal.classList.add('active');
    const firstFocusable = modal.querySelector(
      'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), a[href]'
    );
    (firstFocusable || dialog)?.focus();
  },

  closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (!modal) return;
    modal.classList.remove('active');
    if (modalId === 'taskModal') {
      this.state.editingTaskId = null;
      const title = document.getElementById('taskModalTitle');
      const submit = document.getElementById('btnSubmitTask');
      if (title) title.textContent = '⚡ Organizar Tarea / Tiempo de Estudio';
      if (submit) submit.textContent = '⚡ Organizar Automáticamente';
    }
    if (modalId === 'uploadModal' && this.state.pendingImport) this.cancelExcelImport();
    if (this.state.activeModalId === modalId) {
      this.state.activeModalId = null;
      if (this.state.modalTrigger?.isConnected) this.state.modalTrigger.focus();
      else if (modalId === 'taskModal') document.getElementById('btnQuickAddTask')?.focus();
      else document.getElementById('btnOpenTaskModal')?.focus();
      this.state.modalTrigger = null;
    }
  },

  escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, char => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    })[char]);
  },

  setStorageStatus(mode) {
    const status = document.getElementById('storageStatus');
    if (!status) return;
    if (mode === 'cloud' &&
        (SupabaseService.getLocalTasks().length > 0 || localStorage.getItem('cronosys_schedule'))) {
      mode = 'local';
    }
    const labels = {
      cloud: 'Sincronizado en la nube',
      local: 'Solo en este dispositivo',
      preview: 'Demo: no se guardan cambios'
    };
    status.textContent = labels[mode] || 'Aún no guardado';
    status.dataset.mode = mode;
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

    toast.innerHTML = iconSvg;
    const text = document.createElement('span');
    text.textContent = message;
    toast.appendChild(text);
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
