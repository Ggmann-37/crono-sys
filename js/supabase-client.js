/**
 * Crono-SYS - Cliente de Supabase & Gestión de Estado
 * Integra Supabase Auth y base de datos PostgreSQL con RLS (Row Level Security)
 * Incluye fallback automático a almacenamiento local (LocalStorage)
 */

const SUPABASE_CONFIG = {
  url: 'https://hplvbifzybthoclneumn.supabase.co',
  anonKey: 'sb_publishable_FEyFH-AVquoJFfYoO40UBw_GaryX6C4'
};

const SupabaseService = {
  client: null,
  currentUser: null,
  isInitialized: false,

  /**
   * Inicializa el cliente oficial de Supabase
   */
  init() {
    try {
      if (window.supabase && typeof window.supabase.createClient === 'function') {
        this.client = window.supabase.createClient(
          SUPABASE_CONFIG.url,
          SUPABASE_CONFIG.anonKey,
          {
            auth: {
              persistSession: true,
              autoRefreshToken: true
            }
          }
        );
        this.isInitialized = true;
        console.log('Crono-SYS: Supabase Client inicializado correctamente.');
      } else {
        console.warn('Crono-SYS: Librería de Supabase no disponible en window.');
      }
    } catch (err) {
      console.error('Error al inicializar Supabase:', err);
    }
  },

  /**
   * Obtiene la sesión actual
   */
  async getSession() {
    if (!this.client) return null;
    try {
      const { data, error } = await this.client.auth.getSession();
      if (error) throw error;
      this.currentUser = data.session ? data.session.user : null;
      return data.session;
    } catch (err) {
      console.warn('No se pudo recuperar la sesión de Supabase:', err.message);
      return null;
    }
  },

  /**
   * Registro con correo y contraseña
   */
  async signUp(email, password, fullName = '', captchaToken = null) {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const { data, error } = await this.client.auth.signUp({
      email,
      password,
      options: {
        data: { full_name: fullName },
        captchaToken
      }
    });
    if (error) throw error;
    this.currentUser = data.session ? data.user : null;
    return data;
  },

  /**
   * Reenvía el correo de confirmación para un registro pendiente
   */
  async resendConfirmation(email) {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const { error } = await this.client.auth.resend({
      type: 'signup',
      email
    });
    if (error) throw error;
  },

  /**
   * Inicio de sesión con correo y contraseña
   */
  async signIn(email, password, captchaToken = null) {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const { data, error } = await this.client.auth.signInWithPassword({
      email,
      password,
      options: { captchaToken }
    });
    if (error) throw error;
    this.currentUser = data.user;
    return data;
  },

  /**
   * Inicio de sesión o registro mediante el proveedor Google
   */
  async signInWithGoogle() {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const redirectUrl = `${window.location.origin}${window.location.pathname}`;
    const { error } = await this.client.auth.signInWithOAuth({
      provider: 'google',
      options: {
        redirectTo: redirectUrl,
        queryParams: {
          prompt: 'select_account'
        }
      }
    });
    if (error) throw error;
  },

  /**
   * Inicio de sesión o registro mediante el proveedor GitHub
   */
  async signInWithGitHub() {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const redirectUrl = `${window.location.origin}${window.location.pathname}`;
    const { error } = await this.client.auth.signInWithOAuth({
      provider: 'github',
      options: {
        redirectTo: redirectUrl
      }
    });
    if (error) throw error;
  },

  /**
   * Cierre de sesión
   */
  async signOut() {
    if (!this.client) return;
    const { error } = await this.client.auth.signOut();
    this.currentUser = null;
    if (error) throw error;
  },

  /**
   * Actualiza la contraseña del usuario actual
   */
  async updatePassword(newPassword) {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const { data, error } = await this.client.auth.updateUser({ password: newPassword });
    if (error) throw error;
    return data;
  },

  /**
   * Actualiza el correo electrónico del usuario actual
   */
  async updateEmail(newEmail) {
    if (!this.client) throw new Error('Cliente Supabase no disponible');
    const { data, error } = await this.client.auth.updateUser({ email: newEmail });
    if (error) throw error;
    return data;
  },

  /**
   * Elimina la cuenta y todos sus datos en Supabase y localmente
   */
  async deleteAccount() {
    if (this.currentUser && this.client) {
      const userId = this.currentUser.id;
      try {
        await this.client.from('tasks').delete().eq('user_id', userId);
        await this.client.from('schedule_blocks').delete().eq('user_id', userId);
        await this.client.from('schedules').delete().eq('user_id', userId);
        await this.client.from('profiles').delete().eq('id', userId);
      } catch (e) {
        console.warn('Error al limpiar registros de usuario:', e);
      }
      await this.signOut();
    }
    // Limpieza de datos locales
    localStorage.removeItem('cronosys_schedule');
    localStorage.removeItem('cronosys_tasks');
  },

  /**
   * Escucha cambios en el estado de autenticación
   */
  onAuthStateChange(callback) {
    if (!this.client) return;
    this.client.auth.onAuthStateChange((event, session) => {
      this.currentUser = session ? session.user : null;
      callback(event, session);
    });
  },

  // ==========================================
  // OPERACIONES DE HORARIO (SCHEDULES) CON RLS
  // ==========================================

  /**
   * Guarda un nuevo horario importado y sus bloques en Supabase
   * Si no está logueado, guarda en localStorage
   */
  async saveSchedule(title, fileName, blocks) {
    // Modo local / Fallback
    if (!this.currentUser || !this.client) {
      const localSchedule = {
        id: 'local-schedule-' + Date.now(),
        title: title || 'Mi Horario Semanal',
        fileName: fileName,
        blocks: blocks,
        updatedAt: new Date().toISOString()
      };
      localStorage.setItem('cronosys_schedule', JSON.stringify(localSchedule));
      return { success: true, schedule: localSchedule, isLocal: true };
    }

    try {
      const userId = this.currentUser.id;

      // 1. Desactivar horarios anteriores
      await this.client
        .from('schedules')
        .update({ is_active: false })
        .eq('user_id', userId);

      // 2. Insertar nuevo horario activo
      const { data: scheduleData, error: schedError } = await this.client
        .from('schedules')
        .insert({
          user_id: userId,
          title: title || 'Horario Importado',
          file_name: fileName,
          is_active: true
        })
        .select()
        .single();

      if (schedError) throw schedError;

      const scheduleId = scheduleData.id;

      // 3. Preparar bloques para inserción
      const blocksToInsert = blocks.map(b => ({
        schedule_id: scheduleId,
        user_id: userId,
        day_of_week: b.day,
        start_time: b.startTime + ':00',
        end_time: b.endTime + ':00',
        title: b.title,
        block_type: b.type || 'base',
        color: b.color || '#2563eb',
        notes: b.notes || null
      }));

      // Insertar bloques en lotes
      if (blocksToInsert.length > 0) {
        const { error: blockError } = await this.client
          .from('schedule_blocks')
          .insert(blocksToInsert);

        if (blockError) throw blockError;
      }

      return { success: true, schedule: scheduleData, isLocal: false };
    } catch (err) {
      console.warn('Error al guardar en Supabase, usando almacenamiento local:', err);
      // Fallback a localStorage
      const localSchedule = {
        id: 'local-schedule-' + Date.now(),
        title: title,
        fileName: fileName,
        blocks: blocks,
        updatedAt: new Date().toISOString()
      };
      localStorage.setItem('cronosys_schedule', JSON.stringify(localSchedule));
      return { success: true, schedule: localSchedule, isLocal: true, error: err.message };
    }
  },

  /**
   * Carga el horario activo del usuario (de Supabase o de LocalStorage)
   */
  async loadActiveSchedule() {
    if (!this.currentUser || !this.client) {
      const raw = localStorage.getItem('cronosys_schedule');
      if (raw) {
        try {
          return { success: true, data: JSON.parse(raw), isLocal: true };
        } catch (e) {
          return { success: false, data: null };
        }
      }
      return { success: false, data: null };
    }

    try {
      const userId = this.currentUser.id;

      // 1. Obtener horario activo
      const { data: schedules, error: schedErr } = await this.client
        .from('schedules')
        .select('*')
        .eq('user_id', userId)
        .eq('is_active', true)
        .order('created_at', { ascending: false })
        .limit(1);

      if (schedErr) throw schedErr;
      if (!schedules || schedules.length === 0) {
        return { success: true, data: null, isLocal: false };
      }

      const activeSchedule = schedules[0];

      // 2. Obtener bloques asociados
      const { data: blocks, error: blockErr } = await this.client
        .from('schedule_blocks')
        .select('*')
        .eq('schedule_id', activeSchedule.id)
        .order('start_time', { ascending: true });

      if (blockErr) throw blockErr;

      // Mapear bloques al formato de la aplicación
      const formattedBlocks = (blocks || []).map(b => ({
        id: b.id,
        day: b.day_of_week,
        startTime: b.start_time.substring(0, 5),
        endTime: b.end_time.substring(0, 5),
        title: b.title,
        type: b.block_type,
        color: b.color,
        notes: b.notes,
        taskId: b.task_id
      }));

      return {
        success: true,
        data: {
          id: activeSchedule.id,
          title: activeSchedule.title,
          fileName: activeSchedule.file_name,
          blocks: formattedBlocks
        },
        isLocal: false
      };
    } catch (err) {
      console.warn('Error al cargar de Supabase, consultando local:', err);
      const raw = localStorage.getItem('cronosys_schedule');
      return {
        success: true,
        data: raw ? JSON.parse(raw) : null,
        isLocal: true,
        error: err.message
      };
    }
  },

  // ==========================================
  // OPERACIONES DE TAREAS Y ESTUDIO CON RLS
  // ==========================================

  /**
   * Guarda o actualiza una tarea en Supabase / Local
   */
  async saveTask(task) {
    if (!this.currentUser || !this.client) {
      const tasks = this.getLocalTasks();
      const index = tasks.findIndex(t => t.id === task.id);
      if (index >= 0) {
        tasks[index] = task;
      } else {
        tasks.push(task);
      }
      localStorage.setItem('cronosys_tasks', JSON.stringify(tasks));
      return { success: true, task, isLocal: true };
    }

    try {
      const userId = this.currentUser.id;
      const { data, error } = await this.client
        .from('tasks')
        .upsert({
          id: task.id.includes('local') ? undefined : task.id,
          user_id: userId,
          title: task.title,
          category: task.category,
          duration_minutes: task.durationMinutes,
          preferred_days: task.preferredDays,
          time_preference: task.timePreference,
          custom_start_time: task.customStart ? task.customStart + ':00' : null,
          custom_end_time: task.customEnd ? task.customEnd + ':00' : null,
          max_session_minutes: task.maxSessionMinutes,
          rest_between_minutes: task.restBetweenMinutes,
          priority: task.priority,
          status: task.status || 'scheduled'
        })
        .select()
        .single();

      if (error) throw error;
      return { success: true, task: data, isLocal: false };
    } catch (err) {
      console.warn('Fallo en Supabase saveTask, guardando en local:', err);
      const tasks = this.getLocalTasks();
      tasks.push(task);
      localStorage.setItem('cronosys_tasks', JSON.stringify(tasks));
      return { success: true, task, isLocal: true, error: err.message };
    }
  },

  /**
   * Carga la lista de tareas del usuario
   */
  async loadTasks() {
    if (!this.currentUser || !this.client) {
      return { success: true, tasks: this.getLocalTasks(), isLocal: true };
    }

    try {
      const { data, error } = await this.client
        .from('tasks')
        .select('*')
        .eq('user_id', this.currentUser.id)
        .order('created_at', { ascending: false });

      if (error) throw error;

      const formatted = (data || []).map(t => ({
        id: t.id,
        title: t.title,
        category: t.category,
        durationMinutes: t.duration_minutes,
        preferredDays: t.preferred_days,
        timePreference: t.time_preference,
        customStart: t.custom_start_time ? t.custom_start_time.substring(0, 5) : null,
        customEnd: t.custom_end_time ? t.custom_end_time.substring(0, 5) : null,
        maxSessionMinutes: t.max_session_minutes,
        restBetweenMinutes: t.rest_between_minutes,
        priority: t.priority,
        status: t.status
      }));

      return { success: true, tasks: formatted, isLocal: false };
    } catch (err) {
      console.warn('Error al cargar tareas de Supabase:', err);
      return { success: true, tasks: this.getLocalTasks(), isLocal: true };
    }
  },

  /**
   * Elimina una tarea
   */
  async deleteTask(taskId) {
    if (!this.currentUser || !this.client) {
      let tasks = this.getLocalTasks();
      tasks = tasks.filter(t => t.id !== taskId);
      localStorage.setItem('cronosys_tasks', JSON.stringify(tasks));
      return { success: true };
    }

    try {
      const { error } = await this.client
        .from('tasks')
        .delete()
        .eq('id', taskId)
        .eq('user_id', this.currentUser.id);

      if (error) throw error;
      return { success: true };
    } catch (err) {
      console.error('Error al borrar tarea en Supabase:', err);
      return { success: false, error: err.message };
    }
  },

  /**
   * Utilidad para leer tareas de LocalStorage
   */
  getLocalTasks() {
    const raw = localStorage.getItem('cronosys_tasks');
    if (!raw) return [];
    try {
      return JSON.parse(raw);
    } catch (e) {
      return [];
    }
  }
};

window.SupabaseService = SupabaseService;
