/**
 * Crono-SYS - Motor de Organización Inteligente de Tareas y Estudios
 * Organiza automáticamente los tiempos de estudio, tareas y descansos
 * buscando huecos libres en el horario semanal respetando las restricciones del usuario:
 * - Días preferidos
 * - Franja horaria (mañana, tarde, noche, todo el día o rango personalizado)
 * - Duración máxima de bloque continuo (para no saturar)
 * - Descanso entre sesiones / tras el trabajo
 */

const SmartScheduler = {
  DAYS_ORDER: ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'],

  TIME_PREFERENCES: {
    any: { start: '07:00', end: '23:00' },
    morning: { start: '08:00', end: '14:00' },
    afternoon: { start: '14:00', end: '20:00' },
    evening: { start: '20:00', end: '23:30' },
    workday: { start: '09:00', end: '18:00' }
  },

  /**
   * Convierte "HH:MM" a minutos desde las 00:00
   */
  timeToMinutes(timeStr) {
    if (!timeStr) return 0;
    const [h, m] = timeStr.split(':').map(Number);
    return h * 60 + (m || 0);
  },

  /**
   * Convierte minutos desde las 00:00 a formato "HH:MM"
   */
  minutesToTime(totalMin) {
    const hours = Math.floor(totalMin / 60) % 24;
    const mins = totalMin % 60;
    return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}`;
  },

  /**
   * Divide el tiempo total de una tarea en sesiones lógicas según la duración máxima
   * @param {number} totalMinutes 
   * @param {number} maxSessionMinutes 
   * @returns {number[]} Array con los minutos de cada sesión (ej. [60, 60])
   */
  calculateSessionChunks(totalMinutes, maxSessionMinutes) {
    const maxSession = maxSessionMinutes > 0 ? maxSessionMinutes : 60;
    const chunks = [];
    let remaining = totalMinutes;

    while (remaining > 0) {
      if (remaining <= maxSession) {
        chunks.push(remaining);
        remaining = 0;
      } else {
        chunks.push(maxSession);
        remaining -= maxSession;
      }
    }

    return chunks;
  },

  /**
   * Construye un mapa de ocupación por cada día de la semana
   * @param {Array} existingBlocks - Bloques base del Excel y tareas previas
   * @returns {Object} { lunes: [ [startMin, endMin], ... ], martes: [...] }
   */
  buildOccupiedMap(existingBlocks) {
    const map = {};
    this.DAYS_ORDER.forEach(d => {
      map[d] = [];
    });

    existingBlocks.forEach(b => {
      if (map[b.day]) {
        const start = this.timeToMinutes(b.startTime);
        let end = this.timeToMinutes(b.endTime);
        // Manejar medianoche
        if (end <= start) end = 1440;
        map[b.day].push({ start, end, title: b.title, type: b.type });
      }
    });

    // Ordenar y fusionar intervalos solapados
    this.DAYS_ORDER.forEach(d => {
      map[d].sort((a, b) => a.start - b.start);
    });

    return map;
  },

  /**
   * Encuentra huecos libres disponibles en un día dado dentro del rango deseado
   * @param {Array} dayOccupied - Intervalos ocupados en ese día
   * @param {number} rangeStartMin - Inicio de la franja (en min)
   * @param {number} rangeEndMin - Fin de la franja (en min)
   * @returns {Array} [ { start: min, end: min, duration: min } ]
   */
  findFreeIntervals(dayOccupied, rangeStartMin, rangeEndMin) {
    const freeIntervals = [];
    let currentPointer = rangeStartMin;

    for (let i = 0; i < dayOccupied.length; i++) {
      const occ = dayOccupied[i];

      // Si el bloque ocupado termina antes de nuestro rango, lo saltamos
      if (occ.end <= currentPointer) continue;

      // Si el bloque ocupado empieza después del final de nuestro rango, terminamos
      if (occ.start >= rangeEndMin) break;

      // Si hay espacio entre nuestro puntero y el inicio del bloque ocupado
      if (occ.start > currentPointer) {
        const freeStart = currentPointer;
        const freeEnd = Math.min(occ.start, rangeEndMin);
        if (freeEnd - freeStart >= 15) { // Al menos 15 minutos de hueco
          freeIntervals.push({
            start: freeStart,
            end: freeEnd,
            duration: freeEnd - freeStart
          });
        }
      }

      currentPointer = Math.max(currentPointer, occ.end);
    }

    // Comprobar si queda hueco después del último bloque ocupado
    if (currentPointer < rangeEndMin) {
      if (rangeEndMin - currentPointer >= 15) {
        freeIntervals.push({
          start: currentPointer,
          end: rangeEndMin,
          duration: rangeEndMin - currentPointer
        });
      }
    }

    return freeIntervals;
  },

  /**
   * Organiza una tarea en el horario existente
   * @param {Object} task - Datos de la tarea a organizar
   * @param {Array} currentBlocks - Bloques de horario actuales
   * @returns {Object} Resultado con los bloques nuevos generados y el estado
   */
  scheduleTask(task, currentBlocks) {
    const occupiedMap = this.buildOccupiedMap(currentBlocks);
    const preferredDays = (task.preferredDays && task.preferredDays.length > 0)
      ? task.preferredDays
      : ['lunes', 'martes', 'miercoles', 'jueves', 'viernes'];

    // Determinar la franja horaria permitida
    let rangeStart = '08:00';
    let rangeEnd = '22:00';

    if (task.timePreference === 'custom' && task.customStart && task.customEnd) {
      rangeStart = task.customStart;
      rangeEnd = task.customEnd;
    } else if (this.TIME_PREFERENCES[task.timePreference]) {
      rangeStart = this.TIME_PREFERENCES[task.timePreference].start;
      rangeEnd = this.TIME_PREFERENCES[task.timePreference].end;
    }

    const rangeStartMin = this.timeToMinutes(rangeStart);
    const rangeEndMin = this.timeToMinutes(rangeEnd);

    const sessionChunks = this.calculateSessionChunks(task.durationMinutes, task.maxSessionMinutes || 60);
    const restBetweenMin = Number(task.restBetweenMinutes) || 0;

    const newBlocks = [];
    const newBreaks = [];
    const unplacedChunks = [];

    // Priorizar distribución equitativa entre los días preferidos
    let dayIndex = 0;

    for (let c = 0; c < sessionChunks.length; c++) {
      const chunkDuration = sessionChunks[c];
      const breakDuration = c < sessionChunks.length - 1 ? restBetweenMin : 0;
      const requiredSpan = chunkDuration + breakDuration;
      let placed = false;

      // Intentar ubicar en los días preferidos dando vueltas
      for (let attempt = 0; attempt < preferredDays.length; attempt++) {
        const day = preferredDays[(dayIndex + attempt) % preferredDays.length];
        const dayOccupied = occupiedMap[day] || [];
        const freeIntervals = this.findFreeIntervals(dayOccupied, rangeStartMin, rangeEndMin);

        // Buscar el primer intervalo libre donde quepa la sesión (+ descanso opcional)
        for (let i = 0; i < freeIntervals.length; i++) {
          const interval = freeIntervals[i];
          if (interval.duration >= requiredSpan) {
            // Ubicar la sesión de estudio/tarea
            const sessionStartMin = interval.start;
            const sessionEndMin = sessionStartMin + chunkDuration;

            const sessionBlock = {
              id: 'task-block-' + task.id + '-' + c + '-' + Date.now(),
              taskId: task.id,
              scheduleId: task.scheduleId || null,
              day: day,
              startTime: this.minutesToTime(sessionStartMin),
              endTime: this.minutesToTime(sessionEndMin),
              title: task.title,
              type: task.category === 'estudio' ? 'study' : 'task',
              category: task.category || 'estudio',
              color: task.color || (task.category === 'estudio' ? '#0284c7' : '#0ea5e9'),
              notes: `Sesión ${c + 1} de ${sessionChunks.length} (${chunkDuration} min)`
            };

            newBlocks.push(sessionBlock);

            // Registrar ocupación en el mapa
            occupiedMap[day].push({
              start: sessionStartMin,
              end: sessionEndMin,
              title: sessionBlock.title,
              type: sessionBlock.type
            });

            // Si hay descanso configurado y cabe después de la sesión
            if (breakDuration > 0) {
              const breakBlock = {
                id: 'break-' + task.id + '-' + c + '-' + Date.now(),
                taskId: task.id,
                day: day,
                startTime: this.minutesToTime(sessionEndMin),
                endTime: this.minutesToTime(sessionEndMin + breakDuration),
                title: `Descanso (${breakDuration}m)`,
                type: 'break',
                color: '#10b981', // Verde menta
                notes: 'Pausa para recuperar foco y desconectar'
              };
              newBreaks.push(breakBlock);

              // Registrar el descanso en el mapa de ocupación para no meter otra tarea encima
              occupiedMap[day].push({
                start: sessionEndMin,
                end: sessionEndMin + breakDuration,
                title: breakBlock.title,
                type: 'break'
              });
            }

            // Reordenar ocupación de ese día
            occupiedMap[day].sort((a, b) => a.start - b.start);

            placed = true;
            // Avanzar el índice de día para balancear la carga entre varios días
            dayIndex = (dayIndex + attempt + 1) % preferredDays.length;
            break;
          }
        }

        if (placed) break;
      }

      if (!placed) {
        unplacedChunks.push(chunkDuration);
      }
    }

    const success = unplacedChunks.length === 0;

    return {
      success: success,
      task: task,
      placedBlocks: newBlocks,
      breakBlocks: newBreaks,
      totalPlacedMinutes: sessionChunks.reduce((a, b) => a + b, 0) - unplacedChunks.reduce((a, b) => a + b, 0),
      unplacedMinutes: unplacedChunks.reduce((a, b) => a + b, 0),
      message: success
        ? `Se organizaron con éxito ${newBlocks.length} sesiones de "${task.title}".`
        : `No se encontró suficiente tiempo libre para ${unplacedChunks.reduce((a, b) => a + b, 0)} minutos de "${task.title}". Sugerencia: amplía los días preferidos o el rango horario.`
    };
  },

  /**
   * Reorganiza todas las tareas pendientes sobre la base del horario
   * @param {Array} tasks - Lista de tareas a programar
   * @param {Array} baseBlocks - Bloques base del Excel
   */
  reorganizeAll(tasks, baseBlocks) {
    // Conservar solo bloques base del Excel
    let currentBlocks = [...baseBlocks];
    const allTaskBlocks = [];
    const allBreaks = [];
    const report = [];

    // Ordenar tareas por prioridad (high > medium > low)
    const priorityWeight = { high: 3, medium: 2, low: 1 };
    const sortedTasks = [...tasks].sort((a, b) => {
      const pA = priorityWeight[a.priority] || 2;
      const pB = priorityWeight[b.priority] || 2;
      return pB - pA;
    });

    sortedTasks.forEach(task => {
      const res = this.scheduleTask(task, currentBlocks);
      if (res.placedBlocks.length > 0) {
        allTaskBlocks.push(...res.placedBlocks);
        allBreaks.push(...res.breakBlocks);
        currentBlocks.push(...res.placedBlocks, ...res.breakBlocks);
      }
      report.push({
        taskId: task.id,
        title: task.title,
        success: res.success,
        placedSessions: res.placedBlocks.length,
        unplacedMinutes: res.unplacedMinutes
      });
    });

    return {
      allBlocks: currentBlocks,
      taskBlocks: allTaskBlocks,
      breakBlocks: allBreaks,
      report: report
    };
  }
};

window.SmartScheduler = SmartScheduler;
