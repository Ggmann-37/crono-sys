/**
 * Crono-SYS - Renderizador de Calendario Estilo Microsoft Outlook / Teams
 * Diseñado con estética minimalista, líneas limpias, azul claro y blanco
 */

const CalendarView = {
  container: null,
  currentBlocks: [],
  hourHeight: 52, // Altura en píxeles de cada hora
  activeDays: ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'],
  dayLabels: {
    lunes: 'Lunes',
    martes: 'Martes',
    miercoles: 'Miércoles',
    jueves: 'Jueves',
    viernes: 'Viernes',
    sabado: 'Sábado',
    domingo: 'Domingo'
  },
  onlyWorkdays: false,
  mobileDayOnly: false,
  selectedMobileDay: null,

  /**
   * Inicializa la vista de calendario
   */
  init(containerElement) {
    this.container = containerElement;
    this.render();
    this.setupAutoScroll();
    this.startCurrentTimeIndicator();
  },

  /**
   * Actualiza los bloques a mostrar y redibuja
   */
  setBlocks(blocks) {
    this.currentBlocks = blocks || [];
    this.render();
  },

  /**
   * Alterna entre semana completa (7 días) y días laborales (5 días)
   */
  toggleWorkdaysOnly(workdaysOnly) {
    this.onlyWorkdays = workdaysOnly;
    this.render();
  },

  setMobileDayView(enabled) {
    if (this.mobileDayOnly === enabled) return;
    this.mobileDayOnly = enabled;
    if (!this.selectedMobileDay) this.selectedMobileDay = this.getTodayKey();
    this.render();
  },

  moveMobileDay(offset) {
    const currentIndex = this.activeDays.indexOf(this.selectedMobileDay || this.getTodayKey());
    const nextIndex = (currentIndex + offset + this.activeDays.length) % this.activeDays.length;
    this.selectedMobileDay = this.activeDays[nextIndex];
    this.render();
    this.setupAutoScroll();
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

  /**
   * Determina qué día de la semana es hoy (normalizado)
   */
  getTodayKey() {
    const jsDay = new Date().getDay(); // 0 es domingo, 1 lunes...
    const map = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
    return map[jsDay];
  },

  /**
   * Convierte "HH:MM" a minutos
   */
  timeToMinutes(timeStr) {
    if (!timeStr) return 0;
    const parts = timeStr.split(':').map(Number);
    return parts[0] * 60 + (parts[1] || 0);
  },

  /**
   * Renderiza la estructura del calendario estilo Outlook / Teams
   */
  render() {
    if (!this.container) return;

    const todayKey = this.getTodayKey();
    const visibleDays = this.mobileDayOnly
      ? [this.selectedMobileDay || todayKey]
      : this.onlyWorkdays
        ? ['lunes', 'martes', 'miercoles', 'jueves', 'viernes']
        : this.activeDays;
    const calendarTitle = document.getElementById('calendarTitle');
    if (calendarTitle) {
      calendarTitle.textContent = this.mobileDayOnly
        ? `Horario · ${this.dayLabels[visibleDays[0]]}`
        : 'Horario Semanal';
    }

    let html = `
      <div class="teams-calendar-wrapper">
        <!-- Cabecera de días (Sticky Top) -->
        <div class="calendar-header-row">
          <div class="calendar-time-corner">
            <span class="timezone-label">GMT+1</span>
          </div>
          <div class="calendar-days-header" style="grid-template-columns: repeat(${visibleDays.length}, 1fr);">
    `;

    visibleDays.forEach(day => {
      const isToday = day === todayKey;
      html += `
        <div class="day-column-header ${isToday ? 'is-today' : ''}" data-day="${day}">
          <span class="day-name">${this.dayLabels[day]}</span>
          <span class="day-badge ${isToday ? 'badge-today' : ''}">
            ${isToday ? 'Hoy' : ''}
          </span>
        </div>
      `;
    });

    html += `
          </div>
        </div>

        <div class="mobile-calendar-navigation" aria-label="Navegar por los días">
          <button type="button" data-mobile-day="-1" aria-label="Día anterior">&larr;</button>
          <span>${this.dayLabels[visibleDays[0]]}</span>
          <button type="button" data-mobile-day="1" aria-label="Día siguiente">&rarr;</button>
        </div>

        <!-- Cuerpo del calendario con scroll -->
        <div class="calendar-body-scroll" id="calendarScrollContainer">
          <div class="calendar-body-grid">
            
            <!-- Columna lateral de Horas (00:00 a 23:00) -->
            <div class="calendar-time-scale">
    `;

    for (let h = 0; h < 24; h++) {
      const timeLabel = `${h.toString().padStart(2, '0')}:00`;
      html += `
        <div class="time-cell" style="height: ${this.hourHeight}px;">
          <span class="time-text">${timeLabel}</span>
        </div>
      `;
    }

    html += `
            </div>

            <!-- Grilla de columnas por cada día -->
            <div class="calendar-days-grid" style="grid-template-columns: repeat(${visibleDays.length}, 1fr);">
    `;

    visibleDays.forEach(day => {
      const isToday = day === todayKey;
      html += `
        <div class="day-track-column ${isToday ? 'is-today-track' : ''}" data-day="${day}" style="height: ${24 * this.hourHeight}px;">
          <!-- Líneas de fondo de cada hora -->
          <div class="day-grid-lines">
      `;

      for (let h = 0; h < 24; h++) {
        html += `
          <div class="hour-slot" data-day="${day}" data-hour="${h}" role="button"
               tabindex="${h === 8 && day === visibleDays[0] ? '0' : '-1'}"
               aria-label="Añadir tarea ${this.dayLabels[day]} a las ${h.toString().padStart(2, '0')}:00"
               style="height: ${this.hourHeight}px;"
               title="Añadir tarea a las ${h.toString().padStart(2, '0')}:00">
            <div class="half-hour-line"></div>
          </div>
        `;
      }

      html += `
          </div>

          <!-- Contenedor de Bloques de Eventos para este día -->
          <div class="day-events-layer" id="events-${day}">
      `;

      // Filtrar y renderizar bloques de este día
      const dayBlocks = this.currentBlocks.filter(b => b.day === day);
      dayBlocks.forEach(b => {
        html += this.renderEventBlock(b);
      });

      // Indicador de hora actual si es hoy
      if (isToday) {
        html += `<div class="current-time-marker" id="currentTimeMarker"></div>`;
      }

      html += `
          </div>
        </div>
      `;
    });

    html += `
            </div>
          </div>
        </div>
      </div>
    `;

    this.container.innerHTML = html;
    this.attachEvents();
    this.updateCurrentTimeMarker();
  },

  /**
   * Genera el HTML de un bloque de evento
   */
  renderEventBlock(block) {
    const startMin = this.timeToMinutes(block.startTime);
    let endMin = this.timeToMinutes(block.endTime);
    if (endMin <= startMin) endMin = 1440; // Medianoche

    const topPx = (startMin / 60) * this.hourHeight;
    const heightPx = Math.max(((endMin - startMin) / 60) * this.hourHeight - 3, 24);

    let typeClass = 'event-base';
    let typeBadge = 'Horario Base';
    let iconSvg = `
      <svg class="block-icon" viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
        <line x1="16" y1="2" x2="16" y2="6"></line>
        <line x1="8" y1="2" x2="8" y2="6"></line>
        <line x1="3" y1="10" x2="21" y2="10"></line>
      </svg>
    `;

    if (block.type === 'study') {
      typeClass = 'event-study';
      typeBadge = 'Estudio';
      iconSvg = `
        <svg class="block-icon" viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
          <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path>
        </svg>
      `;
    } else if (block.type === 'task') {
      typeClass = 'event-task';
      typeBadge = 'Tarea';
      iconSvg = `
        <svg class="block-icon" viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M9 11l3 3L22 4"></path>
          <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path>
        </svg>
      `;
    } else if (block.type === 'break') {
      typeClass = 'event-break';
      typeBadge = 'Descanso';
      iconSvg = `
        <svg class="block-icon" viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M18 8h1a4 4 0 0 1 0 8h-1"></path>
          <path d="M2 8h16v9a4 4 0 0 1-4 4H6a4 4 0 0 1-4-4V8z"></path>
          <line x1="6" y1="1" x2="6" y2="4"></line>
          <line x1="10" y1="1" x2="10" y2="4"></line>
          <line x1="14" y1="1" x2="14" y2="4"></line>
        </svg>
      `;
    }

    return `
      <div class="calendar-event-card ${typeClass}" 
           style="top: ${topPx}px; height: ${heightPx}px;"
           data-id="${block.id}"
           data-title="${encodeURIComponent(block.title)}"
           data-time="${block.startTime} - ${block.endTime}"
           data-type="${block.type}"
           data-day="${block.day}"
           role="button"
           aria-label="${this.escapeHtml(`${typeBadge}: ${block.title}, ${this.dayLabels[block.day]}, ${block.startTime} a ${block.endTime}`)}"
           tabindex="0">
        <div class="event-card-content">
          <div class="event-card-header">
            ${iconSvg}
            <span class="event-title" title="${this.escapeHtml(block.title)}">${this.escapeHtml(block.title)}</span>
          </div>
          <div class="event-time-range">${block.startTime} - ${block.endTime}</div>
          ${block.notes ? `<div class="event-notes-snippet">${this.escapeHtml(block.notes)}</div>` : ''}
        </div>
      </div>
    `;
  },

  /**
   * Conecta los eventos de clic y creación
   */
  attachEvents() {
    // Clic en bloque para ver detalle
    this.container.querySelectorAll('.calendar-event-card').forEach(el => {
      const openDetails = (event) => {
        event.stopPropagation();
        const blockId = el.getAttribute('data-id');
        const block = this.currentBlocks.find(b => b.id === blockId);
        if (block && window.App && typeof window.App.showBlockDetailModal === 'function') {
          window.App.showBlockDetailModal(block);
        }
      };
      el.addEventListener('click', (e) => {
        openDetails(e);
      });
      el.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          openDetails(event);
        }
      });
    });

    // Clic en celda vacía para crear evento rápido
    this.container.querySelectorAll('.hour-slot').forEach(el => {
      const openTask = () => {
        const day = el.getAttribute('data-day');
        const hour = parseInt(el.getAttribute('data-hour'), 10);
        if (window.App && typeof window.App.openNewTaskModalWithSlot === 'function') {
          window.App.openNewTaskModalWithSlot(day, hour);
        }
      };
      el.addEventListener('click', openTask);
      el.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          openTask();
          return;
        }
        if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) return;
        event.preventDefault();
        const slots = [...this.container.querySelectorAll('.hour-slot')];
        const currentDay = el.dataset.day;
        const dayIndex = this.activeDays.indexOf(currentDay);
        const hour = Number(el.dataset.hour);
        let nextDay = currentDay;
        let nextHour = hour;
        if (event.key === 'ArrowUp') nextHour = Math.max(0, hour - 1);
        if (event.key === 'ArrowDown') nextHour = Math.min(23, hour + 1);
        if (event.key === 'ArrowLeft') nextDay = this.activeDays[Math.max(0, dayIndex - 1)];
        if (event.key === 'ArrowRight') nextDay = this.activeDays[Math.min(this.activeDays.length - 1, dayIndex + 1)];
        slots.forEach(slot => { slot.tabIndex = -1; });
        const nextSlot = this.container.querySelector(`.hour-slot[data-day="${nextDay}"][data-hour="${nextHour}"]`);
        if (nextSlot) {
          nextSlot.tabIndex = 0;
          nextSlot.focus();
        }
      });
    });

    this.container.querySelectorAll('[data-mobile-day]').forEach(button => {
      button.addEventListener('click', () => this.moveMobileDay(Number(button.dataset.mobileDay)));
    });
  },

  /**
   * Auto-desplaza la vista suavemente a la franja matutina (08:00) para evitar ver medianoche vacía
   */
  setupAutoScroll() {
    setTimeout(() => {
      const scrollBox = document.getElementById('calendarScrollContainer');
      if (scrollBox) {
        const todayKey = this.getTodayKey();
        const currentHour = this.selectedMobileDay === todayKey ? new Date().getHours() : 8;
        const targetHour = this.mobileDayOnly ? Math.max(7, currentHour - 2) : 7.5;
        scrollBox.scrollTop = targetHour * this.hourHeight;
      }
    }, 150);
  },

  /**
   * Mantiene actualizada la línea roja de hora actual
   */
  startCurrentTimeIndicator() {
    setInterval(() => {
      this.updateCurrentTimeMarker();
    }, 60000); // Cada minuto
  },

  updateCurrentTimeMarker() {
    const marker = document.getElementById('currentTimeMarker');
    if (!marker) return;

    const now = new Date();
    const currentMin = now.getHours() * 60 + now.getMinutes();
    const topPx = (currentMin / 60) * this.hourHeight;

    marker.style.top = `${topPx}px`;
    marker.style.display = 'block';
  }
};

window.CalendarView = CalendarView;
