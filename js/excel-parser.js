/**
 * Crono-SYS - Analizador de Archivos Excel con SheetJS
 * Valida de forma estricta el formato requerido:
 * Columnas: Hora/Día, lunes, martes, miercoles, jueves, viernes, sábado, domingo
 * Filas: 00:00:00 hasta 23:00:00
 */

const ExcelParser = {
  // Días esperados en orden normalizado (sin acentos, en minúsculas)
  EXPECTED_DAYS: ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'],

  /**
   * Normaliza un texto eliminando acentos, caracteres especiales y convirtiendo a minúsculas
   */
  normalizeText(text) {
    if (text === null || text === undefined) return '';
    return String(text)
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  },

  /**
   * Valida si la extensión del archivo es permitida (.xlsx o .xls)
   */
  isValidExtension(fileName) {
    if (!fileName) return false;
    const ext = fileName.toLowerCase().split('.').pop();
    return ext === 'xlsx' || ext === 'xls';
  },

  /**
   * Procesa un archivo Excel cargado mediante FileReader y SheetJS
   * @param {File} file - Archivo .xlsx o .xls
   * @returns {Promise<{valid: boolean, error?: string, schedule?: any[]}>}
   */
  async parseExcelFile(file) {
    // 1. Validar extensión de archivo
    if (!this.isValidExtension(file.name)) {
      return {
        valid: false,
        error: 'formato no valido'
      };
    }

    try {
      const data = await file.arrayBuffer();
      // Leer con SheetJS (XLSX global)
      if (typeof XLSX === 'undefined') {
        throw new Error('SheetJS (XLSX) no está cargado');
      }

      const workbook = XLSX.read(data, { type: 'array' });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        return { valid: false, error: 'formato no valido' };
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convertir a matriz 2D de celdas
      const rawRows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: ''
      });

      if (!rawRows || rawRows.length < 2) {
        return { valid: false, error: 'formato no valido' };
      }

      // 2. Validar Cabeceras (Fila 0)
      const headerRow = rawRows[0];
      const validation = this.validateHeader(headerRow);
      if (!validation.valid) {
        return { valid: false, error: 'formato no valido' };
      }

      // 3. Validar y procesar filas de horas
      const rowsData = rawRows.slice(1);
      const parseResult = this.extractScheduleBlocks(rowsData);

      if (!parseResult.valid) {
        return { valid: false, error: 'formato no valido' };
      }

      return {
        valid: true,
        fileName: file.name,
        blocks: parseResult.blocks,
        totalEntries: parseResult.blocks.length,
        rawRows: rowsData
      };
    } catch (err) {
      console.error('Error al procesar el archivo Excel:', err);
      return { valid: false, error: 'formato no valido' };
    }
  },

  /**
   * Valida estrictamente la cabecera:
   * Debe contener 'Hora/Día', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sábado', 'domingo'
   */
  validateHeader(headerRow) {
    if (!Array.isArray(headerRow) || headerRow.length < 8) {
      return { valid: false };
    }

    const firstCol = this.normalizeText(headerRow[0]);
    // Debe empezar por hora/dia o contener hora y dia
    if (!firstCol.includes('hora') && !firstCol.includes('dia')) {
      return { valid: false };
    }

    // Comprobar los 7 días en las columnas 1 a 7
    for (let i = 0; i < 7; i++) {
      const colValue = this.normalizeText(headerRow[i + 1]);
      const expectedDay = this.EXPECTED_DAYS[i];
      if (colValue !== expectedDay) {
        return { valid: false };
      }
    }

    return { valid: true };
  },

  /**
   * Normaliza formato de hora admitiendo múltiples variantes de Excel:
   * "08:00:00", "8:00:00", "08:00", "8:00", "08:00 AM", o número serial de Excel
   */
  normalizeTime(timeStr) {
    if (timeStr === null || timeStr === undefined || timeStr === '') return null;

    // Caso número decimal serial de Excel (ej. 0.375 -> 09:00)
    if (typeof timeStr === 'number' || (!isNaN(Number(timeStr)) && String(timeStr).includes('.'))) {
      const num = Number(timeStr);
      if (num >= 0 && num < 1) {
        const totalMin = Math.round(num * 24 * 60);
        const h = Math.floor(totalMin / 60) % 24;
        const m = totalMin % 60;
        return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
      }
    }

    const str = String(timeStr).trim().toLowerCase();

    // Comprobar formato AM / PM
    const ampmMatch = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)$/i);
    if (ampmMatch) {
      let hours = parseInt(ampmMatch[1], 10);
      const minutes = parseInt(ampmMatch[2], 10);
      const isPm = ampmMatch[4].toLowerCase() === 'pm';
      if (isPm && hours < 12) hours += 12;
      if (!isPm && hours === 12) hours = 0;
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
    }

    // Coincidencia con HH:MM:SS o HH:MM
    const match = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (match) {
      const hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
      }
    }

    // Si es solo un entero de hora (ej: "8" o "08")
    const hourOnlyMatch = str.match(/^(\d{1,2})$/);
    if (hourOnlyMatch) {
      const hours = parseInt(hourOnlyMatch[1], 10);
      if (hours >= 0 && hours <= 23) {
        return `${hours.toString().padStart(2, '0')}:00`;
      }
    }

    return null;
  },

  /**
   * Extrae los bloques de horario de las filas
   */
  extractScheduleBlocks(rows) {
    const blocks = [];
    const daysKeys = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
    let validHourCount = 0;

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const timeRaw = row[0];
      const startTime = this.normalizeTime(timeRaw);

      if (!startTime) {
        // Si hay una fila vacía al final, la ignoramos; si está en medio y no tiene hora válida:
        if (r < 24 && String(timeRaw).trim() !== '') {
          return { valid: false };
        }
        continue;
      }

      validHourCount++;

      // Calcular hora de fin (1 hora después)
      const [h, m] = startTime.split(':').map(Number);
      const endH = (h + 1) % 24;
      const endTime = `${endH.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;

      // Recorrer los 7 días
      for (let d = 0; d < 7; d++) {
        const cellValue = row[d + 1];
        if (cellValue && String(cellValue).trim() !== '') {
          const title = String(cellValue).trim();
          blocks.push({
            id: 'base-' + d + '-' + startTime,
            day: daysKeys[d],
            startTime: startTime,
            endTime: endTime,
            title: title,
            type: 'base',
            color: '#2563eb' // Azul base
          });
        }
      }
    }

    // Si no contiene ninguna hora válida, no cumple la plantilla
    if (validHourCount === 0) {
      return { valid: false };
    }

    // Opcional: fusionar horas continuas con el mismo nombre para que en Teams/Outlook
    // se vea un bloque unificado (ej. "Trabajo" de 09:00 a 14:00)
    const mergedBlocks = this.mergeConsecutiveBlocks(blocks);

    return {
      valid: true,
      blocks: mergedBlocks
    };
  },

  /**
   * Agrupa bloques contiguos del mismo día con el mismo título
   */
  mergeConsecutiveBlocks(blocks) {
    if (!blocks || blocks.length <= 1) return blocks;

    const days = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
    const merged = [];

    days.forEach(day => {
      const dayBlocks = blocks
        .filter(b => b.day === day)
        .sort((a, b) => a.startTime.localeCompare(b.startTime));

      let current = null;

      dayBlocks.forEach(b => {
        if (!current) {
          current = { ...b };
        } else if (
          current.title.toLowerCase() === b.title.toLowerCase() &&
          current.endTime === b.startTime
        ) {
          // Fusionar
          current.endTime = b.endTime;
        } else {
          merged.push(current);
          current = { ...b };
        }
      });

      if (current) {
        merged.push(current);
      }
    });

    return merged;
  },

  /**
   * Genera y descarga la plantilla oficial de Crono-SYS en formato .xlsx
   */
  downloadTemplate() {
    if (typeof XLSX === 'undefined') {
      alert('La librería SheetJS aún se está cargando. Por favor espera.');
      return;
    }

    const headers = ['Hora/Día', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sábado', 'domingo'];
    const data = [headers];

    // Generar 24 filas desde 00:00:00 hasta 23:00:00 con celdas vacías listas para rellenar
    for (let h = 0; h < 24; h++) {
      const timeStr = `${h.toString().padStart(2, '0')}:00:00`;
      const row = [timeStr, '', '', '', '', '', '', ''];
      data.push(row);
    }

    // Crear libro con SheetJS
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Ajustar anchos de columnas
    ws['!cols'] = [
      { wch: 12 }, // Hora
      { wch: 18 }, // Lunes
      { wch: 18 }, // Martes
      { wch: 18 }, // Miércoles
      { wch: 18 }, // Jueves
      { wch: 18 }, // Viernes
      { wch: 18 }, // Sábado
      { wch: 18 }  // Domingo
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Horario');

    // Descargar archivo
    XLSX.writeFile(wb, 'plantilla_horario_crono_sys.xlsx');
  }
};

window.ExcelParser = ExcelParser;
