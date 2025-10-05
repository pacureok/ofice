import { useState, useMemo } from 'react';
import React from 'react'; 

// --- Configuración de la Cuadrícula para simular un tamaño "Infinito" ---
const ROWS = 200; // Filas hacia abajo
const COLS = 70;  // Columnas hacia la derecha (A hasta BR)

// Tipo para almacenar los datos de la hoja de cálculo
interface SheetData {
  [key: string]: string; // Clave: "A1", "B2", Valor: Contenido o fórmula
}

// Lista de pestañas disponibles en el Ribbon
type RibbonTab = 'Archivo' | 'Inicio' | 'Insertar' | 'Dibujar' | 'Disposicion' | 'Formulas' | 'Datos' | 'Revisar' | 'Vista' | 'Automatizar' | 'Ayuda';

/**
 * Genera los encabezados de las columnas (A, B, C, AA, AB, ...)
 */
const getColHeaders = (count: number) => {
  const headers: string[] = [];
  for (let i = 0; i < count; i++) {
    let header = '';
    let num = i;
    while (num >= 0) {
      header = String.fromCharCode((num % 26) + 'A'.charCodeAt(0)) + header;
      num = Math.floor(num / 26) - 1;
    }
    headers.push(header);
  }
  return headers;
};

// Expresión regular para encontrar referencias de celda (Ej: A1, B10, AA1)
const CELL_REFERENCE_REGEX = /([A-Z]+[0-9]+)/g; 

// --- Componente PacurHoja ---
const PacurHoja: React.FC = () => {
  const [data, setData] = useState<SheetData>({});
  const [activeCell, setActiveCell] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<RibbonTab>('Inicio'); // Pestaña activa
  const [zoomLevel, setZoomLevel] = useState(100); // Nivel de zoom para la barra de estado
  
  const colHeaders = useMemo(() => getColHeaders(COLS), []);

  // Maneja la entrada de datos en una celda
  const handleCellChange = (key: string, value: string) => {
    setData(prevData => ({
      ...prevData,
      [key]: value
    }));
  };

  /**
   * Resuelve el valor de una celda, buscando referencias circulares si es necesario.
   * Ahora soporta +, -, *, /, y ^ (exponenciación) con referencias a otras celdas.
   */
  const calculateValue = (key: string, path: string[] = []): string => {
    const content = data[key] || '';

    // 1. Detectar referencia circular
    if (path.includes(key)) {
      return '#CIRCULAR';
    }

    // 2. Si no es fórmula, devolver el contenido
    if (!content.startsWith('=')) {
      return content;
    }

    const formula = content.substring(1).trim();
    
    // 3. Sustituir referencias de celda por sus valores
    const resolvedFormula = formula.replace(CELL_REFERENCE_REGEX, (match) => {
      const referencedValue = calculateValue(match, [...path, key]);
      
      const numValue = parseFloat(referencedValue);
      
      if (isNaN(numValue)) {
        return '0'; 
      }
      return numValue.toString();
    });

    // 4. Reemplazar el operador de exponenciación (^) por **
    const finalFormula = resolvedFormula.replace(/\^/g, '**');

    // 5. Evaluar la fórmula (¡Advertencia: usar 'eval' en producción no es seguro!)
    try {
      // Intentar una evaluación más segura (aunque sigue siendo eval)
      const result = new Function('return ' + finalFormula)();
      
      if (typeof result === 'number' && !isNaN(result) && isFinite(result)) {
          return result.toFixed(2).replace(/\.00$/, ''); 
      }
      return '#ERROR_MATH';
    } catch (e) {
      return '#FÓRMULA_INVÁLIDA';
    }
  };


  // Función para guardar como .aph
  const saveSheet = () => {
    const filename = "hoja_calculo.aph";
    const content = JSON.stringify(data, null, 2); 

    const blob = new Blob([content], { type: 'application/json' }); 
    const a = document.createElement('a');
    a.download = filename;
    a.href = URL.createObjectURL(blob);
    
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    // Utilizamos un mensaje box personalizado
    const messageBox = document.createElement('div');
    messageBox.style.cssText = `
        position: fixed; top: 20px; right: 20px; background-color: #107c41; color: white;
        padding: 15px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        z-index: 1000; transition: opacity 0.5s; font-family: 'Inter', sans-serif;
    `;
    messageBox.textContent = `¡Hoja de cálculo ${filename} guardada con éxito!`;
    document.body.appendChild(messageBox);
    setTimeout(() => {
        messageBox.style.opacity = '0';
        setTimeout(() => document.body.removeChild(messageBox), 500);
    }, 3000);
  };

  // Función de marcador de posición para acciones de la barra de herramientas
  const handleToolbarAction = (action: string) => {
    const messageBox = document.createElement('div');
    messageBox.style.cssText = `
        position: fixed; top: 20px; right: 20px; background-color: #0078d4; color: white;
        padding: 10px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        z-index: 1000; transition: opacity 0.5s; font-family: 'Inter', sans-serif;
    `;
    messageBox.textContent = `Acción: ${action} - (Lógica no implementada)`;
    document.body.appendChild(messageBox);
    setTimeout(() => {
        messageBox.style.opacity = '0';
        setTimeout(() => document.body.removeChild(messageBox), 1000);
    }, 2000);
  };
  
  // --- Contenido del Ribbon para cada Pestaña ---
  const renderRibbonContent = () => {
    switch (activeTab) {
      case 'Inicio':
        return (
          <div className="ribbon-content">
            {/* GRUPO: PORTAPAPELES (Inicio) */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Pegar")} title="Pegar" className="large-button">
                    <span className="icon-xl">📋</span><br/>Pegar
                </button>
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Cortar")} title="Cortar">✂️</button>
                    <button onClick={() => handleToolbarAction("Copiar")} title="Copiar">📝</button>
                    <button onClick={() => handleToolbarAction("Copia de formato")} title="Copia de formato">🪟</button>
                </div>
                <div className="group-label">Portapapeles</div>
            </div>

            {/* GRUPO: FUENTE (Inicio) */}
            <div className="toolbar-group">
                <div className="horizontal-group input-row">
                    <select defaultValue="Aptos Narrow" title="Fuente" className="font-select">
                        <option>Aptos Narrow</option><option>Arial</option><option>Calibri</option>
                    </select>
                    <select defaultValue="11" title="Tamaño" className="size-select">
                        <option>11</option><option>12</option><option>14</option>
                    </select>
                </div>
                <div className="horizontal-group">
                    <button onClick={() => handleToolbarAction("Negrita")} title="Negrita"><b>N</b></button>
                    <button onClick={() => handleToolbarAction("Cursiva")} title="Cursiva"><i>K</i></button>
                    <button onClick={() => handleToolbarAction("Subrayado")} title="Subrayado"><u>S</u></button>
                    <button onClick={() => handleToolbarAction("Bordes")} title="Bordes de celda">🖼️</button>
                    <button onClick={() => handleToolbarAction("Relleno")} title="Color de Relleno" style={{color: 'rgb(240, 240, 0)'}}>🎨</button>
                    <button onClick={() => handleToolbarAction("Color Fuente")} title="Color de Fuente" style={{color: 'red'}}>🅰️</button>
                </div>
                <div className="group-label">Fuente</div>
            </div>

            {/* GRUPO: ALINEACIÓN (Inicio) */}
            <div className="toolbar-group">
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Alinear Superior")} title="Alinear Arriba">↖</button>
                    <button onClick={() => handleToolbarAction("Alinear Medio")} title="Alinear Medio">↔</button>
                    <button onClick={() => handleToolbarAction("Alinear Inferior")} title="Alinear Abajo">↙</button>
                </div>
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Izquierda")} title="Alinear Izquierda">⏴</button>
                    <button onClick={() => handleToolbarAction("Centrar")} title="Centrar">☰</button>
                    <button onClick={() => handleToolbarAction("Derecha")} title="Alinear Derecha">⏵</button>
                </div>
                <div className="group-label">Alineación</div>
            </div>

            {/* GRUPO: NÚMERO (Inicio) */}
            <div className="toolbar-group">
                <select defaultValue="General" title="Formato de Número" className="number-select">
                    <option>General</option><option>Número</option><option>Moneda</option><option>Porcentaje</option>
                </select>
                <div className="horizontal-group">
                    <button onClick={() => handleToolbarAction("Moneda")} title="Formato Moneda">$</button>
                    <button onClick={() => handleToolbarAction("Porcentaje")} title="Estilo Porcentual">%</button>
                    <button onClick={() => handleToolbarAction("Comas")} title="Estilo Millares">, </button>
                </div>
                <div className="group-label">Número</div>
            </div>

            {/* GRUPO: ESTILOS (Inicio) */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Formato Condicional")} title="Formato Condicional" className="small-icon-button">📊</button>
                <button onClick={() => handleToolbarAction("Dar Formato Como Tabla")} title="Dar Formato como Tabla" className="small-icon-button">📋</button>
                <button onClick={() => handleToolbarAction("Estilos de Celda")} title="Estilos de Celda" className="small-icon-button">🎨</button>
                <div className="group-label">Estilos</div>
            </div>

            {/* GRUPO: CELDAS (Inicio) */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Insertar")} title="Insertar Celdas/Filas" className="small-icon-button">➕</button>
                <button onClick={() => handleToolbarAction("Eliminar")} title="Eliminar Celdas/Filas" className="small-icon-button">➖</button>
                <button onClick={() => handleToolbarAction("Formato")} title="Formato de Fila/Columna" className="small-icon-button">⚙️</button>
                <div className="group-label">Celdas</div>
            </div>

            {/* GRUPO: EDICIÓN (Inicio) */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Autosuma")} title="Autosuma" className="small-icon-button">Σ</button>
                <button onClick={() => handleToolbarAction("Ordenar y Filtrar")} title="Ordenar y Filtrar" className="small-icon-button">⬇️⬆️</button>
                <button onClick={() => handleToolbarAction("Buscar y Seleccionar")} title="Buscar y Seleccionar" className="small-icon-button">🔍</button>
                <div className="group-label">Edición</div>
            </div>
            
          </div>
        );
      
      case 'Insertar':
        return (
            <div className="ribbon-content">
                {/* GRUPO: TABLAS (Insertar) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Tabla")} title="Tabla" className="large-button">
                        <span className="icon-xl">📅</span><br/>Tabla
                    </button>
                    <button onClick={() => handleToolbarAction("Tablas Dinámicas")} title="Tablas Dinámicas" className="large-button">
                        <span className="icon-xl">🗃️</span><br/>Tablas Dinámicas
                    </button>
                    <div className="group-label">Tablas</div>
                </div>
                {/* GRUPO: ILUSTRACIONES (Insertar) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Imágenes")} title="Imágenes" className="large-button"><span className="icon-xl">🖼️</span><br/>Imágenes</button>
                    <button onClick={() => handleToolbarAction("Formas")} title="Formas" className="large-button"><span className="icon-xl">🔺</span><br/>Formas</button>
                    <button onClick={() => handleToolbarAction("Iconos")} title="Iconos" className="large-button"><span className="icon-xl">🌟</span><br/>Iconos</button>
                    <div className="group-label">Ilustraciones</div>
                </div>
                {/* GRUPO: GRÁFICOS (Insertar) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Gráfico")} title="Gráficos Recomendados" className="large-button">
                        <span className="icon-xl">📈</span><br/>Gráficos
                    </button>
                    <button onClick={() => handleToolbarAction("Gráfico Dinámico")} title="Gráfico Dinámico" className="large-button">
                        <span className="icon-xl">📊</span><br/>Gráfico Dinámico
                    </button>
                    <div className="group-label">Gráficos</div>
                </div>
            </div>
        );
        
      case 'Formulas':
        return (
            <div className="ribbon-content">
                {/* GRUPO: BIBLIOTECA DE FUNCIONES (Fórmulas) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Insertar Función")} title="Insertar Función" className="large-button">
                        <span className="icon-xl">ƒx</span><br/>Insertar Función
                    </button>
                    <button onClick={() => handleToolbarAction("Autosuma")} title="Autosuma" className="large-button">
                        <span className="icon-xl">Σ</span><br/>Autosuma
                    </button>
                    <div className="vertical-group icon-only-group">
                        <button onClick={() => handleToolbarAction("Financieras")} title="Financieras">🏦</button>
                        <button onClick={() => handleToolbarAction("Lógicas")} title="Lógicas">✅</button>
                        <button onClick={() => handleToolbarAction("Texto")} title="Texto">Añ</button>
                        <button onClick={() => handleToolbarAction("Matemáticas")} title="Matemáticas">π</button>
                    </div>
                    <div className="group-label">Biblioteca de funciones</div>
                </div>
                {/* GRUPO: NOMBRES DEFINIDOS (Fórmulas) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Administrador de Nombres")} title="Administrador de Nombres" className="large-button"><span className="icon-xl">🏷️</span><br/>Adm. Nombres</button>
                    <button onClick={() => handleToolbarAction("Asignar Nombre")} title="Asignar Nombre" className="large-button"><span className="icon-xl">📝</span><br/>Asignar Nombre</button>
                    <div className="group-label">Nombres definidos</div>
                </div>
                {/* GRUPO: AUDITORÍA DE FÓRMULAS (Fórmulas) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Rastrear Precedentes")} title="Rastrear Precedentes" className="large-button"><span className="icon-xl">⬅️</span><br/>Rastrear Precedentes</button>
                    <button onClick={() => handleToolbarAction("Mostrar Fórmulas")} title="Mostrar Fórmulas" className="large-button"><span className="icon-xl">📜</span><br/>Mostrar Fórmulas</button>
                    <button onClick={() => handleToolbarAction("Comprobación de Errores")} title="Comprobación de Errores" className="large-button"><span className="icon-xl">⚠️</span><br/>Comprobación de Errores</button>
                    <div className="group-label">Auditoría de fórmulas</div>
                </div>
            </div>
        );

      case 'Datos':
        return (
            <div className="ribbon-content">
                {/* GRUPO: OBTENER Y TRANSFORMAR DATOS (Datos) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Obtener Datos")} title="Obtener Datos" className="large-button"><span className="icon-xl">💾</span><br/>Obtener Datos</button>
                    <button onClick={() => handleToolbarAction("Desde Texto/CSV")} title="Desde Texto/CSV" className="large-button"><span className="icon-xl">📄</span><br/>Desde Texto</button>
                    <button onClick={() => handleToolbarAction("De la Web")} title="De la Web" className="large-button"><span className="icon-xl">🌐</span><br/>De la Web</button>
                    <div className="group-label">Obtener y transformar datos</div>
                </div>
                {/* GRUPO: ORDENAR Y FILTRAR (Datos) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Ordenar A-Z")} title="Ordenar A-Z" className="large-button"><span className="icon-xl">⬇️</span><br/>Ordenar</button>
                    <button onClick={() => handleToolbarAction("Filtro")} title="Filtro" className="large-button"><span className="icon-xl">🔽</span><br/>Filtro</button>
                    <div className="group-label">Ordenar y Filtrar</div>
                </div>
                {/* GRUPO: HERRAMIENTAS DE DATOS (Datos) */}
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction("Texto en Columnas")} title="Texto en Columnas" className="large-button"><span className="icon-xl">🗒️</span><br/>Texto en Columnas</button>
                    <button onClick={() => handleToolbarAction("Quitar Duplicados")} title="Quitar Duplicados" className="large-button"><span className="icon-xl">🗑️</span><br/>Quitar Duplicados</button>
                    <div className="group-label">Herramientas de datos</div>
                </div>
            </div>
        );

      default:
        // Pestañas con contenido básico o no implementado
        return (
            <div className="ribbon-content">
                <div className="toolbar-group">
                    <button onClick={() => handleToolbarAction(`Acción en ${activeTab}`)} title={`Acción en ${activeTab}`} className="large-button">
                        <span className="icon-xl">🛠️</span><br/>{activeTab}
                    </button>
                    <div className="group-label">Contenido Básico</div>
                </div>
            </div>
        );
    }
  };


  return (
    <div className="pacur-hoja-container">
    <style jsx global>{`
        /* Global Reset and Font */
        :root {
            --excel-dark-bg: #1e1e1e;
            --excel-mid-bg: #2d2d2d;
            --excel-light-bg: #3c3c3c;
            --excel-grid-line: #444;
            --excel-text: #fff;
            --excel-active-tab: #0078d4;
            --excel-button-hover: #4d4d4d;
            --excel-formula-bar: #333;
            --excel-header-bg: #333;
            --excel-header-border: #555;
            --excel-active-cell: #0078d4;
            --excel-active-tab-indicator: #0078d4;
        }

        body, html, #root {
            margin: 0;
            padding: 0;
            height: 100%;
            width: 100%;
            font-family: 'Inter', sans-serif;
            background-color: var(--excel-dark-bg);
            color: var(--excel-text);
        }

        .pacur-hoja-container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            overflow: hidden;
            background-color: var(--excel-dark-bg);
        }

        /* 1. RIBBON (Barra de Herramientas Superior) */
        .ribbon {
            background-color: var(--excel-mid-bg);
            border-bottom: 1px solid var(--excel-grid-line);
            padding-top: 5px;
            user-select: none;
        }

        .ribbon-tabs {
            display: flex;
            align-items: flex-end;
            padding: 0 10px;
        }

        .ribbon-tab {
            padding: 8px 15px;
            cursor: pointer;
            color: var(--excel-text);
            font-size: 0.85rem;
            border-radius: 4px 4px 0 0;
            margin-right: 2px;
            transition: background-color 0.2s;
        }

        .ribbon-tab:hover {
            background-color: var(--excel-button-hover);
        }

        .ribbon-tab.active {
            background-color: var(--excel-dark-bg);
            border-bottom: 2px solid var(--excel-active-tab-indicator);
            font-weight: 600;
            padding-bottom: 7px; /* Compensar el border-bottom */
        }
        
        .ribbon-content {
            background-color: var(--excel-dark-bg);
            padding: 8px 10px;
            display: flex;
            flex-wrap: nowrap;
            overflow-x: auto;
            border-bottom: 1px solid var(--excel-grid-line);
        }

        .toolbar-group {
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 0 15px;
            border-right: 1px solid var(--excel-mid-bg);
            position: relative;
        }

        .group-label {
            font-size: 0.7rem;
            color: #ccc;
            margin-top: 5px;
            text-align: center;
            width: 100%;
        }

        .large-button {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 5px 8px;
            height: 60px;
            width: 60px;
            font-size: 0.75rem;
        }
        
        /* Estilos generales de botones del Ribbon */
        .toolbar-group button,
        .ribbon-tabs button {
            background-color: transparent;
            border: 1px solid transparent;
            color: var(--excel-text);
            padding: 3px 5px;
            margin: 1px;
            border-radius: 3px;
            cursor: pointer;
            transition: background-color 0.1s, border-color 0.1s;
        }

        .toolbar-group button:hover,
        .ribbon-tabs button:hover {
            background-color: var(--excel-button-hover);
            border-color: var(--excel-header-border);
        }
        
        .icon-xl {
            font-size: 1.2rem;
            line-height: 1;
        }

        .vertical-group {
            display: flex;
            flex-direction: column;
            margin-left: 5px;
        }
        
        .horizontal-group {
            display: flex;
            margin-bottom: 2px;
        }

        .input-row {
            margin-bottom: 4px;
        }

        .font-select, .size-select, .number-select {
            background-color: var(--excel-mid-bg);
            color: var(--excel-text);
            border: 1px solid var(--excel-header-border);
            padding: 2px 5px;
            border-radius: 3px;
            font-size: 0.8rem;
            margin-right: 2px;
        }

        .font-select { width: 120px; }
        .size-select { width: 50px; }
        .number-select { width: 90px; }


        /* 2. BARRA DE FÓRMULAS */
        .formula-bar {
            display: flex;
            align-items: center;
            background-color: var(--excel-formula-bar);
            padding: 4px 10px;
            border-bottom: 1px solid var(--excel-grid-line);
        }

        .cell-name-box {
            background-color: var(--excel-mid-bg);
            border: 1px solid var(--excel-grid-line);
            padding: 4px 8px;
            margin-right: 10px;
            font-weight: bold;
            min-width: 50px;
            text-align: center;
            border-radius: 2px;
            font-size: 0.9rem;
        }

        .formula-input {
            flex-grow: 1;
            background-color: var(--excel-dark-bg);
            color: var(--excel-text);
            border: 1px solid var(--excel-grid-line);
            padding: 4px 8px;
            font-size: 0.9rem;
            border-radius: 2px;
            box-shadow: inset 0 1px 3px rgba(0, 0, 0, 0.5);
        }


        /* 3. HOJA DE CÁLCULO (Cuadrícula) */
        .spreadsheet-grid {
            overflow: auto;
            flex-grow: 1;
            background-color: var(--excel-dark-bg);
            position: relative;
        }
        
        .header-row {
            display: flex;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        .data-row {
            display: flex;
            height: 25px; /* Altura fija de la fila */
        }
        
        .cell {
            min-width: 80px; /* Ancho estándar de columna */
            height: 25px;
            border: 1px solid var(--excel-grid-line);
            border-top: none;
            border-left: none;
            display: flex;
            align-items: center;
            padding: 0 5px;
            box-sizing: border-box;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            background-color: var(--excel-dark-bg);
            color: var(--excel-text);
            cursor: pointer;
        }

        .header-cell {
            position: sticky;
            background-color: var(--excel-header-bg);
            color: #ddd;
            font-weight: normal;
            justify-content: center;
            z-index: 20; 
            border: 1px solid var(--excel-header-border);
            border-top: none;
        }

        .corner-cell {
            position: sticky;
            left: 0;
            top: 0;
            z-index: 30;
            min-width: 50px;
            background-color: var(--excel-header-bg);
            border: 1px solid var(--excel-header-border);
            border-top: none;
        }

        .data-row > .header-cell {
            position: sticky;
            left: 0;
            z-index: 20; 
            min-width: 50px;
            justify-content: center;
        }
        
        .data-cell {
            background-color: var(--excel-dark-bg);
            border-color: var(--excel-grid-line);
        }

        .data-cell:hover {
            outline: 1px solid #777;
            z-index: 5;
        }

        .data-cell.active {
            outline: 2px solid var(--excel-active-cell);
            background-color: #000;
            z-index: 6;
        }


        /* 4. BARRA DE ESTADO (Parte Inferior) */
        .status-bar {
            background-color: var(--excel-header-bg);
            border-top: 1px solid var(--excel-grid-line);
            padding: 5px 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.8rem;
            color: #ccc;
            height: 30px;
        }

        .status-left {
            display: flex;
            align-items: center;
            gap: 15px;
        }

        .sheet-tabs {
            display: flex;
            align-items: center;
            gap: 5px;
            margin-left: 10px;
        }
        
        .sheet-tab {
            padding: 4px 10px;
            cursor: pointer;
            border-radius: 4px;
        }

        .sheet-tab.active-sheet {
            color: var(--excel-active-cell);
            border-bottom: 2px solid var(--excel-active-cell);
            font-weight: 600;
            padding-bottom: 2px;
        }

        .status-right {
            display: flex;
            align-items: center;
            gap: 15px;
        }

        .zoom-control {
            display: flex;
            align-items: center;
            gap: 5px;
            margin-left: 10px;
        }
        
        .status-bar button {
            background: transparent;
            border: none;
            color: #ccc;
            padding: 4px 8px;
            border-radius: 3px;
            cursor: pointer;
            transition: background-color 0.1s;
        }
        
        .status-bar button:hover {
            background-color: var(--excel-button-hover);
            color: var(--excel-text);
        }
    `}</style>
      
      {/* 1. Barra de Herramientas (Ribbon COMPLETO) */}
      <div className="toolbar ribbon">
        {/* Pestañas (Simulación) */}
        <div className="ribbon-tabs">
            {['Archivo', 'Inicio', 'Insertar', 'Dibujar', 'Disposicion', 'Formulas', 'Datos', 'Revisar', 'Vista', 'Automatizar', 'Ayuda'].map(tab => (
                 <span 
                    key={tab}
                    className={`ribbon-tab ${activeTab === tab ? 'active' : ''}`}
                    onClick={() => setActiveTab(tab as RibbonTab)}
                >
                    {tab}
                </span>
            ))}
            
            {/* Botones de Compartir/Comentarios */}
            <div style={{marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center'}}>
                <button onClick={() => handleToolbarAction("Comentarios")} title="Comentarios">💬 Comentarios</button>
                <button onClick={() => handleToolbarAction("Compartir")} title="Compartir" style={{backgroundColor: '#107c41', color: 'white'}}>
                    📤 Compartir
                </button>
                {/* Botón de Guardar */}
                <button onClick={saveSheet} title="Guardar como .aph" style={{backgroundColor: '#0078d4'}}>💾</button>
            </div>
        </div>

        {/* Contenido de la Pestaña Activa */}
        {renderRibbonContent()}
      </div>

      {/* 2. Barra de Fórmulas */}
      <div className="formula-bar">
        <div className="cell-name-box">
            {activeCell || 'A1'}
        </div>
        <input 
            type="text" 
            placeholder="Fórmula o valor" 
            value={activeCell ? data[activeCell] || '' : ''}
            onChange={(e) => activeCell && handleCellChange(activeCell, e.target.value)}
            className="formula-input"
        />
      </div>

      {/* 3. Cuadrícula de la Hoja de Cálculo */}
      {/* La transformación del zoom está aquí para que solo afecte a la cuadrícula, no al ribbon ni a la barra de estado */}
      <div className="spreadsheet-grid" style={{overflow: 'auto'}}> 
        <div style={{transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top left', minWidth: `${(COLS * 80) + 50}px`, minHeight: `${(ROWS * 25) + 25}px`}}>
            <div className="header-row">
            <div className="cell header-cell corner-cell"></div>
            {colHeaders.map(header => (
                <div key={header} className="cell header-cell">{header}</div>
            ))}
            </div>

            {Array.from({ length: ROWS }, (_, rIndex) => (
            <div key={rIndex} className="data-row">
                <div className="cell header-cell">{(rIndex + 1)}</div>
                
                {colHeaders.map(cHeader => {
                const cellKey = `${cHeader}${rIndex + 1}`;
                const displayValue = calculateValue(cellKey);
                
                return (
                    <div 
                    key={cellKey}
                    className={`cell data-cell ${activeCell === cellKey ? 'active' : ''}`}
                    onClick={() => setActiveCell(cellKey)}
                    >
                    {displayValue}
                    </div>
                );
                })}
            </div>
            ))}
        </div>
      </div>
      
      {/* 4. Barra de Estado (Parte Inferior) */}
      <div className="status-bar">
        <div className="status-left">
            <span>Listo</span>
            <div className="sheet-tabs">
                {/* Flechas de navegación (opcional) */}
                <button onClick={() => handleToolbarAction("Hoja Anterior")} title="Hoja Anterior">{"<"}</button>
                <button onClick={() => handleToolbarAction("Hoja Siguiente")} title="Hoja Siguiente">{">"}</button>
                <span className="sheet-tab active-sheet">Hoja1</span>
                <button onClick={() => handleToolbarAction("Añadir Hoja")} title="Nueva Hoja">➕</button>
            </div>
        </div>
        <div className="status-right">
            <button onClick={() => handleToolbarAction("Vista Normal")} title="Vista Normal">📄</button>
            <button onClick={() => handleToolbarAction("Vista Diseño")} title="Vista Diseño de página">📑</button>
            <button onClick={() => handleToolbarAction("Vista Salto de página")} title="Vista previa de salto de página">🎚️</button>

            <div className="zoom-control">
                <button onClick={() => setZoomLevel(Math.max(10, zoomLevel - 10))} title="Alejar">➖</button>
                <span>{zoomLevel}%</span>
                <button onClick={() => setZoomLevel(Math.min(200, zoomLevel + 10))} title="Acercar">➕</button>
            </div>
        </div>
      </div>
    </div>
  );
};

export default PacurHoja;
