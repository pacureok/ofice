import { useState, useMemo, useCallback, useEffect } from 'react';
import React from 'react'; 

// --- Configuración de la Cuadrícula para simular un tamaño "Infinito" ---
const ROWS = 200; // Filas hacia abajo
const COLS = 70;  // Columnas hacia la derecha (A hasta BR)

// Tipo para almacenar los estilos de una celda
interface CellStyle {
  fontWeight?: 'bold' | 'normal';
  fontStyle?: 'italic' | 'normal';
  textDecoration?: 'underline' | 'none';
  textAlign?: 'left' | 'center' | 'right';
  backgroundColor?: string;
  color?: string;
  // Nuevos estilos para Fuente y Número
  fontSize?: string; // Ej: '11px'
  fontFamily?: string; // Ej: 'Aptos Narrow'
  numberFormat?: string; // Ej: 'General', 'Moneda', 'Porcentaje'
  
  // Nuevo estilo para Bordes
  borderStyle?: 'all' | 'none' | 'bottom' | 'top' | 'left' | 'right' | 'outside' | 'thickOutside';
}

// Tipo para almacenar los datos de la hoja de cálculo
interface SheetData {
  [key: string]: string; // Clave: "A1", "B2", Valor: Contenido o fórmula
}

// Tipo para almacenar el estado completo de la hoja de cálculo (para guardar/cargar)
interface WorkbookState {
    data: SheetData;
    styles: { [key: string]: CellStyle };
    fileName: string;
}

// Tipos para el estado y la configuración
type RibbonTab = 'Archivo' | 'Inicio' | 'Insertar' | 'Dibujar' | 'Disposicion' | 'Formulas' | 'Datos' | 'Revisar' | 'Vista' | 'Automatizar' | 'Ayuda';
type ContextMenu = {
    visible: boolean;
    x: number;
    y: number;
    targetType: 'cell' | 'row' | 'col';
    cellKey: string | null;
}
// Tipos para las vistas de Backstage
type BackstageView = 'Inicio' | 'Abrir' | 'Guardar como' | 'Imprimir' | 'Opciones';


// Tipos para el control de bordes
type BorderOption = 'Ninguno' | 'Todos los bordes' | 'Borde inferior' | 'Borde superior' | 'Borde izquierdo' | 'Borde derecho' | 'Bordes externos' | 'Borde exterior grueso';


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
// Expresión regular para encontrar funciones de rango
const FUNCTION_REGEX = /(SUMA|PROMEDIO)\(([^)]+)\)/g;

// Estructura de libro de trabajo por defecto
const INITIAL_STATE: WorkbookState = {
    data: {},
    styles: {},
    fileName: 'Libro1.xlsx',
}


// --- Estilos por defecto ---
const defaultStyles: CellStyle = {
    fontWeight: 'normal',
    fontStyle: 'normal',
    textDecoration: 'none',
    textAlign: 'left',
    backgroundColor: 'inherit',
    color: 'var(--excel-text)',
    fontSize: '11px',
    fontFamily: 'Aptos Narrow',
    numberFormat: 'General',
    borderStyle: 'none', // Valor por defecto
}

// --- Componente PacurHoja ---
const PacurHoja: React.FC = () => {
  const [data, setData] = useState<SheetData>(INITIAL_STATE.data);
  const [cellStyles, setCellStyles] = useState<{ [key: string]: CellStyle }>(INITIAL_STATE.styles);
  const [fileName, setFileName] = useState(INITIAL_STATE.fileName);
  const [isDirty, setIsDirty] = useState(false); // Indica si hay cambios sin guardar

  const [activeCell, setActiveCell] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<RibbonTab>('Inicio');
  const [zoomLevel, setZoomLevel] = useState(100);
  
  const [selectedRows, _setSelectedRows] = useState<number[]>([]); 
  const [selectedCols, _setSelectedCols] = useState<string[]>([]); 
  
  const [contextMenu, setContextMenu] = useState<ContextMenu>({ 
    visible: false, x: 0, y: 0, targetType: 'cell', cellKey: null 
  });
  
  // Nuevo estado para la vista Backstage y su sub-navegación
  const [backstageVisible, setBackstageVisible] = useState(false);
  const [backstageView, setBackstageView] = useState<BackstageView>('Inicio');
  
  // Estados para los selectores de color
  const [selectedFillColor, setSelectedFillColor] = useState<string>('#4d4d4d'); 
  const [selectedTextColor, setSelectedTextColor] = useState<string>('#ffffff'); 
  
  // Estado para el menú desplegable de bordes
  const [borderMenuVisible, setBorderMenuVisible] = useState(false);
  
  // Estado para la simulación de Imprimir
  const [printRange, setPrintRange] = useState('1A:5G'); // Rango por defecto para impresión
  
  const colHeaders = useMemo(() => getColHeaders(COLS), []);

  // Efecto para marcar el archivo como modificado al cambiar datos o estilos
  useEffect(() => {
    if (data !== INITIAL_STATE.data || cellStyles !== INITIAL_STATE.styles) {
        setIsDirty(true);
    }
  }, [data, cellStyles]);
    
  // --- Lógica de Aplicación de Estilo ---

  /**
   * Obtiene los estilos consolidados de la celda activa para la barra de herramientas.
   */
  const currentCellStyles = useMemo(() => {
    if (!activeCell) return defaultStyles;
    // Fusiona los estilos por defecto con los estilos específicos de la celda
    return { ...defaultStyles, ...(cellStyles[activeCell] || {}) };
  }, [activeCell, cellStyles]);


  /**
   * Mapea la opción de borde de texto a la propiedad CSS que almacenamos.
   */
  const getBorderStyleValue = (option: BorderOption): CellStyle['borderStyle'] => {
    switch (option) {
        case 'Todos los bordes':
            return 'all';
        case 'Borde inferior':
            return 'bottom';
        case 'Borde superior':
            return 'top';
        case 'Borde izquierdo':
            return 'left';
        case 'Borde derecho':
            return 'right';
        case 'Bordes externos':
            return 'outside';
        case 'Borde exterior grueso':
            return 'thickOutside';
        case 'Ninguno':
        default:
            return 'none';
    }
  }

  /**
   * Aplica un estilo al activeCell, manejando el toggle para N, K, S, y la configuración directa para otros.
   */
  const applyStyleToActiveCell = (styleKey: keyof CellStyle, value: string | undefined) => {
    if (!activeCell) {
        showMessageBox("Selecciona una celda primero para aplicar formato.", true);
        return;
    }

    setCellStyles(prevStyles => {
        const currentStyle = prevStyles[activeCell] || defaultStyles;
        
        // Manejar el toggle para Negrita, Cursiva, Subrayado
        if (styleKey === 'fontWeight' || styleKey === 'fontStyle' || styleKey === 'textDecoration') {
            const currentValue = currentStyle[styleKey] === value; 
            const newValue = currentValue 
                ? (styleKey === 'fontWeight' ? 'normal' : 'none') 
                : value; 
            
            return {
                ...prevStyles,
                [activeCell]: {
                    ...currentStyle,
                    [styleKey]: newValue
                }
            };
        }
        
        // Manejar Alineación, Colores, Fuente, Tamaño y Bordes
        return {
            ...prevStyles,
            [activeCell]: {
                ...currentStyle,
                [styleKey]: value as any
            }
        };
    });
    setIsDirty(true);
  };

  /**
   * Aplica el estilo de borde a la celda activa.
   */
  const applyBorderStyle = (option: BorderOption) => {
    const borderValue = getBorderStyleValue(option);
    applyStyleToActiveCell('borderStyle', borderValue);
    setBorderMenuVisible(false); // Cierra el menú después de la selección
    showMessageBox(`Borde: ${option} aplicado a ${activeCell}`);
  }

  /**
   * Resuelve el valor de una celda (cálculo de fórmulas).
   */
  const calculateValue = (key: string, path: string[] = []): string => {
    const content = data[key] || '';
    if (path.includes(key)) return '#CIRCULAR';
    if (!content.startsWith('=')) return content;

    let formula = content.substring(1).trim();
    
    // 1. Resolución de Funciones (SUMA, PROMEDIO)
    formula = formula.replace(FUNCTION_REGEX, (match, funcName, rangeStr) => {
        const rangeKeys = parseRange(rangeStr.trim());
        const values: number[] = [];

        for (const cellKey of rangeKeys) {
            const valueStr = calculateValue(cellKey, [...path, key]);
            const numValue = parseFloat(valueStr);

            if (!isNaN(numValue) && isFinite(numValue)) {
                values.push(numValue);
            }
        }
        
        if (values.length === 0) return '0';

        const func = funcName.toUpperCase();
        if (func === 'SUMA') return values.reduce((a, b) => a + b, 0).toString();
        if (func === 'PROMEDIO') return (values.reduce((a, b) => a + b, 0) / values.length).toString();

        return match; 
    });

    // 2. Sustitución de Referencias Simples (Ej: A1 + B2)
    const resolvedFormula = formula.replace(CELL_REFERENCE_REGEX, (match) => {
      const referencedValue = calculateValue(match, [...path, key]);
      const numValue = parseFloat(referencedValue);
      // Solo reemplazamos la referencia si es un número válido o no tiene error.
      if (isNaN(numValue) || referencedValue.startsWith('#')) return '0'; 
      return numValue.toString();
    });

    // 3. Evaluar la fórmula
    try {
      const finalFormula = resolvedFormula.replace(/\^/g, '**');
      // eslint-disable-next-line no-new-func
      const result = new Function('return ' + finalFormula)();
      
      if (typeof result === 'number' && !isNaN(result) && isFinite(result)) {
          // Obtener el formato de número de la celda actual
          const format = cellStyles[key]?.numberFormat || 'General';

          let formattedResult = result.toString();

          switch(format) {
              case 'Moneda':
                  formattedResult = `$${result.toFixed(2)}`;
                  break;
              case 'Porcentaje':
                  formattedResult = `${(result * 100).toFixed(2).replace(/\.00$/, '')}%`;
                  break;
              case 'Milésimas':
                  formattedResult = result.toLocaleString('es-MX', { maximumFractionDigits: 2 });
                  break;
              // General se aplica por defecto
          }

          return formattedResult; 
      }
      return '#ERROR_MATH';
    } catch (e) {
      return '#FÓRMULA_INVÁLIDA';
    }
  };


  // Función para mostrar mensajes temporales (reemplaza alert/confirm)
  const showMessageBox = (message: string, isError: boolean = false) => {
    const messageBox = document.createElement('div');
    messageBox.style.cssText = `
        position: fixed; top: 20px; right: 20px; 
        background-color: ${isError ? '#e53e3e' : '#107c41'}; 
        color: white;
        padding: 15px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        z-index: 2000; transition: opacity 0.5s; font-family: 'Inter', sans-serif;
    `;
    messageBox.textContent = message;
    document.body.appendChild(messageBox);
    setTimeout(() => {
        messageBox.style.opacity = '0';
        setTimeout(() => document.body.removeChild(messageBox), 500);
    }, 3000);
  };

  // --- LÓGICA DE GESTIÓN DE ARCHIVOS ---

  const handleOpen = () => {
    if (isDirty) {
        showMessageBox("Tienes cambios sin guardar. Guardando automáticamente 'Libro Anterior'.", false);
        // Simular que guardamos el estado actual antes de cargarlo
        localStorage.setItem('temp_workbook_state', JSON.stringify({ data, styles, fileName }));
    }

    const previousFile = localStorage.getItem('last_saved_workbook');
    if (previousFile) {
        const state: WorkbookState = JSON.parse(previousFile);
        setData(state.data);
        setCellStyles(state.styles);
        setFileName(state.fileName);
        setIsDirty(false);
        setBackstageVisible(false);
        setActiveTab('Inicio');
        showMessageBox(`Archivo "${state.fileName}" cargado correctamente.`);
    } else {
        showMessageBox("No se encontró ningún archivo guardado previamente.", true);
        setBackstageVisible(false);
        setActiveTab('Inicio');
    }
  };

  const handleSave = () => {
    if (!isDirty) {
        showMessageBox("No hay cambios para guardar.", false);
        setBackstageVisible(false);
        return;
    }
    
    // Simular guardado en localStorage
    const currentState: WorkbookState = { data, styles: cellStyles, fileName };
    localStorage.setItem('last_saved_workbook', JSON.stringify(currentState));
    setIsDirty(false);
    setBackstageVisible(false);
    showMessageBox(`Archivo "${fileName}" guardado correctamente.`);
  };

  const handleSaveAs = (newFileName: string, format: 'xlsx' | 'pacur') => {
    const fullFileName = `${newFileName}.${format}`;
    
    // Simular la descarga del archivo (o persistencia)
    const currentState: WorkbookState = { data, styles: cellStyles, fileName: fullFileName };
    
    // Usamos localStorage para simular el guardado
    localStorage.setItem('last_saved_workbook', JSON.stringify(currentState));
    
    setFileName(fullFileName);
    setIsDirty(false);
    setBackstageVisible(false);
    showMessageBox(`Archivo guardado como "${fullFileName}" en formato ${format.toUpperCase()}.`);
  }
  
  // --- Manejadores de Interacción General ---

  // Ocultar el menú contextual y el menú de bordes
  const hideContextMenu = useCallback(() => {
    setContextMenu(prev => ({ ...prev, visible: false }));
    setBorderMenuVisible(false);
  }, []);

  // Maneja la entrada de datos en una celda
  const handleCellChange = (key: string, value: string) => {
    setData(prevData => ({
      ...prevData,
      [key]: value
    }));
    setIsDirty(true);
  };
  
  // Placeholder para acciones de Portapapeles
  const handleClipboardAction = (action: 'Cortar' | 'Copiar' | 'Pegar') => {
      if (!activeCell) {
          showMessageBox("Selecciona una celda primero.", true);
          return;
      }
      showMessageBox(`${action} en ${activeCell} - (Acción simulada)`);
  }

  // Manejador principal para acciones de la barra de herramientas
  const handleToolbarAction = (action: string) => {
    switch (action) {
        case 'Negrita':
            applyStyleToActiveCell('fontWeight', 'bold');
            break;
        case 'Cursiva':
            applyStyleToActiveCell('fontStyle', 'italic');
            break;
        case 'Subrayado':
            applyStyleToActiveCell('textDecoration', 'underline');
            break;
        case 'Alinear Izquierda':
            applyStyleToActiveCell('textAlign', 'left');
            break;
        case 'Alinear Centro':
            applyStyleToActiveCell('textAlign', 'center');
            break;
        case 'Alinear Derecha':
            applyStyleToActiveCell('textAlign', 'right');
            break;
        case 'Moneda':
            applyStyleToActiveCell('numberFormat', 'Moneda');
            break;
        case 'Porcentaje':
            applyStyleToActiveCell('numberFormat', 'Porcentaje');
            break;
        case 'Milésimas':
            applyStyleToActiveCell('numberFormat', 'Milésimas');
            break;
        case 'Bordes':
            setBorderMenuVisible(prev => !prev);
            break;
        case 'Aumentar Decimal':
        case 'Disminuir Decimal':
        case 'Formato Condicional':
        case 'Insertar Celda':
        case 'Ordenar y Filtrar':
        case 'Formato de celdas':
            showMessageBox(`Acción: ${action} - (Lógica no implementada)`);
            break;
        default:
            showMessageBox(`Acción: ${action} - (Lógica no implementada)`);
            break;
    }
  };
  
  // --- Contenido del Ribbon para la Pestaña Inicio ---
  const renderRibbonContent = () => {
    
    // Opciones del menú de bordes (simulando la imagen de Office)
    const borderOptions: { name: BorderOption, icon: string }[] = [
        { name: 'Ninguno', icon: '⧇' },
        { name: 'Borde inferior', icon: '⏏' },
        { name: 'Borde superior', icon: '⏎' },
        { name: 'Borde izquierdo', icon: '⏴' },
        { name: 'Borde derecho', icon: '⏵' },
        { name: 'Todos los bordes', icon: '╬' },
        { name: 'Bordes externos', icon: '⊞' },
        { name: 'Borde exterior grueso', icon: '▩' },
    ];
    
    const BorderMenu = () => (
        <div className="border-dropdown-menu">
            {borderOptions.map((option) => (
                <div 
                    key={option.name} 
                    className="dropdown-item"
                    onClick={(_) => {
                        _.stopPropagation();
                        applyBorderStyle(option.name);
                    }}
                >
                    <span className="dropdown-icon">{option.icon}</span>
                    {option.name}
                </div>
            ))}
        </div>
    );

    switch (activeTab) {
      case 'Inicio':
        return (
          <div className="ribbon-content">
            {/* GRUPO: PORTAPAPELES (Inicio) */}
            <div className="toolbar-group">
                <button onClick={() => handleClipboardAction("Pegar")} title="Pegar" className="large-button">
                    <span className="icon-xl">📋</span><br/>Pegar
                </button>
                <div className="vertical-group">
                    <button onClick={() => handleClipboardAction("Cortar")} title="Cortar">✂️</button>
                    <button onClick={() => handleClipboardAction("Copiar")} title="Copiar">📝</button>
                </div>
                <div className="group-label">Portapapeles</div>
            </div>
            
            {/* GRUPO: FUENTE (Inicio) */}
            <div className="toolbar-group">
                <div className="horizontal-group input-row">
                    <select 
                        value={currentCellStyles.fontFamily} 
                        onChange={(e) => applyStyleToActiveCell('fontFamily', e.target.value)}
                        title="Fuente" 
                        className="font-select"
                    >
                        <option>Aptos Narrow</option>
                        <option>Arial</option>
                        <option>Verdana</option>
                    </select>
                    <select 
                        value={currentCellStyles.fontSize?.replace('px', '')} 
                        onChange={(e) => applyStyleToActiveCell('fontSize', `${e.target.value}px`)}
                        title="Tamaño" 
                        className="size-select"
                    >
                        {[8, 10, 11, 12, 14, 18, 24].map(s => <option key={s}>{s}</option>)}
                    </select>
                </div>
                <div className="horizontal-group">
                    {/* Botón Negrita */}
                    <button 
                        onClick={() => handleToolbarAction("Negrita")} 
                        title="Negrita (Ctrl+N)"
                        className={currentCellStyles.fontWeight === 'bold' ? 'active-style' : ''}
                    >
                        <b>N</b>
                    </button>
                    {/* Botón Cursiva */}
                    <button 
                        onClick={() => handleToolbarAction("Cursiva")} 
                        title="Cursiva (Ctrl+K)"
                        className={currentCellStyles.fontStyle === 'italic' ? 'active-style' : ''}
                    >
                        <i>K</i>
                    </button>
                    {/* Botón Subrayado */}
                    <button 
                        onClick={() => handleToolbarAction("Subrayado")} 
                        title="Subrayado (Ctrl+S)"
                        className={currentCellStyles.textDecoration === 'underline' ? 'active-style' : ''}
                    >
                        <u>S</u>
                    </button>
                    
                    {/* BOTÓN DE BORDES */}
                    <div className="border-button-wrapper">
                         <button 
                            onClick={(_) => {
                                _.stopPropagation(); // Evita que hideContextMenu se active inmediatamente
                                handleToolbarAction("Bordes");
                            }} 
                            title="Bordes" 
                            className={`border-button ${borderMenuVisible ? 'active-style' : ''}`}
                        >
                            <span className="border-icon">▣</span>
                        </button>
                        {borderMenuVisible && <BorderMenu />}
                    </div>

                    {/* Color de Relleno */}
                    <div className="color-picker-wrapper">
                        <input 
                            type="color" 
                            value={currentCellStyles.backgroundColor === 'inherit' ? selectedFillColor : currentCellStyles.backgroundColor || selectedFillColor} 
                            onChange={(e) => {
                                setSelectedFillColor(e.target.value); 
                                applyStyleToActiveCell('backgroundColor', e.target.value);
                            }}
                            className="color-input"
                            id="fillColorPicker"
                        />
                        <button 
                            onClick={(_) => document.getElementById('fillColorPicker')?.click()} 
                            title="Color de Relleno" 
                            style={{backgroundColor: currentCellStyles.backgroundColor || selectedFillColor}} 
                            className="color-button fill-color-indicator"
                        >
                            🎨
                        </button>
                    </div>

                    {/* Color de Fuente */}
                    <div className="color-picker-wrapper">
                         <input 
                            type="color" 
                            value={currentCellStyles.color === 'var(--excel-text)' ? selectedTextColor : currentCellStyles.color || selectedTextColor} 
                            onChange={(e) => {
                                setSelectedTextColor(e.target.value); 
                                applyStyleToActiveCell('color', e.target.value);
                            }}
                            className="color-input"
                            id="textColorPicker"
                        />
                        <button 
                            onClick={(_) => document.getElementById('textColorPicker')?.click()} 
                            title="Color de Fuente" 
                            style={{color: currentCellStyles.color || selectedTextColor}} 
                            className="color-button text-color-button"
                        >
                            🅰️
                        </button>
                    </div>
                </div>
                <div className="group-label">Fuente</div>
            </div>
            
            {/* GRUPO: ALINEACIÓN (Inicio) */}
            <div className="toolbar-group">
                {/* Alineación Horizontal */}
                <div className="horizontal-group input-row">
                    <button 
                        onClick={() => handleToolbarAction("Alinear Izquierda")} 
                        title="Alinear texto a la izquierda"
                        className={currentCellStyles.textAlign === 'left' ? 'active-style' : ''}
                    >
                        &#x2261;
                    </button>
                    <button 
                        onClick={() => handleToolbarAction("Alinear Centro")} 
                        title="Centrar texto"
                        className={currentCellStyles.textAlign === 'center' ? 'active-style' : ''}
                    >
                        &#x2263; 
                    </button>
                    <button 
                        onClick={() => handleToolbarAction("Alinear Derecha")} 
                        title="Alinear texto a la derecha"
                        className={currentCellStyles.textAlign === 'right' ? 'active-style' : ''}
                    >
                        &#x2261;
                    </button>
                </div>
                <div className="group-label">Alineación</div>
            </div>
            
            {/* GRUPO: NÚMERO (Inicio) */}
            <div className="toolbar-group">
                <select 
                    value={currentCellStyles.numberFormat}
                    onChange={(e) => applyStyleToActiveCell('numberFormat', e.target.value)} 
                    title="Formato de Número" 
                    className="number-select"
                >
                    <option>General</option>
                    <option>Moneda</option>
                    <option>Porcentaje</option>
                    <option>Milésimas</option>
                </select>
                <div className="horizontal-group">
                    <button onClick={() => handleToolbarAction("Moneda")} title="Formato de contabilidad">💲</button>
                    <button onClick={() => handleToolbarAction("Porcentaje")} title="Estilo porcentual"> % </button>
                    <button onClick={() => handleToolbarAction("Milésimas")} title="Estilo de millares"> .00 </button>
                    <button onClick={() => handleToolbarAction("Aumentar Decimal")} title="Aumentar decimales"> .0 </button>
                    <button onClick={() => handleToolbarAction("Disminuir Decimal")} title="Disminuir decimales"> 0. </button>
                </div>
                <div className="group-label">Número</div>
            </div>
            {/* GRUPO: ESTILOS, CELDAS, EDICIÓN */}
             <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Formato Condicional")} title="Formato Condicional" className="large-button">
                    <span className="icon-xl">📊</span><br/>Estilos
                </button>
                <div className="group-label">Estilos</div>
            </div>
             <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Insertar Celda")} title="Insertar Celda" className="large-button">
                    <span className="icon-xl">➕</span><br/>Celdas
                </button>
                <div className="group-label">Celdas</div>
            </div>
             <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Ordenar y Filtrar")} title="Ordenar y Filtrar" className="large-button">
                    <span className="icon-xl">⬇️⬆️</span><br/>Edición
                </button>
                <div className="group-label">Edición</div>
            </div>
          </div>
        );
      
      // Contenido básico para otras pestañas no implementadas
      default:
        return (
            <div className="ribbon-content">
                <div className="toolbar-group">
                    <button onClick={() => showMessageBox(`Has seleccionado la pestaña ${activeTab}`)} title={`Acción en ${activeTab}`} className="large-button">
                        <span className="icon-xl">🛠️</span><br/>{activeTab}
                    </button>
                    <div className="group-label">Funcionalidad</div>
                </div>
            </div>
        );
    }
  };
  
  // --- Componente de la Vista "Guardar como" ---
  const SaveAsView: React.FC = () => {
    const [name, setName] = useState(fileName.split('.')[0] || 'Libro Nuevo');
    const [format, setFormat] = useState<'xlsx' | 'pacur'>('xlsx');

    return (
        <div className="backstage-content save-as-view">
            <h2 className="greeting">Guardar como</h2>
            <div className="save-container">
                <h3 className="section-title">Guardar en Pacur</h3>
                <div className="save-form-group">
                    <label>Nombre del archivo:</label>
                    <input 
                        type="text" 
                        value={name} 
                        onChange={(e) => setName(e.target.value)} 
                        className="backstage-input"
                    />
                </div>
                <div className="save-form-group">
                    <label>Tipo de archivo:</label>
                    <select 
                        value={format} 
                        onChange={(e) => setFormat(e.target.value as 'xlsx' | 'pacur')}
                        className="backstage-select"
                    >
                        <option value="xlsx">Libro de Excel (*.xlsx)</option>
                        <option value="pacur">Libro de Pacur (*.pacur)</option>
                    </select>
                </div>
                <button 
                    className="backstage-save-button"
                    onClick={() => name ? handleSaveAs(name, format) : showMessageBox("El nombre del archivo no puede estar vacío.", true)}
                >
                    Guardar
                </button>
                
                <h3 className="section-title" style={{marginTop: '30px'}}>Otras ubicaciones</h3>
                <p style={{color: '#aaa'}}>Compartidos, OneDrive, Este PC...</p>
            </div>
        </div>
    );
  };
  
  // --- Componente de la Vista "Imprimir" ---
  const PrintView: React.FC = () => {
      const handlePrint = () => {
          // Lógica simple para validar el rango: A1:B10
          const rangePattern = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i;
          if (printRange.trim() === 'Hoja actual') {
              showMessageBox("Imprimiendo la hoja actual completa... (Simulación)");
          } else if (rangePattern.test(printRange.trim())) {
              showMessageBox(`Imprimiendo el rango: ${printRange.trim()} (Simulación)`);
          } else {
              showMessageBox("Rango de impresión inválido. Usa el formato CELDA_INICIO:CELDA_FIN (ej: 1A:5G)", true);
              return;
          }
          setBackstageVisible(false);
          setActiveTab('Inicio');
      }
      
      return (
          <div className="backstage-print-overlay">
              <div className="backstage-print-content">
                  {/* Barra Superior de la vista de impresión */}
                  <div className="print-header">
                      <span>Configuración de impresión</span>
                      <span className="page-info">Total: 1 página</span>
                      <div className="print-actions">
                          <button onClick={() => setBackstageVisible(false)} className="print-cancel-button">CANCELAR</button>
                          <button onClick={handlePrint} className="print-submit-button">IMPRIMIR</button>
                      </div>
                  </div>
                  
                  {/* Área principal: Vista previa y Opciones */}
                  <div className="print-main-area">
                      {/* Vista previa (Simulada) */}
                      <div className="print-preview-pane">
                          <div className="print-paper">
                              <span style={{textAlign: 'center', color: '#999'}}>Vista Previa del Rango: {printRange}</span>
                          </div>
                      </div>
                      
                      {/* Opciones de Impresión */}
                      <div className="print-options-pane">
                          <div className="print-option-group">
                              <label>Qué imprimir</label>
                              <input 
                                  type="text" 
                                  value={printRange}
                                  onChange={(e) => setPrintRange(e.target.value)}
                                  placeholder="Ej: Hoja actual, 1A:5G"
                                  className="print-input"
                              />
                          </div>
                          <div className="print-option-group">
                              <label>Orientación de la página</label>
                              <div className="radio-group">
                                  <label><input type="radio" name="orientation" defaultChecked/> Horizontal</label>
                                  <label><input type="radio" name="orientation"/> Vertical</label>
                              </div>
                          </div>
                          
                          <div className="print-option-group">
                              <label>Tamaño del papel</label>
                              <select className="print-select">
                                  <option>A4 (21 cm x 29.7 cm)</option>
                                  <option>Carta</option>
                              </select>
                          </div>
                          
                           <div className="print-option-group">
                              <label>Márgenes</label>
                              <select className="print-select">
                                  <option>Normales</option>
                                  <option>Estrechos</option>
                              </select>
                          </div>
                          
                           <button className="print-link">ESTABLECER SALTOS DE PÁGINA PERSON.</button>

                      </div>
                  </div>
              </div>
          </div>
      );
  }

  // --- Renderizado del Backstage (Menú Archivo) ---
  const renderBackstage = () => {
    if (!backstageVisible) return null;

    if (backstageView === 'Imprimir') {
        return <PrintView />;
    }

    const navItems: { name: string, view: BackstageView, action?: () => void }[] = [
        { name: 'Nuevo', view: 'Inicio' },
        { name: 'Abrir', view: 'Abrir', action: handleOpen },
        { name: 'Guardar', view: 'Inicio', action: handleSave }, // Guarda y vuelve a Inicio
        { name: 'Guardar como', view: 'Guardar como' },
        { name: 'Imprimir', view: 'Imprimir' },
        { name: 'Exportar', view: 'Opciones' },
        { name: 'Cerrar', view: 'Opciones' },
        { name: 'Cuenta', view: 'Opciones' },
        { name: 'Opciones', view: 'Opciones' },
    ];
    
    // Contenido de la vista Backstage principal (Inicio, Abrir, Opciones)
    const MainBackstageContent = () => {
        if (backstageView === 'Guardar como') {
            return <SaveAsView />;
        }
        
        return (
            <div className="backstage-content">
                <h2 className="greeting">
                    {backstageView === 'Inicio' ? 'Buenas tardes' : backstageView}
                    {isDirty && <span className="dirty-indicator"> • Cambios sin guardar</span>}
                </h2>
                
                {/* Contenido de la vista Abrir */}
                {backstageView === 'Abrir' && (
                    <>
                        <h3 className="section-title">Archivos Recientes</h3>
                        <p style={{color: '#aaa'}}>Haz clic en **Abrir** a la izquierda para cargar el último archivo guardado.</p>
                        <div className="recent-list">
                            <p className="recent-file" onClick={handleOpen}>Libro Anterior.xlsx - (Hacer clic para Cargar)</p>
                            <p>Reporte Mensual.xlsx - Ayer</p>
                        </div>
                    </>
                )}
                
                {/* Contenido de la vista Inicio (Nuevo) */}
                {backstageView === 'Inicio' && (
                    <>
                        <h3 className="section-title">Nuevo</h3>
                        <div className="template-grid">
                            <div className="template-card" onClick={() => { setData(INITIAL_STATE.data); setCellStyles(INITIAL_STATE.styles); setFileName(INITIAL_STATE.fileName); setIsDirty(false); setBackstageVisible(false); }}>
                                <div className="template-icon">📄</div>
                                Libro en blanco
                            </div>
                            <div className="template-card">
                                <div className="template-icon">💡</div>
                                Realizar un recorrido
                            </div>
                            <div className="template-card">
                                <div className="template-icon">∑</div>
                                Tutorial de fórmula
                            </div>
                            <div className="template-card">
                                <div className="template-icon">📊</div>
                                Tabla dinámica
                            </div>
                        </div>
                        <h3 className="section-title">Recientes</h3>
                        <div className="recent-list">
                            <p>Documento1.xlsx - Hoy</p>
                            <p>Reporte Mensual.xlsx - Ayer</p>
                        </div>
                    </>
                )}
            </div>
        );
    }


    // Cuando el backstage está visible, forzamos que la pestaña "Archivo" esté activa
    // y oscurecemos el contenido principal
    return (
        <div className="backstage-overlay">
            <div className="backstage-menu">
                {navItems.map(item => (
                    <div 
                        key={item.name} 
                        className={`backstage-item ${backstageView === item.view && item.name !== 'Guardar' ? 'active-item' : ''}`}
                        onClick={() => {
                            if (item.action) {
                                item.action();
                            } else {
                                setBackstageView(item.view);
                            }
                        }}
                    >
                        {item.name}
                    </div>
                ))}
            </div>
            {MainBackstageContent()}
        </div>
    );
  };


  return (
    <div className="pacur-hoja-container" onClick={hideContextMenu}>
    <style>{`
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
            --excel-select-bg: #2a3a5a; 
            --excel-context-bg: #363636;
            --excel-context-hover: #0078d4;
            --excel-backstage-bg: #0d0d0d; /* Fondo más oscuro para el backstage */
            --excel-border-color: #f0f0f0; /* Color claro para los bordes definidos */
            --excel-border-thick: 2px solid var(--excel-border-color);
            --excel-border-thin: 1px solid var(--excel-border-color);
        }

        body, html, #root {
            margin: 0;
            padding: 0;
            height: 100%;
            width: 100%;
            font-family: 'Aptos Narrow', 'Inter', sans-serif;
            background-color: var(--excel-dark-bg);
            color: var(--excel-text);
        }

        .pacur-hoja-container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            width: 100vw;
            overflow: hidden;
            background-color: var(--excel-dark-bg);
            position: relative; /* Necesario para el overlay del backstage */
        }
        
        .workbook-title {
            margin-left: 20px;
            font-weight: 300;
            font-size: 0.9rem;
            color: #ddd;
        }

        /* 1. RIBBON (Barra de Herramientas Superior) */
        .ribbon {
            background-color: var(--excel-mid-bg);
            border-bottom: 1px solid var(--excel-grid-line);
            padding-top: 5px;
            user-select: none;
            flex-shrink: 0;
            z-index: 100; /* Asegura que esté sobre la cuadrícula */
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
            position: relative;
        }
        
        /* Estilo especial para la pestaña Archivo */
        .ribbon-tab.file-tab {
            background-color: #107c41; /* Verde de Office */
            color: white;
            font-weight: bold;
            border-radius: 4px;
            padding: 8px 20px;
            margin-right: 15px; /* Más espacio después de Archivo */
        }
        
        .ribbon-tab.active {
            background-color: var(--excel-dark-bg);
            border-bottom: 2px solid var(--excel-active-tab-indicator);
            font-weight: 600;
            padding-bottom: 7px;
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
        
        .toolbar-group button {
            background-color: transparent;
            border: 1px solid transparent;
            color: var(--excel-text);
            padding: 3px 5px;
            margin: 1px;
            border-radius: 3px;
            cursor: pointer;
            transition: background-color 0.1s;
        }

        .toolbar-group button:hover {
            background-color: var(--excel-button-hover);
        }
        
        .icon-xl {
            font-size: 1.2rem;
            line-height: 1;
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
            cursor: pointer;
        }

        .font-select { width: 120px; }
        .size-select { width: 50px; }
        .number-select { width: 90px; }

        /* Estilos para Botones de Formato Activos (N, K, S, Alineación) */
        .active-style {
            background-color: var(--excel-active-tab) !important;
            border-color: var(--excel-active-tab) !important;
        }
        
        /* Estilos para Selectores de Color */
        .color-picker-wrapper {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-left: 5px;
        }

        .color-input {
            position: absolute;
            width: 0;
            height: 0;
            opacity: 0;
            pointer-events: none; 
        }

        .color-button {
            padding: 5px 8px !important; 
            font-size: 1rem !important;
            line-height: 1 !important;
            transition: none !important;
        }
        
        .fill-color-indicator {
            border-bottom: 3px solid white; 
        }

        .text-color-button {
            font-size: 1.2rem !important;
            font-weight: bold;
        }
        
        .data-cell {
            /* Aplicar estilos de celda directamente */
            transition: background-color 0.1s, border-color 0.1s;
        }

        /* Estilos de Bordes */
        .border-button-wrapper {
            position: relative;
            display: inline-block;
        }
        
        .border-icon {
            font-size: 1rem;
            line-height: 1;
        }

        .border-dropdown-menu {
            position: absolute;
            top: 35px; /* Debajo del botón */
            left: 0;
            background-color: var(--excel-context-bg);
            border: 1px solid #555;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.5);
            z-index: 1001;
            min-width: 150px;
            padding: 5px 0;
            border-radius: 4px;
        }
        
        .dropdown-item {
            padding: 8px 10px;
            cursor: pointer;
            font-size: 0.85rem;
            transition: background-color 0.1s;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .dropdown-item:hover {
            background-color: var(--excel-context-hover);
            color: white;
        }
        
        .dropdown-icon {
            font-size: 1.1rem;
            width: 20px;
            text-align: center;
        }


        /* 2. BARRA DE FÓRMULAS */
        .formula-bar {
            display: flex;
            align-items: center;
            background-color: var(--excel-formula-bar);
            padding: 4px 10px;
            border-bottom: 1px solid var(--excel-grid-line);
            flex-shrink: 0;
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
        
        /* ... Estilos de cuadrícula (header-row, data-row, cell, etc.) se mantienen ... */
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
            
            /* Resetear bordes de cuadrícula por defecto para poder aplicar estilos */
            border-right: 1px solid var(--excel-grid-line);
            border-bottom: 1px solid var(--excel-grid-line);
        }
        
        /* Estilos específicos de borde */
        .border-none {
            /* Mantiene los bordes de la cuadrícula si no hay estilo de borde definido */
            border-right: 1px solid var(--excel-grid-line) !important;
            border-bottom: 1px solid var(--excel-grid-line) !important;
            border-top: none !important;
            border-left: none !important;
        }
        .border-all {
            border: var(--excel-border-thin) !important;
        }
        .border-bottom {
            border-bottom: var(--excel-border-thin) !important;
        }
        .border-top {
            border-top: var(--excel-border-thin) !important;
        }
        .border-left {
            border-left: var(--excel-border-thin) !important;
        }
        .border-right {
            border-right: var(--excel-border-thin) !important;
        }
        .border-outside {
            border: var(--excel-border-thin) !important; /* Simple para simulación */
        }
        .border-thickOutside {
            border: var(--excel-border-thick) !important; /* Borde grueso para simulación */
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
            cursor: pointer;
        }
        
        .data-cell.active {
            outline: 2px solid var(--excel-active-cell);
            background-color: #000;
            z-index: 6;
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

        /* 4. MENÚ CONTEXTUAL */
        .context-menu {
            position: fixed;
            background-color: var(--excel-context-bg);
            border: 1px solid #555;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.5);
            z-index: 1000;
            min-width: 200px;
            padding: 5px 0;
            border-radius: 4px;
        }

        .context-menu-item {
            padding: 8px 15px;
            cursor: pointer;
            font-size: 0.9rem;
            transition: background-color 0.1s;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .context-menu-item:hover {
            background-color: var(--excel-context-hover);
            color: white;
        }
        
        .context-separator {
            height: 1px;
            background-color: #555;
            margin: 5px 0;
        }


        /* 5. BARRA DE ESTADO (Parte Inferior) */
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
            flex-shrink: 0;
        }
        
        /* 6. BACKSTAGE VIEW (Menú Archivo) */
        .backstage-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: flex;
            background-color: var(--excel-backstage-bg); 
            z-index: 500;
        }

        .backstage-menu {
            width: 250px;
            background-color: var(--excel-dark-bg);
            display: flex;
            flex-direction: column;
            border-right: 1px solid var(--excel-grid-line);
            flex-shrink: 0;
        }

        .backstage-item {
            padding: 15px 20px;
            font-size: 1rem;
            cursor: pointer;
            transition: background-color 0.1s;
            color: #ccc;
        }
        
        .backstage-item.active-item {
            background-color: var(--excel-mid-bg);
            color: var(--excel-text);
            border-left: 3px solid var(--excel-active-tab-indicator);
            padding-left: 17px;
            font-weight: 600;
        }

        .backstage-item:hover {
            background-color: var(--excel-mid-bg);
            color: var(--excel-text);
        }

        .backstage-content {
            flex-grow: 1;
            padding: 20px 40px;
            overflow-y: auto;
        }
        
        .greeting {
            font-size: 1.5rem;
            font-weight: 300;
            color: var(--excel-text);
            margin-bottom: 30px;
            border-bottom: 1px solid #444;
            padding-bottom: 10px;
            display: flex;
            align-items: center;
        }
        
        .dirty-indicator {
            font-size: 0.8rem;
            color: #e53e3e;
            margin-left: 10px;
        }

        .section-title {
            font-size: 1.1rem;
            color: #ddd;
            margin-top: 20px;
            margin-bottom: 15px;
        }

        .template-grid {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }

        .template-card {
            width: 150px;
            height: 120px;
            background-color: var(--excel-mid-bg);
            border-radius: 6px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
            cursor: pointer;
            transition: transform 0.1s, box-shadow 0.1s;
            padding: 10px;
            font-size: 0.85rem;
        }

        .template-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.4);
        }
        
        .template-icon {
            font-size: 2.5rem;
            margin-bottom: 10px;
            color: var(--excel-active-tab);
        }
        
        .recent-list p {
            padding: 8px 0;
            border-bottom: 1px solid #333;
            cursor: pointer;
            transition: color 0.1s;
        }
        
        .recent-list p:hover {
            color: var(--excel-active-tab);
        }
        
        /* Estilos de Guardar como */
        .save-container {
            max-width: 400px;
        }
        
        .save-form-group {
            margin-bottom: 15px;
        }
        
        .save-form-group label {
            display: block;
            margin-bottom: 5px;
            font-size: 0.9rem;
            color: #ccc;
        }
        
        .backstage-input, .backstage-select {
            width: 100%;
            padding: 8px 10px;
            background-color: var(--excel-mid-bg);
            border: 1px solid #555;
            border-radius: 4px;
            color: var(--excel-text);
            font-size: 1rem;
        }
        
        .backstage-save-button {
            background-color: #107c41;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            font-size: 1rem;
            cursor: pointer;
            transition: background-color 0.1s;
        }
        .backstage-save-button:hover {
            background-color: #0c5e31;
        }

        /* Estilos de Imprimir (PrintView) - Basado en la imagen */
        .backstage-print-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: flex;
            background-color: #ccc; /* Fondo gris claro */
            z-index: 500;
        }
        
        .backstage-print-content {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
        }
        
        .print-header {
            background-color: white;
            color: black;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 1.1rem;
            font-weight: 500;
            flex-shrink: 0;
        }
        
        .page-info {
            font-size: 0.9rem;
            font-weight: 300;
            margin-left: 20px;
            color: #666;
        }
        
        .print-actions {
            display: flex;
            gap: 10px;
        }
        
        .print-cancel-button {
            background-color: transparent;
            color: #333;
            border: none;
            padding: 8px 15px;
            cursor: pointer;
        }
        
        .print-submit-button {
            background-color: var(--excel-active-tab);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 15px;
            font-weight: 500;
            cursor: pointer;
        }

        .print-main-area {
            flex-grow: 1;
            display: flex;
        }
        
        .print-preview-pane {
            flex-grow: 1;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        
        .print-paper {
            width: 700px; 
            height: 1000px;
            max-width: 90%;
            max-height: 90%;
            background-color: white;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
            border: 1px solid #aaa;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        .print-options-pane {
            width: 300px;
            min-width: 250px;
            background-color: #f0f0f0; /* Panel de opciones más claro */
            color: #333;
            padding: 20px;
            overflow-y: auto;
            border-left: 1px solid #aaa;
            flex-shrink: 0;
        }
        
        .print-option-group {
            margin-bottom: 20px;
        }
        
        .print-option-group label {
            display: block;
            font-weight: 600;
            margin-bottom: 5px;
            font-size: 0.9rem;
        }
        
        .print-input, .print-select {
            width: 100%;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            background-color: white;
            color: #333;
            font-size: 0.9rem;
        }
        
        .radio-group label {
            font-weight: normal;
            display: block;
            margin-bottom: 5px;
        }
        
        .print-link {
            color: var(--excel-active-tab);
            background: none;
            border: none;
            cursor: pointer;
            padding: 0;
            font-size: 0.8rem;
            margin-top: 10px;
            text-decoration: underline;
        }

    `}</style>
      
      {/* 6. Renderizar Backstage View (si está activo) */}
      {renderBackstage()}

      {/* 1. Barra de Herramientas (Ribbon) */}
      <div className="toolbar ribbon" style={{ visibility: backstageVisible ? 'hidden' : 'visible' }}>
        <div className="ribbon-tabs">
            {['Archivo', 'Inicio', 'Insertar', 'Disposicion', 'Formulas', 'Datos', 'Revisar', 'Vista', 'Ayuda'].map(tab => (
                 <span 
                    key={tab}
                    className={`ribbon-tab ${tab === 'Archivo' ? 'file-tab' : ''} ${activeTab === tab ? 'active' : ''}`}
                    onClick={() => {
                        if (tab === 'Archivo') {
                            setBackstageView('Inicio'); // Siempre ir a Inicio al entrar en Archivo
                            setBackstageVisible(true);
                        } else {
                            setBackstageVisible(false);
                            setActiveTab(tab as RibbonTab);
                        }
                    }}
                >
                    {tab}
                </span>
            ))}
            
            {/* Título y Acciones de la derecha */}
            <span className="workbook-title">{fileName} {isDirty ? '(Modificado)' : ''}</span>

            <div style={{marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center'}}>
                <button onClick={() => showMessageBox("Comentarios - (Lógica no implementada)")} title="Comentarios">💬 Comentarios</button>
                <button onClick={() => showMessageBox("Compartir - (Lógica no implementada)")} title="Compartir" style={{backgroundColor: '#107c41'}}>➡️ Compartir</button>
            </div>
        </div>
        {renderRibbonContent()}
      </div>

      {/* 2. Barra de Fórmulas */}
      <div className="formula-bar" style={{ visibility: backstageVisible ? 'hidden' : 'visible' }}>
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
      <div className="spreadsheet-grid" onContextMenu={(_) => {/* Lógica de menú contextual */}}>
        <div 
            style={{
                transform: `scale(${zoomLevel / 100})`, 
                transformOrigin: 'top left', 
                minWidth: `${(COLS * 80) + 50}px`, 
                minHeight: `${(ROWS * 25) + 25}px`,
                // Ocultar la cuadrícula cuando el backstage está visible
                visibility: backstageVisible ? 'hidden' : 'visible'
            }}
        >
            
            {/* Encabezados de Columna */}
            <div className="header-row">
                <div className="cell header-cell corner-cell" onClick={() => setActiveCell(null)}></div>
                {colHeaders.map(header => (
                    <div 
                        key={header} 
                        className={`cell header-cell ${selectedCols.includes(header) ? 'selected' : ''}`}
                        onClick={() => {/* Lógica de selección de columna */}}
                        onContextMenu={(_) => {/* Lógica de menú contextual */}}
                    >
                        {header}
                    </div>
                ))}
            </div>

            {/* Filas de Datos */}
            {Array.from({ length: ROWS }, (_, rIndex) => {
                const rowIndex = rIndex + 1;
                const isRowSelected = selectedRows.includes(rowIndex);
                
                return (
                    <div key={rIndex} className="data-row">
                        {/* Encabezado de fila */}
                        <div 
                            className={`cell header-cell ${isRowSelected ? 'selected' : ''}`}
                            onClick={() => {/* Lógica de selección de fila */}}
                            onContextMenu={(_) => {/* Lógica de menú contextual */}}
                        >
                            {rowIndex}
                        </div>
                        
                        {colHeaders.map(cHeader => {
                            const cellKey = `${cHeader}${rowIndex}`;
                            // Usar 'key' para obtener el formato correcto en calculateValue
                            const displayValue = calculateValue(cellKey);
                            const isSelected = isRowSelected || selectedCols.includes(cHeader);
                            const styles = { ...defaultStyles, ...(cellStyles[cellKey] || {}) };
                            
                            // Determinar la clase de borde para aplicar los estilos visuales
                            const borderClass = `border-${styles.borderStyle || 'none'}`;
                            
                            return (
                                <div 
                                    key={cellKey}
                                    className={`cell data-cell ${activeCell === cellKey ? 'active' : ''} ${isSelected ? 'selected-cell' : ''} ${borderClass}`}
                                    onClick={() => setActiveCell(cellKey)}
                                    onContextMenu={(_) => {/* Lógica de menú contextual */}}
                                    style={{
                                        fontWeight: styles.fontWeight,
                                        fontStyle: styles.fontStyle,
                                        textDecoration: styles.textDecoration,
                                        textAlign: styles.textAlign,
                                        backgroundColor: styles.backgroundColor === 'inherit' ? 'var(--excel-dark-bg)' : styles.backgroundColor,
                                        color: styles.color === 'var(--excel-text)' ? 'var(--excel-text)' : styles.color,
                                        fontSize: styles.fontSize,
                                        fontFamily: styles.fontFamily,
                                    }}
                                >
                                    {displayValue}
                                </div>
                            );
                        })}
                    </div>
                );
            })}
        </div>
      </div>
      
      {/* 5. Barra de Estado (Parte Inferior) */}
      <div className="status-bar" style={{ visibility: backstageVisible ? 'hidden' : 'visible' }}>
        <div className="status-left">
            <span>Listo</span>
            <div className="sheet-tabs">
                <span className="sheet-tab active-sheet">Hoja1</span>
                <button onClick={() => showMessageBox("Añadir Hoja - (Lógica no implementada)")} title="Nueva Hoja">➕</button>
            </div>
        </div>
        <div className="status-right">
            <div className="zoom-control">
                <button onClick={() => setZoomLevel(Math.max(10, zoomLevel - 10))} title="Alejar">➖</button>
                <span>{zoomLevel}%</span>
                <button onClick={() => setZoomLevel(Math.min(200, zoomLevel + 10))} title="Acercar">➕</button>
            </div>
        </div>
      </div>
      
      {/* 5. Menú Contextual (Solo visible si el backstage NO está visible) */}
      {contextMenu.visible && !backstageVisible && (
        <div 
            className="context-menu" 
            style={{ left: contextMenu.x, top: contextMenu.y }}
        >
            <div className="context-menu-item" onClick={() => handleToolbarAction("Cortar")}>✂️ Cortar</div>
            <div className="context-menu-item" onClick={() => handleToolbarAction("Copiar")}>📝 Copiar</div>
            <div className="context-menu-item" onClick={() => handleToolbarAction("Pegar")}>📋 Opciones de pegado...</div>
            
            <div className="context-separator"></div>
            
            <div className="context-menu-item" onClick={() => handleToolbarAction("Formato de celdas")}>⚙️ Formato de celdas...</div>
        </div>
      )}
    </div>
  );
};

// Función auxiliar para parsear rangos (simplificado)
const parseRange = (range: string): string[] => {
    const parts = range.split(':');
    if (parts.length !== 2) return [];

    const [start, end] = parts;

    // Función rudimentaria para obtener coordenadas (Columna, Fila) de una clave de celda (ej: A1)
    const cellToCoords = (cellKey: string): [number, number] | null => {
        const match = cellKey.match(/^([A-Z]+)([0-9]+)$/i);
        if (!match) return null;

        const colStr = match[1].toUpperCase();
        const rowNum = parseInt(match[2], 10);

        let colIndex = 0;
        for (let i = 0; i < colStr.length; i++) {
            colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return [colIndex, rowNum];
    };
    
    // Función rudimentaria para obtener clave de celda de coordenadas
    const coordsToCell = (colIndex: number, rowIndex: number): string => {
        let colStr = '';
        let num = colIndex;
        while (num > 0) {
            let remainder = (num - 1) % 26;
            colStr = String.fromCharCode('A'.charCodeAt(0) + remainder) + colStr;
            num = Math.floor((num - 1) / 26);
        }
        return `${colStr}${rowIndex}`;
    };

    const startCoords = cellToCoords(start);
    const endCoords = cellToCoords(end);

    if (!startCoords || !endCoords) return [];

    const [startCol, startRow] = startCoords;
    const [endCol, endRow] = endCoords;
    
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);

    const keys: string[] = [];
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            keys.push(coordsToCell(c, r));
        }
    }
    return keys;
};

export default PacurHoja;
