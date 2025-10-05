import React, { useState, useRef, useEffect, useCallback } from 'react';

// Carga de iconos de Font Awesome para el Ribbon
// Necesitarás agregar esta línea a tu HTML principal si no usas una herramienta de empaquetado:
// <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" xintegrity="sha512-SnH5WK+bZxgPHs44uWIX+LLMDJ8yKzS0Qj1VlPstn3t0x12oN/cE2O/C/0P1uD2A" crossorigin="anonymous" referrerpolicy="no-referrer" />

// --- TIPOS Y CONSTANTES ---

// Define las propiedades de estilo que puede tener una celda
type CellStyle = {
    fontFamily?: string;
    fontSize?: number;
    fontWeight?: 'normal' | 'bold';
    fontStyle?: 'normal' | 'italic';
    textDecoration?: 'none' | 'underline' | 'line-through';
    textAlign?: 'left' | 'center' | 'right';
    backgroundColor?: string;
    color?: string;
    borderStyle?: 'none' | 'all' | 'bottom' | 'top' | 'left' | 'right' | 'outside' | 'thickOutside';
    numberFormat?: 'General' | 'Currency' | 'Percentage' | 'Thousands';
    // Estilos de la tabla de ejemplo
    padding?: string;
    borderRadius?: string;
    border?: string;
};

// Define la estructura de los datos de la celda
type CellData = {
    value: string;
    formula?: string; // Para almacenar la fórmula original si empieza con '='
    calculatedValue?: string | number; // Resultado de la fórmula
    styles: CellStyle;
};

// Mapa de celdas: clave es "A1", "B5", etc.
type SheetData = Record<string, CellData>;

// Estado del documento (para simular Guardar/Abrir)
interface DocumentState {
    fileName: string;
    isDirty: boolean;
    data: SheetData;
    activeSheet: string;
}

// Interfaz para la vista de impresión
interface PrintSettings {
    range: string;
    orientation: 'vertical' | 'horizontal';
}

const DEFAULT_ROWS = 20;
const DEFAULT_COLS = 15;
const COL_HEADERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

// --- UTILIDADES ---

/**
 * Convierte coordenadas (row, col) a notación de celda (A1, B5)
 * @param row Índice de fila (base 0)
 * @param col Índice de columna (base 0)
 * @returns Notación de celda (string)
 */
const coordsToCell = (row: number, col: number): string => {
    let colStr = '';
    let c = col;
    while (c >= 0) {
        colStr = COL_HEADERS[c % 26] + colStr;
        c = Math.floor(c / 26) - 1;
    }
    return `${colStr}${row + 1}`;
};

/**
 * Convierte notación de celda (A1, B5) a coordenadas (row, col)
 * @param cellId Notación de celda (string)
 * @returns Coordenadas [row, col] (base 0) o [null, null] si es inválido
 */
const cellToCoords = (cellId: string): [number | null, number | null] => {
    const match = cellId.match(/([A-Z]+)(\d+)/i);
    if (!match) return [null, null];

    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10) - 1;

    let colIndex = 0;
    for (let i = 0; i < colStr.length; i++) {
        colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    colIndex -= 1;

    return [row, colIndex];
};


/**
 * Parsea un rango de celda (ej: A1:B5) y devuelve una lista de claves de celda.
 * @param rangeString Rango de celda (ej: A1:B5)
 * @returns Array de claves de celda (ej: ['A1', 'A2', ...])
 */
const parseRange = (rangeString: string): string[] => {
    if (!rangeString || !rangeString.includes(':')) return [];

    const [startCell, endCell] = rangeString.split(':');
    const [startRow, startCol] = cellToCoords(startCell);
    const [endRow, endCol] = cellToCoords(endCell);

    if (startRow === null || startCol === null || endRow === null || endCol === null) return [];

    const cells: string[] = [];

    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            cells.push(coordsToCell(r, c));
        }
    }
    return cells;
};

/**
 * Calcula el valor de una celda, manejando fórmulas simples y funciones.
 * @param expression Expresión a evaluar (ej: "10+5", "=SUMA(A1:B2)")
 * @param sheetData Datos de la hoja para buscar referencias
 * @returns El valor calculado
 */
const calculateValue = (expression: string, sheetData: SheetData): string | number => {
    if (!expression || expression[0] !== '=') return expression;

    let formula = expression.substring(1).toUpperCase();

    // 1. Manejo de funciones (SUMA, PROMEDIO)
    const functionMatch = formula.match(/^(SUMA|PROMEDIO)\((.+)\)$/);
    if (functionMatch) {
        const funcName = functionMatch[1];
        const rangeString = functionMatch[2];
        const cells = parseRange(rangeString);

        const values = cells.map(cellId => {
            const cell = sheetData[cellId];
            if (cell) {
                // Si la celda tiene un valor calculado, úsalo directamente.
                // Si no, intenta parsear el valor.
                const val = cell.calculatedValue !== undefined
                    ? cell.calculatedValue
                    : cell.value;

                // Intentar convertir el valor a número
                const numVal = parseFloat(String(val).replace(/[^0-9.-]+/g, ""));
                return isNaN(numVal) ? 0 : numVal;
            }
            return 0;
        }).filter(val => !isNaN(val));

        if (funcName === 'SUMA') {
            return values.reduce((sum, current) => sum + current, 0);
        }
        if (funcName === 'PROMEDIO' && values.length > 0) {
            return values.reduce((sum, current) => sum + current, 0) / values.length;
        }
        return `#ERROR!`;
    }

    // 2. Manejo de expresiones matemáticas y referencias de celda
    let evaluatedExpression = formula;

    // Sustituir referencias de celda por sus valores
    evaluatedExpression = evaluatedExpression.replace(/[A-Z]+[0-9]+/g, (match) => {
        const cell = sheetData[match];
        if (cell) {
            // Obtener el valor ya calculado o el valor sin fórmula
            const val = cell.calculatedValue !== undefined
                ? cell.calculatedValue
                : cell.value;

            // Usar solo el número para la evaluación
            return String(val).replace(/[^0-9.-]+/g, "");
        }
        return '0';
    });

    try {
        // eslint-disable-next-line no-new-func
        const result = new Function(`return ${evaluatedExpression}`)();
        return isFinite(result) ? result : `#REF!`;
    } catch (error) {
        return `#FÓRMULA INVÁLIDA`;
    }
};

// --- COMPONENTES AUXILIARES PARA VISTA BACKSTAGE ---

// Componente de Vista "Guardar como"
const SaveAsView: React.FC<{
    docState: DocumentState;
    handleSaveAs: (fileName: string, format: string) => void;
    onClose: () => void;
}> = ({ docState, handleSaveAs, onClose }) => {
    const [newFileName, setNewFileName] = useState(docState.fileName);
    const [format, setFormat] = useState('excel');

    const onSave = () => {
        handleSaveAs(newFileName, format);
    }

    return (
        <div className="absolute inset-0 bg-gray-900 dark:bg-[#1e1e1e] p-10 z-50 overflow-auto text-white">
            <h2 className="text-3xl font-light mb-8">Guardar como</h2>
            <div className="bg-gray-800 dark:bg-gray-700 p-6 rounded-lg max-w-xl space-y-4">
                <p className="text-lg">Nombre del Archivo:</p>
                <input
                    type="text"
                    className="w-full p-2 bg-gray-600 border border-gray-500 rounded text-white"
                    value={newFileName}
                    onChange={(e) => setNewFileName(e.target.value)}
                />
                <p className="text-lg">Tipo de Archivo:</p>
                <select
                    className="w-full p-2 bg-gray-600 border border-gray-500 rounded text-white"
                    value={format}
                    onChange={(e) => setFormat(e.target.value)}
                >
                    <option value="excel">Libro de Excel (*.xlsx)</option>
                    <option value="pacur">Archivo PacurHoja (*.pacur)</option>
                    <option value="csv">CSV (Delimitado por comas) (*.csv)</option>
                </select>
                <div className="flex justify-end space-x-3 pt-4">
                     <button
                        className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold"
                        onClick={onSave}
                    >
                        Guardar
                    </button>
                    <button
                        className="px-4 py-2 bg-gray-600 hover:bg-gray-500 rounded-md"
                        onClick={onClose}
                    >
                        Cancelar
                    </button>
                </div>
            </div>
        </div>
    );
};

// Componente de Vista "Imprimir"
const PrintView: React.FC<{
    printSettings: PrintSettings;
    setPrintSettings: React.Dispatch<React.SetStateAction<PrintSettings>>;
    onClose: () => void;
}> = ({ printSettings, setPrintSettings, onClose }) => {

    const isValidRange = (range: string) => {
        return range === 'Hoja actual' || range === 'Selección' || /^[A-Z]+\d+:[A-Z]+\d+$/i.test(range);
    }

    const handlePrint = () => {
        if(isValidRange(printSettings.range)) {
            // Usar console.log en lugar de alert()
            console.log(`[ALERTA] Simulando imprimir rango: ${printSettings.range} en orientación ${printSettings.orientation}.`);
            onClose(); // Cerrar después de simular la impresión
        } else {
            // Usar console.log en lugar de alert()
            console.log('[ALERTA] Rango de impresión inválido.');
        }
    }

    return (
        <div className="absolute inset-0 bg-gray-900 dark:bg-[#1e1e1e] p-10 z-50 overflow-auto text-white flex">
            {/* Panel de Configuración de Impresión (Izquierda) */}
            <div className="w-1/4 bg-gray-800 dark:bg-gray-700 p-6 space-y-6 flex flex-col justify-between rounded-lg">
                <div>
                    <h2 className="text-2xl font-semibold mb-4 border-b border-gray-600 pb-2">Imprimir</h2>
                    <div className="space-y-4">
                        <div>
                            <label className="block text-sm font-medium mb-1">Imprimir:</label>
                            <select
                                className="w-full p-2 bg-gray-600 border border-gray-500 rounded text-white"
                                value={printSettings.range}
                                onChange={(e) => setPrintSettings(p => ({ ...p, range: e.target.value }))}
                            >
                                <option value="Hoja actual">Hoja actual</option>
                                <option value="Selección">Selección</option>
                                <option value="Rango personalizado">Rango personalizado...</option>
                            </select>
                            {printSettings.range === 'Rango personalizado' && (
                                <input
                                    type="text"
                                    className="w-full mt-2 p-2 bg-gray-600 border border-gray-500 rounded text-white text-sm"
                                    placeholder="Ej: A1:G5"
                                    value={printSettings.range}
                                    onChange={(e) => setPrintSettings(p => ({ ...p, range: e.target.value }))}
                                />
                            )}
                        </div>
                        <div>
                            <label className="block text-sm font-medium mb-1">Orientación:</label>
                            <select
                                className="w-full p-2 bg-gray-600 border border-gray-500 rounded text-white"
                                value={printSettings.orientation}
                                onChange={(e) => setPrintSettings(p => ({ ...p, orientation: e.target.value as 'vertical' | 'horizontal' }))}
                            >
                                <option value="vertical">Vertical</option>
                                <option value="horizontal">Horizontal</option>
                            </select>
                        </div>
                        <div className="text-xs text-gray-400 pt-4">
                            Ejemplo de rango personalizado: **A1:G5** imprime el área de A1 a G5.
                        </div>
                    </div>
                </div>

                <div className="pt-6 border-t border-gray-600">
                     <button
                        className={`w-full px-4 py-2 bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold ${!isValidRange(printSettings.range) ? 'opacity-50 cursor-not-allowed' : ''}`}
                        onClick={handlePrint}
                        disabled={!isValidRange(printSettings.range)}
                    >
                        Imprimir
                    </button>
                    <button
                        className="w-full mt-2 px-4 py-2 bg-gray-600 hover:bg-gray-500 rounded-md"
                        onClick={() => { onClose(); setPrintSettings({ range: 'Hoja actual', orientation: 'vertical' }); }} // Reset settings on close
                    >
                        Cerrar Vista Previa
                    </button>
                </div>
            </div>
            {/* Vista Previa de Impresión (Derecha) */}
            <div className="flex-1 ml-6 p-4 bg-gray-200 dark:bg-gray-800 rounded-lg overflow-hidden">
                <div className={`w-full h-full bg-white dark:bg-gray-900 border border-dashed border-gray-400 flex items-center justify-center text-gray-500 ${printSettings.orientation === 'horizontal' ? 'aspect-video' : 'aspect-[3/4]'}`}>
                    Vista Previa de Impresión ({printSettings.range})
                </div>
            </div>
         </div>
    );
};


// --- COMPONENTES AUXILIARES DEL RIBBON (Extraídos para evitar violación de Hooks) ---

interface DropdownProps {
    label: string;
    items: { label: string, action: () => void }[];
    primaryIconName: string;
}

/**
 * Control de menú desplegable para el Ribbon.
 * Usa Hooks (useState, useRef, useEffect) por lo que DEBE estar fuera del componente principal.
 */
const DropdownControl: React.FC<DropdownProps> = ({ label, items, primaryIconName }) => {
     const [isOpen, setIsOpen] = useState(false);
     const dropdownRef = useRef<HTMLDivElement>(null);

     useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
                setIsOpen(false);
            }
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
     }, []);

     return (
         <div className="relative inline-block" ref={dropdownRef}>
             <button
                 title={label}
                 className="p-2 rounded-md hover:bg-gray-600 text-gray-200 flex items-center"
                 onClick={() => setIsOpen(!isOpen)}
             >
                 <i className={`fas fa-${primaryIconName}`}></i>
                 <i className="fas fa-caret-down ml-1 text-xs"></i>
             </button>
             {isOpen && (
                 <div className="absolute top-full left-0 mt-1 w-48 bg-gray-700 border border-gray-600 rounded-md shadow-lg z-50">
                     {items.map(item => (
                         <div
                             key={item.label}
                             className="px-3 py-2 text-sm text-gray-100 hover:bg-indigo-600 cursor-pointer"
                             onClick={() => { item.action(); setIsOpen(false); }}
                         >
                             {item.label}
                         </div>
                     ))}
                 </div>
             )}
         </div>
     );
};

interface ColorPickerProps {
    label: string;
    styleKey: 'backgroundColor' | 'color';
    defaultColor: string;
    activeCell: string | null;
    currentStyles: CellStyle;
    applyStyleToActiveCell: (styleKey: keyof CellStyle, value: any) => void;
}

/**
 * Control para seleccionar color de fuente o relleno.
 * Usa Hooks (useRef) por lo que DEBE estar fuera del componente principal.
 */
const ColorPickerControl: React.FC<ColorPickerProps> = ({
    label, styleKey, defaultColor, activeCell, currentStyles, applyStyleToActiveCell
}) => {
    const colorRef = useRef<HTMLInputElement>(null);
    const initialColor = currentStyles[styleKey] || defaultColor;

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const newColor = e.target.value;
        if(activeCell) {
            applyStyleToActiveCell(styleKey, newColor);
        }
    };

    const isFontColor = styleKey === 'color';

    return (
        <div className="relative inline-block" title={label}>
            <button
                className="p-2 rounded-md hover:bg-gray-600 text-gray-200 text-lg"
                onClick={() => colorRef.current?.click()}
                disabled={!activeCell}
            >
                 {/* Icono con una barra que muestra el color actual */}
                <div className="relative">
                    {/* Usar fa-a para color de fuente, fa-fill-drip para color de relleno */}
                    <i className={`fas ${isFontColor ? 'fa-a' : 'fa-fill-drip'}`}></i>
                    <div
                        className="absolute -bottom-1 left-0 right-0 h-[3px] rounded-full"
                        style={{ backgroundColor: initialColor, border: '1px solid #333' }}
                    ></div>
                </div>
            </button>
            <input
                type="color"
                ref={colorRef}
                value={initialColor}
                onChange={handleChange}
                className="absolute opacity-0 pointer-events-none w-0 h-0"
            />
        </div>
    );
};


// --- COMPONENTE DE CELDA (MOVIDO FUERA DEL COMPONENTE PRINCIPAL) ---

interface CellProps {
    row: number;
    col: number;
    sheetData: SheetData;
    activeCell: string | null;
    selectedRows: Set<number>;
    selectedCols: Set<number>;
    cellInput: string;
    updateCellValue: () => void;
    handleCellClick: (cellId: string) => void;
    handleContextMenu: (e: React.MouseEvent, target: 'cell' | 'row' | 'col', cellId?: string) => void;
    setCellInput: React.Dispatch<React.SetStateAction<string>>;
    handleKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => void;
    inputRef: React.RefObject<HTMLInputElement>;
}

const Cell: React.FC<CellProps> = React.memo(({
    row, col, sheetData, activeCell, selectedRows, selectedCols, cellInput,
    updateCellValue, handleCellClick, handleContextMenu, setCellInput, handleKeyDown, inputRef
}) => {
    const cellId = coordsToCell(row, col);
    const data = sheetData[cellId];
    const isActive = activeCell === cellId;
    const isSelectedRow = selectedRows.has(row);
    const isSelectedCol = selectedCols.has(col);
    const isSelected = isSelectedRow || isSelectedCol;

    // Determinar valor a mostrar
    const displayedValue = data?.calculatedValue !== undefined ? String(data.calculatedValue) : data?.value || '';

    // Aplicar formato de número
    const formatValue = (value: string | number, format: CellStyle['numberFormat']) => {
        let num = typeof value === 'string' ? parseFloat(value.replace(/[^0-9.-]+/g, "")) : value;
        if (isNaN(num)) return value;

        if (format === 'Currency') {
            return num.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2 });
        }
        if (format === 'Percentage') {
            // Si el valor ya es decimal (ej. 0.5), lo multiplicamos por 100
            const numPercent = num < 1 && num > -1 ? num * 100 : num;
            return numPercent.toFixed(2) + '%';
        }
        if (format === 'Thousands') {
             return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
        }
        return String(value);
    };

    const finalDisplayedValue = formatValue(displayedValue, data?.styles.numberFormat);

    // Generar estilos CSS
    const cellStyles: React.CSSProperties = {
        fontFamily: data?.styles.fontFamily || 'Inter, sans-serif',
        fontSize: data?.styles.fontSize ? `${data.styles.fontSize}px` : '14px',
        fontWeight: data?.styles.fontWeight || 'normal',
        fontStyle: data?.styles.fontStyle || 'normal',
        textDecoration: data?.styles.textDecoration || 'none',
        textAlign: data?.styles.textAlign || 'left',
        backgroundColor: data?.styles.backgroundColor || (isSelected ? '#333333' : '#1e1e1e'),
        color: data?.styles.color || '#fff',
        // Borde de celda por defecto (para separación)
        border: isActive ? '2px solid #1a73e8' : '1px solid #333333',
    };

    // Sobreescribir el borde si se aplica un estilo de borde específico
    if (data?.styles.borderStyle === 'all') {
        cellStyles.border = isActive ? '2px solid #1a73e8' : '1px solid #555';
    } else if (data?.styles.borderStyle === 'thickOutside') {
        cellStyles.border = isActive ? '2px solid #1a73e8' : '2px solid #fff';
    }

    return (
        <div
            className={`flex-shrink-0 w-32 h-6 overflow-hidden whitespace-nowrap px-1 py-0.5 transition-colors duration-100 ${isActive ? 'z-20' : ''}`}
            style={cellStyles}
            onClick={() => handleCellClick(cellId)}
            onContextMenu={(e) => handleContextMenu(e, 'cell', cellId)}
        >
            {/* Si la celda está activa, se muestra el input para la edición. 
                Si no está activa, se muestra el valor renderizado.
            */}
            {isActive && activeCell === cellId ? (
                <input
                    ref={inputRef}
                    type="text"
                    className="w-full h-full bg-transparent outline-none border-none text-white text-sm"
                    style={{ ...cellStyles, textAlign: cellStyles.textAlign || 'left' }}
                    // Mostrar la fórmula si existe, de lo contrario, el valor
                    value={cellInput}
                    onChange={(e) => setCellInput(e.target.value)}
                    onBlur={updateCellValue} // Aplicar valor al perder el foco
                    onKeyDown={handleKeyDown}
                    autoFocus
                />
            ) : (
                finalDisplayedValue
            )}
        </div>
    );
});


// --- COMPONENTE PRINCIPAL PacurHoja ---

const PacurHoja: React.FC = () => {
    const [sheetData, setSheetData] = useState<SheetData>({});
    const [activeCell, setActiveCell] = useState<string | null>(null);
    const [cellInput, setCellInput] = useState('');
    const [activeTab, setActiveTab] = useState('Inicio');
    const [backstageVisible, setBackstageVisible] = useState(false);
    const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
    const [selectedCols, setSelectedCols] = useState<Set<number>>(new Set());
    const [docState, setDocState] = useState<DocumentState>({
        fileName: 'Libro1',
        isDirty: false,
        data: {},
        activeSheet: 'Hoja1',
    });
    const [contextMenuVisible, setContextMenuVisible] = useState(false);
    const [contextMenuPosition, setContextMenuPosition] = useState({ x: 0, y: 0 });
    const [contextMenuTarget, setContextMenuTarget] = useState<'cell' | 'row' | 'col' | null>(null);
    const [saveAsVisible, setSaveAsVisible] = useState(false);
    const [printVisible, setPrintVisible] = useState(false);
    const [printSettings, setPrintSettings] = useState<PrintSettings>({
        range: 'Hoja actual',
        orientation: 'vertical',
    });
    const menuRef = useRef<HTMLDivElement>(null);
    const inputRef = useRef<HTMLInputElement>(null);
    const gridRef = useRef<HTMLDivElement>(null);

    // --- MANEJO DE ESTADO Y CÁLCULO ---

    // Función para recálculo completo de la hoja
    const recalculateSheet = useCallback((currentData: SheetData): SheetData => {
        const recalculatedData: SheetData = {};

        Object.keys(currentData).forEach(cellId => {
            const cell = currentData[cellId];
            if (cell.formula) {
                // Si tiene fórmula, calcula el valor
                recalculatedData[cellId] = {
                    ...cell,
                    // Usar currentData para calcular (evitar recursión infinita, pero mantener dependencias)
                    calculatedValue: calculateValue(cell.formula, currentData),
                };
            } else {
                // Si no tiene fórmula, el valor calculado es el valor
                recalculatedData[cellId] = {
                    ...cell,
                    calculatedValue: cell.value,
                };
            }
        });
        return recalculatedData;
    }, []);


    // --- EFECTOS DE CONTROL ---

    // Sincronizar input con celda activa y enfocar el input
    useEffect(() => {
        if (activeCell) {
            const data = sheetData[activeCell];
            // Mostrar la fórmula si existe, de lo contrario, el valor
            setCellInput(data?.formula || data?.value || '');
             // Enfocar el input solo si está visible (manejado dentro de Cell)
             // Nota: El input dentro de Cell está enfocado automáticamente si isActive
        } else {
            setCellInput('');
        }
    }, [activeCell, sheetData]);

    // Ocultar menú contextual al hacer clic fuera
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
                setContextMenuVisible(false);
            }
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [menuRef]);

    // Función que aplica el valor del input a la celda activa y recalcula
    const updateCellValue = useCallback(() => {
        if (!activeCell) return;

        // No hacer nada si el valor no ha cambiado
        const currentData = sheetData[activeCell];
        const currentValue = currentData?.formula || currentData?.value || '';
        if (cellInput === currentValue) {
            return;
        }

        const isFormula = cellInput.startsWith('=');

        // 1. Crear una versión temporal de los datos con el nuevo valor
        const tempUpdatedData = {
            ...sheetData,
            [activeCell]: {
                ...(sheetData[activeCell] || { value: '', styles: {} }),
                value: cellInput,
                formula: isFormula ? cellInput : undefined,
                // El valor calculado de la celda activa se actualizará en la próxima etapa
                calculatedValue: undefined,
            },
        };

        // 2. Recalcular toda la hoja de cálculo usando los datos temporales
        const finalUpdatedData = recalculateSheet(tempUpdatedData);

        setSheetData(finalUpdatedData);
        setDocState(prev => ({ ...prev, isDirty: true })); // Marcar como modificado
    }, [activeCell, cellInput, sheetData, recalculateSheet]);


    // Manejar tecla Enter para salir de la edición de celda
    const handleKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            updateCellValue();
            // Mover la celda activa hacia abajo (simulación básica de Excel)
            if (activeCell) {
                 const [row, col] = cellToCoords(activeCell);
                 if (row !== null && col !== null && row < DEFAULT_ROWS - 1) {
                    setActiveCell(coordsToCell(row + 1, col));
                 } else {
                    // Si es la última fila, desenfocar
                    setActiveCell(null);
                 }
            }
        }
    }, [activeCell, updateCellValue]);


    // --- MANEJADORES DE ESTADO DE ARCHIVO (Guardar/Abrir) ---

    // Función para marcar el documento como modificado
    const setDirty = useCallback((dirty: boolean) => {
        setDocState(prev => ({ ...prev, isDirty: dirty }));
    }, []);

    // Simulación de Guardar (usa localStorage)
    const handleSave = () => {
        if (!docState.isDirty) {
            console.log(`[ALERTA] Archivo '${docState.fileName}' ya está guardado.`);
            return;
        }
        try {
            // Guardamos solo los datos crudos (valor y estilos), no los calculados
            const dataToSave: SheetData = {};
            Object.keys(sheetData).forEach(cellId => {
                const cell = sheetData[cellId];
                dataToSave[cellId] = {
                    value: cell.value,
                    formula: cell.formula,
                    styles: cell.styles
                };
            });

            localStorage.setItem(`pacur_sheet_${docState.fileName}`, JSON.stringify(dataToSave));
            setDirty(false);
            console.log(`Archivo '${docState.fileName}' guardado con éxito.`);
        } catch (error) {
            console.error('Error al guardar:', error);
            console.log('[ALERTA] Error al intentar guardar el archivo.');
        }
    };

    // Simulación de Abrir
    const handleOpen = () => {
        setBackstageVisible(false);
        const savedData = localStorage.getItem(`pacur_sheet_${docState.fileName}`);
        if (savedData) {
            try {
                const loadedSheetData: SheetData = JSON.parse(savedData);
                // Recalcular la hoja después de cargar los datos crudos
                const updatedData = recalculateSheet(loadedSheetData);

                setSheetData(updatedData);
                setDirty(false);
                console.log(`Archivo '${docState.fileName}' abierto.`);
            } catch (error) {
                console.error('Error al cargar datos:', error);
                console.log('[ALERTA] Error al cargar el archivo guardado.');
            }
        } else {
            console.log(`[ALERTA] No se encontró el archivo '${docState.fileName}' en el almacenamiento local.`);
        }
    };

    // Simulación de Guardar Como (muestra la interfaz)
    const handleSaveAs = (newFileName: string, format: string) => {
        const fileExtension = format === 'excel' ? '.xlsx' : format === 'pacur' ? '.pacur' : '.csv';
        console.log(`[ALERTA] Simulando Guardar como: ${newFileName}${fileExtension}.`);
        setDocState(prev => ({ ...prev, fileName: newFileName, isDirty: false }));
        setSaveAsVisible(false);
        setBackstageVisible(false);
    };

    // --- MANEJADORES DE LA HOJA DE CÁLCULO ---

    /**
     * Aplica estilos de formato a la celda activa.
     * @param styleKey Clave del estilo a aplicar (ej: 'fontWeight')
     * @param value Valor del estilo (ej: 'bold', 'center')
     */
    const applyStyleToActiveCell = (styleKey: keyof CellStyle, value: any) => {
        if (!activeCell) return;

        setSheetData(prevData => {
            const currentCell = prevData[activeCell] || { value: '', styles: {} };
            const currentStyles = currentCell.styles;

            let newStyles = { ...currentStyles };

            // Lógica de toggle para estilos binarios (Negrita, Cursiva, Subrayado, Alineación)
            if (['fontWeight', 'fontStyle', 'textDecoration', 'textAlign'].includes(styleKey)) {
                // Si el estilo ya tiene el valor, se desactiva (undefined); si no, se aplica.
                newStyles[styleKey] = (currentStyles[styleKey] === value) ? undefined : value;
            } else {
                // Aplicar color de fuente, color de relleno, tamaño, formato numérico, etc.
                newStyles[styleKey] = value;
            }


            return {
                ...prevData,
                [activeCell]: {
                    ...currentCell,
                    styles: newStyles,
                },
            };
        });
        setDirty(true);
    };

    // --- MANEJADORES DE SELECCIÓN ---

    const handleCellClick = useCallback((cellId: string) => {
        if (contextMenuVisible) setContextMenuVisible(false);
        // Aplicar el valor de la celda previamente activa ANTES de cambiar
        if (activeCell && activeCell !== cellId) {
             updateCellValue();
        }

        setActiveCell(cellId);
        setSelectedRows(new Set());
        setSelectedCols(new Set());
    }, [activeCell, contextMenuVisible, updateCellValue]);

    const handleRowHeaderClick = (row: number) => {
        // Al hacer clic en un encabezado, aplica el valor de la celda activa primero.
        if (activeCell) updateCellValue();

        setSelectedRows(new Set([row]));
        setSelectedCols(new Set());
        setActiveCell(coordsToCell(row, 0)); // Mover el foco al inicio de la fila
        setContextMenuVisible(false);
    };

    const handleColHeaderClick = (col: number) => {
        if (activeCell) updateCellValue();

        setSelectedCols(new Set([col]));
        setSelectedRows(new Set());
        setActiveCell(coordsToCell(0, col)); // Mover el foco al inicio de la columna
        setContextMenuVisible(false);
    };

    // --- MENÚ CONTEXTUAL ---

    const handleContextMenu = (e: React.MouseEvent, target: 'cell' | 'row' | 'col', cellId?: string) => {
        e.preventDefault();
        setContextMenuTarget(target);
        setContextMenuPosition({ x: e.clientX, y: e.clientY });
        setContextMenuVisible(true);

        if (target === 'cell' && cellId) {
            // Asegurarse de que la celda clickeada sea la activa para las acciones
            handleCellClick(cellId);
        }
    };

    const renderContextMenu = () => {
        if (!contextMenuVisible) return null;

        const options = [];
        if (contextMenuTarget === 'cell') {
            options.push('Cortar', 'Copiar', 'Pegar', 'Insertar comentario', 'Borrar contenido');
        } else if (contextMenuTarget === 'row') {
            options.push('Insertar fila', 'Eliminar fila', 'Ocultar fila', 'Alto de fila...');
        } else if (contextMenuTarget === 'col') {
            options.push('Insertar columna', 'Eliminar columna', 'Ocultar columna', 'Ancho de columna...');
        }

        const handleMenuAction = (action: string) => {
            // Lógica simulada para las acciones del menú contextual
            if (action === 'Borrar contenido' && activeCell) {
                setSheetData(prevData => {
                    const newData = { ...prevData };
                    const cell = newData[activeCell];
                    if (cell) {
                         // Restablecer el valor, fórmula y valor calculado
                        newData[activeCell] = { ...cell, value: '', formula: undefined, calculatedValue: '' };
                    }
                    // Recalcular toda la hoja para propagar el borrado
                    return recalculateSheet(newData);
                });
                setCellInput('');
                setDirty(true);
            } else {
                console.log(`[ALERTA] Acción simulada: ${action} en ${contextMenuTarget}`);
            }
            setContextMenuVisible(false);
        };

        return (
            <div
                ref={menuRef}
                className="absolute z-50 bg-white dark:bg-gray-700 shadow-lg rounded-md text-sm border border-gray-300 dark:border-gray-600"
                style={{ top: contextMenuPosition.y, left: contextMenuPosition.x }}
            >
                {options.map((option) => (
                    <div
                        key={option}
                        className="px-4 py-2 hover:bg-indigo-50 dark:hover:bg-indigo-600 text-gray-800 dark:text-gray-100 dark:hover:text-white cursor-pointer"
                        onClick={() => handleMenuAction(option)}
                    >
                        {option}
                    </div>
                ))}
            </div>
        );
    };


    // --- RENDERING AUXILIARES ---

    /**
     * Componente para dibujar la vista Backstage (Archivo)
     */
    const renderBackstage = () => {
        // Renderiza vistas modales anidadas
        if (saveAsVisible) {
            return (
                <SaveAsView
                    docState={docState}
                    handleSaveAs={handleSaveAs}
                    onClose={() => { setSaveAsVisible(false); setBackstageVisible(true); }}
                />
            );
        }

        if (printVisible) {
             return (
                 <PrintView
                    printSettings={printSettings}
                    setPrintSettings={setPrintSettings}
                    onClose={() => { setPrintVisible(false); setBackstageVisible(true); }}
                 />
             );
        }

        // Menú principal del Backstage (Archivo)
        return (
            <div className="absolute inset-0 bg-gray-900 dark:bg-[#1e1e1e] flex z-50 text-white">
                {/* Menú de Opciones */}
                <div className="w-64 bg-gray-800 dark:bg-gray-700 p-4 shadow-xl">
                    {['Nuevo', 'Abrir', 'Guardar', 'Guardar como', 'Imprimir', 'Compartir', 'Opciones'].map(option => (
                        <button
                            key={option}
                            className="w-full text-left p-3 my-1 rounded-md hover:bg-indigo-600 font-medium"
                            onClick={() => {
                                if (option === 'Abrir') return handleOpen();
                                if (option === 'Guardar') return handleSave();
                                if (option === 'Guardar como') return setSaveAsVisible(true);
                                if (option === 'Imprimir') return setPrintVisible(true);
                                if (option === 'Nuevo') {
                                    // Usar console.log en lugar de window.confirm()
                                    console.log('[ALERTA] Simulando confirmación de nuevo libro...');
                                    setSheetData({});
                                    setDocState({ fileName: 'Libro Nuevo', isDirty: false, data: {}, activeSheet: 'Hoja1' });
                                    setActiveCell(null);
                                    return setBackstageVisible(false);
                                }
                                console.log(`[ALERTA] Acción simulada: ${option}`);
                            }}
                        >
                            {option}
                        </button>
                    ))}
                    <button
                        className="w-full text-left p-3 my-1 mt-4 rounded-md bg-indigo-600 hover:bg-indigo-700 font-bold"
                        onClick={() => setBackstageVisible(false)}
                    >
                        &larr; Volver a la hoja de cálculo
                    </button>
                </div>

                {/* Contenido de la Opción */}
                <div className="flex-1 p-8 overflow-auto">
                    <h2 className="text-2xl font-light mb-6">Información</h2>
                    <div className="bg-gray-800 dark:bg-gray-700 p-6 rounded-lg max-w-lg">
                        <p className="text-xl mb-4">Estado del archivo: **{docState.fileName}** {docState.isDirty ? '*(Modificado)*' : '*(Guardado)*'}</p>
                        <p className="text-md text-gray-400">Haga clic en 'Guardar' para guardar el estado actual de la hoja de cálculo en el navegador (LocalStorage).</p>
                        <div className="mt-4">
                            <button className="px-4 py-2 bg-green-600 hover:bg-green-700 rounded-md font-semibold" onClick={handleSave}>Guardar Ahora</button>
                        </div>
                    </div>
                </div>
            </div>
        );
    };

    /**
     * Componente para dibujar la barra de herramientas (Ribbon)
     */
    const renderToolbar = () => {
        const cellData = activeCell ? sheetData[activeCell] : null;
        const currentStyles = cellData?.styles || {};

        const renderButton = (label: string, iconName: string, action: () => void, isActive: boolean = false, className: string = '') => (
            <button
                title={label}
                className={`p-2 rounded-md transition duration-150 ${isActive ? 'bg-indigo-700 text-white' : 'hover:bg-gray-600 text-gray-200'} ${className}`}
                onClick={action}
                disabled={!activeCell && iconName !== 'paste' && iconName !== 'cut' && iconName !== 'copy'}
            >
                <i className={`fas fa-${iconName}`}></i>
            </button>
        );

        // --- PESTAÑAS Y CONTENIDO ---

        const tabContents: Record<string, JSX.Element> = {
            'Inicio': (
                <div className="flex space-x-4">
                    {/* Grupo Portapapeles */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 items-center min-w-[100px] py-1">
                        {renderButton('Pegar', 'paste', () => console.log('[ALERTA] Simulando Pegar...'), false, 'h-8 w-1/2')}
                        <div className="flex mt-1 space-x-1">
                            {renderButton('Cortar', 'cut', () => console.log('[ALERTA] Simulando Cortar...'))}
                            {renderButton('Copiar', 'copy', () => console.log('[ALERTA] Simulando Copiar...'))}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Portapapeles</span>
                    </div>

                    {/* Grupo Fuente */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 items-center min-w-[280px] py-1">
                        <div className="flex space-x-2 mb-1 items-center">
                             {/* Selector de Fuente */}
                            <select
                                className="p-1 bg-gray-600 border border-gray-700 rounded text-gray-100 text-sm w-36 cursor-pointer"
                                value={currentStyles.fontFamily || 'Inter'}
                                onChange={(e) => applyStyleToActiveCell('fontFamily', e.target.value)}
                                disabled={!activeCell}
                            >
                                {['Inter', 'Arial', 'Times New Roman', 'Courier New'].map(font => <option key={font} value={font}>{font}</option>)}
                            </select>
                             {/* Selector de Tamaño */}
                            <select
                                className="p-1 bg-gray-600 border border-gray-700 rounded text-gray-100 text-sm w-16 cursor-pointer"
                                value={currentStyles.fontSize || 14}
                                onChange={(e) => applyStyleToActiveCell('fontSize', parseInt(e.target.value))}
                                disabled={!activeCell}
                            >
                                {[8, 10, 12, 14, 16, 18, 20, 24, 36].map(size => <option key={size} value={size}>{size}</option>)}
                            </select>
                        </div>
                        <div className="flex items-center space-x-1">
                             {/* Negrita, Cursiva, Subrayado */}
                            {renderButton('Negrita', 'bold', () => applyStyleToActiveCell('fontWeight', 'bold'), currentStyles.fontWeight === 'bold', 'font-bold')}
                            {renderButton('Cursiva', 'italic', () => applyStyleToActiveCell('fontStyle', 'italic'), currentStyles.fontStyle === 'italic', 'italic')}
                            {renderButton('Subrayado', 'underline', () => applyStyleToActiveCell('textDecoration', 'underline'), currentStyles.textDecoration === 'underline', 'underline')}
                            {/* Tachado */}
                             {renderButton('Tachado', 'strikethrough', () => applyStyleToActiveCell('textDecoration', 'line-through'), currentStyles.textDecoration === 'line-through')}
                            {/* Bordes Dropdown */}
                            <DropdownControl
                                label="Bordes"
                                primaryIconName="border-all"
                                items={[
                                    { label: 'Sin borde', action: () => applyStyleToActiveCell('borderStyle', 'none') },
                                    { label: 'Todos los bordes', action: () => applyStyleToActiveCell('borderStyle', 'all') },
                                    { label: 'Borde inferior', action: () => applyStyleToActiveCell('borderStyle', 'bottom') },
                                    { label: 'Borde exterior grueso', action: () => applyStyleToActiveCell('borderStyle', 'thickOutside') },
                                ]}
                            />
                            {/* Color de Relleno */}
                            <ColorPickerControl
                                label="Color de Relleno"
                                styleKey='backgroundColor'
                                defaultColor='#1e1e1e'
                                activeCell={activeCell}
                                currentStyles={currentStyles}
                                applyStyleToActiveCell={applyStyleToActiveCell}
                            />
                            {/* Color de Fuente */}
                            <ColorPickerControl
                                label="Color de Fuente"
                                styleKey='color'
                                defaultColor='#ffffff'
                                activeCell={activeCell}
                                currentStyles={currentStyles}
                                applyStyleToActiveCell={applyStyleToActiveCell}
                            />
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Fuente</span>
                    </div>

                    {/* Grupo Alineación */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 items-center min-w-[150px] py-1">
                        <div className="flex space-x-1">
                             {/* Botones de Alineación Horizontal */}
                            {renderButton('Alinear a la izquierda', 'align-left', () => applyStyleToActiveCell('textAlign', 'left'), currentStyles.textAlign === 'left')}
                            {renderButton('Alinear al centro', 'align-center', () => applyStyleToActiveCell('textAlign', 'center'), currentStyles.textAlign === 'center')}
                            {renderButton('Alinear a la derecha', 'align-right', () => applyStyleToActiveCell('textAlign', 'right'), currentStyles.textAlign === 'right')}
                        </div>
                         <div className="flex space-x-1 mt-1">
                            {renderButton('Ajustar texto', 'text-wrap', () => console.log('[ALERTA] Simulando Ajustar texto...'))}
                            {renderButton('Combinar y centrar', 'compress', () => console.log('[ALERTA] Simulando Combinar y centrar...'))}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Alineación</span>
                    </div>

                    {/* Grupo Número */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 items-center min-w-[180px] py-1">
                         <div className="flex space-x-1">
                            <DropdownControl
                                label="Formato de Número"
                                primaryIconName="list-ul"
                                items={[
                                    { label: 'General', action: () => applyStyleToActiveCell('numberFormat', 'General') },
                                    { label: 'Moneda', action: () => applyStyleToActiveCell('numberFormat', 'Currency') },
                                    { label: 'Porcentaje', action: () => applyStyleToActiveCell('numberFormat', 'Percentage') },
                                    { label: 'Número (separador de miles)', action: () => applyStyleToActiveCell('numberFormat', 'Thousands') },
                                ]}
                            />
                            {renderButton('Formato Moneda', 'dollar-sign', () => applyStyleToActiveCell('numberFormat', 'Currency'), currentStyles.numberFormat === 'Currency')}
                            {renderButton('Estilo Porcentual', 'percent', () => applyStyleToActiveCell('numberFormat', 'Percentage'), currentStyles.numberFormat === 'Percentage')}
                            {renderButton('Estilo Millares', 'comma', () => applyStyleToActiveCell('numberFormat', 'Thousands'), currentStyles.numberFormat === 'Thousands')}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Número</span>
                    </div>

                    {/* Grupo Edición */}
                    <div className="flex flex-col items-center min-w-[100px] py-1">
                        {renderButton('Autosuma', 'calculator', () => console.log('[ALERTA] Simulando Autosuma...'), false, 'h-8')}
                        <div className="flex mt-1 space-x-1">
                            {renderButton('Ordenar y filtrar', 'sort', () => console.log('[ALERTA] Simulando Ordenar...'))}
                            {renderButton('Buscar', 'search', () => console.log('[ALERTA] Simulando Buscar...'))}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Edición</span>
                    </div>
                </div>
            ),
            'Insertar': <div className="p-2 text-gray-400">Opciones de Insertar... (Gráficos, Tablas, Imágenes)</div>,
            'Fórmulas': <div className="p-2 text-gray-400">Opciones de Fórmulas... (Autosuma, Lógicas, Financieras)</div>,
            'Datos': <div className="p-2 text-gray-400">Opciones de Datos... (Obtener Datos, Filtros, Análisis de hipótesis)</div>,
        };

        return (
            <div className="w-full bg-gray-800 dark:bg-[#252526] shadow-lg sticky top-0 z-10">
                {/* Pestañas de Navegación */}
                <div className="flex border-b border-gray-700 dark:border-gray-600">
                    {['Archivo', 'Inicio', 'Insertar', 'Fórmulas', 'Datos'].map(tab => (
                        <button
                            key={tab}
                            className={`px-4 py-2 text-sm font-medium ${
                                tab === 'Archivo'
                                    ? 'text-white bg-indigo-700 hover:bg-indigo-600'
                                    : activeTab === tab
                                    ? 'bg-gray-700 dark:bg-gray-600 text-white border-b-2 border-indigo-500'
                                    : 'text-gray-300 hover:bg-gray-700 dark:hover:bg-gray-600'
                            }`}
                            onClick={() => (tab === 'Archivo' ? setBackstageVisible(true) : setActiveTab(tab))}
                        >
                            {tab}
                        </button>
                    ))}
                </div>

                {/* Contenido de la Pestaña */}
                <div className="flex p-2 items-start overflow-x-auto">
                    {tabContents[activeTab]}
                </div>
            </div>
        );
    };

    // --- CUERPO PRINCIPAL DEL COMPONENTE ---

    if (backstageVisible) {
        return renderBackstage();
    }

    return (
        <div className="flex flex-col h-screen w-full bg-[#1e1e1e] text-white font-sans overflow-hidden">
            {/* 1. RIBBON (Barra de Herramientas) */}
            {renderToolbar()}

            {/* 2. BARRA DE FÓRMULA */}
            <div className="flex items-center p-2 bg-gray-700 dark:bg-[#2e2e2e] border-b border-gray-600 shadow-md">
                <div className="w-16 h-6 flex-shrink-0 bg-gray-600 rounded-sm text-center font-bold text-sm leading-6 mr-2 cursor-pointer border border-gray-500"
                     onClick={() => activeCell && handleCellClick(activeCell)}>
                    {activeCell || 'A1'}
                </div>
                <div className="flex-1 bg-gray-800 border border-gray-600 rounded-sm p-1">
                    <input
                        type="text"
                        ref={inputRef}
                        className="w-full bg-transparent outline-none text-white text-sm"
                        placeholder="Fórmula o valor"
                        value={cellInput}
                        onChange={(e) => setCellInput(e.target.value)}
                        onBlur={updateCellValue}
                        onKeyDown={handleKeyDown}
                        // La celda activa se encarga de llamar a `inputRef.current.focus()`
                        readOnly={!activeCell}
                    />
                </div>
            </div>

            {/* 3. CUADRÍCULA DE LA HOJA DE CÁLCULO */}
            <div className="flex-1 overflow-auto relative">
                <div className="min-w-max min-h-max absolute" ref={gridRef}>
                    {/* Encabezado de la esquina superior izquierda */}
                    <div className="sticky top-0 left-0 z-30 w-8 h-6 bg-gray-800 dark:bg-[#252526] border-b border-r border-gray-600"></div>

                    {/* Encabezados de Columna */}
                    <div className="sticky top-0 left-8 z-20 flex bg-gray-800 dark:bg-[#252526] shadow-md border-b border-gray-600">
                        {Array.from({ length: DEFAULT_COLS }).map((_, col) => (
                            <div
                                key={col}
                                className={`w-32 h-6 text-center font-bold text-xs leading-6 border-r border-gray-600 cursor-pointer transition-colors duration-100 flex-shrink-0
                                    ${selectedCols.has(col) ? 'bg-indigo-700 text-white' : 'hover:bg-gray-700 text-gray-400'}`}
                                onClick={() => handleColHeaderClick(col)}
                                onContextMenu={(e) => handleContextMenu(e, 'col')}
                            >
                                {coordsToCell(0, col).replace(/[0-9]/g, '')}
                            </div>
                        ))}
                    </div>

                    {/* Filas */}
                    <div className="flex">
                        {/* Encabezados de Fila (Izquierda) */}
                        <div className="sticky left-0 top-6 z-20 flex flex-col bg-gray-800 dark:bg-[#252526] shadow-md border-r border-gray-600">
                            {Array.from({ length: DEFAULT_ROWS }).map((_, row) => (
                                <div
                                    key={row}
                                    className={`w-8 h-6 text-center font-bold text-xs leading-6 border-b border-gray-600 cursor-pointer transition-colors duration-100
                                        ${selectedRows.has(row) ? 'bg-indigo-700 text-white' : 'hover:bg-gray-700 text-gray-400'}`}
                                    onClick={() => handleRowHeaderClick(row)}
                                    onContextMenu={(e) => handleContextMenu(e, 'row')}
                                >
                                    {row + 1}
                                </div>
                            ))}
                        </div>

                        {/* Celdas de Datos */}
                        <div className="flex flex-col">
                            {Array.from({ length: DEFAULT_COLS }).map((_, col) => (
                                <div key={row} className="flex">
                                    {Array.from({ length: DEFAULT_COLS }).map((_, col) => (
                                        <Cell
                                            key={col}
                                            row={row}
                                            col={col}
                                            // Props para Cell
                                            sheetData={sheetData}
                                            activeCell={activeCell}
                                            selectedRows={selectedRows}
                                            selectedCols={selectedCols}
                                            cellInput={cellInput}
                                            updateCellValue={updateCellValue}
                                            handleCellClick={handleCellClick}
                                            handleContextMenu={handleContextMenu}
                                            setCellInput={setCellInput}
                                            handleKeyDown={handleKeyDown}
                                            inputRef={inputRef}
                                        />
                                    ))}
                                </div>
                            ))}
                        </div>
                    </div>
                </div>
            </div>

            {/* 4. BARRA DE ESTADO */}
            <div className="h-6 flex items-center justify-between px-4 bg-gray-800 dark:bg-[#2e2e2e] border-t border-gray-600 text-xs text-gray-400">
                <span>{docState.activeSheet} | {docState.fileName} {docState.isDirty ? '(*Modificado)' : ''}</span>
                <span>Listo</span>
            </div>

            {/* Menú Contextual */}
            {renderContextMenu()}
        </div>
    );
};

export default PacurHoja;
