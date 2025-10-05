import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';

// --- TIPOS Y CONSTANTES ---

// Define las propiedades de estilo que puede tener una celda
type CellStyle = {
    fontFamily?: string;
    fontSize?: number;
    fontWeight?: 'normal' | 'bold';
    fontStyle?: 'normal' | 'italic';
    textDecoration?: 'none' | 'underline';
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
const DEFAULT_COLS = 10;
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
 * Parsea un rango de celda (ej: A1:B5) y devuelve una lista de claves de celda.
 * @param rangeString Rango de celda (ej: A1:B5)
 * @returns Array de claves de celda (ej: ['A1', 'A2', ...])
 */
const parseRange = (rangeString: string): string[] => {
    if (!rangeString || !rangeString.includes(':')) return [];

    const [startCell, endCell] = rangeString.split(':');
    const startMatch = startCell.match(/([A-Z]+)(\d+)/);
    const endMatch = endCell.match(/([A-Z]+)(\d+)/);

    if (!startMatch || !endMatch) return [];

    // Conversión de letras de columna a índice (simplificado para A-J)
    const colToIndex = (col: string): number => {
        let index = 0;
        for (let i = 0; i < col.length; i++) {
            index = index * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return index - 1;
    };

    const startCol = colToIndex(startMatch[1]);
    const startRow = parseInt(startMatch[2], 10) - 1;
    const endCol = colToIndex(endMatch[1]);
    const endRow = parseInt(endMatch[2], 10) - 1;

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

// --- LÓGICA DE CÁLCULO DE FÓRMULAS ---

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
                // Recursivamente calcular el valor de la celda referenciada
                const val = calculateValue(cell.formula || cell.value, sheetData);
                return typeof val === 'number' ? val : parseFloat(String(val).replace(/[^0-9.-]+/g, ""));
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
            const val = calculateValue(cell.formula || cell.value, sheetData);
            return String(val).replace(/[^0-9.-]+/g, ""); // Usar solo el número
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

    // Definición de estilos estáticos (Corrige TS18004)
    const styles: React.CSSProperties = {}; 

    // --- EFECTOS DE CONTROL ---

    // Sincronizar input con celda activa
    useEffect(() => {
        if (activeCell) {
            const data = sheetData[activeCell];
            setCellInput(data?.formula || data?.value || '');
            inputRef.current?.focus();
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

    // --- MANEJADORES DE ESTADO DE ARCHIVO (Guardar/Abrir) ---

    // Función para marcar el documento como modificado
    const setDirty = useCallback((dirty: boolean) => {
        setDocState(prev => ({ ...prev, isDirty: dirty }));
    }, []);

    // Simulación de Guardar (usa localStorage)
    const handleSave = () => {
        if (!docState.isDirty) {
            console.log('El archivo ya está guardado.');
            return;
        }
        try {
            const dataToSave = JSON.stringify(sheetData);
            localStorage.setItem(`pacur_sheet_${docState.fileName}`, dataToSave);
            setDirty(false);
            console.log(`Archivo '${docState.fileName}' guardado con éxito.`);
            // Usar console.log en lugar de alert()
            console.log(`[ALERTA] Archivo '${docState.fileName}' guardado con éxito.`);
        } catch (error) {
            console.error('Error al guardar:', error);
            // Usar console.log en lugar de alert()
            console.log('[ALERTA] Error al intentar guardar el archivo.');
        }
    };

    // Simulación de Abrir
    const handleOpen = () => {
        const savedData = localStorage.getItem(`pacur_sheet_${docState.fileName}`);
        if (savedData) {
            try {
                const loadedSheetData: SheetData = JSON.parse(savedData);
                setSheetData(loadedSheetData);
                setDirty(false);
                setBackstageVisible(false);
                console.log(`Archivo '${docState.fileName}' abierto.`);
            } catch (error) {
                console.error('Error al cargar datos:', error);
                // Usar console.log en lugar de alert()
                console.log('[ALERTA] Error al cargar el archivo guardado.');
            }
        } else {
            setBackstageVisible(false);
            console.log('No hay datos guardados para cargar.');
        }
    };

    // Simulación de Guardar Como (muestra la interfaz)
    const handleSaveAs = (newFileName: string, format: string) => {
        // En una app real, aquí se llamaría a una API de descarga/guardado.
        const fileExtension = format === 'excel' ? '.xlsx' : '.pacur';
        console.log(`Simulando Guardar como: ${newFileName}${fileExtension}`);
        // Usar console.log en lugar de alert()
        console.log(`[ALERTA] Simulando Guardar como: ${newFileName}${fileExtension}. Ubicación: [Elegir]`);
        setDocState(prev => ({ ...prev, fileName: newFileName, isDirty: false }));
        setSaveAsVisible(false);
        setBackstageVisible(false);
    };

    // --- MANEJADORES DE LA HOJA DE CÁLCULO ---

    /**
     * Aplica el valor del input (fórmula o texto) a la celda activa.
     */
    const updateCellValue = () => {
        if (!activeCell) return;

        const isFormula = cellInput.startsWith('=');
        // Recalcular el valor para manejar dependencias
        const calculated = isFormula
            ? calculateValue(cellInput, sheetData)
            : cellInput;

        setSheetData(prevData => {
            const currentStyles = prevData[activeCell]?.styles || {};
            // Forzar el recálculo de toda la hoja si es una fórmula (para propagar cambios)
            const updatedData = {
                ...prevData,
                [activeCell]: {
                    value: cellInput,
                    formula: isFormula ? cellInput : undefined,
                    calculatedValue: calculated,
                    styles: currentStyles,
                },
            };

            // Recalcular todas las celdas que contienen fórmulas
            const recalculatedData: SheetData = {};
            Object.keys(updatedData).forEach(cellId => {
                const cell = updatedData[cellId];
                if (cell.formula) {
                    recalculatedData[cellId] = {
                        ...cell,
                        calculatedValue: calculateValue(cell.formula, updatedData),
                    };
                } else {
                    recalculatedData[cellId] = cell;
                }
            });

            return recalculatedData;
        });
        setDirty(true);
    };

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

            if (styleKey === 'fontWeight' || styleKey === 'fontStyle' || styleKey === 'textDecoration') {
                // Toggle para estilos de fuente (N, K, S)
                newStyles[styleKey] = (currentStyles[styleKey] === value) ? undefined : value;
            } else if (styleKey === 'textAlign') {
                 // Toggle para alineación (solo un valor activo a la vez)
                newStyles[styleKey] = (currentStyles[styleKey] === value) ? undefined : value;
            } else if (styleKey === 'numberFormat') {
                // Aplicar formato de número
                newStyles[styleKey] = value;
            } else {
                // Aplicar color de fuente, color de relleno, tamaño, etc.
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

    const handleCellClick = (cellId: string) => {
        if (contextMenuVisible) setContextMenuVisible(false);
        setActiveCell(cellId);
        setSelectedRows(new Set());
        setSelectedCols(new Set());
        // No llamar a updateCellValue aquí, solo al perder el foco del input o al presionar Enter.
    };

    const handleRowHeaderClick = (row: number) => {
        setSelectedRows(new Set([row]));
        setSelectedCols(new Set());
        setActiveCell(null);
        setContextMenuVisible(false);
    };

    const handleColHeaderClick = (col: number) => {
        setSelectedCols(new Set([col]));
        setSelectedRows(new Set());
        setActiveCell(null);
        setContextMenuVisible(false);
    };

    // --- MENÚ CONTEXTUAL ---

    const handleContextMenu = (e: React.MouseEvent, target: 'cell' | 'row' | 'col') => {
        e.preventDefault();
        setContextMenuTarget(target);
        setContextMenuPosition({ x: e.clientX, y: e.clientY });
        setContextMenuVisible(true);
    };

    const renderContextMenu = () => {
        if (!contextMenuVisible) return null;

        const options = [];
        if (contextMenuTarget === 'cell') {
            options.push('Cortar', 'Copiar', 'Pegar', 'Insertar comentario');
        } else if (contextMenuTarget === 'row') {
            options.push('Insertar', 'Eliminar', 'Ocultar', 'Alto de fila...');
        } else if (contextMenuTarget === 'col') {
            options.push('Insertar', 'Eliminar', 'Ocultar', 'Ancho de columna...');
        }

        const handleMenuAction = (action: string) => {
            // Usar console.log en lugar de alert()
            console.log(`[ALERTA] Acción simulada: ${action} en ${contextMenuTarget}`);
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
        if (saveAsVisible) {
            // Interfaz de Guardar Como
            const [newFileName, setNewFileName] = useState(docState.fileName);
            const [format, setFormat] = useState('excel');
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
                                onClick={() => handleSaveAs(newFileName, format)}
                            >
                                Guardar
                            </button>
                            <button
                                className="px-4 py-2 bg-gray-600 hover:bg-gray-500 rounded-md"
                                onClick={() => setSaveAsVisible(false)}
                            >
                                Cancelar
                            </button>
                        </div>
                    </div>
                </div>
            );
        }

        if (printVisible) {
             // Interfaz de Imprimir
             const isValidRange = (range: string) => {
                 return range === 'Hoja actual' || /^[A-Z]+\d+:[A-Z]+\d+$/i.test(range);
             }

             return (
                 <div className="absolute inset-0 bg-gray-900 dark:bg-[#1e1e1e] p-10 z-50 overflow-auto text-white flex">
                    {/* Panel de Configuración de Impresión (Izquierda) */}
                    <div className="w-1/4 bg-gray-800 dark:bg-gray-700 p-6 space-y-6 flex flex-col justify-between">
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
                                    Ejemplo de rango personalizado: **1A:5G** imprime el área de A1 a G5.
                                </div>
                            </div>
                        </div>

                        <div className="pt-6 border-t border-gray-600">
                             <button
                                className={`w-full px-4 py-2 bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold ${!isValidRange(printSettings.range) ? 'opacity-50 cursor-not-allowed' : ''}`}
                                onClick={() => {
                                    if(isValidRange(printSettings.range)) {
                                        // Usar console.log en lugar de alert()
                                        console.log(`[ALERTA] Simulando imprimir rango: ${printSettings.range} en orientación ${printSettings.orientation}.`);
                                    } else {
                                        // Usar console.log en lugar de alert()
                                        console.log('[ALERTA] Rango de impresión inválido.');
                                    }
                                }}
                                disabled={!isValidRange(printSettings.range)}
                            >
                                Imprimir
                            </button>
                            <button
                                className="w-full mt-2 px-4 py-2 bg-gray-600 hover:bg-gray-500 rounded-md"
                                onClick={() => setPrintVisible(false)}
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
                                    // Usar console.log en lugar de window.confirm
                                    if (docState.isDirty && !window.confirm('Hay cambios sin guardar. ¿Desea crear un nuevo libro?')) return;
                                    setSheetData({});
                                    setDocState({ fileName: 'Libro Nuevo', isDirty: false, data: {}, activeSheet: 'Hoja1' });
                                    return setBackstageVisible(false);
                                }
                                // Usar console.log en lugar de alert()
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
                        <p className="text-xl mb-4">Estado del archivo: {docState.fileName} {docState.isDirty ? '(Modificado)' : '(Guardado)'}</p>
                        <p className="text-md text-gray-400">Haga clic en 'Guardar' para guardar el estado actual de la hoja de cálculo en el navegador.</p>
                        <div className="mt-4">
                            <button className="px-4 py-2 bg-green-600 hover:bg-green-700 rounded-md font-semibold" onClick={handleSave}>Guardar Ahora</button>
                        </div>
                    </div>
                    <div className="mt-8">
                        <h3 className="text-xl mb-4">Plantillas Recientes</h3>
                        <div className="grid grid-cols-2 gap-4 max-w-2xl">
                            <div className="bg-gray-700 dark:bg-gray-600 p-4 rounded-lg cursor-pointer hover:bg-indigo-600 transition"
                                 onClick={() => setBackstageVisible(false)}>
                                <div className="text-4xl mb-2">📄</div>
                                <p className="font-semibold">Libro en blanco</p>
                                <p className="text-xs text-gray-400">Comience desde cero</p>
                            </div>
                            <div className="bg-gray-700 dark:bg-gray-600 p-4 rounded-lg cursor-pointer hover:bg-indigo-600 transition"
                                 onClick={() => console.log('[ALERTA] Simulando abrir plantilla de Presupuesto...')}>
                                <div className="text-4xl mb-2">💰</div>
                                <p className="font-semibold">Presupuesto Personal</p>
                                <p className="text-xs text-gray-400">Use una plantilla</p>
                            </div>
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

        const renderButton = (label: string, icon: string, action: () => void, isActive: boolean = false, className: string = '') => (
            <button
                title={label}
                className={`p-2 rounded-md transition duration-150 ${isActive ? 'bg-indigo-700 text-white' : 'hover:bg-gray-600 text-gray-200'} ${className}`}
                onClick={action}
            >
                <i className={`fas fa-${icon}`}></i>
            </button>
        );

        const renderDropdown = (label: string, items: { label: string, action: () => void }[], icon: string) => {
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
                         <i className={`fas fa-${icon}`}></i>
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

        const renderColorPicker = (label: string, icon: string, styleKey: 'backgroundColor' | 'color') => {
            const [color, setColor] = useState(currentStyles[styleKey] || '#ffffff');
            const colorRef = useRef<HTMLInputElement>(null);

            const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
                const newColor = e.target.value;
                setColor(newColor);
                applyStyleToActiveCell(styleKey, newColor);
            };

            return (
                <div className="relative inline-block" title={label}>
                    <button
                        className="p-2 rounded-md hover:bg-gray-600 text-gray-200 text-lg"
                        onClick={() => colorRef.current?.click()}
                    >
                         {/* Icono con una barra que muestra el color actual */}
                        <div className="relative">
                            <i className={`fas fa-fill-drip ${styleKey === 'color' ? 'fa-a' : ''}`}></i>
                            <div className="absolute -bottom-1 left-0 right-0 h-[3px] rounded-full" style={{ backgroundColor: currentStyles[styleKey] || 'transparent' }}></div>
                        </div>
                    </button>
                    <input
                        type="color"
                        ref={colorRef}
                        value={color}
                        onChange={handleChange}
                        className="absolute opacity-0 pointer-events-none w-0 h-0"
                    />
                </div>
            );
        };


        // --- PESTAÑAS Y CONTENIDO ---

        const tabContents: Record<string, JSX.Element> = {
            'Inicio': (
                <>
                    {/* Grupo Portapapeles */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 mr-3 items-center">
                        {renderButton('Pegar', 'paste', () => console.log('[ALERTA] Simulando Pegar...'), false, 'h-8')}
                        <div className="flex mt-1">
                            {renderButton('Cortar', 'cut', () => console.log('[ALERTA] Simulando Cortar...'))}
                            {renderButton('Copiar', 'copy', () => console.log('[ALERTA] Simulando Copiar...'))}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Portapapeles</span>
                    </div>

                    {/* Grupo Fuente */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 mr-3 items-center">
                        <div className="flex space-x-1 mb-1">
                             {/* Selector de Fuente */}
                            <select
                                className="p-1 bg-gray-600 border border-gray-700 rounded text-gray-100 text-sm"
                                value={currentStyles.fontFamily || 'Inter'}
                                onChange={(e) => applyStyleToActiveCell('fontFamily', e.target.value)}
                            >
                                {['Inter', 'Arial', 'Times New Roman'].map(font => <option key={font} value={font}>{font}</option>)}
                            </select>
                             {/* Selector de Tamaño */}
                            <select
                                className="p-1 bg-gray-600 border border-gray-700 rounded text-gray-100 text-sm"
                                value={currentStyles.fontSize || 14}
                                onChange={(e) => applyStyleToActiveCell('fontSize', parseInt(e.target.value))}
                            >
                                {[10, 12, 14, 16, 18, 20].map(size => <option key={size} value={size}>{size}</option>)}
                            </select>
                        </div>
                        <div className="flex items-center space-x-1">
                             {/* Negrita, Cursiva, Subrayado */}
                            {renderButton('Negrita', 'bold', () => applyStyleToActiveCell('fontWeight', 'bold'), currentStyles.fontWeight === 'bold', 'font-bold')}
                            {renderButton('Cursiva', 'italic', () => applyStyleToActiveCell('fontStyle', 'italic'), currentStyles.fontStyle === 'italic', 'italic')}
                            {renderButton('Subrayado', 'underline', () => applyStyleToActiveCell('textDecoration', 'underline'), currentStyles.textDecoration === 'underline', 'underline')}
                            {/* Bordes Dropdown */}
                            {renderDropdown('Bordes', [
                                { label: 'Sin borde', action: () => applyStyleToActiveCell('borderStyle', 'none') },
                                { label: 'Todos los bordes', action: () => applyStyleToActiveCell('borderStyle', 'all') },
                                { label: 'Borde inferior', action: () => applyStyleToActiveCell('borderStyle', 'bottom') },
                                { label: 'Borde exterior grueso', action: () => applyStyleToActiveCell('borderStyle', 'thickOutside') },
                            ], 'border-all')}
                            {/* Color de Relleno */}
                            {renderColorPicker('Color de Relleno', 'fill-drip', 'backgroundColor')}
                            {/* Color de Fuente */}
                            {renderColorPicker('Color de Fuente', 'a', 'color')}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Fuente</span>
                    </div>

                    {/* Grupo Alineación */}
                    <div className="flex flex-col border-r border-gray-700 pr-3 mr-3 items-center">
                        <div className="flex space-x-1">
                             {/* Botones de Alineación */}
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
                    <div className="flex flex-col border-r border-gray-700 pr-3 mr-3 items-center">
                         <div className="flex space-x-1">
                            {renderDropdown('Formato de Número', [
                                { label: 'General', action: () => applyStyleToActiveCell('numberFormat', 'General') },
                                { label: 'Moneda', action: () => applyStyleToActiveCell('numberFormat', 'Currency') },
                                { label: 'Porcentaje', action: () => applyStyleToActiveCell('numberFormat', 'Percentage') },
                                { label: 'Número (separador de miles)', action: () => applyStyleToActiveCell('numberFormat', 'Thousands') },
                            ], 'hashtag')}
                            {renderButton('Formato Moneda', 'dollar-sign', () => applyStyleToActiveCell('numberFormat', 'Currency'), currentStyles.numberFormat === 'Currency')}
                            {renderButton('Estilo Porcentual', 'percent', () => applyStyleToActiveCell('numberFormat', 'Percentage'), currentStyles.numberFormat === 'Percentage')}
                            {renderButton('Estilo Millares', 'comma', () => applyStyleToActiveCell('numberFormat', 'Thousands'), currentStyles.numberFormat === 'Thousands')}
                        </div>
                        <span className="text-xs text-gray-400 mt-1">Número</span>
                    </div>
                </>
            ),
            'Insertar': <span className="text-gray-400">Opciones de Insertar...</span>,
            'Fórmulas': <span className="text-gray-400">Opciones de Fórmulas...</span>,
            'Datos': <span className="text-gray-400">Opciones de Datos...</span>,
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

    /**
     * Componente para renderizar una celda individual.
     */
    const Cell: React.FC<{ row: number, col: number }> = useMemo(() => {
        return ({ row, col }) => {
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
                const num = typeof value === 'string' ? parseFloat(value.replace(/[^0-9.-]+/g, "")) : value;
                if (isNaN(num)) return value; // Devolver tal cual si no es numérico

                if (format === 'Currency') {
                    return num.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
                }
                if (format === 'Percentage') {
                    return (num * 100).toFixed(2) + '%';
                }
                if (format === 'Thousands') {
                     return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
                }
                return String(value);
            };

            const finalDisplayedValue = formatValue(displayedValue, data?.styles.numberFormat);

            // Generar estilos CSS
            const cellStyles: React.CSSProperties = {
                ...data?.styles,
                fontFamily: data?.styles.fontFamily || 'Inter, sans-serif',
                fontSize: data?.styles.fontSize ? `${data.styles.fontSize}px` : '14px',
                fontWeight: data?.styles.fontWeight || 'normal',
                fontStyle: data?.styles.fontStyle || 'normal',
                textDecoration: data?.styles.textDecoration || 'none',
                textAlign: data?.styles.textAlign || 'left',
                backgroundColor: data?.styles.backgroundColor || (isSelected ? '#353535' : '#1e1e1e'),
                color: data?.styles.color || '#fff',
                border: isActive ? '2px solid #1a73e8' : '1px solid #333',
                // Manejo de bordes
                ...(data?.styles.borderStyle === 'none' && { border: isActive ? '2px solid #1a73e8' : '1px solid #333' }),
                ...(data?.styles.borderStyle === 'all' && { border: isActive ? '2px solid #1a73e8' : '1px solid #555' }),
                ...(data?.styles.borderStyle === 'bottom' && { borderBottom: '2px solid #fff' }),
                ...(data?.styles.borderStyle === 'top' && { borderTop: '2px solid #fff' }),
                ...(data?.styles.borderStyle === 'left' && { borderLeft: '2px solid #fff' }),
                ...(data?.styles.borderStyle === 'right' && { borderRight: '2px solid #fff' }),
                ...(data?.styles.borderStyle === 'outside' && { boxShadow: `inset 0 0 0 1px #fff` }),
                ...(data?.styles.borderStyle === 'thickOutside' && { boxShadow: `inset 0 0 0 2px #fff` }),
            };

            // Asegurar que las celdas seleccionadas mantengan su estilo de borde principal si no hay un borde activo
            if (isSelected && !isActive && !data?.styles.borderStyle) {
                 cellStyles.backgroundColor = '#353535';
                 cellStyles.border = '1px solid #555';
            }


            return (
                <div
                    className={`h-7 px-1 whitespace-nowrap overflow-hidden flex items-center ${isActive ? 'z-20' : 'z-10'}`}
                    style={cellStyles}
                    onClick={() => handleCellClick(cellId)}
                    onDoubleClick={() => setActiveCell(cellId)}
                    onContextMenu={(e) => handleContextMenu(e, 'cell')}
                >
                    {isActive ? (
                        <input
                            ref={inputRef}
                            type="text"
                            value={cellInput}
                            onChange={(e) => setCellInput(e.target.value)}
                            onBlur={updateCellValue}
                            onKeyDown={(e) => e.key === 'Enter' && updateCellValue()}
                            className="w-full h-full bg-transparent border-none outline-none text-white p-0 m-0"
                            style={{ ...cellStyles, backgroundColor: 'transparent', textAlign: cellStyles.textAlign }}
                        />
                    ) : (
                        finalDisplayedValue
                    )}
                </div>
            );
        };
    }, [activeCell, cellInput, sheetData, selectedRows, selectedCols, updateCellValue]); // Se añade updateCellValue a las dependencias.

    const renderGrid = () => {
        const rows = Array.from({ length: DEFAULT_ROWS }, (_, r) => r);
        const cols = Array.from({ length: DEFAULT_COLS }, (_, c) => c);

        return (
            <div className="overflow-auto flex-1 p-2 bg-[#1e1e1e]">
                <div className="inline-grid gap-0" style={{
                    gridTemplateColumns: `40px repeat(${DEFAULT_COLS}, 100px)`
                }}>
                    {/* Celda superior izquierda vacía */}
                    <div className="h-7 w-10 bg-gray-800 dark:bg-[#333] border-b border-r border-gray-600 dark:border-gray-500 sticky top-0 left-0 z-30"></div>

                    {/* Encabezados de Columna */}
                    {cols.map((col) => (
                        <div
                            key={col}
                            className={`h-7 flex items-center justify-center font-bold text-gray-300 bg-gray-800 dark:bg-[#333] border-b border-r border-gray-600 dark:border-gray-500 cursor-pointer sticky top-0 z-30 ${selectedCols.has(col) ? 'bg-indigo-800 hover:bg-indigo-700' : 'hover:bg-gray-700'}`}
                            onClick={() => handleColHeaderClick(col)}
                            onContextMenu={(e) => handleContextMenu(e, 'col')}
                        >
                            {COL_HEADERS[col % 26]}
                        </div>
                    ))}

                    {/* Filas y Celdas */}
                    {rows.map((row) => (
                        <React.Fragment key={row}>
                            {/* Encabezado de Fila */}
                            <div
                                className={`h-7 flex items-center justify-center font-bold text-gray-300 bg-gray-800 dark:bg-[#333] border-r border-gray-600 dark:border-gray-500 cursor-pointer sticky left-0 z-30 ${selectedRows.has(row) ? 'bg-indigo-800 hover:bg-indigo-700' : 'hover:bg-gray-700'}`}
                                onClick={() => handleRowHeaderClick(row)}
                                onContextMenu={(e) => handleContextMenu(e, 'row')}
                            >
                                {row + 1}
                            </div>
                            {/* Celdas */}
                            {cols.map((col) => (
                                <Cell key={col} row={row} col={col} />
                            ))}
                        </React.Fragment>
                    ))}
                </div>
            </div>
        );
    };

    return (
        <div className="flex flex-col h-screen w-full bg-[#1e1e1e] font-sans">
            {/* CORRECCIÓN: Se eliminan los atributos 'jsx' y 'global' de la etiqueta <style>
                para evitar la advertencia: "Received `true` for a non-boolean attribute `jsx`."
            */}
            <style>{`
                /* Estilos globales para la hoja de cálculo */
                body {
                    margin: 0;
                    padding: 0;
                    background-color: #1e1e1e;
                    color: white;
                }
                .ribbon-icon {
                    width: 20px;
                    height: 20px;
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                }
                .active-style {
                    background-color: #1f3a5f !important; /* Azul oscuro para estilos activos */
                }
            `}</style>
            
            {/* 1. Vista Backstage (Archivo) */}
            {backstageVisible && renderBackstage()}

            {/* 2. Barra de Herramientas (Ribbon) */}
            {!backstageVisible && renderToolbar()}

            {/* 3. Barra de Fórmulas y Nombre de Celda */}
            {!backstageVisible && (
                <div className="flex h-8 bg-gray-800 dark:bg-[#333] items-center px-2 shadow-inner border-y border-gray-700 dark:border-gray-600">
                    <div className="w-20 h-6 bg-gray-900 dark:bg-[#444] rounded flex items-center justify-center text-sm font-semibold text-white mr-2 shadow-inner">
                        {activeCell || 'A1'}
                    </div>
                    <div className="flex-1 h-6 bg-gray-900 dark:bg-[#444] rounded flex items-center px-2 text-white text-sm shadow-inner">
                        {/* Se mantiene el input principal para que la edición no desaparezca al mover el foco */}
                        {sheetData[activeCell || '']?.formula || sheetData[activeCell || '']?.value || ''}
                    </div>
                </div>
            )}

            {/* 4. Cuadrícula Principal */}
            {!backstageVisible && renderGrid()}

            {/* 5. Barra de Estado Inferior */}
            {!backstageVisible && (
                <div className="flex h-6 bg-gray-800 dark:bg-[#333] items-center justify-between text-xs text-gray-400 px-3 border-t border-gray-700 dark:border-gray-600">
                    <div className="flex space-x-4">
                        <span className="cursor-pointer hover:text-white">✅ Listo</span>
                        <span className="cursor-pointer hover:text-white">Hoja {docState.activeSheet}</span>
                        <span className="text-white font-medium ml-4">{docState.fileName} {docState.isDirty ? '(Modificado)' : ''}</span>
                    </div>
                    <div className="flex space-x-3 items-center">
                        {/* Vistas */}
                        <i className="fas fa-grip-vertical cursor-pointer hover:text-white" title="Vista normal"></i>
                        <i className="fas fa-pager cursor-pointer hover:text-white" title="Vista diseño de página"></i>
                        <i className="fas fa-layer-group cursor-pointer hover:text-white" title="Vista previa de salto de página"></i>
                        {/* Zoom */}
                        <span className="ml-3">Zoom: 100%</span>
                    </div>
                </div>
            )}

            {/* Renderizar menú contextual */}
            {renderContextMenu()}
        </div>
    );
};

// Componente Wrapper para demostración
const DefaultApp = () => {
    // No necesitamos items de ejemplo para esta versión, ya que el componente se enfoca en la interfaz de hoja de cálculo.
    return (
      <div className="h-screen w-screen">
          <PacurHoja />
      </div>
    )
}

export default DefaultApp;
