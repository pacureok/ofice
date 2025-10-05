import { useState, useMemo } from 'react';
import React from 'react'; 

// --- Configuración de la Cuadrícula para simular un tamaño "Infinito" ---
const ROWS = 200; // Filas hacia abajo
const COLS = 70;  // Columnas hacia la derecha (A hasta BR)

// Tipo para almacenar los datos de la hoja de cálculo
interface SheetData {
  [key: string]: string; // Clave: "A1", "B2", Valor: Contenido o fórmula
}

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

    // Utilizamos un mensaje box personalizado en lugar de alert
    const messageBox = document.createElement('div');
    messageBox.style.cssText = `
        position: fixed; top: 20px; right: 20px; background-color: #4CAF50; color: white;
        padding: 15px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        z-index: 1000; transition: opacity 0.5s;
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
    // Implementar lógica aquí (ej. cambiar estilo, llamar API, etc.)
    const messageBox = document.createElement('div');
    messageBox.style.cssText = `
        position: fixed; top: 20px; right: 20px; background-color: #0078d4; color: white;
        padding: 10px; border-radius: 5px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        z-index: 1000; transition: opacity 0.5s;
    `;
    messageBox.textContent = `Acción: ${action} - (Lógica no implementada aún)`;
    document.body.appendChild(messageBox);
    setTimeout(() => {
        messageBox.style.opacity = '0';
        setTimeout(() => document.body.removeChild(messageBox), 1000);
    }, 2000);
  };

  return (
    <div className="pacur-hoja-container">
      
      {/* 1. Barra de Herramientas (Ribbon COMPLETO) */}
      <div className="toolbar ribbon">
        {/* Pestañas (Simulación) */}
        <div className="ribbon-tabs">
            <span className="ribbon-tab active">Inicio</span>
            <span className="ribbon-tab">Insertar</span>
            <span className="ribbon-tab">Dibujar</span>
            <span className="ribbon-tab">Disposición de página</span>
            <span className="ribbon-tab">Fórmulas</span>
            <span className="ribbon-tab">Datos</span>
            <span className="ribbon-tab">Vista</span>
        </div>

        <div className="ribbon-content">
            {/* GRUPO: PORTAPAPELES */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Pegar")} title="Pegar" className="large-button">
                    <span style={{fontSize: '1.2rem'}}>📋</span><br/>Pegar
                </button>
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Cortar")} title="Cortar">✂️</button>
                    <button onClick={() => handleToolbarAction("Copiar")} title="Copiar">📝</button>
                </div>
                <div className="group-label">Portapapeles</div>
            </div>

            {/* GRUPO: FUENTE */}
            <div className="toolbar-group">
                <div className="horizontal-group">
                    <select defaultValue="Aptos Narrow" title="Fuente">
                        <option>Aptos Narrow</option><option>Arial</option><option>Calibri</option>
                    </select>
                    <select defaultValue="11" title="Tamaño">
                        <option>11</option><option>12</option>
                    </select>
                </div>
                <div className="horizontal-group">
                    <button onClick={() => handleToolbarAction("Negrita")} title="Negrita"><b>N</b></button>
                    <button onClick={() => handleToolbarAction("Cursiva")} title="Cursiva"><i>K</i></button>
                    <button onClick={() => handleToolbarAction("Subrayado")} title="Subrayado"><u>S</u></button>
                    <button onClick={() => handleToolbarAction("Bordes")} title="Bordes de celda">🖼️</button>
                    <button onClick={() => handleToolbarAction("Relleno")} title="Color de Relleno">🎨</button>
                    <button onClick={() => handleToolbarAction("Color Fuente")} title="Color de Fuente">🅰️</button>
                </div>
                <div className="group-label">Fuente</div>
            </div>

            {/* GRUPO: ALINEACIÓN */}
            <div className="toolbar-group">
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Alinear Superior")} title="Alinear Arriba">⬆️</button>
                    <button onClick={() => handleToolbarAction("Alinear Medio")} title="Alinear Medio">↔</button>
                    <button onClick={() => handleToolbarAction("Alinear Inferior")} title="Alinear Abajo">⬇️</button>
                </div>
                <div className="vertical-group">
                    <button onClick={() => handleToolbarAction("Izquierda")} title="Alinear Izquierda">⏴</button>
                    <button onClick={() => handleToolbarAction("Centrar")} title="Centrar">☰</button>
                    <button onClick={() => handleToolbarAction("Derecha")} title="Alinear Derecha">⏵</button>
                </div>
                <div className="group-label">Alineación</div>
            </div>

            {/* GRUPO: NÚMERO */}
            <div className="toolbar-group">
                <select defaultValue="General" title="Formato de Número" style={{width: '90px'}}>
                    <option>General</option><option>Número</option><option>Moneda</option><option>Porcentaje</option>
                </select>
                <div className="horizontal-group">
                    <button onClick={() => handleToolbarAction("Moneda")} title="Formato Moneda">$</button>
                    <button onClick={() => handleToolbarAction("Porcentaje")} title="Estilo Porcentual">%</button>
                    <button onClick={() => handleToolbarAction("Comas")} title="Estilo Millares">, </button>
                </div>
                <div className="group-label">Número</div>
            </div>

            {/* GRUPO: ESTILOS */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Formato Condicional")} title="Formato Condicional">📊</button>
                <button onClick={() => handleToolbarAction("Dar Formato Como Tabla")} title="Dar Formato como Tabla">📋</button>
                <button onClick={() => handleToolbarAction("Estilos de Celda")} title="Estilos de Celda">🎨</button>
                <div className="group-label">Estilos</div>
            </div>

            {/* GRUPO: CELDAS */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Insertar")} title="Insertar Celdas/Filas">➕</button>
                <button onClick={() => handleToolbarAction("Eliminar")} title="Eliminar Celdas/Filas">➖</button>
                <button onClick={() => handleToolbarAction("Formato")} title="Formato de Fila/Columna">⚙️</button>
                <div className="group-label">Celdas</div>
            </div>

            {/* GRUPO: EDICIÓN */}
            <div className="toolbar-group">
                <button onClick={() => handleToolbarAction("Autosuma")} title="Autosuma">Σ</button>
                <button onClick={() => handleToolbarAction("Ordenar y Filtrar")} title="Ordenar y Filtrar">⬇️⬆️</button>
                <button onClick={() => handleToolbarAction("Buscar y Seleccionar")} title="Buscar y Seleccionar">🔍</button>
                <div className="group-label">Edición</div>
            </div>
            
            {/* GRUPO: GUARDAR (Fuera de la cinta de opciones para mayor visibilidad) */}
            <div className="toolbar-group" style={{marginLeft: 'auto', marginRight: '20px'}}>
                <button onClick={saveSheet} title="Guardar como .aph" style={{backgroundColor: '#0078d4'}}>💾 Guardar</button>
            </div>
        </div>
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
      <div className="spreadsheet-grid">
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
  );
};

export default PacurHoja;
