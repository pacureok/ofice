import { useState, useMemo } from 'react';
import React from 'react'; 

// Constantes para la cuadrícula
const ROWS = 50; 
const COLS = 15; 

// Tipo para almacenar los datos de la hoja de cálculo
interface SheetData {
  [key: string]: string; // Clave: "A1", "B2", Valor: Contenido o fórmula
}

// Genera los encabezados de las columnas (A, B, C...)
const getColHeaders = (count: number) => {
  return Array.from({ length: count }, (_, i) => 
    String.fromCharCode('A'.charCodeAt(0) + i)
  );
};

// Expresión regular para encontrar referencias de celda (Ej: A1, B10, C1)
// NOTA: Esta expresión también debe manejar referencias de varias letras si se extiende más allá de Z
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
   * @param key La clave de la celda (ej: "A1").
   * @param path Historial de celdas visitadas para detectar referencias circulares.
   * @returns El valor calculado o un mensaje de error.
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
      
      // Si la celda referenciada es un error o texto, se devuelve 0 o se propaga el error
      const numValue = parseFloat(referencedValue);
      
      if (isNaN(numValue)) {
        // Si el valor referenciado es texto o error, se trata como 0 para las operaciones matemáticas
        return '0'; 
      }
      return numValue.toString();
    });

    // 4. Reemplazar el operador de exponenciación (^) por Math.pow
    const finalFormula = resolvedFormula.replace(/\^/g, '**');

    // 5. Evaluar la fórmula (¡Advertencia: usar 'eval' en producción no es seguro!)
    try {
      const result = new Function('return ' + finalFormula)();
      
      if (typeof result === 'number' && !isNaN(result) && isFinite(result)) {
          // Limitar decimales para mejor visualización
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
    // Solo guardamos los datos, no los valores calculados
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

  return (
    <div className="pacur-hoja-container">
      
      {/* 1. Barra de Herramientas (Ribbon simplificado) */}
      <div className="toolbar">
        {/* Herramientas de Fuente */}
        <select defaultValue="Aptos Narrow" title="Fuente">
            <option>Aptos Narrow</option>
            <option>Arial</option>
            <option>Calibri</option>
        </select>
        <select defaultValue="11" title="Tamaño">
            <option>10</option>
            <option>11</option>
            <option>12</option>
        </select>
        <button onClick={() => alert("Negrita")} title="Negrita"><b>N</b></button>
        <button onClick={() => alert("Cursiva")} title="Cursiva"><i>K</i></button>
        <button onClick={() => alert("Subrayado")} title="Subrayado"><u>S</u></button>
        
        {/* Alineación */}
        <button onClick={() => alert("Izquierda")} title="Alinear Izquierda">⏴</button>
        <button onClick={() => alert("Centrar")} title="Centrar">☰</button>
        <button onClick={() => alert("Derecha")} title="Alinear Derecha">⏵</button>

        {/* Número y Estilos */}
        <select defaultValue="General" title="Formato de Número">
            <option>General</option>
            <option>Número</option>
            <option>Moneda</option>
            <option>Porcentaje</option>
        </select>
        <button onClick={() => alert("Moneda")} title="Formato Moneda">$</button>
        <button onClick={() => alert("Porcentaje")} title="Estilo Porcentual">%</button>
        
        {/* Botón de Guardar */}
        <button onClick={saveSheet} title="Guardar como .aph">💾 Guardar</button>
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
            // Importante: El input ahora muestra la fórmula/valor del estado 'data'
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
              // Llamamos a calculateValue para obtener el resultado
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
