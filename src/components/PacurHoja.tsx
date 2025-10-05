import React, { useState, useMemo } from 'react';

// Constantes para la cuadrícula
const ROWS = 10; // 10 filas visibles
const COLS = 5;  // 5 columnas (A a E)

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

// --- Componente PacurHoja ---
const PacurHoja: React.FC = () => {
  // Estado para guardar el contenido de cada celda
  const [data, setData] = useState<SheetData>({});
  // Estado para la celda actualmente seleccionada/editada
  const [activeCell, setActiveCell] = useState<string | null>(null);
  
  const colHeaders = useMemo(() => getColHeaders(COLS), []);

  // Función simple para manejar la entrada de datos en una celda
  const handleCellChange = (key: string, value: string) => {
    setData(prevData => ({
      ...prevData,
      [key]: value
    }));
  };

  // Función de CÁLCULO muy simple (solo suma)
  const calculateValue = (key: string): string => {
    const content = data[key] || '';
    
    // Si la celda empieza con "=", asumimos que es una fórmula
    if (content.startsWith('=')) {
      try {
        // Ejemplo de cálculo súper básico: =10+20
        const formula = content.substring(1);
        
        // **AQUÍ VA LA LÓGICA COMPLEJA DE PARSEO DE FÓRMULAS**
        // Por ahora, solo evaluamos una expresión simple.
        // ADVERTENCIA: Usar 'eval' directamente no es seguro en producción.
        const result = eval(formula); 
        return isNaN(result) ? '#ERROR' : result.toString();
        
      } catch (e) {
        return '#FÓRMULA_INVÁLIDA';
      }
    }
    // Si no es una fórmula, devuelve el contenido
    return content;
  };

  // Función para guardar como .aph
  const saveSheet = () => {
    const filename = "hoja_calculo.aph";
    // Convertimos los datos a una cadena JSON
    const content = JSON.stringify(data, null, 2); 

    const blob = new Blob([content], { type: 'application/json' }); 
    const a = document.createElement('a');
    a.download = filename;
    a.href = URL.createObjectURL(blob);
    
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    alert(`¡Hoja de cálculo ${filename} guardada con éxito!`);
  };

  return (
    <div className="pacur-hoja-container">
      <div className="toolbar">
        {/* Celda de entrada para mostrar/editar la fórmula */}
        <input 
            type="text" 
            placeholder="Fórmula o valor" 
            value={activeCell ? data[activeCell] || '' : ''}
            onChange={(e) => activeCell && handleCellChange(activeCell, e.target.value)}
            className="formula-input"
        />
        <button onClick={saveSheet} title="Guardar como .aph">💾 Guardar (.aph)</button>
      </div>

      <div className="spreadsheet-grid">
        {/* Encabezados de columna */}
        <div className="header-row">
          <div className="cell header-cell corner-cell"></div>
          {colHeaders.map(header => (
            <div key={header} className="cell header-cell">{header}</div>
          ))}
        </div>

        {/* Filas de datos */}
        {Array.from({ length: ROWS }, (_, rIndex) => (
          <div key={rIndex} className="data-row">
            {/* Encabezado de fila */}
            <div className="cell header-cell">{(rIndex + 1)}</div>
            
            {/* Celdas de datos */}
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
                  {/* Campo de edición invisible, se activa con un doble click o al enfocar el input principal */}
                </div>
              );
            })}
          </div>
        ))}
      </div>
      <p className="nota-hoja">Nota: La lógica de cálculo es extremadamente simple (usa `eval`). Es el siguiente gran reto.</p>
    </div>
  );
};

export default PacurHoja;
