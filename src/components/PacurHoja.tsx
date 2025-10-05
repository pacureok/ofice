import { useState, useMemo } from 'react'; // Solo se importa lo que se usa
import React from 'react'; // Necesario para definir React.FC

// Constantes para la cuadrícula
const ROWS = 50; // Aumentamos las filas para que se parezca a una hoja real
const COLS = 15; // Aumentamos las columnas (A a O)

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

  // Maneja la entrada de datos en una celda
  const handleCellChange = (key: string, value: string) => {
    setData(prevData => ({
      ...prevData,
      [key]: value
    }));
  };

  // Función de CÁLCULO simple (solo para demostración de la interfaz)
  const calculateValue = (key: string): string => {
    const content = data[key] || '';
    
    if (content.startsWith('=')) {
      try {
        // En una implementación real, aquí se usaría un parser seguro, NO 'eval'
        const formula = content.substring(1);
        const result = eval(formula); 
        return isNaN(result) ? '#ERROR' : result.toString();
        
      } catch (e) {
        return '#FÓRMULA_INVÁLIDA';
      }
    }
    return content;
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

    alert(`¡Hoja de cálculo ${filename} guardada con éxito!`);
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

