import React, { useState } from 'react';
import PacurDoc from './components/PacurDoc';
import PacurHoja from './components/PacurHoja'; // ¡Importa el nuevo componente!
import './App.css'; 

function App() {
  // Estado para simular el enrutamiento
  const [currentPage, setCurrentPage] = useState<'home' | 'doc' | 'sheet'>('home');

  const renderContent = () => {
    switch (currentPage) {
      case 'doc':
        return <PacurDoc />;
      case 'sheet':
        return <PacurHoja />; // Renderiza la hoja de cálculo
      default: // 'home'
        return (
          <div className="menu-principal">
            <div className="opcion">
              <h2>PacurDoc (Documento)</h2>
              <p>Crea, edita y guarda documentos con la extensión **.apd** (tipo Word).</p>
              <button onClick={() => setCurrentPage('doc')} className="boton-acceso">
                Abrir PacurDoc
              </button>
            </div>
            <div className="opcion">
              <h2>PacurHoja (Hoja de Cálculo)</h2>
              <p>Crea, edita y guarda hojas de cálculo con la extensión **.aph** (tipo Excel).</p>
              <button onClick={() => setCurrentPage('sheet')} className="boton-acceso">
                Abrir PacurHoja
              </button>
            </div>
          </div>
        );
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1 onClick={() => setCurrentPage('home')} style={{cursor: 'pointer'}}>
          Pacur Suite Ofimática
        </h1>
      </header>
      
      <main className="app-main">
        {renderContent()}
      </main>
    </div>
  );
}

export default App;
