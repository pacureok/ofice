import React, { useState } from 'react';
import PacurDoc from './components/PacurDoc';
import './App.css'; // Usaremos este archivo para todos los estilos

function App() {
  // Usaremos el estado para simular un enrutamiento (ej: '/' o '/doc')
  const [currentPage, setCurrentPage] = useState<'home' | 'doc'>('home');

  return (
    <div className="app-container">
      <header className="app-header">
        <h1 onClick={() => setCurrentPage('home')} style={{cursor: 'pointer'}}>
          Pacur Suite Ofimática
        </h1>
      </header>
      
      <main className="app-main">
        {currentPage === 'home' ? (
          <div className="menu-principal">
            <div className="opcion">
              <h2>PacurDoc (Documento)</h2>
              <p>Crea, edita y guarda documentos con la extensión .apd.</p>
              <button onClick={() => setCurrentPage('doc')} className="boton-acceso">
                Abrir PacurDoc
              </button>
            </div>
            <div className="opcion">
              <h2>PacurHoja (Hoja de Cálculo)</h2>
              <p>Crea, edita y guarda hojas de cálculo con la extensión .aph.</p>
              <button disabled className="boton-acceso deshabilitado" title="Próximamente">
                Abrir PacurHoja (Próximamente)
              </button>
            </div>
          </div>
        ) : (
          <PacurDoc />
        )}
      </main>
    </div>
  );
}

export default App;
