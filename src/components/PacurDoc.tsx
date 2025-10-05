import React from 'react';

// Define el tipo para las propiedades del componente
interface PacurDocProps {
  initialContent?: string;
}

// Componente principal del editor de documentos
const PacurDoc: React.FC<PacurDocProps> = ({ initialContent = '<h1>Empieza tu documento aquí</h1><p>Usa los botones para dar formato.</p>' }) => {

  // Función para aplicar formato (Negrita, Cursiva, etc.)
  const formatDoc = (command: string) => {
    // document.execCommand es la forma más sencilla de aplicar formato en el navegador
    document.execCommand(command, false, undefined);
    // Vuelve a enfocar el editor después de la acción
    document.getElementById('editor')?.focus();
  };

  // Función para guardar como archivo .apd
  const saveDoc = () => {
    const editor = document.getElementById('editor');
    if (!editor) return;

    // Almacenamos el HTML interno del editor, que mantiene el formato
    const content = editor.innerHTML;
    const filename = "nuevo_documento.apd";

    // Creamos un Blob (paquete de datos) para descargar el archivo
    const blob = new Blob([content], { type: 'text/html' }); 
    const a = document.createElement('a');
    a.download = filename;
    a.href = URL.createObjectURL(blob);
    
    // Simulamos el click para iniciar la descarga
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    alert(`¡Documento ${filename} guardado con éxito!`);
  };

  return (
    <div className="pacur-doc-container">
      
      {/* Barra de herramientas */}
      <div className="toolbar">
        <button onClick={() => formatDoc('bold')} title="Negrita"><b>B</b></button>
        <button onClick={() => formatDoc('italic')} title="Cursiva"><i>I</i></button>
        <button onClick={() => formatDoc('underline')} title="Subrayado"><u>U</u></button>
        <button onClick={saveDoc} title="Guardar como .apd">💾 Guardar</button>
      </div>

      {/* Área de edición con contentEditable */}
      <div 
        id="editor" 
        className="editor-area"
        contentEditable={true} // Permite la edición de texto enriquecido
        dangerouslySetInnerHTML={{ __html: initialContent }} // Establece el contenido inicial
        spellCheck={true}
      />
    </div>
  );
};

export default PacurDoc;
