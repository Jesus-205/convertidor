import React, { useState, useRef, useMemo } from 'react';
import { Upload, FileDown, RefreshCcw, FileSpreadsheet, CheckCircle2, AlertCircle, Search } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { parseExcel, exportToExcel } from './utils/excelParser';
import useSound from 'use-sound';
import './App.css';

function App() {
  const [data, setData] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const fileInputRef = useRef(null);

  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return data;
    const term = searchTerm.toLowerCase().trim();
    return data.filter((item) => {
      const barcode = String(item['Codigo de barras'] || '').toLowerCase();
      const nombre = String(item['Nombre de producto'] || '').toLowerCase();
      const categoria = String(item.Categoria || '').toLowerCase();
      const clase = String(item.Clase || '').toLowerCase();
      return barcode.includes(term) || nombre.includes(term) || categoria.includes(term) || clase.includes(term);
    });
  }, [data, searchTerm]);

  const handleFileUpload = async (file) => {
    if (!file) return;
    setLoading(true);
    setError(null);
    try {
      const cleanedData = await parseExcel(file);
      setData(cleanedData);
    } catch (err) {
      console.error(err);
      setError('Error al procesar el archivo. Asegúrate de que sea un Excel válido.');
    } finally {
      setLoading(false);
    }
  };

  const onDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileUpload(file);
  };

  const onFileChange = (e) => {
    const file = e.target.files[0];
    if (file) handleFileUpload(file);
  };

  const handleDownload = () => {
    exportToExcel(data, 'inventario_convertido.xlsx');
  };

  const reset = () => {
    setData([]);
    setError(null);
    setSearchTerm('');
  };

  return (
    <div className="container">
      <header>
        <motion.h1
          initial={{ opacity: 0, y: -20 }}
          animate={{ opacity: 1, y: 0 }}
        >
          Excel Converter
        </motion.h1>
        <p className="subtitle">Transforma inventarios desordenados en tablas profesionales al instante.</p>
      </header>

      <main style={{ width: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
        <AnimatePresence mode="wait">
          {data.length === 0 ? (
            <motion.div
              key="uploader"
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className={`upload-card ${isDragging ? 'dragging' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current.click()}
            >
              <div className="icon-wrapper">
                {loading ? (
                  <RefreshCcw className="animate-spin text-white" size={40} />
                ) : (
                  <Upload className="text-white" size={40} />
                )}
              </div>
              <h2>{loading ? 'Procesando...' : 'Arrastra tu archivo aquí'}</h2>
              <p>O haz clic para seleccionar un archivo Excel de tu equipo</p>
              <input
                type="file"
                ref={fileInputRef}
                onChange={onFileChange}
                accept=".xlsx, .xls"
                style={{ display: 'none' }}
              />
              {error && (
                <div style={{ marginTop: '1.5rem', color: '#ef4444', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem' }}>
                  <AlertCircle size={18} />
                  <span>{error}</span>
                </div>
              )}
            </motion.div>
          ) : (
            <motion.div
              key="results"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="results-section"
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <CheckCircle2 style={{ color: '#10b981' }} />
                  <h2 style={{ fontSize: '1.5rem' }}>Conversión exitosa</h2>
                </div>
                <p style={{ color: '#9ca3af' }}>
                  {searchTerm ? `${filteredData.length} de ${data.length}` : `${data.length}`} registros
                </p>
              </div>

              <div className="search-bar">
                <Search size={20} style={{ color: '#6b7280', flexShrink: 0 }} />
                <input
                  type="text"
                  placeholder="Buscar por código de barras, nombre, clase o categoría..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
                {searchTerm && (
                  <button
                    className="search-clear"
                    onClick={() => setSearchTerm('')}
                  >
                    ✕
                  </button>
                )}
              </div>

              <div className="table-container">
                <div style={{ overflowX: 'auto', maxHeight: '60vh' }}>
                  <table>
                    <thead>
                      <tr>
                        <th>Código de barras</th>
                        <th>Clase</th>
                        <th>Categoría</th>
                        <th>Producto</th>
                        <th>Medida</th>
                        <th>Densidad</th>
                        <th>Peso Vacío</th>
                        <th>Peso Lleno</th>

                      </tr>
                    </thead>
                    <tbody>
                      {filteredData.length === 0 ? (
                        <tr>
                          <td colSpan={8} style={{ textAlign: 'center', padding: '2rem', color: '#6b7280' }}>
                            No se encontraron registros para "{searchTerm}"
                          </td>
                        </tr>
                      ) : (
                        filteredData.map((item, index) => (
                          <tr key={index}>
                            <td style={{ fontFamily: 'monospace', fontSize: '0.85rem' }}>{item['Codigo de barras'] || '—'}</td>
                            <td>{item.Clase}</td>
                            <td>{item.Categoria}</td>
                            <td style={{ fontWeight: 600 }}>{item['Nombre de producto'] || '—'}</td>
                            <td>{item.Medida ? `${item.Medida}` : '—'}</td>
                            <td>{item.Densidad || '—'}</td>
                            <td>{item['Peso Vacio'] || '—'}</td>
                            <td>{item['Peso lleno'] || '—'}</td>

                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="actions">
                <button className="btn btn-secondary" onClick={reset}>
                  <RefreshCcw size={20} />
                  Subir otro
                </button>
                <button className="btn btn-primary" onClick={handleDownload}>
                  <FileDown size={20} />
                  Descargar Excel Limpio
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer style={{ marginTop: '4rem', color: '#4b5563', fontSize: '0.9rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <FileSpreadsheet size={16} />
          <span>Soporta formatos .xlsx y .xls de auditoría estándar.</span>
        </div>
      </footer>
    </div>
  );
}

export default App;
