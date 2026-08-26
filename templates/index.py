<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Gestión de Documentación</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            padding: 30px;
        }
        
        h1 {
            color: #333;
            margin-bottom: 10px;
            text-align: center;
        }
        
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
        }
        
        .tabs {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        
        .tab {
            padding: 10px 20px;
            background: #f0f0f0;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s;
        }
        
        .tab:hover {
            background: #e0e0e0;
            transform: translateY(-2px);
        }
        
        .tab.active {
            background: #667eea;
            color: white;
        }
        
        .tab-content {
            display: none;
        }
        
        .tab-content.active {
            display: block;
        }
        
        .form-section {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        
        .form-group {
            margin-bottom: 15px;
        }
        
        .form-group label {
            display: block;
            font-weight: 600;
            margin-bottom: 5px;
            color: #333;
        }
        
        .form-group input, .form-group select {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 14px;
        }
        
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }
        
        .btn {
            background: #667eea;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 5px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s;
        }
        
        .btn:hover {
            background: #5a67d8;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .btn-danger {
            background: #e53e3e;
        }
        
        .btn-danger:hover {
            background: #c53030;
        }
        
        .table-container {
            overflow-x: auto;
            margin-top: 20px;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        
        th {
            background: #667eea;
            color: white;
            padding: 12px;
            text-align: left;
        }
        
        td {
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        tr:hover {
            background: #f8f9fa;
        }
        
        .status-ok {
            color: #38a169;
            font-weight: 600;
        }
        
        .status-warning {
            color: #dd6b20;
            font-weight: 600;
        }
        
        .status-danger {
            color: #e53e3e;
            font-weight: 600;
        }
        
        .status-expired {
            color: #c53030;
            font-weight: 600;
            background: #fed7d7;
            padding: 3px 8px;
            border-radius: 3px;
        }
        
        .file-link {
            color: #667eea;
            text-decoration: none;
        }
        
        .file-link:hover {
            text-decoration: underline;
        }
        
        .loading {
            text-align: center;
            padding: 20px;
            color: #666;
        }
        
        .alert {
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        
        .alert-success {
            background: #c6f6d5;
            color: #22543d;
            border: 1px solid #9ae6b4;
        }
        
        .alert-error {
            background: #fed7d7;
            color: #742a2a;
            border: 1px solid #fc8181;
        }
        
        .required {
            color: #e53e3e;
        }
        
        @media (max-width: 768px) {
            .form-row {
                grid-template-columns: 1fr;
            }
            
            .tabs {
                flex-direction: column;
            }
            
            .tab {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🚛 Sistema de Gestión de Documentación</h1>
        <p class="subtitle">Control de vencimientos de documentación</p>
        
        <div id="alert-container"></div>
        
        <div class="tabs" id="tabs">
            <button class="tab active" data-sheet="Camion T1">🚚 Camión T1</button>
            <button class="tab" data-sheet="Camion T2">🚛 Camión T2</button>
            <button class="tab" data-sheet="Autoelevadores">🏗️ Autoelevadores</button>
            <button class="tab" data-sheet="Choferes y Ayudantes">👨‍✈️ Choferes y Ayudantes</button>
        </div>
        
        <div id="tab-content">
            <!-- El contenido de las tabs se carga dinámicamente -->
        </div>
    </div>
    
    <script>
        // Variables globales
        let currentSheet = 'Camion T1';
        let dataCache = {};
        
        // Función para mostrar alertas
        function showAlert(message, type = 'success') {
            const container = document.getElementById('alert-container');
            const alert = document.createElement('div');
            alert.className = `alert alert-${type}`;
            alert.textContent = message;
            container.appendChild(alert);
            
            setTimeout(() => {
                alert.remove();
            }, 5000);
        }
        
        // Función para formatear fecha
        function formatDate(dateString) {
            if (!dateString) return '';
            const date = new Date(dateString);
            return date.toLocaleDateString('es-ES', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit'
            });
        }
        
        // Función para calcular estado del vencimiento
        function getStatus(expirationDate) {
            if (!expirationDate) return { class: '', text: 'Sin fecha' };
            
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const expDate = new Date(expirationDate);
            expDate.setHours(0, 0, 0, 0);
            
            const diffDays = Math.ceil((expDate - today) / (1000 * 60 * 60 * 24));
            
            if (diffDays < 0) {
                return { class: 'status-expired', text: '⚠️ VENCIDO' };
            } else if (diffDays <= 30) {
                return { class: 'status-danger', text: `🔴 ${diffDays} días` };
            } else if (diffDays <= 90) {
                return { class: 'status-warning', text: `🟡 ${diffDays} días` };
            } else {
                return { class: 'status-ok', text: `✅ ${diffDays} días` };
            }
        }
        
        // Función para cargar datos de una hoja
        async function loadSheetData(sheetName) {
            const content = document.getElementById('tab-content');
            content.innerHTML = '<div class="loading">Cargando datos...</div>';
            
            try {
                const response = await fetch(`/api/sheet/${encodeURIComponent(sheetName)}`);
                const data = await response.json();
                dataCache[sheetName] = data;
                renderTable(sheetName, data);
            } catch (error) {
                content.innerHTML = `<div class="alert alert-error">Error al cargar datos: ${error.message}</div>`;
            }
        }
        
        // Función para renderizar la tabla
        function renderTable(sheetName, data) {
            const content = document.getElementById('tab-content');
            
            // Verificar que data sea un array
            if (!Array.isArray(data)) {
                data = [];
            }
            
            // Encontrar el índice de la columna "Fecha de carga" si existe
            const headers = data.length > 0 ? Object.keys(data[0]) : [];
            
            let html = `
                <div class="form-section">
                    <h3>Agregar nuevo documento</h3>
                    <form id="addForm">
                        <div class="form-row">
                            <div class="form-group">
                                <label>Patente <span class="required">*</span></label>
                                <input type="text" id="patente" placeholder="ABC123" required>
                            </div>
                            <div class="form-group">
                                <label>Documento <span class="required">*</span></label>
                                <input type="text" id="documento" placeholder="Nombre del documento" required>
                            </div>
                        </div>
                        <div class="form-row">
                            <div class="form-group">
                                <label>Fecha de Vencimiento <span class="required">*</span></label>
                                <input type="date" id="fechaVencimiento" required>
                            </div>
                            <div class="form-group">
                                <label>Foto del documento <span class="required">*</span></label>
                                <input type="file" id="foto" accept="image/*" required>
                            </div>
                        </div>
                        <input type="hidden" id="sheetName" value="${sheetName}">
                        <button type="submit" class="btn">Agregar Documento</button>
                    </form>
                </div>
            `;
            
            html += `<div class="table-container"><table>`;
            
            if (data.length === 0) {
                html += `<tr><td colspan="6" style="text-align: center; padding: 20px;">No hay documentos cargados</td></tr>`;
            } else {
                // Mostrar todas las columnas disponibles
                const allColumns = Object.keys(data[0]);
                
                // Definir columnas a mostrar (todas las disponibles)
                const columnsToShow = allColumns.filter(col => col !== 'Fecha de carga');
                
                html += `<thead><tr>`;
                // Siempre mostrar Patente, Documento, Fecha Vencimiento, Link, Estado
                const columnOrder = ['Patente', 'Documento', 'Fecha Vencimiento', 'Link'];
                const displayColumns = columnOrder.filter(col => allColumns.includes(col));
                // Agregar columnas adicionales que no están en el orden definido
                const extraColumns = allColumns.filter(col => !columnOrder.includes(col) && col !== 'Fecha de carga');
                const finalColumns = [...displayColumns, ...extraColumns];
                
                // Agregar columna de Estado
                html += `<th>Estado</th>`;
                
                // Agregar headers para las columnas finales
                finalColumns.forEach(col => {
                    html += `<th>${col}</th>`;
                });
                html += `</tr></thead><tbody>`;
                
                // Mostrar datos
                data.forEach((row, index) => {
                    html += `<tr>`;
                    
                    // Columna de estado
                    const fechaVenc = row['Fecha Vencimiento'] || row['FechaVencimiento'] || '';
                    const status = getStatus(fechaVenc);
                    html += `<td><span class="${status.class}">${status.text}</span></td>`;
                    
                    finalColumns.forEach(col => {
                        let value = row[col] || '';
                        if (col === 'Link' && value) {
                            value = `<a href="${value}" target="_blank" class="file-link">📎 Ver archivo</a>`;
                        }
                        html += `<td>${value}</td>`;
                    });
                    html += `</tr>`;
                });
            }
            
            html += `</tbody></table></div>`;
            content.innerHTML = html;
            
            // Event listener para el formulario
            document.getElementById('addForm').addEventListener('submit', handleAddDocument);
        }
        
        // Función para manejar el envío del formulario
        async function handleAddDocument(e) {
            e.preventDefault();
            
            const form = e.target;
            const formData = new FormData(form);
            
            // Validar campos obligatorios
            const patente = formData.get('patente');
            const documento = formData.get('documento');
            const fechaVencimiento = formData.get('fechaVencimiento');
            const foto = formData.get('foto');
            
            if (!patente || !documento || !fechaVencimiento || !foto || foto.size === 0) {
                showAlert('Todos los campos son obligatorios incluyendo la foto', 'error');
                return;
            }
            
            try {
                const response = await fetch('/api/add', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    showAlert('Documento agregado correctamente', 'success');
                    form.reset();
                    // Recargar los datos
                    const sheetName = formData.get('sheet_name') || document.getElementById('sheetName').value;
                    loadSheetData(sheetName);
                } else {
                    showAlert(`Error: ${result.error}`, 'error');
                }
            } catch (error) {
                showAlert(`Error al enviar el formulario: ${error.message}`, 'error');
            }
        }
        
        // Configurar tabs
        document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', function() {
                // Actualizar tab activa
                document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                this.classList.add('active');
                
                // Cargar datos de la hoja seleccionada
                const sheetName = this.dataset.sheet;
                currentSheet = sheetName;
                loadSheetData(sheetName);
            });
        });
        
        // Cargar datos iniciales
        loadSheetData('Camion T1');
    </script>
</body>
</html>