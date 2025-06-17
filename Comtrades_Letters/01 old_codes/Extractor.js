const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');

class CorrespondenciaExtractor {
  constructor() {
    this.browser = null;
    this.page = null;
  }

  async inicializar() {
    this.browser = await chromium.launch({ 
      headless: false // Cambiar a true para modo sin interfaz
    });
    this.page = await this.browser.newPage();
    
    // Configurar timeouts
    this.page.setDefaultTimeout(30000);
  }

  async buscarCorrespondencia(correlativo, period = '', docType = 'T') {
    const url = `https://correspondencia.coordinador.cl/correspondencia/busqueda?query=${correlativo}&period=${period}&doc_type=${docType}`;
    
    console.log(`Buscando: ${correlativo}`);
    
    try {
      await this.page.goto(url);
      
      // Esperar a que cargue el contenido
      await this.page.waitForTimeout(3000);
      
      // Extraer datos de la página
      const datos = await this.page.evaluate(() => {
        const extraerFecha = (texto) => {
          const patrones = [/\d{2}\/\d{2}\/\d{4}/, /\d{2}-\d{2}-\d{4}/, /\d{4}-\d{2}-\d{2}/];
          for (const patron of patrones) {
            const match = texto.match(patron);
            if (match) return match[0];
          }
          return null;
        };

        const extraerCorrelativo = (texto) => {
          const patrones = [/[A-Z]{2}\d{5}-\d{2}/, /[A-Z]+\d+-\d+/];
          for (const patron of patrones) {
            const match = texto.match(patron);
            if (match) return match[0];
          }
          return null;
        };

        const extraerEmpresa = (textos) => {
          const keywords = ['S.A.', 'LTDA', 'SPA', 'ENEL', 'CGE', 'CHILQUINTA'];
          for (const texto of textos) {
            if (keywords.some(keyword => texto.toUpperCase().includes(keyword))) {
              return texto.trim();
            }
          }
          return textos.find(t => t.length > 10) || null;
        };

        const verificarComtrade = (texto) => {
          const textoUpper = texto.toUpperCase();
          if (textoUpper.includes('COMTRADE')) {
            if (['ENVÍA', 'ENVIA', 'SÍ', 'SI'].some(palabra => textoUpper.includes(palabra))) {
              return 'Sí envía';
            } else if (['NO ENVÍA', 'NO ENVIA', 'NO'].some(palabra => textoUpper.includes(palabra))) {
              return 'No envía';
            }
            return 'Menciona Comtrade';
          }
          return 'Sin información';
        };

        // Buscar tabla o filas de resultados
        const filas = Array.from(document.querySelectorAll('tr, .result-row, .correspondence-item'));
        
        if (filas.length === 0) {
          console.log('No se encontraron filas');
          return [];
        }

        const resultados = [];
        
        filas.forEach((fila, index) => {
          const celdas = Array.from(fila.querySelectorAll('td, div, span'));
          const textos = celdas.map(celda => celda.textContent.trim()).filter(texto => texto);
          
          if (textos.length === 0) return;
          
          const textoCompleto = textos.join(' ');
          
          const resultado = {
            numeroFila: index + 1,
            fechaEnvio: extraerFecha(textoCompleto),
            empresa: extraerEmpresa(textos),
            correlativo: extraerCorrelativo(textoCompleto),
            comtradeRespuesta: verificarComtrade(textoCompleto),
            textoCompleto: textos.join(' | '),
            correlativoBuscado: window.correlativoBuscado || ''
          };
          
          // Solo agregar si tiene información útil
          if (resultado.correlativo || resultado.fechaEnvio || textos.length > 2) {
            resultados.push(resultado);
          }
        });
        
        return resultados;
      });
      
      // Pasar el correlativo buscado al contexto de la página
      await this.page.evaluate((corr) => {
        window.correlativoBuscado = corr;
      }, correlativo);
      
      return datos;
      
    } catch (error) {
      console.error(`Error buscando ${correlativo}:`, error.message);
      return [];
    }
  }

  async buscarMultiples(correlativos) {
    const todosResultados = [];
    
    for (const correlativo of correlativos) {
      const resultados = await this.buscarCorrespondencia(correlativo);
      
      // Agregar correlativo buscado a cada resultado
      resultados.forEach(resultado => {
        resultado.correlativoBuscado = correlativo;
      });
      
      todosResultados.push(...resultados);
      
      // Pausa entre búsquedas
      await this.page.waitForTimeout(2000);
    }
    
    return todosResultados;
  }

  async exportarAExcel(datos, nombreArchivo = 'correspondencia_extraida.xlsx') {
    try {
      if (!datos || datos.length === 0) {
        console.log('No hay datos para exportar');
        return;
      }

      // Preparar datos para Excel
      const datosParaExcel = datos.map(item => ({
        'Correlativo Buscado': item.correlativoBuscado || '',
        'Correlativo Encontrado': item.correlativo || '',
        'Fecha de Envío': item.fechaEnvio || '',
        'Empresa': item.empresa || '',
        'Respuesta Comtrade': item.comtradeRespuesta || '',
        'Número de Fila': item.numeroFila || '',
        'Texto Completo': item.textoCompleto || '',
        'Fecha de Extracción': new Date().toLocaleString('es-CL')
      }));

      // Crear workbook y worksheet
      const ws = XLSX.utils.json_to_sheet(datosParaExcel);
      const wb = XLSX.utils.book_new();
      
      // Ajustar ancho de columnas
      const colWidths = [
        { wch: 18 }, // Correlativo Buscado
        { wch: 18 }, // Correlativo Encontrado
        { wch: 12 }, // Fecha de Envío
        { wch: 30 }, // Empresa
        { wch: 15 }, // Respuesta Comtrade
        { wch: 8 },  // Número de Fila
        { wch: 50 }, // Texto Completo
        { wch: 20 }  // Fecha de Extracción
      ];
      ws['!cols'] = colWidths;
      
      XLSX.utils.book_append_sheet(wb, ws, 'Correspondencia');
      
      // Guardar archivo
      XLSX.writeFile(wb, nombreArchivo);
      
      console.log(`✅ Excel exportado exitosamente: ${nombreArchivo}`);
      console.log(`📊 Total de registros: ${datosParaExcel.length}`);
      
      return nombreArchivo;
      
    } catch (error) {
      console.error('Error exportando a Excel:', error);
      throw error;
    }
  }

  async exportarConFormato(datos, nombreArchivo = 'correspondencia_formateada.xlsx') {
    const ExcelJS = require('exceljs');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Correspondencia', {
      pageSetup: { paperSize: 9, orientation: 'landscape' }
    });

    // Definir columnas
    worksheet.columns = [
      { header: 'Correlativo Buscado', key: 'correlativoBuscado', width: 18 },
      { header: 'Correlativo Encontrado', key: 'correlativo', width: 18 },
      { header: 'Fecha de Envío', key: 'fechaEnvio', width: 12 },
      { header: 'Empresa', key: 'empresa', width: 30 },
      { header: 'Respuesta Comtrade', key: 'comtradeRespuesta', width: 15 },
      { header: 'Texto Completo', key: 'textoCompleto', width: 50 }
    ];

    // Agregar datos
    datos.forEach(row => {
      worksheet.addRow({
        correlativoBuscado: row.correlativoBuscado || '',
        correlativo: row.correlativo || '',
        fechaEnvio: row.fechaEnvio || '',
        empresa: row.empresa || '',
        comtradeRespuesta: row.comtradeRespuesta || '',
        textoCompleto: row.textoCompleto || ''
      });
    });

    // Formatear header
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };

    // Formatear filas alternadas
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && rowNumber % 2 === 0) {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF2F2F2' }
        };
      }
    });

    // Guardar archivo
    await workbook.xlsx.writeFile(nombreArchivo);
    console.log(`✅ Excel con formato exportado: ${nombreArchivo}`);
  }

  async cerrar() {
    if (this.browser) {
      await this.browser.close();
    }
  }
}

// Función principal
async function main() {
  const extractor = new CorrespondenciaExtractor();
  
  try {
    await extractor.inicializar();
    
    // Ejemplo 1: Buscar un solo correlativo
    console.log('🔍 Buscando correlativo único...');
    const resultadosUnico = await extractor.buscarCorrespondencia('DE01746-25');
    
    if (resultadosUnico.length > 0) {
      await extractor.exportarAExcel(resultadosUnico, 'correspondencia_unica.xlsx');
    }
    
    // Ejemplo 2: Buscar múltiples correlativos
    console.log('\n🔍 Buscando múltiples correlativos...');
    const correlativos = ['DE01746-25', 'DE01747-25', 'DE01748-25'];
    const resultadosMultiples = await extractor.buscarMultiples(correlativos);
    
    if (resultadosMultiples.length > 0) {
      await extractor.exportarAExcel(resultadosMultiples, 'correspondencia_multiple.xlsx');
      // También exportar con formato
      await extractor.exportarConFormato(resultadosMultiples, 'correspondencia_formateada.xlsx');
    }
    
    console.log('\n✅ Proceso completado exitosamente!');
    
  } catch (error) {
    console.error('❌ Error en el proceso:', error);
  } finally {
    await extractor.cerrar();
  }
}

// Ejecutar si es llamado directamente
if (require.main === module) {
  main();
}

module.exports = CorrespondenciaExtractor;