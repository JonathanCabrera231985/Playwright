import { test, expect } from '@playwright/test';
import { WordReporterTest ,  TableCellOptions} from '../helper/WordReporterTest';
import { AlignmentType } from 'docx';

test('test', async ({ page }) => {
    test.setTimeout(60000); // 60 segundos
 let loginTime = 0; 
 let totalTime = 0;// Variable para sumar tiempos
 // 1. Inicializar el reporter
  const reporter = new WordReporterTest();
  // 1. Llama al nuevo método para obtener la fecha formateada
  const fechaFormateada = reporter.getFormattedDate();

  try {
    // 1. Crea el encabezado primero, especificando la ruta a tu logo.
    reporter.createHeader('./assets/LogoTest.png');
  
  
  try {
      // 1. Crea el encabezado primero, especificando la ruta a tu logo.
      
    reporter.createHeader('./assets/LogoTest.png');

    reporter.addLineBreak(4); // Agrega dos saltos de línea después del título
    reporter.addTitle('Reporte de Prueba Automatizada - Expalsa', true, 36, '4F81BD', AlignmentType.CENTER);
    reporter.addLineBreak(6); // Agrega tres saltos de línea después del título
    reporter.addTitle('QA-IPA - Ticket 27926-Novedad con retenciones LatamFoods   -Agente de retención - Expalsa-Soporte', false, 24, '4F81BD', AlignmentType.CENTER);
    reporter.addLineBreak(8); // Agrega ocho saltos de línea después del título
    reporter.addTitle(fechaFormateada, true, 36, '4F81BD', AlignmentType.CENTER);
    reporter.addLineBreak(20); // Agrega veinte saltos de línea después del título
    reporter.addTitle('VERSIÓN 1.0', true, 36, '4F81BD', AlignmentType.CENTER);
    reporter.addPageBreak();  

    

  await page.goto('http://cm78qwinv3sipe:9091/cuenta/Login.aspx');
  
// O medir otra acción
   const loginTime = await reporter.measureStep(page, 'Página de Portal cargada', async () => {
        await page.locator('#MainContent_tbRfcuser').click();
        await page.locator('#MainContent_tbRfcuser').fill('emple110001');
        await page.locator('#MainContent_tbRfcuser').press('Tab');
        await page.locator('#MainContent_tbPass').fill('123456');
    });
  
    // 2. Imprime el valor en la consola y úsalo para una aserción
    console.log(`El login tardó: ${loginTime} segundos.`);
    //expect(loginTime).toBeLessThan(5); // Valida que el login no tarde más de 5 segundos

    totalTime += loginTime; // Suma el tiempo al total
  
    await page.waitForTimeout(1000);   
  reporter.addStep(page, 'Clic en inicio de sesión con usuario emple110001');


  await page.getByRole('button', { name: 'Ingresar' }).click();
  await page.getByRole('link', { name: 'Documentos' }).click();
  await page.getByRole('link', { name: 'Consulta' }).click();
  await page.waitForTimeout(1000);   
  reporter.addStep(page, 'El componente de consulta de documentos se muestra correctamente');

// 1. Define los datos y el formato de tu tabla
    const tablaDeDatos = [
      // Fila 1 (Cabecera)
      [
        { text: 'Módulo', backgroundColor: '4F81BD', textColor: 'FFFFFF', bold: true },
        { text: 'Resultado', backgroundColor: '4F81BD', textColor: 'FFFFFF', bold: true },
        { text: 'Tiempo (s)', backgroundColor: '4F81BD', textColor: 'FFFFFF', bold: true },
      ],
      // Fila 2
      [
        { text: 'Inicio de Sesión' },
        { text: 'Exitoso', backgroundColor: 'C6EFCE', textColor: '006100' }, // Fondo verde, texto verde
        { text: `${totalTime}s.`},
      ],
      // Fila 3
      [
        { text: 'Carga de Archivos' },
        { text: 'Fallido', backgroundColor: 'FFC7CE', textColor: '9C0006' }, // Fondo rojo, texto rojo
        { text: '5.4s' },
      ],
       // Fila 4
       [
        { text: 'Observaciones Generales', bold: true, columnSpan: 3, backgroundColor: 'D3D3D3' },
      ],
    ];

    // 2. Llama al método para crear la tabla
    reporter.createCustomTable(tablaDeDatos);
    reporter.addLineBreak(4); // Agrega dos saltos de línea después del título

/*
    // 1. Agrega la tabla de cabecera con toda la infoxrmación del caso de prueba.
    reporter.addFeatureTable({
      functionality: 'Procesar documentos Parque de la paz',
      scenario: 'Verificar el guardado de datos',
      type: 'Stored Procedure',
      preconditions: '- Los ambientes deben estar sin errores de otros procesos.',
      requiredData: 'Comprobantes para procesar',
    });*/
  } catch (error) {
    console.error('Error al agregar la tabla de cabecera:', error);
  }




  await page.locator('#MainContent_tbFolioAnterior').click();
  await page.locator('#MainContent_tbFolioAnterior').fill('001001004000013');
  await page.getByRole('button', { name: 'Buscar' }).click();
  //await page.screenshot({ path: 'screenshots/search-results.png' });
  await page.getByRole('link', { name: 'Cerrar Sesion' }).click();
  //await page.screenshot({ path: 'screenshots/logged-out.png' });  
  await page.waitForTimeout(1000);   
  await reporter.save('Informe de Pruebas de Aceptación1.docx');

  } catch (error) {
    // --- BLOQUE CATCH: SE ACTIVA SI ALGO FALLA ---

    // 1. Añade un subtítulo rojo para destacar el error.
    reporter.addSubtitle('LA PRUEBA HA FALLADO');
    
    // 2. Agrega el mensaje de error técnico al reporte.
    if (error instanceof Error) {
        reporter.addTitle('Mensaje de Error:', false, 24, 'FF0000', AlignmentType.LEFT);
        // Usamos addTitle aquí porque permite más personalización de formato
        reporter.addTitle(error.message, false, 20, '000000', AlignmentType.LEFT);
    }
    
    // 3. Toma una captura de pantalla final del momento del fallo.
    await reporter.addStep(page, 'Captura de pantalla final en el momento del error.');

    // 4. Vuelve a lanzar el error para que Playwright marque la prueba como "Fallida".
    // Si omites esta línea, la prueba aparecerá como "Exitosa".
    throw error; 
    
  } finally {
    // --- BLOQUE FINALLY: SE EJECUTA SIEMPRE ---
    
    // Guarda el archivo de Word, ya sea que la prueba haya pasado o fallado.
    await reporter.save('reporte-con-manejo-de-fallos.docx');
  }
  
});