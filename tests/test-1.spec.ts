import { test, expect } from '@playwright/test';
import { WordReporter } from '../helper/WordReporter'; // Asegúrate que la ruta sea correcta

test('Generar reporte en Word ingresando a la pagina de Playwright', async ({ page }) => {
  // 1. Inicializar el reporter
  const reporter = new WordReporter();
  // 2. Iniciar la grabación de la prueba
  await page.goto('https://playwright.dev/');
    await page.waitForTimeout(1000);    

    try {
      // 1. Crea el encabezado primero, especificando la ruta a tu logo.
      
    reporter.createHeader('./assets/logo.png');
    
    reporter.addTitle('Reporte de Prueba Automatizada - Playwright');

    // 1. Agrega la tabla de cabecera con toda la infoxrmación del caso de prueba.
    reporter.addFeatureTable({
      functionality: 'Procesar documentos Parque de la paz',
      scenario: 'Verificar el guardado de datos',
      type: 'Stored Procedure',
      preconditions: '- Los ambientes deben estar sin errores de otros procesos.',
      requiredData: 'Comprobantes para procesar',
    });
  } catch (error) {
    console.error('Error al agregar la tabla de cabecera:', error);
  }
  reporter.addLineBreak();
  reporter.addStep(page, 'Página de Playwright cargada');  
  // Esperar a que el título contenga "Playwright"
  //await expect(page).toHaveTitle(/Playwright/);
  //reporter.addStep(page, 'Título verificado: Contiene "Playwright"');
  // Click the get started link.


 
  await page.getByRole('link', { name: 'Get started' }).click();
  await page.waitForTimeout(1000);    

   // --- Insertar el Salto de Página ---
    reporter.addPageBreak();

  reporter.addStep(page, 'Clic en "Get started" and show Installation heading');
    // Expects page to have a heading with the name of Installation.
  await expect(page.getByRole('heading', { name: 'Installation' })).toBeVisible();
  // Expects the URL to contain intro.
  //await expect(page).toHaveURL(/.*intro/);
  //reporter.addStep(page, 'URL verificada: Contiene "intro"');

  
  await page.waitForTimeout(1000);    

  // 5. Guardar el reporte en un archivo .docx
  await reporter.save('Reporte_Prueba_DuckDuckGo.docx');
});