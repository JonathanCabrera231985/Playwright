import { test, expect } from '@playwright/test';
import { WordReporter } from '../helper/WordReporter'; // Asegúrate que la ruta sea correcta

test('Generar reporte en Word ingresando a la pagina de Playwright', async ({ page }) => {
  // 1. Inicializar el reporter
  const reporter = new WordReporter();
  // 2. Iniciar la grabación de la prueba
  await page.goto('https://playwright.dev/');
    await page.waitForTimeout(1000);    
  reporter.addStep(page, 'Página de Playwright cargada');  
  // Esperar a que el título contenga "Playwright"
  //await expect(page).toHaveTitle(/Playwright/);
  //reporter.addStep(page, 'Título verificado: Contiene "Playwright"');
  // Click the get started link.


 
  await page.getByRole('link', { name: 'Get started' }).click();
    await page.waitForTimeout(1000);    
    
   // --- Insertar el Salto de Página ---
    reporter.addPageBreak();

  reporter.addStep(page, 'Clic en "Get started"');
    // Expects page to have a heading with the name of Installation.
  await expect(page.getByRole('heading', { name: 'Installation' })).toBeVisible();
  // Expects the URL to contain intro.
  //await expect(page).toHaveURL(/.*intro/);
  //reporter.addStep(page, 'URL verificada: Contiene "intro"');

  
  await page.waitForTimeout(1000);    

  // 5. Guardar el reporte en un archivo .docx
  await reporter.save('Reporte_Prueba_DuckDuckGo.docx');
});