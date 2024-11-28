function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu('Generar Diapositivas')
    .addItem('Generar...', 'generarDiapositivas')
    .addToUi();
}

function generarDiapositivas() {
  const ui = SlidesApp.getUi();

  // Solicitar número de diapositivas
  const numDiapositivasResponse = ui.prompt(
    'Número de Diapositivas',
    '¿Cuántas diapositivas deseas generar?',
    ui.ButtonSet.OK_CANCEL
  );

  if (numDiapositivasResponse.getSelectedButton() !== ui.Button.OK) return;

  const numDiapositivas = parseInt(numDiapositivasResponse.getResponseText().trim(), 10);

  if (isNaN(numDiapositivas) || numDiapositivas <= 0) {
    ui.alert('Por favor, ingresa un número válido de diapositivas (mayor a 0).');
    return;
  }

  const presentation = SlidesApp.getActivePresentation();

  for (let i = 0; i < numDiapositivas; i++) {
    // Solicitar título
    const tituloResponse = ui.prompt(
      `Diapositiva ${i + 1} - Título`,
      `Ingresa el título para la diapositiva ${i + 1}:`,
      ui.ButtonSet.OK_CANCEL
    );

    if (tituloResponse.getSelectedButton() !== ui.Button.OK) return;

    const titulo = tituloResponse.getResponseText().trim();
    if (!titulo) {
      ui.alert(`El título para la diapositiva ${i + 1} no puede estar vacío.`);
      return;
    }

    // Solicitar puntos
    const puntosResponse = ui.prompt(
      `Diapositiva ${i + 1} - Puntos`,
      'Ingresa tres puntos separados por punto y coma (;):',
      ui.ButtonSet.OK_CANCEL
    );

    if (puntosResponse.getSelectedButton() !== ui.Button.OK) return;

    const puntos = puntosResponse.getResponseText()
      .split(';')
      .map(punto => punto.trim())
      .filter(punto => punto);

    if (puntos.length === 0) {
      ui.alert(`No se ingresaron puntos válidos para la diapositiva ${i + 1}.`);
      return;
    }

    // Crear la diapositiva
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    // Añadir el título
    const titleBox = slide.insertTextBox(titulo);
    titleBox.setHeight(50);
    titleBox.setWidth(500);
    titleBox.setLeft(100);
    titleBox.setTop(50);

    const titleText = titleBox.getText();
    if (SlidesApp.TextAlignment && SlidesApp.TextAlignment.CENTER) {
      titleText.setTextAlignment(SlidesApp.TextAlignment.CENTER);
    }
    titleText.getTextStyle().setBold(true);

    // Añadir los puntos como lista con viñetas manuales
    const textBox = slide.insertTextBox('');
    textBox.setHeight(300);
    textBox.setWidth(500);
    textBox.setLeft(100);
    textBox.setTop(150);

    const textRange = textBox.getText();
    const puntosConViñetas = puntos.map(punto => `• ${punto}`).join('\n'); // Añade "• " manualmente
    textRange.setText(puntosConViñetas); // Inserta los puntos con viñetas
  }

  ui.alert('Diapositivas generadas exitosamente.');
}