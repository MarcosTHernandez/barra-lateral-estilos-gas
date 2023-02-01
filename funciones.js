// Variable donde guardaremos todos los estilos
var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Aula en la nube')
    .addItem('Mostrar barra lateral', 'mostrarBarraLateral')
    .addToUi();
}

function mostrarBarraLateral()
{
  var barra = HtmlService.createTemplateFromFile('BarraLateral')
    .evaluate()
    .setTitle('Barra lateral Aulaenlanube');
    SpreadsheetApp.getUi().showSidebar(barra);
}

function aplicarEstilo(estilo)
{
  // Primero borramos el estilo de las celdas activas
  borrarEstilos();

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas     = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('colorLetra'+estilo))
        .setBackground(estilos_sheet.getProperty('colorFondo'+estilo))
        .setFontSize(estilos_sheet.getProperty('sizeFuente'+estilo))
        .setValue('Estilo '+estilo);

  // APLICAR BORDES
  // Borde sup
  if(comprobarBordes('sup', estilo))
    celdas.setBorder(true,null,null,null,null,true,estilos_sheet.getProperty('BordeSupCo'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeSupSt'+estilo)));
  
// Borde izq
  if(comprobarBordes('izq', estilo))
    celdas.setBorder(null,true,null,null,true,null,estilos_sheet.getProperty('BordeIzqCo'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeIzqSt'+estilo)));

  // Borde inf
  if(comprobarBordes('inf', estilo))
    celdas.setBorder(null,null,true,null,null,true,estilos_sheet.getProperty('BordeInfCo'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeInfSt'+estilo)));

  // Borde der
  if(comprobarBordes('der', estilo))
    celdas.setBorder(null,null,null,true,true,null,estilos_sheet.getProperty('BordeDerCo'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeDerSt'+estilo)));
}

function comprobarBordes(borde, estilo)
{
  switch(borde)
  {
    case 'sup': return estilos_sheet.getProperty('BordeSupCo'+estilo) != null;
    case 'izq': return estilos_sheet.getProperty('BordeIzqCo'+estilo) != null;
    case 'inf': return estilos_sheet.getProperty('BordeInfCo'+estilo) != null;
    case 'der': return estilos_sheet.getProperty('BordeDerCo'+estilo) != null;
  }
}

function obtenerEnumBorde(tipoBorde)
{
  switch(tipoBorde)
  {
    case 'DOTTED': return SpreadsheetApp.BorderStyle.DOTTED;
    case 'DASHED': return SpreadsheetApp.BorderStyle.DASHED;
    case 'SOLID': return SpreadsheetApp.BorderStyle.SOLID;
    case 'SOLID_MEDIUM': return SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    case 'SOLID_THICK': return SpreadsheetApp.BorderStyle.SOLID_THICK;
    case 'DOUBLE': return SpreadsheetApp.BorderStyle.DOUBLE;
    default: return null;
  }
}

function guardarEstilo(estilo)
{
  // Borramos previamente los estilos
  eliminarEstilo(estilo);

  // Obtenemos la celda activa
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  // Guardamos los bordes
  guardarBordes(celda, estilo);

  // Guardamos colores y tamaños
  estilos_sheet.setProperty('colorLetra'+estilo, celda.getFontColor())
               .setProperty('colorFondo'+estilo, celda.getBackground())
               .setProperty('sizeFuente'+estilo, celda.getFontSize()+'');
  
  return{ colorFondo: estilos_sheet.getProperty('colorFondo'+estilo),
          colorLetra: estilos_sheet.getProperty('colorLetra'+estilo),
          BordeSupCo: estilos_sheet.getProperty('BordeSupCo'+estilo),
          BordeSupSt: estilos_sheet.getProperty('BordeSupSt'+estilo),
          BordeInfCo: estilos_sheet.getProperty('BordeInfCo'+estilo),
          BordeInfSt: estilos_sheet.getProperty('BordeInfSt'+estilo),
          BordeDerCo: estilos_sheet.getProperty('BordeDerCo'+estilo),
          BordeDerSt: estilos_sheet.getProperty('BordeDerSt'+estilo),
          BordeIzqCo: estilos_sheet.getProperty('BordeIzqCo'+estilo),
          BordeIzqSt: estilos_sheet.getProperty('BordeIzqSt'+estilo)
        };
}

function cargarEstilos()
{
  return estilos_sheet.getProperties();
}

function guardarBordes(celda, estilo)
{ 
  // Obtenemos los bordes
  var bordes = celda.getBorder();

  if(bordes != null)
  {
    var borde_sup = bordes.getTop();
    var borde_inf = bordes.getBottom();
    var borde_izq = bordes.getLeft();
    var borde_der = bordes.getRight();

    // Borde sup
    if(borde_sup.getColor() != null && borde_sup.getBorderStyle() != null)
    {
      estilos_sheet.setProperty('BordeSupCo'+estilo, borde_sup.getColor().asRgbColor().asHexString())
                   .setProperty('BordeSupSt'+estilo, borde_sup.getBorderStyle());
    }

    // Borde inf
    if(borde_inf.getColor() != null && borde_inf.getBorderStyle() != null)
    {
      estilos_sheet.setProperty('BordeInfCo'+estilo, borde_inf.getColor().asRgbColor().asHexString())
                   .setProperty('BordeInfSt'+estilo, borde_inf.getBorderStyle());
    }
    
    // Borde der
    if(borde_der.getColor() != null && borde_der.getBorderStyle() != null)
    {
      estilos_sheet.setProperty('BordeDerCo'+estilo, borde_der.getColor().asRgbColor().asHexString())
                   .setProperty('BordeDerSt'+estilo, borde_der.getBorderStyle());
    }
    // Borde izq
    if(borde_izq.getColor() != null && borde_izq.getBorderStyle() != null)
    {
      estilos_sheet.setProperty('BordeIzqCo'+estilo, borde_izq.getColor().asRgbColor().asHexString())
                   .setProperty('BordeIzqSt'+estilo, borde_izq.getBorderStyle());
    }

  }
}

function eliminarEstilo(estilo)
{
  // Colores
  estilos_sheet.deleteProperty('colorLetra'+estilo);
  estilos_sheet.deleteProperty('colorFondo'+estilo);
  estilos_sheet.deleteProperty('sizeFuente'+estilo);

  // Bordes
  estilos_sheet.deleteProperty('BordeSupCo'+estilo);
  estilos_sheet.deleteProperty('BordeSupSt'+estilo);
  estilos_sheet.deleteProperty('BordeInfCo'+estilo);
  estilos_sheet.deleteProperty('BordeInfSt'+estilo);
  estilos_sheet.deleteProperty('BordeIzqCo'+estilo);
  estilos_sheet.deleteProperty('BordeIzqSt'+estilo);
  estilos_sheet.deleteProperty('BordeDerCo'+estilo);
  estilos_sheet.deleteProperty('BordeDerSt'+estilo);
  
}

function borrarEstilos()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true});
}

function borrarTodo()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}