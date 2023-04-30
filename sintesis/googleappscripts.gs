// !!!!!!!!!!!!
// Cal tenir en compte que aquest script s'ha d'executar directament desde el menú de la taula, ja que s'integra amb aquesta.
// !!!!!!!!!!!!
var id = SpreadsheetApp.getActiveSpreadsheet().getId()
var idDocumentBase = "1XbfInNzt8ogDZovfTQeYPsAtsob-U43b0o1eOvVBi7c"
var idNotesBase = "14WlKmIv8bvP8L0DcHhf4Pp5sf09ronCyJ6aYpOQoC5U"
var idCarpetaProjecte = "1ojuimQw6S4pm5Le59wcn4WWeCSavbFKk" // Projecte 08 - ChatGPT
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var sheetAlum = sheets[0]
var sheetProf = sheets[1]
var sheetFami = sheets[2]
var sheetNota = sheets[3]

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu("Script")
      .addItem("Executa l'script", "ejecutar")
      .addSeparator()
      .addItem("Borra les carpetes", "eliminarCarpetes")
      .addToUi();
  SpreadsheetApp.getActiveSpreadsheet().toast("El menú ha cargado correctamente.", "⚠️ - Info", -1)
}

function ejecutar() {
  crearCarpetesTutories() // ya funciona
  compartirCarpetes() // ya funciona ((HACE FALTA DESTACAR QUE SI HAY MUCHOS PROFESORES/CARPETAS DE ALUMNOS, GOOGLE NOS METERÁ UN COOLDOWN Y NO PODREMOS COMPARTIR MAS))
  crearCarpetesAlumnes() // ya funciona (al menos con dos alumnos)
  crearButlletinsNotes() // ya funciona (se podrian añadir mas notas pero no sé)

  SpreadsheetApp.getActiveSpreadsheet().toast("El script s'ha executat sense errors!", "✅ - Oleee", -1)
}

function crearCarpetesTutories() {
  SpreadsheetApp.getActiveSpreadsheet().toast("El script se está iniciando.", "⚠️ - Info", -1)

  var data = sheetAlum.getDataRange().offset(1, 0, sheetAlum.getDataRange().getNumRows() - 1, sheetAlum.getDataRange().getNumColumns()).getValues();
  var archivoEnDrive = DriveApp.getFileById(id); 
  var drive = archivoEnDrive.getParents();

  if (drive.hasNext()) {
    var carpetaProyecto = drive.next();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("El spreadsheet no tiene carpeta madre.", "❌ - Liada", -1)
  }


  if (!carpetaProyecto.getFoldersByName("Carpeta tutoria").hasNext()) {
    carpetaProyecto = carpetaProyecto.createFolder("Carpeta tutoria")
  } else {
    carpetaProyecto = carpetaProyecto.getFoldersByName("Carpeta tutoria").next()
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Creando carpetas de cursos...", "⚠️ - Info", -1)
  data.forEach(function (fila) {
    var letra = calcularCurs(fila)
    if (!carpetaProyecto.getFoldersByName(letra).hasNext()) {
      carpetaProyecto.createFolder(letra)
    }
  })
}

function compartirCarpetes() {
  SpreadsheetApp.getActiveSpreadsheet().toast(`Inicializando variables de permisos de carpetas.`, "⚠️ - Info", -1)
  var dataProfes = sheetProf.getDataRange().offset(1, 0, sheetProf.getDataRange().getNumRows() - 1, sheetProf.getDataRange().getNumColumns()).getValues();
  carpetaProjecte = DriveApp.getFolderById(idCarpetaProjecte)
  carpeta = carpetaProjecte.getFoldersByName("Carpeta tutoria").next()
  var carpetasCursos = carpeta.getFolders();
  var arrCarpetas = []
  while (carpetasCursos.hasNext()) {
    var carpetaCurso = carpetasCursos.next();
    arrCarpetas.push(carpetaCurso)
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Modificando permisos de carpetas de cursos... ${arrCarpetas.length} carpetas restantes.`, "⚠️ - Info", -1)
  arrCarpetas.forEach(function (carpetaCurso, index) {
    dataProfes.forEach(function (profe) {
      if (carpetaCurso.getName() == profe[3]) {
        carpetaCurso.addEditor(profe[6])
      } else {
        carpetaCurso.addViewer(profe[6])
      }
    })
    SpreadsheetApp.getActiveSpreadsheet().toast(`Modificando permisos de carpetas de cursos... ${arrCarpetas.length-(index+1)} carpetas restantes.`, "⚠️ - Info", -1)
  })
  SpreadsheetApp.getActiveSpreadsheet().toast("Los permisos de las carpetas se han modificado correctamente.", "✅ - Oleee", -1)
}

function crearCarpetesAlumnes() {
  var carpetaTutoria = DriveApp.getFoldersByName("Carpeta tutoria").next()
  SpreadsheetApp.getActiveSpreadsheet().toast("Creando carpetas de alumnos...", "⚠️ - Info", -1)
  var dataAlum = sheetAlum.getDataRange().offset(1, 0, sheetAlum.getDataRange().getNumRows() - 1, sheetAlum.getDataRange().getNumColumns()).getValues();

  dataAlum.forEach((alumno) => {
    var letra = calcularCurs(alumno)
    var nombreAlumno = `${alumno[2]} ${alumno[3]}, ${alumno[4]}`
    var carpetaCurso = carpetaTutoria.getFoldersByName(letra).next()

    if (!carpetaCurso.getFoldersByName(nombreAlumno).hasNext()) {
      var carpetaAlumno = carpetaCurso.createFolder(nombreAlumno)
    }
    else {
      var carpetaAlumno = carpetaCurso.getFoldersByName(nombreAlumno).next()
    }
    if (!carpetaAlumno.getFoldersByName("Justificants").hasNext()) {
      var carpetaTrimestre = carpetaAlumno.createFolder("Justificants")
    }
    if (!carpetaAlumno.getFoldersByName("Trimestre 1").hasNext()) {
      var carpetaTrimestre = carpetaAlumno.createFolder("Trimestre 1")
    }
    if (!carpetaAlumno.getFoldersByName("Trimestre 2").hasNext()) {
      var carpetaTrimestre = carpetaAlumno.createFolder("Trimestre 2")
    }
    if (!carpetaAlumno.getFoldersByName("Trimestre 3").hasNext()) {
      var carpetaTrimestre = carpetaAlumno.createFolder("Trimestre 3")
    }
  })
}

function crearButlletinsNotes() {
  var carpetaTutoria = DriveApp.getFoldersByName("Carpeta tutoria").next()
  SpreadsheetApp.getActiveSpreadsheet().toast("Creando bulletines de nota...", "⚠️ - Info", -1)
  var dataNota = sheetNota.getDataRange().offset(1, 0, sheetNota.getDataRange().getNumRows() - 1, sheetNota.getDataRange().getNumColumns()).getValues();

  dataNota.forEach((alumno) => {
    var nombreAlumno = `${alumno[2]} ${alumno[3]}, ${alumno[4]}`
    var letra = calcularCurs(alumno) 
    var carpetaCurso = carpetaTutoria.getFoldersByName(letra).next()
    var carpetaAlumno = carpetaCurso.getFoldersByName(nombreAlumno).next()
    if (carpetaAlumno.getFoldersByName("Trimestre 1").hasNext()) {

      var trimestre = 1
      var carpetaTrimestre = carpetaAlumno.getFoldersByName("Trimestre 1").next()
      var butlleti = DriveApp.getFileById(idNotesBase).makeCopy(`Notes ${alumno[4]} - T${trimestre.toString()}`, carpetaTrimestre)
      var archivoNotas = DocumentApp.openById(butlleti.getId())
      var archivoEnviar1 = DriveApp.getFileById(butlleti.getId())

      afegirNotes(alumno, trimestre, archivoNotas)
    }
    if (carpetaAlumno.getFoldersByName("Trimestre 2").hasNext()) {

      var trimestre = 2
      var carpetaTrimestre = carpetaAlumno.getFoldersByName("Trimestre 2").next()
      var butlleti = DriveApp.getFileById(idNotesBase).makeCopy(`Notes ${alumno[4]} - T${trimestre.toString()}`, carpetaTrimestre)
      var archivoNotas = DocumentApp.openById(butlleti.getId())
      var archivoEnviar2 = DriveApp.getFileById(butlleti.getId())

      afegirNotes(alumno, trimestre, archivoNotas)
    }
    if (carpetaAlumno.getFoldersByName("Trimestre 3").hasNext()) {

      var trimestre = 3
      var carpetaTrimestre = carpetaAlumno.getFoldersByName("Trimestre 3").next()
      var butlleti = DriveApp.getFileById(idNotesBase).makeCopy(`Notes ${alumno[4]} - T${trimestre.toString()}`, carpetaTrimestre)
      var archivoNotas = DocumentApp.openById(butlleti.getId())
      var archivoEnviar3 = DriveApp.getFileById(butlleti.getId())

      afegirNotes(alumno, trimestre, archivoNotas)
    }
    enviarButlleti(alumno, archivoEnviar1, archivoEnviar2, archivoEnviar3)
  })
}

// Funciones de utilidades

function eliminarCarpetes() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Eliminando las carpetas...", "⚠️ - Info", -1)
  var ui = SpreadsheetApp.getUi();
  var archivoEnDrive = DriveApp.getFileById(id); 
  var drive = archivoEnDrive.getParents();


  while (drive.hasNext()) {
    var carpetaProyecto = drive.next();
 }

  if (!carpetaProyecto.getFoldersByName("Carpeta tutoria").hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast("La carpeta que intentas eliminar no existe.", "❌ - Liada", -1)
  } else {
    carpetaProyecto.getFoldersByName("Carpeta tutoria").next().setTrashed(true)
    SpreadsheetApp.getActiveSpreadsheet().toast("Las carpetas se han eliminado correctamente.", "✅ - Oleee", -1)
  }
}

function afegirNotes(alumno, trimestre, documento) {
  var cuerpo = documento.getBody();
  cuerpo.replaceText("inputalumne", `${alumno[4]} ${alumno[2]} ${alumno[3]}`);
  cuerpo.replaceText("inputaval", `T${trimestre.toString()}`);
  cuerpo.replaceText("inputcurs", calcularCurs(alumno));
  cuerpo.replaceText("inputm4uf1", alumno[6]);
  cuerpo.replaceText("inputm6uf1", alumno[7]);
  cuerpo.replaceText("inputm7uf2", alumno[8]);
  cuerpo.replaceText("inputm8uf1", alumno[9]);
  cuerpo.replaceText("inputm14uf1", alumno[10]);
  documento.saveAndClose()
}

function enviarButlleti(alumno, documento1, documento2, documento3) {
  GmailApp.sendEmail(alumno[11], `Notes de ${alumno[4]} ${alumno[2]} ${alumno[3]}`, 'Jo no li habria aprovat...', {
    attachments: [documento1.getAs(MimeType.PDF), documento2.getAs(MimeType.PDF), documento3.getAs(MimeType.PDF)],
    name: documento1.getName()
  });
}

function calcularCurs(fila) {
  switch(fila[0]) {
    case 1:
      letra = "r";
      break;
    case 2:
      letra = "n";
      break;
    case 3:
      letra = "r";
      break;
    case 4:
      letra = "t";
      break;
    default:
      letra = "";
    }
  return fila[0]+letra+fila[1]
}
