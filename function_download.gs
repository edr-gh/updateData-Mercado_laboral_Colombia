function descargarArchivos() {

  // Obteniendo los URL de mercado laboral 

  /* Datos de fecha para URL de empleo y desempleo */
    let fecha = new Date()
    let month = fecha.getMonth()+1
    let year =  fecha.getFullYear()

    let meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]

    let previous_month

    // Año y mes de los datos
    if (month===1){
      previous_month = "dic"
      year = year -1
    }
    else{
      previous_month = meses[month-2] 
    }

  let url_empleo_y_desempleo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIH-"+previous_month+year+".xlsx"

  // Datos para fecha de las demás URLs
    // mes inicial
      let initial_month

      if(month-4<=0){
        let ref1 = (month-4)*-1
        initial_month = meses.reverse()[ref1]
        meses.reverse()
      }

      else{
        initial_month = meses[month-5]
      }
    
    /* ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]    meses*/ 
    /* ["dic","nov","oct","sep","ago","jul","jun","may","abr","mar","feb","ene"]    meses revertido*/

    // mes final
    let final_month
    
    if(month-2<=0){
      let ref2 = (month-2)*-1
      final_month = meses.reverse()[ref2]
      meses.reverse()
    }

    else{
      final_month = meses[month-3]
    }
    
    let initial_year
    if(initial_month === "nov" || initial_month === "dic"){
      initial_year = fecha.getFullYear()-1
    }
    else{
      initial_year = ""
    }

    let final_year
    if(initial_month === "sep" || initial_month ==="oct"){
      final_year = fecha.getFullYear()-1
    }
    else{
      final_year = fecha.getFullYear()
    }

    
    let url_mercado_segun_sexo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
    let url_mercado_juventud = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLJ-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
    let url_informalidad = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHEISS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"

    Logger.log("URL's iniciales")
    Logger.log(url_empleo_y_desempleo)
    Logger.log(url_informalidad)
    Logger.log(url_mercado_juventud)
    Logger.log(url_mercado_segun_sexo)

    // Descargando los archivos de las URLs

    let params = {muteHttpExceptions:true, method: "get"}

    archivos = {
      "Empleo y desempleo.xlsx": UrlFetchApp.fetch(url_empleo_y_desempleo, params),
      "Mercado laboral de la juventud.xlsx": UrlFetchApp.fetch(url_mercado_juventud, params),
      "Mercado laboral según sexo.xlsx":UrlFetchApp.fetch(url_mercado_segun_sexo, params),
      "Empleo informal y seguridad social.xlsx": UrlFetchApp.fetch(url_informalidad, params)
    }

  // Verificar que los archivos a descargar existen en la web

    Object.entries(archivos).forEach(function([nombreArchivo, solicitud]){
      let responseCode = solicitud.getResponseCode()
      
      if(responseCode !== 200){
        Logger.log("Responde code: %s", responseCode)
        Logger.log(nombreArchivo)
        let choice =  Browser.msgBox("El archivo "+nombreArchivo+" no está disponible ¿Desea descargar el último archivo disponible?",
        Browser.Buttons.YES_NO_CANCEL)
        if(choice === "cancel" || choice === "no"){
          Logger.log("Se evitó la descarga de los últimos archivos disponibles")
        }
        else{
          Browser.msgBox("Descargando...", Browser.Buttons.OK)
          if(nombreArchivo === "Empleo y desempleo.xlsx"){
            if (month===1){
              previous_month = "nov"
            }
            else{
              previous_month = meses[month-3]
            }

            url_empleo_y_desempleo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIH-"+previous_month+year+".xlsx"
          }
          else{

            if(month-5<=0){
              let ref1 = (month-5)*-1
              initial_month = meses.reverse()[ref1]
              meses.reverse()
            }

            else{
              initial_month = meses[month-6]
            }
            if(month-3<=0){
              let ref2 = (month-3)*-1
              final_month = meses.reverse()[ref2]
              meses.reverse()
            }

            else{
              final_month = meses[month-4]
            }
          
            if(initial_month === "nov" || initial_month === "dic"){
              initial_year = fecha.getFullYear()-1
            }
            else{
              initial_year = ""
            }
            if(initial_month === "sep" || initial_month ==="oct"){
              final_year = fecha.getFullYear()-1
            }
            else{
              final_year = fecha.getFullYear()
            }
          }
          url_mercado_segun_sexo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
          url_mercado_juventud = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLJ-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
          url_informalidad = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHEISS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
        }
      }
    })
  
    archivos = {
      "Empleo y desempleo.xlsx": UrlFetchApp.fetch(url_empleo_y_desempleo, params),
      "Mercado laboral de la juventud.xlsx": UrlFetchApp.fetch(url_mercado_juventud, params),
      "Mercado laboral según sexo.xlsx":UrlFetchApp.fetch(url_mercado_segun_sexo, params),
      "Empleo informal y seguridad social.xlsx": UrlFetchApp.fetch(url_informalidad, params)
  }
  
    Logger.log("URL's nuevos")
    Logger.log(url_empleo_y_desempleo)
    Logger.log(url_informalidad)
    Logger.log(url_mercado_juventud)
    Logger.log(url_mercado_segun_sexo)


  // Verificar si ya existen los archivos a descargar en el directorio
  
  let folder = DriveApp.getFilesByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next().getParents().next()
  let listaNombres = folder.getFiles()
  let nombres = []

  while(listaNombres.hasNext()){
    nombres.push(listaNombres.next().getName())
  }
  nombres = nombres.filter(nombre => nombre !== "Análisis Montería")

  Logger.log(nombres)

  if(nombres.length !== 0){

    let command = Browser.msgBox("La carpeta ya contiene archivos ¿Desea reemplazarlos?", Browser.Buttons.YES_NO_CANCEL)
    if(command === "cancel" || command === "no"){
      Logger.log("Orden denegada o cancelada")
      return;
    }
    else{
      for(let i = 0; i < nombres.length; i++){
        folder.getFilesByName(nombres[i]).next().setTrashed(true)
      }
      Logger.log("Archivos reemplazados")
    }
  }
    // Esta parte ejecuta las descargas

 Object.entries(archivos).forEach(function([nombreArchivo,solicitud]){
      let archivo = solicitud.getBlob()
      archivo.setName(nombreArchivo)
      folder.createFile(archivo)
      }
    )  

}