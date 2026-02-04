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


  // Directorio de descarga y sus archivos
    let folder = DriveApp.getFilesByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next().getParents().next()
    let listaArchivos = folder.getFiles()
    let nombresArchivos = []

    while(listaArchivos.hasNext()){
      nombresArchivos.push(listaArchivos.next().getName())
    }
    nombresArchivos = nombresArchivos.filter(nombre => nombre !== "Análisis Montería")
    
    Logger.log(nombresArchivos)

  // Clase para gestión de descargas

    class urlClass{
      constructor(url, name){
        this.url = url
        this.name = name
        this.folder = DriveApp.getFilesByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next().getParents().next()
        this.request = UrlFetchApp.fetch(this.url, {muteHttpExceptions:true, method:"get"})
      }

      download(){
        let archivo = this.request.getBlob()
        archivo.setName(this.name)
        this.folder.createFile(archivo)
      }

      responseCode(){
        return this.request.getResponseCode()
      }
    }

  // Descargando datos de la ciudad
    let empleo_y_desempleo = new urlClass(url=url_empleo_y_desempleo, name="Empleo y desempleo.xlsx")

    if(empleo_y_desempleo.responseCode() == 200 && nombresArchivos.includes("Empleo y desempleo.xlsx")){
      let choice = Browser.msgBox("El archivo '"+empleo_y_desempleo.name+"' ya se encuentra en el directorio ¿Desea reemplazarlo?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
        
      }
      else{
        folder.getFilesByName(empleo_y_desempleo.name).next().setTrashed(true)
        empleo_y_desempleo.download()
        Browser.msgBox("Archivo '"+empleo_y_desempleo.name+"' descargado en: "+ empleo_y_desempleo.folder.getName())
      }
    }

    if(empleo_y_desempleo.responseCode() == 200 && !nombresArchivos.includes("Empleo y desempleo.xlsx")){
      Browser.msgBox("Descargando archivo '"+empleo_y_desempleo.name+"' de: " +url_empleo_y_desempleo)
      empleo_y_desempleo.download()
      Browser.msgBox("Archivo '"+empleo_y_desempleo.name+"' descargado en: "+ empleo_y_desempleo.folder.getName())
    }

    if(empleo_y_desempleo.responseCode() !== 200 && nombresArchivos.includes("Empleo y desempleo.xlsx")){
      let choice = Browser.msgBox("El archivo '"+url_empleo_y_desempleo+"' aún no está disponible ¿Descargar la última versión?", Browser.Buttons.YES_NO)
      if(choice === "no"){

      }
      else{
        if (month-3<0){
          ref = (month-2)*-1
          previous_month = meses.reverse()[ref]
          meses.reverse()

          year = year - 1 
        }
    
        else{
          previous_month = meses[month-3]
        }
        folder.getFilesByName(empleo_y_desempleo.name).next().setTrashed(true)
        url_empleo_y_desempleo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIH-"+previous_month+year+".xlsx"
        Logger.log("Nueva url empleo y desempleo: %s", url_empleo_y_desempleo)
        new urlClass(url=url_empleo_y_desempleo, name="Empleo y desempleo.xlsx").download()
        Browser.msgBox("Archivo '"+empleo_y_desempleo.name+"' descargado en: "+ empleo_y_desempleo.folder.getName())
      }
    }
    
    if(empleo_y_desempleo.responseCode() !== 200 && !nombresArchivos.includes("Empleo y desempleo.xlsx")){
      Browser.msgBox("Descargando la última versión disponible del archivo '"+empleo_y_desempleo.name+"'")
        if (month-3<0){
          ref = (month-2)*-1
          previous_month = meses.reverse()[ref]
          meses.reverse()
          
          year = year - 1
        }
        else{
          previous_month = meses[month-3]
        }
      url_empleo_y_desempleo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIH-"+previous_month+year+".xlsx"
      Logger.log("Nueva url empleo y desempleo: %s", url_empleo_y_desempleo)
      new urlClass(url=url_empleo_y_desempleo, name="Empleo y desempleo.xlsx").download()
      Browser.msgBox("Archivo '"+empleo_y_desempleo.name+"' descargado en: "+ empleo_y_desempleo.folder.getName())
    }

  // Descargando datos de informalidad
    let informalidad = new urlClass(url=url_informalidad, name="Empleo informal y seguridad social.xlsx")

    if(informalidad.responseCode() == 200 && nombresArchivos.includes("Empleo informal y seguridad social.xlsx")){
      let choice = Browser.msgBox("El archivo '"+informalidad.name+"' ya se encuentra en el directorio ¿Desea reemplazarlo?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
        
      }
      else{
        folder.getFilesByName(informalidad.name).next().setTrashed(true)
        informalidad.download()
        Browser.msgBox("Archivo '"+informalidad.name+ "' descargado en: "+ informalidad.folder.getName())
      }
    }

    if(informalidad.responseCode() == 200 && !nombresArchivos.includes("Empleo informal y seguridad social.xlsx")){
      Browser.msgBox("Descargando archivo '"+informalidad.name+"' de: " +url_informalidad)
      informalidad.download()
      Browser.msgBox("Archivo '"+informalidad.name +"' descargado en: "+ informalidad.folder.getName())
    }  

    if(informalidad.responseCode() !== 200 && nombresArchivos.includes("Empleo informal y seguridad social.xlsx")){
      let choice = Browser.msgBox("El archivo '"+url_informalidad+"' aún no está disponible ¿Descargar la última versión?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
    
      }
      else{ // Modificar lógica del trimestre movil
        if(month-5<=0){

          let ref1 = (month-5)*-1
          initial_month = meses.reverse()[ref1]
          meses.reverse()
        }

        else{
          initial_month = meses[month-6]
        }

        // mes final
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
        folder.getFilesByName(informalidad.name).next().setTrashed(true)
        url_informalidad = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHEISS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
        Logger.log("Nueva url empleo informal y seguridad social: %s", url_informalidad)
        new urlClass(url=url_informalidad, name="Empleo informal y seguridad social.xlsx").download()
        Browser.msgBox("Archivo '"+informalidad.name+"' descargado en: "+ informalidad.folder.getName())
      }
    }

    if(informalidad.responseCode() !== 200 && !nombresArchivos.includes("Empleo informal y seguridad social.xlsx")){
      Browser.msgBox("Descargando la última versión disponible del archivo '"+informalidad.name+"'")
        if(month-5<=0){
          let ref1 = (month-5)*-1
          initial_month = meses.reverse()[ref1]
          meses.reverse()
        }

        else{
          initial_month = meses[month-6]
        }

        // mes final
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

      url_informalidad = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHEISS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
      Logger.log("Nueva url empleo informal y seguridad social: %s", url_informalidad)
      new urlClass(url=url_informalidad, name="Empleo informal y seguridad social.xlsx").download()
      Browser.msgBox("Archivo '"+informalidad.name+"' descargado en: "+ informalidad.folder.getName())
    }

  // Descargando datos del mercado según sexo
    let mercado_segun_sexo = new urlClass(url=url_mercado_segun_sexo, name="Mercado laboral según sexo.xlsx")

    if(mercado_segun_sexo.responseCode == 200 && nombresArchivos.includes("Mercado laboral según sexo.xlsx")){
      let choice = Browser.msgBox("El archivo '"+mercado_segun_sexo.name+"' ya se encuentra en el directorio ¿Desea reemplazarlo?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
        
      }
      else{
        folder.getFilesByName(mercado_segun_sexo.name).next().setTrashed(true)
        mercado_segun_sexo.download()
        Browser.msgBox("Archivo '"+mercado_segun_sexo.name+ "' descargado en: "+ mercado_segun_sexo.folder.getName())
      }
    }

    if(mercado_segun_sexo.responseCode == 200 && !nombresArchivos.includes("Mercado laboral según sexo.xlsx")){
      Browser.msgBox("Descargando archivo '"+mercado_segun_sexo.name+"' de: " +url_mercado_segun_sexo)
      mercado_segun_sexo.download()
      Browser.msgBox("Archivo '"+mercado_segun_sexo.name +"' descargado en: "+ mercado_segun_sexo.folder.getName())
    }

    if(mercado_segun_sexo.responseCode !== 200 && nombresArchivos.includes("Mercado laboral según sexo.xlsx")){
      let choice = Browser.msgBox("El archivo '"+url_mercado_segun_sexo+"' aún no está disponible ¿Descargar la última versión?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
    
      }
      else{ // Modificar lógica del trimestre movil
        if(month-5<=0){

          let ref1 = (month-5)*-1
          initial_month = meses.reverse()[ref1]
          meses.reverse()
        }

        else{
          initial_month = meses[month-6]
        }

        // mes final
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
        folder.getFilesByName(mercado_segun_sexo.name).next().setTrashed(true)
        url_mercado_segun_sexo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
        Logger.log("Nueva url Mercado laboral según sexo: %s", url_mercado_segun_sexo)
        new urlClass(url=url_mercado_segun_sexo, name="Mercado laboral según sexo.xlsx").download()
        Browser.msgBox("Archivo '"+mercado_segun_sexo.name+"' descargado en: "+ mercado_segun_sexo.folder.getName())
      }
    }

    if(mercado_segun_sexo.responseCode !== 200 && !nombresArchivos.includes("Mercado laboral según sexo.xlsx")){
        Browser.msgBox("Descargando la última versión disponible del archivo '"+mercado_segun_sexo.name+"'")
        if(month-5<=0){
          let ref1 = (month-5)*-1
          initial_month = meses.reverse()[ref1]
          meses.reverse()
        }

        else{
          initial_month = meses[month-6]
        }

        // mes final
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

      url_mercado_segun_sexo = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLS-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
      Logger.log("Nueva url Mercado laboral según sexo: %s", url_mercado_segun_sexo)
      new urlClass(url=url_mercado_segun_sexo, name="Mercado laboral según sexo.xlsx").download()
      Browser.msgBox("Archivo '"+mercado_segun_sexo.name+"' descargado en: "+ mercado_segun_sexo.folder.getName())
    }

  // Descargar datos del mercado de la juventud

    let mercado_juventud = new urlClass(url=url_mercado_juventud, name="Mercado laboral de la juventud.xlsx")

    if(mercado_juventud.responseCode() == 200 && nombresArchivos.includes("Mercado laboral de la juventud.xlsx")){
      let choice = Browser.msgBox("El archivo '"+mercado_juventud.name+"' ya se encuentra en el directorio ¿Desea reemplazarlo?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
        
      }
      else{
        folder.getFilesByName(mercado_juventud.name).next().setTrashed(true)
        mercado_juventud.download()
        Browser.msgBox("Archivo '"+mercado_juventud.name+ "' descargado en: "+ mercado_juventud.folder.getName())
      }
    }

    if(mercado_juventud.responseCode() == 200 && !nombresArchivos.includes("Mercado laboral de la juventud.xlsx")){
      Browser.msgBox("Descargando archivo '"+mercado_juventud.name+"' de: " +url_mercado_juventud)
      mercado_juventud.download()
      Browser.msgBox("Archivo '"+mercado_juventud.name +"' descargado en: "+ mercado_juventud.folder.getName())
    }

    if(mercado_juventud.responseCode() !== 200 && nombresArchivos.includes("Mercado laboral de la juventud.xlsx")){
      let choice = Browser.msgBox("El archivo '"+url_mercado_juventud+"' aún no está disponible ¿Descargar la última versión?",
      Browser.Buttons.YES_NO)
      if(choice === "no"){
    
      }
      else{ // Modificar lógica del trimestre movil
        if(month-5<=0){

          let ref1 = (month-5)*-1
          initial_month = meses.reverse()[ref1]
          meses.reverse()
        }

        else{
          initial_month = meses[month-6]
        }

        // mes final
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
        folder.getFilesByName(mercado_juventud.name).next().setTrashed(true)
        url_mercado_juventud = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLJ-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
        Logger.log("Nueva url Mercado laboral de la juventud: %s", url_mercado_segun_sexo)
        new urlClass(url=url_mercado_juventud, name="Mercado laboral de la juventud.xlsx").download()
        Browser.msgBox("Archivo '"+mercado_juventud.name+"' descargado en: "+ mercado_juventud.folder.getName())
      }
    }

    if(mercado_juventud.responseCode() !== 200 && !nombresArchivos.includes("Mercado laboral de la juventud.xlsx")){
      Browser.msgBox("Descargando la última versión disponible del archivo '"+mercado_juventud.name+"'")
      if(month-5<=0){
        let ref1 = (month-5)*-1
        initial_month = meses.reverse()[ref1]
        meses.reverse()
      }

      else{
        initial_month = meses[month-6]
      }

      // mes final
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

      url_mercado_juventud = "https://www.dane.gov.co/files/operaciones/GEIH/anex-GEIHMLJ-"+initial_month+initial_year+"-"+final_month+final_year+".xlsx"
      Logger.log("Nueva url Mercado laboral de la juventud: %s", url_mercado_juventud)
      new urlClass(url=url_mercado_juventud, name="Mercado laboral de la juventud.xlsx").download()
      Browser.msgBox("Archivo '"+mercado_juventud.name+"' descargado en: "+ mercado_juventud.folder.getName())
    }

  // Realizando nuevamente la lectura de los archivos en el directorio
  
    if(nombresArchivos.length == 0){
      listaArchivos = folder.getFiles()

      while(listaArchivos.hasNext()){
        nombresArchivos.push(listaArchivos.next().getName())
      }
      
      nombresArchivos = nombresArchivos.filter(nombre => nombre !== "Análisis Montería")
    }

  // Convertir los archivos Excel a Google Sheets
    for(let i=0; i<nombresArchivos.length; i++){
      let excelFile = DriveApp.getFilesByName(nombresArchivos[i].toString()).next()
      Logger.log(nombresArchivos[i])
      let blob = excelFile.getBlob()

      let config_api = {
        name: excelFile.getName().replace(/\.xlsx$/i,""),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [folder.getId()]
      }

      try{
        let hojaNueva= Drive.Files.create(config_api, blob, {fields:"id, name, webViewLink"}) //sin el objeto con fields, no puedes ver el link
        Browser.msgBox(`Archivo Excel '`+nombresArchivos[i]+`' convertido a Google Sheets \n: URL: ${hojaNueva.webViewLink}`)
        Logger.log("URL: %s", hojaNueva.webViewLink)
      }
      catch(e){
        Browser.msgBox(`Error al convertir ${e.toString()}`)
        Logger.log("Error al convertir el archivo %s", e.toString())
      }
    }
  // Eliminar los archivos Excel
    for(let i=0; i<nombresArchivos.length; i++){
      if(nombresArchivos[i].match(/\.xlsx$/)){
        folder.getFilesByName(nombresArchivos[i]).next().setTrashed(true)
      }
    }
}