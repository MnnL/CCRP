var nombre_hoja = "Completar"
var columnas_hoja = null
var letra_primera_columna_encabezado = "A"
var letra_ultima_columna_encabezado = "K" //Columna con el dato de respuesta
var indice_columna_inspeccionar = 9 //respuesta
var letra_columna_inspeccionar = "J" //respuesta
var indice_primera_fila_encabezado = 1
var cantidad_filas_cargar_por_ronda = 30 // filas x peticion 
var delimitador_csv = ";"


function traer_indice_columna_por_nombre(nombre_columna){
    let respuesta = columnas_hoja.indexOf(nombre_columna)
    // if(respuesta == -1){
    //     throw `${nombre_columna} no encontrado en la hoja ${nombre_hoja}`
    // }
    return respuesta;
}

async function traer_casos_no_procesados_ultimoRango(ultimo_indice){
    let resultado = []
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        let ultimo_indice_encontrado = false
        ultimo_indice = Number(ultimo_indice)
        let rango_encabezados = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`)
                                    .getUsedRange()
                                    .load("values")
        await context.sync();
        //cambiar para que traiga solo hasta la columna de corte
        resultado.push(["indice_excel"].concat(rango_encabezados.values[0]))      
        while(ultimo_indice_encontrado == false){
            indice_inferior = ultimo_indice + cantidad_filas_cargar_por_ronda 
            
            let indice_rangos = `${nombre_hoja}!${letra_primera_columna_encabezado}${ultimo_indice}:${letra_ultima_columna_encabezado}${indice_inferior}`
            
            let rangos = hoja.getRange(indice_rangos).load("values")
            await context.sync()

            let valores_columna = rangos.values.entries()

            for(var [indice,fila] of valores_columna){
                if(fila[0] == ""){
                    ultimo_indice_encontrado = true 
                    break;
                }
                let index_fila = ultimo_indice + indice
                let indice_fila = `${letra_ultima_columna_encabezado}${index_fila}` //E1 formato
                resultado.push([indice_fila].concat(fila))
            }
            
            ultimo_indice = indice_inferior
        }
    });
    return resultado;
}

//Trae todas las filas hasta que el valor de la fila en la columna indice_columna_inspeccionar no este vacio
//La verificacion se hace a partir de la columna letra_primera_columna_encabezado, si los datos se mueven a los lados el codigo se rompe
async function traer_casos_no_procesados(){
    let resultado = []
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        let ultimo_indice_encontrado = false
        //Cambiar a dinamico, para encontrar indice de columna necesario no escribirlo directamente
        
        //Devuelve el indice de excel del ultimo valor en la columna
        let indice_ultimo_rango = hoja.getRange(`${letra_primera_columna_encabezado}:${letra_primera_columna_encabezado}`)
                                    .getUsedRange()
                                    .getLastRow()
                                    .load("address")
        
        
        let rango_encabezados = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`)
                                    .getUsedRange()
                                    .load("values")
        
        await context.sync();
        //cambiar para que traiga solo hasta la columna de corte
        resultado.push(["indice_excel"].concat(rango_encabezados.values[0]))


        let indice_superior = /[!][A-Z]+(\d+)/.exec(indice_ultimo_rango.address)[1] 
        indice_superior = Number(indice_superior)
        let indice_inferior = 0

        while(ultimo_indice_encontrado == false){
            indice_inferior = indice_superior - cantidad_filas_cargar_por_ronda 
            indice_inferior = indice_inferior < 1 ? 1 : indice_inferior //evitar que baje del minimo
            let indice_rangos = `${nombre_hoja}!${letra_primera_columna_encabezado}${indice_superior}:${letra_ultima_columna_encabezado}${indice_inferior}`
            
            let rangos = hoja.getRange(indice_rangos).load("values")
            await context.sync()

            let valores_columna_invertida = rangos.values.reverse().entries()

            for(var [indice,fila] of valores_columna_invertida){
                if(fila[indice_columna_inspeccionar] != ""){
                    ultimo_indice_encontrado = true 
                    break;
                }
                let index_fila = indice_superior - indice
                let indice_fila = `${letra_columna_inspeccionar}${index_fila}` //E1 formato
                resultado.push([indice_fila].concat(fila))
            }
            
            indice_superior = indice_inferior
        }
    });
    return resultado;
}

function resultado_a_csv(resultado){
    return resultado.map(e => e.join(delimitador_csv)).join("\r\t")
    //return resultado.map(e => e.join(delimitador_csv)).join("\r\n")
}

async function pegar_csv(texto_csv){
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem(nombre_hoja);
        let valores_fila = texto_csv.split('\n').slice(1) //Quitar los encabezados
        for(let fila of valores_fila){
            fila = fila.split(delimitador_csv)
            let range = sheet.getRange(fila[0])
            range.values = fila[indice_columna_inspeccionar+1]//Quitar el indice_excel 
        }
        
        await context.sync();
    });

}
//DEBUG
//traer_casos_no_procesados().then(e => console.log(resultado_a_csv(e)))

Office.onReady(async (info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();

        columnas_hoja = await Excel.run(async (context) => {
            let hoja = context.workbook.worksheets.getItem(nombre_hoja);
            let temp_indice_columnas = hoja.getRange(`A1:Z1`)
                                        .load("values")
            await context.sync();
            return temp_indice_columnas.values[0] //Se retorna el 0 por temas de indices, 1:1
        })
    }
});


