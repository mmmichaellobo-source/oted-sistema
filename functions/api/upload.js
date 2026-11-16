import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

export async function onRequestPost(context) {
  try {
    const formData = await context.request.formData();
    const file = formData.get('excelFile');
    const tipoExcel = formData.get('tipoExcel'); 
    
    if (!file) {
      return new Response(JSON.stringify({ error: 'No se envió archivo' }), { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(sheet, { defval: null, range: 8 }); // Empezamos desde la fila 9 (índice 8) donde están los datos reales

    if (filas.length === 0) {
      return new Response(JSON.stringify({ error: 'El Excel está vacío o no tiene el formato correcto' }), { status: 400 });
    }

    const supabase = createClient(context.env.SUPABASE_URL, context.env.SUPABASE_SERVICE_ROLE_KEY);
    let tablaDestino = '';
    let datosParaGuardar = [];

    if (tipoExcel === 'consolidado') {
      tablaDestino = 'consolidado_repuestos';
      // El código del consolidado de 41 columnas (mantenlo igual al que ya tenías)
      datosParaGuardar = filas.map(fila => {
        const codigoPersona = String(fila['PERSONA_LINEA_PRODUCCION'] || '');
        return {
          "ID_ALMACEN": fila["ID_ALMACEN"], "NOMBRE_ALMACEN": fila["NOMBRE_ALMACEN"], "FECHA": fila["FECHA"] ? new Date(fila["FECHA"]).toISOString().split('T')[0] : null,
          "TIPO_MOVIMIENTO": fila["TIPO_MOVIMIENTO"], "GESTION": fila["GESTION"], "TIPO_DOCTO": fila["TIPO_DOCTO"], "NOMBRE_DOCTO": fila["NOMBRE_DOCTO"],
          "NUMERO": fila["NUMERO"], "ITEM": fila["ITEM"], "ID_LINEA": fila["ID_LINEA"], "NOMBRE_LINEA": fila["NOMBRE_LINEA"],
          "CODIGO_PRODUCTO": fila["CODIGO_PRODUCTO"], "NOMBRE_PRODUCTO": fila["NOMBRE_PRODUCTO"], "CANTIDAD": fila["CANTIDAD"],
          "UNIDAD_INVENTARIO": fila["UNIDAD_INVENTARIO"], "CANTIDAD_2DA": fila["CANTIDAD_2DA"], "UNIDAD_2DA": fila["UNIDAD_2DA"],
          "COSTO": fila["COSTO"], "PERSONA_LINEA_PRODUCCION": codigoPersona, "NOMBRE_DESCRIPCION": fila["NOMBRE_DESCRIPCION"],
          "DOC_FTE_GRUPO_CA": fila["DOC_FTE_GRUPO_CA"], "NRO_DOC_FTE_CA": fila["NRO_DOC_FTE_CA"], "REGIONAL_UNIDAD_NEGOCIO": fila["REGIONAL_UNIDAD_NEGOCIO"],
          "CCO_ALMACEN": fila["CCO_ALMACEN"], "NOMBRE_CCO_ALMACEN": fila["NOMBRE_CCO_ALMACEN"], "CODIGO_CUENTA": fila["CODIGO_CUENTA"],
          "NOMBRE_CUENTA": fila["NOMBRE_CUENTA"], "CODIGO_SKU": fila["CODIGO_SKU"], "CODIGO_BARRAS": fila["CODIGO_BARRAS"],
          "PRECIO_BS": fila["PRECIO_BS"], "PRECIO_US": fila["PRECIO_US"], "ID_DIVISION": fila["ID_DIVISION"], "NOMBRE_DIVISION": fila["NOMBRE_DIVISION"],
          "ID_MARKETING": fila["ID_MARKETING"], "NOMBRE_MARKETING": fila["NOMBRE_MARKETING"], "ID_VENDEDOR": fila["ID_VENDEDOR"], "VENDEDOR": fila["VENDEDOR"],
          "FORMA_PAGO": fila["FORMA_PAGO"], "VENCIMIENTO": fila["VENCIMIENTO"], "OBSERVACION": fila["OBSERVACION"],
          // El cruce con vendedores se mantiene
          "DIVISION_MAESTRA": 'PENDIENTE', "SEMANA": "Semana Actual"
        };
      });

    } else if (tipoExcel === 'caja_chica') {
      tablaDestino = 'caja_chica';
      datosParaGuardar = filas.map(fila => {
        // 1. Obtenemos la unidad de negocio
        const unidadNegocio = String(fila["UNIDAD NEGOCIO"] || '').toUpperCase();
        
        // 2. Aplicamos la regla: Si NO es POST VENTA, es FABRICA
        let clasificacionFinal = "FABRICA";
        if (unidadNegocio.includes("POST VENTA")) {
          clasificacionFinal = "POST VENTA";
        }

        return {
          "N": String(fila["N"] || ''),
          "FECHA": fila["FECHA"] ? new Date(fila["FECHA"]).toISOString().split('T')[0] : null,
          "UNIDAD_NEGOCIO": clasificacionFinal, // Guardamos la clasificación ya procesada
          "REPUESTO_SERVICIO": fila["REPUESTO SERVICIO"] || fila["REPUESTO_SERVICIO"],
          "PROVEEDOR": fila["PROVEEDOR"],
          "CONCEPTO": fila["CONCEPTO"],
          "IMPORTE": fila["IMPORTE"]
        };
      });
    }

    const { error: errorInsert } = await supabase.from(tablaDestino).insert(datosParaGuardar);
    if (errorInsert) throw new Error("Error al guardar: " + errorInsert.message);

    return new Response(JSON.stringify({ 
      success: true, 
      mensaje: `Se procesaron ${filas.length} filas en la tabla ${tablaDestino}.` 
    }), { headers: { 'Content-Type': 'application/json' } });

  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), { status: 500 });
  }
}
