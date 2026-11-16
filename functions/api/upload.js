import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

export async function onRequestPost(context) {
  try {
    const formData = await context.request.formData();
    const file = formData.get('excelFile');
    
    if (!file) {
      return new Response(JSON.stringify({ error: 'No se envió archivo' }), { status: 400 });
    }

    // 1. Leer el Excel
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Convertir a JSON (array de objetos)
    const filas = XLSX.utils.sheet_to_json(sheet, { defval: null });

    if (filas.length === 0) {
      return new Response(JSON.stringify({ error: 'El Excel está vacío' }), { status: 400 });
    }

    // 2. Conectar a Supabase
    const supabase = createClient(context.env.SUPABASE_URL, context.env.SUPABASE_SERVICE_ROLE_KEY);

    // 3. Traer la tabla de divisiones (Maestra)
    const { data: divisiones, error: errorDiv } = await supabase.from('vendedores_division').select('*');
    if (errorDiv) throw new Error("Error al leer divisiones: " + errorDiv.message);

    // Crear un diccionario rápido para buscar el código -> división
    const mapaDivisiones = {};
    divisiones.forEach(d => {
      mapaDivisiones[String(d.PERSONA_LINEA_PRODUCCION)] = d.DIVISION;
    });

    // 4. Preparar los datos para guardar
    const datosParaGuardar = filas.map(fila => {
      const codigoPersona = String(fila['PERSONA_LINEA_PRODUCCION'] || '');
      const divisionEncontrada = mapaDivisiones[codigoPersona] || 'SIN ASIGNAR';

      // Mapeamos las columnas del Excel (si alguna no existe, se guarda null)
      return {
        "ID_ALMACEN": fila["ID_ALMACEN"],
        "NOMBRE_ALMACEN": fila["NOMBRE_ALMACEN"],
        "FECHA": fila["FECHA"] ? new Date(fila["FECHA"]).toISOString().split('T')[0] : null,
        "TIPO_MOVIMIENTO": fila["TIPO_MOVIMIENTO"],
        "GESTION": fila["GESTION"],
        "TIPO_DOCTO": fila["TIPO_DOCTO"],
        "NOMBRE_DOCTO": fila["NOMBRE_DOCTO"],
        "NUMERO": fila["NUMERO"],
        "ITEM": fila["ITEM"],
        "ID_LINEA": fila["ID_LINEA"],
        "NOMBRE_LINEA": fila["NOMBRE_LINEA"],
        "CODIGO_PRODUCTO": fila["CODIGO_PRODUCTO"],
        "NOMBRE_PRODUCTO": fila["NOMBRE_PRODUCTO"],
        "CANTIDAD": fila["CANTIDAD"],
        "UNIDAD_INVENTARIO": fila["UNIDAD_INVENTARIO"],
        "CANTIDAD_2DA": fila["CANTIDAD_2DA"],
        "UNIDAD_2DA": fila["UNIDAD_2DA"],
        "COSTO": fila["COSTO"],
        "PERSONA_LINEA_PRODUCCION": codigoPersona,
        "NOMBRE_DESCRIPCION": fila["NOMBRE_DESCRIPCION"],
        "DOC_FTE_GRUPO_CA": fila["DOC_FTE_GRUPO_CA"],
        "NRO_DOC_FTE_CA": fila["NRO_DOC_FTE_CA"],
        "REGIONAL_UNIDAD_NEGOCIO": fila["REGIONAL_UNIDAD_NEGOCIO"],
        "CCO_ALMACEN": fila["CCO_ALMACEN"],
        "NOMBRE_CCO_ALMACEN": fila["NOMBRE_CCO_ALMACEN"],
        "CODIGO_CUENTA": fila["CODIGO_CUENTA"],
        "NOMBRE_CUENTA": fila["NOMBRE_CUENTA"],
        "CODIGO_SKU": fila["CODIGO_SKU"],
        "CODIGO_BARRAS": fila["CODIGO_BARRAS"],
        "PRECIO_BS": fila["PRECIO_BS"],
        "PRECIO_US": fila["PRECIO_US"],
        "ID_DIVISION": fila["ID_DIVISION"],
        "NOMBRE_DIVISION": fila["NOMBRE_DIVISION"],
        "ID_MARKETING": fila["ID_MARKETING"],
        "NOMBRE_MARKETING": fila["NOMBRE_MARKETING"],
        "ID_VENDEDOR": fila["ID_VENDEDOR"],
        "VENDEDOR": fila["VENDEDOR"],
        "FORMA_PAGO": fila["FORMA_PAGO"],
        "VENCIMIENTO": fila["VENCIMIENTO"],
        "OBSERVACION": fila["OBSERVACION"],
        "DIVISION_MAESTRA": divisionEncontrada, // AQUÍ INYECTAMOS LA DIVISIÓN CRUZADA
        "SEMANA": "Semana Actual" // Esto lo podemos hacer dinámico luego
      };
    });

    // 5. Guardar en Supabase
    const { error: errorInsert } = await supabase.from('notas_entrega').insert(datosParaGuardar);
    if (errorInsert) throw new Error("Error al guardar en base de datos: " + errorInsert.message);

    return new Response(JSON.stringify({ 
      success: true,
      mensaje: `Se procesaron ${filas.length} filas correctamente.`
    }), {
      headers: { 'Content-Type': 'application/json' }
    });

  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), { status: 500 });
  }
}
