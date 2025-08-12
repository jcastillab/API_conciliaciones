from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
from io import BytesIO

from procesar_pdf import procesar_pdf
from unir_archivos import conciliar_movimientos

app = FastAPI()

@app.post("/conciliacion-unificada/")
async def conciliacion_unificada(
    pdf_file: UploadFile = File(...),
    contabilidad_file: UploadFile = File(...)
):
    try:
        # Procesar PDF directamente desde UploadFile (ya definido en procesar_pdf)
        df_extracto = procesar_pdf(pdf_file)

        # Leer Excel contabilidad directamente desde UploadFile en memoria
        contabilidad_bytes = await contabilidad_file.read()
        df_contabilidad = pd.read_excel(BytesIO(contabilidad_bytes))

        # Llamar función conciliación que retorna bytes del Excel final
        excel_bytes = conciliar_movimientos(df_contabilidad, df_extracto)

        # Preparar respuesta streaming del archivo excel
        return StreamingResponse(
            BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Conciliacion_bancaria.xlsx"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error en proceso unificado: {str(e)}")
