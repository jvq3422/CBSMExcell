"""
CBSMExcell - Backend Empresarial (FastAPI + SQLAlchemy)
Pronto para integração com bancos de dados relacionais corporativos:
PostgreSQL, SQL Server, MySQL, SQLite, Oracle, etc.
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io

app = FastAPI(
    title="CBSMExcell - API Empresarial",
    description="Serviço para ingestão, consolidação e auditoria de planilhas empresariais",
    version="1.0.0"
)

# Habilitar CORS para comunicação com o frontend web
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {
        "status": "online",
        "service": "CBSMExcell API",
        "description": "Pronto para persistência em Banco de Dados Empresarial"
    }

@app.post("/api/planilhas/processar")
async def processar_planilha(file: UploadFile = File(...)):
    """
    Recebe qualquer arquivo Excel (.xlsx, .xls) ou CSV,
    processa suas abas, gera métricas e prepara a inserção no banco.
    """
    contents = await file.read()
    filename = file.filename.lower()
    
    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(contents))
            sheets = {"Dados": df}
        else:
            sheets = pd.read_excel(io.BytesIO(contents), sheet_name=None)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Erro ao processar arquivo Excel: {str(e)}")

    resumo_geral = []
    
    for sheet_name, df_sheet in sheets.items():
        # Limpeza e identificação básica
        colunas = list(df_sheet.columns)
        total_linhas = len(df_sheet)
        
        # Identificação de colunas financeiras
        colunas_numericas = df_sheet.select_dtypes(include=['number']).columns.tolist()
        soma_total = float(df_sheet[colunas_numericas[0]].sum()) if colunas_numericas else 0.0

        # Aqui pode ser adicionada a inserção direta no banco:
        # df_sheet.to_sql("historico_gastos", con=database_engine, if_exists="append", index=False)

        resumo_geral.append({
            "aba": sheet_name,
            "total_linhas": total_linhas,
            "colunas": colunas,
            "soma_estimada": round(soma_total, 2)
        })

    return {
        "arquivo": file.filename,
        "total_abas": len(sheets),
        "detalhes": resumo_geral
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("backend_exemplo:app", host="127.0.0.1", port=8000, reload=True)
