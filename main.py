from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from openpyxl import load_workbook
from datetime import datetime
import io, base64, os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Modelo Excel em base64 (embutido no código)
MODELO_B64 = os.environ.get("MODELO_INSITU_B64", "")

class DadosEnsaio(BaseModel):
    amostra: Optional[str] = ""
    camada: Optional[str] = ""
    tipoMaterial: Optional[str] = ""
    densLab: Optional[float] = 0
    umidadeOtima: Optional[float] = 0
    metodoUmidade: Optional[str] = "direto"
    frigideiraUmido: Optional[float] = 0
    frigideiraSeco: Optional[float] = 0
    tara: Optional[float] = 0
    frascoAntes: Optional[float] = 0
    frascoDepois: Optional[float] = 0
    funil: Optional[float] = 0
    densAreia: Optional[float] = 0
    amostraUmida: Optional[float] = 0
    ident: Optional[str] = ""
    cliente: Optional[str] = ""
    construtora: Optional[str] = ""
    data: Optional[str] = ""

class PayloadExcel(BaseModel):
    lista_ensaios: List[DadosEnsaio]
    relatorio_obs: Optional[str] = ""

@app.get("/")
def health():
    return {"status": "ok", "servico": "LabCQ Excel API"}

@app.post("/gerar-excel")
def gerar_excel(payload: PayloadExcel):
    COLUNAS = ['C','D','E','F','G','H','I','J','K']

    # Carrega modelo
    modelo_bytes = base64.b64decode(MODELO_B64)
    wb = load_workbook(io.BytesIO(modelo_bytes))
    ws = wb.worksheets[0]

    def set_cell(addr, value):
        for merge in ws.merged_cells.ranges:
            if addr in merge:
                ws.cell(row=merge.min_row, column=merge.min_col).value = value
                return
        ws[addr].value = value

    lista = payload.lista_ensaios
    if not lista:
        return {"erro": "Nenhum ensaio fornecido"}

    d0 = lista[0]

    # Cabeçalho
    ws['B6'].value = f"OBRA: {d0.ident or ''}"
    ws['C7'].value = d0.cliente or ''
    ws['B8'].value = f"CONSTRUTORA: {d0.construtora or ''}"
    ws['B9'].value = f"TRECHO: {d0.amostra or ''}"
    ws['B10'].value = f"CAMADA: {d0.camada or ''}"
    ws['B11'].value = f"LOCAL: {d0.ident or ''}"

    # Data emissão
    if d0.data:
        try:
            ws['J4'].value = datetime.strptime(d0.data, '%Y-%m-%d')
        except:
            ws['J4'].value = d0.data

    # OBS do Claude
    if payload.relatorio_obs:
        ws['C51'].value = payload.relatorio_obs

    # Preenche cada determinação
    for idx, d in enumerate(lista[:9]):
        c = COLUNAS[idx]
        num = idx + 1

        ws[f'{c}14'].value = num
        if d.data:
            try:
                ws[f'{c}15'].value = datetime.strptime(d.data, '%Y-%m-%d')
            except:
                ws[f'{c}15'].value = d.data

        ws[f'{c}16'].value = d.amostra or ''
        ws[f'{c}17'].value = 1
        ws[f'{c}18'].value = d.camada or ''
        ws[f'{c}19'].value = d.tipoMaterial or ''
        ws[f'{c}20'].value = 0.2

        # Densidade lab — inteiro sem decimais
        dens_cell = ws[f'{c}22']
        dens_cell.value = int(round(d.densLab or 0))
        dens_cell.number_format = '0'

        ws[f'{c}23'].value = d.umidadeOtima or 0

        # Umidade
        if d.metodoUmidade == 'frigideira':
            ws[f'{c}25'].value = d.frigideiraUmido or 0
            ws[f'{c}26'].value = d.frigideiraSeco or 0
            ws[f'{c}27'].value = d.tara or 0

        # Fórmulas umidade
        ws[f'{c}28'].value = f'={c}25-{c}27'
        ws[f'{c}29'].value = f'={c}26-{c}27'
        ws[f'{c}30'].value = f'={c}28-{c}29'
        ws[f'{c}31'].value = f'=IF({c}29=0,0,{c}30/{c}29*100)'

        # Execução
        ws[f'{c}33'].value = d.frascoAntes or 0
        ws[f'{c}34'].value = d.frascoDepois or 0
        ws[f'{c}35'].value = f'={c}33-{c}34'
        ws[f'{c}36'].value = d.funil or 0
        ws[f'{c}37'].value = f'={c}35-{c}36'
        ws[f'{c}38'].value = d.densAreia or 0
        ws[f'{c}39'].value = f'=IF({c}38=0,0,{c}37/{c}38)'
        ws[f'{c}40'].value = d.amostraUmida or 0
        ws[f'{c}41'].value = 0
        ws[f'{c}42'].value = 0

        # Fórmulas resultado
        ws[f'{c}43'].value = f'=IF({c}41=0,0,{c}41/{c}42)'
        ws[f'{c}44'].value = f'=IF({c}41=0,{c}40,{c}40-{c}41)'
        ws[f'{c}45'].value = f'=IF({c}41=0,{c}39,{c}39-{c}43)'
        ws[f'{c}46'].value = f'=IF({c}45=0,0,{c}44/{c}45)'
        ws[f'{c}47'].value = f'=IF({c}31=0,{c}46,{c}46/(1+({c}31/100)))'
        ws[f'{c}48'].value = f'=IF({c}22=0,0,({c}47/{c}22)*100)'
        ws[f'{c}49'].value = f'={c}31-{c}23'

    # Retorna arquivo
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    nome = f"InSitu_{(d0.amostra or d0.ident or 'ensaio').replace(' ','_')}.xlsx"
    return StreamingResponse(
        buf,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{nome}"'}
    )
