from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
from sqlalchemy import create_engine, text
from fpdf import FPDF
import io
import re
import zipfile
from datetime import date
import traceback
from google import genai

# Imports para E-mail com anexo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

app = FastAPI(title="Gestão de Contratos TCE-GO")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 1. CONEXÃO COM O BANCO DE DADOS
DATABASE_URL = "postgresql://neondb_owner:npg_foy20hbmuILp@ep-proud-sea-ay8wz1z6.c-5.us-east-2.aws.neon.tech/neondb?sslmode=require"
engine = create_engine(DATABASE_URL)

# 2. CHAVE DO GEMINI
GEMINI_API_KEY = "AIzaSyAeNilkxFpLu0ExOmgFXD-QA16G5GGAQa4" 

def limpar_valor(valor):
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    v = str(valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
    try: return float(v)
    except: return 0.0

def classificar_com_ia(objeto: str) -> bool:
    if not GEMINI_API_KEY or GEMINI_API_KEY == "AIzaSyAeNilkxFpLu0ExOmgFXD-QA16G5GGAQa4":
        return False
        
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt = f"""
        Você é um auditor e especialista em Licitações Públicas (Lei 14.133/2021).
        Analise a natureza do seguinte objeto de contrato administrativo: "{objeto}"
        
        Sua tarefa é classificar se este contrato é PRORROGÁVEL (serviços contínuos, assinaturas, locações, suportes, manutenção de atividade administrativa - art. 106 e 107) ou NÃO PRORROGÁVEL (escopo fixo, aquisição pontual de bens/equipamentos, obras, entrega com resultado específico).
        
        Responda APENAS com a palavra PRORROGAVEL ou NAO_PRORROGAVEL. Não adicione nenhuma outra palavra, ponto ou explicação.
        """
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        resposta_ia = response.text.strip().upper()
        
        if "NAO_PRORROGAVEL" in resposta_ia or "NÃO_PRORROGAVEL" in resposta_ia:
            return False
        return True
    except Exception:
        return False

def limpar_texto_pdf(texto: str) -> str:
    if not texto: return ""
    texto = str(texto)
    reps = {'–': '-', '—': '-', '“': '"', '”': '"', '‘': "'", '’': "'"}
    for k, v in reps.items(): texto = texto.replace(k, v)
    return texto

def limpar_excel_corrompido(file_bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as zin:
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, 'w') as zout:
            for item in zin.infolist():
                if item.filename != 'xl/styles.xml':
                    zout.writestr(item, zin.read(item.filename))
        out_buf.seek(0)
        return out_buf

@app.post("/api/importar")
async def importar_planilha(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="O arquivo deve ser .xlsx")
    
    try:
        file_bytes = await file.read()
        try:
            clean_buf = limpar_excel_corrompido(file_bytes)
            df = pd.read_excel(clean_buf, engine='openpyxl', header=1)
        except Exception:
            df = pd.read_excel(io.BytesIO(file_bytes), header=1)
            
        df.columns = df.columns.str.strip()
        
        def find_col(possiveis):
            for p in possiveis:
                for col in df.columns:
                    if p.lower() in col.lower(): return col
            return None

        col_contrato = find_col(['CONTRATADO', 'Fornecedor'])
        col_aditivos = find_col(['ADITIVO'])
        col_processo = find_col(['PROCESSO'])
        col_objeto = find_col(['OBJETO'])
        col_inicio = find_col(['INÍCIO', 'Inicio'])
        col_fim = find_col(['TÉRMINO', 'Fim'])
        col_valor_anual = find_col(['VALOR ANUAL', 'Valor'])
        col_gestor = find_col(['GESTOR'])
        col_fiscal = find_col(['FISCAL'])
        col_data_contrato = find_col(['DT CONTRATO'])

        df_banco = pd.DataFrame()
        df_banco['ano_contrato'] = pd.to_datetime(df[col_data_contrato], dayfirst=True, errors='coerce').dt.year.fillna(2026).astype(int) if col_data_contrato else 2026
        df_banco['numero_contrato'] = "S/N" 
        df_banco['qtd_aditivos'] = pd.to_numeric(df[col_aditivos], errors='coerce').fillna(0).astype(int) if col_aditivos else 0
        df_banco['numero_processo'] = df[col_processo].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() if col_processo else "0"
        df_banco['setor_demandante'] = df[col_gestor].astype(str) if col_gestor else "Geral"
        df_banco['objeto_contrato'] = df[col_objeto].astype(str) if col_objeto else ""
        df_banco['fornecedor'] = df[col_contrato].astype(str) if col_contrato else ""
        
        df_banco['inicio_vigencia'] = pd.to_datetime(df[col_inicio], dayfirst=True, errors='coerce') if col_inicio else pd.to_datetime('2026-01-01')
        df_banco['fim_vigencia'] = pd.to_datetime(df[col_fim], dayfirst=True, errors='coerce') if col_fim else pd.to_datetime('2026-12-31')
        df_banco['valor_contrato'] = df[col_valor_anual].apply(limpar_valor) if col_valor_anual else 0.0
        df_banco['gestor'] = df[col_gestor].astype(str) if col_gestor else ""
        df_banco['fiscal'] = df[col_fiscal].astype(str) if col_fiscal else ""

        df_banco = df_banco.dropna(subset=['numero_processo', 'inicio_vigencia', 'fim_vigencia'])
        df_banco = df_banco.drop_duplicates(subset=['numero_processo'], keep='first')

        with engine.begin() as conn:
            conn.execute(text("DROP TABLE IF EXISTS contratos_tce;"))

        df_banco.to_sql('contratos_tce', con=engine, if_exists='replace', index=False)
        return {"mensagem": "Sucesso", "contratos_importados": len(df_banco)}
        
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/contratos")
def listar_contratos():
    try:
        with engine.connect() as conn:
            query = text("SELECT * FROM contratos_tce")
            result = conn.execute(query)
            contratos_processados = []
            hoje = date.today()
            
            for row in result:
                try:
                    fim_dt = pd.to_datetime(row.fim_vigencia)
                    fim = fim_dt.date() if pd.notnull(fim_dt) else None
                except: fim = None
                    
                try:
                    inicio_dt = pd.to_datetime(row.inicio_vigencia)
                    inicio = inicio_dt.date() if pd.notnull(inicio_dt) else None
                except: inicio = None

                dias_restantes = (fim - hoje).days if fim else 0
                
                if dias_restantes < 0: status = "Encerrado"
                elif dias_restantes <= 120: status = "Crítico"
                elif dias_restantes <= 180: status = "Atenção"
                else: status = "No Prazo"
                
                try: valor_bruto = float(row.valor_contrato)
                except: valor_bruto = 0.0
                    
                contratos_processados.append({
                    "ano": int(row.ano_contrato) if row.ano_contrato else 2026,
                    "contrato": str(row.numero_contrato),
                    "aditivos": int(row.qtd_aditivos) if row.qtd_aditivos else 0,
                    "processo": str(row.numero_processo),
                    "objeto": str(row.objeto_contrato),
                    "fornecedor": str(row.fornecedor),
                    "vigencia_inicio": inicio.strftime("%d/%m/%Y") if inicio else "-",
                    "vigencia_fim": fim.strftime("%d/%m/%Y") if fim else "-",
                    "valor_formatado": f"R$ {valor_bruto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    "valor_bruto": valor_bruto,
                    "status": status,
                    "dias": dias_restantes,
                    "gestor": str(row.gestor) if hasattr(row, 'gestor') else "",
                    "fiscal": str(row.fiscal) if hasattr(row, 'fiscal') else ""
                })
            return contratos_processados
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

def gerar_bytes_pdf(processo: str, conn) -> bytes:
    query = text(f"SELECT * FROM contratos_tce WHERE numero_processo = '{processo}' LIMIT 1")
    result = conn.execute(query).fetchone()
    if not result:
        raise ValueError("Contrato não encontrado")

    prorrogavel = classificar_com_ia(result.objeto_contrato)
    
    try:
        inicio_dt = pd.to_datetime(result.inicio_vigencia)
        data_inicio = inicio_dt.strftime('%d/%m/%Y') if pd.notnull(inicio_dt) else "-"
    except: data_inicio = "-"
        
    try:
        fim_dt = pd.to_datetime(result.fim_vigencia)
        data_fim = fim_dt.strftime('%d/%m/%Y') if pd.notnull(fim_dt) else "-"
    except: data_fim = "-"
    
    fornecedor_raw = limpar_texto_pdf(result.fornecedor)
    nome_empresa = re.sub(r'\s*[-–]\s*[\*\d\.]+/[*\d]+-[*\d]+', '', fornecedor_raw).strip()
    if not nome_empresa: nome_empresa = fornecedor_raw
        
    objeto = limpar_texto_pdf(result.objeto_contrato).rstrip('. ')
    setor = limpar_texto_pdf(result.gestor if hasattr(result, 'gestor') else result.setor_demandante).upper()
    
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_margins(left=30, top=25, right=30)
    pdf.set_auto_page_break(auto=True, margin=25)
    pdf.add_page()
    
    h_linha = 5
    pdf.set_font("helvetica", 'B', 11)
    
    pdf.cell(0, h_linha, "MEMORANDO - SERV-CONTRATACOES", ln=True, align='C')
    
    pdf.ln(h_linha * 3) 
    
    pdf.cell(0, h_linha, "DE: SERVIÇO DE CONTRATAÇÕES", ln=True)
    pdf.cell(0, h_linha, f"PARA: {setor}", ln=True)
    pdf.cell(0, h_linha, "ASSUNTO: AVISO DE VENCIMENTO DO CONTRATO", ln=True)
    pdf.ln(h_linha * 2) 

    pdf.set_font("helvetica", '', 11)
    pdf.cell(0, h_linha, "        Prezado(a) Gestor(a),", ln=True)
    pdf.ln(h_linha)

    if prorrogavel:
        texto = (
            f"Informamos que o contrato celebrado com a empresa **{nome_empresa}**, formalizado por meio do processo "
            f"nº **{result.numero_processo}**, terá sua vigência encerrada em **{data_fim}**.\n\n"
            f"Destacamos que o referido Contrato tem por objeto {objeto}, "
            f"com vigência de **{data_inicio}** a **{data_fim}**.\n\n"
            f"Portanto, caso haja interesse na renovação do presente Contrato, solicitamos gentilmente que seja elaborado "
            f"um **Documento de Formalização de Demanda (DFD)**, por meio de Memorando, e encaminhado à **Diretoria de Administração**, "
            f"acompanhado da documentação pertinente, conforme estabelecido na **Ordem de Serviço nº 01/2026 - GPRES**.\n\n"
            f"Tendo em vista o tempo que costuma ser exigido para a conclusão dos trâmites legais relativos à formalização de "
            f"ajustes contratuais, solicita-se que a renovação seja requerida com a maior brevidade possível.\n\n"
            f"Permanecemos à disposição para quaisquer esclarecimentos adicionais que se façam necessários.\n\n"
            f"        Goiânia, data da assinatura eletrônica."
        )
    else:
        texto = (
            f"Informamos que o contrato celebrado com a empresa **{nome_empresa}**, formalizado por meio do processo "
            f"nº **{result.numero_processo}**, terá sua vigência impreterivelmente encerrada em **{data_fim}**.\n\n"
            f"O referido contrato tem por objeto o {objeto}, compreendido entre **{data_inicio}** e **{data_fim}**, "
            f"--não sendo juridicamente possível sua prorrogação.--\n\n"
            f"Assim, a eventual continuidade dos serviços dependerá da realização de nova contratação, observados os procedimentos "
            f"previstos na legislação vigente e nos normativos internos deste Tribunal.\n\n"
            f"Diante disso, solicitamos que essa Diretoria manifeste formalmente seu interesse, ou não, na continuidade dos serviços "
            f"atualmente prestados.\n\n"
            f"Em caso de interesse, deverá ser elaborado o respectivo **Documento de Formalização de Demanda (DFD)** e encaminhado "
            f"à **Diretoria de Administração**, juntamente com a documentação pertinente, conforme disposto na **Ordem de Serviço "
            f"nº 01/2026 - GPRES**, para viabilizar, em tempo hábil, a instauração do novo procedimento de contratação.\n\n"
            f"Ressaltamos que a definição antecipada da demanda é essencial para evitar a descontinuidade dos serviços após o "
            f"término da vigência contratual.\n\n"
            f"Permanecemos à disposição para prestar os esclarecimentos que se fizerem necessários.\n\n"
            f"        Goiânia, data da assinatura eletrônica."
        )

    pdf.multi_cell(0, h_linha, texto, markdown=True)
    return bytes(pdf.output())

@app.get("/api/gerar-memorando/{processo}")
def gerar_memorando(processo: str):
    try:
        with engine.connect() as conn:
            pdf_bytes = gerar_bytes_pdf(processo, conn)
            
        return StreamingResponse(
            io.BytesIO(pdf_bytes), 
            media_type="application/pdf", 
            # Alterado aqui para baixar com nome fixo
            headers={"Content-Disposition": "attachment; filename=Memorando.pdf"}
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/enviar-notificacao/{processo}")
def enviar_notificacao(processo: str):
    try:
        with engine.connect() as conn:
            query = text(f"SELECT fornecedor FROM contratos_tce WHERE numero_processo = '{processo}' LIMIT 1")
            result = conn.execute(query).fetchone()
            if not result:
                raise HTTPException(status_code=404, detail="Contrato não encontrado")
            
            nome_fornecedor = result.fornecedor
            pdf_bytes = gerar_bytes_pdf(processo, conn)

        sender_email = "arturedu22@gmail.com"
        sender_password = "wdxt xixe aqav vmkf"
        receiver_email = "aeduardo@tce.go.gov.br"

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        
        # Alterado aqui para título genérico
        msg['Subject'] = "Aviso de Vencimento Contratual"

        body = f"""À consideração do(a) Gestor(a),

Encaminhamos, em anexo, o Memorando referente ao aviso de vencimento do contrato vinculado ao processo nº {processo} ({nome_fornecedor}).

Solicitamos a gentileza de verificar as providências necessárias, conforme detalhado no documento anexo.

Atenciosamente,

Serviço de Contratações
Tribunal de Contas do Estado de Goiás (TCE-GO)
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
        # Alterado aqui para anexo com nome fixo
        attachment.add_header('Content-Disposition', 'attachment', filename="Memorando.pdf")
        msg.attach(attachment)

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.quit()
            
            return {"status": "sucesso", "mensagem": "Memorando gerado e enviado por e-mail com sucesso!"}
        except Exception as e_smtp:
            print(f"Erro no SMTP do Google: {e_smtp}")
            return {"status": "erro", "mensagem": "Falha de autenticação no Google. Verifique a senha de app."}

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)