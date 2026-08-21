import os
import pandas as pd
from datetime import datetime, date
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import create_engine, Column, Integer, String, Float, Date
from sqlalchemy.orm import declarative_base, sessionmaker
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 1. CONEXÃO COM O BANCO DE DADOS (Neon PostgreSQL)
DATABASE_URL = "postgresql://neondb_owner:npg_foy20hbmuILp@ep-proud-sea-ay8wz1z6.c-5.us-east-2.aws.neon.tech/neondb?sslmode=require"
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# Modelo da Tabela
class Contrato(Base):
    __tablename__ = "contratos"
    id = Column(Integer, primary_key=True, index=True)
    fornecedor = Column(String)
    objeto = Column(String)
    processo = Column(String, unique=True, index=True)
    vigencia_inicio = Column(Date)
    vigencia_fim = Column(Date)
    valor_bruto = Column(Float)
    gestor = Column(String)
    status = Column(String)
    dias = Column(Integer)

Base.metadata.create_all(bind=engine)

# 2. CONFIGURAÇÃO DO FASTAPI
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- FUNÇÕES AUXILIARES ---
def calcular_status_e_dias(data_fim):
    if not data_fim or pd.isna(data_fim):
        return "Indefinido", 0
    hoje = date.today()
    dias_restantes = (data_fim - hoje).days
    
    if dias_restantes < 0:
        return "Encerrado", dias_restantes
    elif dias_restantes <= 60:
        return "Crítico", dias_restantes
    elif dias_restantes <= 120:
        return "Atenção", dias_restantes
    else:
        return "No Prazo", dias_restantes

# --- ROTAS DA API ---

@app.post("/api/importar")
async def importar_planilha(file: UploadFile = File(...)):
    try:
        df = pd.read_excel(file.file)
        
        db = SessionLocal()
        for index, row in df.iterrows():
            processo = str(row.get('Processo', '')).strip()
            if not processo:
                continue
                
            vigencia_fim = pd.to_datetime(row.get('Término'), errors='coerce').date() if pd.notna(row.get('Término')) else None
            vigencia_inicio = pd.to_datetime(row.get('Início'), errors='coerce').date() if pd.notna(row.get('Início')) else None
            status, dias = calcular_status_e_dias(vigencia_fim)
            valor = float(row.get('Valor Anual', 0)) if pd.notna(row.get('Valor Anual')) else 0.0

            contrato_existente = db.query(Contrato).filter(Contrato.processo == processo).first()
            
            if contrato_existente:
                contrato_existente.fornecedor = str(row.get('Fornecedor', ''))
                contrato_existente.objeto = str(row.get('Objeto do Contrato', ''))
                contrato_existente.vigencia_inicio = vigencia_inicio
                contrato_existente.vigencia_fim = vigencia_fim
                contrato_existente.valor_bruto = valor
                contrato_existente.gestor = str(row.get('Gestor', ''))
                contrato_existente.status = status
                contrato_existente.dias = dias
            else:
                novo_contrato = Contrato(
                    fornecedor=str(row.get('Fornecedor', '')),
                    objeto=str(row.get('Objeto do Contrato', '')),
                    processo=processo,
                    vigencia_inicio=vigencia_inicio,
                    vigencia_fim=vigencia_fim,
                    valor_bruto=valor,
                    gestor=str(row.get('Gestor', '')),
                    status=status,
                    dias=dias
                )
                db.add(novo_contrato)
                
        db.commit()
        db.close()
        return {"status": "sucesso", "mensagem": "Planilha processada com sucesso!"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "erro", "mensagem": str(e)})

@app.get("/api/contratos")
def listar_contratos():
    db = SessionLocal()
    contratos = db.query(Contrato).all()
    db.close()
    
    resultado = []
    for c in contratos:
        resultado.append({
            "fornecedor": c.fornecedor,
            "objeto": c.objeto,
            "processo": c.processo,
            "vigencia_inicio": c.vigencia_inicio.strftime('%d/%m/%Y') if c.vigencia_inicio else '-',
            "vigencia_fim": c.vigencia_fim.strftime('%d/%m/%Y') if c.vigencia_fim else '-',
            "valor_bruto": c.valor_bruto,
            "valor_formatado": f"R$ {c.valor_bruto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "gestor": c.gestor,
            "status": c.status,
            "dias": c.dias
        })
    return resultado

@app.get("/api/gerar-memorando/{processo}")
def gerar_memorando(processo: str, tipo: str = "continuo"):
    db = SessionLocal()
    contrato = db.query(Contrato).filter(Contrato.processo == processo).first()
    db.close()
    
    if not contrato:
        raise HTTPException(status_code=404, detail="Contrato não encontrado no banco de dados.")

    # Criação do PDF
    pdf = FPDF()
    pdf.add_page()
    
    # Cabeçalho Limpo
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 10, "Memorando - SERV-CONTRATACOES", ln=True, align="C")
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 5, f"Goiânia, {date.today().strftime('%d/%m/%Y')}", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("helvetica", size=11)
    
    data_fim_formatada = contrato.vigencia_fim.strftime('%d/%m/%Y') if contrato.vigencia_fim else 'Não definida'

    # Modelos Baseados na Seleção do Usuário
    if tipo == "continuo":
        # MODELO "CLARO" - SERVIÇOS CONTÍNUOS E PRORROGÁVEIS
        texto = f"""**Ao Gestor do Contrato:** {contrato.gestor}
**Assunto:** Aviso de Vencimento e Avaliação de Prorrogação (Serviço Contínuo)

Informamos que o contrato referente ao processo **{contrato.processo}**, firmado com a empresa **{contrato.fornecedor}** para o objeto "{contrato.objeto}", possui término de vigência previsto para **{data_fim_formatada}**.

Tratando-se de serviço de natureza contínua, solicitamos que seja avaliada a necessidade, a vantajosidade e o interesse da Administração na **prorrogação contratual**, nos termos da legislação vigente. 

Em caso positivo, solicitamos o envio das justificativas técnicas, pesquisas de preço (se aplicável) e demais documentações pertinentes com a antecedência necessária para a formalização do termo aditivo antes do encerramento da vigência atual.

Atenciosamente,

Serviço de Contratações
Tribunal de Contas do Estado de Goiás"""

    else:
        # MODELO "BULL LTDA" - ESCOPO FECHADO / NÃO CONTÍNUO
        texto = f"""**Ao Gestor do Contrato:** {contrato.gestor}
**Assunto:** Aviso de Vencimento de Contrato (Escopo Fechado / Entrega)

Informamos que o contrato referente ao processo **{contrato.processo}**, firmado com a empresa **{contrato.fornecedor}** para o objeto "{contrato.objeto}", possui término de vigência previsto para **{data_fim_formatada}**.

Tratando-se de contratação por escopo ou entrega não contínua, solicitamos que seja providenciado o ateste conclusivo dos serviços prestados ou bens entregues para fins de liquidação financeira e encerramento regular do processo.

Caso haja atrasos justificáveis por parte da contratada ou pendências que exijam a extensão do prazo, solicitamos a imediata comunicação e o envio das justificativas para análise de possível termo aditivo de prazo, ressaltando que o pedido deve ocorrer **antes** do término da vigência atual.

Atenciosamente,

Serviço de Contratações
Tribunal de Contas do Estado de Goiás"""

    # Gerar o texto com negrito (markdown)
    pdf.multi_cell(0, 7, texto, markdown=True)
    
    # Salvar e Enviar
    filename = f"Memorando_{processo.replace('/', '-')}.pdf"
    filepath = f"/tmp/{filename}" if os.path.exists("/tmp") else filename
    pdf.output(filepath)
    
    return FileResponse(filepath, filename=filename, media_type='application/pdf')

@app.post("/api/enviar-notificacao/{processo}")
def enviar_notificacao(processo: str):
    db = SessionLocal()
    contrato = db.query(Contrato).filter(Contrato.processo == processo).first()
    db.close()
    
    if not contrato:
        return JSONResponse(status_code=404, content={"status": "erro", "mensagem": "Contrato não encontrado."})

    # Puxando os dados de email das variáveis do Render
    EMAIL_REMETENTE = os.getenv("EMAIL")
    SENHA_APP = os.getenv("SENHA_APP")
    EMAIL_DESTINO = "aeduardo@tce.go.gov.br" # E-mail do seu pai para testes reais
    
    if not EMAIL_REMETENTE or not SENHA_APP:
        return JSONResponse(status_code=500, content={"status": "erro", "mensagem": "Credenciais de e-mail não configuradas no servidor."})

    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINO
        msg['Subject'] = f"ALERTA TCE-GO: Vencimento do Processo {processo}"

        data_fim = contrato.vigencia_fim.strftime('%d/%m/%Y') if contrato.vigencia_fim else 'Não definida'
        corpo_email = f"""
        Olá,

        O contrato da empresa {contrato.fornecedor} (Processo: {processo}) vencerá em {data_fim}.
        Status atual: {contrato.status} ({contrato.dias} dias restantes).
        
        Por favor, acesse o Sistema de Gestão de Contratos para emitir o memorando e tomar as providências necessárias.

        Atenciosamente,
        Sistema Automatizado - Serviço de Contratações
        """
        
        msg.attach(MIMEText(corpo_email, 'plain'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP)
        server.send_message(msg)
        server.quit()
        
        return {"status": "sucesso", "mensagem": f"E-mail enviado com sucesso para {EMAIL_DESTINO}!"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "erro", "mensagem": f"Erro ao enviar e-mail: {str(e)}"})