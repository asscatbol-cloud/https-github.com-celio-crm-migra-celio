import sqlite3
import pandas as pd
import os
from datetime import datetime

DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "MIGRA 1 CELIO OUT.xlsx")
DB_PATH = os.path.join(os.path.dirname(__file__), "data", "interactions.db")
PROPOSALS_DIR = os.path.join(os.path.dirname(__file__), "data", "proposals")

def init_db():
    """Inicializa o banco de dados SQLite para salvar interações e propostas."""
    os.makedirs(PROPOSALS_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS interactions (
            row_id INTEGER PRIMARY KEY,
            cnpj TEXT,
            cliente TEXT,
            status TEXT,
            atendeu TEXT,
            observacao TEXT,
            consultor TEXT,
            proposal_notes TEXT,
            proposal_image_path TEXT,
            proposal_image_date TIMESTAMP,
            updated_at TIMESTAMP
        )
    """)
    
    # Adicionar colunas se tabela já existia sem elas
    for col_def in [
        ("proposal_notes", "TEXT"),
        ("proposal_image_path", "TEXT"),
        ("proposal_image_date", "TIMESTAMP")
    ]:
        try:
            cursor.execute(f"ALTER TABLE interactions ADD COLUMN {col_def[0]} {col_def[1]}")
        except sqlite3.OperationalError:
            pass

    conn.commit()
    conn.close()

def load_dataset():
    """Carrega o Excel original e mescla com o banco de dados de interações e propostas."""
    init_db()
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError(f"Arquivo não encontrado em: {DATA_PATH}")

    df = pd.read_excel(DATA_PATH)
    df["ROW_ID"] = df.index

    # Carregar interações e propostas existentes
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT 
            row_id as ROW_ID, 
            status as STATUS_CHAMADA, 
            atendeu as ATENDEU, 
            observacao as INTERACAO_CONSULTOR, 
            consultor as CONSULTOR, 
            proposal_notes as PROPOSAL_NOTES,
            proposal_image_path as PROPOSAL_IMAGE_PATH,
            proposal_image_date as PROPOSAL_IMAGE_DATE,
            updated_at as DATA_ATENDIMENTO 
        FROM interactions
    """
    interactions_df = pd.read_sql_query(query, conn)
    conn.close()

    # Mesclar com os dados do Excel
    merged = pd.merge(df, interactions_df, on="ROW_ID", how="left")
    
    # Preencher NaN em colunas de interação e propostas
    merged["STATUS_CHAMADA"] = merged["STATUS_CHAMADA"].fillna("Pendente")
    merged["ATENDEU"] = merged["ATENDEU"].fillna("Não Registrado")
    merged["INTERACAO_CONSULTOR"] = merged["INTERACAO_CONSULTOR"].fillna("")
    merged["CONSULTOR"] = merged["CONSULTOR"].fillna("")
    merged["PROPOSAL_NOTES"] = merged["PROPOSAL_NOTES"].fillna("")
    merged["PROPOSAL_IMAGE_PATH"] = merged["PROPOSAL_IMAGE_PATH"].fillna("")
    merged["PROPOSAL_IMAGE_DATE"] = merged["PROPOSAL_IMAGE_DATE"].fillna("")
    merged["DATA_ATENDIMENTO"] = merged["DATA_ATENDIMENTO"].fillna("")

    # Limpeza de campos nulos em strings para exibição limpa no app
    string_cols = [
        "CLIENTE", "CNPJ_CLIENTE", "CONTATO", "NR_TELEFONE", "DS_MUNICIPIO", "UF",
        "NR_ENDERECO", "Nº", "NR_CEP", "PLANO", "TIPO_REDE", "TEM_COBERTURA_BANDA_LARGA",
        "TA NA MANCHA ", "APARELHO_TRAFEGO_MARCA", "APARELHO_TRAFEGO_MODELO",
        "RECOMENDACAO_APARELHO_LINHA"
    ]
    for col in string_cols:
        if col in merged.columns:
            merged[col] = merged[col].astype(str).replace(["nan", "NaN", "None", "<NA>"], "").str.strip()

    return merged

def save_proposal(row_id, notes, image_bytes=None, filename=None):
    """Salva notas e imagem da proposta para determinado cliente."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_dt = datetime.now()
    now_str = now_dt.strftime("%d/%m/%Y às %H:%M:%S")

    # Verificar dados atuais
    cursor.execute("SELECT proposal_image_path, proposal_image_date FROM interactions WHERE row_id=?", (int(row_id),))
    existing = cursor.fetchone()
    current_img_path = existing[0] if existing and existing[0] else ""
    current_img_date = existing[1] if existing and existing[1] else ""

    # Se uma nova imagem foi enviada
    if image_bytes is not None and filename:
        ext = os.path.splitext(filename)[1] or ".png"
        saved_filename = f"proposal_{row_id}_{now_dt.strftime('%Y%m%d_%H%M%S')}{ext}"
        target_path = os.path.join(PROPOSALS_DIR, saved_filename)
        
        with open(target_path, "wb") as f:
            f.write(image_bytes)
            
        current_img_path = target_path
        current_img_date = now_str

    cursor.execute("""
        INSERT INTO interactions (row_id, proposal_notes, proposal_image_path, proposal_image_date, updated_at)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(row_id) DO UPDATE SET
            proposal_notes=excluded.proposal_notes,
            proposal_image_path=CASE WHEN excluded.proposal_image_path != '' THEN excluded.proposal_image_path ELSE interactions.proposal_image_path END,
            proposal_image_date=CASE WHEN excluded.proposal_image_date != '' THEN excluded.proposal_image_date ELSE interactions.proposal_image_date END,
            updated_at=excluded.updated_at
    """, (int(row_id), str(notes), current_img_path, current_img_date, now_dt.strftime("%Y-%m-%d %H:%M:%S")))

    conn.commit()
    conn.close()

def save_interaction(row_id, cnpj, cliente, status, atendeu, observacao, consultor="Consultor"):
    """Salva ou atualiza a interação do consultor para determinado cliente."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT INTO interactions (row_id, cnpj, cliente, status, atendeu, observacao, consultor, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(row_id) DO UPDATE SET
            status=excluded.status,
            atendeu=excluded.atendeu,
            observacao=excluded.observacao,
            consultor=excluded.consultor,
            updated_at=excluded.updated_at
    """, (int(row_id), str(cnpj), str(cliente), str(status), str(atendeu), str(observacao), str(consultor), now))

    conn.commit()
    conn.close()

def export_results():
    """Gera um DataFrame pronto para exportação em Excel ou CSV."""
    return load_dataset()
