import sqlite3
import pandas as pd
import os
import io
import base64
import streamlit as st
from datetime import datetime
from PIL import Image

DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "MIGRA 1 CELIO OUT.xlsx")
DB_PATH = os.path.join(os.path.dirname(__file__), "data", "interactions.db")
PROPOSALS_DIR = os.path.join(os.path.dirname(__file__), "data", "proposals")

def init_db():
    """Inicializa o banco de dados SQLite para salvar interações, propostas e histórico de imagens."""
    os.makedirs(PROPOSALS_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabela principal de interações
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
            callback_date TEXT,
            callback_time TEXT,
            updated_at TIMESTAMP
        )
    """)
    
    # Tabela de histórico de imagens por cliente
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proposal_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_id INTEGER,
            cnpj TEXT,
            image_path TEXT,
            image_date TEXT,
            image_b64 TEXT,
            created_at TIMESTAMP
        )
    """)
    
    # Adicionar colunas se tabela interações ou proposal_images já existiam sem elas
    for table_name, col_def in [
        ("interactions", ("proposal_notes", "TEXT")),
        ("interactions", ("proposal_image_path", "TEXT")),
        ("interactions", ("proposal_image_date", "TIMESTAMP")),
        ("interactions", ("callback_date", "TEXT")),
        ("interactions", ("callback_time", "TEXT")),
        ("proposal_images", ("image_b64", "TEXT"))
    ]:
        try:
            cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_def[0]} {col_def[1]}")
        except sqlite3.OperationalError:
            pass

    conn.commit()
    conn.close()

def compress_image_bytes(image_bytes, max_dim=1200, quality=82):
    """Comprime e redimensiona a imagem enviada para garantir carregamento instantâneo."""
    try:
        with Image.open(io.BytesIO(image_bytes)) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            out_buf = io.BytesIO()
            img.save(out_buf, format='JPEG', quality=quality, optimize=True)
            return out_buf.getvalue(), ".jpg"
    except Exception:
        return image_bytes, ".png"

@st.cache_data(ttl=600, show_spinner=False)
def load_dataset():
    """Carrega o Excel original e mescla com o banco de dados de interações (com Caching inteligente)."""
    init_db()
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError(f"Arquivo não encontrado em: {DATA_PATH}")

    df = pd.read_excel(DATA_PATH)
    df.columns = [str(c).strip() for c in df.columns]
    df["ROW_ID"] = df.index

    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = df[col].astype(str).replace(["nan", "NaN", "None", "<NA>"], "").str.strip()

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
            callback_date as CALLBACK_DATE,
            callback_time as CALLBACK_TIME,
            updated_at as DATA_ATENDIMENTO 
        FROM interactions
    """
    interactions_df = pd.read_sql_query(query, conn)
    conn.close()

    merged = pd.merge(df, interactions_df, on="ROW_ID", how="left")
    
    merged["STATUS_CHAMADA"] = merged["STATUS_CHAMADA"].fillna("Pendente")
    merged["ATENDEU"] = merged["ATENDEU"].fillna("Não Registrado")
    merged["INTERACAO_CONSULTOR"] = merged["INTERACAO_CONSULTOR"].fillna("")
    merged["CONSULTOR"] = merged["CONSULTOR"].fillna("")
    merged["PROPOSAL_NOTES"] = merged["PROPOSAL_NOTES"].fillna("")
    merged["PROPOSAL_IMAGE_PATH"] = merged["PROPOSAL_IMAGE_PATH"].fillna("")
    merged["PROPOSAL_IMAGE_DATE"] = merged["PROPOSAL_IMAGE_DATE"].fillna("")
    merged["CALLBACK_DATE"] = merged["CALLBACK_DATE"].fillna("")
    merged["CALLBACK_TIME"] = merged["CALLBACK_TIME"].fillna("")
    merged["DATA_ATENDIMENTO"] = merged["DATA_ATENDIMENTO"].fillna("")

    return merged

def save_proposal(row_id, notes, image_bytes=None, filename=None, cnpj=""):
    """Salva notas e adiciona uma imagem de proposta otimizada ao histórico do cliente."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_dt = datetime.now()
    now_str = now_dt.strftime("%d/%m/%Y às %H:%M:%S")

    current_img_name = ""
    current_img_date = ""

    if image_bytes is not None and filename:
        compressed_bytes, ext = compress_image_bytes(image_bytes)
        saved_filename = f"proposal_{row_id}_{now_dt.strftime('%Y%m%d_%H%M%S')}{ext}"
        target_path = os.path.join(PROPOSALS_DIR, saved_filename)
        
        with open(target_path, "wb") as f:
            f.write(compressed_bytes)
            
        b64_str = base64.b64encode(compressed_bytes).decode('utf-8')
        current_img_name = saved_filename
        current_img_date = now_str

        # Inserir no histórico de imagens (guardando apenas o nome do arquivo + base64 backup)
        cursor.execute("""
            INSERT INTO proposal_images (row_id, cnpj, image_path, image_date, image_b64, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (int(row_id), str(cnpj), saved_filename, now_str, b64_str, now_dt.strftime("%Y-%m-%d %H:%M:%S")))

    # Atualizar notas e última imagem na tabela principal (guardando apenas o nome do arquivo)
    cursor.execute("""
        INSERT INTO interactions (row_id, cnpj, proposal_notes, proposal_image_path, proposal_image_date, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(row_id) DO UPDATE SET
            proposal_notes=excluded.proposal_notes,
            proposal_image_path=CASE WHEN excluded.proposal_image_path != '' THEN excluded.proposal_image_path ELSE interactions.proposal_image_path END,
            proposal_image_date=CASE WHEN excluded.proposal_image_date != '' THEN excluded.proposal_image_date ELSE interactions.proposal_image_date END,
            updated_at=excluded.updated_at
    """, (int(row_id), str(cnpj), str(notes), current_img_name, current_img_date, now_dt.strftime("%Y-%m-%d %H:%M:%S")))

    conn.commit()
    conn.close()
    st.cache_data.clear()

@st.cache_data(ttl=300, show_spinner=False)
def get_proposal_images(row_id):
    """Retorna todas as imagens da proposta associadas a um cliente (com restauração automática de backup)."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT image_path, image_date, image_b64 
        FROM proposal_images 
        WHERE row_id=? 
        ORDER BY id DESC
    """, (int(row_id),))
    
    rows = cursor.fetchall()
    images = []
    seen_paths = set()
    
    for r in rows:
        raw_path, dt, b64_str = r[0], r[1], r[2]
        if not raw_path:
            continue
            
        img_name = os.path.basename(raw_path)
        target_path = os.path.join(PROPOSALS_DIR, img_name)
        
        # Se a imagem sumiu do disco, restaurar a partir do b64 armazenado no banco SQLite
        if not os.path.exists(target_path) and b64_str:
            try:
                img_data = base64.b64decode(b64_str)
                with open(target_path, "wb") as f:
                    f.write(img_data)
            except Exception:
                pass

        if os.path.exists(target_path) and target_path not in seen_paths:
            images.append({"path": target_path, "date": dt})
            seen_paths.add(target_path)

    # Backup/Fallback: Se tabela proposal_images não tem registros, checar tabela principal interactions
    if not images:
        cursor.execute("SELECT proposal_image_path, proposal_image_date FROM interactions WHERE row_id=?", (int(row_id),))
        row = cursor.fetchone()
        if row and row[0]:
            img_name = os.path.basename(row[0])
            target_path = os.path.join(PROPOSALS_DIR, img_name)
            if os.path.exists(target_path):
                images.append({"path": target_path, "date": row[1] or "Data não registrada"})

    conn.close()
    return images

def delete_proposal_image(row_id, image_path):
    """Exclui uma imagem específica do cliente do disco e do banco de dados SQLite."""
    init_db()
    
    img_name = os.path.basename(image_path)
    target_path = os.path.join(PROPOSALS_DIR, img_name)
    
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception:
            pass
            
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Deletar da tabela proposal_images (tanto por nome quanto por caminho completo antigo)
    cursor.execute("DELETE FROM proposal_images WHERE row_id=? AND (image_path=? OR image_path=?)", (int(row_id), img_name, image_path))
    
    # Buscar imagem restante mais recente
    cursor.execute("SELECT image_path, image_date FROM proposal_images WHERE row_id=? ORDER BY id DESC LIMIT 1", (int(row_id),))
    remaining = cursor.fetchone()
    
    new_path = remaining[0] if remaining else ""
    new_date = remaining[1] if remaining else ""
    
    # Atualizar tabela principal interactions
    cursor.execute("UPDATE interactions SET proposal_image_path=?, proposal_image_date=? WHERE row_id=?", (new_path, new_date, int(row_id)))
    
    conn.commit()
    conn.close()
    st.cache_data.clear()

def import_new_dataset(file_bytes, filename):
    """Recebe um novo arquivo Excel/CSV e atualiza a planilha base de dados."""
    os.makedirs(os.path.dirname(DATA_PATH), exist_ok=True)
    
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))
        
    df.to_excel(DATA_PATH, index=False)
    st.cache_data.clear()
    return len(df)

def save_interaction(row_id, cnpj, cliente, status, atendeu, observacao, consultor="Consultor", callback_date="", callback_time=""):
    """Salva ou atualiza a interação do consultor para determinado cliente."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT INTO interactions (row_id, cnpj, cliente, status, atendeu, observacao, consultor, callback_date, callback_time, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(row_id) DO UPDATE SET
            status=excluded.status,
            atendeu=excluded.atendeu,
            observacao=excluded.observacao,
            consultor=excluded.consultor,
            callback_date=excluded.callback_date,
            callback_time=excluded.callback_time,
            updated_at=excluded.updated_at
    """, (int(row_id), str(cnpj), str(cliente), str(status), str(atendeu), str(observacao), str(consultor), str(callback_date), str(callback_time), now))

    conn.commit()
    conn.close()
    st.cache_data.clear()

def export_results():
    """Gera um DataFrame pronto para exportação em Excel ou CSV."""
    return load_dataset()
