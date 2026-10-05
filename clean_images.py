import os
import sqlite3

proposal_dir = os.path.join(os.path.dirname(__file__), "data", "proposals")
db_path = os.path.join(os.path.dirname(__file__), "data", "interactions.db")

# 1. Apagar arquivos na pasta data/proposals
if os.path.exists(proposal_dir):
    for f in os.listdir(proposal_dir):
        fp = os.path.join(proposal_dir, f)
        if os.path.isfile(fp):
            os.remove(fp)
    print("Todas as imagens da pasta proposals foram apagadas.")

# 2. Resetar banco de dados SQLite
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM proposal_images")
    cursor.execute("UPDATE interactions SET proposal_image_path='', proposal_image_date=''")
    conn.commit()
    conn.close()
    print("Registros de imagens no banco de dados resetados com sucesso.")
