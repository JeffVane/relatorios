import sqlite3


def create_user_eliete():
    conn = sqlite3.connect(r"\\srvsql\banco fisc\fiscais.db")
    cursor = conn.cursor()

    # Verificar se o usuário "ELIETE" já existe
    cursor.execute("SELECT name FROM fiscals WHERE name = 'ELIETE'")
    user_exists = cursor.fetchone()

    if not user_exists:
        # Inserir o usuário "ELIETE" com senha "123456" e como administrador
        cursor.execute("INSERT INTO fiscals (name, password, is_admin) VALUES (?, ?, ?)", ('ELIETE', '123456', 1))

        # Criar a tabela de procedimentos para o usuário "ELIETE" com a coluna 'quantidade'
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS procedimentos_ELIETE (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                coluna_1 TEXT,
                    coluna_2 TEXT,
                    coluna_3 TEXT,
                    coluna_4 TEXT,
                    coluna_5 TEXT,
                    coluna_6 TEXT,
                procedimento TEXT,
                quantidade INTEGER,
                realizado INTEGER DEFAULT 0,
                meta_anual_cfc INTEGER DEFAULT 0,
                crcdf_30 INTEGER DEFAULT 0,
                a_realizar INTEGER DEFAULT 0
            )
        ''')
        conn.commit()
        print("Usuário 'ELIETE' criado com a coluna 'quantidade' na tabela de procedimentos.")
    else:
        print("Usuário 'ELIETE' já existe no banco de dados.")

    conn.close()


# Executar a função
create_user_eliete()
