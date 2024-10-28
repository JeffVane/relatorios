import sqlite3

# Conectar ao banco de dados
db_path = 'fiscais.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# 1. Criar uma nova tabela sem as colunas indesejadas
cursor.execute('''
CREATE TABLE procedimentos_MFA_nova (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    procedimento TEXT
)
''')

# 2. Copiar os dados da tabela antiga para a nova tabela
cursor.execute('''
INSERT INTO procedimentos_MFA_nova (procedimento)
SELECT procedimento FROM procedimentos_MFA
''')

# 3. Excluir a tabela antiga
cursor.execute('DROP TABLE procedimentos_MFA')

# 4. Renomear a nova tabela para o nome da tabela antiga
cursor.execute('ALTER TABLE procedimentos_MFA_nova RENAME TO procedimentos_MFA')

# Commit e fechamento da conexão
conn.commit()
conn.close()
