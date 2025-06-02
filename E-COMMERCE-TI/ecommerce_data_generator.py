import random
import uuid
from faker import Faker
import psycopg2
from datetime import datetime, timedelta

fake = Faker('pt_BR')

# Conexão com PostgreSQL
conn = psycopg2.connect(
    dbname="E-commerce-TI",
    user="postgres",
    password="12345",
    host="localhost",
    port="5432"
)
cur = conn.cursor()

# Limpar dados antigos (opcional)
cur.execute("TRUNCATE TABLE vendas, produtos_servicos, vendedores, clientes RESTART IDENTITY CASCADE")

# Categorias e produtos
categorias = {
    "Nuvem": [
        "AWS S3", "Azure Backup", "Google Cloud VM", "Dropbox Business",
        "Cloudflare CDN", "Oracle Cloud Storage", "IBM Cloud Functions"
    ],
    "Antivírus": [
        "Kaspersky Total Security", "McAfee LiveSafe", "Norton 360 Deluxe",
        "Bitdefender Premium", "Avast Ultimate", "ESET Internet Security"
    ],
    "Hardware": [
        "Notebook Lenovo i7", "Monitor LG UltraWide", "SSD 1TB NVMe", "HD 2TB Seagate",
        "Placa de Vídeo RTX 4060", "Processador Ryzen 7", "Fonte 750W Corsair",
        "Memória RAM 16GB DDR4", "Gabinete Gamer RGB", "Placa Mãe B550", "Cooler WaterCooler 240mm"
    ],
    "Periféricos": [
        "Mouse Gamer Razer", "Teclado Mecânico Redragon", "Headset HyperX",
        "Webcam Logitech 1080p", "Microfone Blue Yeti", "Mousepad XXL RGB",
        "Controle Xbox Wireless", "Hub USB-C", "Suporte Monitor Articulado"
    ],
    "Software": [
        "MS Office 365", "Adobe Photoshop CC", "Windows 11 Pro", "Visual Studio Pro",
        "AutoCAD LT", "CorelDRAW Graphics Suite", "Vegas Pro", "SketchUp Pro"
    ],
    "Licenças e Ativação": [
        "Licença Windows 10", "Chave Office Home", "Ativador Norton", "Licença Adobe PDF",
        "Serial Microsoft SQL Server", "Licença ESET Business"
    ],
    "Serviços Técnicos": [
        "Instalação Remota", "Backup em Nuvem", "Formatação Profissional",
        "Limpeza Física Completa", "Montagem de PC Gamer", "Otimização de Sistema"
    ],
    "Acessórios": [
        "Cabo HDMI 2.1", "Adaptador USB para Ethernet", "Case HD Externo", "Bateria para NoBreak",
        "Filtro de Linha", "Suporte para Headset", "Organizador de Cabos"
    ]
}


# Gerar vendedores
vendedores = []
for _ in range(500):
    nome = fake.name()
    vendedores.append(nome)
    cur.execute("INSERT INTO vendedores (nome) VALUES (%s)", (nome,))

# Gerar clientes
# Gerar clientes com localização internacional e dados básicos
clientes = []
for _ in range(200):
    sexo = random.choice(['Masculino', 'Feminino'])
    nome = fake.name_male() if sexo == 'Masculino' else fake.name_female()
    email = fake.email()
    telefone = fake.phone_number()
    data_nasc = fake.date_of_birth(minimum_age=18, maximum_age=65)
    
    # Localização internacional
    pais = fake.country()
    estado = fake.state()
    cidade = fake.city()

    clientes.append((nome, email, telefone, sexo, data_nasc, pais, estado, cidade))
    cur.execute("""
        INSERT INTO clientes (nome, email, telefone, sexo, data_nascimento, pais, estado, cidade)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """, (nome, email, telefone, sexo, data_nasc, pais, estado, cidade))

# Gerar produtos/serviços
produtos = []
for categoria, itens in categorias.items():
    for item in itens:
        preco = round(random.uniform(100, 5000), 2)
        produtos.append((item, categoria, preco))
        cur.execute("""
            INSERT INTO produtos_servicos (nome, categoria, preco)
            VALUES (%s, %s, %s)
        """, (item, categoria, preco))

# Obter IDs
cur.execute("SELECT id FROM clientes")
clientes_ids = [row[0] for row in cur.fetchall()]

cur.execute("SELECT id FROM vendedores")
vendedores_ids = [row[0] for row in cur.fetchall()]

cur.execute("SELECT id, preco FROM produtos_servicos")
produtos_ids_precos = cur.fetchall()

# Gerar vendas
for _ in range(10000):
    cliente_id = random.choice(clientes_ids)
    vendedor_id = random.choice(vendedores_ids)
    produto_id, preco_unitario = random.choice(produtos_ids_precos)
    quantidade = random.randint(1, 5)
    total = round(quantidade * preco_unitario, 2)
    data_venda = fake.date_between(start_date='-1y', end_date='today')

    cur.execute("""
        INSERT INTO vendas (cliente_id, vendedor_id, produto_id, data_venda, quantidade, total)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (cliente_id, vendedor_id, produto_id, data_venda, quantidade, total))

# Commit
conn.commit()
cur.close()
conn.close()

print("População finalizada com sucesso.")
