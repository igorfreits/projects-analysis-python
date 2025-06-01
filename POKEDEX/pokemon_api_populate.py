import requests
import psycopg2
from psycopg2.extras import execute_values
import time

# --- Conexão com o banco ---
conn = psycopg2.connect(
    host="localhost",
    database="pokedex",
    user="postgres",
    password="12345",
    port="5432"
)
cur = conn.cursor()

# --- Criação das tabelas ---
cur.execute("""
CREATE TABLE IF NOT EXISTS regioes (
    id SERIAL PRIMARY KEY,
    nome VARCHAR(50) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS tipos (
    id SERIAL PRIMARY KEY,
    nome VARCHAR(50) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS pokemons (
    id INT PRIMARY KEY,
    nome VARCHAR(50),
    altura INT,
    peso INT,
    hp INT,
    ataque INT,
    defesa INT,
    velocidade INT,
    regiao_id INT REFERENCES regioes(id)
);

CREATE TABLE IF NOT EXISTS pokemon_tipos (
    pokemon_id INT REFERENCES pokemons(id),
    tipo_id INT REFERENCES tipos(id),
    PRIMARY KEY (pokemon_id, tipo_id)
);

CREATE TABLE IF NOT EXISTS imagens (
    pokemon_id INT PRIMARY KEY REFERENCES pokemons(id),
    url VARCHAR(255)
);

CREATE TABLE IF NOT EXISTS evolucoes (
    pokemon_id INT REFERENCES pokemons(id),
    evolui_para_id INT REFERENCES pokemons(id),
    PRIMARY KEY (pokemon_id, evolui_para_id)
);
""")
conn.commit()

# --- Inserção de regiões fixas ---
regioes = [
    (1, 'Kanto'), (2, 'Johto'), (3, 'Hoenn'), (4, 'Sinnoh'),
    (5, 'Unova'), (6, 'Kalos'), (7, 'Alola'), (8, 'Galar'), (9, 'Paldea')
]
execute_values(cur,
    "INSERT INTO regioes (id, nome) VALUES %s ON CONFLICT (id) DO NOTHING;",
    regioes
)
conn.commit()

# --- Funções auxiliares ---

def get_regiao_id_por_nome(nome_regiao):
    cur.execute("SELECT id FROM regioes WHERE nome = %s", (nome_regiao,))
    res = cur.fetchone()
    return res[0] if res else None

def get_regiao_pokemon(pokemon_id):
    url = f'https://pokeapi.co/api/v2/pokemon-species/{pokemon_id}/'
    resp = requests.get(url)
    if resp.status_code != 200:
        return None
    data = resp.json()
    gen_name = data['generation']['name']
    gen_regiao_map = {
        'generation-i': 'Kanto',
        'generation-ii': 'Johto',
        'generation-iii': 'Hoenn',
        'generation-iv': 'Sinnoh',
        'generation-v': 'Unova',
        'generation-vi': 'Kalos',
        'generation-vii': 'Alola',
        'generation-viii': 'Galar',
        'generation-ix': 'Paldea'
    }
    return gen_regiao_map.get(gen_name, None)

def get_or_create_tipo(nome_tipo):
    cur.execute("SELECT id FROM tipos WHERE nome = %s", (nome_tipo,))
    res = cur.fetchone()
    if res:
        return res[0]
    cur.execute("INSERT INTO tipos (nome) VALUES (%s) RETURNING id", (nome_tipo,))
    tipo_id = cur.fetchone()[0]
    conn.commit()
    return tipo_id

def popular_pokemon_sem_evolucao(pokemon_id):
    url = f'https://pokeapi.co/api/v2/pokemon/{pokemon_id}/'
    resp = requests.get(url)
    if resp.status_code != 200:
        print(f'Pokémon {pokemon_id} não encontrado')
        return

    data = resp.json()
    nome = data['name']
    altura = data['height']
    peso = data['weight']
    stats = {stat['stat']['name']: stat['base_stat'] for stat in data['stats']}
    hp = stats.get('hp', 0)
    ataque = stats.get('attack', 0)
    defesa = stats.get('defense', 0)
    velocidade = stats.get('speed', 0)

    regiao_nome = get_regiao_pokemon(pokemon_id)
    regiao_id = get_regiao_id_por_nome(regiao_nome) if regiao_nome else None

    # Inserir Pokémon
    cur.execute("""
        INSERT INTO pokemons (id, nome, altura, peso, hp, ataque, defesa, velocidade, regiao_id)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            nome = EXCLUDED.nome,
            altura = EXCLUDED.altura,
            peso = EXCLUDED.peso,
            hp = EXCLUDED.hp,
            ataque = EXCLUDED.ataque,
            defesa = EXCLUDED.defesa,
            velocidade = EXCLUDED.velocidade,
            regiao_id = EXCLUDED.regiao_id;
    """, (pokemon_id, nome, altura, peso, hp, ataque, defesa, velocidade, regiao_id))
    conn.commit()

    # Popular tipos e vincular
    tipos = [t['type']['name'] for t in data['types']]
    for tipo_nome in tipos:
        tipo_id = get_or_create_tipo(tipo_nome)
        cur.execute("""
            INSERT INTO pokemon_tipos (pokemon_id, tipo_id)
            VALUES (%s, %s)
            ON CONFLICT (pokemon_id, tipo_id) DO NOTHING;
        """, (pokemon_id, tipo_id))
    conn.commit()

    # Inserir imagem
    imagem_url = data['sprites']['other']['official-artwork']['front_default']
    if imagem_url:
        cur.execute("""
            INSERT INTO imagens (pokemon_id, url)
            VALUES (%s, %s)
            ON CONFLICT (pokemon_id) DO UPDATE SET url = EXCLUDED.url;
        """, (pokemon_id, imagem_url))
        conn.commit()

    time.sleep(0.2)  # evitar rate limit

def popular_evolucoes(pokemon_id):
    url = f'https://pokeapi.co/api/v2/pokemon-species/{pokemon_id}/'
    resp = requests.get(url)
    if resp.status_code != 200:
        print(f'Pokémon species {pokemon_id} não encontrado para evolução')
        return

    data = resp.json()
    chain_url = data['evolution_chain']['url']
    chain_resp = requests.get(chain_url)
    if chain_resp.status_code != 200:
        print(f'Cadeia de evolução não encontrada para {pokemon_id}')
        return
    chain_data = chain_resp.json()['chain']

    def parse_chain(chain):
        current_id = int(chain['species']['url'].split('/')[-2])
        evolutions = []
        for evo in chain['evolves_to']:
            evo_id = int(evo['species']['url'].split('/')[-2])
            evolutions.append((current_id, evo_id))
            evolutions.extend(parse_chain(evo))
        return evolutions

    evolucoes = parse_chain(chain_data)
    for base_id, evo_id in evolucoes:
        try:
            cur.execute("""
                INSERT INTO evolucoes (pokemon_id, evolui_para_id)
                VALUES (%s, %s)
                ON CONFLICT (pokemon_id, evolui_para_id) DO NOTHING;
            """, (base_id, evo_id))
            conn.commit()
        except Exception as e:
            print(f'Erro ao inserir evolução {base_id} -> {evo_id}: {e}')
            conn.rollback()

    time.sleep(0.2)

# --- Execução principal ---

print('Populando pokémons (1 a 649)...')
for i in range(1, 650):
    try:
        popular_pokemon_sem_evolucao(i)
    except Exception as e:
        print(f'Erro no Pokémon {i}: {e}')
        conn.rollback()

print('Populando evoluções (1 a 649)...')
for i in range(1, 650):
    try:
        popular_evolucoes(i)
    except Exception as e:
        print(f'Erro na evolução do Pokémon {i}: {e}')
        conn.rollback()

print('Carga concluída com sucesso!')

cur.close()
conn.close()
