import requests

# Autenticar
auth = requests.post('https://app.pipefy.com/oauth/token', data={
    'grant_type': 'client_credentials',
    'client_id': 'ofgUSnXFhXadEzrDd_ZtUzXsV8-Crv-0NFboRn0CbrU',
    'client_secret': 'DyLPVER8t6SIeVpO7lQiDqTzoquM3UqLDOUgOFtHFpw'
}, headers={'Content-Type': 'application/x-www-form-urlencoded'})
token = auth.json()['access_token']

# Buscar todos os cards
query = '''query { cards(pipe_id: 306527874, first: 50) { edges { node { phases_history { phase { name } duration } } } pageInfo { hasNextPage endCursor } } }'''

all_cards = []
cursor = None
while True:
    q = query if not cursor else query.replace('first: 50', f'first: 50, after: "{cursor}"')
    resp = requests.post('https://api.pipefy.com/graphql', headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}, json={'query': q})
    data = resp.json()['data']['cards']
    all_cards.extend([e['node'] for e in data['edges']])
    if not data['pageInfo']['hasNextPage']:
        break
    cursor = data['pageInfo']['endCursor']

print(f'Total de cards: {len(all_cards)}')

# Calcular media por fase
tempos_por_fase = {}
for card in all_cards:
    for ph in card.get('phases_history', []):
        fase = ph['phase']['name']
        duration = ph.get('duration') or 0
        dias = duration / 86400
        if dias > 0:
            if fase not in tempos_por_fase:
                tempos_por_fase[fase] = []
            tempos_por_fase[fase].append(dias)

print()
print('MEDIA DE TEMPO POR FASE (API):')
print('=' * 60)
for fase, tempos in sorted(tempos_por_fase.items()):
    media = sum(tempos) / len(tempos)
    print(f'{fase}:')
    print(f'  Media: {media:.6f} dias')
    print(f'  Total registros: {len(tempos)}')
    print()
