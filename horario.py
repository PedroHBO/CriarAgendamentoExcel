import pandas as pd
from datetime import datetime, timedelta

# Definir data de início e fim
data_inicio = datetime(2023, 11, 22)
data_fim = datetime(2025, 12, 30)

# Inicializar listas para data e hora
lista_datas = []
lista_horarios = []

# Criar o DataFrame
df = pd.DataFrame(columns=['Data', 'Hora', 'Agendado', 'Nome'])

# Função para verificar se a data é um sábado ou domingo
def is_weekend(date):
    return date.weekday() >= 5  # Retorna True se for sábado ou domingo

# Preencher o DataFrame com datas e horários, excluindo sábados e domingos
data_atual = data_inicio
while data_atual <= data_fim:
    if not is_weekend(data_atual):
        hora_inicio = datetime.strptime('08:00', '%H:%M')
        hora_fim = datetime.strptime('18:00', '%H:%M')
        
        while hora_inicio <= hora_fim:
            lista_datas.append(data_atual.strftime('%d/%m/%Y'))
            lista_horarios.append(hora_inicio.strftime('%H:%M'))
            hora_inicio += timedelta(hours=1)
    
    data_atual += timedelta(days=1)

# Preencher o DataFrame com as listas de datas e horários
df['Data'] = lista_datas
df['Hora'] = lista_horarios

# Salvar o DataFrame como um arquivo Excel
df.to_excel('horarios.xlsx', sheet_name='Horarios', index=False, engine='openpyxl')
print("Arquivo Excel gerado com sucesso.")
