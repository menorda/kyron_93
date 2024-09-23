#! python3

versão = '2.4'

banner = r'''
    ____  __._____.___.__________ ________    _______     ________________
    |    |/ _|\__  |   |\______   \\_____  \   \      \   /   __   \_____  \ 
    |      <   /   |   | |       _/ /   |   \  /   |   \  \____    / _(__  < 
    |    |  \  \____   | |    |   \/    |    \/    |    \    /    / /       \
    |____|__ \ / ______| |____|_  /\_______  /\____|__  /   /____/ /______  /
            \/ \/               \/         \/         \/                  \/  ''' + versão + '''
                                                                 csv -> kml
'''
print(banner)

with open("memento.txt", "w") as f:
    f.write(f'''
{banner}

    ********************************** BEM-VINDO! **********************************

    Este é um script para gerar um arquivo '.kml' a partir de um arquivo '.csv'.

    Fara fins de entendimento:
    -   'raiz' é a pasta onde está o intalador 'K93_install.py' e o script 'KYRON_93.bat'.
    -   'pasta do projeto' refere-se à pasta onde estará o arquivo '.csv' a ser tratado. Ela é uma subpasta da raiz, e pode ter qualquer nome.
    -   É desejável que cada arquivo '.csv' seja colocado em uma 'pasta do projeto' exclusiva.

    ********************************** INSTALAÇÃO **********************************

    A primeira execução do KYRON_93 deverá ser realizada por meio da linha de comando (tela preta), para fins de instalação.
    Para isso:
	1. Instalar o python3.12 (ou superior) no computador;
	2. Na pasta 'raiz', clicar com o botão direito do mouse, e selecionar 'Abrir no terminal';
   	3. Digitar: 'python K93_install.py';

    Será gerado o arquivo 'KYRON_93.bat', que servirá para a execução do script sem a necessidade da linha de comando.


    ********************************* CONFIGURAÇÃO *********************************

    O arquivo 'Conf.csv' é o arquivo de configuração, e deve estar na raiz.
    caso não exista, esse arquivo é gerado automaticamente.
    
    É composto de quatro Headers:
    'Cat'............Indica o nome das categorias cadastradas.
    'Icone'..........Indica o nome do arquivo a ser usado como icone, ou o link ('http' ou 'https') para acesso ao icone.
    'Regex'..........Indica a regex a ser usada para categorização automática dos placemarks, a partir da descrição.
    'Ocultar'........Indica se o nome de cada ponto dessa categoria deve aparecer permanentemente ou somente quando sobreposto pelo cursor. (True ou False)
    
    ************************************ ÍCONES ************************************

    Os ícones locais devem estar na pasta 'icones', subpasta da raiz.

    *********************************** HEADERS ************************************

    Considere usar os seguintes Headers:

    'Descrição' ou 'Evento'     Descreve o placemark, e serve como referência para a categorização e nomeação automáticas.
    'Coordenadas' ou 'Coord'    Indica a latitude e a longitude, em coordenadas geográficas, caso os valores estejam na mesma célula.
    'Latitude' ou 'Lat'         Indica a latitude, em coordenadas geográficas, caso esteja em célula separada da longitude. 
    'Longitude' ou 'Lon'        Indica a longitude, em coordenadas geográficas, caso esteja em célula separada da latitude.
    'Categoria' ou 'Cat'        Deve corresponder a uma das categorias cadastradas em 'Conf.csv'
    'Nome'                      Indica o título de cada ponto. Caso não conste, será gerado automaticamente com base na categoria.
    'Imagem' ou 'Img'           Indica o nome do arquivo de imagem '.jpg' ou '.png', ou o link (http ou https) a ser anexado ao placemark. Não é necessário indicar a extensão do arquivo. Se houver metadados de geolocalização na imagem 'jpg', eles podem ser usados para a geração do kml, não sendo necessários os campos 'Coord' ou 'Lat' e 'Lon';
    'Alt'                       Indica a altura do ponto
    
    Nenhum dos headers é obrigatório; entretanto, caso não seja possível obter dados de geolocalização, o arquivo kml não será gerado.

    ******************************* ANEXANDO IMAGENS *******************************

    No caso de uso da coluna 'Imagem', os arquivos referenciados devem estar dentro da pasta 'fotos', subpasta da 'pasta do projeto'. O KYRON_93 converterá as imagens automaticamente para o formato '.png'. O KYRON_93 gerará automaticamente um arquivo '.csv' listando os arquivos nessa pasta, e indicando se há presença de metadados.

    *************************** CATEGORIZAÇÃO AUTOMÁTICA ***************************

    Caso não seja indicada explicitamente uma categoria (header 'Cat') no arquivo de origem, o KYRON_93 atribuirá automaticamente uma das categorias cadastradas em 'Conf.csv'. Para isso, tomará por base o texto da coluna 'Evento', percorrendo as categorias cadastradas, desde o topo até a base da planilha. Havendo uma correspondência com as expressões regulares cadastradas, a categoria é atribuída. Dessa forma, mesmo que o texto contido em 'Evento' corresponda à regex de mais de uma categoria, será atribuída a categoria que se encontrar em uma posição superior em 'Conf.csv'. Em caso de dúvidas quanto ao uso de Regex, sugere-se solicitar ao ChatGPT "crie uma regex que seja compatível com os seguintes termos: ...".

    ************************************ FILTRO ************************************

    Em cada execução, o KYRON_93 gerará um arquivo '_filtro.csv', compilando os dados obtidos para cada item listado no arquivo '.csv' de origem. O arquivo não será sobrescrito, devendo ser deletado a cada vez que se desejar gerá-lo novamente.

    ***************************** NÃO ENTENDEU AINDA? ******************************

    Passo a passo:
    1. Instalar o 'KYRON_93', por meio do arquivo 'K93_install.py' (conforme o item 'INSTALAÇÃO' acima);
    1. Criar uma 'pasta do projeto', dentro da 'raiz' (pasta onde se localiza o script), com qualquer nome;
    2. Abrir uma planilha no MS Excel ou LibreOffice Calc;
    3. Nomear as colunas com os dados disponíveis (conforme o item 'HEADERS' acima);
    4. Lançar os dados disponíveis referentes aos pontos de interesse;
    5. Se houver fotos:
        a. Se houver imagens, criar dentro dessa pasta nova uma pasta chamada 'fotos';
        b. Salvar as imagens na pasta 'fotos';
        c. Criar, na planilha, uma coluna 'Imagens';
        d. Escrever no campo 'Imagens' o nome ou o link externo (http ou https) do arquivo da foto;
    6. Salvar a planilha no formato COMMA SEPARATED VALUES (.csv), com codificação 'utf-8', dentro da 'pasta do projeto';
    7. Executar o arquivo 'KYRON_93.bat'


    Desenvolvido no 6º BIM, em SET 2024.
    
    "OMNIA POSSIBILIA SUNT CREDENTI"

    MTM, em parceria com ChatGPT.
    ''')

import os, re, subprocess, sys, csv, logging
logging.getLogger('chardet.charsetprober').setLevel(logging.INFO)
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', encoding='UTF-8')

import os
import sys

default_timezone = 'Q'
esconde_nome = True

def criar_arquivo_bat():
    # Nome do arquivo .bat a ser criado
    nome_bat = 'KYRON_93.bat'

    # Caminho completo do script Python sendo executado
    caminho_script = os.path.abspath(sys.argv[0])

    # Descobrir o caminho do Python instalado
    caminho_python = sys.executable

    # Verifica se o arquivo .bat já existe
    if not os.path.exists(nome_bat):
        # Conteúdo do arquivo .bat
        conteudo_bat = f'@echo off\n"{caminho_python}" "{caminho_script}"\npause'

        # Cria o arquivo .bat
        with open(nome_bat, 'w') as arquivo_bat:
            arquivo_bat.write(conteudo_bat)

        logging.info(f"Arquivo '{nome_bat}' criado com sucesso!")
    else:
        logging.info(f"Arquivo '{nome_bat}' já existe.")

# Função para detectar a codificação do arquivo
def preopen(arquivo):
    encode = 'UTF-8'
    try:
        with open(arquivo, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encode = result['encoding']
    except Exception as e:
        print(f"Erro ao detectar a codificação: {e}")
    return encode

# Função para normalizar os cabeçalhos do CSV
def normalize_headers(headers):
    header_map = {
        'Cat': r'\b(cat(egoria|egorizad[oa])?s?|tipos?)\b',
        'Lat': r'\b(lat(itude)?s?)\b',
        'Lon': r'\b(lon(g|gitude)?s?)\b',
        'Decri': r'\b(des?cri([çcÇ]([aãÃ]o|[oõÕ]es))?|eventos?)\b',
        'Nome': r'\b(nomes?|t[íÍ]tulos?)\b',
        'Icone': r'\b([iíÍ]cone?s?)\b',
        'Regex': r'\b(e[sx]press([Ãã]o|[oõÕ]es)(\sregular(?:es)?]?)?|regex)\b',
        'Img': r'\b(ima?g(em|ens)?)\b',
        'Coor': r'\b(coo?rd?(enada[s]?)?)\b',
        'Alt': r'\b(alt(u|ura|itude)?)\b',
        'Data': r'\b(dat[ae]s?|GDH)',
        'Hora': r'\b(hor[aá]s?(?:rio)?)\b',
        'Hid': r'\b(escond(e|er|idos?)|ocult(os?|ar))\b'
    }
    
    normalized_headers = {}
    for header in headers:
        matched = False
        for key, pattern in header_map.items():
            if re.search(pattern, header, re.IGNORECASE):
                normalized_headers[header] = key
                logging.info(f'{header} convertido para {key}')
                matched = True
                break
        if not matched:
            logging.info(f'{header} mantido como {header}')
            normalized_headers[header] = header
    return normalized_headers

# Função para ler e normalizar o CSV lido
def ler_e_normalizar_csv(csv_path, limpar=None):
    try:
        encoding = preopen(csv_path)
        df_temp = pd.read_csv(csv_path, encoding=encoding, nrows=0)  # Lê apenas a primeira linha (cabeçalhos)
        original_headers = df_temp.columns.tolist()
        
        # Normaliza os cabeçalhos
        logging.info(f'Normalizando os cabeçalhos de {csv_path}')
        normalized_headers = normalize_headers(original_headers)
        
        # Lê o CSV novamente, agora com cabeçalhos normalizados
        df = pd.read_csv(csv_path, encoding=encoding, header=0)
        df.rename(columns=normalized_headers, inplace=True)

        if limpar != None:
            # Limpeza de linhas mortas
            try:
                df.dropna(subset=[limpar], inplace=True)
            except:
                logging.info(f'Não foi possível excluir as linhas mortas de {csv_path}')
        return df

    except pd.errors.EmptyDataError:
        logging.info(f"O arquivo CSV '{csv_path}' está vazio.")
        return
    except Exception as e:
        logging.info(f"Erro ao ler o arquivo CSV '{csv_path}': {e}")
        return

# Função para checar se o pip está instalado
def check_and_install_pip():
    try:
        # Verifica se o pip está instalado, rodando o comando "pip --version"
       subprocess.check_call([sys.executable, '-m', 'pip', '--version'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        print("pip não está instalado. Tentando instalar...")

        try:
            # Tenta instalar o pip usando o módulo ensurepip
            subprocess.check_call([sys.executable, '-m', 'ensurepip', '--upgrade'])
            print("pip foi instalado com sucesso.")
            
            # Atualiza o pip para a versão mais recente
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
            print("pip foi atualizado com sucesso.")
        except subprocess.CalledProcessError as e:
            print(f"Falha ao instalar o pip. Erro: {e}")
            sys.exit(1)

# Função para verificar e instalar bibliotecas
def check_and_install_libraries():
    libraries = ['pandas', 'exifread', 'pillow', 'chardet']
    
    for lib in libraries:
        try:
            if lib == 'pillow':
                __import__('PIL')
            else:
                __import__(lib)

        except ImportError:
            print(f"{lib} não encontrado. Instalando...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', lib])

# Chama a função para verificar e instalar o pip
check_and_install_pip()

# Chama a função para verificar e instalar as bibliotecas
check_and_install_libraries()

import chardet
import pandas as pd
import exifread
from PIL import Image, PngImagePlugin, ExifTags
# from PIL.ExifTags import TAGS, GPSTAGS

# Configuração de categorias
cat_to_icon = {'Sem Categoria': 'https://cdn-icons-png.flaticon.com/512/5999/5999673.png'}
cat_hide = {'Sem Categoria': True}
cat_filter = {}
filtro = None
cat_temp = {}
bolis = ['0x4d', '0x54', '0x4d', '0x34', '0x37', '0x30']

# Inicia o arquivo de log
def log_start(file='log.txt'):
    with open(file, 'w', encoding = 'UTF-8') as logbook:
        logbook.write(f'{banner}\n\n')

# Função para criar a estrutura de pastas de exemplo
def cria_estrutura():
    pasta_exemplo = os.path.join('.', 'exemplo')
    pasta_icones = os.path.join('.', 'icones')

    if not os.path.exists(pasta_icones):
        os.makedirs(pasta_icones)
        logging.info(f"Pasta '{pasta_icones}' criada.")
    
    if not os.path.exists(pasta_exemplo):
        os.makedirs(pasta_exemplo)
        logging.info(f"Pasta '{pasta_exemplo}' criada.")
        
        pasta_fotos = os.path.join(pasta_exemplo, 'fotos')
        os.makedirs(pasta_fotos)
        logging.info(f"Pasta '{pasta_fotos}' criada.")

        exemplo_csv = os.path.join(pasta_exemplo, 'exemplo.csv')
        dados_csv = {
            'Lat': ["28° 52' 23.6\" S"],
            'Lon': ["55° 31' 28.2\" W"],
            'Decri': ["Sítio Sepé Tiaraju"],
            'Imagem':["https://leouve.com.br/wp-content/uploads/2021/09/19826583.jpg"]
        }
        df = pd.DataFrame(dados_csv)
        df.to_csv(exemplo_csv, index=False)

# Função para inicializar o arquivo de configuração
def setup():
    # Caso não haja arquivo Conf.csv, cria um default
    if not 'Conf.csv' in os.listdir('.'):
        with open('Conf.csv', mode='w', newline='', encoding='UTF-8') as f:
            escritor_csv = csv.writer(f)
            # Escreve os cabeçalhos
            conf_headers = ('Cat', 'Icone','Regex','Ocultar')
            escritor_csv.writerow(conf_headers)
            # Escreve a categoria Casario, a titulo de exemplo
            conf_casario = ('Casario', r'https://cdn-icons-png.freepik.com/256/4056/4056445.png?semt=ais_hybrid',r'(?i)\b(fazenda|s[Ííi]tio|casa(rio)?)[s]?\b','True')
            escritor_csv.writerow (conf_casario)     
    reader = ler_e_normalizar_csv('Conf.csv')

    if reader is None:
        print("Erro ao processar 'Conf.csv'.")
        sys.exit(1)

    # Itera sobre cada linha no arquivo CSV, montando os dicionários categoria-ícone e placemarks
    for index, row in reader.iterrows():
        # Dicionário categoria-ícone
        cat_to_icon[row.get('Cat', 'Sem Categoria')] = row.get('Icone', 'logo.png')

        hidden = row.get('Hid', 'True')
        if hidden == '' or str(hidden) == 'nan':
            hidden = True
        cat_hide[row.get('Cat', 'Sem Categoria')] = hidden

        # Cria o dicionário para filtragem
        if 'Regex' in row:
            filtro = 'Regex'
            if row[filtro] != '':
                cat_filter[row.get('Cat', 'Sem Categoria')] = row[filtro]
        elif 'Keywords' in row:
            filtro = 'Keywords'
            if row[filtro] != '':
                cat_filter[row.get('Cat', 'Sem Categoria')] = row[filtro]

# Função para converter coordenadas geográficas em grau decimal
def gms2dec(coord):
    g = coord[0]
    m = coord[1]
    s = coord[2]
    d = coord[3]

    if g >= 0:
            newcoord = g + m / 60 + s / 3600
    else:
        newcoord = g - m / 60 - s / 3600
    if d in ['S', 'W', 'O'] and g >= 0:
        newcoord = -newcoord
    return newcoord

# Função para desconverter coordenadas geográficas, de grau decimal para graus minutos e segundos
def dec2gms(coord, tipo):
    # Define a direção
    dir = 'N'
    if coord < 0:
        if tipo == 'lat':
            dir = 'S'
        else:
            dir = 'W'
    elif tipo == 'lon':
        dir = 'E'

    coord = abs(coord)

    # Extrai a parte inteira dos graus
    graus = int(coord)
    
    # Calcula a parte decimal
    decimal = coord - graus
    
    # Converte a parte decimal para minutos
    minutos_decimais = decimal * 60
    minutos = int(minutos_decimais)
    
    # Calcula a parte decimal dos minutos
    segundos_decimais = (minutos_decimais - minutos) * 60
    segundos_formatados = round(segundos_decimais, 2)

    # Formata GMS, com 2 casas decimais
    string = f'{graus}° {minutos}\' {segundos_formatados:.2f}\" {dir}'

    return string

# Função para separar latitude e longitude
def dividir(coord):
    # Troca vírgulas decimais por pontos
    regex = r'(?<=[0-9]),(?=[0-9])'
    coord = re.sub(regex, '.', coord)
    
    # Fraciona Lat e Long
    # regex = r'((?<=[NSLEOW\-\/])[,\/\s]+|(?<=[0-9]),\s|(?<=[0-9])[,\s](?=[+\-]|[0-9]))'
    regex = r'((?<=[SLOWEN0-9])(\s?[,\/\-]\s?)|(?<=[SLOWEN])\s(?=[+\-\d]))'
    dividido = re.split(regex, coord)
    
    if len(dividido) < 2:
            logging.error(f"Não foi possível fracionar a coordenada {coord}. Separe Lat e Long no arquivo de origem.")
            return 0, 0
    lat = dividido[0]
    lon = dividido[-1] 
    newcoord = (lat, lon)
    return newcoord

# Função para capturar coordenadas geográficas
def capturar(coord):
    coord = coord.strip()
    coord = re.sub(r'\s{2,}', ' ', coord)
    coord = re.sub(',', '.', coord)
    regex = r'^([+-]?\d+\.?\d*)[°ºo]?\s*(\d*\.?\d*)?[\'’`]*\s*(\d*\.?\d*)?["“”]?\s*([NSEOWL]?)$'
    match = re.match(regex, coord)
    if match:
        deg = float(match.group(1))
        minutes = float(match.group(2)) if match.group(2) else 0
        seconds = float(match.group(3)) if match.group(3) else 0
        direction = match.group(4)
        newcoord = (deg, minutes, seconds, direction)
        newcoord = gms2dec(newcoord)
    else:
        try:
            newcoord = float(coord)
        except ValueError:
            logging.error(f'Coordenada inválida: {coord}')
            return
    return newcoord

def get_datetime(img_path):
    # Obtém os dados exif
    tags = get_exif_data(img_path)
    if tags == None:
        return None, None
    # Extrai a data e hora dos metadados EXIF
    datetime = tags.get('DateTimeOriginal') or tags.get('DateTime')
    
    if not datetime:
        return None, None
    
    # Converte a data e hora para o formato desejado AAAA-MM-DD HH:MM:SS
    def format_datetime(datetime_str):
        # O formato original do EXIF geralmente é "AAAA:MM:DD HH:MM:SS"
        date_part, time_part = datetime_str.split(" ")
        date_part = date_part.replace(":", "-")  # Substitui ':' por '-' na data
        return (date_part, time_part)
    
    # Formata e retorna a data e hora
    return format_datetime(datetime)

# Função para extrair todos os metadados EXIF da imagem
def get_exif_data(img_path):
    try:
        img = Image.open(img_path)
    except:
        logging.error(f'A imagem {img} não pode ser aberta.')
        return
    exif_data = {}
    info = img._getexif()
    if info:
        for tag, value in info.items():
            tag_name = ExifTags.TAGS.get(tag, tag)
            exif_data[tag_name] = value
    return exif_data

# Função para extração de metadados de geolocalização
def get_coordinates(tags):
    lat = tags.get('GPS GPSLatitude')
    lat_ref = tags.get('GPS GPSLatitudeRef')
    lon = tags.get('GPS GPSLongitude')
    lon_ref = tags.get('GPS GPSLongitudeRef')

    # <--- PRECEDE, GUIA E LIDERA!
    def convert_to_degrees(value_tags, dir_tag):
        g = float(value_tags[0].num) / float(value_tags[0].den)
        m = float(value_tags[1].num) / float(value_tags[1].den)
        s = float(value_tags[2].num) / float(value_tags[2].den)
        d = dir_tag.values[0]
        tudo = (g, m, s, d)
        return tudo

    if not lat or not lon:
        return None
    lat = convert_to_degrees(lat.values, lat_ref)
    lon = convert_to_degrees(lon.values, lon_ref)
    return (lat, lon) 

def geoextract(img_path):
    # Abre a imagem
    try:
        with open(img_path, 'rb') as f:
            tags = exifread.process_file(f)
    except FileNotFoundError:
        logging.error(f'A imagem {img_path} não foi encontrada.')
        return

    # Extrai as coordenadas
    coordinates = get_coordinates(tags)
    if not coordinates:
        logging.error(f'Metadados de geolocalização não encontrados em {img_path}')
        return
    return coordinates

# Função de filtragem por categoria
def autocat(texto, dicionario):
    if texto == '' or str(texto) == 'nan':
        return 'Sem Categoria'
    for key in dicionario.keys():
        if dicionario[key] == '':
            continue
        reg = re.compile(dicionario[key])
        if reg.search(texto):
            logging.debug(f'Categoria {key} atribuída ao texto {texto}')
            return key
    logging.debug(f'Não foi encontrada categoria para o texto {texto}')
    return 'Sem Categoria'

# Função para leitura de headers
def head_read(arquivo):
    with open(arquivo, mode='r', newline='', encoding='utf-8') as arquivo_csv:
        leitor = csv.reader(arquivo_csv)
        # Lê a primeira linha, que contém os headers
        headers = next(leitor, [])
        headers = normalize_headers(headers)
    logging.debug(f'Headers lidos: {headers}')
    return list(headers.values())

# Função para reordenar os placemarks
def reord_cat(places):
    outros = places.pop('Sem Categoria')
    for key in places:
        cat_temp[key] = places[key]
    places = {}
    for key in sorted(cat_temp):
        places[key] = cat_temp[key]
    places['Sem Categoria'] = outros
    return places

# Função para gerar o kml de cada placemark
def kml_placemark_gen(nome, cat, lat, lon, decri, raw_img='', alt=0, date=None, time=None, timezone=None):
    cod_img = f'<img style=\"max-width:500px;\" src=\"'
    if raw_img == '' or str(raw_img) == 'nan' or raw_img == None:
        cod_img = ''
    elif raw_img.startswith('http'):
        cod_img += f"{raw_img}\">"
    else:
        regex_ext = re.compile(r'\.[a-zA-Z]{3,}$')
        raw_img = os.path.basename(raw_img)
        if regex_ext.search(raw_img):
            img = regex_ext.sub('.png', raw_img)
        else:
            img = raw_img + '.png'
        raiz = '.'
        caminho_img = os.path.join(raiz, 'fotos', 'png', img)
        cod_img += f"{caminho_img}\">"
        logging.debug(f'Código de imagem: {cod_img}')
    
    # Define o código do icone
    if cat_to_icon[cat].startswith('http'):
        icone = cat_to_icon[cat]
    else:
        icone = os.path.join(os.path.abspath('.'), 'icones', cat_to_icon[cat])

    # Define o código de data
    if not time:
        cod_time = ''
    else:
        if not timezone:
            time = time + default_timezone
        elif not time.endswith(timezone):
            time = time + timezone
        cod_time = f'T{time}'
    if not date:
        cod_data = ''
    else:
        cod_data = f'''
        <TimeStamp>
            <when>{date}{cod_time}</when>
        </TimeStamp>'''
    
    if esconde_nome:
        style = f'''
        <styleUrl>#{cat}_onoff</styleUrl>'''
    else:
        # Define o código do icone
        if cat_to_icon[cat].startswith('http'):
            icone = cat_to_icon[cat]
        else:
            icone = os.path.join(os.path.abspath('.'), 'icones', cat_to_icon[cat])
        style = f'''
        <Style>
            <IconStyle>
            <Icon>
                <href>{icone}</href>
            </Icon>
            </IconStyle>
        </Style>'''

    # Define o código do placemark
    placemark = f'''
        <Placemark>
        <name>{nome}</name>
        <description>
        <![CDATA[
            {cod_img}
            <p>{decri}</p>
            <br><br>
            <p align=right><sub><i>Gerado pelo script KYRON 93<br>Desenvolvimento: 6º BIM</i></sub></p>
            ]]>
        </description>
        {style} 
        <Point>
            <coordinates>{lon},{lat},{alt}</coordinates>
        </Point>{cod_data}
        </Placemark>
        '''
    return placemark

# Função para gerar o kml de todas as pastas, usando o dict 'placemarks'
def kml_folders_gen(placemarks):
    kml = ''
    placemarks = reord_cat(placemarks)
    valid_cats = []
    try:
        for cat in placemarks.keys():
            if placemarks[cat] != '':
                buffer = f''' 
                <Folder>
                <name>{cat}</name>
                {placemarks[cat]}
                </Folder>
                ''' 
                kml += buffer
                valid_cats.append(cat)
    except:
        print('Erro ao gerar o arquivo kml.')
        logging.error('Não foi possível gerar o arquivo kml. Verifique o arquivo de entrada.')
    return kml, valid_cats

# Função para realizar conversão de imagem jpg para png, acrescentando metadados de geolocalização
def jpg2png(entrada, saida, lat, lon):
    print(f'Convertendo imagem de {entrada} para {saida}')
    if not os.path.exists(entrada):
        logging.error(f'{entrada} não existe.')
        return
    
    try:
        with Image.open(entrada) as img:
            # Verifica a orientação da imagem nos metadados EXIF
            exif = img._getexif()
            if exif is not None:
                for tag, value in exif.items():
                    tag_name = ExifTags.TAGS.get(tag, tag)
                    if tag_name == 'Orientation':
                        # Aplica a rotação conforme a orientação
                        if value == 3:
                            img = img.rotate(180, expand=True)
                        elif value == 6:
                            img = img.rotate(270, expand=True)
                        elif value == 8:
                            img = img.rotate(90, expand=True)

            img = img.convert('RGB')
            
            # Salva a imagem PNG e adiciona metadados
            png_info = PngImagePlugin.PngInfo()
            png_info.add_text('lat', str(lat))
            png_info.add_text('lon', str(lon))
            
            # Re-salva a imagem com os metadados
            img.save(saida, 'PNG', pnginfo=png_info)

            logging.info(f'Imagem convertida e salva em {saida}')
    except Exception as e:
        logging.error(f'Erro ao converter imagem: {e}')

# Função para realizar cópia de imagem png para png, acrescentando metadados de geolocalização
def png2png(entrada, saida, lat, lon):
    if not os.path.exists(entrada):
        logging.error(f'{entrada} não existe.')
        return
    try:
        with Image.open(entrada) as img:
            img = img.convert('RGB')
            
            # Cria um objeto PngInfo para adicionar metadados
            png_info = PngInfo()
            png_info.add_text('lat', str(lat))
            png_info.add_text('lon', str(lon))
            
            # Salva a imagem PNG com metadados
            img.save(saida, 'PNG', pnginfo=png_info)

            logging.info(f'Imagem convertida e salva em {saida}')
    except Exception as e:
        logging.error(f'Erro ao converter imagem: {e}')

# Função para verificar e ajustar estrutura de imagens
def preimg(folder, i, row, headers, lat=None, lon=None):
    input_folder = os.path.join(folder, 'fotos')
    output_folder = os.path.join(input_folder, 'png')
    if 'Img' in headers:
        raw_img = row['Img']
        if raw_img == '' or str(raw_img) == 'nan':
            logging.info(f'{folder} sem a imagem para a linha {i}')
            return
        if not os.path.exists(input_folder):
            logging.error(f'A pasta {input_folder} não existe!')
            return
        if raw_img.lower().endswith('.jpg') or raw_img.lower().endswith('.jpeg'):
            raw_img = raw_img
        else:
            raw_img = raw_img + '.jpg'
    else:
        logging.info(f'O campo \'Img\' não existe no arquivo.')
        return

    # se a imagem não é um link
    if not raw_img.lower().startswith('http'):
        raw_img = os.path.basename(raw_img)
        raw_img_path = os.path.join(input_folder, raw_img)
        # ajusta a extensão
        regex_ext = re.compile(r'\.[a-zA-Z]{3,}$')
        extension = regex_ext.search(raw_img)
        if extension:
            png_img = regex_ext.sub('.png', raw_img)
        else:
            png_img = raw_img + '.png'
            extension = ['.jpg']
        logging.debug(f'nome de imagem: {png_img}')
        png_img_path = os.path.join(output_folder, png_img)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            logging.info(f'Pasta {output_folder} criada.')
        if extension[0] == '.jpg' or extension[0] == '.jpeg':
            # se a imagem em png não está na pasta png
            if not os.path.exists(png_img_path):
                # se há uma imagem de origem
                if os.path.exists(raw_img_path):
                    jpg2png(raw_img_path, png_img_path, lat, lon)
                else:
                    logging.error(f'Imagem {raw_img} não localizada para conversão.')
                    return
        elif extension == '.png':
            png2png(raw_img_path, png_img_path, lat, lon)    
    # caso seja um link http
    else:
        raw_img_path = raw_img
        logging.debug(f'O caminho para a imagem {raw_img} é um link.')
        return raw_img_path
    logging.debug(f'Caminho para a imagem {png_img}: {png_img_path}')
    return png_img_path

# Função geradora de arquivo csv com dados filtrados
def gerar_filtro(locais, nome_arquivo):
    # Definir os cabeçalhos do CSV
    cabecalhos = ['Nome', 'Cat', 'Lat', 'Lon', 'Lat_gms', 'Lon_gms', 'Alt', 'Decri', 'Data', 'Hora', 'Gdh', 'Img']
    #nome, cat, lat, lon, lat_gms, lon_gms, alt, decri, date, time, gdh, img
    
    if os.path.exists(nome_arquivo):
        logging.info(f'O arquivo {nome_arquivo} já existe. Remova-o se desejar gerá-lo novamente.')
        print(f'O arquivo {nome_arquivo} já existe. Remova-o se desejar gerá-lo novamente.')
        return

    # Criar e abrir o arquivo CSV para escrita
    with open(nome_arquivo, mode='w', newline='', encoding='utf-8') as arquivo_csv:
        escritor_csv = csv.writer(arquivo_csv)
        
        # Escrever os cabeçalhos
        escritor_csv.writerow(cabecalhos)
        
        # Escrever as linhas com os dados de 'locais'
        for local in locais:
            escritor_csv.writerow(local)
    
    print(f"Arquivo '{nome_arquivo}' criado com sucesso.")

# Função para captura e ajuste de data para o formato AAAA-MM-DD
# Função para captura e ajuste de data para o formato AAAA-MM-DD
def captura_timedate(date, time=None, img = None):
    day = hour = minute = timezone = month = century = year = None
    if date:
        date = str(date)
    if time:
        time = str(time)
    meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAIO', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    
    if not date and not time and not img:
        return None

    # Se houver imagem, obtém os metadados
    if img:
        geodate, geotime = get_datetime(img)
        # Se não houver data ou hora, puxa os valores dos metadados
        if not date and geodate:
            date = geodate    
        if not time and geotime:
            time = geotime
    
    # Regex para capturar dia, mês e ano (formato DD/MM/AAAA ou DD-MM-AAAA)
    regex_data1 = re.compile(r'\b([0-3]?\d(?!\d))[\/-:]([01]?\d)[\/-:]([0-2]?\d(?=\d{2}))?(\d{2})\b')

    # Regex para capturar dia, mês e ano (formato AAAA-MM-DD)
    regex_data2 = re.compile(r'\b([0-2]\d)(\d{2})[-\/:]([01]\d)[-\/:]([0-3]\d)\b')
    
    # Regex para capturar hora, minuto e fuso-horário
    regex_time = re.compile(r'\b([0-2]?\d)\s*[h\:]?\s*([0-5]\d)\s?(?:min)?(?:(?:(?<=\s)|(?<=\d)|(?<=min)\s?)([A-Z]))?\b')

    # Regex para data no formato grupo data-hora (GDH) (dia, hora, minuto, fuso, mês, século, ano)
    regex_gdh = re.compile(r'(\d{2})\s?(\d{2})(\d{2})([A-Z]\s?(?=[A-Za-z^(?:AIO)]{3}))?\s?(\w{3}[Oo]?)\s?(\d{2})?(\d{2})')

    # Primeiro, tenta a correspondência como GDH
    date_match = regex_gdh.search(date)
    if date_match:
        day, hour, minute, timezone, month, century, year = date_match.groups()
        logging.debug(f'regex match: {day, hour, minute, timezone, month, century, year}')
        if century:
            full_year = f'{century}{year}'
        else:
            full_year = f'20{year}'
        logging.debug(f'Full year: {full_year}')
        month = meses.index(month.upper()) + 1  # Para ajustar ao formato numérico de mês
        time = True 
    else:
        # A seguir, tenta a correspondência na string de data DD-MM-AAAA
        date_match = regex_data1.search(date)
        # Se não correspondeu, tenta o padrão AAAA-MM-DD
        if not date_match:
            date_match = regex_data2.search(date)
            if date_match:
                century, year, month, day = date_match.groups()
        else:
            day, month, century, year = date_match.groups()
            logging.debug(f'date match: {day, month, century, year}')

    if not date_match and not time:
        return None
    if time:
        time_match = regex_time.search(time)
        if time_match:
            hour, minute, timezone = time_match.groups()
            logging.debug(f'time match: {hour, minute, timezone}')
    
    timedate = (day, hour, minute, timezone, month, century, year)
    
    return timedate

def formata_timedate(timedate):
    meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAIO', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    if not timedate or timedate == (None, None, None, None, None, None, None):
        return None, None, None, None
    day, hour, minute, timezone, month, century, year = timedate
    if century:
        full_year = f'{century}{year}'
    else:
        full_year = f'20{year}'
    newdate = f'{full_year}-{int(month):02d}-{int(day):02d}'  # Formatação de mês e dia com dois dígitos
    if hour:
        newtime = f'{hour}:{minute}:00{timezone if timezone else ""}'.strip()
        if century and century != '20':
            year = full_year
        gdh = f'{day}{hour}{minute}{timezone if timezone else ""}{meses[int(month) - 1]}{year}'
    else:
        newtime = None
        gdh = None
    return newdate, newtime, timezone, gdh
 
# Função geradora de estilos
def style_gen(valid_cats, cat_to_icon):
    if not esconde_nome:
        return None
    style_map = ''

    for cat in valid_cats:
        if cat_to_icon[cat].startswith('http'):
            icone = cat_to_icon[cat]
        else:
            icone = os.path.join(os.path.abspath('.'), 'icones', cat_to_icon[cat])
        mapa = f'''
            <Style id="{cat}_hide">
                <LabelStyle>
                    <scale>0</scale>
                </LabelStyle>   
                <IconStyle>
                    <Icon>
                        <href>{icone}</href>
                    </Icon>
                </IconStyle>
            </Style>
            <Style id="{cat}_show">
                <LabelStyle>
                    <scale>1.1</scale>
                </LabelStyle>
                <IconStyle>
                    <Icon>
                        <href>{icone}</href>
                    </Icon>
                </IconStyle>                
            </Style>
            <StyleMap id="{cat}_onoff">
                <Pair>
                    <key>normal</key>
                    <styleUrl>#{cat}_hide</styleUrl>
                </Pair>
                <Pair>
                    <key>highlight</key>
                    <styleUrl>#{cat}_show</styleUrl>
                </Pair>
            </StyleMap>'''
        style_map += mapa
    return style_map

# Função geradora de kml completo
def kml_gen(csv_path):
    save = True
    locais = []
    placemarks = {key: '' for key in cat_to_icon}
    placemarks.setdefault('Sem Categoria', '')
    cat_count = {key: 0 for key in placemarks}
    folder_path = os.path.dirname(csv_path)
    arquivo = os.path.basename(csv_path)
    nome_filtro = arquivo.replace('.csv', '_filtrado.csv')
    filtro_path = os.path.join(folder_path, nome_filtro)

    kml_header = f'''<?xml version="1.0" encoding="UTF-8"?>
    <kml xmlns="http://www.opengis.net/kml/2.2">
    <Document>
    <name>{arquivo}</name>'''

    kml_footer = '''</Document>
    </kml>'''

    assert csv_path.lower().endswith('.csv')

    if not os.path.isfile(csv_path):
        logging.info(f"O arquivo csv'{csv_path}' não foi encontrado.")
        return
    if os.path.getsize(csv_path) == 0:
        logging.info(f"O arquivo csv '{csv_path}' está vazio.")
        return
    
    try:
        df = ler_e_normalizar_csv(csv_path)
    except pd.errors.EmptyDataError:
        logging.info(f"O arquivo CSV '{csv_path}' não contém dados.")
        return
    except Exception as e:
        logging.info(f"Erro ao ler o arquivo CSV '{csv_path}': {e}")
        return
    
    headers = head_read(csv_path)

    # inicia a iteração para cada local marcado (cada linha da tabela)
    for i, row in df.iterrows():
        img_path = None

        # preparação dos dados de geolocalização
        if 'Coor' in headers and row['Coor'] != '' and str(row['Coor']) != 'nan':
            logging.info(f'Linha {i}: MODO 1')
            raw_coord = dividir(row['Coor'])
        elif 'Lat' in headers and 'Lon' in headers and str(row['Lat']) != 'nan' and str(row['Lon']) != 'nan':
            logging.info(f'Linha {i}: MODO 2')
            raw_coord = (row['Lat'], row['Lon'])
        elif 'Img' in headers and row['Img'] != '' and str(row['Img']) != 'nan':
            logging.info(f'Linha {i}: MODO 3')
            raw_img = row['Img']
            if raw_img.lower().endswith('.jpg') or raw_img.lower().endswith('.jpeg'):
                raw_img = raw_img
            else:
                raw_img = raw_img + '.jpg'
            img_path = os.path.join(folder_path, 'fotos', raw_img)
            coord = geoextract(img_path)
            logging.debug(f'{img_path}: {coord}')
            if coord:
                raw_coord = [0, 0]
                raw_coord[0] = gms2dec(coord[0])
                raw_coord[1] = gms2dec(coord[1]) 
            else:
                raw_coord = None         
        else:
            logging.info(f'Linha {i}: MODO 4')
            raw_coord = None
            logging.error(f'Linha {i}: Não foi possível obter dados de geolocalização para {csv_path}, linha {i}')  
        dec_coord = [0,0]
        try:
            dec_coord[0] = float(raw_coord[0])
            dec_coord[1] = float(raw_coord[1])
            logging.info(f'Linha {i}: Coordenadas de {csv_path} já eram decimais.')
        except Exception as e:
            logging.debug(e)
            logging.info(f'Linha {i}: Convertendo as coordenadas de {csv_path} por meio de Regex.')
            try:
                dec_coord[0] = capturar(raw_coord[0])
                dec_coord[1] = capturar(raw_coord[1])
            except Exception as e:
                logging.debug(e)
                dec_coord = None
        logging.info(f'Dec Coord: {dec_coord}')
        if dec_coord == None:
            logging.info(f'Linha {i}: As coordenadas fornecidas em {csv_path}, estão em um formato não reconhecido')
            continue
        
        lat = dec_coord[0]
        lon = dec_coord[1]
        # preparação das imagens(SFC)
        img = preimg(folder_path, i, row, headers, lat, lon)

        # Define a altura do placemark
        alt = 0
        if 'Alt' in headers:
            if row['Alt'] != '' and str(row['Alt']) != 'nan':
                alt = row['Alt']
        logging.debug(f'Altura definida para {alt}')

        # Define a categoria
        cat = 'Sem Categoria'
        if 'Cat' not in headers:
            if 'Decri' not in headers:
                logging.info(f'Não foi possível categorizar {csv_path}, linha {i}.')
            else:
                cat = autocat(str(row['Decri']), cat_filter)
        elif row['Cat'] != '' and str(row['Cat']) != 'nan' and row['Cat'] in cat_to_icon.keys():
                cat = row['Cat']
        cat_count[cat] += 1
        logging.debug(f'Categoria definida como {cat}')

        esconde_nome = cat_hide[cat]

        # Define o nome do ponto
        nome = f'Ponto {i + 1}'
        if 'Nome' not in headers:
            if cat != 'Sem Categoria':
                nome = f'{cat} {cat_count[cat]}'
        elif row['Nome'] != '' and str(row['Nome']) != 'nan':
            nome = row['Nome']
            
        # Define a descrição    
        decri = row.get('Decri', cat)
        if str(decri) == 'nan':
            decri = ''
        lat_gms = dec2gms(lat, 'lat')
        lon_gms = dec2gms(lon, 'lon')

        # Define a data-hora
        date, time, timezone, gdh = None, None, None, None
        raw_date = row.get('Data', None)
        raw_time = row.get('Hora', None)
        timedate_full = captura_timedate(raw_date, raw_time, img_path)
        date, time, timezone, gdh = formata_timedate(timedate_full)
        
        # Registra o local
        locais.append((nome, cat, lat, lon, lat_gms, lon_gms, alt, decri, date, time, gdh, img))

        # Gera o placemark do local e adiciona ao kml da categoria
        kml = kml_placemark_gen(nome, cat, lat, lon, decri, img, alt, date, time, timezone)
        placemarks[cat] += kml
    #gera as pastas e monta o kml
    kml, valid_cats = kml_folders_gen(placemarks)
    estilo = ''
    if esconde_nome:
        estilo = style_gen(valid_cats, cat_to_icon)
    kml = kml_header + estilo + kml + kml_footer
    cont_locais = len(locais)
    logging.debug(f'Foram registrados {cont_locais} locais: {locais}')
    # Caso nenhum local tenha sido registrado
    if cont_locais == 0:
        logging.error('O kml não pôde ser gerado, devido à inexistência de pontos a serem marcados')
        save = False

    gerar_filtro(locais, filtro_path)

    if not save:
        return
    return kml

# Função para salvar o arquivo kml
def kml_save(kml, kml_path):
    if os.path.exists(kml_path):
        logging.info(f'O arquivo {kml_path} já existe. Remova-o se desejar gerá-lo novamente.')
        print(f'O arquivo {kml_path} já existe. Remova-o se desejar gerá-lo novamente.')
        return
    with open(kml_path, 'w', encoding='UTF-8') as f:
        f.write(kml)
        logging.info(f'Arquivo KML gerado: {kml_path}')
        print(f'Arquivo KML gerado: {kml_path}')

# Função para mapear os arquivos nas pastas 'fotos'
def cria_geo_metadata(root_dir='.'):
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if 'fotos' in dirnames:
            fotos_path = os.path.join(dirpath, 'fotos')
            csv_file_path = os.path.join(fotos_path, 'metadados_geo.csv')
            
            with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(['Imagem', 'Meta_Geo'])  # Cabeçalho
                for filename in os.listdir(fotos_path):
                    if filename.lower().endswith(('.jpg', '.jpeg', '.png')):  # Verifica extensões de imagem
                        file_path = os.path.join(fotos_path, filename)
                        has_geo = geoextract(file_path)
                        writer.writerow([filename, 'Sim' if has_geo else 'Não'])
            
            logging.info(f'Arquivo CSV gerado em: {csv_file_path}')

# Função principal
def main():
    log_start()
    criar_arquivo_bat()
    cria_estrutura()
    setup()
    cria_geo_metadata()
    for pasta in os.listdir('.'):
        caminho = os.path.join('.', pasta)
        if pasta == 'icones' or pasta == 'venv' or not os.path.isdir(caminho):
            continue
        arquivos = os.listdir(caminho)

        # Verifica se há arquivos com a extensão '.csv' e que não terminam com '_filtrado.csv'
        arquivos_csv = [arquivo for arquivo in arquivos if arquivo.endswith('.csv') and not arquivo.endswith('_filtrado.csv')]

        if not arquivos_csv:
            logging.info(f"Na pasta '{pasta}', não há arquivos '.csv'.")
            continue

        for arquivo in arquivos_csv:
            caminho_csv = os.path.join(caminho, arquivo)
            arquivo_kml = arquivo.split('.')[0] + '.kml'
            caminho_kml = os.path.join(caminho, arquivo_kml)
            kml = kml_gen(caminho_csv)
            if kml == None:
                logging.error(f'Não foi possível gerar o arquivo kml a partir de {arquivo}')
                continue
            kml_save(kml, caminho_kml)

    print('\nMissão cumprida!\n\nConsulte \'log.txt\' para detalhes.\t\t\t"ADTI"\n')

if __name__ == "__main__":
    main()    