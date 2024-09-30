

    ____  __._____.___.__________ ________    _______     ________________
    |    |/ _|\__  |   |\______   \\_____  \   \      \   /   __   \_____  \ 
    |      <   /   |   | |       _/ /   |   \  /   |   \  \____    / _(__  < 
    |    |  \  \____   | |    |   \/    |    \/    |    \    /    / /       \
    |____|__ \ / ______| |____|_  /\_______  /\____|__  /   /____/ /______  /
            \/ \/               \/         \/         \/                  \/  3.0
                                                                 csv -> kml


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
    
