# ExcelToKML

Criado á partir da documentação do simplekml: https://simplekml.readthedocs.io/en/latest/index.html#
Criado á partir da documentação do polycircles: https://polycircles.readthedocs.io/en/latest/getting_started.html

Cria círculos (ou polígonos), marcadores,links P2P e múltiplos links P2P a partir de coordenadas nesta pasta de trabalho do Excel.  
[Creates circles (or polygons), markers, P2P links and multiples P2P links from coordinates on this Excel workbook.]

Requisitos:
Habilitar macros.
Python
  Executar o arquivo em lotes "Dependencies.bat" para:
    Instalar plugin "simplekml".
    Instalar plugin "polycircles"
    Instalar plugin "numpy"
Estrutura:
ExcelToKML --
            |-- KML: Pasta onde ficarão os arquivos de polígonos criados pela planilha "ExcelToKML".
            |-- MP2PPlan: Pasta repositório dos arquivos relativos á esta funcionalidade.
                 |-- Exec.bat: Arquivo em lotes que executará o arquivo "MP2PPlanning.py" que será criado neste diretório em tempo de execução.
                 |-- MP2P.png: Imagem do marcador do concentrador de múltiplos links P2P (unidade local).
                 |-- MP2PPlan.kml: Arquivo KML criado pela execução do arquivo "MP2PPlanning.py" quando o "Exec.bat" é executado.
                 |-- MP2PPlanning.py: Arquivo .py criado em tempo de execução contendo as intruções para geração do arquivo de múltiplos links P2P.
                 |-- P2P.png: Imagem do marcador de cada unidade remota da infraestrutura.
                 |-- Os arquivos KML desta funcionalidade serão exibidos neste nível.
            |-- P2PPlan: Pasta repositório dos arquivos relativos á esta funcionalidade.
                 |-- Exec.bat: Arquivo em lotes que executará o arquivo "#P2Planning.py" que será criado neste diretório em tempo de execução.
                 |-- #P2PPlanning.py: Arquivo .py criado em tempo de execução contendo as intruções para geração dos arquivos de links P2P.
                 |-- P2P.png: Imagem do marcador de cada link da infraestrutura.
                 |-- Os arquivos KML desta funcionalidade serão exibidos neste nível.
            |-- Points: Pasta repositório de cada marcador gerado e dos arquivos relativos á esta funcionalidade.
                 |-- Groups: Pasta repositório de cada ícone de cada grupo previsto da funcionalidade "Points".
                 |-- Exec.bat: Arquivo em lotes que executará o arquivo "Points.py" que será criado neste diretório em tempo de execução.
                 |-- Points.py: Arquivo .py criado em tempo de execução contendo as intruções para geração do arquivo contendo cada um dos pontos definidos nesta funcionalidade.
                 |-- Os arquivos KML desta funcionalidade serão exibidos neste nível.
            |-- Python Scripts: Pasta repositório dos arquivos relativos á esta funcionalidade.
                |-- Exec.bat: Criado em tempo de execução, este arquivo executará cada arquivo ".py" criado pela macro.
                |-- Os arquivos .py desta funcionalidade serão exibidos neste nível e executados pelo arquivo "Exec.bat". Os arquivos de saída serão gerados na pasta um nível acima "KML"
            |-- Dependencies.bat: Arquivo em lotes que deverá ser executado uma única vez para a instalação dos plugins "simplelml, polycircles e numpy".
            |-- ExcelToKML: Pasta de trabalho contendo as planilhas com as funcionalidades deste projeto.
![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/bb9b33d0-442c-4a4d-aba4-85d3958bf7cd)

Planilhas presentes:            
XLSTOKML:
Esta planilha cria um polígono com X lados e Y raio á partir de coordenadas (latitude e longitude). Há a opção de personalizar a cor da linha, de preenchimento do polígono, do rótulo, da altitude, da perspectiva e de alongamento até o solo.
  Nome: Identificação do nome do polígono.
  Latitude: Coordenada da latitude do centro do polígono.
  Longitude: Coordenada da longitude do centro do polígono.
  Raio (m): Raio do polígono á partir do centro do mesmo.	
  Número de Vértices: Número de lados do polígono (mínimo de 3 lados).
  Cor do Polígono: Cor do preenchimento do polígono.	
  Cor da Linha: Cor da linha do polígono.
  Tamanho da Linha: Espessura da linha do polígono.
  Escala do rótulo: Tamanho do rótulo.
  Descrição ⚠: Descrição do polígono
  Altitude (m) e Topologia: Altitude do polígono em relação ao solo (inteiro e em metros) seguido da topologia do polígono (preso, absoluto ou relativo ao solo).
  Alongar até o solo? Sim/Não: Opção de alongar traçado do polígono até o solo.
  Formato de Saída: Descrição - Nome
![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/602ce66a-727a-4ccf-9cfa-9e8adad6c822)
Dois pontos ao redos dos quais serão gerados os polígonos com 600 lados, a título de exemplo.


![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/3b5d9d1e-44e6-4e7a-be73-2e854c10eaca)
Arquivos gerados em tempo de execução na pasta "Python Scripts".


![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/a3e00b22-46d0-4b39-87c3-35460b67b056)
Arquivos gerados em tempo de execução na pasta "KML".


![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/88d042f3-9a28-41da-9602-d51ac2f19947)
Note que os polígonos acompanham o relevo do local. Esta imagem indica que o lado direito do polígono da RM-01 fica parcialmente bloqueado por causa disso.
A RM-02 fica apenas com parte do seu lado esquerdo bloquado por uma árvore.


![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/4ad16b05-3f3c-4868-b380-5c1cea8e55cc)
Observe a perspectiva do polígono á partir da altitude estipulada.


![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/9bf9c430-b8a5-4741-93ad-6e5f1f4a248b)
Alterando a quantidade de vértices, este é o resultado.


Points:
  Cria marcadores (pontos) á partir das corrdenadas (latitude e longitude)
  Nome: Identificação do marcador.
  Latitude: Coordenada de laitude do marcador.
  Longitude: Coordenada de longitude do marcador.
  Escala do rótulo: Tamanho do rótulo.
  Escala do ícone: Tamanho do ícone.
  Selecione o Grupo: Grupo do marcador. Edite esta opção de acordo com sua necessidade, mas deverá existir um ícone (em formato .png) na pasta "Groups" com o mesmo nome do grupo selecionado. Edite os grupos em "Dados -> Validação dos Dados -> Configurações".

P2PPlanning:
Cria um link P2P ou linha entre dois pontos, estando cada um á uma determinada altitude, possibilitando avaliar se há visada entre ambos.
  Unidade local: Arbitrário, mas é o primeiro ponto do link.
    Latitude: Coordenada de latitude do ponto da unidade local ou de referência.
    Longitude: Coordenada de longitude do ponto da unidade local ou de referência.
    Altitude: Altitude do ponto da unidade local ou de referência.
  Unidade remota:
    Latitude: Coordenada de latitude do ponto da unidade remota.
    Longitude: Coordenada de longitude do ponto da unidade remota.
    Altitude: Altitude do ponto da unidade remota.
  Escala dos marcadores: Tamanho dos marcadores do link.
  Escala dos rótulos: Tamanho dos rótulos dos marcadores do link.
  Escala da linha: Espessura da linha do link.
  Cor (da linha): Cor da linha do link.

MP2PPlanning:
  Cria múltiplos links entra a "Unidade Local" e as "Unidades Remotas", estando a unidade local em uma determinada altitude e as remotas á altitudes diferentes entre si ou iguais.
  Unidade Local: Nome do concentrador de links P2P.
    Latitude: Coordenada de latitude do ponto da unidade local ou de referência.
    Longitude: Coordenada de longitude do ponto da unidade local ou de referência.
    Altitude: Altitude do ponto da unidade local ou de referência.
  Unidades Remotas: 
    Latitude: Coordenada de latitude do ponto da unidade remota.
    Longitude: Coordenada de longitude do ponto da unidade remota.
    Altitude: Altitude do ponto da unidade remota.
  Escala dos marcadores: Tamanho dos ícone dos marcadores.
  Escala dos rótulos: Tamanho dos rótulos dos marcadores.
  Escala da linha: Espessura da linha do link.
  Cor (da linha): Cor da linha do link.

Auxiliar:
  Relação das cores utilizadas pelos polígonos e linhas.


