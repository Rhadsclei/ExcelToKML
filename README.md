# <h1 align=center>  ExcelToKML </h1>

## Criado á partir da documentação do simplekml: https://simplekml.readthedocs.io/en/latest/index.html
## Criado á partir da documentação do polycircles: https://polycircles.readthedocs.io/en/latest/getting_started.html
<strong> Cria círculos (ou polígonos), marcadores, links P2P e múltiplos links P2P a partir de coordenadas nesta pasta de trabalho do Excel.  </strong>

[Creates circles (or polygons), markers, P2P links and multiples P2P links from coordinates on this Excel workbook.]

## Conteúdo
- [Requisitos](#requisitos)
- [Estrutura](#estrutura)
- [Planilhas presentes](#planilhas presentes)

## Requisitos:
       Habilitar macros.
       Python
       Executar o arquivo em lotes "Dependencies.bat" para:
              Instalar plugin "simplekml".
              Instalar plugin "polycircles"
              Instalar plugin "numpy"
## Estrutura:
![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/30896b82-4aed-4818-8cb3-e0af086f5146)

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




## Planilhas presentes: 

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/a492a5b5-c8b2-41de-ac76-b85e73567c53)

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
       
Parâmetros de entrada:
![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/67459519-17df-4486-bf09-e112277dbcca)
       
Arquivos gerados na pasta "Python Scripts":

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/df0a9fa3-ce37-491f-ac39-732c03e5059c)

Arquivos gerados na pasta "KML":

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/d94d3bf2-885b-4d13-a078-ccf433a00346)

       
Visualização:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/20f0ba5f-5e18-4d35-976b-e2be86c9a087)


Visualização com marcadores, polígonos de 3 e 4 lados com 50 de altura e link P2P:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/ca9a85c4-315a-4365-a4a1-8f868aafe681)


       Points:
       Cria marcadores (pontos) á partir das corrdenadas (latitude e longitude)
       Nome: Identificação do marcador.
       Latitude: Coordenada de laitude do marcador.
       Longitude: Coordenada de longitude do marcador.
       Escala do rótulo: Tamanho do rótulo.
       Escala do ícone: Tamanho do ícone.
       Selecione o Grupo: Grupo do marcador. Edite esta opção de acordo com sua necessidade, mas deverá existir um ícone (em formato .png) na pasta "Groups" com o mesmo nome do grupo selecionado. Edite os grupos em "Dados -> Validação dos Dados -> Configurações".
Parâmetros de entrada:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/c47abde4-4694-449b-bb4f-ab1f6e988a0a)

Arquivos gerados na pasta "Points":

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/78785434-98ae-405c-b50b-a2ac5f0cd99d)

Visualização:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/a9ce2dbe-9a61-4366-b1ce-ce6ff5e1ce54)


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
       
Parâmetros de entrada:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/e7312a27-7c07-467b-bdb8-6c40585a6ff8)


Arquivos gerados na pasta "P2PPlanning":

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/fb94e41f-2963-4cd6-9cf6-9813191815e0)


Visualização:

![image](https://github.com/Rhadsclei/ExcelToKML/assets/143188137/2996f2e8-bac2-4525-b390-fd8fdacec6f9)


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


