# 🛡️ Mesa de Inteligência Forense - LCFO

## Mapa da Arquitetura (Quem faz o quê?)

* **app.py (A Vitrine / A Tela)**
  - Responsabilidade: Apenas desenhar a interface visual no navegador.
  - O que fica aqui: Botões do Streamlit, abas (`tabs`), cabeçalho HUD, seletores da barra lateral e PyDeck (mapa).
  - Regra de ouro: Nenhuma limpeza de texto ou cálculo pesado de SQL deve ser escrito aqui.

* **database.py (O Almoxarifado / O Banco)**
  - Responsabilidade: Conversar com os arquivos do Excel e com o SQLite.
  - O que fica aqui: Caminhos de rede (`T:\...`), criação das tabelas, comandos `INSERT/SELECT`, e a função de cruzamento entre a Blacklist e as Criações Diárias.
  - Regra de ouro: Se precisar criar uma nova coluna ou ler uma pasta nova, mexe apenas aqui.

* **graph_engine.py (O Cérebro Relacional)**
  - Responsabilidade: Inteligência de redes e montagem visual do grafo.
  - O que fica aqui: NetworkX (cálculo de hubs, pontes com betweenness, modularidade de comunidades) e o script Vis.js (HTML com os layouts em grade).
  - Regra de ouro: Tudo o que envolve matemática de redes e desenho de linhas/nós fica isolado aqui.

* **utils.py (A Caixa de Ferramentas)**
  - Responsabilidade: Saneamento e tratamento de dados brutos.
  - O que fica aqui: Limpeza de CPF, máscara de telefone, validação de placa de 7 dígitos, correção de ano (`0026` -> `2026`) e coordenadas de cidades.
  - Regra de ouro: Funções "puras" que apenas recebem um texto sujo e devolvem um texto limpo.