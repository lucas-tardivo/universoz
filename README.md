# Universo Z Online
Universo Z foi um MMORPG baseado em Dragon Ball Z criado por boasfesta e Pirata_254 em 2013 como uma homenagem por fãs. Este repositório é sua engine original criada em VB6 através da distribuição Eclipse Advanced.
A sua publicação tem como intuito contribuir com a comunidade de desenvolvimento indie de jogos de forma acadêmica e também permitindo todas as regras compostas na licença MIT.

## Considerações
Este projeto foi desenvolvido em 2013~2014 e permaneceu em desenvolvimento de maneira intermitente ao longo dos anos, fazendo parte inclusive de mais de um projeto e objetivo. Portanto, **NÃO** possui padrão de desenvolvimento ou obrigação quanto á qualidade de código. As atuais correções de bugs e otimizações de código estão á mercê da comunidade na qual este projeto possui e são muito bem-vindas via Pull Request neste repositório.

## Contribuições
Para contribuir publicamente com este projeto, crie um fork do repositório e abra um Pull Request descrevendo a modificação efetuada.

## Obtendo o projeto
### Código-fonte
Clone o repositório utilizando o aplicativo **Github Desktop**. Evite baixar o código por ZIP uma vez de que O VB6 possui problemas quanto á codificação após a compressão de arquivos.

### Jogo compilado
Baixe uma versão oficial da página de **Releases** do repositório.

## Estrutura do projeto
### Cliente
O cliente se encontra na pasta "Cliente" e possui duas versões, sendo elas: Cliente e Suite. O cliente possui toda a experiência de jogo para os jogadores, enquanto o Suite permite uma experiência limitada de jogo para administração/desenvolvimento interno do jogo.

### Servidor
O servidor se encontra na pasta "Servidor" e possui acesso ao WebService que permite sua interação com um web site. Este WebService é responsável pelo cadastro de novas contas, integração com a loja e sincronia dos rankings.

### Utilitários
Nesta seção é possível ter acesso aos utilitários e seus respectivos códigos-fonte. Entre eles estão o serviço de criptografia de gráficos, driver de conexão com MySQL, arquivos necessários, editor de multilinguagem, conector do WebService e editor manual de personagem. 

## Utilização
- Abra o servidor localizado em Servidor/Universo Z Server.exe
- Abra o WebManager.exe
- Clique em Comandos > Add Manual
- Preencha os campos e crie uma nova conta
- Abra o jogo em Cliente/Universo Z.exe
- Acesse a conta com a credenciais preenchidas no WebManager
- Crie seu personagem
- Feche o jogo
- Abra o editor de personagem em Servidor/Editor de Personagens.exe
- Insira o nome da sua conta no campo e clique em Carregar
- Altere o acesso para 15
- Clique em salvar
- Abra o editor em Cliente/Universo Z Suite.exe
- Acesse sua conta
- Edite o jogo como preferir

## Licença
Este projeto acompanha as regras da Licença MIT: https://pt.wikipedia.org/wiki/Licença_MIT  
Para a distribuição com fins lucrativos á partir deste, é **obrigatório a adição dos devidos créditos** ao repositório do projeto e seus respectivos criadores.

## Sistemas inclusos
- DirectX 8
- Gráficos criptografados
- Animação com tremor na tela
- Efeitos de transformação melhorada
- Transformação com cabelos customizáveis
- Kaioken (RGB automático)
- Oozaru
- Sistema de partículas
- Efeito visuais de buracos na tela
- Scouter com efeito visual
- Movimento diagonal
- Evolução de skills
- Desafio diário
- Efeito de tela de quase morte e flash de dano
- Antihack integrado
- Multiservidor
- Multilinguagem
- Títulos com ícones
- Itens animados no inventário
- Item com bonus de EXP temporário
- Shenlong
- Casas
- Conquistas
- Quests
- Transportes dinâmicos (Barco e avião)
- Pesca
- VIP com evolução
- Sistema de voo
- Efeito flutuante de voo para jogadores e NPCs
- Customização de personagem
- Menu interativo com animação
- 22 animações de personagem
- Bonus diário
- Máquina de gravidade para AFK
- Level de divindade (duas evoluções)
- Eventos globais
- Suporte in-game e feedbacks
- Caixas surpresa com itens fixos ou aleatórios
- Editor de personagem completo
- Habilidades especiais
  - Invocação de lacaios
  - Chuva de meteoros
  - Magia linear com rotação automática e particionada (Base, Corpo, Cabeça)
  - Transformação de NPCs
  - Empurrão de inimigos
  - Teleporte para as costas
- Guilds
  - Ícone customizável
  - Banco da Guild
  - Evolução da guild com EXP e Level
  - Notícia diária
- Integração com Website
  - Cadastro
  - Compras online
  - Tabela de rankings
- Ambientação dos mapas
  - Animais (Pássaros, Morcegos e Gaivotas)
  - Nuvens

## Créditos
#### Programação
- boasfesta
#### Apoio e desenvolvimento
- Pirata_254
