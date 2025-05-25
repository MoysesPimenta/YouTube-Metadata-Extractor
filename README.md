# YouTube Playlist Extractor

Um projeto simples para extrair dados e capturas de tela de playlists do YouTube.

## Estrutura do Projeto

Vou criar uma aplicação web moderna e clean que permitirá:
1. Inserir o link de uma playlist do YouTube
2. Extrair automaticamente os dados de cada vídeo
3. Gerar uma planilha Excel com os dados
4. Criar um documento Word com capturas de tela para comprovação
5. Baixar os arquivos gerados

## Arquitetura

A aplicação terá duas partes principais:

1. **Frontend**: Interface de usuário moderna e responsiva
   - HTML/CSS/JavaScript puro (sem frameworks pesados para facilitar GitHub Pages)
   - Design minimalista e profissional
   - Formulário para entrada do link da playlist
   - Visualização do progresso
   - Download dos arquivos gerados

2. **Backend**: Processamento dos dados (via Worker)
   - Como o GitHub Pages só suporta conteúdo estático, usaremos uma abordagem de worker
   - O processamento será feito no navegador do usuário
   - Utilizaremos APIs públicas do YouTube para obter os dados
   - Bibliotecas JS para geração de Excel e Word

## Considerações Técnicas

- Sem dependências de servidores externos
- Código totalmente aberto e documentado
- Design responsivo para desktop e mobile
- Tratamento de erros e feedback ao usuário
- Privacidade dos dados (processamento local)

## Próximos Passos

1. Criar a estrutura básica do projeto
2. Implementar a interface de usuário
3. Desenvolver a lógica de extração de dados
4. Integrar geração de Excel e Word
5. Testar e otimizar
