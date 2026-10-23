# Clicky 0.2.0

![Ícone oficial do Clicky](src-tauri/icons/icon.png)

**Editor de textos open source, offline-first, para Windows e macOS.** Layout inspirado no fluxo do Google Docs, aparência minimalista com cantos arredondados. Licença MIT.

![Biblioteca de documentos](preview.png)

## Nesta versão

### Biblioteca de projetos

- Tela inicial com documentos independentes (não substitui o documento anterior ao criar outro).
- Pastas coloridas (criar, renomear, mudar cor, excluir preservando os documentos).
- Visualização em **grade** ou **lista**, busca por título, renomear, mover, duplicar e excluir documentos.
- Salvamento automático local em **IndexedDB**, índice separado dos documentos para manter a listagem leve.
- Migração automática do projeto único do Clicky 0.1, desde que aberto com o **mesmo armazenamento do aplicativo/origem**.
- Backup por documento `.clicky` e backup de toda a biblioteca `.clicky-library` (botão `⋯` na tela inicial), com restauração **sem apagar** os documentos já existentes. Os backups podem conter imagens e fontes; guarde-os em local seguro.

### Editor e organização

- Menus no alto, barra de formatação, régua, página e guias na lateral, como no protótipo 0.1.
- Guias internas de um documento (por exemplo: Original, Remake, English), com duplicação, renomeação e exclusão.
- Destaque de texto após seleção, ativável/desativável, com diversas cores; menu de contexto para editar e limpar.
- Comentários **locais** ancorados em trechos selecionados, inclusive responder, resolver, reabrir, excluir e ver histórico na barra à direita; `Ctrl+Alt+M` no Windows ou `⌘+Option+M` no Mac.
- Mecanismo nativo de verificação ortográfica (`spellcheck`) configurado por idioma pt-BR/en-US/es-ES/ja-JP. **A qualidade e disponibilidade das sugestões são determinadas pelo WebView e pelo sistema operacional.**
- Revisão rápida local de alguns erros comuns em português e de palavras repetidas; **não é um corretor gramatical completo**.
- Dicionário pessoal para palavras aceitas pelo revisor rápido. A consulta de significado é feita **mediante clique** no Wikcionário e requer internet; palavras pessoais **não são integradas ao dicionário nativo do sistema**.
- Atalhos de edição, formatação, alinhamento, comentários, pesquisa, dicionário e exportação. Lista pesquisável em **Ctrl/Cmd + /**; menu de comandos em Alt + /. Ver `docs/ATALHOS.md`.
- Arquivos de fontes próprias (`.ttf`, `.otf`, `.woff`, `.woff2`), além de fontes disponíveis no sistema.
- Imagens e GIFs livres com arrastar, redimensionar, cadeado e magnetismo opcional, como na v0.1.
- Temas prontos, personalização de cor e fundo, exportação/importação de temas.

### Exportações

- `.clicky`: **formato próprio de preservação**, com guias, comentários, anexos, fontes e suas posições.
- `.clicky-library`: backup de **todos** os documentos e pastas do dispositivo.
- `.docx`: Word/OOXML editável; **posicionamento livre de imagens e comentários não são exportados**. As guias viram seções sequenciais e imagens são colocadas ao final de cada guia.
- `.txt`, `.html` e impressão/salvar PDF da guia ativa.

## Google Drive: etapa opcional, ainda não implementada

Esta versão **não tem login Google**, leitura de `.docx` diretamente do Drive, sincronização automática nem colaboração em tempo real. Nenhuma tela de login fictícia é apresentada.

Para uma integração real, o plano de implementação é usar **OAuth 2.0 de aplicativo instalado**, autorização de arquivos selecionados com o escopo `drive.file` e Google Picker, importação fiel de `.docx`, identificação de revisão/ETag e resolução de conflitos antes de sobrescrever cópias remotas. Nunca guardar senha Google no aplicativo. O Google Drive não fornece edição colaborativa automática para qualquer editor externo. Consulte:
- https://developers.google.com/workspace/drive/api/guides/api-specific-auth
- https://developers.google.com/workspace/drive/picker/guides/desktop-mobile-picker

Para mover documentos hoje entre computadores, exporte o `.clicky-library`, copie o arquivo e importe no outro computador; **isso não sincroniza alterações automaticamente**.

## Teste no navegador

Requer Node.js 20+:

```bash
npm run dev
```

Abra `http://127.0.0.1:5173`. O modo navegador e o desktop Tauri possuem armazenamentos locais **separados**. Não limpe os dados do aplicativo/navegador sem exportar backup.

## Compilar desktop

Requer Node.js 20+, Rust, WebView2/MSVC no Windows e Xcode Command Line Tools no macOS. Veja https://v2.tauri.app/start/prerequisites/

```bash
npm install
npm run tauri dev
npm run tauri build
```

Também existe `.github/workflows/build.yml` configurado para Windows e macOS Intel/Apple Silicon. Faça o upload de todos os arquivos da pasta do projeto para um repositório GitHub, acesse **Actions → Clicky - Windows e macOS → Run workflow** e consulte **Artifacts**. Não há binários assinados; os sistemas podem exibir avisos.

## Limitações técnicas e próximos passos

- O editor ainda utiliza `contenteditable` e `document.execCommand`. Para documentos longos, paginação correta e fidelidade DOCX, o núcleo deve migrar para ProseMirror/Tiptap ou equivalente.
- Comentários nesta versão são **locais e individuais**; não há colaboração nem rastreamento de revisões entre usuários.
- Verificação gramatical ampla, importação DOCX de terceiros, Drive e edição colaborativa são **pendentes**.
- Formatação complexa, GIF animado, transparências, objetos flutuantes e comentários não têm correspondência fiel no DOCX exportado.
- O projeto não foi compilado nem testado nativamente nos instaladores Windows/macOS neste ambiente. Foram executados `npm test`, `npm run build` e testes da interface Chromium em ambiente de testes isolado; o salvamento real do IndexedDB deve ser conferido no instalador.
- Faça backups periódicos: mesmo com autosave, o navegador/sistema pode remover o armazenamento local.

## Estrutura

```text
Clicky/
├── index.html                  # Editor, biblioteca e painel de comentários
├── src/app.js                 # Comandos, biblioteca, comentários, revisão, UI
├── src/storage.js             # IndexedDB: docs, índice, pastas, versões
├── src/docx.js                # Exportador OOXML
├── src/style.css              # Interface e temas
├── src-tauri/                 # Desktop Windows/macOS
├── scripts/                   # Servidor local e build
├── docs/ATALHOS.md            # Atalhos atuais e limitações
├── tests/                     # Testes automatizados do projeto
└── .github/workflows/build.yml
```

Licença: [MIT](LICENSE).

## Instalar e compilar

O ícone oficial utiliza `src-tauri/icons/clicky-master.png`. Gere os formatos específicos antes de executar o build local:

```bash
npm install
npm run tauri -- icon src-tauri/icons/clicky-master.png
npm run tauri build
```

Os instaladores para Windows x64, macOS Apple Silicon e macOS Intel são criados pelo workflow em `.github/workflows/build.yml`. Baixe-os em **Actions > último run > Artifacts**. Builds sem assinatura podem exigir confirmação de segurança no primeiro acesso.
