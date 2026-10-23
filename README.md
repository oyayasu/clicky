# Clicky — Editor de documentos

Aplicativo desktop de edição local em desenvolvimento, inspirado na disposição do Google Docs, com identidade de cor rosa. Open source (MIT). Repositório exclusivo do Clicky; não modifica o projeto Docses.

## O que funciona na versão 0.3.0

- Biblioteca local com pastas coloridas, pesquisa, grade e lista, miniaturas de texto sem logotipos.
- Cabeçalho, menus, régua, página A4 de 816 × 1056 CSS px e barra de formatação reorganizados para lembrar o Google Docs.
- Guias independentes por documento, comentários ancorados com respostas e resolução, destaque automático tipo Kindle, editor rico, fontes locais/importadas, ortografia do WebView e dicionário pessoal.
- Abrir arquivos `.docx` por **Arquivo > Abrir** ou pela tela inicial; o importador local extrai texto, estilos comuns, títulos, alinhamentos, tabelas e imagens PNG/JPEG/GIF/WebP quando reconhecidas. Importação é aproximada: **não preserva com fidelidade integral layouts complexos, caixas de texto, rodapés, imagens ancoradas ou numeração avançada**.
- **Salvar como** em `.docx`, `.txt` ou `.clicky`, usando o seletor nativo de pasta/arquivo no Tauri. No navegador, usa `showSaveFilePicker` quando disponível; o fallback é o download padrão do navegador. `Ctrl/Cmd+S` abre Salvar como DOCX; `Ctrl/Cmd+Shift+S` salva backup `.clicky`.
- Salvamento automático local em IndexedDB continua ativo para recuperação. Escolher um arquivo com Salvar como cria uma cópia no caminho desejado, mas **não sincroniza automaticamente futuras alterações ao mesmo arquivo externo**.
- GIFs e imagens móveis, ícones vetoriais de cadeado aberto/fechado, magnetismo agressivo e guias horizontais e diagonais durante o arraste.
- Ícone do aplicativo rosa com fundo transparente, sem azulejo branco ao redor.

## Compilação

Requer Node.js 22 e Rust/Tauri 2.

```sh
npm install
npm test
npm run dev
npm run tauri build
```

Para gerar instaladores para Windows x64, macOS Intel e macOS Apple Silicon sem instalar o SDK localmente, abra **GitHub > Actions > Compilar Clicky - Windows e macOS** e execute o workflow. Os artefatos `.exe`, `.msi`, `.dmg` estarão na execução concluída.

## Restrições atuais

Não é uma cópia do motor de layout do Google Docs: alguns itens de menu e formatos avançados ainda não têm paridade funcional. DOCX é importação/exportação com perdas de layout; os comentários e anexos posicionados livremente são preservados de forma completa no backup `.clicky`, **não no DOCX**. Integração Google Drive e colaboração simultânea ainda não estão implementadas. Os instaladores não são assinados nem notarizados.

O nome do aplicativo só é apresentado em janelas/sobre/documentação, evitando marcas d'água nas páginas ou miniaturas.