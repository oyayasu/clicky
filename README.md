# Clicky — editor de textos open source

**v0.5.0** — Aplicativo local para Windows e macOS, feito com Tauri 2, com biblioteca de documentos, guias, outline por títulos, comentários, importação/exportação de DOCX, fontes e imagens.

## Iniciar em modo desenvolvimento

```bash
npm install
npm run dev
```

## Testar e compilar

```bash
npm test
npm run build
npm run tauri build
```

Os executáveis também são compilados em `.github/workflows/build.yml` (GitHub Actions). Artifacts para Windows (`exe`/`msi`) e macOS (`dmg`) ficam disponíveis no workflow.

## Mudanças da v0.5.0

- Seleção da esquerda para a direita: aplica destaque, se Grifo rápido estiver ativado. Da direita para a esquerda: remove destaque.
- Botão direito no destaque: menu contextual com **Excluir, Destacar, Consultar, Comentar, Copiar**; a cor pode ser escolhida em uma paleta.
- Notas são comentários. Os comentários são salvos na biblioteca local, não são colaboração em tempo real.
- Outline dinâmico dentro da guia ativa: títulos de nível 1 a 6 aparecem sob a guia e permitem navegar no documento.
- Caixa de confirmação, entrada de URL e renomeação agora usam diálogos próprios em vez de janelas de navegador.
- Menu contextual da biblioteca oferece Abrir DOCX, Nova pasta e Novo documento.
- Tela inicial mostra os controles de minimizar/maximizar/fechar no desktop.
- Interface oferece Português e Inglês nas funções principais; partes secundárias ainda têm rótulos somente em português.
- Temas continuam disponíveis como opções predefinidas; o editor, importação e exportação de temas personalizados foram removidos.

**Limitações conhecidas:** A fidelidade DOCX não é integral (principalmente layouts complexos, flutuação de imagens e comentários). Documentos locais são salvos automaticamente com IndexedDB; use o backup `.clicky` para preservar a estrutura completa. As compilações do GitHub não substituem testes manuais de instalação e interação.

## Referências de comportamento

- [Guias de documento](https://support.google.com/docs/answer/15499791?hl=pt-BR)
- [Estrutura de tópicos](https://support.google.com/docs/answer/6367684?hl=pt-BR)
- [Atalhos](https://support.google.com/docs/answer/179738?hl=pt-BR)

Licença MIT. Clicky e Docses são projetos independentes.