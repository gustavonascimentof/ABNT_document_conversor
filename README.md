# 📄 Conversor ABNT

Aplicação web que transforma documentos `.docx` (padronizados ou não) em um modelo formatado de acordo com as normas da ABNT — ou permite começar um trabalho acadêmico do zero, já com a estrutura ABNT pronta.

🔗 **Acesse:** [abnt-document-conversor.vercel.app](https://abnt-document-conversor.vercel.app/)

## O problema

Formatar um trabalho acadêmico segundo a ABNT — margens, fonte, espaçamento, capa, folha de rosto, sumário, citações, referências — é repetitivo e cheio de detalhes fáceis de errar. Este projeto automatiza essa formatação: você envia um `.docx` já estruturado com os estilos nativos do Word (Título 1, Título 2, Citação), e a aplicação devolve um documento pronto, com todas as regras aplicadas.

## Funcionalidades

- **Conversão automática**: lê um `.docx` enviado, identifica títulos, citações e a seção de referências a partir dos estilos do Word, e reconstrói o documento já formatado em ABNT.
- **Começar do zero**: gera um documento modelo, com capa, folha de rosto, sumário e estrutura (Introdução/Desenvolvimento/Conclusão/Referências) prontos para preencher.
- **Capa e folha de rosto** geradas a partir de um formulário (título, autor, instituição, curso, orientador, cidade, ano).
- **Sumário automático** como campo nativo do Word (atualiza os números de página ao abrir no Word).
- **Numeração progressiva de seções** (1, 1.1, 1.1.1...), conforme NBR 6024.
- **Citações longas** com recuo de 4cm, fonte reduzida e espaçamento simples, conforme NBR 10520.
- **Referências** com espaçamento correto, conforme NBR 6023.
- **Modelo de inserção de imagem** (placeholder), com legenda acima e indicação de fonte abaixo, no padrão exigido para ilustrações.
- Todo o processamento acontece **no navegador** — nenhum arquivo é enviado para um servidor.

## Normas aplicadas

| Norma | Escopo |

| NBR 14724:2024 | Apresentação geral (margens, fonte, espaçamento, capa, folha de rosto) |
| NBR 6024:2012 | Numeração progressiva das seções |
| NBR 10520:2023 | Citações |
| NBR 6023 | Referências (layout) |

## Stack técnica

- **Next.js 16** (App Router) + **React 19** com React Compiler
- **TypeScript**
- **Tailwind CSS**
- [`docx`](https://www.npmjs.com/package/docx) — geração do `.docx` final
- [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) — leitura do XML interno do `.docx` enviado
- [`file-saver`](https://www.npmjs.com/package/file-saver) — download do arquivo gerado
- Hospedado na **Vercel**

Arquitetura 100% client-side: não há backend nem upload para servidor — a leitura, a formatação e a geração do documento acontecem inteiramente no navegador do usuário.

## Como funciona (resumo técnico)

1. **Parser** (`src/lib/abnt/parser.ts`): abre o `.docx` como um `.zip`, lê `word/document.xml` e `word/styles.xml`, e classifica cada parágrafo por tipo (título, citação, referência) com base no estilo do Word aplicado.
2. **Modelo de blocos** (`src/lib/abnt/types.ts`): representação intermediária do documento, independente do formato de origem.
3. **Gerador** (`src/lib/abnt/generator.ts`): percorre os blocos e monta um novo `.docx`, aplicando as regras da ABNT (`src/lib/abnt/rules.ts`) a cada tipo de bloco.

## Rodando localmente

\`\`\`bash
git clone https://github.com/gustavonascimentof/ABNT_document_conversor.git
cd ABNT_document_conversor
npm install
npm run dev
\`\`\`

Acesse `http://localhost:3000`.

## Limitações conhecidas / próximos passos

- Listas (marcadores/numeração), células mescladas verticalmente em tabelas e legendas das imagens do documento original ainda não são tratadas. Texto, negrito/itálico, imagens e tabelas já são preservados.
- Normalização do *conteúdo* das referências (capitalização, itálico de títulos) ainda não é feita — hoje só o espaçamento/layout é corrigido.
- Gerador de referências a partir de campos estruturados (planejado).
