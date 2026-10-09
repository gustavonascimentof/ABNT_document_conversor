import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  HeadingLevel,
  AlignmentType,
  TableOfContents,
  PageBreak,
} from "docx";
import { DocBlock, DocumentMetadata } from "./types";
import { ABNT_RULES } from "./rules";

function removerAcentos(texto: string): string {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const SECOES_SEM_NUMERO = ["referencias", "anexo", "anexos", "apendice", "apendices", "glossario"];

function ehSecaoSemNumero(texto: string): boolean {
  const limpo = removerAcentos(texto.toLowerCase());
  return SECOES_SEM_NUMERO.some((s) => limpo.includes(s));
}

function numerarTitulos(blocks: DocBlock[]): DocBlock[] {
  const contadores = [0, 0, 0];
  return blocks.map((block) => {
    if (block.type === "imagemExemplo" || ehSecaoSemNumero(block.text)) return block;

    if (block.type === "heading1") {
      contadores[0] += 1;
      contadores[1] = 0;
      contadores[2] = 0;
      return { ...block, text: `${contadores[0]} ${block.text.toUpperCase()}` };
    }
    if (block.type === "heading2") {
      contadores[1] += 1;
      contadores[2] = 0;
      return { ...block, text: `${contadores[0]}.${contadores[1]} ${block.text}` };
    }
    if (block.type === "heading3") {
      contadores[2] += 1;
      return { ...block, text: `${contadores[0]}.${contadores[1]}.${contadores[2]} ${block.text}` };
    }
    return block;
  });
}

function buildCapa(meta: DocumentMetadata): Paragraph[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: meta.instituicao.toUpperCase(), bold: true })],
    }),
    ...Array(6).fill(new Paragraph({ text: "" })),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: meta.autor.toUpperCase() })],
    }),
    ...Array(8).fill(new Paragraph({ text: "" })),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: meta.titulo.toUpperCase(), bold: true, size: ABNT_RULES.font.sizeBodyHalfPt + 4 })],
    }),
    ...Array(10).fill(new Paragraph({ text: "" })),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.cidade })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.ano })] }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function buildFolhaDeRosto(meta: DocumentMetadata): Paragraph[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: meta.autor.toUpperCase() })],
    }),
    ...Array(6).fill(new Paragraph({ text: "" })),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: meta.titulo.toUpperCase(), bold: true })],
    }),
    ...Array(4).fill(new Paragraph({ text: "" })),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: 4536 },
      children: [
        new TextRun({
          text: meta.natureza || `Trabalho apresentado como requisito para o curso de ${meta.curso}.`,
          size: ABNT_RULES.font.sizeSmallHalfPt,
        }),
      ],
    }),
    ...(meta.orientador
      ? [
          new Paragraph({
            indent: { left: 4536 },
            children: [new TextRun({ text: `Orientador(a): ${meta.orientador}`, size: ABNT_RULES.font.sizeSmallHalfPt })],
          }),
        ]
      : []),
    ...Array(10).fill(new Paragraph({ text: "" })),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.cidade })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.ano })] }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function buildSumario(): Paragraph[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: "SUMÁRIO", bold: true })],
    }),
    new TableOfContents("Sumário", { hyperlink: true, headingStyleRange: "1-3" }) as unknown as Paragraph,
  ];
}

// ---------- Placeholder de imagem (modelo coringa) ----------

const PLACEHOLDER_WIDTH_PX = 500;
const PLACEHOLDER_HEIGHT_PX = 320;

function criarImagemPlaceholder(width = PLACEHOLDER_WIDTH_PX, height = PLACEHOLDER_HEIGHT_PX): Uint8Array {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Não foi possível gerar o placeholder de imagem neste navegador.");

  ctx.fillStyle = "#ebebeb";
  ctx.fillRect(0, 0, width, height);

  ctx.strokeStyle = "#9a9a9a";
  ctx.setLineDash([10, 6]);
  ctx.lineWidth = 2;
  ctx.strokeRect(6, 6, width - 12, height - 12);

  ctx.fillStyle = "#5c5c5c";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.font = "18px Arial";
  ctx.fillText("Substitua por sua imagem", width / 2, height / 2 - 14);
  ctx.font = "14px Arial";
  ctx.fillText("(clique com o botão direito → Alterar imagem)", width / 2, height / 2 + 14);

  const dataUrl = canvas.toDataURL("image/png");
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function buildImagemExemplo(numero: number, legenda: string): Paragraph[] {
  const imagemBuffer = criarImagemPlaceholder();

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: ABNT_RULES.spacing.singleLine, lineRule: "auto", before: 240 },
      children: [new TextRun({ text: `Figura ${numero} – ${legenda}`, size: ABNT_RULES.font.sizeSmallHalfPt })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          type: "png",
          data: imagemBuffer,
          transformation: { width: PLACEHOLDER_WIDTH_PX, height: PLACEHOLDER_HEIGHT_PX },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: ABNT_RULES.spacing.singleLine, lineRule: "auto", after: 240 },
      children: [new TextRun({ text: "Fonte: elaborado pelo autor (2026).", size: ABNT_RULES.font.sizeSmallHalfPt })],
    }),
  ];
}

function buildImagemDoOriginal(imagem: NonNullable<DocBlock["imagem"]>): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240 },
    children: [
      new ImageRun({
        type: imagem.tipo,
        data: imagem.data,
        transformation: { width: imagem.width, height: imagem.height },
      }),
    ],
  });
}

// ---------- Corpo do documento ----------

function blockParaParagrafo(block: DocBlock): Paragraph {
  const { type, text } = block;

  if (type === "heading1") {
    const semNumero = ehSecaoSemNumero(text);
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      alignment: semNumero ? AlignmentType.CENTER : AlignmentType.LEFT,
      pageBreakBefore: true,
      spacing: { before: 240, after: 240 },
      children: [new TextRun({ text, bold: true, size: ABNT_RULES.font.sizeBodyHalfPt })],
    });
  }
  if (type === "heading2") {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 240 },
      children: [new TextRun({ text, bold: true, size: ABNT_RULES.font.sizeBodyHalfPt })],
    });
  }
  if (type === "heading3") {
    return new Paragraph({
      heading: HeadingLevel.HEADING_3,
      spacing: { before: 240, after: 240 },
      children: [new TextRun({ text, italics: true, size: ABNT_RULES.font.sizeBodyHalfPt })],
    });
  }
  if (type === "citacaoLonga") {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: ABNT_RULES.indent.quoteLeftTwip },
      spacing: { line: ABNT_RULES.spacing.singleLine, lineRule: "auto" },
      children: [new TextRun({ text, size: ABNT_RULES.font.sizeSmallHalfPt })],
    });
  }
  if (type === "referencia") {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: ABNT_RULES.spacing.singleLine, lineRule: "auto", after: 240 },
      children: [new TextRun({ text, size: ABNT_RULES.font.sizeBodyHalfPt })],
    });
  }

    const runsRenderizados =
    block.runs && block.runs.length > 0
      ? block.runs.map(
          (r) => new TextRun({ text: r.text, size: ABNT_RULES.font.sizeBodyHalfPt, bold: r.bold, italics: r.italic })
        )
      : [new TextRun({ text, size: ABNT_RULES.font.sizeBodyHalfPt })];

  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: ABNT_RULES.indent.firstLineTwip },
    spacing: { line: ABNT_RULES.spacing.bodyLine, lineRule: "auto" },
    children: runsRenderizados,
  });
}

export async function generateAbntDocx(blocks: DocBlock[], meta: DocumentMetadata): Promise<Blob> {
  const blocksNumerados = numerarTitulos(blocks);

  let contadorFigura = 0;
  const corpo: Paragraph[] = [];
  for (const block of blocksNumerados) {
    if (block.type === "imagemExemplo") {
      contadorFigura += 1;
      corpo.push(...buildImagemExemplo(contadorFigura, block.text));
    } else if (block.type === "imagem" && block.imagem) {
      corpo.push(buildImagemDoOriginal(block.imagem));
    } else {
      corpo.push(blockParaParagrafo(block));
    }
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: ABNT_RULES.font.family, size: ABNT_RULES.font.sizeBodyHalfPt, color: "000000" },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: ABNT_RULES.font.family, size: ABNT_RULES.font.sizeBodyHalfPt, bold: true, color: "000000" },
          paragraph: { spacing: { before: 240, after: 240 } },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: ABNT_RULES.font.family, size: ABNT_RULES.font.sizeBodyHalfPt, bold: true, color: "000000" },
          paragraph: { spacing: { before: 240, after: 240 } },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: ABNT_RULES.font.family, size: ABNT_RULES.font.sizeBodyHalfPt, italics: true, color: "000000" },
          paragraph: { spacing: { before: 240, after: 240 } },
        },
      ],
      characterStyles: [
        {
          id: "Hyperlink",
          name: "Hyperlink",
          basedOn: "DefaultParagraphFont",
          run: { color: "000000" },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: ABNT_RULES.margin.topTwip,
              bottom: ABNT_RULES.margin.bottomTwip,
              left: ABNT_RULES.margin.leftTwip,
              right: ABNT_RULES.margin.rightTwip,
            },
          },
        },
        children: [...buildCapa(meta), ...buildFolhaDeRosto(meta), ...buildSumario(), ...corpo],
      },
    ],
  });

  return Packer.toBlob(doc);
}