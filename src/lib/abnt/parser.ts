import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { DocBlock, BlockType, TextRunData, CelulaTabela, DadosTabela } from "./types";

const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: false,
  isArray: (name: string) => ["w:p", "w:r", "w:t", "w:tbl", "w:tr", "w:tc"].includes(name),
};

function extractStyles(stylesXml: string): Map<string, string> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const parsed = parser.parse(stylesXml);

  const styles = parsed?.["w:styles"]?.["w:style"];
  const styleArray = Array.isArray(styles) ? styles : styles ? [styles] : [];

  const styleMap = new Map<string, string>();
  for (const style of styleArray) {
    const id = style["@_w:styleId"];
    const name = style["w:name"]?.["@_w:val"];
    if (id && name) styleMap.set(id, String(name).toLowerCase());
  }
  return styleMap;
}

function extractRelationships(relsXml: string): Map<string, string> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const parsed = parser.parse(relsXml);

  const relationships = parsed?.["Relationships"]?.["Relationship"];
  const relArray = Array.isArray(relationships) ? relationships : relationships ? [relationships] : [];

  const map = new Map<string, string>();
  for (const rel of relArray) {
    const id = rel["@_Id"];
    const target = rel["@_Target"];
    if (id && target) map.set(id, target);
  }
  return map;
}

// Descobre a ordem original entre parágrafos ("p") e tabelas ("tbl") no corpo do documento
function ordemDoCorpo(documentXml: string): ("p" | "tbl")[] {
  const parser = new XMLParser({ preserveOrder: true, ignoreAttributes: false });
  const parsed = parser.parse(documentXml) as any[];

  const doc = parsed.find((n) => n["w:document"]);
  const body = doc?.["w:document"]?.find((n: any) => n["w:body"]);
  const filhos: any[] = body?.["w:body"] ?? [];

  const ordem: ("p" | "tbl")[] = [];
  for (const filho of filhos) {
    if (filho["w:p"]) ordem.push("p");
    else if (filho["w:tbl"]) ordem.push("tbl");
  }
  return ordem;
}

function removerAcentos(texto: string): string {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizar(valor?: string): string {
  return removerAcentos((valor || "").toLowerCase()).replace(/\s+/g, "");
}

function classifyStyle(styleId?: string, styleName?: string): BlockType {
  const combinado = `${normalizar(styleId)} ${normalizar(styleName)}`;

  if (combinado.includes("heading1") || combinado.includes("titulo1")) return "heading1";
  if (combinado.includes("heading2") || combinado.includes("titulo2")) return "heading2";
  if (combinado.includes("heading3") || combinado.includes("titulo3")) return "heading3";
  if (combinado.includes("quote") || combinado.includes("citac")) return "citacaoLonga";
  return "paragraph";
}

function propriedadeAtiva(prop: any): boolean {
  if (prop === undefined) return false;
  if (typeof prop === "object") {
    const val = prop["@_w:val"];
    if (val === undefined) return true;
    return val === "1" || val === "true" || val === "on";
  }
  return true;
}

function extractParagraphRuns(paragraph: any): TextRunData[] {
  const runs = paragraph["w:r"];
  const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];

  const resultado: TextRunData[] = [];
  for (const run of runArray) {
    const rPr = run["w:rPr"];
    const bold = propriedadeAtiva(rPr?.["w:b"]);
    const italic = propriedadeAtiva(rPr?.["w:i"]);

    const t = run["w:t"];
    const items = Array.isArray(t) ? t : t ? [t] : [];
    let texto = "";
    for (const item of items) {
            texto += typeof item === "object" ? String(item?.["#text"] ?? "") : String(item ?? "");
    }
    if (!texto) continue;

    resultado.push({ text: texto, bold, italic });
  }
  return resultado;
}

function normalizarTipoImagem(ext: string): "jpg" | "png" | "gif" | "bmp" {
  const e = ext.toLowerCase();
  if (e === "jpeg") return "jpg";
  if (e === "jpg" || e === "png" || e === "gif" || e === "bmp") return e as "jpg" | "png" | "gif" | "bmp";
  return "png";
}

async function extractImagemDoRun(
  run: any,
  relMap: Map<string, string>,
  zip: JSZip
): Promise<DocBlock["imagem"] | null> {
  const drawing = run["w:drawing"];
  if (!drawing) return null;

  const container = drawing["wp:inline"] ?? drawing["wp:anchor"];
  if (!container) return null;

  const blip = container?.["a:graphic"]?.["a:graphicData"]?.["pic:pic"]?.["pic:blipFill"]?.["a:blip"];
  const relId = blip?.["@_r:embed"];
  if (!relId) return null;

  const target = relMap.get(relId);
  if (!target) return null;

  const caminho = `word/${target.replace(/^\/?word\//, "")}`;
  const arquivoImagem = zip.file(caminho) ?? zip.file(target);
  if (!arquivoImagem) return null;

  const data = await arquivoImagem.async("uint8array");

  const extent = container?.["wp:extent"];
  const cx = Number(extent?.["@_cx"] ?? 0);
  const cy = Number(extent?.["@_cy"] ?? 0);

  const EMU_POR_PIXEL = 9525;
  let width = cx > 0 ? Math.round(cx / EMU_POR_PIXEL) : 400;
  let height = cy > 0 ? Math.round(cy / EMU_POR_PIXEL) : 300;

  const LARGURA_MAXIMA = 560;
  if (width > LARGURA_MAXIMA) {
    const fator = LARGURA_MAXIMA / width;
    width = LARGURA_MAXIMA;
    height = Math.round(height * fator);
  }

  const extensao = target.split(".").pop() || "png";

  return { data, width, height, tipo: normalizarTipoImagem(extensao) };
}

// ---------- Tabelas ----------

function ehNumerico(texto: string): boolean {
  return /\d/.test(texto) && /^[\s\d.,%R$€+\-–()]+$/.test(texto);
}

// Tabela (IBGE) = dados majoritariamente numéricos; Quadro = conteúdo textual
function classificarTabela(linhas: CelulaTabela[][]): "tabela" | "quadro" {
  let total = 0;
  let numericas = 0;
  for (const linha of linhas.slice(1)) {
    for (const celula of linha.slice(1)) {
      if (!celula.text) continue;
      total += 1;
      if (ehNumerico(celula.text)) numericas += 1;
    }
  }
  return total > 0 && numericas / total >= 0.5 ? "tabela" : "quadro";
}

function extractTabela(tbl: any): DadosTabela | null {
  const linhasXml: any[] = tbl["w:tr"] ?? [];
  const linhas: CelulaTabela[][] = [];

  for (const tr of linhasXml) {
    const celulasXml: any[] = tr["w:tc"] ?? [];
    const linha: CelulaTabela[] = [];

    for (const tc of celulasXml) {
      const paragrafos: any[] = tc["w:p"] ?? [];
      const texto = paragrafos
        .map((p) => extractParagraphRuns(p).map((r) => r.text).join(""))
        .filter(Boolean)
        .join(" ")
        .trim();
      const span = Number(tc["w:tcPr"]?.["w:gridSpan"]?.["@_w:val"] ?? 1);
      linha.push({ text: texto, colSpan: span > 1 ? span : undefined });
    }
    if (linha.length > 0) linhas.push(linha);
  }

  if (linhas.length === 0) return null;
  return { linhas, tipo: classificarTabela(linhas) };
}

const REGEX_TITULO_TABELA = /^(tabela|quadro)\s*\d*\s*[–—-]\s*(.+)$/i;
const REGEX_FONTE = /^fonte\s*:\s*(.+)$/i;

// Se o usuário já escreveu "Tabela 1 – título" acima e "Fonte: ..." abaixo, reaproveita
function associarLegendas(blocks: DocBlock[]): DocBlock[] {
  const resultado: DocBlock[] = [];

  for (let i = 0; i < blocks.length; i++) {
    const bloco = blocks[i];
    if (bloco.type !== "tabela" || !bloco.tabela) {
      resultado.push(bloco);
      continue;
    }

    const anterior = resultado[resultado.length - 1];
    if (anterior && anterior.type === "paragraph") {
      const m = anterior.text.match(REGEX_TITULO_TABELA);
      if (m) {
        bloco.tabela.titulo = m[2].trim();
        bloco.tabela.tipo = m[1].toLowerCase() === "quadro" ? "quadro" : "tabela";
        resultado.pop();
      }
    }

    const proximo = blocks[i + 1];
    if (proximo && proximo.type === "paragraph") {
      const m = proximo.text.match(REGEX_FONTE);
      if (m) {
        bloco.tabela.fonte = m[1].trim();
        i += 1;
      }
    }

    resultado.push(bloco);
  }

  return resultado;
}

// ---------- Função principal ----------

export async function parseDocx(file: File): Promise<DocBlock[]> {
  const zip = await JSZip.loadAsync(file);

  const documentXmlFile = zip.file("word/document.xml");
  const stylesXmlFile = zip.file("word/styles.xml");
  const relsXmlFile = zip.file("word/_rels/document.xml.rels");
  if (!documentXmlFile) {
    throw new Error("Arquivo .docx inválido: word/document.xml não encontrado.");
  }

  const documentXml = await documentXmlFile.async("text");
  const stylesXml = stylesXmlFile ? await stylesXmlFile.async("text") : "";
  const relsXml = relsXmlFile ? await relsXmlFile.async("text") : "";

  const styleMap = stylesXml ? extractStyles(stylesXml) : new Map<string, string>();
  const relMap = relsXml ? extractRelationships(relsXml) : new Map<string, string>();

  const parser = new XMLParser(parserOptions);
  const parsed = parser.parse(documentXml);
  const body = parsed?.["w:document"]?.["w:body"];

  const paragraphs = body?.["w:p"];
  const paragraphArray: any[] = Array.isArray(paragraphs) ? paragraphs : paragraphs ? [paragraphs] : [];
  const tables = body?.["w:tbl"];
  const tableArray: any[] = Array.isArray(tables) ? tables : tables ? [tables] : [];

  const ordem = ordemDoCorpo(documentXml);
  const sequencia: ("p" | "tbl")[] = ordem.length > 0 ? ordem : paragraphArray.map(() => "p" as const);

  const blocks: DocBlock[] = [];
  let insideReferencias = false;
  let contadorFiguraOriginal = 0;
  let iP = 0;
  let iT = 0;

  for (const item of sequencia) {
    if (item === "tbl") {
      const dados = tableArray[iT] ? extractTabela(tableArray[iT]) : null;
      iT += 1;
      if (dados) blocks.push({ type: "tabela", text: "", tabela: dados });
      continue;
    }

    const p = paragraphArray[iP];
    iP += 1;
    if (!p) continue;

    const runsBrutos = p["w:r"];
    const runArray = Array.isArray(runsBrutos) ? runsBrutos : runsBrutos ? [runsBrutos] : [];

    let imagemEncontrada: DocBlock["imagem"] | null = null;
    for (const run of runArray) {
      try {
        const resultado = await extractImagemDoRun(run, relMap, zip);
        if (resultado) {
          imagemEncontrada = resultado;
          break;
        }
      } catch {
        // se uma imagem específica falhar ao extrair, ignora ela e segue o documento
      }
    }

    if (imagemEncontrada) {
      contadorFiguraOriginal += 1;
      blocks.push({
        type: "imagem",
        text: `Imagem ${contadorFiguraOriginal} do documento original`,
        imagem: imagemEncontrada,
      });
      continue;
    }

    const styleId: string | undefined = p["w:pPr"]?.["w:pStyle"]?.["@_w:val"];
    const styleName = styleId ? styleMap.get(styleId) : undefined;
    const runsDoParagrafo = extractParagraphRuns(p);
    const text = runsDoParagrafo.map((r) => r.text).join("").trim();
    if (!text) continue;

    let type = classifyStyle(styleId, styleName);

    if (type === "heading1" || type === "heading2") {
      insideReferencias = removerAcentos(text.toLowerCase()).includes("referencia");
    } else if (insideReferencias && type === "paragraph") {
      type = "referencia";
    }

    blocks.push({ type, text, runs: runsDoParagrafo });
  }

  return associarLegendas(blocks);
}