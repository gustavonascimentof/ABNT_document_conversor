import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { DocBlock, BlockType } from "./types";

const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  isArray: (name: string) => ["w:p", "w:r", "w:t"].includes(name),
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

function extractParagraphText(paragraph: any): string {
  const runs = paragraph["w:r"];
  const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];

  let text = "";
  for (const run of runArray) {
    const t = run["w:t"];
    const items = Array.isArray(t) ? t : t ? [t] : [];
    for (const item of items) {
      text += typeof item === "string" ? item : item?.["#text"] ?? "";
    }
  }
  return text;
}

export async function parseDocx(file: File): Promise<DocBlock[]> {
  const zip = await JSZip.loadAsync(file);

  const documentXmlFile = zip.file("word/document.xml");
  const stylesXmlFile = zip.file("word/styles.xml");
  if (!documentXmlFile) {
    throw new Error("Arquivo .docx inválido: word/document.xml não encontrado.");
  }

  const documentXml = await documentXmlFile.async("text");
  const stylesXml = stylesXmlFile ? await stylesXmlFile.async("text") : "";
  const styleMap = stylesXml ? extractStyles(stylesXml) : new Map<string, string>();

  const parser = new XMLParser(parserOptions);
  const parsed = parser.parse(documentXml);

  const paragraphs = parsed?.["w:document"]?.["w:body"]?.["w:p"];
  const paragraphArray: any[] = Array.isArray(paragraphs) ? paragraphs : paragraphs ? [paragraphs] : [];

  const blocks: DocBlock[] = [];
  let insideReferencias = false;

  for (const p of paragraphArray) {
    const styleId: string | undefined = p["w:pPr"]?.["w:pStyle"]?.["@_w:val"];
    const styleName = styleId ? styleMap.get(styleId) : undefined;
    const text = extractParagraphText(p).trim();
    if (!text) continue;

    let type = classifyStyle(styleId, styleName);
    
    if (type === "heading1" || type === "heading2") {
      insideReferencias = removerAcentos(text.toLowerCase()).includes("referencia");
    } else if (insideReferencias && type === "paragraph") {
      type = "referencia";
    }

    blocks.push({ type, text });
  }

  return blocks;
}