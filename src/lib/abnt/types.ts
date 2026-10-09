export type BlockType =
  | "heading1"
  | "heading2"
  | "heading3"
  | "paragraph"
  | "citacaoLonga"
  | "referencia"
  | "imagemExemplo"
  | "imagem"
  | "tabela";

export interface TextRunData {
  text: string;
  bold?: boolean;
  italic?: boolean;
}

export interface CelulaTabela {
  text: string;
  colSpan?: number;
}

export interface DadosTabela {
  linhas: CelulaTabela[][];
  tipo: "tabela" | "quadro";
  titulo?: string;
  fonte?: string;
}

export interface DocBlock {
  type: BlockType;
  text: string;
  runs?: TextRunData[];
  imagem?: {
    data: Uint8Array;
    width: number;
    height: number;
    tipo: "png" | "jpg" | "gif" | "bmp";
  };
  tabela?: DadosTabela;
}

export interface DocumentMetadata {
  titulo: string;
  autor: string;
  instituicao: string;
  curso: string;
  cidade: string;
  ano: string;
  orientador?: string;
  natureza?: string;
}