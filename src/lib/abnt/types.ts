export type BlockType =
  | "heading1"
  | "heading2"
  | "heading3"
  | "paragraph"
  | "citacaoLonga"
  | "referencia"
  | "imagemExemplo"
  | "imagem";

export interface TextRunData {
  text: string;
  bold?: boolean;
  italic?: boolean;
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