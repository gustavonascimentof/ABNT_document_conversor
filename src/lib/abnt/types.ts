export type BlockType =
  | "heading1"
  | "heading2"
  | "heading3"
  | "paragraph"
  | "citacaoLonga"
  | "referencia"
  | "imagemExemplo";

export interface DocBlock {
  type: BlockType;
  text: string;
}

export interface DocumentMetadata {
  titulo: string;
  autor: string;
  instituicao: string;
  curso: string;
  cidade: string;
  ano: string;
  orientador?: string;
  natureza?: string; // ex: "Trabalho apresentado à disciplina X como requisito..."
}