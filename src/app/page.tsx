"use client";

import { useState } from "react";
import { saveAs } from "file-saver";
import { parseDocx } from "@/lib/abnt/parser";
import { generateAbntDocx } from "@/lib/abnt/generator";
import { DocBlock, DocumentMetadata } from "@/lib/abnt/types";

const METADATA_INICIAL: DocumentMetadata = {
  titulo: "",
  autor: "",
  instituicao: "",
  curso: "",
  cidade: "",
  ano: new Date().getFullYear().toString(),
  orientador: "",
  natureza: "",
};

const BLOCOS_TEMPLATE_DO_ZERO: DocBlock[] = [
  { type: "heading1", text: "INTRODUÇÃO" },
  { type: "paragraph", text: "Escreva aqui a introdução do seu trabalho." },
  { type: "heading1", text: "DESENVOLVIMENTO" },
  { type: "paragraph", text: "Escreva aqui o desenvolvimento do seu trabalho." },
  { type: "imagemExemplo", text: "Exemplo de legenda descritiva da imagem" },
  { type: "paragraph", text: "Continue o texto normalmente depois da imagem, explicando o que ela demonstra." },
  { type: "heading1", text: "CONCLUSÃO" },
  { type: "paragraph", text: "Escreva aqui a conclusão do seu trabalho." },
  { type: "heading1", text: "REFERÊNCIAS" },
  { type: "referencia", text: "SOBRENOME, Nome. Título da obra. Cidade: Editora, ano." },
];
export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [fromScratch, setFromScratch] = useState(false);
  const [meta, setMeta] = useState<DocumentMetadata>(METADATA_INICIAL);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  function handleMetaChange(campo: keyof DocumentMetadata, valor: string) {
    setMeta((atual) => ({ ...atual, [campo]: valor }));
  }

  async function handleGenerate() {
    setError(null);

    if (!fromScratch && !file) {
      setError("Selecione um arquivo .docx ou marque a opção 'Começar do zero'.");
      return;
    }
    if (!meta.titulo || !meta.autor || !meta.instituicao) {
      setError("Preencha pelo menos título, autor e instituição.");
      return;
    }

    setLoading(true);
    try {
      const blocks: DocBlock[] = fromScratch
        ? BLOCOS_TEMPLATE_DO_ZERO
        : await parseDocx(file as File);

      const blob = await generateAbntDocx(blocks, meta);
      saveAs(blob, `${meta.titulo || "documento"}-abnt.docx`);
    } catch (e) {
      console.error(e);
      setError("Não foi possível gerar o documento. Confira se o arquivo é um .docx válido.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="mx-auto flex max-w-2xl flex-col gap-6 p-8">
      <h1 className="text-2xl font-bold">Conversor ABNT</h1>

      <div className="flex flex-col gap-2">
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={fromScratch}
            onChange={(e) => setFromScratch(e.target.checked)}
          />
          Começar um documento do zero (sem enviar arquivo)
        </label>

        {!fromScratch && (
          <input
            type="file"
            accept=".docx"
            onChange={(e) => setFile(e.target.files?.[0] ?? null)}
          />
        )}
      </div>

      <fieldset className="flex flex-col gap-3 rounded border p-4">
        <legend className="px-1 font-semibold">Dados da capa</legend>

        <input
          className="rounded border p-2"
          placeholder="Título do trabalho"
          value={meta.titulo}
          onChange={(e) => handleMetaChange("titulo", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Autor(a)"
          value={meta.autor}
          onChange={(e) => handleMetaChange("autor", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Instituição"
          value={meta.instituicao}
          onChange={(e) => handleMetaChange("instituicao", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Curso"
          value={meta.curso}
          onChange={(e) => handleMetaChange("curso", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Cidade"
          value={meta.cidade}
          onChange={(e) => handleMetaChange("cidade", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Ano"
          value={meta.ano}
          onChange={(e) => handleMetaChange("ano", e.target.value)}
        />
        <input
          className="rounded border p-2"
          placeholder="Orientador(a) (opcional)"
          value={meta.orientador}
          onChange={(e) => handleMetaChange("orientador", e.target.value)}
        />
      </fieldset>

      <button
        onClick={handleGenerate}
        disabled={loading}
        className="rounded bg-black px-4 py-2 text-white disabled:opacity-50"
      >
        {loading ? "Gerando..." : "Gerar documento ABNT"}
      </button>

      {error && <p className="text-red-600">{error}</p>}

      <p className="text-sm text-gray-500">
        Depois de baixar, abra o Word, clique com o botão direito no Sumário e escolha
        &quot;Atualizar campo&quot; para os números de página aparecerem.
      </p>
    </main>
  );
}