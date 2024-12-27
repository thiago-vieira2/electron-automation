import { useState } from "react";
import "./App.scss";

function App() {
  const [arquivo, setArquivo] = useState<File | null>(null);
  const [mensagem, setMensagem] = useState("");
  const [carregando, setCarregando] = useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target && e.target.files) {
      setArquivo(e.target.files[0]);
    }
  };

  const handleUpload = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!arquivo) {
      setMensagem("Por favor, selecione um arquivo Excel.");
      return;
    }

    setCarregando(true);
    setMensagem("Processando arquivo...");

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const buffer = event.target?.result;
        if (buffer) {
          const resultado = await window.api.iniciarNavegador(buffer);
          setMensagem(resultado || "Arquivo processado com sucesso!");
        }
      };
      reader.onerror = () => {
        setMensagem("Erro ao ler o arquivo. Tente novamente.");
      };
      reader.readAsArrayBuffer(arquivo);
    } catch (erro) {
      setMensagem("Erro ao processar o arquivo.");
      console.error(erro);
    } finally {
      setCarregando(false);
    }
  };

  return (
    <div className="App">
      <h1>Cadastramento Automático de Notas</h1>
      <form onSubmit={handleUpload}>
        <div>
          <input type="file" id="arquivo" accept=".xlsx" onChange={handleFileChange} />
        </div>
        <button className="Enviar" type="submit" disabled={carregando}>
          {carregando ? "Processando..." : "Enviar Arquivo"}
        </button>
      </form>
      {mensagem && <p>{mensagem}</p>}
    </div>
  );
}

export default App;
