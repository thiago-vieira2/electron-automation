import { useState } from "react";
import "./App.scss";

function App() {
  const [arquivo, setArquivo] = useState<{ path?: string }>({});
  const [mensagem, setMensagem] = useState("");
  const [carregando, setCarregando] = useState(false);


  const handleFileChange = (e) => {
    setArquivo(e.target.files[0]);
  };

  const handleUpload = async (e) => {
    e.preventDefault();

    if (!arquivo) {
      setMensagem("Por favor, selecione um arquivo Excel.");
      return;
    }

    window.api.iniciarNavegador(arquivo.path);
    setCarregando(true);
    setMensagem("Processando arquivo...");

    try {
      // Chama o processo principal via Electron
      const resultado = '';
      setMensagem(resultado);
    } catch (erro) {
    } finally {
      setCarregando(false);
    }
  };
  
  
  return (  
    <div className="App">
      <h1>Automação do Ministério da Fazenda</h1>
      <form onSubmit={handleUpload}>
        <div>
          <label htmlFor="arquivo">Selecione o arquivo Excel:</label>
          <input type="file" id="arquivo" accept=".xlsx" onChange={handleFileChange} />
        </div>
        <button type="submit" disabled={carregando}>
          {carregando ? "Processando..." : "Enviar Arquivo"}
        </button>
      </form>
      {mensagem && <p>{mensagem}</p>}
    </div>
  );
}

export default App;