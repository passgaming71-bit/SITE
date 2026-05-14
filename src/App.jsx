import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { usuarios } from './usuarios';

const COLORS = {
  bg: '#0D0D0D',
  card: '#151515',
  cardAlt: '#1A1A1A',
  blue: '#0D6EFD',
  blueLight: '#3BA0FF',
  text: '#FFFFFF',
  muted: '#BFC7D5',
  border: '#2A2A2A',
  successBg: '#0F2D1D',
  successText: '#34D399',
  dangerBg: '#2B1215',
  dangerText: '#F87171',
};

const STORAGE_KEYS = {
  lista: 'odax_web_lista',
  index: 'odax_web_index',
  aba: 'odax_web_aba',
};

export default function App() {
  const [lista, setLista] = useState([]);
  const [index, setIndex] = useState(0);
  const [quantidade, setQuantidade] = useState('');
  const [aba, setAba] = useState('inicio');
  const [carregado, setCarregado] = useState(false);
  const [usuarioLogado, setUsuarioLogado] = useState(null);
  const [login, setLogin] = useState('');
  const [senha, setSenha] = useState('');

  const fileInputRef = useRef(null);
  const item = lista[index] || null;

  const total = lista.length;
  const contados = lista.filter((i) => i.contagem !== null).length;
  const diferencasCount = lista.filter((i) => i.contagem !== null && i.contagem !== i.estoque).length;
  const pendentes = total - contados;
  const progresso = total === 0 ? 0 : Math.round((contados / total) * 100);

  const diferencas = useMemo(() => {
    return lista.filter((i) => i.contagem !== null && i.contagem !== i.estoque);
  }, [lista]);

  useEffect(() => {
    try {
      const usuarioSalvo = localStorage.getItem('odax_usuario');
      const listaSalva = localStorage.getItem(STORAGE_KEYS.lista);
      const indexSalvo = localStorage.getItem(STORAGE_KEYS.index);
      const abaSalva = localStorage.getItem(STORAGE_KEYS.aba);

      if (usuarioSalvo) {
        setUsuarioLogado(JSON.parse(usuarioSalvo));
      }

      if (listaSalva) {
        const parsed = JSON.parse(listaSalva);
        setLista(Array.isArray(parsed) ? parsed : []);
      }

      if (indexSalvo !== null) setIndex(Number(indexSalvo) || 0);
      if (abaSalva) setAba(abaSalva);
    } catch (error) {
      console.error('Erro ao carregar dados:', error);
    } finally {
      setCarregado(true);
    }
  }, []);

  useEffect(() => {
    if (!carregado) return;

    const timer = setTimeout(() => {
      localStorage.setItem(STORAGE_KEYS.lista, JSON.stringify(lista));
      localStorage.setItem(STORAGE_KEYS.index, String(index));
      localStorage.setItem(STORAGE_KEYS.aba, aba);
    }, 500);

    return () => clearTimeout(timer);
  }, [lista, index, aba, carregado]);

  function ordenar(listaOrdenar) {
    return [...listaOrdenar].sort((a, b) =>
      String(a.endereco || '').localeCompare(String(b.endereco || ''), 'pt-BR', { numeric: true })
    );
  }

  function calcularDiferenca(contagem, estoque) {
    if (contagem === null || contagem === undefined) return null;
    return contagem - estoque;
  }

  function normalizarNumero(valor) {
    if (valor === null || valor === undefined || valor === '') return 0;
    const convertido = Number(String(valor).replace(',', '.'));
    return Number.isNaN(convertido) ? 0 : convertido;
  }


  function normalizarChave(chave) {
    return String(chave || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .trim();
  }

  function buscarCampo(objeto, nomesPossiveis) {
    const nomesNormalizados = nomesPossiveis.map(normalizarChave);

    const campo = Object.keys(objeto).find((key) =>
      nomesNormalizados.includes(normalizarChave(key))
    );

    return campo ? objeto[campo] : '';
  }

  function buscarPrimeiroCampo(objeto, nomesPossiveis, valorPadrao = '') {
    const valor = buscarCampo(objeto, nomesPossiveis);
    return valor === null || valor === undefined || valor === '' ? valorPadrao : valor;
  }

  async function importarExcel(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

      const dados = json.map((i) => {
        const endereco = buscarPrimeiroCampo(i, [
          'Endereços',
          'Endereços ',
          'Endereco',
          'Endereço',
          'Enderecos',
          'Endereços',
          'MEZANINO',
          'Mezanino',
          'Rua',
          'RUA',
          'Bin',
          'BIN',
          'Local',
          'LOCAL',
          'Prateleira',
          'PRATELEIRA',
          'Posição',
          'Posicao',
          'POSIÇÃO',
          'POSICAO',
          'End',
          'END'
        ], '');

        return {
          codigo: buscarPrimeiroCampo(i, [
            'Material',
            'MATERIAL',
            'Código',
            'Codigo',
            'CÓDIGO',
            'CODIGO'
          ], ''),

          material: buscarPrimeiroCampo(i, [
            'Texto breve material',
            'DESCRIÇÃO',
            'Descrição',
            'Descricao',
            'DESCRIPTION',
            'Material Descrição',
            'Material Descricao'
          ], 'Sem nome'),

          unidade: buscarPrimeiroCampo(i, [
            'Unid.medida basi',
            'Unidade',
            'UN',
            'UMB',
            'Unid',
            'Unidade Medida'
          ], 'UN'),

          endereco: String(endereco || '').trim(),

          estoque: normalizarNumero(buscarPrimeiroCampo(i, [
            'Utilização livre',
            'Utilizacao livre',
            'Saldo no Sistema',
            'Saldo Sistema',
            'Estoque',
            'ESTOQUE',
            'Saldo',
            'SALDO',
            'Livre utilização',
            'Livre utilizacao'
          ], 0)),

          contagem: null,
        };
      });

      const ordenado = ordenar(dados);
      setLista(ordenado);
      setIndex(0);
      setQuantidade('');
      setAba('contagem');
      alert('Base importada com sucesso!');
    } catch (error) {
      console.error(error);
      alert('Falha ao importar base.');
    } finally {
      event.target.value = '';
    }
  }

  function salvarItem(valor) {
    if (!item) return;
    if (valor === '' || valor === null || valor === undefined) return;

    const numero = Number(String(valor).replace(',', '.'));
    if (Number.isNaN(numero)) {
      alert('Digite uma quantidade válida.');
      return;
    }

    const nova = [...lista];
    nova[index].contagem = numero;
    setLista(nova);
    setQuantidade('');

    if (numero === item.estoque) {
      if (index + 1 < nova.length) {
        setIndex(index + 1);
      } else {
        alert('Você chegou ao último item.');
      }
    } else {
      alert('Diferença encontrada. Corrija antes de continuar.');
    }
  }

  function voltarItem() {
    if (index > 0) {
      const novoIndex = index - 1;
      setIndex(novoIndex);
      const contagemAnterior = lista[novoIndex]?.contagem;
      setQuantidade(contagemAnterior === null || contagemAnterior === undefined ? '' : String(contagemAnterior));
      return;
    }

    setAba('inicio');
  }

  function proximoItem() {
    if (index + 1 < lista.length) {
      const novoIndex = index + 1;
      setIndex(novoIndex);
      const contagemProxima = lista[novoIndex]?.contagem;
      setQuantidade(contagemProxima === null || contagemProxima === undefined ? '' : String(contagemProxima));
    }
  }

  function exportarArquivo(dados, nomeArquivo, nomeAba) {
    if (!dados.length) {
      alert('Não há dados para exportar.');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nomeAba);
    XLSX.writeFile(wb, nomeArquivo);
  }

  function exportarDiferencas() {
    const dados = diferencas.map((i) => ({
      Codigo: i.codigo,
      Material: i.material,
      Endereco: i.endereco,
      SaldoSistema: i.estoque,
      ContagemFisica: i.contagem,
      Resultado: i.contagem - i.estoque,
    }));

    exportarArquivo(dados, 'diferencas.xlsx', 'Diferencas');
  }

  function exportarInventarioCompleto() {
    const dados = lista.map((i) => ({
      Codigo: i.codigo,
      Material: i.material,
      Endereco: i.endereco,
      SaldoSistema: i.estoque,
      ContagemFisica: i.contagem === null ? '' : i.contagem,
      Resultado: i.contagem === null ? '' : i.contagem - i.estoque,
    }));

    exportarArquivo(dados, 'inventario_completo.xlsx', 'InventarioCompleto');
  }

  function limparProgresso() {
    const confirmar = window.confirm('Deseja encerrar o inventário?');
    if (!confirmar) return;

    localStorage.removeItem(STORAGE_KEYS.lista);
    localStorage.removeItem(STORAGE_KEYS.index);
    localStorage.removeItem(STORAGE_KEYS.aba);
    setLista([]);
    setIndex(0);
    setQuantidade('');
    setAba('inicio');
  }

  function renderResultado(valor) {
    if (valor === null) return <span className="muted">-</span>;

    const cor = valor === 0 ? COLORS.successText : COLORS.dangerText;
    return <span style={{ color: cor, fontWeight: 900 }}>{valor > 0 ? `+${valor}` : valor}</span>;
  }

  function renderResumo() {
    return (
      <>
        <div className="progressBox">
          <div className="progressBar">
            <div className="progressFill" style={{ width: `${progresso}%` }} />
          </div>
          <div className="progressText">{progresso}% concluído</div>
        </div>

        <div className="summary">
          <span>Total: {total}</span>
          <span>Conferidos: {contados}</span>
          <span>Pendentes: {pendentes}</span>
          <span>Diferenças: {diferencasCount}</span>
        </div>
      </>
    );
  }

  function fazerLogin(event) {
    event.preventDefault();

    const usuarioEncontrado = usuarios.find(
      (u) =>
        String(u.usuario).trim().toLowerCase() === String(login).trim().toLowerCase() &&
        String(u.senha) === String(senha)
    );

    if (!usuarioEncontrado) {
      alert('Usuário ou senha inválidos.');
      return;
    }

    localStorage.setItem('odax_usuario', JSON.stringify(usuarioEncontrado));
    setUsuarioLogado(usuarioEncontrado);
    setLogin('');
    setSenha('');
  }

  function sairSistema() {
    localStorage.removeItem('odax_usuario');
    setUsuarioLogado(null);
    setLogin('');
    setSenha('');
    setAba('inicio');
  }

  function renderLogin() {
    return (
      <main className="loginPage">
        <form className="loginCard" onSubmit={fazerLogin}>
          <div className="loginBrand">
            <div className="loginLogo">✓</div>
            <h1>ODAX Estoque</h1>
            <p>Acesse seu inventário profissional</p>
          </div>

          <label>
            Usuário
            <input
              type="text"
              placeholder="Digite seu usuário"
              value={login}
              onChange={(e) => setLogin(e.target.value)}
              autoFocus
            />
          </label>

          <label>
            Senha
            <input
              type="password"
              placeholder="Digite sua senha"
              value={senha}
              onChange={(e) => setSenha(e.target.value)}
            />
          </label>

          <button type="submit">Entrar</button>
        </form>
      </main>
    );
  }

  function renderMenu() {
    return (
      <main className="page">
        <header className="header">
          <div>
            <h1>ODAX Estoque</h1>
            <p>Inventário profissional</p>
          </div>

          <div className="userBox">
            <span>👤 {usuarioLogado?.usuario}</span>
            <button onClick={sairSistema}>Sair</button>
          </div>
        </header>

        <section className="grid">
          <button className="menuCard" onClick={() => fileInputRef.current?.click()}>
            <strong>📥 Importar Base</strong>
            <small>Selecionar planilha Excel</small>
          </button>

          <button className="menuCard" onClick={() => lista.length ? setAba('contagem') : alert('Importe uma base primeiro.')}>
            <strong>📋 Contagem</strong>
            <small>Continuar inventário</small>
          </button>

          <button className="menuCard" onClick={() => lista.length ? setAba('divergencias') : alert('Importe uma base primeiro.')}>
            <strong>⚠️ Diferenças</strong>
            <small>Consultar divergências</small>
          </button>

          <button className="menuCard" onClick={exportarInventarioCompleto}>
            <strong>📤 Exportar</strong>
            <small>Baixar inventário completo</small>
          </button>

          <button className="menuCard danger" onClick={limparProgresso}>
            <strong>🧹 Encerrar</strong>
            <small>Limpar inventário atual</small>
          </button>

          <button className="menuCard" onClick={() => alert('Configurações em breve.')}>
            <strong>⚙️ Ajustes</strong>
            <small>Configurações do sistema</small>
          </button>
        </section>

        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={importarExcel} />
      </main>
    );
  }


  function escolherItemPorNumero() {
    if (!lista.length) return;

    const escolhido = window.prompt(`Digite o número do item entre 1 e ${lista.length}:`, String(index + 1));

    if (escolhido === null) return;

    const numero = Number(String(escolhido).trim());

    if (!Number.isInteger(numero) || numero < 1 || numero > lista.length) {
      alert(`Número inválido. Digite um número entre 1 e ${lista.length}.`);
      return;
    }

    const novoIndex = numero - 1;
    setIndex(novoIndex);

    const contagemAtual = lista[novoIndex]?.contagem;
    setQuantidade(contagemAtual === null || contagemAtual === undefined ? '' : String(contagemAtual));
  }

  function renderContagem() {
    if (!item) {
      return (
        <main className="page">
          <button className="backButton" onClick={() => setAba('inicio')}>← Voltar ao menu</button>
          <div className="emptyCard">
            <h2>Nenhuma base importada</h2>
            <p>Importe uma planilha para iniciar a contagem.</p>
            <button className="primaryButton" onClick={() => fileInputRef.current?.click()}>Importar Base</button>
          </div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={importarExcel} />
        </main>
      );
    }

    const resultado = calcularDiferenca(item.contagem, item.estoque);
    const statusClass = item.contagem === null ? '' : item.contagem === item.estoque ? 'success' : 'danger';

    return (
      <main className="page">
        <button
          className="backButton"
          onClick={() => setAba('inicio')}
        >
          ← Voltar ao menu
        </button>

        <div style={{ display: 'flex', justifyContent: 'center' }}>
          <button
            className="itemSelector"
            onClick={escolherItemPorNumero}
            title="Clique para escolher o número do item"
          >
            Item {index + 1} / {lista.length}
          </button>
        </div>

        <section className={`itemCard ${statusClass}`}>
          <h3>{item.codigo} - {item.material}</h3>
          <p>Endereço: {item.endereco || '-'}</p>
          <p>Saldo Sistema: {item.estoque} {item.unidade}</p>
          <p>Contagem Física: {item.contagem === null ? '-' : `${item.contagem} ${item.unidade}`}</p>
          <p>RESULTADO: {renderResultado(resultado)}</p>
          {item.contagem !== null && <p className="okText">✔ Conferido</p>}
        </section>

        <div className="inputRow">
          <input
            className="input"
            type="number"
            placeholder="Digite a quantidade"
            value={quantidade}
            onChange={(e) => setQuantidade(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') salvarItem(quantidade);
            }}
          />
          <button className="primaryButton" onClick={() => salvarItem(quantidade)}>Salvar</button>
        </div>

        <div className="navRow">
          <button className="secondaryButton" onClick={voltarItem}>← Anterior</button>
          <button className="secondaryButton" onClick={proximoItem}>Próximo →</button>
        </div>
      </main>
    );
  }

  function renderDivergencias() {
    return (
      <main className="page">
        <button className="backButton" onClick={() => setAba('inicio')}>← Voltar ao menu</button>

        <h2>Diferenças</h2>
        {renderResumo()}

        <div className="actions">
          <button className="primaryButton" onClick={exportarDiferencas}>Exportar Diferenças</button>
          <button className="primaryButton" onClick={exportarInventarioCompleto}>Exportar Completo</button>
          <button className="secondaryButton" onClick={limparProgresso}>Encerrar Inventário</button>
        </div>

        <div className="list">
          {diferencas.length === 0 ? (
            <p className="muted">Nenhuma diferença encontrada.</p>
          ) : (
            diferencas.map((divItem, i) => {
              const d = calcularDiferenca(divItem.contagem, divItem.estoque);
              return (
                <button
                  key={`${divItem.codigo}-${divItem.endereco}-${i}`}
                  className="diffCard"
                  onClick={() => {
                    const idx = lista.findIndex((item) =>
                      String(item.codigo) === String(divItem.codigo) &&
                      String(item.endereco) === String(divItem.endereco)
                    );
                    if (idx !== -1) {
                      setIndex(idx);
                      setAba('contagem');
                    }
                  }}
                >
                  <strong>{divItem.codigo} - {divItem.material}</strong>
                  <span>Endereço: {divItem.endereco || '-'}</span>
                  <span>Saldo Sistema: {divItem.estoque} {divItem.unidade}</span>
                  <span>Contagem Física: {divItem.contagem} {divItem.unidade}</span>
                  <span>RESULTADO: {renderResultado(d)}</span>
                </button>
              );
            })
          )}
        </div>
      </main>
    );
  }

  if (!carregado) {
    return (
      <main className="page center">
        <div className="loader" />
        <p>Carregando sistema...</p>
      </main>
    );
  }

  if (!usuarioLogado) return renderLogin();

  if (aba === 'inicio') return renderMenu();
  if (aba === 'contagem') return renderContagem();
  if (aba === 'divergencias') return renderDivergencias();

  return renderMenu();
}
