const ABA_CADASTROS = "CADASTROS";
const ABA_OPERADORES = "OPERADORES";
const ABA_LOG = "LOG_OPERADORES";

const ID_PLANILHA = "COLOQUE_AQUI_O_ID_DA_PLANILHA";

/* ================== WEB ================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Painel de Operadores')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/* ================== UTIL ================== */

function normalizar_(txt) {
  return String(txt || '').trim();
}

function normalizarCPF_(txt) {
  return String(txt || '').replace(/\D/g, '').trim();
}

function formatarCPF_(valor) {
  const cpf = normalizarCPF_(valor);
  if (cpf.length !== 11) return normalizar_(valor);
  return cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function formatarDataHora_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  }
  return normalizar_(valor);
}

function formatarData_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return normalizar_(valor);
}

function converterDataInputParaDate_(valor) {
  const txt = normalizar_(valor);

  if (!txt) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) {
    return new Date(txt + 'T00:00:00');
  }

  const partes = txt.split('/');
  if (partes.length === 3) {
    return new Date(partes[2], Number(partes[1]) - 1, partes[0]);
  }

  const data = new Date(txt);
  return isNaN(data.getTime()) ? '' : data;
}

function hashSenha_(senha) {
  const texto = normalizar_(senha);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto);
  return Utilities.base64Encode(digest);
}

function obterAba_(nome) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(nome);
  if (!aba) throw new Error('Aba "' + nome + '" não encontrada.');
  return aba;
}

function garantirEstruturaOperadores_() {
  const aba = obterAba_(ABA_OPERADORES);
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, 9).setValues([[
      'Matrícula',
      'Nome',
      'Senha',
      'Perfil',
      'Status',
      'Último acesso',
      'Data criação',
      'Observação',
      'Trocar senha?'
    ]]);
    return;
  }

  if (aba.getLastColumn() < 9) {
    aba.getRange(1, 9).setValue('Trocar senha?');
  } else {
    const atual = normalizar_(aba.getRange(1, 9).getDisplayValue());
    if (!atual) aba.getRange(1, 9).setValue('Trocar senha?');
  }
}

function obterMapaColunasCadastros_() {
  const aba = obterAba_(ABA_CADASTROS);
  const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0];
  const mapa = {};

  cabecalhos.forEach(function(nome, i) {
    mapa[normalizar_(nome).toLowerCase()] = i + 1;
  });

  return {
    id: mapa['id'] || 1,
    dataCadastro: mapa['data cadastro'] || 2,
    dataAtualizacao: mapa['data atualização'] || mapa['data atualizacao'] || 3,
    nome: mapa['nome completo'] || 4,
    cpf: mapa['cpf'] || 5,
    cpfNormalizado: mapa['cpf normalizado'] || 6,
    rg: mapa['rg'] || 7,
    dataNasc: mapa['data nascimento'] || 8,
    telefone: mapa['telefone'] || 9,
    email: mapa['e-mail'] || mapa['email'] || 10,
    cep: mapa['cep'] || 11,
    endereco: mapa['endereço'] || mapa['endereco'] || 12,
    genero: mapa['gênero'] || mapa['genero'] || 13,
    idade: mapa['idade'] || 14,
    escolaridade: mapa['escolaridade'] || 15,
    cursos: mapa['cursos'] || 16,
    experiencia: mapa['experiência'] || mapa['experiencia'] || 17,
    areaAtuacao: mapa['área de atuação'] || mapa['area de atuação'] || mapa['área de atuacao'] || mapa['area de atuacao'] || 18,
    objetivo: mapa['objetivo profissional'] || 19,
    linkFoto: mapa['link da foto'] || 20,
    linkCurriculo: mapa['link do currículo'] || mapa['link do curriculo'] || 21,
    linkComprovante: mapa['link do comprovante'] || 22,
    status: mapa['status'] || 23,
    estadoCivil: mapa['estado civil'] || 24,
    filhos: mapa['possui filhos'] || 25,
    pcd: mapa['é pcd?'] || mapa['e pcd?'] || 26,
    cnh: mapa['possui cnh'] || 27,
    categoriaCnh: mapa['categoria da cnh'] || 28,
    idioma: mapa['fala algum idioma'] || 29,
    primeiroEmprego: mapa['1º emprego'] || mapa['1o emprego'] || 30
  };
}

function serializarValorAuditoria_(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  }
  return normalizar_(valor);
}

function montarAuditoriaCampos_(antes, depois) {
  const campos = Object.keys(depois);
  const alteracoes = [];

  campos.forEach(function(campo) {
    const valorAntes = serializarValorAuditoria_(antes[campo]);
    const valorDepois = serializarValorAuditoria_(depois[campo]);
    if (valorAntes !== valorDepois) {
      alteracoes.push({
        campo: campo,
        antes: valorAntes,
        depois: valorDepois
      });
    }
  });

  return alteracoes;
}

function tentarJson_(texto) {
  try {
    return JSON.parse(texto);
  } catch (e) {
    return null;
  }
}

/* ================== LOG ================== */

function registrarLog_(matricula, nome, perfil, acao, cpf, detalhe) {
  const aba = obterAba_(ABA_LOG);
  aba.appendRow([
    new Date(),
    normalizar_(matricula),
    normalizar_(nome),
    normalizar_(perfil),
    normalizar_(acao),
    normalizar_(cpf),
    normalizar_(detalhe)
  ]);
}

/* ================== OPERADORES ================== */

function localizarOperadorPorMatricula_(matricula) {
  garantirEstruturaOperadores_();
  const aba = obterAba_(ABA_OPERADORES);
  const dados = aba.getDataRange().getValues();
  const matriculaBusca = normalizar_(matricula);

  for (let i = 1; i < dados.length; i++) {
    if (normalizar_(dados[i][0]) === matriculaBusca) {
      return {
        linha: i + 1,
        matricula: normalizar_(dados[i][0]),
        nome: normalizar_(dados[i][1]),
        senhaHash: normalizar_(dados[i][2]),
        perfil: normalizar_(dados[i][3]).toUpperCase(),
        status: normalizar_(dados[i][4]).toUpperCase(),
        ultimoAcesso: dados[i][5],
        dataCriacao: dados[i][6],
        observacao: normalizar_(dados[i][7]),
        trocarSenha: normalizar_(dados[i][8]).toUpperCase() === 'SIM'
      };
    }
  }

  return null;
}

function autenticarAdmin_(matricula, senha) {
  const login = loginOperador(matricula, senha, true);

  if (!login.ok) {
    throw new Error(login.msg || 'Falha na autenticação do administrador.');
  }

  if (String(login.perfil || '').toUpperCase() !== 'ADMIN') {
    throw new Error('Acesso negado. Apenas ADMIN pode executar esta ação.');
  }

  return login;
}

function listarOperadoresAdmin(adminMatricula, adminSenha) {
  autenticarAdmin_(adminMatricula, adminSenha);
  const aba = obterAba_(ABA_OPERADORES);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const dados = aba.getRange(2, 1, ultimaLinha - 1, Math.max(9, aba.getLastColumn())).getValues();

  return dados.map(function(linha) {
    return {
      matricula: normalizar_(linha[0]),
      nome: normalizar_(linha[1]),
      perfil: normalizar_(linha[3]).toUpperCase(),
      status: normalizar_(linha[4]).toUpperCase(),
      ultimoAcesso: formatarDataHora_(linha[5]),
      dataCriacao: formatarDataHora_(linha[6]),
      observacao: normalizar_(linha[7]),
      trocarSenha: normalizar_(linha[8]).toUpperCase() === 'SIM' ? 'SIM' : 'NÃO'
    };
  }).sort(function(a, b) {
    return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
  });
}

function salvarOperadorAdmin(adminMatricula, adminSenha, dadosOperador) {
  const admin = loginOperador(adminMatricula, adminSenha, true);
  if (!admin.ok) throw new Error('Falha na autenticação.');
  const aba = obterAba_(ABA_OPERADORES);

  const matricula = normalizar_(dadosOperador.matricula);
  const nome = normalizar_(dadosOperador.nome).toUpperCase();
  const senha = normalizar_(dadosOperador.senha);
  const perfil = normalizar_(dadosOperador.perfil).toUpperCase() || 'OPERADOR';
  const status = normalizar_(dadosOperador.status).toUpperCase() || 'ATIVO';
  const observacao = normalizar_(dadosOperador.observacao);
  const trocarSenha = normalizar_(dadosOperador.trocarSenha).toUpperCase() === 'SIM' ? 'SIM' : 'NÃO';

  if (!matricula) throw new Error('Informe a matrícula.');
  if (!nome) throw new Error('Informe o nome.');
  if (['ADMIN', 'OPERADOR'].indexOf(perfil) === -1) throw new Error('Perfil inválido.');
  if (['ATIVO', 'INATIVO'].indexOf(status) === -1) throw new Error('Status inválido.');

  const operadorExistente = localizarOperadorPorMatricula_(matricula);

  if (operadorExistente) {
    const antes = {
      nome: operadorExistente.nome,
      perfil: operadorExistente.perfil,
      status: operadorExistente.status,
      observacao: operadorExistente.observacao,
      trocarSenha: operadorExistente.trocarSenha ? 'SIM' : 'NÃO',
      senha: senha ? '[DEFINIDA]' : '[SEM ALTERAÇÃO]'
    };

    aba.getRange(operadorExistente.linha, 2).setValue(nome);
    aba.getRange(operadorExistente.linha, 4).setValue(perfil);
    aba.getRange(operadorExistente.linha, 5).setValue(status);
    aba.getRange(operadorExistente.linha, 8).setValue(observacao);
    aba.getRange(operadorExistente.linha, 9).setValue(trocarSenha);

    let senhaDepois = '[SEM ALTERAÇÃO]';

    if (senha) {
      if (senha.length < 4) throw new Error('Senha mínimo 4 caracteres.');
      aba.getRange(operadorExistente.linha, 3).setValue(hashSenha_(senha));
      senhaDepois = '[REDEFINIDA]';
    }

    const depois = {
      nome: nome,
      perfil: perfil,
      status: status,
      observacao: observacao,
      trocarSenha: trocarSenha,
      senha: senhaDepois
    };

    registrarLog_(
      admin.matricula,
      admin.nome,
      admin.perfil,
      'ADMIN_EDITOU_OPERADOR',
      '',
      JSON.stringify({
        operador: matricula,
        alteracoes: montarAuditoriaCampos_(antes, depois)
      })
    );

    return 'Operador atualizado com sucesso.';
  }

  if (!senha) throw new Error('Informe a senha do novo operador.');
  if (senha.length < 4) throw new Error('Senha mínimo 4 caracteres.');

  aba.appendRow([
    matricula,
    nome,
    hashSenha_(senha),
    perfil,
    status,
    '',
    new Date(),
    observacao,
    trocarSenha
  ]);

  registrarLog_(
    admin.matricula,
    admin.nome,
    admin.perfil,
    'ADMIN_CADASTROU_OPERADOR',
    '',
    JSON.stringify({
      operador: matricula,
      alteracoes: [
        { campo: 'matricula', antes: '', depois: matricula },
        { campo: 'nome', antes: '', depois: nome },
        { campo: 'perfil', antes: '', depois: perfil },
        { campo: 'status', antes: '', depois: status },
        { campo: 'observacao', antes: '', depois: observacao },
        { campo: 'trocarSenha', antes: '', depois: trocarSenha },
        { campo: 'senha', antes: '', depois: '[DEFINIDA]' }
      ]
    })
  );

  return 'Operador cadastrado com sucesso.';
}

function trocarSenhaObrigatoria(matricula, senhaAtual, novaSenha) {
  garantirEstruturaOperadores_();
  const aba = obterAba_(ABA_OPERADORES);
  const operador = localizarOperadorPorMatricula_(matricula);

  if (!operador) throw new Error('Operador não encontrado.');
  if (operador.status !== 'ATIVO') throw new Error('Usuário inativo.');
  if (hashSenha_(senhaAtual) !== operador.senhaHash) throw new Error('Senha atual incorreta.');

  const novaSenhaLimpa = normalizar_(novaSenha);
  if (!novaSenhaLimpa) throw new Error('Informe a nova senha.');
  if (novaSenhaLimpa.length < 4) throw new Error('Nova senha mínimo 4 caracteres.');
  if (hashSenha_(novaSenhaLimpa) === operador.senhaHash) throw new Error('A nova senha não pode ser igual à senha atual.');

  aba.getRange(operador.linha, 3).setValue(hashSenha_(novaSenhaLimpa));
  aba.getRange(operador.linha, 9).setValue('NÃO');

  registrarLog_(
    operador.matricula,
    operador.nome,
    operador.perfil,
    'TROCOU_SENHA_OBRIGATORIA',
    '',
    'Senha alterada no primeiro acesso'
  );

  return { ok: true, msg: 'Senha alterada com sucesso.' };
}

/* ================== LOGIN ================== */

function loginOperador(matricula, senha, ignorarTrocaObrigatoria) {
  garantirEstruturaOperadores_();
  const aba = obterAba_(ABA_OPERADORES);
  const dados = aba.getDataRange().getValues();

  const matriculaInformada = normalizar_(matricula);
  const hashInformado = hashSenha_(senha);

  for (let i = 1; i < dados.length; i++) {
    const mat = normalizar_(dados[i][0]);
    const nome = normalizar_(dados[i][1]);
    const senhaSalva = normalizar_(dados[i][2]);
    const perfil = normalizar_(dados[i][3]);
    const status = normalizar_(dados[i][4]).toUpperCase();
    const trocarSenha = normalizar_(dados[i][8]).toUpperCase() === 'SIM';

    if (mat === matriculaInformada) {
      if (status !== 'ATIVO') {
        registrarLog_(mat, nome, perfil, 'LOGIN_BLOQUEADO', '', 'Usuário inativo');
        return { ok: false, msg: 'Usuário inativo.' };
      }

      if (senhaSalva !== hashInformado) {
        registrarLog_(mat, nome, perfil, 'LOGIN_ERRO', '', 'Senha incorreta');
        return { ok: false, msg: 'Senha incorreta.' };
      }

      if (aba.getLastColumn() >= 6) {
        aba.getRange(i + 1, 6).setValue(new Date());
      }

      registrarLog_(mat, nome, perfil, 'LOGIN_OK', '', 'Login realizado com sucesso');

      return {
        ok: true,
        nome: nome,
        perfil: perfil,
        matricula: mat,
        isAdmin: String(perfil || '').toUpperCase() === 'ADMIN',
        exigirTrocaSenha: ignorarTrocaObrigatoria ? false : trocarSenha
      };
    }
  }

  return { ok: false, msg: 'Matrícula não encontrada.' };
}

/* ================== LISTAGEM ================== */

function listarCadastros() {
  const aba = obterAba_(ABA_CADASTROS);
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  const col = obterMapaColunasCadastros_();

  if (ultimaLinha < 2) return [];

  const dados = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
  const lista = [];

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const nome = linha[col.nome - 1] || '';
    const cpf = linha[col.cpf - 1] || '';
    const status = linha[col.status - 1] || 'ATIVO';

    if (!nome && !cpf) continue;

    lista.push({
      id: linha[col.id - 1] || '',
      dataCadastro: formatarDataHora_(linha[col.dataCadastro - 1]),
      dataAtualizacao: formatarDataHora_(linha[col.dataAtualizacao - 1]),
      nome: nome,
      cpf: cpf,
      cpfNormalizado: normalizarCPF_(linha[(col.cpfNormalizado - 1)] || cpf),
      telefone: linha[col.telefone - 1] || '',
      email: linha[col.email - 1] || '',
      areaAtuacao: linha[col.areaAtuacao - 1] || '',
      escolaridade: linha[col.escolaridade - 1] || '',
      linkFoto: linha[col.linkFoto - 1] || '',
      pcd: linha[col.pcd - 1] || '',
      cnh: linha[col.cnh - 1] || '',
      categoriaCnh: linha[col.categoriaCnh - 1] || '',
      status: status
    });
  }

  lista.sort(function(a, b) {
    return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
  });

  return lista;
}

/* ================== DETALHE ================== */

function obterCadastroPorLinha_(linha, col) {
  return {
    id: linha[col.id - 1] || '',
    dataCadastro: formatarDataHora_(linha[col.dataCadastro - 1]),
    dataAtualizacao: formatarDataHora_(linha[col.dataAtualizacao - 1]),
    nome: linha[col.nome - 1] || '',
    cpf: linha[col.cpf - 1] || '',
    rg: linha[col.rg - 1] || '',
    dataNasc: formatarData_(linha[col.dataNasc - 1]),
    telefone: linha[col.telefone - 1] || '',
    email: linha[col.email - 1] || '',
    cep: linha[col.cep - 1] || '',
    endereco: linha[col.endereco - 1] || '',
    genero: linha[col.genero - 1] || '',
    idade: linha[col.idade - 1] || '',
    escolaridade: linha[col.escolaridade - 1] || '',
    cursos: linha[col.cursos - 1] || '',
    experiencia: linha[col.experiencia - 1] || '',
    areaAtuacao: linha[col.areaAtuacao - 1] || '',
    objetivo: linha[col.objetivo - 1] || '',
    linkFoto: linha[col.linkFoto - 1] || '',
    linkCurriculo: linha[col.linkCurriculo - 1] || '',
    linkComprovante: linha[col.linkComprovante - 1] || '',
    status: linha[col.status - 1] || 'ATIVO',
    estadoCivil: linha[col.estadoCivil - 1] || '',
    filhos: linha[col.filhos - 1] || '',
    pcd: linha[col.pcd - 1] || '',
    cnh: linha[col.cnh - 1] || '',
    categoriaCnh: linha[col.categoriaCnh - 1] || '',
    idioma: linha[col.idioma - 1] || '',
    primeiroEmprego: linha[col.primeiroEmprego - 1] || ''
  };
}

function obterCadastroPorCPF(cpf, matriculaOperador, nomeOperador, perfilOperador) {
  const aba = obterAba_(ABA_CADASTROS);
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  const col = obterMapaColunasCadastros_();

  if (ultimaLinha < 2) return null;

  const cpfBusca = normalizarCPF_(cpf);
  const dados = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const cpfLinhaNormalizado = normalizarCPF_(linha[(col.cpfNormalizado - 1)] || '');
    const cpfLinhaFormatado = normalizarCPF_(linha[col.cpf - 1] || '');

    if (cpfLinhaNormalizado === cpfBusca || cpfLinhaFormatado === cpfBusca) {
      registrarLog_(
        matriculaOperador,
        nomeOperador,
        perfilOperador,
        'ABRIU_CADASTRO',
        cpfBusca,
        'Abriu ficha do cadastro'
      );

      return obterCadastroPorLinha_(linha, col);
    }
  }

  return null;
}

/* ================== EDIÇÃO ADMIN DE CADASTROS ================== */

function atualizarCadastroAdmin(adminMatricula, adminSenha, dadosCadastro) {
  const admin = loginOperador(adminMatricula, adminSenha, true);
  if (!admin.ok) throw new Error('Falha na autenticação.');
  const aba = obterAba_(ABA_CADASTROS);
  const col = obterMapaColunasCadastros_();
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();

  if (ultimaLinha < 2) throw new Error('Nenhum cadastro encontrado.');

  const idBusca = normalizar_(dadosCadastro.id);
  const cpfOriginal = normalizarCPF_(dadosCadastro.cpfOriginal || dadosCadastro.cpf);
  const cpfNovo = normalizarCPF_(dadosCadastro.cpf);
  const dados = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();

  if (!normalizar_(dadosCadastro.nome)) throw new Error('Informe o nome completo.');
  if (!cpfNovo) throw new Error('Informe o CPF.');

  let indiceLinha = -1;

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const idLinha = normalizar_(linha[col.id - 1]);
    const cpfLinha = normalizarCPF_(linha[col.cpfNormalizado - 1] || linha[col.cpf - 1]);

    if ((idBusca && idLinha === idBusca) || (!idBusca && cpfLinha === cpfOriginal)) {
      indiceLinha = i;
      break;
    }
  }

  if (indiceLinha === -1) throw new Error('Cadastro não encontrado para edição.');

  for (let j = 0; j < dados.length; j++) {
    if (j === indiceLinha) continue;
    const cpfExistente = normalizarCPF_(dados[j][col.cpfNormalizado - 1] || dados[j][col.cpf - 1]);
    if (cpfNovo && cpfExistente === cpfNovo) throw new Error('Já existe outro cadastro com este CPF.');
  }

  const linhaPlanilha = indiceLinha + 2;
  const antesObj = obterCadastroPorLinha_(dados[indiceLinha], col);

  aba.getRange(linhaPlanilha, col.nome).setValue(normalizar_(dadosCadastro.nome));
  aba.getRange(linhaPlanilha, col.cpf).setValue(formatarCPF_(cpfNovo));
  aba.getRange(linhaPlanilha, col.cpfNormalizado).setValue(cpfNovo);
  aba.getRange(linhaPlanilha, col.rg).setValue(normalizar_(dadosCadastro.rg));
  aba.getRange(linhaPlanilha, col.dataNasc).setValue(converterDataInputParaDate_(dadosCadastro.dataNasc));
  aba.getRange(linhaPlanilha, col.telefone).setValue(normalizar_(dadosCadastro.telefone));
  aba.getRange(linhaPlanilha, col.email).setValue(normalizar_(dadosCadastro.email));
  aba.getRange(linhaPlanilha, col.cep).setValue(normalizar_(dadosCadastro.cep));
  aba.getRange(linhaPlanilha, col.endereco).setValue(normalizar_(dadosCadastro.endereco));
  aba.getRange(linhaPlanilha, col.genero).setValue(normalizar_(dadosCadastro.genero));
  aba.getRange(linhaPlanilha, col.idade).setValue(normalizar_(dadosCadastro.idade));
  aba.getRange(linhaPlanilha, col.escolaridade).setValue(normalizar_(dadosCadastro.escolaridade));
  if(normalizar_(dadosCadastro.estadoCivil) !== '') aba.getRange(linhaPlanilha, col.estadoCivil).setValue(normalizar_(dadosCadastro.estadoCivil));
  if(normalizar_(dadosCadastro.filhos) !== '') aba.getRange(linhaPlanilha, col.filhos).setValue(normalizar_(dadosCadastro.filhos));
  if(normalizar_(dadosCadastro.pcd) !== '') aba.getRange(linhaPlanilha, col.pcd).setValue(normalizar_(dadosCadastro.pcd));
  if(normalizar_(dadosCadastro.cnh) !== '') aba.getRange(linhaPlanilha, col.cnh).setValue(normalizar_(dadosCadastro.cnh));
  if(normalizar_(dadosCadastro.categoriaCnh) !== '') aba.getRange(linhaPlanilha, col.categoriaCnh).setValue(normalizar_(dadosCadastro.categoriaCnh));
  if(normalizar_(dadosCadastro.idioma) !== '') aba.getRange(linhaPlanilha, col.idioma).setValue(normalizar_(dadosCadastro.idioma));
  if(normalizar_(dadosCadastro.primeiroEmprego) !== '') aba.getRange(linhaPlanilha, col.primeiroEmprego).setValue(normalizar_(dadosCadastro.primeiroEmprego));
  aba.getRange(linhaPlanilha, col.cursos).setValue(normalizar_(dadosCadastro.cursos));
  aba.getRange(linhaPlanilha, col.experiencia).setValue(normalizar_(dadosCadastro.experiencia));
  aba.getRange(linhaPlanilha, col.areaAtuacao).setValue(normalizar_(dadosCadastro.areaAtuacao));
  aba.getRange(linhaPlanilha, col.objetivo).setValue(normalizar_(dadosCadastro.objetivo));
  const statusPermitidos = [
  'ATIVO',
  'INATIVO',
  'AGUARDANDO CARTA',
  'ENCAMINHADO',
  'CONTRATADO'
];

let statusFinal = normalizar_(dadosCadastro.status).toUpperCase();

if (statusPermitidos.indexOf(statusFinal) === -1) {
  statusFinal = 'ATIVO';
}

aba.getRange(linhaPlanilha, col.status).setValue(statusFinal);

  if (col.dataAtualizacao) aba.getRange(linhaPlanilha, col.dataAtualizacao).setValue(new Date());

  const linhaAtualizada = aba.getRange(linhaPlanilha, 1, 1, ultimaColuna).getValues()[0];
  const depoisObj = obterCadastroPorLinha_(linhaAtualizada, col);

  const alteracoesCadastro = montarAuditoriaCampos_(antesObj, depoisObj);

  registrarLog_(
    admin.matricula,
    admin.nome,
    admin.perfil,
    'ADMIN_EDITOU_CADASTRO',
    cpfNovo,
    JSON.stringify({
      id: depoisObj.id,
      alteracoes: alteracoesCadastro
    })
  );

  const statusAntes = normalizar_(antesObj.status).toUpperCase();
  const statusDepois = normalizar_(depoisObj.status).toUpperCase();

  if (statusAntes !== statusDepois) {
    registrarLog_(
      admin.matricula,
      admin.nome,
      admin.perfil,
      'STATUS_ALTERADO',
      cpfNovo,
      JSON.stringify({
        id: depoisObj.id,
        nome: depoisObj.nome,
        statusAntes: statusAntes,
        statusDepois: statusDepois
      })
    );
  }

  return depoisObj;
}

/* ================== AUDITORIA ADMIN ================== */

function listarLogsAdmin(adminMatricula, adminSenha, filtros) {
  autenticarAdmin_(adminMatricula, adminSenha);
  const aba = obterAba_(ABA_LOG);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const quantidade = Math.min(ultimaLinha - 1, 1000);
  const inicio = ultimaLinha - quantidade + 1;
  const dados = aba.getRange(inicio, 1, quantidade, 7).getValues();

  const termo = normalizar_(filtros && filtros.termo).toLowerCase();
  const acaoFiltro = normalizar_(filtros && filtros.acao).toUpperCase();
  const matriculaFiltro = normalizar_(filtros && filtros.matricula);
  const cpfFiltro = normalizarCPF_(filtros && filtros.cpf);

  const lista = dados.map(function(linha) {
    const detalheBruto = normalizar_(linha[6]);
    const json = tentarJson_(detalheBruto);

    return {
      dataHora: formatarDataHora_(linha[0]),
      matricula: normalizar_(linha[1]),
      nome: normalizar_(linha[2]),
      perfil: normalizar_(linha[3]),
      acao: normalizar_(linha[4]).toUpperCase(),
      cpf: normalizar_(linha[5]),
      detalhe: detalheBruto,
      detalheJson: json
    };
  }).reverse();

  return lista.filter(function(item) {
    const resumo = item.detalheJson && item.detalheJson.alteracoes ? item.detalheJson.alteracoes.map(function(a) {
      return a.campo + ': ' + a.antes + ' -> ' + a.depois;
    }).join(' | ') : item.detalhe;

    const blob = [
      item.dataHora,
      item.matricula,
      item.nome,
      item.perfil,
      item.acao,
      item.cpf,
      item.detalhe,
      resumo
    ].join(' ').toLowerCase();

    const passouTermo = !termo || blob.indexOf(termo) !== -1;
    const passouAcao = !acaoFiltro || item.acao === acaoFiltro;
    const passouMatricula = !matriculaFiltro || item.matricula === matriculaFiltro;
    const passouCpf = !cpfFiltro || normalizarCPF_(item.cpf).indexOf(cpfFiltro) !== -1;

    return passouTermo && passouAcao && passouMatricula && passouCpf;
  });
}

function listarHistoricoStatusAdmin(adminMatricula, adminSenha, filtros) {
  const admin = loginOperador(adminMatricula, adminSenha, true);
  if (!admin.ok) throw new Error('Falha na autenticação.');
  const aba = obterAba_(ABA_LOG);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const quantidade = Math.min(ultimaLinha - 1, 1000);
  const inicio = ultimaLinha - quantidade + 1;
  const dados = aba.getRange(inicio, 1, quantidade, 7).getValues();

  const cpfFiltro = normalizarCPF_(filtros && filtros.cpf);
  const nomeFiltro = normalizar_(filtros && filtros.nome).toLowerCase();
  const statusFiltro = normalizar_(filtros && filtros.status).toUpperCase();

  const lista = dados.map(function(linha) {
    const acao = normalizar_(linha[4]).toUpperCase();
    const detalheBruto = normalizar_(linha[6]);
    const json = tentarJson_(detalheBruto);

    return {
      dataHora: formatarDataHora_(linha[0]),
      matricula: normalizar_(linha[1]),
      operador: normalizar_(linha[2]),
      perfil: normalizar_(linha[3]),
      acao: acao,
      cpf: normalizar_(linha[5]),
      detalheJson: json
    };
  }).filter(function(item) {
    return item.acao === 'STATUS_ALTERADO' && item.detalheJson;
  }).map(function(item) {
    return {
      dataHora: item.dataHora,
      matricula: item.matricula,
      operador: item.operador,
      perfil: item.perfil,
      cpf: item.cpf,
      id: normalizar_(item.detalheJson.id),
      nome: normalizar_(item.detalheJson.nome),
      statusAntes: normalizar_(item.detalheJson.statusAntes).toUpperCase(),
      statusDepois: normalizar_(item.detalheJson.statusDepois).toUpperCase()
    };
  }).reverse();

  return lista.filter(function(item) {
    const passouCpf = !cpfFiltro || normalizarCPF_(item.cpf).indexOf(cpfFiltro) !== -1;
    const passouNome = !nomeFiltro || String(item.nome || '').toLowerCase().indexOf(nomeFiltro) !== -1;
    const passouStatus = !statusFiltro || item.statusAntes === statusFiltro || item.statusDepois === statusFiltro;
    return passouCpf && passouNome && passouStatus;
  });
}
