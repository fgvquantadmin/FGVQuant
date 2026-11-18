/**
 * Agente do Instagram. Fica no mesmo projeto do Apps Script que atualizar-site.gs e formulario.gs.
 *
 * Todo dia às 5h (antes da atualização das 6h):
 *  1. busca os posts novos do Instagram da liga (Instagram API com login do Instagram);
 *  2. o Claude lê legenda e imagem de cada post e decide se é Evento, Newsletter, Fundo mensal ou nada;
 *  3. preenche os campos da aba certa e põe o post na aba "Instagram" com o status "Revisar".
 * Ninguém publica nada sozinho: alguém confere a linha, corrige o que precisar na coluna "Campos",
 * marca "Aprovar" e usa o menu Site → Instagram → Publicar aprovados. A linha vai para a aba de
 * destino pelo mesmo caminho do formulário (salvarLinha), com as mesmas validações, e o site é
 * atualizado na hora.
 *
 * Regra do site: nada inventado. O Claude só preenche o que está no post; o que faltar fica vazio
 * e aparece na coluna "Observação".
 *
 * Propriedades do script (menu Site → Instagram → Configurar): IG_TOKEN, IG_HOST, IG_ID,
 * ANTHROPIC_KEY, IG_DESDE. O token pode ser de dois tipos, e o Configurar descobre qual é:
 *  - token do Instagram (graph.instagram.com, vale 60 dias e é renovado toda semana);
 *  - token de usuário do sistema do login unificado (graph.facebook.com, não expira).
 */

const MODELO = "claude-sonnet-5-5";
const ABA_IG = "Instagram";
const ABAS_DESTINO = ["Eventos", "Newsletter", "Fundo mensal"];
const CABECALHO_IG = [
  "Data do post",
  "Link",
  "Legenda",
  "Aba",
  "Campos",
  "Observação",
  "Aprovar",
  "Status",
  "ID",
];
const COL = Object.fromEntries(CABECALHO_IG.map((c, i) => [c, i + 1]));
const IG = "https://graph.instagram.com";
const FB = "https://graph.facebook.com/v25.0";
const ID_PADRAO = "17841448100380830"; // @fgvquant

function igHost() {
  return PropertiesService.getScriptProperties().getProperty("IG_HOST") || IG;
}
function igId() {
  return PropertiesService.getScriptProperties().getProperty("IG_ID") || "me";
}

/** Testa o token nos dois tipos de API e devolve {host, id, usuario}, ou null. */
function detectarToken(token, id) {
  const tentativas = [
    [IG, "me"],
    [FB, id],
    [IG, id],
  ];
  for (const [host, conta] of tentativas) {
    const res = UrlFetchApp.fetch(host + "/" + conta + "?fields=username&access_token=" + token, {
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200)
      return { host: host, id: conta, usuario: JSON.parse(res.getContentText()).username };
  }
  return null;
}

// ── configuração ────────────────────────────────────────────────────────────────

function configurarInstagram() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const t = ui.prompt(
    "Token do Instagram",
    "Cole o token gerado no Meta for Developers (Casos de uso → API do Instagram → Gerar tokens de acesso → Gerar token):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (t.getSelectedButton() !== ui.Button.OK) return;
  const token = t.getResponseText().trim();
  const conta = detectarToken(token, ID_PADRAO);
  if (!conta)
    return ui.alert(
      "O Instagram recusou esse token. Gere outro no painel da Meta (Gerar token) e tente de novo.",
    );
  const usuario = conta.usuario;

  const k = ui.prompt(
    "Chave da API do Claude",
    "Cole a chave criada no Claude Console (console.anthropic.com → API keys). Começa com sk-ant-:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (k.getSelectedButton() !== ui.Button.OK) return;
  const chave = k.getResponseText().trim();
  if (!/^sk-ant-/.test(chave)) return ui.alert("Essa não parece uma chave do Claude.");

  const d = ui.prompt(
    "A partir de quando?",
    "Buscar posts publicados desde qual data? (dd/mm/aaaa; vazio = últimos 30 dias)",
    ui.ButtonSet.OK_CANCEL,
  );
  if (d.getSelectedButton() !== ui.Button.OK) return;
  let desde = new Date(Date.now() - 30 * 864e5);
  const m = d.getResponseText().trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) desde = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  props.setProperties({
    IG_TOKEN: token,
    IG_HOST: conta.host,
    IG_ID: conta.id,
    IG_TOKEN_DATA: String(Date.now()),
    ANTHROPIC_KEY: chave,
    IG_DESDE: desde.toISOString(),
  });

  // Só mexe nos agendamentos deste arquivo; o das 6h (atualizarSite) continua como está.
  ScriptApp.getProjectTriggers()
    .filter((g) => ["buscarInstagram", "renovarTokenInstagram"].includes(g.getHandlerFunction()))
    .forEach((g) => ScriptApp.deleteTrigger(g));
  ScriptApp.newTrigger("buscarInstagram").timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger("renovarTokenInstagram").timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();

  abaInstagram();
  ui.alert(
    "Pronto, conectado a @" +
      usuario +
      ".\n\nTodo dia às 5h o agente busca os posts novos e coloca na aba Instagram para revisão. " +
      "Para buscar agora: menu Site → Instagram → Buscar posts agora.",
  );
}

/** O token do Instagram vale 60 dias; renovar toda semana mantém ele vivo para sempre. */
function renovarTokenInstagram() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("IG_TOKEN");
  if (!token || igHost() !== IG) return; // token de usuário do sistema não expira
  const res = UrlFetchApp.fetch(
    IG + "/refresh_access_token?grant_type=ig_refresh_token&access_token=" + token,
    { muteHttpExceptions: true },
  );
  if (res.getResponseCode() !== 200) {
    avisarPorEmail(
      "O token do Instagram não renovou",
      "Resposta do Instagram:\n" +
        res.getContentText() +
        "\n\nGere um token novo e use o menu Site → Instagram → Configurar.",
    );
    return;
  }
  props.setProperties({
    IG_TOKEN: JSON.parse(res.getContentText()).access_token,
    IG_TOKEN_DATA: String(Date.now()),
  });
}

// ── buscar e classificar ────────────────────────────────────────────────────────

function buscarInstagram() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("IG_TOKEN");
  const chave = props.getProperty("ANTHROPIC_KEY");
  if (!token || !chave) throw new Error("Falta configurar: menu Site → Instagram → Configurar.");
  const desde = new Date(props.getProperty("IG_DESDE") || 0);

  const aba = abaInstagram();
  // Posts que deram erro na leitura (API fora do ar, por exemplo) são lidos de novo.
  const vistos = new Set();
  const comErro = [];
  if (aba.getLastRow() > 1)
    aba
      .getRange(2, 1, aba.getLastRow() - 1, CABECALHO_IG.length)
      .getValues()
      .forEach((l, i) => {
        if (/^Erro ao ler o post/.test(String(l[COL.Observação - 1]))) comErro.push(i + 2);
        else vistos.add(String(l[COL.ID - 1]));
      });
  comErro.reverse().forEach((linha) => aba.deleteRow(linha));

  const posts = postsDesde(token, desde).filter((p) => !vistos.has(p.id));
  if (!posts.length) return toast("Nenhum post novo no Instagram.");

  const contexto = contextoDasAbas();
  let paraRevisar = 0;
  // do mais antigo para o mais novo, para a aba ficar em ordem
  posts.reverse().forEach((p) => {
    let r;
    try {
      r = classificar(p, contexto, chave);
    } catch (e) {
      r = { aba: "", campos: {}, observacao: "Erro ao ler o post: " + e.message };
    }
    const destino = ABAS_DESTINO.includes(r.aba) ? r.aba : "";
    const faltando = destino
      ? (OBRIGATORIAS[destino] || []).filter((c) => !r.campos[c] && c !== "Foto")
      : [];
    const obs = [r.observacao, faltando.length ? "Falta: " + faltando.join(", ") : ""]
      .filter(Boolean)
      .join(" · ");
    aba.appendRow([
      new Date(p.timestamp),
      p.permalink,
      p.caption || "",
      destino || "(ignorar)",
      Object.entries(r.campos)
        .filter(([, v]) => v !== "" && v !== null && v !== undefined)
        .map(([c, v]) => c + ": " + String(v).replace(/\s*\n\s*/g, " "))
        .join("\n"),
      obs,
      false,
      destino ? "Revisar" : "Ignorado",
      p.id,
    ]);
    aba.getRange(aba.getLastRow(), COL.Aprovar).insertCheckboxes();
    if (destino) paraRevisar++;
  });

  const msg =
    posts.length +
    " post(s) novo(s) no Instagram, " +
    paraRevisar +
    " para revisar na aba Instagram.";
  toast(msg);
  if (paraRevisar)
    avisarPorEmail(
      "FGV Quant: " + paraRevisar + " post(s) do Instagram para revisar",
      msg +
        "\n\nConfira a coluna Campos, marque Aprovar e use o menu Site → Instagram → Publicar aprovados.\n\n" +
        SpreadsheetApp.getActive().getUrl(),
    );
}

function postsDesde(token, desde) {
  const campos = "id,caption,media_type,media_url,thumbnail_url,permalink,timestamp";
  let url = igHost() + "/" + igId() + "/media?fields=" + campos + "&limit=25&access_token=" + token;
  const posts = [];
  for (let pagina = 0; url && pagina < 8; pagina++) {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200)
      throw new Error("O Instagram recusou o pedido: " + res.getContentText());
    const json = JSON.parse(res.getContentText());
    const lote = json.data || [];
    posts.push(...lote.filter((p) => new Date(p.timestamp) >= desde));
    // a API devolve do mais novo para o mais antigo: parou quando passou da data
    if (lote.some((p) => new Date(p.timestamp) < desde)) break;
    url = json.paging && json.paging.next;
  }
  return posts;
}

/** Colunas, tipos, listas e notas das abas de destino, lidos da própria planilha (formulario.gs). */
function contextoDasAbas() {
  const ss = SpreadsheetApp.getActive();
  const abas = ABAS_DESTINO.map((nome) => {
    const campos = camposDaAba(nome).filter((c) => c.tipo !== "arquivo");
    const linhas = campos.map((c) => {
      let s = "  - " + c.nome + " (" + FORMATO[c.tipo] + ")";
      if (c.opcoes) s += " opções: " + c.opcoes.join(" | ");
      if (c.obrigatoria) s += " [obrigatório]";
      if (c.nota) s += " — " + c.nota.replace(/\s+/g, " ");
      return s;
    });
    return "Aba " + nome + ":\n" + linhas.join("\n");
  }).join("\n\n");
  const parceiros = ss.getSheetByName("Parceiros");
  const nomes =
    parceiros && parceiros.getLastRow() > 4
      ? parceiros
          .getRange(5, 1, parceiros.getLastRow() - 4, 1)
          .getValues()
          .map((r) => r[0])
          .filter(Boolean)
      : [];
  return abas + "\n\nParceiros cadastrados (use só estes nomes, exatamente assim): " + nomes.join(", ");
}

const FORMATO = {
  texto: "texto curto",
  longo: "texto, uma linha só",
  lista: "uma das opções",
  data: "dd/mm/aaaa",
  mes: "mm/aaaa",
  semestre: "aaaa.1 ou aaaa.2",
  numero: "número, ponto decimal",
  link: "link começando com https://",
};

function classificar(post, contexto, chave) {
  const conteudo = [];
  const img = imagemDoPost(post);
  if (img) conteudo.push({ type: "image", source: { type: "base64", media_type: img.tipo, data: img.base64 } });
  conteudo.push({
    type: "text",
    text:
      "Post do Instagram da FGV Quant (liga de finanças quantitativas da FGV-SP).\n" +
      "Publicado em: " +
      post.timestamp +
      "\nLink do post: " +
      post.permalink +
      "\nTipo: " +
      post.media_type +
      "\nLegenda:\n" +
      (post.caption || "(sem legenda)"),
  });

  const sistema =
    "Você alimenta a planilha que gera o site da FGV Quant. Decida em qual aba este post entra e preencha os campos dela.\n\n" +
    "Abas:\n" +
    "- Eventos: palestra, aula, visita, competição ou marco da liga que aconteceu ou vai acontecer.\n" +
    "- Newsletter: anúncio de uma edição da newsletter FGV Quant Insights.\n" +
    "- Fundo mensal: resultado mensal do FGV Quant Capital (fundo simulado), só se o post trouxer o retorno do mês em número.\n" +
    "- ignorar: processo seletivo, data comemorativa, meme, repost, qualquer outra coisa.\n\n" +
    "Regras:\n" +
    "1. Nunca invente. Preencha um campo só se a informação estiver na legenda ou na imagem. Na dúvida, deixe de fora.\n" +
    "2. Respeite o formato de cada campo e use só as opções da lista quando houver.\n" +
    "3. Colunas '(inglês)': traduza o campo correspondente que você preencheu.\n" +
    "4. Fonte (Eventos) e Link (Newsletter): use o link que aparecer na legenda; se não houver, o link do post.\n" +
    "5. Newsletter: Edição é o mês da edição (mm/aaaa); Resumo com até 400 caracteres, nas suas palavras; não preencha Texto.\n" +
    "6. Data do evento é a data em que ele acontece, não a do post. Se o post não disser o dia, use só o que souber e explique em observacao.\n" +
    "7. observacao: uma frase curta para quem vai revisar (o que ficou incerto). Vazio se estiver tudo claro.\n" +
    "8. Texto em português do Brasil, sem hífen como conector entre frases.\n\n" +
    "Responda sempre chamando a ferramenta registrar, uma única vez, mesmo quando a aba for ignorar.\n\n" +
    "Campos de cada aba:\n" +
    contexto;

  const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": chave, "anthropic-version": "2023-06-01" },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: MODELO,
      max_tokens: 2000,
      system: sistema,
      tools: [
        {
          name: "registrar",
          description: "Registra em qual aba o post entra e os campos preenchidos.",
          input_schema: {
            type: "object",
            properties: {
              aba: { type: "string", enum: ABAS_DESTINO.concat(["ignorar"]) },
              campos: {
                type: "object",
                description: "Nome exato da coluna → valor. Só os campos que o post sustenta.",
                additionalProperties: { type: "string" },
              },
              observacao: { type: "string" },
            },
            required: ["aba", "campos", "observacao"],
          },
        },
      ],
      tool_choice: { type: "auto" },
      messages: [{ role: "user", content: conteudo }],
    }),
  });
  if (res.getResponseCode() !== 200) throw new Error("API do Claude: " + res.getContentText());
  const blocos = JSON.parse(res.getContentText()).content || [];
  const uso = blocos.find((b) => b.type === "tool_use");
  if (uso) return uso.input;
  // sem a ferramenta: tenta ler um JSON no texto da resposta
  const texto = blocos.filter((b) => b.type === "text").map((b) => b.text).join("\n");
  const json = texto.match(/\{[\s\S]*\}/);
  if (json) {
    const r = JSON.parse(json[0]);
    if (r.aba) return { aba: r.aba, campos: r.campos || {}, observacao: r.observacao || "" };
  }
  throw new Error("o Claude não devolveu a classificação");
}

/** Primeira imagem do post (a capa, em vídeo ou carrossel), em base64. */
function imagemDoPost(post) {
  const url = post.media_type === "VIDEO" ? post.thumbnail_url : post.media_url;
  if (!url) return null;
  try {
    const blob = UrlFetchApp.fetch(url).getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 4.5 * 1024 * 1024) return null; // limite da API
    const tipo = blob.getContentType();
    if (!/^image\/(jpeg|png|webp|gif)$/.test(tipo)) return null;
    return { tipo: tipo, base64: Utilities.base64Encode(bytes) };
  } catch (e) {
    return null;
  }
}

// ── publicar o que foi aprovado ─────────────────────────────────────────────────

function publicarAprovados() {
  const aba = abaInstagram();
  const n = aba.getLastRow() - 1;
  if (n < 1) return toast("A aba Instagram está vazia.");
  const linhas = aba.getRange(2, 1, n, CABECALHO_IG.length).getValues();
  const token = PropertiesService.getScriptProperties().getProperty("IG_TOKEN");
  let publicadas = 0;
  let erros = 0;

  linhas.forEach((l, i) => {
    const linhaIg = i + 2;
    const destino = l[COL.Aba - 1];
    const status = String(l[COL.Status - 1]);
    if (l[COL.Aprovar - 1] !== true || /^Publicado/.test(status)) return;
    try {
      if (!ABAS_DESTINO.includes(destino))
        throw new Error("a coluna Aba precisa ser " + ABAS_DESTINO.join(", ") + ".");
      const lidos = lerCampos(l[COL.Campos - 1]);
      const campos = camposDaAba(destino);
      const valores = {};
      const arquivos = {};
      Object.entries(lidos).forEach(([nome, v]) => {
        const c = campos.find((x) => x.nome === nome);
        if (!c) throw new Error("a aba " + destino + " não tem a coluna '" + nome + "'.");
        valores[c.coluna] = paraFormulario(c, v);
      });
      const foto = campos.find((c) => c.nome === "Foto");
      if (foto && !valores[foto.coluna] && token) {
        const img = imagemDoPost(postAtualizado(l[COL.ID - 1], token));
        if (img)
          arquivos[foto.coluna] = {
            nome: "instagram." + img.tipo.split("/")[1].replace("jpeg", "jpg"),
            tipo: img.tipo,
            base64: img.base64,
          };
      }
      const linha = salvarLinha(destino, valores, arquivos);
      aba.getRange(linhaIg, COL.Status).setValue("Publicado (" + destino + ", linha " + linha + ")");
      publicadas++;
    } catch (e) {
      aba.getRange(linhaIg, COL.Status).setValue("Erro: " + e.message);
      erros++;
    }
  });

  if (!publicadas && !erros) return toast("Nenhuma linha marcada em Aprovar.");
  if (publicadas) {
    try {
      atualizarSite();
    } catch (e) {
      // sem Deploy Hook configurado: entra na atualização das 6h
    }
  }
  toast(
    publicadas +
      " publicada(s)" +
      (erros ? ", " + erros + " com erro (veja a coluna Status)" : "") +
      ". O site atualiza em uns 3 minutos.",
  );
}

/** "Campo: valor", um por linha, como o agente escreve e quem revisa edita. */
function lerCampos(texto) {
  const campos = {};
  String(texto || "")
    .split("\n")
    .forEach((linha) => {
      const i = linha.indexOf(":");
      if (i < 1) return;
      const nome = linha.slice(0, i).trim();
      const valor = linha.slice(i + 1).trim();
      if (valor) campos[nome] = valor;
    });
  return campos;
}

/** Converte do formato da planilha (dd/mm/aaaa, mm/aaaa) para o que salvarLinha espera. */
function paraFormulario(c, v) {
  let m;
  if (c.tipo === "data") {
    if ((m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)))
      return m[3] + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0");
    throw new Error("'" + c.nome + "': use dd/mm/aaaa.");
  }
  if (c.tipo === "mes") {
    if ((m = v.match(/^(\d{1,2})\/(\d{4})$/))) return m[2] + "-" + m[1].padStart(2, "0");
    throw new Error("'" + c.nome + "': use mm/aaaa.");
  }
  return v;
}

/** Os links de imagem do Instagram expiram; pede um novo na hora de publicar. */
function postAtualizado(id, token) {
  const res = UrlFetchApp.fetch(
    igHost() + "/" + id + "?fields=media_type,media_url,thumbnail_url&access_token=" + token,
    { muteHttpExceptions: true },
  );
  return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : {};
}

// ── utilidades ──────────────────────────────────────────────────────────────────

function abaInstagram() {
  const ss = SpreadsheetApp.getActive();
  let aba = ss.getSheetByName(ABA_IG);
  if (aba) return aba;
  aba = ss.insertSheet(ABA_IG);
  aba.getRange(1, 1, 1, CABECALHO_IG.length).setValues([CABECALHO_IG]).setFontWeight("bold");
  aba.setFrozenRows(1);
  aba.setColumnWidth(COL.Legenda, 320);
  aba.setColumnWidth(COL.Campos, 380);
  aba.setColumnWidth(COL.Observação, 260);
  aba.getRange("A:A").setNumberFormat("dd/mm/yyyy");
  aba.getRange("C:F").setWrap(true).setVerticalAlignment("top");
  aba
    .getRange(2, COL.Aba, aba.getMaxRows() - 1, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(ABAS_DESTINO.concat(["(ignorar)"]))
        .build(),
    );
  aba.hideColumns(COL.ID);
  aba
    .getRange(1, COL.Campos)
    .setNote(
      "Um campo por linha, no formato 'Coluna: valor'. Pode editar à vontade antes de aprovar. " +
        "Datas em dd/mm/aaaa, meses em mm/aaaa.",
    );
  return aba;
}

function toast(msg) {
  try {
    SpreadsheetApp.getActive().toast(msg, "Instagram", 10);
  } catch (e) {
    // rodando pelo agendamento, sem planilha aberta
  }
}

function avisarPorEmail(assunto, corpo) {
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), assunto, corpo);
  } catch (e) {
    // sem permissão de e-mail: fica só na planilha
  }
}
