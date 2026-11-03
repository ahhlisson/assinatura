/*
 * Assinaturas de e-mail para o Outlook
 * Arquivo gerado pelo gerador.html. Para atualizar funcionários ou o visual,
 * use o gerador e substitua este arquivo no site onde o add-in está hospedado.
 *
 * Funcionamento: quando alguém cria, responde ou encaminha um e-mail, o Outlook
 * dispara o evento OnNewMessageCompose e este código insere a assinatura da pessoa.
 */

var CONFIG = {
  "empresa": "Sua Empresa",
  "cor": "#1e407c",
  "logo": "",
  "larguraLogo": 120,
  "site": "www.suaempresa.com.br",
  "telefoneGeral": "",
  "endereco": "",
  "linkedin": "",
  "instagram": "",
  "aviso": "",
  "assinaturaCurtaNasRespostas": true
};
var FUNCIONARIOS = {
  "ana.souza@suaempresa.com.br": {
    "email": "ana.souza@suaempresa.com.br",
    "nome": "Ana Souza",
    "cargo": "Gerente de Marketing",
    "departamento": "Marketing",
    "telefone": "(21) 3000-0001"
  }
};

function esc(valor) {
  return String(valor == null ? "" : valor)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function telLink(numero) {
  var limpo = String(numero || "").replace(/[^\d+]/g, "");
  return '<a href="tel:' + esc(limpo) + '" style="color:#333333;text-decoration:none;">' + esc(numero) + "</a>";
}

function urlCompleta(url) {
  url = String(url || "").trim();
  if (!url) return "";
  return /^https?:\/\//i.test(url) ? url : "https://" + url;
}

function urlSemProtocolo(url) {
  return String(url || "").replace(/^https?:\/\//i, "").replace(/\/$/, "");
}

function dataISO(d) {
  function dois(n) { return (n < 10 ? "0" : "") + n; }
  return d.getFullYear() + "-" + dois(d.getMonth() + 1) + "-" + dois(d.getDate());
}

function bannerAtivo(hoje) {
  var b = CONFIG.banner || {};
  if (!b.imagem) return null;
  var dia = dataISO(hoje || new Date());
  if (b.inicio && dia < b.inicio) return null;
  if (b.fim && dia > b.fim) return null;
  return b;
}

function montarAssinatura(p, tipo, hoje) {
  var cor = esc(CONFIG.cor || "#1a1a1a");
  var empresa = esc(CONFIG.empresa || "");
  var fonte = "font-family:Arial,Helvetica,sans-serif;";
  var linhaCargo = [p.cargo, p.departamento].filter(Boolean).map(esc).join(" &middot; ");

  var telefones = [];
  if (p.telefone || CONFIG.telefoneGeral) telefones.push("Tel " + telLink(p.telefone || CONFIG.telefoneGeral));
  if (p.celular) telefones.push("Cel " + telLink(p.celular));

  if (tipo === "resposta") {
    var curta = '<div style="' + fonte + 'font-size:13px;color:#333333;line-height:1.5;">' +
      '<span style="font-weight:bold;color:' + cor + ';">' + esc(p.nome) + "</span>";
    var linha2 = [linhaCargo, empresa].filter(Boolean).join(" | ");
    if (linha2) curta += "<br>" + linha2;
    if (telefones.length) curta += "<br>" + telefones.join(" &nbsp;|&nbsp; ");
    return curta + "</div>";
  }

  var info = '<div style="font-size:16px;font-weight:bold;color:' + cor + ';">' + esc(p.nome) + "</div>";
  if (linhaCargo) info += '<div style="color:#555555;">' + linhaCargo + "</div>";
  info += '<div style="height:8px;line-height:8px;font-size:1px;">&nbsp;</div>';
  if (telefones.length) info += "<div>" + telefones.join(" &nbsp;|&nbsp; ") + "</div>";
  if (p.email) info += '<div><a href="mailto:' + esc(p.email) + '" style="color:#333333;text-decoration:none;">' + esc(p.email) + "</a></div>";
  if (CONFIG.site) info += '<div><a href="' + esc(urlCompleta(CONFIG.site)) + '" style="color:' + cor + ';text-decoration:none;font-weight:bold;">' + esc(urlSemProtocolo(CONFIG.site)) + "</a></div>";
  if (CONFIG.endereco) info += '<div style="color:#777777;font-size:12px;">' + esc(CONFIG.endereco) + "</div>";

  var redes = [];
  if (CONFIG.linkedin) redes.push('<a href="' + esc(urlCompleta(CONFIG.linkedin)) + '" style="color:' + cor + ';text-decoration:none;">LinkedIn</a>');
  if (CONFIG.instagram) redes.push('<a href="' + esc(urlCompleta(CONFIG.instagram)) + '" style="color:' + cor + ';text-decoration:none;">Instagram</a>');
  if (redes.length) info += '<div style="margin-top:6px;font-size:12px;">' + redes.join(" &nbsp;&middot;&nbsp; ") + "</div>";

  var html = '<table cellpadding="0" cellspacing="0" border="0" style="' + fonte + 'font-size:13px;color:#333333;line-height:1.45;border-collapse:collapse;">' +
    "<tr>";
  if (CONFIG.logo) {
    html += '<td style="vertical-align:top;padding:0 16px 0 0;border-right:3px solid ' + cor + ';">' +
      '<img src="' + esc(CONFIG.logo) + '" alt="' + empresa + '" width="' + (parseInt(CONFIG.larguraLogo, 10) || 120) + '" style="display:block;border:0;"></td>' +
      '<td style="vertical-align:top;padding:0 0 0 16px;">';
  } else {
    html += '<td style="vertical-align:top;padding:0 0 0 12px;border-left:3px solid ' + cor + ';">';
  }
  html += info + "</td></tr></table>";

  var banner = bannerAtivo(hoje);
  if (banner) {
    var img = '<img src="' + esc(banner.imagem) + '" alt="' + esc(banner.alt || empresa) + '" width="' + (parseInt(banner.largura, 10) || 500) + '" style="display:block;border:0;max-width:100%;">';
    html += '<div style="margin-top:12px;">' + (banner.link ? '<a href="' + esc(urlCompleta(banner.link)) + '">' + img + "</a>" : img) + "</div>";
  }

  if (CONFIG.aviso) {
    html += '<div style="' + fonte + 'margin-top:12px;max-width:560px;font-size:10px;line-height:1.4;color:#999999;">' + esc(CONFIG.aviso) + "</div>";
  }
  return html;
}

function dadosDoUsuario() {
  var perfil = Office.context.mailbox.userProfile;
  var email = String(perfil.emailAddress || "").toLowerCase();
  var p = FUNCIONARIOS[email];
  if (!p) p = { nome: perfil.displayName || "" };
  var copia = {};
  for (var k in p) if (Object.prototype.hasOwnProperty.call(p, k)) copia[k] = p[k];
  if (!copia.nome) copia.nome = perfil.displayName || "";
  if (!copia.email) copia.email = perfil.emailAddress || "";
  return copia;
}

function onNewMessageComposeHandler(event) {
  var concluido = false;
  function concluir() {
    if (!concluido) { concluido = true; event.completed(); }
  }

  try {
    var item = Office.context.mailbox.item;
    var pessoa = dadosDoUsuario();

    var aplicar = function (tipo) {
      var html = montarAssinatura(pessoa, tipo, new Date());
      var inserir = function () {
        item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function () { concluir(); });
      };
      // Desliga a assinatura própria do Outlook neste e-mail para não duplicar.
      if (item.disableClientSignatureAsync) item.disableClientSignatureAsync(function () { inserir(); });
      else inserir();
    };

    if (CONFIG.assinaturaCurtaNasRespostas !== false && item.getComposeTypeAsync) {
      item.getComposeTypeAsync(function (r) {
        var tipo = "completa";
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.composeType && r.value.composeType !== "newMail") {
          tipo = "resposta";
        }
        aplicar(tipo);
      });
    } else {
      aplicar("completa");
    }
  } catch (e) {
    concluir();
  }
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { montarAssinatura: montarAssinatura, onNewMessageComposeHandler: onNewMessageComposeHandler, CONFIG: CONFIG, FUNCIONARIOS: FUNCIONARIOS };
}
