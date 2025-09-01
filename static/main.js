// ========= Helper: manter/reativar a mesma aba do WhatsApp =========
let zapWin = null;
const ZAP_WINDOW_NAME = "WHATS_ZHAZ";

function openOrReuseWhats(phoneE164, text) {
  const digits = String(phoneE164 || "").replace(/\D/g, "");
  const sendUrl = `https://web.whatsapp.com/send?${digits ? `phone=${digits}&` : ""}text=${encodeURIComponent(text || "")}`;

  // 1) Já temos a aba sob controle?
  if (zapWin && !zapWin.closed) {
    try {
      zapWin.location.replace(sendUrl); // navega sem empilhar histórico
      zapWin.focus();
      return;
    } catch (e) { /* continua */ }
  }

  // 2) Tentar “reatar” com a aba nomeada se o handle foi perdido
  const existing = window.open("", ZAP_WINDOW_NAME);
  if (existing && !existing.closed) {
    zapWin = existing;
    try {
      zapWin.location.replace(sendUrl);
      zapWin.focus();
      return;
    } catch (e) { /* continua */ }
  }

  // 3) Não existe nossa aba: cria a aba NOMEADA no WhatsApp raiz
  //    e só depois navega para /send (evita abrir duas guias)
  zapWin = window.open("https://web.whatsapp.com", ZAP_WINDOW_NAME);
  if (zapWin) {
    // pequena espera para a aba inicializar e anexar o nome corretamente
    setTimeout(() => {
      try {
        zapWin.location.href = sendUrl;
        zapWin.focus();
      } catch (e) {
        // fallback, reabre se algo bloquear
        zapWin = window.open(sendUrl, ZAP_WINDOW_NAME);
      }
    }, 300);
  } else {
    alert("O navegador bloqueou o pop-up do WhatsApp Web. Permita pop-ups para este site.");
  }
}


// ================== UPLOAD ==================
async function uploadFiles() {
  const fr = document.getElementById("f-rollout").files[0];
  const fl = document.getElementById("f-lojas").files[0];
  const fd = new FormData();
  if (fr) fd.append("rollout", fr);
  if (fl) fd.append("lojas", fl);

  const r = await fetch("/upload", { method: "POST", body: fd });
  const j = await r.json();
  alert(j.ok ? "Upload concluído!" : "Falha no upload.");
  location.reload();
}

// ================== BUSCAR ==================
async function buscar() {
  const numInput = document.getElementById("numero");
  const alertBox = document.getElementById("alert");
  const ta = document.getElementById("mensagem");
  const whats = document.getElementById("whats");

  let num = (numInput.value || "").trim();
  if (!num) { alert("Informe o número da loja."); return; }
  num = num.padStart(3, "0"); // ex.: 73 -> 073

  // reset UI
  alertBox.textContent = "";
  ta.value = "";
  whats.disabled = true;
  whats.onclick = null;
  whats.removeAttribute("href");
  whats.dataset.payload = "";

  const r = await fetch(`/buscar?numero=${encodeURIComponent(num)}`);
  if (!r.ok) {
    alertBox.textContent = `Erro ${r.status} ao buscar.`;
    return;
  }
  const j = await r.json();

  if (j && j.ok === false) {
    alertBox.textContent = j.error || "Não foi possível obter a loja.";
    ta.value = j.error ? `*** ${j.error} ***` : "";
    return;
  }

  // mensagem pronta
  ta.value = j.mensagem || "";

  // monta payload p/ registrar (usa primeiro item de "dados")
  const first = Array.isArray(j.dados) && j.dados.length ? j.dados[0] : {};
  const payload = {
    numero: first?.loja_numero ?? num,
    loja_nome: first?.loja_nome ?? "",
    regional: first?.regional ?? "",
    destinatario: j.destinatario || "",
    mensagem: j.mensagem || ""
  };
  whats.dataset.payload = JSON.stringify(payload);

  if (j.concluida) {
    // 100% → não dispara
    alertBox.textContent = `Loja ${num} está 100% concluída — não precisa de contato.`;
    whats.disabled = true;
    return;
  }

  // Há pendências → habilita botão e usa aba do WhatsApp reaproveitável
  alertBox.textContent = "Pendências encontradas. Clique para abrir o WhatsApp Web.";
  if (payload.destinatario && payload.mensagem) {
    whats.disabled = false;
    whats.onclick = (e) => {
      e.preventDefault();
      openOrReuseWhats(payload.destinatario, payload.mensagem);
    };
  }
}

// ================== REGISTRAR ENVIO ==================
async function registrarEnvio() {
  const whats = document.getElementById("whats");
  const payload = whats.dataset.payload ? JSON.parse(whats.dataset.payload) : null;
  if (!payload) { alert("Busque a loja antes de registrar."); return; }

  const r = await fetch("/log", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const j = await r.json();
  alert(j.ok ? "Envio registrado!" : "Falha ao registrar envio.");
}

// ================== EVENTOS ==================
document.getElementById("btn-enviar").addEventListener("click", uploadFiles);
document.getElementById("btn-buscar").addEventListener("click", buscar);
document.getElementById("btn-log").addEventListener("click", registrarEnvio);
