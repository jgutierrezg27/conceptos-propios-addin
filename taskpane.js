import { dictionary } from "./dictionary";

Office.onReady((info) => {
  console.log(`✅ Office está listo By Alfeizar. Host: ${info.host}, Plataforma: ${info.platform}`);

  const btnNormalizar = document.getElementById("btn-normalizar");

  window.mostrarEstado = function(mensaje, tipo = 'info') {
    const statusDiv = document.getElementById("status");
    const statusMessage = document.getElementById("status-message");
    if (!statusDiv || !statusMessage) return;
    statusMessage.textContent = mensaje;
    statusDiv.style.display = "block";
    if (tipo === "success") {
      statusDiv.style.backgroundColor = "#dff6dd";
      statusDiv.style.color = "#107c10";
      statusDiv.style.border = "1px solid #107c10";
      setTimeout(() => { statusDiv.style.display = "none"; }, 5000);
    } else if (tipo === "error") {
      statusDiv.style.backgroundColor = "#fde7e9";
      statusDiv.style.color = "#a80000";
      statusDiv.style.border = "1px solid #a80000";
    } else {
      statusDiv.style.backgroundColor = "rgba(255,255,255,0.08)";
      statusDiv.style.color = "rgba(255,255,255,0.85)";
      statusDiv.style.border = "1px solid rgba(255,255,255,0.2)";
    }
  };

  if (btnNormalizar) {
    btnNormalizar.addEventListener("click", normalizarTextoWord);
  }
});

const isWordChar = (ch) => {
  if (!ch) return false;
  return /\p{L}|\p{N}/u.test(ch);
};

function boundaryOk(fullText, matchLower) {
  const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(escapeRegExp(matchLower), 'giu');
  let m;
  while ((m = regex.exec(fullText)) !== null) {
    const prevChar = m.index > 0 ? fullText[m.index - 1] : null;
    const nextChar = m.index + matchLower.length < fullText.length ? fullText[m.index + matchLower.length] : null;
    if (!isWordChar(prevChar) && !isWordChar(nextChar)) return true;
  }
  return false;
}

async function normalizarTextoWord() {
  const btnNormalizar = document.getElementById("btn-normalizar");

  if (btnNormalizar) {
    btnNormalizar.dataset._origHtml = btnNormalizar.innerHTML;
    btnNormalizar.disabled = true;
    btnNormalizar.innerHTML = `<svg class="spinner" width="16" height="16" viewBox="0 0 50 50" aria-hidden="true"><circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="4" stroke-linecap="round" stroke-dasharray="31.4 31.4" transform="rotate(-90 25 25)"></circle></svg> Corrigiendo documento`;
  }

  function restoreButton() {
    if (!btnNormalizar) return;
    if (btnNormalizar.dataset._origHtml) {
      btnNormalizar.innerHTML = btnNormalizar.dataset._origHtml;
      delete btnNormalizar.dataset._origHtml;
    }
    btnNormalizar.disabled = false;
  }

  try {
    console.log("🔄 Iniciando normalización en Word...");
    mostrarEstado("Procesando...", "info");

    // --- Paso 1: construir mapa de variantes ---
    const matchToRule = new Map();
    const allMatches = [];
    for (const rule of dictionary) {
      for (const m of rule.matches) {
        const key = m.toLocaleLowerCase();
        if (!matchToRule.has(key)) {
          matchToRule.set(key, rule);
          allMatches.push(m);
        }
      }
    }

    if (allMatches.length === 0) {
      mostrarEstado('No hay reglas en el diccionario', 'info');
      restoreButton();
      return;
    }

    allMatches.sort((a, b) => b.length - a.length);

    // --- Paso 2: obtener TODO el texto del documento en UN solo sync ---
    // Esto nos permite filtrar localmente qué variantes realmente aparecen
    // antes de hacer búsquedas en Word — elimina el 90%+ de búsquedas innecesarias
    let fullText = '';
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      fullText = body.text.normalize('NFC').toLowerCase();
    });

    // --- Paso 3: filtrar solo las variantes que aparecen en el texto ---
    const presentMatches = allMatches.filter(m => {
      const key = m.normalize('NFC').toLowerCase();
      if (!fullText.includes(key)) return false;
      return boundaryOk(fullText, key);
    });

    console.log(`📋 Variantes en diccionario: ${allMatches.length} | Con match real: ${presentMatches.length}`);

    if (presentMatches.length === 0) {
      mostrarEstado("No se encontraron textos que coincidan con el diccionario", "info");
      restoreButton();
      return;
    }

    // --- Paso 4: buscar y reemplazar SOLO las variantes presentes, en lotes ---
    let cambiosRealizados = 0;
    const BATCH_SIZE = 50;

    await Word.run(async (context) => {
      for (let batchStart = 0; batchStart < presentMatches.length; batchStart += BATCH_SIZE) {
        const batch = presentMatches.slice(batchStart, batchStart + BATCH_SIZE);

        // Lanzar todas las búsquedas del lote sin sync
        const searchResultsMap = [];
        for (const matchText of batch) {
          const rule = matchToRule.get(matchText.toLocaleLowerCase());
          if (!rule) continue;

          const searchResults = context.document.body.search(matchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          searchResults.load('items/text,items/paragraphs/text');
          searchResultsMap.push({ matchText, rule, searchResults });
        }

        // Un solo sync para todo el lote
        await context.sync();

        // Aplicar reemplazos con boundary check sobre el párrafo
        for (const { matchText, rule, searchResults } of searchResultsMap) {
          if (searchResults.items.length === 0) continue;

          for (const range of searchResults.items) {
            const paragraphText = range.paragraphs.items.length > 0
              ? range.paragraphs.items.map(p => p.text).join(' ').normalize('NFC').toLowerCase()
              : range.text.normalize('NFC').toLowerCase();

            if (!boundaryOk(paragraphText, matchText.normalize('NFC').toLowerCase())) continue;

            const parts = rule.replacement;
            if (parts.length === 1) {
              const inserted = range.insertText(parts[0].text, Word.InsertLocation.replace);
              inserted.font.bold = parts[0].bold || false;
            } else {
              const firstRange = range.insertText(parts[0].text, Word.InsertLocation.replace);
              firstRange.font.bold = parts[0].bold || false;
              let currentRange = firstRange;
              for (let i = 1; i < parts.length; i++) {
                const inserted = currentRange.insertText(parts[i].text, Word.InsertLocation.after);
                inserted.font.bold = parts[i].bold || false;
                currentRange = inserted;
              }
            }
            cambiosRealizados++;
          }
        }

        // Un solo sync para aplicar los reemplazos del lote
        await context.sync();
      }
    });

    if (cambiosRealizados > 0) {
      console.log(`✅ ${cambiosRealizados} cambio(s) realizado(s).`);
      mostrarEstado(`✅ ${cambiosRealizados} cambio(s) realizado(s)`, "success");
    } else {
      console.log("ℹ️ No se encontraron coincidencias.");
      mostrarEstado("No se encontraron textos que coincidan con el diccionario", "info");
    }

  } catch (error) {
    console.error("❌ Error en Word:", error);
    mostrarEstado("Error: " + error.message, "error");
  } finally {
    restoreButton();
  }
}