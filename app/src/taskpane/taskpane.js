let startPosition = "";
let endPosition = "";
let newText = "";

async function sendToOllama(textToEdit, userPrompt) {
  const sys_prompt =
    'В ответе используй русский язык. Должен быть только один вариант. В ответе не должно быть фраз - "вот переписанный фрагмент", или "вот что получилось"';

  const fullPrompt = `\n\n${sys_prompt}\n\n${textToEdit}\n\n${userPrompt}`;
  const response = await fetch("https://localhost/ollama/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "gemma3:1b",
      options: {
        num_ctx: 32000,
      },
      prompt: fullPrompt,
      stream: false,
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Ollama API error ${response.status}: ${text}`);
  }
  const result = await response.json();
  return result.response;
}

function markdownToWord(context, markdown) {
  const parsedHtml = marked.parse(markdown);
  const range = context.document.getSelection();
  range.insertHtml(parsedHtml, Word.InsertLocation.replace);
}

Office.onReady(() => {
  document.getElementById("sendButton").onclick = async function () {
    const userPrompt = document.getElementById("userInput").value;
    document.getElementById("responseOutput").innerText = "⌛ Анализируем текст...";
    document.getElementById("diffPreview").style.display = "none";

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.font.highlightColor = Word.HighlightColor.yellow;
        selection.load("text, start, end");
        await context.sync();

        startPosition = selection.start
        endPosition = selection.end

        let textToEdit = selection.text?.trim();

        if (!textToEdit) {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          textToEdit = body.text.trim();
        }

        const markedText = textToEdit;

        newText = await sendToOllama(textToEdit, userPrompt);
        document.getElementById("err2").innerText = startPosition + ", " + endPosition;


        document.getElementById("originalText").innerText = markedText;
        document.getElementById("updatedText").innerText = newText;
        document.getElementById("diffPreview").style.display = "block";
        document.getElementById("responseOutput").innerText = "✅ Готово. Подтвердите изменения.";
      });
    } catch (err) {
      console.error("❌ Ошибка:", err);
      document.getElementById("responseOutput").innerText = "Ошибка: " + err.message;
    }
  };

  document.getElementById("applyChangesButton").onclick = async function () {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();
        selection.load("text, font/highlightColor");
        await context.sync();

        if (typeof newText !== "string" || newText.length === 0) {
          throw new Error("Нет сгенерированного текста для вставки");
        }

        if (selection.text && selection.text.length > 0) {
          // Remove highlight first to ensure it's cleared even if selection collapses after replace
          selection.font.highlightColor = Word.HighlightColor.noColor;
          selection.insertText(newText, Word.InsertLocation.replace);
        } else {
          body.insertText(newText, Word.InsertLocation.end);
        }

        document.getElementById("diffPreview").style.display = "none";
        document.getElementById("responseOutput").innerText = "✅ Изменения применены.";
        await context.sync();
      });
    } catch (err) {
      console.error("❌ Ошибка применения:", err);
      document.getElementById("responseOutput").innerText = "Ошибка применения: " + err.message;
    }
  };

  // Reset highlight button
  const resetBtn = document.getElementById("reset");
  if (resetBtn) {
    resetBtn.onclick = async function () {
      try {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.font.highlightColor = Word.HighlightColor.noColor;

          const body = context.document.body;
          const search = body.search("*", { matchWildcards: true });
          search.load("items/font/highlightColor");
          await context.sync();

          for (let i = 0; i < search.items.length; i++) {
            const r = search.items[i];
            if (r.font && r.font.highlightColor && r.font.highlightColor !== Word.HighlightColor.noColor) {
              r.font.highlightColor = Word.HighlightColor.noColor;
            }
          }

          document.getElementById("diffPreview").style.display = "none";
          document.getElementById("responseOutput").innerText = "Подсветка сброшена.";
          await context.sync();
        });
      } catch (err) {
        console.error("❌ Ошибка сброса подсветки:", err);
        document.getElementById("responseOutput").innerText = "Ошибка сброса: " + err.message;
      }
    };
  }

  // Cancel preview button also clears selection highlight
  const cancelBtn = document.getElementById("cancelChangesButton");
  if (cancelBtn) {
    cancelBtn.onclick = async function () {
      try {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.font.highlightColor = Word.HighlightColor.noColor;
          document.getElementById("diffPreview").style.display = "none";
          document.getElementById("responseOutput").innerText = "Изменения отменены.";
          await context.sync();
        });
      } catch (err) {
        console.error("❌ Ошибка отмены изменений:", err);
        document.getElementById("responseOutput").innerText = "Ошибка отмены: " + err.message;
      }
    };
  }
});
