const API_URL = "https://your-firebase-functions-url.cloudfunctions.net/processText";

Office.onReady(() => {
  document.getElementById("insertTemplateBtn").onclick = insertTemplate;
  document.getElementById("rewriteBtn").onclick = rewriteSelectedText;
});

async function insertTemplate() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("## Software Requirements Document", Word.InsertLocation.end);
    body.insertParagraph("### Overview", Word.InsertLocation.end);
    body.insertParagraph("Describe the project goals and purpose here.", Word.InsertLocation.end);
    body.insertParagraph("### Stakeholders", Word.InsertLocation.end);
    body.insertTable(3, 2, Word.InsertLocation.end, (table) => {
      table.getCell(0,0).insertText("Role", Word.InsertLocation.start);
      table.getCell(0,1).insertText("Contact", Word.InsertLocation.start);
      table.getCell(1,0).insertText("Developer", Word.InsertLocation.start);
      table.getCell(1,1).insertText("dev@example.com", Word.InsertLocation.start);
      table.getCell(2,0).insertText("Product Owner", Word.InsertLocation.start);
      table.getCell(2,1).insertText("po@example.com", Word.InsertLocation.start);
    }).load();
    await context.sync();

    // Name the table for audience extraction
    const tables = context.document.body.tables;
    tables.load();
    await context.sync();
    const lastTable = tables.items[tables.items.length - 1];
    lastTable.title = "AudienceTable"; // store role info here

    await context.sync();
  });
}

async function rewriteSelectedText() {
  document.getElementById("loading").style.display = "block";
  document.getElementById("output").textContent = "";

  try {
    const selectedText = await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text;
    });

    if (!selectedText || selectedText.trim().length === 0) {
      alert("Please select some text to rewrite.");
      document.getElementById("loading").style.display = "none";
      return;
    }

    const audience = await getAudienceRoles();

    const payload = {
      text: selectedText,
      audienceRoles: audience,
      mode: "rewrite"
    };

    const response = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await response.json();

    if (data.result) {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(data.result, "Replace");
        await context.sync();
      });

      document.getElementById("output").textContent = "Rewrite successful!";
    } else {
      document.getElementById("output").textContent = "No result from AI.";
    }
  } catch (error) {
    document.getElementById("output").textContent = "Error: " + error.message;
  }

  document.getElementById("loading").style.display = "none";
}

async function getAudienceRoles() {
  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items/title, items/rowCount, items/columnCount");
    await context.sync();

    for (const table of tables.items) {
      if (table.title === "AudienceTable") {
        table.load("rows/items/cells/items/body/text");
        await context.sync();

        const roles = [];
        // Skip header row (assumed at index 0)
        for (let i = 1; i < table.rows.items.length; i++) {
          const roleCell = table.rows.items[i].cells.items[0];
          roleCell.load("body/text");
          await context.sync();
          roles.push(roleCell.body.text.trim());
        }
        return roles;
      }
    }
    return [];
  });
}
