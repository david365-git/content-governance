
function codeshowSidebar() {
  const html = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h3>ce_ File Bundler</h3>
      <p>Because Google is blocking the automatic "All-File" read, please follow these 3 quick steps:</p>
      <ol>
        <li>Click a <b>ce_</b> file in your left sidebar.</li>
        <li>Press <b>Ctrl+A</b> then <b>Ctrl+C</b>.</li>
        <li>Paste it into the box below and click "Add to Bundle".</li>
      </ol>
      <textarea id="input" style="width:100%; height:100px;"></textarea><br><br>
      <button onclick="addFile()" style="background: #4285f4; color: white; border: none; padding: 10px; cursor: pointer;">Add to Bundle</button>
      <hr>
      <p><b>Your Final Bundle (Paste this to NotebookLM):</b></p>
      <textarea id="output" style="width:100%; height:200px; background: #f0f0f0;" readonly></textarea>
      <script>
        function addFile() {
          const input = document.getElementById('input');
          const output = document.getElementById('output');
          const fileName = prompt("What is the filename? (e.g. ce_Main.gs)");
          if (fileName) {
            output.value += "\\n\\n--- FILE_START: " + fileName + " ---\\n\\n" + input.value + "\\n--- FILE_END ---";
            input.value = "";
            alert(fileName + " added!");
          }
        }
      </script>
    </div>
  `;
  const userInterface = HtmlService.createHtmlOutput(html).setTitle('NotebookLM Bundler').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(userInterface);
}